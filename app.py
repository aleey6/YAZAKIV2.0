"""Interface web Streamlit pour l'extracteur de factures IAM YAZAKI.

Lancement :
    streamlit run app.py
"""

from __future__ import annotations

import io
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# Imports des modules refactorisés
from config import load_config, safe_filename, fmt_money, DATA_DIR, ROOT_DIR
from ui_components import apply_custom_css, init_session_state, render_info_card
from excel_export import contracts_to_excel
from budget_utils import read_budget_from_excel
from invoice_utils import (
    contracts_dataframe, 
    append_to_master, 
    list_saved_invoices, 
    save_invoice_record,
    load_ocr_engine
)
from invoice_parser import parse_invoice
from expenses import calculate_expenses
import database as db  # Votre base de données existante

# Initialisation
init_session_state()
apply_custom_css()
db.init_db()

# Configuration de la page
st.set_page_config(
    page_title="YAZAKI · Extracteur de factures IAM",
    page_icon="🟥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Logo
logo_path = ROOT_DIR / "public" / "logo.png"
if logo_path.exists():
    st.logo(str(logo_path), size="large")

# En-tête
st.title("YAZAKI · Extracteur de factures IAM")
st.caption(
    "Téléchargez une facture Maroc Telecom multipage → catégorisez-la par "
    "Usine / Département / Projet → enregistrez en JSON → calculez "
    "Budget vs. Dépenses."
)

# Barre latérale
config = load_config()
plants_data = config.get("plants", {})

with st.sidebar:
    st.header("Catégorisation")
    st.caption("Choisissez l'Usine, le Département et le Projet pour cette facture.")
    
    plant = st.selectbox(
        "Usine",
        options=[""] + list(plants_data.keys()),
        format_func=lambda x: x if x else "— Choisir une usine —",
        key="plant",
    )
    
    departments = list(plants_data.get(plant, {}).keys()) if plant else []
    dept = st.selectbox(
        "Département",
        options=[""] + departments,
        format_func=lambda x: x if x else "— Choisir un département —",
        disabled=not plant,
        key="department",
    )
    
    projects = plants_data.get(plant, {}).get(dept, []) if (plant and dept) else []
    project = st.selectbox(
        "Projet",
        options=[""] + list(projects),
        format_func=lambda x: x if x else "— Choisir un projet —",
        disabled=not (plant and dept),
        key="project",
    )
    
    st.divider()
    st.caption("Modifiez les options dans `config.json`, puis actualisez la page.")

# Onglets
tab_extract, tab_expenses, tab_history = st.tabs(
    ["Extraire la facture", "Calculer les dépenses", "Historique"]
)

# ============================================================== ONGLET EXTRAIRE
with tab_extract:
    upload_col, info_col = st.columns([2, 1])

    with upload_col:
        uploaded = st.file_uploader(
            "Téléchargez une facture IAM multipage (PDF)",
            type=["pdf"],
            accept_multiple_files=False,
        )

    with info_col:
        render_info_card(
            "Calcul du Total",
            "Somme du TOTAL CONTRAT de chaque page de contrat — "
            "c.-à-d. toutes les lignes sauf la première (récapitulatif global), "
            "projetées sur la colonne Total Contrat."
        )

    # Stocker les bytes du PDF
    if uploaded is not None:
        file_key = f"pdf_bytes_{uploaded.name}_{uploaded.size}"
        if st.session_state.get("pdf_file_key") != file_key:
            st.session_state.pdf_bytes = uploaded.read()
            st.session_state.pdf_file_key = file_key
        pdf_bytes = st.session_state.pdf_bytes
    else:
        pdf_bytes = None
        st.session_state.pdf_bytes = None
        st.session_state.pdf_file_key = None

    # Détection OCR
    ocr_available = True
    ocr_import_error = None
    try:
        import easyocr  # noqa: F401
        import fitz     # noqa: F401
    except ImportError as e:
        ocr_available = False
        ocr_import_error = str(e)

    use_ocr = False
    if pdf_bytes is not None:
        try:
            import fitz
            pdf_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
            is_scanned = all(
                len(pdf_preview[i].get_text().strip()) < 80
                for i in range(min(3, len(pdf_preview)))
            )
            pdf_preview.close()
        except Exception:
            is_scanned = False

        if is_scanned and ocr_available:
            st.warning(
                "⚠️ Ce PDF semble **scanné** (pas de texte extractible). "
                "Le mode OCR a été activé automatiquement."
            )
            use_ocr = True
        elif is_scanned and not ocr_available:
            st.error(
                f"⚠️ Ce PDF semble scanné mais OCR non disponible : `{ocr_import_error}`. "
                "Lancez `pip install easyocr PyMuPDF` puis redémarrez."
            )
        elif not ocr_available:
            st.caption(f"ℹ️ OCR non disponible (`{ocr_import_error}`). Installez easyocr et PyMuPDF.")
        elif ocr_available:
            use_ocr = st.toggle(
                "🔍 Forcer le mode OCR (PDF scanné)",
                value=False,
                help="Activez si des pages sont manquantes ou vides dans le résultat.",
            )

    extract_clicked = st.button(
        "Extraire les données",
        type="primary",
        disabled=pdf_bytes is None,
        use_container_width=False,
    )

    if extract_clicked and pdf_bytes is not None:
        ocr_engine = None
        if use_ocr:
            with st.spinner("⏳ Chargement du moteur OCR..."):
                ocr_engine, ocr_error = load_ocr_engine()
            if ocr_engine is None:
                st.error(f"❌ Impossible de charger EasyOCR : `{ocr_error}`")
                st.stop()

        with st.spinner("Analyse du PDF en cours..." if not use_ocr else "🔍 OCR en cours..."):
            progress_bar = st.progress(0, text="Démarrage…")
            status_text = st.empty()

            def on_progress(current, total, status):
                pct = int((current / total) * 100) if total else 0
                progress_bar.progress(pct, text=f"{status} ({pct}%)")
                status_text.caption(status)

            try:
                buf = io.BytesIO(pdf_bytes)
                invoice = parse_invoice(buf, ocr_engine=ocr_engine,
                                        progress_callback=on_progress)
                invoice.source_file = uploaded.name if uploaded else "inconnu"
                st.session_state.invoice = invoice
                st.session_state.uploaded_name = invoice.source_file

                progress_bar.progress(100, text="✅ Extraction terminée !")
                status_text.empty()

                if len(invoice.contracts) == 0:
                    st.warning("⚠️ 0 contrats extraits.")
                else:
                    st.success(f"✅ {len(invoice.contracts)} pages de contrat extraites.")
            except Exception as exc:
                progress_bar.empty()
                status_text.empty()
                st.session_state.invoice = None
                st.error(f"Échec de l'analyse du PDF : {exc}")

    invoice = st.session_state.invoice

    if invoice is not None:
        st.subheader("Récapitulatif")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("N° Facture", invoice.invoice_number or "—")
        m2.metric("Date Facture", invoice.invoice_date or "—")
        m3.metric("Nombre de contrats", len(invoice.contracts))
        m4.metric("Total (MAD)", fmt_money(invoice.total))

        st.subheader("Détail des contrats")
        df = contracts_dataframe(invoice)
        st.dataframe(df, use_container_width=True, hide_index=True)

        # Téléchargement Excel
        excel_bytes = contracts_to_excel(invoice, plant, dept, project)
        st.download_button(
            "📥 Télécharger le détail des contrats (Excel)",
            data=excel_bytes,
            file_name=safe_filename(
                f"{plant or 'usine'}_{dept or 'dept'}_{project or 'projet'}"
                f"_{invoice.invoice_number or 'contrats'}"
            ) + "_contrats.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Afficher le JSON complet extrait"):
            st.json(invoice.to_dict(), expanded=False)

        st.divider()
        st.subheader("Enregistrer cette facture")

        record = save_invoice_record(invoice, plant, dept, project)
        json_bytes = json.dumps(record, indent=2, ensure_ascii=False).encode("utf-8")

        # Auto-sauvegarde
        if plant and dept and project:
            invoice_key = f"{plant}_{dept}_{project}_{invoice.invoice_number}"
            if not st.session_state.get(f"saved_{invoice_key}"):
                DATA_DIR.mkdir(exist_ok=True)
                filename = safe_filename(f"{plant}_{dept}_{project}_{invoice.invoice_number or 'facture'}") + ".json"
                per_file = DATA_DIR / filename
                per_file.write_text(json.dumps(record, indent=2, ensure_ascii=False), encoding="utf-8")
                append_to_master(record)
                try:
                    invoice_id = db.save_invoice(record)
                    st.session_state[f"saved_{invoice_key}"] = True
                    st.success(f"✅ Sauvegarde automatique effectuée — SQLite (id={invoice_id})")
                except Exception as exc:
                    st.warning(f"Fichiers JSON enregistrés, mais SQLite a échoué : {exc}")
            else:
                st.success(f"✅ Déjà sauvegardé pour {plant} / {dept} / {project}.")

            st.download_button(
                "📥 Télécharger le JSON",
                data=json_bytes,
                file_name=safe_filename(f"{plant}_{dept}_{project}_{invoice.invoice_number or 'facture'}") + ".json",
                mime="application/json",
            )
        else:
            st.info("Choisissez Usine → Département → Projet dans la barre latérale pour activer la sauvegarde.")
            st.download_button("📥 Télécharger le JSON (sans catégorie)", data=json_bytes, file_name="facture.json", mime="application/json")
    else:
        st.info("Téléchargez un PDF puis cliquez sur Extraire les données pour commencer.")
# ============================================================ ONGLET DÉPENSES
with tab_expenses:
    st.subheader("Calculer le budget restant")
    
    if not (plant and dept and project):
        st.info("⚠️ Veuillez sélectionner une Usine, un Département et un Projet dans la barre latérale.")
    else:
        budget_file = st.file_uploader(
            f"Téléchargez le fichier Budget pour {plant} / {dept} / {project}", 
            type=["xlsx", "xls"]
        )
        
        if budget_file is not None:
            try:
                with st.spinner("Analyse de la matrice budgétaire..."):
                    # On utilise .read() car budget_utils attend des octets
                    budget_data = read_budget_from_excel(budget_file.read(), plant, dept, project)
                
                if budget_data is not None:
                    st.success("✅ Données budgétaires chargées avec succès !")
                    
                    allocated_budget = budget_data["budget"]
                    cc_id = budget_data["cost_center_id"] or "N/A"
                    
                    st.markdown(f"**Centre de coût (Cost Center ID) :** `{cc_id}`")
                    
                    # Récupération de la facture en cours
                    current_invoice = st.session_state.get("invoice")
                    
                    if current_invoice is not None:
                        st.subheader("📊 Écart Budget vs Dépenses")
                        
                        # CORRECTION ICI : On convertit l'objet InvoiceData en dictionnaire
                        invoice_dict = current_invoice.to_dict()
                        
                        # Appel de la fonction avec le dictionnaire compatible
                        analysis = calculate_expenses(invoice_dict, allocated_budget)
                        
                        total_invoice = analysis["total"]
                        remaining_budget = analysis["expenses"]
                        
                        # Affichage des métriques (KPIs)
                        kpi1, kpi2, kpi3 = st.columns(3)
                        with kpi1:
                            st.metric(label="Budget Alloué", value=fmt_money(allocated_budget))
                        with kpi2:
                            st.metric(label="Dépenses Facture (TTC)", value=fmt_money(total_invoice), delta=f"{fmt_money(total_invoice)} Utilisés", delta_color="inverse")
                        with kpi3:
                            color_status = "normal" if remaining_budget >= 0 else "inverse"
                            st.metric(label="Budget Restant Évalué", value=fmt_money(remaining_budget), delta=f"{fmt_money(remaining_budget)} Restants", delta_color=color_status)
                            
                        # Barre de progression
                        if allocated_budget > 0:
                            progress_percent = min(1.0, max(0.0, total_invoice / allocated_budget))
                            st.progress(progress_percent, text=f"Utilisation du budget : {progress_percent * 100:.1f}%")
                            if remaining_budget < 0:
                                st.error(f"🚨 Dépassement budgétaire détecté de {fmt_money(abs(remaining_budget))} !")
                    else:
                        st.info("💡 Facture non détectée. Importez et extrayez d'abord une facture Maroc Telecom dans l'onglet principal pour calculer l'écart.")
                        st.metric(label="Budget Alloué Trouvé", value=fmt_money(allocated_budget))
                        
                else:
                    st.warning(f"⚠️ Aucune ligne correspondante trouvée dans le fichier Excel pour : {plant} -> {dept} -> {project}.")
                    
            except Exception as e:
                st.error(f"Erreur lors du calcul budgétaire : {e}")
        else:
            st.info("💡 Importez votre fichier Excel `BUDGET.xlsx` pour lancer la confrontation Budget vs Dépenses.")
# ========================================================== ONGLET HISTORIQUE
with tab_history:
    st.subheader("Factures enregistrées (base SQLite)")
    
    col_refresh, _ = st.columns([1, 4])
    with col_refresh:
        refresh_clicked = st.button("🔄 Actualiser l'historique", use_container_width=True)
        
    try:
        # Call the actual function from your database.py module
        records = db.list_invoices()
        
        if records:
            # Convert list of dicts to a clean DataFrame
            history_df = pd.DataFrame(records)
            
            # Optional: Clean up column display orders or names if needed
            # e.g., history_df = history_df[['id', 'plant', 'department', 'project', 'invoice_number', 'total', 'created_at']]
            
            st.dataframe(history_df, use_container_width=True, hide_index=True)
            
            # Metrics quick-glance helper
            st.caption(f"Total de factures indexées : {len(history_df)}")
        else:
            st.info("📂 Aucune facture n'a encore été enregistrée dans la base SQLite.")
            
    except Exception as e:
        st.error(f"Impossible de récupérer l'historique depuis la base de données : {e}")