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
from excel_export import contracts_to_excel, contracts_to_excel_batch
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
# ============================================================== ONGLET EXTRAIRE
with tab_extract:
    upload_col, info_col = st.columns([2, 1])

    with upload_col:
        uploaded_files = st.file_uploader(
            "Téléchargez une ou plusieurs factures IAM (PDF)",
            type=["pdf"],
            accept_multiple_files=True,
        )

    with info_col:
        render_info_card(
            "Calcul du Total",
            "Extraction en masse (Batch) : L'application va traiter chaque PDF l'un après l'autre."
        )

    # Stocker les bytes du PDF (pour compatibilité avec le code existant)
    if uploaded_files and len(uploaded_files) > 0:
        # Pour la compatibilité, on garde une référence au premier fichier
        first_file = uploaded_files[0]
        file_key = f"pdf_bytes_{first_file.name}_{first_file.size}"
        if st.session_state.get("pdf_file_key") != file_key:
            st.session_state.pdf_bytes = first_file.read()
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
        disabled=not uploaded_files,
        use_container_width=False,
    )

    # Initialiser une liste pour stocker toutes les factures extraites dans le session_state
    if "all_invoices" not in st.session_state:
        st.session_state.all_invoices = []

    if extract_clicked and uploaded_files:
        ocr_engine = None
        if use_ocr:
            with st.spinner("⏳ Chargement du moteur OCR..."):
                ocr_engine, ocr_error = load_ocr_engine()
            if ocr_engine is None:
                st.error(f"❌ Impossible de charger EasyOCR : `{ocr_error}`")
                st.stop()

        # Nettoyer les anciennes extractions f-l-glessa
        st.session_state.all_invoices = []
        
        # SÉCURITÉ : Vider l'ancienne clé de sauvegarde pour forcer SQLite à réenregistrer le nouveau batch
        for key in list(st.session_state.keys()):
            if key.startswith("saved_"):
                del st.session_state[key]
        
        # Boucle sur TOUS les fichiers uploadés
        for idx, file in enumerate(uploaded_files):
            st.markdown(f"### 📄 Traitement du fichier ({idx+1}/{len(uploaded_files)}) : `{file.name}`")
            
            progress_bar = st.progress(0, text="Démarrage…")
            status_text = st.empty()

            def on_progress(current, total, status):
                pct = int((current / total) * 100) if total else 0
                progress_bar.progress(pct, text=f"{status} ({pct}%)")
                status_text.caption(status)

            try:
                # Réinitialiser le pointeur du fichier et lire le contenu
                file.seek(0)
                pdf_bytes_current = file.read()
                buf = io.BytesIO(pdf_bytes_current)
                
                # Extraction du PDF actuel
                invoice = parse_invoice(buf, ocr_engine=ocr_engine, progress_callback=on_progress)
                invoice.source_file = file.name
                
                # 🟢 COPIER / COLLER ICI :
                from datetime import datetime
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                clean_filename = file.name.replace(".pdf", "").replace(".PDF", "").replace(" ", "_")
                invoice.invoice_number = f"{clean_filename}_{current_time}"
                # ----------------------------------------------------------------------------

                # Sauvegarder dans la liste globale de session
                st.session_state.all_invoices.append(invoice)

                progress_bar.progress(100, text="✅ Extraction terminée !")
                status_text.empty()
                
                if len(invoice.contracts) == 0:
                    st.warning(f"⚠️ 0 contrats extraits pour {file.name}.")
                else:
                    st.success(f"✅ {len(invoice.contracts)} pages de contrat extraites pour {file.name}.")
                
                # 5. AUTO-SAUVEGARDE DIRECTE DANS SQLITE POUR CHAQUE FICHIER DISTINCT
                if plant and dept and project:
                    record = save_invoice_record(invoice, plant, dept, project)
                    try:
                        invoice_id = db.save_invoice(record)
                        st.success(f"💾 Sauvegardé avec succès dans SQLite — Ligne ID: {invoice_id}")
                    except Exception as exc:
                        st.warning(f"Problème d'insertion dans SQLite pour {file.name} : {exc}")
                        
            except Exception as exc:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ Échec de l'analyse du fichier {file.name} : {exc}")
    # --- AFFICHAGE DU RÉCAPITULATIF GLOBAL ---
    invoices_list = st.session_state.all_invoices
    if invoices_list:
        st.divider()
        st.subheader("📊 Récapitulatif de l'extraction en masse")
        
        # Calculer le total cumulé de toutes les factures
        total_global = sum(inv.total for inv in invoices_list)
        
        m1, m2 = st.columns(2)
        m1.metric("Nombre de factures", len(invoices_list))
        m2.metric("Total cumulé (MAD)", fmt_money(total_global))

        # Combiner tous les dataframes de contrats en un seul
        all_dfs = []
        for inv in invoices_list:
            all_dfs.append(contracts_dataframe(inv))
        
        if all_dfs:
            import pandas as pd
            final_df = pd.concat(all_dfs, ignore_index=True)
            st.dataframe(final_df, use_container_width=True, hide_index=True)
            
            # Téléchargement Excel pour toutes les factures
            # Note: Cette partie nécessiterait une fonction pour combiner tous les contrats en Excel
            st.download_button(
                "📥 Télécharger toutes les données (Excel)",
                data=contracts_to_excel_batch(invoices_list, plant, dept, project),
                file_name=safe_filename(
                    f"{plant or 'usine'}_{dept or 'dept'}_{project or 'projet'}_batch_contrats"
                ) + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with st.expander("Afficher le JSON complet de toutes les factures"):
            all_invoices_dict = [inv.to_dict() for inv in invoices_list]
            st.json(all_invoices_dict, expanded=False)

        st.divider()
        st.subheader("Enregistrer toutes ces factures")

       
    else:
        st.info("Téléchargez un ou plusieurs PDF puis cliquez sur Extraire les données pour commencer.")
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
                    
                    # Récupération de la facture en cours (support single et batch)
                    current_invoice = st.session_state.get("invoice")
                    all_invoices = st.session_state.get("all_invoices", [])
                    
                    # Déterminer quelle facture utiliser
                    if current_invoice is not None:
                        # Mode single facture
                        invoices_to_analyze = [current_invoice]
                        st.info("📄 Analyse de la facture unique")
                    elif all_invoices:
                        # Mode batch - analyser toutes les factures extraites
                        invoices_to_analyze = all_invoices
                        st.info(f"📊 Analyse batch : {len(all_invoices)} facture(s) extraite(s)")
                    else:
                        invoices_to_analyze = []
                    
                    if invoices_to_analyze:
                        st.subheader("📊 Écart Budget vs Dépenses")
                        
                        # Calculer le total cumulé de toutes les factures
                        total_all_invoices = 0
                        all_analyses = []
                        
                        for inv in invoices_to_analyze:
                            # CORRECTION ICI : On convertit l'objet InvoiceData en dictionnaire
                            invoice_dict = inv.to_dict()
                            
                            # Appel de la fonction avec le dictionnaire compatible
                            analysis = calculate_expenses(invoice_dict, allocated_budget)
                            all_analyses.append(analysis)
                            total_all_invoices += analysis["total"]
                        
                        total_invoice = total_all_invoices
                        remaining_budget = allocated_budget - total_invoice
                        
                        # Affichage des métriques (KPIs)
                        kpi1, kpi2, kpi3 = st.columns(3)
                        with kpi1:
                            st.metric(label="Budget Alloué", value=fmt_money(allocated_budget))
                        with kpi2:
                            facture_label = f"Dépenses Facture{'s' if len(invoices_to_analyze) > 1 else ''} (TTC)"
                            delta_text = f"{fmt_money(total_invoice)} Utilisé{'s' if len(invoices_to_analyze) > 1 else ''}"
                            st.metric(label=facture_label, value=fmt_money(total_invoice), delta=delta_text, delta_color="inverse")
                        with kpi3:
                            color_status = "normal" if remaining_budget >= 0 else "inverse"
                            restant_label = f"Budget Restant Évalué ({len(invoices_to_analyze)} facture{'s' if len(invoices_to_analyze) > 1 else ''})"
                            st.metric(label=restant_label, value=fmt_money(remaining_budget), delta=f"{fmt_money(remaining_budget)} Restants", delta_color=color_status)
                            
                        # Barre de progression
                        if allocated_budget > 0:
                            progress_percent = min(1.0, max(0.0, total_invoice / allocated_budget))
                            st.progress(progress_percent, text=f"Utilisation du budget : {progress_percent * 100:.1f}%")
                            if remaining_budget < 0:
                                st.error(f"🚨 Dépassement budgétaire détecté de {fmt_money(abs(remaining_budget))} !")
                        
                        # Afficher le détail par facture si batch
                        if len(invoices_to_analyze) > 1:
                            with st.expander("📋 Détail par facture"):
                                detail_data = []
                                for idx, (inv, analysis) in enumerate(zip(invoices_to_analyze, all_analyses)):
                                    detail_data.append({
                                        "Facture": inv.invoice_number or f"Facture {idx+1}",
                                        "Date": inv.invoice_date or "—",
                                        "Montant TTC": fmt_money(analysis["total"]),
                                        "% Budget": f"{(analysis['total']/allocated_budget*100):.1f}%" if allocated_budget > 0 else "N/A"
                                    })
                                st.dataframe(detail_data, use_container_width=True)
                    else:
                        st.info("💡 Facture non détectée. Importez et extrayez d'abord une ou plusieurs factures Maroc Telecom dans l'onglet principal pour calculer l'écart.")
                        st.metric(label="Budget Alloué Trouvé", value=fmt_money(allocated_budget))
                        
                else:
                    st.warning(f"⚠️ Aucune ligne correspondante trouvée dans le fichier Excel pour : {plant} -> {dept} -> {project}.")
                    
            except Exception as e:
                st.error(f"Erreur lors du calcul budgétaire : {e}")
        else:
            st.info("💡 Importez votre fichier Excel `BUDGET.xlsx` pour lancer la confrontation Budget vs Dépenses.")
# ========================================================== ONGLET HISTORIQUE
with tab_history:
    st.subheader("📚 Liste complète des factures enregistrées")
    
    # 🟢 1. Remplacer col_refresh, _ par deux vraies colonnes pour les boutons
    col_refresh, col_clear, _ = st.columns([1.5, 1.5, 3])
    
    with col_refresh:
        refresh_clicked = st.button("🔄 Actualiser l'historique", use_container_width=True)
        if refresh_clicked:
            st.rerun()
            
    with col_clear:
        # Bouton rouge pour vider l'historique
        clear_clicked = st.button("🗑️ Vider l'historique", type="primary", use_container_width=True)
        
    # 🟢 2. Action d l-bouton jdid
    if clear_clicked:
        try:
            db.clear_all_invoices()
            st.success("💥 La base de données SQLite a été vidée avec succès !")
            st.rerun()
        except Exception as e:
            st.error(f"Erreur lors du nettoyage de la base : {e}")
        
    try:
        records = db.list_invoices()
        
        if records:
            history_df = pd.DataFrame(records)
            
            # Réorganiser l'affichage pour mettre l'ID et la date de sauvegarde au début
            all_cols = history_df.columns.tolist()
            desired_order = ['id', 'saved_at', 'plant', 'department', 'project', 'invoice_number', 'invoice_date', 'total', 'montant_ht', 'montant_ttc', 'contracts_count']
            display_cols = [c for c in desired_order if c in all_cols] + [c for c in all_cols if c not in desired_order]
            
            # Affichage du tableau propre
            st.dataframe(
                history_df[display_cols].sort_values(by="id", ascending=False),
                use_container_width=True,
                hide_index=True
            )
            
            st.caption(f"📊 Total de factures indexées distinctes f SQLite : {len(history_df)}")
            
        else:
            st.info("📂 Aucune facture n'a encore été enregistrée dans la base SQLite.")
            
    except Exception as e:
        st.error(f"Impossible de récupérer l'historique depuis la base de données : {e}")