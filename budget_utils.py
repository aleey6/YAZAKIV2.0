"""Utilitaires pour la gestion du budget et des dépenses."""

from __future__ import annotations

import io
import pandas as pd
from config import _normalize

def load_budget_df(file_bytes: bytes) -> pd.DataFrame:
    """Charge le fichier Excel en détectant automatiquement la ligne d'en-tête."""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
    
    for i, row in df_raw.iterrows():
        row_norm = [_normalize(str(v)) for v in row.values]
        if any(k in row_norm for k in ("usine", "plant", "budget", "departement")):
            df_raw.columns = [str(v).strip() for v in df_raw.iloc[i].values]
            return df_raw.iloc[i + 1:].reset_index(drop=True)
    
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=0)

def read_budget_from_excel(file_bytes: bytes, plant: str, dept: str, project: str) -> dict | None:
    """Lit le budget ET le cost center depuis un fichier Excel."""
    try:
        df = load_budget_df(file_bytes)
        df.columns = [str(c).strip() for c in df.columns]
        col_map = {}
        for col in df.columns:
            n = _normalize(col)
            if n in ("usine", "plant", "factory"): col_map["usine"] = col
            elif n in ("departement", "department", "dept"): col_map["dept"] = col
            elif n in ("projet", "project"): col_map["projet"] = col
            elif n in ("budget", "montant", "amount"): col_map["budget"] = col
            elif n in ("costcenterid", "costcenter", "centredecoût", "cc"): col_map["cc"] = col
        
        if not col_map.get("budget"):
            return None
        
        # Matching normalisé
        mask = pd.Series([True] * len(df))
        if col_map.get("usine") and plant:
            mask &= df[col_map["usine"]].astype(str).apply(_normalize) == _normalize(plant)
        if col_map.get("dept") and dept:
            mask &= df[col_map["dept"]].astype(str).apply(_normalize) == _normalize(dept)
        if col_map.get("projet") and project:
            mask &= df[col_map["projet"]].astype(str).apply(_normalize) == _normalize(project)
        matched = df[mask]
        
        # Fallback sans projet
        if matched.empty and (plant or dept):
            mask2 = pd.Series([True] * len(df))
            if col_map.get("usine") and plant:
                mask2 &= df[col_map["usine"]].astype(str).apply(_normalize) == _normalize(plant)
            if col_map.get("dept") and dept:
                mask2 &= df[col_map["dept"]].astype(str).apply(_normalize) == _normalize(dept)
            matched = df[mask2]
        
        if matched.empty:
            return None
        
        row_data = matched.iloc[0]
        raw_budget = str(row_data[col_map["budget"]]).replace(",", ".").replace(" ", "")
        budget_val = float(raw_budget)
        cc_val = None
        if col_map.get("cc"):
            cc_raw = row_data.get(col_map["cc"])
            cc_val = str(cc_raw).strip() if cc_raw is not None and str(cc_raw) != "nan" else None
        found_project = str(row_data.get(col_map.get("projet", ""), "")).strip() \
            if col_map.get("projet") else project
        return {"budget": budget_val, "cost_center_id": cc_val, "matched_project": found_project}
    except Exception:
        return None