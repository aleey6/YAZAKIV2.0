"""Composants réutilisables de l'interface utilisateur."""

from __future__ import annotations

import streamlit as st
from config import COLORS

def apply_custom_css():
    """Applique le CSS personnalisé à l'application."""
    st.markdown(
        f"""
        <style>
          :root {{
            --rouge: #{COLORS['rouge']};
            --rouge-fonce: #991b1b;
            --noir: #{COLORS['noir']};
            --blanc: #{COLORS['blanc']};
            --gris-clair: #{COLORS['gris']};
            --gris-bord: #{COLORS['gris_bord']};
          }}
          /* ... tout votre CSS existant ... */
        </style>
        """,
        unsafe_allow_html=True,
    )

def render_info_card(title: str, content: str):
    """Affiche une carte d'information stylisée."""
    st.markdown(
        f"""
        <div class='carte-info'>
          <b>{title}</b><br/>
          <span>{content}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

def init_session_state():
    """Initialise les variables de session Streamlit."""
    if "invoice" not in st.session_state:
        st.session_state.invoice = None
    if "uploaded_name" not in st.session_state:
        st.session_state.uploaded_name = None
    if "auto_saved" not in st.session_state:
        st.session_state.auto_saved = False
    if "cost_center_id" not in st.session_state:
        st.session_state.cost_center_id = None
    if "pdf_bytes" not in st.session_state:
        st.session_state.pdf_bytes = None
    if "pdf_file_key" not in st.session_state:
        st.session_state.pdf_file_key = None