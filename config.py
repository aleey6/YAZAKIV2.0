"""Configuration et constantes de l'application."""

from __future__ import annotations

import json
import re
from pathlib import Path

import streamlit as st

ROOT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = ROOT_DIR / "config.json"
DATA_DIR = ROOT_DIR / "data"
MASTER_FILE = DATA_DIR / "invoices.json"

# Couleurs YAZAKI
COLORS = {
    "rouge": "DC2626",
    "noir": "0A0A0A",
    "blanc": "FFFFFF",
    "gris": "F5F5F5",
    "gris_bord": "E5E5E5"
}

def load_config() -> dict:
    """Charge la configuration depuis config.json."""
    if not CONFIG_PATH.exists():
        return {"plants": {}}
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))

def safe_filename(value: str) -> str:
    """Nettoie une chaîne pour en faire un nom de fichier valide."""
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._-")
    return cleaned or "facture"

def fmt_money(value: float) -> str:
    """Formate un nombre en monnaie (MAD)."""
    return f"{value:,.2f} MAD".replace(",", " ")

def _normalize(s: str) -> str:
    """Normalise une chaîne : minuscules, sans accents, sans tirets/espaces."""
    import unicodedata
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[\s\-_/]+", "", s)