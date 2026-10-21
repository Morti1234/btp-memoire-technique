"""
==========================================================================
 Agent IA - Générateur de Mémoire Technique BTP
 Application Streamlit pour automatiser la réponse aux appels d'offres
==========================================================================
 Stack     : Streamlit · pdfplumber · OpenAI / Gemini · python-docx
 Hébergement : Streamlit Community Cloud (Snowflake)
 Mode      : Usage personnel – profil entreprise dans st.secrets
==========================================================================
"""

import io
import json
import math
import os
import re
import textwrap
from datetime import datetime
from pathlib import Path
from typing import Optional

import pdfplumber
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import markdown
import base64
import shutil
import tempfile
import zipfile

import video_tools as vt
import email_automation as ea

# ─────────────────────────────────────────────────────────────────────────
# 0. PERSISTENCE – Compatible local ET Streamlit Cloud
# ─────────────────────────────────────────────────────────────────────────
# En cloud, le filesystem est éphémère. On utilise st.secrets pour les
# données permanentes (clé API, profil entreprise) et st.session_state
# pour les modifications en session.
# En local, on garde le fallback config.json.
# ─────────────────────────────────────────────────────────────────────────

DEFAULT_ENTREPRISE = {
    "nom": "",
    "forme_juridique": "",
    "siret": "",
    "adresse": "",
    "telephone": "",
    "email": "",
    "site_web": "",
    "dirigeant": "",
    "date_creation": "",
    "effectif": "",
    "chiffre_affaires": "",
    "activites_principales": "",
    "zone_intervention": "",
    "certifications": "",
    "qualifications": "",
    "assurances": "",
    "references_chantiers": "",
    "moyens_materiels": "",
    "presentation_libre": "",
}


def _read_secrets_entreprise() -> dict:
    """
    Lit le profil entreprise depuis st.secrets (section [entreprise]).
    Retourne un dict avec les valeurs trouvées, complétées par les défauts.
    """
    ent = dict(DEFAULT_ENTREPRISE)
    try:
        secrets_ent = dict(st.secrets.get("entreprise", {}))
        for key in ent:
            if key in secrets_ent:
                ent[key] = str(secrets_ent[key])
    except Exception:
        pass
    return ent


def load_config() -> dict:
    """
    Charge la configuration en cascade :
    1. st.session_state (modifications en cours de session)
    2. st.secrets (configuration cloud persistante)
    3. config.json local (fallback développement)
    4. Valeurs par défaut
    """
    # Si déjà chargé en session, on le retourne
    if "config" in st.session_state:
        return st.session_state["config"]

    config = {
        "api_key": "",
        "llm_provider": "OpenAI (GPT-4o)",
        "dark_mode": False,
        "entreprise": dict(DEFAULT_ENTREPRISE),
    }

    # ── Source 1 : st.secrets (Streamlit Cloud) ─────────────────────────
    try:
        config["api_key"] = st.secrets.get("api_key", "")
        config["llm_provider"] = st.secrets.get("llm_provider", "OpenAI (GPT-4o)")
        config["entreprise"] = _read_secrets_entreprise()
    except Exception:
        pass

    # ── Source 2 : config.json local (fallback dev) ─────────────────────
    config_file = Path(__file__).parent / "config.json"
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                saved = json.load(f)
            # Les secrets ont priorité, on ne surcharge que les champs vides
            if not config["api_key"]:
                config["api_key"] = saved.get("api_key", "")
            if config["llm_provider"] == "OpenAI (GPT-4o)" and "llm_provider" in saved:
                config["llm_provider"] = saved["llm_provider"]
            if "dark_mode" in saved and isinstance(saved["dark_mode"], bool):
                config["dark_mode"] = saved["dark_mode"]
            for key in config["entreprise"]:
                if not config["entreprise"][key]:
                    config["entreprise"][key] = saved.get("entreprise", {}).get(key, "")
        except (json.JSONDecodeError, KeyError):
            pass

    # Stocker en session pour les modifications futures
    st.session_state["config"] = config
    return config


def save_config(config: dict):
    """
    Sauvegarde la config :
    - Toujours dans st.session_state (immédiat)
    - Dans config.json si le filesystem le permet (local)
    """
    st.session_state["config"] = config

    # Tenter la sauvegarde locale (échoue silencieusement en cloud)
    config_file = Path(__file__).parent / "config.json"
    try:
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except OSError:
        pass  # Filesystem en lecture seule en cloud, c'est normal


def build_profil_text(entreprise: dict) -> Optional[str]:
    """
    Construit le texte du profil entreprise à injecter dans le prompt LLM.
    Retourne None si le profil est quasiment vide.
    """
    fields_labels = {
        "nom": "Nom de l'entreprise",
        "forme_juridique": "Forme juridique",
        "siret": "SIRET",
        "adresse": "Adresse",
        "telephone": "Téléphone",
        "email": "Email",
        "site_web": "Site web",
        "dirigeant": "Dirigeant / Gérant",
        "date_creation": "Date de création",
        "effectif": "Effectif",
        "chiffre_affaires": "Chiffre d'affaires annuel",
        "activites_principales": "Activités principales",
        "zone_intervention": "Zone d'intervention géographique",
        "certifications": "Certifications (RGE, ISO, etc.)",
        "qualifications": "Qualifications (Qualibat, Qualifelec, etc.)",
        "assurances": "Assurances (décennale, RC Pro, etc.)",
        "references_chantiers": "Références chantiers similaires",
        "moyens_materiels": "Moyens matériels",
        "presentation_libre": "Présentation libre / Historique",
    }

    lines = []
    filled = 0
    for key, label in fields_labels.items():
        value = entreprise.get(key, "").strip()
        if value:
            lines.append(f"**{label}** : {value}")
            filled += 1

    if filled < 2:
        return None

    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────
# 1. CONFIGURATION DE LA PAGE STREAMLIT & THÈME (MODE SOMBRE ON/OFF)
# ─────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MarketBuilding Pro - Mémoire Technique BTP",
    page_icon="🏗️",
    layout="wide",
)

config = load_config()
if "dark_mode" not in st.session_state:
    st.session_state["dark_mode"] = bool(config.get("dark_mode", False))


def get_marketbuilding_logo_svg(dark_mode: bool = False) -> str:
    """Retourne le logo SVG MarketBuilding adapté au fond (Clair par défaut pour la navigation)."""
    win_fill = "#0e1117" if dark_mode else "#ffffff"
    title_fill = "#ffffff" if dark_mode else "#0F172A"
    sub_fill = "#94A3B8" if dark_mode else "#475569"
    tag_fill = "#64748B" if dark_mode else "#64748B"
    return f'''
    <div style="text-align:center; padding: 6px 0 10px 0;">
    <svg width="215" height="90" viewBox="0 0 210 90" xmlns="http://www.w3.org/2000/svg">
      <!-- Bâtiment 1 (petit, gauche) -->
      <rect x="12" y="30" width="16" height="36" rx="1.5" fill="#3B82F6"/>
      <rect x="15" y="34" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="21" y="34" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="15" y="41" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="21" y="41" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="15" y="48" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="21" y="48" width="4" height="4" rx="0.5" fill="{win_fill}"/>

      <!-- Bâtiment 2 (grand, centre) -->
      <rect x="31" y="12" width="22" height="54" rx="1.5" fill="#2563EB"/>
      <rect x="35" y="16" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="42" y="16" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="35" y="23" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="42" y="23" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="35" y="30" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="42" y="30" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="35" y="37" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="42" y="37" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="35" y="44" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="42" y="44" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="35" y="51" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="42" y="51" width="4" height="4" rx="0.5" fill="{win_fill}"/>

      <!-- Bâtiment 3 (moyen, droite) -->
      <rect x="56" y="24" width="18" height="42" rx="1.5" fill="#60A5FA"/>
      <rect x="59" y="28" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="66" y="28" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="59" y="35" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="66" y="35" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="59" y="42" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="66" y="42" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="59" y="49" width="4" height="4" rx="0.5" fill="{win_fill}"/>
      <rect x="66" y="49" width="4" height="4" rx="0.5" fill="{win_fill}"/>

      <!-- Grue de chantier -->
      <line x1="78" y1="66" x2="78" y2="5" stroke="#F59E0B" stroke-width="2.5"/>
      <line x1="78" y1="5" x2="56" y2="5" stroke="#F59E0B" stroke-width="2"/>
      <line x1="78" y1="5" x2="88" y2="5" stroke="#F59E0B" stroke-width="2"/>
      <line x1="60" y1="5" x2="60" y2="16" stroke="#F59E0B" stroke-width="1" stroke-dasharray="2,1"/>
      <rect x="84" y="3" width="6" height="5" rx="1" fill="#D97706"/>
      <rect x="73" y="62" width="10" height="4" rx="1" fill="#D97706"/>
      <line x1="5" y1="66" x2="90" y2="66" stroke="#475569" stroke-width="1.5"/>

      <!-- Typographie -->
      <text x="100" y="34" font-family="'Segoe UI', 'Helvetica Neue', Arial, sans-serif" font-size="19" font-weight="700" fill="{title_fill}" letter-spacing="0.5">Market</text>
      <text x="100" y="54" font-family="'Segoe UI', 'Helvetica Neue', Arial, sans-serif" font-size="19" font-weight="400" fill="{sub_fill}" letter-spacing="0.5">Building</text>
      <line x1="100" y1="61" x2="200" y2="61" stroke="#F59E0B" stroke-width="1.5" opacity="0.75"/>
      <text x="100" y="74" font-family="'Segoe UI', 'Helvetica Neue', Arial, sans-serif" font-size="7" font-weight="600" fill="{tag_fill}" letter-spacing="1.5">MÉMOIRES TECHNIQUES BTP</text>
    </svg>
    </div>
    '''


def inject_custom_css(dark_mode: bool = False):
    """
    Injecte la feuille de style complète et unifiée :
    Quand Mode Sombre est OFF -> TOUT le site (barre latérale, navigation, onglets, contenus ouverts,
    formulaires, zones d'upload, menus déroulants, expanders) est en Thème Clair.
    Quand Mode Sombre est ON  -> TOUT le site (barre latérale, navigation, onglets, contenus ouverts,
    formulaires, zones d'upload, menus déroulants, expanders) passe en Thème Sombre.
    """
    if dark_mode:
        vars_css = """
            --bg-app: #0B0F17;
            --bg-sidebar: #111827;
            --bg-card: #151C28;
            --bg-card-alt: #1E293B;
            --bg-hero: linear-gradient(135deg, #111827 0%, #1E293B 100%);
            --border-color: #263143;
            --border-hover: #3B82F6;
            --text-primary: #F8FAFC;
            --text-secondary: #CBD5E1;
            --text-muted: #94A3B8;
            --nav-item-bg: #151C28;
            --nav-item-hover: #1E293B;
            --nav-active-bg: rgba(37, 99, 235, 0.22);
            --nav-active-text: #60A5FA;
            --tag-bg: rgba(59, 130, 246, 0.16);
            --tag-text: #93C5FD;
            --tag-border: rgba(59, 130, 246, 0.35);
            --shadow-card: 0 8px 24px rgba(0, 0, 0, 0.35);
            --input-bg: #0F1520;
            --uploader-bg: #111722;
        """
    else:
        vars_css = """
            --bg-app: #F8FAFC;
            --bg-sidebar: #FFFFFF;
            --bg-card: #FFFFFF;
            --bg-card-alt: #F1F5F9;
            --bg-hero: linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 100%);
            --border-color: #E2E8F0;
            --border-hover: #2563EB;
            --text-primary: #0F172A;
            --text-secondary: #334155;
            --text-muted: #64748B;
            --nav-item-bg: #F8FAFC;
            --nav-item-hover: #EFF6FF;
            --nav-active-bg: #EFF6FF;
            --nav-active-text: #1D4ED8;
            --tag-bg: #EFF6FF;
            --tag-text: #1D4ED8;
            --tag-border: #BFDBFE;
            --shadow-card: 0 6px 20px rgba(15, 23, 42, 0.05);
            --input-bg: #FFFFFF;
            --uploader-bg: #F8FAFC;
        """

    st.markdown(
        f"""
        <style>
            :root {{
                {vars_css}
            }}

            /* Masquer les éléments parasites de Streamlit */
            #MainMenu {{visibility: hidden;}}
            footer {{visibility: hidden;}}
            .stAppDeployButton {{display: none;}}
            [data-testid="stHeader"] {{background: transparent !important;}}

            /* Fond global et typographie */
            .stApp,
            [data-testid="stAppViewContainer"],
            [data-testid="stMain"] {{
                background-color: var(--bg-app) !important;
                color: var(--text-primary) !important;
            }}

            .block-container {{
                padding-top: 1.2rem !important;
                padding-bottom: 3rem !important;
                max-width: 1180px !important;
            }}

            /* ── BARRE LATÉRALE & NAVIGATION PRINCIPALE (CLAIR ET SOMBRE UNIFIÉS) ── */
            section[data-testid="stSidebar"],
            section[data-testid="stSidebar"] > div {{
                background-color: var(--bg-sidebar) !important;
                border-right: 1px solid var(--border-color) !important;
            }}
            section[data-testid="stSidebar"] * {{
                color: var(--text-primary) !important;
            }}
            section[data-testid="stSidebar"] .stCaption,
            section[data-testid="stSidebar"] small {{
                color: var(--text-muted) !important;
            }}
            section[data-testid="stSidebar"] hr {{
                border-color: var(--border-color) !important;
            }}

            /* Boutons radio de la Navigation (Barre latérale & Barre horizontale) */
            div[role="radiogroup"] label {{
                background-color: var(--nav-item-bg) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 10px !important;
                padding: 10px 14px !important;
                margin-bottom: 6px !important;
                transition: all 0.15s ease !important;
                font-weight: 600 !important;
                color: var(--text-primary) !important;
            }}
            div[role="radiogroup"] label p,
            div[role="radiogroup"] label span {{
                color: var(--text-primary) !important;
                font-weight: 600 !important;
            }}
            div[role="radiogroup"] label:hover {{
                background-color: var(--nav-item-hover) !important;
                border-color: var(--border-hover) !important;
            }}
            div[role="radiogroup"] label[data-checked="true"],
            div[role="radiogroup"] label:has(input:checked) {{
                background-color: var(--nav-active-bg) !important;
                border: 1.5px solid #2563EB !important;
                border-left: 4px solid #2563EB !important;
                color: var(--nav-active-text) !important;
                box-shadow: 0 2px 8px rgba(37, 99, 235, 0.15) !important;
            }}
            div[role="radiogroup"] label:has(input:checked) p,
            div[role="radiogroup"] label:has(input:checked) span {{
                color: var(--nav-active-text) !important;
                font-weight: 700 !important;
            }}

            /* Titres et textes de toutes les pages et onglets */
            h1, h2, h3, h4, h5, h6,
            [data-testid="stMarkdownContainer"] h1,
            [data-testid="stMarkdownContainer"] h2,
            [data-testid="stMarkdownContainer"] h3,
            [data-testid="stMarkdownContainer"] h4 {{
                color: var(--text-primary) !important;
                letter-spacing: -0.01em;
            }}
            [data-testid="stMarkdownContainer"] p,
            [data-testid="stMarkdownContainer"] li,
            label, .stCaption, [data-testid="stCaptionContainer"] * {{
                color: var(--text-secondary) !important;
            }}

            /* ── ONGLETS (TABS) ET LEUR CONTENU QUAND ON LES OUVRE ── */
            div[data-testid="stTabs"] {{
                background: transparent !important;
            }}
            div[data-baseweb="tab-list"] {{
                background-color: var(--bg-card) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 12px !important;
                padding: 6px 10px !important;
                margin-bottom: 16px !important;
                gap: 8px !important;
                box-shadow: var(--shadow-card) !important;
            }}
            button[data-baseweb="tab"] {{
                background-color: transparent !important;
                color: var(--text-muted) !important;
                font-weight: 600 !important;
                font-size: 14px !important;
                border-radius: 8px !important;
                padding: 8px 14px !important;
            }}
            button[data-baseweb="tab"] * {{
                color: var(--text-muted) !important;
            }}
            button[data-baseweb="tab"]:hover {{
                background-color: var(--nav-item-hover) !important;
            }}
            button[data-baseweb="tab"][aria-selected="true"] {{
                background-color: var(--nav-active-bg) !important;
                color: var(--nav-active-text) !important;
                font-weight: 700 !important;
            }}
            button[data-baseweb="tab"][aria-selected="true"] * {{
                color: var(--nav-active-text) !important;
                font-weight: 700 !important;
            }}
            div[data-testid="stTabs"] button:not([data-baseweb="tab"]) {{
                background-color: var(--bg-card-alt) !important;
                color: var(--text-primary) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 8px !important;
            }}
            div[data-testid="stTabs"] button:not([data-baseweb="tab"]) svg {{
                fill: var(--text-primary) !important;
                color: var(--text-primary) !important;
            }}
            div[data-baseweb="tab-panel"] {{
                background-color: transparent !important;
                color: var(--text-primary) !important;
                padding-top: 6px !important;
            }}
            code {{
                background-color: var(--bg-card-alt) !important;
                color: var(--nav-active-text) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 5px !important;
                padding: 1px 6px !important;
                font-size: 0.9em !important;
            }}

            /* ── EXPANDERS (VOLETS DÉROULANTS QUAND ON LES OUVRE) & FORMULAIRES ── */
            div[data-testid="stForm"],
            div[data-testid="stExpander"],
            details {{
                background-color: var(--bg-card) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 14px !important;
                box-shadow: var(--shadow-card) !important;
                color: var(--text-primary) !important;
            }}
            div[data-testid="stExpander"] details summary,
            div[data-testid="stExpander"] details summary * {{
                background-color: var(--bg-card) !important;
                color: var(--text-primary) !important;
                font-weight: 600 !important;
                border-radius: 14px !important;
            }}
            div[data-testid="stExpanderDetails"] {{
                background-color: var(--bg-card) !important;
                color: var(--text-primary) !important;
                border-top: 1px solid var(--border-color) !important;
                border-radius: 0 0 14px 14px !important;
                padding: 14px !important;
            }}

            /* ── CHAMPS DE SAISIE, TEXTAREAS, SELECTBOX & MENUS DÉROULANTS ── */
            div[data-baseweb="input"] > div,
            div[data-baseweb="base-input"],
            div[data-baseweb="textarea"] > div,
            div[data-baseweb="select"] > div {{
                background-color: var(--input-bg) !important;
                border: 1px solid var(--border-color) !important;
                color: var(--text-primary) !important;
                border-radius: 10px !important;
            }}
            div[data-baseweb="input"] input,
            div[data-baseweb="textarea"] textarea,
            div[data-baseweb="select"] * {{
                color: var(--text-primary) !important;
                background-color: transparent !important;
            }}
            div[data-baseweb="popover"] > div,
            ul[data-baseweb="menu"],
            ul[data-baseweb="menu"] li {{
                background-color: var(--bg-card) !important;
                color: var(--text-primary) !important;
                border-color: var(--border-color) !important;
            }}
            ul[data-baseweb="menu"] li:hover {{
                background-color: var(--nav-item-hover) !important;
            }}

            /* ── ZONES DE DRAG & DROP (FILE UPLOADER) ET FICHIERS CHARGÉS ── */
            [data-testid="stFileUploader"] section,
            [data-testid="stFileUploaderDropzone"] {{
                background-color: var(--uploader-bg) !important;
                border: 1.5px dashed var(--border-color) !important;
                border-radius: 12px !important;
                padding: 14px !important;
                transition: all 0.2s ease;
            }}
            [data-testid="stFileUploader"] section:hover,
            [data-testid="stFileUploaderDropzone"]:hover {{
                border-color: var(--border-hover) !important;
            }}
            [data-testid="stFileUploader"] section *,
            [data-testid="stFileUploaderDropzone"] * {{
                color: var(--text-secondary) !important;
            }}
            [data-testid="stFileUploader"] section button,
            [data-testid="stFileUploaderDropzone"] button {{
                background-color: var(--bg-card) !important;
                border: 1px solid var(--border-color) !important;
                color: var(--text-primary) !important;
                border-radius: 8px !important;
                font-weight: 600 !important;
            }}
            [data-testid="stFileUploaderFile"] {{
                background-color: var(--bg-card-alt) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 8px !important;
                color: var(--text-primary) !important;
            }}
            [data-testid="stFileUploaderFile"] * {{
                color: var(--text-primary) !important;
            }}

            /* ── BOUTONS ── */
            button[kind="primary"] {{
                background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
                color: #FFFFFF !important;
                border: none !important;
                border-radius: 11px !important;
                padding: 0.65rem 1.2rem !important;
                font-weight: 700 !important;
                letter-spacing: 0.2px !important;
                box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
                transition: transform 0.15s ease, box-shadow 0.15s ease !important;
            }}
            button[kind="primary"] * {{
                color: #FFFFFF !important;
            }}
            button[kind="primary"]:hover {{
                transform: translateY(-1px);
                box-shadow: 0 6px 20px rgba(37, 99, 235, 0.5) !important;
            }}
            button[kind="secondary"] {{
                background-color: var(--bg-card) !important;
                color: var(--text-primary) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 10px !important;
                font-weight: 600 !important;
            }}
            button[kind="secondary"] * {{
                color: var(--text-primary) !important;
            }}

            /* ── COMPOSANTS PERSONNALISÉS MARKETBUILDING ── */
            .mb-hero {{
                background: var(--bg-hero);
                border: 1px solid var(--border-color);
                border-left: 5px solid #2563EB;
                border-radius: 16px;
                padding: 22px 26px;
                margin-bottom: 20px;
                box-shadow: var(--shadow-card);
            }}
            .mb-hero-badge {{
                display: inline-block;
                background: var(--tag-bg);
                color: var(--tag-text);
                border: 1px solid var(--tag-border);
                font-size: 11.5px;
                font-weight: 700;
                letter-spacing: 0.8px;
                text-transform: uppercase;
                padding: 4px 10px;
                border-radius: 999px;
                margin-bottom: 8px;
            }}
            .mb-hero-title {{
                color: var(--text-primary);
                font-size: 24px;
                font-weight: 800;
                margin: 0 0 6px 0;
                line-height: 1.25;
            }}
            .mb-hero-sub {{
                color: var(--text-secondary);
                font-size: 14px;
                margin: 0;
                line-height: 1.55;
            }}

            /* Grille d'étapes (Vidéo & Méthodologie) */
            .mb-steps-grid {{
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 14px;
                margin-top: 12px;
            }}
            @media (max-width: 768px) {{
                .mb-steps-grid {{
                    grid-template-columns: 1fr;
                }}
            }}
            .mb-step-card {{
                background: var(--bg-card);
                border: 1px solid var(--border-color);
                border-radius: 12px;
                padding: 14px 16px;
                box-shadow: var(--shadow-card);
            }}
            .mb-step-title {{
                font-weight: 700;
                font-size: 13px;
                margin-bottom: 5px;
                letter-spacing: 0.2px;
            }}
            .mb-step-desc {{
                color: var(--text-secondary);
                font-size: 12.5px;
                line-height: 1.45;
            }}

            /* Cartes des 5 zones d'upload et articles juridiques */
            .mb-zone-header {{
                background: var(--bg-card);
                border: 1px solid var(--border-color);
                border-radius: 14px;
                padding: 16px 18px 12px 18px;
                margin-bottom: 8px;
                box-shadow: var(--shadow-card);
            }}
            .mb-zone-top {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 8px;
                margin-bottom: 6px;
            }}
            .mb-zone-num {{
                display: inline-flex;
                align-items: center;
                gap: 6px;
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 0.7px;
                padding: 3px 9px;
                border-radius: 6px;
            }}
            .mb-zone-title {{
                color: var(--text-primary);
                font-size: 16px;
                font-weight: 700;
                margin: 2px 0 6px 0;
            }}
            .mb-zone-desc {{
                color: var(--text-secondary);
                font-size: 13px;
                line-height: 1.6;
                margin-bottom: 10px;
            }}
            .mb-tag-row {{
                display: flex;
                flex-wrap: wrap;
                gap: 6px;
                margin-top: 6px;
            }}
            .mb-tag {{
                background: var(--bg-card-alt);
                border: 1px solid var(--border-color);
                color: var(--text-secondary);
                font-size: 11.5px;
                font-weight: 600;
                padding: 3px 8px;
                border-radius: 6px;
            }}

            /* Barre récapitulative des 5 catégories */
            .mb-summary-bar {{
                display: grid;
                grid-template-columns: repeat(5, 1fr);
                gap: 10px;
                margin: 16px 0 18px 0;
            }}
            @media (max-width: 900px) {{
                .mb-summary-bar {{
                    grid-template-columns: repeat(2, 1fr);
                }}
            }}
            .mb-summary-pill {{
                background: var(--bg-card);
                border: 1px solid var(--border-color);
                border-radius: 12px;
                padding: 10px 12px;
                text-align: center;
                box-shadow: var(--shadow-card);
            }}
            .mb-summary-pill.active {{
                border-color: #10B981;
                background: rgba(16, 185, 129, 0.10);
            }}
            .mb-summary-label {{
                font-size: 11.5px;
                font-weight: 700;
                color: var(--text-muted);
                text-transform: uppercase;
                letter-spacing: 0.4px;
            }}
            .mb-summary-val {{
                font-size: 14px;
                font-weight: 800;
                color: var(--text-primary);
                margin-top: 3px;
            }}

            /* Carte Login */
            .password-card {{
                max-width: 460px;
                margin: 30px auto 20px auto;
                background: var(--bg-card);
                border: 1px solid var(--border-color);
                border-top: 4px solid #2563EB;
                border-radius: 16px;
                padding: 24px 28px;
                text-align: center;
                box-shadow: var(--shadow-card);
            }}

            /* ── BARRE EXÉCUTIVE SUPÉRIEURE & CARTES KPI (DESIGN PRO B2B) ── */
            .mb-exec-topbar {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                flex-wrap: wrap;
                gap: 10px;
                background: var(--bg-card);
                border: 1px solid var(--border-color);
                border-radius: 12px;
                padding: 10px 16px;
                margin-bottom: 16px;
                box-shadow: var(--shadow-card);
            }}
            .mb-exec-brand {{
                display: flex;
                align-items: center;
                gap: 10px;
                font-size: 13px;
                font-weight: 700;
                color: var(--text-primary);
            }}
            .mb-exec-pills {{
                display: flex;
                align-items: center;
                flex-wrap: wrap;
                gap: 8px;
            }}
            .mb-exec-pill {{
                display: inline-flex;
                align-items: center;
                gap: 6px;
                background: var(--bg-card-alt);
                border: 1px solid var(--border-color);
                border-radius: 999px;
                padding: 4px 11px;
                font-size: 11.5px;
                font-weight: 600;
                color: var(--text-secondary);
            }}
            .mb-kpi-grid {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 12px;
                margin-bottom: 18px;
            }}
            @media (max-width: 900px) {{
                .mb-kpi-grid {{
                    grid-template-columns: repeat(2, 1fr);
                }}
            }}
            .mb-kpi-card {{
                background: var(--bg-card);
                border: 1px solid var(--border-color);
                border-radius: 14px;
                padding: 14px 16px;
                box-shadow: var(--shadow-card);
            }}
            .mb-kpi-label {{
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 0.6px;
                color: var(--text-muted);
                margin-bottom: 4px;
            }}
            .mb-kpi-val {{
                font-size: 22px;
                font-weight: 800;
                color: var(--text-primary);
                line-height: 1.2;
            }}
            .mb-kpi-sub {{
                font-size: 11.5px;
                color: var(--text-secondary);
                margin-top: 4px;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_custom_css(st.session_state["dark_mode"])

ADMIN_SECRET_CODE = "morti333"


def _on_theme_toggle_change(toggle_key: str):
    new_val = bool(st.session_state.get(toggle_key, False))
    st.session_state["dark_mode"] = new_val
    config["dark_mode"] = new_val
    save_config(config)


def render_theme_toggle(location_key: str):
    """Affiche un bouton ON/OFF clair pour basculer tout le site (navigation + onglets + pages) entre Clair et Sombre."""
    current_dark = bool(st.session_state.get("dark_mode", False))
    toggle_key = f"theme_toggle_{location_key}"
    st.session_state[toggle_key] = current_dark
    label_mode = "🌙 Mode Sombre : ON" if current_dark else "☀️ Mode Clair (Sombre OFF)"
    st.toggle(
        label_mode,
        key=toggle_key,
        on_change=_on_theme_toggle_change,
        args=(toggle_key,),
        help="Activez (ON) pour passer toute la navigation et tous les onglets en Thème Sombre, ou désactivez (OFF) pour tout passer en Thème Clair.",
    )


# ─────────────────────────────────────────────────────────────────────────
# 0.5 SÉCURITÉ - ACCÈS PRIVÉ PAR MOT DE PASSE
# ─────────────────────────────────────────────────────────────────────────
def check_password():
    """Returns `True` if the user had the correct password."""
    try:
        if "app_password" not in st.secrets:
            return True
    except Exception:
        return True

    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    # Interrupteur Mode Sombre ON/OFF sur l'écran de connexion
    top_col1, top_col2 = st.columns([3.6, 1.4])
    with top_col2:
        render_theme_toggle("login")

    st.markdown(get_marketbuilding_logo_svg(st.session_state["dark_mode"]), unsafe_allow_html=True)
    st.markdown(
        """
        <div class="password-card">
            <div class="mb-hero-badge">PORTAIL CLIENT SÉCURISÉ</div>
            <div style="font-size:20px; font-weight:800; color:var(--text-primary); margin-bottom:6px;">
                MarketBuilding Pro · Mémoires Techniques BTP
            </div>
            <div style="font-size:13px; color:var(--text-secondary); margin-bottom:4px;">
                Croisement documentaire RC × CCTP × DPGF × PGC × Entreprise
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _, center_col, _ = st.columns([1.2, 2, 1.2])
    with center_col:
        st.text_input(
            "🔑 Code d'accès client :",
            placeholder="Entrez votre mot de passe...",
            type="password",
            on_change=password_entered,
            key="password",
        )
        if "password_correct" in st.session_state and not st.session_state["password_correct"]:
            st.error("😕 Mot de passe incorrect. Veuillez réessayer.")
    return False


if not check_password():
    _feat = vt.featured_video()
    if _feat:
        st.markdown(
            "<div style='text-align:center; margin-top:20px; margin-bottom:10px;'>"
            "<span style='color:var(--text-muted); font-size:13px; font-weight:700; letter-spacing:1px; text-transform:uppercase;'>"
            "🎬 Démonstration : Du Dossier de Consultation au Téléchargement du Mémoire Technique"
            "</span></div>",
            unsafe_allow_html=True,
        )
        st.video(str(_feat), autoplay=True, loop=True, muted=True)
        st.markdown(
            """
            <div class="mb-steps-grid">
                <div class="mb-step-card" style="border-left:4px solid #3B82F6;">
                    <div class="mb-step-title" style="color:#3B82F6;">ÉTAPE 1 · DÉPÔT MULTI-DOCUMENTS</div>
                    <div class="mb-step-desc">Déposez vos pièces (<b>RC, CCTP, DPGF, Annexes, Entreprise</b> en PDF, Word ou Excel) et lancez la génération.</div>
                </div>
                <div class="mb-step-card" style="border-left:4px solid #F59E0B;">
                    <div class="mb-step-title" style="color:#F59E0B;">ÉTAPE 2 · TÉLÉCHARGEMENT DIRECT</div>
                    <div class="mb-step-desc">Une fois rédigé, cliquez sur <b>« 📥 Télécharger en PDF »</b> ou <b>« 📝 Télécharger en Word (.docx) »</b>.</div>
                </div>
                <div class="mb-step-card" style="border-left:4px solid #10B981;">
                    <div class="mb-step-title" style="color:#10B981;">ÉTAPE 3 · APERÇU PDF & ARCHIVES</div>
                    <div class="mb-step-desc">Prévisualisez les pages du PDF officiel directement à l'écran ou retrouvez tous vos mémoires dans l'onglet <b>Archives</b>.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.stop()


# ─────────────────────────────────────────────────────────────────────────
# 2. BARRE LATÉRALE – Navigation Principale (suit le thème Clair / Sombre)
# ─────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(get_marketbuilding_logo_svg(st.session_state["dark_mode"]), unsafe_allow_html=True)

    # Interrupteur Mode Clair / Sombre ON/OFF unique dans la barre latérale
    render_theme_toggle("sidebar")
    st.caption(
        f"Thème actif sur tout le site : **{'🌙 Mode Sombre (ON)' if st.session_state['dark_mode'] else '☀️ Mode Clair (OFF)'}**"
    )

    st.divider()
    st.markdown(
        "<div style='font-size:11.5px; font-weight:700; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.8px; margin-bottom:6px;'>Navigation Principale</div>",
        unsafe_allow_html=True,
    )

    # Navigation accessible à tous (clients et administrateur) en Mode Clair comme en Mode Sombre
    CLIENT_NAV_LABELS = {
        "Génération": "📄  Générer un Mémoire",
        "Automatisation": "🤖  Automatisation (Mails)",
        "Profil": "🏢  Profil Entreprise",
        "Archives": "📂  Archives & Livrables",
        "CGU": "📜  Conditions Générales (CGU)",
        "Confidentialité": "🔒  Politique de Confidentialité",
    }

    # Paramètres Système (Admin) visible uniquement après saisie du code secret "morti333"
    is_owner = bool(st.session_state.get("owner_unlocked", False))
    if is_owner:
        CLIENT_NAV_LABELS["Paramètres"] = "⚙️  Paramètres Système (Admin)"
        CLIENT_NAV_LABELS["Vidéos"] = "🎬  Studio Vidéos (Admin)"

    nav_options = list(CLIENT_NAV_LABELS.keys())
    if "_jump_to_nav" in st.session_state:
        target_nav = st.session_state.pop("_jump_to_nav")
        if target_nav in nav_options:
            st.session_state["main_nav_radio"] = target_nav
    if st.session_state.get("main_nav_radio") not in nav_options:
        st.session_state["main_nav_radio"] = "Génération"

    page = st.radio(
        "Navigation",
        options=nav_options,
        format_func=lambda k: CLIENT_NAV_LABELS[k],
        key="main_nav_radio",
        label_visibility="collapsed",
    )

    st.divider()

    # Carte d'état client adaptée au thème Clair / Sombre (sans aucune mention de Moteur IA)
    entreprise_ok = bool(config["entreprise"].get("nom", "").strip())
    os.makedirs("archives", exist_ok=True)
    nb_archives = len([f for f in os.listdir("archives") if f.endswith((".pdf", ".docx"))])
    _auto_st = ea.load_automation_state()
    _watcher_on = bool(_auto_st.get("watcher_enabled", False))
    _auto_mode_lbl = "100% Auto" if _auto_st.get("automation_mode", "auto") == "auto" else "Validation Manuelle"

    st.markdown(
        f"""
        <div style="background:var(--bg-card-alt); border:1px solid var(--border-color); border-radius:12px; padding:12px 14px;">
            <div style="font-size:11.5px; font-weight:700; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.6px; margin-bottom:8px;">
                Dossier Entreprise & Mails
            </div>
            <div style="font-size:12.5px; color:var(--text-secondary); margin-bottom:6px;">
                {'🟢' if entreprise_ok else '🔵'} <b>Société :</b> {config['entreprise'].get('nom') or 'Via pièces déposées'}
            </div>
            <div style="font-size:12.5px; color:var(--text-secondary); margin-bottom:6px;">
                📂 <b>Livrables archivés :</b> {nb_archives} fichier(s)
            </div>
            <div style="font-size:12.5px; color:var(--text-secondary);">
                {'🟢' if _watcher_on else '⏸️'} <b>Écoute Mails :</b> {'Active' if _watcher_on else 'En pause'} ({_auto_mode_lbl})
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.divider()
    st.caption("MarketBuilding Pro v3.5 · Édition Client BTP")

    # Zone protégée par le code "morti333" pour afficher Paramètres Système (Admin)
    with st.expander("🔑 Accès Administrateur", expanded=False):
        if not is_owner:
            owner_code = st.text_input(
                "Code secret administrateur",
                type="password",
                key="owner_pin_input",
                placeholder="Entrez votre code...",
            )
            if owner_code:
                if owner_code.strip() == ADMIN_SECRET_CODE:
                    st.session_state["owner_unlocked"] = True
                    st.session_state["_jump_to_nav"] = "Paramètres"
                    st.rerun()
                else:
                    st.error("Code administrateur incorrect.")
        else:
            st.success("🔓 Accès Admin déverrouillé")
            if st.button("🔒 Verrouiller l'accès Admin", use_container_width=True):
                st.session_state["owner_unlocked"] = False
                if st.session_state.get("main_nav_radio") in ("Paramètres", "Vidéos"):
                    st.session_state["_jump_to_nav"] = "Génération"
                st.rerun()


# ─────────────────────────────────────────────────────────────────────────
# 3. FONCTIONS D'EXTRACTION MULTI-FORMATS (PDF, WORD, EXCEL, CSV)
# ─────────────────────────────────────────────────────────────────────────

SUPPORTED_DOC_TYPES = ["pdf", "docx", "xlsx", "xls", "csv", "txt"]


def extract_text_from_pdf(uploaded_file) -> str:
    """
    Extrait le texte brut ET les tableaux d'un PDF uploadé via Streamlit.
    Utilise pdfplumber pour un parsing robuste (gère les tableaux, colonnes…).
    """
    text_pages: list[str] = []
    raw_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    try:
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            for i, page in enumerate(pdf.pages):
                parts: list[str] = []
                page_text = page.extract_text()
                if page_text:
                    parts.append(page_text)
                # Extraction complémentaire des tableaux (utile pour DPGF / barèmes RC en PDF)
                try:
                    tables = page.extract_tables() or []
                    for t_idx, table in enumerate(tables, start=1):
                        cleaned_rows = []
                        for row in table:
                            if not row:
                                continue
                            cells = [re.sub(r"\s+", " ", (c or "").strip()) for c in row]
                            if any(cells):
                                cleaned_rows.append(" | ".join(cells))
                        if len(cleaned_rows) >= 2:
                            parts.append(f"[Tableau page {i + 1}.{t_idx}]\n" + "\n".join(cleaned_rows))
                except Exception:
                    pass
                if parts:
                    text_pages.append(f"--- Page {i + 1} ---\n" + "\n\n".join(parts))
    except Exception as exc:
        raise ValueError(
            f"Impossible de lire le fichier PDF ({getattr(uploaded_file, 'name', 'document')}). "
            f"Détail : {exc}"
        ) from exc

    if not text_pages:
        raise ValueError(
            f"Le PDF « {getattr(uploaded_file, 'name', 'document')} » ne contient aucun texte extractible "
            "(document scanné sans OCR)."
        )

    return "\n\n".join(text_pages)


def extract_text_from_docx(uploaded_file) -> str:
    """Extrait les paragraphes et les tableaux d'un fichier Word (.docx)."""
    raw_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    try:
        doc = Document(io.BytesIO(raw_bytes))
    except Exception as exc:
        raise ValueError(f"Impossible de lire le fichier Word ({getattr(uploaded_file, 'name', '')}) : {exc}") from exc

    blocks: list[str] = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            blocks.append(txt)

    for t_idx, table in enumerate(doc.tables, start=1):
        rows_txt = []
        for row in table.rows:
            cells = [re.sub(r"\s+", " ", c.text.strip()) for c in row.cells]
            if any(cells):
                rows_txt.append(" | ".join(cells))
        if rows_txt:
            blocks.append(f"[Tableau Word {t_idx}]\n" + "\n".join(rows_txt))

    if not blocks:
        raise ValueError(f"Le fichier Word « {getattr(uploaded_file, 'name', '')} » est vide.")
    return "\n\n".join(blocks)


def extract_text_from_excel(uploaded_file) -> str:
    """Extrait toutes les feuilles et lignes d'un fichier Excel (.xlsx, .xls) ou CSV (DPGF / DQE / listes)."""
    import pandas as pd

    name = getattr(uploaded_file, "name", "tableau.xlsx")
    suffix = Path(name).suffix.lower()
    raw_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

    try:
        if suffix == ".csv":
            try:
                df = pd.read_csv(io.BytesIO(raw_bytes))
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(raw_bytes), encoding="latin-1", sep=None, engine="python")
            sheets = {"Feuille_CSV": df}
        else:
            sheets = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None)
    except Exception as exc:
        raise ValueError(f"Impossible de lire le tableur « {name} » : {exc}") from exc

    out_sheets: list[str] = []
    for sheet_name, df in sheets.items():
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            continue
        df = df.fillna("")
        headers = [str(col).strip() for col in df.columns]
        lines = [f"--- Feuille : {sheet_name} ---", " | ".join(headers)]
        for _, row in df.iterrows():
            vals = [re.sub(r"\s+", " ", str(v).strip()) for v in row.values]
            if any(vals):
                lines.append(" | ".join(vals))
        out_sheets.append("\n".join(lines))

    if not out_sheets:
        raise ValueError(f"Le tableur « {name} » ne contient aucune donnée exploitable.")
    return "\n\n".join(out_sheets)


def extract_text_from_document(uploaded_file) -> str:
    """Extrait le texte de n'importe quel fichier supporté (PDF, Word .docx, Excel .xlsx/.xls/.csv, .txt)."""
    name = getattr(uploaded_file, "name", "document.pdf")
    suffix = Path(name).suffix.lower()
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    if suffix == ".pdf":
        return extract_text_from_pdf(uploaded_file)
    elif suffix == ".docx":
        return extract_text_from_docx(uploaded_file)
    elif suffix in (".xlsx", ".xls", ".csv"):
        return extract_text_from_excel(uploaded_file)
    elif suffix == ".txt":
        raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
        return raw.decode("utf-8", errors="replace")
    else:
        raise ValueError(f"Format de fichier non supporté ({suffix}) pour « {name} ».")


def extract_category_documents(files_list, category_label: str) -> tuple[str, list[str]]:
    """
    Extrait et concatène un ou plusieurs fichiers d'une même catégorie
    avec des en-têtes clairs indiquant le nom de chaque fichier source.
    Retourne (texte_concatene, liste_des_logs_par_fichier).
    """
    if not files_list:
        return "", []
    if not isinstance(files_list, list):
        files_list = [files_list]

    extracted_parts: list[str] = []
    logs: list[str] = []
    for f in files_list:
        if f is None:
            continue
        txt = extract_text_from_document(f)
        extracted_parts.append(
            f"================ {category_label.upper()} : {f.name} ================\n{txt}"
        )
        logs.append(f"`{f.name}` ({len(txt):,} caractères)")
    return "\n\n".join(extracted_parts), logs


def chunk_text(text: str, max_chars: int = 60_000) -> list[str]:
    """
    Découpe un texte long en morceaux (chunks) de taille maîtrisée.
    Chaque chunk fait au maximum `max_chars` caractères.
    """
    if len(text) <= max_chars:
        return [text]

    pages = re.split(r"(?=--- Page \d+|--- Feuille :|================ )", text)
    chunks: list[str] = []
    current_chunk = ""

    for page in pages:
        if len(current_chunk) + len(page) > max_chars and current_chunk:
            chunks.append(current_chunk.strip())
            current_chunk = page
        else:
            current_chunk += page

    if current_chunk.strip():
        chunks.append(current_chunk.strip())

    return chunks


# ─────────────────────────────────────────────────────────────────────────
# 4. PROMPTS D'ANALYSE ET DE RÉDACTION MULTI-DOCUMENTS
# ─────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT_ANALYSE = textwrap.dedent("""\
Tu es un ingénieur études de prix et directeur technique expert en marchés publics et privés du BTP en France.
On te fournit les pièces d'un Dossier de Consultation des Entreprises (DCE) et/ou les documents de l'entreprise candidate, répartis jusqu'en 5 catégories :
1. Règlement de Consultation (RC)
2. Cahiers des Clauses Techniques Particulières (CCTP Général Lot 00 + CCTP des lots techniques)
3. DPGF / DQE (Devis Quantitatif Estimatif : quantités m², m³, ml, unités)
4. Annexes Chantier (PGC Sécurité, Plan de masse, Étude de sol)
5. Documents de l'Entreprise (Ancien mémoire, organigramme, CV, personnel, parc matériel, CACES/AIPR, certifications Qualibat/RGE/FNTP, fournisseurs, déchèteries)

Ta mission : extraire et croiser TOUTES les informations clés sous forme d'un JSON structuré exhaustif.

Le JSON doit contenir EXACTEMENT les clés suivantes :
{
  "nom_projet": "Nom ou intitulé exact du projet / marché",
  "maitre_ouvrage": "Nom du maître d'ouvrage (commanditaire) et maître d'œuvre si mentionné",
  "lieu_execution": "Adresse exacte et localisation du chantier",
  "dates_cles": {
    "date_remise_offres": "Date limite de remise des offres",
    "delai_execution": "Durée globale ou par phase / délai d'exécution des travaux",
    "date_debut_prevue": "Date de démarrage prévue"
  },
  "reglement_consultation_rc": {
    "criteres_valeur_technique": [
      {"ordre": 1, "critere": "Intitulé exact du critère ou sous-critère dans l'ordre du RC", "bareme_points": "Pondération ou points", "attendus": "Détails exigés par l'acheteur"}
    ],
    "plan_impose": ["Titres et sous-titres imposés dans l'ordre exact du RC"],
    "nombre_pages_max": "Nombre de pages maximum imposé par le RC (ou 'Non spécifié')"
  },
  "contraintes_chantier": [
    "Contraintes précises : adresse, accès, gabarits, horaires de travail, riverains, site occupé, maintien en exploitation, réseaux"
  ],
  "lots_concernes": ["Liste des lots traités (ex: Lot 00 Général, Lot 01 Gros Œuvre, Lot 02 VRD...)"],
  "exigences_techniques_par_lot": [
    {
      "lot": "Numéro et intitulé du lot",
      "description_travaux": "Travaux demandés",
      "materiaux_requis": ["Matériaux précis, marques/références ou performances exigées"],
      "normes_dtu": ["DTU, normes NF/EN, Eurocodes, RGE, Qualibat exigés"],
      "methodes_pose_autocontroles": ["Méthodes de mise en œuvre, tolérances, essais et autocontrôles exigés"]
    }
  ],
  "quantites_dpgf": [
    {"poste": "Désignation de l'ouvrage ou de la tâche", "quantite": "Quantité chiffrée exacte", "unite": "m², m³, ml, u, forfait...", "lot": "Lot concerné"}
  ],
  "annexes_securite_installation": {
    "risques_specifiques_pgc": ["Risques sécurité spécifiques tirés du PGC / étude de sol"],
    "coactivite": ["Contraintes de coactivité entre corps d'état ou avec l'exploitant"],
    "installation_chantier": ["Base vie, cantonnements, clôtures, zones de stockage, grues/levage, accès engins"],
    "contraintes_sol_reseaux": ["Données géotechniques (étude de sol), blindage, nappe, réseaux enterrés/aériens"]
  },
  "gestion_dechets_sosed": [
    "Exigences SOSED/SOGED du CCTP, filières de tri, traçabilité Bordereaux de Suivi des Déchets"
  ],
  "ressources_entreprise_extraites": {
    "identite_et_chiffres": "Nom, SIRET, effectif, CA, implantations extraits des documents entreprise",
    "organigramme_et_cv": ["Noms, fonctions, ancienneté, diplômes et CV de l'encadrement et des chefs d'équipe"],
    "effectifs_et_habilitations": ["Équipes d'exécution et diplômes sécurité : CACES, AIPR, SST, habilitations électriques, travail en hauteur"],
    "parc_materiel_engins": ["Liste exacte des engins, camions, grues, échafaudages, outillages et petits matériels"],
    "certifications": ["Certifications détenues : Qualibat, RGE, FNTP, ISO, MASE..."],
    "fournisseurs_habituels": ["Noms des fournisseurs de matériaux et marques utilisés par l'entreprise"],
    "decheteries_partenaires": ["Noms, adresses et distances des déchèteries / centres de tri habituels de l'entreprise"],
    "references_chantiers": ["Chantiers similaires réalisés par l'entreprise"]
  },
  "criteres_notation": [
    {"critere": "Nom du critère", "ponderation": "Pondération ou points"}
  ],
  "exigences_techniques": ["Synthèse des exigences techniques majeures"],
  "materiaux_requis": [
    {"materiau": "Nom du matériau", "specification": "Détail / norme / référence"}
  ],
  "ressources_humaines": [
    {"profil": "Type de profil requis", "detail": "Qualification, certification…"}
  ],
  "normes_certifications": ["Liste des normes et certifications exigées"],
  "contraintes_securite_environnement": ["Synthèse des contraintes SSE"]
}

IMPORTANT :
- Sois ultra-précis : conserve tous les chiffres, quantités DPGF (m², m³, ml, u), numéros de DTU, noms propres, marques, modèles d'engins et distances.
- Si une information n'est pas mentionnée dans les documents fournis, mets "Non spécifié".
- Réponds UNIQUEMENT avec le JSON valide, sans texte avant ni après.
""")

SYSTEM_PROMPT_REDACTION = textwrap.dedent("""\
Tu es un Directeur Technique BTP expert en rédaction de Mémoires Techniques gagnants pour les appels d'offres publics et privés en France.

On te fournit :
1. L'analyse structurée et croisée (JSON) issue des documents uploadés :
   - Règlement de Consultation (RC)
   - Un ou plusieurs CCTP (CCTP Général Lot 00 + CCTP des lots techniques)
   - La DPGF / DQE (quantités exactes : m², m³, ml, unités)
   - Les Annexes Chantier (PGC Sécurité, Plan de masse, Étude de sol)
   - Les Documents de l'Entreprise (Ancien mémoire, organigramme, CV, personnel, parc matériel, CACES/AIPR, certifications, fournisseurs, déchèteries)
2. Le profil de l'entreprise candidate et des extraits directs des pièces fournies.

Ta mission : rédiger un MÉMOIRE TECHNIQUE complet, ultra-précis, personnalisé et structuré pour obtenir la note technique maximale.

═══════════════════════════════════════════════════════════════════════════
RÈGLES IMPÉRATIVES DE RÉDACTION ET DE CROISEMENT DES DOCUMENTS :
═══════════════════════════════════════════════════════════════════════════

1. SUIVRE L'ORDRE EXACT DU RÈGLEMENT DE CONSULTATION (RC) :
   - Si le RC impose un plan ou un ordre précis de critères/sous-critères de la "Valeur Technique" (voir `reglement_consultation_rc` dans le JSON), le sommaire et les grands titres (`# 1. ...`, `# 2. ...`) de ton mémoire technique DOIVENT reprendre **exactement l'ordre et les intitulés des critères de notation du RC** (ainsi que la contrainte éventuelle de nombre de pages).
   - Si aucun RC n'a été fourni ou si le RC ne précise pas de plan, utilise exactement les 8 sections obligatoires détaillées plus bas.

2. FUSIONNER LES DIFFÉRENTS CCTP :
   - Si plusieurs CCTP ont été fournis (ex: CCTP Général Lot 00 + CCTP de lots techniques comme Gros Œuvre, VRD, Électricité, Plomberie…), combine systématiquement les contraintes générales du chantier (accès, horaires, riverains, site occupé, gestion des interfaces) avec les prescriptions techniques d'exécution propres à chaque lot.

3. CROISER OBLIGATOIREMENT CCTP + DPGF + ENTREPRISE (ZÉRO PHRASE VAGUE) :
   - **Interdiction formelle de faire des phrases génériques ou vagues.**
   - Pour chaque tâche ou ouvrage demandé dans le CCTP, tu DOIS croiser dans le même paragraphe :
     a) **La technique et la norme DTU** exigées au CCTP,
     b) **La quantité exacte** trouvée dans la DPGF / DQE (`m²`, `m³`, `ml`, `unités`),
     c) **Le matériel précis** (engins, camions, outillage) et **le personnel dédié** (noms/fonctions, qualifications, CACES, AIPR, habilitations) trouvés dans les documents de l'entreprise.
   - Exemple attendu : *"Pour la réalisation des 450 m² de cloisons acoustiques (poste 3.2 de la DPGF) conformément au DTU 25.41, nous mobiliserons une équipe de 2 plaquistes qualifiés encadrés par notre chef de chantier [Nom], équipés de lève-plaques et lasers rotatifs de notre parc matériel, à une cadence de 45 m²/jour soit 10 jours ouvrés."*

4. SECTIONS OBLIGATOIRES À RÉDIGER (intégrées dans l'ordre du RC, ou comme structure principale si pas de plan RC) :
   # 1. COMPRÉHENSION DES ENJEUX ET CONTRAINTES DU CHANTIER
      - Adresse exacte, contexte, accès chantier, circulation, riverains, horaires, site occupé ou maintien en activité, synthèse des lots traités.
   # 2. INSTALLATION DE CHANTIER ET LOGISTIQUE
      - Base vie / cantonnements, zones de stockage et de déchargement, gestion des flux et approvisionnements, clôtures, signalisation, raccordements provisoires (en croisant CCTP Lot 00 + Plan de masse / PGC).
   # 3. MOYENS HUMAINS ET MOYENS MATÉRIELS DÉDIÉS AU CHANTIER
      - Organigramme de chantier nominatif, CV et références de l'encadrement, composition des équipes par tâche, habilitations sécurité (CACES, AIPR, SST, habilitations électriques, travail en hauteur), certifications de l'entreprise (Qualibat, RGE, FNTP).
      - Liste précise des engins, camions, moyens de levage, échafaudages et outillages affectés au chantier.
   # 4. MÉTHODOLOGIE D'EXÉCUTION DÉTAILLÉE PHASE PAR PHASE ET AUTOCONTRÔLES QUALITÉ
      - Détail opérationnel phase par phase et lot par lot selon le(s) CCTP, en citant systématiquement les quantités de la DPGF, les normes DTU applicables, les points d'arrêt et les fiches d'autocontrôle qualité (PAQ).
   # 5. PROVENANCE DES MATÉRIAUX, FOURNISSEURS ET FICHES TECHNIQUES PRÉVUES AU CCTP
      - Tableau ou liste structurée des matériaux exigés au CCTP avec leurs caractéristiques techniques, normes/marquages CE/NF, et les fournisseurs habituels de l'entreprise.
   # 6. PLANNING PRÉVISIONNEL ET CADENCES BASÉS SUR LES QUANTITÉS DE LA DPGF
      - Justification réaliste des durées et cadences (quantité DPGF ÷ rendement journalier de l'équipe = durée en jours), phasage temporel, chemin critique, jalons et livrables (DOE, DICT, PPSPS).
   # 7. SÉCURITÉ ET SANTÉ SUR LE CHANTIER (PGC / PPSPS)
      - Analyse des risques spécifiques au chantier tirés du PGC et de l'étude de sol (travaux en hauteur, fouilles, réseaux, levage, poussières/bruit), gestion de la coactivité, protections collectives et individuelles (EPI), accueil sécurité.
   # 8. ENVIRONNEMENT ET GESTION DES DÉCHETS (SOSED)
      - Schéma d'Organisation et de Suivi de l'Élimination des Déchets (SOSED) : estimation et tri par catégorie (DID, DIB, inertes), bennes sur site, bordereaux BSD, nom et distance exacte des déchèteries / centres de valorisation habituels, réduction des nuisances (bruit, poussières, eaux de lavage).

5. RÈGLE STRICTE ANTI-INVENTION :
   - Tu ne dois **JAMAIS inventer** une donnée propre à l'entreprise (nom de salarié, numéro SIRET, marque d'engin non citée, nom de fournisseur, nom ou distance de déchèterie, référence chantier) ou propre au chantier (adresse, quantité, date, contrainte) si elle n'est présente dans aucun des documents uploadés ni dans le profil entreprise.
   - Chaque fois qu'une information spécifique à l'entreprise ou au chantier manque dans les pièces fournies, écris explicitement un marqueur bien visible de la forme :
     `[À COMPLÉTER : préciser ici ...]`
     (Exemples : `[À COMPLÉTER : préciser ici le nom et l'ancienneté du conducteur de travaux]`, `[À COMPLÉTER : préciser ici le nom et la distance en km de la déchèterie partenaire]`, `[À COMPLÉTER : préciser ici la quantité exacte issue de la DPGF]`).
""")


def _call_openai(api_key: str, system_prompt: str, user_content: str, model: str = "gpt-4o") -> str:
    """Appelle le service de génération (OpenAI)."""
    from openai import OpenAI, APIError, AuthenticationError

    client = OpenAI(api_key=api_key)
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content},
            ],
            temperature=0.3,
            max_tokens=16_000,
        )
        return response.choices[0].message.content.strip()
    except AuthenticationError:
        raise ValueError(
            "❌ Clé d'authentification système invalide. Veuillez contacter l'administrateur."
        )
    except APIError as e:
        raise ValueError(f"❌ Erreur du service de traitement : {e.message}")


def _call_deepseek(api_key: str, system_prompt: str, user_content: str) -> str:
    """Appelle le service de génération (DeepSeek)."""
    from openai import OpenAI, AuthenticationError

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content},
            ],
            temperature=0.3,
            max_tokens=8000,
        )
        return response.choices[0].message.content.strip()
    except AuthenticationError:
        raise ValueError("❌ Clé d'authentification système invalide. Veuillez contacter l'administrateur.")
    except Exception as e:
        raise ValueError(f"❌ Erreur du service de traitement : {e}")


def _call_gemini(api_key: str, system_prompt: str, user_content: str, model: str = "gemini-2.0-flash") -> str:
    """Appelle le service de génération (Gemini)."""
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    try:
        response = client.models.generate_content(
            model=model,
            contents=user_content,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                temperature=0.3,
                max_output_tokens=16_000,
            ),
        )
        return response.text.strip()
    except Exception as e:
        error_msg = str(e).lower()
        if "api key" in error_msg or "unauthorized" in error_msg or "403" in error_msg:
            raise ValueError(
                "❌ Clé d'authentification système invalide. Veuillez contacter l'administrateur."
            )
        raise ValueError(f"❌ Erreur du service de traitement : {e}")


def call_llm(api_key: str, provider: str, system_prompt: str, user_content: str) -> str:
    """Dispatch vers le moteur de traitement configuré."""
    if "openai" in provider.lower():
        return _call_openai(api_key, system_prompt, user_content)
    elif "deepseek" in provider.lower():
        return _call_deepseek(api_key, system_prompt, user_content)
    else:
        return _call_gemini(api_key, system_prompt, user_content)


# ─────────────────────────────────────────────────────────────────────────
# 5. TÂCHE A : ANALYSE MULTI-DOCUMENTS (RC, CCTP, DPGF, ANNEXES, ENTREPRISE) → JSON
# ─────────────────────────────────────────────────────────────────────────

def _merge_json_values(base_val, new_val):
    """Fusionne récursivement deux valeurs JSON extraites de différents chunks."""
    empty_markers = (None, "", "Non spécifié", "Non spécifié dans cette section", [], {})
    if isinstance(base_val, list) and isinstance(new_val, list):
        existing = {json.dumps(item, ensure_ascii=False, sort_keys=True) for item in base_val}
        for item in new_val:
            serialized = json.dumps(item, ensure_ascii=False, sort_keys=True)
            if serialized not in existing:
                base_val.append(item)
                existing.add(serialized)
        return base_val
    if isinstance(base_val, dict) and isinstance(new_val, dict):
        for k, v in new_val.items():
            if k not in base_val or base_val[k] in empty_markers:
                base_val[k] = v
            else:
                base_val[k] = _merge_json_values(base_val[k], v)
        return base_val
    if base_val in empty_markers and new_val not in empty_markers:
        return new_val
    return base_val


def analyse_cctp(api_key: str, provider: str, dossier_text: str) -> dict:
    """
    Envoie le texte combiné des documents uploadés (RC, CCTP, DPGF, Annexes, Entreprise)
    au moteur d'analyse pour extraire et croiser toutes les informations clés au format JSON.
    Gère le chunking si le dossier complet est très volumineux.
    """
    chunks = chunk_text(dossier_text)
    all_analyses: list[dict] = []

    for i, chunk in enumerate(chunks):
        if len(chunks) > 1:
            user_msg = (
                f"[Partie {i + 1}/{len(chunks)} du Dossier de Consultation et Entreprise]\n\n"
                f"{chunk}\n\n"
                f"Analyse cette partie et extrais toutes les informations trouvées (RC, CCTP, DPGF, Annexes, Entreprise) "
                f"au format JSON demandé. Si certaines informations ne figurent pas dans cette partie, "
                f"mets 'Non spécifié' ou une liste vide []."
            )
        else:
            user_msg = (
                "Voici l'ensemble des documents uploadés (RC, CCTP, DPGF/DQE, Annexes Chantier, Documents Entreprise) à analyser et croiser :\n\n"
                f"{chunk}"
            )

        raw_response = call_llm(api_key, provider, SYSTEM_PROMPT_ANALYSE, user_msg)

        json_match = re.search(r"```json\s*(.*?)\s*```", raw_response, re.DOTALL)
        json_str = json_match.group(1) if json_match else raw_response

        try:
            parsed = json.loads(json_str)
            all_analyses.append(parsed)
        except json.JSONDecodeError:
            repair_prompt = (
                "Le texte suivant devait être un JSON valide mais il est malformé. "
                "Corrige-le et renvoie UNIQUEMENT le JSON valide :\n\n" + json_str
            )
            repaired = call_llm(api_key, provider, "Tu es un assistant qui corrige du JSON.", repair_prompt)
            json_match2 = re.search(r"```json\s*(.*?)\s*```", repaired, re.DOTALL)
            json_str2 = json_match2.group(1) if json_match2 else repaired
            try:
                parsed = json.loads(json_str2)
                all_analyses.append(parsed)
            except json.JSONDecodeError:
                raise ValueError(
                    "❌ Le système n'a pas pu structurer les données après deux tentatives. "
                    "Essayez de réduire la taille des fichiers déposés."
                )

    if len(all_analyses) == 1:
        return all_analyses[0]

    merged = all_analyses[0]
    for analysis in all_analyses[1:]:
        merged = _merge_json_values(merged, analysis)

    return merged


# ─────────────────────────────────────────────────────────────────────────
# 6. TÂCHE B : RÉDACTION DU MÉMOIRE TECHNIQUE CROISÉ
# ─────────────────────────────────────────────────────────────────────────

def rediger_memoire(
    api_key: str,
    provider: str,
    analyse_json: dict,
    profil_entreprise: Optional[str] = None,
    extraits_par_categorie: Optional[dict[str, str]] = None,
) -> str:
    """
    Demande au LLM de rédiger le mémoire technique complet en croisant :
    1. L'analyse JSON multi-documents (RC + CCTP + DPGF + Annexes + Docs Entreprise)
    2. Le profil de l'entreprise enregistré dans l'application (s'il existe)
    3. Les extraits directs des 5 catégories de documents uploadés
    """
    user_content = (
        "## 1. Analyse structurée et croisée des documents uploadés (JSON)\n"
        f"```json\n{json.dumps(analyse_json, ensure_ascii=False, indent=2)}\n```\n\n"
    )

    if profil_entreprise and profil_entreprise.strip():
        user_content += f"## 2. Profil enregistré de l'entreprise candidate\n{profil_entreprise}\n\n"
    else:
        user_content += (
            "## 2. Profil enregistré de l'entreprise candidate\n"
            "Aucun profil pré-enregistré dans l'onglet Profil. Utilise exclusivement les informations "
            "extraites de la catégorie 5 « Documents de l'Entreprise » (organigramme, CV, personnel, "
            "parc matériel, CACES/AIPR, certifications, fournisseurs, déchèteries). "
            "Si une information propre à l'entreprise ou au chantier n'est présente dans aucun document uploadé, "
            "n'invente rien et insère obligatoirement `[À COMPLÉTER : préciser ici ...]`.\n\n"
        )

    if extraits_par_categorie:
        user_content += "## 3. Extraits directs des pièces uploadées par catégorie\n"
        for cat_label, cat_text in extraits_par_categorie.items():
            if cat_text and cat_text.strip():
                # On inclut un extrait généreux par catégorie pour garantir une fidélité maximale aux chiffres et textes
                extrait_borne = cat_text[:18_000]
                user_content += f"\n### {cat_label}\n{extrait_borne}\n"
        user_content += "\n"

    user_content += (
        "Rédige maintenant le Mémoire Technique complet, personnalisé et ultra-précis en appliquant strictement :\n"
        "- L'ordre exact des critères de notation du Règlement de Consultation (RC) s'il est fourni (sinon les 8 sections obligatoires),\n"
        "- La fusion de tous les CCTP (Général Lot 00 + lots techniques),\n"
        "- Le croisement systématique CCTP (méthode + norme DTU) + DPGF (quantités exactes m², m³, ml, u) + Entreprise (personnel habilité + matériel précis),\n"
        "- La règle anti-invention : écris `[À COMPLÉTER : préciser ici...]` pour toute donnée spécifique absente des documents."
    )

    return call_llm(api_key, provider, SYSTEM_PROMPT_REDACTION, user_content)


# ─────────────────────────────────────────────────────────────────────────
# 7. GÉNÉRATION DU FICHIER WORD (.docx)
# ─────────────────────────────────────────────────────────────────────────

def generate_docx(memoire_text: str, analyse_json: dict, entreprise: dict) -> io.BytesIO:
    """
    Convertit le texte Markdown du mémoire technique en un fichier Word formaté.
    Ajoute les infos de l'entreprise sur la page de garde.
    """
    doc = Document()

    # ── Styles du document ──────────────────────────────────────────────
    style_normal = doc.styles["Normal"]
    font = style_normal.font
    font.name = "Calibri"
    font.size = Pt(11)

    # ── Page de garde ───────────────────────────────────────────────────
    titre_garde = doc.add_paragraph()
    titre_garde.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titre_garde.add_run("MÉMOIRE TECHNIQUE")
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_paragraph()

    # Nom du projet
    nom_projet = analyse_json.get("nom_projet", "Projet")
    p_projet = doc.add_paragraph()
    p_projet.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_p = p_projet.add_run(nom_projet)
    run_p.font.size = Pt(16)
    run_p.font.color.rgb = RGBColor(0, 51, 102)

    # Maître d'ouvrage
    mo = analyse_json.get("maitre_ouvrage", "")
    if mo and mo != "Non spécifié":
        p_mo = doc.add_paragraph()
        p_mo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_mo = p_mo.add_run(f"Maître d'ouvrage : {mo}")
        run_mo.font.size = Pt(12)
        run_mo.italic = True

    doc.add_paragraph()

    # Infos entreprise sur la page de garde (profil ou extrait des documents entreprise)
    res_ent = analyse_json.get("ressources_entreprise_extraites", {}) if isinstance(analyse_json.get("ressources_entreprise_extraites"), dict) else {}
    nom_entreprise = entreprise.get("nom", "").strip()
    if not nom_entreprise:
        extracted_nom = str(res_ent.get("nom_entreprise", "")).strip()
        if extracted_nom and extracted_nom not in ("Non spécifié", "Non spécifié dans cette section"):
            nom_entreprise = extracted_nom

    if nom_entreprise:
        p_ent = doc.add_paragraph()
        p_ent.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_ent = p_ent.add_run(f"Présenté par : {nom_entreprise}")
        run_ent.font.size = Pt(14)
        run_ent.bold = True

        details = []
        if entreprise.get("adresse", "").strip():
            details.append(entreprise["adresse"].strip())
        if entreprise.get("telephone", "").strip():
            details.append(f"Tél : {entreprise['telephone'].strip()}")
        if entreprise.get("email", "").strip():
            details.append(f"Email : {entreprise['email'].strip()}")
        if details:
            p_details = doc.add_paragraph()
            p_details.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_details = p_details.add_run(" · ".join(details))
            run_details.font.size = Pt(10)

        certifs = []
        if entreprise.get("certifications", "").strip():
            certifs.append(entreprise["certifications"].strip())
        if entreprise.get("qualifications", "").strip():
            certifs.append(entreprise["qualifications"].strip())
        if not certifs and isinstance(res_ent.get("certifications"), list) and res_ent["certifications"]:
            certifs.extend([str(c) for c in res_ent["certifications"][:5]])
        if certifs:
            p_certif = doc.add_paragraph()
            p_certif.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_certif = p_certif.add_run(" | ".join(certifs))
            run_certif.font.size = Pt(10)
            run_certif.italic = True

    # Date
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_date = p_date.add_run(f"Date : {datetime.now().strftime('%d/%m/%Y')}")
    run_date.font.size = Pt(12)

    doc.add_page_break()

    # ── Sommaire ────────────────────────────────────────────────────────
    doc.add_heading("SOMMAIRE", level=1)
    doc.add_paragraph(
        "[Le sommaire sera généré automatiquement dans Word : "
        "onglet Références → Table des matières]"
    )
    doc.add_page_break()

    def _add_formatted_runs(paragraph, text_str: str):
        """Ajoute le texte dans un paragraphe Word en mettant en gras **...** et en évidence [À COMPLÉTER...]."""
        tokens = re.split(r"(\[À COMPLÉTER[^\]]*\]|\*\*.*?\*\*)", text_str)
        for tok in tokens:
            if not tok:
                continue
            if tok.startswith("[À COMPLÉTER") and tok.endswith("]"):
                r = paragraph.add_run(tok)
                r.bold = True
                r.font.color.rgb = RGBColor(217, 83, 79)
            elif tok.startswith("**") and tok.endswith("**"):
                r = paragraph.add_run(tok[2:-2])
                r.bold = True
            else:
                paragraph.add_run(tok)

    # ── Contenu du mémoire ──────────────────────────────────────────────
    lines = memoire_text.split("\n")
    for line in lines:
        stripped = line.strip()

        if not stripped:
            doc.add_paragraph("")
            continue

        if stripped.startswith("### "):
            doc.add_heading(stripped[4:], level=3)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:], level=2)
        elif stripped.startswith("# "):
            doc.add_heading(stripped[2:], level=1)
        elif stripped.startswith("- ") or stripped.startswith("* "):
            bullet_text = stripped[2:]
            p = doc.add_paragraph(style="List Bullet")
            _add_formatted_runs(p, bullet_text)
        elif re.match(r"^\d+\.\s", stripped):
            text = re.sub(r"^\d+\.\s", "", stripped)
            p = doc.add_paragraph(style="List Number")
            _add_formatted_runs(p, text)
        else:
            p = doc.add_paragraph()
            _add_formatted_runs(p, stripped)

    # ── Sauvegarde en mémoire ───────────────────────────────────────────
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer



# ─────────────────────────────────────────────────────────────────────────
# 8. GÉNÉRATION DU FICHIER PDF (HAUTE QUALITÉ)
# ─────────────────────────────────────────────────────────────────────────
def generate_pdf(memoire_text: str, analyse_json: dict, entreprise: dict) -> io.BytesIO:
    font_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    reg_font = os.path.join(font_dir, "CustomFont-Regular.ttf")
    bold_font = os.path.join(font_dir, "CustomFont-Bold.ttf")
    italic_font = os.path.join(font_dir, "CustomFont-Italic.ttf")

    res_ent = analyse_json.get("ressources_entreprise_extraites", {}) if isinstance(analyse_json.get("ressources_entreprise_extraites"), dict) else {}
    extracted_nom = str(res_ent.get("nom_entreprise", "")).strip()
    if extracted_nom in ("Non spécifié", "Non spécifié dans cette section"):
        extracted_nom = ""
    nom_entreprise = entreprise.get("nom", "").strip() or extracted_nom or "MarketBuilding"
    nom_projet = str(analyse_json.get("nom_projet", "Projet BTP")).strip()

    class ExecutiveBTPPDF(FPDF):
        def __init__(self):
            super().__init__()
            if os.path.exists(reg_font) and os.path.exists(bold_font) and os.path.exists(italic_font):
                self.add_font("AppFont", "", reg_font)
                self.add_font("AppFont", "B", bold_font)
                self.add_font("AppFont", "I", italic_font)
                self.main_font = "AppFont"
            else:
                self.main_font = "Helvetica"

        def header(self):
            if self.page_no() > 1:
                self.set_font(self.main_font, "B", 8)
                self.set_text_color(30, 64, 175)
                self.cell(90, 8, nom_entreprise.upper(), align="L")
                self.set_font(self.main_font, "I", 8)
                self.set_text_color(100, 116, 139)
                proj_title = (nom_projet[:45] + "...") if len(nom_projet) > 45 else nom_projet
                self.cell(0, 8, f"Mémoire Technique - {proj_title}", align="R", new_x="LMARGIN", new_y="NEXT")
                self.set_draw_color(203, 213, 225)
                self.set_line_width(0.3)
                self.line(10, 18, 200, 18)
                self.ln(5)

        def footer(self):
            if self.page_no() > 1:
                self.set_y(-15)
                self.set_font(self.main_font, "I", 7)
                self.set_text_color(148, 163, 184)
                self.cell(100, 10, "Document confidentiel - Remis dans le cadre de la consultation", align="L")
                self.set_font(self.main_font, "", 8)
                self.cell(0, 10, f"Page {self.page_no()}", align="R")

    pdf = ExecutiveBTPPDF()
    pdf.set_auto_page_break(auto=True, margin=20)

    # ── Page de garde ───────────────────────────────────────────────────
    pdf.add_page()

    # Bandeau supérieur
    pdf.set_fill_color(30, 58, 138)
    pdf.rect(0, 0, 210, 36, "F")
    pdf.set_xy(15, 11)
    pdf.set_font(pdf.main_font, "B", 11)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 6, "RÉPONSE À CONSULTATION - MARCHÉS DU BTP", align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(15)
    pdf.set_font(pdf.main_font, "I", 8)
    pdf.set_text_color(224, 231, 255)
    pdf.cell(0, 5, "DOSSIER D'OFFRE TECHNIQUE & MÉTHODOLOGIQUE", align="L", new_x="LMARGIN", new_y="NEXT")

    # Titre central
    pdf.set_y(65)
    pdf.set_font(pdf.main_font, "B", 24)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 12, "MÉMOIRE TECHNIQUE", align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.set_draw_color(59, 130, 246)
    pdf.set_line_width(0.8)
    pdf.line(75, 80, 135, 80)
    pdf.ln(12)

    # Nom du projet
    pdf.set_font(pdf.main_font, "B", 14)
    pdf.set_text_color(30, 64, 175)
    pdf.multi_cell(0, 7, nom_projet, align="C")
    pdf.ln(6)

    # Maître d'ouvrage
    mo = analyse_json.get("maitre_ouvrage", "").strip()
    if mo and mo != "Non spécifié":
        pdf.set_font(pdf.main_font, "I", 10)
        pdf.set_text_color(71, 85, 105)
        pdf.cell(0, 6, f"Maître d'ouvrage : {mo}", align="C", new_x="LMARGIN", new_y="NEXT")

    # Encadré Entreprise Soumissionnaire
    pdf.set_y(145)
    pdf.set_draw_color(226, 232, 240)
    pdf.set_fill_color(248, 250, 252)
    pdf.rect(20, 145, 170, 80, "DF")

    pdf.set_xy(30, 152)
    pdf.set_font(pdf.main_font, "B", 10)
    pdf.set_text_color(30, 64, 175)
    pdf.cell(0, 6, "ENTREPRISE CANDIDATE", align="L", new_x="LMARGIN", new_y="NEXT")

    pdf.set_x(30)
    pdf.set_font(pdf.main_font, "B", 13)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, nom_entreprise, align="L", new_x="LMARGIN", new_y="NEXT")

    pdf.set_x(30)
    pdf.set_font(pdf.main_font, "", 9)
    pdf.set_text_color(71, 85, 105)

    forme_siret = []
    if entreprise.get("forme_juridique"):
        forme_siret.append(entreprise["forme_juridique"])
    if entreprise.get("siret"):
        forme_siret.append(f"SIRET : {entreprise['siret']}")
    if forme_siret:
        pdf.cell(0, 5, " | ".join(forme_siret), align="L", new_x="LMARGIN", new_y="NEXT")
        pdf.set_x(30)

    if entreprise.get("adresse"):
        pdf.cell(0, 5, f"Adresse : {entreprise['adresse']}", align="L", new_x="LMARGIN", new_y="NEXT")
        pdf.set_x(30)

    contact_parts = []
    if entreprise.get("telephone"):
        contact_parts.append(entreprise["telephone"])
    if entreprise.get("email"):
        contact_parts.append(entreprise["email"])
    if contact_parts:
        pdf.cell(0, 5, f"Contact : {' | '.join(contact_parts)}", align="L", new_x="LMARGIN", new_y="NEXT")
        pdf.set_x(30)

    if entreprise.get("activites_principales"):
        act = entreprise["activites_principales"][:100]
        pdf.cell(0, 5, f"Activités : {act}", align="L", new_x="LMARGIN", new_y="NEXT")
        pdf.set_x(30)

    # Pied de page de garde
    pdf.set_y(250)
    pdf.set_font(pdf.main_font, "", 8)
    pdf.set_text_color(148, 163, 184)
    date_str = datetime.now().strftime("%d/%m/%Y")
    pdf.cell(0, 5, f"Date de remise : {date_str}  |  Document officiel pour instruction", align="C")

    # ── Corps du document ───────────────────────────────────────────────
    pdf.add_page()
    pdf.set_text_color(15, 23, 42)
    pdf.set_font(pdf.main_font, "", 10)

    # Nettoyage et formatage HTML pour write_html
    cleaned_md = memoire_text
    # Assurer que les listes à puces Markdown ont un saut de ligne propre
    cleaned_md = re.sub(r"(\n-\s)", r"\n\n- ", cleaned_md)
    html_content = markdown.markdown(cleaned_md, extensions=["tables"])

    try:
        pdf.write_html(html_content)
    except Exception as e:
        pdf.multi_cell(0, 5, f"Note de conversion : {e}")

    buffer = io.BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    return buffer


# ─────────────────────────────────────────────────────────────────────────
# 8.5 DÉMARRAGE DU WORKER D'ÉCOUTE AUTOMATIQUE DES MAILS & VUE PROJET ÉDITABLE
# ─────────────────────────────────────────────────────────────────────────
PIPELINE_CALLBACKS = {
    "extract_text": extract_text_from_document,
    "extract_category": extract_category_documents,
    "analyse_cctp": analyse_cctp,
    "rediger_memoire": rediger_memoire,
    "generate_docx": generate_docx,
    "generate_pdf": generate_pdf,
    "build_profil_text": build_profil_text,
    "call_llm": call_llm,
}

ea.ensure_background_imap_worker(PIPELINE_CALLBACKS, load_config)


@st.cache_data(show_spinner=False)
def _build_cached_archives_zip(file_signatures: tuple[tuple[str, float], ...]) -> bytes:
    """Construit et met en cache le pack .ZIP des archives pour éviter toute latence à l'ouverture de l'onglet Archives."""
    zip_arch_buf = io.BytesIO()
    with zipfile.ZipFile(zip_arch_buf, "w", zipfile.ZIP_DEFLATED) as zarch:
        for fname, _ in file_signatures:
            fpath = os.path.join("archives", fname)
            if os.path.exists(fpath):
                zarch.write(fpath, arcname=fname)
    return zip_arch_buf.getvalue()


def render_interactive_project_viewer(proj: dict, key_prefix: str):
    """
    Permet d'ouvrir, lire, modifier et télécharger (.docx et .pdf) tout mémoire technique
    synchronisé sur le site (avec badge 'Automatisé via Mail' ou 'Créé sur le site'),
    sans alourdir le DOM par des iframes PDF multiples.
    """
    pid = proj.get("project_id", "proj")
    origin_mail = proj.get("origin") == "email"
    badge_txt = "📧 Automatisé via Mail" if origin_mail else "🖥️ Créé sur le site"
    badge_bg = "rgba(139, 92, 246, 0.16)" if origin_mail else "rgba(37, 99, 235, 0.14)"
    badge_col = "#8B5CF6" if origin_mail else "#2563EB"
    lots_str = ", ".join(proj.get("lots_vises") or [])

    st.markdown(
        f"""
        <div class="mb-zone-header" style="border-left:4px solid {badge_col}; margin-bottom:10px;">
            <div class="mb-zone-top">
                <span class="mb-zone-num" style="background:{badge_bg}; color:{badge_col};">{badge_txt}</span>
                <span class="mb-zone-badge">📅 {proj.get('created_at', '')}</span>
            </div>
            <div class="mb-zone-title">{proj.get('nom_projet', 'Mémoire Technique')}</div>
            <div class="mb-zone-desc" style="margin-bottom:0;">
                <b>Pièces croisées :</b> <code>{proj.get('source_name', 'Pièces DCE')}</code>
                {f"<br><b>Expéditeur :</b> <code>{proj.get('sender_email', '')}</code> &nbsp;·&nbsp; <b>Lot(s) :</b> {lots_str or 'Tous lots'}" if origin_mail else ""}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    pdf_path = proj.get("pdf_path", "")
    docx_path = proj.get("docx_path", "")
    pdf_bytes = Path(pdf_path).read_bytes() if pdf_path and os.path.exists(pdf_path) else b""
    docx_bytes = Path(docx_path).read_bytes() if docx_path and os.path.exists(docx_path) else b""
    nom_fichier = proj.get("nom_fichier", "Memoire_BTP")

    dl_c1, dl_c2 = st.columns(2)
    with dl_c1:
        if pdf_bytes:
            st.download_button(
                "📥 Télécharger en PDF (.pdf)",
                data=pdf_bytes,
                file_name=f"Memoire_Technique_{nom_fichier}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
                key=f"{key_prefix}_dl_pdf_{pid}",
            )
    with dl_c2:
        if docx_bytes:
            st.download_button(
                "📝 Télécharger en Word (.docx)",
                data=docx_bytes,
                file_name=f"Memoire_Technique_{nom_fichier}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"{key_prefix}_dl_docx_{pid}",
            )

    vt1, vt2, vt3, vt4 = st.tabs([
        "📖 Lire le Mémoire",
        "✏️ Modifier & Régénérer (.docx / .pdf)",
        "👁️ Aperçu PDF",
        "🔍 5 Points de Vigilance & Balises [À COMPLÉTER]",
    ])

    with vt1:
        st.markdown(proj.get("memoire_md", ""))

    with vt2:
        st.caption(
            "Modifiez directement le texte du mémoire ci-dessous (ex: remplacez les balises `[À COMPLÉTER : ...]`). "
            "Cliquez sur **Enregistrer** pour mettre à jour instantanément le fichier **Word (.docx)** et le **PDF (.pdf)** sur le site."
        )
        edited_md = st.text_area(
            "Contenu du Mémoire Technique (Markdown)",
            value=proj.get("memoire_md", ""),
            height=420,
            key=f"{key_prefix}_edit_md_{pid}",
        )
        if st.button(
            "💾 Enregistrer les modifications & Régénérer Word (.docx) et PDF (.pdf)",
            type="primary",
            use_container_width=True,
            key=f"{key_prefix}_save_md_{pid}",
        ):
            ent_dict = proj.get("entreprise") or config.get("entreprise", {})
            analyse_dict = proj.get("analyse") or {"nom_projet": proj.get("nom_projet", "Projet BTP")}
            new_docx = generate_docx(edited_md, analyse_dict, ent_dict).getvalue()
            new_pdf = generate_pdf(edited_md, analyse_dict, ent_dict).getvalue()
            if pdf_path:
                with open(pdf_path, "wb") as f:
                    f.write(new_pdf)
            if docx_path:
                with open(docx_path, "wb") as f:
                    f.write(new_docx)
            new_balises = list(dict.fromkeys(re.findall(r"\[À COMPLÉTER[^\]]*\]", edited_md)))
            proj["memoire_md"] = edited_md
            proj["balises_a_completer"] = new_balises
            ea.save_project_to_db(proj)
            if proj.get("job_id"):
                ea.update_job_status(proj["job_id"], balises_a_completer=new_balises)
            st.success("✅ Mémoire modifié ! Les fichiers PDF et Word (.docx) ont été régénérés et synchronisés.")
            st.rerun()

    with vt3:
        if pdf_bytes:
            show_pdf_iframe = st.toggle(
                "👁️ Charger la visionneuse PDF haute définition",
                value=False,
                key=f"{key_prefix}_toggle_pdf_{pid}",
                help="Charge le lecteur PDF uniquement à la demande pour garantir une navigation ultra-fluide sans latence.",
            )
            if show_pdf_iframe:
                b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
                st.markdown(
                    f'<iframe src="data:application/pdf;base64,{b64_pdf}#view=FitH" width="100%" height="680" type="application/pdf"></iframe>',
                    unsafe_allow_html=True,
                )
            else:
                st.info("💡 Activez l'interrupteur ci-dessus pour prévisualiser les pages du PDF directement dans le navigateur, ou cliquez sur **« 📥 Télécharger en PDF »**.")
        else:
            st.info("Aperçu PDF non disponible.")

    with vt4:
        vig_pts = proj.get("vigilance_points") or ea.extract_five_vigilance_points(proj.get("analyse") or {})
        st.markdown("#### 🛡️ Les 5 Points de Vigilance Majeurs du Chantier (RC & CCTP)")
        for vp in vig_pts:
            st.markdown(f"- **{vp['titre']}** : {vp['detail']}")
        balises = proj.get("balises_a_completer") or list(dict.fromkeys(re.findall(r"\[À COMPLÉTER[^\]]*\]", proj.get("memoire_md", ""))))
        st.markdown(f"#### ⚠️ Balises `[À COMPLÉTER]` restantes ({len(balises)})")
        if balises:
            for b in balises:
                st.markdown(f"- `{b}`")
        else:
            st.success("✅ Aucune balise `[À COMPLÉTER]` restante dans ce mémoire.")


# ═══════════════════════════════════════════════════════════════════════
#  BARRE EXÉCUTIVE SUPÉRIEURE (STATUT TEMPS RÉEL BTP)
# ═══════════════════════════════════════════════════════════════════════
nom_ent_top = config["entreprise"].get("nom", "").strip() or "Espace Candidat BTP"
_top_projects = ea.load_projects_db()
_top_mail_count = sum(1 for p in _top_projects if p.get("origin") == "email")
_top_watcher_str = "🟢 Écoute Mail Active" if _watcher_on else "⏸️ Écoute Mail en Pause"
st.markdown(
    f"""
    <div class="mb-exec-topbar">
        <div class="mb-exec-brand">
            <span>🏗️ MarketBuilding Pro</span>
            <span style="color:var(--text-muted); font-weight:500;">|</span>
            <span style="color:#2563EB;">{nom_ent_top}</span>
        </div>
        <div class="mb-exec-pills">
            <span class="mb-exec-pill">{_top_watcher_str}</span>
            <span class="mb-exec-pill">📧 {_top_mail_count} Automatisé(s) via Mail</span>
            <span class="mb-exec-pill">📂 {nb_archives} Livrables (.pdf / .docx)</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════════════════════
#  PAGES DE L'APPLICATION
# ═══════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────
# PAGE : ⚙️ RÉGLAGES (PROTÉGÉ PAR LE CODE "morti333")
# ─────────────────────────────────────────────────────────────────────────
if page == "Paramètres":
    if not st.session_state.get("owner_unlocked", False):
        st.error("🔒 Accès réservé à l'administrateur. Veuillez saisir votre code dans « 🔑 Accès Administrateur ».")
        st.stop()
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">⚙️ CONFIGURATION SYSTÈME</div>
            <div class="mb-hero-title">Paramètres du Moteur d'Intelligence Artificielle</div>
            <p class="mb-hero-sub">
                Configurez votre clé d'authentification et choisissez le modèle d'IA utilisé pour l'analyse croisée des DCE et la rédaction des mémoires techniques.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Montrer si les secrets cloud sont détectés
    secrets_detected = False
    try:
        if st.secrets.get("api_key"):
            secrets_detected = True
    except Exception:
        pass

    if secrets_detected:
        st.success("☁️ **Secrets Cloud détectés** – Votre clé API est lue depuis les secrets Streamlit Cloud.")
    else:
        st.info("💻 **Mode local actif** – Vos réglages et votre thème sont enregistrés dans `config.json`.")

    with st.form("form_reglages"):
        new_provider = st.selectbox(
            "Moteur de traitement IA",
            options=["OpenAI (GPT-4o)", "Google Gemini (gemini-2.0-flash)", "DeepSeek (deepseek-chat)"],
            index=0 if "openai" in config["llm_provider"].lower() else (2 if "deepseek" in config["llm_provider"].lower() else 1),
            help="Choisissez l'algorithme d'analyse et de rédaction à utiliser.",
        )

        new_api_key = st.text_input(
            "🔑 Clé d'authentification système",
            value=config["api_key"],
            type="password",
            help="En cloud, configurez plutôt vos secrets dans le dashboard Streamlit.",
        )

        submitted = st.form_submit_button("💾 Sauvegarder les réglages", type="primary", use_container_width=True)

        if submitted:
            config["api_key"] = new_api_key
            config["llm_provider"] = new_provider
            save_config(config)
            st.success("✅ Réglages sauvegardés avec succès !")
            st.rerun()

    # Guide secrets cloud
    with st.expander("📖 Comment configurer les secrets sur Streamlit Cloud"):
        st.markdown(
            "Dans le dashboard Streamlit Cloud → votre app → **Settings** → **Secrets**, "
            "collez le contenu suivant au format TOML :\n\n"
            "```toml\n"
            "# Clé d'authentification système et fournisseur\n"
            'api_key = "sk-votre-cle-ici"\n'
            'llm_provider = "OpenAI (GPT-4o)"\n'
            "\n"
            "# Profil entreprise\n"
            "[entreprise]\n"
            'nom = "Nom de votre entreprise"\n'
            'forme_juridique = "SARL"\n'
            'siret = "123 456 789 00001"\n'
            'adresse = "12 rue du Bâtiment, 75001 Paris"\n'
            'telephone = "01 23 45 67 89"\n'
            'email = "contact@monentreprise.fr"\n'
            'dirigeant = "Jean Dupont"\n'
            'certifications = "RGE QualiPAC, RGE QualiBois"\n'
            'qualifications = "Qualibat 5112, Qualibat 5212"\n'
            "# ... ajoutez tous les champs souhaités\n"
            "```\n\n"
            "Les secrets sont **chiffrés** et persistent indéfiniment."
        )


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 🏢 MON ENTREPRISE
# ─────────────────────────────────────────────────────────────────────────
elif page == "Profil":
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">🏢 IDENTITÉ & RÉFÉRENCES ENTREPRISE</div>
            <div class="mb-hero-title">Profil de l'Entreprise Candidate</div>
            <p class="mb-hero-sub">
                Renseignez votre socle d'informations permanentes (SIRET, effectifs, assurances, certifications, parc matériel). Ces données seront croisées avec les pièces que vous déposez dans la zone 5 lors de chaque appel d'offres.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    ent = config["entreprise"]

    with st.form("form_entreprise"):
        st.subheader("📋 Identité administrative")

        col1, col2 = st.columns(2)
        with col1:
            ent["nom"] = st.text_input("Nom de l'entreprise *", value=ent.get("nom", ""))
            ent["forme_juridique"] = st.text_input("Forme juridique (SARL, SAS…)", value=ent.get("forme_juridique", ""))
            ent["siret"] = st.text_input("SIRET", value=ent.get("siret", ""))
            ent["dirigeant"] = st.text_input("Dirigeant / Gérant", value=ent.get("dirigeant", ""))
        with col2:
            ent["adresse"] = st.text_input("Adresse du siège / dépôt", value=ent.get("adresse", ""))
            ent["telephone"] = st.text_input("Téléphone", value=ent.get("telephone", ""))
            ent["email"] = st.text_input("Email", value=ent.get("email", ""))
            ent["site_web"] = st.text_input("Site web", value=ent.get("site_web", ""))

        st.subheader("📊 Chiffres clés")
        col3, col4 = st.columns(2)
        with col3:
            ent["date_creation"] = st.text_input("Date de création", value=ent.get("date_creation", ""))
            ent["effectif"] = st.text_input("Effectif (nombre de salariés)", value=ent.get("effectif", ""))
        with col4:
            ent["chiffre_affaires"] = st.text_input("Chiffre d'affaires annuel", value=ent.get("chiffre_affaires", ""))
            ent["zone_intervention"] = st.text_input("Zone d'intervention géographique", value=ent.get("zone_intervention", ""))

        st.subheader("🏆 Qualifications & Certifications")
        ent["activites_principales"] = st.text_area(
            "Activités principales / Corps d'état",
            value=ent.get("activites_principales", ""),
            height=80,
            help="Ex : Gros Œuvre, VRD, Électricité, Plomberie CVC, Menuiserie, Rénovation énergétique…",
        )
        col5, col6 = st.columns(2)
        with col5:
            ent["certifications"] = st.text_area(
                "Certifications (RGE, ISO, FNTP, MASE…)",
                value=ent.get("certifications", ""),
                height=80,
            )
        with col6:
            ent["qualifications"] = st.text_area(
                "Qualifications & Habilitations (Qualibat, Qualifelec, CACES, AIPR…)",
                value=ent.get("qualifications", ""),
                height=80,
            )

        ent["assurances"] = st.text_area(
            "Assurances (Décennale, RC Professionnelle, compagnie & n° de police)",
            value=ent.get("assurances", ""),
            height=80,
        )

        st.subheader("💼 Références, Matériel & Déchèteries / Fournisseurs")
        ent["references_chantiers"] = st.text_area(
            "Références chantiers similaires",
            value=ent.get("references_chantiers", ""),
            height=110,
        )
        ent["moyens_materiels"] = st.text_area(
            "Parc matériel, engins, camions, outillage",
            value=ent.get("moyens_materiels", ""),
            height=100,
        )

        st.subheader("📝 Présentation libre (Fournisseurs, Déchèteries SOSED, Historique)")
        ent["presentation_libre"] = st.text_area(
            "Fournisseurs habituels, centres de tri/déchèteries partenaires (nom et distance), valeurs…",
            value=ent.get("presentation_libre", ""),
            height=130,
        )

        submitted = st.form_submit_button("💾 Sauvegarder le profil entreprise", type="primary", use_container_width=True)

        if submitted:
            config["entreprise"] = ent
            save_config(config)
            st.success("✅ Profil entreprise sauvegardé !")
            st.rerun()

    # Aperçu du profil
    profil_preview = build_profil_text(ent)
    if profil_preview:
        with st.expander("👁️ Aperçu du profil tel qu'il sera intégré au document"):
            st.markdown(profil_preview)

    # Bouton d'export TOML pour faciliter la copie dans les secrets cloud
    if any(v.strip() for v in ent.values()):
        with st.expander("☁️ Exporter le profil au format Secrets Cloud (TOML)"):
            toml_lines = ["[entreprise]"]
            for key, value in ent.items():
                if value.strip():
                    escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
                    toml_lines.append(f'{key} = "{escaped}"')
            toml_output = "\n".join(toml_lines)
            st.code(toml_output, language="toml")
            st.caption("Copiez ce bloc dans vos Secrets Streamlit Cloud (Settings → Secrets) pour rendre le profil permanent.")


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 📄 GÉNÉRER UN MÉMOIRE
# ─────────────────────────────────────────────────────────────────────────
elif page == "Génération":
    nom_ent = config["entreprise"].get("nom", "").strip()
    sub_hero = (
        f"Bienvenue <b>{nom_ent}</b> · Déposez simultanément vos pièces de marché (<b>RC, CCTP, DPGF/DQE, Annexes</b>) et vos <b>Documents Entreprise</b> (PDF, Word ou Excel). "
        "Le moteur d'ingénierie les croise automatiquement pour rédiger un <b>Mémoire Technique</b> complet, chiffré et conforme au RC."
        if nom_ent
        else "Déposez simultanément vos pièces de marché (<b>RC, CCTP, DPGF/DQE, Annexes</b>) et vos <b>Documents Entreprise</b> (PDF, Word ou Excel). "
             "Le moteur d'ingénierie les croise automatiquement pour rédiger un <b>Mémoire Technique</b> complet, chiffré et conforme au RC."
    )
    st.markdown(
        f"""
        <div class="mb-hero">
            <div class="mb-hero-badge">🏗️ INGÉNIERIE D'APPELS D'OFFRES BTP · CROISEMENT MULTI-DOCUMENTS</div>
            <div class="mb-hero-title">Générateur de Mémoires Techniques sur-mesure</div>
            <p class="mb-hero-sub">{sub_hero}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not config["api_key"].strip():
        st.warning("⚠️ **Configuration système incomplète.** Veuillez contacter l'administrateur de la plateforme.")

    _feat_main = vt.featured_video()
    if _feat_main:
        with st.expander("🎬 Démonstration Vidéo : Du dépôt des pièces au téléchargement du Mémoire Technique", expanded=False):
            st.markdown(
                """
                <div class="mb-steps-grid">
                    <div class="mb-step-card" style="border-left:4px solid #3B82F6;">
                        <div class="mb-step-title" style="color:#3B82F6;">1. Importer vos documents (RC, CCTP, DPGF…)</div>
                        <div class="mb-step-desc">Déposez vos fichiers (PDF, Word, Excel) dans les 5 zones dédiées puis cliquez sur <b>« 🚀 Croiser les documents & Générer »</b>.</div>
                    </div>
                    <div class="mb-step-card" style="border-left:4px solid #F59E0B;">
                        <div class="mb-step-title" style="color:#F59E0B;">2. Boutons de téléchargement PDF & Word</div>
                        <div class="mb-step-desc">En bas du mémoire généré, cliquez sur le bouton bleu <b>« 📥 Télécharger en PDF »</b> ou sur <b>« 📝 Télécharger en Word (.docx) »</b>.</div>
                    </div>
                    <div class="mb-step-card" style="border-left:4px solid #10B981;">
                        <div class="mb-step-title" style="color:#10B981;">3. Visionneuse PDF & Onglet Archives</div>
                        <div class="mb-step-desc">Consultez l'<b>Aperçu du document PDF</b> (bouton <b>⬇️</b> intégré au lecteur) ou retrouvez tous vos mémoires dans l'onglet <b>Archives</b>.</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.toggle("▶️ Charger le lecteur vidéo HD (6,5 Mo)", value=False, key="toggle_hero_demo_video"):
                st.video(str(_feat_main), autoplay=True, loop=True, muted=True)

    ALLOWED_DOC_TYPES = ["pdf", "docx", "xlsx", "xls", "csv"]

    mode_tab1, mode_tab2 = st.tabs([
        "🧩 Dossier Complet Croisé (5 Catégories DCE + Entreprise)",
        "📚 Génération en Lot (Plusieurs marchés d'un coup)",
    ])

    with mode_tab1:
        # Bandeau des 4 engagements méthodologiques
        st.markdown(
            """
            <div class="mb-steps-grid" style="grid-template-columns: repeat(4, 1fr); margin-top:4px; margin-bottom:18px;">
                <div class="mb-step-card" style="border-top:3px solid #3B82F6;">
                    <div class="mb-step-title" style="color:#3B82F6;">1 · Ordre exact du RC</div>
                    <div class="mb-step-desc">Sommaire calqué sur l'ordre des critères et sous-critères de la <i>Valeur Technique</i> et respect du nombre de pages max.</div>
                </div>
                <div class="mb-step-card" style="border-top:3px solid #2563EB;">
                    <div class="mb-step-title" style="color:#2563EB;">2 · Fusion Multi-CCTP</div>
                    <div class="mb-step-desc">Combinaison du CCTP Général (Lot 00) avec les prescriptions techniques, matériaux et DTU de chaque lot.</div>
                </div>
                <div class="mb-step-card" style="border-top:3px solid #F59E0B;">
                    <div class="mb-step-title" style="color:#F59E0B;">3 · Croisement CCTP × DPGF</div>
                    <div class="mb-step-desc">Zéro phrase vague : chaque tâche cite la quantité exacte de la DPGF (<code>m²</code>, <code>m³</code>, <code>ml</code>) et les moyens dédiés.</div>
                </div>
                <div class="mb-step-card" style="border-top:3px solid #10B981;">
                    <div class="mb-step-title" style="color:#10B981;">4 · Garantie Anti-Invention</div>
                    <div class="mb-step-desc">Aucune donnée inventée : toute information absente des pièces est signalée par <code>[À COMPLÉTER : ...]</code>.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ── ZONE 1 & ZONE 2 ─────────────────────────────────────────────────
        col_z1, col_z2 = st.columns(2, gap="medium")
        with col_z1:
            st.markdown(
                """
                <div class="mb-zone-header" style="border-left:4px solid #3B82F6;">
                    <div class="mb-zone-top">
                        <span class="mb-zone-num" style="background:rgba(59,130,246,0.15); color:#3B82F6;">CATÉGORIE 1 · 1 FICHIER</span>
                        <span style="font-size:11.5px; color:var(--text-muted); font-weight:600;">PDF · Word · Excel</span>
                    </div>
                    <div class="mb-zone-title">📋 Le Règlement de Consultation (RC)</div>
                    <div class="mb-zone-desc">
                        Permet de structurer le mémoire exactement selon les attentes de l'acheteur public ou privé.
                    </div>
                    <div class="mb-tag-row">
                        <span class="mb-tag">🎯 Critères Valeur Technique</span>
                        <span class="mb-tag">⚖️ Barème de points</span>
                        <span class="mb-tag">📑 Plan imposé</span>
                        <span class="mb-tag">📏 Limite de pages</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            rc_file = st.file_uploader(
                "Déposer le Règlement de Consultation (RC)",
                type=ALLOWED_DOC_TYPES,
                key="upload_rc",
                accept_multiple_files=False,
                help="Format accepté : PDF, Word (.docx) ou Excel (.xlsx, .xls, .csv) — 1 fichier.",
            )

        with col_z2:
            st.markdown(
                """
                <div class="mb-zone-header" style="border-left:4px solid #2563EB;">
                    <div class="mb-zone-top">
                        <span class="mb-zone-num" style="background:rgba(37,99,235,0.15); color:#2563EB;">CATÉGORIE 2 · UPLOAD MULTIPLE</span>
                        <span style="font-size:11.5px; color:var(--text-muted); font-weight:600;">PDF · Word · Excel</span>
                    </div>
                    <div class="mb-zone-title">📐 Les CCTP (Lot 00 Général + Lots Techniques)</div>
                    <div class="mb-zone-desc">
                        Déposez le CCTP Général (Lot 00) ainsi qu'un ou plusieurs CCTP de lots techniques (Gros Œuvre, VRD, Électricité…).
                    </div>
                    <div class="mb-tag-row">
                        <span class="mb-tag">📍 Adresse & Accès</span>
                        <span class="mb-tag">🕒 Horaires & Site occupé</span>
                        <span class="mb-tag">🧱 Matériaux & DTU</span>
                        <span class="mb-tag">♻️ Déchets SOSED</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            cctp_multi_files = st.file_uploader(
                "Déposer le CCTP Général (Lot 00) et/ou vos CCTP de lots techniques",
                type=ALLOWED_DOC_TYPES,
                key="upload_cctp_multi",
                accept_multiple_files=True,
                help="Vous pouvez déposer simultanément le CCTP Général (Lot 00) et plusieurs CCTP de lots techniques.",
            )

        st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

        # ── ZONE 3 & ZONE 4 ─────────────────────────────────────────────────
        col_z3, col_z4 = st.columns(2, gap="medium")
        with col_z3:
            st.markdown(
                """
                <div class="mb-zone-header" style="border-left:4px solid #F59E0B;">
                    <div class="mb-zone-top">
                        <span class="mb-zone-num" style="background:rgba(245,158,11,0.15); color:#F59E0B;">CATÉGORIE 3 · UPLOAD MULTIPLE</span>
                        <span style="font-size:11.5px; color:var(--text-muted); font-weight:600;">Excel · PDF · Word</span>
                    </div>
                    <div class="mb-zone-title">📊 La DPGF / DQE (Devis Quantitatif)</div>
                    <div class="mb-zone-desc">
                        Extrait les métrés exacts du chantier pour justifier scientifiquement vos cadences, durées et effectifs.
                    </div>
                    <div class="mb-tag-row">
                        <span class="mb-tag">📏 Quantités (m², m³, ml, u)</span>
                        <span class="mb-tag">⏱️ Cadences & Rendements</span>
                        <span class="mb-tag">📅 Dimensionnement planning</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            dpgf_files = st.file_uploader(
                "Déposer la DPGF / le DQE (fichiers Excel, PDF ou Word)",
                type=ALLOWED_DOC_TYPES,
                key="upload_dpgf",
                accept_multiple_files=True,
                help="Déposez un ou plusieurs fichiers DPGF / DQE (tableaux Excel .xlsx/.xls/.csv, PDF ou Word).",
            )

        with col_z4:
            st.markdown(
                """
                <div class="mb-zone-header" style="border-left:4px solid #8B5CF6;">
                    <div class="mb-zone-top">
                        <span class="mb-zone-num" style="background:rgba(139,92,246,0.15); color:#8B5CF6;">CATÉGORIE 4 · OPTIONNEL MULTIPLE</span>
                        <span style="font-size:11.5px; color:var(--text-muted); font-weight:600;">PDF · Word · Excel</span>
                    </div>
                    <div class="mb-zone-title">🦺 Les Annexes Chantier (PGC, Plan de masse, Sol)</div>
                    <div class="mb-zone-desc">
                        Personnalise la sécurité (PPSPS/PGC) et le plan d'installation de chantier (base vie, grue, stockage).
                    </div>
                    <div class="mb-tag-row">
                        <span class="mb-tag">⚠️ Risques spécifiques PGC</span>
                        <span class="mb-tag">🤝 Coactivité</span>
                        <span class="mb-tag">🏗️ Base vie & Stockage</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            annexes_files = st.file_uploader(
                "Déposer les Annexes Chantier (PGC Sécurité, Plan de masse, Étude de sol…)",
                type=ALLOWED_DOC_TYPES,
                key="upload_annexes",
                accept_multiple_files=True,
                help="Optionnel : ajoutez le PGC, le plan d'installation de chantier, l'étude géotechnique, etc.",
            )

        st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

        # ── ZONE 5 : DOCUMENTS DE L'ENTREPRISE ──────────────────────────────
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #10B981;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num" style="background:rgba(16,185,129,0.15); color:#10B981;">CATÉGORIE 5 · UPLOAD MULTIPLE</span>
                    <span style="font-size:11.5px; color:var(--text-muted); font-weight:600;">PDF · Word · Excel</span>
                </div>
                <div class="mb-zone-title">🏢 Les Documents de l'Entreprise (Ancien mémoire, Personnel, Matériel, Certifications)</div>
                <div class="mb-zone-desc">
                    Le système y puise vos ressources réelles pour associer à chaque tâche du CCTP votre personnel habilité, vos engins et vos filières de tri.
                </div>
                <div class="mb-tag-row">
                    <span class="mb-tag">👥 Organigramme & CV</span>
                    <span class="mb-tag">🚜 Engins, Camions & Outils</span>
                    <span class="mb-tag">🪪 CACES, AIPR, Habilitations</span>
                    <span class="mb-tag">🏅 Qualibat, RGE, FNTP</span>
                    <span class="mb-tag">🚚 Fournisseurs & Déchèteries</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        entreprise_files = st.file_uploader(
            "Déposer les documents de votre entreprise (ancien mémoire technique, liste personnel, parc matériel, attestations Qualibat/RGE/CACES…)",
            type=ALLOWED_DOC_TYPES,
            key="upload_entreprise_docs",
            accept_multiple_files=True,
            help="Déposez tous vos justificatifs et documents internes (PDF, Word ou Excel). Ils seront croisés avec le CCTP et la DPGF.",
        )

        # ── TABLEAU DE BORD VISUEL DES PIÈCES PRÊTES À CROISER ──────────────
        nb_rc = 1 if rc_file else 0
        nb_cctp_m = len(cctp_multi_files) if cctp_multi_files else 0
        nb_dpgf = len(dpgf_files) if dpgf_files else 0
        nb_ann = len(annexes_files) if annexes_files else 0
        nb_ent_docs = len(entreprise_files) if entreprise_files else 0
        total_uploaded = nb_rc + nb_cctp_m + nb_dpgf + nb_ann + nb_ent_docs

        st.markdown(
            f"""
            <div class="mb-summary-bar">
                <div class="mb-summary-pill {'active' if nb_rc > 0 else ''}">
                    <div class="mb-summary-label">1. Règlement (RC)</div>
                    <div class="mb-summary-val">{'✅ 1 fichier' if nb_rc > 0 else '⚪ 0 fichier'}</div>
                </div>
                <div class="mb-summary-pill {'active' if nb_cctp_m > 0 else ''}">
                    <div class="mb-summary-label">2. CCTP (Lot 00 + Lots)</div>
                    <div class="mb-summary-val">{f'✅ {nb_cctp_m} fichier(s)' if nb_cctp_m > 0 else '⚪ 0 fichier'}</div>
                </div>
                <div class="mb-summary-pill {'active' if nb_dpgf > 0 else ''}">
                    <div class="mb-summary-label">3. DPGF / DQE</div>
                    <div class="mb-summary-val">{f'✅ {nb_dpgf} fichier(s)' if nb_dpgf > 0 else '⚪ 0 fichier'}</div>
                </div>
                <div class="mb-summary-pill {'active' if nb_ann > 0 else ''}">
                    <div class="mb-summary-label">4. Annexes Chantier</div>
                    <div class="mb-summary-val">{f'✅ {nb_ann} fichier(s)' if nb_ann > 0 else '⚪ Optionnel'}</div>
                </div>
                <div class="mb-summary-pill {'active' if (nb_ent_docs > 0 or bool(nom_ent)) else ''}">
                    <div class="mb-summary-label">5. Docs Entreprise</div>
                    <div class="mb-summary-val">{f'✅ {nb_ent_docs} doc(s)' + (' + Profil' if nom_ent else '') if (nb_ent_docs > 0 or nom_ent) else '⚪ 0 fichier'}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        generate_crossed_btn = st.button(
            f"🚀 Croiser les {total_uploaded} document(s) & Générer le Mémoire Technique Complet"
            if total_uploaded > 0
            else "🚀 Croiser les documents & Générer le Mémoire Technique Complet",
            type="primary",
            use_container_width=True,
            key="btn_generate_crossed",
        )

        if generate_crossed_btn:
            api_key = config["api_key"].strip()
            llm_provider = config["llm_provider"]

            if not api_key:
                st.error("❌ Clé d'authentification système manquante. Allez dans **Paramètres** pour la configurer.")
                st.stop()

            if not cctp_multi_files and not rc_file and not dpgf_files:
                st.error("❌ Veuillez uploader au moins un **CCTP** (ou un RC / DPGF) pour lancer la génération.")
                st.stop()

            profil_text = build_profil_text(config["entreprise"])
            if profil_text:
                st.info(f"🏢 Profil enregistré **{config['entreprise'].get('nom', '')}** combiné aux documents uploadés.")
            elif nb_ent_docs > 0:
                st.info(f"🏢 **{nb_ent_docs} document(s) d'entreprise** utilisé(s) pour personnaliser les moyens humains, matériels et certifications.")
            else:
                st.warning(
                    "ℹ️ Aucun profil ni document d'entreprise fourni : conformément à la règle anti-invention, "
                    "les informations propres à votre entreprise seront indiquées par `[À COMPLÉTER : préciser ici...]`."
                )

            categories_input = [
                ("1. RÈGLEMENT DE CONSULTATION (RC)", [rc_file] if rc_file else []),
                ("2. CAHIERS DES CLAUSES TECHNIQUES PARTICULIÈRES (CCTP GÉNÉRAL + LOTS)", cctp_multi_files or []),
                ("3. DPGF / DQE (QUANTITÉS DU MARCHÉ)", dpgf_files or []),
                ("4. ANNEXES CHANTIER (PGC SÉCURITÉ, PLAN DE MASSE, ÉTUDE DE SOL)", annexes_files or []),
                ("5. DOCUMENTS DE L'ENTREPRISE (ANCIEN MÉMOIRE, PERSONNEL, MATÉRIEL, CERTIFICATIONS)", entreprise_files or []),
            ]

            combined_sections: list[str] = []
            extraits_par_categorie: dict[str, str] = {}
            source_filenames: list[str] = []

            with st.status("📖 Extraction multi-formats (PDF, Word, Excel) des 5 catégories de documents…", expanded=True) as status:
                try:
                    for cat_title, file_list in categories_input:
                        if not file_list:
                            st.write(f"⚪ **{cat_title}** : *Aucun fichier déposé*")
                            continue
                        cat_text, cat_logs = extract_category_documents(file_list, cat_title)
                        for f in file_list:
                            if f is not None:
                                source_filenames.append(f.name)
                        for log_line in cat_logs:
                            st.write(f"✅ {cat_title} → {log_line}")
                        if cat_text:
                            combined_sections.append(cat_text)
                            extraits_par_categorie[cat_title] = cat_text

                    if not combined_sections:
                        raise ValueError("Aucun texte exploitable n'a pu être extrait des documents fournis.")

                    full_dossier_text = "\n\n".join(combined_sections)
                    status.update(
                        label=f"✅ Extraction terminée ({len(source_filenames)} fichier(s), {len(full_dossier_text):,} caractères extraits)",
                        state="complete",
                    )
                except ValueError as e:
                    status.update(label="❌ Erreur lors de l'extraction des documents", state="error")
                    st.error(str(e))
                    st.stop()

            # ── Tâche A : Analyse croisée ───────────────────────────────────
            with st.status("🔍 Analyse et croisement sémantique (RC × CCTP × DPGF × Annexes × Entreprise)…", expanded=True) as status:
                try:
                    nb_chunks = len(chunk_text(full_dossier_text))
                    if nb_chunks > 1:
                        st.write(f"Dossier complet volumineux → analyse structurée en {nb_chunks} segments.")
                    st.write("Extraction des critères du RC, des prescriptions CCTP, des quantités DPGF, du PGC et des moyens de l'entreprise…")
                    analyse = analyse_cctp(api_key, llm_provider, full_dossier_text)
                    st.write("✅ Croisement JSON terminé.")
                    status.update(label="✅ Analyse croisée des documents terminée", state="complete")
                except ValueError as e:
                    status.update(label="❌ Erreur d'analyse", state="error")
                    st.error(str(e))
                    st.stop()

            # ── Tâche B : Rédaction du Mémoire Technique ────────────────────
            with st.status("✍️ Rédaction du Mémoire Technique sur-mesure (ordre du RC, quantités DPGF, règle anti-invention)…", expanded=True) as status:
                try:
                    st.write("Rédaction détaillée phase par phase en croisant CCTP + DPGF + Moyens Entreprise…")
                    memoire_md = rediger_memoire(
                        api_key,
                        llm_provider,
                        analyse,
                        profil_text,
                        extraits_par_categorie=extraits_par_categorie,
                    )
                    st.write("✅ Rédaction du Mémoire Technique terminée.")
                    status.update(label="✅ Mémoire Technique rédigé avec succès", state="complete")
                except ValueError as e:
                    status.update(label="❌ Erreur de rédaction", state="error")
                    st.error(str(e))
                    st.stop()

            # ── Compilation Word et PDF ─────────────────────────────────────
            with st.status("📄 Génération des livrables PDF & Word (.docx)…", expanded=True) as status:
                try:
                    docx_buffer = generate_docx(memoire_md, analyse, config["entreprise"])
                    pdf_buffer = generate_pdf(memoire_md, analyse, config["entreprise"])
                    status.update(label="✅ Fichiers PDF et Word générés avec succès", state="complete")
                except Exception as e:
                    status.update(label="❌ Erreur de formatage PDF/Word", state="error")
                    st.error(f"Erreur lors de la génération des fichiers : {e}")
                    st.stop()

            default_stem = Path(source_filenames[0]).stem if source_filenames else "Dossier_BTP"
            nom_projet = str(analyse.get("nom_projet", "")).strip()
            if not nom_projet or nom_projet in ("Non spécifié", "Non spécifié dans cette section"):
                nom_projet = default_stem
            nom_fichier = re.sub(r"[^\w\s-]", "", nom_projet).strip().replace(" ", "_")[:60] or "Memoire_BTP"

            os.makedirs("archives", exist_ok=True)
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_bytes = pdf_buffer.getvalue()
            docx_bytes = docx_buffer.getvalue()

            pdf_path_saved = f"archives/Memoire_{nom_fichier}_{date_str}.pdf"
            docx_path_saved = f"archives/Memoire_{nom_fichier}_{date_str}.docx"
            with open(pdf_path_saved, "wb") as f:
                f.write(pdf_bytes)
            with open(docx_path_saved, "wb") as f:
                f.write(docx_bytes)

            ea.save_project_to_db({
                "project_id": f"site_{date_str}",
                "created_at": datetime.now().strftime("%d/%m/%Y à %H:%M"),
                "origin": "site",
                "badge": "Créé sur le site",
                "nom_projet": nom_projet,
                "nom_fichier": nom_fichier,
                "source_name": ", ".join(source_filenames),
                "analyse": analyse,
                "memoire_md": memoire_md,
                "entreprise": config["entreprise"],
                "pdf_path": pdf_path_saved,
                "docx_path": docx_path_saved,
                "vigilance_points": ea.extract_five_vigilance_points(analyse),
                "balises_a_completer": list(dict.fromkeys(re.findall(r"\[À COMPLÉTER[^\]]*\]", memoire_md))),
            })

            st.session_state["generated_results"] = [{
                "source_name": ", ".join(source_filenames),
                "nom_projet": nom_projet,
                "nom_fichier": nom_fichier,
                "analyse": analyse,
                "memoire_md": memoire_md,
                "pdf_bytes": pdf_bytes,
                "docx_bytes": docx_bytes,
            }]

    with mode_tab2:
        st.subheader("📚 Générer plusieurs Mémoires Techniques en lot (1 fichier CCTP = 1 Mémoire)")
        st.caption(
            "Utilisez ce mode si vous souhaitez traiter plusieurs appels d'offres distincts à la suite en quelques clics."
        )
        cctp_files = st.file_uploader(
            "Uploadez plusieurs CCTP (PDF, Word ou Excel) pour générer un mémoire technique par fichier",
            type=ALLOWED_DOC_TYPES,
            key="cctp_batch",
            accept_multiple_files=True,
            help="Glissez-déposez plusieurs CCTP simultanément pour les traiter en lot.",
        )

        nb_files = len(cctp_files) if cctp_files else 0
        if nb_files > 1:
            st.info(f"📚 **{nb_files} fichiers sélectionnés** : cliquez sur le bouton ci-dessous pour générer les {nb_files} mémoires techniques à la suite automatiquement.")

        btn_label = (
            f"🚀 Générer les {nb_files} Mémoires Techniques en lot"
            if nb_files > 1
            else "🚀 Générer le Mémoire Technique"
        )
        generate_btn = st.button(
            btn_label,
            type="primary",
            use_container_width=True,
            key="btn_generate_batch",
        )

        if generate_btn:
            api_key = config["api_key"].strip()
            llm_provider = config["llm_provider"]

            if not api_key:
                st.error("❌ Clé d'authentification système manquante. Allez dans **Paramètres** pour la configurer.")
                st.stop()

            if not cctp_files:
                st.error("❌ Veuillez uploader au moins un fichier CCTP.")
                st.stop()

            profil_text = build_profil_text(config["entreprise"])
            batch_results = []
            os.makedirs("archives", exist_ok=True)
            batch_progress = st.progress(0.0, text=f"Démarrage du traitement (0/{nb_files})…") if nb_files > 1 else None

            for idx, cctp_file in enumerate(cctp_files, start=1):
                prefix = f"[{idx}/{nb_files}] {cctp_file.name} — " if nb_files > 1 else ""
                if batch_progress:
                    batch_progress.progress((idx - 1) / nb_files, text=f"Traitement du dossier {idx}/{nb_files} : {cctp_file.name}…")

                with st.status(f"📖 {prefix}Extraction du texte…", expanded=(nb_files == 1)) as status:
                    try:
                        cctp_text = extract_text_from_document(cctp_file)
                        status.update(label=f"✅ {prefix}Extraction terminée ({len(cctp_text):,} caractères)", state="complete")
                    except ValueError as e:
                        status.update(label=f"❌ {prefix}Erreur d'extraction", state="error")
                        st.error(f"{cctp_file.name} : {e}")
                        continue

                with st.status(f"🔍 {prefix}Analyse sémantique en cours…", expanded=(nb_files == 1)) as status:
                    try:
                        analyse = analyse_cctp(api_key, llm_provider, cctp_text)
                        status.update(label=f"✅ {prefix}Analyse terminée", state="complete")
                    except ValueError as e:
                        status.update(label=f"❌ {prefix}Erreur d'analyse", state="error")
                        st.error(f"{cctp_file.name} : {e}")
                        continue

                with st.status(f"✍️ {prefix}Rédaction du Mémoire Technique…", expanded=(nb_files == 1)) as status:
                    try:
                        memoire_md = rediger_memoire(
                            api_key,
                            llm_provider,
                            analyse,
                            profil_text,
                            extraits_par_categorie={"CCTP": cctp_text},
                        )
                        status.update(label=f"✅ {prefix}Mémoire Technique rédigé", state="complete")
                    except ValueError as e:
                        status.update(label=f"❌ {prefix}Erreur de rédaction", state="error")
                        st.error(f"{cctp_file.name} : {e}")
                        continue

                with st.status(f"📄 {prefix}Compilation PDF & Word…", expanded=(nb_files == 1)) as status:
                    try:
                        docx_buffer = generate_docx(memoire_md, analyse, config["entreprise"])
                        pdf_buffer = generate_pdf(memoire_md, analyse, config["entreprise"])
                        status.update(label=f"✅ {prefix}Document généré avec succès", state="complete")
                    except Exception as e:
                        status.update(label=f"❌ {prefix}Erreur de formatage", state="error")
                        st.error(f"Erreur lors de la génération du document ({cctp_file.name}) : {e}")
                        continue

                nom_projet = str(analyse.get("nom_projet", "")).strip()
                if not nom_projet or nom_projet in ("Non spécifié", "Non spécifié dans cette section"):
                    nom_projet = Path(cctp_file.name).stem
                nom_fichier = re.sub(r"[^\w\s-]", "", nom_projet).strip().replace(" ", "_")[:60] or f"Memoire_{idx}"

                date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                pdf_bytes = pdf_buffer.getvalue()
                docx_bytes = docx_buffer.getvalue()

                pdf_batch_path = f"archives/Memoire_{nom_fichier}_{date_str}.pdf"
                docx_batch_path = f"archives/Memoire_{nom_fichier}_{date_str}.docx"
                with open(pdf_batch_path, "wb") as f:
                    f.write(pdf_bytes)
                with open(docx_batch_path, "wb") as f:
                    f.write(docx_bytes)

                ea.save_project_to_db({
                    "project_id": f"batch_{date_str}_{idx}",
                    "created_at": datetime.now().strftime("%d/%m/%Y à %H:%M"),
                    "origin": "site",
                    "badge": "Créé sur le site",
                    "nom_projet": nom_projet,
                    "nom_fichier": nom_fichier,
                    "source_name": cctp_file.name,
                    "analyse": analyse,
                    "memoire_md": memoire_md,
                    "entreprise": config["entreprise"],
                    "pdf_path": pdf_batch_path,
                    "docx_path": docx_batch_path,
                    "vigilance_points": ea.extract_five_vigilance_points(analyse),
                    "balises_a_completer": list(dict.fromkeys(re.findall(r"\[À COMPLÉTER[^\]]*\]", memoire_md))),
                })

                batch_results.append({
                    "source_name": cctp_file.name,
                    "nom_projet": nom_projet,
                    "nom_fichier": nom_fichier,
                    "analyse": analyse,
                    "memoire_md": memoire_md,
                    "pdf_bytes": pdf_bytes,
                    "docx_bytes": docx_bytes,
                })

            if batch_progress:
                batch_progress.progress(1.0, text=f"✅ {len(batch_results)}/{nb_files} Mémoire(s) Technique(s) généré(s) avec succès !")

            if batch_results:
                st.session_state["generated_results"] = batch_results

    # ── Affichage des résultats générés (persistant au clic de téléchargement) ──
    results = st.session_state.get("generated_results", [])
    if results:
        st.divider()
        if len(results) == 1:
            res = results[0]
            st.success(f"🎉 **Mémoire Technique généré avec succès** (Pièces croisées : `{res.get('source_name', '')}`).")

            # Détection des balises anti-invention [À COMPLÉTER : ...]
            marqueurs_a_completer = re.findall(r"\[À COMPLÉTER[^\]]*\]", res["memoire_md"])
            if marqueurs_a_completer:
                uniques_marqueurs = list(dict.fromkeys(marqueurs_a_completer))
                with st.expander(
                    f"⚠️ {len(uniques_marqueurs)} information(s) à compléter identifiée(s) (Règle Anti-Invention)",
                    expanded=True,
                ):
                    st.caption(
                        "Ces éléments n'étaient présents dans aucun des documents uploadés. "
                        "Ils apparaissent en rouge dans le fichier Word (.docx) pour que vous puissiez les compléter facilement :"
                    )
                    for m in uniques_marqueurs:
                        st.markdown(f"- `{m}`")

            with st.expander("📊 Voir l'analyse croisée extraite des documents (RC, CCTP, DPGF, PGC, Entreprise)", expanded=False):
                st.json(res["analyse"])

            with st.expander("📝 Aperçu du Mémoire Technique (Markdown)", expanded=True):
                st.markdown(res["memoire_md"])

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Télécharger en PDF",
                    data=res["pdf_bytes"],
                    file_name=f"Memoire_Technique_{res['nom_fichier']}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True,
                    key="dl_single_pdf",
                )
            with col2:
                st.download_button(
                    label="📝 Télécharger en Word (.docx)",
                    data=res["docx_bytes"],
                    file_name=f"Memoire_Technique_{res['nom_fichier']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_single_docx",
                )

            st.caption("💡 Le format Word (.docx) met en évidence en rouge gras les balises `[À COMPLÉTER : ...]` et vous permet d'ajouter votre logo ou vos photos.")

            st.subheader("👁️ Aperçu du document PDF")
            b64_pdf = base64.b64encode(res["pdf_bytes"]).decode("utf-8")
            pdf_display = f'<iframe src="data:application/pdf;base64,{b64_pdf}#view=FitH" width="100%" height="800" type="application/pdf"></iframe>'
            st.markdown(pdf_display, unsafe_allow_html=True)

        else:
            st.success(f"🎉 **{len(results)} Mémoires Techniques générés avec succès !** Téléchargez-les tous en un clic ou individuellement ci-dessous.")

            # Création des archives ZIP (Tous les PDF + Pack Complet PDF & Word)
            zip_pdf_buf = io.BytesIO()
            zip_all_buf = io.BytesIO()
            with zipfile.ZipFile(zip_pdf_buf, "w", zipfile.ZIP_DEFLATED) as zpdf, \
                 zipfile.ZipFile(zip_all_buf, "w", zipfile.ZIP_DEFLATED) as zall:
                for i, r in enumerate(results, start=1):
                    pdf_name = f"{i:02d}_Memoire_Technique_{r['nom_fichier']}.pdf"
                    docx_name = f"{i:02d}_Memoire_Technique_{r['nom_fichier']}.docx"
                    zpdf.writestr(pdf_name, r["pdf_bytes"])
                    zall.writestr(pdf_name, r["pdf_bytes"])
                    zall.writestr(docx_name, r["docx_bytes"])

            zcol1, zcol2 = st.columns(2)
            with zcol1:
                st.download_button(
                    label=f"📦 Télécharger les {len(results)} PDF d'un coup (.ZIP)",
                    data=zip_pdf_buf.getvalue(),
                    file_name=f"Lot_{len(results)}_Memoires_Techniques_PDF_{datetime.now().strftime('%Y%m%d')}.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                    key="dl_batch_zip_pdf",
                )
            with zcol2:
                st.download_button(
                    label=f"📦 Télécharger le pack complet PDF + Word (.ZIP)",
                    data=zip_all_buf.getvalue(),
                    file_name=f"Lot_{len(results)}_Memoires_Techniques_Complet_{datetime.now().strftime('%Y%m%d')}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="dl_batch_zip_all",
                )

            st.markdown("### 📋 Détail par appel d'offres")
            tab_labels = [f"📄 {i}. {r['nom_projet'][:32]}" for i, r in enumerate(results, start=1)]
            tabs = st.tabs(tab_labels)
            for i, (tab, res) in enumerate(zip(tabs, results), start=1):
                with tab:
                    st.markdown(f"**Projet :** {res['nom_projet']}  \n**Fichier CCTP source :** `{res['source_name']}`")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            label="📥 Télécharger en PDF",
                            data=res["pdf_bytes"],
                            file_name=f"Memoire_Technique_{res['nom_fichier']}.pdf",
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True,
                            key=f"dl_batch_pdf_{i}",
                        )
                    with c2:
                        st.download_button(
                            label="📝 Télécharger en Word (.docx)",
                            data=res["docx_bytes"],
                            file_name=f"Memoire_Technique_{res['nom_fichier']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key=f"dl_batch_docx_{i}",
                        )
                    with st.expander("📝 Voir le texte rédigé (Markdown)", expanded=False):
                        st.markdown(res["memoire_md"])
                    if st.toggle(f"👁️ Afficher l'aperçu PDF intégré ({res['nom_projet'][:28]})", value=False, key=f"batch_pdf_prev_{i}"):
                        b64_pdf = base64.b64encode(res["pdf_bytes"]).decode("utf-8")
                        st.markdown(
                            f'<iframe src="data:application/pdf;base64,{b64_pdf}#view=FitH" width="100%" height="700" type="application/pdf"></iframe>',
                            unsafe_allow_html=True,
                        )

    # ── Synchronisation en temps réel des projets "Automatisé via Mail" sur la page principale ──
    email_synced_projects = [p for p in ea.load_projects_db() if p.get("origin") == "email"]
    if email_synced_projects:
        st.divider()
        st.subheader(f"📧 Derniers Mémoires Techniques Automatisés via Mail ({len(email_synced_projects)})")
        st.caption(
            "Ces mémoires ont été générés automatiquement à la réception d'un e-mail. "
            "Sélectionnez un dossier ci-dessous pour l'ouvrir, le lire, le modifier (avec régénération immédiate du `.docx` et du `.pdf`) ou le télécharger."
        )
        selected_em_idx = st.selectbox(
            "Sélectionner un mémoire automatisé par mail :",
            options=list(range(min(len(email_synced_projects), 10))),
            format_func=lambda idx: f"📧 [Automatisé via Mail] {email_synced_projects[idx].get('nom_projet', 'Mémoire')} — de {email_synced_projects[idx].get('sender_email', '')} ({email_synced_projects[idx].get('created_at', '')})",
            key="gen_sync_select_idx",
        )
        em_proj = email_synced_projects[selected_em_idx]
        with st.expander(
            f"📧 Ouvrir le dossier : {em_proj.get('nom_projet', 'Mémoire')} ({em_proj.get('created_at', '')})",
            expanded=(not results),
        ):
            render_interactive_project_viewer(em_proj, key_prefix=f"gen_sync_{selected_em_idx}")


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 🤖 AUTOMATISATION (PIPELINE EMAIL -> SITE WEB -> EMAIL)
# ─────────────────────────────────────────────────────────────────────────
elif page == "Automatisation":
    auto_state = ea.load_automation_state()
    watcher_active = bool(auto_state.get("watcher_enabled", False))
    current_mode = auto_state.get("automation_mode", "auto")
    jobs_list = auto_state.get("jobs", [])
    logs_list = auto_state.get("logs", [])
    pending_manual_jobs = [j for j in jobs_list if j.get("statut") == "En attente de validation manuelle"]
    sent_jobs_count = sum(1 for j in jobs_list if j.get("statut") == "Mémoire envoyé")
    mail_projects_count = sum(1 for p in ea.load_projects_db() if p.get("origin") == "email")

    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">🤖 PIPELINE AUTOMATISÉ · EMAIL ➔ TRI IA BTP ➔ SITE WEB ➔ EMAIL</div>
            <div class="mb-hero-title">Panneau de Contrôle · Automatisation des Mails</div>
            <p class="mb-hero-sub">
                Surveillance automatique de votre boîte mail (IMAP), décompression des archives <code>.zip</code> DCE, tri intelligent des pièces (RC, CCTP Lot 00 + Lot visé, DPGF, Annexes, Entreprise), génération en temps réel sur le site et réponse automatique par e-mail (SMTP) avec le mémoire <code>.docx</code> et <code>.pdf</code>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Bandeau KPI Exécutif
    st.markdown(
        f"""
        <div class="mb-kpi-grid">
            <div class="mb-kpi-card">
                <div class="mb-kpi-label">📩 Mails DCE Reçus</div>
                <div class="mb-kpi-val">{len(jobs_list)}</div>
                <div class="mb-kpi-sub">Dossiers analysés par l'Auto-Triage BTP</div>
            </div>
            <div class="mb-kpi-card">
                <div class="mb-kpi-label">📧 Mémoires sur le Site</div>
                <div class="mb-kpi-val" style="color:#3B82F6;">{mail_projects_count}</div>
                <div class="mb-kpi-sub">Synchronisés avec badge « Automatisé via Mail »</div>
            </div>
            <div class="mb-kpi-card">
                <div class="mb-kpi-label">⏳ En Attente Validation</div>
                <div class="mb-kpi-val" style="color:#F59E0B;">{len(pending_manual_jobs)}</div>
                <div class="mb-kpi-sub">Prêts à partir en 1 clic (Mode Manuel)</div>
            </div>
            <div class="mb-kpi-card">
                <div class="mb-kpi-label">📤 Réponses Envoyées</div>
                <div class="mb-kpi-val" style="color:#10B981;">{sent_jobs_count}</div>
                <div class="mb-kpi-sub">Livrables .docx + .pdf + synthèse vigilance</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── 1. BARRE DE CONTRÔLE PRINCIPALE : ACTIVER/PAUSE, MODE & SIMULATEUR ──
    ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([1.35, 1.55, 1.3], gap="medium")

    with ctrl_col1:
        status_color = "#10B981" if watcher_active else "#F59E0B"
        status_label = "🟢 SURVEILLANCE ACTIVE" if watcher_active else "⏸️ SURVEILLANCE EN PAUSE"
        st.markdown(
            f"""
            <div class="mb-zone-header" style="border-left:4px solid {status_color}; margin-bottom:10px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num" style="background:rgba(16,185,129,0.14); color:{status_color};">{status_label}</span>
                </div>
                <div class="mb-zone-title">1. Écoute Boîte Mail (Worker IMAP)</div>
                <div class="mb-zone-desc" style="margin-bottom:4px;">
                    Dernière vérification : <b>{auto_state.get('last_check') or 'En attente'}</b><br>
                    Fréquence : toutes les <b>{auto_state.get('poll_interval_seconds', 30)}s</b>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        btn_toggle_label = "⏸️ Mettre en pause la surveillance des mails" if watcher_active else "▶️ Activer la surveillance automatique des mails"
        if st.button(btn_toggle_label, type="primary" if not watcher_active else "secondary", use_container_width=True, key="btn_toggle_watcher"):
            auto_state["watcher_enabled"] = not watcher_active
            ea.save_automation_state(auto_state)
            ea.add_automation_log(
                "Surveillance activée" if not watcher_active else "Surveillance en pause",
                "L'écoute automatique IMAP en arrière-plan a été " + ("activée." if not watcher_active else "mise en pause."),
            )
            st.rerun()

        if st.button("🔄 Vérifier la boîte mail maintenant (IMAP)", use_container_width=True, key="btn_poll_imap_now"):
            with st.spinner("Connexion IMAP et recherche de nouveaux mails non lus…"):
                nb_proc, err_imap = ea.poll_imap_inbox_once(PIPELINE_CALLBACKS, load_config)
            if err_imap:
                st.warning(f"ℹ️ {err_imap}")
            elif nb_proc > 0:
                st.success(f"✅ {nb_proc} nouveau(x) mail(s) traité(s) avec succès !")
            else:
                st.info("📭 Aucun nouvel e-mail non lu dans la boîte de réception.")
            st.rerun()

    with ctrl_col2:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:10px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">MODE D'EXPÉDITION</span>
                </div>
                <div class="mb-zone-title">2. Sélecteur de Mode d'Automatisation</div>
                <div class="mb-zone-desc" style="margin-bottom:4px;">
                    Choisissez si l'e-mail de réponse part immédiatement ou après votre validation sur le site.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        mode_options = ["auto", "manual"]
        mode_labels = {
            "auto": "⚡ Mode 100 % Automatique (Génère sur le site & renvoie le mail tout seul)",
            "manual": "👁️ Mode Validation Manuelle (Génère sur le site & attend « Valider et Envoyer »)",
        }
        selected_mode = st.radio(
            "Mode d'automatisation",
            options=mode_options,
            index=0 if current_mode == "auto" else 1,
            format_func=lambda k: mode_labels[k],
            key="radio_automation_mode",
            label_visibility="collapsed",
        )
        if selected_mode != current_mode:
            auto_state["automation_mode"] = selected_mode
            ea.save_automation_state(auto_state)
            ea.add_automation_log(
                "Changement de mode",
                f"Mode d'automatisation basculé sur : {mode_labels[selected_mode]}",
            )
            st.rerun()

    with ctrl_col3:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #8B5CF6; margin-bottom:10px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num" style="background:rgba(139,92,246,0.15); color:#8B5CF6;">TEST EN 1 CLIC</span>
                </div>
                <div class="mb-zone-title">3. Simuler la réception d'un mail</div>
                <div class="mb-zone-desc" style="margin-bottom:4px;">
                    Testez tout le circuit (ZIP DCE multi-lots ➔ Tri CCTP ➔ Site ➔ Mail) directement depuis l'interface.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button(
            "🧪 Simuler la réception d'un mail (DCE .ZIP complet)",
            type="primary",
            use_container_width=True,
            key="btn_simulate_email_1click",
        ):
            with st.status("📨 Simulation de la réception d'un e-mail client BTP avec dossier DCE .ZIP…", expanded=True) as sim_status:
                st.write("1️⃣ **Mail reçu** de `contact@durand-tp-vrd.fr` — Objet : *Demande de mémoire technique - Groupe Scolaire Jean Moulin (Lot 02 VRD)*")
                zip_fname, zip_bytes = ea.build_realistic_dce_zip_simulation()
                st.write(f"2️⃣ **Décompression automatique du fichier `{zip_fname}`** et **Tri intelligent des CCTP** (Lot 00 Général + Lot 02 VRD retenus, Lot 01 Gros Œuvre écarté)…")
                sim_body = (
                    "Bonjour,\n\n"
                    "Veuillez trouver ci-joint le dossier DCE complet en .zip pour le marché de réhabilitation "
                    "du Groupe Scolaire Jean Moulin à Lyon 3e.\n"
                    "Nous répondons uniquement sur le Lot 02 VRD & Aménagements Extérieurs (cour d'école perméable).\n"
                    "Consignes spécifiques : bien mettre en avant notre équipe certifiée AIPR et CACES R482, "
                    "l'interdiction de livraison aux heures d'entrée/sortie des élèves (8h00-8h45 et 16h15-17h00), "
                    "et notre exutoire de recyclage Valomat Vénissieux à 5,5 km.\n\n"
                    "Bien cordialement,\nMarc Durand — DURAND TP & VRD"
                )
                st.write("3️⃣ **Génération du Mémoire Technique (.docx & .pdf)** et **Synchronisation sur le site** (badge `Automatisé via Mail`)…")
                job_res = ea.process_incoming_email_pipeline(
                    sender_email="contact@durand-tp-vrd.fr",
                    sender_name="Marc Durand (DURAND TP & VRD)",
                    subject="Demande de mémoire technique - Groupe Scolaire Jean Moulin (Lot 02 VRD)",
                    body_text=sim_body,
                    raw_attachments=[(zip_fname, zip_bytes)],
                    config=config,
                    extract_text_fn=extract_text_from_document,
                    extract_category_fn=extract_category_documents,
                    analyse_cctp_fn=analyse_cctp,
                    rediger_memoire_fn=rediger_memoire,
                    generate_docx_fn=generate_docx,
                    generate_pdf_fn=generate_pdf,
                    build_profil_text_fn=build_profil_text,
                    call_llm_fn=call_llm,
                    is_simulation=True,
                )
                sim_status.update(
                    label=f"✅ Circuit terminé — Statut : {job_res.get('statut', 'Terminé')} (Projet : {job_res.get('nom_projet', '')})",
                    state="complete",
                )
            st.rerun()

    # ── Volet optionnel : Simulation personnalisée ou test "Pièce manquante (sans CCTP)" ──
    with st.expander("🛠️ Personnaliser la simulation d'e-mail (avec vos propres fichiers .zip / .pdf / .docx / .xlsx ou tester le cas « Sans CCTP »)", expanded=False):
        sim_c1, sim_c2 = st.columns(2)
        with sim_c1:
            custom_sender = st.text_input("Adresse e-mail de l'expéditeur simulé", value="etudes@entreprise-btp.fr", key="sim_custom_sender")
            custom_subject = st.text_input("Objet de l'e-mail", value="DCE Appel d'offres - Lot 02 VRD", key="sim_custom_subject")
        with sim_c2:
            custom_body = st.text_area(
                "Corps du mail (précisez le lot visé et vos consignes spécifiques)",
                value="Bonjour, nous souhaitons répondre au Lot 02 VRD. Merci de croiser le RC, le CCTP Lot 00, le CCTP Lot 02 et notre DPGF.",
                height=105,
                key="sim_custom_body",
            )
        custom_files = st.file_uploader(
            "Pièces jointes du mail simulé (archive .zip DCE, ou fichiers PDF, Word, Excel en vrac — laissez vide pour utiliser le pack DCE .zip de démonstration)",
            type=["zip", "pdf", "docx", "xlsx", "xls", "csv", "txt"],
            accept_multiple_files=True,
            key="sim_custom_files",
        )
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            if st.button("🚀 Lancer la simulation avec ces paramètres", type="primary", use_container_width=True, key="btn_run_custom_sim"):
                if custom_files:
                    raw_att_list = [(f.name, f.getvalue()) for f in custom_files]
                else:
                    zname, zbytes = ea.build_realistic_dce_zip_simulation()
                    raw_att_list = [(zname, zbytes)]
                with st.spinner("Exécution du pipeline complet (Décompression ➔ Auto-Triage ➔ Génération ➔ Site ➔ Mail)…"):
                    ea.process_incoming_email_pipeline(
                        sender_email=custom_sender,
                        sender_name=custom_sender.split("@")[0],
                        subject=custom_subject,
                        body_text=custom_body,
                        raw_attachments=raw_att_list,
                        config=config,
                        extract_text_fn=extract_text_from_document,
                        extract_category_fn=extract_category_documents,
                        analyse_cctp_fn=analyse_cctp,
                        rediger_memoire_fn=rediger_memoire,
                        generate_docx_fn=generate_docx,
                        generate_pdf_fn=generate_pdf,
                        build_profil_text_fn=build_profil_text,
                        call_llm_fn=call_llm,
                        is_simulation=True,
                    )
                st.rerun()
        with b_col2:
            if st.button("⚠️ Tester le cas « Mail reçu sans CCTP (Pièce manquante) »", use_container_width=True, key="btn_sim_missing_cctp"):
                with st.spinner("Simulation d'un e-mail incomplet (sans CCTP)…"):
                    ea.process_incoming_email_pipeline(
                        sender_email=custom_sender,
                        sender_name=custom_sender.split("@")[0],
                        subject="Demande de mémoire technique (Dossier incomplet sans CCTP)",
                        body_text="Bonjour, pouvez-vous nous rédiger le mémoire technique pour ce chantier ?",
                        raw_attachments=[],
                        config=config,
                        extract_text_fn=extract_text_from_document,
                        extract_category_fn=extract_category_documents,
                        analyse_cctp_fn=analyse_cctp,
                        rediger_memoire_fn=rediger_memoire,
                        generate_docx_fn=generate_docx,
                        generate_pdf_fn=generate_pdf,
                        build_profil_text_fn=build_profil_text,
                        call_llm_fn=call_llm,
                        is_simulation=True,
                    )
                st.rerun()

    # ── 2. FILE D'ATTENTE DE VALIDATION MANUELLE (SI MODE VALIDATION MANUELLE) ──
    if pending_manual_jobs:
        st.divider()
        st.subheader(f"⏳ Mémoire(s) en attente de votre validation avant envoi du mail ({len(pending_manual_jobs)})")
        for pjob in pending_manual_jobs:
            jid = pjob["job_id"]
            st.markdown(
                f"""
                <div class="mb-zone-header" style="border-left:4px solid #F59E0B; margin-bottom:10px;">
                    <div class="mb-zone-top">
                        <span class="mb-zone-num" style="background:rgba(245,158,11,0.16); color:#F59E0B;">⏳ EN ATTENTE DE VALIDATION MANUELLE</span>
                        <span class="mb-zone-badge">Reçu le {pjob.get('received_at', '')}</span>
                    </div>
                    <div class="mb-zone-title">{pjob.get('nom_projet', pjob.get('subject', ''))}</div>
                    <div class="mb-zone-desc" style="margin-bottom:0;">
                        <b>Destinataire de la réponse :</b> <code>{pjob.get('sender_email', '')}</code> &nbsp;·&nbsp;
                        <b>Lot(s) traité(s) :</b> {', '.join(pjob.get('lots_vises') or ['Tous lots'])}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(
                f"✅ Valider et Envoyer le Mémoire (.docx + .pdf) par mail à {pjob.get('sender_email', '')}",
                type="primary",
                use_container_width=True,
                key=f"btn_approve_send_{jid}",
            ):
                ok_send, msg_send = ea.approve_and_send_manual_job(jid)
                if ok_send:
                    st.success(f"✅ {msg_send}")
                else:
                    st.error(f"❌ {msg_send}")
                st.rerun()

    st.divider()

    # ── 3. ONGLETS DU PANNEAU D'AUTOMATISATION : JOURNAL EN DIRECT, MÉMOIRES AUTOMATISÉS & CONFIG IMAP/SMTP ──
    auto_tab1, auto_tab2, auto_tab3 = st.tabs([
        f"📡 Journal en Direct (Logs & Statuts des Mails — {len(jobs_list)} dossier(s))",
        "📧 Mémoires « Automatisé via Mail » (Ouvrir, Lire, Modifier & Télécharger)",
        "⚙️ Configuration Boîte Mail (IMAP / SMTP / .env)",
    ])

    with auto_tab1:
        st.markdown("#### 📡 Suivi en direct des e-mails reçus et traités")
        if not jobs_list:
            st.info(
                "Aucun e-mail reçu pour le moment. Cliquez sur **« 🧪 Simuler la réception d'un mail (DCE .ZIP complet) »** "
                "ci-dessus pour tester immédiatement toute la chaîne d'automatisation !"
            )
        else:
            for idx_j, job in enumerate(jobs_list):
                st_str = job.get("statut", "Mail reçu")
                if st_str == "Mémoire envoyé":
                    st_col = "#10B981"
                    st_icon = "✅"
                elif st_str == "En attente de validation manuelle":
                    st_col = "#F59E0B"
                    st_icon = "⏳"
                elif "Pièce manquante" in st_str:
                    st_col = "#EF4444"
                    st_icon = "⚠️"
                else:
                    st_col = "#3B82F6"
                    st_icon = "🔄"

                with st.expander(
                    f"{st_icon} [{st_str}] {job.get('subject', 'Sans objet')} — de {job.get('sender_email', '')} ({job.get('received_at', '')})",
                    expanded=(idx_j == 0),
                ):
                    # Barre de progression visuelle des 4 étapes du pipeline
                    steps_html = f"""
                    <div class="mb-summary-bar" style="grid-template-columns: repeat(4, 1fr); margin-top:4px; margin-bottom:14px;">
                        <div class="mb-summary-pill active">
                            <div class="mb-summary-label">Étape 1</div>
                            <div class="mb-summary-val">📩 Mail reçu</div>
                        </div>
                        <div class="mb-summary-pill {'active' if st_str != 'Mail reçu' else ''}">
                            <div class="mb-summary-label">Étape 2</div>
                            <div class="mb-summary-val">🗂️ Tri des CCTP</div>
                        </div>
                        <div class="mb-summary-pill {'active' if st_str in ('Génération sur le site', 'En attente de validation manuelle', 'Mémoire envoyé') else ''}">
                            <div class="mb-summary-label">Étape 3</div>
                            <div class="mb-summary-val">⚙️ Génération sur le site</div>
                        </div>
                        <div class="mb-summary-pill {'active' if st_str == 'Mémoire envoyé' else ''}">
                            <div class="mb-summary-label">Étape 4</div>
                            <div class="mb-summary-val">{'⏳ Validation requise' if st_str == 'En attente de validation manuelle' else '📤 Mémoire envoyé'}</div>
                        </div>
                    </div>
                    """
                    st.markdown(steps_html, unsafe_allow_html=True)

                    st.markdown(
                        f"- **Expéditeur :** `{job.get('sender_email', '')}` ({job.get('entreprise_source', '')})\n"
                        f"- **Lot(s) ciblé(s) dans le mail :** `{', '.join(job.get('lots_vises') or ['Non spécifié'])}`\n"
                        f"- **Consignes extraites du mail :** {job.get('consignes_specifiques') or 'Aucune consigne particulière'}"
                    )

                    if job.get("classifications"):
                        st.markdown("##### 🗂️ Résultat du Tri Intelligent des Pièces Jointes (RC / CCTP / DPGF / Annexes / Entreprise)")
                        for c_item in job["classifications"]:
                            cat_name = c_item.get("categorie", "")
                            if cat_name == "IGNORED_CCTP":
                                st.markdown(f"- 🚫 **[ÉCARTÉ — HORS LOT VISÉ]** `{c_item['filename']}` — *{c_item.get('motif', '')}*")
                            else:
                                st.markdown(f"- ✅ **[{cat_name}]** `{c_item['filename']}` — *{c_item.get('motif', '')}*")

                    if job.get("vigilance_points"):
                        st.markdown("##### 🛡️ Synthèse des 5 Points de Vigilance Majeurs envoyée dans le mail")
                        for vp in job["vigilance_points"]:
                            st.markdown(f"- **{vp['titre']}** : {vp['detail']}")

                    if job.get("balises_a_completer") is not None and job.get("project_id"):
                        bal_list = job.get("balises_a_completer") or []
                        st.markdown(f"##### ⚠️ Balises `[À COMPLÉTER]` signalées dans le corps du mail ({len(bal_list)})")
                        for b in bal_list:
                            st.markdown(f"- `{b}`")

                    if job.get("email_body_plain_out"):
                        with st.expander("✉️ Aperçu de l'e-mail de réponse automatique envoyé à l'entreprise", expanded=False):
                            st.markdown(f"**Objet :** `{job.get('email_subject_out', '')}`")
                            if job.get("smtp_info"):
                                st.caption(f"Statut SMTP : {job['smtp_info']}")
                            st.code(job["email_body_plain_out"], language="text")

                    if job.get("statut") == "En attente de validation manuelle":
                        if st.button(
                            f"✅ Valider et Envoyer maintenant à {job.get('sender_email', '')}",
                            type="primary",
                            use_container_width=True,
                            key=f"log_approve_{job['job_id']}",
                        ):
                            ea.approve_and_send_manual_job(job["job_id"])
                            st.rerun()

        st.markdown("#### 📜 Historique chronologique des événements (Logs en direct)")
        if logs_list:
            for log_item in logs_list[:30]:
                st.markdown(
                    f"`{log_item.get('timestamp', '')}` · **[{log_item.get('statut', '')}]** — {log_item.get('message', '')}"
                )
        else:
            st.caption("Aucun événement enregistré.")

    with auto_tab2:
        all_projects = ea.load_projects_db()
        mail_projects = [p for p in all_projects if p.get("origin") == "email"]
        if not mail_projects:
            st.info("Aucun mémoire automatisé via mail pour le moment. Utilisez le bouton **« 🧪 Simuler la réception d'un mail »** ci-dessus pour en générer un.")
        else:
            sel_mp_idx = st.selectbox(
                "📂 Sélectionnez le mémoire automatisé à inspecter, modifier ou télécharger :",
                options=list(range(len(mail_projects))),
                format_func=lambda idx: f"📧 {mail_projects[idx].get('nom_projet', 'Mémoire')} — de {mail_projects[idx].get('sender_email', '')} ({mail_projects[idx].get('created_at', '')})",
                key="autotab_select_proj_idx",
            )
            render_interactive_project_viewer(mail_projects[sel_mp_idx], key_prefix=f"autotab_{sel_mp_idx}")

    with auto_tab3:
        st.markdown("#### ⚙️ Paramètres du Serveur Mail (IMAP & SMTP)")
        st.caption(
            "Ces paramètres peuvent être définis dans un fichier `.env` (`IMAP_HOST`, `IMAP_USER`, `IMAP_PASSWORD`, `SMTP_HOST`, `SMTP_USER`, `SMTP_PASSWORD`) "
            "ou directement dans le formulaire ci-dessous."
        )
        with st.form("form_imap_smtp_config"):
            c_im1, c_im2 = st.columns(2)
            with c_im1:
                st.markdown("**📥 Réception des mails (IMAP)**")
                new_imap_host = st.text_input("Serveur IMAP", value=auto_state.get("imap_host", "imap.gmail.com"))
                new_imap_port = st.number_input("Port IMAP SSL", value=int(auto_state.get("imap_port", 993)), min_value=1, max_value=65535)
                new_imap_user = st.text_input("Adresse e-mail surveillée (IMAP)", value=auto_state.get("imap_user", ""))
                new_imap_pass = st.text_input("Mot de passe d'application IMAP", value=auto_state.get("imap_password", ""), type="password")
            with c_im2:
                st.markdown("**📤 Envoi automatique des réponses (SMTP)**")
                new_smtp_host = st.text_input("Serveur SMTP", value=auto_state.get("smtp_host", "smtp.gmail.com"))
                new_smtp_port = st.number_input("Port SMTP (587 TLS / 465 SSL)", value=int(auto_state.get("smtp_port", 587)), min_value=1, max_value=65535)
                new_smtp_user = st.text_input("Utilisateur SMTP", value=auto_state.get("smtp_user", ""))
                new_smtp_pass = st.text_input("Mot de passe d'application SMTP", value=auto_state.get("smtp_password", ""), type="password")

            new_poll = st.slider(
                "Intervalle de vérification automatique de la boîte mail (secondes)",
                min_value=10,
                max_value=300,
                value=int(auto_state.get("poll_interval_seconds", 30)),
                step=5,
            )
            save_mail_cfg = st.form_submit_button("💾 Enregistrer la configuration IMAP / SMTP", type="primary", use_container_width=True)
            if save_mail_cfg:
                auto_state["imap_host"] = new_imap_host.strip()
                auto_state["imap_port"] = int(new_imap_port)
                auto_state["imap_user"] = new_imap_user.strip()
                auto_state["imap_password"] = new_imap_pass
                auto_state["smtp_host"] = new_smtp_host.strip()
                auto_state["smtp_port"] = int(new_smtp_port)
                auto_state["smtp_user"] = new_smtp_user.strip() or new_imap_user.strip()
                auto_state["smtp_password"] = new_smtp_pass or new_imap_pass
                auto_state["smtp_from"] = new_smtp_user.strip() or new_imap_user.strip()
                auto_state["poll_interval_seconds"] = int(new_poll)
                ea.save_automation_state(auto_state)
                st.success("✅ Configuration IMAP / SMTP enregistrée avec succès !")
                st.rerun()


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 📂 ARCHIVES
# ─────────────────────────────────────────────────────────────────────────
elif page == "Archives":
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">📂 GESTION DOCUMENTAIRE & LIVRABLES SYNCHRONISÉS</div>
            <div class="mb-hero-title">Archives & Projets (Site Web & Automatisés via Mail)</div>
            <p class="mb-hero-sub">
                Ouvrez, lisez, modifiez et téléchargez tous vos mémoires techniques (<code>.docx</code> et <code>.pdf</code>), qu'ils aient été générés manuellement sur le site ou automatiquement par e-mail.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    arch_tab1, arch_tab2 = st.tabs([
        "📑 Projets Synchronisés (Ouvrir, Lire, Modifier & Télécharger — Site & Mail)",
        "📦 Tous les Fichiers Livrables (.pdf & .docx)",
    ])

    with arch_tab1:
        synced_projects = ea.load_projects_db()
        if not synced_projects:
            st.info("Aucun projet synchronisé pour le moment. Générez un mémoire dans « Génération » ou lancez une simulation dans « Automatisation (Mails) ».")
        else:
            sel_sp_idx = st.selectbox(
                "📂 Sélectionnez un dossier de mémoire technique à ouvrir, modifier ou télécharger :",
                options=list(range(len(synced_projects))),
                format_func=lambda idx: (
                    f"{'📧 [Automatisé via Mail]' if synced_projects[idx].get('origin') == 'email' else '🖥️ [Créé sur le Site]'} "
                    f"{synced_projects[idx].get('nom_projet', 'Mémoire')} — {synced_projects[idx].get('created_at', '')}"
                ),
                key="arch_select_proj_idx",
            )
            render_interactive_project_viewer(synced_projects[sel_sp_idx], key_prefix=f"arch_proj_{sel_sp_idx}")

    with arch_tab2:
        os.makedirs("archives", exist_ok=True)
        fichiers = [f for f in os.listdir("archives") if f.endswith((".pdf", ".docx"))]
        if not fichiers:
            st.info("Aucun mémoire technique généré pour le moment. Rendez-vous dans l'onglet « Génération ».")
        else:
            fichiers.sort(key=lambda x: os.path.getmtime(os.path.join("archives", x)), reverse=True)

            # Bouton de téléchargement groupé en ZIP (mis en cache via @st.cache_data pour zéro latence)
            file_sigs = tuple(
                (f, os.path.getmtime(os.path.join("archives", f)))
                for f in fichiers
            )
            zip_arch_bytes = _build_cached_archives_zip(file_sigs)
            st.download_button(
                label=f"📦 Tout télécharger en un clic ({len(fichiers)} fichiers .ZIP)",
                data=zip_arch_bytes,
                file_name=f"Archives_Memoires_Techniques_{datetime.now().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="dl_all_archives_zip",
            )
            st.divider()

            for fichier in fichiers:
                chemin = os.path.join("archives", fichier)
                date_mod = datetime.fromtimestamp(os.path.getmtime(chemin)).strftime("%d/%m/%Y à %H:%M")
                is_automail_file = "AutoMail" in fichier
                c1, c2, c3 = st.columns([3.2, 1.1, 0.9])
                with c1:
                    badge_ext = "PDF OFFICIEL" if fichier.endswith(".pdf") else "WORD ÉDITABLE"
                    badge_col = "#3B82F6" if fichier.endswith(".pdf") else "#10B981"
                    mail_tag = (
                        "<span class='mb-zone-num' style='background:rgba(139,92,246,0.16); color:#8B5CF6; margin-right:8px;'>📧 AUTOMATISÉ VIA MAIL</span>"
                        if is_automail_file
                        else ""
                    )
                    st.markdown(
                        f"<div class='mb-zone-header' style='padding:10px 14px; margin-bottom:6px; border-left:4px solid {badge_col};'>"
                        f"{mail_tag}"
                        f"<span class='mb-zone-num' style='background:rgba(59,130,246,0.12); color:{badge_col}; margin-right:8px;'>{badge_ext}</span>"
                        f"<b style='color:var(--text-primary); font-size:14px;'>{fichier}</b><br>"
                        f"<span style='color:var(--text-muted); font-size:12px;'>Généré le {date_mod} &nbsp;•&nbsp; {os.path.getsize(chemin) / 1024:.0f} Ko</span>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                with c2:
                    with open(chemin, "rb") as f:
                        st.download_button(
                            "📥 Télécharger", data=f.read(), file_name=fichier,
                            mime="application/pdf" if fichier.endswith(".pdf")
                            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True, key=f"dl_{fichier}",
                        )
                with c3:
                    if st.button("🗑️ Supprimer", key=f"del_{fichier}", use_container_width=True):
                        os.remove(chemin)
                        st.rerun()


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 🎬 VIDÉOS
# ─────────────────────────────────────────────────────────────────────────
elif page == "Vidéos":
    if not st.session_state.get("owner_unlocked", False):
        st.error("🔒 Accès réservé à l'administrateur. Veuillez saisir votre code dans « 🔑 Accès Administrateur ».")
        st.stop()
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">🎬 STUDIO DE PRÉSENTATION CLIENT</div>
            <div class="mb-hero-title">Vidéos de Démonstration & Tutoriels</div>
            <p class="mb-hero-sub">
                Importez vos enregistrements d'écran (.mov Retina, .mp4), optimisez-les pour le web et générez vos vidéos de présentation commerciale en un clic.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    DESKTOP_OUT = Path.home() / "Desktop" / "MarketBuilding Vidéos"
    VIDEO_EXT = (".mov", ".mp4", ".m4v", ".webm", ".avi", ".mkv")

    def _desktop_videos() -> list:
        found = []
        for folder in [Path.home() / "Desktop", Path.home() / "Downloads", Path.home() / "Movies"]:
            try:
                found += [p for p in folder.iterdir() if p.suffix.lower() in VIDEO_EXT and p.is_file()]
            except (PermissionError, FileNotFoundError, OSError):
                continue  # dossier protégé par macOS : on l'ignore
        return sorted(found, key=lambda p: p.stat().st_mtime, reverse=True)

    def _progress_ui():
        bar = st.progress(0.0, text="Préparation…")
        return lambda frac, label: bar.progress(frac, text=f"{label} — {int(frac * 100)} %")

    def _copy_to_desktop(p: Path) -> Path:
        DESKTOP_OUT.mkdir(parents=True, exist_ok=True)
        dst = DESKTOP_OUT / p.name
        shutil.copy2(p, dst)
        return dst

    tab_gal, tab_imp, tab_gen = st.tabs(["Galerie", "Importer", "Générer une présentation"])

    # ── GALERIE ───────────────────────────────────────────────────────
    with tab_gal:
        meta = vt.load_meta()
        videos = vt.list_videos()
        if not videos:
            st.info("Aucune vidéo pour l'instant. Utilisez l'onglet « Importer ».")
        for v in videos:
            titre = meta.get("titles", {}).get(v.name, v.stem)
            is_feat = meta.get("featured") == v.name
            st.markdown(f"**{'⭐ ' if is_feat else ''}{titre}**")
            st.video(str(v))
            info = vt.media_info(v)
            st.caption(f"{info['width']}×{info['height']} • {info['duration']:.0f} s • {v.stat().st_size / 1e6:.1f} Mo")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                with open(v, "rb") as f:
                    st.download_button("Télécharger", f.read(), file_name=v.name, mime="video/mp4",
                                       use_container_width=True, key=f"vdl_{v.name}")
            with c2:
                if st.button("Retirer de l'accueil" if is_feat else "Mettre sur l'accueil",
                             key=f"feat_{v.name}", use_container_width=True):
                    meta["featured"] = None if is_feat else v.name
                    vt.save_meta(meta)
                    st.rerun()
            with c3:
                if st.button("Copier sur le Bureau", key=f"desk_{v.name}", use_container_width=True):
                    try:
                        st.success(f"Copiée dans : {_copy_to_desktop(v)}")
                    except OSError:
                        st.warning("macOS refuse l'accès au Bureau : utilisez « Télécharger ».")
            with c4:
                if st.button("Supprimer", key=f"vdel_{v.name}", use_container_width=True):
                    vt.delete_video(v)
                    st.rerun()
            st.divider()

    # ── IMPORTER ──────────────────────────────────────────────────────
    with tab_imp:
        st.subheader("Depuis votre Mac (Bureau, Téléchargements, Vidéos)")
        st.caption("Recommandé pour les gros fichiers : aucun envoi, conversion directe.")
        locaux = _desktop_videos()
        if not locaux:
            st.info("Aucune vidéo trouvée sur le Bureau, dans Téléchargements ou Vidéos.")
        else:
            labels = {f"{p.name}  ({p.stat().st_size / 1e6:.0f} Mo — {p.parent.name})": p for p in locaux}
            choix = st.multiselect("Vidéos à importer", list(labels.keys()))
            titre_imp = st.text_input("Titre (optionnel)", key="titre_local")
            if st.button("Importer et optimiser", type="primary", disabled=not choix):
                for i, lab in enumerate(choix, 1):
                    st.write(f"**{labels[lab].name}**")
                    t = f"{titre_imp} {i}".strip() if titre_imp and len(choix) > 1 else titre_imp
                    try:
                        out = vt.optimize_for_web(labels[lab], title=t, progress=_progress_ui())
                        st.success(f"Prête : {out.name} ({out.stat().st_size / 1e6:.1f} Mo)")
                    except Exception as e:
                        st.error(str(e))

        st.divider()
        st.subheader("Envoyer un fichier")
        uploads = st.file_uploader("Glissez vos vidéos ici (jusqu'à 2 Go)",
                                   type=[e.strip(".") for e in VIDEO_EXT], accept_multiple_files=True)
        if uploads and st.button("Optimiser les fichiers envoyés", type="primary"):
            for up in uploads:
                with tempfile.NamedTemporaryFile(suffix=Path(up.name).suffix, delete=False) as tmpf:
                    shutil.copyfileobj(up, tmpf)
                    tmp_path = Path(tmpf.name)
                try:
                    out = vt.optimize_for_web(tmp_path, title=Path(up.name).stem, progress=_progress_ui())
                    st.success(f"Prête : {out.name} ({out.stat().st_size / 1e6:.1f} Mo)")
                except Exception as e:
                    st.error(str(e))
                finally:
                    tmp_path.unlink(missing_ok=True)

    # ── GÉNÉRER ───────────────────────────────────────────────────────
    with tab_gen:
        sources = {f"[Galerie] {vt.load_meta().get('titles', {}).get(p.name, p.stem)}": p for p in vt.list_videos()
                   if not p.name.startswith("Presentation_")}
        sources.update({f"[Mac] {p.name}": p for p in _desktop_videos()})
        if not sources:
            st.info("Importez d'abord un enregistrement d'écran.")
        else:
            ent = config["entreprise"]
            src_label = st.selectbox("Enregistrement à utiliser", list(sources.keys()))
            src = sources[src_label]
            dur = vt.media_info(src)["duration"] or 1.0

            fmt = st.selectbox("Format", list(vt.FORMATS.keys()), index=0)
            c1, c2 = st.columns(2)
            with c1:
                g_title = st.text_input("Titre d'introduction", "Vos mémoires techniques BTP en quelques minutes")
                g_sub = st.text_input("Sous-titre", "Importez votre CCTP, récupérez un mémoire prêt à déposer")
                g_caption = st.text_input("Bandeau pendant la démo (optionnel)",
                                          "Analyse du CCTP  →  Rédaction  →  PDF & Word téléchargeables")
            with c2:
                g_outro = st.text_input("Titre de fin", "Gagnez plus d'appels d'offres")
                contact_def = " • ".join(x for x in [ent.get("telephone", ""), ent.get("email", "")] if x) \
                    or "marckethit67000@gmail.com"
                g_contact = st.text_input("Contact affiché à la fin", contact_def)
                g_speed = st.select_slider("Vitesse de la démo", options=[1.0, 1.25, 1.5, 2.0, 3.0], value=1.5,
                                           format_func=lambda x: f"×{x:g}")

            g_range = st.slider("Extrait de l'enregistrement (secondes)", 0.0, float(round(dur, 1)),
                                (0.0, float(round(dur, 1))), step=0.5)
            g_audio = st.file_uploader("Musique ou voix-off (optionnel)", type=["mp3", "m4a", "wav", "aac"])
            final_d = 3.5 + (g_range[1] - g_range[0]) / g_speed + 4.0 - 1.2
            st.caption(f"Durée finale estimée : {final_d:.0f} secondes")

            if st.button("Générer la vidéo de présentation", type="primary", use_container_width=True):
                audio_path = None
                if g_audio:
                    with tempfile.NamedTemporaryFile(suffix=Path(g_audio.name).suffix, delete=False) as af:
                        af.write(g_audio.getvalue())
                        audio_path = Path(af.name)
                try:
                    out = vt.generate_presentation(
                        src, fmt, g_title, g_sub, g_caption, g_outro, g_contact,
                        start=g_range[0], end=g_range[1], speed=g_speed,
                        audio=audio_path, progress=_progress_ui(),
                    )
                    try:
                        st.success(f"Vidéo générée et enregistrée sur votre Bureau : {_copy_to_desktop(out)}")
                    except OSError:
                        st.success("Vidéo générée (disponible dans la Galerie et via « Télécharger »).")
                    st.video(str(out))
                    with open(out, "rb") as f:
                        st.download_button("Télécharger la vidéo", f.read(), file_name=out.name,
                                           mime="video/mp4", type="primary", use_container_width=True)
                except Exception as e:
                    st.error(f"Échec de la génération : {e}")
                finally:
                    if audio_path:
                        audio_path.unlink(missing_ok=True)


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 📜 CONDITIONS GÉNÉRALES D'UTILISATION (CGU) EN DÉTAIL
# ─────────────────────────────────────────────────────────────────────────
elif page == "CGU":
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">📜 CADRE JURIDIQUE, CONTRACTUEL & TRANSPARENCE IA</div>
            <div class="mb-hero-title">Conditions Générales d'Utilisation (CGU)</div>
            <p class="mb-hero-sub">
                Version en vigueur : Octobre 2026 · Les présentes Conditions Générales d'Utilisation régissent en détail l'accès, le fonctionnement et l'utilisation de la plateforme <b>MarketBuilding Pro</b> dédiée à la génération automatisée de mémoires techniques BTP assistée par Intelligence Artificielle.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    cgu_tab1, cgu_tab2, cgu_tab3 = st.tabs([
        "📋 Articles 1 à 4 · Objet, Fonctionnement & Intelligence Artificielle (IA)",
        "⚖️ Articles 5 à 7 · Obligations, Validation Humaine & Propriété Intellectuelle",
        "🛡️ Articles 8 à 10 · Responsabilité, Disponibilité & Droit Applicable",
    ])

    with cgu_tab1:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 1</span>
                    <span class="mb-zone-badge">DÉFINITIONS & OBJET DU CONTRAT</span>
                </div>
                <div class="mb-zone-title">Article 1 · Définitions et Objet de la Plateforme</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    <b>1.1 Définitions :</b><br>
                    • <b>« Plateforme » :</b> désigne l'application web professionnelle <b>MarketBuilding Pro</b> dédiée à l'analyse de dossiers d'appels d'offres BTP et à la rédaction assistée de mémoires techniques.<br>
                    • <b>« Utilisateur » ou « Entreprise Soumissionnaire » :</b> désigne toute personne morale ou physique professionnelle du secteur du Bâtiment, du Génie Civil ou des Travaux Publics (BTP) utilisant la Plateforme.<br>
                    • <b>« DCE (Dossier de Consultation des Entreprises) » :</b> désigne l'ensemble des pièces d'un marché public ou privé importées par l'Utilisateur (Règlement de Consultation, CCTP, DPGF/DQE, PGC, étude de sol, plan de masse).<br>
                    • <b>« Livrable » :</b> désigne le projet de Mémoire Technique généré par la Plateforme aux formats PDF et Word (<code>.docx</code>).<br><br>
                    <b>1.2 Objet :</b> Les présentes Conditions Générales d'Utilisation (CGU) définissent les conditions juridiques et techniques dans lesquelles l'Utilisateur accède à la Plateforme et utilise ses fonctionnalités de génération unitaire ou multi-lots de mémoires techniques. Toute utilisation de la Plateforme implique l'acceptation sans réserve des présentes CGU.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #3B82F6; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 2</span>
                    <span class="mb-zone-badge">ARCHITECTURE DOCUMENTAIRE EN 5 ZONES</span>
                </div>
                <div class="mb-zone-title">Article 2 · Description Détaillée des Fonctionnalités et des Pièces Traitées</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    La Plateforme permet à l'Utilisateur d'importer simultanément et de croiser jusqu'à cinq (5) catégories de documents aux formats PDF, Word (<code>.docx</code>) ou Excel (<code>.xlsx</code>, <code>.xls</code>, <code>.csv</code>) :<br>
                    1. <b>Le Règlement de Consultation (RC) :</b> extraction des critères et sous-critères de jugement de la « Valeur Technique », du barème de notation, du plan imposé et de la limite de pages.<br>
                    2. <b>Les Cahiers des Clauses Techniques Particulières (CCTP Général Lot 00 et CCTP des lots techniques) :</b> identification de l'adresse du chantier, des contraintes d'accès et d'horaires, du caractère occupé ou non du site, des matériaux prescrits, des normes (DTU, Eurocodes), des modes opératoires et de la gestion des déchets (SOSED).<br>
                    3. <b>La DPGF / le DQE (Devis Quantitatif Estimatif) :</b> extraction des quantités exactes (m², m³, ml, unités, forfaits) afin de dimensionner le planning d'exécution, les cadences et les équipes.<br>
                    4. <b>Les Annexes Chantier (PGC Sécurité, Plan de masse, Étude géotechnique) :</b> prise en compte des risques spécifiques, de la coactivité, du levage, du cantonnement (base vie) et des zones de stockage.<br>
                    5. <b>Les Documents Internes de l'Entreprise :</b> exploitation de l'organigramme, des CV de l'encadrement, de la liste du personnel et des habilitations sécurité (CACES, AIPR, habilitations électriques, travail en hauteur), du parc matériel, des certifications (Qualibat, RGE, FNTP, ISO) ainsi que des fournisseurs et exutoires déchets habituels.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #F59E0B; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num" style="background:rgba(245,158,11,0.16); color:#F59E0B;">ARTICLE 3 · CLAUSE ESSENTIELLE</span>
                    <span class="mb-zone-badge">TRANSPARENCE SUR L'INTELLIGENCE ARTIFICIELLE (RÈGLEMENT UE 2024/1689 « AI ACT »)</span>
                </div>
                <div class="mb-zone-title">Article 3 · Recours à l'Intelligence Artificielle (IA) Générative et Algorithmique</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    <b>3.1 Information explicite sur l'usage de l'IA :</b> Conformément aux obligations de transparence prévues par le Règlement européen sur l'Intelligence Artificielle (UE 2024/1689 dit « AI Act »), l'Utilisateur est expressément informé et reconnaît que <b>la rédaction, la structuration et le croisement documentaire réalisés par la Plateforme MarketBuilding Pro sont produits par des systèmes d'Intelligence Artificielle (IA) générative et d'analyse sémantique automatisée</b>.<br><br>
                    <b>3.2 Fonctionnement de l'Intelligence Artificielle :</b><br>
                    • L'IA extrait le texte brut et les tableaux des documents déposés par l'Utilisateur ;<br>
                    • L'IA structure le sommaire selon l'ordre exact des critères du Règlement de Consultation (RC) ;<br>
                    • L'IA fusionne les prescriptions des différents CCTP et croise chaque poste technique avec les quantités de la DPGF et les ressources humaines et matérielles de l'Entreprise.<br><br>
                    <b>3.3 Limites inhérentes aux technologies d'Intelligence Artificielle :</b> L'Utilisateur reconnaît que tout contenu généré par une Intelligence Artificielle repose sur des modèles probabilistes de traitement du langage naturel. Malgré les consignes de rigueur appliquées au système, l'IA peut occasionnellement produire des formulations incomplètes, des erreurs d'interprétation de tableaux complexes, des omissions ou des imprécisions techniques nécessitant impérativement l'intervention d'un professionnel qualifié du BTP.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #10B981; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 4</span>
                    <span class="mb-zone-badge">RÈGLE ANTI-INVENTION & MARQUEURS</span>
                </div>
                <div class="mb-zone-title">Article 4 · Mécanisme « Anti-Invention » et Balises [À COMPLÉTER]</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    Afin de prévenir le risque d'hallucination documentaire propre aux modèles d'Intelligence Artificielle, la Plateforme applique une règle stricte dite <b>« Anti-Invention »</b> :<br>
                    • Lorsqu'une donnée propre au chantier ou à l'entreprise (nom d'un conducteur de travaux, distance d'une déchèterie, marque d'un matériau, référence d'un certificat, quantité absente) ne figure dans aucun des fichiers importés ni dans le Profil Entreprise, l'IA a pour instruction formelle de ne pas inventer cette information et d'insérer une balise apparente de type <code>[À COMPLÉTER : préciser ici...]</code>.<br>
                    • La présence ou l'absence de ces balises constitue une aide visuelle mais ne dispense en aucun cas l'Utilisateur d'un contrôle intégral de la véracité de chaque affirmation contenue dans le Livrable.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with cgu_tab2:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #10B981; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 5</span>
                    <span class="mb-zone-badge">CONTRÔLE HUMAIN OBLIGATOIRE</span>
                </div>
                <div class="mb-zone-title">Article 5 · Obligation Impérative de Relecture, Complétion et Validation Humaine</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    Les mémoires techniques générés par l'Intelligence Artificielle de la Plateforme constituent exclusivement des <b>projets préparatoires d'aide à la rédaction</b>. Ils n'ont pas valeur d'engagement contractuel automatique.<br><br>
                    Avant toute signature, exportation finale ou dépôt d'une offre sur un profil d'acheteur public (PLACE, Maximilien, AWS, Mégalis, Chorus Pro, etc.) ou auprès d'un maître d'ouvrage privé, l'Utilisateur a l'obligation absolue de :<br>
                    1. <b>Relire intégralement</b> l'ensemble des chapitres, tableaux et annexes du mémoire technique généré ;<br>
                    2. <b>Remplacer toutes les balises</b> <code>[À COMPLÉTER : ...]</code> par les données exactes et vérifiées de son entreprise et du chantier ;<br>
                    3. <b>Contrôler la conformité technique et réglementaire</b> des méthodes décrites (respect des DTU, Eurocodes, avis techniques CSTB, prescriptions du CCTP et du PGC) ;<br>
                    4. <b>Vérifier la cohérence arithmétique et opérationnelle</b> entre le mémoire technique, le DPGF/DQE chiffré, le bordereau des prix unitaires (BPU) et le planning d'exécution.<br><br>
                    L'Entreprise Soumissionnaire assume seule la responsabilité juridique, technique et financière des engagements contenus dans le mémoire technique qu'elle remet au Maître d'Ouvrage.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 6</span>
                    <span class="mb-zone-badge">RESPONSABILITÉ DE L'UTILISATEUR</span>
                </div>
                <div class="mb-zone-title">Article 6 · Qualité, Licéité et Exactitude des Documents Importés</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    L'Utilisateur garantit détenir l'ensemble des droits et autorisations nécessaires pour importer sur la Plateforme les pièces du DCE ainsi que les documents internes de son entreprise (CV, diplômes, CACES, certificats Qualibat/RGE, références chantiers).<br><br>
                    L'Utilisateur s'engage à fournir des informations sincères et à jour dans son « Profil Entreprise ». Toute fausse déclaration dans le cadre d'un marché public (notamment sur les effectifs, les qualifications professionnelles ou les certifications) relève de la seule responsabilité pénale et administrative de l'Entreprise Soumissionnaire conformément au Code de la commande publique.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #8B5CF6; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 7</span>
                    <span class="mb-zone-badge">DROITS D'AUTEUR & LIVRABLES</span>
                </div>
                <div class="mb-zone-title">Article 7 · Propriété Intellectuelle</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    <b>7.1 Propriété de la Plateforme :</b> L'architecture logicielle, le code source, l'interface graphique, les modèles de mise en page, les algorithmes d'orchestration et la marque <b>MarketBuilding Pro</b> demeurent la propriété exclusive de l'Éditeur.<br><br>
                    <b>7.2 Propriété des Documents et Livrables :</b> L'Utilisateur conserve l'entière propriété intellectuelle des fichiers qu'il téléverse sur la Plateforme. Dès leur génération, l'Utilisateur dispose d'un droit d'usage, de modification, de reproduction et de diffusion plein, entier et exclusif sur les Mémoires Techniques (PDF et Word) générés pour les besoins de ses réponses aux appels d'offres.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with cgu_tab3:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #EF4444; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 8</span>
                    <span class="mb-zone-badge">GARANTIES & LIMITES DE RESPONSABILITÉ</span>
                </div>
                <div class="mb-zone-title">Article 8 · Obligation de Moyens et Limitation de Responsabilité</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    <b>8.1 Obligation de moyens :</b> L'Éditeur fournit un outil technologique d'assistance par Intelligence Artificielle et est soumis à une stricte obligation de moyens. Il ne garantit en aucun cas l'attribution d'un marché public ou privé, ni l'obtention d'une note minimale sur le critère de la « Valeur Technique ».<br><br>
                    <b>8.2 Exclusion des dommages indirects :</b> En aucun cas l'Éditeur ne pourra être tenu responsable :<br>
                    • Du rejet d'une offre pour irrégularité, incomplétude ou non-conformité au Règlement de Consultation ;<br>
                    • D'une perte de chance de remporter un appel d'offres, d'une perte de chiffre d'affaires, de marge ou d'exploitation ;<br>
                    • De pénalités de retard, de malfaçons, de surcoûts d'exécution ou de litiges avec le Maître d'Ouvrage, le Maître d'Œuvre ou le coordonnateur SPS résultant d'engagements pris dans le mémoire technique sans vérification humaine préalable par l'Utilisateur.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #3B82F6; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 9</span>
                    <span class="mb-zone-badge">CONTINUITÉ DE SERVICE & SÉCURITÉ</span>
                </div>
                <div class="mb-zone-title">Article 9 · Disponibilité du Service et Maintenance</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    La Plateforme est accessible 24h/24 et 7j/7, sous réserve des opérations de maintenance évolutive ou corrective et de la disponibilité des infrastructures réseaux et des moteurs d'Intelligence Artificielle tiers. L'Utilisateur est invité à générer et vérifier ses mémoires techniques avec une marge d'anticipation suffisante avant l'heure limite de remise des plis fixée au Règlement de Consultation.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 10</span>
                    <span class="mb-zone-badge">JURIDICTION COMPÉTENTE</span>
                </div>
                <div class="mb-zone-title">Article 10 · Modification des CGU, Droit Applicable et Litiges</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    L'Éditeur se réserve le droit d'adapter ou de modifier à tout moment les présentes CGU afin de tenir compte des évolutions législatives (notamment en matière d'Intelligence Artificielle et de commande publique) ou fonctionnelles de la Plateforme.<br><br>
                    Les présentes CGU sont régies exclusivement par le <b>droit français</b>. En cas de différend relatif à leur interprétation ou à leur exécution, et à défaut de résolution amiable dans un délai de trente (30) jours, compétence expresse est attribuée aux tribunaux compétents du ressort du siège social de l'Éditeur.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 🔒 POLITIQUE DE CONFIDENTIALITÉ & RGPD EN DÉTAIL
# ─────────────────────────────────────────────────────────────────────────
elif page == "Confidentialité":
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">🔒 PROTECTION DES DONNÉES, SECRET DES AFFAIRES & RGPD</div>
            <div class="mb-hero-title">Politique de Confidentialité & Protection des Données</div>
            <p class="mb-hero-sub">
                Version en vigueur : Octobre 2026 · La présente Politique de Confidentialité détaille comment <b>MarketBuilding Pro</b> protège vos dossiers d'appels d'offres (DCE), le secret industriel et commercial de votre entreprise BTP ainsi que les données personnelles conformément au Règlement Général sur la Protection des Données (RGPD - Règlement UE 2016/679) et à la loi Informatique et Libertés.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    priv_tab1, priv_tab2, priv_tab3 = st.tabs([
        "🗂️ Articles 1 à 3 · Données Collectées, Finalités & Secret des Affaires",
        "🤖 Articles 4 à 6 · Traitement par l'IA, Sécurité & Durée de Conservation",
        "👤 Articles 7 à 8 · Droits RGPD des Utilisateurs, Cookies & Contact CNIL",
    ])

    with priv_tab1:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 1</span>
                    <span class="mb-zone-badge">RESPONSABLE DU TRAITEMENT</span>
                </div>
                <div class="mb-zone-title">Article 1 · Identité et Engagements du Responsable de Traitement</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    Dans le cadre de l'exploitation de la plateforme <b>MarketBuilding Pro</b>, l'Éditeur agit en qualité de :
                    <br>• <b>Responsable de traitement</b> pour les données de compte et de paramétrage de l'Utilisateur ;
                    <br>• <b>Sous-traitant au sens de l'article 28 du RGPD</b> pour les documents professionnels et les données à caractère personnel figurant dans les pièces d'appels d'offres (DCE) et les justificatifs d'entreprise importés par l'Utilisateur pour générer ses mémoires techniques.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #3B82F6; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 2</span>
                    <span class="mb-zone-badge">INVENTAIRE DES DONNÉES TRAITÉES</span>
                </div>
                <div class="mb-zone-title">Article 2 · Catégories de Données Collectées et Traitées</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    La Plateforme traite strictement les catégories de données suivantes fournies par l'Utilisateur :<br><br>
                    <b>2.1 Données d'identification de l'Entreprise (Profil Entreprise) :</b><br>
                    • Raison sociale, forme juridique, numéro SIRET, adresse du siège social, téléphone, adresse email professionnelle, nom du signataire/dirigeant, assurances décennales et RC Pro, chiffre d'affaires, labels et certifications (Qualibat, RGE, FNTP, Qualifelec, ISO).<br><br>
                    <b>2.2 Données à caractère personnel relatives aux équipes BTP (Documents Entreprise) :</b><br>
                    • Noms, prénoms, fonctions, ancienneté, parcours professionnel (CV de conducteurs de travaux, chefs de chantier, compagnons) et titres d'habilitation réglementaire (CACES, AIPR, habilitations électriques, SST, travail en hauteur) présents dans les organigrammes et justificatifs importés.<br><br>
                    <b>2.3 Données techniques et économiques de marché (Pièces DCE) :</b><br>
                    • Règlement de Consultation (RC), CCTP, quantitatifs DPGF/DQE, PGC sécurité, études de sol, plans de masse, listes du parc matériel, fiches matériaux et références de chantiers similaires.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #10B981; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 3</span>
                    <span class="mb-zone-badge">BASES LÉGALES (ART. 6 RGPD)</span>
                </div>
                <div class="mb-zone-title">Article 3 · Finalités du Traitement et Bases Juridiques</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    Vos données et documents sont traités exclusivement pour les finalités suivantes :<br>
                    1. <b>Exécution du service demandé par l'Utilisateur (Article 6.1.b du RGPD) :</b> extraction automatisée du texte des pièces DCE, croisement avec les ressources de l'entreprise et génération des mémoires techniques aux formats PDF et Word.<br>
                    2. <b>Conservation dans l'espace « Profil Entreprise » et « Archives » (Article 6.1.b du RGPD) :</b> permettre à l'Utilisateur de retrouver ses informations pré-remplies et de retélécharger ses mémoires techniques générés.<br>
                    3. <b>Sécurité et prévention des abus (Article 6.1.f du RGPD - Intérêt légitime) :</b> protection de l'intégrité technique de la Plateforme.<br><br>
                    <b>Exclusion totale de prospection et de revente :</b> Aucune donnée importée sur MarketBuilding Pro n'est vendue, louée, cédée ou partagée à des tiers commerciaux, courtiers en données ou entreprises concurrentes.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with priv_tab2:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #F59E0B; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num" style="background:rgba(245,158,11,0.16); color:#F59E0B;">ARTICLE 4 · CONFIDENTIALITÉ IA & SECRET DES AFFAIRES</span>
                    <span class="mb-zone-badge">PROTECTION DES MÉTHODES & CHIFFRAGES BTP</span>
                </div>
                <div class="mb-zone-title">Article 4 · Encadrement du Traitement par Intelligence Artificielle (IA)</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    La génération du mémoire technique fait appel à des interfaces de programmation (API) professionnelles d'Intelligence Artificielle soumises à des engagements stricts de confidentialité :<br><br>
                    • <b>Isolation stricte entre entreprises :</b> Les documents déposés par une entreprise (méthodologies propres, DPGF, prix, CV, liste matériel, anciens mémoires techniques) sont traités dans un contexte d'exécution isolé. Ils ne sont <b>jamais</b> accessibles ni réutilisés pour générer les mémoires techniques d'une autre entreprise utilisatrice.<br>
                    • <b>Absence d'entraînement public sur vos secrets d'affaires :</b> Les appels d'API professionnels utilisés par la Plateforme excluent l'utilisation des documents téléversés par les clients pour entraîner des modèles publics d'Intelligence Artificielle.<br>
                    • <b>Traitement en mémoire vive lors de l'extraction :</b> L'extraction textuelle des fichiers PDF, Word et Excel importés dans les 5 zones de dépôt s'effectue à la volée en mémoire vive le temps de la génération du livrable.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 5</span>
                    <span class="mb-zone-badge">RÉTENTION & PURGE DES DONNÉES</span>
                </div>
                <div class="mb-zone-title">Article 5 · Durée de Conservation et Suppression Immédiate</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    • <b>Fichiers DCE téléversés (RC, CCTP, DPGF, PGC, Annexes) :</b> Les fichiers déposés dans les zones d'upload sont traités en mémoire vive pendant la session de génération et sont automatiquement purgés à la fermeture ou au réinitialisation de la session.<br>
                    • <b>Mémoires générés dans l'onglet « Archives » :</b> Les fichiers PDF et Word générés sont conservés dans le répertoire d'archives de l'Utilisateur afin de permettre leur téléchargement ultérieur. L'Utilisateur peut les <b>supprimer définitivement en un clic</b> à tout moment depuis l'onglet <b>« 🗂️ Archives »</b> via le bouton <i>« 🗑️ Supprimer »</i>.<br>
                    • <b>Données du « Profil Entreprise » :</b> Conservées tant que l'Utilisateur utilise la Plateforme, et modifiables ou effaçables à tout moment depuis l'onglet <b>« 🏢 Profil Entreprise »</b>.
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #8B5CF6; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 6</span>
                    <span class="mb-zone-badge">ARTICLE 32 DU RGPD</span>
                </div>
                <div class="mb-zone-title">Article 6 · Mesures de Sécurité Techniques et Organisationnelles</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    Conformément à l'article 32 du RGPD et à la protection du secret des affaires (articles L. 151-1 et suivants du Code de commerce), la Plateforme met en œuvre les mesures de sécurité suivantes :<br>
                    • <b>Chiffrement des communications :</b> Flux chiffrés en transit via les protocoles TLS/HTTPS vers les services d'analyse ;<br>
                    • <b>Cloisonnement des accès :</b> Séparation stricte entre l'interface Utilisateur et l'espace d'administration système protégé par authentification propriétaire ;<br>
                    • <b>Contrôle d'intégrité des fichiers :</b> Validation des formats de fichiers autorisés (PDF, DOCX, XLSX, CSV, TXT) lors de l'importation.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with priv_tab3:
        st.markdown(
            """
            <div class="mb-zone-header" style="border-left:4px solid #10B981; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 7</span>
                    <span class="mb-zone-badge">ARTICLES 15 À 22 DU RGPD</span>
                </div>
                <div class="mb-zone-title">Article 7 · Vos Droits sur vos Données Personnelles (RGPD)</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    Conformément au Règlement Général sur la Protection des Données (RGPD) et à la Loi Informatique et Libertés modifiée, toute personne physique dont les données figurent sur la Plateforme dispose des droits suivants :<br><br>
                    • <b>Droit d'accès (Art. 15 RGPD) :</b> obtenir la confirmation que des données vous concernant sont traitées et en recevoir une copie ;<br>
                    • <b>Droit de rectification (Art. 16 RGPD) :</b> modifier ou corriger immédiatement toute donnée inexacte ou incomplète (directement via l'onglet <i>« 🏢 Profil Entreprise »</i>) ;<br>
                    • <b>Droit à l'effacement / « Droit à l'oubli » (Art. 17 RGPD) :</b> supprimer à tout moment vos mémoires archivés depuis l'onglet <i>« 🗂️ Archives »</i> ou demander la suppression intégrale de vos données ;<br>
                    • <b>Droit à la limitation et d'opposition (Art. 18 et 21 RGPD) :</b> vous opposer au traitement de vos données pour des motifs légitimes ;<br>
                    • <b>Droit à la portabilité (Art. 20 RGPD) :</b> récupérer vos données et vos livrables dans un format structuré, couramment utilisé et lisible par machine (PDF, DOCX, JSON).
                </div>
            </div>

            <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:14px;">
                <div class="mb-zone-top">
                    <span class="mb-zone-num">ARTICLE 8</span>
                    <span class="mb-zone-badge">COOKIES TECHNIQUES & RECOURS CNIL</span>
                </div>
                <div class="mb-zone-title">Article 8 · Cookies Techniques de Session et Réclamation auprès de la CNIL</div>
                <div class="mb-zone-desc" style="margin-bottom:0;">
                    <b>8.1 Absence de traceurs publicitaires :</b> La plateforme <b>MarketBuilding Pro</b> n'utilise aucun cookie publicitaire, aucun traceur de ciblage marketing ni aucun outil de profilage commercial tiers. Seuls des cookies techniques strictement nécessaires au fonctionnement de la session d'application et à la mémorisation de votre préférence d'affichage (Mode Clair / Mode Sombre) sont utilisés.<br><br>
                    <b>8.2 Exercice de vos droits et contact CNIL :</b> Vous pouvez exercer vos droits directement depuis les outils de gestion intégrés à la Plateforme (modification du Profil Entreprise, suppression immédiate dans Archives) ou auprès de l'administrateur de votre espace. Si vous estimez, après nous avoir contactés, que vos droits « Informatique et Libertés » ne sont pas respectés, vous avez la faculté d'introduire une réclamation auprès de l'autorité de contrôle compétente en France : la <b>Commission Nationale de l'Informatique et des Libertés (CNIL)</b> — 3 Place de Fontenoy, TSA 80715, 75334 Paris Cedex 07 (www.cnil.fr).
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# Pied de page présent sur toutes les pages
st.markdown(
    "<div style='text-align:center; margin-top:36px; padding-top:14px; border-top:1px solid var(--border-color); color:var(--text-muted); font-size:12px;'>"
    "MarketBuilding Pro · Plateforme soumise aux <b>Conditions Générales d'Utilisation (CGU)</b> et à la <b>Politique de Confidentialité (RGPD)</b> consultables dans le menu principal."
    "</div>",
    unsafe_allow_html=True,
)
