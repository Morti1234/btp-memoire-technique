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
    Injecte la feuille de style globale :
    - Toute la Navigation Principale (barre latérale & onglets de navigation) est TOUJOURS en Mode Clair
      pour une lisibilité et une présentation client irréprochables.
    - L'espace de travail central bascule entre Mode Clair (OFF) et Mode Sombre (ON) selon le switch.
    """
    if dark_mode:
        vars_css = """
            --bg-app: #0B0F17;
            --bg-card: #151C28;
            --bg-card-alt: #1B2434;
            --bg-hero: linear-gradient(135deg, #111827 0%, #1E293B 100%);
            --border-color: #263143;
            --border-hover: #3B82F6;
            --text-primary: #F8FAFC;
            --text-secondary: #CBD5E1;
            --text-muted: #94A3B8;
            --tag-bg: rgba(59, 130, 246, 0.14);
            --tag-text: #93C5FD;
            --tag-border: rgba(59, 130, 246, 0.32);
            --shadow-card: 0 8px 24px rgba(0, 0, 0, 0.35);
            --input-bg: #0F1520;
            --uploader-bg: #111722;
        """
    else:
        vars_css = """
            --bg-app: #F8FAFC;
            --bg-card: #FFFFFF;
            --bg-card-alt: #F1F5F9;
            --bg-hero: linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 100%);
            --border-color: #E2E8F0;
            --border-hover: #2563EB;
            --text-primary: #0F172A;
            --text-secondary: #334155;
            --text-muted: #64748B;
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
            .stApp {{
                background-color: var(--bg-app) !important;
                color: var(--text-primary) !important;
            }}

            .block-container {{
                padding-top: 1.5rem !important;
                padding-bottom: 3rem !important;
                max-width: 1180px !important;
            }}

            /* ── NAVIGATION PRINCIPALE TOUJOURS EN MODE CLAIR ── */
            section[data-testid="stSidebar"] {{
                background-color: #FFFFFF !important;
                border-right: 1px solid #E2E8F0 !important;
                box-shadow: 4px 0 20px rgba(15, 23, 42, 0.04) !important;
            }}
            section[data-testid="stSidebar"] * {{
                color: #0F172A !important;
            }}
            section[data-testid="stSidebar"] .stCaption,
            section[data-testid="stSidebar"] small {{
                color: #64748B !important;
            }}
            section[data-testid="stSidebar"] hr {{
                border-color: #E2E8F0 !important;
            }}
            /* Boutons radio de la Navigation Principale en Mode Clair structuré */
            section[data-testid="stSidebar"] div[role="radiogroup"] label {{
                background-color: #F8FAFC !important;
                border: 1px solid #E2E8F0 !important;
                border-radius: 10px !important;
                padding: 10px 14px !important;
                margin-bottom: 6px !important;
                transition: all 0.15s ease !important;
                font-weight: 600 !important;
            }}
            section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{
                background-color: #EFF6FF !important;
                border-color: #93C5FD !important;
            }}
            section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"],
            section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {{
                background-color: #EFF6FF !important;
                border: 1.5px solid #2563EB !important;
                border-left: 4px solid #2563EB !important;
                color: #1D4ED8 !important;
                box-shadow: 0 2px 8px rgba(37, 99, 235, 0.12) !important;
            }}
            section[data-testid="stSidebar"] div[data-testid="stExpander"] {{
                background-color: #F8FAFC !important;
                border: 1px solid #E2E8F0 !important;
                box-shadow: none !important;
            }}
            section[data-testid="stSidebar"] div[data-baseweb="input"] > div {{
                background-color: #FFFFFF !important;
                border-color: #CBD5E1 !important;
                color: #0F172A !important;
            }}

            /* Titres et textes de la zone principale */
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
            label, .stCaption {{
                color: var(--text-secondary) !important;
            }}

            /* Champs de saisie, formulaires, sélecteurs, expanders */
            div[data-testid="stForm"],
            div[data-testid="stExpander"] {{
                background-color: var(--bg-card) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 14px !important;
                box-shadow: var(--shadow-card) !important;
            }}
            div[data-testid="stExpander"] details summary {{
                color: var(--text-primary) !important;
                font-weight: 600 !important;
            }}

            div[data-baseweb="input"] > div,
            div[data-baseweb="textarea"] > div,
            div[data-baseweb="select"] > div {{
                background-color: var(--input-bg) !important;
                border-color: var(--border-color) !important;
                color: var(--text-primary) !important;
                border-radius: 10px !important;
            }}
            div[data-baseweb="input"] input,
            div[data-baseweb="textarea"] textarea {{
                color: var(--text-primary) !important;
            }}

            /* Zone de drag & drop (File Uploader) */
            [data-testid="stFileUploader"] section {{
                background-color: var(--uploader-bg) !important;
                border: 1.5px dashed var(--border-color) !important;
                border-radius: 12px !important;
                padding: 14px !important;
                transition: all 0.2s ease;
            }}
            [data-testid="stFileUploader"] section:hover {{
                border-color: var(--border-hover) !important;
            }}
            [data-testid="stFileUploader"] section * {{
                color: var(--text-secondary) !important;
            }}
            [data-testid="stFileUploader"] section button {{
                background-color: var(--bg-card) !important;
                border: 1px solid var(--border-color) !important;
                color: var(--text-primary) !important;
                border-radius: 8px !important;
                font-weight: 600 !important;
            }}

            /* Boutons principaux */
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

            /* Onglets (Tabs) de navigation principale toujours en mode clair */
            div[data-baseweb="tab-list"] {{
                background-color: #FFFFFF !important;
                border: 1px solid #E2E8F0 !important;
                border-radius: 12px !important;
                padding: 4px 8px !important;
                margin-bottom: 14px !important;
            }}
            button[data-baseweb="tab"] {{
                color: #475569 !important;
                font-weight: 600 !important;
                font-size: 14px !important;
            }}
            button[data-baseweb="tab"][aria-selected="true"] {{
                color: #1D4ED8 !important;
                font-weight: 700 !important;
            }}

            /* Composants personnalisés MarketBuilding */
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

            /* Cartes des 5 zones d'upload */
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
                font-size: 12.8px;
                line-height: 1.45;
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
                background: rgba(16, 185, 129, 0.08);
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
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_custom_css(st.session_state["dark_mode"])


def render_theme_toggle(location_key: str):
    """Affiche un bouton ON/OFF clair pour activer ou désactiver le Mode Sombre."""
    current_dark = st.session_state.get("dark_mode", False)
    new_dark = st.toggle(
        "🌙 Mode Sombre (ON / OFF)",
        value=current_dark,
        key=f"theme_toggle_{location_key}",
        help="Basculez l'espace de travail entre le Mode Clair (OFF) et le Mode Sombre (ON).",
    )
    if new_dark != current_dark:
        st.session_state["dark_mode"] = new_dark
        config["dark_mode"] = new_dark
        save_config(config)
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────
# 0.5 SÉCURITÉ - ACCÈS PRIVÉ PAR MOT DE PASSE
# ─────────────────────────────────────────────────────────────────────────
def check_password():
    """Returns `True` if the user had the correct password."""
    if "app_password" not in st.secrets:
        return True

    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    # Barre supérieure sur l'écran de connexion avec switch Mode Sombre ON/OFF
    top_col1, top_col2 = st.columns([4, 1.3])
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
# 2. BARRE LATÉRALE – Navigation Principale (Mode Clair) & Accès Propriétaire
# ─────────────────────────────────────────────────────────────────────────
with st.sidebar:
    # Le logo dans la barre latérale est toujours rendu en Mode Clair (dark_mode=False)
    st.markdown(get_marketbuilding_logo_svg(False), unsafe_allow_html=True)

    # Interrupteur Mode Sombre ON/OFF pour l'espace de travail
    render_theme_toggle("sidebar")
    st.caption(f"Affichage : **{'🌙 Mode Sombre (ON)' if st.session_state['dark_mode'] else '☀️ Mode Clair (OFF)'}**")

    st.divider()
    st.markdown(
        "<div style='font-size:11.5px; font-weight:700; color:#64748B; text-transform:uppercase; letter-spacing:0.8px; margin-bottom:6px;'>Navigation Principale</div>",
        unsafe_allow_html=True,
    )

    # Navigation visible par les clients (aucune mention d'IA ni de Paramètres IA)
    CLIENT_NAV_LABELS = {
        "Génération": "📄  Générer un Mémoire",
        "Profil": "🏢  Profil Entreprise",
        "Archives": "📂  Archives & Livrables",
        "CGU": "⚖️  Conditions Générales (CGU)",
    }

    # Accès Propriétaire réservé à l'administrateur (?admin=1 ou déverrouillage propriétaire)
    is_owner = str(st.query_params.get("admin", "")).strip() == "1" or st.session_state.get("owner_unlocked", False)
    if is_owner:
        CLIENT_NAV_LABELS["Vidéos"] = "🎬  Studio Vidéos (Admin)"
        CLIENT_NAV_LABELS["Paramètres"] = "⚙️  Paramètres Système (Admin)"

    page = st.radio(
        "Navigation",
        options=list(CLIENT_NAV_LABELS.keys()),
        format_func=lambda k: CLIENT_NAV_LABELS[k],
        index=0,
        label_visibility="collapsed",
    )

    st.divider()

    # Carte d'état client (sans aucune mention de Moteur IA)
    entreprise_ok = bool(config["entreprise"].get("nom", "").strip())
    os.makedirs("archives", exist_ok=True)
    nb_archives = len([f for f in os.listdir("archives") if f.endswith((".pdf", ".docx"))])

    st.markdown(
        f"""
        <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px;">
            <div style="font-size:11.5px; font-weight:700; color:#64748B; text-transform:uppercase; letter-spacing:0.6px; margin-bottom:8px;">
                Dossier Entreprise
            </div>
            <div style="font-size:12.5px; color:#334155; margin-bottom:6px;">
                {'🟢' if entreprise_ok else '🔵'} <b>Société :</b> {config['entreprise'].get('nom') or 'Via pièces déposées'}
            </div>
            <div style="font-size:12.5px; color:#334155;">
                📂 <b>Livrables archivés :</b> {nb_archives} fichier(s)
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.divider()
    st.caption("MarketBuilding Pro v3.5 · Édition Client BTP")

    # Zone discrète réservée uniquement au propriétaire en bas de barre latérale
    with st.expander("🔒 Accès Propriétaire", expanded=False):
        if not is_owner:
            owner_code = st.text_input("Code propriétaire", type="password", key="owner_pin_input", placeholder="Code admin...")
            if owner_code:
                expected_pin = st.secrets.get("app_password", "admin") if "app_password" in st.secrets else "admin"
                if owner_code == expected_pin:
                    st.session_state["owner_unlocked"] = True
                    st.rerun()
                else:
                    st.error("Code incorrect")
        else:
            st.success("Mode Propriétaire actif")
            if st.button("Verrouiller l'accès propriétaire", use_container_width=True):
                st.session_state["owner_unlocked"] = False
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


def extract_category_documents(files_list, category_label: str) -> str:
    """
    Extrait et concatène un ou plusieurs fichiers d'une même catégorie
    avec des en-têtes clairs indiquant le nom de chaque fichier source.
    """
    if not files_list:
        return ""
    if not isinstance(files_list, list):
        files_list = [files_list]

    extracted_parts: list[str] = []
    for f in files_list:
        if f is None:
            continue
        txt = extract_text_from_document(f)
        extracted_parts.append(
            f"================ {category_label.upper()} : {f.name} ================\n{txt}"
        )
    return "\n\n".join(extracted_parts)


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

# ═══════════════════════════════════════════════════════════════════════
#  BARRE SUPÉRIEURE GLOBALE (THÈME SOMBRE ON/OFF & STATUT CLIENT)
# ═══════════════════════════════════════════════════════════════════════
top_bar_left, top_bar_right = st.columns([3.6, 1.4])
with top_bar_left:
    nom_ent_top = config["entreprise"].get("nom", "").strip() or "Espace Candidat BTP"
    st.markdown(
        f"<div style='font-size:12.5px; color:var(--text-muted); font-weight:600; padding-top:6px;'>"
        f"🏗️ <b>MarketBuilding Pro</b> &nbsp;·&nbsp; {nom_ent_top} &nbsp;·&nbsp; Réponse aux Appels d'Offres Publics & Privés"
        f"</div>",
        unsafe_allow_html=True,
    )
with top_bar_right:
    render_theme_toggle("topbar")

# ═══════════════════════════════════════════════════════════════════════
#  PAGES DE L'APPLICATION
# ═══════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────
# PAGE : ⚙️ RÉGLAGES
# ─────────────────────────────────────────────────────────────────────────
if page == "Paramètres":
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
            st.video(str(_feat_main), autoplay=True, loop=True, muted=True)
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

    ALLOWED_DOC_TYPES = ["pdf", "docx", "xlsx", "xls", "csv"]

    mode_tab1, mode_tab2 = st.tabs([
        "🧩 Dossier Complet Croisé (5 Catégories : RC + CCTP + DPGF + Annexes + Entreprise)",
        "📚 Génération en Lot (Plusieurs marchés distincts d'un coup)",
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

            with open(f"archives/Memoire_{nom_fichier}_{date_str}.pdf", "wb") as f:
                f.write(pdf_bytes)
            with open(f"archives/Memoire_{nom_fichier}_{date_str}.docx", "wb") as f:
                f.write(docx_bytes)

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

                with open(f"archives/Memoire_{nom_fichier}_{date_str}.pdf", "wb") as f:
                    f.write(pdf_bytes)
                with open(f"archives/Memoire_{nom_fichier}_{date_str}.docx", "wb") as f:
                    f.write(docx_bytes)

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
                    with st.expander("👁️ Aperçu du document PDF", expanded=True):
                        b64_pdf = base64.b64encode(res["pdf_bytes"]).decode("utf-8")
                        st.markdown(
                            f'<iframe src="data:application/pdf;base64,{b64_pdf}#view=FitH" width="100%" height="700" type="application/pdf"></iframe>',
                            unsafe_allow_html=True,
                        )


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 📂 ARCHIVES
# ─────────────────────────────────────────────────────────────────────────
elif page == "Archives":
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">📂 GESTION DOCUMENTAIRE & LIVRABLES</div>
            <div class="mb-hero-title">Archives des Mémoires Techniques</div>
            <p class="mb-hero-sub">
                Retrouvez l'ensemble de vos mémoires techniques générés (versions officielles PDF et Word .docx), téléchargeables individuellement ou en pack .ZIP.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    os.makedirs("archives", exist_ok=True)
    fichiers = [f for f in os.listdir("archives") if f.endswith((".pdf", ".docx"))]
    if not fichiers:
        st.info("Aucun mémoire technique généré pour le moment. Rendez-vous dans l'onglet « Génération ».")
    else:
        fichiers.sort(key=lambda x: os.path.getmtime(os.path.join("archives", x)), reverse=True)

        # Bouton de téléchargement groupé en ZIP
        zip_arch_buf = io.BytesIO()
        with zipfile.ZipFile(zip_arch_buf, "w", zipfile.ZIP_DEFLATED) as zarch:
            for fichier in fichiers:
                zarch.write(os.path.join("archives", fichier), arcname=fichier)
        st.download_button(
            label=f"📦 Tout télécharger en un clic ({len(fichiers)} fichiers .ZIP)",
            data=zip_arch_buf.getvalue(),
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
            c1, c2, c3 = st.columns([3.2, 1.1, 0.9])
            with c1:
                badge_ext = "PDF OFFICIEL" if fichier.endswith(".pdf") else "WORD ÉDITABLE"
                badge_col = "#3B82F6" if fichier.endswith(".pdf") else "#10B981"
                st.markdown(
                    f"<div class='mb-zone-header' style='padding:10px 14px; margin-bottom:6px; border-left:4px solid {badge_col};'>"
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
# PAGE : ⚖️ CONDITIONS GÉNÉRALES D'UTILISATION (CGU)
# ─────────────────────────────────────────────────────────────────────────
elif page == "CGU":
    st.markdown(
        """
        <div class="mb-hero">
            <div class="mb-hero-badge">⚖️ CADRE CONTRACTUEL & MENTIONS LÉGALES</div>
            <div class="mb-hero-title">Conditions Générales d'Utilisation (CGU)</div>
            <p class="mb-hero-sub">
                Les présentes Conditions Générales d'Utilisation régissent l'accès et l'utilisation de la plateforme <b>MarketBuilding Pro</b> dédiée à l'élaboration de mémoires techniques pour les marchés publics et privés du BTP.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="mb-zone-header" style="border-left:4px solid #2563EB; margin-bottom:14px;">
            <div class="mb-zone-title">Article 1 · Objet du Service</div>
            <div class="mb-zone-desc" style="margin-bottom:0;">
                La plateforme <b>MarketBuilding Pro</b> met à la disposition des entreprises du Bâtiment et des Travaux Publics (BTP) un outil d'assistance documentaire permettant d'importer les pièces d'un Dossier de Consultation des Entreprises (<i>Règlement de Consultation, CCTP, DPGF/DQE, PGC et annexes chantier</i>) ainsi que les documents internes de l'entreprise candidate afin de générer une proposition de <b>Mémoire Technique</b> structurée aux formats PDF et Word (<code>.docx</code>).
            </div>
        </div>

        <div class="mb-zone-header" style="border-left:4px solid #F59E0B; margin-bottom:14px;">
            <div class="mb-zone-top">
                <span class="mb-zone-num" style="background:rgba(245,158,11,0.15); color:#F59E0B;">INFORMATION TRANSPARENCE TECHNOLOGIQUE (RÈGLEMENT IA / AI ACT)</span>
            </div>
            <div class="mb-zone-title">Article 2 · Recours à l'Intelligence Artificielle (IA)</div>
            <div class="mb-zone-desc" style="margin-bottom:0;">
                L'Utilisateur est expressément informé et reconnaît que le fonctionnement de la plateforme <b>MarketBuilding Pro</b> repose sur des <b>systèmes d'Intelligence Artificielle (IA) générative et d'analyse sémantique automatisée</b>.<br><br>
                • <b>Traitement par Intelligence Artificielle :</b> Les algorithmes d'IA analysent, extraient et croisent automatiquement les données textuelles et chiffrées issues des documents déposés (critères du RC, prescriptions des CCTP, quantités de la DPGF, contraintes du PGC et ressources de l'entreprise) pour rédiger le contenu du mémoire technique.<br>
                • <b>Règle Anti-Invention et balises <code>[À COMPLÉTER]</code> :</b> Bien que l'IA soit paramétrée avec une consigne stricte lui interdisant d'inventer des données absentes des pièces fournies (en insérant des marqueurs explicites <code>[À COMPLÉTER : préciser ici...]</code>), tout système d'Intelligence Artificielle comporte des limites inhérentes et peut produire des approximations, omissions ou erreurs d'interprétation technique.
            </div>
        </div>

        <div class="mb-zone-header" style="border-left:4px solid #10B981; margin-bottom:14px;">
            <div class="mb-zone-title">Article 3 · Obligation de Relecture et Validation Humaine par l'Utilisateur</div>
            <div class="mb-zone-desc" style="margin-bottom:0;">
                Les documents générés par l'Intelligence Artificielle constituent des <b>projets préparatoires d'aide à la rédaction</b> et ne sauraient remplacer l'expertise technique du soumissionnaire.<br><br>
                Avant toute remise d'offre auprès d'un Maître d'Ouvrage public ou privé, l'Utilisateur s'engage impérativement à :<br>
                1. <b>Relire intégralement</b> le mémoire technique généré (PDF et Word) ;<br>
                2. <b>Renseigner toutes les balises</b> <code>[À COMPLÉTER : ...]</code> signalées par le système ;<br>
                3. <b>Vérifier l'exactitude technique, réglementaire et quantitative</b> (normes DTU, cadences, effectifs, matériel, sécurité PGC, filières déchets SOSED) des engagements pris dans le document final.<br><br>
                L'entreprise soumissionnaire demeure seule et unique responsable du contenu de l'offre qu'elle dépose dans le cadre d'un marché.
            </div>
        </div>

        <div class="mb-zone-header" style="border-left:4px solid #3B82F6; margin-bottom:14px;">
            <div class="mb-zone-title">Article 4 · Confidentialité des Pièces et Propriété Intellectuelle</div>
            <div class="mb-zone-desc" style="margin-bottom:0;">
                • <b>Propriété des documents :</b> L'Utilisateur conserve l'entière propriété intellectuelle des pièces qu'il importe sur la plateforme ainsi que des mémoires techniques générés pour son compte.<br>
                • <b>Confidentialité :</b> Les pièces de marché (DCE) et les justificatifs d'entreprise (CV, organigrammes, parc matériel, certifications) sont traités de manière strictement confidentielle aux seules fins de générer les livrables demandés par l'Utilisateur.
            </div>
        </div>

        <div class="mb-zone-header" style="border-left:4px solid #8B5CF6; margin-bottom:14px;">
            <div class="mb-zone-title">Article 5 · Limitation de Responsabilité et Droit Applicable</div>
            <div class="mb-zone-desc" style="margin-bottom:0;">
                L'éditeur de la plateforme <b>MarketBuilding Pro</b> est tenu à une obligation de moyens dans la fourniture du service d'assistance rédactionnelle par Intelligence Artificielle. Sa responsabilité ne pourra en aucun cas être engagée en cas de rejet d'une candidature ou d'une offre, de perte de chance de remporter un marché, de pénalités d'exécution ou d'inexactitude non corrigée par l'Utilisateur lors de sa relecture.<br><br>
                Les présentes CGU sont soumises au droit français. L'utilisation de la plateforme vaut acceptation pleine et entière des présentes Conditions Générales d'Utilisation.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Pied de page discret rappelant les CGU sur toutes les pages
st.markdown(
    "<div style='text-align:center; margin-top:36px; padding-top:14px; border-top:1px solid var(--border-color); color:var(--text-muted); font-size:12px;'>"
    "MarketBuilding Pro · Plateforme soumise aux <b>Conditions Générales d'Utilisation (CGU)</b> (consultables dans le menu de navigation)."
    "</div>",
    unsafe_allow_html=True,
)
