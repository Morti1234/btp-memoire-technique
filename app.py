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
# 1. CONFIGURATION DE LA PAGE STREAMLIT
# ─────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MarketBuilding Pro - Mémoire Technique",
    page_icon="🏗️",
    layout="wide",
)


# ─────────────────────────────────────────────────────────────────────────
# 0.5 SÉCURITÉ - ACCÈS PRIVÉ PAR MOT DE PASSE
# ─────────────────────────────────────────────────────────────────────────
def check_password():
    """Returns `True` if the user had the correct password."""
    # Si aucun mot de passe n'est configuré dans les secrets, on laisse passer (pour le dev local)
    if "app_password" not in st.secrets:
        return True

    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.markdown('<div class="password-screen">', unsafe_allow_html=True)
        st.text_input(
            "🔑 Accès sécurisé :",
            placeholder="Entrez votre mot de passe...",
            type="password",
            on_change=password_entered,
            key="password"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.markdown('<div class="password-screen">', unsafe_allow_html=True)
        st.text_input(
            "🔑 Accès sécurisé :",
            placeholder="Entrez votre mot de passe...",
            type="password",
            on_change=password_entered,
            key="password"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        st.error("😕 Mot de passe incorrect")
        return False
    else:
        # Password correct.
        return True

if not check_password():
    _feat = vt.featured_video()
    if _feat:
        st.markdown(
            "<div style='text-align:center; margin-top:16px; margin-bottom:10px;'>"
            "<span style='color:#94A3B8; font-size:14px; letter-spacing:1px; text-transform:uppercase;'>"
            "Démonstration : Du CCTP au téléchargement du Mémoire Technique"
            "</span></div>",
            unsafe_allow_html=True,
        )
        st.video(str(_feat), autoplay=True, loop=True, muted=True)
        st.markdown(
            """
            <div style='display:grid; grid-template-columns:repeat(3, 1fr); gap:12px; margin-top:14px;'>
                <div style='background:#12151b; border:1px solid #262b36; border-left:3px solid #3B82F6; border-radius:10px; padding:12px;'>
                    <div style='color:#3B82F6; font-weight:700; font-size:13px;'>ÉTAPE 1 · IMPORT DU CCTP</div>
                    <div style='color:#cbd5e1; font-size:12.5px; margin-top:4px;'>Déposez votre fichier <b>CCTP (PDF)</b> et cliquez sur <b>« Générer le Mémoire Technique »</b>.</div>
                </div>
                <div style='background:#12151b; border:1px solid #262b36; border-left:3px solid #F59E0B; border-radius:10px; padding:12px;'>
                    <div style='color:#F59E0B; font-weight:700; font-size:13px;'>ÉTAPE 2 · TÉLÉCHARGEMENT DIRECT</div>
                    <div style='color:#cbd5e1; font-size:12.5px; margin-top:4px;'>Une fois rédigé, descendez en bas du document et cliquez sur <b>« 📥 Télécharger en PDF »</b> ou <b>« 📝 Télécharger en Word (.docx) »</b>.</div>
                </div>
                <div style='background:#12151b; border:1px solid #262b36; border-left:3px solid #10B981; border-radius:10px; padding:12px;'>
                    <div style='color:#10B981; font-weight:700; font-size:13px;'>ÉTAPE 3 · APERÇU PDF & ARCHIVES</div>
                    <div style='color:#cbd5e1; font-size:12.5px; margin-top:4px;'>Prévisualisez les pages du PDF officiel directement à l'écran (icône <b>⬇️</b> en haut à droite) ou retrouvez-le dans l'onglet <b>Archives</b>.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.stop()



# Charger la config (secrets → config.json → défauts)

# ─────────────────────────────────────────────────────────────────────────
# 0.7 INJECTION CSS (MODERNISATION & MINIMALISME)
# ─────────────────────────────────────────────────────────────────────────
def inject_custom_css():
    st.markdown('''
        <style>
            /* Masquer les éléments par défaut de Streamlit (Menu, Footer, Header) */
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            
            /* Conserver le header mais masquer le bouton Deploy et le menu droite */
            .stAppDeployButton {display: none;}
            [data-testid="stHeader"] {background: transparent;}

            
            /* Réduire la marge en haut de l'écran */
            .block-container {
                padding-top: 2rem !important;
                padding-bottom: 2rem !important;
                max-width: 1000px;
            }
            
            /* Améliorer les bordures et ombres des conteneurs (minimalisme) */
            div[data-testid="stForm"] {
                border-radius: 12px;
                border: 1px solid #333;
                background-color: #12151b;
            }
            
            /* Styliser l'écran de mot de passe pour faire "Login Screen" */
            .password-screen {
                max-width: 400px;
                margin: 100px auto;
                text-align: center;
            }
        </style>
    ''', unsafe_allow_html=True)

inject_custom_css()

MARKETBUILDING_LOGO_SVG = '''
<div style="text-align:center; padding: 10px 0;">
<svg width="210" height="95" viewBox="0 0 210 95" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="skyGrad" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" style="stop-color:#3B82F6;stop-opacity:1" />
      <stop offset="100%" style="stop-color:#1E3A5F;stop-opacity:1" />
    </linearGradient>
  </defs>

  <!-- Bâtiment 1 (petit, gauche) -->
  <rect x="12" y="32" width="16" height="36" rx="1" fill="#3B82F6"/>
  <rect x="15" y="36" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="21" y="36" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="15" y="43" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="21" y="43" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="15" y="50" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="21" y="50" width="4" height="4" rx="0.5" fill="#0e1117"/>

  <!-- Bâtiment 2 (grand, centre) -->
  <rect x="31" y="14" width="22" height="54" rx="1" fill="#2563EB"/>
  <rect x="35" y="18" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="42" y="18" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="35" y="25" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="42" y="25" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="35" y="32" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="42" y="32" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="35" y="39" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="42" y="39" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="35" y="46" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="42" y="46" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="35" y="53" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="42" y="53" width="4" height="4" rx="0.5" fill="#0e1117"/>

  <!-- Bâtiment 3 (moyen, droite) -->
  <rect x="56" y="26" width="18" height="42" rx="1" fill="#60A5FA"/>
  <rect x="59" y="30" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="66" y="30" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="59" y="37" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="66" y="37" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="59" y="44" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="66" y="44" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="59" y="51" width="4" height="4" rx="0.5" fill="#0e1117"/>
  <rect x="66" y="51" width="4" height="4" rx="0.5" fill="#0e1117"/>

  <!-- Grue de chantier -->
  <line x1="78" y1="68" x2="78" y2="6" stroke="#F59E0B" stroke-width="2.5"/>
  <line x1="78" y1="6" x2="56" y2="6" stroke="#F59E0B" stroke-width="2"/>
  <line x1="78" y1="6" x2="88" y2="6" stroke="#F59E0B" stroke-width="2"/>
  <!-- Câble de la grue -->
  <line x1="60" y1="6" x2="60" y2="18" stroke="#F59E0B" stroke-width="1" stroke-dasharray="2,1"/>
  <!-- Contrepoids -->
  <rect x="84" y="4" width="6" height="5" rx="1" fill="#D97706"/>
  <!-- Base grue -->
  <rect x="73" y="64" width="10" height="4" rx="1" fill="#D97706"/>

  <!-- Sol -->
  <line x1="5" y1="68" x2="90" y2="68" stroke="#475569" stroke-width="1.5"/>

  <!-- Typographie -->
  <text x="100" y="36" font-family="'Segoe UI', 'Helvetica Neue', Arial, sans-serif" font-size="19" font-weight="700" fill="#ffffff" letter-spacing="0.5">Market</text>
  <text x="100" y="56" font-family="'Segoe UI', 'Helvetica Neue', Arial, sans-serif" font-size="19" font-weight="300" fill="#94A3B8" letter-spacing="0.5">Building</text>
  <!-- Ligne déco -->
  <line x1="100" y1="63" x2="200" y2="63" stroke="#F59E0B" stroke-width="1.5" opacity="0.6"/>
  <!-- Tagline -->
  <text x="100" y="76" font-family="'Segoe UI', 'Helvetica Neue', Arial, sans-serif" font-size="7" fill="#64748B" letter-spacing="1.5">MÉMOIRES TECHNIQUES BTP</text>
</svg>
</div>
'''



config = load_config()

# ─────────────────────────────────────────────────────────────────────────
# 2. BARRE LATERALE – Navigation par onglets
# ─────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(MARKETBUILDING_LOGO_SVG, unsafe_allow_html=True)
    # st.title("MarketBuilding Pro")

    page = st.radio(
        "Navigation",
        options=["Génération", "Profil", "Archives", "Vidéos", "Paramètres"],
        index=0,
        label_visibility="collapsed",
    )

    st.divider()

    # Indicateurs de statut rapide
    entreprise_ok = bool(config["entreprise"].get("nom", "").strip())
    api_ok = bool(config["api_key"].strip())

    st.markdown("**Statut de la configuration**")
    st.markdown(f"{'✅' if api_ok else '❌'} Clé d'authentification système configurée")
    st.markdown(f"{'✅' if entreprise_ok else 'ℹ️'} Profil entreprise {'renseigné' if entreprise_ok else '(optionnel)'}")

    st.divider()
    st.caption("MarketBuilding Engine v3.0")
    st.caption("Système Sécurisé")


# ─────────────────────────────────────────────────────────────────────────
# 3. FONCTIONS D'EXTRACTION PDF
# ─────────────────────────────────────────────────────────────────────────

def extract_text_from_pdf(uploaded_file) -> str:
    """
    Extrait le texte brut d'un PDF uploadé via Streamlit.
    Utilise pdfplumber pour un parsing robuste (gère les tableaux, colonnes…).
    Lève une ValueError si le PDF est scanné / ne contient pas de texte.
    """
    text_pages: list[str] = []
    try:
        with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text_pages.append(f"--- Page {i + 1} ---\n{page_text}")
    except Exception as exc:
        raise ValueError(
            f"Impossible de lire le fichier PDF. Vérifiez qu'il n'est pas corrompu. "
            f"Détail : {exc}"
        ) from exc

    if not text_pages:
        raise ValueError(
            "Le PDF uploadé ne contient aucun texte extractible. "
            "Il s'agit probablement d'un document scanné (image). "
            "Veuillez fournir un PDF avec du texte sélectionnable (non scanné) "
            "ou effectuer un OCR au préalable."
        )

    full_text = "\n\n".join(text_pages)
    return full_text


def chunk_text(text: str, max_chars: int = 60_000) -> list[str]:
    """
    Découpe un texte long en morceaux (chunks) de taille maîtrisée.
    Chaque chunk fait au maximum `max_chars` caractères.
    Le découpage se fait par page (marqueur '--- Page') pour conserver le contexte.
    """
    if len(text) <= max_chars:
        return [text]

    pages = re.split(r"(?=--- Page \d+)", text)
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
# 4. FONCTIONS D'APPEL AU LLM (OpenAI / Gemini)
# ─────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT_ANALYSE = textwrap.dedent("""\
Tu es un expert en marchés publics et privés du BTP en France.
On te fournit le texte d'un Cahier des Clauses Techniques Particulières (CCTP).

Ta mission : extraire TOUTES les informations clés sous forme de JSON structuré.

Le JSON doit contenir EXACTEMENT les clés suivantes :
{
  "nom_projet": "Nom ou intitulé du projet / marché",
  "maitre_ouvrage": "Nom du maître d'ouvrage (commanditaire)",
  "lieu_execution": "Adresse ou localisation des travaux",
  "dates_cles": {
    "date_remise_offres": "Date limite de remise des offres (si mentionnée)",
    "delai_execution": "Durée ou délai d'exécution des travaux",
    "date_debut_prevue": "Date de début prévue (si mentionnée)"
  },
  "lots_concernes": ["Liste des lots concernés"],
  "criteres_notation": [
    {"critere": "Nom du critère", "ponderation": "Pondération en %"}
  ],
  "exigences_techniques": ["Liste des exigences techniques majeures"],
  "materiaux_requis": [
    {"materiau": "Nom du matériau", "specification": "Détail / norme / référence"}
  ],
  "ressources_humaines": [
    {"profil": "Type de profil requis", "detail": "Qualification, certification…"}
  ],
  "normes_certifications": ["Liste des normes et certifications exigées (NF, CE, RGE, Qualibat…)"],
  "contraintes_securite_environnement": ["Contraintes SSE mentionnées"]
}

IMPORTANT :
- Sois exhaustif, n'oublie aucun matériau ni aucune exigence.
- Si une information n'est pas mentionnée, mets "Non spécifié".
- Réponds UNIQUEMENT avec le JSON, sans texte avant ni après.
""")

SYSTEM_PROMPT_REDACTION = textwrap.dedent("""\
Tu es un rédacteur technique expert en mémoires techniques pour les marchés publics
et privés du BTP en France. Tu rédiges de manière professionnelle, structurée et
convaincante pour maximiser la note technique du dossier.

On te fournit :
1. L'analyse structurée (JSON) du CCTP.
2. Le profil de l'entreprise candidate (si disponible).

Ta mission : rédiger un MÉMOIRE TECHNIQUE complet, professionnel et PRÊT À L'EMPLOI.

RÈGLE D'OR ABSOLUE :
1. ADAPTATION TOTALE AU CLIENT : Tu dois te baser en priorité sur les informations fournies dans la section "Profil de l'entreprise candidate". Le mémoire technique doit donner l'impression d'avoir été écrit SPÉCIFIQUEMENT pour cette entreprise (utilise son vrai nom, son adresse, ses spécificités si elles sont fournies).
2. PAS DE RÉSUMÉ : Tu ne DOIS PAS faire un simple résumé du CCTP. Tu DOIS rédiger une VÉRITABLE OFFRE TECHNIQUE ET COMMERCIALE du point de vue de l''entreprise candidate qui répond à l'appel d'offres.
- Utilise systématiquement le pronom "Nous" ou "Notre entreprise".
- Argumente, explique COMMENT l'entreprise va réaliser les travaux, avec quelles méthodes, pour prouver son expertise.
- Sois très convaincant, professionnel, détaillé et orienté "solutions". C'est un document pour GAGNER un marché.
- Tu dois être OMNISCIENT : Même si le profil de l'entreprise est
incomplet ou absent, tu dois rédiger un mémoire technique complet et crédible.
Pour les informations manquantes de l'entreprise :
- Rédige des sections génériques mais professionnelles et réalistes.
- Propose des contenus types qu'une PME du BTP compétente inclurait.
- Utilise des formulations comme "Notre entreprise", "Notre équipe", "Nos références".
- NE METS JAMAIS de marqueurs [À COMPLÉTER]. Le document doit être prêt à être lu.
- Déduis intelligemment le type d'entreprise nécessaire à partir du CCTP
  (plomberie, électricité, gros œuvre, etc.) et rédige en conséquence.

STRUCTURE OBLIGATOIRE du mémoire (utilise ces titres exacts) :

# 1. PRÉSENTATION DE L'ENTREPRISE
   - Présentation générale (utilise les infos du profil si fournies, sinon rédige
     une présentation type cohérente avec les lots du CCTP)
   - Moyens humains et organigramme de chantier
   - Références et expériences similaires (propose des références crédibles
     en rapport avec le type de projet du CCTP)
   - Certifications et qualifications (RGE, Qualibat… adaptées au CCTP)

# 2. COMPRÉHENSION DU PROJET
   - Analyse du contexte et des enjeux
   - Synthèse des exigences du CCTP

# 3. MÉTHODOLOGIE ET ORGANISATION DU CHANTIER
   - Phasage des travaux (avec planning indicatif)
   - Méthodologie technique détaillée pour chaque lot
   - Gestion des interfaces entre corps d'état

# 4. MOYENS HUMAINS AFFECTÉS AU CHANTIER
   - Organigramme de l'équipe dédiée
   - Qualifications du personnel affecté
   - Plan de formation si nécessaire

# 5. MOYENS MATÉRIELS ET MATÉRIAUX
   - Liste du matériel et équipements mobilisés
   - Fiches techniques des matériaux proposés (en utilisant ceux du CCTP)
   - Justification des choix techniques

# 6. GESTION DE LA QUALITÉ
   - Plan d'Assurance Qualité (PAQ)
   - Procédures de contrôle et autocontrôle
   - Gestion des non-conformités

# 7. SÉCURITÉ ET PRÉVENTION DES RISQUES
   - Plan Particulier de Sécurité et Protection de la Santé (PPSPS)
   - Analyse des risques et mesures de prévention
   - EPI et formations sécurité

# 8. GESTION ENVIRONNEMENTALE
   - Plan de gestion des déchets (SOGED)
   - Mesures de réduction de l'impact environnemental
   - Engagements développement durable

# 9. PLANNING PRÉVISIONNEL
   - Macro-planning des travaux
   - Jalons et livrables

CONSIGNES DE RÉDACTION :
- Sois concret, précis et professionnel.
- Utilise les matériaux et normes extraits du CCTP (fournis dans le JSON).
- Personnalise le mémoire avec les infos de l'entreprise SI elles sont fournies.
- Si les infos ne sont pas fournies, rédige un contenu professionnel et réaliste.
- NE METS JAMAIS de marqueurs [À COMPLÉTER] ou de crochets vides.
- Vise un document de 3 000 à 5 000 mots.
- Le mémoire doit pouvoir être imprimé et soumis tel quel.
""")


def _call_openai(api_key: str, system_prompt: str, user_content: str, model: str = "gpt-4o") -> str:
    """Appelle l'API OpenAI Chat Completions."""
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
            "❌ Clé d'authentification système OpenAI invalide. Allez dans Paramètres pour la corriger."
        )
    except APIError as e:
        raise ValueError(f"❌ Erreur API OpenAI : {e.message}")




def _call_deepseek(api_key: str, system_prompt: str, user_content: str) -> str:
    """Appelle l'API DeepSeek via la librairie openai."""
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
        raise ValueError("❌ Clé d'authentification système DeepSeek invalide. Vérifiez vos réglages (ou vos secrets).")
    except Exception as e:
        raise ValueError(f"❌ Erreur DeepSeek : {e}")

def _call_gemini(api_key: str, system_prompt: str, user_content: str, model: str = "gemini-2.0-flash") -> str:
    """Appelle l'API Google Gemini."""
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
                "❌ Clé d'authentification système Google Gemini invalide. Allez dans Paramètres pour la corriger."
            )
        raise ValueError(f"❌ Erreur API Gemini : {e}")


def call_llm(api_key: str, provider: str, system_prompt: str, user_content: str) -> str:
    """Dispatch vers le bon fournisseur LLM."""
    if "openai" in provider.lower():
        return _call_openai(api_key, system_prompt, user_content)
    elif "deepseek" in provider.lower():
        return _call_deepseek(api_key, system_prompt, user_content)
    else:
        return _call_gemini(api_key, system_prompt, user_content)


# ─────────────────────────────────────────────────────────────────────────
# 5. TÂCHE A : ANALYSE DU CCTP → JSON
# ─────────────────────────────────────────────────────────────────────────

def analyse_cctp(api_key: str, provider: str, cctp_text: str) -> dict:
    """
    Envoie le texte du CCTP au LLM pour extraire les informations clés.
    Gère le chunking si le document est trop long.
    """
    chunks = chunk_text(cctp_text)
    all_analyses: list[dict] = []

    for i, chunk in enumerate(chunks):
        if len(chunks) > 1:
            user_msg = (
                f"[Partie {i + 1}/{len(chunks)} du CCTP]\n\n"
                f"{chunk}\n\n"
                f"Analyse cette partie et extrais les informations au format JSON demandé. "
                f"Si certaines informations ne sont pas dans cette partie, mets 'Non spécifié dans cette section'."
            )
        else:
            user_msg = f"Voici le CCTP complet à analyser :\n\n{chunk}"

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
                    "❌ Le LLM n'a pas retourné un JSON valide après deux tentatives. "
                    "Essayez de réduire la taille du PDF ou changez de fournisseur LLM."
                )

    if len(all_analyses) == 1:
        return all_analyses[0]

    merged = all_analyses[0]
    for analysis in all_analyses[1:]:
        for key, value in analysis.items():
            if isinstance(value, list) and isinstance(merged.get(key), list):
                existing = [json.dumps(item, ensure_ascii=False) for item in merged[key]]
                for item in value:
                    if json.dumps(item, ensure_ascii=False) not in existing:
                        merged[key].append(item)
            elif isinstance(value, dict) and isinstance(merged.get(key), dict):
                for k, v in value.items():
                    if merged[key].get(k) in (None, "", "Non spécifié", "Non spécifié dans cette section"):
                        merged[key][k] = v
            elif merged.get(key) in (None, "", "Non spécifié", "Non spécifié dans cette section"):
                merged[key] = value

    return merged


# ─────────────────────────────────────────────────────────────────────────
# 6. TÂCHE B : RÉDACTION DU MÉMOIRE TECHNIQUE
# ─────────────────────────────────────────────────────────────────────────

def rediger_memoire(
    api_key: str,
    provider: str,
    analyse_json: dict,
    profil_entreprise: Optional[str] = None,
) -> str:
    """
    Demande au LLM de rédiger le mémoire technique complet
    à partir de l'analyse JSON du CCTP et du profil entreprise.
    """
    user_content = f"## Analyse du CCTP (JSON)\n```json\n{json.dumps(analyse_json, ensure_ascii=False, indent=2)}\n```\n\n"

    if profil_entreprise:
        user_content += f"## Profil de l'entreprise candidate\n{profil_entreprise}\n\n"
    else:
        user_content += (
            "## Profil de l'entreprise candidate\n"
            "Aucun profil spécifique fourni. Tu dois rédiger le mémoire technique de manière "
            "autonome et omnisciente. Déduis le type d'entreprise nécessaire à partir du CCTP "
            "(corps de métier, taille probable, qualifications requises) et rédige un profil "
            "d'entreprise crédible et professionnel. Utilise 'Notre entreprise' comme nom. "
            "Le mémoire doit être complet et prêt à être soumis sans aucune modification.\n\n"
        )

    user_content += "Rédige maintenant le Mémoire Technique complet en suivant la structure demandée."

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

    # Infos entreprise sur la page de garde
    nom_entreprise = entreprise.get("nom", "").strip()
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
            parts = re.split(r"(\*\*.*?\*\*)", bullet_text)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)
        elif re.match(r"^\d+\.\s", stripped):
            text = re.sub(r"^\d+\.\s", "", stripped)
            p = doc.add_paragraph(style="List Number")
            parts = re.split(r"(\*\*.*?\*\*)", text)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)
        else:
            p = doc.add_paragraph()
            parts = re.split(r"(\*\*.*?\*\*)", stripped)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)

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

    nom_entreprise = entreprise.get("nom", "").strip() or "MarketBuilding"
    nom_projet = analyse_json.get("nom_projet", "Projet BTP").strip()

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
#  PAGES DE L'APPLICATION
# ═══════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────
# PAGE : ⚙️ RÉGLAGES
# ─────────────────────────────────────────────────────────────────────────
if page == "Paramètres":
    st.title("Paramètres")
    st.markdown(
        "Configurez votre clé API et votre fournisseur LLM.\n\n"
        "💡 **En cloud** : pour une persistance permanente, configurez vos secrets dans le "
        "[dashboard Streamlit Cloud](https://share.streamlit.io/) → votre app → Settings → Secrets."
    )

    st.divider()

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
        st.info("💻 **Mode local** – La configuration est stockée dans `config.json`.")

    with st.form("form_reglages"):
        new_provider = st.selectbox(
            "Moteur de traitement",
            options=["OpenAI (GPT-4o)", "Google Gemini (gemini-2.0-flash)", "DeepSeek (deepseek-chat)"],
            index=0 if "openai" in config["llm_provider"].lower() else (2 if "deepseek" in config["llm_provider"].lower() else 1),
            help="Choisissez l'algorithme d'analyse à utiliser.",
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
            st.success("✅ Réglages sauvegardés pour cette session !")
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
    st.title("Profil")
    st.markdown(
        "Renseignez les informations de votre entreprise. "
        "Elles seront automatiquement intégrées dans chaque mémoire technique.\n\n"
        "💡 **En cloud** : ces données sont conservées pour la session en cours. "
        "Pour les rendre permanentes, copiez-les dans les [Secrets Streamlit Cloud]"
        "(voir Paramètres)."
    )

    st.divider()

    ent = config["entreprise"]

    with st.form("form_entreprise"):
        st.subheader("📋 Identité")

        col1, col2 = st.columns(2)
        with col1:
            ent["nom"] = st.text_input("Nom de l'entreprise *", value=ent.get("nom", ""))
            ent["forme_juridique"] = st.text_input("Forme juridique (SARL, SAS…)", value=ent.get("forme_juridique", ""))
            ent["siret"] = st.text_input("SIRET", value=ent.get("siret", ""))
            ent["dirigeant"] = st.text_input("Dirigeant / Gérant", value=ent.get("dirigeant", ""))
        with col2:
            ent["adresse"] = st.text_input("Adresse", value=ent.get("adresse", ""))
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
            "Activités principales",
            value=ent.get("activites_principales", ""),
            height=80,
            help="Ex : Plomberie, chauffage, climatisation, rénovation énergétique…",
        )
        col5, col6 = st.columns(2)
        with col5:
            ent["certifications"] = st.text_area(
                "Certifications (RGE, ISO…)",
                value=ent.get("certifications", ""),
                height=80,
            )
        with col6:
            ent["qualifications"] = st.text_area(
                "Qualifications (Qualibat, Qualifelec…)",
                value=ent.get("qualifications", ""),
                height=80,
            )

        ent["assurances"] = st.text_area(
            "Assurances (décennale, RC Pro…)",
            value=ent.get("assurances", ""),
            height=80,
        )

        st.subheader("💼 Références & Moyens")
        ent["references_chantiers"] = st.text_area(
            "Références chantiers similaires",
            value=ent.get("references_chantiers", ""),
            height=120,
        )
        ent["moyens_materiels"] = st.text_area(
            "Moyens matériels",
            value=ent.get("moyens_materiels", ""),
            height=100,
        )

        st.subheader("📝 Présentation libre")
        ent["presentation_libre"] = st.text_area(
            "Historique, valeurs, points forts…",
            value=ent.get("presentation_libre", ""),
            height=150,
        )

        submitted = st.form_submit_button("💾 Sauvegarder le profil", type="primary", use_container_width=True)

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
                    # Echapper les guillemets et retours à la ligne pour TOML
                    escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
                    toml_lines.append(f'{key} = "{escaped}"')
            toml_output = "\n".join(toml_lines)
            st.code(toml_output, language="toml")
            st.caption("Copiez ce bloc dans vos Secrets Streamlit Cloud (Settings → Secrets) pour rendre le profil permanent.")


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 📄 GÉNÉRER UN MÉMOIRE
# ─────────────────────────────────────────────────────────────────────────
elif page == "Génération":
    st.title("Générateur de Mémoires Techniques")

    nom_ent = config["entreprise"].get("nom", "").strip()
    if nom_ent:
        st.markdown(f"Bienvenue **{nom_ent}** · Uploadez votre CCTP et générez votre mémoire technique en un clic.")
    else:
        st.markdown("Uploadez votre **CCTP** et le système générera un **Mémoire Technique** complet et prêt à l'emploi.")

    if not config["api_key"].strip():
        st.warning("⚠️ **Clé d'authentification système manquante.** Allez dans **Paramètres** pour la configurer.")

    _feat_main = vt.featured_video()
    if _feat_main:
        with st.expander("🎬 Tutoriel vidéo : Comment générer et télécharger votre Mémoire Technique", expanded=True):
            st.video(str(_feat_main), autoplay=True, loop=True, muted=True)
            st.markdown(
                """
                <div style='display:grid; grid-template-columns:repeat(3, 1fr); gap:12px; margin-top:10px;'>
                    <div style='background:#12151b; border:1px solid #262b36; border-left:3px solid #3B82F6; border-radius:10px; padding:12px;'>
                        <div style='color:#3B82F6; font-weight:700; font-size:13px;'>1. Importer le CCTP (Début de la vidéo)</div>
                        <div style='color:#cbd5e1; font-size:12.5px; margin-top:4px;'>Cliquez sur <b>Upload</b>, sélectionnez votre <b>CCTP (.pdf)</b> puis lancez <b>« 🚀 Générer le Mémoire Technique »</b>.</div>
                    </div>
                    <div style='background:#12151b; border:1px solid #262b36; border-left:3px solid #F59E0B; border-radius:10px; padding:12px;'>
                        <div style='color:#F59E0B; font-weight:700; font-size:13px;'>2. Boutons de téléchargement PDF & Word</div>
                        <div style='color:#cbd5e1; font-size:12.5px; margin-top:4px;'>En bas du mémoire généré, cliquez sur le bouton bleu <b>« 📥 Télécharger en PDF »</b> ou sur <b>« 📝 Télécharger en Word (.docx) »</b>.</div>
                    </div>
                    <div style='background:#12151b; border:1px solid #262b36; border-left:3px solid #10B981; border-radius:10px; padding:12px;'>
                        <div style='color:#10B981; font-weight:700; font-size:13px;'>3. Visionneuse PDF & Onglet Archives</div>
                        <div style='color:#cbd5e1; font-size:12.5px; margin-top:4px;'>Consultez l'<b>Aperçu du document PDF</b> (bouton <b>⬇️</b> intégré au lecteur) ou retrouvez tous vos mémoires dans l'onglet <b>Archives</b>.</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.divider()

    st.subheader("📄 CCTP de l'appel d'offres (Simple ou Multi-CCTP)")
    cctp_files = st.file_uploader(
        "Uploadez un ou plusieurs CCTP (PDF) pour générer plusieurs mémoires en une seule fois",
        type=["pdf"],
        key="cctp",
        accept_multiple_files=True,
        help="Glissez-déposez un seul CCTP ou plusieurs CCTP simultanément pour les traiter en lot.",
    )

    nb_files = len(cctp_files) if cctp_files else 0
    if nb_files > 1:
        st.info(f"📚 **{nb_files} CCTP sélectionnés** : cliquez sur le bouton ci-dessous pour générer les {nb_files} mémoires techniques à la suite automatiquement.")

    st.divider()

    btn_label = (
        f"🚀 Générer les {nb_files} Mémoires Techniques"
        if nb_files > 1
        else "🚀 Générer le Mémoire Technique"
    )
    generate_btn = st.button(
        btn_label,
        type="primary",
        use_container_width=True,
    )

    if generate_btn:
        api_key = config["api_key"].strip()
        llm_provider = config["llm_provider"]

        if not api_key:
            st.error("❌ Clé d'authentification système manquante. Allez dans **Paramètres** pour la configurer.")
            st.stop()

        if not cctp_files:
            st.error("❌ Veuillez uploader au moins un fichier CCTP (PDF).")
            st.stop()

        profil_text = build_profil_text(config["entreprise"])
        if profil_text:
            st.info(f"🏢 Profil **{config['entreprise'].get('nom', '')}** utilisé automatiquement.")
        else:
            st.info("⚙️ **Mode autonome** : le système générera un mémoire complet en se basant uniquement sur chaque CCTP.")

        batch_results = []
        os.makedirs("archives", exist_ok=True)
        batch_progress = st.progress(0.0, text=f"Démarrage du traitement (0/{nb_files})…") if nb_files > 1 else None

        for idx, cctp_file in enumerate(cctp_files, start=1):
            prefix = f"[{idx}/{nb_files}] {cctp_file.name} — " if nb_files > 1 else ""
            if batch_progress:
                batch_progress.progress((idx - 1) / nb_files, text=f"Traitement du dossier {idx}/{nb_files} : {cctp_file.name}…")

            # ── Extraction du PDF ───────────────────────────────────────────
            with st.status(f"📖 {prefix}Extraction du texte du CCTP…", expanded=(nb_files == 1)) as status:
                try:
                    st.write(f"Lecture de **{cctp_file.name}**…")
                    cctp_text = extract_text_from_pdf(cctp_file)
                    nb_pages_cctp = cctp_text.count("--- Page ")
                    nb_chars_cctp = len(cctp_text)
                    st.write(f"✅ CCTP : {nb_pages_cctp} pages extraites ({nb_chars_cctp:,} caractères)")
                    status.update(label=f"✅ {prefix}Extraction terminée", state="complete")
                except ValueError as e:
                    status.update(label=f"❌ {prefix}Erreur d'extraction", state="error")
                    st.error(f"{cctp_file.name} : {e}")
                    continue

            # ── Tâche A : Analyse du CCTP ──────────────────────────────────
            with st.status(f"🔍 {prefix}Analyse sémantique du CCTP en cours…", expanded=(nb_files == 1)) as status:
                try:
                    nb_chunks = len(chunk_text(cctp_text))
                    if nb_chunks > 1:
                        st.write(f"Document volumineux → découpage en {nb_chunks} parties pour l'analyse.")
                    st.write("Extraction algorithmique des exigences techniques…")
                    analyse = analyse_cctp(api_key, llm_provider, cctp_text)
                    st.write("✅ Analyse terminée.")
                    status.update(label=f"✅ {prefix}Analyse du CCTP terminée", state="complete")
                except ValueError as e:
                    status.update(label=f"❌ {prefix}Erreur d'analyse", state="error")
                    st.error(f"{cctp_file.name} : {e}")
                    continue

            # ── Tâche B : Rédaction du Mémoire ──────────────────────────────
            with st.status(f"✍️ {prefix}Rédaction du Mémoire Technique…", expanded=(nb_files == 1)) as status:
                try:
                    st.write("Compilation et formatage du document final personnalisé…")
                    memoire_md = rediger_memoire(api_key, llm_provider, analyse, profil_text)
                    st.write("✅ Rédaction terminée.")
                    status.update(label=f"✅ {prefix}Mémoire Technique rédigé", state="complete")
                except ValueError as e:
                    status.update(label=f"❌ {prefix}Erreur de rédaction", state="error")
                    st.error(f"{cctp_file.name} : {e}")
                    continue

            # ── Génération des fichiers Word et PDF ────────────────────────
            with st.status(f"📄 {prefix}Compilation PDF & Word…", expanded=(nb_files == 1)) as status:
                try:
                    docx_buffer = generate_docx(memoire_md, analyse, config["entreprise"])
                    pdf_buffer = generate_pdf(memoire_md, analyse, config["entreprise"])
                    status.update(label=f"✅ {prefix}Document généré avec succès", state="complete")
                except Exception as e:
                    status.update(label=f"❌ {prefix}Erreur de formatage", state="error")
                    st.error(f"Erreur lors de la génération du document ({cctp_file.name}) : {e}")
                    continue

            nom_projet = analyse.get("nom_projet", "") or Path(cctp_file.name).stem
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
            st.success("🎉 Document généré avec succès.")

            with st.expander("📊 Voir l'analyse extraite du CCTP", expanded=False):
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

            st.caption("💡 Le format Word vous permet d'ajouter le logo de l'entreprise ou des photos de chantier avant envoi.")

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
    st.title("Archives des mémoires")
    st.markdown("Tous les mémoires techniques générés, retéléchargeables à tout moment.")
    st.divider()

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
            c1, c2, c3 = st.columns([3, 1, 1])
            with c1:
                st.markdown(f"**{'📄' if fichier.endswith('.pdf') else '📝'} {fichier}**")
                st.caption(f"Généré le {date_mod} • {os.path.getsize(chemin) / 1024:.0f} Ko")
            with c2:
                with open(chemin, "rb") as f:
                    st.download_button(
                        "Télécharger", data=f.read(), file_name=fichier,
                        mime="application/pdf" if fichier.endswith(".pdf")
                        else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True, key=f"dl_{fichier}",
                    )
            with c3:
                if st.button("Supprimer", key=f"del_{fichier}", use_container_width=True):
                    os.remove(chemin)
                    st.rerun()


# ─────────────────────────────────────────────────────────────────────────
# PAGE : 🎬 VIDÉOS
# ─────────────────────────────────────────────────────────────────────────
elif page == "Vidéos":
    st.title("Vidéos de présentation")
    st.markdown(
        "Importez vos enregistrements (même très lourds, .mov Retina inclus), "
        "ils sont automatiquement optimisés pour le web. Générez ensuite une vidéo de "
        "présentation professionnelle du service en un clic."
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
