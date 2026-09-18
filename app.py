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
    page_title="Mémoire Technique BTP – Agent IA",
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
        st.text_input(
            "🔒 Veuillez entrer le mot de passe pour accéder à l'application :",
            type="password",
            on_change=password_entered,
            key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "🔒 Veuillez entrer le mot de passe pour accéder à l'application :",
            type="password",
            on_change=password_entered,
            key="password"
        )
        st.error("😕 Mot de passe incorrect")
        return False
    else:
        # Password correct.
        return True

if not check_password():
    st.stop()



# Charger la config (secrets → config.json → défauts)
config = load_config()

# ─────────────────────────────────────────────────────────────────────────
# 2. BARRE LATERALE – Navigation par onglets
# ─────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/color/96/building-with-rooftop-terrace.png", width=80)
    st.title("🏗️ Mon Assistant BTP")

    page = st.radio(
        "Navigation",
        options=["📄 Générer un Mémoire", "🏢 Mon Entreprise", "⚙️ Réglages"],
        index=0,
        label_visibility="collapsed",
    )

    st.divider()

    # Indicateurs de statut rapide
    entreprise_ok = bool(config["entreprise"].get("nom", "").strip())
    api_ok = bool(config["api_key"].strip())

    st.markdown("**Statut de la configuration**")
    st.markdown(f"{'✅' if api_ok else '❌'} Clé API configurée")
    st.markdown(f"{'✅' if entreprise_ok else 'ℹ️'} Profil entreprise {'renseigné' if entreprise_ok else '(optionnel)'}")

    st.divider()
    st.caption("Agent IA Assistant de Chantier v2.1")
    st.caption("Hébergé sur Streamlit Cloud · Snowflake")


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

RÈGLE D'OR : Tu dois être OMNISCIENT. Même si le profil de l'entreprise est
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
            "❌ Clé API OpenAI invalide. Allez dans ⚙️ Réglages pour la corriger."
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
        raise ValueError("❌ Clé API DeepSeek invalide. Vérifiez vos réglages (ou vos secrets).")
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
                "❌ Clé API Google Gemini invalide. Allez dans ⚙️ Réglages pour la corriger."
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


# ═══════════════════════════════════════════════════════════════════════
#  PAGES DE L'APPLICATION
# ═══════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────
# PAGE : ⚙️ RÉGLAGES
# ─────────────────────────────────────────────────────────────────────────
if page == "⚙️ Réglages":
    st.title("⚙️ Réglages")
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
            "Fournisseur LLM",
            options=["OpenAI (GPT-4o)", "Google Gemini (gemini-2.0-flash)", "DeepSeek (deepseek-chat)"],
            index=0 if "openai" in config["llm_provider"].lower() else (2 if "deepseek" in config["llm_provider"].lower() else 1),
            help="Choisissez le modèle IA à utiliser.",
        )

        new_api_key = st.text_input(
            "🔑 Clé API",
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
            '# Clé API et fournisseur\n'
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
elif page == "🏢 Mon Entreprise":
    st.title("🏢 Mon Entreprise")
    st.markdown(
        "Renseignez les informations de votre entreprise. "
        "Elles seront automatiquement intégrées dans chaque mémoire technique.\n\n"
        "💡 **En cloud** : ces données sont conservées pour la session en cours. "
        "Pour les rendre permanentes, copiez-les dans les [Secrets Streamlit Cloud]"
        "(voir ⚙️ Réglages)."
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
        with st.expander("👁️ Aperçu du profil tel qu'il sera transmis à l'IA"):
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
elif page == "📄 Générer un Mémoire":
    st.title("🏗️ Générateur de Mémoire Technique")

    nom_ent = config["entreprise"].get("nom", "").strip()
    if nom_ent:
        st.markdown(f"Bienvenue **{nom_ent}** · Uploadez votre CCTP et générez votre mémoire technique en un clic.")
    else:
        st.markdown("Uploadez votre **CCTP** et l'IA générera un **Mémoire Technique** complet et prêt à l'emploi.")

    if not config["api_key"].strip():
        st.warning("⚠️ **Clé API manquante.** Allez dans **⚙️ Réglages** pour la configurer.")

    st.divider()

    st.subheader("📄 CCTP de l'appel d'offres")
    cctp_file = st.file_uploader(
        "Uploadez le CCTP (PDF)",
        type=["pdf"],
        key="cctp",
        help="Le cahier des charges techniques de l'appel d'offres au format PDF.",
    )

    st.divider()

    generate_btn = st.button(
        "🚀 Générer le Mémoire Technique",
        type="primary",
        use_container_width=True,
    )

    if generate_btn:
        api_key = config["api_key"].strip()
        llm_provider = config["llm_provider"]

        if not api_key:
            st.error("❌ Clé API manquante. Allez dans **⚙️ Réglages** pour la configurer.")
            st.stop()

        if not cctp_file:
            st.error("❌ Veuillez uploader le fichier CCTP (PDF).")
            st.stop()

        # ── Extraction du PDF ───────────────────────────────────────────
        with st.status("📖 Extraction du texte du CCTP…", expanded=True) as status:
            try:
                st.write("Lecture du CCTP…")
                cctp_text = extract_text_from_pdf(cctp_file)
                nb_pages_cctp = cctp_text.count("--- Page ")
                nb_chars_cctp = len(cctp_text)
                st.write(f"✅ CCTP : {nb_pages_cctp} pages extraites ({nb_chars_cctp:,} caractères)")
                status.update(label="✅ Extraction terminée", state="complete")
            except ValueError as e:
                status.update(label="❌ Erreur d'extraction", state="error")
                st.error(str(e))
                st.stop()

        profil_text = build_profil_text(config["entreprise"])

        if profil_text:
            st.info(f"🏢 Profil **{config['entreprise'].get('nom', '')}** utilisé automatiquement.")
        else:
            st.info("🤖 **Mode autonome** : l'IA générera un mémoire complet en se basant uniquement sur le CCTP.")

        # ── Tâche A : Analyse du CCTP ──────────────────────────────────
        with st.status("🧠 Analyse IA du CCTP en cours…", expanded=True) as status:
            try:
                nb_chunks = len(chunk_text(cctp_text))
                if nb_chunks > 1:
                    st.write(f"Document volumineux → découpage en {nb_chunks} parties pour l'analyse.")

                st.write("Envoi au LLM pour extraction des exigences…")
                analyse = analyse_cctp(api_key, llm_provider, cctp_text)
                st.write("✅ Analyse terminée.")
                status.update(label="✅ Analyse du CCTP terminée", state="complete")
            except ValueError as e:
                status.update(label="❌ Erreur d'analyse", state="error")
                st.error(str(e))
                st.stop()

        with st.expander("📊 Voir l'analyse extraite du CCTP", expanded=False):
            st.json(analyse)

        # ── Tâche B : Rédaction du Mémoire ──────────────────────────────
        with st.status("✍️ Rédaction du Mémoire Technique…", expanded=True) as status:
            try:
                st.write("Le LLM rédige votre mémoire technique personnalisé…")
                memoire_md = rediger_memoire(api_key, llm_provider, analyse, profil_text)
                st.write("✅ Rédaction terminée.")
                status.update(label="✅ Mémoire Technique rédigé", state="complete")
            except ValueError as e:
                status.update(label="❌ Erreur de rédaction", state="error")
                st.error(str(e))
                st.stop()

        with st.expander("📝 Aperçu du Mémoire Technique (Markdown)", expanded=True):
            st.markdown(memoire_md)

        # ── Génération du fichier Word ──────────────────────────────────
        with st.status("📄 Génération du fichier Word…", expanded=True) as status:
            try:
                docx_buffer = generate_docx(memoire_md, analyse, config["entreprise"])
                status.update(label="✅ Fichier Word généré", state="complete")
            except Exception as e:
                status.update(label="❌ Erreur de génération Word", state="error")
                st.error(f"Erreur lors de la génération du fichier Word : {e}")
                st.stop()

        # ── Bouton de téléchargement ────────────────────────────────────
        st.divider()
        st.success("🎉 Votre Mémoire Technique est prêt !")

        nom_fichier = re.sub(r"[^\w\s-]", "", analyse.get("nom_projet", "Memoire_Technique"))
        nom_fichier = nom_fichier.strip().replace(" ", "_")[:60]

        st.download_button(
            label="📥 Télécharger le Mémoire Technique (.docx)",
            data=docx_buffer,
            file_name=f"Memoire_Technique_{nom_fichier}_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

        st.caption(
            "💡 Ouvrez le fichier dans Word, personnalisez avec vos logos et photos, "
            "puis générez le sommaire automatique (onglet Références → Table des matières)."
        )
