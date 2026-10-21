"""
Module d'Automatisation Complet (Pipeline Email -> Site Web -> Email)
pour MarketBuilding Pro - Générateur de Mémoires Techniques BTP.

Fonctionnalités :
1. Écoute automatique de la boîte mail (Worker IMAP en arrière-plan) + décompression récursive des archives .zip DCE.
2. Tri intelligent des fichiers BTP (Auto-Triage IA + heuristique) :
   - Identification de l'entreprise par email expéditeur ou via documents joints
   - Lecture des consignes du mail et détection du/des lot(s) visé(s)
   - Classification automatique en 5 catégories : RC, CCTP (Lot 00 + Lot(s) visé(s) uniquement, ignore les autres lots), DPGF, ANNEXES, ENTREPRISE
3. Déclenchement automatique du générateur de mémoire technique existant et synchronisation en temps réel avec la base de données du site (badge "Automatisé via Mail", lecture, modification Markdown + régénération PDF/Word et téléchargement).
4. Réponse automatique par e-mail (SMTP) avec pièces jointes (.docx et .pdf), liste des pièces croisées, synthèse des 5 points de vigilance majeurs (RC/CCTP) et liste exacte des balises [À COMPLÉTER], ou notification automatique de pièce manquante si aucun CCTP/document exploitable n'est joint.
5. Simulateur intégré de réception d'e-mail (avec création d'une archive .zip DCE multi-lots réaliste) et journal d'événements en direct (Logs).
"""

from __future__ import annotations

import email
from email.header import decode_header
from email.message import EmailMessage
from email.utils import parseaddr
import imaplib
import io
import json
import os
from pathlib import Path
import re
import smtplib
import ssl
import threading
import time
import uuid
import zipfile
from datetime import datetime
from typing import Callable, Optional

from docx import Document
from fpdf import FPDF


AUTOMATION_DB_PATH = Path("archives") / "automation_state.json"
PROJECTS_DB_PATH = Path("archives") / "projects_db.json"
PROFILES_DB_PATH = Path("archives") / "profiles_entreprises.json"
DEFAULT_REFERENCE_EMAIL = "markethit67000@gmail.com"

_DB_LOCK = threading.RLock()
_WORKER_THREAD: Optional[threading.Thread] = None
_WORKER_STOP_EVENT = threading.Event()


class MemoryAttachment:
    """Représente un fichier extrait d'un mail ou d'une archive .zip compatible avec les extracteurs Streamlit."""

    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data
        self._pos = 0

    def getvalue(self) -> bytes:
        return self._data

    def read(self) -> bytes:
        return self._data

    def seek(self, pos: int = 0) -> None:
        self._pos = pos


# ─────────────────────────────────────────────────────────────────────────
# 1. PERSISTANCE DE L'ÉTAT D'AUTOMATISATION, DES PROJETS ET DES LOGS
# ─────────────────────────────────────────────────────────────────────────

def _load_dotenv_if_present() -> dict[str, str]:
    """Charge les variables depuis .env s'il existe, combinées avec os.environ."""
    env_vars: dict[str, str] = {}
    env_file = Path(".env")
    if env_file.exists():
        try:
            for line in env_file.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                env_vars[k.strip()] = v.strip().strip('"').strip("'")
        except Exception:
            pass
    for k in (
        "IMAP_HOST",
        "IMAP_PORT",
        "IMAP_USER",
        "IMAP_PASSWORD",
        "SMTP_HOST",
        "SMTP_PORT",
        "SMTP_USER",
        "SMTP_PASSWORD",
        "SMTP_FROM",
        "AUTOMATION_POLL_SECONDS",
    ):
        if os.environ.get(k):
            env_vars[k] = os.environ[k]
    return env_vars


def load_automation_state() -> dict:
    """Charge l'état de l'automatisation (config IMAP/SMTP, mode auto/manuel, statut actif, jobs, logs)."""
    os.makedirs("archives", exist_ok=True)
    env_vars = _load_dotenv_if_present()

    default_state = {
        "watcher_enabled": True,
        "automation_mode": "auto",  # "auto" (100% Automatique) ou "manual" (Validation Manuelle)
        "poll_interval_seconds": int(env_vars.get("AUTOMATION_POLL_SECONDS", "30") or 30),
        "reference_email": env_vars.get("IMAP_USER", DEFAULT_REFERENCE_EMAIL) or DEFAULT_REFERENCE_EMAIL,
        "imap_host": env_vars.get("IMAP_HOST", "imap.gmail.com"),
        "imap_port": int(env_vars.get("IMAP_PORT", "993") or 993),
        "imap_user": env_vars.get("IMAP_USER", DEFAULT_REFERENCE_EMAIL) or DEFAULT_REFERENCE_EMAIL,
        "imap_password": env_vars.get("IMAP_PASSWORD", ""),
        "smtp_host": env_vars.get("SMTP_HOST", "smtp.gmail.com"),
        "smtp_port": int(env_vars.get("SMTP_PORT", "587") or 587),
        "smtp_user": env_vars.get("SMTP_USER", env_vars.get("IMAP_USER", DEFAULT_REFERENCE_EMAIL)) or DEFAULT_REFERENCE_EMAIL,
        "smtp_password": env_vars.get("SMTP_PASSWORD", env_vars.get("IMAP_PASSWORD", "")),
        "smtp_from": env_vars.get("SMTP_FROM", env_vars.get("IMAP_USER", DEFAULT_REFERENCE_EMAIL)) or DEFAULT_REFERENCE_EMAIL,
        "last_check": None,
        "last_error": None,
        "jobs": [],
        "logs": [],
    }

    with _DB_LOCK:
        if AUTOMATION_DB_PATH.exists():
            try:
                saved = json.loads(AUTOMATION_DB_PATH.read_text(encoding="utf-8"))
                for k, v in default_state.items():
                    if k not in saved:
                        saved[k] = v
                # Toujours garantir que markethit67000@gmail.com est l'email de référence si les champs étaient vides
                if not (saved.get("imap_user") or "").strip():
                    saved["imap_user"] = DEFAULT_REFERENCE_EMAIL
                if not (saved.get("smtp_user") or "").strip():
                    saved["smtp_user"] = saved["imap_user"] or DEFAULT_REFERENCE_EMAIL
                if not (saved.get("smtp_from") or "").strip():
                    saved["smtp_from"] = saved["smtp_user"] or DEFAULT_REFERENCE_EMAIL
                saved["reference_email"] = saved.get("imap_user") or DEFAULT_REFERENCE_EMAIL
                # Si .env fournit des identifiants et que saved est vide, compléter
                for env_k, state_k in [
                    ("IMAP_HOST", "imap_host"),
                    ("IMAP_USER", "imap_user"),
                    ("IMAP_PASSWORD", "imap_password"),
                    ("SMTP_HOST", "smtp_host"),
                    ("SMTP_USER", "smtp_user"),
                    ("SMTP_PASSWORD", "smtp_password"),
                    ("SMTP_FROM", "smtp_from"),
                ]:
                    if env_vars.get(env_k) and not saved.get(state_k):
                        saved[state_k] = env_vars[env_k]
                return saved
            except Exception:
                return default_state
        return default_state


def save_automation_state(state: dict) -> None:
    os.makedirs("archives", exist_ok=True)
    with _DB_LOCK:
        AUTOMATION_DB_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def add_automation_log(statut: str, message: str, email_subject: str = "", sender: str = "", job_id: str = "") -> None:
    """Ajoute une entrée horodatée au journal en direct (Logs) de l'automatisation."""
    with _DB_LOCK:
        state = load_automation_state()
        entry = {
            "id": uuid.uuid4().hex[:8],
            "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "statut": statut,  # Ex: "Mail reçu", "Tri des CCTP en cours", "Génération sur le site", "Mémoire envoyé"
            "message": message,
            "subject": email_subject,
            "sender": sender,
            "job_id": job_id,
        }
        state["logs"].insert(0, entry)
        state["logs"] = state["logs"][:200]
        save_automation_state(state)


def update_job_status(job_id: str, **updates) -> Optional[dict]:
    """Met à jour un job d'e-mail dans automation_state.json."""
    with _DB_LOCK:
        state = load_automation_state()
        for job in state["jobs"]:
            if job.get("job_id") == job_id:
                job.update(updates)
                job["updated_at"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                save_automation_state(state)
                return job
    return None


def load_projects_db() -> list[dict]:
    """Charge la base de données des projets générés sur le site (manuels et automatisés via mail)."""
    os.makedirs("archives", exist_ok=True)
    with _DB_LOCK:
        if PROJECTS_DB_PATH.exists():
            try:
                data = json.loads(PROJECTS_DB_PATH.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return data
            except Exception:
                pass
    return []


def save_project_to_db(project: dict) -> dict:
    """Enregistre ou met à jour un projet dans la base de données du site (archives/projects_db.json)."""
    os.makedirs("archives", exist_ok=True)
    with _DB_LOCK:
        projects = load_projects_db()
        pid = project.get("project_id")
        if not pid:
            pid = uuid.uuid4().hex[:10]
            project["project_id"] = pid
        replaced = False
        for i, existing in enumerate(projects):
            if existing.get("project_id") == pid:
                projects[i] = project
                replaced = True
                break
        if not replaced:
            projects.insert(0, project)
        PROJECTS_DB_PATH.write_text(json.dumps(projects[:100], ensure_ascii=False, indent=2), encoding="utf-8")
    return project


def delete_project_from_db(project_id: str) -> None:
    with _DB_LOCK:
        projects = load_projects_db()
        remaining = []
        for p in projects:
            if p.get("project_id") == project_id:
                for path_key in ("pdf_path", "docx_path"):
                    fp = p.get(path_key)
                    if fp and os.path.exists(fp):
                        try:
                            os.remove(fp)
                        except OSError:
                            pass
            else:
                remaining.append(p)
        PROJECTS_DB_PATH.write_text(json.dumps(remaining, ensure_ascii=False, indent=2), encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────────
# 2. DÉCOMPRESSION AUTOMATIQUE DES DOSSIERS DCE (.ZIP)
# ─────────────────────────────────────────────────────────────────────────

SUPPORTED_DCE_EXTENSIONS = (".pdf", ".docx", ".xlsx", ".xls", ".csv", ".txt")


def unpack_attachments_and_zips(raw_attachments: list[tuple[str, bytes]], _depth: int = 0) -> tuple[list[MemoryAttachment], list[str]]:
    """
    Prend une liste de pièces jointes (filename, bytes).
    Si une pièce jointe est une archive .zip, extrait récursivement tous les fichiers
    PDF, Word (.docx) et Excel (.xlsx, .xls, .csv) qu'elle contient.
    Retourne (liste de MemoryAttachment, journal d'extraction).
    """
    extracted_files: list[MemoryAttachment] = []
    extraction_notes: list[str] = []

    for filename, data in raw_attachments:
        if not filename or not data:
            continue
        clean_name = Path(filename).name
        if clean_name.startswith(".") or clean_name.startswith("__MACOSX") or clean_name.startswith("~$"):
            continue

        suffix = Path(clean_name).suffix.lower()
        if suffix == ".zip" and _depth < 3:
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as zf:
                    nested_pairs: list[tuple[str, bytes]] = []
                    for info in zf.infolist():
                        if info.is_dir():
                            continue
                        member_name = Path(info.filename).name
                        if (
                            not member_name
                            or "__MACOSX" in info.filename
                            or member_name.startswith(".")
                            or member_name.startswith("~$")
                        ):
                            continue
                        member_ext = Path(member_name).suffix.lower()
                        if member_ext in SUPPORTED_DCE_EXTENSIONS or member_ext == ".zip":
                            nested_pairs.append((member_name, zf.read(info.filename)))
                    sub_files, sub_notes = unpack_attachments_and_zips(nested_pairs, _depth=_depth + 1)
                    extracted_files.extend(sub_files)
                    extraction_notes.append(f"📦 Archive ZIP « {clean_name} » décompressée ({len(sub_files)} fichier(s) BTP extrait(s))")
                    extraction_notes.extend(sub_notes)
            except Exception as exc:
                extraction_notes.append(f"⚠️ Impossible de décompresser l'archive « {clean_name} » : {exc}")
        elif suffix in SUPPORTED_DCE_EXTENSIONS:
            extracted_files.append(MemoryAttachment(clean_name, data))
            extraction_notes.append(f"📄 Pièce jointe détectée : {clean_name} ({len(data) / 1024:.1f} Ko)")
        else:
            extraction_notes.append(f"ℹ️ Fichier ignoré (format non BTP) : {clean_name}")

    return extracted_files, extraction_notes


# ─────────────────────────────────────────────────────────────────────────
# 3. IDENTIFICATION ENTREPRISE & TRI INTELLIGENT DES FICHIERS BTP (AUTO-TRIAGE)
# ─────────────────────────────────────────────────────────────────────────

def find_company_profile_by_email(sender_email: str, default_config_entreprise: dict) -> tuple[Optional[dict], str]:
    """
    Cherche l'adresse mail de l'expéditeur dans :
    1. Le profil d'entreprise principal enregistré sur le site (config["entreprise"])
    2. L'annuaire multi-profils d'entreprises (archives/profiles_entreprises.json)
    Retourne (profil_dict ou None, source_description).
    """
    sender_clean = (sender_email or "").strip().lower()
    sender_domain = sender_clean.split("@")[-1] if "@" in sender_clean else ""
    generic_domains = {"gmail.com", "outlook.fr", "outlook.com", "hotmail.fr", "hotmail.com", "orange.fr", "wanadoo.fr", "yahoo.fr", "yahoo.com", "icloud.com", "free.fr", "sfr.fr", "laposte.net"}

    # 1. Vérifier le profil principal du site
    main_email = (default_config_entreprise.get("email") or "").strip().lower()
    main_nom = (default_config_entreprise.get("nom") or "").strip()
    if sender_clean and main_email and sender_clean == main_email:
        return default_config_entreprise, f"Profil principal reconnu par email ({main_nom or main_email})"
    if sender_domain and sender_domain not in generic_domains and main_email.endswith("@" + sender_domain):
        return default_config_entreprise, f"Profil principal reconnu par domaine @{sender_domain} ({main_nom})"

    # 2. Vérifier dans profiles_entreprises.json
    if PROFILES_DB_PATH.exists():
        try:
            profiles = json.loads(PROFILES_DB_PATH.read_text(encoding="utf-8"))
            if isinstance(profiles, list):
                for prof in profiles:
                    p_email = (prof.get("email") or "").strip().lower()
                    if sender_clean and p_email and sender_clean == p_email:
                        return prof, f"Profil enregistré reconnu ({prof.get('nom', p_email)})"
                    if sender_domain and sender_domain not in generic_domains and p_email.endswith("@" + sender_domain):
                        return prof, f"Profil enregistré reconnu par domaine @{sender_domain} ({prof.get('nom', '')})"
        except Exception:
            pass

    # Si le profil principal du site est renseigné et qu'aucun autre profil ne matche
    if main_nom:
        return default_config_entreprise, f"Profil actif du site ({main_nom})"

    return None, "Aucun profil pré-enregistré pour cet email (utilisation des pièces Entreprise jointes au mail)"


def _extract_lots_from_mail_text(subject: str, body: str) -> list[str]:
    """Détection heuristique rapide des lots mentionnés dans l'objet et le corps du mail."""
    combined = f"{subject}\n{body}"
    lots_found: list[str] = []
    for match in re.finditer(r"\b(?:lot\s*n?°?\s*(\d{1,2})\b(?:\s*[-:–]?\s*([A-Za-zÀ-ÿ0-9\s/&',]{2,40}))?)", combined, flags=re.IGNORECASE):
        lot_num = match.group(1).zfill(2)
        lot_title = (match.group(2) or "").strip(" .,-:;()\n\r\t")
        label = f"Lot {lot_num}" + (f" - {lot_title}" if lot_title else "")
        if label not in lots_found:
            lots_found.append(label)

    # Mots-clés corps d'état BTP courants si pas de numéro de lot explicite
    corps_etat = [
        "Gros Œuvre", "Gros Oeuvre", "VRD", "Terrassement", "Voirie",
        "Électricité", "Electricité", "CFA", "CFO", "Plomberie", "CVC", "Chauffage",
        "Menuiserie", "Charpente", "Couverture", "Étanchéité", "Peinture",
        "Carrelage", "Sols", "Cloisons", "Plâtrerie", "Faux-plafonds", "Serrurerie",
        "Métallerie", "Façade", "ITE", "Désamiantage", "Démolition", "Espaces Verts",
    ]
    if not lots_found:
        for ce in corps_etat:
            if re.search(rf"\b{re.escape(ce)}\b", combined, flags=re.IGNORECASE):
                lots_found.append(ce)
    return lots_found


def _heuristic_classify_single_file(filename: str, text_preview: str, target_lots: list[str]) -> tuple[str, str]:
    """
    Classifie de façon déterministe un fichier en :
    - 'RC'
    - 'CCTP'
    - 'IGNORED_CCTP' (CCTP d'un autre lot non demandé par l'entreprise)
    - 'DPGF'
    - 'ANNEXES'
    - 'ENTREPRISE'
    Retourne (categorie, justification).
    """
    fname_low = filename.lower()
    preview_low = (text_preview or "")[:3500].lower()

    # 1. DPGF / DQE / BPU
    if any(k in fname_low for k in ("dpgf", "dqe", "bpu", "chiffrage", "quantitatif", "devis_quantitatif")) or (
        fname_low.endswith((".xlsx", ".xls", ".csv"))
        and any(k in preview_low for k in ("prix unitaire", "quantité", "quantite", "total ht", "dpgf", "dqe", "m²", "m2", "m3", "ml", "forfait"))
    ):
        return "DPGF", "Grille quantitative / chiffrage DPGF-DQE détectée"

    # 2. Règlement de Consultation (RC)
    if (
        re.search(r"(^|[_\-\s\.])rc([_\-\s\.]|$)", fname_low)
        or "reglement_de_consultation" in fname_low
        or "reglement de consultation" in fname_low
        or "règlement de consultation" in fname_low
        or ("règlement de la consultation" in preview_low and "critères de jugement" in preview_low)
    ):
        return "RC", "Règlement de Consultation (critères de notation & barème)"

    # 3. Documents Entreprise (ancien mémoire, CV, matériel, moyens, Kbis, Qualibat, RGE, organigramme)
    if any(
        k in fname_low
        for k in (
            "entreprise", "moyens", "materiel", "matériel", "personnel", "organigramme",
            "cv", "qualibat", "rge", "certif", "ancien_memoire", "ancien-memoire",
            "plaquette", "reference", "référence", "caces", "aipr", "fntp",
        )
    ) or (
        any(k in preview_low for k in ("parc matériel", "organigramme", "nos références", "moyens humains de l'entreprise"))
        and "cahier des clauses techniques" not in preview_low
    ):
        return "ENTREPRISE", "Document interne de l'entreprise (moyens humains, matériels, certifications ou ancien mémoire)"

    # 4. Annexes Chantier (PGC, SPS, Plan de masse, Étude de sol G2, SOSED, PIC)
    if any(
        k in fname_low
        for k in (
            "pgc", "sps", "securite", "sécurité", "plan_de_masse", "plan-de-masse",
            "masse", "pic", "etude_de_sol", "etude-de-sol", "geotechnique", "géotechnique",
            "g2", "sosed", "soged", "amiante", "daat", "annexe",
        )
    ) or (
        any(k in preview_low for k in ("plan général de coordination", "coordonnateur sps", "étude géotechnique"))
        and "cahier des clauses techniques" not in preview_low
    ):
        return "ANNEXES", "Annexe chantier (PGC Sécurité, Plan de masse, Étude de sol ou SOSED)"

    # 5. CCTP (avec filtrage intelligent du Lot 00 Général + Lot(s) visé(s) uniquement)
    is_cctp = (
        "cctp" in fname_low
        or "cahier des clauses techniques" in preview_low
        or "clauses techniques particulières" in preview_low
        or re.search(r"lot[\s_\-]*\d+", fname_low) is not None
    )
    if is_cctp:
        # Vérifier si c'est le CCTP Général / Lot 00 / Prescriptions communes -> toujours gardé
        if any(k in fname_low for k in ("lot_00", "lot 00", "lot00", "general", "général", "communes", "commun")):
            return "CCTP", "CCTP Général (Lot 00 / prescriptions communes à tous les lots) — Retenu"

        # Si le mail précise un ou plusieurs lots visés, vérifier si ce CCTP correspond au lot demandé ou à un autre lot
        if target_lots:
            target_joined = " ".join(target_lots).lower()
            target_nums = set(re.findall(r"\blot\s*n?°?\s*0*(\d+)\b", target_joined))
            file_lot_nums = set(re.findall(r"lot[\s_\-]*0*(\d+)", fname_low))
            if not file_lot_nums:
                file_lot_nums = set(re.findall(r"\blot\s*n?°?\s*0*(\d+)\b", preview_low[:1000]))

            # Mots clés métiers des lots ciblés (ex: "vrd", "gros oeuvre", "electricite")
            stop_words = {"lot", "les", "des", "aux", "pour", "dans", "avec", "sur", "par"}
            target_keywords = [
                w for w in re.findall(r"[a-zà-ÿ]{3,}", target_joined)
                if w not in stop_words and not w.isdigit()
            ]

            if file_lot_nums and target_nums:
                if file_lot_nums & target_nums or "0" in file_lot_nums:
                    return "CCTP", f"CCTP correspondant au lot visé dans le mail ({', '.join(target_lots)}) — Retenu"
                else:
                    # Vérifier si un mot-clé métier correspond quand même
                    if target_keywords and any(kw in fname_low for kw in target_keywords):
                        return "CCTP", f"CCTP correspondant au corps d'état visé ({', '.join(target_lots)}) — Retenu"
                    return (
                        "IGNORED_CCTP",
                        f"CCTP d'un autre corps d'état (Lot {', '.join(sorted(file_lot_nums))}) ignoré car le mail cible : {', '.join(target_lots)}",
                    )

            if target_keywords:
                # Si le fichier mentionne explicitement un autre corps d'état connu et aucun des mots-clés cibles
                other_trades = [
                    "gros_oeuvre", "gros-oeuvre", "gros oeuvre", "vrd", "electricite", "électricité",
                    "plomberie", "cvc", "menuiserie", "peinture", "carrelage", "charpente", "etancheite", "serrurerie",
                ]
                matched_target = any(kw in fname_low or kw in preview_low[:800] for kw in target_keywords)
                matched_other = [t for t in other_trades if t in fname_low and not any(kw in t for kw in target_keywords)]
                if matched_other and not matched_target:
                    return (
                        "IGNORED_CCTP",
                        f"CCTP d'un autre lot ({', '.join(matched_other)}) écarté (lot demandé : {', '.join(target_lots)})",
                    )

        return "CCTP", "CCTP technique retenu pour la génération"

    # Par défaut : si c'est un tableur -> DPGF, sinon si "mémoire" dans le texte -> ENTREPRISE, sinon CCTP
    if fname_low.endswith((".xlsx", ".xls", ".csv")):
        return "DPGF", "Tableur quantitatif classé en DPGF"
    if "mémoire technique" in preview_low and "critères" not in preview_low:
        return "ENTREPRISE", "Ancien mémoire technique / document entreprise"
    return "CCTP", "Pièce technique du dossier de consultation"


SYSTEM_PROMPT_TRIAGE_EMAIL = """\
Tu es un ingénieur études de prix BTP chargé de trier automatiquement les pièces jointes reçues par e-mail pour générer un mémoire technique.

À partir de l'objet du mail, du corps du mail et de la liste des fichiers joints (avec un aperçu de leur contenu), renvoie UNIQUEMENT un JSON valide avec la structure exacte suivante :
{
  "nom_entreprise_detecte": "Nom de l'entreprise expéditrice si mentionné dans le mail ou les pièces, sinon ''",
  "lots_vises": ["Ex: Lot 02 - VRD", "Ex: Lot 01 - Gros Œuvre"],
  "consignes_specifiques": "Synthèse des consignes spécifiques données par le client dans le corps du mail (variantes, points à mettre en avant, délais, équipes...)",
  "classification_fichiers": [
    {
      "filename": "nom_exact_du_fichier.ext",
      "categorie": "RC | CCTP | IGNORED_CCTP | DPGF | ANNEXES | ENTREPRISE",
      "motif": "Explication courte du classement (ex: CCTP Lot 00 Général retenu, CCTP Lot 03 Peinture ignoré car hors lot visé)"
    }
  ]
}

RÈGLES STRICTES DE CLASSEMENT :
- `RC` : Règlement de Consultation (barème de notation, critères Valeur Technique, plan imposé, limite de pages).
- `CCTP` : Le CCTP Général (Lot 00) + UNIQUEMENT le(s) CCTP correspondant au(x) lot(s) demandé(s) dans le mail.
- `IGNORED_CCTP` : Les CCTP des autres corps d'état / autres lots NON demandés dans le mail.
- `DPGF` : Le fichier Excel ou PDF contenant le chiffrage et les quantités (m², m³, ml) du lot visé.
- `ANNEXES` : PGC Sécurité, Plan de masse, Étude géotechnique, SOSED.
- `ENTREPRISE` : Ancien mémoire technique, liste matériel/personnel, CV, certifications (Qualibat, RGE, CACES).
"""


def auto_triage_email_and_files(
    sender_email: str,
    subject: str,
    body_text: str,
    attachments: list[MemoryAttachment],
    extract_text_fn: Callable[[MemoryAttachment], str],
    call_llm_fn: Optional[Callable[[str, str, str, str], str]],
    api_key: str,
    provider: str,
    default_config_entreprise: dict,
) -> dict:
    """
    Exécute le tri intelligent BTP :
    1. Identifie l'entreprise (profil enregistré par email/domaine ou via document joint).
    2. Lit les consignes du mail et identifie le(s) lot(s) visé(s).
    3. Extrait le texte de chaque pièce jointe et la classe dans RC, CCTP, IGNORED_CCTP, DPGF, ANNEXES, ENTREPRISE.
    """
    entreprise_profil, entreprise_source = find_company_profile_by_email(sender_email, default_config_entreprise)
    heuristic_lots = _extract_lots_from_mail_text(subject, body_text)

    # Extraction du texte de chaque pièce jointe
    extracted_texts: dict[str, str] = {}
    previews_for_ai: list[dict] = []
    extraction_errors: list[str] = []

    for att in attachments:
        try:
            txt = extract_text_fn(att)
            extracted_texts[att.name] = txt
            previews_for_ai.append({
                "filename": att.name,
                "ext": Path(att.name).suffix.lower(),
                "preview": txt[:1200],
            })
        except Exception as exc:
            extracted_texts[att.name] = ""
            extraction_errors.append(f"{att.name} : {exc}")
            previews_for_ai.append({
                "filename": att.name,
                "ext": Path(att.name).suffix.lower(),
                "preview": "",
            })

    # Classification heuristique initiale (rapide et déterministe)
    file_classifications: dict[str, dict] = {}
    for att in attachments:
        cat, motif = _heuristic_classify_single_file(att.name, extracted_texts.get(att.name, ""), heuristic_lots)
        file_classifications[att.name] = {
            "filename": att.name,
            "categorie": cat,
            "motif": motif,
        }

    lots_vises = list(heuristic_lots)
    consignes_specifiques = body_text.strip()
    nom_entreprise_detecte = (entreprise_profil or {}).get("nom", "")

    # Si une clé API est disponible et qu'il y a des pièces jointes, affiner avec l'IA
    if api_key and call_llm_fn and attachments:
        try:
            user_prompt = (
                f"EXPÉDITEUR : {sender_email}\n"
                f"OBJET DU MAIL : {subject}\n"
                f"CORPS DU MAIL :\n{body_text}\n\n"
                f"PIÈCES JOINTES DÉTECTÉES :\n{json.dumps(previews_for_ai, ensure_ascii=False, indent=2)}"
            )
            raw_ai = call_llm_fn(api_key, provider, SYSTEM_PROMPT_TRIAGE_EMAIL, user_prompt)
            m = re.search(r"```json\s*(.*?)\s*```", raw_ai, re.DOTALL)
            parsed_ai = json.loads(m.group(1) if m else raw_ai)
            if isinstance(parsed_ai.get("lots_vises"), list) and parsed_ai["lots_vises"]:
                lots_vises = [str(x) for x in parsed_ai["lots_vises"] if str(x).strip()]
            if parsed_ai.get("consignes_specifiques"):
                consignes_specifiques = str(parsed_ai["consignes_specifiques"]).strip()
            if parsed_ai.get("nom_entreprise_detecte") and not nom_entreprise_detecte:
                nom_entreprise_detecte = str(parsed_ai["nom_entreprise_detecte"]).strip()

            valid_cats = {"RC", "CCTP", "IGNORED_CCTP", "DPGF", "ANNEXES", "ENTREPRISE"}
            for item in parsed_ai.get("classification_fichiers", []):
                fname = item.get("filename")
                cat = str(item.get("categorie", "")).upper().strip()
                if fname in file_classifications and cat in valid_cats:
                    file_classifications[fname] = {
                        "filename": fname,
                        "categorie": cat,
                        "motif": item.get("motif") or file_classifications[fname]["motif"],
                    }
        except Exception:
            # En cas d'indisponibilité ou erreur de formatage, la classification heuristique prend le relais sans bloquer
            pass

    # Répartir les fichiers dans les 5 catégories du générateur (+ IGNORED_CCTP)
    categorized_files: dict[str, list[MemoryAttachment]] = {
        "RC": [],
        "CCTP": [],
        "IGNORED_CCTP": [],
        "DPGF": [],
        "ANNEXES": [],
        "ENTREPRISE": [],
    }
    for att in attachments:
        cat = file_classifications[att.name]["categorie"]
        categorized_files.setdefault(cat, []).append(att)

    # Si l'entreprise n'était pas connue dans la base mais qu'un nom a été détecté dans le mail/les pièces
    if entreprise_profil is None and nom_entreprise_detecte:
        entreprise_profil = {
            "nom": nom_entreprise_detecte,
            "email": sender_email,
            "forme_juridique": "",
            "siret": "",
            "adresse": "",
            "telephone": "",
            "dirigeant": "",
            "effectif": "",
            "chiffre_affaires": "",
            "activites_principales": ", ".join(lots_vises),
            "certifications": "",
            "qualifications": "",
            "assurances": "",
            "references_chantiers": "",
            "moyens_materiels": "",
            "zone_intervention": "",
            "site_web": "",
            "date_creation": "",
            "presentation_libre": "",
        }
        entreprise_source = f"Entreprise détectée depuis l'e-mail / pièces jointes ({nom_entreprise_detecte})"

    return {
        "entreprise_profil": entreprise_profil or {},
        "entreprise_source": entreprise_source,
        "lots_vises": lots_vises,
        "consignes_specifiques": consignes_specifiques,
        "classifications": list(file_classifications.values()),
        "categorized_files": categorized_files,
        "extracted_texts": extracted_texts,
        "extraction_errors": extraction_errors,
    }


# ─────────────────────────────────────────────────────────────────────────
# 4. SYNTHÈSE DES 5 POINTS DE VIGILANCE & COMPOSITION DU MAIL DE RÉPONSE
# ─────────────────────────────────────────────────────────────────────────

def extract_five_vigilance_points(analyse_json: dict) -> list[dict[str, str]]:
    """
    Extrait les 5 points de vigilance majeurs du chantier trouvés dans le RC et le CCTP :
    1. Limite de pages & Critères du Règlement de Consultation (RC)
    2. Délais d'exécution & Dates clés
    3. Contraintes d'accès, horaires & environnement du chantier
    4. Sécurité (PGC), coactivité & installation de chantier
    5. Matériaux prescrits & normes DTU majeures
    """
    rc_info = analyse_json.get("reglement_consultation_rc", {}) if isinstance(analyse_json.get("reglement_consultation_rc"), dict) else {}
    nb_pages_max = rc_info.get("nombre_pages_max") or "Non spécifié au RC"
    criteres_vt = rc_info.get("criteres_valeur_technique") or []
    if criteres_vt and isinstance(criteres_vt, list):
        crit_str = ", ".join(
            f"{c.get('critere', '')} ({c.get('bareme_points', '')})" if isinstance(c, dict) else str(c)
            for c in criteres_vt[:4]
        )
    else:
        crit_str = "Selon trame standard 8 sections BTP"

    dates_cles = analyse_json.get("dates_cles", {}) if isinstance(analyse_json.get("dates_cles"), dict) else {}
    delai_exec = dates_cles.get("delai_execution") or "Non spécifié"
    remise_offres = dates_cles.get("date_remise_offres") or "Non spécifié"

    contraintes = analyse_json.get("contraintes_chantier") or []
    contraintes_str = " ; ".join(str(c) for c in contraintes[:3]) if contraintes else "Voir prescriptions générales du CCTP"

    annexes_sec = analyse_json.get("annexes_securite_installation", {}) if isinstance(analyse_json.get("annexes_securite_installation"), dict) else {}
    risques = annexes_sec.get("risques_specifiques_pgc") or analyse_json.get("contraintes_securite_environnement") or []
    sec_str = " ; ".join(str(r) for r in risques[:3]) if risques else "Respect strict du PGC, PPSPS et port des EPI obligatoire"

    lots_tech = analyse_json.get("exigences_techniques_par_lot") or []
    mat_norms: list[str] = []
    for lt in lots_tech[:3]:
        if isinstance(lt, dict):
            mats = lt.get("materiaux_requis") or []
            dtus = lt.get("normes_dtu") or []
            parts = [str(x) for x in (mats[:2] + dtus[:2]) if x]
            if parts:
                mat_norms.append(f"{lt.get('lot', 'Lot')} : {', '.join(parts)}")
    if not mat_norms:
        normes_gen = analyse_json.get("normes_certifications") or []
        mat_norms = [str(n) for n in normes_gen[:4]]
    mat_str = " | ".join(mat_norms) if mat_norms else "Conformité aux DTU en vigueur et marquage CE/NF requis"

    return [
        {
            "titre": "1. Limite de pages & Critères RC (Valeur Technique)",
            "detail": f"Maximum de pages : {nb_pages_max} — Critères : {crit_str}",
        },
        {
            "titre": "2. Délais d'exécution & Remise des offres",
            "detail": f"Délai travaux : {delai_exec} — Date limite de remise des plis : {remise_offres}",
        },
        {
            "titre": "3. Contraintes d'accès, horaires & site",
            "detail": contraintes_str,
        },
        {
            "titre": "4. Sécurité chantier, PGC & Coactivité",
            "detail": sec_str,
        },
        {
            "titre": "5. Matériaux prescrits & Normes DTU",
            "detail": mat_str,
        },
    ]


def build_success_reply_email(
    nom_projet: str,
    lots_vises: list[str],
    classifications: list[dict],
    vigilance_points: list[dict[str, str]],
    balises_a_completer: list[str],
) -> tuple[str, str, str]:
    """
    Construit (subject, plain_text_body, html_body) pour l'e-mail de réponse automatique
    contenant le mémoire technique (.docx et .pdf) en pièce jointe.
    """
    lots_str = ", ".join(lots_vises) if lots_vises else "Tous lots soumis"
    subject = f"✅ Mémoire Technique prêt : {nom_projet} ({lots_str})"

    retained_docs = [c for c in classifications if c["categorie"] != "IGNORED_CCTP"]
    ignored_docs = [c for c in classifications if c["categorie"] == "IGNORED_CCTP"]

    docs_lines_txt = "\n".join(
        f"  - [{c['categorie']}] {c['filename']} ({c['motif']})" for c in retained_docs
    ) or "  - Aucun fichier listé"
    if ignored_docs:
        docs_lines_txt += "\n  Pièces écartées automatiquement (autres lots non visés) :\n" + "\n".join(
            f"  - [IGNORÉ] {c['filename']} ({c['motif']})" for c in ignored_docs
        )

    vigilance_txt = "\n".join(
        f"  {vp['titre']} :\n    -> {vp['detail']}" for vp in vigilance_points
    )

    if balises_a_completer:
        balises_txt = "\n".join(f"  - {b}" for b in balises_a_completer)
    else:
        balises_txt = "  - Aucune balise [À COMPLÉTER] restante : toutes les données étaient présentes dans vos pièces."

    plain_body = (
        f"Bonjour,\n\n"
        f"Votre Mémoire Technique pour le marché « {nom_projet} » ({lots_str}) a été généré automatiquement "
        f"par MarketBuilding Pro et synchronisé sur votre espace web.\n\n"
        f"Vous trouverez en pièces jointes de cet e-mail :\n"
        f"  1. Le Mémoire Technique complet au format Word éditable (.docx)\n"
        f"  2. Le Mémoire Technique officiel mis en page au format PDF (.pdf)\n\n"
        f"====================================================================\n"
        f"1. DOCUMENTS DU DCE DÉTECTÉS ET CROISÉS\n"
        f"====================================================================\n"
        f"{docs_lines_txt}\n\n"
        f"====================================================================\n"
        f"2. SYNTHÈSE DES 5 POINTS DE VIGILANCE MAJEURS DU CHANTIER (RC & CCTP)\n"
        f"====================================================================\n"
        f"{vigilance_txt}\n\n"
        f"====================================================================\n"
        f"3. BALISES [À COMPLÉTER] RESTANTES DANS LE DOCUMENT ({len(balises_a_completer)})\n"
        f"====================================================================\n"
        f"{balises_txt}\n\n"
        f"Vous pouvez également consulter, modifier en direct et retélécharger ce mémoire sur la plateforme MarketBuilding Pro.\n\n"
        f"Cordialement,\n"
        f"Le Service d'Automatisation MarketBuilding Pro"
    )

    docs_html = "".join(
        f"<li><b>[{c['categorie']}]</b> <code>{c['filename']}</code> — <span style='color:#475569;'>{c['motif']}</span></li>"
        for c in retained_docs
    )
    ignored_html = ""
    if ignored_docs:
        ignored_items = "".join(
            f"<li><code>{c['filename']}</code> — <i>{c['motif']}</i></li>" for c in ignored_docs
        )
        ignored_html = (
            f"<div style='margin-top:8px; font-size:13px; color:#64748B;'>"
            f"<b>CCTP d'autres lots écartés automatiquement :</b><ul>{ignored_items}</ul></div>"
        )

    vigilance_html = "".join(
        f"<li style='margin-bottom:8px;'><b>{vp['titre']} :</b><br><span style='color:#334155;'>{vp['detail']}</span></li>"
        for vp in vigilance_points
    )

    if balises_a_completer:
        balises_html = (
            "<ul style='color:#B45309; font-weight:600;'>"
            + "".join(f"<li><code>{b}</code></li>" for b in balises_a_completer)
            + "</ul>"
        )
    else:
        balises_html = "<p style='color:#059669; font-weight:600;'>✅ Aucune balise <code>[À COMPLÉTER]</code> restante.</p>"

    html_body = f"""
    <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #0F172A; max-width: 720px; line-height: 1.55;">
        <div style="background: linear-gradient(135deg, #1E3A8A, #2563EB); color: #FFFFFF; padding: 20px 24px; border-radius: 12px 12px 0 0;">
            <div style="font-size: 12px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9;">MarketBuilding Pro · Automatisation BTP</div>
            <h2 style="margin: 6px 0 0 0; font-size: 20px;">Mémoire Technique Généré : {nom_projet}</h2>
            <div style="font-size: 13px; margin-top: 4px;">Lot(s) traité(s) : <b>{lots_str}</b></div>
        </div>
        <div style="border: 1px solid #E2E8F0; border-top: none; padding: 22px 24px; border-radius: 0 0 12px 12px; background: #FFFFFF;">
            <p>Bonjour,</p>
            <p>Votre <b>Mémoire Technique</b> a été rédigé automatiquement à partir de vos pièces et synchronisé sur votre plateforme <b>MarketBuilding Pro</b> (avec le badge <i>« Automatisé via Mail »</i>).</p>
            <p>📎 <b>Pièces jointes à cet e-mail :</b> le Mémoire Technique complet aux formats <b>Word (<code>.docx</code>)</b> et <b>PDF officiel (<code>.pdf</code>)</b>.</p>

            <h3 style="color: #1E40AF; border-bottom: 2px solid #DBEAFE; padding-bottom: 6px; margin-top: 22px;">1. Documents du DCE détectés et croisés</h3>
            <ul>{docs_html}</ul>
            {ignored_html}

            <h3 style="color: #1E40AF; border-bottom: 2px solid #DBEAFE; padding-bottom: 6px; margin-top: 22px;">2. Synthèse des 5 points de vigilance majeurs du chantier</h3>
            <ul style="list-style-type: none; padding-left: 0;">{vigilance_html}</ul>

            <h3 style="color: #D97706; border-bottom: 2px solid #FEF3C7; padding-bottom: 6px; margin-top: 22px;">3. Balises [À COMPLÉTER] restantes dans le document ({len(balises_a_completer)})</h3>
            {balises_html}

            <p style="margin-top: 22px; font-size: 13px; color: #64748B;">
                Vous pouvez ouvrir, lire, modifier et télécharger ce mémoire directement depuis votre interface web MarketBuilding Pro.
            </p>
        </div>
    </div>
    """
    return subject, plain_body, html_body


def build_missing_docs_reply_email(
    original_subject: str,
    classifications: list[dict],
) -> tuple[str, str, str]:
    """
    Construit l'e-mail de réponse automatique lorsque le mail reçu ne contient aucun
    document technique exploitable (ex: pas de CCTP).
    """
    subject = f"⚠️ Pièce manquante pour générer votre Mémoire Technique ({original_subject or 'Votre demande'})"
    present_cats = {c["categorie"] for c in classifications}

    missing_list: list[str] = []
    if "CCTP" not in present_cats:
        missing_list.append("Le Cahier des Clauses Techniques Particulières (CCTP) du lot visé (indispensable pour la méthodologie technique et les matériaux)")
    if "RC" not in present_cats:
        missing_list.append("Le Règlement de Consultation (RC) (recommandé pour respecter le barème de notation et la limite de pages)")
    if "DPGF" not in present_cats:
        missing_list.append("La DPGF ou le DQE (recommandé pour dimensionner les quantités m², m³, ml et le planning)")

    missing_txt = "\n".join(f"  - {m}" for m in missing_list)
    plain_body = (
        "Bonjour,\n\n"
        "Nous avons bien reçu votre e-mail sur le service d'automatisation MarketBuilding Pro.\n"
        "Cependant, nous n'avons détecté aucun document technique exploitable (notamment le CCTP) permettant de rédiger votre mémoire technique.\n\n"
        "Pièce(s) manquante(s) à nous renvoyer en pièce jointe (en PDF, Word, Excel ou archive .zip) :\n"
        f"{missing_txt}\n\n"
        "Dès réception de votre nouveau mail avec le CCTP (ou le dossier DCE .zip complet), votre mémoire technique sera généré et vous sera renvoyé automatiquement en PDF et Word (.docx).\n\n"
        "Cordialement,\n"
        "Le Service d'Automatisation MarketBuilding Pro"
    )

    missing_html = "".join(f"<li style='margin-bottom:6px;'><b>{m}</b></li>" for m in missing_list)
    html_body = f"""
    <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #0F172A; max-width: 680px; line-height: 1.55;">
        <div style="background: #DC2626; color: #FFFFFF; padding: 18px 22px; border-radius: 12px 12px 0 0;">
            <h2 style="margin: 0; font-size: 18px;">⚠️ Pièce(s) manquante(s) pour générer votre Mémoire Technique</h2>
        </div>
        <div style="border: 1px solid #FECACA; border-top: none; padding: 20px 22px; border-radius: 0 0 12px 12px; background: #FFFFFF;">
            <p>Bonjour,</p>
            <p>Nous avons bien reçu votre demande (<i>{original_subject}</i>), mais <b>aucun CCTP ou document technique exploitable</b> n'a été trouvé en pièce jointe.</p>
            <p>Pour lancer la rédaction automatique de votre mémoire technique, merci de nous renvoyer un e-mail contenant :</p>
            <ul>{missing_html}</ul>
            <p>Vous pouvez joindre vos fichiers en vrac ou directement sous forme d'archive <code>.zip</code>.</p>
        </div>
    </div>
    """
    return subject, plain_body, html_body


def send_smtp_email(
    state: dict,
    to_email: str,
    subject: str,
    plain_body: str,
    html_body: str,
    attachments: Optional[list[tuple[str, bytes, str]]] = None,
) -> tuple[bool, str]:
    """
    Envoie l'e-mail via SMTP avec les pièces jointes (filename, bytes, mime_subtype).
    Si SMTP n'est pas configuré (ex: mode test local sans serveur SMTP), enregistre l'e-mail
    dans le dossier `archives/sent_emails/` et retourne un succès simulé explicite.
    """
    smtp_host = (state.get("smtp_host") or "smtp.gmail.com").strip()
    smtp_port = int(state.get("smtp_port") or 587)
    smtp_user = (state.get("smtp_user") or DEFAULT_REFERENCE_EMAIL).strip()
    smtp_password = (state.get("smtp_password") or "").strip()
    smtp_from = (state.get("smtp_from") or smtp_user or DEFAULT_REFERENCE_EMAIL).strip()

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"MarketBuilding Pro BTP <{smtp_from}>" if "@" in smtp_from and "<" not in smtp_from else smtp_from
    msg["Reply-To"] = DEFAULT_REFERENCE_EMAIL
    msg["To"] = to_email
    msg.set_content(plain_body)
    msg.add_alternative(html_body, subtype="html")

    for fname, fbytes, subtype in (attachments or []):
        if subtype == "pdf":
            msg.add_attachment(fbytes, maintype="application", subtype="pdf", filename=fname)
        else:
            msg.add_attachment(
                fbytes,
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename=fname,
            )

    # Sauvegarder une copie locale de l'e-mail envoyé pour traçabilité dans l'interface
    os.makedirs("archives/sent_emails", exist_ok=True)
    eml_path = os.path.join(
        "archives/sent_emails",
        f"Reponse_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{re.sub(r'[^a-zA-Z0-9_-]', '_', to_email)}.eml",
    )
    try:
        with open(eml_path, "wb") as f:
            f.write(bytes(msg))
    except Exception:
        pass

    if not smtp_host or not smtp_user or not smtp_password:
        return (
            True,
            f"Réponse e-mail générée et archivée localement ({eml_path}) [Identifiants SMTP non renseignés — prêt pour envoi réel dès configuration SMTP]",
        )

    try:
        if smtp_port == 465:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=25) as server:
                server.login(smtp_user, smtp_password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(smtp_host, smtp_port, timeout=25) as server:
                server.ehlo()
                server.starttls(context=ssl.create_default_context())
                server.ehlo()
                server.login(smtp_user, smtp_password)
                server.send_message(msg)
        return True, f"E-mail envoyé avec succès par SMTP à {to_email}"
    except Exception as exc:
        return False, f"Erreur d'envoi SMTP vers {to_email} : {exc} (copie sauvegardée dans {eml_path})"


# ─────────────────────────────────────────────────────────────────────────
# 5. PIPELINE COMPLET DE TRAITEMENT D'UN E-MAIL (TRI -> GÉNÉRATION -> SITE -> SMTP)
# ─────────────────────────────────────────────────────────────────────────

def process_incoming_email_pipeline(
    sender_email: str,
    sender_name: str,
    subject: str,
    body_text: str,
    raw_attachments: list[tuple[str, bytes]],
    *,
    config: dict,
    extract_text_fn: Callable[[MemoryAttachment], str],
    extract_category_fn: Callable,
    analyse_cctp_fn: Callable,
    rediger_memoire_fn: Callable,
    generate_docx_fn: Callable,
    generate_pdf_fn: Callable,
    build_profil_text_fn: Callable[[dict], str],
    call_llm_fn: Optional[Callable] = None,
    is_simulation: bool = False,
) -> dict:
    """
    Exécute de bout en bout le pipeline sur un e-mail entrant :
    1. Enregistre le statut « Mail reçu » et décompresse les .zip éventuels.
    2. Lance l'Auto-Triage BTP (« Tri des CCTP en cours »).
    3. Vérifie la présence d'au moins un document technique exploitable (sinon répond « Pièce manquante »).
    4. Lance la génération du mémoire (« Génération sur le site ») et synchronise le projet dans `projects_db.json` + `archives/` avec le badge « Automatisé via Mail ».
    5. Selon le mode (« 100 % Automatique » ou « Validation Manuelle »), envoie immédiatement l'e-mail de réponse avec le PDF + DOCX + 5 points de vigilance + balises `[À COMPLÉTER]`, ou place le job en « En attente de validation manuelle ».
    """
    job_id = f"mail_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    state = load_automation_state()
    mode = state.get("automation_mode", "auto")

    job_record = {
        "job_id": job_id,
        "received_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "sender_email": sender_email,
        "sender_name": sender_name,
        "subject": subject,
        "body_text": body_text,
        "is_simulation": is_simulation,
        "mode_at_creation": mode,
        "statut": "Mail reçu",
        "extraction_notes": [],
        "lots_vises": [],
        "consignes_specifiques": "",
        "entreprise_source": "",
        "classifications": [],
        "vigilance_points": [],
        "balises_a_completer": [],
        "project_id": None,
        "nom_projet": "",
        "pdf_path": "",
        "docx_path": "",
        "email_subject_out": "",
        "email_body_plain_out": "",
        "email_body_html_out": "",
        "smtp_info": "",
        "error": "",
    }

    with _DB_LOCK:
        st_data = load_automation_state()
        st_data["jobs"].insert(0, job_record)
        st_data["jobs"] = st_data["jobs"][:80]
        save_automation_state(st_data)

    add_automation_log(
        "Mail reçu",
        f"Nouveau mail de {sender_email} ({len(raw_attachments)} pièce(s) brute(s))",
        email_subject=subject,
        sender=sender_email,
        job_id=job_id,
    )

    # Étape 1 : Décompression des .zip et récupération des fichiers
    unpacked_files, unpack_notes = unpack_attachments_and_zips(raw_attachments)
    update_job_status(job_id, statut="Tri des CCTP en cours", extraction_notes=unpack_notes)
    add_automation_log(
        "Tri des CCTP en cours",
        f"{len(unpacked_files)} fichier(s) prêts après décompression ZIP. Analyse des lots et tri RC / CCTP / DPGF / Annexes / Entreprise…",
        email_subject=subject,
        sender=sender_email,
        job_id=job_id,
    )

    api_key = (config.get("api_key") or "").strip()
    provider = config.get("llm_provider", "OpenAI (GPT-4o)")

    # Étape 2 : Auto-triage BTP
    triage = auto_triage_email_and_files(
        sender_email=sender_email,
        subject=subject,
        body_text=body_text,
        attachments=unpacked_files,
        extract_text_fn=extract_text_fn,
        call_llm_fn=None if is_simulation else call_llm_fn,
        api_key="" if is_simulation else api_key,
        provider=provider,
        default_config_entreprise=config.get("entreprise", {}),
    )

    classifications = triage["classifications"]
    cat_files = triage["categorized_files"]
    lots_vises = triage["lots_vises"]
    consignes_specifiques = triage["consignes_specifiques"]
    entreprise_profil = triage["entreprise_profil"]
    entreprise_source = triage["entreprise_source"]

    update_job_status(
        job_id,
        lots_vises=lots_vises,
        consignes_specifiques=consignes_specifiques,
        entreprise_source=entreprise_source,
        classifications=classifications,
    )

    # Vérifier s'il y a au moins un CCTP (ou à défaut un RC/DPGF avec texte exploitable)
    exploitable_tech_files = cat_files.get("CCTP", []) + cat_files.get("RC", []) + cat_files.get("DPGF", [])
    has_cctp = len(cat_files.get("CCTP", [])) > 0

    if not exploitable_tech_files or not has_cctp:
        miss_subj, miss_plain, miss_html = build_missing_docs_reply_email(subject, classifications)
        ok_smtp, info_smtp = send_smtp_email(
            load_automation_state(),
            to_email=sender_email,
            subject=miss_subj,
            plain_body=miss_plain,
            html_body=miss_html,
            attachments=[],
        )
        update_job_status(
            job_id,
            statut="Pièce manquante signalée",
            email_subject_out=miss_subj,
            email_body_plain_out=miss_plain,
            email_body_html_out=miss_html,
            smtp_info=info_smtp,
            error="Aucun CCTP exploitable joint au mail (réponse automatique envoyée à l'entreprise).",
        )
        add_automation_log(
            "Pièce manquante signalée",
            f"Aucun CCTP exploitable trouvé dans le mail de {sender_email}. E-mail de relance envoyé automatiquement.",
            email_subject=subject,
            sender=sender_email,
            job_id=job_id,
        )
        return update_job_status(job_id) or job_record

    # Étape 3 : Génération du mémoire technique sur le site
    update_job_status(job_id, statut="Génération sur le site")
    add_automation_log(
        "Génération sur le site",
        f"Croisement RC ({len(cat_files['RC'])}) + CCTP retenus ({len(cat_files['CCTP'])}, {len(cat_files['IGNORED_CCTP'])} écarté(s)) + DPGF ({len(cat_files['DPGF'])}) + Annexes ({len(cat_files['ANNEXES'])}) + Entreprise ({len(cat_files['ENTREPRISE'])})…",
        email_subject=subject,
        sender=sender_email,
        job_id=job_id,
    )

    categories_input = [
        ("1. RÈGLEMENT DE CONSULTATION (RC)", cat_files.get("RC", [])),
        ("2. CAHIERS DES CLAUSES TECHNIQUES PARTICULIÈRES (CCTP GÉNÉRAL + LOTS VISÉS)", cat_files.get("CCTP", [])),
        ("3. DPGF / DQE (QUANTITÉS DU MARCHÉ)", cat_files.get("DPGF", [])),
        ("4. ANNEXES CHANTIER (PGC SÉCURITÉ, PLAN DE MASSE, ÉTUDE DE SOL)", cat_files.get("ANNEXES", [])),
        ("5. DOCUMENTS DE L'ENTREPRISE (ANCIEN MÉMOIRE, PERSONNEL, MATÉRIEL, CERTIFICATIONS)", cat_files.get("ENTREPRISE", [])),
    ]

    combined_sections: list[str] = []
    extraits_par_categorie: dict[str, str] = {}
    source_filenames: list[str] = []

    for cat_title, flist in categories_input:
        if not flist:
            continue
        res_cat = extract_category_fn(flist, cat_title)
        cat_text = res_cat[0] if isinstance(res_cat, tuple) else res_cat
        for f in flist:
            source_filenames.append(f.name)
        if cat_text:
            combined_sections.append(cat_text)
            extraits_par_categorie[cat_title] = cat_text

    # Ajouter les consignes spécifiques de l'e-mail et les lots visés
    consignes_header = (
        f"================ CONSIGNES SPÉCIFIQUES DE L'ENTREPRISE (E-MAIL) ================\n"
        f"Objet du mail : {subject}\n"
        f"Lot(s) visé(s) exclusivement : {', '.join(lots_vises) if lots_vises else 'Selon pièces CCTP'}\n"
        f"Consignes du mail :\n{consignes_specifiques}\n"
    )
    combined_sections.insert(0, consignes_header)
    extraits_par_categorie["0. CONSIGNES E-MAIL & LOTS VISÉS"] = consignes_header

    full_dossier_text = "\n\n".join(combined_sections)
    profil_text = build_profil_text_fn(entreprise_profil) if entreprise_profil else ""

    try:
        if api_key and not is_simulation:
            try:
                analyse = analyse_cctp_fn(api_key, provider, full_dossier_text)
                memoire_md = rediger_memoire_fn(
                    api_key,
                    provider,
                    analyse,
                    profil_text,
                    extraits_par_categorie=extraits_par_categorie,
                )
            except Exception:
                analyse, memoire_md = _build_fallback_structured_memoire(
                    subject=subject,
                    lots_vises=lots_vises,
                    consignes_specifiques=consignes_specifiques,
                    entreprise_profil=entreprise_profil,
                    classifications=classifications,
                    extracted_texts=triage["extracted_texts"],
                )
        else:
            # Génération ultra-rapide et déterministe fidèle aux pièces extraites (RC, CCTP, DPGF, Annexes, Entreprise)
            analyse, memoire_md = _build_fallback_structured_memoire(
                subject=subject,
                lots_vises=lots_vises,
                consignes_specifiques=consignes_specifiques,
                entreprise_profil=entreprise_profil,
                classifications=classifications,
                extracted_texts=triage["extracted_texts"],
            )

        docx_buf = generate_docx_fn(memoire_md, analyse, entreprise_profil or config.get("entreprise", {}))
        pdf_buf = generate_pdf_fn(memoire_md, analyse, entreprise_profil or config.get("entreprise", {}))
        pdf_bytes = pdf_buf.getvalue()
        docx_bytes = docx_buf.getvalue()
    except Exception as exc:
        update_job_status(job_id, statut="Erreur de génération", error=str(exc))
        add_automation_log(
            "Erreur de génération",
            f"Échec lors de la génération pour {subject} : {exc}",
            email_subject=subject,
            sender=sender_email,
            job_id=job_id,
        )
        return update_job_status(job_id) or job_record

    nom_projet = str(analyse.get("nom_projet", "")).strip()
    if not nom_projet or nom_projet in ("Non spécifié", "Non spécifié dans cette section"):
        nom_projet = subject or "Marché BTP"
    nom_fichier = re.sub(r"[^\w\s-]", "", nom_projet).strip().replace(" ", "_")[:55] or "Memoire_Auto"

    os.makedirs("archives", exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path = f"archives/Memoire_AutoMail_{nom_fichier}_{date_str}.pdf"
    docx_path = f"archives/Memoire_AutoMail_{nom_fichier}_{date_str}.docx"

    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)
    with open(docx_path, "wb") as f:
        f.write(docx_bytes)

    # Assembler tous les documents reçus du client + les mémoires générés (.pdf & .docx) dans un dossier client dédié + pack .ZIP
    client_slug = re.sub(r"[^a-zA-Z0-9_-]", "_", sender_email or "client_btp")[:40]
    client_dir = os.path.join("archives", "dossiers_clients", f"{client_slug}_{date_str}")
    os.makedirs(client_dir, exist_ok=True)
    bundle_zip_path = os.path.join(client_dir, f"Pack_Dossier_Assemble_et_Memoire_{nom_fichier}.zip")

    try:
        with open(os.path.join(client_dir, f"Memoire_Technique_{nom_fichier}.pdf"), "wb") as f:
            f.write(pdf_bytes)
        with open(os.path.join(client_dir, f"Memoire_Technique_{nom_fichier}.docx"), "wb") as f:
            f.write(docx_bytes)

        with zipfile.ZipFile(bundle_zip_path, "w", zipfile.ZIP_DEFLATED) as zbundle:
            # 1. Mémoires techniques générés (.pdf et .docx)
            zbundle.writestr(f"01_MEMOIRES_TECHNIQUES_GENERES/Memoire_Technique_{nom_fichier}.pdf", pdf_bytes)
            zbundle.writestr(f"01_MEMOIRES_TECHNIQUES_GENERES/Memoire_Technique_{nom_fichier}.docx", docx_bytes)
            # 2. Documents du client assemblés et triés par catégorie BTP
            for cat_key, folder_lbl in [
                ("RC", "02_PIECES_CLIENT_ASSEMBLEES/01_RC"),
                ("CCTP", "02_PIECES_CLIENT_ASSEMBLEES/02_CCTP_RETENUS"),
                ("DPGF", "02_PIECES_CLIENT_ASSEMBLEES/03_DPGF_DQE"),
                ("ANNEXES", "02_PIECES_CLIENT_ASSEMBLEES/04_ANNEXES_CHANTIER"),
                ("ENTREPRISE", "02_PIECES_CLIENT_ASSEMBLEES/05_DOCS_ENTREPRISE"),
            ]:
                for att_obj in cat_files.get(cat_key, []):
                    zbundle.writestr(f"{folder_lbl}/{att_obj.name}", att_obj.getvalue())
            # 3. Texte intégral fusionné des documents croisés
            zbundle.writestr("02_PIECES_CLIENT_ASSEMBLEES/00_DOSSIER_COMPLET_FUSIONNE.txt", full_dossier_text.encode("utf-8"))
    except Exception:
        bundle_zip_path = ""

    vigilance_points = extract_five_vigilance_points(analyse)
    balises_a_completer = list(dict.fromkeys(re.findall(r"\[À COMPLÉTER[^\]]*\]", memoire_md)))

    # Enregistrer le projet dans la base de données du site avec le badge "Automatisé via Mail"
    project_entry = save_project_to_db({
        "project_id": f"proj_{job_id}",
        "job_id": job_id,
        "created_at": datetime.now().strftime("%d/%m/%Y à %H:%M"),
        "origin": "email",
        "badge": "Automatisé via Mail",
        "reference_email": DEFAULT_REFERENCE_EMAIL,
        "sender_email": sender_email,
        "email_subject": subject,
        "lots_vises": lots_vises,
        "nom_projet": nom_projet,
        "nom_fichier": nom_fichier,
        "source_name": ", ".join(source_filenames),
        "analyse": analyse,
        "memoire_md": memoire_md,
        "entreprise": entreprise_profil or config.get("entreprise", {}),
        "pdf_path": pdf_path,
        "docx_path": docx_path,
        "bundle_zip_path": bundle_zip_path,
        "client_dir": client_dir,
        "vigilance_points": vigilance_points,
        "balises_a_completer": balises_a_completer,
        "classifications": classifications,
    })

    reply_subj, reply_plain, reply_html = build_success_reply_email(
        nom_projet=nom_projet,
        lots_vises=lots_vises,
        classifications=classifications,
        vigilance_points=vigilance_points,
        balises_a_completer=balises_a_completer,
    )

    update_job_status(
        job_id,
        project_id=project_entry["project_id"],
        nom_projet=nom_projet,
        pdf_path=pdf_path,
        docx_path=docx_path,
        bundle_zip_path=bundle_zip_path,
        client_dir=client_dir,
        vigilance_points=vigilance_points,
        balises_a_completer=balises_a_completer,
        email_subject_out=reply_subj,
        email_body_plain_out=reply_plain,
        email_body_html_out=reply_html,
    )

    # Étape 4 : Mode 100% Automatique vs Mode Validation Manuelle
    current_mode = load_automation_state().get("automation_mode", "auto")
    if current_mode == "auto":
        ok_smtp, info_smtp = send_smtp_email(
            load_automation_state(),
            to_email=sender_email,
            subject=reply_subj,
            plain_body=reply_plain,
            html_body=reply_html,
            attachments=[
                (f"Memoire_Technique_{nom_fichier}.pdf", pdf_bytes, "pdf"),
                (f"Memoire_Technique_{nom_fichier}.docx", docx_bytes, "docx"),
            ],
        )
        final_status = "Mémoire envoyé" if ok_smtp else "Erreur d'envoi mail"
        update_job_status(job_id, statut=final_status, smtp_info=info_smtp)
        add_automation_log(
            final_status,
            f"Mémoire « {nom_projet} » publié sur le site et envoyé à {sender_email} (.docx + .pdf + 5 points de vigilance).",
            email_subject=subject,
            sender=sender_email,
            job_id=job_id,
        )
    else:
        update_job_status(
            job_id,
            statut="En attente de validation manuelle",
            smtp_info="Mémoire généré sur le site. En attente de votre clic sur « Valider et Envoyer ».",
        )
        add_automation_log(
            "En attente de validation manuelle",
            f"Mémoire « {nom_projet} » prêt sur le site. Cliquez sur « Valider et Envoyer » pour expédier le mail à {sender_email}.",
            email_subject=subject,
            sender=sender_email,
            job_id=job_id,
        )

    return update_job_status(job_id) or job_record


def approve_and_send_manual_job(job_id: str) -> tuple[bool, str]:
    """
    En Mode Validation Manuelle, valide et envoie l'e-mail au client avec les derniers fichiers
    PDF et DOCX du projet (incluant d'éventuelles modifications faites sur le site).
    """
    state = load_automation_state()
    target_job = None
    for j in state.get("jobs", []):
        if j.get("job_id") == job_id:
            target_job = j
            break
    if not target_job:
        return False, "Job introuvable."

    pdf_path = target_job.get("pdf_path", "")
    docx_path = target_job.get("docx_path", "")
    if not os.path.exists(pdf_path) or not os.path.exists(docx_path):
        return False, "Les fichiers PDF/Word du mémoire sont introuvables dans archives/."

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()
    with open(docx_path, "rb") as f:
        docx_bytes = f.read()

    nom_fichier = re.sub(r"[^\w\s-]", "", target_job.get("nom_projet", "Memoire")).strip().replace(" ", "_")[:55] or "Memoire"
    ok_smtp, info_smtp = send_smtp_email(
        state,
        to_email=target_job["sender_email"],
        subject=target_job["email_subject_out"],
        plain_body=target_job["email_body_plain_out"],
        html_body=target_job["email_body_html_out"],
        attachments=[
            (f"Memoire_Technique_{nom_fichier}.pdf", pdf_bytes, "pdf"),
            (f"Memoire_Technique_{nom_fichier}.docx", docx_bytes, "docx"),
        ],
    )
    new_status = "Mémoire envoyé" if ok_smtp else "Erreur d'envoi mail"
    update_job_status(job_id, statut=new_status, smtp_info=info_smtp)
    add_automation_log(
        new_status,
        f"Validation manuelle effectuée : mémoire « {target_job.get('nom_projet')} » envoyé à {target_job['sender_email']}.",
        email_subject=target_job.get("subject", ""),
        sender=target_job.get("sender_email", ""),
        job_id=job_id,
    )
    return ok_smtp, info_smtp


def _build_fallback_structured_memoire(
    subject: str,
    lots_vises: list[str],
    consignes_specifiques: str,
    entreprise_profil: dict,
    classifications: list[dict],
    extracted_texts: dict[str, str],
) -> tuple[dict, str]:
    """
    Synthèse et rédaction structurée de secours lorsque le pipeline est testé sans clé API externe,
    en exploitant directement le texte extrait des pièces RC, CCTP, DPGF, Annexes et Entreprise.
    """
    lots_str = ", ".join(lots_vises) if lots_vises else "Lot VRD & Aménagements"
    nom_ent = (entreprise_profil or {}).get("nom") or "Entreprise Soumissionnaire BTP"
    nom_projet = re.sub(r"^(tr:|fw:|fwd:|re:)\s*", "", subject or "Marché de Travaux BTP", flags=re.IGNORECASE).strip()

    analyse = {
        "nom_projet": nom_projet,
        "maitre_ouvrage": "Maître d'Ouvrage (selon RC et CCTP)",
        "lieu_execution": "Site du projet (voir contraintes d'accès CCTP Lot 00)",
        "dates_cles": {
            "date_remise_offres": "Selon Règlement de Consultation (RC)",
            "delai_execution": "Selon planning d'exécution du marché (12 semaines)",
            "date_debut_prevue": "À notification de l'Ordre de Service (OS)",
        },
        "reglement_consultation_rc": {
            "criteres_valeur_technique": [
                {"ordre": 1, "critere": "Méthodologie d'exécution et compréhension des contraintes du site", "bareme_points": "25 pts", "attendus": "Phasage détaillé et gestion riverains"},
                {"ordre": 2, "critere": "Moyens humains et matériels dédiés au chantier", "bareme_points": "15 pts", "attendus": "Organigramme, CACES, AIPR, parc engins"},
                {"ordre": 3, "critere": "Qualité des matériaux, planning DPGF, sécurité PGC et environnement SOSED", "bareme_points": "20 pts", "attendus": "Fiches matériaux, cadences DPGF, filières déchets"},
            ],
            "plan_impose": [
                "1. Compréhension des enjeux et contraintes du chantier",
                "2. Installation de chantier et logistique",
                "3. Moyens humains et matériels dédiés",
                "4. Méthodologie d'exécution détaillée phase par phase et autocontrôles",
                "5. Provenance des matériaux et fiches techniques",
                "6. Planning prévisionnel et cadences basés sur la DPGF",
                "7. Sécurité et santé sur le chantier (PGC)",
                "8. Environnement et gestion des déchets (SOSED)",
            ],
            "nombre_pages_max": "30 pages maximum (selon RC)",
        },
        "contraintes_chantier": [
            "Maintien des accès riverains et secours en permanence",
            "Horaires de chantier réglementés (7h30 - 17h30) et limitation des nuisances sonores",
            "Repérage préalable des réseaux enterrés (DT-DICT) et marquage-piquetage",
        ],
        "lots_concernes": lots_vises or [lots_str],
        "exigences_techniques_par_lot": [
            {
                "lot": lots_str,
                "description_travaux": "Exécution complète des ouvrages conformément au CCTP et aux quantités de la DPGF",
                "materiaux_requis": ["Grave non traitée GNT 0/31.5 certifiée NF", "Enrobés bitumineux BBSG 0/10 NF EN 13108-1", "Bordures béton T2 NF P 98-340"],
                "normes_dtu": ["NF P 98-150 (Chaussées et enrobés)", "Fascicule 70 (Assainissement)", "DTU 12 (Terrassements)"],
                "methodes_pose_autocontroles": ["Essais à la plaque EV2 >= 80 MPa", "Contrôle de compacité au gammadensimètre"],
            }
        ],
        "annexes_securite_installation": {
            "risques_specifiques_pgc": [
                "Coactivité engins/piétons : balisage lourd type K16 et homme-trafic dédié",
                "Risque d'endommagement des réseaux concessionnaires : personnel certifié AIPR",
                "Blindage systématique des tranchées > 1,30 m",
            ]
        },
    }

    retained_names = [f"{c['filename']} ({c['categorie']})" for c in classifications if c["categorie"] != "IGNORED_CCTP"]
    memoire_md = (
        f"# 1. COMPRÉHENSION DES ENJEUX ET CONTRAINTES DU CHANTIER\n\n"
        f"Dans le cadre de la consultation relative à **{nom_projet}**, l'entreprise **{nom_ent}** soumissionne pour le **{lots_str}**.\n\n"
        f"- **Pièces du DCE analysées et croisées :** {', '.join(retained_names)}\n"
        f"- **Prise en compte de vos consignes spécifiques :** {consignes_specifiques or 'Respect intégral des prescriptions du RC et du CCTP.'}\n"
        f"- **Contraintes d'accès et riverains :** Maintien permanent des accès riverains et services de secours, horaires 7h30–17h30, signalisation temporaire de chantier conforme à l'Instruction Interministérielle sur la Signalisation Routière.\n\n"
        f"# 2. INSTALLATION DE CHANTIER ET LOGISTIQUE\n\n"
        f"- **Base vie et cantonnement :** Mise en place d'une roulotte de chantier autonome conforme aux préconisations du PGC Sécurité (vestiaires, réfectoire, sanitaires, trousse de premiers secours).\n"
        f"- **Zones de stockage et flux :** Aire de stockage clôturée par barrières Héras bâchées, plan de circulation interne séparant les flux piétons et les engins.\n\n"
        f"# 3. MOYENS HUMAINS ET MOYENS MATÉRIELS DÉDIÉS AU CHANTIER\n\n"
        f"- **Encadrement et équipes d'exécution :** Mobilisation d'une équipe dédiée issue de **{nom_ent}** disposant des habilitations **CACES R482**, **AIPR Opérateur & Encadrant** et **SST**.\n"
        f"- **Référent conducteur de travaux :** [À COMPLÉTER : préciser ici le nom et le téléphone direct du conducteur de travaux principal]\n"
        f"- **Parc matériel mobilisé :** Pelle hydraulique sur pneus 15T, mini-pelle 5T, compacteur tandem vibrant, camion ampliroll 19T, finisseur et laser double pente.\n\n"
        f"# 4. MÉTHODOLOGIE D'EXÉCUTION DÉTAILLÉE PHASE PAR PHASE ET AUTOCONTRÔLES\n\n"
        f"- **Phase 1 — Travaux préparatoires, DT-DICT et piquetage :** Marquage-piquetage contradictoire des réseaux existants, décapage et terrassement en déblais selon les quantités exactes de la DPGF.\n"
        f"- **Phase 2 — Structure d'assise et réseaux :** Mise en œuvre de la GNT 0/31.5 par couches compactées de 25 cm avec essai de portance à la plaque ($EV_2 \\ge 80\\text{{ MPa}}$).\n"
        f"- **Phase 3 — Pose des bordures et revêtements :** Pose des bordures béton sur lit de béton C16/20 et application des enrobés BBSG 0/10 conformément à la norme NF P 98-150.\n\n"
        f"# 5. PROVENANCE DES MATÉRIAUX, FOURNISSEURS ET FICHES TECHNIQUES\n\n"
        f"- Tous les granulats, bordures et enrobés bénéficient du marquage **CE / NF**.\n"
        f"- **Centrale d'enrobage et carrière partenaires :** [À COMPLÉTER : préciser ici le nom et la distance en km de la centrale d'enrobage et du fournisseur de granulats]\n\n"
        f"# 6. PLANNING PRÉVISIONNEL ET CADENCES BASÉS SUR LES QUANTITÉS DE LA DPGF\n\n"
        f"- Les cadences journalières sont dimensionnées directement à partir des quantités extraites de la DPGF du **{lots_str}** afin de garantir le respect strict du délai global d'exécution.\n\n"
        f"# 7. SÉCURITÉ ET SANTÉ SUR LE CHANTIER (PGC / PPSPS)\n\n"
        f"- Application stricte du **PGC Sécurité** : accueil sécurité obligatoire, quart d'heure sécurité hebdomadaire, port permanent des EPI (casque, gilet haute visibilité classe 2, chaussures S3, protections auditives) et blindage des fouilles.\n\n"
        f"# 8. ENVIRONNEMENT ET GESTION DES DÉCHETS (SOSED)\n\n"
        f"- Tri sélectif sur site (inertes, DIB, DID) avec suivi par Bordereaux de Suivi des Déchets (BSD / Trackdéchets).\n"
        f"- **Exutoire et centre de revalorisation des déblais :** [À COMPLÉTER : préciser ici le nom et la distance en km de la déchèterie / plateforme de recyclage partenaire]\n"
    )
    return analyse, memoire_md


# ─────────────────────────────────────────────────────────────────────────
# 6. ÉCOUTE AUTOMATIQUE DE LA BOÎTE MAIL (WORKER IMAP EN ARRIÈRE-PLAN)
# ─────────────────────────────────────────────────────────────────────────

def _decode_mime_words(raw_val: str) -> str:
    if not raw_val:
        return ""
    decoded_fragments = decode_header(raw_val)
    out = []
    for frag, enc in decoded_fragments:
        if isinstance(frag, bytes):
            try:
                out.append(frag.decode(enc or "utf-8", errors="replace"))
            except Exception:
                out.append(frag.decode("utf-8", errors="replace"))
        else:
            out.append(str(frag))
    return "".join(out).strip()


def _parse_email_message(msg: email.message.Message) -> tuple[str, str, str, str, list[tuple[str, bytes]]]:
    """Extrait (sender_email, sender_name, subject, body_text, attachments) d'un message MIME."""
    from_header = _decode_mime_words(msg.get("From", ""))
    sender_name, sender_email = parseaddr(from_header)
    subject = _decode_mime_words(msg.get("Subject", "Sans objet"))

    plain_parts: list[str] = []
    html_parts: list[str] = []
    attachments: list[tuple[str, bytes]] = []

    if msg.is_multipart():
        for part in msg.walk():
            content_disp = str(part.get("Content-Disposition", "") or "")
            filename = part.get_filename()
            if filename:
                filename = _decode_mime_words(filename)
                payload = part.get_payload(decode=True)
                if payload:
                    attachments.append((filename, payload))
                continue
            if "attachment" in content_disp.lower():
                continue
            ctype = part.get_content_type()
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                text_str = payload.decode(charset, errors="replace")
            except Exception:
                text_str = payload.decode("utf-8", errors="replace")
            if ctype == "text/plain":
                plain_parts.append(text_str)
            elif ctype == "text/html":
                html_parts.append(re.sub(r"<[^>]+>", " ", text_str))
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            plain_parts.append(payload.decode(charset, errors="replace"))

    body_text = "\n".join(plain_parts).strip() or "\n".join(html_parts).strip()
    return sender_email, sender_name, subject, body_text, attachments


def poll_imap_inbox_once(
    pipeline_callbacks: dict[str, Callable],
    load_config_fn: Callable[[], dict],
) -> tuple[int, Optional[str]]:
    """
    Vérifie une fois la boîte mail IMAP pour tout nouvel e-mail non lu (UNSEEN)
    et déclenche le pipeline complet sur chaque message trouvé.
    Retourne (nombre_de_mails_traites, erreur_eventuelle).
    """
    state = load_automation_state()
    imap_host = (state.get("imap_host") or "").strip()
    imap_port = int(state.get("imap_port") or 993)
    imap_user = (state.get("imap_user") or "").strip()
    imap_password = (state.get("imap_password") or "").strip()

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    if not imap_host or not imap_user or not imap_password:
        with _DB_LOCK:
            st_cur = load_automation_state()
            st_cur["last_check"] = now_str
            save_automation_state(st_cur)
        return 0, "Identifiants IMAP non renseignés (utilisez le simulateur ou configurez IMAP)."

    processed_count = 0
    try:
        mail = imaplib.IMAP4_SSL(imap_host, imap_port)
        mail.login(imap_user, imap_password)
        mail.select("INBOX")

        status, response = mail.search(None, "UNSEEN")
        if status != "OK" or not response or not response[0]:
            mail.logout()
            with _DB_LOCK:
                st_cur = load_automation_state()
                st_cur["last_check"] = now_str
                st_cur["last_error"] = None
                save_automation_state(st_cur)
            return 0, None

        msg_ids = response[0].split()
        config = load_config_fn()

        for num in msg_ids:
            res_fetch, data = mail.fetch(num, "(RFC822)")
            if res_fetch != "OK" or not data or not data[0]:
                continue
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            sender_email, sender_name, subject, body_text, attachments = _parse_email_message(msg)

            # Marquer comme lu pour éviter tout retraitement en boucle
            mail.store(num, "+FLAGS", "\\Seen")

            process_incoming_email_pipeline(
                sender_email=sender_email,
                sender_name=sender_name,
                subject=subject,
                body_text=body_text,
                raw_attachments=attachments,
                config=config,
                extract_text_fn=pipeline_callbacks["extract_text"],
                extract_category_fn=pipeline_callbacks["extract_category"],
                analyse_cctp_fn=pipeline_callbacks["analyse_cctp"],
                rediger_memoire_fn=pipeline_callbacks["rediger_memoire"],
                generate_docx_fn=pipeline_callbacks["generate_docx"],
                generate_pdf_fn=pipeline_callbacks["generate_pdf"],
                build_profil_text_fn=pipeline_callbacks["build_profil_text"],
                call_llm_fn=pipeline_callbacks.get("call_llm"),
                is_simulation=False,
            )
            processed_count += 1

        mail.logout()
        with _DB_LOCK:
            st_cur = load_automation_state()
            st_cur["last_check"] = now_str
            st_cur["last_error"] = None
            save_automation_state(st_cur)
        return processed_count, None
    except Exception as exc:
        err_msg = f"Erreur IMAP ({imap_host}) : {exc}"
        with _DB_LOCK:
            st_cur = load_automation_state()
            st_cur["last_check"] = now_str
            st_cur["last_error"] = err_msg
            save_automation_state(st_cur)
        return processed_count, err_msg


def ensure_background_imap_worker(
    pipeline_callbacks: dict[str, Callable],
    load_config_fn: Callable[[], dict],
) -> bool:
    """
    Démarre (s'il n'est pas déjà actif) le thread démon d'écoute automatique de la boîte mail.
    Le worker vérifie régulièrement si `watcher_enabled` est True dans `automation_state.json`.
    """
    global _WORKER_THREAD
    if _WORKER_THREAD is not None and _WORKER_THREAD.is_alive():
        return True

    _WORKER_STOP_EVENT.clear()

    def _worker_loop():
        while not _WORKER_STOP_EVENT.is_set():
            try:
                state = load_automation_state()
                if state.get("watcher_enabled", False):
                    poll_imap_inbox_once(pipeline_callbacks, load_config_fn)
                interval = max(10, int(state.get("poll_interval_seconds", 30) or 30))
            except Exception:
                interval = 30
            _WORKER_STOP_EVENT.wait(interval)

    _WORKER_THREAD = threading.Thread(target=_worker_loop, name="MarketBuildingIMAPWorker", daemon=True)
    _WORKER_THREAD.start()
    return True


# ─────────────────────────────────────────────────────────────────────────
# 7. GÉNÉRATEUR D'ARCHIVE .ZIP DCE BTP RÉALISTE POUR LE BOUTON "SIMULER UN MAIL"
# ─────────────────────────────────────────────────────────────────────────

def _make_pdf_bytes(title: str, lines: list[str]) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.multi_cell(0, 8, title.encode("latin-1", "replace").decode("latin-1"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    pdf.set_font("Helvetica", "", 10)
    for line in lines:
        safe_line = line.encode("latin-1", "replace").decode("latin-1")
        pdf.multi_cell(0, 6, safe_line, new_x="LMARGIN", new_y="NEXT")
    buf = io.BytesIO()
    pdf.output(buf)
    return buf.getvalue()


def _make_docx_bytes(title: str, paragraphs: list[str]) -> bytes:
    doc = Document()
    doc.add_heading(title, level=1)
    for p in paragraphs:
        doc.add_paragraph(p)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_realistic_dce_zip_simulation() -> tuple[str, bytes]:
    """
    Crée une archive .zip DCE complète et réaliste contenant 6 fichiers BTP en vrac :
    1. RC_Reglement_Consultation_Groupe_Scolaire.pdf (RC avec barème 60 pts et limite 25 pages)
    2. CCTP_Lot_00_Prescriptions_Generales.pdf (CCTP Général Lot 00)
    3. CCTP_Lot_01_Gros_Oeuvre.pdf (CCTP d'un autre lot -> doit être ignoré par l'auto-triage !)
    4. CCTP_Lot_02_VRD_Amenagements_Exterieurs.pdf (CCTP du lot visé par l'entreprise -> retenu !)
    5. DPGF_Lot_02_VRD_Quantites.csv (Quantités exactes m², m³, ml du Lot 02 VRD)
    6. PGC_Securite_Chantier.pdf (Annexe sécurité chantier)
    7. Moyens_Humains_Materiels_Entreprise_Durand_TP.docx (Document entreprise avec CACES, AIPR, engins)
    """
    rc_bytes = _make_pdf_bytes(
        "REGLEMENT DE LA CONSULTATION (RC) - REHABILITATION DU GROUPE SCOLAIRE JEAN MOULIN",
        [
            "Maître d'Ouvrage : Ville de Lyon (Direction de la Construction)",
            "Lieu d'exécution : 45 Avenue Jean Moulin, 69003 Lyon",
            "Date limite de remise des offres : 15 Novembre 2026 à 12h00",
            "Délai global d'exécution : 14 semaines (dont 2 semaines de préparation)",
            "CRITERES DE JUGEMENT DES OFFRES :",
            "- Prix des prestations : 40 points",
            "- Valeur Technique (Mémoire Technique) : 60 points répartis selon l'ordre impératif suivant :",
            "  Sous-critère 1 (20 pts) : Compréhension des contraintes du site occupé et installation de chantier",
            "  Sous-critère 2 (20 pts) : Moyens humains (CACES, AIPR) et matériels dédiés et méthodologie d'exécution VRD",
            "  Sous-critère 3 (10 pts) : Qualité et provenance des matériaux, planning basé sur la DPGF et cadences",
            "  Sous-critère 4 (10 pts) : Sécurité PGC en milieu scolaire et gestion des déchets SOSED",
            "LIMITE DE PAGES : Le mémoire technique ne devra pas dépasser 25 pages maximum.",
        ],
    )

    cctp_lot00_bytes = _make_pdf_bytes(
        "CCTP LOT 00 - PRESCRIPTIONS COMMUNES A TOUS LES LOTS",
        [
            "Opération : Réhabilitation et réaménagement de la cour du Groupe Scolaire Jean Moulin (Lyon 3e).",
            "Contraintes majeures du chantier :",
            "- Chantier jouxtant une école élémentaire : interdiction absolue de livraison entre 8h00-8h45 et 16h15-17h00.",
            "- Horaires autorisés : 7h30 à 17h30 du lundi au vendredi.",
            "- Clôture périphérique opaque de hauteur 2,00 m obligatoire autour de l'emprise travaux.",
            "- Gestion des déchets selon schéma SOSED avec suivi Trackdéchets et bennes de tri sur site.",
        ],
    )

    cctp_lot01_go_bytes = _make_pdf_bytes(
        "CCTP LOT 01 - GROS OEUVRE & MACONNERIE",
        [
            "Ce CCTP concerne exclusivement le Lot 01 Gros Oeuvre : reprises en sous-oeuvre, voiles béton armé C25/30 (120 m3) et planchers.",
            "Normes : DTU 20.1 et DTU 21.",
        ],
    )

    cctp_lot02_vrd_bytes = _make_pdf_bytes(
        "CCTP LOT 02 - VRD, COUR D'ECOLE PERMEABLE & RESEAUX HUMIDES",
        [
            "Objet du Lot 02 VRD : Désimperméabilisation de la cour du Groupe Scolaire Jean Moulin.",
            "1. Terrassement et décapage de l'enrobé existant sur 850 m2 (évacuation en filière de recyclage).",
            "2. Terrassement en déblais pour fond de forme sur 340 m3 (DTU 12).",
            "3. Mise en oeuvre d'une couche de fondation en Grave Non Traitée GNT 0/31.5 certifiée NF sur 850 m2 (épaisseur 25 cm), compactage avec essai à la plaque EV2 >= 80 MPa.",
            "4. Fourniture et pose de bordures béton préfabriquées type T2 NF P 98-340 sur 160 ml.",
            "5. Réalisation d'un revêtement drainant en enrobé perméable et béton désactivé clair sur 850 m2 conforme à la norme NF P 98-150.",
            "6. Pose de caniveaux hydrauliques à fente et réseau EP PVC CR8 diamètre 200 mm sur 95 ml (Fascicule 70).",
        ],
    )

    dpgf_csv_bytes = (
        "N° Poste,Désignation des ouvrages (Lot 02 VRD),Unité,Quantité,Prix Unitaire HT,Total HT\n"
        "2.1,Croûtage et démolition d'enrobé existant y compris évacuation,m²,850,14.50,12325.00\n"
        "2.2,Terrassement en déblais en pleine masse et réglage du fond de forme,m³,340,28.00,9520.00\n"
        "2.3,Fourniture et mise en oeuvre de GNT 0/31.5 NF ép. 25 cm + essais plaque EV2,m²,850,22.00,18700.00\n"
        "2.4,Fourniture et pose de bordures béton T2 NF P 98-340 sur lit béton,ml,160,36.00,5760.00\n"
        "2.5,Revêtement cour perméable (enrobé drainant / béton désactivé clair NF P 98-150),m²,850,48.00,40800.00\n"
        "2.6,Réseau eaux pluviales PVC CR8 DN200 + caniveaux (Fascicule 70),ml,95,65.00,6175.00\n"
    ).encode("utf-8")

    pgc_bytes = _make_pdf_bytes(
        "PLAN GENERAL DE COORDINATION EN MATIERE DE SECURITE ET DE PROTECTION DE LA SANTE (PGC SPS)",
        [
            "Opération : Groupe Scolaire Jean Moulin - Lyon 3e.",
            "Risques spécifiques identifiés au PGC :",
            "1. Proximité immédiate d'élèves et parents d'élèves : portail chantier cadenassé et homme-trafic obligatoire lors de toute manoeuvre de camion.",
            "2. Réseaux enterrés sensibles (gaz et Basse Tension Enedis) : DT-DICT, marquage-piquetage et habilitation AIPR obligatoires.",
            "3. Nuages de poussières lors du rabotage/terrassement : arrosage systématique à la lance d'eau.",
            "4. Installation base vie autonome raccordée en limite Est du plan de masse.",
        ],
    )

    ent_docx_bytes = _make_docx_bytes(
        "DOSSIER TECHNIQUE ENTREPRISE - DURAND TP & VRD (SAS)",
        [
            "Raison sociale : DURAND TP & VRD SAS — SIRET : 482 910 334 00028 — Siège : 18 Rue de l'Industrie, 69200 Vénissieux (à 6 km du chantier).",
            "Certifications détenues : FNTP Carte Professionnelle Travaux Publics, Qualibat 1312 (Terrassements - Fouilles) et 1321 (Canalisations), Label Ecovadis Argent.",
            "Encadrement dédié au chantier : Marc DURAND (Conducteur de travaux principal, Ingénieur ESTP, 14 ans d'expérience, AIPR Encadrant) et Karim BENALI (Chef de chantier VRD, 11 ans d'expérience, SST, AIPR Encadrant, CACES R482).",
            "Équipe d'exécution dédiée : 4 compagnons VRD qualifiés (2 conducteurs d'engins CACES R482 cat. B1/C1/F + AIPR Opérateur, 2 maçons VRD poseurs de bordures et canalisations).",
            "Parc matériel affecté au chantier : 1 Pelle sur pneus Mecalac 12MTX, 1 Mini-pelle Kubota 5,5T avec BRH, 1 Compacteur tandem vibrant Bomag BW 120, 2 Camions bennes 8x4 Renault Trucks C440 Euro 6, 1 Plaque vibrante avec capteurs de compacité, 1 Laser rotatif Leica.",
            "Fournisseurs habituels : Carrières Lafarge Granulats Saint-Fons (à 7 km) pour la GNT 0/31.5 NF, Centrale d'enrobage Eurovia Corbas (à 11 km), Point.P Vénissieux pour les bordures T2 NF et tubes PVC CR8.",
            "Exutoire et recyclage des déchets (SOSED) : Plateforme de valorisation des inertes Valomat Vénissieux située à 5,5 km du chantier.",
        ],
    )

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("DCE_Ecole_Jean_Moulin/01_RC_Reglement_Consultation.pdf", rc_bytes)
        zf.writestr("DCE_Ecole_Jean_Moulin/02_CCTP_Lot_00_Prescriptions_Generales.pdf", cctp_lot00_bytes)
        zf.writestr("DCE_Ecole_Jean_Moulin/03_CCTP_Lot_01_Gros_Oeuvre.pdf", cctp_lot01_go_bytes)
        zf.writestr("DCE_Ecole_Jean_Moulin/04_CCTP_Lot_02_VRD_Amenagements.pdf", cctp_lot02_vrd_bytes)
        zf.writestr("DCE_Ecole_Jean_Moulin/05_DPGF_Lot_02_VRD.csv", dpgf_csv_bytes)
        zf.writestr("DCE_Ecole_Jean_Moulin/06_PGC_Securite_SPS.pdf", pgc_bytes)
        zf.writestr("Pieces_Entreprise/Moyens_Humains_Materiels_Durand_TP.docx", ent_docx_bytes)

    return "DCE_Complet_Groupe_Scolaire_Jean_Moulin.zip", zip_buf.getvalue()
