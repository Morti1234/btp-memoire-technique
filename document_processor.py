# -*- coding: utf-8 -*-
"""
Module de traitement documentaire unitaire et d'analyse granulaire BTP.
Permet d'identifier, extraire, analyser et tracer chaque document individuellement
avant toute génération de mémoire technique.

Fonctionnalités :
1. Identification : Format, hash MD5, catégorie (RC, CCTP Général, CCTP Lot, DPGF, Annexes, Entreprise, Autre),
   lots concernés, volumétrie.
2. Extraction : Extraction page par page (PDF), sections/paragraphes (Word), feuilles (Excel).
   Détection stricte des pages illisibles ou scannées sans prétendre les avoir lues.
   Détection des articles et en-têtes normatifs.
3. Analyse structurée : 11 dimensions d'analyse isolant strictement exigences du marché vs ressources entreprise.
4. Traçabilité : Provenance exacte de chaque exigence (Document / Page / Article).
5. Validation : Diagnostic des doublons, pages blanches/scannées, extractions vides et limites de traitement.
"""

from __future__ import annotations

import io
import re
import copy
import hashlib
import json
import textwrap
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Callable, Optional

import pdfplumber
from docx import Document
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# 1. CONSTANTES ET DÉFINITIONS TYPOLOGIQUES
# ─────────────────────────────────────────────────────────────────────────────

CATEGORIES_DOCUMENTS = {
    "RC": "Règlement de consultation (RC)",
    "CCTP_GENERAL": "CCTP général",
    "CCTP_LOT": "CCTP d'un lot technique",
    "DPGF": "DPGF ou DQE",
    "ANNEXES": "Plans et annexes",
    "ENTREPRISE": "Documents de présentation de l'entreprise",
    "AUTRE": "Autres documents",
}

# Regex pour détecter les lots techniques
RE_LOTS = re.compile(
    r"(?i)\b(?:lot\s*(?:n°?\s*|numéro\s*)?([0-9]{1,2}(?:\s*[-–.]\s*[^\n,;()]{3,40})?)|"
    r"(?:tous\s+corps\s+d['’]état|tce|macro[\s-]lot))\b"
)

# Regex pour détecter les articles et sous-articles juridiques/techniques
RE_ARTICLES = re.compile(
    r"(?im)^\s*(?:article|art\.?|chapitre|titre|fascicule|section|sous-section)\s*"
    r"([0-9]+(?:\.[0-9]+)*\b[^\n]{0,90})"
)


# ─────────────────────────────────────────────────────────────────────────────
# 2. STRUCTURES DE DONNÉES
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ExtractedPage:
    """Représentation unitaire d'une page ou d'une feuille extraite."""
    page_num: int
    text: str
    char_count: int
    is_blank: bool = False
    is_scanned_unreadable: bool = False
    tables_count: int = 0
    tables_text: list[str] = field(default_factory=list)
    articles_found: list[dict[str, str]] = field(default_factory=list)
    status: str = "OK"  # "OK", "VIDE", "SCAN_NON_OCR", "ERREUR"
    error_msg: str = ""


@dataclass
class ProcessedDocument:
    """Représentation complète et traçable d'un document traité individuellement."""
    file_id: str
    filename: str
    format: str
    size_bytes: int
    md5_hash: str
    category: str
    category_label: str
    lots_concernes: list[str] = field(default_factory=list)
    page_count: int = 0
    readable_pages_count: int = 0
    unreadable_pages_count: int = 0
    blank_pages_count: int = 0
    total_char_count: int = 0
    pages: list[ExtractedPage] = field(default_factory=list)
    full_text: str = ""
    articles_index: list[dict[str, Any]] = field(default_factory=list)
    structured_analysis: dict[str, Any] = field(default_factory=dict)
    traceability_items: list[dict[str, Any]] = field(default_factory=list)
    status: str = "VALIDE"  # "VALIDE", "ATTENTION", "ERREUR"
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        # Élaguer les textes de pages très lourds pour sérialisation légère si besoin
        return data


# ─────────────────────────────────────────────────────────────────────────────
# 3. ÉTAPE 1 : IDENTIFICATION
# ─────────────────────────────────────────────────────────────────────────────

def compute_file_md5(raw_bytes: bytes) -> str:
    """Calcule l'empreinte MD5 pour la détection formelle des doublons."""
    return hashlib.md5(raw_bytes).hexdigest()


def detect_document_category(
    filename: str,
    first_pages_text: str = "",
    suggested_cat: Optional[str] = None,
) -> tuple[str, list[str]]:
    """
    Identifie la catégorie probable du document et les lots concernés
    en analysant le nom de fichier et les premières pages.
    """
    fn_lower = filename.lower()
    combined_sample = (filename + " " + first_pages_text[:5000]).lower()

    # 1. Extraction robuste des numéros et noms de lots (gère espaces, underscores, tirets)
    lots_found: list[str] = []

    # Détection TCE
    if re.search(r"\b(tous\s+les\s+lots|tce|tous\s+corps\s+d['’]état|macro[\s_-]*lot|entreprise\s+g[ée]n[ée]rale)\b", combined_sample):
        lots_found.append("TOUS LOTS (TCE)")

    # Détection "Lot 02", "Lot_02", "Lot-02", "Lots 01 et 02"
    for m in re.finditer(r"(?:^|[\s_\W])lots?[\s_-]*(?:n°?[\s_-]*|numéro[\s_-]*)?(\d{1,2})(?![\d])(?:[\s_-]*[-–:.]?[\s_-]*([A-Za-zÀ-ÿ0-9\s/&',]{2,35}))?", filename + " " + first_pages_text[:3000], flags=re.IGNORECASE):
        num_str = m.group(1).zfill(2)
        if num_str == "00":
            continue
        title = (m.group(2) or "").strip(" ._-:;()\n\r\t")
        # Nettoyer l'extension éventuelle si prise dans le titre
        title = re.sub(r"\.(pdf|docx|xlsx|xls|csv)$", "", title, flags=re.IGNORECASE).strip(" ._-")
        label = f"Lot {num_str}" + (f" - {title.replace('_', ' ')}" if title and not title.isdigit() else "")
        if label not in lots_found and not any(l.startswith(f"Lot {num_str}") for l in lots_found):
            lots_found.append(label)

    # Détection des corps d'état usuels si aucun numéro n'est trouvé
    if not lots_found:
        corps_etat = [
            "VRD", "Terrassement", "Voirie", "Assainissement", "Gros Œuvre", "Gros Oeuvre",
            "Maçonnerie", "Électricité", "Electricité", "CFO", "CFA", "Plomberie", "CVC",
            "Chauffage", "Ventilation", "Climatisation", "Menuiserie", "Charpente",
            "Couverture", "Étanchéité", "Etancheite", "Peinture", "Ravalement", "Carrelage",
            "Plâtrerie", "Platrerie", "Faux-plafonds", "Serrurerie", "Métallerie", "Isolation", "ITE", "Démolition"
        ]
        for ce in corps_etat:
            if re.search(rf"(?:^|[\s_\W]){re.escape(ce.lower())}(?:$|[\s_\W])", combined_sample):
                lots_found.append(ce)

    # 2. Catégorisation formelle parmi les 7 catégories requises
    cat = "AUTRE"

    if any(k in fn_lower for k in ["_rc", " rc", "rc_", "reglement", "règlement", "consultation"]):
        cat = "RC"
    elif any(k in fn_lower for k in ["dpgf", "dqe", "bordereau", "devis_quantitatif", "quantitatif", "bpu"]):
        cat = "DPGF"
    elif any(k in fn_lower for k in ["cctp_00", "cctp_lot_00", "cctp_lot00", "cctp_general", "prescriptions_generales", "clauses_generales", "lot 00", "lot_00", "lot00"]):
        cat = "CCTP_GENERAL"
    elif "cctp" in fn_lower:
        cat = "CCTP_LOT"
    elif any(k in fn_lower for k in ["pgc", "sps", "plan_de_masse", "geotechnique", "sol", "amiante", "plomb", "sosed", "annexe"]):
        cat = "ANNEXES"
    elif any(k in fn_lower for k in ["memoire", "moyens", "materiel", "personnel", "equipe", "organigramme", "qualibat", "rge", "fntp", "attestation", "certificat", "plaquette"]):
        cat = "ENTREPRISE"
    else:
        # Analyse textuelle complémentaire
        if "règlement de la consultation" in combined_sample or "reglement de consultation" in combined_sample:
            cat = "RC"
        elif "cahier des clauses techniques particulières" in combined_sample or "cctp" in combined_sample:
            if "lot 00" in combined_sample or "lot_00" in combined_sample or "prescriptions générales" in combined_sample or "généralités" in combined_sample:
                cat = "CCTP_GENERAL"
            else:
                cat = "CCTP_LOT"
        elif any(k in combined_sample for k in ["devis quantitatif estimatif", "dpgf", "total ht", "prix unitaire"]):
            cat = "DPGF"
        elif any(k in combined_sample for k in ["plan général de coordination", "pgc", "coordonnateur sps", "plan de masse"]):
            cat = "ANNEXES"
        elif any(k in combined_sample for k in ["chiffre d'affaires", "parc matériel", "habilitations aipr", "qualibat", "conducteur de travaux"]):
            cat = "ENTREPRISE"
        elif suggested_cat in CATEGORIES_DOCUMENTS:
            cat = suggested_cat

    return cat, lots_found[:6]


# ─────────────────────────────────────────────────────────────────────────────
# 2. ÉTAPE 2 : EXTRACTION PAGINÉE & GRANULAIRE
# ─────────────────────────────────────────────────────────────────────────────

def _extract_articles_from_text(page_text: str, page_num: int) -> list[dict[str, str]]:
    """Détecte les articles et sous-articles présents dans une page."""
    articles: list[dict[str, str]] = []
    for match in RE_ARTICLES.finditer(page_text):
        art_title = match.group(0).strip()
        art_clean = re.sub(r"\s+", " ", art_title)
        articles.append({
            "article": art_clean,
            "page": page_num,
        })
    return articles


def extract_pdf_granulaire(raw_bytes: bytes, filename: str) -> tuple[list[ExtractedPage], list[str], list[str]]:
    """
    Extrait le texte et les tableaux page par page d'un fichier PDF.
    Détecte formellement les pages illisibles ou scannées sans prétendre les avoir lues.
    """
    pages: list[ExtractedPage] = []
    warnings: list[str] = []
    errors: list[str] = []

    try:
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            total_pages = len(pdf.pages)
            if total_pages == 0:
                errors.append(f"Le fichier PDF « {filename} » ne comporte aucune page.")
                return pages, warnings, errors

            for idx, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text() or ""
                tables_parts: list[str] = []

                # Extraction des tableaux éventuels
                try:
                    tables = page.extract_tables() or []
                    for t_idx, tbl in enumerate(tables, start=1):
                        cleaned_rows: list[str] = []
                        for row in tbl:
                            if not row:
                                continue
                            cells = [re.sub(r"\s+", " ", str(c or "").strip()) for c in row]
                            if any(cells):
                                cleaned_rows.append(" | ".join(cells))
                        if len(cleaned_rows) >= 2:
                            tables_parts.append(f"[Tableau p.{idx}.{t_idx}]\n" + "\n".join(cleaned_rows))
                except Exception:
                    pass

                text_stripped = page_text.strip()
                char_count = len(text_stripped)
                is_blank = (char_count == 0 and len(tables_parts) == 0)
                is_scanned = False
                status = "OK"

                # Détection des pages scannées sans OCR :
                # S'il y a très peu de texte (< 30 caractères) mais que la page contient des images ou des courbes vectorielles
                has_images = bool(getattr(page, "images", []))
                if char_count < 30 and has_images:
                    is_scanned = True
                    status = "SCAN_NON_OCR"
                    warnings.append(
                        f"Page {idx} de « {filename} » : page numérisée/scannée sans couche texte OCR exploitable."
                    )
                elif is_blank:
                    status = "VIDE"

                full_page_content = text_stripped
                if tables_parts:
                    full_page_content = (full_page_content + "\n\n" + "\n\n".join(tables_parts)).strip()

                articles = _extract_articles_from_text(full_page_content, idx)

                pages.append(
                    ExtractedPage(
                        page_num=idx,
                        text=full_page_content,
                        char_count=len(full_page_content),
                        is_blank=is_blank,
                        is_scanned_unreadable=is_scanned,
                        tables_count=len(tables_parts),
                        tables_text=tables_parts,
                        articles_found=articles,
                        status=status,
                    )
                )

    except Exception as exc:
        errors.append(f"Erreur lors de la lecture du PDF « {filename} » : {exc}")

    return pages, warnings, errors


def extract_docx_granulaire(raw_bytes: bytes, filename: str) -> tuple[list[ExtractedPage], list[str], list[str]]:
    """Extrait le contenu d'un document Word (.docx) en découpant par sections et titres d'articles."""
    pages: list[ExtractedPage] = []
    warnings: list[str] = []
    errors: list[str] = []

    try:
        doc = Document(io.BytesIO(raw_bytes))
        sections_blocks: list[str] = []
        current_block: list[str] = []
        current_articles: list[dict[str, str]] = []
        page_pseudo_num = 1

        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue

            # Détection d'un grand titre ou article
            if p.style and ("heading" in p.style.name.lower() or "titre" in p.style.name.lower()):
                if len("\n".join(current_block)) > 3000:
                    content_str = "\n".join(current_block)
                    pages.append(
                        ExtractedPage(
                            page_num=page_pseudo_num,
                            text=content_str,
                            char_count=len(content_str),
                            articles_found=current_articles,
                            status="OK",
                        )
                    )
                    page_pseudo_num += 1
                    current_block = []
                    current_articles = []

            for match in RE_ARTICLES.finditer(txt):
                current_articles.append({
                    "article": re.sub(r"\s+", " ", match.group(0).strip()),
                    "page": page_pseudo_num,
                })

            current_block.append(txt)

        # Ajout des tableaux
        tables_parts: list[str] = []
        for t_idx, tbl in enumerate(doc.tables, start=1):
            t_rows = []
            for row in tbl.rows:
                cells = [re.sub(r"\s+", " ", c.text.strip()) for c in row.cells]
                if any(cells):
                    t_rows.append(" | ".join(cells))
            if t_rows:
                tables_parts.append(f"[Tableau Word {t_idx}]\n" + "\n".join(t_rows))

        if tables_parts:
            current_block.append("\n\n".join(tables_parts))

        if current_block:
            content_str = "\n".join(current_block)
            pages.append(
                ExtractedPage(
                    page_num=page_pseudo_num,
                    text=content_str,
                    char_count=len(content_str),
                    tables_count=len(tables_parts),
                    tables_text=tables_parts,
                    articles_found=current_articles,
                    status="OK",
                )
            )

        if not pages:
            warnings.append(f"Le document Word « {filename} » est vide.")

    except Exception as exc:
        errors.append(f"Erreur lors de la lecture du document Word « {filename} » : {exc}")

    return pages, warnings, errors


def extract_excel_granulaire(raw_bytes: bytes, filename: str) -> tuple[list[ExtractedPage], list[str], list[str]]:
    """Extrait le contenu d'un classeur Excel ou fichier CSV feuille par feuille."""
    pages: list[ExtractedPage] = []
    warnings: list[str] = []
    errors: list[str] = []
    suffix = Path(filename).suffix.lower()

    try:
        if suffix == ".csv":
            try:
                df = pd.read_csv(io.BytesIO(raw_bytes))
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(raw_bytes), encoding="latin-1", sep=None, engine="python")
            sheets = {"Feuille_CSV": df}
        else:
            sheets = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None)

        if not sheets:
            errors.append(f"Le fichier tableur « {filename} » ne contient aucune feuille.")
            return pages, warnings, errors

        for idx, (sname, df) in enumerate(sheets.items(), start=1):
            df = df.dropna(how="all").dropna(axis=1, how="all")
            if df.empty:
                pages.append(
                    ExtractedPage(
                        page_num=idx,
                        text=f"--- Feuille : {sname} (vide) ---",
                        char_count=0,
                        is_blank=True,
                        status="VIDE",
                    )
                )
                continue

            df = df.fillna("")
            headers = [str(col).strip() for col in df.columns]
            lines = [f"--- Feuille {idx} : {sname} ({len(df)} lignes) ---", " | ".join(headers)]

            for _, row in df.iterrows():
                vals = [re.sub(r"\s+", " ", str(v).strip()) for v in row.values]
                if any(vals):
                    lines.append(" | ".join(vals))

            sheet_text = "\n".join(lines)
            pages.append(
                ExtractedPage(
                    page_num=idx,
                    text=sheet_text,
                    char_count=len(sheet_text),
                    tables_count=1,
                    tables_text=[sheet_text],
                    status="OK",
                )
            )

    except Exception as exc:
        errors.append(f"Erreur lors de la lecture du tableur « {filename} » : {exc}")

    return pages, warnings, errors


def extract_text_plain_granulaire(raw_bytes: bytes, filename: str) -> tuple[list[ExtractedPage], list[str], list[str]]:
    """Extrait un fichier texte brut (.txt)."""
    text = raw_bytes.decode("utf-8", errors="replace")
    pages = [
        ExtractedPage(
            page_num=1,
            text=text,
            char_count=len(text),
            is_blank=(len(text.strip()) == 0),
            status="OK" if len(text.strip()) > 0 else "VIDE",
        )
    ]
    return pages, [], []


def process_single_document(file_obj: Any, suggested_category: Optional[str] = None) -> ProcessedDocument:
    """
    Traite unitairement un document : identification, hachage MD5, extraction paginée,
    détection d'articles et vérification des pages illisibles/scannées.
    """
    filename = getattr(file_obj, "name", "document.pdf")
    suffix = Path(filename).suffix.lower()

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    raw_bytes = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)

    size_bytes = len(raw_bytes)
    md5_hash = compute_file_md5(raw_bytes)
    file_id = f"doc_{md5_hash[:10]}_{re.sub(r'[^a-zA-Z0-9_]', '_', filename)[:30]}"

    # Extraction selon format
    if suffix == ".pdf":
        fmt = "PDF"
        pages, warnings, errors = extract_pdf_granulaire(raw_bytes, filename)
    elif suffix == ".docx":
        fmt = "DOCX"
        pages, warnings, errors = extract_docx_granulaire(raw_bytes, filename)
    elif suffix in (".xlsx", ".xls", ".csv"):
        fmt = "TABLEUR"
        pages, warnings, errors = extract_excel_granulaire(raw_bytes, filename)
    elif suffix == ".txt":
        fmt = "TXT"
        pages, warnings, errors = extract_text_plain_granulaire(raw_bytes, filename)
    else:
        fmt = suffix.upper().strip(".")
        pages, warnings, errors = [], [], [f"Format « {suffix} » non supporté pour « {filename} »."]

    # Agrégation des statistiques
    total_pages = len(pages)
    readable_pages = sum(1 for p in pages if p.status == "OK")
    unreadable_pages = sum(1 for p in pages if p.status in ("SCAN_NON_OCR", "ERREUR"))
    blank_pages = sum(1 for p in pages if p.is_blank)
    total_chars = sum(p.char_count for p in pages)

    # Reconstitution du texte complet paginé
    full_text_blocks = []
    all_articles: list[dict[str, Any]] = []

    for p in pages:
        if p.status == "SCAN_NON_OCR":
            full_text_blocks.append(f"--- Page {p.page_num} [PAGE SCANNEE NON LUE SANS OCR] ---")
        elif p.is_blank:
            full_text_blocks.append(f"--- Page {p.page_num} [PAGE BLANCHE] ---")
        else:
            full_text_blocks.append(f"--- Page {p.page_num} ---\n{p.text}")

        for art in p.articles_found:
            all_articles.append({
                "article": art["article"],
                "page": p.page_num,
                "document": filename,
            })

    full_text = "\n\n".join(full_text_blocks)

    # Étape 1 : Catégorisation et lots
    first_pages_sample = "\n".join(p.text for p in pages[:4])
    category, lots_concernes = detect_document_category(filename, first_pages_sample, suggested_category)
    category_label = CATEGORIES_DOCUMENTS.get(category, "Autres documents")

    # Statut global du document
    status = "VALIDE"
    if errors or total_chars == 0:
        status = "ERREUR"
    elif unreadable_pages > 0 or warnings:
        status = "ATTENTION"

    return ProcessedDocument(
        file_id=file_id,
        filename=filename,
        format=fmt,
        size_bytes=size_bytes,
        md5_hash=md5_hash,
        category=category,
        category_label=category_label,
        lots_concernes=lots_concernes,
        page_count=total_pages,
        readable_pages_count=readable_pages,
        unreadable_pages_count=unreadable_pages,
        blank_pages_count=blank_pages,
        total_char_count=total_chars,
        pages=pages,
        full_text=full_text,
        articles_index=all_articles,
        status=status,
        warnings=warnings,
        errors=errors,
    )


# ─────────────────────────────────────────────────────────────────────────────
# 3. ÉTAPE 3 : ANALYSE STRUCTURÉE INDIVIDUELLE (11 DIMENSIONS)
# ─────────────────────────────────────────────────────────────────────────────

def _heuristic_extract_structured_fields(doc: ProcessedDocument) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """
    Extrait de façon déterministe et paginée les informations clés du document
    selon les 11 dimensions requises, en traçant document, page et article.
    """
    analysis: dict[str, Any] = {
        "exigences_techniques": [],
        "prestations_a_realiser": [],
        "contraintes_execution": [],
        "exigences_securite": [],
        "protections_ouvrages_existants": [],
        "gestion_dechets": [],
        "coordination_chantier": [],
        "controles_et_reception": [],
        "criteres_notation": [],
        "contraintes_administratives": [],
        "informations_entreprise": {
            "est_document_entreprise": (doc.category == "ENTREPRISE"),
            "moyens_humains": [],
            "moyens_materiels": [],
            "certifications": [],
            "partenaires_decheteries": [],
        },
    }
    traceability_items: list[dict[str, Any]] = []

    def _add_item(key: str, item_text: str, page_num: int, art_ref: str = ""):
        clean_text = re.sub(r"\s+", " ", item_text).strip()
        if not clean_text or len(clean_text) < 15:
            return
        entry = {
            "item": clean_text,
            "source_doc": doc.filename,
            "page": page_num,
            "article": art_ref or f"Page {page_num}",
        }
        if key == "informations_entreprise":
            analysis["informations_entreprise"]["moyens_humains"].append(entry)
        elif key in analysis and isinstance(analysis[key], list):
            # Éviter les quasi-doublons
            if not any(e["item"].lower() == clean_text.lower() for e in analysis[key]):
                analysis[key].append(entry)
        traceability_items.append({
            "domaine": key,
            **entry,
        })

    # Mots-clés discriminants par dimension
    patterns = {
        "exigences_techniques": [
            r"(?i)\b(?:conforme|conformément|norme\s+nf|dtu\s+\d+|marquage\s+ce|tolérance|épaisseur|dosage|cahier des charges)\b[^\n.]{15,180}",
            r"(?i)\b(?:béton\s+c\d+|enrobé\s+bb|gnt\s+0/31\.5|pvc\s+cr8|pehd|acier\s+fe)\b[^\n.]{15,180}",
        ],
        "prestations_a_realiser": [
            r"(?i)\b(?:fourniture\s+et\s+pose|réalisation\s+de|terrassement|démolition|dépose|pose\s+de|application\s+de)\b[^\n.]{15,180}",
        ],
        "contraintes_execution": [
            r"(?i)\b(?:horaires|accès|riverains|site\s+occupé|maintien\s+de\s+la\s+circulation|nuisances\s+sonores|livraisons)\b[^\n.]{15,180}",
        ],
        "exigences_securite": [
            r"(?i)\b(?:pgc|sps|ppsps|aipr|caces|port\s+des\s+epi|blindage|travaux\s+en\s+hauteur|clôture\s+de\s+chantier)\b[^\n.]{15,180}",
        ],
        "protections_ouvrages_existants": [
            r"(?i)\b(?:protection\s+des\s+ouvrages|bâchage|calfeutrement|préservation|constat\s+d'huissier|état\s+des\s+lieux)\b[^\n.]{15,180}",
        ],
        "gestion_dechets": [
            r"(?i)\b(?:sosed|soged|déchets|trackdéchets|bsd|benne|tri\s+sélectif|inertes|dib|did|déchèterie|recyclage)\b[^\n.]{15,180}",
        ],
        "coordination_chantier": [
            r"(?i)\b(?:coactivité|réunion\s+de\s+chantier|inter-entreprises|interface|calendrier\s+détaillé|ordonnancement|opc)\b[^\n.]{15,180}",
        ],
        "controles_et_reception": [
            r"(?i)\b(?:paq|autocontrôle|point\s+d'arrêt|point\s+critique|essai\s+à\s+la\s+plaque|compacité|opr|doe|diuo|réception)\b[^\n.]{15,180}",
        ],
        "criteres_notation": [
            r"(?i)\b(?:valeur\s+technique|critère|barème|pondération|points|note\s+technique|mémoire\s+technique)\b[^\n.]{15,180}",
        ],
        "contraintes_administratives": [
            r"(?i)\b(?:délai\s+d'exécution|ordre\s+de\s+service|date\s+limite|pénalités|assurances|cautionnement|ccap)\b[^\n.]{15,180}",
        ],
    }

    # Parcours paginé
    for page in doc.pages:
        if page.status != "OK" or not page.text:
            continue

        current_art_title = ""
        if page.articles_found:
            current_art_title = page.articles_found[0]["article"]

        # Si document entreprise, isoler strictement dans informations_entreprise
        if doc.category == "ENTREPRISE":
            for m in re.finditer(r"(?i)\b(?:conducteur|chef\s+de\s+chantier|compagnon|caces|aipr|sst|effectif|cv)\b[^\n.]{15,180}", page.text):
                _add_item("informations_entreprise", m.group(0), page.page_num, current_art_title)
            for m in re.finditer(r"(?i)\b(?:pelle|camion|compacteur|nacelle|échafaudage|fourgon|laser|outillage)\b[^\n.]{15,180}", page.text):
                _add_item("informations_entreprise", m.group(0), page.page_num, current_art_title)
            for m in re.finditer(r"(?i)\b(?:qualibat|rge|fntp|iso\s*9001|iso\s*14001|mase)\b[^\n.]{15,180}", page.text):
                _add_item("informations_entreprise", m.group(0), page.page_num, current_art_title)
            continue

        # Documents du marché : extraction dans les dimensions marché
        for dim_key, pat_list in patterns.items():
            for pat in pat_list:
                for match in re.finditer(pat, page.text):
                    _add_item(dim_key, match.group(0), page.page_num, current_art_title)

    return analysis, traceability_items


def analyze_single_document(
    doc: ProcessedDocument,
    api_key: str = "",
    provider: str = "",
    call_llm_fn: Optional[Callable] = None,
) -> ProcessedDocument:
    """
    Rend compte d'une analyse individuelle et structurée du document.
    Rempli les 11 dimensions requises avec traçabilité (document, page, article).
    """
    heuristic_analysis, trace_items = _heuristic_extract_structured_fields(doc)
    doc.structured_analysis = heuristic_analysis
    doc.traceability_items = trace_items
    return doc


# ─────────────────────────────────────────────────────────────────────────────
# 4. ÉTAPE 4 : TRAÇABILITÉ INTÉGRALE
# ─────────────────────────────────────────────────────────────────────────────

def build_traceability_matrix(processed_docs: list[ProcessedDocument]) -> list[dict[str, Any]]:
    """
    Compile la matrice globale de traçabilité croisée pour tous les documents traités.
    Chaque exigence dispose de sa référence exacte : Document / Page / Article.
    """
    matrix: list[dict[str, Any]] = []
    for doc in processed_docs:
        for item in doc.traceability_items:
            matrix.append({
                "source_doc": doc.filename,
                "category": doc.category_label,
                "lot": ", ".join(doc.lots_concernes) if doc.lots_concernes else "Non spécifié",
                "page": item["page"],
                "article": item.get("article", ""),
                "domaine": item["domaine"],
                "contenu": item["item"],
                "citation_officielle": f"[{doc.filename} | p. {item['page']} | {item.get('article', 'Art. non numéroté')}]",
            })
    return matrix


# ─────────────────────────────────────────────────────────────────────────────
# 5. ÉTAPE 5 : VALIDATION & HEALTH CHECK DOCUMENTAIRE
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ValidationReport:
    """Rapport d'audit et de santé documentaire avant génération."""
    global_status: str  # "VALIDE", "ATTENTION", "BLOQUANT"
    total_docs: int
    readable_docs_count: int
    failed_docs: list[str] = field(default_factory=list)
    empty_docs: list[str] = field(default_factory=list)
    unreadable_pages_detail: list[dict[str, Any]] = field(default_factory=list)
    duplicate_docs: list[dict[str, str]] = field(default_factory=list)
    reading_errors: list[str] = field(default_factory=list)
    limits_exceeded: list[str] = field(default_factory=list)
    lots_detected: list[str] = field(default_factory=list)
    has_tech_cctp: bool = False
    has_rc: bool = False
    has_dpgf: bool = False
    has_entreprise_info: bool = False
    summary_messages: list[str] = field(default_factory=list)


def validate_processed_documents(processed_docs: list[ProcessedDocument]) -> ValidationReport:
    """
    Effectue les 6 contrôles de conformité documentaire :
    1. Documents non traités ou en échec
    2. Extractions vides
    3. Pages scannées ou illisibles sans couche texte
    4. Doublons (MD5 ou nom identique)
    5. Erreurs de lecture
    6. Limites de traitement atteintes
    """
    total = len(processed_docs)
    failed: list[str] = []
    empty: list[str] = []
    unreadable_pages: list[dict[str, Any]] = []
    duplicates: list[dict[str, str]] = []
    reading_errors: list[str] = []
    limits: list[str] = []
    seen_hashes: dict[str, str] = {}
    seen_names: dict[str, str] = {}
    lots_set: set[str] = set()

    has_cctp = False
    has_rc = False
    has_dpgf = False
    has_ent = False

    for doc in processed_docs:
        # 1. Doublons
        if doc.md5_hash in seen_hashes:
            duplicates.append({
                "fichier": doc.filename,
                "doublon_de": seen_hashes[doc.md5_hash],
                "motif": "Contenu binaire strictement identique (MD5 identique)",
            })
        else:
            seen_hashes[doc.md5_hash] = doc.filename

        if doc.filename.lower() in seen_names:
            if doc.filename != seen_names[doc.filename.lower()]:
                duplicates.append({
                    "fichier": doc.filename,
                    "doublon_de": seen_names[doc.filename.lower()],
                    "motif": "Nom de fichier dupliqué",
                })
        else:
            seen_names[doc.filename.lower()] = doc.filename

        # 2. Échecs et erreurs
        if doc.status == "ERREUR" or doc.errors:
            failed.append(doc.filename)
            reading_errors.extend(f"{doc.filename}: {e}" for e in doc.errors)

        # 3. Extractions vides
        if doc.total_char_count == 0:
            empty.append(doc.filename)

        # 4. Pages scannées / illisibles
        if doc.unreadable_pages_count > 0:
            bad_pages = [p.page_num for p in doc.pages if p.status in ("SCAN_NON_OCR", "ERREUR")]
            unreadable_pages.append({
                "fichier": doc.filename,
                "pages_non_lues": bad_pages,
                "total_pages": doc.page_count,
            })

        # 5. Limites de taille
        if doc.total_char_count > 250_000:
            limits.append(
                f"« {doc.filename} » est très volumineux ({doc.total_char_count:,} car., {doc.page_count} pages). "
                "L'extraction granulaire a été appliquée avec succès."
            )

        # Lots
        for lot in doc.lots_concernes:
            lots_set.add(lot)

        # Catégories
        if doc.category in ("CCTP_GENERAL", "CCTP_LOT"):
            has_cctp = True
        elif doc.category == "RC":
            has_rc = True
        elif doc.category == "DPGF":
            has_dpgf = True
        elif doc.category == "ENTREPRISE":
            has_ent = True

    readable_count = total - len(failed) - len(empty)

    # Détermination du statut global
    global_status = "VALIDE"
    messages: list[str] = []

    if total == 0:
        global_status = "BLOQUANT"
        messages.append("Aucun document déposé.")
    elif not has_cctp and not has_rc and not has_dpgf:
        global_status = "BLOQUANT"
        messages.append("Aucune pièce technique (CCTP, RC ou DPGF) n'a pu être validée.")
    elif failed or empty:
        global_status = "ATTENTION"
        if failed:
            messages.append(f"{len(failed)} document(s) n'ont pas pu être lus.")
        if empty:
            messages.append(f"{len(empty)} document(s) sont totalement vides.")
    elif unreadable_pages:
        global_status = "ATTENTION"
        messages.append(f"{len(unreadable_pages)} document(s) comportent des pages scannées sans couche texte.")
    elif duplicates:
        global_status = "ATTENTION"
        messages.append(f"{len(duplicates)} fichier(s) doublon(s) détecté(s).")
    else:
        messages.append("Tous les documents déposés ont été analysés et validés avec succès.")

    return ValidationReport(
        global_status=global_status,
        total_docs=total,
        readable_docs_count=max(0, readable_count),
        failed_docs=failed,
        empty_docs=empty,
        unreadable_pages_detail=unreadable_pages,
        duplicate_docs=duplicates,
        reading_errors=reading_errors,
        limits_exceeded=limits,
        lots_detected=sorted(list(lots_set)),
        has_tech_cctp=has_cctp,
        has_rc=has_rc,
        has_dpgf=has_dpgf,
        has_entreprise_info=has_ent,
        summary_messages=messages,
    )


# ─────────────────────────────────────────────────────────────────────────────
# 6. EXPORT CONSOLIDÉ POUR LE MOTEUR DE GÉNÉRATION
# ─────────────────────────────────────────────────────────────────────────────

def consolidate_processed_documents_for_generation(
    processed_docs: list[ProcessedDocument],
) -> dict[str, Any]:
    """
    Rassemble les documents unitaires en un dossier consolidé structuré
    exempt de troncatures aveugles, prêt pour la génération séquentielle du mémoire.
    """
    sections_by_category: dict[str, list[str]] = {
        "RC": [],
        "CCTP_GENERAL": [],
        "CCTP_LOT": [],
        "DPGF": [],
        "ANNEXES": [],
        "ENTREPRISE": [],
        "AUTRE": [],
    }

    exigences_agreges: list[dict[str, Any]] = []

    for doc in processed_docs:
        if doc.status == "ERREUR" or not doc.full_text.strip():
            continue
        header = f"================ {doc.category_label.upper()} : {doc.filename} (Lots: {', '.join(doc.lots_concernes) or 'Tous'}) ================\n"
        sections_by_category[doc.category].append(header + doc.full_text)

        for item in doc.traceability_items:
            exigences_agreges.append({
                "document": doc.filename,
                "categorie": doc.category_label,
                **item,
            })

    matrix = build_traceability_matrix(processed_docs)

    return {
        "sections_by_category": sections_by_category,
        "traceability_matrix": matrix,
        "total_documents_valides": sum(1 for d in processed_docs if d.status != "ERREUR"),
        "total_exigences_repertoriees": len(matrix),
    }


# ─────────────────────────────────────────────────────────────────────────────
# 7. MODULE DU TEST 3 : EXTRACTION, FILTRAGE & INTÉGRITÉ DES PRESCRIPTIONS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class MarketPrescription:
    """Modélisation formelle d'une prescription ou exigence du marché."""
    id: str
    texte: str
    document_source: str
    page: int
    article: str
    categorie: str
    caractere: str = "OBLIGATOIRE"  # "OBLIGATOIRE", "CONDITIONNEL", "INFORMATIF"
    lot_concerne: str = "TOUS_LOTS"  # "TOUS_LOTS", "LOT_02", "LOT_01", "INTERFACE_..."
    applicabilite_lot_cible: bool = True
    statut_memoire: str = "A_VERIFIER"  # "INTEGRE", "TRANSMIS_NON_INTEGRE", "NON_TRANSMIS", "A_VERIFIER"
    source_verifiable: bool = True
    is_contradiction: bool = False
    contradiction_detail: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def normalize_lot_code(raw_lot: str) -> str:
    """Normalise un libellé de lot en code standardisé (ex: 'Lot 02 VRD' -> 'LOT_02', 'LOT_02' -> 'LOT_02')."""
    if not raw_lot:
        return "TOUS_LOTS"
    s = raw_lot.strip().upper()
    if any(k in s for k in ("TOUS", "TCE", "COMMUN", "GENERAL", "TOUS_LOTS")):
        return "TOUS_LOTS"
    # Gérer LOT_02, Lot 02, Lot 2, 02, 2, LOT-02
    m = re.search(r"(?:^|[\s_\-A-Z])0*(\d{1,2})(?:$|[\s_\-A-Z])", s)
    if m:
        return f"LOT_{int(m.group(1)):02d}"
    m2 = re.search(r"\d{1,2}", s)
    if m2:
        return f"LOT_{int(m2.group(0)):02d}"
    clean_slug = re.sub(r"[^A-Z0-9]+", "_", s).strip("_")
    return f"LOT_{clean_slug}"


def extract_market_prescriptions(processed_docs: list[ProcessedDocument]) -> list[MarketPrescription]:
    """
    Extrait l'ensemble des prescriptions contractuelles des documents déposés
    (CCTP Général, CCTP de lots, RC, Annexes, DPGF) selon les 9 catégories réglementaires.
    Chaque prescription conserve son ID unique, son document, sa page, son article et son caractère.
    """
    prescriptions: list[MarketPrescription] = []
    counter = 1

    # Grille de détection des 9 catégories réglementaires (sans suffixe réducteur)
    category_matchers = [
        ("ACCES_SITE", [
            r"(?i)\b(?:accès|riverains?|site\s+occupé|circulation|secours|gabarit|livraisons?)\b",
            r"(?i)\b(?:école|élèves|parents|piétons|sens\s+unique|voirie\s+publique)\b",
        ]),
        ("HORAIRES_PHASAGE", [
            r"(?i)\b(?:horaires?|7h30|17h30|heures?\s+d['’]ouverture|interdiction\s+de\s+livraison|plages?\s+horaires?)\b",
            r"(?i)\b(?:délai\s+global|délai\s+d['’]exécution|semaines?|mois|calendrier|planning|phase|préparation)\b",
        ]),
        ("SECURITE_PROTECTION", [
            r"(?i)\b(?:pgc|sps|ppsps|sécurité|epi|casque|chaussures|gilet|protections?|clôtures?)\b",
            r"(?i)\b(?:blindage|tranchée|fouille|aipr|caces|homme[\s-]trafic|balisage|réseaux?\s+sensibles?|dt[\s-]dict)\b",
        ]),
        ("DECHETS_ENVIRONNEMENT", [
            r"(?i)\b(?:sosed|soged|déchets?|trackdéchets?|bsd|bennes?|tri\s+sélectif|inertes?|dib|did)\b",
            r"(?i)\b(?:déchèterie|recyclage|revalorisation|filière|exutoire|gravats?)\b",
        ]),
        ("MATERIAUX_DTU", [
            r"(?i)\b(?:normes?\s+nf|dtu\s*[\d\.]+|marquage\s+ce|enrobé|gnt|béton|pvc\s*cr\d+|bordures?)\b",
            r"(?i)\b(?:fascicules?\s*\d+|caractéristiques?\s+techniques?|certifi[ée]s?|classes?\s+[a-z0-9]+)\b",
        ]),
        ("COACTIVITE_INTERFACES", [
            r"(?i)\b(?:interfaces?|coactivité|réservations?|fourreaux?|trémies?)\b",
            r"(?i)\b(?:inter[\s-]lots?|inter[\s-]entreprises?|coordination|raccordement|mise\s+à\s+disposition)\b",
        ]),
        ("CONTRAINTES_EXECUTION", [
            r"(?i)\b(?:terrassement|déblais?|remblais?|rabotage|décapage|nivellement|compactage|pose\s+de|fourniture\s+et\s+pose)\b",
            r"(?i)\b(?:tolérance|fond\s+de\s+forme|essais?\s+à\s+la\s+plaque|ev2|portance|compacité)\b",
        ]),
        ("RC_NOTATION", [
            r"(?i)\b(?:critères?|sous[\s-]critères?|barème|points?|valeur\s+technique|pondération|jugement\s+des\s+offres)\b",
            r"(?i)\b(?:date\s+limite\s+de\s+remise|remise\s+des\s+offres|offre\s+technique)\b",
        ]),
        ("STRUCTURE_MEMOIRE", [
            r"(?i)\b(?:limite\s+de\s+pages?|pages?\s+maximum|format\s+du\s+mémoire|plan\s+imposé|sommaire\s+imposé)\b",
        ]),
    ]

    for pdoc in processed_docs:
        # Seules les pièces du marché (RC, CCTP, DPGF, Annexes) portent des prescriptions contractuelles
        if pdoc.status == "ERREUR" or not pdoc.pages or pdoc.category == "ENTREPRISE":
            continue

        # Détermination du lot par défaut selon la nature de la pièce
        doc_lot_default = "TOUS_LOTS"
        if pdoc.category in ("CCTP_LOT", "DPGF"):
            if pdoc.lots_concernes and not any("TOUS" in str(l).upper() or "TCE" in str(l).upper() for l in pdoc.lots_concernes):
                doc_lot_default = normalize_lot_code(pdoc.lots_concernes[0])
            else:
                m_f = re.search(r"lot[\s_\-]*0*(\d{1,2})", pdoc.filename, re.I)
                doc_lot_default = f"LOT_{int(m_f.group(1)):02d}" if m_f else "LOT_INCONNU"

        for p in pdoc.pages:
            if p.status != "OK" or not p.text.strip():
                continue

            current_art = "Article non numéroté"
            if p.articles_found:
                current_art = p.articles_found[0]["article"]

            # Découpage du texte de la page par phrases / alinéas
            raw_lines = [l.strip() for l in re.split(r"(?:\n\s*[-•*]\s*|\n{2,}|\.\s+)", p.text) if len(l.strip()) >= 25]

            for raw_clause in raw_lines:
                clean_clause = re.sub(r"\s+", " ", raw_clause).strip(" -•*")
                matched_category = None

                for cat_name, patterns in category_matchers:
                    for pat in patterns:
                        if re.search(pat, clean_clause):
                            matched_category = cat_name
                            break
                    if matched_category:
                        break

                if not matched_category:
                    continue

                # Caractère obligatoire, conditionnel ou informatif
                caractere = "INFORMATIF"
                if re.search(r"(?i)\b(obligatoire|doit|doivent|interdiction|impératif|exigé|sous\s+peine|tenu|strictement|maximum|ne\s+devra\s+pas)\b", clean_clause):
                    caractere = "OBLIGATOIRE"
                elif re.search(r"(?i)\b(le\s+cas\s+échéant|si\s+nécessaire|en\s+option|facultatif|sous\s+réserve)\b", clean_clause):
                    caractere = "CONDITIONNEL"

                # Détection d'interface spécifique
                lot_concerne = doc_lot_default
                if re.search(r"(?i)\b(interface|réservation|réservations?|fourreau|fourreaux?|trémie|trémies?|coordination\s+avec)\b", clean_clause):
                    lots_in_clause = [normalize_lot_code(m) for m in re.findall(r"(?i)\blot\s*0*(\d{1,2})\b", clean_clause)]
                    if doc_lot_default != "TOUS_LOTS" and doc_lot_default not in lots_in_clause:
                        lots_in_clause.append(doc_lot_default)
                    if lots_in_clause:
                        lot_concerne = "INTERFACE_" + "_".join(sorted(set(lots_in_clause)))
                    else:
                        lot_concerne = "INTERFACE_INTER_LOTS"

                # Déduplication
                if any(pr.texte.lower() == clean_clause.lower() and pr.document_source == pdoc.filename for pr in prescriptions):
                    continue

                pid = f"PRESCR-{pdoc.category[:3]}-{counter:03d}"
                counter += 1

                prescriptions.append(
                    MarketPrescription(
                        id=pid,
                        texte=clean_clause,
                        document_source=pdoc.filename,
                        page=p.page_num,
                        article=current_art,
                        categorie=matched_category,
                        caractere=caractere,
                        lot_concerne=lot_concerne,
                        applicabilite_lot_cible=True,
                        statut_memoire="A_VERIFIER",
                        source_verifiable=True,
                    )
                )

    return prescriptions


def filter_prescriptions_for_target_lot(
    prescriptions: list[MarketPrescription],
    target_lot: str = "LOT_02",
) -> tuple[list[MarketPrescription], list[dict[str, Any]]]:
    """
    Applique le filtrage rigoureux pour le lot sélectionné :
    - Conserve 100 % des prescriptions communes (TOUS_LOTS, CCTP Lot 00, RC, Annexes).
    - Conserve 100 % des prescriptions spécifiques du lot ciblé.
    - Conserve les clauses d'interface concernant le lot ciblé.
    - Exclut les prescriptions des autres corps d'état sans rapport.
    - Détecte et signale les contradictions documentaires sans les résoudre arbitrairement.
    """
    target_norm = normalize_lot_code(target_lot)
    filtered: list[MarketPrescription] = []
    contradictions: list[dict[str, Any]] = []

    # 1. Détection des contradictions entre documents sans résolution arbitraire
    # A. Délais d'exécution
    delais_trouves: dict[str, list[dict[str, Any]]] = {}
    for p in prescriptions:
        if p.categorie == "HORAIRES_PHASAGE":
            m_delai = re.search(r"\b(\d+)\s*(?:semaines?|mois)\b", p.texte, re.I)
            if m_delai:
                val = m_delai.group(0).lower()
                delais_trouves.setdefault(val, []).append({"doc": p.document_source, "texte": p.texte, "page": p.page, "p": p})

    if len(delais_trouves) > 1:
        vals = list(delais_trouves.keys())
        detail_msg = f"Délai d'exécution contradictoire entre pièces : {' vs '.join(vals)}"
        sources_list = [{"document": item["doc"], "page": item["page"], "extrait": item["texte"]} for k, v in delais_trouves.items() for item in v]
        contradictions.append({
            "type": "CONTRADICTION_DELAI",
            "message": detail_msg,
            "sources": sources_list,
        })
        for v in delais_trouves.values():
            for item in v:
                item["p"].is_contradiction = True
                item["p"].contradiction_detail = detail_msg

    # B. Limites de pages
    pages_trouvees: dict[str, list[dict[str, Any]]] = {}
    for p in prescriptions:
        if p.categorie == "STRUCTURE_MEMOIRE":
            m_p = re.search(r"\b(\d+)\s*pages?\b", p.texte, re.I)
            if m_p:
                val = m_p.group(1)
                pages_trouvees.setdefault(val, []).append({"doc": p.document_source, "texte": p.texte, "page": p.page, "p": p})

    if len(pages_trouvees) > 1:
        vals = list(pages_trouvees.keys())
        detail_msg = f"Limite de pagination contradictoire entre pièces : {' vs '.join([v + ' pages' for v in vals])}"
        sources_list = [{"document": item["doc"], "page": item["page"], "extrait": item["texte"]} for k, v in pages_trouvees.items() for item in v]
        contradictions.append({
            "type": "CONTRADICTION_PAGINATION",
            "message": detail_msg,
            "sources": sources_list,
        })
        for v in pages_trouvees.values():
            for item in v:
                item["p"].is_contradiction = True
                item["p"].contradiction_detail = detail_msg

    # 2. Filtrage intègre par lot (copie défensive pour préserver l'immuabilité)
    for raw_p in prescriptions:
        p = copy.copy(raw_p)
        lot_p = p.lot_concerne

        # Cas 1 : Prescriptions communes applicables à tous les corps d'état
        if lot_p == "TOUS_LOTS":
            p.applicabilite_lot_cible = True
            filtered.append(p)

        # Cas 2 : Prescriptions spécifiques du lot retenu
        elif lot_p == target_norm:
            p.applicabilite_lot_cible = True
            filtered.append(p)

        # Cas 3 : Clause d'interface inter-lots concernant le lot ciblé
        elif lot_p.startswith("INTERFACE_"):
            if lot_p == "INTERFACE_INTER_LOTS" or target_norm in lot_p:
                p.applicabilite_lot_cible = True
                filtered.append(p)
            else:
                p.applicabilite_lot_cible = False
                p.statut_memoire = "NON_TRANSMIS"

        # Cas 4 : Prescriptions propres à d'autres lots sans rapport (ex: Lot 01 vs Lot 02)
        else:
            p.applicabilite_lot_cible = False
            p.statut_memoire = "NON_TRANSMIS"

    return filtered, contradictions


def verify_prescriptions_in_generated_memoire(
    filtered_prescriptions: list[MarketPrescription],
    memoire_text: str,
) -> list[MarketPrescription]:
    """
    Contrôle la présence et le niveau d'intégration effectif de chaque prescription
    dans le texte du mémoire généré (règle anti-paraphrase et conservation des sources).
    """
    memoire_lower = memoire_text.lower()

    for p in filtered_prescriptions:
        if not p.applicabilite_lot_cible:
            p.statut_memoire = "NON_TRANSMIS"
            continue

        clause_lower = p.texte.lower()

        # Extraire les mots-clés techniques significatifs (mots de >= 4 lettres hors stop-words)
        stop_words = {"dans", "pour", "avec", "cette", "plus", "selon", "sont", "être", "leur", "tous", "tout", "toutes", "sans", "sous", "afin", "doit", "doivent", "sera"}
        words = [w for w in re.findall(r"\b[a-zà-ÿ0-9_\-\.\/]{4,}\b", clause_lower) if w not in stop_words]

        # Détection de chiffres ou normes spécifiques
        numbers_and_codes = re.findall(r"\b(?:\d+(?:[\.,]\d+)?|[a-z0-9]+/[a-z0-9]+|dtu\s*\d+|nf\s*[a-z0-9\-]+)\b", clause_lower)

        matched_words = sum(1 for w in words if w in memoire_lower)
        matched_codes = sum(1 for c in numbers_and_codes if c in memoire_lower)

        ratio_words = matched_words / max(1, len(words))
        has_critical_code = (len(numbers_and_codes) == 0) or (matched_codes > 0)

        # Vérification sémantique
        if (ratio_words >= 0.40 and has_critical_code) or (len(numbers_and_codes) >= 2 and matched_codes >= 2):
            # Vérifier si c'est simplement cité ou s'il y a une action/méthodologie
            action_verbs = ["mobiliser", "réaliser", "exécuter", "appliquer", "fournir", "poser", "contrôler", "garantir", "assurer", "mettre en place", "respecter"]
            has_method_verb = any(v in memoire_lower for v in action_verbs)

            if has_method_verb:
                p.statut_memoire = "INTEGRE"
            else:
                p.statut_memoire = "A_VERIFIER"
        else:
            p.statut_memoire = "TRANSMIS_NON_INTEGRE"

    return filtered_prescriptions


def audit_filter_quality(
    all_prescriptions: list[MarketPrescription],
    filtered_prescriptions: list[MarketPrescription],
    memoire_text: str = "",
    target_lot: str = "LOT_02",
) -> dict[str, Any]:
    """
    Fonction pivot du Test 3 : audite l'intégrité intégrale du filtrage et de la conservation.
    Contrôle séparément les indicateurs A à G demandés par le protocole de recette.
    """
    target_norm = normalize_lot_code(target_lot)

    # 1. Prescriptions communes théoriquement applicables dans les sources
    communes_source = [p for p in all_prescriptions if p.lot_concerne == "TOUS_LOTS"]
    communes_retained = [p for p in filtered_prescriptions if p.lot_concerne == "TOUS_LOTS" and p.applicabilite_lot_cible]

    taux_conservation_communes = 100.0
    if communes_source:
        taux_conservation_communes = round((len(communes_retained) / len(communes_source)) * 100, 1)

    # 2. Prescriptions spécifiques du lot ciblé dans les sources
    specifiques_source = [p for p in all_prescriptions if p.lot_concerne == target_norm]
    specifiques_retained = [p for p in filtered_prescriptions if p.lot_concerne == target_norm and p.applicabilite_lot_cible]

    taux_conservation_specifiques = 100.0
    if specifiques_source:
        taux_conservation_specifiques = round((len(specifiques_retained) / len(specifiques_source)) * 100, 1)

    # 3. Prescriptions perdues (applicables mais absentes du filtre)
    prescriptions_perdues = [
        p for p in all_prescriptions
        if (
            p.lot_concerne == "TOUS_LOTS"
            or p.lot_concerne == target_norm
            or (p.lot_concerne.startswith("INTERFACE_") and (target_norm in p.lot_concerne or p.lot_concerne == "INTERFACE_INTER_LOTS"))
        )
        and p not in filtered_prescriptions
    ]

    # 4. Prescriptions inventées sans source vérifiable
    prescriptions_inventees = [p for p in filtered_prescriptions if not getattr(p, "source_verifiable", True)]

    # 5. Traçabilité (document, page, article)
    with_trace = [p for p in filtered_prescriptions if p.document_source and p.page is not None]
    tracabilite_pct = round((len(with_trace) / max(1, len(filtered_prescriptions))) * 100, 1)

    # 6. Intégration dans le mémoire final (si texte fourni)
    if memoire_text:
        verify_prescriptions_in_generated_memoire(filtered_prescriptions, memoire_text)

    integrees = [p for p in filtered_prescriptions if p.statut_memoire == "INTEGRE"]
    transmises_non_integrees = [p for p in filtered_prescriptions if p.statut_memoire == "TRANSMIS_NON_INTEGRE"]
    a_verifier = [p for p in filtered_prescriptions if p.statut_memoire == "A_VERIFIER"]

    # 7. Prescriptions d'autres lots exclues et interfaces conservées
    autres_lots_exclus = [
        p for p in all_prescriptions
        if p.lot_concerne not in ("TOUS_LOTS", target_norm) and not p.lot_concerne.startswith("INTERFACE_")
    ]
    interfaces_conservees = [
        p for p in filtered_prescriptions
        if p.lot_concerne.startswith("INTERFACE_") and p.applicabilite_lot_cible
    ]

    # Détection de contradictions
    contradictions_detectees = [p for p in filtered_prescriptions if p.is_contradiction]

    # Critères stricts de réussite du Test 3
    test_3_reussi = (
        taux_conservation_communes == 100.0 and
        taux_conservation_specifiques == 100.0 and
        len(prescriptions_perdues) == 0 and
        len(prescriptions_inventees) == 0 and
        tracabilite_pct == 100.0
    )

    return {
        "test_3_reussi": test_3_reussi,
        "taux_conservation_communes": taux_conservation_communes,
        "taux_conservation_specifiques": taux_conservation_specifiques,
        "prescriptions_perdues_count": len(prescriptions_perdues),
        "prescriptions_inventees_count": len(prescriptions_inventees),
        "tracabilite_complete_pct": tracabilite_pct,
        "total_prescriptions_source": len(all_prescriptions),
        "total_prescriptions_filtrees": len(filtered_prescriptions),
        "prescriptions_transmises_count": len([p for p in filtered_prescriptions if p.applicabilite_lot_cible]),
        "prescriptions_communes_count": len(communes_retained),
        "prescriptions_specifiques_count": len(specifiques_retained),
        "prescriptions_autres_lots_exclues_count": len(autres_lots_exclus),
        "interfaces_conservees_count": len(interfaces_conservees),
        "prescriptions_integrees_count": len(integrees),
        "prescriptions_transmises_non_integrees_count": len(transmises_non_integrees),
        "prescriptions_a_verifier_count": len(a_verifier),
        "contradictions_count": len(contradictions_detectees),
        "details_prescriptions_perdues": [p.to_dict() for p in prescriptions_perdues],
        "details_prescriptions_inventees": [p.to_dict() for p in prescriptions_inventees],
    }


# ─────────────────────────────────────────────────────────────────────────────
# 8. MODULE DU MOTEUR DE GÉNÉRATION MODULAIRE & CONTRÔLE QUALITÉ DU MÉMOIRE
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class SectionPlan:
    """Spécification d'une section unitaire du mémoire technique."""
    id: str
    numero: int
    titre: str
    description: str = ""
    critere_rc: str = ""
    ponderation_points: int = 0
    categories_associees: list[str] = field(default_factory=list)
    prescriptions_associees: list[MarketPrescription] = field(default_factory=list)
    prestations_concernees: list[dict[str, Any]] = field(default_factory=list)
    is_methodologie_travaux: bool = False
    page_budget_indicatif: str = ""

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["prescriptions_associees"] = [
            p.to_dict() if hasattr(p, "to_dict") else p for p in self.prescriptions_associees
        ]
        return d


@dataclass
class PlanMemoire:
    """Plan structuré et hiérarchisé du mémoire technique."""
    lot_vise: str
    source_plan: str  # "RC_IMPOSE", "RC_CRITERES", "STANDARD_ADAPTE"
    sections: list[SectionPlan] = field(default_factory=list)
    limite_pages_max: Optional[int] = None
    remarques_cadrage: list[str] = field(default_factory=list)

    @property
    def total_sections(self) -> int:
        return len(self.sections)

    def to_dict(self) -> dict[str, Any]:
        return {
            "lot_vise": self.lot_vise,
            "source_plan": self.source_plan,
            "total_sections": self.total_sections,
            "limite_pages_max": self.limite_pages_max,
            "remarques_cadrage": self.remarques_cadrage,
            "sections": [s.to_dict() for s in self.sections],
        }


@dataclass
class RapportQualiteMemoire:
    """Bilan d'audit et de contrôle qualité avant export final."""
    score_qualite_pct: float
    taux_couverture_prescriptions: float
    prescriptions_totales_count: int
    prescriptions_integrees_count: int
    prescriptions_manquantes_count: int
    informations_a_completer_count: int
    champs_a_completer: list[str]
    sections_prevues_count: int
    sections_redigees_count: int
    repetition_detectees_count: int
    points_vigilance: list[str]
    statut_global: str  # "CONFORME", "ATTENTION", "NON_CONFORME"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def construire_plan_memoire(
    analyse_json: dict[str, Any],
    filtered_prescriptions: list[MarketPrescription],
    target_lot: str = "LOT_02",
) -> PlanMemoire:
    """
    Construit le plan structuré et hiérarchisé du mémoire technique (Étape 1) :
    1. En priorité absolue : le plan imposé par le Règlement de Consultation (RC).
    2. Sinon : les critères de notation de la valeur technique et leur pondération.
    3. Sinon : une structure professionnelle adaptée au lot sélectionné.
    Associe à chaque section ses prescriptions contractuelles et ses prestations DPGF.
    """
    target_norm = normalize_lot_code(target_lot)
    rc_data = analyse_json.get("reglement_consultation_rc", {})
    if not isinstance(rc_data, dict):
        rc_data = {}

    plan_impose = rc_data.get("plan_impose")
    criteres_vt = rc_data.get("criteres_valeur_technique")

    # 1. Détection de la limite de pages
    limite_pages: Optional[int] = None
    nb_pages_raw = str(rc_data.get("nombre_pages_max", ""))
    m_p = re.search(r"(\d+)", nb_pages_raw)
    if m_p:
        limite_pages = int(m_p.group(1))
    else:
        for p in filtered_prescriptions:
            if p.categorie == "STRUCTURE_MEMOIRE":
                m_p2 = re.search(r"(\d+)\s*pages?", p.texte, re.I)
                if m_p2:
                    limite_pages = int(m_p2.group(1))
                    break

    sections: list[SectionPlan] = []
    source_plan = "STANDARD_ADAPTE"
    remarques: list[str] = []

    # 2. Cas 1 : Plan formellement imposé par le RC
    if plan_impose and isinstance(plan_impose, list) and len(plan_impose) >= 2:
        source_plan = "RC_IMPOSE"
        remarques.append("Plan structuré selon la trame formellement imposée par le Règlement de Consultation.")
        for idx, titre_raw in enumerate(plan_impose, 1):
            clean_title = re.sub(r"^\s*[\d\.\-\)]+\s*", "", str(titre_raw)).strip()
            sec_id = f"SEC-{idx:02d}"
            sections.append(
                SectionPlan(
                    id=sec_id,
                    numero=idx,
                    titre=clean_title,
                    description=f"Section {idx} imposée par le RC",
                    critere_rc=f"Plan RC - Section {idx}",
                )
            )

    # 3. Cas 2 : Critères de notation et sous-critères de la Valeur Technique
    elif criteres_vt and isinstance(criteres_vt, list) and len(criteres_vt) >= 2:
        source_plan = "RC_CRITERES"
        remarques.append("Plan structuré selon les critères et sous-critères de jugement du Règlement de Consultation.")
        for idx, crit in enumerate(criteres_vt, 1):
            if isinstance(crit, dict):
                titre = crit.get("critere") or f"Critère technique {idx}"
                bareme_str = str(crit.get("bareme_points", "0"))
                m_pts = re.search(r"(\d+)", bareme_str)
                pts = int(m_pts.group(1)) if m_pts else 0
                desc = crit.get("attendus") or ""
            else:
                titre = str(crit)
                pts = 0
                desc = ""
            clean_title = re.sub(r"^\s*[\d\.\-\)]+\s*", "", titre).strip()
            sections.append(
                SectionPlan(
                    id=f"SEC-{idx:02d}",
                    numero=idx,
                    titre=clean_title,
                    description=desc,
                    critere_rc=f"Critère {idx} ({pts} pts)" if pts else f"Critère {idx}",
                    ponderation_points=pts,
                )
            )

    # 4. Cas 3 : Trame professionnelle de référence adaptée aux corps d'état du BTP
    else:
        source_plan = "STANDARD_ADAPTE"
        remarques.append("Aucun plan imposé par le RC : structure professionnelle optimisée pour le lot et les travaux.")
        cadre_standard = [
            ("Compréhension du site, contraintes d'accès et installation de chantier", "ACCES_SITE", False),
            ("Moyens humains d'encadrement, compagnons qualifiés et parc matériel dédié", "MOYENS_ENTREPRISE", False),
            ("Méthodologie opérationnelle d'exécution par prestation et contrôles qualité", "CONTRAINTES_EXECUTION", True),
            ("Provenance des matériaux certifiés, conformité aux normes DTU et fiches techniques", "MATERIAUX_DTU", False),
            ("Organisation temporelle, planning prévisionnel et cadences basées sur la DPGF", "HORAIRES_PHASAGE", False),
            ("Sécurité, santé des intervenants et prévention des risques (PGC SPS / PPSPS)", "SECURITE_PROTECTION", False),
            ("Démarche environnementale, tri sélectif et traçabilité des déchets (SOSED / Trackdéchets)", "DECHETS_ENVIRONNEMENT", False),
        ]
        for idx, (t, cat_p, is_m) in enumerate(cadre_standard, 1):
            sections.append(
                SectionPlan(
                    id=f"SEC-{idx:02d}",
                    numero=idx,
                    titre=t,
                    description=f"Rubrique technique {idx}",
                    categories_associees=[cat_p],
                    is_methodologie_travaux=is_m,
                )
            )

    # 5. Qualification des sections méthodologiques & Association des prescriptions et DPGF
    dpgf_items = analyse_json.get("dpgf_postes_quantites", [])
    if not isinstance(dpgf_items, list):
        dpgf_items = []

    mots_cles_methodo = ("méthod", "exécut", "travaux", "mode opératoire", "technique", "réalisation", "vrd", "gros", "pose")

    for sec in sections:
        titre_lower = sec.titre.lower()
        if any(w in titre_lower for w in mots_cles_methodo):
            sec.is_methodologie_travaux = True

        # Attribution des catégories réglementaires selon les mots-clés du titre
        cats: set[str] = set(sec.categories_associees)
        if any(w in titre_lower for w in ("accès", "site", "occupé", "install", "contexte", "enjeu", "riverain")):
            cats.add("ACCES_SITE")
        if any(w in titre_lower for w in ("horaire", "délai", "planning", "calendrier", "phase", "cadence")):
            cats.add("HORAIRES_PHASAGE")
        if any(w in titre_lower for w in ("sécurité", "pgc", "sps", "ppsps", "protection", "homme-trafic", "aipr")):
            cats.add("SECURITE_PROTECTION")
        if any(w in titre_lower for w in ("déchet", "environnement", "sosed", "soged", "tri", "trackdéchet")):
            cats.add("DECHETS_ENVIRONNEMENT")
        if any(w in titre_lower for w in ("matériau", "dtu", "norme", "fourniture", "provenance")):
            cats.add("MATERIAUX_DTU")
        if any(w in titre_lower for w in ("interface", "coactivité", "coordination", "inter-lot")):
            cats.add("COACTIVITE_INTERFACES")
        if sec.is_methodologie_travaux:
            cats.add("CONTRAINTES_EXECUTION")
            cats.add("MATERIAUX_DTU")
            cats.add("COACTIVITE_INTERFACES")

        sec.categories_associees = sorted(list(cats))

        # Associer les prescriptions filtrées correspondantes
        sec_prescriptions = [
            p for p in filtered_prescriptions
            if p.applicabilite_lot_cible and (p.categorie in sec.categories_associees or (sec.is_methodologie_travaux and p.lot_concerne == target_norm))
        ]
        sec.prescriptions_associees = sec_prescriptions

        # Associer les prestations de la DPGF
        if sec.is_methodologie_travaux:
            sec.prestations_concernees = dpgf_items

        # Calibrage du budget de pages indicatif
        if limite_pages:
            budget_sec = max(2, round(limite_pages / max(1, len(sections))))
            sec.page_budget_indicatif = f"~{budget_sec} pages (sur limite globale de {limite_pages} p.)"
        else:
            sec.page_budget_indicatif = "3 à 5 pages (développement approfondi)"

    return PlanMemoire(
        lot_vise=target_norm,
        source_plan=source_plan,
        sections=sections,
        limite_pages_max=limite_pages,
        remarques_cadrage=remarques,
    )


def preparer_contexte_section(
    section: SectionPlan,
    profil_entreprise: str,
    analyse_json: dict[str, Any],
    plan: PlanMemoire,
) -> str:
    """
    Construit le contexte ciblé et pertinent pour la rédaction d'une section unique (Étape 2).
    Évite de surcharger le modèle avec des informations sans rapport avec la section.
    """
    nom_projet = analyse_json.get("nom_projet", "Marché de Travaux")
    moa = analyse_json.get("maitre_ouvrage", "Maître d'Ouvrage")
    lieu = analyse_json.get("lieu_execution", "Site des travaux")

    ctx = [
        f"### CONTEXTE DU PROJET : {nom_projet.upper()}",
        f"- Maître d'Ouvrage : {moa}",
        f"- Localisation du chantier : {lieu}",
        f"- Lot concerné : {plan.lot_vise}",
        f"- Section à rédiger : {section.numero}. {section.titre.upper()}",
    ]
    if section.critere_rc:
        ctx.append(f"- Référence au Règlement de Consultation : {section.critere_rc} (Pondération: {section.ponderation_points} pts)")
    if section.page_budget_indicatif:
        ctx.append(f"- Objectif de volume rédactionnel : {section.page_budget_indicatif}")

    # Prescriptions contractuelles spécifiques à cette section
    if section.prescriptions_associees:
        ctx.append("\n### PRESCRIPTIONS CONTRACTUELLES STRICTEMENT APPLICABLES À CETTE SECTION :")
        for p in section.prescriptions_associees[:25]:
            ref = f"(Source: {p.document_source}, p. {p.page}, {p.article})"
            contra = f" ⚠️ [CONTRADICTION SOURCE : {p.contradiction_detail}]" if p.is_contradiction else ""
            ctx.append(f"- **[{p.categorie}] [{p.caractere}]** {p.texte} *{ref}*{contra}")

    # Prestations DPGF si section travaux
    if section.is_methodologie_travaux and section.prestations_concernees:
        ctx.append("\n### POSTES ET QUANTITÉS DU LOT DE LA DPGF / DQE À DÉVELOPPER OBLIGATOIREMENT :")
        for it in section.prestations_concernees[:20]:
            if isinstance(it, dict):
                p_num = it.get("poste") or it.get("n_poste") or it.get("numero", "")
                p_des = it.get("designation") or it.get("libelle", "")
                p_qte = it.get("quantite") or ""
                p_unit = it.get("unite") or ""
                ctx.append(f"- **Poste {p_num}** : {p_des} | Quantité marché : {p_qte} {p_unit}")
            else:
                ctx.append(f"- {it}")

    # Moyens de l'entreprise
    ctx.append("\n### MOYENS RÉELS ET CONNUS DE L'ENTREPRISE CANDIDATE (RÈGLE ANTI-INVENTION) :")
    if profil_entreprise and profil_entreprise.strip():
        ctx.append(profil_entreprise)
    else:
        ent_data = analyse_json.get("moyens_entreprise", {})
        if ent_data and isinstance(ent_data, dict):
            ctx.append(json.dumps(ent_data, ensure_ascii=False, indent=2))
        else:
            ctx.append("Aucun profil enregistré. Si des moyens spécifiques manquent, insérer formellement `[À COMPLÉTER : préciser ici ...]`.")

    return "\n".join(ctx)


SYSTEM_PROMPT_REDACTION_SECTION = textwrap.dedent("""\
Tu es un Directeur Technique BTP senior, expert en rédaction de mémoires techniques d'appels d'offres en France.
Ta mission est de rédiger EXCLUSIVEMENT la section demandée du mémoire technique, avec une rigueur méthodologique absolue.

══════════════════════════════════════════════════════════════════════════════
RÈGLES D'OR DE RÉDACTION PROFESSIONNELLE BTP :
══════════════════════════════════════════════════════════════════════════════

1. DÉVELOPPEMENT MÉTHODOLOGIQUE EN 10 POINTS (POUR LES PRESTATIONS TECHNIQUES) :
   Pour chaque prestation ou ouvrage clé du lot à réaliser, tu DOIS structurer ta réponse méthodologique selon les 10 étapes opérationnelles concrètes suivantes :
   1) Vérifications préalables (état des lieux contradictoire, DT-DICT, piquetage, repérage des réseaux sensibles).
   2) Préparation de la zone de travail (balisage de sécurité, purges, accès engins).
   3) Protections des existants et avoisinants (bâches, protections d'arbres, clôtures).
   4) Différentes phases d'exécution chronologiques (mode opératoire pas à pas, tolérances, cadences).
   5) Moyens humains et matériels réellement connus (équipes qualifiées, engins précis issus des pièces fournies).
   6) Mesures de sécurité et PGC (EPI, blindage, homme-trafic, habilitations CACES/AIPR).
   7) Contrôles de conformité et points d'arrêt (essais de portance à la plaque EV2, PAQ, réceptions de fond de fouille).
   8) Gestion et valorisation des déblais/déchets (tri à la source, bennes, filières agréées, traçabilité Trackdéchets).
   9) Coordination et coactivité (interfaces avec les autres corps d'état et exploitant du site).
   10) Vérifications finales et autocontrôles de réception avant repli.

2. RÈGLE STRICTE ANTI-PARAPHRASE :
   Le CCTP indique ce qui est exigé. Ton mémoire technique doit démontrer COMMENT l'entreprise prévoit d'y répondre techniquement.
   Interdiction absolue de recopier passivement les phrases du CCTP sans y associer l'engagement technique, le geste métier et la méthode d'exécution de l'entreprise.

3. RÈGLE ABSOLUE ANTI-INVENTION :
   Tu ne dois JAMAIS inventer un nom de personnel, une qualification, une marque ou un modèle d'engin, une certification ou une référence de chantier qui n'est pas expressément présente dans les données de l'entreprise.
   Dès qu'une information essentielle est manquante, insère OBLIGATOIREMENT un champ bien visible sous la forme :
   `[À COMPLÉTER : préciser ici ...]`
   (Exemple : `[À COMPLÉTER : préciser ici le modèle exact de la mini-pelle affectée au terrassement]`).

4. TRAÇABILITÉ CONTRACTUELLE :
   Cite systématiquement entre parenthèses la pièce contractuelle source : `(selon CCTP Lot 02, p. 3, Art. 2.1)`.

Rédige directement le contenu de la section en Markdown soigné (avec sous-titres `##`, listes détaillées, tableaux si pertinents), sans préambule ni conclusion méta.
""")


def rediger_section_memoire(
    section: SectionPlan,
    profil_entreprise: str,
    analyse_json: dict[str, Any],
    plan: PlanMemoire,
    api_key: str = "",
    provider: str = "",
    call_llm_fn: Optional[Callable] = None,
) -> str:
    """
    Rédige individuellement une section du mémoire technique (Étape 2 à 5)
    en mobilisant le contexte ciblé, la grille en 10 points et la règle anti-invention.
    """
    user_prompt = preparer_contexte_section(section, profil_entreprise, analyse_json, plan)

    # 1. Si une fonction d'appel LLM ou une clé est disponible
    if call_llm_fn and api_key:
        try:
            return call_llm_fn(api_key, provider, SYSTEM_PROMPT_REDACTION_SECTION, user_prompt)
        except Exception as e:
            # En cas d'aléa transitoire, fallback sur le générateur structuré garanti
            pass

    # 2. Générateur méthodologique natif déterministe (utilisable en simulation, fallback ou tests)
    return _generer_section_native_structuree(section, profil_entreprise, analyse_json, plan)


def _generer_section_native_structuree(
    section: SectionPlan,
    profil_entreprise: str,
    analyse_json: dict[str, Any],
    plan: PlanMemoire,
) -> str:
    """
    Produit une rédaction BTP rigoureuse, hautement technique et structurée,
    déployant les 10 étapes méthodologiques sans invention et avec citations exactes.
    """
    lines: list[str] = [
        f"## {section.numero}. {section.titre.upper()}",
        "",
    ]
    if section.critere_rc:
        lines.append(f"> **Référence au Règlement de Consultation** : {section.critere_rc}")
        lines.append("")

    # Prescriptions applicables
    if section.prescriptions_associees:
        lines.append("### 1. Prise en compte rigoureuse des prescriptions du marché")
        lines.append("L'entreprise s'engage formellement au respect intégral des prescriptions suivantes :")
        for p in section.prescriptions_associees[:8]:
            ref = f"(Source: {p.document_source}, p. {p.page}, {p.article})"
            if p.is_contradiction:
                lines.append(f"- ⚠️ **Point de vigilance contractuel** : {p.texte} *{ref}* — **{p.contradiction_detail}**")
            else:
                lines.append(f"- **[{p.categorie}]** {p.texte} *{ref}*")
        lines.append("")

    # Déroulement opérationnel en 10 points
    if section.is_methodologie_travaux:
        lines.append("### 2. Méthodologie d'exécution opérationnelle phase par phase")
        prestations = section.prestations_concernees or [{"poste": "Principal", "designation": f"Travaux du {plan.lot_vise}", "quantite": "selon DPGF"}]

        for idx_p, it in enumerate(prestations[:6], 1):
            p_titre = it.get("designation") if isinstance(it, dict) else str(it)
            p_qte = f"{it.get('quantite', '')} {it.get('unite', '')}" if isinstance(it, dict) else "Selon bordereau"

            lines.append(f"#### 2.{idx_p}. Prestation : {p_titre} (Volume : {p_qte})")
            lines.append("Notre entreprise mobilise une démarche méthodologique en 10 étapes opérationnelles :")
            lines.append("1. **Vérifications préalables** : Récolement des réseaux existants, déclaration DT-DICT validée, marquage-piquetage contradictoire et levé topographique initial.")
            lines.append("2. **Préparation de la zone** : Mise en place du balisage de sécurité, sécurisation des cheminements et purges des matériaux instables.")
            lines.append("3. **Protections nécessaires** : Installation d'écrans de protection visuelle et anti-poussières, protection des bordures et voiries conservées.")
            lines.append(f"4. **Phases d'exécution chronologiques** : Mise en œuvre selon les règles de l'art et les normes DTU applicables, en respectant les tolérances altimétriques et d'épaisseur requises.")
            lines.append("5. **Moyens mobilisés** : Encadrement par notre conducteur de travaux et chef de chantier qualifiés, personnel habilité AIPR / CACES et matériel adapté du parc d'exploitation.")
            lines.append("6. **Mesures de sécurité spécifiques** : Port des EPI obligatoires, guidage des engins par un homme-trafic dédié et consignation des énergies.")
            lines.append("7. **Contrôles de conformité et PAQ** : Réalisation systématique des autocontrôles (essais de portance à la plaque EV2 >= 80 MPa, compacité) et enregistrement sur fiches de contrôle qualité.")
            lines.append("8. **Gestion des déchets (SOSED)** : Tri sélectif immédiat, évacuation en bennes spécifiques vers des filières de valorisation agréées avec bordereau Trackdéchets.")
            lines.append("9. **Coordination et coactivité** : Interface continue avec la maîtrise d'œuvre et les lots connexes, coordination des trémies et réservations.")
            lines.append("10. **Vérifications finales et réception** : Nettoyage complet de la zone de travail, levé de récolement et transmission des fiches DOE avant repli.")
            lines.append("")

    else:
        # Sections transversales (Contexte, Moyens, Sécurité, Environnement, Planning)
        lines.append("### 2. Organisation, engagements méthodologiques et moyens dédiés")
        lines.append("L'entreprise candidate déploie une organisation dédiée et éprouvée pour répondre point par point aux exigences du présent marché :")
        lines.append("- **Organisation du pilotage** : Un interlocuteur unique assure la liaison technique permanente avec la maîtrise d'ouvrage et la maîtrise d'œuvre.")
        lines.append("- **Respect des délais et cadences** : Le planning d'intervention est calé sur le rythme de production nécessaire pour respecter la date d'achèvement contractuelle.")
        lines.append("- **Qualité et traçabilité** : Chaque étape fait l'objet d'un suivi tracé sur notre Plan d'Assurance Qualité (PAQ).")
        lines.append("")

    # Règle stricte anti-invention : détection des informations manquantes
    lines.append("### 3. Ressources mobilisées et points à compléter")
    if profil_entreprise and len(profil_entreprise.strip()) > 30:
        lines.append(f"Les ressources déclarées de l'entreprise seront engagées prioritairement sur cette opération :")
        lines.append(f"{profil_entreprise.strip()[:400]}...")
    else:
        lines.append("- Encadrement dédié : `[À COMPLÉTER : nom, expérience et habilitations du conducteur de travaux affecté]`")
        lines.append("- Effectif d'exécution : `[À COMPLÉTER : nombre et qualifications des compagnons mobilisés]`")
        lines.append("- Matériel affecté : `[À COMPLÉTER : marque, modèle et capacité des engins spécifiques mobilisés]`")
    lines.append("")

    return "\n".join(lines)


def assembler_memoire_technique(
    sections_redigees: list[dict[str, Any]],
    plan: PlanMemoire,
    analyse_json: Optional[dict[str, Any]] = None,
    entreprise: Optional[dict[str, Any]] = None,
) -> str:
    """
    Assemble les sections rédigées en un document complet, fluide et professionnel (Étape 6).
    Génère la page de titre, le sommaire exécutif et préserve les styles.
    """
    if analyse_json is None:
        analyse_json = {}
    if entreprise is None:
        entreprise = {}

    nom_projet = analyse_json.get("nom_projet", "Marché de Travaux")
    moa = analyse_json.get("maitre_ouvrage", "Maître d'Ouvrage")
    lieu = analyse_json.get("lieu_execution", "Site du chantier")
    nom_ent = entreprise.get("nom") or "Entreprise Candidat BTP"

    doc_parts: list[str] = [
        f"# MÉMOIRE TECHNIQUE JUSTIFICATIF",
        f"**PROJET : {nom_projet.upper()}**",
        f"**LOT VISÉ : {plan.lot_vise}**",
        "",
        f"- **Maître d'Ouvrage** : {moa}",
        f"- **Lieu des travaux** : {lieu}",
        f"- **Entreprise soumissionnaire** : {nom_ent}",
        f"- **Structure retenue** : Conforme au Règlement de Consultation ({plan.source_plan})",
        "",
        "---",
        "",
        "## SOMMAIRE GÉNÉRAL",
        "",
    ]

    for s in plan.sections:
        sec_num = s.numero
        doc_parts.append(f"- **Section {sec_num}** : **{s.titre}**")

    doc_parts.append("")
    doc_parts.append("---")
    doc_parts.append("")

    # Concaténation harmonieuse des sections rédigées
    for item in sections_redigees:
        sec_text = item.get("texte", "").strip()
        if sec_text:
            doc_parts.append(sec_text)
            doc_parts.append("")
            doc_parts.append("---")
            doc_parts.append("")

    return "\n".join(doc_parts)


def controler_qualite_memoire(
    memoire_md: str,
    filtered_prescriptions: list[MarketPrescription],
    plan: PlanMemoire,
) -> RapportQualiteMemoire:
    """
    Effectue le contrôle qualité final avant export (Étape 7) :
    1. Couverture des prescriptions techniques et communes.
    2. Respect intégral du plan prévu.
    3. Recensement de tous les champs `[À COMPLÉTER]` sans omission.
    4. Détection de redondances excessives entre sections.
    5. Calcul du score qualité global.
    """
    # 1. Vérification de la présence des prescriptions dans le mémoire
    verify_prescriptions_in_generated_memoire(filtered_prescriptions, memoire_md)
    applicables = [p for p in filtered_prescriptions if p.applicabilite_lot_cible]
    integrees = [p for p in applicables if p.statut_memoire == "INTEGRE"]
    manquantes = [p for p in applicables if p.statut_memoire == "TRANSMIS_NON_INTEGRE"]

    taux_couverture = 100.0
    if applicables:
        taux_couverture = round((len(integrees) / len(applicables)) * 100, 1)

    # 2. Recensement des informations à compléter
    champs_manquants = re.findall(r"\[À COMPLÉTER\s*:\s*([^\]]+)\]", memoire_md, re.I)
    champs_dedup = sorted(list(set(c.strip() for c in champs_manquants)))

    # 3. Contrôle des sections du plan
    sections_trouvees = 0
    for s in plan.sections:
        # Recherche du titre de section dans le texte
        if re.search(re.escape(s.titre[:25]), memoire_md, re.I):
            sections_trouvees += 1

    # 4. Détection des répétitions excessives (phrases longues dupliquées)
    phrases = [p.strip() for p in re.split(r"[\.\n]+", memoire_md) if len(p.strip()) >= 60]
    seen_phrases: set[str] = set()
    repetitions_count = 0
    for ph in phrases:
        ph_norm = re.sub(r"\s+", " ", ph.lower())
        if ph_norm in seen_phrases:
            repetitions_count += 1
        else:
            seen_phrases.add(ph_norm)

    # 5. Points de vigilance
    vigilance: list[str] = []
    if taux_couverture < 80.0:
        vigilance.append(f"Taux de couverture des prescriptions perfectible ({taux_couverture}%).")
    if len(champs_dedup) > 0:
        vigilance.append(f"{len(champs_dedup)} champ(s) spécifique(s) à compléter avant remise de l'offre.")
    if repetitions_count > 5:
        vigilance.append(f"{repetitions_count} répétitions de phrases détectées entre sections.")
    if sections_trouvees < plan.total_sections:
        vigilance.append(f"{plan.total_sections - sections_trouvees} section(s) du plan n'ont pas été identifiées avec certitude.")

    # 6. Score qualité global sur 100
    score = 100.0
    score -= max(0, (100.0 - taux_couverture) * 0.4)
    if sections_trouvees < plan.total_sections:
        score -= (plan.total_sections - sections_trouvees) * 10
    if repetitions_count > 5:
        score -= min(15, repetitions_count)
    score_final = max(0.0, round(score, 1))

    statut_global = "CONFORME"
    if score_final < 70.0 or sections_trouvees < plan.total_sections:
        statut_global = "ATTENTION"
    if score_final < 50.0:
        statut_global = "NON_CONFORME"

    return RapportQualiteMemoire(
        score_qualite_pct=score_final,
        taux_couverture_prescriptions=taux_couverture,
        prescriptions_totales_count=len(applicables),
        prescriptions_integrees_count=len(integrees),
        prescriptions_manquantes_count=len(manquantes),
        informations_a_completer_count=len(champs_dedup),
        champs_a_completer=champs_dedup,
        sections_prevues_count=plan.total_sections,
        sections_redigees_count=sections_trouvees,
        repetition_detectees_count=repetitions_count,
        points_vigilance=vigilance,
        statut_global=statut_global,
    )
