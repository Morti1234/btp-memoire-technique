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
import hashlib
import json
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
