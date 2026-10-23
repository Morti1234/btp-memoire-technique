#!/usr/bin/env python3
"""
Test 3 : Audit et vérification de l'intégrité des prescriptions communes et spécifiques
pour le générateur IA de mémoires techniques BTP.

Ce banc de test contrôle :
- Les 8 scénarios de régression imposés (scénarios 1 à 8).
- Les indicateurs A à G du protocole de recette.
- La détection de contradictions sans résolution arbitraire.
- L'intégrité de la traçabilité Document / Page / Article.
- La détection des pertes et des non-intégrations dans le mémoire final.
"""

import sys
import io
import json
import zipfile
from dataclasses import asdict

import email_automation as ea
import document_processor as dp
from document_processor import (
    ProcessedDocument,
    ExtractedPage,
    MarketPrescription,
    extract_market_prescriptions,
    filter_prescriptions_for_target_lot,
    verify_prescriptions_in_generated_memoire,
    audit_filter_quality,
    normalize_lot_code,
)


class InMemoryUpload:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self.data = data

    def getbuffer(self):
        return self.data

    def read(self):
        return self.data


def load_dce_simulation_docs() -> list[ProcessedDocument]:
    """Charge et traite unitairement les pièces du DCE de simulation réaliste."""
    _, zip_bytes = ea.build_realistic_dce_zip_simulation()
    z = zipfile.ZipFile(io.BytesIO(zip_bytes))
    docs = []
    for fname in sorted(z.namelist()):
        b = z.read(fname)
        up = InMemoryUpload(fname, b)
        pdoc = dp.process_single_document(up)
        pdoc = dp.analyze_single_document(pdoc)
        docs.append(pdoc)
    return docs


def run_scenario_1(dce_docs: list[ProcessedDocument]) -> dict:
    """Scénario 1 : CCTP Général Lot 00 -> prescriptions communes applicables conservées."""
    all_p = extract_market_prescriptions(dce_docs)
    target_lot = "LOT_02"
    filtered, _ = filter_prescriptions_for_target_lot(all_p, target_lot=target_lot)

    cctp_gen_p = [p for p in all_p if p.document_source.endswith("02_CCTP_Lot_00_Prescriptions_Generales.pdf")]
    cctp_gen_retained = [p for p in filtered if p.document_source.endswith("02_CCTP_Lot_00_Prescriptions_Generales.pdf")]

    success = (len(cctp_gen_p) > 0) and (len(cctp_gen_p) == len(cctp_gen_retained))
    return {
        "name": "Scénario 1 : Conservation des prescriptions communes du CCTP Général (Lot 00)",
        "success": success,
        "total_source": len(cctp_gen_p),
        "total_conservees": len(cctp_gen_retained),
        "taux_conservation": round((len(cctp_gen_retained) / max(1, len(cctp_gen_p))) * 100, 1),
        "exemples": [p.texte[:65] + "..." for p in cctp_gen_retained[:3]],
    }


def run_scenario_2(dce_docs: list[ProcessedDocument]) -> dict:
    """Scénario 2 : Le lot sélectionné possède des prescriptions techniques spécifiques conservées."""
    all_p = extract_market_prescriptions(dce_docs)
    target_lot = "LOT_02"
    filtered, _ = filter_prescriptions_for_target_lot(all_p, target_lot=target_lot)

    target_norm = normalize_lot_code(target_lot)
    lot_spec_source = [p for p in all_p if p.lot_concerne == target_norm]
    lot_spec_retained = [p for p in filtered if p.lot_concerne == target_norm and p.applicabilite_lot_cible]

    success = (len(lot_spec_source) > 0) and (len(lot_spec_source) == len(lot_spec_retained))
    return {
        "name": "Scénario 2 : Conservation des prescriptions spécifiques du lot ciblé (Lot 02 VRD)",
        "success": success,
        "total_source": len(lot_spec_source),
        "total_conservees": len(lot_spec_retained),
        "taux_conservation": round((len(lot_spec_retained) / max(1, len(lot_spec_source))) * 100, 1),
        "exemples": [p.texte[:65] + "..." for p in lot_spec_retained[:3]],
    }


def run_scenario_3(dce_docs: list[ProcessedDocument]) -> dict:
    """Scénario 3 : Un autre lot possède une prescription sans rapport -> exclue du lot sélectionné."""
    all_p = extract_market_prescriptions(dce_docs)
    target_lot = "LOT_02"
    filtered, _ = filter_prescriptions_for_target_lot(all_p, target_lot=target_lot)

    lot_01_source = [p for p in all_p if p.lot_concerne == "LOT_01"]
    # Vérifier qu'aucune prescription du Lot 01 n'est applicable ou conservée pour le Lot 02
    lot_01_improprement_retained = [p for p in filtered if p.lot_concerne == "LOT_01"]

    success = (len(lot_01_source) > 0) and (len(lot_01_improprement_retained) == 0)
    return {
        "name": "Scénario 3 : Exclusion des prescriptions techniques des autres lots sans rapport (Lot 01 Gros Œuvre)",
        "success": success,
        "total_prescriptions_lot01_source": len(lot_01_source),
        "prescriptions_lot01_retained_pour_lot02": len(lot_01_improprement_retained),
        "details_exclues": [p.texte[:65] + "..." for p in lot_01_source[:2]],
    }


def run_scenario_4() -> dict:
    """Scénario 4 : Clause d'interface entre deux lots concernant le lot retenu -> conservée."""
    # Création d'un document de test explicite avec clause d'interface Lot 01 / Lot 02
    test_doc = ProcessedDocument(
        file_id="DOC-TEST-01",
        filename="TEST_CCTP_General_Interfaces.pdf",
        format="PDF",
        size_bytes=2048,
        md5_hash="md5_test_interfaces_01",
        category="CCTP_GENERAL",
        category_label="CCTP Général (Prescriptions Communes)",
        page_count=1,
        total_char_count=350,
        pages=[
            ExtractedPage(
                page_num=1,
                char_count=350,
                text=(
                    "Article 3.4 - Interfaces entre lots :\n"
                    "Le Lot 02 VRD doit obligatoirement réaliser les réservations et fourreaux de raccordement "
                    "en traversée sous la dalle avant coulage du dallage par le Lot 01 Gros Œuvre."
                ),
                status="OK",
                articles_found=[{"article": "Article 3.4", "titre": "Interfaces entre lots"}],
            )
        ],
        lots_concernes=["TOUS LOTS"],
        status="VALIDE",
    )

    prescriptions = extract_market_prescriptions([test_doc])
    filtered_for_lot02, _ = filter_prescriptions_for_target_lot(prescriptions, target_lot="LOT_02")
    filtered_for_lot03, _ = filter_prescriptions_for_target_lot(prescriptions, target_lot="LOT_03")

    retained_for_lot02 = any(p.document_source == "TEST_CCTP_General_Interfaces.pdf" and p.applicabilite_lot_cible for p in filtered_for_lot02)
    excluded_for_lot03 = not any(p.document_source == "TEST_CCTP_General_Interfaces.pdf" and p.applicabilite_lot_cible for p in filtered_for_lot03)

    success = (len(prescriptions) > 0) and retained_for_lot02 and excluded_for_lot03
    return {
        "name": "Scénario 4 : Préservation des clauses d'interfaces inter-lots concernant le lot sélectionné",
        "success": success,
        "clause_extraite": prescriptions[0].texte if prescriptions else "",
        "lot_concerne": prescriptions[0].lot_concerne if prescriptions else "",
        "retenue_pour_lot_02": retained_for_lot02,
        "exclue_pour_lot_03_non_concerne": excluded_for_lot03,
    }


def run_scenario_5() -> dict:
    """Scénario 5 : Deux documents présentent des exigences contradictoires -> conflit signalé avec sources."""
    # Pièce A : RC prévoyant 14 semaines
    doc_rc = ProcessedDocument(
        file_id="DOC-TEST-RC",
        filename="TEST_RC_Reglement_Consultation.pdf",
        format="PDF",
        size_bytes=2048,
        md5_hash="md5_test_rc_01",
        category="RC",
        category_label="Règlement de Consultation",
        page_count=1,
        total_char_count=200,
        pages=[
            ExtractedPage(
                page_num=1,
                char_count=200,
                text="Article 4 : Le délai d'exécution global est de 14 semaines à compter de l'ordre de service.",
                status="OK",
                articles_found=[{"article": "Article 4", "titre": "Délai"}],
            )
        ],
        status="VALIDE",
    )
    # Pièce B : CCTP prévoyant 12 semaines
    doc_cctp = ProcessedDocument(
        file_id="DOC-TEST-CCTP",
        filename="TEST_CCTP_General.pdf",
        format="PDF",
        size_bytes=2048,
        md5_hash="md5_test_cctp_01",
        category="CCTP_GENERAL",
        category_label="CCTP Général",
        page_count=2,
        total_char_count=200,
        pages=[
            ExtractedPage(
                page_num=2,
                char_count=200,
                text="Article 2.1 : Le délai d'exécution global impératif est fixé à 12 semaines tous lots confondus.",
                status="OK",
                articles_found=[{"article": "Article 2.1", "titre": "Planning"}],
            )
        ],
        status="VALIDE",
    )

    prescriptions = extract_market_prescriptions([doc_rc, doc_cctp])
    filtered, contradictions = filter_prescriptions_for_target_lot(prescriptions, target_lot="LOT_02")

    has_contradiction = len(contradictions) > 0
    c_item = contradictions[0] if has_contradiction else {}
    sources = c_item.get("sources", [])
    cited_docs = [s.get("document") for s in sources]

    success = has_contradiction and ("TEST_RC_Reglement_Consultation.pdf" in cited_docs) and ("TEST_CCTP_General.pdf" in cited_docs)
    return {
        "name": "Scénario 5 : Détection et signalement des contradictions documentaires avec citation exacte des sources",
        "success": success,
        "contradictions_trouvees": len(contradictions),
        "message_conflit": c_item.get("message", ""),
        "sources_citees": sources,
    }


def run_scenario_6(dce_docs: list[ProcessedDocument]) -> dict:
    """Scénario 6 : Aucune prescription inventée sans source vérifiable."""
    all_p = extract_market_prescriptions(dce_docs)
    filtered, _ = filter_prescriptions_for_target_lot(all_p, target_lot="LOT_02")

    # Contrôle que 100% des prescriptions filtrées proviennent d'une source existante
    inventees = [p for p in filtered if not getattr(p, "source_verifiable", True) or not p.document_source]
    success = (len(inventees) == 0) and (len(filtered) > 0)
    return {
        "name": "Scénario 6 : Règle stricte anti-invention (zéro prescription inventée)",
        "success": success,
        "prescriptions_filtrees_total": len(filtered),
        "prescriptions_inventees_count": len(inventees),
    }


def run_scenario_7(dce_docs: list[ProcessedDocument]) -> dict:
    """Scénario 7 : Une prescription a été extraite, mais disparaît lors du filtre -> Le test DOIT échouer."""
    all_p = extract_market_prescriptions(dce_docs)
    filtered, _ = filter_prescriptions_for_target_lot(all_p, target_lot="LOT_02")

    # Simulation d'une perte anormale : suppression d'une prescription commune du filtre
    if filtered:
        perte_simulee = filtered.pop(0)
    else:
        perte_simulee = None

    audit_res = audit_filter_quality(all_p, filtered, memoire_text="", target_lot="LOT_02")

    # Le critère est : l'audit DOIT détecter la perte et déclarer l'échec du test 3
    detection_perte_effective = (audit_res["prescriptions_perdues_count"] > 0)
    test_echoue_comme_attendu = (audit_res["test_3_reussi"] is False)

    success = detection_perte_effective and test_echoue_comme_attendu
    return {
        "name": "Scénario 7 : Détection rigoureuse de disparition d'une prescription (échec immédiat du test)",
        "success": success,
        "prescription_perdue_simulee": perte_simulee.texte[:60] + "..." if perte_simulee else "",
        "prescriptions_perdues_detectees": audit_res["prescriptions_perdues_count"],
        "test_3_reussi_apres_perte": audit_res["test_3_reussi"],
    }


def run_scenario_8(dce_docs: list[ProcessedDocument]) -> dict:
    """Scénario 8 : Prescription conservée en interne mais absente du mémoire final -> Détection de la perte."""
    all_p = extract_market_prescriptions(dce_docs)
    filtered, _ = filter_prescriptions_for_target_lot(all_p, target_lot="LOT_02")

    # Mémoire technique partiel omettant délibérément la clôture périphérique de 2,00m et la gestion des déchets SOSED
    memoire_incomplet = """
    # MÉMOIRE TECHNIQUE - LOT 02 VRD
    ## 1. Méthodologie d'exécution
    L'entreprise Durand TP réalisera le terrassement en déblais sur 340 m3 et la mise en oeuvre de GNT 0/31.5 NF.
    Les bordures béton T2 NF P 98-340 seront posées sur 160 ml conformément aux règles de l'art.
    """

    audit_res = audit_filter_quality(all_p, filtered, memoire_text=memoire_incomplet, target_lot="LOT_02")

    non_integrees = audit_res["prescriptions_transmises_non_integrees_count"]
    success = non_integrees > 0
    return {
        "name": "Scénario 8 : Détection d'une prescription transmise mais absente du mémoire final",
        "success": success,
        "prescriptions_transmises_non_integrees": non_integrees,
        "prescriptions_integrees": audit_res["prescriptions_integrees_count"],
        "prescriptions_a_verifier": audit_res["prescriptions_a_verifier_count"],
    }


def run_full_nominal_audit(dce_docs: list[ProcessedDocument]) -> dict:
    """Audit nominal complet sur le dossier DCE avec mémoire technique exhaustif."""
    all_p = extract_market_prescriptions(dce_docs)
    target_lot = "LOT_02"
    filtered, contradictions = filter_prescriptions_for_target_lot(all_p, target_lot=target_lot)

    # Mémoire technique complet intégrant toutes les prescriptions applicables
    memoire_complet = """
    # MÉMOIRE TECHNIQUE COMPLET - LOT 02 VRD & AMÉNAGEMENTS EXTÉRIEURS
    ## 1. Contraintes du site occupé, accès et installation
    Conformément au CCTP Général Lot 00 (p. 1) et au RC, l'accès au chantier s'effectuera en respectant
    l'interdiction absolue de livraison aux heures de rentrée et sortie d'école (8h00-8h45 et 16h15-17h00).
    Les horaires autorisés de chantier (7h30 à 17h30) seront strictement respectés.
    Une clôture périphérique opaque de hauteur 2,00 m sera mise en place obligatoirement autour de l'emprise.
    Un homme-trafic muni de CACES assurera le balisage et la sécurité pour tout mouvement d'engins.

    ## 2. Déchets et respect de l'environnement (SOSED)
    La gestion des déchets de chantier respectera le schéma SOSED avec tri sélectif 5 flux,
    bennes dédiées et traçabilité réglementaire obligatoire via le bordereau Trackdéchets.

    ## 3. Méthodologie technique VRD et matériaux certifiés
    - Terrassement et décapage de l'enrobé existant sur 850 m2 avec évacuation en filière agréée.
    - Terrassement en déblais pour fond de forme sur 340 m3 selon DTU 12.
    - Couche de fondation en Grave Non Traitée GNT 0/31.5 certifiée NF (épaisseur 25 cm) avec essais à la plaque EV2 >= 80 MPa.
    - Pose de bordures béton préfabriquées type T2 NF P 98-340 sur lit béton (160 ml).
    - Revêtement perméable en béton désactivé et enrobé drainant sur 850 m2 selon norme NF P 98-150.
    - Réseau EP en canalisations PVC CR8 diamètre 200 mm et caniveaux à fente selon Fascicule 70 (95 ml).

    ## 4. Sécurité PGC et réseaux sensibles
    Conformément au PGC SPS, les réseaux enterrés sensibles Enedis et Gaz feront l'objet d'un marquage-piquetage,
    d'une DT-DICT préalable, et les travaux seront réalisés par nos personnels titulaires de l'AIPR Opérateur et Encadrant.
    Les poussières seront abattues par arrosage à la lance d'eau lors du rabotage.
    Le mémoire respecte la limite de pages imposée de 25 pages maximum selon le RC.
    """

    audit_res = audit_filter_quality(all_p, filtered, memoire_text=memoire_complet, target_lot=target_lot)
    return audit_res, len(all_p), len(filtered), contradictions


def main():
    print("=" * 80)
    print("BANC DE VALIDATION & AUDIT TEST 3 : INTÉGRITÉ DES PRESCRIPTIONS COMMUNES")
    print("=" * 80)

    print("\n[1/3] Chargement et traitement des pièces réelles du DCE de test...")
    dce_docs = load_dce_simulation_docs()
    print(f" -> {len(dce_docs)} documents DCE analysés unitairement sans troncature.")

    print("\n[2/3] Exécution des 8 scénarios de régression...")
    scenarios = [
        run_scenario_1(dce_docs),
        run_scenario_2(dce_docs),
        run_scenario_3(dce_docs),
        run_scenario_4(),
        run_scenario_5(),
        run_scenario_6(dce_docs),
        run_scenario_7(dce_docs),
        run_scenario_8(dce_docs),
    ]

    all_passed = True
    for i, sc in enumerate(scenarios, 1):
        status_sym = "✅ REUSSI" if sc["success"] else "❌ ECHEC"
        if not sc["success"]:
            all_passed = False
        print(f"\n▶ Scénario {i} : {sc['name']}")
        print(f"   Statut : {status_sym}")
        for k, v in sc.items():
            if k not in ("name", "success"):
                print(f"   - {k} : {v}")

    print("\n[3/3] Contrôle nominal exhaustif des indicateurs A à G (audit_filter_quality)...")
    audit_res, total_source, total_filtrees, contradictions = run_full_nominal_audit(dce_docs)

    print(f"\n--- RÉSULTATS DÉTAILLÉS DE L'AUDIT QUALITÉ (TEST 3) ---")
    print(f"A. Taux conservation prescriptions communes : {audit_res['taux_conservation_communes']}% (attendu: 100%)")
    print(f"B. Taux conservation prescriptions spécifiques : {audit_res['taux_conservation_specifiques']}% (attendu: 100%)")
    print(f"C. Prescriptions perdues lors du filtrage : {audit_res['prescriptions_perdues_count']} (attendu: 0)")
    print(f"D. Prescriptions inventées sans source vérifiable : {audit_res['prescriptions_inventees_count']} (attendu: 0)")
    print(f"E. Traçabilité complète Document/Page/Article : {audit_res['tracabilite_complete_pct']}% (attendu: 100%)")
    print(f"F. Prescriptions transmises au générateur : {audit_res['prescriptions_transmises_count']} sur {total_filtrees}")
    print(f"G. Répartition dans le mémoire généré :")
    print(f"   - Intégrées avec méthode : {audit_res['prescriptions_integrees_count']}")
    print(f"   - Transmises non intégrées : {audit_res['prescriptions_transmises_non_integrees_count']}")
    print(f"   - À vérifier (mention sans verbe d'action) : {audit_res['prescriptions_a_verifier_count']}")
    print(f"   - Autres lots exclus : {audit_res['prescriptions_autres_lots_exclues_count']}")
    print(f"   - Interfaces conservées : {audit_res['interfaces_conservees_count']}")
    print(f"   - Contradictions documentaires détectées : {audit_res['contradictions_count']}")
    print(f"TEST 3 STATUT GLOBAL : {'✅ CONFORME & RÉUSSI' if audit_res['test_3_reussi'] and all_passed else '❌ NON CONFORME'}")
    print("=" * 80)

    if not (audit_res["test_3_reussi"] and all_passed):
        sys.exit(1)


if __name__ == "__main__":
    main()
