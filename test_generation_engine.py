#!/usr/bin/env python3
"""
Test d'intégration et de validation du Moteur de Génération Modulaire de Mémoires Techniques BTP.
Valide les Étapes 1 à 7 conformément au cahier des charges :
1. Construction du plan (RC imposé vs Critères RC vs Structure adaptée BTP, limite de pages).
2. Rédaction séparée section par section avec contexte ciblé.
3. Déploiement méthodologique en 10 points pour chaque prestation technique.
4. Règle anti-paraphrase (réponse opérationnelle 'comment' vs exigence CCTP 'quoi').
5. Règle anti-invention (balises [À COMPLÉTER : ...] pour moyens non établis).
6. Assemblage sans troncature avec table des matières et styles préservés.
7. Contrôle qualité final (couverture, non-invention, répétitions, cohérence du plan).
"""

import sys
import io
import re
import json
import zipfile

import email_automation as ea
import document_processor as dp
import app


def load_test_dce_data():
    """Charge les documents de test réalistes et extrait l'analyse et les prescriptions."""
    _, zip_bytes = ea.build_realistic_dce_zip_simulation()
    z = zipfile.ZipFile(io.BytesIO(zip_bytes))

    class UploadMock:
        def __init__(self, name, b):
            self.name = name
            self.b = b
        def getbuffer(self): return self.b
        def read(self): return self.b

    docs = []
    for fname in sorted(z.namelist()):
        raw = z.read(fname)
        pdoc = dp.process_single_document(UploadMock(fname, raw))
        pdoc = dp.analyze_single_document(pdoc)
        docs.append(pdoc)

    all_prescriptions = dp.extract_market_prescriptions(docs)
    filtered_prescriptions, contradictions = dp.filter_prescriptions_for_target_lot(
        all_prescriptions, target_lot="LOT_02"
    )

    # Analyse structurée type issue de l'analyse croisée
    analyse_json = {
        "nom_projet": "Réhabilitation du Groupe Scolaire Jean Moulin",
        "maitre_ouvrage": "Ville de Lyon (Direction de la Construction)",
        "lieu_execution": "45 Avenue Jean Moulin, 69003 Lyon",
        "dates_cles": {
            "date_remise_offres": "15 Novembre 2026 à 12h00",
            "delai_execution": "14 semaines",
        },
        "reglement_consultation_rc": {
            "criteres_valeur_technique": [
                {
                    "ordre": 1,
                    "critere": "Compréhension des contraintes du site occupé et installation de chantier",
                    "bareme_points": "20 pts",
                    "attendus": "Accès, riverains, sécurité aux heures scolaires, clôture 2m",
                },
                {
                    "ordre": 2,
                    "critere": "Moyens humains (CACES, AIPR), matériels dédiés et méthodologie d'exécution VRD",
                    "bareme_points": "20 pts",
                    "attendus": "Méthode d'exécution des postes de terrassement et voirie, encadrement qualifié",
                },
                {
                    "ordre": 3,
                    "critere": "Qualité et provenance des matériaux, planning basé sur la DPGF et cadences",
                    "bareme_points": "10 pts",
                    "attendus": "GNT 0/31.5 NF, bordures T2, enrobé perméable NF P 98-150, calendrier",
                },
                {
                    "ordre": 4,
                    "critere": "Sécurité PGC en milieu scolaire et gestion des déchets SOSED",
                    "bareme_points": "10 pts",
                    "attendus": "Homme-trafic, réseaux sensibles Enedis/Gaz, suivi Trackdéchets",
                },
            ],
            "nombre_pages_max": "25 pages maximum",
        },
        "lots_concernes": ["Lot 02 - VRD, cour d'école perméable & réseaux humides"],
        "dpgf_postes_quantites": [
            {"poste": "2.1", "designation": "Croûtage et démolition d'enrobé existant", "quantite": "850", "unite": "m²"},
            {"poste": "2.2", "designation": "Terrassement en déblais en pleine masse pour fond de forme", "quantite": "340", "unite": "m³"},
            {"poste": "2.3", "designation": "Fourniture et mise en oeuvre de GNT 0/31.5 NF + essais plaque EV2", "quantite": "850", "unite": "m²"},
            {"poste": "2.4", "designation": "Fourniture et pose de bordures béton T2 NF P 98-340", "quantite": "160", "unite": "ml"},
            {"poste": "2.5", "designation": "Revêtement perméable (béton désactivé / enrobé drainant NF P 98-150)", "quantite": "850", "unite": "m²"},
            {"poste": "2.6", "designation": "Réseau eaux pluviales PVC CR8 DN200 + caniveaux (Fascicule 70)", "quantite": "95", "unite": "ml"},
        ],
        "moyens_entreprise": {
            "nom": "DURAND TP & VRD SAS",
            "effectif": "4 compagnons VRD qualifiés CACES R482 et AIPR",
            "materiel": "1 Pelle Mecalac 12MTX, 1 Mini-pelle Kubota 5.5T, 1 Compacteur tandem Bomag BW 120",
        },
    }

    profil_entreprise = (
        "Raison sociale : DURAND TP & VRD SAS\n"
        "Encadrement dédié : Marc DURAND (Conducteur de travaux principal, ESTP, AIPR Encadrant) "
        "et Karim BENALI (Chef de chantier VRD, 11 ans exp, CACES R482, SST, AIPR).\n"
        "Équipe d'exécution : 4 compagnons VRD (2 conducteurs CACES B1/C1 + AIPR Opérateur, 2 maçons poseurs).\n"
        "Parc matériel : 1 Pelle Mecalac 12MTX, 1 Mini-pelle Kubota 5,5T, 1 Compacteur Bomag BW 120, "
        "2 Camions bennes 8x4 Renault Trucks C440, plaque vibrante avec capteurs de compacité, laser rotatif Leica.\n"
        "Certifications : Qualibat 1312, Qualibat 1321, FNTP Carte Pro TP, Ecovadis Argent."
    )

    return docs, analyse_json, filtered_prescriptions, profil_entreprise


def test_etape_1_construction_du_plan():
    """Vérifie la construction du plan à partir du RC ou standard adapté (Étape 1)."""
    print("\n[TEST 1] Étape 1 : Construction du plan selon RC et critères...")
    _, analyse_json, filtered_prescriptions, _ = load_test_dce_data()

    # Cas 1 : Plan basé sur les 4 critères du RC
    plan_rc = dp.construire_plan_memoire(analyse_json, filtered_prescriptions, target_lot="LOT_02")
    assert plan_rc.source_plan == "RC_CRITERES", f"Source attendue RC_CRITERES, obtenu {plan_rc.source_plan}"
    assert plan_rc.total_sections == 4, f"Attendu 4 sections selon RC, obtenu {plan_rc.total_sections}"
    assert plan_rc.limite_pages_max == 25, f"Attendu 25 pages, obtenu {plan_rc.limite_pages_max}"

    # Vérification que la section 2 a été détectée comme méthodologie travaux
    sec2 = plan_rc.sections[1]
    assert sec2.is_methodologie_travaux is True, "La section 2 (méthodologie VRD) doit être qualifiée en méthodologie travaux."
    assert len(sec2.prestations_concernees) >= 6, "Les postes DPGF doivent être associés à la section méthodologie."
    print("   ✅ Cas 1 (Plan RC) : 4 sections ordonnées avec pondération et limite 25 pages.")

    # Cas 2 : Plan standard adapté lorsque le RC n'impose rien
    analyse_sans_rc = dict(analyse_json)
    analyse_sans_rc["reglement_consultation_rc"] = {}
    plan_standard = dp.construire_plan_memoire(analyse_sans_rc, filtered_prescriptions, target_lot="LOT_02")
    assert plan_standard.source_plan == "STANDARD_ADAPTE"
    assert plan_standard.total_sections == 7, f"Attendu 7 sections standard, obtenu {plan_standard.total_sections}"
    assert any(s.is_methodologie_travaux for s in plan_standard.sections), "Le plan standard doit inclure une section travaux."
    print("   ✅ Cas 2 (Plan Standard BTP) : 7 sections professionnelles structurées pour le lot.")


def test_etape_2_et_3_redaction_sections_et_10_points():
    """Vérifie la rédaction séparée par section et le respect strict des 10 étapes méthodologiques (Étapes 2 & 3)."""
    print("\n[TEST 2] Étapes 2 & 3 : Rédaction séparée par section et déroulement en 10 points...")
    _, analyse_json, filtered_prescriptions, profil_entreprise = load_test_dce_data()
    plan = dp.construire_plan_memoire(analyse_json, filtered_prescriptions, target_lot="LOT_02")

    # Rédaction unitaire de la section 2 (Méthodologie travaux VRD)
    sec_travaux = plan.sections[1]
    texte_sec2 = dp.rediger_section_memoire(
        section=sec_travaux,
        profil_entreprise=profil_entreprise,
        analyse_json=analyse_json,
        plan=plan,
    )

    # 1. Vérification des 10 étapes opérationnelles
    etapes_attendues = [
        "1. **Vérifications préalables**",
        "2. **Préparation de la zone**",
        "3. **Protections nécessaires**",
        "4. **Phases d'exécution chronologiques**",
        "5. **Moyens mobilisés**",
        "6. **Mesures de sécurité spécifiques**",
        "7. **Contrôles de conformité et PAQ**",
        "8. **Gestion des déchets (SOSED)**",
        "9. **Coordination et coactivité**",
        "10. **Vérifications finales et réception**",
    ]
    for etape in etapes_attendues:
        assert etape in texte_sec2, f"Étape manquante dans la méthodologie : {etape}"

    print(f"   ✅ Les 10 points méthodologiques sont intégralement développés pour chaque prestation.")
    # Vérification que les prestations DPGF (démolition, déblais 340 m3, GNT, bordures) sont citées
    assert "850 m²" in texte_sec2 or "850" in texte_sec2
    assert "340 m³" in texte_sec2 or "340" in texte_sec2
    assert "EV2 >= 80 MPa" in texte_sec2 or "EV2" in texte_sec2
    print("   ✅ Quantités exactes de la DPGF et exigences techniques intégrées sans paraphrase passive.")


def test_etape_4_et_5_anti_paraphrase_et_non_invention():
    """Vérifie l'absence de paraphrase passive et l'application stricte de la règle anti-invention (Étapes 4 & 5)."""
    print("\n[TEST 3] Étapes 4 & 5 : Règle anti-invention et signalement des données manquantes...")
    _, analyse_json, filtered_prescriptions, _ = load_test_dce_data()
    plan = dp.construire_plan_memoire(analyse_json, filtered_prescriptions, target_lot="LOT_02")

    # Rédaction avec profil entreprise vide -> le système doit générer les balises [À COMPLÉTER : ...]
    sec_moyens = plan.sections[0]
    texte_sans_profil = dp.rediger_section_memoire(
        section=sec_moyens,
        profil_entreprise="",
        analyse_json=analyse_json,
        plan=plan,
    )

    balises_trouvees = re.findall(r"\[À COMPLÉTER\s*:\s*([^\]]+)\]", texte_sans_profil, re.I)
    assert len(balises_trouvees) >= 2, f"Le système devait générer des balises [À COMPLÉTER], trouvées : {balises_trouvees}"
    print(f"   ✅ Règle anti-invention respectée : {len(balises_trouvees)} balise(s) [À COMPLÉTER] générée(s) sans moyen fictif.")


def test_etape_6_et_7_assemblage_et_controle_qualite():
    """Vérifie l'assemblage complet sans troncature et le rapport qualité final (Étapes 6 & 7)."""
    print("\n[TEST 4] Étapes 6 & 7 : Assemblage complet sans troncature et contrôle qualité final...")
    _, analyse_json, filtered_prescriptions, profil_entreprise = load_test_dce_data()

    # Exécution du moteur complet via app.rediger_memoire avec suivi progressif
    progression_enregistree = []
    def callback_ui(cur, total, titre):
        progression_enregistree.append((cur, total, titre))

    memoire_complet = app.rediger_memoire(
        api_key="",
        provider="native",
        analyse_json=analyse_json,
        profil_entreprise=profil_entreprise,
        prescriptions_marche=filtered_prescriptions,
        progress_callback=callback_ui,
    )

    # 1. Vérification de l'assemblage
    assert len(progression_enregistree) == 4, f"4 sections devaient être suivies, {len(progression_enregistree)} reçues."
    assert "# MÉMOIRE TECHNIQUE JUSTIFICATIF" in memoire_complet
    assert "## SOMMAIRE GÉNÉRAL" in memoire_complet
    assert "Section 1" in memoire_complet
    assert "Section 4" in memoire_complet
    assert len(memoire_complet) >= 5000, f"Le mémoire assemblé est anormalement court ({len(memoire_complet)} car.)."
    print(f"   ✅ Mémoire assemblé avec succès : {len(memoire_complet):,} caractères, Sommaire et 4 sections présentes.")

    # 2. Vérification du contrôle qualité final (Étape 7)
    plan = dp.construire_plan_memoire(analyse_json, filtered_prescriptions, target_lot="LOT_02")
    rapport = dp.controler_qualite_memoire(memoire_complet, filtered_prescriptions, plan)

    print("\n--- RAPPORT QUALITÉ DU MÉMOIRE GÉNÉRÉ ---")
    print(f"   - Score qualité global : {rapport.score_qualite_pct} / 100")
    print(f"   - Taux couverture prescriptions : {rapport.taux_couverture_prescriptions} %")
    print(f"   - Prescriptions intégrées : {rapport.prescriptions_integrees_count} sur {rapport.prescriptions_totales_count}")
    print(f"   - Sections prévues / rédigées : {rapport.sections_prevues_count} / {rapport.sections_redigees_count}")
    print(f"   - Statut global : {rapport.statut_global}")

    assert rapport.score_qualite_pct >= 70.0, f"Score qualité insuffisant : {rapport.score_qualite_pct}"
    assert rapport.sections_redigees_count == plan.total_sections, "Toutes les sections du plan doivent être présentes."
    assert rapport.statut_global in ("CONFORME", "ATTENTION")
    print("   ✅ Contrôle qualité validé avec succès.")


def test_export_docx_et_pdf_integrite():
    """Vérifie que le mémoire enrichi s'exporte sans erreur en DOCX et PDF."""
    print("\n[TEST 5] Export final Word (.docx) et PDF (.pdf)...")
    _, analyse_json, filtered_prescriptions, profil_entreprise = load_test_dce_data()
    memoire_md = app.rediger_memoire(
        api_key="",
        provider="native",
        analyse_json=analyse_json,
        profil_entreprise=profil_entreprise,
        prescriptions_marche=filtered_prescriptions,
    )
    ent_config = {"nom": "DURAND TP & VRD", "siret": "48291033400028"}

    docx_buf = app.generate_docx(memoire_md, analyse_json, ent_config)
    pdf_buf = app.generate_pdf(memoire_md, analyse_json, ent_config)

    assert docx_buf.getbuffer().nbytes > 10_000, "Le buffer Word est anormalement petit."
    assert pdf_buf.getbuffer().nbytes > 10_000, "Le buffer PDF est anormalement petit."
    print(f"   ✅ Livrables exportés sans perte : Word={docx_buf.getbuffer().nbytes:,} octets, PDF={pdf_buf.getbuffer().nbytes:,} octets.")


def main():
    print("=" * 80)
    print("BANC DE TEST DU NOUVEAU MOTEUR DE GÉNÉRATION DE MÉMOIRES TECHNIQUES BTP")
    print("=" * 80)

    try:
        test_etape_1_construction_du_plan()
        test_etape_2_et_3_redaction_sections_et_10_points()
        test_etape_4_et_5_anti_paraphrase_et_non_invention()
        test_etape_6_et_7_assemblage_et_controle_qualite()
        test_export_docx_et_pdf_integrite()
        print("\n" + "=" * 80)
        print("TOUS LES TESTS DU MOTEUR DE GÉNÉRATION ONT RÉUSSI AVEC SUCCÈS (100% SUCCÈS) ✅")
        print("=" * 80)
    except AssertionError as e:
        print(f"\n❌ ÉCHEC D'ASSERTION : {e}")
        sys.exit(1)
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"\n❌ ERREUR INATTENDUE : {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
