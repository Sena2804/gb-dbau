"""Génère le rapport Word des incohérences entre quotas et candidatures."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


def set_cell(cell, text, bold=False, color=None, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def add_table(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        set_cell(table.rows[0].cells[i], h, bold=True)
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            set_cell(table.rows[r_idx + 1].cells[c_idx], val)
    return table


def main():
    doc = Document()

    # -- Styles --
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(4)

    # -- En-tête --
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("RAPPORT D'ANALYSE")
    run.bold = True
    run.font.size = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Incohérences entre le fichier des quotas de bourses\net le tableau des candidatures CNaBAU")
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(80, 80, 80)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Session CNaBAU 2026 — Bourse de Coopération Russe")
    run.font.size = Pt(11)
    run.italic = True

    doc.add_paragraph()

    # -- Contexte --
    doc.add_heading("1. Contexte", level=1)
    doc.add_paragraph(
        "Ce rapport présente les incohérences identifiées entre deux documents de référence :"
    )
    doc.add_paragraph("Le fichier des quotas de bourses (quotas_bourses.xlsx) — 150 places réparties sur 57 filières", style="List Bullet")
    doc.add_paragraph("Le tableau des candidatures CNaBAU (Tableau_liste candidature_CNaBAU.xlsx) — 644 candidats répartis sur 53 filières", style="List Bullet")
    doc.add_paragraph(
        "Ces incohérences empêchent l'application correcte des quotas dans l'outil de gestion "
        "de la session. Des décisions sont nécessaires pour chaque point soulevé ci-dessous."
    )

    # -- Section 1 : Classification Doctorat / Spécialisation --
    doc.add_heading("2. Désaccord de classification : Doctorat vs Spécialisation", level=1)
    doc.add_paragraph(
        "Quatre filières sont classées différemment entre les deux fichiers. "
        "Le fichier des quotas regroupe Doctorat et Spécialisation sous une même section "
        "(\"DOCTORAT ET SPECIALISATION\"), tandis que le tableau des candidatures les sépare "
        "en deux niveaux distincts."
    )

    add_table(doc,
        ["Filière", "Classement dans le fichier Quotas", "Classement dans le fichier Candidatures"],
        [
            ["Sécurité Informatique", "Spécialisation (code 1.2.4)", "DOCTORAT"],
            ["Cardiologie Pédiatrique", "Doctorat (code 31.08.36)", "SPECIALITE"],
            ["Dentisterie pédiatrique", "Doctorat (code 31.08.76)", "SPECIALITE"],
            ["Oncologie", "Doctorat (code 31.08.57)", "SPECIALITE"],
        ]
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Question 1 : ")
    run.bold = True
    run.font.color.rgb = RGBColor(180, 40, 40)
    p.add_run("Quel classement doit faire foi ? Celui du fichier des quotas ou celui du tableau des candidatures ? "
              "Les quotas doivent-ils être appliqués selon le niveau indiqué dans quel document ?")

    # -- Section 2 : Différences de noms de filières --
    doc.add_heading("3. Différences dans les noms de filières", level=1)
    doc.add_paragraph(
        "Plusieurs filières portent des noms légèrement différents entre les deux fichiers "
        "(fautes de frappe, accents, pluriels). Cela peut empêcher la correspondance automatique des quotas."
    )

    doc.add_heading("3.1 Licence", level=2)
    add_table(doc,
        ["Nom dans le fichier Quotas", "Nom dans le fichier Candidatures", "Type d'écart"],
        [
            ["Aménagement foncier et cadastre", "Aménagement Foncier et Cadastre", "Casse uniquement"],
            ["Bioressources aquatiques et acquaculture", "Bioressources Aquatiques et Acquaculture", "Casse uniquement"],
            ["Cartographie et géoinformatique", "Cartographie et Géoinformatique", "Casse uniquement"],
            ["Chimie et pédologie agricoles", "Chimie et Pédologie Agricole", "Pluriel (agricoles → Agricole)"],
            ["Métiers d'art appliqué et métiers traditionnels", "Métiers d'art appliqué et métiers tradionnels", "Faute de frappe (tradionnels)"],
            ["Mécanique et modèles mathématiques", "Mécanique et modèle mathématiques", "Pluriel (modèles → modèle)"],
            ["Cultures physiques pour les personnes handicapées", "Cultures Physiques pour les personnes handicapées", "Casse uniquement"],
        ]
    )

    doc.add_heading("3.2 Master", level=2)
    add_table(doc,
        ["Nom dans le fichier Quotas", "Nom dans le fichier Candidatures", "Type d'écart"],
        [
            ["Machines et outillages technologiques", "Machine et Outillage Technologiques", "Pluriel + casse"],
            ["Électronique et nanoélectronique", "Electronique et Nanoélectrique", "Faute (Nanoélectrique)"],
            ["Énergétique et électrotechnique", "Energétique et Electrotecnique", "Faute (Electrotecnique)"],
            ["Bioressources aquatiques et aquaculture", "Bioressources aquatique et aquaculture", "Pluriel (aquatiques → aquatique)"],
            ["Exploitation technique de systèmes électriques d'aéronefs...", "Exploitation Technique de Systènes Electriques d'Aeronefs...", "Fautes (Systènes, Aeronefs)"],
        ]
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Question 2 : ")
    run.bold = True
    run.font.color.rgb = RGBColor(180, 40, 40)
    p.add_run("Les noms de filières du tableau des candidatures sont-ils les noms officiels à utiliser ? "
              "Ou faut-il se baser sur les noms du fichier des quotas ? "
              "Cela permettra d'uniformiser les appellations dans l'outil.")

    # -- Section 3 : Filières avec candidats mais sans quota --
    doc.add_heading("4. Filières avec candidats mais SANS quota attribué", level=1)
    doc.add_paragraph(
        "Les filières suivantes comptent des candidats dans le tableau CNaBAU, "
        "mais aucune place ne leur est attribuée dans le fichier des quotas. "
        "Les candidats de ces filières ne pourront pas recevoir d'avis favorable si un quota est requis."
    )

    doc.add_heading("4.1 Licence (5 filières)", level=2)
    add_table(doc,
        ["Filière", "Nombre de candidats"],
        [
            ["Agronomie", "Présents dans le tableau"],
            ["Entraînement Physique", "Présents dans le tableau"],
            ["Mécanique Aéronautique", "Présents dans le tableau"],
            ["Soins infirmiers", "Présents dans le tableau"],
            ["Utilisation des terres et registre cadastral", "Présents dans le tableau"],
        ]
    )

    doc.add_heading("4.2 Master (4 filières)", level=2)
    add_table(doc,
        ["Filière", "Observation"],
        [
            ["Agronomie", "Pas de quota Master pour cette filière"],
            ["Informatique fondamentale et technologie de l'information", "Pas de quota"],
            ["Ingénierie Électrique et Génie Électrique", "Pas de quota"],
            ["Génie Énergétique", "Le quota existe en Licence mais pas en Master"],
        ]
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Question 3 : ")
    run.bold = True
    run.font.color.rgb = RGBColor(180, 40, 40)
    p.add_run("Quel traitement appliquer aux candidats de ces filières sans quota ? "
              "Doivent-ils être automatiquement marqués \"Défavorable\" ? "
              "Ou un quota doit-il leur être attribué ? Si oui, combien de places par filière ?")

    # -- Section 4 : Filières avec quota mais sans candidats --
    doc.add_heading("5. Filières avec quota attribué mais SANS candidat", level=1)
    doc.add_paragraph(
        "Les filières suivantes ont des places attribuées dans le fichier des quotas, "
        "mais aucun candidat n'a postulé dans le tableau CNaBAU."
    )

    add_table(doc,
        ["Niveau", "Filière", "Quota (places)"],
        [
            ["Master", "Design (par industrie)", "3"],
            ["Master", "Restauration", "3"],
            ["Doctorat", "Expertise médico-légale", "1"],
            ["Doctorat", "Médecine transfusionnelle", "1"],
            ["Doctorat", "Radiothérapie", "2"],
            ["Doctorat", "Pharmacologie clinique", "1"],
            ["Doctorat", "Génétique de laboratoire", "1"],
            ["Doctorat", "Diagnostic et traitement radiographiques endovasculaires", "1"],
        ]
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Question 4 : ")
    run.bold = True
    run.font.color.rgb = RGBColor(180, 40, 40)
    p.add_run("Ces places non utilisées peuvent-elles être réaffectées à d'autres filières ? "
              "Ou doivent-elles rester réservées ?")

    # -- Section 5 : Doublons dans le fichier quotas --
    doc.add_heading("6. Doublons dans le fichier des quotas", level=1)
    doc.add_paragraph(
        "Certaines filières apparaissent deux fois dans le fichier des quotas au sein du même niveau, "
        "avec des quotas potentiellement différents. Dans l'outil, les quotas ont été additionnés."
    )

    add_table(doc,
        ["Niveau", "Filière", "1re occurrence", "2e occurrence", "Total appliqué"],
        [
            ["Licence", "Chimie et pédologie agricoles", "2", "2", "4"],
            ["Master", "Mécanique et modélisation mathématique", "4", "3", "7"],
            ["Master", "Génie agricole", "4", "4", "8"],
            ["Master", "Écologie et gestion de l'environnement", "4", "3", "7"],
        ]
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Question 5 : ")
    run.bold = True
    run.font.color.rgb = RGBColor(180, 40, 40)
    p.add_run("Ces doublons sont-ils intentionnels (les quotas doivent être additionnés) "
              "ou s'agit-il d'erreurs de saisie ? Quel est le quota correct pour chaque filière concernée ?")

    # -- Récapitulatif --
    doc.add_heading("7. Récapitulatif des questions", level=1)
    questions = [
        "Quel classement (Doctorat vs Spécialisation) fait foi pour les 4 filières en désaccord ?",
        "Quels noms de filières sont les noms officiels à utiliser dans l'outil ?",
        "Quel traitement pour les 9 filières ayant des candidats mais aucun quota ?",
        "Les 8 filières ayant un quota mais aucun candidat peuvent-elles être réaffectées ?",
        "Les doublons dans le fichier quotas sont-ils intentionnels ou des erreurs ?",
    ]
    for i, q in enumerate(questions, 1):
        doc.add_paragraph(f"Question {i} : {q}", style="List Number")

    # -- Sauvegarde --
    output = "data/rapport_incoherences_quotas_candidatures.docx"
    doc.save(output)
    print(f"Rapport généré : {output}")


if __name__ == "__main__":
    main()
