"""Module de gestion SQLite pour les candidatures CNBAU."""

import io
import json
import re
import sqlite3
from pathlib import Path

import pandas as pd

DB_PATH = "cnbau_session.db"

NIVEAU_MAP = {
    "LICENCE": "Licence",
    "MASTER": "Master",
    "DOCTORAT": "Doctorat",
    "SPECIALITE": "Spécialisation",
    "SPÉCIALISATION": "Spécialisation",
    "SPECIALISATION": "Spécialisation",
    "SPECIALITE MEDICALE": "Spécialisation",
    "SPÉCIALITÉ MÉDICALE": "Spécialisation",
}

# Politique de gestion des doublons par catégorie.
# "keep" = conserver les deux entrées, "drop" = ne garder que la première.
DUPLICATE_POLICY = {
    "identical": "keep",                # Même personne, même filière/niveau
    "same_person_diff_filiere": "keep", # Même personne, filières/niveaux différents
    "different_people": "keep",         # Personnes différentes, même ID russe
}


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS candidatures (
            id_demande TEXT PRIMARY KEY,
            id_russe TEXT,
            numero INTEGER,
            sexe TEXT,
            name TEXT NOT NULL, 
            date_lieu_naissance TEXT,
            diplome_filiere_annee TEXT,
            moyenne TEXT,
            observation TEXT,
            filiere TEXT NOT NULL,
            niveau_etudes TEXT NOT NULL,
            avis TEXT DEFAULT 'En attente'
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS quotas (
            niveau_etudes TEXT NOT NULL,
            filiere TEXT NOT NULL,
            nb_places INTEGER NOT NULL,
            PRIMARY KEY (niveau_etudes, filiere)
        )
    """)
    conn.commit()
    conn.close()


def _normalize_niveau(raw: str) -> str:
    """Normalise 'LICENCE' → 'Licence', 'SPECIALITE' → 'Spécialisation', etc."""
    key = raw.strip().upper()
    return NIVEAU_MAP.get(key, raw.strip().title())


def _normalize_filiere(name: str) -> str:
    """Normalise les espaces multiples dans un nom de filière."""
    return re.sub(r'\s+', ' ', name).strip()


def _build_filiere_lookup() -> dict:
    """Construit un index {clé_normalisée: nom_json} depuis quotas.json pour matcher
    les noms de filières indépendamment de la casse et des espaces."""
    quotas_path = Path(__file__).resolve().parent / "quotas.json"
    if not quotas_path.exists():
        return {}
    with open(quotas_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    lookup = {}
    for filieres in data.values():
        for fil in filieres:
            key = _normalize_filiere(fil).lower()
            lookup[key] = fil
    return lookup


def _parse_real_excel(excel_path: str) -> list[dict]:
    """Parse le fichier CNaBAU structuré avec lignes de séparation niveau/filière."""
    from openpyxl import load_workbook

    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

    filiere_lookup = _build_filiere_lookup()
    current_niveau = ""
    current_filiere = ""
    candidates = []

    for row in ws.iter_rows(min_row=2, values_only=False):
        cell_a = row[0].value
        if cell_a is None:
            continue

        val_a = str(cell_a).strip()

        # Ligne de séparation NIVEAU
        if "NIVEAU" in val_a.upper() and ":" in val_a:
            raw_niveau = val_a.split(":", 1)[-1].strip()
            current_niveau = _normalize_niveau(raw_niveau)
            continue

        # Ligne de séparation Filière (avec préfixe "Filiere :" ou "Filiere:")
        val_lower = val_a.lower()
        if val_lower.startswith("filiere") or val_lower.startswith("filière"):
            if ":" in val_a:
                raw_fil = val_a.split(":", 1)[-1].strip()
            else:
                raw_fil = val_a
            # Nettoyer : enlever "(N bourses)" puis "- code", normaliser les espaces
            raw_fil = re.sub(r'\(\s*\d+\s*bourses?\)', '', raw_fil).strip()
            raw_fil = re.sub(r'\s*-\s*[\d.]+\s*$', '', raw_fil).strip()
            raw_fil = _normalize_filiere(raw_fil)
            # Matcher avec le nom officiel du JSON (casse de référence)
            current_filiere = filiere_lookup.get(raw_fil.lower(), raw_fil)
            continue

        # Ligne de données : la colonne A doit être un numéro
        try:
            num = int(cell_a)
        except (ValueError, TypeError):
            # Peut être un nom de filière orphelin (sans préfixe "Filiere :")
            rest_empty = all(c.value is None for c in row[1:9])
            if rest_empty and len(val_a) > 3:
                current_filiere = val_a.strip()
            continue

        # C'est un candidat
        def cell_str(idx):
            v = row[idx].value if idx < len(row) else None
            return str(v).strip() if v is not None else ""

        id_russe = cell_str(2)
        candidates.append({
            "id_demande": f"{num:04d}",
            "id_russe": id_russe,
            "numero": num,
            "sexe": cell_str(1),
            "name": cell_str(3),
            "date_lieu_naissance": cell_str(4),
            "diplome_filiere_annee": cell_str(5),
            "moyenne": cell_str(6),
            "observation": cell_str(7),
            "filiere": current_filiere,
            "niveau_etudes": current_niveau,
            "avis": cell_str(8) or "En attente",
        })

    wb.close()
    return candidates


def _apply_duplicate_policy(candidates: list[dict]) -> list[dict]:
    """Applique la politique de déduplication selon DUPLICATE_POLICY."""
    from collections import defaultdict

    by_id_russe = defaultdict(list)
    for c in candidates:
        id_russe = c.get("id_russe", "")
        if id_russe:
            by_id_russe[id_russe].append(c)

    to_drop = set()  # id_demande à supprimer

    for id_russe, group in by_id_russe.items():
        if len(group) < 2:
            continue
        a, b = group[0], group[1]

        # Classifier la paire
        name_a = set(a["name"].upper().split())
        name_b = set(b["name"].upper().split())
        common = name_a & name_b
        similarity = len(common) / max(len(name_a), len(name_b)) if name_a or name_b else 0

        if similarity < 0.3:
            category = "different_people"
        elif a["filiere"] == b["filiere"] and a["niveau_etudes"] == b["niveau_etudes"]:
            category = "identical"
        else:
            category = "same_person_diff_filiere"

        if DUPLICATE_POLICY.get(category) == "drop":
            # On garde la première entrée (plus petit numéro), on supprime la seconde
            to_drop.add(b["id_demande"])

    if to_drop:
        candidates = [c for c in candidates if c["id_demande"] not in to_drop]

    return candidates


def _is_real_cnabau_file(excel_path: str) -> bool:
    """Détecte si c'est le vrai fichier CNaBAU (structuré) ou un fichier plat."""
    from openpyxl import load_workbook

    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    # Vérifie si les premières lignes contiennent du texte institutionnel
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=1, values_only=True):
        val = str(row[0] or "")
        if "NIVEAU" in val.upper() or "CNaBAU" in val or "BOURSE" in val.upper():
            wb.close()
            return True
    wb.close()
    return False


def load_excel_to_db(excel_path: str) -> int:
    """Charge un fichier Excel dans la base SQLite.
    Détecte automatiquement si c'est le vrai fichier CNaBAU ou un fichier plat (mock).
    """
    if _is_real_cnabau_file(excel_path):
        return _load_real_excel(excel_path)
    return _load_flat_excel(excel_path)


def _load_real_excel(excel_path: str) -> int:
    """Charge le vrai fichier CNaBAU structuré avec gestion des doublons."""
    candidates = _parse_real_excel(excel_path)
    candidates = _apply_duplicate_policy(candidates)
    conn = get_connection()
    for c in candidates:
        conn.execute(
            """INSERT OR REPLACE INTO candidatures
               (id_demande, id_russe, numero, sexe, name, date_lieu_naissance,
                diplome_filiere_annee, moyenne, observation, filiere, niveau_etudes, avis)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (c["id_demande"], c["id_russe"], c["numero"], c["sexe"], c["name"],
             c["date_lieu_naissance"], c["diplome_filiere_annee"], c["moyenne"],
             c["observation"], c["filiere"], c["niveau_etudes"], c["avis"]),
        )
    conn.commit()
    conn.close()
    return len(candidates)


def _load_flat_excel(excel_path: str) -> int:
    """Charge un fichier Excel plat (mock / format simple)."""
    df = pd.read_excel(excel_path, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    if "avis" in df.columns:
        df["avis"] = df["avis"].fillna("En attente").replace("", "En attente")
    else:
        df["avis"] = "En attente"

    conn = get_connection()
    for _, row in df.iterrows():
        conn.execute(
            """INSERT OR REPLACE INTO candidatures
               (id_demande, name, filiere, niveau_etudes, avis)
               VALUES (?, ?, ?, ?, ?)""",
            (row["id_demande"], row["name"], row["filiere"],
             row["niveau_etudes"], row["avis"]),
        )
    conn.commit()
    conn.close()
    return len(df)


def load_quotas(quotas_path: str):
    """Charge les quotas depuis un fichier JSON."""
    with open(quotas_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    conn = get_connection()
    for niveau, filieres in data.items():
        for filiere, nb_places in filieres.items():
            conn.execute(
                """INSERT OR REPLACE INTO quotas (niveau_etudes, filiere, nb_places)
                   VALUES (?, ?, ?)""",
                (niveau, filiere, nb_places),
            )
    conn.commit()
    conn.close()


def get_all_candidatures() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM candidatures", conn)
    conn.close()
    return df


def get_quotas() -> dict:
    """Retourne {(niveau, filiere): nb_places}."""
    conn = get_connection()
    rows = conn.execute("SELECT niveau_etudes, filiere, nb_places FROM quotas").fetchall()
    conn.close()
    return {(r["niveau_etudes"], r["filiere"]): r["nb_places"] for r in rows}


def get_favorables_count() -> dict:
    """Retourne {(niveau, filiere): count} des avis favorables."""
    conn = get_connection()
    rows = conn.execute(
        """SELECT niveau_etudes, filiere, COUNT(*) as n
           FROM candidatures WHERE avis = 'Favorable'
           GROUP BY niveau_etudes, filiere"""
    ).fetchall()
    conn.close()
    return {(r["niveau_etudes"], r["filiere"]): r["n"] for r in rows}


def update_avis(id_demande: str, avis: str):
    conn = get_connection()
    conn.execute(
        "UPDATE candidatures SET avis = ? WHERE id_demande = ?",
        (avis, id_demande),
    )
    conn.commit()
    conn.close()


def search_by_field(field: str, query: str) -> dict | None:
    """Recherche exacte sur un champ précis."""
    conn = get_connection()
    row = None
    if field == "numero":
        try:
            num = int(query)
            row = conn.execute(
                "SELECT * FROM candidatures WHERE numero = ?", (num,)
            ).fetchone()
        except ValueError:
            pass
    elif field == "id_russe":
        row = conn.execute(
            "SELECT * FROM candidatures WHERE id_russe = ?", (query,)
        ).fetchone()
    elif field == "name":
        row = conn.execute(
            "SELECT * FROM candidatures WHERE name = ? COLLATE NOCASE", (query,)
        ).fetchone()
    conn.close()
    return dict(row) if row else None


def search_by_field_fuzzy(field: str, query: str) -> list[dict]:
    """Recherche floue sur un champ précis."""
    conn = get_connection()
    if field == "numero":
        rows = conn.execute(
            "SELECT * FROM candidatures WHERE CAST(numero AS TEXT) LIKE ? ORDER BY numero LIMIT 20",
            (f"%{query}%",),
        ).fetchall()
    elif field == "id_russe":
        rows = conn.execute(
            "SELECT * FROM candidatures WHERE id_russe LIKE ? ORDER BY numero LIMIT 20",
            (f"%{query}%",),
        ).fetchall()
    elif field == "name":
        rows = conn.execute(
            "SELECT * FROM candidatures WHERE name LIKE ? COLLATE NOCASE ORDER BY numero LIMIT 20",
            (f"%{query}%",),
        ).fetchall()
    else:
        rows = []
    conn.close()
    return [dict(r) for r in rows]


def get_stats() -> dict:
    conn = get_connection()
    total = conn.execute("SELECT COUNT(*) FROM candidatures").fetchone()[0]
    favorables = conn.execute(
        "SELECT COUNT(*) FROM candidatures WHERE avis = 'Favorable'"
    ).fetchone()[0]
    defavorables = conn.execute(
        "SELECT COUNT(*) FROM candidatures WHERE avis = 'Défavorable'"
    ).fetchone()[0]
    suppleants = conn.execute(
        "SELECT COUNT(*) FROM candidatures WHERE avis = 'Suppléant'"
    ).fetchone()[0]
    conn.close()
    return {
        "total": total,
        "traites": favorables + defavorables + suppleants,
        "favorables": favorables,
        "defavorables": defavorables,
        "suppleants": suppleants,
        "restants": total - favorables - defavorables - suppleants,
    }


NIVEAU_ORDER = ["Licence", "Master", "Doctorat", "Spécialisation"]


def _create_base_docx():
    """Crée un Document Word de base avec logo, en-tête, marges et numéros de page.
    Retourne (doc, add_table_for_section, set_run_font).
    """
    from itertools import groupby

    from docx import Document
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, Twips

    doc = Document()

    # --- Page margins ---
    for section in doc.sections:
        section.top_margin = Twips(720)
        section.bottom_margin = Twips(720)
        section.left_margin = Twips(720)
        section.right_margin = Twips(720)

    # --- Page numbers (centered in footer) ---
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.font.name = "Trebuchet MS"
        run.font.size = Pt(9)
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")
        run._r.append(fldChar1)
        instrText = OxmlElement("w:instrText")
        instrText.set(qn("xml:space"), "preserve")
        instrText.text = " PAGE "
        run._r.append(instrText)
        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")
        run._r.append(fldChar2)

    def set_run_font(run, size=10, bold=False, underline=False):
        run.font.name = "Trebuchet MS"
        run.font.size = Pt(size)
        run.bold = bold
        run.underline = underline

    def set_cell_shading(cell, color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    def merge_row_cells(row, table):
        """Merge all cells in a row using gridSpan."""
        n_cols = len(table.columns)
        tc = row.cells[0]._tc
        tcPr = tc.get_or_add_tcPr()
        gs = OxmlElement("w:gridSpan")
        gs.set(qn("w:val"), str(n_cols))
        tcPr.append(gs)
        tr = row._tr
        tcs = tr.findall(qn("w:tc"))
        for extra_tc in tcs[1:]:
            tr.remove(extra_tc)

    def add_table_for_section(candidates_list):
        table = doc.add_table(rows=1, cols=4, style="Table Grid")

        col_widths = [Twips(567), Twips(4254), Twips(4223), Twips(2127)]
        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        headers = ["N° ", "FILIERE", "NOM ET PRENOMS", "OBSERVATIONS"]
        for i, text in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_run_font(run, size=11)
            set_cell_shading(cell, "BFBFBF")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        def sort_key(c):
            niv = c["niveau_etudes"]
            idx = NIVEAU_ORDER.index(niv) if niv in NIVEAU_ORDER else 99
            return (idx, c["filiere"], c.get("numero", 0) or 0)

        candidates_list = sorted(candidates_list, key=sort_key)

        num = 1
        for niveau, niveau_group in groupby(candidates_list, key=lambda c: c["niveau_etudes"]):
            niveau_group = list(niveau_group)
            row = table.add_row()
            cell = row.cells[0]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(niveau.upper())
            set_run_font(run, size=11, bold=True)
            set_cell_shading(cell, "E7E6E6")
            merge_row_cells(row, table)

            for filiere, fil_group in groupby(niveau_group, key=lambda c: c["filiere"]):
                fil_group = list(fil_group)
                first_row_idx = len(table.rows)

                for i_in_fil, c in enumerate(fil_group):
                    row = table.add_row()
                    cell_num = row.cells[0]
                    cell_num.text = ""
                    p = cell_num.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(str(num))
                    set_run_font(run, size=11)

                    cell_fil = row.cells[1]
                    cell_fil.text = ""
                    if i_in_fil == 0:
                        p = cell_fil.paragraphs[0]
                        run = p.add_run(c["filiere"])
                        set_run_font(run, size=11)

                    cell_name = row.cells[2]
                    cell_name.text = ""
                    p = cell_name.paragraphs[0]
                    run = p.add_run(c["name"])
                    set_run_font(run, size=11)

                    cell_obs = row.cells[3]
                    cell_obs.text = ""
                    p = cell_obs.paragraphs[0]
                    obs = c.get("observation") or ""
                    run = p.add_run(obs)
                    set_run_font(run, size=11)

                    num += 1

                last_row_idx = len(table.rows) - 1
                if last_row_idx > first_row_idx:
                    table.cell(first_row_idx, 1).merge(table.cell(last_row_idx, 1))

        return table

    # --- Logo ---
    logo_path = Path(__file__).parent / "assets" / "logo.png"
    if logo_path.exists():
        section = doc.sections[0]
        avail_width = section.page_width - section.left_margin - section.right_margin
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(logo_path), width=avail_width)

    # --- Document header ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("DIRECTION DES BOURSES ET AIDES UNIVERSITAIRES ")
    set_run_font(run, size=10, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        "LISTE DES ETUDIANTS PRESELECTIONNES POUR BENEFICIER DE LA BOURSE "
        "DE COOPERATION RUSSE AU TITRE DE L\u2019ANNEE ACADEMIQUE 2025-2026"
    )
    set_run_font(run, size=10, bold=True)

    doc.add_paragraph()

    return doc, add_table_for_section, set_run_font


def export_to_docx(output_path: str) -> str:
    """Exporte les candidatures Favorable/Suppléant dans un document Word (.docx)."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    conn = get_connection()
    rows = conn.execute(
        """SELECT * FROM candidatures
           WHERE avis IN ('Favorable', 'Suppléant')
           ORDER BY niveau_etudes, filiere, numero"""
    ).fetchall()
    conn.close()
    candidates = [dict(r) for r in rows]

    favorables = [c for c in candidates if c["avis"] == "Favorable"]
    suppleants = [c for c in candidates if c["avis"] == "Suppléant"]

    doc, add_table_for_section, set_run_font = _create_base_docx()

    # --- Section Titulaires ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run("LISTE DES CANDIDATS TITULAIRES ")
    set_run_font(run, size=13, underline=True)

    add_table_for_section(favorables)

    doc.add_paragraph()

    # --- Section Suppléants ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run("LISTE DES CANDIDATS SUPPLÉANTS")
    set_run_font(run, size=13, underline=True)

    add_table_for_section(suppleants)

    doc.save(output_path)
    return output_path


def export_all_avis_to_docx(output_path: str) -> str:
    """Exporte TOUTES les candidatures (Favorable, Suppléant, Défavorable) dans un document Word."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    conn = get_connection()
    rows = conn.execute(
        """SELECT * FROM candidatures
           WHERE avis IN ('Favorable', 'Suppléant', 'Défavorable')
           ORDER BY niveau_etudes, filiere, numero"""
    ).fetchall()
    conn.close()
    candidates = [dict(r) for r in rows]

    favorables = [c for c in candidates if c["avis"] == "Favorable"]
    suppleants = [c for c in candidates if c["avis"] == "Suppléant"]
    defavorables = [c for c in candidates if c["avis"] == "Défavorable"]

    doc, add_table_for_section, set_run_font = _create_base_docx()

    # --- Section Titulaires ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run("LISTE DES CANDIDATS TITULAIRES ")
    set_run_font(run, size=13, underline=True)
    add_table_for_section(favorables)
    doc.add_paragraph()

    # --- Section Suppléants ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run("LISTE DES CANDIDATS SUPPLÉANTS")
    set_run_font(run, size=13, underline=True)
    add_table_for_section(suppleants)
    doc.add_paragraph()

    # --- Section Défavorables ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run("LISTE DES CANDIDATS NON RETENUS")
    set_run_font(run, size=13, underline=True)
    add_table_for_section(defavorables)

    doc.save(output_path)
    return output_path


def export_avis_to_xlsx(output_path: str) -> str:
    """Exporte les candidatures par avis dans un fichier Excel avec une feuille par avis."""
    from itertools import groupby

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    conn = get_connection()

    avis_config = [
        ("Favorable", "Favorables (Titulaires)"),
        ("Suppléant", "Suppléants"),
        ("Défavorable", "Défavorables"),
    ]

    wb = Workbook()
    wb.remove(wb.active)

    header_font = Font(name="Trebuchet MS", bold=True, size=11)
    header_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
    niveau_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    headers = ["N°", "Filière", "Nom et Prénoms", "Observation"]
    col_widths = [8, 45, 35, 25]

    for avis_value, sheet_name in avis_config:
        rows = conn.execute(
            """SELECT numero, name, filiere, niveau_etudes, observation
               FROM candidatures
               WHERE avis = ?
               ORDER BY niveau_etudes, filiere, numero""",
            (avis_value,),
        ).fetchall()

        ws = wb.create_sheet(title=sheet_name)

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        for col_idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        candidates = [dict(r) for r in rows]

        def sort_key(c):
            niv = c["niveau_etudes"]
            idx = NIVEAU_ORDER.index(niv) if niv in NIVEAU_ORDER else 99
            return (idx, c["filiere"], c.get("numero", 0) or 0)

        candidates.sort(key=sort_key)

        current_row = 2
        for niveau, niveau_group in groupby(candidates, key=lambda c: c["niveau_etudes"]):
            cell = ws.cell(row=current_row, column=1, value=niveau.upper())
            cell.font = Font(name="Trebuchet MS", bold=True, size=11)
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=current_row, column=col_idx).fill = niveau_fill
            ws.merge_cells(
                start_row=current_row, start_column=1,
                end_row=current_row, end_column=len(headers),
            )
            current_row += 1

            for c in niveau_group:
                ws.cell(row=current_row, column=1, value=c.get("numero", ""))
                ws.cell(row=current_row, column=2, value=c.get("filiere", ""))
                ws.cell(row=current_row, column=3, value=c.get("name", ""))
                ws.cell(row=current_row, column=4, value=c.get("observation", ""))
                current_row += 1

    conn.close()
    wb.save(output_path)
    return output_path


def export_quotas_to_xlsx(output_path: str) -> str:
    """Exporte la grille des quotas par filière et par niveau dans un fichier Excel."""
    from itertools import groupby

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    conn = get_connection()
    quotas_rows = conn.execute(
        "SELECT niveau_etudes, filiere, nb_places FROM quotas ORDER BY niveau_etudes, filiere"
    ).fetchall()

    fav_rows = conn.execute(
        """SELECT niveau_etudes, filiere, COUNT(*) as n
           FROM candidatures WHERE avis = 'Favorable'
           GROUP BY niveau_etudes, filiere"""
    ).fetchall()
    conn.close()

    fav_counts = {(r["niveau_etudes"], r["filiere"]): r["n"] for r in fav_rows}

    wb = Workbook()
    ws = wb.active
    ws.title = "Quotas par Filière"

    headers = ["Niveau", "Filière", "Places (Quota)", "Favorables", "Restantes"]
    header_font = Font(name="Trebuchet MS", bold=True, size=11)
    header_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
    niveau_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    col_widths = [18, 50, 16, 14, 14]

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for col_idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    quotas_data = [dict(r) for r in quotas_rows]

    def sort_key(q):
        niv = q["niveau_etudes"]
        idx = NIVEAU_ORDER.index(niv) if niv in NIVEAU_ORDER else 99
        return (idx, q["filiere"])

    quotas_data.sort(key=sort_key)

    current_row = 2
    total_places = 0
    total_fav = 0
    total_restantes = 0

    for niveau, group in groupby(quotas_data, key=lambda q: q["niveau_etudes"]):
        cell = ws.cell(row=current_row, column=1, value=niveau.upper())
        cell.font = Font(name="Trebuchet MS", bold=True, size=12)
        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=current_row, column=col_idx).fill = niveau_fill
        ws.merge_cells(
            start_row=current_row, start_column=1,
            end_row=current_row, end_column=len(headers),
        )
        current_row += 1

        for q in group:
            key = (q["niveau_etudes"], q["filiere"])
            fav = fav_counts.get(key, 0)
            restantes = q["nb_places"] - fav

            total_places += q["nb_places"]
            total_fav += fav
            total_restantes += restantes

            ws.cell(row=current_row, column=1, value=q["niveau_etudes"])
            ws.cell(row=current_row, column=2, value=q["filiere"])
            ws.cell(row=current_row, column=3, value=q["nb_places"]).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=4, value=fav).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=5, value=restantes).alignment = Alignment(horizontal="center")
            current_row += 1

    # --- Ligne de total ---
    total_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
    total_font = Font(name="Trebuchet MS", bold=True, size=11)
    ws.cell(row=current_row, column=2, value="TOTAL").font = total_font
    for col_idx in range(1, len(headers) + 1):
        ws.cell(row=current_row, column=col_idx).fill = total_fill
    ws.cell(row=current_row, column=3, value=total_places).font = total_font
    ws.cell(row=current_row, column=3).alignment = Alignment(horizontal="center")
    ws.cell(row=current_row, column=4, value=total_fav).font = total_font
    ws.cell(row=current_row, column=4).alignment = Alignment(horizontal="center")
    ws.cell(row=current_row, column=5, value=total_restantes).font = total_font
    ws.cell(row=current_row, column=5).alignment = Alignment(horizontal="center")

    wb.save(output_path)
    return output_path


def get_total_quota() -> int:
    """Retourne la somme totale de tous les quotas."""
    conn = get_connection()
    total = conn.execute("SELECT COALESCE(SUM(nb_places), 0) FROM quotas").fetchone()[0]
    conn.close()
    return total


def transfer_quota(source_niveau: str, source_filiere: str,
                   dest_niveau: str, dest_filiere: str,
                   nb_places: int) -> dict:
    """Transfère nb_places de la source vers la destination de manière transactionnelle."""
    if nb_places <= 0:
        return {"success": False, "error": "Le nombre de places doit être supérieur à 0."}

    if source_niveau == dest_niveau and source_filiere == dest_filiere:
        return {"success": False, "error": "La source et la destination doivent être différentes."}

    conn = get_connection()
    try:
        # Lire le quota source
        row_src = conn.execute(
            "SELECT nb_places FROM quotas WHERE niveau_etudes = ? AND filiere = ?",
            (source_niveau, source_filiere),
        ).fetchone()
        if not row_src:
            return {"success": False, "error": f"Quota source introuvable ({source_niveau}, {source_filiere})."}

        quota_source = row_src["nb_places"]

        # Compter les favorables déjà attribués à la source
        fav_row = conn.execute(
            "SELECT COUNT(*) as n FROM candidatures WHERE avis = 'Favorable' AND niveau_etudes = ? AND filiere = ?",
            (source_niveau, source_filiere),
        ).fetchone()
        fav_source = fav_row["n"]

        disponibles = quota_source - fav_source
        if nb_places > disponibles:
            return {
                "success": False,
                "error": f"Places disponibles insuffisantes. Quota : {quota_source}, Favorables : {fav_source}, Disponibles : {disponibles}.",
            }

        # Lire le quota destination
        row_dest = conn.execute(
            "SELECT nb_places FROM quotas WHERE niveau_etudes = ? AND filiere = ?",
            (dest_niveau, dest_filiere),
        ).fetchone()
        if not row_dest:
            return {"success": False, "error": f"Quota destination introuvable ({dest_niveau}, {dest_filiere})."}

        quota_dest = row_dest["nb_places"]

        # Transaction : retirer de la source, ajouter à la destination
        conn.execute(
            "UPDATE quotas SET nb_places = nb_places - ? WHERE niveau_etudes = ? AND filiere = ?",
            (nb_places, source_niveau, source_filiere),
        )
        conn.execute(
            "UPDATE quotas SET nb_places = nb_places + ? WHERE niveau_etudes = ? AND filiere = ?",
            (nb_places, dest_niveau, dest_filiere),
        )
        conn.commit()

        return {
            "success": True,
            "source_nouveau": quota_source - nb_places,
            "dest_nouveau": quota_dest + nb_places,
        }
    except Exception as e:
        conn.rollback()
        return {"success": False, "error": str(e)}
    finally:
        conn.close()


def is_db_loaded() -> bool:
    if not Path(DB_PATH).exists():
        return False
    conn = get_connection()
    try:
        count = conn.execute("SELECT COUNT(*) FROM candidatures").fetchone()[0]
    except sqlite3.OperationalError:
        conn.close()
        return False
    conn.close()
    return count > 0


def reset_db():
    """Supprime et recrée les tables."""
    if Path(DB_PATH).exists():
        Path(DB_PATH).unlink()
