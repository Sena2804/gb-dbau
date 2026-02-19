"""Module de gestion SQLite pour les candidatures CNBAU."""

import json
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


def _parse_real_excel(excel_path: str) -> list[dict]:
    """Parse le fichier CNaBAU structuré avec lignes de séparation niveau/filière."""
    from openpyxl import load_workbook

    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

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
                current_filiere = val_a.split(":", 1)[-1].strip()
            else:
                current_filiere = val_a
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
                diplome_filiere_annee, observation, filiere, niveau_etudes, avis)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (c["id_demande"], c["id_russe"], c["numero"], c["sexe"], c["name"],
             c["date_lieu_naissance"], c["diplome_filiere_annee"],
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
    conn.close()
    return {
        "total": total,
        "traites": favorables + defavorables,
        "favorables": favorables,
        "defavorables": defavorables,
        "restants": total - favorables - defavorables,
    }


def export_to_excel(output_path: str) -> str:
    """Exporte les candidatures triées par numéro de demande."""
    df = get_all_candidatures()
    df = df.drop(columns=["id_demande"])
    df = df.sort_values("numero")
    col_order = [
        "numero", "sexe", "id_russe", "name", "date_lieu_naissance",
        "diplome_filiere_annee", "observation", "filiere", "niveau_etudes", "avis",
    ]
    df = df[[c for c in col_order if c in df.columns]]
    df = df.rename(columns={
        "numero": "NUMÉRO",
        "sexe": "SEXE",
        "id_russe": "ID RUSSE",
        "name": "NOM & PRÉNOM",
        "date_lieu_naissance": "DATE & LIEU DE NAISSANCE",
        "diplome_filiere_annee": "DIPLÔME, FILIÈRE & ANNÉE",
        "observation": "OBSERVATION",
        "filiere": "FILIÈRE",
        "niveau_etudes": "NIVEAU D'ÉTUDES",
        "avis": "AVIS",
    })
    df.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


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
