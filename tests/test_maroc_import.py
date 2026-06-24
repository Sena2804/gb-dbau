import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

import database as db


def create_maroc_workbook(path: Path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Tableau COE"
    sheet.append([
        "N°",
        "SEXE",
        "NOM ET PRENOM (S)",
        "DATE & LIEU DE NAISSANCE",
        "DIPLOME / FILIERE / ANNEE",
        "MOYENNE / MENTION",
        "OBSERVATION",
        "AVIS CNaBAU",
    ])
    sheet.append(["NIVEAU : LICENCE"])
    sheet.append(["Filière : Sciences de l'Ingénieur (02 places)"])
    sheet.append([
        1,
        "F",
        "CANDIDATE TEST",
        "01/01/2008 à Cotonou",
        "BAC/C/2025",
        "16,25 Très-bien",
        "RAS",
        None,
    ])
    sheet.append(["NIVEAU : SPECIALITE MEDICALE"])
    sheet.append(["Filière : Filière : Pneumologie (01 place)"])
    sheet.append([
        2,
        "M",
        "CANDIDAT TEST",
        "01/01/1999 à Porto-Novo",
        "DOCTORAT EN MEDECINE/2025",
        None,
        None,
        None,
    ])
    workbook.save(path)


def create_professional_workbook(path: Path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Tableau COE"
    sheet.append([
        "N°",
        "SEXE",
        "NOM ET PRENOM (S)",
        "DATE & LIEU DE NAISSANCE",
        "DIPLOME / FILIERE / ANNEE",
        "MOYENNE / MENTION",
        "OBSERVATION",
        "AVIS CNaBAU",
    ])
    sheet.append(["NIVEAU : BAC + 2 ans"])
    sheet.append(["Filière : Génie Electrique (02 places)"])
    sheet.append([
        1,
        "M",
        "PREMIER DOSSIER",
        "01/01/2006 à Cotonou",
        "BAC D 2025",
        "13,00 Assez bien",
        "RAS",
        None,
    ])
    sheet.append([
        2,
        "F",
        "SECOND DOSSIER",
        "02/02/2006 à Porto-Novo",
        "BAC C 2025",
        "14,00 Bien",
        "RAS",
        "BAC ETRANGER",
    ])
    sheet.append([
        1,
        "F",
        "DOUBLON À IGNORER",
        "03/03/2006 à Parakou",
        "BAC D 2025",
        "12,00 Assez bien",
        "RAS",
        None,
    ])
    workbook.save(path)


class MarocImportTest(unittest.TestCase):
    def test_imports_candidates_and_quotas(self):
        original_db_path = db.DB_PATH
        try:
            with tempfile.TemporaryDirectory() as tmp:
                db.DB_PATH = str(Path(tmp) / "session.db")
                workbook_path = Path(tmp) / "maroc.xlsx"
                create_maroc_workbook(workbook_path)
                db.init_db()

                count = db.load_excel_to_db(str(workbook_path))
                candidates = db.get_all_candidatures()
                quotas = db.get_quotas()

                self.assertEqual(count, 2)
                self.assertEqual(len(candidates), 2)
                self.assertEqual(len(quotas), 2)
                self.assertEqual(sum(quotas.values()), 3)
                self.assertEqual(candidates["numero"].nunique(), 2)
                self.assertEqual(
                    set(candidates["niveau_etudes"]),
                    {"Licence", "Spécialité médicale"},
                )
                self.assertTrue(
                    all(
                        (row.niveau_etudes, row.filiere) in quotas
                        for row in candidates.itertuples()
                    )
                )
                first = candidates.sort_values("numero").iloc[0]
                self.assertEqual(first["id_demande"], "MAR-0001/26")
                self.assertEqual(first["name"], "CANDIDATE TEST")
        finally:
            db.DB_PATH = original_db_path

    def test_imports_professional_file_and_keeps_first_duplicate_number(self):
        original_db_path = db.DB_PATH
        try:
            with tempfile.TemporaryDirectory() as tmp:
                db.DB_PATH = str(Path(tmp) / "session.db")
                workbook_path = Path(tmp) / "maroc_fp.xlsx"
                create_professional_workbook(workbook_path)
                db.init_db()

                count = db.load_excel_to_db(str(workbook_path))
                candidates = db.get_all_candidatures().sort_values("numero")
                quotas = db.get_quotas()

                self.assertEqual(count, 2)
                self.assertEqual(len(candidates), 2)
                self.assertEqual(set(candidates["niveau_etudes"]), {"Bac + 2 ans"})
                self.assertEqual(candidates.iloc[0]["name"], "PREMIER DOSSIER")
                self.assertEqual(candidates.iloc[1]["avis"], "En attente")
                self.assertIn("BAC ETRANGER", candidates.iloc[1]["observation"])
                self.assertEqual(quotas, {("Bac + 2 ans", "Génie Electrique"): 2})
        finally:
            db.DB_PATH = original_db_path


if __name__ == "__main__":
    unittest.main()
