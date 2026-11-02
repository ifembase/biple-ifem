# -*- coding: utf-8 -*-
"""Tests du moteur (bibliothèque standard uniquement) :
    python -m unittest discover -s tests -v
Vérifie sur les 15 templates que les valeurs sont écrites et qu'AUCUN
exemple ni accolade {{ }} ne subsiste dans le document généré."""
import io
import os
import re
import sys
import unittest
import zipfile
from html import unescape

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docgen  # noqa: E402

# Valeurs d'exemple présentes dans les templates fournis.
EXEMPLES = [
    "RANDRIANTSEHENO", "TNR 227", "Andramasina", "01/02/1984", "MIANDRY", "TNR 192", "Vohipeno", "31/08/1972",
    "RATINAHARIMALALA", "TNR 224", "15/10/1974", "RAHELIARIMALALA", "ASN 032", "29/12/1985", "RAZANAMIADANA",
    "TNR 229", "14/09/1979", "RAZAFIMAHATRADRAIBE", "16/06/1988", "MAHASOABE", "AMB 001", "RAZANAJAFY", "AMB 007",
    "AMB 025", "SOLOFOHARIVAO", "AMB 029", "RASOATAHINJANAHARY", "AMB 041", "TNR 158", "RAJONIRIANA", "ABM 052",
    "RABARIOELINIRINA", "VNO 025", "ABP 002", "TIANAHARILAZA", "ABP 051", "RAKOTONIAINA",
]
BASE = dict(matricule="ZZZ 099", nom="ANDRIAMAMPIANINA Fanja & Co", mention="Très-Bien",
            date_naissance="05/07/1990", lieu_naissance="Fianarantsoa <II>")


def _texte(docx: bytes) -> str:
    xml = zipfile.ZipFile(io.BytesIO(docx)).read("word/document.xml").decode("utf-8")
    return unescape("".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", xml)))


class TestDocgen(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        docgen.load_all()

    def _verifie(self, kind, code, data):
        t = _texte(docgen.generate(kind, code, data))
        self.assertNotIn("{{", t)
        self.assertNotIn("}}", t)
        for ex in EXEMPLES:
            self.assertNotIn(ex, t, f"{kind}/{code} : exemple {ex!r} encore présent")
        return t

    def test_diplomes_et_authentifications(self):
        for kind in ("diplome", "authentification"):
            for code in docgen.NIVEAUX:
                t = self._verifie(kind, code, BASE)
                for v in ("ZZZ 099", "ANDRIAMAMPIANINA Fanja & Co", "Très-Bien", "05/07/1990", "Fianarantsoa <II>"):
                    self.assertIn(v, t, f"{kind}/{code} : {v!r} absent")

    def test_releves(self):
        for code in docgen.NIVEAUX:
            n = len(docgen.releve_modules(code))
            t = self._verifie("releve", code, dict(BASE, notes=[12] * n, note_soutenance=16))
            self.assertIn("ZZZ 099", t)

    def test_releve_calcul(self):
        # Mêmes notes que le relevé L2 d'origine -> mêmes totaux (214,5 / 13,41 / 15,45).
        d = dict(BASE, notes=["15", "15", "11,5", "12,5", "12", "14"], note_soutenance="17,5")
        t = self._verifie("releve", "L2", d)
        for v in ("214,5", "13,41", "15,45"):
            self.assertIn(v, t)

    def test_releve_mauvais_nombre_de_notes(self):
        with self.assertRaises(ValueError):
            docgen.generate("releve", "L1", dict(BASE, notes=[10, 11], note_soutenance=12))
        with self.assertRaises(ValueError):
            docgen.generate("releve", "L1", dict(BASE, notes=[10, 11, 12, 13, 14, 25], note_soutenance=12))


class TestListe(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        docgen.load_all()

    @staticmethod
    def _lignes(docx):
        xml = zipfile.ZipFile(io.BytesIO(docx)).read("word/document.xml").decode("utf-8")
        res = []
        for tr in re.findall(r"<w:tr[ >].*?</w:tr>", xml, re.S):
            cells = [unescape("".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", tc))).strip()
                     for tc in re.findall(r"<w:tc[ >].*?</w:tc>", tr, re.S)]
            if len(cells) >= 4 and re.fullmatch(r"\d+", cells[0] or "x"):
                res.append((cells[0], cells[1], cells[2], cells[-1]))
        return res

    def test_capacite(self):
        for code in docgen.NIVEAUX:
            self.assertEqual(docgen.liste_capacite(code), 5)

    def test_liste_complete_partielle(self):
        for code in docgen.NIVEAUX:
            for n in (5, 3, 1):
                etu = [dict(matricule=f"ZZ {i}", nom=f"NOM{i}") for i in range(1, n + 1)]
                lignes = self._lignes(docgen.generate("liste", code, dict(etudiants=etu)))
                self.assertEqual([l[0] for l in lignes], ["1", "2", "3", "4", "5"])  # NB toujours 1..5
                for i, l in enumerate(lignes):
                    if i < n:
                        self.assertEqual(l[1:], (f"ZZ {i+1}", f"NOM{i+1}", "Admis"))
                    else:
                        self.assertEqual(l[1:3], ("_", "_"))
                        self.assertNotEqual(l[3], "Admis")

    def test_liste_trop_d_etudiants(self):
        etu = [dict(matricule=f"ZZ {i}", nom=f"NOM{i}") for i in range(6)]
        with self.assertRaises(ValueError):
            docgen.generate("liste", "L2", dict(etudiants=etu))
        with self.assertRaises(ValueError):
            docgen.generate("liste", "L2", dict(etudiants=[]))

    def test_date_naissance_approximative(self):
        """« vers 1990 » : le mot « le » du modèle est retiré (« Né(e) vers 1990 »)."""
        import io, re, zipfile
        def texte(b):
            xml = zipfile.ZipFile(io.BytesIO(b)).read("word/document.xml").decode("utf-8")
            return re.sub(r"<[^>]+>", "", xml)
        for kind in ("diplome", "authentification"):
            for code in docgen.NIVEAUX:
                base = dict(matricule="A 1", nom="X Y", mention="Bien", lieu_naissance="Tana")
                t = texte(docgen.generate(kind, code, dict(base, date_naissance="vers 1990")))
                self.assertRegex(t, r"[Nn]é\(e\)\s+vers 1990")
                self.assertNotRegex(t, r"[Nn]é\(e\)\s+le\s+vers")
                t = texte(docgen.generate(kind, code, dict(base, date_naissance="16/02/1974")))
                self.assertRegex(t, r"[Nn]é\(e\)\s+le\s+16/02/1974")


class TestArchive(unittest.TestCase):
    def test_nom_archive(self):
        self.assertEqual(docgen.nom_archive("Ambovombe"), "AMBOVOMBE.zip")
        self.assertEqual(docgen.nom_archive("Port Bergé"), "PORT_BERGE.zip")
        self.assertEqual(docgen.nom_archive("  nosy varika "), "NOSY_VARIKA.zip")
        self.assertEqual(docgen.nom_archive(None), "DIPLOMES.zip")
        self.assertEqual(docgen.nom_archive("   "), "DIPLOMES.zip")

    def test_zip_un_dossier_par_niveau(self):
        docgen.load_all()
        base = dict(mention="Bien", date_naissance="05/07/1990", lieu_naissance="Tana")
        entrees = []
        for dossier, code, mat in [("DT3", "L1", "A 1"), ("DTS", "L2", "B 2"), ("DT3", "L1", "A 3"), ("DT3", "L1", "A 1")]:
            data = docgen.generate("diplome", code, dict(base, matricule=mat, nom="NOM " + mat))
            entrees.append((dossier, f"diplome_{code}_{docgen.slugify(mat)}.docx", data))
        z = zipfile.ZipFile(io.BytesIO(docgen.assembler_zip(entrees)))
        self.assertEqual(sorted(z.namelist()), [
            "DT3/diplome_L1_A_1.docx", "DT3/diplome_L1_A_1_1.docx", "DT3/diplome_L1_A_3.docx",
            "DTS/diplome_L2_B_2.docx"])  # doublon suffixé, rien d'écrasé
        self.assertIsNone(z.testzip())
        self.assertIn("NOM A 3", _texte(z.read("DT3/diplome_L1_A_3.docx")))


if __name__ == "__main__":
    unittest.main()
