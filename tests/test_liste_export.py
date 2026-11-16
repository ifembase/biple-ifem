import hashlib
import io
import json
import unittest
import zipfile
from concurrent.futures import ThreadPoolExecutor
from xml.etree import ElementTree as ET

from pydantic import ValidationError
import liste_export
from main import ExportListeRequest, export_liste_etudiants

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def ligne(index=1, niveau='DT3', bassin_id='1', **extra):
    return dict(bassin_id=bassin_id, bassin='Bassin Test ' + bassin_id, niveau=niveau,
                matricule=f'TEST-{index}', nom=f'NOM TEST {index}', prenoms='Prénom & Test',
                date_naissance='vers 1990', lieu_naissance='Ville <Test>',
                note_ecrit=12.5, note_soutenance=0, mention='Assez-Bien', annee='2025-2026', **extra)


def payload(lignes=None, **extra):
    data = dict(etudiants=lignes or [ligne()], niveaux=['DT3'], colonnes=list(liste_export.COLONNES),
                critere='tous', regroupement='bassin_niveau', session='Août 2026')
    data.update(extra)
    return ExportListeRequest(**data).model_dump()


def document(donnees):
    return zipfile.ZipFile(io.BytesIO(liste_export.generer_export(donnees)))


class TestListeExport(unittest.TestCase):
    def test_cinq_en_tetes_et_modele_mixte(self):
        for niveau in liste_export.NIVEAUX:
            with self.subTest(niveau=niveau), document(payload([ligne(niveau=niveau)], niveaux=[niveau])) as z:
                xml = z.read('word/document.xml').decode('utf-8')
                self.assertIn('REPOBLIKAN', xml)
                self.assertNotIn('{{SESSION}}', xml)
                self.assertNotIn(liste_export.SLOT, xml)
                self.assertIn('TEST-1', xml)
                self.assertIn(niveau + '-IFEM-SESSION-', z.read('word/header1.xml').decode('utf-8'))
        with document(payload([ligne(), ligne(2, 'M2')], niveaux=['DT3', 'M2'])) as z:
            xml = z.read('word/document.xml').decode('utf-8')
            self.assertIn('LISTE DES ÉTUDIANTS', xml)
            self.assertNotIn('AUTORISÉS', xml)

    def test_colonnes_choisies_uniquement_et_prenoms_separes(self):
        with document(payload(colonnes=['matricule', 'nom', 'prenoms'])) as z:
            doc = ET.fromstring(z.read('word/document.xml'))
            entete = doc.find(f'.//{W}tbl/{W}tr')
            cellules = [''.join(c.itertext()) for c in entete.findall(W+'tc')]
            self.assertEqual(cellules, ['Matricule', 'Nom', 'Prénoms'])
            xml = z.read('word/document.xml').decode()
            self.assertNotIn('Ville &lt;Test&gt;', xml)
            self.assertIn('Prénom &amp; Test', xml)

    def test_aucune_limite_de_cinq_etudiants_et_entetes_repetes(self):
        with document(payload([ligne(i) for i in range(1, 76)])) as z:
            doc = ET.fromstring(z.read('word/document.xml'))
            self.assertEqual(len(doc.findall(f'.//{W}tr')), 76)
            self.assertIsNotNone(doc.find(f'.//{W}tblHeader'))
            self.assertEqual(len(doc.findall(f'.//{W}cantSplit')), 76)

    def test_trois_regroupements(self):
        lignes = [ligne(), ligne(2, 'DTS'), ligne(3, 'DT3', '2')]
        for mode, nombre in [('bassin_niveau', 3), ('niveau', 2), ('liste_unique', 1)]:
            with self.subTest(mode=mode), document(payload(lignes, niveaux=['DT3', 'DTS'], regroupement=mode)) as z:
                doc = ET.fromstring(z.read('word/document.xml'))
                self.assertEqual(len(doc.findall(f'.//{W}tbl')), nombre)
                self.assertEqual(len(doc.findall(f'.//{W}tr')), nombre + 3)

    def test_logos_styles_et_parties_preservees(self):
        manifeste = json.loads((liste_export.TEMPLATES/'provenance.json').read_text(encoding='utf-8'))
        for modele in manifeste:
            code = modele['niveau']
            data = payload() if code == 'DT3' else payload([ligne(niveau=code)], niveaux=[code]) if code != 'MIXTE' else payload([ligne()], niveaux=['DT3', 'DTS'])
            with self.subTest(modele=code), document(data) as z:
                for nom, sha in modele['parties_preservees'].items():
                    self.assertEqual(hashlib.sha256(z.read(nom)).hexdigest(), sha, nom)
                self.assertFalse(any(n.startswith('customXml/') for n in z.namelist()))
                self.assertIn(b'updateFields', z.read('word/settings.xml'))

    def test_modeles_ne_contiennent_aucune_ligne_source(self):
        for path in liste_export.TEMPLATES.glob('*.docx'):
            with self.subTest(modele=path.name), zipfile.ZipFile(path) as z:
                self.assertEqual(ET.fromstring(z.read('word/document.xml')).findall(f'.//{W}tr'), [])
                self.assertNotIn(b'lastModifiedBy', z.read('docProps/core.xml'))

    def test_zero_et_notes_manquantes(self):
        row = ligne()
        row.update(note_ecrit=None, note_soutenance=0)
        with document(payload([row])) as z:
            xml = z.read('word/document.xml').decode('utf-8')
            self.assertIn('0,00', xml)
            self.assertIn('—', xml)

    def test_validation_selection(self):
        cas = [dict(colonnes=[]), dict(colonnes=['mot_de_passe']), dict(colonnes=['nom', 'nom']),
               dict(niveaux=['M2']), dict(niveaux=['DT3', 'DT3']), dict(etudiants=[]),
               dict(etudiants=[ligne(), ligne()])]
        for changement in cas:
            with self.subTest(changement=changement), self.assertRaises(ValidationError):
                data = payload()
                data.update(changement)
                ExportListeRequest(**data)
        for note in [-1, 21, float('nan'), float('inf')]:
            with self.subTest(note=note), self.assertRaises(ValidationError):
                row = ligne()
                row['note_ecrit'] = note
                ExportListeRequest(**dict(payload(), etudiants=[row]))

    def test_generation_concurrente(self):
        with ThreadPoolExecutor(max_workers=4) as pool:
            resultats = list(pool.map(lambda i: liste_export.generer_export(payload([ligne(i)])), range(1, 13)))
        for i, contenu in enumerate(resultats, 1):
            with zipfile.ZipFile(io.BytesIO(contenu)) as z:
                self.assertIsNone(z.testzip())
                self.assertIn(f'TEST-{i}', z.read('word/document.xml').decode())

    def test_top_trois_par_niveau_et_bassin_et_ordre_conserve_dans_word(self):
        lignes = [dict(ligne(i), moyenne_generale=score, nom=nom) for i, score, nom in [
            (1, 12, 'Alpha'), (2, 19, 'Zulu'), (3, 17, 'Bravo'), (4, 15, 'Charlie'), (5, 10, 'Delta')]]
        lignes += [ligne(6, 'DTS', moyenne_generale=8), ligne(7, 'DTS', moyenne_generale=0)]
        lignes += [ligne(i, bassin_id='2', moyenne_generale=score) for i, score in [(8, 10), (9, 18), (10, 16), (11, 14)]]
        data = payload(lignes, niveaux=['DT3', 'DTS'], colonnes=['matricule', 'nom'], classement='top3_bassin_niveau')
        with document(data) as z:
            doc = ET.fromstring(z.read('word/document.xml'))
            tableaux = doc.findall(f'.//{W}tbl')
            matricules = lambda table: [''.join(row.find(W+'tc').itertext()) for row in table.findall(W+'tr')[1:]]
            self.assertEqual([matricules(table) for table in tableaux], [
                ['TEST-2', 'TEST-3', 'TEST-4'], ['TEST-6', 'TEST-7'], ['TEST-9', 'TEST-10', 'TEST-11']])
            self.assertIn('8 étudiant(s)', ''.join(doc.itertext()))
            self.assertIn('Les 3 premiers par niveau et par bassin', ''.join(doc.itertext()))

    def test_top_trois_exige_une_moyenne_generale_et_un_ecrit_complet(self):
        for extra in [dict(moyenne_generale=None), dict(moyenne_generale=15, note_ecrit=None),
                      dict(moyenne_generale=float('nan')), dict(moyenne_generale=21)]:
            with self.subTest(extra=extra), self.assertRaises(ValidationError):
                payload([dict(ligne(), **extra)], classement='top3_bassin_niveau')
        with self.assertRaises(ValidationError):
            payload([ligne(moyenne_generale=15)], classement='top3_bassin_niveau', regroupement='niveau')

    def test_route_renvoie_un_document_word(self):
        response = export_liste_etudiants(ExportListeRequest(**payload()))
        self.assertIn('.docx', response.headers['content-disposition'])
        self.assertIn('wordprocessingml', response.media_type)
        with zipfile.ZipFile(io.BytesIO(response.body)) as z:
            self.assertIsNotNone(ET.fromstring(z.read('word/document.xml')))


if __name__ == '__main__': unittest.main()
