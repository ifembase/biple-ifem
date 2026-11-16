"""Export Word de listes completes, a partir d'en-tetes IFEM anonymises.

Les modeles fournis sont preserves : seul le contenu du tableau et les champs
de session/titre changent. Aucun nom des listes sources n'est embarque.
"""
import io
import re
import unicodedata
import zipfile
from copy import copy
from collections import OrderedDict
from decimal import Decimal, ROUND_HALF_UP
from functools import lru_cache
from pathlib import Path
from xml.sax.saxutils import escape

NIVEAUX = ("DT3", "DTS", "L3", "M1", "M2")
COLONNES = OrderedDict([
    ("matricule", ("Matricule", 1.4)), ("niveau", ("Niveau", .75)),
    ("nom", ("Nom", 2.6)), ("prenoms", ("Prénoms", 2.4)),
    ("date_naissance", ("Date de naissance", 1.25)), ("lieu_naissance", ("Lieu de naissance", 2.8)),
    ("note_ecrit", ("Note écrit", 1.0)), ("note_soutenance", ("Note soutenance", 1.35)),
    ("mention", ("Mention", 1.6)),
])
CRITERES = {
    "tous": "Tous les étudiants", "avec_soutenance": "Avec une note de soutenance",
    "ecrit_sans_soutenance": "Écrit terminé, sans soutenance",
    "ecrit_termine": "Écrit terminé, avec ou sans soutenance",
    "admis": "Admis avec écrit et soutenance terminés", "en_cours": "Évaluation écrite en cours",
    "sans_notes": "Sans aucune note",
}
TEMPLATES = Path(__file__).parent / "templates" / "export"
SLOT = "<!--IFEM_EXPORT_CONTENT-->"


def _texte(value):
    # Les caracteres de controle XML 1.0 ne doivent pas corrompre un fichier Word.
    return escape("".join(c for c in str(value) if c in "\t\n\r" or 0x20 <= ord(c) <= 0xD7FF or 0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF))


def _cle(value):
    return "".join(c for c in unicodedata.normalize("NFKD", value).casefold() if not unicodedata.combining(c))


@lru_cache(maxsize=6)
def _modele(code):
    with zipfile.ZipFile(TEMPLATES / f"Export_{code}.docx") as archive:
        parties = tuple((info, archive.read(info.filename)) for info in archive.infolist())
    document = next(data.decode("utf-8") for info, data in parties if info.filename == "word/document.xml")
    if document.count(SLOT) != 1:
        raise RuntimeError(f"Modele d'export {code} invalide")
    taille = re.search(r'<w:pgSz\b[^>]*w:w="(\d+)"', document)
    marges = re.search(r'<w:pgMar\b([^>]*)', document)
    largeur = int(taille.group(1)) if taille else 16838
    if marges:
        for cote in ("left", "right"):
            valeur = re.search(rf'w:{cote}="(\d+)"', marges.group(1))
            largeur -= int(valeur.group(1)) if valeur else 1134
    return parties, largeur


def _paragraphe(texte, *, gras=False, taille=20, centre=False, suivant=False, font="Calibri", cellule=False):
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' if cellule else
        '<w:p><w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/>'
    ) + (
        f'<w:keepNext w:val="{1 if suivant else 0}"/><w:keepLines/><w:pageBreakBefore w:val="0"/>'
        '<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
        + f'<w:jc w:val="{"center" if centre else "left"}"/></w:pPr><w:r><w:rPr>'
        + f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{taille}"/>'
        + ('<w:b/>' if gras else '') + f'</w:rPr><w:t xml:space="preserve">{_texte(texte)}</w:t></w:r></w:p>'
    )


def _valeur(ligne, colonne):
    valeur = ligne.get(colonne)
    if valeur is None or valeur == "":
        return "—"
    if colonne in ("note_ecrit", "note_soutenance", "moyenne_generale"):
        return str(Decimal(str(valeur)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)).replace(".", ",")
    return str(valeur)


def _tableau(lignes, colonnes, largeur, mixte, *, definitions=None, compact=False):
    definitions = COLONNES if definitions is None else definitions
    poids = sum(definitions[c][1] for c in colonnes)
    largeurs = [int(largeur * definitions[c][1] / poids) for c in colonnes]
    largeurs[-1] += largeur - sum(largeurs)
    couleur = "D9D9D9" if mixte else "000000"
    marge_verticale = 30 if compact else 65
    bords = ''.join(f'<w:{c} w:val="single" w:sz="6" w:color="{couleur}"/>' for c in ("top", "left", "bottom", "right", "insideH", "insideV"))
    xml = [f'<w:tbl><w:tblPr><w:tblW w:w="{largeur}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
           f'<w:tblBorders>{bords}</w:tblBorders><w:tblCellMar><w:top w:w="{marge_verticale}" w:type="dxa"/><w:bottom w:w="{marge_verticale}" w:type="dxa"/>'
           '<w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>']
    xml += [f'<w:gridCol w:w="{w}"/>' for w in largeurs]
    xml.append('</w:tblGrid>')
    for index, ligne in enumerate([None, *lignes]):
        entete = ligne is None
        xml.append('<w:tr><w:trPr><w:cantSplit/>' + ('<w:tblHeader/>' if entete else '') + '</w:trPr>')
        for colonne, w in zip(colonnes, largeurs):
            fond = "203C73" if mixte and entete else "DCE6F1" if mixte and index % 2 == 0 else "FFFFFF"
            texte = definitions[colonne][0] if entete else _valeur(ligne, colonne)
            centre = entete or colonne in ("niveau", "rang", "date_naissance", "note_ecrit", "note_soutenance", "moyenne_generale")
            p = _paragraphe(texte, gras=entete, taille=18 if compact else (18 if entete else 19) if len(colonnes) > 6 else 21,
                            centre=centre, suivant=entete, font="Times New Roman" if mixte else "Calibri", cellule=True)
            if mixte and entete:
                p = p.replace('<w:rPr>', '<w:rPr><w:color w:val="FFFFFF"/>')
            bords_cellule = ''.join(f'<w:{c} w:val="single" w:sz="6" w:color="{couleur}"/>' for c in ('top', 'left', 'bottom', 'right'))
            xml.append(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:tcBorders>{bords_cellule}</w:tcBorders><w:shd w:fill="{fond}"/><w:vAlign w:val="center"/></w:tcPr>{p}</w:tc>')
        xml.append('</w:tr>')
    xml.append('</w:tbl>')
    return ''.join(xml)


def _groupes(lignes, mode, classement="tous"):
    groupes = OrderedDict()
    top3 = classement == "top3_bassin_niveau"
    if top3:
        cle = lambda l: (_cle(l["bassin"]), l["bassin_id"], NIVEAUX.index(l["niveau"]),
                         -l["moyenne_generale"], _cle(l["nom"]), _cle(l["prenoms"]), l["matricule"])
    elif mode == "bassin_niveau":
        cle = lambda l: (_cle(l["bassin"]), l["bassin_id"], NIVEAUX.index(l["niveau"]), _cle(l["nom"]), _cle(l["prenoms"]), l["matricule"])
    elif mode == "niveau":
        cle = lambda l: (NIVEAUX.index(l["niveau"]), _cle(l["nom"]), _cle(l["prenoms"]), l["matricule"])
    else:
        cle = lambda l: (_cle(l["nom"]), _cle(l["prenoms"]), l["matricule"])
    for ligne in sorted(lignes, key=cle):
        identifiant = (ligne["bassin_id"], ligne["niveau"]) if mode == "bassin_niveau" else (ligne["niveau"],) if mode == "niveau" else ("unique",)
        groupe = groupes.setdefault(identifiant, [])
        if not top3 or len(groupe) < 3:
            groupe.append(ligne)
    return list(groupes.values())


def _palmares_global(lignes):
    """Trois meilleurs de chaque niveau, tous bassins selectionnes confondus.

    Les trois premiers locaux suffisent comme candidats : un etudiant exclu
    du top trois de son bassin ne peut pas entrer dans le top trois global.
    """
    effectifs = {}
    palmares = []
    for ligne in sorted(lignes, key=lambda l: (
        NIVEAUX.index(l["niveau"]), -l["moyenne_generale"], _cle(l["nom"]),
        _cle(l["prenoms"]), l["matricule"], l["bassin_id"],
    )):
        niveau = ligne["niveau"]
        rang = effectifs.get(niveau, 0) + 1
        if rang > 3:
            continue
        effectifs[niveau] = rang
        palmares.append(dict(ligne, rang=rang, nom_complet=" ".join(
            valeur for valeur in (ligne["nom"], ligne["prenoms"]) if valeur)))
    return palmares


def _tableau_palmares(lignes, largeur):
    colonnes = OrderedDict([
        ("niveau", ("Niveau", .7)), ("rang", ("Rang", .45)),
        ("matricule", ("Matricule", 1.3)), ("nom_complet", ("Noms et prénoms", 3.5)),
        ("bassin", ("Bassin", 2.1)), ("moyenne_generale", ("Moyenne générale", 1.2)),
    ])
    return _tableau(_palmares_global(lignes), list(colonnes), largeur, True,
                    definitions=colonnes, compact=True)


def generer_export(data):
    """Un unique DOCX, avec autant de tableaux et de pages que necessaire."""
    code = data["niveaux"][0] if len(data["niveaux"]) == 1 else "MIXTE"
    parties, largeur = _modele(code)
    classement = data.get("classement", "tous")
    groupes = _groupes(data["etudiants"], data["regroupement"], classement)
    lignes = [ligne for groupe in groupes for ligne in groupe]
    session = data.get("session") or "Non précisée"
    titre = "LISTE DES ADMIS" if data["critere"] == "admis" else "LISTE DES ÉTUDIANTS"
    # Les en-tetes des references contiennent des zones de texte flottantes.
    # Reprendre le flux sous ces zones avant le titre et les tableaux exportes.
    contenu = ['<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:keepNext/></w:pPr><w:r><w:br w:type="textWrapping" w:clear="all"/></w:r></w:p>']
    if code != "MIXTE":
        contenu.append(_paragraphe(titre, gras=True, taille=24, centre=True, suivant=True))
    contenu.append(_paragraphe(f'Sélection : {CRITERES[data["critere"]]} · {len(lignes)} étudiant(s)', centre=True, suivant=True))
    bassins = sorted({l["bassin"] for l in lignes}, key=_cle)
    annees = sorted({l.get("annee", "") for l in lignes if l.get("annee")})
    contenu.append(_paragraphe(
        f'Bassins sélectionnés : {len(bassins)}' if classement == "top3_bassin_niveau" else "Bassins : " + ", ".join(bassins),
        taille=18, suivant=True))
    if code == "MIXTE":
        contenu.append(_paragraphe("Niveaux : " + ", ".join(n for n in NIVEAUX if n in data["niveaux"]) + " · Session : " + session, taille=18, suivant=True))
    if annees:
        contenu.append(_paragraphe("Année(s) universitaire(s) : " + ", ".join(annees), taille=18, suivant=True))
    if classement == "top3_bassin_niveau":
        contenu.append(_paragraphe("Les 3 meilleurs étudiants par niveau", gras=True, taille=24, centre=True, suivant=True))
        contenu.append(_paragraphe("Tous les bassins sélectionnés confondus — Moyenne générale décroissante",
                                   taille=18, centre=True, suivant=True))
        contenu.append(_tableau_palmares(lignes, largeur))
        # Le detail par bassin commence sur une nouvelle page.
        titre_detail = _paragraphe("Les 3 premiers par niveau et par bassin", gras=True, taille=24, centre=True, suivant=True)
        contenu.append(titre_detail.replace('<w:pageBreakBefore w:val="0"/>', '<w:pageBreakBefore w:val="1"/>'))
    for groupe in groupes:
        premier = groupe[0]
        intitule = (f'Bassin {premier["bassin"]} — Niveau {premier["niveau"]}' if data["regroupement"] == "bassin_niveau"
                    else f'Niveau {premier["niveau"]} — Bassins regroupés' if data["regroupement"] == "niveau"
                    else "Tous les bassins et niveaux sélectionnés")
        contenu.append(_paragraphe(f'{intitule} · {len(groupe)} étudiant(s)', gras=True, taille=22, suivant=True))
        contenu.append(_tableau(groupe, data["colonnes"], largeur, code == "MIXTE"))
    remplacement = {
        '{{SESSION}}': _texte(session), '{{TITRE}}': _texte(titre),
        '{{ENTETE_SESSION}}': _texte(f'{code}-IFEM-SESSION-{session}'),
    }
    resultat = io.BytesIO()
    with zipfile.ZipFile(resultat, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for info, original in parties:
            contenu_partie = original
            if info.filename == 'word/document.xml' or re.match(r'word/(header|footer)\d+\.xml$', info.filename):
                xml = original.decode('utf-8')
                # Remplacer avant insertion des valeurs utilisateurs : des accolades
                # saisies dans un nom ne doivent pas devenir des variables.
                for marqueur, valeur in remplacement.items():
                    xml = xml.replace(marqueur, valeur)
                if info.filename == 'word/document.xml':
                    if code != 'MIXTE':
                        # Les colonnes repetees ne doivent pas masquer le titre
                        # courant des modeles dont la marge haute est tres courte.
                        def marge_haute(match):
                            partie = match.group(0)
                            top = re.search(r'w:top="(\d+)"', partie)
                            header = re.search(r'w:header="(\d+)"', partie)
                            minimum = (int(header.group(1)) if header else 708) + 400
                            return re.sub(r'w:top="\d+"', f'w:top="{max(int(top.group(1)), minimum)}"', partie) if top else partie
                        xml = re.sub(r'<w:pgMar\b[^>]*/>', marge_haute, xml)
                    fin = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
                    xml = xml.replace(SLOT + '<w:p/>', ''.join(contenu) + fin)
                contenu_partie = xml.encode('utf-8')
            archive.writestr(copy(info), contenu_partie)
    return resultat.getvalue()
