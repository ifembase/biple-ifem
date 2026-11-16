"""Export XLSX des notes par bassin, d'apres les en-tetes du modele fourni.

Le classeur est assemble en OOXML avec la bibliotheque standard. Aucun fichier
Excel contenant les anciens etudiants n'est embarque et aucune dependance de
tableur n'est necessaire sur Render.
"""
import io
import json
import re
import unicodedata
import zipfile
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from functools import lru_cache
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DOSSIER = Path(__file__).parent / "templates" / "source_notes"
NIVEAUX = ("DT3", "DTS", "L3", "M1", "M2")
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"


class Schema(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True, allow_inf_nan=False)


def _nombre_decimal(valeur):
    if isinstance(valeur, str):
        valeur = valeur.strip().replace('\u00a0', '').replace('\u202f', '').replace(',', '.')
        return valeur if valeur else None
    return valeur


class MatiereSource(Schema):
    nom: str = Field(..., min_length=1, max_length=300)
    coefficient: float = Field(..., gt=0, le=1000)

    @field_validator("coefficient", mode="before")
    @classmethod
    def coefficient_decimal(cls, valeur):
        return _nombre_decimal(valeur)


class EtudiantSource(Schema):
    matricule: str = Field(..., min_length=1, max_length=100)
    nom: str = Field(..., min_length=1, max_length=200)
    prenoms: str = Field("", max_length=200)
    date_naissance: str = Field("", max_length=100)
    lieu_naissance: str = Field("", max_length=500)
    annee: str = Field("", max_length=100)
    notes: list[float | None] = Field(..., max_length=100)
    # Un export doit conserver les anciennes valeurs, meme hors plage, sans
    # les transformer en notes /20. Elles seront signalees dans le classeur.
    note_soutenance: float | None = None

    @field_validator("notes", mode="before")
    @classmethod
    def notes_decimales(cls, valeurs):
        return [_nombre_decimal(v) for v in valeurs] if isinstance(valeurs, (list, tuple)) else valeurs

    @field_validator("note_soutenance", mode="before")
    @classmethod
    def soutenance_decimale(cls, valeur):
        return _nombre_decimal(valeur)


class NiveauSource(Schema):
    code: str = Field(..., min_length=1, max_length=80)
    intitule: str = Field("", max_length=300)
    matieres: list[MatiereSource] = Field(..., max_length=100)
    etudiants: list[EtudiantSource] = Field(..., max_length=10000)

    @model_validator(mode="after")
    def notes_alignees(self):
        if any(len(e.notes) != len(self.matieres) for e in self.etudiants):
            raise ValueError("Le nombre de notes doit correspondre aux matieres du niveau")
        return self


class SourceNotesRequest(Schema):
    bassin: str = Field(..., min_length=1, max_length=200)
    niveaux: list[NiveauSource] = Field(..., min_length=1, max_length=50)

    @model_validator(mode="after")
    def effectif_limite(self):
        if sum(len(n.etudiants) for n in self.niveaux) > 10000:
            raise ValueError("Un export est limite a 10000 etudiants")
        return self


@lru_cache(maxsize=1)
def _modele():
    return json.loads((DOSSIER / "modele.json").read_text(encoding="utf8"))


@lru_cache(maxsize=1)
def _styles_avec_alerte():
    xml = (DOSSIER / 'styles.xml').read_text(encoding='utf8')
    dxf = '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill></dxf>'
    motif = re.search(r'<dxfs\b([^>]*?)(?:\s*/>|>(.*?)</dxfs>)', xml, re.DOTALL)
    if motif:
        compte = re.search(r'count="(\d+)"', motif.group(1))
        index = int(compte.group(1)) if compte else 0
        bloc = f'<dxfs count="{index+1}">{motif.group(2) or ""}{dxf}</dxfs>'
        xml = xml[:motif.start()] + bloc + xml[motif.end():]
    else:
        index = 0
        xml = xml.replace('<tableStyles', f'<dxfs count="1">{dxf}</dxfs><tableStyles', 1)
    return xml.encode('utf8'), index


def _hors_plage(note):
    return note is not None and not 0 <= note <= 20


def _texte(value):
    return escape("".join(c for c in str(value) if c in "\t\n\r" or 0x20 <= ord(c) <= 0xD7FF or 0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF))


def _col(index):
    resultat = ""
    while index:
        index, reste = divmod(index - 1, 26)
        resultat = chr(65 + reste) + resultat
    return resultat


def _cell(adresse, valeur, style, formule=None):
    attrs = f'r="{adresse}" s="{style}"'
    if formule is not None:
        numerique = isinstance(valeur, (int, float, Decimal))
        return f'<c {attrs} t="{"n" if numerique else "str"}"><f>{_texte(formule)}</f><v>{_texte(valeur) if valeur is not None else ""}</v></c>'
    if valeur is None:
        return f'<c {attrs}/>'
    if isinstance(valeur, (int, float, Decimal)):
        return f'<c {attrs}><v>{_texte(valeur)}</v></c>'
    return f'<c {attrs} t="inlineStr"><is><t xml:space="preserve">{_texte(valeur)}</t></is></c>'


def _date_naissance(valeur):
    for forme, extrait in (("%Y-%m-%d", valeur[:10]), ("%d/%m/%Y", valeur)):
        try:
            jour = datetime.strptime(extrait, forme).date()
            if jour.year >= 1900:
                serial = (jour - date(1899, 12, 30)).days
                return serial - (1 if jour < date(1900, 3, 1) else 0), True
        except ValueError:
            pass
    return valeur, False


def _code(code):
    cle = "".join(c for c in unicodedata.normalize("NFKD", code.upper()) if c.isalnum() and not unicodedata.combining(c))
    return {"L1": "DT3", "DTSE": "DT3", "DT3SE": "DT3", "L2": "DTS", "DTSSE": "DTS",
            "L3SE": "L3", "M1SE": "M1", "M2SE": "M2"}.get(cle, cle)


def _mention(moyenne, saisies):
    if moyenne is None:
        return "Évaluation en cours" if saisies else "En attente"
    for seuil, texte in ((8, "Ajourné"), (10, "Rattrapage"), (12, "Passable"), (14, "Assez-Bien"), (16, "Bien"), (18, "Très-Bien")):
        if moyenne < seuil:
            return texte
    return "Honorable"


def _ligne(etudiant, matieres, numero, rang, styles):
    cellules = [_cell(f'A{numero}', rang, styles['centre']), _cell(f'B{numero}', etudiant['matricule'], styles['texte']),
                _cell(f'C{numero}', ' '.join(v for v in (etudiant['nom'], etudiant['prenoms']) if v), styles['texte'])]
    naissance, exacte = _date_naissance(etudiant['date_naissance'])
    cellules.extend([_cell(f'D{numero}', naissance, styles['date'] if exacte else styles['centre']),
                     _cell(f'E{numero}', etudiant['lieu_naissance'], styles['texte'])])
    entete = numero - rang
    brutes, ponderees, coefficients = [], [], []
    total = Decimal(0)
    saisies = 0
    for index, (matiere, note) in enumerate(zip(matieres, etudiant['notes'])):
        col = 6 + index * 2
        brute, ponderee = f'{_col(col)}{numero}', f'{_col(col+1)}{numero}'
        coeff = f'${_col(col+1)}${entete}'
        brutes.append(brute)
        ponderees.append(ponderee)
        coefficients.append(coeff)
        valeur = Decimal(str(note)) * Decimal(str(matiere['coefficient'])) if note is not None and not _hors_plage(note) else None
        if valeur is not None:
            total += valeur
            saisies += 1
        cellules.extend([_cell(brute, note, styles['nombre']),
                         _cell(ponderee, valeur, styles['nombre'], f'IF(OR({brute}="",{brute}<0,{brute}>20),"",{brute}*{coeff})')])
    debut = 6 + len(matieres) * 2
    r, s, t, u, v, w = [f'{_col(debut+i)}{numero}' for i in range(6)]
    complet = bool(matieres) and saisies == len(matieres)
    somme_coeff = sum(Decimal(str(m['coefficient'])) for m in matieres)
    moyenne = (total / somme_coeff).quantize(Decimal('.01'), rounding=ROUND_HALF_UP) if complet else None
    soutenance = etudiant['note_soutenance']
    ecrit_invalide = any(_hors_plage(note) for note in etudiant['notes'])
    soutenance_invalide = _hors_plage(soutenance)
    generale = None if soutenance_invalide else ((moyenne + Decimal(str(soutenance))) / 2).quantize(Decimal('.01'), rounding=ROUND_HALF_UP) if moyenne is not None and soutenance is not None else moyenne
    if matieres:
        compte = 'COUNT(' + ','.join(brutes) + ')'
        hors_plage = 'OR(' + ','.join(f'{cellule}<0,{cellule}>20' for cellule in brutes) + ')'
        mention_valide = f'IF({v}<8,"Ajourné",IF({v}<10,"Rattrapage",IF({v}<12,"Passable",IF({v}<14,"Assez-Bien",IF({v}<16,"Bien",IF({v}<18,"Très-Bien","Honorable"))))))'
        formules = [f'IF(COUNT({",".join(ponderees)})=0,"",SUM({",".join(ponderees)}))',
                    f'IF(AND({compte}={len(matieres)},NOT({hors_plage})),ROUND({r}/SUM({",".join(coefficients)}),2),"")',
                    f'IF({hors_plage},"Note hors plage à vérifier",IF({s}="","",IF({s}>=10,"Admis(e)","Non-admis(e)")))', None,
                    f'IF(OR({s}="",{u}<0,{u}>20),"",IF({u}="",{s},ROUND(({s}+{u})/2,2)))',
                    f'IF({hors_plage},"Note hors plage à vérifier",IF(OR({u}<0,{u}>20),"Soutenance à vérifier",IF({s}="",IF({compte}=0,"En attente","Évaluation en cours"),{mention_valide})))']
    else:
        formules = [None] * 6
    examen = 'Note hors plage à vérifier' if ecrit_invalide else None if moyenne is None else 'Admis(e)' if moyenne >= 10 else 'Non-admis(e)'
    mention = 'Note hors plage à vérifier' if ecrit_invalide else 'Soutenance à vérifier' if soutenance_invalide else _mention(generale, saisies)
    valeurs = [total if saisies else None, moyenne, examen, soutenance, generale, mention]
    for index, (adresse, valeur, formule) in enumerate(zip((r, s, t, u, v, w), valeurs, formules)):
        cellules.append(_cell(adresse, valeur, styles['texte'] if index in (2, 5) else styles['nombre'], formule))
    return f'<row r="{numero}" ht="30" customHeight="1">' + ''.join(cellules) + '</row>'


def _feuille(niveau, bassin, annees, modele, dxf_alerte):
    h = modele['entete']
    matieres = niveau['matieres']
    etudiants = sorted(niveau['etudiants'], key=lambda e: (e['nom'].casefold(), e['prenoms'].casefold(), e['matricule']))
    fin = h + len(etudiants)
    dernier = 11 + 2 * len(matieres)
    noms = ['N°', 'Matricule', 'Noms et prénoms', 'Date de naissance', 'Lieu de naissance']
    styles = modele['styles_entete'][:5]
    largeurs = [max(v, minimum) for v, minimum in zip(modele['largeurs'][:5], (7, 15, 45, 16, 28))]
    for index, matiere in enumerate(matieres):
        noms.extend([matiere['nom'], matiere['coefficient']])
        styles.extend(modele['styles_entete'][5:7])
        largeurs.extend([18, 11])
    noms.extend(['Total', 'Note écrite', 'Mention Examen', 'Note soutenance', 'Moyenne', 'Mention'])
    styles.extend(modele['styles_entete'][17:23])
    largeurs.extend([14, 15, 18, 16, 15, 20])
    rows = []
    for ligne in modele['lignes']:
        cellules = []
        for cellule in ligne['cellules']:
            texte = cellule['valeur'].replace('{{BASSIN}}', bassin.upper()).replace('{{ANNEES}}', ', '.join(annees) or 'Non renseignée')
            if texte.startswith('Niveau :') and _code(niveau['code']) not in NIVEAUX:
                texte = 'Niveau : ' + (niveau['intitule'] or niveau['code'])
            cellules.append(_cell(f'{_col(cellule["col"])}{ligne["numero"]}', texte, cellule['style']))
        rows.append(f'<row r="{ligne["numero"]}" ht="{ligne["hauteur"]}" customHeight="1">' + ''.join(cellules) + '</row>')
    rows.append(f'<row r="{h}" ht="90" customHeight="1">' + ''.join(_cell(f'{_col(i)}{h}', nom, style) for i, (nom, style) in enumerate(zip(noms, styles), 1)) + '</row>')
    rows.extend(_ligne(etudiant, matieres, h+i, i, _modele()['styles']) for i, etudiant in enumerate(etudiants, 1))
    fusions = ''.join(f'<mergeCell ref="{ref}"/>' for ref in modele['fusions'])
    plages = ' '.join(f'{_col(c)}{h+1}:{_col(c)}{fin}' for c in [*[6+i*2 for i in range(len(matieres))], 9+len(matieres)*2])
    validations = []
    if etudiants:
        validations.append(f'<dataValidation type="decimal" operator="between" allowBlank="1" showErrorMessage="1" errorTitle="Note invalide" error="Saisissez une note entre 0 et 20." sqref="{plages}"><formula1>0</formula1><formula2>20</formula2></dataValidation>')
    if matieres:
        cases_coefficients = ' '.join(f'{_col(7+i*2)}{h}' for i in range(len(matieres)))
        validations.append(f'<dataValidation type="decimal" operator="greaterThan" showErrorMessage="1" errorTitle="Coefficient invalide" error="Le coefficient doit etre positif." sqref="{cases_coefficients}"><formula1>0</formula1></dataValidation>')
    validation = f'<dataValidations count="{len(validations)}">{"".join(validations)}</dataValidations>' if validations else ''
    xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{S}" xmlns:r="{R}">'
    xml += f'<dimension ref="A1:{_col(dernier)}{fin}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="5" ySplit="{h}" topLeftCell="F{h+1}" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="F{h+1}" sqref="F{h+1}"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
    xml += '<cols>' + ''.join(f'<col min="{i}" max="{i}" width="{largeur}" customWidth="1"/>' for i, largeur in enumerate(largeurs, 1)) + '</cols>'
    xml += '<sheetData>' + ''.join(rows) + '</sheetData>'
    xml += f'<autoFilter ref="A{h}:{_col(dernier)}{fin}"/><mergeCells count="{len(modele["fusions"])}">{fusions}</mergeCells>'
    if etudiants:
        xml += (f'<conditionalFormatting sqref="{plages}">'
                f'<cfRule type="cellIs" dxfId="{dxf_alerte}" priority="1" operator="lessThan"><formula>0</formula></cfRule>'
                f'<cfRule type="cellIs" dxfId="{dxf_alerte}" priority="2" operator="greaterThan"><formula>20</formula></cfRule>'
                '</conditionalFormatting>')
    xml += validation
    xml += '<printOptions horizontalCentered="1"/><pageMargins left="0.25" right="0.25" top="0.35" bottom="0.35" header="0.15" footer="0.15"/><pageSetup paperSize="8" orientation="landscape" fitToWidth="1" fitToHeight="0"/><drawing r:id="rId1"/></worksheet>'
    # sheetPr precedes dimension in SpreadsheetML.
    xml = xml.replace('<dimension ', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ', 1)
    return xml


def generer_source_notes(data):
    modele = _modele()
    reste = list(data['niveaux'])
    sections = []
    for code in NIVEAUX:
        niveau = next((n for n in reste if _code(n['code']) == code), None)
        if niveau is not None:
            reste.remove(niveau)
        sections.append(niveau or dict(code=code, intitule='', matieres=[], etudiants=[]))
    sections.extend(reste)
    annees = sorted({e['annee'] for n in sections for e in n['etudiants'] if e['annee']})
    feuilles, relations, definitions, contenus = [], [], [], []
    noms_utilises = set()
    resultat = io.BytesIO()
    styles, dxf_alerte = _styles_avec_alerte()
    with zipfile.ZipFile(resultat, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('xl/styles.xml', styles)
        archive.writestr('xl/theme/theme1.xml', (DOSSIER / 'theme.xml').read_bytes())
        for i in range(1, 6):
            archive.writestr(f'xl/media/image{i}.png', (DOSSIER / f'image{i}.png').read_bytes())
        for i, niveau in enumerate(sections, 1):
            code = _code(niveau['code'])
            modele_feuille = modele['niveaux'].get(code, modele['niveaux']['L3'])
            base = re.sub(r'[\[\]:*?/\\]', '_', code).strip("'")[:31] or 'Niveau'
            nom, compteur = base, 1
            while nom.casefold() in noms_utilises:
                compteur += 1
                suffixe = f'_{compteur}'
                nom = base[:31-len(suffixe)] + suffixe
            noms_utilises.add(nom.casefold())
            feuilles.append(f'<sheet name={quoteattr(nom)} sheetId="{i}" r:id="rId{i}"/>')
            relations.append(f'<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>')
            titre = "'" + nom.replace("'", "''") + "'!$1:$" + str(modele_feuille['entete'])
            definitions.append(f'<definedName name="_xlnm.Print_Titles" localSheetId="{i-1}">{_texte(titre)}</definedName>')
            contenus.extend([f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
                             f'<Override PartName="/xl/drawings/drawing{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'])
            archive.writestr(f'xl/worksheets/sheet{i}.xml', _feuille(niveau, data['bassin'], annees, modele_feuille, dxf_alerte))
            archive.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels', f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/drawing" Target="../drawings/drawing{i}.xml"/></Relationships>')
            dessin = modele_feuille['dessin']
            archive.writestr(f'xl/drawings/drawing{i}.xml', (DOSSIER / f'drawing{dessin}.xml').read_bytes())
            archive.writestr(f'xl/drawings/_rels/drawing{i}.xml.rels', f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/image" Target="../media/image{dessin}.png"/></Relationships>')
        total = len(sections)
        relations.extend([f'<Relationship Id="rId{total+1}" Type="{R}/styles" Target="styles.xml"/>', f'<Relationship Id="rId{total+2}" Type="{R}/theme" Target="theme/theme1.xml"/>'])
        archive.writestr('xl/workbook.xml', f'<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>{"".join(feuilles)}</sheets><definedNames>{"".join(definitions)}</definedNames><calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>')
        archive.writestr('xl/_rels/workbook.xml.rels', f'<Relationships xmlns="{P}">{"".join(relations)}</Relationships>')
        archive.writestr('_rels/.rels', f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' + ''.join(contenus) + '</Types>')
    return resultat.getvalue()
