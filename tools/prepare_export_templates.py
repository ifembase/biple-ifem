"""Deriver des en-tetes sans embarquer les listes nominatives fournies.

Cet outil est execute manuellement lors d'un changement des modeles.
Le service utilise uniquement les six petits modeles anonymises produits.
"""
import argparse
import hashlib
import json
import re
import zipfile
from html import escape, unescape
from pathlib import Path
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
T = re.compile(r'<w:t((?:\s[^>]*)?)>(.*?)</w:t>', re.S)


def remplacer_texte(xml, expression, remplacement):
    textes = list(T.finditer(xml))
    valeurs = [unescape(m.group(2)) for m in textes]
    joint = ''.join(valeurs)
    debuts, position = [], 0
    for valeur in valeurs:
        debuts.append(position)
        position += len(valeur)
    nouvelles = list(valeurs)
    for correspondance in reversed(list(re.finditer(expression, joint, re.S))):
        indices = [i for i, debut in enumerate(debuts) if debut < correspondance.end() and debut + len(valeurs[i]) > correspondance.start()]
        for i in indices:
            a = max(0, correspondance.start() - debuts[i])
            b = min(len(valeurs[i]), correspondance.end() - debuts[i])
            nouvelles[i] = nouvelles[i][:a] + (remplacement if i == indices[0] else '') + nouvelles[i][b:]
    sortie, dernier = [], 0
    for match, ancien, nouveau in zip(textes, valeurs, nouvelles):
        sortie.append(xml[dernier:match.start()])
        sortie.append(match.group(0) if ancien == nouveau else f'<w:t{match.group(1)}>{escape(nouveau, quote=False)}</w:t>')
        dernier = match.end()
    return ''.join(sortie) + xml[dernier:]


def preparer(source, destination, code):
    with zipfile.ZipFile(source) as archive:
        parties = {info.filename: (info, archive.read(info.filename)) for info in archive.infolist()}
    original = parties['word/document.xml'][1].decode('utf-8')
    debut_table = re.search(r'<w:tbl(?:\s|>)', original).start()
    prefixe = original[:debut_table]
    if code == 'MIXTE':
        prefixe = re.sub(r'<w:p(?:\s[^>]*)?>.*?</w:p>', lambda m: '' if 'IMPORTANT' in ''.join(unescape(t.group(2)) for t in T.finditer(m.group(0))) else m.group(0), prefixe, flags=re.S)
        prefixe = remplacer_texte(prefixe, r'LISTE DES ADMIS AUX .*?MÉMOIRE', '{{TITRE}}')
        if '{{TITRE}}' not in prefixe: raise ValueError('Titre generique non retrouve')
    else:
        prefixe = remplacer_texte(prefixe, r'Août\s+2026', '{{SESSION}}')
        if '{{SESSION}}' not in prefixe: raise ValueError(f'Session non retrouvee : {code}')
    # Retirer les signets qui pourraient se terminer dans le contenu supprime.
    prefixe = re.sub(r'<w:bookmark(?:Start|End)\b[^>]*/>', '', prefixe)
    section = re.findall(r'<w:sectPr\b.*?</w:sectPr>', original, re.S)[-1]
    document = prefixe + '<!--IFEM_EXPORT_CONTENT--><w:p/>' + section + '</w:body></w:document>'
    ET.fromstring(document)
    modifiees = {'word/document.xml': document.encode('utf-8')}
    for nom, (_, contenu) in parties.items():
        if re.match(r'word/header\d+\.xml$', nom) and code != 'MIXTE':
            entete = remplacer_texte(contenu.decode('utf-8'), r'(?:DT-SE|DTS|L3|M1|M2)-IFEM-SESSION-2026', '{{ENTETE_SESSION}}')
            modifiees[nom] = entete.encode('utf-8')
    utilises = set(re.findall(r'\br:(?:id|embed|link)="([^"]+)"', document))
    relations = parties['word/_rels/document.xml.rels'][1].decode('utf-8')
    medias = set()
    def relation(m):
        noeud = ET.fromstring(m.group(0))
        cible, type_rel, identifiant = (noeud.get(k, '') for k in ('Target', 'Type', 'Id'))
        if 'customXml' in type_rel or type_rel.endswith('/comments'): return ''
        if type_rel.endswith('/image'):
            if identifiant not in utilises: return ''
            medias.add('word/' + cible.lstrip('/'))
        return m.group(0)
    relations = re.sub(r'<Relationship\b[^>]*/>', relation, relations)
    modifiees['word/_rels/document.xml.rels'] = relations.encode('utf-8')
    for nom, (_, contenu) in parties.items():
        if re.match(r'word/_rels/(header|footer)\d+\.xml.rels', nom):
            for rel in ET.fromstring(contenu):
                if rel.get('Type', '').endswith('/image'):
                    medias.add('word/' + rel.get('Target', '').lstrip('/'))
    supprimees = {nom for nom in parties if nom.startswith('customXml/') or nom.startswith('docProps/thumbnail')
                  or nom.startswith('word/comments') or (nom.startswith('word/media/') and nom not in medias)}
    types = parties['[Content_Types].xml'][1].decode('utf-8')
    types = re.sub(r'<Override\b[^>]*/>', lambda m: '' if ET.fromstring(m.group(0)).get('PartName', '').lstrip('/') in supprimees else m.group(0), types)
    modifiees['[Content_Types].xml'] = types.encode('utf-8')
    if 'word/settings.xml' in parties:
        settings = parties['word/settings.xml'][1].decode('utf-8')
        settings = re.sub(r'<w:updateFields\b[^>]*/>', '', settings)
        settings = settings.replace('</w:settings>', '<w:updateFields w:val="true"/></w:settings>')
        modifiees['word/settings.xml'] = settings.encode('utf-8')
    modifiees['docProps/core.xml'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>'
    modifiees['docProps/app.xml'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>IFEM HUB</Application></Properties>'
    with zipfile.ZipFile(destination, 'w', compression=zipfile.ZIP_DEFLATED) as sortie:
        for nom, (info, contenu) in parties.items():
            if nom not in supprimees:
                sortie.writestr(info, modifiees.get(nom, contenu))
    # Le modele ne contient aucune ligne nominative et conserve ses logos.
    with zipfile.ZipFile(destination) as resultat:
        arbre = ET.fromstring(resultat.read('word/document.xml'))
        if arbre.findall(f'.//{W}tr'): raise ValueError('Des lignes source subsistent')
        if not any(n.startswith('word/media/') for n in resultat.namelist()): raise ValueError('Logo perdu')
    return {'niveau': code, 'source': source.name, 'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
            'modele': destination.name, 'modele_sha256': hashlib.sha256(destination.read_bytes()).hexdigest(),
            'parties_preservees': {nom: hashlib.sha256(contenu).hexdigest() for nom, (_, contenu) in parties.items() if nom not in supprimees and nom not in modifiees}}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--references-dir', type=Path, required=True)
    parser.add_argument('--mixed-reference', type=Path, required=True)
    parser.add_argument('--output-dir', type=Path, default=Path(__file__).resolve().parents[1] / 'templates/export')
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    manifeste = []
    for code in ('DT3', 'DTS', 'L3', 'M1', 'M2', 'MIXTE'):
        source = args.mixed_reference if code == 'MIXTE' else args.references_dir / f'Liste_Des_Admis_{code}_2025-2026.docx'
        manifeste.append(preparer(source, args.output_dir / f'Export_{code}.docx', code))
    (args.output_dir / 'provenance.json').write_text(json.dumps(manifeste, ensure_ascii=False, indent=2), encoding='utf-8')
    print('6 modeles anonymises prepares ; aucun etudiant source embarque.')


if __name__ == '__main__': main()
