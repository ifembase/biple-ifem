#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utilitaire de maintenance : à lancer UNIQUEMENT si vous avez modifié un
template .docx dans Word et que le backend renvoie une erreur du type
« N variable(s) {{}} trouvée(s), M attendue(s) » au démarrage.

Pourquoi : quand on tape du texte dans Word (correction orthographique,
retouche de mise en forme...), Word découpe souvent le texte d'un même
paragraphe en plusieurs balises <w:r> (runs) consécutives. Le backend sait
lire un marqueur {{...}} même s'il est coupé en deux runs (un run qui finit
par "{{", un run qui commence par la valeur), mais pas s'il est fragmenté
en trois runs ou plus. Ce script fusionne les runs consécutifs qui ont
exactement le même formatage, ce qui réduit ce risque de fragmentation.

Usage :
    python tools/merge_runs.py templates/template_L1.docx
    (écrase le fichier en place ; une copie de sauvegarde .bak est créée)
"""
import re
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)


def _tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def merge_runs_in_xml(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)

    for para in root.iter(_tag("p")):
        children = list(para)
        merged = []
        for child in children:
            if (
                child.tag == _tag("r")
                and merged
                and merged[-1].tag == _tag("r")
                and _rpr_key(merged[-1]) == _rpr_key(child)
                and _only_text_run(merged[-1])
                and _only_text_run(child)
            ):
                # fusionne le texte de `child` dans le dernier run gardé
                prev_t = merged[-1].find(_tag("t"))
                cur_t = child.find(_tag("t"))
                if prev_t is not None and cur_t is not None:
                    prev_t.text = (prev_t.text or "") + (cur_t.text or "")
                    prev_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    continue  # `child` absorbé, ne pas le garder
            merged.append(child)
            para.remove(child) if False else None  # noop (on reconstruit ci-dessous)

        # reconstruit l'ordre des enfants du paragraphe
        for c in children:
            para.remove(c)
        for c in merged:
            para.append(c)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _rpr_key(run_elem):
    rpr = run_elem.find(_tag("rPr"))
    return ET.tostring(rpr, encoding="unicode") if rpr is not None else ""


def _only_text_run(run_elem) -> bool:
    """True si le run ne contient qu'un <w:t> (pas de <w:br>, <w:tab>,
    <w:drawing>... qu'on ne veut pas fusionner à l'aveugle)."""
    tags = [c.tag for c in run_elem if c.tag != _tag("rPr")]
    return tags == [_tag("t")]


def main(path: str) -> None:
    shutil.copy(path, path + ".bak")
    with zipfile.ZipFile(path, "r") as zin:
        items = {i.filename: zin.read(i.filename) for i in zin.infolist()}
        infos = zin.infolist()

    xml = items["word/document.xml"]
    before = len(re.findall(rb"<w:r>|<w:r ", xml))
    new_xml = merge_runs_in_xml(xml)
    after = len(re.findall(rb"<w:r>|<w:r ", new_xml))
    items["word/document.xml"] = new_xml

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, items[info.filename])

    print(f"{path} : {before} runs -> {after} runs. Sauvegarde : {path}.bak")
    print("Vérifiez le nombre de variables {{}} détectées en relançant le backend "
          "(voir /health ou les logs au démarrage).")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python tools/merge_runs.py <fichier.docx>")
        sys.exit(1)
    main(sys.argv[1])
