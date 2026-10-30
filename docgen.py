#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Moteur de génération des documents IFEM (diplômes, authentifications,
relevés de notes) à partir de templates Word (.docx).

Principe
--------
Les templates contiennent des marqueurs {{...}} dont le contenu est un
*exemple* de valeur (ex. {{TNR 192}}). Les marqueurs sont remplacés dans
l'ordre d'apparition dans le document, par les valeurs de la requête.

Robustesse : Word découpe souvent un même texte en plusieurs "runs" <w:r>
(ex. "{{" dans un run, "TNR 192" dans un autre, "}}" dans un troisième).
Le moteur ne travaille donc PAS sur le XML brut mais sur la suite des textes
<w:t> concaténés : un marqueur est retrouvé même s'il est réparti sur
n'importe quel nombre de runs, l'exemple est toujours supprimé, et la valeur
est écrite dans le run qui portait l'exemple (donc avec son formatage : gras,
taille, couleur...).

Ce module n'a aucune dépendance externe (bibliothèque standard uniquement).
"""
import io
import os
import re
import zipfile
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from html import unescape
from typing import Dict, List, Optional, Tuple

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

NIVEAUX = ["L1", "L2", "L3", "M1", "M2"]

# --------------------------------------------------------------------------
# Configuration des documents
# --------------------------------------------------------------------------
# Pour chaque type de document et chaque niveau : (fichier, ordre des
# variables {{}} TELLES QU'ELLES APPARAISSENT DANS LE DOCUMENT). Les noms sont
# ceux des champs de la requête. Le nom apparaît deux fois sur les diplômes
# L2/L3/M1/M2 (signature en filigrane + ligne « Délivré à »).

_ORDRE_DIPLOME_L1 = ["matricule", "nom", "mention", "date_naissance", "lieu_naissance"]
_ORDRE_DIPLOME_L2_M2 = ["matricule", "nom", "nom", "mention", "date_naissance", "lieu_naissance"]

DIPLOME: Dict[str, Tuple[str, List[str]]] = {
    "L1": ("template_L1.docx", _ORDRE_DIPLOME_L1),
    "L2": ("template_L2.docx", _ORDRE_DIPLOME_L2_M2),
    "L3": ("template_L3.docx", _ORDRE_DIPLOME_L2_M2),
    "M1": ("template_M1.docx", _ORDRE_DIPLOME_L2_M2),
    "M2": ("template_M2.docx", _ORDRE_DIPLOME_L2_M2),
}

# Authentification : « ... que Madame/Monsieur : {{nom}} né(e) le {{date}} à
# {{lieu}}, ... numéro matricule : {{matricule}}/26 avec la mention {{mention}} »
_ORDRE_AUTH = ["nom", "date_naissance", "lieu_naissance", "matricule", "mention"]

AUTHENTIFICATION: Dict[str, Tuple[str, List[str]]] = {
    code: (os.path.join("auth", f"Authentification_{code}.docx"), _ORDRE_AUTH)
    for code in NIVEAUX
}

# Relevé de notes : l'ordre est construit à partir du template (nombre de
# modules lu dans le tableau) : matricule, nom, [note, note×coeff] × modules,
# TOTAL, MOYENNE, note de soutenance, moyenne générale, mention.
# Style des nombres imprimés, reprenant celui des templates fournis :
# (séparateur décimal, supprimer les zéros inutiles ?).
RELEVE: Dict[str, Tuple[str, Tuple[str, bool]]] = {
    "L1": (os.path.join("releve", "Releve_L1.docx"), (".", False)),   # 15.50
    "L2": (os.path.join("releve", "Releve_L2.docx"), (",", True)),    # 11,5
    "L3": (os.path.join("releve", "Releve_L3.docx"), (".", False)),
    "M1": (os.path.join("releve", "Releve_M1.docx"), (".", False)),
    "M2": (os.path.join("releve", "Releve_M2.docx"), (",", True)),
}


# Liste des admis : un fichier par niveau. Le nombre de lignes (capacité) est
# détecté dans le tableau du template (5 aujourd'hui). Voir _prepare_liste().
LISTE: Dict[str, str] = {code: os.path.join("liste", f"Liste_{code}.docx") for code in NIVEAUX}

LISTE_VIDE = "_"            # remplace matricule et nom d'une ligne sans étudiant
LISTE_OBS_PLEIN = "Admis"   # colonne « Observation » d'une ligne remplie
LISTE_OBS_VIDE = "-------"  # colonne « Observation » d'une ligne vide

# --------------------------------------------------------------------------
# Moteur de remplissage
# --------------------------------------------------------------------------

_T_RE = re.compile(r"<w:t((?:\s[^>]*)?)>(.*?)</w:t>", re.S)
_VAR_RE = re.compile(r"\{\{(.*?)\}\}", re.S)


def _xml_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _scan(xml: str):
    items = list(_T_RE.finditer(xml))
    texts = [unescape(m.group(2)) for m in items]
    return items, texts


def count_variables(xml: str) -> int:
    """Nombre de marqueurs {{...}} du document, même coupés en plusieurs runs."""
    _, texts = _scan(xml)
    return len(_VAR_RE.findall("".join(texts)))


def _retirer_le(new: List[str], first: int, a_first: int) -> None:
    """Retire le mot « le » (et ses espaces) juste avant la variable, pour
    transformer « Né(e) le {{date}} » en « Né(e) vers 1990 » quand la date est
    approximative. Le texte précédant la variable peut être réparti sur
    plusieurs runs ; `new` est modifié sur place."""
    j = first
    head, tail = new[j][:a_first], new[j][a_first:]
    while head.strip() == "" and j > 0:
        new[j] = head + tail
        j -= 1
        head, tail = new[j], ""
    m = re.search(r"(^|[^\wÀ-ÿ])le\s*$", head, re.IGNORECASE)
    if m:
        head = head[:m.start()] + m.group(1)
        if not head.endswith(" "):
            head += " "
    new[j] = head + tail


def fill_xml(xml: str, values: List[str], sans_le: Optional[set] = None) -> str:
    """Remplace, dans l'ordre d'apparition, chaque {{...}} par la valeur
    correspondante. L'exemple contenu entre les accolades est supprimé."""
    items, texts = _scan(xml)
    full = "".join(texts)
    spans = list(_VAR_RE.finditer(full))
    if len(spans) != len(values):
        raise ValueError(f"{len(spans)} variable(s) trouvée(s), {len(values)} valeur(s) fournie(s)")

    starts, pos = [], 0
    for t in texts:
        starts.append(pos)
        pos += len(t)
    ends = [s + len(t) for s, t in zip(starts, texts)]

    new = list(texts)
    # On traite de la dernière variable à la première : les modifications
    # ne décalent ainsi jamais les positions des variables restantes.
    for idx_var, (m, val) in reversed(list(enumerate(zip(spans, values)))):
        s, e = m.start(), m.end()
        inner_s, inner_e = s + 2, e - 2
        overlap = [i for i in range(len(texts)) if starts[i] < e and ends[i] > s]
        # Hôte = premier run qui porte réellement l'exemple (pas seulement
        # "{{" ou des espaces) : la valeur hérite ainsi de son formatage.
        host = None
        for i in overlap:
            a, b = max(starts[i], inner_s), min(ends[i], inner_e)
            if a < b and full[a:b].strip():
                host = i
                break
        if host is None:
            host = overlap[0]
        for i in overlap:
            a = max(starts[i], s) - starts[i]
            b = min(ends[i], e) - starts[i]
            ins = val if i == host else ""
            new[i] = new[i][:a] + ins + new[i][b:]
        if sans_le and idx_var in sans_le:
            first = overlap[0]
            _retirer_le(new, first, max(starts[first], s) - starts[first])

    if _VAR_RE.search("".join(new)) or "{{" in "".join(new) or "}}" in "".join(new):
        raise ValueError("accolades {{ }} résiduelles après remplissage")

    out, last = [], 0
    for m, old, nt in zip(items, texts, new):
        out.append(xml[last:m.start()])
        if nt == old:
            out.append(m.group(0))
        else:
            attrs = m.group(1)
            if nt != nt.strip() and "xml:space" not in attrs:
                attrs += ' xml:space="preserve"'
            out.append(f"<w:t{attrs}>{_xml_escape(nt)}</w:t>")
        last = m.end()
    out.append(xml[last:])
    return "".join(out)



# --------------------------------------------------------------------------
# Préparation des listes
# --------------------------------------------------------------------------

_TR_RE = re.compile(r"<w:tr[ >].*?</w:tr>", re.S)
_TC_RE = re.compile(r"<w:tc[ >].*?</w:tc>", re.S)


def _cell_text(tc: str) -> str:
    return "".join(unescape(m.group(2)) for m in _T_RE.finditer(tc)).strip()


def _set_cell_marker(tc: str, marker: str) -> Optional[str]:
    """Remplace le texte de la cellule par `marker` en gardant le formatage du
    premier run non vide. Renvoie None si la cellule n'a aucun texte."""
    state = {"done": False}

    def repl(m):
        if not unescape(m.group(2)):
            return m.group(0)
        if not state["done"] and unescape(m.group(2)).strip():
            state["done"] = True
            return f"<w:t{m.group(1)}>{marker}</w:t>"
        return f"<w:t{m.group(1)}></w:t>"

    new = _T_RE.sub(repl, tc)
    return new if state["done"] else None


def _prepare_liste(xml: str) -> Tuple[str, List[Tuple[int, str]]]:
    """Normalise un template de liste : chaque ligne d'étudiant reçoit les
    marqueurs {{nb}}, {{matricule}}, {{nom}}, {{obs}}.

    Une « ligne d'étudiant » est une ligne du tableau dont les colonnes 2 et 3
    sont soit un marqueur {{...}}, soit un « _ » écrit en dur (ligne vide déjà
    présente dans le template). Cela rend le moteur indifférent à l'état du
    template (ligne vide déjà en « _ », numérotation qui ne commence pas à 1,
    « Admis » écrit en dur...). Renvoie (xml, [(n° de ligne, champ), ...]) dans
    l'ordre d'apparition."""
    fields: List[Tuple[int, str]] = []
    slot = {"n": 0}

    def is_slot_cell(t: str) -> bool:
        return t == LISTE_VIDE or (t.startswith("{{") and t.endswith("}}"))

    def fix_row(m):
        tr = m.group(0)
        tcs = _TC_RE.findall(tr)
        if len(tcs) < 4:
            return tr
        texts = [_cell_text(tc) for tc in tcs]
        if not (is_slot_cell(texts[1]) and is_slot_cell(texts[2])):
            return tr
        i = slot["n"]
        slot["n"] += 1
        idx = {"n": -1}

        def fix_cell(cm):
            idx["n"] += 1
            k, tc = idx["n"], cm.group(0)
            if k == 0:
                field, marker = "nb", "{{nb}}"
            elif k in (1, 2):
                field = "matricule" if k == 1 else "nom"
                if texts[k] == LISTE_VIDE:
                    marker = "{{" + field + "}}"
                else:  # déjà un marqueur : on le conserve
                    fields.append((i, field))
                    return tc
            elif k == len(tcs) - 1:
                field, marker = "obs", "{{obs}}"
            else:
                return tc
            new = _set_cell_marker(tc, marker)
            if new is None:
                return tc
            fields.append((i, field))
            return new

        return _TC_RE.sub(fix_cell, tr)

    return _TR_RE.sub(fix_row, xml), fields

# --------------------------------------------------------------------------
# Templates chargés en mémoire
# --------------------------------------------------------------------------

class Template:
    def __init__(self, kind: str, code: str, relpath: str, order: Optional[List[str]] = None):
        self.kind, self.code = kind, code
        path = os.path.join(TEMPLATES_DIR, relpath)
        if not os.path.isfile(path):
            raise RuntimeError(f"Template introuvable : {path}")
        with open(path, "rb") as f:
            self.raw = f.read()
        zf = zipfile.ZipFile(io.BytesIO(self.raw))
        self.xml = zf.read("word/document.xml").decode("utf-8")
        self.relpath = relpath
        self.modules: List[Tuple[str, Decimal]] = []
        self.capacity = 0  # nombre de lignes d'une liste

        if kind == "liste":
            self.xml, self.liste_fields = _prepare_liste(self.xml)
            self.capacity = len({i for i, _ in self.liste_fields})
            if self.capacity == 0:
                raise RuntimeError(f"Template {relpath} : aucune ligne d'étudiant trouvée dans le tableau.")
            self.order = [f"{f}{i}" for i, f in self.liste_fields]
        elif kind == "releve":
            self.modules = extract_modules(self.xml)
            if not self.modules:
                raise RuntimeError(f"Template {relpath} : tableau des modules introuvable.")
            n = len(self.modules)
            self.order = ["matricule", "nom"] + [x for i in range(n) for x in (f"note{i}", f"pond{i}")] + [
                "total", "moyenne", "soutenance", "moyenne_generale", "mention"]
        else:
            self.order = list(order or [])
        self.found = count_variables(self.xml)

        if self.found != len(self.order):
            raise RuntimeError(
                f"Template {kind}/{code} ({relpath}) : {self.found} variable(s) {{}} trouvée(s) "
                f"dans le document, {len(self.order)} attendue(s). Vérifiez le template ou "
                f"l'ordre déclaré dans docgen.py (voir README.md)."
            )

    def render(self, values: List[str], sans_le: Optional[set] = None) -> bytes:
        new_xml = fill_xml(self.xml, values, sans_le)
        src = zipfile.ZipFile(io.BytesIO(self.raw))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "word/document.xml":
                    data = new_xml.encode("utf-8")
                out_zip.writestr(item, data)
        return buf.getvalue()


def extract_modules(xml: str) -> List[Tuple[str, Decimal]]:
    """Lit le tableau du relevé : lignes dont les colonnes NOTE et NOTE AVEC
    COEFF sont des variables. Renvoie [(nom du module, coefficient), ...]."""
    modules = []
    for tr in re.findall(r"<w:tr[ >].*?</w:tr>", xml, re.S):
        cells = []
        for tc in re.findall(r"<w:tc[ >].*?</w:tc>", tr, re.S):
            txt = "".join(unescape(t[1]) for t in _T_RE.findall(tc)).strip()
            cells.append(txt)
        if len(cells) >= 4 and "{{" in cells[2] and "{{" in cells[3]:
            try:
                coef = Decimal(cells[1].replace(",", "."))
            except InvalidOperation:
                raise RuntimeError(f"Coefficient illisible dans le relevé : {cells[1]!r}")
            modules.append((cells[0], coef))
    return modules


_REGISTRY: Dict[Tuple[str, str], Template] = {}


def load_all() -> None:
    """Charge et valide tous les templates (appelé au démarrage)."""
    _REGISTRY.clear()
    for code, (rel, order) in DIPLOME.items():
        _REGISTRY[("diplome", code)] = Template("diplome", code, rel, order)
    for code, (rel, order) in AUTHENTIFICATION.items():
        _REGISTRY[("authentification", code)] = Template("authentification", code, rel, order)
    for code, (rel, _style) in RELEVE.items():
        _REGISTRY[("releve", code)] = Template("releve", code, rel)
    for code, rel in LISTE.items():
        _REGISTRY[("liste", code)] = Template("liste", code, rel)


def loaded() -> Dict[str, List[str]]:
    res: Dict[str, List[str]] = {}
    for (kind, code) in _REGISTRY:
        res.setdefault(kind, []).append(code)
    return {k: sorted(v) for k, v in res.items()}


_DATE_COMPLETE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")


def liste_capacite(code: str) -> int:
    """Nombre maximal d'étudiants d'une liste du niveau donné."""
    return _REGISTRY[("liste", code)].capacity


def releve_modules(code: str) -> List[dict]:
    tpl = _REGISTRY[("releve", code)]
    return [{"module": n, "coefficient": float(c)} for n, c in tpl.modules]


# --------------------------------------------------------------------------
# Nombres (relevés)
# --------------------------------------------------------------------------

def _dec(v, label: str) -> Decimal:
    if isinstance(v, Decimal):
        return v
    s = str(v).strip().replace(",", ".").replace(" ", "")
    try:
        d = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"{label} : nombre invalide ({v!r})")
    if not d.is_finite():
        raise ValueError(f"{label} : nombre invalide ({v!r})")
    return d


def _fmt(d: Decimal, decimal: str, trim: bool) -> str:
    s = f"{d.quantize(Decimal('0.01'), ROUND_HALF_UP):.2f}"
    if trim:
        s = s.rstrip("0").rstrip(".")
    return s.replace(".", decimal)


def _releve_values(tpl: Template, data: dict) -> List[str]:
    decimal, trim = RELEVE[tpl.code][1]
    mods = tpl.modules
    n = len(mods)

    notes_in = data.get("notes")
    if not isinstance(notes_in, (list, tuple)) or len(notes_in) != n:
        raise ValueError(
            f"'notes' doit contenir {n} valeur(s) (une par module, dans l'ordre du relevé "
            f"{tpl.code} : {', '.join(m[0] for m in mods)})"
        )
    notes = [_dec(x, f"notes[{i}]") for i, x in enumerate(notes_in)]
    for i, x in enumerate(notes):
        if not (Decimal(0) <= x <= Decimal(20)):
            raise ValueError(f"notes[{i}] ({mods[i][0]}) : la note doit être comprise entre 0 et 20")

    sout = _dec(data.get("note_soutenance"), "note_soutenance")
    if not (Decimal(0) <= sout <= Decimal(20)):
        raise ValueError("note_soutenance : la note doit être comprise entre 0 et 20")

    pond_in = data.get("notes_ponderees")
    if pond_in is not None:
        if len(pond_in) != n:
            raise ValueError(f"'notes_ponderees' doit contenir {n} valeur(s)")
        pond = [_dec(x, f"notes_ponderees[{i}]") for i, x in enumerate(pond_in)]
    else:
        pond = [x * c for x, (_, c) in zip(notes, mods)]

    total = _dec(data["total"], "total") if data.get("total") not in (None, "") else sum(pond, Decimal(0))
    somme_coef = sum((c for _, c in mods), Decimal(0))
    moyenne = _dec(data["moyenne"], "moyenne") if data.get("moyenne") not in (None, "") else total / somme_coef
    mg = (
        _dec(data["moyenne_generale"], "moyenne_generale")
        if data.get("moyenne_generale") not in (None, "")
        else (moyenne + sout) / 2
    )

    f = lambda d: _fmt(d, decimal, trim)
    vals = {"matricule": data["matricule"], "nom": data["nom"], "mention": data["mention"],
            "total": f(total), "moyenne": f(moyenne), "soutenance": f(sout), "moyenne_generale": f(mg)}
    for i in range(n):
        vals[f"note{i}"] = f(notes[i])
        vals[f"pond{i}"] = f(pond[i])
    return [vals[k] for k in tpl.order]


def _liste_values(tpl: Template, data: dict) -> List[str]:
    etudiants = data.get("etudiants")
    if not isinstance(etudiants, (list, tuple)) or len(etudiants) == 0:
        raise ValueError("'etudiants' doit contenir au moins un étudiant")
    if len(etudiants) > tpl.capacity:
        raise ValueError(
            f"une liste {tpl.code} contient au maximum {tpl.capacity} étudiant(s) : "
            f"{len(etudiants)} reçus. Sélectionnez {tpl.capacity} étudiant(s) ou moins."
        )
    vals = {}
    for i in range(tpl.capacity):
        if i < len(etudiants):
            e = etudiants[i]
            vals[f"matricule{i}"] = str(e["matricule"]).strip() or LISTE_VIDE
            vals[f"nom{i}"] = str(e["nom"]).strip() or LISTE_VIDE
            vals[f"obs{i}"] = LISTE_OBS_PLEIN
        else:
            vals[f"matricule{i}"] = LISTE_VIDE
            vals[f"nom{i}"] = LISTE_VIDE
            vals[f"obs{i}"] = LISTE_OBS_VIDE
        vals[f"nb{i}"] = str(i + 1)
    return [vals[k] for k in tpl.order]


# --------------------------------------------------------------------------
# API publique
# --------------------------------------------------------------------------

def generate(kind: str, code: str, data: dict) -> bytes:
    """Génère un document .docx. `kind` : diplome | authentification | releve | liste.
    `data` : dictionnaire des champs (voir les schémas de requête dans main.py).
    Lève ValueError si les données sont invalides."""
    tpl = _REGISTRY.get((kind, code))
    if tpl is None:
        raise ValueError(f"template non chargé : {kind}/{code}")
    if kind == "releve":
        values = _releve_values(tpl, data)
    elif kind == "liste":
        values = _liste_values(tpl, data)
    else:
        values = [str(data[k]).strip() for k in tpl.order]
    # Date de naissance approximative (« vers 1990 ») : « Né(e) le » devient « Né(e) ».
    sans_le = {
        i for i, k in enumerate(tpl.order)
        if k == "date_naissance" and not _DATE_COMPLETE_RE.match(values[i])
    }
    return tpl.render(values, sans_le)
