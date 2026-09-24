#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Backend de génération de diplômes IFEM.

Expose une API HTTP (FastAPI) qui remplit les templates Word (.docx) fournis
(variables entourées de {{}}) avec les données d'un(e) candidat(e) et renvoie
directement le fichier .docx généré en téléchargement.

Lancement local :
    uvicorn main:app --reload

Déploiement : voir README.md (Render).
"""
import io
import os
import re
import unicodedata
import zipfile
from typing import Dict, List

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field, field_validator

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# Pour chaque niveau : fichier template + ordre des variables {{}} telles
# qu'elles apparaissent DANS LE DOCUMENT (déterminé une fois par inspection
# du template — voir README.md "Ajouter un nouveau template"). Le champ
# "nom" apparaît deux fois dans les templates L2/L3/M1/M2 (signature en
# filigrane + ligne "Délivré à").
LEVELS: Dict[str, dict] = {
    "L1": dict(file="template_L1.docx", order=["matricule", "nom", "mention", "date", "lieu"]),
    "L2": dict(file="template_L2.docx", order=["matricule", "nom", "nom", "mention", "date", "lieu"]),
    "L3": dict(file="template_L3.docx", order=["matricule", "nom", "nom", "mention", "date", "lieu"]),
    "M1": dict(file="template_M1.docx", order=["matricule", "nom", "nom", "mention", "date", "lieu"]),
    "M2": dict(file="template_M2.docx", order=["matricule", "nom", "nom", "mention", "date", "lieu"]),
}

# Alias acceptés pour "niveau" côté front (les intitulés de la feuille de
# notes source ne correspondent pas toujours au nom du template).
NIVEAU_ALIASES: Dict[str, str] = {
    "L1": "L1", "DT3": "L1", "DT-SE": "L1",
    "L2": "L2", "DTS": "L2",
    "L3": "L3",
    "M1": "M1",
    "M2": "M2",
}

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")

# Origine(s) autorisée(s) pour les appels CORS depuis le front. En prod,
# définir la variable d'environnement ALLOWED_ORIGINS (séparées par des
# virgules) plutôt que de laisser "*".
ALLOWED_ORIGINS = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",") if o.strip()]

# --------------------------------------------------------------------------
# Chargement des templates (une fois, au démarrage)
# --------------------------------------------------------------------------

_template_cache: Dict[str, dict] = {}


def _load_templates() -> None:
    for code, cfg in LEVELS.items():
        path = os.path.join(TEMPLATES_DIR, cfg["file"])
        if not os.path.isfile(path):
            raise RuntimeError(f"Template introuvable : {path}")
        with open(path, "rb") as f:
            raw = f.read()
        zf = zipfile.ZipFile(io.BytesIO(raw))
        xml = zf.read("word/document.xml").decode("utf-8")
        found = len(re.findall(r"\{\{.*?\}\}", xml, re.DOTALL))
        expected = len(cfg["order"])
        if found != expected:
            raise RuntimeError(
                f"Template {code} ({cfg['file']}) : {found} variable(s) {{}} trouvée(s) "
                f"dans le document, {expected} attendue(s). Le template a peut-être été "
                f"modifié après édition dans Word (les runs XML ont pu se refragmenter) — "
                f"voir README.md « Ajouter / mettre à jour un template »."
            )
        _template_cache[code] = {
            "raw": raw,
            "xml": xml,
            "names": [i for i in zf.namelist()],
            "infolist": zf.infolist(),
        }


# --------------------------------------------------------------------------
# Remplissage du template
# --------------------------------------------------------------------------

def _xml_escape(s: str) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _fill_xml(xml: str, values: List[str]) -> str:
    """Remplace, dans l'ordre d'apparition, chaque {{...}} par la valeur
    correspondante — en conservant le formatage (gras, couleur, taille...)
    porté par le run qui contenait la valeur d'origine."""
    pattern = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)
    matches = list(pattern.finditer(xml))
    if len(matches) != len(values):
        raise ValueError(f"{len(matches)} variable(s) trouvée(s), {len(values)} valeur(s) fournie(s)")

    out = []
    last_end = 0
    for m, new_val in zip(matches, values):
        out.append(xml[last_end:m.start()])
        inner = m.group(1)
        last_gt = inner.rfind(">")
        if last_gt == -1:
            out.append(_xml_escape(new_val))
        else:
            out.append(inner[: last_gt + 1] + _xml_escape(new_val))
        last_end = m.end()
    out.append(xml[last_end:])
    return "".join(out)


def _build_docx_bytes(code: str, values: List[str]) -> io.BytesIO:
    cfg = _template_cache[code]
    new_xml = _fill_xml(cfg["xml"], values)
    orig_zip = zipfile.ZipFile(io.BytesIO(cfg["raw"]))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for item in orig_zip.infolist():
            data = orig_zip.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml.encode("utf-8")
            out_zip.writestr(item, data)
    buf.seek(0)
    return buf


def _slugify(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s or "x"


# --------------------------------------------------------------------------
# Schéma de requête
# --------------------------------------------------------------------------

class DiplomeRequest(BaseModel):
    niveau: str = Field(..., description="L1, L2, L3, M1, M2 (ou DT3, DTS)")
    matricule: str = Field(..., min_length=1)
    nom: str = Field(..., min_length=1, description="Nom et prénoms du/de la candidat(e)")
    mention: str = Field(..., min_length=1, description="Ex. Passable, Assez-bien, Bien, Très-bien")
    date_naissance: str = Field(..., description="Format JJ/MM/AAAA")
    lieu_naissance: str = Field(..., min_length=1)
    bassin: str | None = Field(None, description="Optionnel, informatif uniquement (non imprimé)")

    @field_validator("niveau")
    @classmethod
    def _niveau_connu(cls, v: str) -> str:
        code = NIVEAU_ALIASES.get(v.strip().upper())
        if code is None:
            raise ValueError(f"niveau inconnu : {v!r}. Valeurs acceptées : {sorted(NIVEAU_ALIASES)}")
        return code

    @field_validator("date_naissance")
    @classmethod
    def _date_valide(cls, v: str) -> str:
        v = v.strip()
        if not DATE_RE.match(v):
            raise ValueError("date_naissance doit être au format JJ/MM/AAAA")
        return v

    @field_validator("matricule", "nom", "mention", "lieu_naissance")
    @classmethod
    def _non_vide(cls, v: str) -> str:
        v = v.strip()
        if not v:
            raise ValueError("champ vide non autorisé")
        return v


# --------------------------------------------------------------------------
# Application
# --------------------------------------------------------------------------

app = FastAPI(
    title="IFEM - Génération de diplômes",
    description="Génère un diplôme .docx à partir d'un template et des données d'un(e) candidat(e).",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


@app.on_event("startup")
def on_startup() -> None:
    _load_templates()


@app.get("/")
def root():
    return {"status": "ok", "service": "ifem-diplomes-backend", "niveaux": sorted(LEVELS)}


@app.get("/health")
def health():
    return {"status": "ok", "templates_charges": sorted(_template_cache)}


@app.get("/niveaux")
def niveaux():
    """Liste des niveaux disponibles et des champs attendus — utile pour construire
    dynamiquement le formulaire côté front."""
    return {
        "niveaux": sorted(LEVELS),
        "alias_acceptes": NIVEAU_ALIASES,
        "champs_requis": ["niveau", "matricule", "nom", "mention", "date_naissance", "lieu_naissance"],
    }


@app.post("/generate-diplome")
def generate_diplome(req: DiplomeRequest):
    code = NIVEAU_ALIASES[req.niveau.strip().upper()]
    if code not in _template_cache:
        raise HTTPException(500, f"Template non chargé pour le niveau {code}")

    values_by_field = {
        "matricule": req.matricule,
        "nom": req.nom,
        "mention": req.mention,
        "date": req.date_naissance,
        "lieu": req.lieu_naissance,
    }
    order = LEVELS[code]["order"]
    values = [values_by_field[k] for k in order]

    try:
        buf = _build_docx_bytes(code, values)
    except ValueError as e:
        raise HTTPException(500, str(e))

    filename = f"diplome_{code}_{_slugify(req.matricule)}_{_slugify(req.nom)}.docx"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Access-Control-Expose-Headers": "Content-Disposition",
    }
    return StreamingResponse(buf, media_type=DOCX_MIME, headers=headers)


@app.exception_handler(Exception)
async def unhandled_exception_handler(request, exc):
    # Évite de renvoyer une trace brute au front ; log côté serveur (stdout,
    # visible dans les logs Render) et réponse JSON propre.
    print(f"[ERREUR] {request.method} {request.url.path} -> {exc!r}")
    return JSONResponse(status_code=500, content={"detail": "Erreur interne du serveur."})
