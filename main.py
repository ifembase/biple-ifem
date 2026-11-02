#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Backend de génération de documents IFEM : diplômes, authentifications et
relevés de notes.

Expose une API HTTP (FastAPI) qui remplit les templates Word (.docx) fournis
(variables entourées de {{}}) avec les données d'un(e) candidat(e) et renvoie
directement le fichier .docx généré en téléchargement (ou une archive .zip
pour les générations en lot).

Le moteur de remplissage est dans docgen.py.

Lancement local :
    uvicorn main:app --reload

Déploiement : voir README.md (Render).
"""
import io
import os
import re
import zipfile
from typing import Dict, List, Optional, Tuple, Union

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field, field_validator, model_validator

import docgen

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# Alias acceptés pour "niveau" côté front (les intitulés de la feuille de
# notes source ne correspondent pas toujours au nom du template).
NIVEAU_ALIASES: Dict[str, str] = {
    "L1": "L1", "DT3": "L1", "DT-SE": "L1",
    "L2": "L2", "DTS": "L2", "DTS-SE": "L2", "DTS – SE": "L2",
    "L3": "L3",
    "M1": "M1",
    "M2": "M2",
}

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
# Date approximative (personnes dont seul l'âge approximatif est connu) :
# « vers 1990 », « Vers 1990 », « vers l'année 1990 » ou simplement « 1990 ».
# Normalisée en « vers AAAA » ; le modèle imprime alors « Né(e) vers 1990 »
# (le mot « le » du modèle est retiré par docgen).
DATE_APPROX_RE = re.compile(r"^(?:vers\s+(?:l['’]\s*ann[ée]e\s+)?)?(\d{4})$", re.IGNORECASE)

# Origine(s) autorisée(s) pour les appels CORS depuis le front. En prod,
# définir la variable d'environnement ALLOWED_ORIGINS (séparées par des
# virgules) plutôt que de laisser "*".
ALLOWED_ORIGINS = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",") if o.strip()]


_slugify = docgen.slugify


def _normaliser_niveau(v: str) -> str:
    code = NIVEAU_ALIASES.get(v.strip().upper())
    if code is None:
        raise ValueError(f"niveau inconnu : {v!r}. Valeurs acceptées : {sorted(NIVEAU_ALIASES)}")
    return code


# --------------------------------------------------------------------------
# Schémas de requête
# --------------------------------------------------------------------------

class DiplomeRequest(BaseModel):
    """Champs communs au diplôme et à l'authentification."""
    niveau: str = Field(..., description="L1, L2, L3, M1, M2 (ou DT3, DTS)")
    matricule: str = Field(..., min_length=1)
    nom: str = Field(..., min_length=1, description="Nom et prénoms du/de la candidat(e)")
    mention: str = Field(..., min_length=1, description="Ex. Passable, Assez-bien, Bien, Très-bien")
    date_naissance: str = Field(..., description="JJ/MM/AAAA, ou date approximative « vers AAAA » (ou « AAAA »)")
    lieu_naissance: str = Field(..., min_length=1)
    bassin: Optional[str] = Field(None, description="Nom du bassin ; nomme le .zip d'un lot (non imprimé)")
    # Niveau tel qu'envoyé par le front (« DT3 », « DTS »...), avant normalisation
    # en L1..M2 : sert à nommer le dossier du niveau dans le .zip d'un lot.
    niveau_libelle: Optional[str] = Field(None, description="Rempli automatiquement depuis 'niveau'")

    @model_validator(mode="before")
    @classmethod
    def _garder_libelle_niveau(cls, data):
        if isinstance(data, dict) and not data.get("niveau_libelle") and isinstance(data.get("niveau"), str):
            data = {**data, "niveau_libelle": data["niveau"].strip()}
        return data

    @field_validator("niveau")
    @classmethod
    def _niveau_connu(cls, v: str) -> str:
        return _normaliser_niveau(v)

    @field_validator("date_naissance")
    @classmethod
    def _date_valide(cls, v: str) -> str:
        v = v.strip()
        if DATE_RE.match(v):
            return v
        m = DATE_APPROX_RE.match(v)
        if m:
            return f"vers {m.group(1)}"
        raise ValueError("date_naissance doit être au format JJ/MM/AAAA, ou « vers AAAA » si la date exacte est inconnue")

    @field_validator("matricule", "nom", "mention", "lieu_naissance")
    @classmethod
    def _non_vide(cls, v: str) -> str:
        v = v.strip()
        if not v:
            raise ValueError("champ vide non autorisé")
        return v


class AuthentificationRequest(DiplomeRequest):
    """Mêmes champs que le diplôme : nom, date/lieu de naissance, matricule, mention."""


Nombre = Union[float, str]  # 15.5, "15.5" ou "15,5" acceptés


class ReleveRequest(BaseModel):
    niveau: str = Field(..., description="L1, L2, L3, M1, M2 (ou DT3, DTS)")
    matricule: str = Field(..., min_length=1)
    nom: str = Field(..., min_length=1, description="Nom et prénoms du/de la candidat(e)")
    mention: str = Field(..., min_length=1)
    notes: List[Nombre] = Field(
        ..., description="Note /20 de chaque module, dans l'ordre du relevé (voir GET /niveaux)"
    )
    note_soutenance: Nombre = Field(..., description="Note de soutenance /20")
    # Optionnels : calculés par le backend s'ils sont absents.
    notes_ponderees: Optional[List[Nombre]] = Field(
        None, description="« Note avec coeff » de chaque module (défaut : note × coefficient)"
    )
    total: Optional[Nombre] = Field(None, description="Défaut : somme des notes pondérées")
    moyenne: Optional[Nombre] = Field(None, description="Défaut : total / somme des coefficients")
    moyenne_generale: Optional[Nombre] = Field(None, description="Défaut : (moyenne + soutenance) / 2")
    bassin: Optional[str] = Field(None, description="Optionnel, informatif uniquement (non imprimé)")

    @field_validator("niveau")
    @classmethod
    def _niveau_connu(cls, v: str) -> str:
        return _normaliser_niveau(v)

    @field_validator("matricule", "nom", "mention")
    @classmethod
    def _non_vide(cls, v: str) -> str:
        v = v.strip()
        if not v:
            raise ValueError("champ vide non autorisé")
        return v


class EtudiantListe(BaseModel):
    matricule: str = Field(..., min_length=1)
    nom: str = Field(..., min_length=1, description="Nom et prénoms")

    @field_validator("matricule", "nom")
    @classmethod
    def _non_vide(cls, v: str) -> str:
        v = v.strip()
        if not v:
            raise ValueError("champ vide non autorisé")
        return v


class ListeRequest(BaseModel):
    """Une liste des admis : au maximum N étudiants (N = capacité du template,
    5 aujourd'hui, voir GET /niveaux). Les lignes non remplies sont imprimées « _ »."""
    niveau: str = Field(..., description="L1, L2, L3, M1, M2 (ou DT3, DTS)")
    etudiants: List[EtudiantListe] = Field(..., min_length=1, description="Les étudiants sélectionnés, dans l'ordre d'impression")

    @field_validator("niveau")
    @classmethod
    def _niveau_connu(cls, v: str) -> str:
        return _normaliser_niveau(v)


class DiplomeBatchRequest(BaseModel):
    """Plusieurs diplômes d'un coup, récupérés dans une seule archive .zip :
    un dossier par niveau, nommée d'après le bassin (en majuscules)."""
    diplomes: List[DiplomeRequest] = Field(..., min_length=1)
    bassin: Optional[str] = Field(None, description="Défaut : le 'bassin' des diplômes")


class AuthentificationBatchRequest(BaseModel):
    authentifications: List[AuthentificationRequest] = Field(..., min_length=1)


class ReleveBatchRequest(BaseModel):
    releves: List[ReleveRequest] = Field(..., min_length=1)


class ListeBatchRequest(BaseModel):
    listes: List[ListeRequest] = Field(..., min_length=1)


# --------------------------------------------------------------------------
# Application
# --------------------------------------------------------------------------

app = FastAPI(
    title="IFEM - Génération de documents",
    description="Génère diplômes, authentifications et relevés de notes (.docx) à partir des templates.",
    version="2.0.0",
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
    docgen.load_all()


@app.get("/")
def root():
    return {
        "status": "ok",
        "service": "ifem-documents-backend",
        "niveaux": docgen.NIVEAUX,
        "documents": ["diplome", "authentification", "releve", "liste"],
    }


@app.get("/health")
def health():
    return {"status": "ok", "templates_charges": docgen.loaded()}


@app.get("/niveaux")
def niveaux():
    """Niveaux disponibles, champs attendus et modules de chaque relevé — utile
    pour construire dynamiquement les formulaires côté front."""
    return {
        "niveaux": docgen.NIVEAUX,
        "alias_acceptes": NIVEAU_ALIASES,
        "champs_requis": {
            "diplome": ["niveau", "matricule", "nom", "mention", "date_naissance", "lieu_naissance"],
            "authentification": ["niveau", "matricule", "nom", "mention", "date_naissance", "lieu_naissance"],
            "releve": ["niveau", "matricule", "nom", "mention", "notes", "note_soutenance"],
            "liste": ["niveau", "etudiants[{matricule, nom}]"],
        },
        "liste_capacite": {code: docgen.liste_capacite(code) for code in docgen.NIVEAUX},
        "releve_modules": {code: docgen.releve_modules(code) for code in docgen.NIVEAUX},
    }


# ---- helpers -------------------------------------------------------------

def _payload(req: BaseModel) -> dict:
    return req.model_dump()


def _generate_one(kind: str, req: BaseModel) -> Tuple[str, bytes]:
    """Génère un document pour UN candidat et renvoie (nom_fichier, contenu_docx)."""
    code = NIVEAU_ALIASES[req.niveau.strip().upper()]
    if kind == "liste":
        first = req.etudiants[0]
        label = f"{req.niveau} / {first.matricule} ({first.nom})"
        stem = f"{kind}_{code}_{_slugify(first.matricule)}_{len(req.etudiants)}etu"
    else:
        label = f"{req.matricule} ({req.nom})"
        stem = f"{kind}_{code}_{_slugify(req.matricule)}_{_slugify(req.nom)}"
    try:
        data = docgen.generate(kind, code, _payload(req))
    except ValueError as e:
        # Données invalides (ex. trop d'étudiants, mauvais nombre de notes) -> 422 explicite
        raise HTTPException(422, f"{label} : {e}")
    return stem + ".docx", data


def _single_response(filename: str, data: bytes) -> StreamingResponse:
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Access-Control-Expose-Headers": "Content-Disposition",
    }
    return StreamingResponse(io.BytesIO(data), media_type=DOCX_MIME, headers=headers)


def _batch_response(kind: str, items: List[BaseModel], zip_name: str, par_niveau: bool = False) -> StreamingResponse:
    entrees = []
    for item in items:
        filename, data = _generate_one(kind, item)
        dossier = None
        if par_niveau:
            # Libellé du front (DT3, DTS...) si connu, sinon le code normalisé.
            libelle = getattr(item, "niveau_libelle", None) or item.niveau
            dossier = _slugify(libelle).upper()
        entrees.append((dossier, filename, data))
    zip_buf = io.BytesIO(docgen.assembler_zip(entrees))
    headers = {
        "Content-Disposition": f'attachment; filename="{zip_name}"',
        "Access-Control-Expose-Headers": "Content-Disposition",
    }
    return StreamingResponse(zip_buf, media_type="application/zip", headers=headers)


# ---- diplômes ------------------------------------------------------------

@app.post("/generate-diplome")
def generate_diplome(req: DiplomeRequest):
    filename, data = _generate_one("diplome", req)
    return _single_response(filename, data)


@app.post("/generate-diplomes-batch")
def generate_diplomes_batch(req: DiplomeBatchRequest):
    """Sélection multiple côté front : un diplôme par candidat, dans un .zip.
    Le .zip porte le nom du bassin en majuscules (ex. PORT_BERGE.zip) et contient
    un dossier par niveau (DT3/, DTS/, L3/...) avec les diplômes du niveau."""
    bassin = req.bassin or next((d.bassin for d in req.diplomes if d.bassin and d.bassin.strip()), None)
    return _batch_response("diplome", req.diplomes, docgen.nom_archive(bassin, "DIPLOMES"), par_niveau=True)


# ---- authentifications ---------------------------------------------------

@app.post("/generate-authentification")
def generate_authentification(req: AuthentificationRequest):
    filename, data = _generate_one("authentification", req)
    return _single_response(filename, data)


@app.post("/generate-authentifications-batch")
def generate_authentifications_batch(req: AuthentificationBatchRequest):
    return _batch_response("authentification", req.authentifications, "authentifications.zip")


# ---- relevés de notes ----------------------------------------------------

@app.post("/generate-releve")
def generate_releve(req: ReleveRequest):
    filename, data = _generate_one("releve", req)
    return _single_response(filename, data)


@app.post("/generate-releves-batch")
def generate_releves_batch(req: ReleveBatchRequest):
    return _batch_response("releve", req.releves, "releves.zip")


# ---- listes des admis ----------------------------------------------------

@app.post("/generate-liste")
def generate_liste(req: ListeRequest):
    """Une liste par appel : au plus N étudiants (422 au-delà), « _ » pour les
    lignes restantes si moins de N."""
    filename, data = _generate_one("liste", req)
    return _single_response(filename, data)


@app.post("/generate-listes-batch")
def generate_listes_batch(req: ListeBatchRequest):
    return _batch_response("liste", req.listes, "listes.zip")


@app.exception_handler(Exception)
async def unhandled_exception_handler(request, exc):
    # Évite de renvoyer une trace brute au front ; log côté serveur (stdout,
    # visible dans les logs Render) et réponse JSON propre.
    print(f"[ERREUR] {request.method} {request.url.path} -> {exc!r}")
    return JSONResponse(status_code=500, content={"detail": "Erreur interne du serveur."})
