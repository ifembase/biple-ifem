# IFEM – Backend de génération de diplômes

API HTTP qui remplit vos templates Word (`template_L1.docx` … `template_M2.docx`,
variables entourées de `{{}}`) avec les données d'un(e) candidat(e) et renvoie
directement le fichier `.docx` généré, prêt à être téléchargé depuis votre front
(liste des étudiants des 58 bassins).

## Structure du projet

```
ifem-diplomes-backend/
├── main.py                # l'API (FastAPI)
├── templates/              # vos 5 templates Word (déjà « nettoyés » — voir plus bas)
│   ├── template_L1.docx
│   ├── template_L2.docx
│   ├── template_L3.docx
│   ├── template_M1.docx
│   └── template_M2.docx
├── tools/
│   └── merge_runs.py       # utilitaire de maintenance (voir "Modifier un template")
├── requirements.txt
├── render.yaml              # blueprint de déploiement Render
├── Procfile                 # filet de sécurité si render.yaml n'est pas utilisé
└── README.md
```

Les templates fournis ici sont déjà prétraités (runs Word fusionnés) pour que le
détecteur de variables `{{}}` fonctionne de façon fiable — voir la section
« Modifier ou ajouter un template » avant de les remplacer.

## 1. Lancer en local

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload
```

Puis testez :

```bash
curl http://127.0.0.1:8000/health
```

## 2. Déployer sur Render

Render déploie à partir d'un dépôt Git (GitHub/GitLab/Bitbucket). Ce dossier
n'est pas encore un dépôt — étapes :

1. **Créer le dépôt et pousser le code**
   ```bash
   cd ifem-diplomes-backend
   git init
   git add .
   git commit -m "Backend génération diplômes IFEM"
   git branch -M main
   git remote add origin <URL_DE_VOTRE_DEPOT_GITHUB>
   git push -u origin main
   ```

2. **Sur [render.com](https://render.com)** : New → **Blueprint**, sélectionnez
   votre dépôt. Render lit automatiquement `render.yaml` et propose le service
   `ifem-diplomes-backend` (plan Free, Python). Cliquez **Apply**.

   *(Si vous préférez configurer à la main plutôt que via le blueprint : New →
   Web Service → votre dépôt → Environment: Python 3 → Build Command:
   `pip install -r requirements.txt` → Start Command:
   `uvicorn main:app --host 0.0.0.0 --port $PORT`.)*

3. **Variable d'environnement `ALLOWED_ORIGINS`** : une fois le service créé,
   allez dans Environment et remplacez `*` par le domaine réel de votre front,
   par exemple `https://mon-front-ifem.vercel.app` (plusieurs origines possibles,
   séparées par des virgules). Ne laissez pas `*` en production si l'API doit
   rester privée à votre organisation.

4. Render vous donne une URL du type `https://ifem-diplomes-backend.onrender.com`.
   C'est cette URL que votre front va appeler.

   ⚠️ Sur le plan gratuit, le service s'endort après une période d'inactivité :
   le premier appel après une pause peut prendre 30 à 60 secondes (temps de
   réveil), les suivants sont rapides. Prévoyez un indicateur de chargement
   côté front pour cette latence.

## 3. API

### `GET /niveaux`
Renvoie les niveaux disponibles et les champs attendus — pratique pour
construire dynamiquement un formulaire.

```json
{
  "niveaux": ["L1", "L2", "L3", "M1", "M2"],
  "alias_acceptes": {"L1": "L1", "DT3": "L1", "DTS": "L2", "L3": "L3", "M1": "M1", "M2": "M2"},
  "champs_requis": ["niveau", "matricule", "nom", "mention", "date_naissance", "lieu_naissance"]
}
```

### `POST /generate-diplome`
Génère un diplôme et renvoie le fichier `.docx` en pièce jointe.

**Corps de la requête (JSON) :**

| Champ | Type | Exemple | Remarque |
|---|---|---|---|
| `niveau` | string | `"L1"` | Accepte aussi `DT3`, `DTS` (alias) |
| `matricule` | string | `"AVD 004"` | |
| `nom` | string | `"RAHARISOA Hyama Bakoly Ollivine"` | Nom et prénoms tels qu'à imprimer |
| `mention` | string | `"Bien"` | |
| `date_naissance` | string | `"16/02/1974"` | Format **JJ/MM/AAAA** obligatoire |
| `lieu_naissance` | string | `"Befelatanana ANTANANARIVO IV"` | |
| `bassin` | string | `"AVARADRANO"` | Optionnel, informatif (non imprimé sur le diplôme) |

```bash
curl -X POST https://ifem-diplomes-backend.onrender.com/generate-diplome \
  -H "Content-Type: application/json" \
  -d '{
        "niveau": "L1",
        "matricule": "AVD 004",
        "nom": "RAHARISOA Hyama Bakoly Ollivine",
        "mention": "Passable",
        "date_naissance": "16/02/1974",
        "lieu_naissance": "Befelatanana ANTANANARIVO IV"
      }' \
  --output diplome.docx
```

**Réponses d'erreur** : `422` (champ manquant ou date mal formatée, détail dans
`{"detail": [...]}`), `400`/`500` niveau inconnu ou erreur serveur (détail
générique, la vraie erreur est dans les logs Render).

## 4. Intégration front — un bouton, un téléchargement

Exemple vanilla JS (adaptable à React/Vue/etc.) : le clic appelle l'API et
déclenche le téléchargement du fichier renvoyé.

```html
<button id="btn-generer">Générer le diplôme</button>

<script>
const API_URL = "https://ifem-diplomes-backend.onrender.com";

async function genererDiplome(etudiant) {
  const btn = document.getElementById("btn-generer");
  btn.disabled = true;
  btn.textContent = "Génération...";
  try {
    const res = await fetch(`${API_URL}/generate-diplome`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        niveau: etudiant.niveau,             // ex. "L1", "DTS", "M2"...
        matricule: etudiant.matricule,
        nom: etudiant.nom,
        mention: etudiant.mention,
        date_naissance: etudiant.dateNaissance,   // JJ/MM/AAAA
        lieu_naissance: etudiant.lieuNaissance,
        bassin: etudiant.bassin
      })
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail ? JSON.stringify(err.detail) : `Erreur ${res.status}`);
    }

    // extrait le nom de fichier proposé par le backend
    const disposition = res.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const filename = match ? match[1] : "diplome.docx";

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (e) {
    alert("Échec de la génération : " + e.message);
  } finally {
    btn.disabled = false;
    btn.textContent = "Générer le diplôme";
  }
}

document.getElementById("btn-generer").addEventListener("click", () => {
  genererDiplome({
    niveau: "L1",
    matricule: "AVD 004",
    nom: "RAHARISOA Hyama Bakoly Ollivine",
    mention: "Passable",
    dateNaissance: "16/02/1974",
    lieuNaissance: "Befelatanana ANTANANARIVO IV",
    bassin: "AVARADRANO"
  });
});
</script>
```

## 5. Modifier ou ajouter un template

Le backend détecte les variables en cherchant `{{...}}` dans le XML du
document. Si vous rouvrez un template dans Word et le modifiez (texte,
police...), Word peut re-fragmenter un paragraphe en plusieurs balises
internes, si bien qu'un marqueur `{{...}}` peut ne plus être reconnu comme un
seul bloc.

Après toute édition dans Word :

```bash
python tools/merge_runs.py templates/template_L1.docx
```

Puis relancez le backend (`uvicorn main:app --reload` en local, ou redéployez
sur Render) : au démarrage, le backend vérifie que chaque template contient
bien le nombre de variables attendu et refuse de démarrer sinon, avec un
message précis (nom du template, nombre trouvé vs attendu).

**Pour ajouter un nouveau niveau** (ex. un 6ᵉ template) : déposez le fichier
dans `templates/`, ajoutez une entrée dans le dictionnaire `LEVELS` de
`main.py` avec l'ordre d'apparition de ses variables `{{}}` (matricule, nom
— une ou deux fois selon le template —, mention, date, lieu), et ajoutez
l'alias correspondant dans `NIVEAU_ALIASES` si besoin.

## 6. Limites volontaires de cette version

- Un appel = un diplôme (pas de génération en lot). Si vous voulez un
  endpoint `/generate-diplomes-lot` qui prend une liste d'étudiants et renvoie
  un `.zip`, c'est une extension simple de `generate_diplome` — dites-le et
  on l'ajoute.
- Aucune donnée n'est stockée côté serveur : chaque requête est traitée en
  mémoire et rien n'est écrit sur disque, donc rien à nettoyer ni à sécuriser
  côté stockage. Les données des candidat(e)s ne transitent que dans la
  requête/réponse HTTPS.
- Pas de vérification d'anomalies (note de soutenance manquante, etc.) côté
  backend — c'est à votre front de ne proposer le bouton « Générer » que pour
  les candidat(e)s admis(es) avec un dossier complet, comme dans les rapports
  TXT produits précédemment.
