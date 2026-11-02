# IFEM – Backend de génération de documents

API HTTP qui remplit vos templates Word (variables entourées de `{{}}`) avec les
données d'un(e) candidat(e) et renvoie directement le fichier `.docx` généré,
prêt à être téléchargé depuis votre front. Trois types de documents, pour
chacun des 5 niveaux (L1, L2, L3, M1, M2) :

| Document | Templates | Endpoint |
|---|---|---|
| Diplôme | `templates/template_<NIVEAU>.docx` | `/generate-diplome` |
| Authentification | `templates/auth/Authentification_<NIVEAU>.docx` | `/generate-authentification` |
| Relevé de notes | `templates/releve/Releve_<NIVEAU>.docx` | `/generate-releve` |
| Liste des admis | `templates/liste/Liste_<NIVEAU>.docx` | `/generate-liste` |


## Structure du projet

```
├── main.py               # l'API (FastAPI) : schémas de requête + routes
├── docgen.py             # le moteur : lecture des templates, remplissage, calcul des relevés
├── templates/            # diplômes, auth/, releve/, liste/
├── tests/test_docgen.py  # tests (stdlib) : python -m unittest discover -s tests -v
├── tools/merge_runs.py   # utilitaire optionnel (voir "Modifier un template")
├── requirements.txt
├── render.yaml           # blueprint de déploiement Render
└── Procfile
```

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

Toutes les routes `generate-*` renvoient un `.docx` en pièce jointe ; les routes
`*-batch` renvoient un `.zip` (un `.docx` par candidat). Le champ `niveau`
accepte `L1`…`M2` et les alias `DT3`, `DT-SE` (→ L1), `DTS`, `DTS-SE` (→ L2).

### `GET /niveaux`
Niveaux, alias, champs requis par type de document et **modules + coefficients
de chaque relevé** (lus directement dans les templates) — pratique pour
construire les formulaires de saisie des notes.

### Diplôme et authentification — mêmes champs

`POST /generate-diplome` · `POST /generate-authentification`
Lots : `POST /generate-diplomes-batch` `{"diplomes": [...]}` ·
`POST /generate-authentifications-batch` `{"authentifications": [...]}`

| Champ | Exemple | Remarque |
|---|---|---|
| `niveau` | `"L2"` | |
| `matricule` | `"AVD 004"` | le « /26 » est déjà dans le template |
| `nom` | `"RAHARISOA Hyama Bakoly Ollivine"` | |
| `mention` | `"Bien"` | |
| `date_naissance` | `"16/02/1974"` ou `"vers 1990"` | **JJ/MM/AAAA**, ou date approximative `vers AAAA` (`1990` et `vers l'année 1990` sont normalisés en `vers 1990`) : le document imprime alors « Né(e) vers 1990 » (le « le » du modèle est retiré) |
| `lieu_naissance` | `"Befelatanana ANTANANARIVO IV"` | |
| `bassin` | `"AVARADRANO"` | optionnel, non imprimé |

#### Archive d'un lot de diplômes

`POST /generate-diplomes-batch` renvoie un `.zip` nommé d'après le **bassin en majuscules**
(`"Port Bergé"` → `PORT_BERGE.zip` ; champ `bassin` du lot, sinon celui des diplômes ;
`DIPLOMES.zip` si aucun bassin) et contenant **un dossier par niveau**, avec le libellé envoyé
par le front :

```
PORT_BERGE.zip
├── DT3/   diplome_L1_….docx
├── DTS/   diplome_L2_….docx
├── L3/    diplome_L3_….docx
└── M1/    diplome_M1_….docx
```

Les lots d'authentifications, de relevés et de listes gardent une archive à plat.

### Relevé de notes

`POST /generate-releve` · lot : `POST /generate-releves-batch` `{"releves": [...]}`

| Champ | Exemple | Remarque |
|---|---|---|
| `niveau`, `matricule`, `nom`, `mention` | | comme ci-dessus |
| `notes` | `[15, 15, "11,5", 12.5, 12, 14]` | **une note /20 par module**, dans l'ordre de `GET /niveaux` (virgule ou point acceptés) |
| `note_soutenance` | `17.5` | /20 |
| `notes_ponderees` | optionnel | « Note avec coeff » par module. Défaut : note × coefficient |
| `total` | optionnel | Défaut : somme des notes pondérées |
| `moyenne` | optionnel | Défaut : total ÷ somme des coefficients |
| `moyenne_generale` | optionnel | Défaut : (moyenne + note de soutenance) ÷ 2 |

Les coefficients viennent du tableau du template. Les calculs se font sur les
valeurs non arrondies puis sont arrondis à 2 décimales ; si votre front a déjà
calculé les valeurs officielles (par ex. note × coeff à partir d'une moyenne
non arrondie), envoyez-les dans `notes_ponderees` / `total` / `moyenne` /
`moyenne_generale` : elles sont imprimées telles quelles. Le style des nombres
(point ou virgule, zéros finaux) suit celui de chaque template — réglable dans
`RELEVE` (`docgen.py`).

```bash
curl -X POST https://ifem-diplomes-backend.onrender.com/generate-releve \
  -H "Content-Type: application/json" \
  -d '{"niveau":"L2","matricule":"ABM 052","nom":"RABARIOELINIRINA Mbola Tina Fiononana",
       "mention":"Bien","notes":[15,15,"11,5","12,5",12,14],"note_soutenance":"17,5"}' \
  --output releve.docx
```

### Liste des admis

`POST /generate-liste` · lot : `POST /generate-listes-batch` `{"listes": [...]}`

```json
{"niveau": "DTS", "etudiants": [{"matricule": "MGA 003", "nom": "RAHARISOA Saida"},
                                {"matricule": "MGA 015", "nom": "RASOAZANANY Manjatiana"}]}
```

Une liste contient **au maximum 5 étudiants** (capacité détectée dans le tableau du
template, visible dans `GET /niveaux` → `liste_capacite`) :

- **5 étudiants** : toutes les lignes sont remplies.
- **moins de 5** : les lignes restantes affichent `_` (matricule et nom) et `-------`
  dans la colonne Observation.
- **plus de 5** (ou aucun) : la génération échoue, `422` avec un message explicite.

Les étudiants sont imprimés dans l'ordre reçu ; la colonne NB est toujours numérotée
1 à 5 ; « Admis » est imprimé pour chaque ligne remplie. Le niveau `DTS` correspond à
`Liste_L2`. Les valeurs `_` / `-------` se règlent en tête de `docgen.py`
(`LISTE_VIDE`, `LISTE_OBS_VIDE`).

**Erreurs** : `422` (champ manquant, date mal formatée, mauvais nombre de notes,
note hors de 0–20…) avec le détail ; `500` erreur serveur (détail dans les logs).

## 4. Intégration front — un bouton, un téléchargement

L'exemple ci-dessous génère un diplôme ; pour une authentification ou un relevé,
changez simplement l'URL (`/generate-authentification`, `/generate-releve`) et le corps de la requête.

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

Dans les templates, le texte entre `{{ }}` est un **exemple** (ex. `{{TNR 192}}`) :
il est remplacé, dans l'ordre d'apparition dans le document, par la valeur de la
requête. Le moteur (`docgen.py`) lit le texte réel du document : un marqueur est
retrouvé même si Word l'a coupé en plusieurs morceaux (`{{` / exemple / `}}`), et
l'exemple est toujours supprimé. La valeur reprend le formatage du texte d'exemple.

Au démarrage, chaque template est vérifié : s'il ne contient pas le nombre de
variables attendu, le backend refuse de démarrer avec un message précis (type de
document, niveau, nombre trouvé vs attendu). Après toute modification d'un
template, lancez `python -m unittest discover -s tests -v`.

- **Diplôme / authentification** : l'ordre des variables est déclaré dans
  `DIPLOME` / `AUTHENTIFICATION` (`docgen.py`). Si vous ajoutez, retirez ou
  déplacez une variable dans le document, mettez cette liste à jour.
- **Liste** : les lignes d'étudiants sont détectées dans le tableau (colonnes Matricule et Nom sous forme de `{{...}}` ou de `_`) ; NB et Observation sont réécrits par le moteur. Ajouter ou retirer des lignes change la capacité automatiquement.
- **Relevé** : le nombre de modules et les coefficients sont lus dans le tableau
  du template ; seules les 5 variables fixes (matricule, nom, TOTAL, MOYENNE,
  soutenance / moyenne générale / mention) doivent rester dans cet ordre.
  Les lignes de modules doivent avoir leurs colonnes NOTE et NOTE AVEC COEFF
  sous forme de `{{...}}`.
- `tools/merge_runs.py` n'est plus nécessaire pour la détection des variables ;
  il reste disponible pour nettoyer un template très fragmenté.

**Nouveau niveau** : déposez le template, ajoutez-le dans les dictionnaires de
`docgen.py` (et `NIVEAUX`), plus l'alias dans `NIVEAU_ALIASES` de `main.py`.

## 6. Limites volontaires

- Aucune donnée n'est stockée côté serveur : chaque requête est traitée en
  mémoire, rien n'est écrit sur disque.
- Pas de vérification d'anomalies métier (admis ou non, note de soutenance
  manquante, mention cohérente avec la moyenne…) : la mention est imprimée telle
  qu'envoyée. C'est au front de ne proposer « Générer » que pour les candidat(e)s
  admis(es) avec un dossier complet.
- Une liste est limitée à la capacité de son template (5) : le découpage d'une sélection plus grande en plusieurs listes doit se faire côté front (un appel par groupe de 5, ou `/generate-listes-batch`).

## Performances

Les templates de diplôme contenaient un aperçu de fichier (`docProps/thumbnail.emf`, 22 Mo
décompressé) inutile pour la génération. Il est retiré en mémoire au chargement (vos fichiers
`.docx` ne sont pas modifiés) : un diplôme passe de ~0,75 s à ~0,04 s de génération et de
4,7 Mo à ~1,6 Mo de téléchargement. Les lots (`*-batch`) sont assemblés sans recompression.
Sur l'offre gratuite de Render, le service s'endort après ~15 min d'inactivité : la première
requête qui suit peut prendre 30 à 60 s (démarrage à froid), indépendamment du code.
