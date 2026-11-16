# Export Word personnalisé des listes IFEM

Depuis la page d’entrée de la **Section Diplôme**, le bouton **Exporter une
liste** ouvre une sélection indépendante des étudiants éligibles au diplôme.

L’utilisateur choisit un ou plusieurs bassins, un ou plusieurs niveaux parmi
DT3, DTS, L3, M1 et M2, les colonnes et les étudiants à inclure. Les trois
organisations proposées sont bassin puis niveau, niveau avec tous les bassins
regroupés, ou liste unique avec niveaux confondus. Le résultat est un seul
fichier `.docx`, paginé automatiquement, sans la limite de cinq étudiants du
parcours historique des listes d’admis.

L’option **Les meilleurs par niveau, tous bassins sélectionnés confondus**
permet de saisir le nombre souhaité (top 3, 10, 20, etc.). Le nombre est un
entier de 1 à 10 000. Le document contient uniquement ce classement global,
sans détail supplémentaire par bassin. Le rang repart de 1 pour chaque niveau.
Si moins de candidats remplissent les filtres, ils sont tous inclus.

Le classement utilise la moyenne générale de la plateforme : moyenne de l’écrit
complet et de la soutenance lorsqu’elle est notée, sinon moyenne de l’écrit
complet. Les égalités sont départagées par nom, prénoms et matricule. Le niveau,
le rang, le bassin et la moyenne générale accompagnent les colonnes choisies.
L’en-tête du modèle reste celui du niveau sélectionné, ou le modèle mixte si
plusieurs niveaux sont sélectionnés. Le classement peut occuper plusieurs pages.

Pour ce mode, l’API reçoit `classement: "top_n_niveau"`, `top_n: 10` (par
exemple), `regroupement: "niveau"` et la `moyenne_generale` de chaque candidat.
L’ancien identifiant `top3_bassin_niveau` reste accepté pour les onglets déjà
ouverts, mais produit uniquement un classement global de trois étudiants par niveau.

Les colonnes disponibles sont matricule, niveau, nom, prénoms, date et lieu de
naissance, note d’écrit, note de soutenance et mention. Une note manquante est
imprimée avec un tiret ; une note égale à zéro reste une vraie note. La note
d’écrit est la moyenne théorique des matières et n’est imprimée que lorsque
toutes les matières sont notées. La mention reprend le résultat affiché dans
la plateforme.

Les filtres distinguent tous les étudiants, soutenance notée, écrit terminé sans
soutenance, écrit terminé avec ou sans soutenance, admis avec écrit et
soutenance terminés, écrit en cours et absence de notes. La plateforme ne
stocke pas de statut distinct « soutenance effectuée » : la présence d’une
note de soutenance, même zéro, est donc utilisée. Un filtre d’année et le texte
de session dans l’en-tête sont également disponibles.

## Modèles et confidentialité

`templates/export/Export_DT3.docx` à `Export_M2.docx` reprennent les cinq modèles
FOP fournis. Lorsqu’un seul niveau est coché, son en-tête est utilisé.
Lorsqu’au moins deux niveaux sont cochés, `Export_MIXTE.docx` reprend l’en-tête
IFEM de la liste des admis déjà présente dans le front. Le titre du document
est « Liste des étudiants », ou « Liste des admis » pour le filtre des admis.

Ces six fichiers sont des copies nettoyées : aucune ligne nominative d’origine,
signature, ancien total ou information d’auteur n’y est conservé. Les logos,
polices et styles utiles restent ceux des références. Le manifeste
`templates/export/provenance.json` conserve les empreintes des références et
des parties préservées. Les modèles sources de l’utilisateur restent inchangés.

L’outil `tools/prepare_export_templates.py` permet de régénérer ces copies à
partir des références originales. Il n’est pas exécuté au démarrage du service.

## API

`POST /export-liste-etudiants` reçoit :

```json
{
  "niveaux": ["DT3"],
  "colonnes": ["matricule", "nom", "prenoms"],
  "critere": "tous",
  "regroupement": "bassin_niveau",
  "session": "Août 2026",
  "etudiants": [{
    "bassin_id": "1", "bassin": "Bassin de démonstration", "niveau": "DT3",
    "matricule": "DEMO-001", "nom": "NOM TEST", "prenoms": "Prénom Test",
    "date_naissance": "vers 1990", "lieu_naissance": "Ville Test",
    "note_ecrit": 12.5, "note_soutenance": null,
    "mention": "En attente", "annee": "2025-2026"
  }]
}
```

Le service valide les niveaux, les colonnes et les notes entre 0 et 20, refuse
les doublons de sélection et accepte jusqu’à 10 000 lignes par requête. La
réponse est un document Word avec un nom de fichier de téléchargement et une
consigne de non-mise en cache.

Comme les autres générations documentaires existantes, cet endpoint transforme
les données reçues ; il ne lit pas la base métier. La sélection des étudiants
provient des inscriptions actuellement chargées par le front, soumises aux
droits d’accès de son utilisateur. Il ne remplace pas une validation officielle
des résultats ni une consultation exhaustive des anciennes inscriptions.

## Validation et déploiement

Les tests couvrent les filtres, le zéro, les colonnes, les regroupements,
les grandes listes, les données invalides, la conservation des logos/styles et
la génération concurrente. La recette navigateur utilise exclusivement des
données fictives et vérifie les téléchargements réels ainsi que le formulaire
sur ordinateur et mobile. Les modèles et les exemples sont aussi rendus dans
Word pour contrôle visuel de la pagination.

Déployer d’abord ce service documentaire, puis le front. Les anciens endpoints
de diplômes, authentifications, relevés et listes restent disponibles. Vérifier
que `VITE_DIPLOME_API_BASE_URL` du front désigne bien ce service déployé.
