# Trésorerie — La Kasbah Salam

Classeur Google Sheets de trésorerie, avec un script Apps Script qui
synchronise l'onglet Caisse depuis Supabase, et des scripts Python qui
lisent et écrivent les données via l'API Sheets.

## État actuel — lire avant toute proposition

Ces points sont réglés. Ne pas proposer de les refaire.

**La migration v15 → v16 est faite.** Le classeur v16 est la référence
unique. Le script y est rattaché (`clasp create-script --parentId`), poussé
et fonctionnel : menus Caisse et CdR présents, propriétés `URL_FONCTION` et
`SECRET` renseignées, `installerDeclencheurs()` lancé. Le déclencheur de
l'ancien projet v15 a été supprimé — un seul classeur tire sur la file
Supabase.

**La synchronisation reste en Apps Script, et c'est voulu.** Elle tourne
dans le cloud de Google, gratuitement, sans machine à maintenir. Ne pas
proposer de la réécrire en Python : cela obligerait à héberger une
exécution quotidienne à 21h (poste allumé ou serveur), en remplaçant
quelque chose qui fonctionne par quelque chose de plus fragile.

**Le script ne disparaît plus d'une version à l'autre.** C'était vrai
avant clasp. `Code.js` vit désormais en local et se rattache à une
nouvelle Sheet en une commande — c'est précisément le problème que clasp
résout.

**Les deux outils ont des rôles distincts, ils ne se remplacent pas :**

- **Apps Script (clasp)** — la synchronisation automatique depuis
  Supabase, les menus du classeur, les formules du compte de résultat.
- **Python (API Sheets)** — la manipulation des données : corrections en
  masse, analyses, vérifications de soldes. L'accès est en place et
  testé ; les scripts s'écrivent au fil des besoins.

## Deux domaines, ne pas les mélanger

```
Tresorerie\
├── Script\          → le code Apps Script (clasp)
│   ├── .clasp.json  → lie ce dossier à la Sheet, ne pas modifier
│   ├── Code.js      → le programme
│   └── appsscript.json
└── data\            → les scripts Python (API Sheets)
    ├── credentials.json      ← SECRET, ne jamais lire ni afficher
    ├── authorized_user.json  ← SECRET, ne jamais lire ni afficher
    └── test_connexion.py     ← modèle d'authentification à réutiliser
```

`clasp push` envoie **tout** le contenu de `Script\` vers Google. Ne jamais
y placer de fichier Python, de secret, ou de script shell.

## Identifiants

- Sheet : `1OUsQHnN9J241pLTb3pyrKmj7MTIrjCI6PNAw4FZ835U`
- Projet Apps Script lié : `11AYtQbMbbngxUbdj0Kv6A2wqd821ieUusBG2KUT0L1vHt_wPuwGb32Xh`

## Modifier le script (Apps Script)

Éditer `Script\Code.js`, puis l'utilisateur lance `clasp push`.
Le changement n'est actif qu'après ce push.

Si `HEURE_CLOTURE` est modifiée, il faut aussi relancer
`installerDeclencheurs()` depuis l'éditeur Apps Script — le déclencheur
existant garde l'ancienne heure sinon.

Le sens est toujours local → Google. Une modification faite directement
dans l'éditeur en ligne sera écrasée au prochain push. `clasp pull` pour
la récupérer d'abord.

## Modifier les données (Python)

Écrire le script dans `data\`, réutiliser le bloc d'authentification de
`test_connexion.py`. L'utilisateur lance `python le_script.py`.

L'écriture est **immédiate et irréversible** — pas d'étape de publication,
pas de Ctrl+Z. Le seul recours est l'historique des versions Google, qui
restaure tout le document.

Règle de travail : d'abord un script qui **affiche** ce qu'il compte
modifier, l'utilisateur valide, ensuite seulement l'écriture.

Écrire en une opération groupée (`update` sur une plage) plutôt qu'en
boucle cellule par cellule : plus rapide, et pas d'état à moitié modifié
si le script plante en cours de route.

## Contraintes du classeur

**Onglet Caisse, colonnes A à H** : écrites par la synchro automatique.
- Colonne D : montant
- Colonne E : catégorie — doit exister dans la liste déroulante de
  validation de données, sinon la synchro refuse d'écrire
- Colonne H : **formule**, jamais une valeur. Toute ligne ajoutée doit
  porter `=H(n-1)+D(n)` pour que le solde cumulé se recalcule.

**Colonnes I et J : ne jamais y toucher.** Réservées à la saisie manuelle
(comptage, projet). Le script Apps Script ne les écrit pas.

**Synchro automatique à 21h** (heure du Maroc). Elle se positionne d'après
la dernière ligne portant un code en colonne A. Éviter les écritures
Python massives autour de cette heure : le verrou `LockService` du script
protège contre deux exécutions Apps Script simultanées, mais il ne
connaît pas Python.

**Onglets sources du compte de résultat** : `Caisse`, `Banque`,
` Saisies CB Karim` (noter l'espace initial dans ce dernier nom).

**Séparateur de formules** : ce classeur est en paramètres régionaux
français, donc Google Sheets attend `;` entre les arguments d'une
fonction, pas `,`. Une formule avec des virgules est rejetée.

## Sécurité

Les fichiers `credentials.json` et `authorized_user.json` sont des
secrets d'authentification. Ne jamais les lire, les afficher, les copier
ailleurs, ni inclure leur contenu dans un fichier ou un message.

Vérifier que `.gitignore` les exclut avant tout `git init` ou premier
commit.

## Avant une opération d'écriture massive

Demander à l'utilisateur de faire une copie du classeur
(Fichier → Créer une copie) avant d'exécuter. C'est sa comptabilité.
