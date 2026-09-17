# Journal des versions

Format [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/),
numérotation [SemVer](https://semver.org/lang/fr/).

## [2.0.0] - 2026-09-17

Première version publiée. L'outil ne se limite plus à une seule colonne
décidée dans le code : il injecte n'importe quelle colonne d'un tableau vers
n'importe quelle colonne d'un autre, à partir d'une clé d'association choisie
dans une interface graphique.

### Ajouté

- **Interface graphique** en quatre étapes (Tkinter/ttk) : tableaux, clé
  d'association, colonnes à injecter, vérification et écriture.
- **Import des deux tableaux** avec choix de la feuille, de la ligne d'en-tête
  et aperçu immédiat des 40 premières lignes. Formats `.xlsx`, `.xlsm`,
  `.csv`, `.txt`, `.tsv`, et `.xls` en lecture seule.
- **Clé d'association composite** : autant de couples de colonnes que
  nécessaire, tous devant correspondre. Un bouton *Deviner* propose les
  colonnes qui portent le même en-tête des deux côtés.
- **Options de comparaison** : casse, espaces de bord, espaces multiples,
  accents, séparateurs (`- _ . /`), zéros de tête, et `1234,0` lu comme `1234`.
- **Injections multiples** : plusieurs colonnes rapatriées en une passe, vers
  une colonne existante **ou une nouvelle colonne créée à la fin**, avec un
  comportement réglable si la cellule cible est déjà remplie, et une valeur
  par défaut facultative quand la clé est introuvable.
- **Analyse d'impact** préalable, sans aucune écriture : statut par ligne
  (À écrire, Identique, Conflit, Orphelin, Source ambiguë, Sans clé),
  compteurs, filtre, recherche, et **décochage ligne à ligne** pour exclure
  une correction avant de l'appliquer.
- **Détection des pièges** : doublons, quasi-doublons (`007`/`7`, `AB-12`/`AB12`),
  formules sans valeur en cache, clés source jamais utilisées, écriture dans
  une colonne qui sert de clé, deux injections visant la même colonne.
- **Production du résultat** : copie horodatée (défaut), chemin libre, ou
  écriture sur place avec sauvegarde préalable. Surlignage vert des cellules
  modifiées, rapport `.txt` détaillé et rapport `.xlsx` (synthèse, détail
  filtrable, anomalies).
- **Profils `.json`** rejouables : toute la configuration est enregistrée, et
  les colonnes y sont mémorisées **par en-tête et par position** — une colonne
  déplacée est retrouvée par son libellé, et le rapport le signale.
- **Mode ligne de commande** (`--profil`, `--simuler`, `--oui`, `--source`,
  `--cible`, `--sortie`, `--surplace`) pour une tâche planifiée.
- Tests du moteur couvrant clé simple, clé composite, écriture, nouvelle
  colonne, écrasement, exclusion de lignes, cible CSV, écriture sur place et
  aller-retour de profil.
- **Exécutable Windows autonome et signé** (`INJXL.exe`, un seul fichier, sans
  installation ni Python), attaché à cette release. Signature SHA256 horodatée
  par DigiCert — la clé publique du certificat est fournie dans `signature/certificat-public.cer`.

### Modifié

- Les colonnes ne sont plus figées dans le code : l'ancien réglage
  (clé en colonne A, injection de L vers O) devient `profils/exemple.json`.
- Le fichier cible n'est plus jamais reconstruit : le classeur est rouvert et
  modifié cellule par cellule, ce qui préserve formules, mises en forme et
  macros.
