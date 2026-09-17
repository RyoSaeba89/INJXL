# INJXL

Injection de colonnes d'un tableau vers un autre, par clé d'association.

Aucune colonne n'est figée dans le code : tout se règle dans une interface graphique,
et rien n'est écrit avant une simulation validée.

Version courante : **2.0.0**

---

## Lancement

| Situation | Commande |
|---|---|
| Interface graphique | double-clic sur `INJXL.bat`, ou `python INJXL.py` |
| Interface préchargée | `python INJXL.py profils\mon_profil.json` |
| Simulation en ligne de commande | `python INJXL.py --profil profils\mon_profil.json --simuler` |
| Exécution automatique | `python INJXL.py --profil mon.json --source A.xlsx --cible B.xlsx --oui` |
| Créer l'exécutable | `build_exe.bat` → `dist\INJXL.exe` (compilé **et signé**) |

Dépendance : `openpyxl` (`pip install openpyxl`). `xlrd` en plus si vous devez **lire** des `.xls`.

---

## Les quatre étapes de l'interface

**1 · Tableaux** — Choix du fichier SOURCE (celui qui détient les données) et du fichier
CIBLE (celui à compléter). Pour chacun : la feuille et la ligne d'en-tête, avec un aperçu
immédiat des 40 premières lignes. Formats : `.xlsx`, `.xlsm`, `.csv`, `.txt`, `.tsv`,
et `.xls` en lecture seule.

**2 · Clé d'association** — Les colonnes qui identifient une même ligne des deux côtés.
Plusieurs couples = clé composite (ex. `Code` + `Année`) : toutes les parties
doivent correspondre. Le bouton *Deviner* propose les colonnes de même en-tête.
Options de comparaison : casse, espaces, accents, séparateurs, zéros de tête, `1234,0` = `1234`.

**3 · Colonnes à injecter** — Autant d'injections que nécessaire. Pour chacune : la colonne
source, la destination (une colonne existante **ou une nouvelle colonne créée à la fin**),
le comportement si la cellule cible est déjà remplie, et une valeur par défaut facultative
quand la clé est introuvable.

**4 · Vérification et écriture** — L'analyse est une **simulation** : elle n'écrit rien.
Elle affiche ligne par ligne la valeur actuelle et la valeur proposée, avec un statut par
ligne, des compteurs, un filtre et une recherche. Chaque ligne « À écrire » peut être
**décochée** pour l'exclure. L'écriture n'a lieu qu'après confirmation explicite.

---

## Statuts

| Statut | Signification |
|---|---|
| **À écrire** | la cellule cible sera renseignée |
| **Identique** | la cible contient déjà la bonne valeur |
| **Conflit** | la cible contient une *autre* valeur ; conservée ou écrasée selon le mode choisi |
| **Orphelin** | la clé de la cible n'existe pas dans la source |
| **Source ambiguë** | la clé apparaît plusieurs fois dans la source avec des valeurs différentes ; écartée par sécurité (arbitrage réglable : première / dernière occurrence) |
| **Sans clé** | ligne non vide mais sans valeur de clé |
| **Sans effet** | la source n'a rien à apporter pour cette ligne |

Sont également signalés : les doublons, les quasi-doublons (`007` / `7`, `AB-12` / `AB12`),
les formules sans valeur en cache, les clés source jamais utilisées, et les configurations
risquées (écrire dans une colonne qui sert de clé, deux injections vers la même colonne).

---

## Production du résultat

- **Copie** (défaut) : nouveau fichier horodaté, ou chemin libre. L'original n'est pas touché.
- **Sur place** : le fichier cible est modifié, avec sauvegarde horodatée préalable.
- Surlignage vert des cellules modifiées, rapport `.txt` détaillé et rapport `.xlsx`
  (synthèse + détail filtrable + anomalies).

Formules, mises en forme et macros du fichier cible sont préservées : le classeur est
rouvert et modifié cellule par cellule, jamais reconstruit.

---

## Profils

Menu **Profil → Enregistrer sous** : toute la configuration (fichiers, feuilles, lignes
d'en-tête, clés, injections, options, sortie) part dans un `.json` rejouable d'un clic —
ou en ligne de commande pour une tâche planifiée.

Les colonnes y sont mémorisées **par en-tête *et* par position** : si une colonne se
déplace dans un fichier, INJXL la retrouve par son libellé et le signale dans le rapport.

`profils/exemple.json` sert de point de départ : clé en colonne A des deux côtés,
injection de la colonne L vers la colonne O, sans écraser l'existant. Les colonnes y sont
repérées par position uniquement (en-têtes laissés vides), donc il s'applique à n'importe
quelle paire de fichiers ayant cette disposition.

---

## Structure

```
INJXL.py              point d'entrée (GUI + ligne de commande)
INJXL.bat             lancement sans console
build_exe.bat         compilation en .exe autonome, puis signature
injxl/
  model.py            profil, colonnes, options, sérialisation JSON
  tables.py           lecture/écriture xlsx, xlsm, csv, xls
  engine.py           normalisation, indexation, analyse d'impact, écriture
  gui.py              interface Tkinter/ttk
profils/              profils d'exemple
tests/                vérifications du moteur, sans interface
signature/            clé publique du certificat de signature
signer.ps1            signature de l'exe (appelé par build_exe.bat)
dist/INJXL.exe        exécutable signé, attaché aux releases
```

---

## Signature

`build_exe.bat` compile puis appelle `signer.ps1`, qui signe en SHA256 et horodate auprès de
DigiCert — l'horodatage garde la signature valide après l'expiration du certificat. À relancer
après **chaque** recompilation : PyInstaller réécrit l'exe et la signature précédente disparaît.

Le certificat est **auto-signé**, pas délivré par une autorité publique : Windows ne le
connaît pas tant que sa clé publique n'a pas été importée. Cette clé publique seule
(aucun secret) est fournie dans `signature/certificat-public.cer` :

```powershell
Import-Certificate -FilePath .\certificat-public.cer -CertStoreLocation Cert:\LocalMachine\Root
Import-Certificate -FilePath .\certificat-public.cer -CertStoreLocation Cert:\LocalMachine\TrustedPublisher
```
