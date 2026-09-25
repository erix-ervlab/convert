# Extraction et pseudonymisation de documents pour analyse par IA

Convertit un lot de documents Word (et PDF) en Markdown structuré, lisible par
une IA, en remplaçant une liste de termes sensibles par des codes stables.

- **Structure conservée** : titres, listes, tableaux (au format Markdown), notes de bas de page.
- **Version finale uniquement** : les modifications suivies sont acceptées ; commentaires,
  métadonnées, en-têtes et pieds de page ne sont pas repris.
- **OCR des images** (optionnel) : le texte lisible des captures d'écran est inséré à
  l'emplacement de l'image, dans un bloc identifié.
- **Pseudonymisation par table** : chaque terme reçoit un code stable (`ORG-01`,
  `CLIENT-03`…), identique dans tous les documents, ce qui permet de comparer les
  documents entre eux sans exposer les noms.
- **Contrôles** : rapport par fichier, détail par terme, occurrences restantes,
  extraits avant/après de chaque remplacement.

Les documents d'origine ne sont jamais modifiés.

## Installation

Deux moteurs de conversion Word sont disponibles :

- **python** (mammoth) : aucun logiciel à installer en dehors de Python, adapté aux
  postes sur lesquels on ne peut rien installer d'autre que des paquets `pip` ;
- **pandoc** : utilisé automatiquement s'il est installé.

Les deux produisent un résultat équivalent (révisions acceptées, tableaux Markdown,
notes de bas de page). Le moteur python retire en plus les lignes de tableau
supprimées en mode révision, que pandoc conserve.

Minimum, pour les `.docx` sans OCR :

```
pip install mammoth beautifulsoup4
```

Prérequis système optionnels :

- [pandoc](https://pandoc.org/installing.html) 3.1 ou plus — autre moteur de conversion des `.docx`
- [Tesseract](https://tesseract-ocr.github.io/tessdoc/Installation.html) avec le pack
  de langue français (`fra`) — OCR des images (sous Linux : `apt install tesseract-ocr tesseract-ocr-fra`)
- LibreOffice (optionnel) — OCR des schémas au format `.emf` / `.wmf`

Pour tout installer (OCR et PDF compris) :

```
pip install -r requirements.txt
```

Paquets par usage : `mammoth` et `beautifulsoup4` (Word sans pandoc), `pytesseract` et
`pillow` (OCR), `pymupdf4llm` (PDF, licence AGPL-3.0).

## Utilisation

1. Préparer la table de pseudonymisation :

   ```
   cp pseudonymisation.exemple.csv pseudonymisation.csv
   ```

   puis la compléter (voir les explications en tête du fichier). **Cette table contient
   les vrais noms : elle est exclue du dépôt par `.gitignore` et ne doit pas être partagée.**

2. Tester la table sur une phrase (affiche aussi le rendu Markdown) :

   ```
   python extraire_anonymiser.py --table pseudonymisation.csv --forme "{{}}" --essai "Le ACMEBOX de ACME"
   ```

3. Faire un essai sur quelques fichiers, avec le détail de chaque remplacement :

   ```
   python extraire_anonymiser.py documents/ sortie/ --table pseudonymisation.csv --forme "{{}}" --note --limite 5 --contexte
   ```

4. Traiter le lot complet :

   ```
   python extraire_anonymiser.py documents/ sortie/ --table pseudonymisation.csv --forme "{{}}" --note --workers 8
   ```

## Options principales

| Option | Rôle |
|---|---|
| `--table FICHIER` | table de pseudonymisation (CSV, séparateur `;`) |
| `--mot TERME --jeton CODE` | sans table : un seul terme à remplacer |
| `--forme "{{}}"` | habillage des codes pour signaler un remplacement (`{ORG-01}`) ; `{}` = emplacement du code |
| `--note` | ajoute en tête de chaque fichier une note expliquant la pseudonymisation à l'IA |
| `--echapper` | échappe les caractères Markdown des codes (utile avec `_{}_`) |
| `--essai TEXTE` | teste la table sur un texte, sans traiter de fichier |
| `--limite N` | ne traite que les N premiers fichiers |
| `--contexte` | produit `contextes.csv` : chaque remplacement avec son extrait avant/après |
| `--moteur auto\|pandoc\|python` | moteur de conversion Word (auto : pandoc s'il est installé, sinon python) |
| `--sans-ocr` | ignore le texte des images |
| `--ocr-confiance N` | confiance minimale par mot OCR (0-100, défaut 70) |
| `--workers N` | nombre de fichiers traités en parallèle |

## Fichiers produits

Dans le dossier de sortie :

- un fichier `.md` par document, dans la même arborescence ;
- `rapport.csv` : statut, nombre de remplacements, occurrences restantes et images lues, par fichier ;
- `occurrences.csv` : remplacements et restes, par fichier et par terme ;
- `contextes.csv` (option `--contexte`) : chaque remplacement avec son contexte.

Ces fichiers contiennent des extraits des documents : ils sont exclus du dépôt.

## Choix du format des codes

Les accolades (`--forme "{{}}"` → `{ORG-01}`) sont recommandées :

- elles n'ont aucune signification en Markdown et s'affichent telles quelles ;
- elles sont rarement présentes dans des documents rédigés, donc tout `{…}` est un remplacement ;
- dans Excel, elles empêchent la conversion automatique de certains codes en dates.

À éviter : `_{}_` sans `--echapper` (interprété comme de l'italique), `[{}]` (devient un
lien s'il est suivi d'une parenthèse), et les guillemets `« »`, fréquents dans les textes.

## Limites

- **Mode `prefixe`** : il remplace tout mot commençant par le terme. Pour un terme court,
  préférer le mode `mot`, et vérifier `contextes.csv`.
- **OCR** : fiable sur les captures de texte (logs, consoles), médiocre sur les schémas,
  nul sur les graphiques. Un terme mal lu par l'OCR peut ne pas être reconnu : le texte
  extrait des images doit être relu.
- **PDF** : la reconstruction des tableaux est moins fiable qu'à partir du Word d'origine
  (tableaux coupés entre deux pages notamment). Préférer les fichiers Word.
- **Cellules fusionnées** dans les tableaux Word : la valeur n'apparaît que dans la première cellule.
- **Pseudonymisation n'est pas anonymisation** : seuls les termes de la table sont
  remplacés ; le contexte technique peut suffire à identifier une organisation.
