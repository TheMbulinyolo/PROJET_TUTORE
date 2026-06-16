# Projet Tutoré - Calcul de poutres isostatiques

Ce projet relie une interface Excel/VBA a un moteur Python de calcul de
poutres isostatiques. La refactorisation actuelle reorganise les dossiers sans
modifier les algorithmes de calcul.

## Structure du projet

```text
PROJET_TUTORE/
├── core/                 # Coeur metier : modeles, solveur, interface JSON
│   ├── models/           # Poutres, appuis, charges et reactions
│   ├── resolver/         # Calculs d'equilibre, efforts et diagrammes
│   └── interface/        # Parseur, orchestrateur et exporteur JSON
├── scripts/              # Points d'entree executables
├── vba/                  # Module VBA a importer dans Excel
├── excel/                # Classeurs Excel de l'interface utilisateur
├── data/                 # Fichiers d'entree/sortie JSON
├── logs/                 # Journaux d'execution
├── tests/                # Tests automatises Python
├── docs/                 # Documentation detaillee
├── config/               # Dependances et configuration de packaging
├── build/                # Artefacts intermediaires PyInstaller
└── dist/                 # Executable genere
```

## Installation Python

Depuis la racine du projet :

```bat
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r config\requirements.txt
```

Sous Linux/macOS :

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r config/requirements.txt
```

## Lancer le moteur Python

Avec les fichiers JSON par defaut :

```bash
python scripts/main_excel.py
```

Avec des chemins explicites :

```bash
python scripts/main_excel.py data/input.json data/output.json
```

Le script lit un fichier JSON d'entree, execute le moteur de calcul, puis ecrit
un fichier JSON de sortie.

## Lancer depuis Excel

1. Ouvrir `excel/EXCEL-INTERFACE.xlsm`.
2. Activer les macros si Excel le demande.
3. Importer `vba/Macros_Isostatique.bas` si le module n'est pas deja present.
4. Executer la macro `CalculerStructure`.

Les macros utilisent :

- `data/input_structure.json` pour les donnees exportees depuis Excel ;
- `data/resultats.json` pour les resultats renvoyes au classeur ;
- `dist/moteur_calcul.exe` pour l'executable PyInstaller ;
- `logs/log.txt` pour le journal d'execution.

## Points d'entree principaux

- `scripts/main_excel.py` : point d'entree Python appele directement ou par
  l'executable.
- `vba/Macros_Isostatique.bas` : macros Excel/VBA.
- `excel/EXCEL-INTERFACE.xlsm` : classeur principal avec interface utilisateur.
- `tests/test_interface_excel.py` : tests du contrat JSON Excel/Python.

## Tests

Depuis la racine du projet :

```bash
python -m unittest tests/test_interface_excel.py
```

La documentation detaillee de l'integration Excel/Python se trouve dans
`docs/README_INTERFACE_EXCEL.md`.
