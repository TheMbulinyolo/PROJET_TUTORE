# Interface Excel / Python

Ce projet permet de saisir une poutre dans Excel, d'executer le moteur de
calcul Python avec des macros VBA, puis de ramener les reactions, efforts et
diagrammes dans le classeur.

La couche d'integration est organisee ainsi :

- `vba/Macros_Isostatique.bas` : macros Excel ;
- `scripts/main_excel.py` : point d'entree appele par VBA ;
- `core/interface/parseur.py` : lecture et validation du JSON ;
- `core/interface/orchestrateur.py` : appel du moteur de calcul ;
- `core/interface/exporteur.py` : creation du fichier de resultats.

## Prerequis

L'execution des macros necessite :

- Windows ;
- Microsoft Excel avec les macros VBA autorisees ;
- Python 3 installe ;
- le classeur `excel/EXCEL-INTERFACE.xlsm` ;
- les dependances Python de `config/requirements.txt`.

Le fichier `excel/Interface .xlsx` ne peut pas executer de macros. Il faut
utiliser le fichier avec l'extension `.xlsm`.

## 1. Preparer Python

Ouvrir PowerShell ou l'invite de commandes dans le dossier racine du projet,
puis creer un environnement virtuel :

```bat
py -3 -m venv .venv
```

Activer l'environnement dans PowerShell :

```powershell
.venv\Scripts\Activate.ps1
```

Ou dans l'invite de commandes (`cmd`) :

```bat
.venv\Scripts\activate.bat
```

Installer les dependances :

```bat
python -m pip install -r config\requirements.txt
```

Verifier que le moteur Python fonctionne :

```bat
python scripts\main_excel.py data\input.json data\output.json
```

Un fichier `data\output.json` doit etre cree avec un statut `OK`.

La macro cherche automatiquement Python dans cet ordre :

1. `.venv\Scripts\python.exe` ;
2. `venv\Scripts\python.exe` ;
3. `ENV\Scripts\python.exe` ;
4. le lanceur Windows `py -3`.

Il est donc recommande de creer `.venv` directement a la racine du projet.

## 2. Ouvrir le classeur

1. Ouvrir `excel/EXCEL-INTERFACE.xlsm` dans Microsoft Excel.
2. Si Windows affiche un avertissement de securite, fermer Excel, faire un
   clic droit sur le fichier, choisir **Proprietes**, cocher **Debloquer**, puis
   rouvrir le classeur.
3. Cliquer sur **Activer la modification**, si ce bouton apparait.
4. Cliquer sur **Activer le contenu** pour autoriser les macros.
5. Conserver le classeur dans le dossier `excel` du projet.

La macro retrouve la racine du projet seulement si le classeur est place :

- directement a la racine du projet ; ou
- dans un sous-dossier direct comme `excel`.

## 3. Importer les macros si necessaire

Le classeur `.xlsm` doit normalement deja contenir les macros. Si elles sont
absentes :

1. Appuyer sur `Alt + F11` pour ouvrir l'editeur VBA.
2. Dans le menu **Fichier**, choisir **Importer un fichier**.
3. Selectionner `vba/Macros_Isostatique.bas`.
4. Verifier que le module `Macros_Isostatique` apparait dans le projet VBA.
5. Enregistrer le classeur au format **Classeur Excel prenant en charge les
   macros (`.xlsm`)**.
6. Fermer l'editeur VBA avec `Alt + Q`.

Aucune reference VBA supplementaire n'est requise : les objets Windows sont
charges automatiquement par le module.

## 4. Saisir un exemple

Dans la feuille `STRUCTURE` :

- `C13` : longueur de la poutre, par exemple `6` ;
- `C14` : nombre de points des diagrammes, par exemple `61` ;
- lignes `20` a `29`, colonnes `B:D` : nom, type et position des appuis.

Exemple :

| Nom | Type | Position |
| --- | --- | ---: |
| A | Pivot | 0 |
| B | Rouleau | 6 |

Dans la feuille `CHARGES`, saisir au moins une charge.

Force ponctuelle, lignes `21` a `28` :

| Nom | Fx | Fy | Position |
| --- | ---: | ---: | ---: |
| P1 | 0 | -10 | 3 |

Charge repartie, lignes `34` a `41` :

| Nom | q | Debut | Fin |
| --- | ---: | ---: | ---: |
| Q1 | -2 | 1 | 5 |

Les charges dirigees vers le bas sont saisies avec une valeur negative.

## 5. Executer la macro principale

La methode recommandee est la macro `CalculerStructure`, car elle enchaine
automatiquement toutes les operations :

1. Appuyer sur `Alt + F8`.
2. Selectionner `CalculerStructure`.
3. Cliquer sur **Executer**.
4. Attendre le message **Calcul termine avec succes**.

La macro effectue successivement :

1. la validation des donnees Excel ;
2. l'export vers `data\input_structure.json` ;
3. l'execution de `scripts\main_excel.py` ;
4. la creation de `data\resultats.json` ;
5. l'import des resultats dans Excel ;
6. l'actualisation des diagrammes ;
7. l'ouverture de la feuille `RESULTATS`.

Les boutons du classeur peuvent aussi appeler directement ces macros s'ils ont
deja ete associes.

## 6. Executer les macros separement

Pour tester chaque etape avec `Alt + F8`, utiliser cet ordre :

1. `ValiderDonneesStructure` : controle les champs saisis.
2. `ExporterDonneesVersJSON` : cree `data\input_structure.json`.
3. `LancerCalculPython` : lance Python et cree `data\resultats.json`.
4. `ImporterResultatsDepuisJSON` : remplit la feuille `RESULTATS`.
5. `ActualiserDiagrammes` : remplit la feuille `DIAGRAMMES` et actualise les
   graphiques.

Les macros `ReinitialiserStructure`, `ReinitialiserCharges`,
`ReinitialiserResultats`, `ReinitialiserDiagrammes` et `ReinitialiserTout`
permettent d'effacer les zones correspondantes.

## 7. Fichiers produits

Pendant l'execution, les fichiers suivants sont utilises :

- `data\input_structure.json` : donnees exportees depuis Excel ;
- `data\resultats.json` : resultats renvoyes par Python ;
- `logs\log.txt` : journal d'execution et details des erreurs.

Le programme Python retourne :

- code `0` : calcul reussi ;
- code `2` : donnees d'entree invalides ;
- code `1` : erreur inattendue.

## Depannage

### Les macros sont bloquees

Verifier que le fichier utilise est bien `EXCEL-INTERFACE.xlsm`, puis autoriser
les macros dans la barre de securite d'Excel. Le deblocage du fichier dans les
proprietes Windows peut aussi etre necessaire.

### Python est introuvable

Creer `.venv` a la racine du projet et reinstaller les dependances :

```bat
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r config\requirements.txt
```

### Le script Python est introuvable

Verifier que le classeur se trouve toujours dans `excel` et que le fichier
`scripts\main_excel.py` existe a la racine du projet.

### Aucun fichier de sortie n'est cree

Tester directement la commande suivante depuis la racine :

```bat
.venv\Scripts\python.exe scripts\main_excel.py data\input_structure.json data\resultats.json
```

Consulter ensuite `logs\log.txt` pour obtenir le detail de l'erreur.

### Le calcul est refuse

Verifier notamment :

- une longueur strictement positive ;
- un nombre de points entier compris entre `2` et `100` ;
- au moins deux appuis ;
- au moins une charge ;
- des identifiants d'appuis et de charges uniques ;
- des positions comprises dans la longueur de la poutre ;
- pour une charge repartie : `0 <= debut < fin <= longueur`.
