# Interface Excel / Python

Le moteur scientifique reste dans `core/models` et `core/resolver`.

La couche d'integration est separee ainsi :

- `core/interface/parseur.py` : lecture, validation et conversion du JSON ;
- `core/interface/orchestrateur.py` : appel du moteur selon les options ;
- `core/interface/exporteur.py` : ecriture atomique de `output.json` ;
- `python/main_excel.py` : point d'entree appele par VBA.

## Execution

Depuis la racine du projet :

```bash
ENV/bin/python python/main_excel.py data/input.json data/output.json
```

Sous Windows avec l'environnement demande pour Excel :

```bat
.venv\Scripts\python.exe python\main_excel.py data\input.json data\output.json
```

Le programme retourne :

- code `0` : calcul reussi ;
- code `2` : donnees d'entree invalides ;
- code `1` : erreur inattendue.

Dans tous les cas ou cela est possible, `data/output.json` est produit avec
`statut` egal a `OK` ou `ERREUR`. Les details techniques sont aussi ecrits dans
`logs/log.txt`.
