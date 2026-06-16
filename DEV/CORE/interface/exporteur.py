"""Serialisation des resultats du moteur pour Excel/VBA."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def ecrire_json(donnees: dict[str, Any], chemin: str | Path) -> None:
    chemin = Path(chemin)
    chemin.parent.mkdir(parents=True, exist_ok=True)

    temporaire = chemin.with_suffix(f"{chemin.suffix}.tmp")
    with temporaire.open("w", encoding="utf-8") as fichier:
        json.dump(
            _normaliser(donnees),
            fichier,
            ensure_ascii=False,
            indent=2,
            allow_nan=False,
        )
        fichier.write("\n")
    temporaire.replace(chemin)


def construire_sortie_erreur(message: str, type_erreur: str = "Erreur") -> dict[str, Any]:
    return {
        "statut": "ERREUR",
        "message": message,
        "type_erreur": type_erreur,
        "isostatique": False,
        "reactions": {},
        "efforts_position": None,
        "maxima": None,
        "diagrammes": None,
    }


def _normaliser(valeur: Any) -> Any:
    if isinstance(valeur, dict):
        return {str(cle): _normaliser(item) for cle, item in valeur.items()}
    if isinstance(valeur, (list, tuple)):
        return [_normaliser(item) for item in valeur]
    if hasattr(valeur, "item"):
        return valeur.item()
    return valeur
