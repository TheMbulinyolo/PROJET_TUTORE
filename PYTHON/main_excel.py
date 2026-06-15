"""Point d'entree appele par Excel/VBA.

Usage:
    python main_excel.py [chemin_input] [chemin_output]
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path


RACINE_PROJET = Path(__file__).resolve().parents[1]
if str(RACINE_PROJET) not in sys.path:
    sys.path.insert(0, str(RACINE_PROJET))

from CORE.interface.exporteur import construire_sortie_erreur, ecrire_json
from CORE.interface.orchestrateur import executer_calcul
from CORE.interface.parseur import ErreurValidation, lire_et_parser


def configurer_journal(chemin: Path) -> None:
    chemin.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=chemin,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        encoding="utf-8",
    )


def main() -> int:
    chemin_entree = (
        Path(sys.argv[1]).resolve()
        if len(sys.argv) > 1
        else RACINE_PROJET / "DATA" / "input.json"
    )
    chemin_sortie = (
        Path(sys.argv[2]).resolve()
        if len(sys.argv) > 2
        else RACINE_PROJET / "DATA" / "output.json"
    )
    configurer_journal(RACINE_PROJET / "LOGS" / "log.txt")

    try:
        logging.info("Lecture de %s", chemin_entree)
        donnees = lire_et_parser(chemin_entree)
        sortie = executer_calcul(donnees)
        ecrire_json(sortie, chemin_sortie)
        logging.info("Calcul termine avec succes : %s", chemin_sortie)
        return 0
    except ErreurValidation as exc:
        logging.warning("Donnees invalides : %s", exc)
        ecrire_json(
            construire_sortie_erreur(str(exc), "ErreurValidation"),
            chemin_sortie,
        )
        return 2
    except Exception as exc:
        logging.exception("Erreur inattendue")
        ecrire_json(
            construire_sortie_erreur(str(exc), type(exc).__name__),
            chemin_sortie,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
