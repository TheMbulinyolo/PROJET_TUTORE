"""Interface JSON entre Excel/VBA et le moteur de calcul."""

from .exporteur import construire_sortie_erreur, ecrire_json
from .orchestrateur import executer_calcul
from .parseur import DonneesCalcul, ErreurValidation, lire_et_parser

__all__ = [
    "DonneesCalcul",
    "ErreurValidation",
    "construire_sortie_erreur",
    "ecrire_json",
    "executer_calcul",
    "lire_et_parser",
]
