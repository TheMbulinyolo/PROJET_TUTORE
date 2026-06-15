from dataclasses import dataclass
from typing import Dict, List
import sympy as sp


@dataclass
class ResultatEquilibre:
    """
    Résultat du calcul des réactions d'appui.
    """
    reactions: Dict[str, float]
    equations: List[sp.Eq]


@dataclass
class ResultatDiagrammes:
    """
    Valeurs numériques des diagrammes le long de la poutre.

    x : positions le long de la poutre
    N : effort normal
    T : effort tranchant
    M : moment fléchissant
    """
    x: List[float]
    N: List[float]
    T: List[float]
    M: List[float]
