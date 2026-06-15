from dataclasses import dataclass, field
from typing import List
from .charges import Charge, ForcePonctuelle, ForceRepartieUniforme
from .supports import Appui
from .reactions import Reaction

@dataclass
class Poutre:
    """
    Représente une poutre en 2D
    """

    longueur: float
    appuis: List[Appui] = field(default_factory=list)
    charges: List[Charge] = field(default_factory=list)

    def ajouter_appui(self, appui: Appui) -> None:
        if appui.x < 0 or appui.x > self.longueur:
            raise ValueError(f"Position de l'appui {appui.nom} hors de la poutre")
        self.appuis.append(appui)

    def ajouter_charge(self, charge: Charge) -> None:
        if isinstance(charge, ForcePonctuelle) and (charge.x < 0 or charge.x > self.longueur):
            raise ValueError("Position de la force ponctuelle hors de la poutre")
        if isinstance(charge, ForceRepartieUniforme) and (charge.debut < 0 or charge.fin > self.longueur):
            raise ValueError("Position de la force répartie hors de la poutre")
        self.charges.append(charge)

    def get_reactions(self) -> List[Reaction]:
        reactions = []
        for appui in self.appuis:
            reactions.extend(appui.get_reactions())
        return reactions
    
    def is_isostatique_2d(self) -> bool:
        """
        Vérifie si la poutre est isostatique en 2D.
        Une poutre est isostatique si le nombre de réactions inconnues est égal au nombre d'équations d'équilibre (3 en 2D).
        """
        return len(self.get_reactions()) == 3
