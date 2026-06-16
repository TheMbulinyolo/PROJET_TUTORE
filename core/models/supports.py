from dataclasses import dataclass
from typing import List
from .reactions import Reaction

@dataclass
class Appui:
    """
    Classe de base pour les appuis.
    """
    nom: str
    x: float
    
    def get_reactions(self) -> List[Reaction]:
        """
        Chaque type d'appui doit reourner ses reactions associées.
        """
        raise NotImplementedError()
    
@dataclass
class Rouleau(Appui):
    """
    Il bloque uniquement le déplacement vertical.
    Il crée donc une seule réaction verticale.
    """
    def get_reactions(self) -> List[Reaction]:
        return [Reaction(nom=f"{self.nom}y", x=self.x, direction="y")]
    
@dataclass
class Rotule(Appui):
    """
    Il bloque les déplacements verticaux et horizontaux.
    Il crée donc deux réactions, une verticale et une horizontale.
    """
    def get_reactions(self) -> List[Reaction]:
        return [
            Reaction(nom=f"{self.nom}y", x=self.x, direction="y"),
            Reaction(nom=f"{self.nom}x", x=self.x, direction="x")
        ]
    
@dataclass
class Encastrement(Appui):
    """
    Il bloque les déplacements verticaux et horizontaux ainsi que la rotation.
    Il crée donc trois réactions, une verticale, une horizontale et un moment d'encastrement.
    """
    def get_reactions(self) -> List[Reaction]:
        return [
            Reaction(nom=f"{self.nom}y", x=self.x, direction="y"),
            Reaction(nom=f"{self.nom}x", x=self.x, direction="x"),
            Reaction(nom=f"M{self.nom}", x=self.x, direction="moment")
        ]