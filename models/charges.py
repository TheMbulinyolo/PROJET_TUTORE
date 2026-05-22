from dataclasses import dataclass

@dataclass
class Charge:
    """
    Classe de base pour les charges.
    """

    def force_resultante(self) -> float:
        """
        Chaque type de charge doit implémenter sa propre méthode pour calculer la force résultante.
        """
        raise NotImplementedError()
    
    def resultant_position(self) -> float:
        """
        Chaque type de charge doit implémenter sa propre méthode pour calculer la position de la force résultante.
        """
        raise NotImplementedError()
    
@dataclass
class ForcePonctuelle(Charge):
    """
    Représente une force ponctuelle appliquée à une position donnée.
    """
    valeur: float
    x: float

    def force_resultante(self) -> float:
        return self.valeur
    
    def position_resultante(self) -> float:
        return self.x
    
@dataclass
class ForceRepartieUniforme(Charge):
    """
    Charge repartie uniformément sur une longueur donnée.

    q < 0 : charge vers le bas
    q > 0 : charge vers le haut
    """
    q: float
    debut : float
    fin: float

    def longuer(self) -> float:
        return self.fin - self.debut
    
    def force_resultante(self) -> float:
        return self.q * self.longuer()
    
    def position_resultante(self) -> float:
        return (self.debut + self.fin) / 2