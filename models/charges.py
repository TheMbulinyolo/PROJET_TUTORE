from dataclasses import dataclass


@dataclass
class Charge:
    """
    Classe de base pour les charges.
    """
    pass


@dataclass
class ForcePonctuelle(Charge):
    """
    Force ponctuelle 2D appliquée en x.

    fx > 0 : force horizontale vers la droite
    fx < 0 : force horizontale vers la gauche

    fy > 0 : force verticale vers le haut
    fy < 0 : force verticale vers le bas
    """
    fx: float
    fy: float
    x: float

    def force_x(self) -> float:
        return self.fx

    def force_y(self) -> float:
        return self.fy


@dataclass
class ForceRepartieUniforme(Charge):
    """
    Charge répartie uniformément verticale.

    q > 0 : vers le haut
    q < 0 : vers le bas
    """
    q: float
    debut: float
    fin: float

    def longueur(self) -> float:
        return self.fin - self.debut

    def force_resultante(self) -> float:
        return self.q * self.longueur()

    def position_resultante(self) -> float:
        return (self.debut + self.fin) / 2