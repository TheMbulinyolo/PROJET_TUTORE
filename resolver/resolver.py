from .equilibre import SolveurEquilibre2D
from .efforts_internes import CalculEffortsInternes2D
from .diagrammes import GenerateurDiagrammes2D

from core.models.poutre import Poutre
from core.models.supports import Rotule, Rouleau
from core.models.charges import ForcePonctuelle, ForceRepartieUniforme


class ResolverIsostatique2D:
    """
    Façade principale du resolver.

    Cette classe garde une API simple pour l'utilisateur,
    mais délègue le travail à des modules spécialisés :
    - equilibre.py : réactions d'appui
    - efforts_internes.py : N(x), T(x), M(x)
    - diagrammes.py : tableaux de valeurs et maxima
    """

    def __init__(self, poutre):
        self.poutre = poutre

        self.solveur_equilibre = SolveurEquilibre2D(poutre)
        self.calcul_efforts = CalculEffortsInternes2D(
            poutre=poutre,
            solveur_equilibre=self.solveur_equilibre,
        )
        self.generateur_diagrammes = GenerateurDiagrammes2D(
            poutre=poutre,
            calcul_efforts=self.calcul_efforts,
        )

    @property
    def reactions(self):
        return self.solveur_equilibre.reactions

    @property
    def resultat_equilibre(self):
        return self.solveur_equilibre.resultat_equilibre

    def verifier_structure(self):
        return self.solveur_equilibre.verifier_structure()

    def construire_equations(self):
        return self.solveur_equilibre.construire_equations()

    def resoudre_reactions(self):
        return self.solveur_equilibre.resoudre_reactions()

    def effort_normal(self, x: float) -> float:
        return self.calcul_efforts.effort_normal(x)

    def effort_tranchant(self, x: float) -> float:
        return self.calcul_efforts.effort_tranchant(x)

    def moment_flechissant(self, x: float) -> float:
        return self.calcul_efforts.moment_flechissant(x)

    def generer_diagrammes(self, nb_points: int = 101):
        return self.generateur_diagrammes.generer_diagrammes(nb_points=nb_points)

    def valeurs_maximales(self, nb_points: int = 1001):
        return self.generateur_diagrammes.valeurs_maximales(nb_points=nb_points)
    

if __name__ == '__main__':
    poutre = Poutre(longueur=10)

    poutre.ajouter_appui(Rotule(nom="A", x=0))
    poutre.ajouter_appui(Rouleau(nom="B", x=10))

    poutre.ajouter_charge(ForcePonctuelle(fx=6, fy=-12, x=2))
    poutre.ajouter_charge(ForceRepartieUniforme(q=-4, debut=4, fin=8))
    poutre.ajouter_charge(ForcePonctuelle(fx=0, fy=-10, x=9))

    resolver = ResolverIsostatique2D(poutre)

    resultat = resolver.resoudre_reactions()
    print("Réactions :", resultat.reactions)

    print("T(3) =", resolver.effort_tranchant(3))
    print("M(3) =", resolver.moment_flechissant(3))

    print("T(6) =", resolver.effort_tranchant(6))
    print("M(6) =", resolver.moment_flechissant(6))

    print("T(9) =", resolver.effort_tranchant(9))
    print("M(9) =", resolver.moment_flechissant(9))

    diagrammes = resolver.generer_diagrammes(nb_points=11)
    maxima = resolver.valeurs_maximales()

    print("Diagrammes :", diagrammes)
    print("Maxima :", maxima)
