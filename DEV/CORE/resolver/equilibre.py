from typing import List
import sympy as sp

from .charge_utils import composantes_charge
from .resultats import ResultatEquilibre


class SolveurEquilibre2D:
    """
    Responsable uniquement du calcul des équations d'équilibre
    et des réactions d'appui.
    """

    def __init__(self, poutre):
        self.poutre = poutre
        self.reactions = poutre.get_reactions()
        self.resultat_equilibre = None

    def verifier_structure(self) -> None:
        if not self.poutre.is_isostatique_2d():
            raise ValueError(
                f"Structure non isostatique en 2D : "
                f"{len(self.reactions)} réactions inconnues au lieu de 3."
            )

    def construire_equations(self) -> List[sp.Eq]:
        symboles = {
            reaction.nom: sp.Symbol(reaction.nom)
            for reaction in self.reactions
        }

        somme_fx = 0
        somme_fy = 0
        somme_moment_o = 0

        # Réactions d'appui
        for reaction in self.reactions:
            R = symboles[reaction.nom]

            if reaction.direction == "x":
                somme_fx += R

            elif reaction.direction == "y":
                somme_fy += R
                somme_moment_o += R * reaction.x

            elif reaction.direction == "moment":
                somme_moment_o += R

            else:
                raise ValueError(
                    f"Direction de réaction inconnue : {reaction.direction}"
                )

        # Charges extérieures
        for charge in self.poutre.charges:
            Fx, Fy, xF = composantes_charge(charge)
            somme_fx += Fx
            somme_fy += Fy

            # Hypothèse poutre 2D : forces appliquées sur l'axe moyen.
            # Donc seule la composante verticale Fy crée un moment fléchissant.
            somme_moment_o += Fy * xF

        return [
            sp.Eq(somme_fx, 0),
            sp.Eq(somme_fy, 0),
            sp.Eq(somme_moment_o, 0),
        ]

    def resoudre_reactions(self) -> ResultatEquilibre:
        self.verifier_structure()

        equations = self.construire_equations()
        inconnues = [sp.Symbol(reaction.nom) for reaction in self.reactions]

        solution = sp.solve(equations, inconnues, dict=True)

        if not solution:
            raise ValueError("Impossible de résoudre les équations d'équilibre.")

        solution = solution[0]
        reactions_calculees = {}

        for reaction in self.reactions:
            symbole = sp.Symbol(reaction.nom)
            valeur = float(solution[symbole])
            reaction.valeur = valeur
            reactions_calculees[reaction.nom] = valeur

        self.resultat_equilibre = ResultatEquilibre(
            reactions=reactions_calculees,
            equations=equations,
        )

        return self.resultat_equilibre

    def assurer_reactions_calculees(self) -> None:
        reactions_non_calculees = [
            reaction for reaction in self.reactions
            if reaction.valeur is None
        ]

        if reactions_non_calculees:
            self.resoudre_reactions()
