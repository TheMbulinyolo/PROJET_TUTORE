from dataclasses import dataclass
from typing import Dict, List
import sympy as sp

from core.models.poutre import Poutre
from core.models.charges import ForcePonctuelle, ForceRepartieUniforme
from core.models.reactions import Reaction
from core.models.supports import Rouleau, Encastrement, Rotule


@dataclass
class ResultatEquilibre:
    """
    Résultat du calcul des réactions d'appui.
    reactions: dictionnaire {nom_reaction: valeur}
    equations: équations d'équilibre utilisées par le solveur
    """
    reactions: Dict[str, float]
    equations: List[sp.Eq]


class ResolverIsostatique2D:
    """
    Solveur pour poutres isostatiques 2D.

    Convention utilisée :
    - axe x : horizontal vers la droite
    - axe y : vertical vers le haut
    - force vers le bas : valeur négative
    - moment positif : sens trigonométrique
    """

    def __init__(self, poutre: Poutre):
        self.poutre = poutre
        self.reactions = poutre.get_reactions()

    def verifier_structure(self) -> None:
        """Vérifie que la structure est calculable par les 3 équations d'équilibre 2D."""
        if not self.poutre.is_isostatique_2d():
            raise ValueError(
                f"Structure non isostatique en 2D : "
                f"{len(self.reactions)} réactions inconnues au lieu de 3."
            )

    def _force_resultante(self, charge) -> float:
        return charge.force_resultante()

    def _position_resultante(self, charge) -> float:
        """
        Compatibilité avec tes noms de méthodes.
        Dans la classe mère tu as resultant_position(),
        mais dans les classes filles tu as position_resultante().
        """
        if hasattr(charge, "position_resultante"):
            return charge.position_resultante()

        raise AttributeError("La charge ne fournit pas de position de résultante.")

    def construire_equations(self) -> List[sp.Eq]:
        """Construit les équations ΣFx=0, ΣFy=0 et ΣM(O)=0."""
        symboles = {reaction.nom: sp.Symbol(reaction.nom) for reaction in self.reactions}

        somme_fx = 0
        somme_fy = 0
        somme_moment_o = 0

        # Contribution des réactions d'appui
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
                raise ValueError(f"Direction de réaction inconnue : {reaction.direction}")

        # Contribution des charges extérieures
        for charge in self.poutre.charges:
            F = self._force_resultante(charge)
            xF = self._position_resultante(charge)

            # Dans l'état actuel de tes modèles, les charges sont verticales.
            somme_fy += F
            somme_moment_o += F * xF

        return [
            sp.Eq(somme_fx, 0),
            sp.Eq(somme_fy, 0),
            sp.Eq(somme_moment_o, 0),
        ]

    def resoudre_reactions(self) -> ResultatEquilibre:
        """Résout les réactions inconnues et les réinjecte dans les objets Reaction."""
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

        return ResultatEquilibre(
            reactions=reactions_calculees,
            equations=equations,
        )

if __name__ == "__main__":
    # Exemple d'utilisation
    poutre = Poutre(longueur=6)
    poutre.ajouter_appui(Rotule(nom="A", x=0))
    poutre.ajouter_appui(Rouleau(nom="B", x=6))
    poutre.ajouter_charge(ForcePonctuelle(valeur=-10, x=2))
    poutre.ajouter_charge(ForceRepartieUniforme(q=-4, debut=3, fin=6))

    resolver = ResolverIsostatique2D(poutre)
    resultat = resolver.resoudre_reactions()
    print(resultat.reactions)

    