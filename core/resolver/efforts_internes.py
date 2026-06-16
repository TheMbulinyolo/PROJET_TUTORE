from .charge_utils import (
    composantes_charge,
    est_force_ponctuelle_2d,
    est_charge_repartie_uniforme,
    est_ancienne_force_verticale,
)


class CalculEffortsInternes2D:
    """
    Responsable du calcul de N(x), T(x) et M(x),
    à partir des réactions déjà calculées.
    """

    def __init__(self, poutre, solveur_equilibre):
        self.poutre = poutre
        self.solveur_equilibre = solveur_equilibre
        self.reactions = solveur_equilibre.reactions

    def effort_normal(self, x: float) -> float:
        """
        N(x) = somme algébrique des forces horizontales à gauche de x.
        """
        self.solveur_equilibre.assurer_reactions_calculees()

        N = 0.0

        for reaction in self.reactions:
            if reaction.direction == "x" and x >= reaction.x:
                N += reaction.valeur

        for charge in self.poutre.charges:
            Fx, _, xF = composantes_charge(charge)
            if x >= xF:
                N += Fx

        return N

    def effort_tranchant(self, x: float) -> float:
        """
        T(x) = somme algébrique des forces verticales à gauche de x.
        """
        self.solveur_equilibre.assurer_reactions_calculees()

        T = 0.0

        for reaction in self.reactions:
            if reaction.direction == "y" and x >= reaction.x:
                T += reaction.valeur

        for charge in self.poutre.charges:
            if est_force_ponctuelle_2d(charge):
                if x >= charge.x:
                    T += charge.force_y()

            elif est_charge_repartie_uniforme(charge):
                if x < charge.debut:
                    pass
                elif x <= charge.fin:
                    T += charge.q * (x - charge.debut)
                else:
                    T += charge.q * (charge.fin - charge.debut)

            elif est_ancienne_force_verticale(charge):
                if x >= charge.x:
                    T += charge.valeur

        return T

    def moment_flechissant(self, x: float) -> float:
        """
        M(x) = somme des moments des actions situées à gauche de x,
        calculés par rapport à la section x.
        """
        self.solveur_equilibre.assurer_reactions_calculees()

        M = 0.0

        for reaction in self.reactions:
            if reaction.direction == "y" and x >= reaction.x:
                M += reaction.valeur * (x - reaction.x)

        for reaction in self.reactions:
            if reaction.direction == "moment" and x >= reaction.x:
                M += reaction.valeur

        for charge in self.poutre.charges:
            if est_force_ponctuelle_2d(charge):
                if x >= charge.x:
                    M += charge.force_y() * (x - charge.x)

            elif est_charge_repartie_uniforme(charge):
                if x < charge.debut:
                    pass
                elif x <= charge.fin:
                    longueur_chargee = x - charge.debut
                    M += charge.q * longueur_chargee**2 / 2
                else:
                    longueur_totale = charge.fin - charge.debut
                    force_equivalente = charge.q * longueur_totale
                    position_equivalente = (charge.debut + charge.fin) / 2
                    M += force_equivalente * (x - position_equivalente)

            elif est_ancienne_force_verticale(charge):
                if x >= charge.x:
                    M += charge.valeur * (x - charge.x)

        return M
