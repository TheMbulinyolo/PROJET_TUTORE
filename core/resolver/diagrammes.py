from typing import Dict

from .resultats import ResultatDiagrammes


class GenerateurDiagrammes2D:
    """
    Responsable de la génération des valeurs numériques
    pour les diagrammes N(x), T(x), M(x).
    """

    def __init__(self, poutre, calcul_efforts):
        self.poutre = poutre
        self.calcul_efforts = calcul_efforts

    def generer_diagrammes(self, nb_points: int = 101) -> ResultatDiagrammes:
        if nb_points < 2:
            raise ValueError("nb_points doit être supérieur ou égal à 2.")

        L = self.poutre.longueur
        xs = [i * L / (nb_points - 1) for i in range(nb_points)]

        N = [self.calcul_efforts.effort_normal(x) for x in xs]
        T = [self.calcul_efforts.effort_tranchant(x) for x in xs]
        M = [self.calcul_efforts.moment_flechissant(x) for x in xs]

        return ResultatDiagrammes(x=xs, N=N, T=T, M=M)

    def valeurs_maximales(self, nb_points: int = 1001) -> Dict[str, float]:
        diagrammes = self.generer_diagrammes(nb_points=nb_points)

        index_nmax = max(range(len(diagrammes.N)), key=lambda i: abs(diagrammes.N[i]))
        index_tmax = max(range(len(diagrammes.T)), key=lambda i: abs(diagrammes.T[i]))
        index_mmax = max(range(len(diagrammes.M)), key=lambda i: abs(diagrammes.M[i]))

        return {
            "Nmax": diagrammes.N[index_nmax],
            "x_Nmax": diagrammes.x[index_nmax],
            "Tmax": diagrammes.T[index_tmax],
            "x_Tmax": diagrammes.x[index_tmax],
            "Mmax": diagrammes.M[index_mmax],
            "x_Mmax": diagrammes.x[index_mmax],
        }
