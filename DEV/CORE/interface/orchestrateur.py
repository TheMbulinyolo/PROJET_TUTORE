"""Orchestration des calculs sans logique de lecture ou d'ecriture JSON."""

from __future__ import annotations

from typing import Any

from CORE.resolver.resolver import ResolverIsostatique2D

from .parseur import DonneesCalcul, calculer_nombre_points


def executer_calcul(donnees: DonneesCalcul) -> dict[str, Any]:
    resolver = ResolverIsostatique2D(donnees.poutre)
    options = donnees.options
    sortie: dict[str, Any] = {
        "statut": "OK",
        "message": "Calcul reussi",
        "isostatique": donnees.poutre.is_isostatique_2d(),
        "unites": {
            "longueur": donnees.unite_longueur,
            "force": donnees.unite_force,
        },
    }

    # Les efforts et diagrammes ont besoin des reactions calculees.
    calcul_mecanique_requis = (
        options.calcul_reactions
        or options.calcul_efforts
        or options.calcul_diagrammes
    )
    if calcul_mecanique_requis:
        resultat_equilibre = resolver.resoudre_reactions()
        sortie["reactions"] = resultat_equilibre.reactions
    else:
        sortie["reactions"] = {}

    if options.calcul_efforts:
        position = (
            options.position_efforts
            if options.position_efforts is not None
            else donnees.poutre.longueur / 2
        )
        sortie["efforts_position"] = {
            "x": position,
            "N": resolver.effort_normal(position),
            "T": resolver.effort_tranchant(position),
            "M": resolver.moment_flechissant(position),
        }
    else:
        sortie["efforts_position"] = None

    if options.calcul_diagrammes:
        nombre_points = calculer_nombre_points(
            donnees.poutre.longueur,
            options.pas_discretisation,
        )
        diagrammes = resolver.generer_diagrammes(nb_points=nombre_points)
        sortie["diagrammes"] = {
            "x": diagrammes.x,
            "N": diagrammes.N,
            "T": diagrammes.T,
            "M": diagrammes.M,
        }
        sortie["maxima"] = resolver.valeurs_maximales(
            nb_points=max(nombre_points, 1001)
        )
    else:
        sortie["diagrammes"] = None
        sortie["maxima"] = None

    return sortie
