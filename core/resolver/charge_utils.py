from typing import Tuple


def position_resultante(charge) -> float:
    """
    Retourne la position d'application équivalente d'une charge.

    Compatible avec :
    - ForcePonctuelle : attribut x
    - ForceRepartieUniforme : position_resultante()
    - anciens modèles éventuels : resultant_position()
    """
    if hasattr(charge, "position_resultante"):
        return charge.position_resultante()

    if hasattr(charge, "resultant_position"):
        return charge.resultant_position()

    if hasattr(charge, "x"):
        return charge.x

    raise AttributeError("Impossible de trouver la position de la charge.")


def composantes_charge(charge) -> Tuple[float, float, float]:
    """
    Retourne Fx, Fy, xF pour une charge donnée.

    Cas 1 : ForcePonctuelle 2D
        - force_x()
        - force_y()
        - x

    Cas 2 : ForceRepartieUniforme verticale
        - force_resultante()
        - position_resultante()

    Cas 3 : ancien modèle de force ponctuelle verticale
        - valeur
        - x
    """
    if hasattr(charge, "force_x") and hasattr(charge, "force_y"):
        return charge.force_x(), charge.force_y(), charge.x

    if hasattr(charge, "force_resultante"):
        return 0.0, charge.force_resultante(), position_resultante(charge)

    if hasattr(charge, "valeur") and hasattr(charge, "x"):
        return 0.0, charge.valeur, charge.x

    raise TypeError("Type de charge non reconnu par le resolver.")


def est_force_ponctuelle_2d(charge) -> bool:
    return hasattr(charge, "force_x") and hasattr(charge, "force_y") and hasattr(charge, "x")


def est_charge_repartie_uniforme(charge) -> bool:
    return hasattr(charge, "q") and hasattr(charge, "debut") and hasattr(charge, "fin")


def est_ancienne_force_verticale(charge) -> bool:
    return hasattr(charge, "valeur") and hasattr(charge, "x")
