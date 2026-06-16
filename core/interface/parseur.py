"""Lecture, validation et conversion des donnees JSON provenant d'Excel."""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from core.models.charges import ForcePonctuelle, ForceRepartieUniforme
from core.models.poutre import Poutre
from core.models.supports import Encastrement, Rotule, Rouleau


class ErreurValidation(ValueError):
    """Erreur de donnees d'entree destinee a etre affichee dans Excel."""


@dataclass(frozen=True)
class OptionsCalcul:
    calcul_reactions: bool = True
    calcul_efforts: bool = True
    calcul_diagrammes: bool = True
    pas_discretisation: float = 0.1
    position_efforts: float | None = None


@dataclass(frozen=True)
class DonneesCalcul:
    poutre: Poutre
    options: OptionsCalcul
    unite_longueur: str
    unite_force: str


def lire_json(chemin: str | Path) -> dict[str, Any]:
    chemin = Path(chemin)
    if not chemin.exists():
        raise ErreurValidation(f"Fichier d'entree introuvable : {chemin}")

    try:
        with chemin.open("r", encoding="utf-8-sig") as fichier:
            donnees = json.load(fichier)
    except json.JSONDecodeError as exc:
        raise ErreurValidation(
            f"JSON invalide a la ligne {exc.lineno}, colonne {exc.colno}."
        ) from exc
    except OSError as exc:
        raise ErreurValidation(f"Impossible de lire le fichier d'entree : {exc}") from exc

    if not isinstance(donnees, dict):
        raise ErreurValidation("La racine du JSON doit etre un objet.")
    return donnees


def lire_et_parser(chemin: str | Path) -> DonneesCalcul:
    return parser_donnees(lire_json(chemin))


def parser_donnees(donnees: dict[str, Any]) -> DonneesCalcul:
    poutre_json = _objet_requis(donnees, "poutre")
    longueur = _nombre_positif(poutre_json.get("longueur"), "poutre.longueur")
    unite_longueur = _texte_optionnel(poutre_json.get("unite"), "m")

    poutre = Poutre(longueur=longueur)
    _ajouter_appuis(poutre, donnees.get("appuis"))
    unite_force = _ajouter_charges(poutre, donnees.get("charges"))
    options = _parser_options(donnees.get("options"), longueur)

    if not poutre.appuis:
        raise ErreurValidation("Au moins un appui est requis.")
    if not poutre.charges:
        raise ErreurValidation("Au moins une charge est requise.")

    return DonneesCalcul(
        poutre=poutre,
        options=options,
        unite_longueur=unite_longueur,
        unite_force=unite_force,
    )


def calculer_nombre_points(longueur: float, pas: float) -> int:
    if pas <= 0:
        raise ErreurValidation("options.pas_discretisation doit etre strictement positif.")
    return max(2, min(10001, math.ceil(longueur / pas) + 1))


def _ajouter_appuis(poutre: Poutre, appuis_json: Any) -> None:
    if not isinstance(appuis_json, list):
        raise ErreurValidation("appuis doit etre une liste.")

    types_appuis = {
        "pivot": Rotule,
        "rotule": Rotule,
        "rouleau": Rouleau,
        "encastrement": Encastrement,
    }
    identifiants: set[str] = set()

    for index, element in enumerate(appuis_json):
        chemin = f"appuis[{index}]"
        if not isinstance(element, dict):
            raise ErreurValidation(f"{chemin} doit etre un objet.")

        identifiant = _texte_requis(element.get("id"), f"{chemin}.id")
        if identifiant in identifiants:
            raise ErreurValidation(f"Identifiant d'appui duplique : {identifiant}")
        identifiants.add(identifiant)

        type_appui = _normaliser_type(element.get("type"), f"{chemin}.type")
        classe_appui = types_appuis.get(type_appui)
        if classe_appui is None:
            raise ErreurValidation(
                f"{chemin}.type inconnu : {type_appui}. "
                "Valeurs acceptees : pivot, rotule, rouleau, encastrement."
            )

        position = _nombre(element.get("position"), f"{chemin}.position")
        try:
            poutre.ajouter_appui(classe_appui(nom=identifiant, x=position))
        except ValueError as exc:
            raise ErreurValidation(str(exc)) from exc


def _ajouter_charges(poutre: Poutre, charges_json: Any) -> str:
    if not isinstance(charges_json, list):
        raise ErreurValidation("charges doit etre une liste.")

    identifiants: set[str] = set()
    unite_force = "kN"

    for index, element in enumerate(charges_json):
        chemin = f"charges[{index}]"
        if not isinstance(element, dict):
            raise ErreurValidation(f"{chemin} doit etre un objet.")

        identifiant = _texte_requis(element.get("id"), f"{chemin}.id")
        if identifiant in identifiants:
            raise ErreurValidation(f"Identifiant de charge duplique : {identifiant}")
        identifiants.add(identifiant)

        type_charge = _normaliser_type(element.get("type"), f"{chemin}.type")
        valeur = _nombre(element.get("valeur"), f"{chemin}.valeur")
        unite_charge = _texte_optionnel(element.get("unite"), unite_force)

        if type_charge in {"ponctuelle", "force_ponctuelle"}:
            unite_force = unite_charge
            position = _nombre(element.get("position"), f"{chemin}.position")
            fx = _nombre_optionnel(element.get("fx"), 0.0, f"{chemin}.fx")
            fy = _nombre_optionnel(element.get("fy"), valeur, f"{chemin}.fy")
            charge = ForcePonctuelle(fx=fx, fy=fy, x=position)
        elif type_charge in {"repartie_uniforme", "uniformement_repartie"}:
            if "/" in unite_charge:
                unite_force = unite_charge.split("/", 1)[0].strip() or unite_force
            debut = _nombre(element.get("debut"), f"{chemin}.debut")
            fin = _nombre(element.get("fin"), f"{chemin}.fin")
            if fin <= debut:
                raise ErreurValidation(f"{chemin}.fin doit etre superieur a debut.")
            charge = ForceRepartieUniforme(q=valeur, debut=debut, fin=fin)
        else:
            raise ErreurValidation(
                f"{chemin}.type inconnu : {type_charge}. "
                "Valeurs acceptees : ponctuelle, repartie_uniforme."
            )

        try:
            poutre.ajouter_charge(charge)
        except ValueError as exc:
            raise ErreurValidation(str(exc)) from exc

    return unite_force


def _parser_options(options_json: Any, longueur: float) -> OptionsCalcul:
    if options_json is None:
        options_json = {}
    if not isinstance(options_json, dict):
        raise ErreurValidation("options doit etre un objet.")

    pas = _nombre_optionnel(
        options_json.get("pas_discretisation"),
        0.1,
        "options.pas_discretisation",
    )
    calculer_nombre_points(longueur, pas)

    position = options_json.get("position_efforts")
    if position is not None:
        position = _nombre(position, "options.position_efforts")
        if position < 0 or position > longueur:
            raise ErreurValidation("options.position_efforts est hors de la poutre.")

    return OptionsCalcul(
        calcul_reactions=_booleen(options_json.get("calcul_reactions"), True),
        calcul_efforts=_booleen(options_json.get("calcul_efforts"), True),
        calcul_diagrammes=_booleen(options_json.get("calcul_diagrammes"), True),
        pas_discretisation=pas,
        position_efforts=position,
    )


def _objet_requis(donnees: dict[str, Any], cle: str) -> dict[str, Any]:
    valeur = donnees.get(cle)
    if not isinstance(valeur, dict):
        raise ErreurValidation(f"{cle} doit etre un objet.")
    return valeur


def _texte_requis(valeur: Any, chemin: str) -> str:
    texte = str(valeur).strip() if valeur is not None else ""
    if not texte:
        raise ErreurValidation(f"{chemin} est requis.")
    return texte


def _texte_optionnel(valeur: Any, defaut: str) -> str:
    texte = str(valeur).strip() if valeur is not None else ""
    return texte or defaut


def _normaliser_type(valeur: Any, chemin: str) -> str:
    return (
        _texte_requis(valeur, chemin)
        .lower()
        .replace("é", "e")
        .replace("è", "e")
        .replace(" ", "_")
        .replace("-", "_")
    )


def _nombre(valeur: Any, chemin: str) -> float:
    if isinstance(valeur, bool):
        raise ErreurValidation(f"{chemin} doit etre un nombre.")
    try:
        resultat = float(valeur)
    except (TypeError, ValueError) as exc:
        raise ErreurValidation(f"{chemin} doit etre un nombre.") from exc
    if not math.isfinite(resultat):
        raise ErreurValidation(f"{chemin} doit etre un nombre fini.")
    return resultat


def _nombre_positif(valeur: Any, chemin: str) -> float:
    resultat = _nombre(valeur, chemin)
    if resultat <= 0:
        raise ErreurValidation(f"{chemin} doit etre strictement positif.")
    return resultat


def _nombre_optionnel(valeur: Any, defaut: float, chemin: str) -> float:
    return defaut if valeur in (None, "") else _nombre(valeur, chemin)


def _booleen(valeur: Any, defaut: bool) -> bool:
    if valeur is None:
        return defaut
    if isinstance(valeur, bool):
        return valeur
    raise ErreurValidation("Les options de calcul doivent etre des booleens.")
