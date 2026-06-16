import json
import tempfile
import unittest
from pathlib import Path

from CORE.interface.orchestrateur import executer_calcul
from CORE.interface.parseur import ErreurValidation, lire_et_parser, parser_donnees


CAS_SIMPLE = {
    "poutre": {"longueur": 6.0, "unite": "m"},
    "appuis": [
        {"id": "A", "type": "pivot", "position": 0.0},
        {"id": "B", "type": "rouleau", "position": 6.0},
    ],
    "charges": [
        {
            "id": "P1",
            "type": "ponctuelle",
            "position": 3.0,
            "valeur": -10.0,
            "unite": "kN",
        },
        {
            "id": "Q1",
            "type": "repartie_uniforme",
            "debut": 1.0,
            "fin": 5.0,
            "valeur": -2.0,
            "unite": "kN/m",
        },
    ],
    "options": {
        "calcul_reactions": True,
        "calcul_efforts": True,
        "calcul_diagrammes": True,
        "pas_discretisation": 0.1,
    },
}


class InterfaceExcelTests(unittest.TestCase):
    def test_calcul_complet(self):
        resultat = executer_calcul(parser_donnees(CAS_SIMPLE))

        self.assertEqual(resultat["statut"], "OK")
        self.assertTrue(resultat["isostatique"])
        self.assertAlmostEqual(resultat["reactions"]["Ay"], 9.0)
        self.assertAlmostEqual(resultat["reactions"]["By"], 9.0)
        self.assertEqual(resultat["unites"]["force"], "kN")
        self.assertEqual(len(resultat["diagrammes"]["x"]), 61)

    def test_fichier_json_utf8(self):
        with tempfile.TemporaryDirectory() as dossier:
            chemin = Path(dossier) / "input.json"
            chemin.write_text(json.dumps(CAS_SIMPLE), encoding="utf-8")
            donnees = lire_et_parser(chemin)
            self.assertEqual(donnees.poutre.longueur, 6.0)

    def test_position_appui_hors_poutre(self):
        donnees = json.loads(json.dumps(CAS_SIMPLE))
        donnees["appuis"][1]["position"] = 8.0

        with self.assertRaises(ErreurValidation):
            parser_donnees(donnees)


if __name__ == "__main__":
    unittest.main()
