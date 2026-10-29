from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fix_simon_attributes_to_shop_model import (  # noqa: E402
    fix_features,
    load_dictionary,
    process_rows,
    symbol_base,
    validate_features,
)
from shop_attribute_values import ShopVocabulary, load_rules as load_value_rules  # noqa: E402


class SymbolBaseTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()

    def test_strips_supplier_suffix_and_color(self) -> None:
        self.assertEqual(symbol_base("TR1/142/KON", self.config), "TR1")
        self.assertEqual(symbol_base("DW7.01/X/11/KON", self.config), "DW7.01/X")
        self.assertEqual(symbol_base("TW6/2.01/X/142/KON", self.config), "TW6/2.01/X")

    def test_bare_symbol_without_color_is_unchanged(self) -> None:
        self.assertEqual(symbol_base("G22/KON", self.config), "G22")


class FixFeaturesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()

    def test_frame_gets_krotnosc_and_loses_subtype(self) -> None:
        features = {
            "Typ produktu": "Ramki",
            "Podtyp produktu": "Ramka 1- krotna",
            "Seria": "Simon 55",
            "Stopień ochrony [IP]": "IP20/IP44",
            "Sposób montażu": "montaż podtynkowy",
            "Wysokość [mm]": "90",
        }
        result = fix_features("TR1/142/KON", "Ramka 1-krotna LINE kaszmirowa Simon 55 TR1/142", features, self.config)
        self.assertEqual(result.after["Typ produktu"], "Ramki")
        self.assertNotIn("Podtyp produktu", result.after)
        self.assertEqual(result.after["Krotność"], "1x")
        self.assertEqual(result.after["Materiał"], "Tworzywo sztuczne")
        self.assertNotIn("Sposób montażu", result.after)
        self.assertEqual(result.after["Wysokość [mm]"], "90")
        self.assertEqual(result.problems, [])

    def test_supplier_type_names_are_replaced_by_shop_terms(self) -> None:
        features = {
            "Typ produktu": "Gniazdka",
            "Podtyp produktu": "Gniazdo komputerowe pojedyncze RJ45 kategoria 6, z przesłoną przeciwkurzową",
            "Seria": "Simon 55",
            "Stopień ochrony [IP]": "IP20",
            "Sposób montażu": "montaż podtynkowy",
        }
        result = fix_features("T61.01/142/KON", "Gniazdo komputerowe pojedyncze RJ45 kat. 6 kaszmirowe", features, self.config)
        self.assertEqual(result.after["Typ produktu"], "Gniazda")
        self.assertEqual(result.after["Podtyp produktu"], "Komputerowe")
        self.assertEqual(result.after["Stopień ochrony [IP]"], "IP 20")
        self.assertEqual(result.after["Sposób montażu"], "Podtynkowy")
        self.assertEqual(result.problems, [])

    def test_keys_without_shop_type_lose_type_and_subtype(self) -> None:
        features = {"Typ produktu": "Inne", "Podtyp produktu": "Klawisze", "Seria": "Simon 54", "Stopień ochrony [IP]": "IP20"}
        result = fix_features("DKW1/42/KON", "Klawisz do mechanizmów SW1M kaszmirowy DKW1/42", features, self.config)
        self.assertNotIn("Typ produktu", result.after)
        self.assertNotIn("Podtyp produktu", result.after)
        self.assertEqual(result.problems, [])

    def test_name_rule_used_when_symbol_unknown(self) -> None:
        features = {"Typ produktu": "Ściemniacze", "Podtyp produktu": "cokolwiek", "Seria": "Simon 55 GO", "Stopień ochrony [IP]": "IP20"}
        result = fix_features("XYZ.99/142/KON", "Ściemniacz do LED nieznany", features, self.config)
        self.assertEqual(result.after["Typ produktu"], "Sterowniki i automatyka")
        self.assertEqual(result.after["Podtyp produktu"], "Ściemniacze")
        self.assertTrue(result.source.startswith("regula nazwy"))

    def test_unknown_frame_takes_krotnosc_from_name(self) -> None:
        features = {"Typ produktu": "Ramki", "Podtyp produktu": "Ramka 3- krotna", "Seria": "Simon 55"}
        result = fix_features("NOWA/142/KON", "Ramka 3-krotna nowa kaszmirowa", features, self.config)
        self.assertEqual(result.after["Krotność"], "3x")
        self.assertNotIn("Podtyp produktu", result.after)
        self.assertEqual(result.problems, [])

    def test_material_kept_for_non_frames(self) -> None:
        features = {"Typ produktu": "Łączniki", "Podtyp produktu": "Łącznik jednobiegunowy", "Seria": "Simon 55", "Materiał": "tworzywo sztuczne, PC, bezhalogenowe"}
        result = fix_features("TW1.01/142/KON", "Łącznik jednobiegunowy pojedynczy kaszmirowy", features, self.config)
        self.assertEqual(result.after["Podtyp produktu"], "Pojedyncze")
        self.assertEqual(result.after["Materiał"], "tworzywo sztuczne, PC, bezhalogenowe")

    def test_untouched_attributes_survive(self) -> None:
        features = {"Typ produktu": "Gniazdka", "Podtyp produktu": "x", "Klasa ETIM": "EC000125", "Seria": "Simon 55"}
        result = fix_features("T61.01/142/KON", "Gniazdo komputerowe", features, self.config)
        self.assertEqual(result.after["Klasa ETIM"], "EC000125")


class MultiTermSubtypeTest(unittest.TestCase):
    def test_sibling_multi_subtype_uses_woo_separator(self) -> None:
        config = load_dictionary()
        features = {"Typ produktu": "Gniazda", "Podtyp produktu": "Antenowe, Komputerowe", "Seria": "Simon 55"}
        result = fix_features("TASFRJ45.01/142/KON", "Gniazdo antenowe F + RJ45 kaszmirowe", features, config)
        self.assertEqual(result.after["Podtyp produktu"], "Antenowe|Komputerowe")
        self.assertEqual(result.problems, [])


class ShopValueCheckTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()
        self.vocabulary = ShopVocabulary()
        for attribute, terms in {
            "Typ produktu": ["Gniazda", "Łączniki"],
            "Podtyp produktu": ["Komputerowe", "Pojedyncze"],
            "Seria": ["Simon 55"],
            "Materiał": ["Tworzywo sztuczne"],
            "Kolor producenta": ["Kaszmir"],
            "Typ zacisków": ["Gwintowe", "Bezgwintowe"],
            "Klasa energetyczna": ["A"],
            "Stopień ochrony [IP]": ["IP 20"],
        }.items():
            for term in terms:
                self.vocabulary.add(attribute, term)
        self.vocabulary.add("Podtyp produktu", "Komputerowe", parent="Gniazda")
        self.vocabulary.add("Podtyp produktu", "Pojedyncze", parent="Łączniki")
        self.rules = load_value_rules()

    def test_new_values_are_mapped_to_shop_terms_or_removed(self) -> None:
        rows = [{
            "sku": "TW1.01/142/KON",
            "name": "Łącznik jednobiegunowy pojedynczy kaszmirowy",
            "features": json.dumps({
                "Typ produktu": "Łączniki", "Podtyp produktu": "Pojedyncze", "Seria": "Simon 55",
                "Materiał": "tworzywo sztuczne, PC, bezhalogenowe", "Kolor producenta": "Kaszmirowy",
                "Typ zacisków": "szybkozłącza", "Klasa energetyczna": "nie podlega",
            }),
        }]
        results, _ = process_rows(rows, self.config, vocabulary=self.vocabulary, value_rules=self.rules)
        after = json.loads(rows[0]["features"])
        self.assertEqual(after["Materiał"], "Tworzywo sztuczne")
        self.assertEqual(after["Kolor producenta"], "Kaszmir")
        self.assertEqual(after["Typ zacisków"], "Bezgwintowe")
        self.assertNotIn("Klasa energetyczna", after)
        self.assertEqual(results[0].changes["Klasa energetyczna"], ("nie podlega", None))
        kinds = {check.attribute: check.kind for check in results[0].value_checks}
        self.assertEqual(kinds["Klasa energetyczna"], "usunieta (brak odpowiednika)")

    def test_unknown_new_value_is_removed_unless_kept(self) -> None:
        features = {"Typ produktu": "Gniazda", "Podtyp produktu": "Komputerowe", "Seria": "Simon 55", "Typ zacisków": "magnetyczne"}
        rows = [{"sku": "T61.01/142/KON", "name": "Gniazdo komputerowe", "features": json.dumps(features)}]
        results, _ = process_rows(rows, self.config, vocabulary=self.vocabulary, value_rules=self.rules)
        self.assertNotIn("Typ zacisków", json.loads(rows[0]["features"]))
        check = next(check for check in results[0].value_checks if check.attribute == "Typ zacisków")
        self.assertEqual(check.status, "DO_SPRAWDZENIA")
        rows = [{"sku": "T61.01/142/KON", "name": "Gniazdo komputerowe", "features": json.dumps(features)}]
        process_rows(rows, self.config, vocabulary=self.vocabulary, value_rules=self.rules, keep_new_values=True)
        self.assertEqual(json.loads(rows[0]["features"])["Typ zacisków"], "magnetyczne")


class UnknownAttributesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()

    def test_unknown_attributes_are_removed_and_reported(self) -> None:
        rows = [{
            "sku": "T61.01/142/KON",
            "name": "Gniazdo komputerowe pojedyncze",
            "features": json.dumps({"Typ produktu": "Gniazda", "Podtyp produktu": "Komputerowe", "Seria": "Simon 55", "Klasa ETIM": "EC000125", "Wysokość [mm]": "75"}),
        }]
        results, unknown = process_rows(rows, self.config)
        after = json.loads(rows[0]["features"])
        self.assertNotIn("Klasa ETIM", after)
        self.assertNotIn("Wysokość [mm]", after)
        self.assertEqual(after["Seria"], "Simon 55")
        self.assertEqual(unknown["Klasa ETIM"]["produktow"], 1)
        self.assertEqual(unknown["Wysokość [mm]"]["sugerowany_atrybut_woo"], "Wysokość")
        self.assertIsNone(unknown["Klasa ETIM"]["sugerowany_atrybut_woo"])
        self.assertEqual(results[0].changes["Klasa ETIM"], ("EC000125", None))

    def test_keep_unknown_flag_leaves_features(self) -> None:
        rows = [{"sku": "T61.01/142/KON", "name": "Gniazdo komputerowe", "features": json.dumps({"Typ produktu": "Gniazda", "Klasa ETIM": "EC000125"})}]
        _, unknown = process_rows(rows, self.config, keep_unknown=True)
        self.assertIn("Klasa ETIM", json.loads(rows[0]["features"]))
        self.assertEqual(unknown, {})

    def test_approved_extra_attribute_is_kept(self) -> None:
        config = dict(self.config)
        config["approved_extra_attributes"] = ["Klasa ETIM"]
        rows = [{"sku": "T61.01/142/KON", "name": "Gniazdo komputerowe", "features": json.dumps({"Typ produktu": "Gniazda", "Klasa ETIM": "EC000125"})}]
        _, unknown = process_rows(rows, config)
        self.assertIn("Klasa ETIM", json.loads(rows[0]["features"]))
        self.assertEqual(unknown, {})


class ValidateFeaturesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()

    def test_reports_type_outside_dictionary(self) -> None:
        self.assertIn("Typ spoza slownika sklepu: Gniazdka", validate_features({"Typ produktu": "Gniazdka"}, self.config))

    def test_reports_subtype_outside_dictionary(self) -> None:
        problems = validate_features({"Typ produktu": "Gniazda", "Podtyp produktu": "Antenowe|Foo"}, self.config)
        self.assertEqual(problems, ["Podtyp spoza slownika dla Gniazda: Foo"])

    def test_comma_joined_subtypes_are_one_unknown_term(self) -> None:
        # Kaszmir 2026-09-15: "Antenowe, Komputerowe" utworzylo w Woo jeden nowy termin.
        problems = validate_features({"Typ produktu": "Gniazda", "Podtyp produktu": "Antenowe, Komputerowe"}, self.config)
        self.assertEqual(problems, ["Podtyp spoza slownika dla Gniazda: Antenowe, Komputerowe"])

    def test_frame_requires_krotnosc(self) -> None:
        self.assertIn("Ramka bez poprawnej Krotnosci", validate_features({"Typ produktu": "Ramki"}, self.config))


if __name__ == "__main__":
    unittest.main()
