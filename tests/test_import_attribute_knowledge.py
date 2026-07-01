from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from import_attribute_knowledge import (  # noqa: E402
    canonicalize_attribute,
    parse_sheet,
    split_outside_parentheses,
)


class ImportAttributeKnowledgeTests(unittest.TestCase):
    def test_splits_attribute_lists_but_not_examples_in_parentheses(self) -> None:
        self.assertEqual(
            split_outside_parentheses("Moc, Strumień świetlny, Barwa"),
            ["Moc", "Strumień świetlny", "Barwa"],
        )
        self.assertEqual(
            split_outside_parentheses("Test (auto, manualny, brak)"),
            ["Test (auto, manualny, brak)"],
        )

    def test_parses_category_roles_and_infers_order_column(self) -> None:
        df = pd.DataFrame(
            [
                ["Oświetlenie wewnętrzne", "", "", ""],
                ["Plafony", "", "", ""],
                [
                    "Atrybuty do filtrowania",
                    "Atrybuty obowiązkowe do opisu",
                    "Atrybuty opcjonalne do opisu",
                    "",
                ],
                ["Producent", "Producent", "Napięcie znamionowe", "Producent"],
                ["Stopień ochrony", "Dane producenta", "Seria", "Dane producenta"],
                ["", "", "", ""],
            ]
        )

        sections = parse_sheet(df, "Oświetlenie wewnętrzne")

        self.assertEqual(len(sections), 1)
        self.assertEqual(sections[0]["category"], "Plafony")
        self.assertEqual(sections[0]["filters"], ["Producent", "Stopień ochrony"])
        self.assertEqual(sections[0]["order_context"], "unspecified")
        self.assertEqual(sections[0]["ordered_attributes"], ["Producent", "Dane producenta"])

    def test_parses_split_headers_without_absorbing_next_category(self) -> None:
        df = pd.DataFrame(
            [
                ["Oświetlenie przemysłowe", "", "", "", ""],
                ["High Bay", "Atrybuty obowiązkowe do opisu", "Atrybuty opcjonalne do opisu", "", ""],
                ["Atrybuty do filtrowania", "Producent", "Czujnik ruchu", "Producent", ""],
                ["Moc [W]", "Moc [W]", "Seria", "Moc [W]", ""],
                ["Oprawy hermetyczne", "Atrybuty obowiązkowe do opisu", "Atrybuty opcjonalne do opisu", "", ""],
                ["Atrybuty do filtrowania", "Producent", "Seria", "Producent", ""],
                ["Stopień ochrony", "Stopień ochrony", "Kolor", "Stopień ochrony", ""],
            ]
        )

        sections = parse_sheet(df, "Przemysłowe")

        self.assertEqual([section["category"] for section in sections], ["High Bay", "Oprawy hermetyczne"])
        self.assertEqual(sections[0]["required_description"], ["Producent", "Moc [W]"])
        self.assertNotIn("Oprawy hermetyczne", sections[0]["filters"])
        self.assertEqual(sections[1]["filters"], ["Stopień ochrony"])

    def test_maps_renamed_catalog_attributes_and_marks_new_ones(self) -> None:
        catalog = {
            "napiecie v": "Napięcie [V]",
            "stopien ochrony ip": "Stopień ochrony [IP]",
        }
        names = set(catalog.values())

        voltage = canonicalize_attribute("Napięcie znamionowe", catalog, names)
        ik = canonicalize_attribute("Klasa odporności IK", catalog, names)

        self.assertEqual(voltage["canonical_name"], "Napięcie [V]")
        self.assertEqual(voltage["match_type"], "CATALOG_ALIAS")
        self.assertEqual(ik["canonical_name"], "Stopień odporności [IK]")
        self.assertEqual(ik["match_type"], "NEW_ATTRIBUTE")


if __name__ == "__main__":
    unittest.main()
