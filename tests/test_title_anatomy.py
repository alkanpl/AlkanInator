from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from optimize_titles import optimize_titles_for_dataframe  # noqa: E402
from title_anatomy import build_title_from_anatomy, duplicate_titles_report  # noqa: E402


ANATOMY_CONFIG = {
    "enabled": True,
    "default_primary_keyword_source": "primary_keyword",
    "primary_keyword_fallback_fields": ["attr_typ"],
    "sku_fields": ["Kod"],
    "producer_fields": ["Producent"],
    "accepted_status_values": ["approved"],
    "rules": [
        {
            "normalized_product_type": "panel led",
            "status": "approved",
            "required_title_attributes": ["seria", "moc", "barwa", "kolor"],
            "optional_title_attributes": [],
            "excluded_if_redundant": {"kolor": ["barwa_swiatla"]},
        }
    ],
}

TITLE_CONFIG = {
    "title_template": "{typ} {seria} {moc} {barwa} {kolor} {producent}",
    "append_sku_after_producer": True,
    "sku_columns": ["Kod"],
    "max_title_length": 120,
    "required_fields": ["typ", "moc", "producent"],
    "default_values": {"producent": "Kanlux"},
}


class TitleAnatomyTest(unittest.TestCase):
    def test_builds_title_from_primary_keyword_and_ordered_attributes(self) -> None:
        row = pd.Series({"primary_keyword": "panel led sufitowy", "attr_typ": "Panel LED", "Kod": "12345", "Producent": "Kanlux"})
        values = {"typ": "Panel LED", "seria": "BLINGO", "moc": "36W", "barwa": "4000K", "kolor": "biały", "producent": "Kanlux"}
        result = build_title_from_anatomy(row, ANATOMY_CONFIG, values)
        self.assertIsNotNone(result)
        assert result is not None
        self.assertEqual(result.title, "Panel led sufitowy BLINGO 36W 4000K biały Kanlux 12345")
        self.assertEqual(result.used_attributes, ["seria", "moc", "barwa", "kolor"])

    def test_falls_back_to_product_type_when_primary_keyword_missing(self) -> None:
        row = pd.Series({"attr_typ": "Panel LED", "Kod": "12345", "Producent": "Kanlux"})
        values = {"typ": "Panel LED", "seria": "BLINGO", "moc": "36W", "barwa": "4000K", "kolor": "biały", "producent": "Kanlux"}
        result = build_title_from_anatomy(row, ANATOMY_CONFIG, values)
        self.assertIsNotNone(result)
        assert result is not None
        self.assertTrue(result.title.startswith("Panel LED BLINGO"))
        # Bez primary_keyword i bez title_keyword reguly: fallback na typ produktu.
        self.assertIn("title_anatomy_missing_primary_keyword", result.warnings)

    def test_skips_redundant_color_when_light_attribute_contains_same_color(self) -> None:
        row = pd.Series({"primary_keyword": "panel led", "attr_typ": "Panel LED", "Kod": "12345", "Producent": "Kanlux"})
        values = {
            "typ": "Panel LED",
            "seria": "BLINGO",
            "moc": "36W",
            "barwa": "biała",
            "barwa_swiatla": "światło biała",
            "kolor": "biała",
            "producent": "Kanlux",
        }
        result = build_title_from_anatomy(row, ANATOMY_CONFIG, values)
        self.assertIsNotNone(result)
        assert result is not None
        self.assertEqual(result.title.count("biała"), 1)
        self.assertIn("kolor", result.skipped_attributes)

    def test_optimize_titles_uses_legacy_template_when_rule_missing(self) -> None:
        df = pd.DataFrame([{"Nazwa": "Old", "attr_typ": "Brak reguly", "Kod": "12345", "Producent": "Kanlux"}])
        result = optimize_titles_for_dataframe(df, "Nazwa", TITLE_CONFIG, ANATOMY_CONFIG)
        self.assertEqual(result.loc[0, "title_source"], "legacy_template")
        self.assertIn("title_anatomy_missing_accepted_rule", result.loc[0, "warnings"])

    def test_duplicate_final_titles_are_reported(self) -> None:
        df = pd.DataFrame(
            [
                {"Nazwa": "Old A", "primary_keyword": "panel led", "attr_typ": "Panel LED", "Kod": "12345", "Producent": "Kanlux"},
                {"Nazwa": "Old B", "primary_keyword": "panel led", "attr_typ": "Panel LED", "Kod": "12345", "Producent": "Kanlux"},
            ]
        )
        result = optimize_titles_for_dataframe(df, "Nazwa", TITLE_CONFIG, ANATOMY_CONFIG)
        self.assertIn("duplicate_final_title:2", result.loc[0, "warnings"])
        report = duplicate_titles_report(result)
        self.assertGreaterEqual(len(report), 2)


if __name__ == "__main__":
    unittest.main()
