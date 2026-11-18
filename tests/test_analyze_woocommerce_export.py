from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from analyze_woocommerce_export import (  # noqa: E402
    build_analysis,
    from_wp_all_export,
    is_wp_all_export,
    split_attribute_values,
)
from export_to_baselinker_csv import build_manufacturer_data_by_producer  # noqa: E402


LABELS = {
    "Atrybut Produktu: Producent": "Producent",
    "Atrybut Produktu: Dane Producenta": "Dane producenta",
    "Atrybut Produktu: Podtyp Produktu": "Podtyp produktu",
    "Atrybut Produktu: Przekrój Żyły [Mm]": "Przekrój żyły [mm]",
}


def pull_frame() -> pd.DataFrame:
    """Dwa wiersze w ukladzie pliku z woo_api_cli.py pull (WP All Export)."""
    return pd.DataFrame(
        [
            {
                "id": "11",
                "Title": "Gniazdo podwojne biale Simon 10",
                "SKU": "CGP2.01/11/KON",
                "Product Type": "simple",
                "Kategorie produktów": "Gniazdka i Łączniki>Gniazdka Kontakt Simon>Kontakt Simon 10",
                "Product Tags": "Gniazdo Simon 10|Kolor - Simon",
                "hwp_product_gtin": "5902787826178",
                "Atrybut Produktu: Producent": "Kontakt Simon",
                "Atrybut Produktu: Dane Producenta": "KONTAKT-SIMON, Bestwińska 21 43-500 Czechowice-Dziedzice, info@kontakt-simon.com.pl",
                "Atrybut Produktu: Podtyp Produktu": "Antenowe|Komputerowe",
                "Atrybut Produktu: Przekrój Żyły [Mm]": "",
            },
            {
                "id": "12",
                "Title": "Przewod YDY 3x2,5",
                "SKU": "YDY325/EKB",
                "Product Type": "simple",
                "Kategorie produktów": "Kable>Kable instalacyjne|Promocje",
                "Product Tags": "",
                "hwp_product_gtin": "",
                "Atrybut Produktu: Producent": "Polski Producent",
                "Atrybut Produktu: Dane Producenta": "",
                "Atrybut Produktu: Podtyp Produktu": "",
                "Atrybut Produktu: Przekrój Żyły [Mm]": "2,5",
            },
        ]
    )


class AnalyzeWooCommerceExportTest(unittest.TestCase):
    def test_detects_pull_layout(self) -> None:
        self.assertTrue(is_wp_all_export(pull_frame()))
        self.assertFalse(is_wp_all_export(from_wp_all_export(pull_frame(), LABELS)))

    def test_pull_file_gives_same_knowledge_shape_as_native_csv(self) -> None:
        analysis = build_analysis(from_wp_all_export(pull_frame(), LABELS))

        categories = {row["category"]: row["count"] for row in analysis["categories"]}
        self.assertEqual(categories["Gniazdka i Łączniki"], 1)
        self.assertEqual(categories["Gniazdka i Łączniki > Gniazdka Kontakt Simon > Kontakt Simon 10"], 1)
        self.assertEqual(categories["Kable"], 1)
        self.assertEqual(categories["Promocje"], 1)

        attributes = {row["attribute"]: row for row in analysis["attributes"]}
        self.assertEqual(attributes["Podtyp produktu"]["top_values"], "Antenowe (1); Komputerowe (1)")
        self.assertEqual(attributes["Przekrój żyły [mm]"]["top_values"], "2,5 (1)")
        self.assertEqual(attributes["Producent"]["products_count"], 2)
        self.assertEqual({row["subtype"] for row in analysis["subtypes"]}, {"Antenowe", "Komputerowe"})
        self.assertEqual(analysis["summary"]["without_ean"], 1)

    def test_manufacturer_data_survives_pull_layout(self) -> None:
        analysis = build_analysis(from_wp_all_export(pull_frame(), LABELS))
        knowledge = {"top_attribute_values_by_category": analysis["top_attribute_values_by_category"]}

        data = build_manufacturer_data_by_producer(knowledge)

        self.assertEqual(
            data["kontakt simon"],
            "KONTAKT-SIMON, Bestwińska 21 43-500 Czechowice-Dziedzice, info@kontakt-simon.com.pl",
        )

    def test_native_csv_escaped_comma_stays_in_value(self) -> None:
        self.assertEqual(split_attribute_values("2\\,5, 4"), ["2,5", "4"])
        self.assertEqual(split_attribute_values("Biały | Czarny"), ["Biały", "Czarny"])


if __name__ == "__main__":
    unittest.main()
