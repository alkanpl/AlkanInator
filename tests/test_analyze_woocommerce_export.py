from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from analyze_woocommerce_export import (  # noqa: E402
    build_analysis,
    from_wp_all_export,
    local_attribute_labels,
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

    def test_ean_column_from_bridge_2_1_1(self) -> None:
        frame = pull_frame().drop(columns=["hwp_product_gtin"])
        frame["EAN"] = ["5902787826178", ""]
        analysis = build_analysis(from_wp_all_export(frame, LABELS))
        self.assertEqual(analysis["summary"]["without_ean"], 1)
        self.assertEqual(analysis["summary"]["invalid_ean"], 0)

    def test_local_attribute_names_from_product_attributes(self) -> None:
        frame = pull_frame()
        frame["Atrybut Produktu: Kompatybilny Z"] = ["Simon 10", ""]
        frame["Product Attributes"] = [
            'a:2:{s:12:"pa_producent";a:6:{s:4:"name";s:12:"pa_producent";s:5:"value";s:0:"";s:8:"position";i:0;'
            's:10:"is_visible";i:1;s:12:"is_variation";i:0;s:11:"is_taxonomy";i:1;}s:14:"kompatybilny-z";a:6:'
            '{s:4:"name";s:14:"Kompatybilny z";s:5:"value";s:8:"Simon 10";s:8:"position";i:1;s:10:"is_visible";i:1;'
            's:12:"is_variation";i:0;s:11:"is_taxonomy";i:0;}}',
            "",
        ]
        labels = local_attribute_labels(frame)
        self.assertEqual(labels, {"Atrybut Produktu: Kompatybilny Z": "Kompatybilny z"})
        analysis = build_analysis(from_wp_all_export(frame, {**labels, **LABELS}))
        self.assertIn("Kompatybilny z", {row["attribute"] for row in analysis["attributes"]})

    def test_native_csv_escaped_comma_stays_in_value(self) -> None:
        self.assertEqual(split_attribute_values("2\\,5, 4"), ["2,5", "4"])
        self.assertEqual(split_attribute_values("Biały | Czarny"), ["Biały", "Czarny"])


if __name__ == "__main__":
    unittest.main()
