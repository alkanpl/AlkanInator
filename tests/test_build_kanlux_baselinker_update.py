from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from build_kanlux_baselinker_update import (  # noqa: E402
    add_woo_feature_fallbacks,
    add_woo_values_to_feature_normalizer,
    apply_pipeline_titles_for_duplicate_names,
    build_control_rows,
    apply_pipeline_titles_for_mismatched_names,
    build_attribute_aliases,
    build_missing_required_rows,
    compare_feature_sets,
    duplicate_name_key,
    remaining_duplicate_name_rows,
    woo_features,
)


class ControlRowsOutsideReferenceTests(unittest.TestCase):
    def test_enriched_woo_products_outside_reference_are_included(self) -> None:
        reference = pd.DataFrame([{"Sku": "100/KAN", "hwp_product_gtin": "", "id": "1", "Title": "Ref 100"}])
        enriched = pd.DataFrame(
            [
                {"Kod": "100", "EAN": "", "new_title": "Plafon 100", "attr_typ": "plafon led"},
                {"Kod": "200", "EAN": "", "new_title": "Plafon 200", "attr_typ": "plafon led"},
            ]
        )
        accepted = pd.DataFrame(columns=["Kod", "EAN"])
        woo = pd.DataFrame(
            [
                {"SKU": "100/KAN", "id": "11", "Title": "Woo 100"},
                {"SKU": "200/KAN", "id": "22", "Title": "Woo 200"},
            ]
        )
        control, _ = build_control_rows(reference, enriched, accepted, woo, {})
        status = dict(zip(control["SKU"], control["source_match_status"]))
        self.assertEqual(status.get("100/KAN"), "MATCHED_MASTER")
        # 200 jest w master+woo, ale nie w referencji -> trafia do Woo (nie do Baselinkera).
        self.assertEqual(status.get("200/KAN"), "MATCHED_OUTSIDE_REFERENCE")
        outside = control[control["source_match_status"] == "MATCHED_OUTSIDE_REFERENCE"].iloc[0]
        self.assertEqual(outside["product_id"], "22")  # id z WooCommerce, nie z Baselinkera


class DuplicateNameDisambiguationTests(unittest.TestCase):
    def test_swaps_to_pipeline_title_when_names_collide_ignoring_code(self) -> None:
        control = pd.DataFrame(
            [
                {"pipeline_title": "Plafon ERTON E27 330mm IP44 biały Kanlux 35771"},
                {"pipeline_title": "Plafon ERTON E27 252mm IP44 biały Kanlux 35770"},
            ]
        )
        exported = [
            {"sku": "35771/KAN", "name": "Plafon ERTON E27 IP44 biały Kanlux 35771", "features": "{}"},
            {"sku": "35770/KAN", "name": "Plafon ERTON E27 IP44 biały Kanlux 35770", "features": "{}"},
        ]
        renamed = apply_pipeline_titles_for_duplicate_names(control, exported)
        self.assertEqual(len(renamed), 2)
        self.assertEqual(exported[0]["name"], "Plafon ERTON E27 330mm IP44 biały Kanlux 35771")
        self.assertEqual(exported[1]["name"], "Plafon ERTON E27 252mm IP44 biały Kanlux 35770")
        self.assertEqual(remaining_duplicate_name_rows(exported), [])

    def test_keeps_name_when_pipeline_title_does_not_distinguish(self) -> None:
        # Pipeline nie rozroznia (ten sam klucz) - nie podmieniamy, zostaje duplikat.
        control = pd.DataFrame(
            [
                {"pipeline_title": "Plafon TUNA E27 IP44 biały Kanlux 4260"},
                {"pipeline_title": "Plafon TUNA E27 IP44 biały Kanlux 8091"},
            ]
        )
        exported = [
            {"sku": "4260/KAN", "name": "Plafon TUNA E27 IP44 biały Kanlux 4260", "features": "{}"},
            {"sku": "8091/KAN", "name": "Plafon TUNA E27 IP44 biały Kanlux 8091", "features": "{}"},
        ]
        renamed = apply_pipeline_titles_for_duplicate_names(control, exported)
        self.assertEqual(renamed, [])
        self.assertEqual(len(remaining_duplicate_name_rows(exported)), 2)

    def test_duplicate_name_key_ignores_trailing_code(self) -> None:
        self.assertEqual(
            duplicate_name_key("Plafon ERTON E27 IP44 biały Kanlux 35770"),
            duplicate_name_key("Plafon ERTON E27 IP44 biały Kanlux 99999"),
        )


class KanluxBaselinkerUpdateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.knowledge = {
            "attribute_mappings": [
                {
                    "source_name": "Stopień ochrony",
                    "canonical_name": "Stopień ochrony [IP]",
                },
                {
                    "source_name": "Wydajność świetlna",
                    "canonical_name": "Skuteczność świetlna [lm/W]",
                },
            ]
        }

    def test_woo_attributes_are_mapped_to_canonical_names(self) -> None:
        aliases = build_attribute_aliases(self.knowledge)
        row = pd.Series(
            {
                "Atrybut Produktu: Stopień ochrony": "IP65",
                "Atrybut Produktu: Wydajność świetlna": "120 lm/W",
                "Atrybut Produktu: Dostępność": "Dostępny",
                "Atrybut Produktu: Trzonek": "NAN",
            }
        )

        result = woo_features(row, aliases)

        self.assertEqual(result["Stopień ochrony [IP]"], "IP65")
        self.assertEqual(result["Skuteczność świetlna [lm/W]"], "120 lm/W")
        self.assertNotIn("Dostępność", result)
        self.assertNotIn("Trzonek", result)

    def test_woo_features_are_added_as_parameter_fallbacks(self) -> None:
        aliases = build_attribute_aliases(self.knowledge)
        row: dict[str, object] = {"attr_moc": "20W"}
        woo_row = pd.Series(
            {
                "id": "123",
                "Title": "Produkt",
                "Atrybut Produktu: Stopień ochrony": "IP65",
            }
        )

        add_woo_feature_fallbacks(row, woo_row, aliases)

        self.assertEqual(row["woo_match_status"], "MATCHED")
        self.assertEqual(row["Parametr: Stopień ochrony [IP]"], "IP65")
        self.assertEqual(json.loads(str(row["woo_features_json"]))["Stopień ochrony [IP]"], "IP65")

    def test_feature_comparison_reports_added_changed_and_removed(self) -> None:
        result = compare_feature_sets(
            '{"Producent":"Kanlux","Moc [W]":"10","Materiał":"Metal"}',
            '{"Producent":"Kanlux","Moc [W]":"12","Trwałość [h]":"50000"}',
        )

        self.assertEqual(result["added"], ["Trwałość [h]"])
        self.assertEqual(result["changed"], ["Moc [W]"])
        self.assertEqual(result["removed"], ["Materiał"])

    def test_existing_woo_values_extend_catalog_normalizer(self) -> None:
        normalizer: dict[str, dict[str, str]] = {}
        woo = pd.DataFrame(
            [
                {
                    "Atrybut Produktu: Materiał": "ABS,PE",
                    "Atrybut Produktu: Stopień ochrony": "IP 65",
                }
            ]
        )

        add_woo_values_to_feature_normalizer(
            normalizer,
            woo,
            build_attribute_aliases(self.knowledge),
        )

        self.assertIn("Materiał", normalizer)
        self.assertIn("abs,pe", normalizer["Materiał"])
        self.assertEqual(normalizer["Materiał"]["abs,pe"], "ABS|PE")
        self.assertEqual(normalizer["Materiał"]["abs"], "ABS")
        self.assertEqual(normalizer["Materiał"]["pe"], "PE")

    def test_missing_required_skips_attributes_dependent_on_bulb(self) -> None:
        attribute_knowledge = {
            "categories": [
                {
                    "category": "Plafony",
                    "ordered_attributes": [],
                    "required_description": [
                        "Producent",
                        "Moc [W]",
                        "Barwa światła",
                        "Strumień świetlny [lm]",
                        "Czujnik ruchu",
                        "Kolor",
                    ],
                    "filters": [],
                    "optional_description": [],
                }
            ]
        }
        control = pd.DataFrame([{"Kategoria": "Plafoniery i oprawy sufitowe"}])
        exported = [
            {
                "product_id": "1",
                "sku": "36504/KAN",
                "name": "Plafon JASMIN 3xE27",
                "category": "Oświetlenie > Plafony LED",
                "features": json.dumps(
                    {
                        "Producent": "Kanlux",
                        "Trzonek": "E27",
                        "Maksymalna moc źródła światła": "20",
                    },
                    ensure_ascii=False,
                ),
            }
        ]

        rows = build_missing_required_rows(control, exported, attribute_knowledge)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["status"], "BRAKI_WYMAGANYCH")
        self.assertEqual(rows[0]["missing_required"], "Kolor")

    def test_pipeline_title_replaces_title_with_stale_ip(self) -> None:
        control = pd.DataFrame(
            [
                {
                    "pipeline_title": (
                        "Panel LED natynkowy BLINGO U 1195x295mm 40W "
                        "4000K IP54 / IP20 Kanlux 39240"
                    )
                }
            ]
        )
        exported = [
            {
                "sku": "39240/KAN",
                "name": "Panel LED podtynkowy BLINGO U 4000K neutralne IP12 Kanlux 39240",
                "features": json.dumps(
                    {
                        "Typ produktu": "Panel LED",
                        "Stopien ochrony [IP]": "IP 54|IP 20",
                    },
                    ensure_ascii=False,
                ),
            }
        ]

        renamed = apply_pipeline_titles_for_mismatched_names(control, exported)

        self.assertEqual(len(renamed), 1)
        self.assertEqual(renamed[0]["reason"], "FEATURE_MISMATCH")
        self.assertIn("IP54 / IP20", exported[0]["name"])
        self.assertNotIn("IP12", exported[0]["name"])

    def test_pipeline_title_replaces_title_with_stale_mounting(self) -> None:
        control = pd.DataFrame(
            [
                {
                    "pipeline_title": (
                        "Panel LED natynkowy BLINGO U kwadratowy 595x595mm "
                        "34W 4080lm 4000K IP20 Kanlux 39172"
                    )
                }
            ]
        )
        exported = [
            {
                "sku": "39172/KAN",
                "name": "Panel LED podtynkowy BLINGO U 34W 4000K neutralne Kanlux 39172",
                "features": json.dumps(
                    {
                        "Typ produktu": "Panel LED",
                        "Sposob montazu": "Natynkowy",
                    },
                    ensure_ascii=False,
                ),
            }
        ]

        renamed = apply_pipeline_titles_for_mismatched_names(control, exported)

        self.assertEqual(len(renamed), 1)
        self.assertEqual(renamed[0]["reason"], "FEATURE_MISMATCH")
        self.assertIn("natynkowy", exported[0]["name"])
        self.assertNotIn("podtynkowy", exported[0]["name"])

    def test_pipeline_title_replaces_title_missing_title_bearing_attribute(self) -> None:
        control = pd.DataFrame(
            [
                {
                    "pipeline_title": "Oprawa kanałowa MILO E27 165x100mm IP54 szara Kanlux 70523"
                }
            ]
        )
        exported = [
            {
                "sku": "70523/KAN",
                "name": "Oprawa kanałowa MILO E27 IP54 szara Kanlux 70523",
                "features": json.dumps(
                    {
                        "Typ produktu": "Oprawa kanałowa",
                        "Wymiary [mm]": "165x100",
                        "Stopien ochrony [IP]": "IP 54",
                    },
                    ensure_ascii=False,
                ),
            }
        ]

        renamed = apply_pipeline_titles_for_mismatched_names(control, exported)

        self.assertEqual(len(renamed), 1)
        self.assertEqual(renamed[0]["reason"], "FEATURE_MISMATCH")
        self.assertIn("165x100mm", exported[0]["name"])

    def test_pipeline_title_replaces_title_missing_model_identity(self) -> None:
        control = pd.DataFrame(
            [
                {
                    "pipeline_title": "Ramka do paneli ADTR SKY-V 60x60 cm kwadratowa złoto-brązowy Kanlux 39755"
                }
            ]
        )
        exported = [
            {
                "sku": "39755/KAN",
                "name": "Ramka do paneli ADTR 60x60 cm Kanlux 39755",
                "features": json.dumps(
                    {
                        "Typ produktu": "Ramka do paneli",
                        "Wymiary [mm]": "600 x 600",
                    },
                    ensure_ascii=False,
                ),
            }
        ]

        renamed = apply_pipeline_titles_for_mismatched_names(control, exported)

        self.assertEqual(len(renamed), 1)
        self.assertEqual(renamed[0]["reason"], "FEATURE_MISMATCH")
        self.assertIn("ADTR SKY-V", exported[0]["name"])

    def test_pipeline_title_replaces_agor_title_missing_beam_type(self) -> None:
        control = pd.DataFrame(
            [
                {
                    "pipeline_title": "Naświetlacz LED FL AGOR PRO asymetryczny 100W 13500lm 4000K IP65 Kanlux 38423"
                }
            ]
        )
        exported = [
            {
                "sku": "38423/KAN",
                "name": "Naświetlacz LED FL AGOR 100W 13500lm 4000K IP65 Kanlux 38423",
                "features": json.dumps(
                    {
                        "Typ produktu": "Naświetlacz LED",
                        "Moc [W]": "100",
                    },
                    ensure_ascii=False,
                ),
            }
        ]

        renamed = apply_pipeline_titles_for_mismatched_names(control, exported)

        self.assertEqual(len(renamed), 1)
        self.assertIn("asymetryczny", exported[0]["name"])


if __name__ == "__main__":
    unittest.main()
