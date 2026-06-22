from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from optimize_titles import optimize_titles_for_dataframe  # noqa: E402
from enrich_from_parameters import build_parameter_index, detect_parameter_columns, enrich_from_parameters  # noqa: E402
from title_anatomy import build_title_from_anatomy, build_title_type_review, duplicate_titles_report  # noqa: E402


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

    def test_solar_keyword_override_skips_power(self) -> None:
        config = {
            **ANATOMY_CONFIG,
            "series_title_keyword": {"SLR": "Naswietlacz solarny LED"},
            "rules": [
                {
                    "normalized_product_type": "naswietlacz led",
                    "status": "approved",
                    "title_keyword": "Naswietlacz LED",
                    "required_title_attributes": ["seria", "moc", "strumien", "lm_w", "barwa", "ip"],
                    "optional_title_attributes": [],
                }
            ],
        }
        row = pd.Series({"attr_typ": "naswietlacz LED", "Kod": "36605", "Producent": "Kanlux"})
        values = {
            "typ": "naswietlacz LED",
            "old_title": "Naswietlacz LED solarny FL SOLNAR SLR 8W",
            "seria": "FL",
            "moc": "8W",
            "strumien": "800lm",
            "lm_w": "100lm/W",
            "barwa": "4000K",
            "ip": "IP54",
            "producent": "Kanlux",
        }
        result = build_title_from_anatomy(row, config, values)
        self.assertIsNotNone(result)
        assert result is not None
        self.assertEqual(result.title, "Naswietlacz solarny LED FL 800lm 4000K IP54 Kanlux 36605")
        self.assertIn("moc", result.skipped_attributes)
        self.assertIn("lm_w", result.skipped_attributes)

    def test_avar_frame_uses_feminine_mounting_and_shape(self) -> None:
        config = {
            **ANATOMY_CONFIG,
            "series_title_keyword": {"AVAR": "Ramka oswietleniowa LED"},
            "rules": [
                {
                    "normalized_product_type": "panel led",
                    "status": "approved",
                    "title_keyword": "Panel LED",
                    "required_title_attributes": ["montaz", "seria", "ksztalt_masculine", "wymiary"],
                    "optional_title_attributes": [],
                }
            ],
        }
        row = pd.Series({"attr_typ": "panel LED", "Kod": "26771", "Producent": "Kanlux"})
        values = {
            "typ": "panel LED",
            "old_title": "Panel LED ramka AVAR 6060 40W-CW",
            "montaz": "podtynkowy",
            "seria": "AVAR",
            "ksztalt_masculine": "kwadratowy",
            "ksztalt_feminine": "kwadratowa",
            "wymiary": "60x60 cm",
            "producent": "Kanlux",
        }
        result = build_title_from_anatomy(row, config, values)
        self.assertIsNotNone(result)
        assert result is not None
        self.assertEqual(result.title, "Ramka oswietleniowa LED podtynkowa AVAR kwadratowa 60x60 cm Kanlux 26771")

    def test_review_reports_approved_rules_without_manual_review_flag_as_not_required(self) -> None:
        df = pd.DataFrame([{"attr_typ": "Panel LED", "new_title": "Panel LED BLINGO 36W Kanlux 12345"}])
        report = build_title_type_review(df, ANATOMY_CONFIG)
        self.assertEqual(report.loc[0, "review_status"], "approved")
        self.assertFalse(bool(report.loc[0, "manual_review_required"]))

    def test_duplicate_titles_are_disambiguated_with_differing_attribute(self) -> None:
        config = {
            "title_template": "Ramka {seria} {wymiary} {ksztalt} {kolor} {producent}",
            "append_sku_after_producer": True,
            "sku_columns": ["Kod"],
            "max_title_length": 120,
            "required_fields": [],
            "duplicate_title_disambiguation_attributes": ["wysokosc"],
        }
        df = pd.DataFrame(
            [
                {
                    "Nazwa": "Old A",
                    "attr_typ": "Ramka",
                    "attr_seria": "ADTR-H",
                    "attr_wymiary": "60x60cm",
                    "attr_ksztalt": "kwadratowa",
                    "attr_kolor": "biała",
                    "attr_wysokosc": "65mm",
                    "attr_producent": "Kanlux",
                    "Kod": "29843",
                },
                {
                    "Nazwa": "Old B",
                    "attr_typ": "Ramka",
                    "attr_seria": "ADTR-H",
                    "attr_wymiary": "60x60cm",
                    "attr_ksztalt": "kwadratowa",
                    "attr_kolor": "biała",
                    "attr_wysokosc": "76mm",
                    "attr_producent": "Kanlux",
                    "Kod": "33398",
                },
            ]
        )
        result = optimize_titles_for_dataframe(df, "Nazwa", config, {})
        self.assertIn("wys. 65mm", result.loc[0, "new_title"])
        self.assertIn("wys. 76mm", result.loc[1, "new_title"])
        self.assertNotIn("duplicate_seo_title_without_code", result.loc[0, "warnings"])
        self.assertNotIn("duplicate_seo_title_without_code", result.loc[1, "warnings"])

    def test_legacy_title_normalizes_wooden_plafon_gender(self) -> None:
        config = {
            "title_template": "Plafon {material} {seria} {producent}",
            "append_sku_after_producer": True,
            "sku_columns": ["Kod"],
            "max_title_length": 120,
            "required_fields": [],
        }
        df = pd.DataFrame(
            [
                {
                    "Nazwa": "Old",
                    "attr_material": "drewniana",
                    "attr_seria": "JASMIN",
                    "attr_producent": "Kanlux",
                    "Kod": "36507",
                }
            ]
        )

        result = optimize_titles_for_dataframe(df, "Nazwa", config, {})

        self.assertEqual(result.loc[0, "new_title"], "Plafon drewniany JASMIN Kanlux 36507")

    def test_power_range_is_used_in_title_while_attribute_can_hold_maximum(self) -> None:
        config = {
            **ANATOMY_CONFIG,
            "rules": [
                {
                    "normalized_product_type": "panel led",
                    "status": "approved",
                    "title_keyword": "Panel LED",
                    "required_title_attributes": ["seria", "wymiary", "moc"],
                    "optional_title_attributes": [],
                }
            ],
        }
        df = pd.DataFrame(
            [
                {
                    "Nazwa": "Old",
                    "attr_typ": "Panel LED",
                    "attr_seria": "BLINGO",
                    "attr_wymiary": "60x60 cm",
                    "attr_moc": "18W",
                    "Moc - zakres": "12 - 18W",
                    "Producent": "Kanlux",
                    "Kod": "11111",
                }
            ]
        )

        result = optimize_titles_for_dataframe(df, "Nazwa", TITLE_CONFIG, config)

        self.assertIn("12-18W", result.loc[0, "new_title"])
        self.assertNotRegex(result.loc[0, "new_title"], r"(?<!-)18W\b")

    def test_panel_title_uses_max_power_when_regular_power_is_missing(self) -> None:
        config = {
            **ANATOMY_CONFIG,
            "rules": [
                {
                    "normalized_product_type": "panel led",
                    "status": "approved",
                    "title_keyword": "Panel LED",
                    "required_title_attributes": ["seria", "wymiary", "moc"],
                    "optional_title_attributes": [],
                }
            ],
        }
        df = pd.DataFrame(
            [
                {
                    "Nazwa": "Old",
                    "attr_typ": "Panel LED",
                    "attr_seria": "BLINGO UHRA",
                    "attr_wymiary": "595x595mm",
                    "attr_moc_max_zrodla": "max 40",
                    "Producent": "Kanlux",
                    "Kod": "39170",
                }
            ]
        )

        result = optimize_titles_for_dataframe(df, "Nazwa", TITLE_CONFIG, config)

        self.assertIn("40W", result.loc[0, "new_title"])

    def test_adtr_frame_preserves_supplier_model_family(self) -> None:
        config = {
            "title_template": "{typ} {seria} {producent}",
            "accessory_title_templates": {
                "Ramka": "{accessory_type} do paneli {seria} {wymiary} {ksztalt} {kolor} {producent}",
            },
            "append_sku_after_producer": True,
            "sku_columns": ["Kod"],
            "max_title_length": 120,
            "required_fields_by_role": {"accessory": []},
        }
        df = pd.DataFrame(
            [
                {
                    "Nazwa": "Old",
                    "product_role": "accessory",
                    "accessory_type": "Ramka",
                    "attr_seria": "ADTR",
                    "attr_wymiary": "60x60 cm",
                    "attr_ksztalt": "kwadratowa",
                    "attr_kolor": "biała",
                    "Nazwa Kanlux": "ADTR SKY 6060 W",
                    "attr_producent": "Kanlux",
                    "Kod": "27618",
                }
            ]
        )

        result = optimize_titles_for_dataframe(df, "Nazwa", config, {})

        self.assertIn("ADTR SKY", result.loc[0, "new_title"])
        self.assertIn("60x60 cm", result.loc[0, "new_title"])

    def test_agor_floodlight_title_gets_beam_type_from_supplier_name(self) -> None:
        config = {
            **ANATOMY_CONFIG,
            "rules": [
                {
                    "normalized_product_type": "naswietlacz led",
                    "status": "approved",
                    "title_keyword": "Naswietlacz LED",
                    "required_title_attributes": ["seria_clean", "optyka", "moc", "kat_swiecenia"],
                    "optional_title_attributes": [],
                }
            ],
        }
        df = pd.DataFrame(
            [
                {
                    "Nazwa": "Old asym",
                    "attr_typ": "naswietlacz LED",
                    "attr_seria": "FL AGOR",
                    "attr_moc": "100W",
                    "attr_kat_swiecenia": "125°",
                    "Nazwa Kanlux": "FL AGOR/A PRO 100W NW",
                    "Producent": "Kanlux",
                    "Kod": "38423",
                },
                {
                    "Nazwa": "Old sym",
                    "attr_typ": "naswietlacz LED",
                    "attr_seria": "FL AGOR",
                    "attr_moc": "100W",
                    "attr_kat_swiecenia": "110°",
                    "Nazwa Kanlux": "FL AGOR PRO 100W NW",
                    "Producent": "Kanlux",
                    "Kod": "38420",
                },
            ]
        )

        result = optimize_titles_for_dataframe(df, "Nazwa", TITLE_CONFIG, config)

        self.assertIn("FL AGOR PRO asymetryczny", result.loc[0, "new_title"])
        self.assertIn("FL AGOR PRO symetryczny", result.loc[1, "new_title"])

    def test_ceiling_fixture_title_adds_pointowa_when_source_says_point_fixture(self) -> None:
        config = {
            "title_template": "{typ} {seria} {gwint} {ip} {producent}",
            "title_strategy_rules": [
                {
                    "type_contains": "Oprawa sufitowa",
                    "template": "Oprawa sufitowa {punktowa} {seria} {gwint} {ip} {producent}",
                    "required_fields": ["producent"],
                }
            ],
            "append_sku_after_producer": True,
            "sku_columns": ["Kod"],
            "max_title_length": 120,
        }
        df = pd.DataFrame(
            [
                {
                    "Nazwa": "Oprawa sufitowa AQILO GU10",
                    "Nazwa B2C / SEO": "Oprawa sufitowa punktowa AQILO GU10",
                    "attr_typ": "Oprawa sufitowa",
                    "attr_seria": "AQILO",
                    "attr_gwint": "GU10",
                    "attr_ip": "IP20",
                    "attr_producent": "Kanlux",
                    "Kod": "12345",
                }
            ]
        )

        result = optimize_titles_for_dataframe(df, "Nazwa", config, {})

        self.assertEqual(result.loc[0, "new_title"], "Oprawa sufitowa punktowa AQILO GU10 IP20 Kanlux 12345")

    def test_parameters_clear_false_sensor_and_map_cable_length(self) -> None:
        products = pd.DataFrame(
            [
                {
                    "Kod": "19001",
                    "Nazwa B2C / SEO": "Plafoniera PIRES ECO E27",
                    "Typ": "plafoniera z czujnikiem ruchu",
                    "Rodzina": "PIRES ECO",
                    "Czujnik ruchu": "z czujnikiem ruchu",
                },
                {
                    "Kod": "36507",
                    "Nazwa B2C / SEO": "Plafoniera drewniana wisząca JASMIN C 470-B",
                    "Typ": "plafoniera drewniana dekoracyjna",
                    "Rodzina": "JASMIN",
                },
            ]
        )
        params = pd.DataFrame(
            [
                {"Kod Kanlux": "19001", "Nazwa": "PIRES ECO", "Nazwa atrybutu": "Trzonek", "Wartość": "E27"},
                {"Kod Kanlux": "36507", "Nazwa": "JASMIN", "Nazwa atrybutu": "Długość przewodu [m]", "Wartość": "1"},
            ]
        )
        index, _ = build_parameter_index(params, detect_parameter_columns(params))
        config = {
            "title_fallback_columns": ["Nazwa Kanlux"],
            "default_values": {"producent": "Kanlux"},
            "input_attribute_columns": {"typ": "Typ", "seria": "Rodzina", "czujnik": "Czujnik ruchu"},
        }
        enriched, _ = enrich_from_parameters(
            products,
            {"sku": "Kod", "title": "Nazwa B2C / SEO", "category": None, "producer": None},
            config,
            index,
        )
        self.assertEqual(enriched.loc[0, "attr_czujnik"], "")
        self.assertEqual(enriched.loc[1, "attr_dlugosc_przewodu"], "1m")


if __name__ == "__main__":
    unittest.main()
