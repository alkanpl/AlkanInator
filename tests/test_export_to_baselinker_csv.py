from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from export_to_baselinker_csv import (  # noqa: E402
    apply_category_attribute_knowledge,
    build_baselinker_rows,
    build_feature_value_normalizer,
    build_features,
    default_feature_review_output,
    filter_features_for_catalog,
    normalize_feature_value,
    split_multi_color_feature,
    write_feature_review_file,
)


class BaselinkerFeatureNormalizationTest(unittest.TestCase):
    def setUp(self) -> None:
        catalog_knowledge = {
            "colors": [
                {"color": "Biały", "count": 10},
                {"color": "Czarny", "count": 8},
            ],
            "ip_values": [{"ip": "IP 65", "count": 5}],
            "top_attribute_values_by_category": [
                {"attribute": "Moc [W]", "value": "20", "count": 5},
                {"attribute": "Kształt", "value": "Okrągły", "count": 5},
                {"attribute": "Barwa światła", "value": "Neutralna", "count": 5},
            ],
            "top_attributes": [
                {"attribute": "Kąt świecenia [°]", "top_values": "90 (2); 120 (1)"},
                {"attribute": "Wysokość [mm]", "top_values": "112 (4); 60 (2)"},
                {"attribute": "Moc [W]", "top_values": "17.5 (2)"},
            ],
        }
        self.normalizer = build_feature_value_normalizer(catalog_knowledge)

    def test_features_are_normalized_before_baselinker_export(self) -> None:
        row = pd.Series(
            {
                "attr_ip": "IP65",
                "attr_moc": "max 20W",
                "attr_wymiary": "60x60cm",
                "attr_ksztalt": "okrągła",
                "attr_kolor": "bialy|czarny",
                "Barwa - kategoria": "neutralne",
            }
        )

        features = build_features(row, False, feature_value_normalizer=self.normalizer)

        self.assertEqual(features["Stopień ochrony [IP]"], "IP 65")
        self.assertEqual(features["Moc [W]"], "20")
        self.assertEqual(features["Wymiary [mm]"], "600x600")
        self.assertEqual(features["Kształt"], "Okrągły")
        self.assertEqual(features["Kolor"], "Biały / Czarny")
        self.assertEqual(features["Barwa światła"], "Neutralna")

    def test_dimension_attributes_are_converted_to_millimeters(self) -> None:
        features = build_features(
            pd.Series(
                {
                    "attr_dlugosc": "1.2m",
                    "attr_szerokosc": "60cm",
                    "attr_wysokosc": "8.5cm",
                    "attr_srednica": "200mm",
                }
            ),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Długość"], "1200 mm")
        self.assertEqual(features["Szerokość"], "600 mm")
        self.assertEqual(features["Wysokość"], "85 mm")
        self.assertEqual(features["Średnica"], "200 mm")

    def test_manufacturer_data_uses_commas(self) -> None:
        self.assertEqual(
            normalize_feature_value("KANLUX; Objazdowa 1-3 41-922 Radzionków; kanlux@kanlux.pl", "Dane producenta"),
            "KANLUX, Objazdowa 1-3 41-922 Radzionków, kanlux@kanlux.pl",
        )

    def test_socket_count_is_split_into_liczba_gniazd(self) -> None:
        features = build_features(
            pd.Series({"attr_gwint": "3xGU10"}),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Trzonek"], "GU10")
        self.assertEqual(features["Liczba źródeł światła"], "3")

    def test_socket_count_is_read_from_title_when_socket_has_no_count(self) -> None:
        features = build_features(
            pd.Series(
                {
                    "attr_gwint": "E27",
                    "new_title": "Plafon drewniany JASMIN 3xE27 okrągły 47cm biały Kanlux 36504",
                }
            ),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Trzonek"], "E27")
        self.assertEqual(features["Liczba źródeł światła"], "3")

    def test_energy_class_and_mounting_are_exported(self) -> None:
        features = build_features(
            pd.Series(
                {
                    "attr_klasa_energetyczna": "F",
                    "attr_sposob_montazu": "Natynkowy",
                    "attr_moc_max_zrodla": "40W",
                }
            ),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Klasa energetyczna"], "F")
        self.assertEqual(features["Sposób montażu"], "Natynkowy")
        self.assertEqual(features["Maksymalna moc źródła światła"], "40")

    def test_max_bulb_power_is_reclassified(self) -> None:
        features = build_features(
            pd.Series({"Parametr: Moc [W]": "3 x max 20", "attr_gwint": "E27"}),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertNotIn("Moc [W]", features)
        self.assertEqual(features["Maksymalna moc źródła światła"], "20")
        self.assertEqual(features["Liczba źródeł światła"], "3")

    def test_voltage_keeps_current_type(self) -> None:
        self.assertEqual(normalize_feature_value("220-240 AC", "Napięcie [V]"), "220-240 AC")
        self.assertEqual(normalize_feature_value("220-240", "Napięcie [V]"), "220-240 AC")
        self.assertEqual(normalize_feature_value("3,2 DC", "Napięcie [V]"), "3,2 DC")
        self.assertEqual(normalize_feature_value("12V", "Napięcie [V]"), "12")
        self.assertEqual(normalize_feature_value("MLS", "Napięcie [V]"), "")

    def test_led_max_power_value_is_reclassified(self) -> None:
        features = build_features(
            pd.Series({"Parametr: Moc [W]": "max 10 LED", "attr_gwint": "GU10"}),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertNotIn("Moc [W]", features)
        self.assertEqual(features["Maksymalna moc źródła światła"], "10")

    def test_fixture_power_is_dropped_without_included_light_source(self) -> None:
        features = build_features(
            pd.Series(
                {
                    "attr_moc": "15W",
                    "attr_moc_max_zrodla": "25W",
                    "attr_gwint": "E27",
                    "attr_zrodlo_w_komplecie": "0",
                }
            ),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertNotIn("Moc [W]", features)
        self.assertEqual(features["Maksymalna moc źródła światła"], "25")

    def test_fixture_power_is_dropped_for_replaceable_light_source(self) -> None:
        features = build_features(
            pd.Series({"attr_moc": "15W", "attr_zrodlo_swiatla": "Wymienne"}),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertNotIn("Moc [W]", features)

    def test_light_source_has_two_values_and_is_derived(self) -> None:
        self.assertEqual(normalize_feature_value("Wymienne", "Źródło światła"), "Nie zintegrowane")
        self.assertEqual(normalize_feature_value("T8 LED", "Źródło światła"), "Nie zintegrowane")
        self.assertEqual(normalize_feature_value("GLS/CFL/LED", "Źródło światła"), "Nie zintegrowane")
        self.assertEqual(normalize_feature_value("LED", "Źródło światła"), "Zintegrowane")

        features = build_features(
            pd.Series({"attr_gwint": "E27"}),
            False,
            feature_value_normalizer=self.normalizer,
        )
        self.assertEqual(features["Źródło światła"], "Nie zintegrowane")

    def test_integrated_source_drops_socket_features(self) -> None:
        features = build_features(
            pd.Series({"attr_zrodlo_swiatla": "LED", "attr_gwint": "2xGU10"}),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Źródło światła"], "Zintegrowane")
        self.assertNotIn("Trzonek", features)
        self.assertNotIn("Liczba źródeł światła", features)

    def test_light_color_is_derived_from_cct(self) -> None:
        for cct, expected in [("3000", "Ciepła"), ("4000", "Neutralna"), ("6500", "Zimna"), ("3000-6500", "Zmienna")]:
            with self.subTest(cct=cct):
                features = build_features(
                    pd.Series({"attr_barwa": cct}),
                    False,
                    feature_value_normalizer=self.normalizer,
                )
                self.assertEqual(features["Barwa światła"], expected)

    def test_power_supply_is_derived(self) -> None:
        solar = build_features(
            pd.Series({"new_title": "Naświetlacz solarny LED FL 2200lm Kanlux 36607"}),
            False,
            feature_value_normalizer=self.normalizer,
        )
        self.assertEqual(solar["Zasilanie"], "Solarne")

        mains = build_features(
            pd.Series({"attr_napiecie": "220-240 AC", "new_title": "Naświetlacz LED IQ 100W"}),
            False,
            feature_value_normalizer=self.normalizer,
        )
        self.assertEqual(mains["Zasilanie"], "Sieciowe")

    def test_rule_matching_prefers_product_type_and_skips_accessories(self) -> None:
        from export_to_baselinker_csv import match_attribute_category_rule

        knowledge = {
            "categories": [
                {"category": "Plafony", "ordered_attributes": ["Moc [W]"], "required_description": [], "filters": [], "optional_description": []},
                {"category": "Oprawy natynkowe", "ordered_attributes": ["Trzonek"], "required_description": [], "filters": [], "optional_description": []},
            ]
        }
        row = pd.Series({"Kategoria": "Plafoniery i oprawy sufitowe"})

        plafon = match_attribute_category_rule(row, "", knowledge, {"Typ produktu": "Plafon"})
        self.assertEqual(plafon["category"], "Plafony")
        kanalowa = match_attribute_category_rule(row, "", knowledge, {"Typ produktu": "Oprawa kanałowa"})
        self.assertEqual(kanalowa["category"], "Oprawy natynkowe")
        sufitowa = match_attribute_category_rule(row, "", knowledge, {"Typ produktu": "Oprawa sufitowa"})
        self.assertEqual(sufitowa["category"], "Oprawy natynkowe")
        ramka = match_attribute_category_rule(row, "", knowledge, {"Typ produktu": "Ramka do paneli"})
        self.assertIsNone(ramka)

    def test_shape_is_derived_from_dimensions(self) -> None:
        round_shape = build_features(
            pd.Series({"attr_srednica": "470mm"}),
            False,
            feature_value_normalizer=self.normalizer,
        )
        self.assertEqual(round_shape["Kształt"], "Okrągły")

        square = build_features(
            pd.Series({"attr_dlugosc": "595mm", "attr_szerokosc": "595mm"}),
            False,
            feature_value_normalizer=self.normalizer,
        )
        self.assertEqual(square["Kształt"], "Kwadrat")

        rect = build_features(
            pd.Series({"attr_wymiary": "1195x295"}),
            False,
            feature_value_normalizer=self.normalizer,
        )
        self.assertEqual(rect["Kształt"], "Prostokątny")

    def test_materials_are_split_into_separate_values(self) -> None:
        self.assertEqual(normalize_feature_value("ABS,PE", "Materiał"), "ABS|PE")
        self.assertEqual(
            normalize_feature_value("Stop aluminium,tworzywo sztuczne", "Materiał"),
            "Metal|Tworzywo sztuczne",
        )
        self.assertEqual(normalize_feature_value("Metal", "Materiał"), "Metal")

    def test_hermetic_fixture_with_socket_gets_fluorescent_title(self) -> None:
        from export_to_baselinker_csv import hermetic_fluorescent_title

        title = "Oprawa hermetyczna LED DICHT 4LED 1270mm 2x36W IP65 Kanlux 31321"
        result = hermetic_fluorescent_title(title, {"Trzonek": "G13"})
        self.assertEqual(result, "Oprawa hermetyczna do świetlówek LED DICHT 4LED 1270mm 2x36W IP65 Kanlux 31321")
        self.assertEqual(hermetic_fluorescent_title(title, {}), title)
        self.assertEqual(hermetic_fluorescent_title(result, {"Trzonek": "G13"}), result)

    def test_multi_color_value_moves_to_producer_color(self) -> None:
        features = {"Kolor": "Czarny / Złoty"}
        split_multi_color_feature(features)
        self.assertEqual(features["Kolor"], "Czarny")
        self.assertEqual(features["Kolor producenta"], "Czarny / Złoty")
        self.assertNotIn("Kolory", features)

        features = {"Kolor": "Biały|Brązowy"}
        split_multi_color_feature(features)
        self.assertEqual(features["Kolor"], "Biały")
        self.assertEqual(features["Kolor producenta"], "Biały / Brązowy")

        features = {"Kolor": "czarno-biały"}
        split_multi_color_feature(features)
        self.assertEqual(features["Kolor"], "Czarny")
        self.assertEqual(features["Kolor producenta"], "Czarny / Biały")

    def test_light_source_included_is_a_separate_boolean_feature(self) -> None:
        features = build_features(
            pd.Series(
                {
                    "attr_zrodlo_swiatla": "LED",
                    "attr_zrodlo_w_komplecie": "1",
                }
            ),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Źródło światła"], "Zintegrowane")
        self.assertEqual(features["Źródło światła w komplecie"], "Tak")

    def test_top_attributes_values_are_used_after_normalization(self) -> None:
        row = pd.Series(
            {
                "attr_kat_swiecenia": "90°",
                "attr_wysokosc": "112mm",
                "attr_moc": "17,5 W",
            }
        )

        features = build_features(row, False, feature_value_normalizer=self.normalizer)
        filtered = filter_features_for_catalog(features, self.normalizer, [])

        self.assertEqual(filtered["Kąt świecenia [°]"], "90")
        self.assertEqual(filtered["Wysokość"], "112 mm")
        self.assertEqual(filtered["Moc [W]"], "17.5")

    def test_power_range_feature_uses_maximum_value(self) -> None:
        features = build_features(
            pd.Series({"attr_moc": "12 - 18W"}),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Moc [W]"], "18")

    def test_panel_max_source_power_is_promoted_to_own_power(self) -> None:
        features = build_features(
            pd.Series(
                {
                    "attr_typ": "Panel LED",
                    "attr_moc_max_zrodla": "max 40",
                }
            ),
            False,
            feature_value_normalizer=self.normalizer,
        )

        self.assertEqual(features["Moc [W]"], "40")

    def test_parametr_columns_use_the_same_normalization(self) -> None:
        row = pd.Series(
            {
                "Parametr: Stopień ochrony IP": "IP65",
                "Parametr: moc": "20 W",
                "Parametr: Materiał": pd.NA,
            }
        )

        features = build_features(row, False, feature_value_normalizer=self.normalizer)

        self.assertEqual(features["Stopień ochrony [IP]"], "IP 65")
        self.assertEqual(features["Moc [W]"], "20")
        self.assertNotIn("Materiał", features)

    def test_category_knowledge_filters_and_orders_lighting_features(self) -> None:
        knowledge = {
            "categories": [
                {
                    "category": "Panele LED",
                    "ordered_attributes": ["Moc [W]", "Wskaźnik olśnienia [UGR]"],
                    "required_description": ["Producent"],
                    "filters": ["Moc [W]"],
                    "optional_description": [],
                }
            ]
        }
        features = {
            "Materiał": "Metal",
            "Wskaźnik olśnienia [UGR]": "<19",
            "Producent": "Kanlux",
            "Moc [W]": "40",
            "Dane producenta": "KANLUX; adres; email",
            "Typ produktu": "Panel LED",
        }

        result = apply_category_attribute_knowledge(
            features,
            pd.Series({"Kategoria": "Panele LED"}),
            "Oświetlenie > Panele LED sufitowe",
            knowledge,
        )

        self.assertEqual(
            list(result),
            ["Producent", "Dane producenta", "Typ produktu", "Moc [W]", "Wskaźnik olśnienia [UGR]"],
        )
        self.assertNotIn("Materiał", result)

    def test_store_category_has_priority_over_source_family(self) -> None:
        knowledge = {
            "categories": [
                {
                    "category": "Plafony",
                    "ordered_attributes": ["Moc [W]"],
                    "required_description": [],
                    "filters": [],
                    "optional_description": [],
                },
                {
                    "category": "Oprawy natynkowe",
                    "ordered_attributes": ["Trzonek", "Materiał"],
                    "required_description": [],
                    "filters": [],
                    "optional_description": [],
                },
            ]
        }
        features = {
            "Producent": "Kanlux",
            "Dane producenta": "KANLUX; adres; email",
            "Typ produktu": "Oprawa natynkowa",
            "Moc [W]": "10",
            "Trzonek": "GU10",
            "Materiał": "Aluminium",
        }

        result = apply_category_attribute_knowledge(
            features,
            pd.Series({"Kategoria": "Plafoniery i oprawy sufitowe"}),
            "Oświetlenie > Oświetlenie wewnętrzne > Oprawy natynkowe",
            knowledge,
        )

        self.assertIn("Trzonek", result)
        self.assertIn("Materiał", result)
        self.assertNotIn("Moc [W]", result)

    def test_compound_ip_is_normalized_and_woo_exact_values_win(self) -> None:
        normalizer = build_feature_value_normalizer(
            {"ip_values": [{"ip": "IP20/IP44", "count": 10}]}
        )

        first = build_features(
            pd.Series({"attr_ip": "IP20/IP44"}),
            False,
            feature_value_normalizer=normalizer,
        )
        second = build_features(
            pd.Series({"attr_ip": "IP54/IP20"}),
            False,
            feature_value_normalizer=normalizer,
        )

        self.assertEqual(first["Stopień ochrony [IP]"], "IP20/IP44")
        self.assertEqual(second["Stopień ochrony [IP]"], "IP 54/IP 20")

    def test_build_rows_writes_normalized_features_json(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "new_title": "Panel LED 20W 60x60 TEST",
                    "Kod": "TEST",
                    "EAN": "5900000000000",
                    "Producent": "Kanlux",
                    "proponowana_kategoria_1": "Oświetlenie > Panele LED",
                    "description_html": "<p>Opis</p>",
                    "attr_ip": "IP65",
                }
            ]
        )

        rows = build_baselinker_rows(
            df,
            sku_format="plain",
            feature_value_normalizer=self.normalizer,
        )
        features = json.loads(rows[0]["features"])

        self.assertEqual(features["Stopień ochrony [IP]"], "IP 65")

    def test_baselinker_features_do_not_include_ean_or_producer_code(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "new_title": "Panel LED 20W 60x60 TEST",
                    "Kod": "TEST",
                    "EAN": "5900000000000",
                    "Producent": "Kanlux",
                    "proponowana_kategoria_1": "Oświetlenie > Panele LED",
                    "description_html": "<p>Opis</p>",
                    "attr_ip": "IP65",
                }
            ]
        )

        rows = build_baselinker_rows(
            df,
            sku_format="plain",
            feature_value_normalizer=self.normalizer,
        )
        features = json.loads(rows[0]["features"])

        self.assertNotIn("EAN (GTIN)", features)
        self.assertNotIn("Kod producenta", features)
        self.assertEqual(rows[0]["ean"], "5900000000000")
        self.assertEqual(rows[0]["sku"], "TEST")

    def test_unknown_attribute_or_value_goes_to_review_instead_of_features(self) -> None:
        review_rows: list[dict[str, str]] = []

        filtered = filter_features_for_catalog(
            {
                "Stopień ochrony [IP]": "IP 65",
                "Nowy atrybut": "Tak",
                "Barwa światła": "Nieznana barwa",
            },
            self.normalizer,
            review_rows,
            {"sku": "TEST", "ean": "5900000000000", "name": "Produkt testowy"},
        )

        self.assertEqual(filtered, {"Stopień ochrony [IP]": "IP 65"})
        self.assertEqual(
            [(row["feature_name"], row["feature_value"], row["reason"]) for row in review_rows],
            [
                ("Nowy atrybut", "Tak", "UNKNOWN_ATTRIBUTE"),
                ("Barwa światła", "Nieznana barwa", "UNKNOWN_VALUE"),
            ],
        )

    def test_review_notes_drive_next_attribute_iteration(self) -> None:
        normalizer = build_feature_value_normalizer(
            {
                "ip_values": [
                    {"ip": "IP 20", "count": 10},
                    {"ip": "IP 54", "count": 10},
                ],
                "top_attributes": [
                    {"attribute": "Czujnik ruchu", "top_values": "Nie (152); Tak (52)"},
                    {"attribute": "Kolor", "top_values": "Złoty (516)"},
                ],
            }
        )
        row = pd.Series(
            {
                "attr_typ": "High Bay przemysłowy",
                "attr_model": "HB PRO STRONG 100W-NW",
                "attr_strumien": "8500-17000 lm",
                "attr_wysokosc": "112mm",
                "attr_moc": "32 W",
                "attr_ilosc_sztuk": "2 szt.",
                "attr_pasuje_do": "do paneli",
                "attr_czujnik": "z czujnikiem ruchu",
                "attr_ik": "IK08",
                "attr_ip": "IP54/IP20",
                "attr_kolor": "złoto-brązowy",
                "attr_kolor_producenta": "złoto-brązowy",
                "attr_gwarancja": "5 lat",
            }
        )
        review_rows: list[dict[str, str]] = []

        features = build_features(row, False, feature_value_normalizer=normalizer)
        filtered = filter_features_for_catalog(features, normalizer, review_rows)

        self.assertEqual(filtered["Typ produktu"], "Oprawa High Bay")
        self.assertNotIn("Model", filtered)
        self.assertEqual(filtered["Strumień świetlny [lm]"], "8500-17000")
        self.assertEqual(filtered["Wysokość"], "112 mm")
        self.assertEqual(filtered["Moc [W]"], "32")
        self.assertEqual(filtered["Opakowanie"], "2 szt.")
        self.assertEqual(filtered["Zastosowanie"], "Do paneli")
        self.assertEqual(filtered["Czujnik ruchu"], "Tak")
        self.assertEqual(filtered["Stopień ochrony [IK]"], "IK 08")
        self.assertEqual(filtered["Stopień ochrony [IP]"], "IP 54|IP 20")
        self.assertEqual(filtered["Kolor"], "Złoty")
        self.assertEqual(filtered["Kolor producenta"], "Złoto-brązowy")
        self.assertEqual(filtered["Gwarancja"], "5 lat")
        self.assertEqual(review_rows, [])

    def test_next_review_iteration_unifies_similar_series_and_types(self) -> None:
        normalizer = build_feature_value_normalizer({})
        rows = [
            pd.Series({"attr_seria": "VARSO LED"}),
            pd.Series({"attr_seria": "IQ-LED FL"}),
            pd.Series({"attr_seria": "DICHT 4LED"}),
            pd.Series({"attr_seria": "BAREV BL"}),
            pd.Series({"attr_seria": "BLINGO UAIO"}),
            pd.Series({"attr_seria": "FL AGOR AIO"}),
            pd.Series({"attr_seria": "ADTR-S-H"}),
            pd.Series({"attr_seria": "MAH LED HI"}),
            pd.Series({"attr_seria": "TP STRONG LED"}),
            pd.Series({"attr_seria": "BORD DLP"}),
            pd.Series({"attr_seria": "PIRES ECO"}),
        ]
        expected = ["VARSO", "IQ-LED", "DICHT", "BAREV", "BLINGO", "AGOR", "ADTR", "MAH", "TP", "BORD", "PIRES"]

        got = []
        for row in rows:
            features = build_features(row, False, feature_value_normalizer=normalizer)
            filtered = filter_features_for_catalog(features, normalizer, [])
            got.append(filtered["Seria"])

        self.assertEqual(got, expected)

    def test_next_review_iteration_maps_product_types_and_drops_context_noise(self) -> None:
        normalizer = build_feature_value_normalizer({})
        row = pd.Series(
            {
                "attr_typ": "Oprawa hermetyczna LED Pro",
                "attr_seria": "MAH PLUS",
                "attr_kat_swiecenia": "145°",
                "attr_wymiary": "3x1mm2",
                "attr_barwa": "3000-6500K",
                "attr_napiecie": "12V",
                "attr_kolor": "Złoty dąb",
            }
        )

        features = build_features(row, False, feature_value_normalizer=normalizer)
        filtered = filter_features_for_catalog(
            features,
            normalizer,
            [],
            {"name": "Oprawa hermetyczna LED MAH PRO 1200mm"},
        )

        self.assertEqual(filtered["Typ produktu"], "Oprawa hermetyczna")
        self.assertNotIn("Kąt świecenia", filtered)
        self.assertNotIn("Wymiary [mm]", filtered)
        self.assertEqual(filtered["Temperatura barwowa [K]"], "3000-6500")
        self.assertEqual(filtered["Napięcie [V]"], "12")
        self.assertEqual(filtered["Kolor"], "Brązowy")
        self.assertEqual(filtered["Kolor producenta"], "Złoty dąb")

    def test_latest_review_update_maps_colors_types_and_context_removals(self) -> None:
        normalizer = build_feature_value_normalizer(
            {
                "colors": [
                    {"color": "Biały", "count": 10},
                    {"color": "Brązowy", "count": 10},
                    {"color": "Czarny", "count": 10},
                    {"color": "Złoty", "count": 10},
                ],
                "ip_values": [{"ip": "IP 20", "count": 10}],
            }
        )
        row = pd.Series(
            {
                "attr_typ": "ADTR PT",
                "attr_kat_swiecenia": "115°",
                "attr_ip": "IP12",
                "attr_kolor": "Biały / Brązowy",
            }
        )

        features = build_features(row, False, feature_value_normalizer=normalizer)
        filtered = filter_features_for_catalog(
            features,
            normalizer,
            [],
            {"name": "Plafon LED CCT ANBAR"},
        )

        self.assertEqual(filtered["Typ produktu"], "Ramka do paneli")
        self.assertNotIn("Kąt świecenia", filtered)
        self.assertEqual(filtered["Stopień ochrony [IP]"], "IP 20")
        self.assertEqual(filtered["Kolor"], "Biały")
        self.assertEqual(filtered["Kolor producenta"], "Biały / Brązowy")
        self.assertNotIn("Kolory", filtered)

        link_features = build_features(pd.Series({"attr_typ": "Linka", "attr_seria": "SPN"}), False, normalizer)
        link_filtered = filter_features_for_catalog(
            link_features,
            normalizer,
            [],
            {"name": "Linka do podwieszania lamp BRAVO"},
        )
        self.assertEqual(link_filtered["Typ produktu"], "Linka do podwieszenia")
        self.assertNotIn("Seria", link_filtered)

    def test_latest_feature_review_notes_are_applied_after_normalization(self) -> None:
        normalizer = build_feature_value_normalizer(
            {
                "colors": [
                    {"color": "Srebrny", "count": 10},
                ],
            }
        )
        review_rows: list[dict[str, str]] = []

        features = build_features(
            pd.Series(
                {
                    "attr_seria": "AL55-SH-NW-MAT-W-NT",
                    "attr_kat_swiecenia": "125°",
                    "attr_kolor": "STAL INOX",
                }
            ),
            False,
            feature_value_normalizer=normalizer,
        )
        filtered = filter_features_for_catalog(
            features,
            normalizer,
            review_rows,
            {"name": "Oprawa sufitowa PHLOX C"},
        )

        self.assertEqual(filtered["Seria"], "ALIN")
        self.assertEqual(filtered["Kąt świecenia [°]"], "125")
        self.assertEqual(filtered["Kolor"], "Srebrny")
        self.assertEqual(filtered["Kolor producenta"], "STAL INOX")
        self.assertEqual(review_rows, [])

    def test_review_notes_drop_series_for_lenses_and_clips(self) -> None:
        normalizer = build_feature_value_normalizer({})
        lens_features = build_features(
            pd.Series({"attr_typ": "Soczewka", "attr_seria": "FLS"}),
            False,
            feature_value_normalizer=normalizer,
        )
        clip_features = build_features(
            pd.Series({"attr_typ": "Zapinka", "attr_seria": "CL-MAH"}),
            False,
            feature_value_normalizer=normalizer,
        )

        lens_filtered = filter_features_for_catalog(
            lens_features,
            normalizer,
            [],
            {"name": "Soczewka do oprawy LED"},
        )
        clip_filtered = filter_features_for_catalog(
            clip_features,
            normalizer,
            [],
            {"name": "Zapinka do oprawy MAH"},
        )

        self.assertEqual(lens_filtered["Typ produktu"], "Soczewka")
        self.assertEqual(clip_filtered["Typ produktu"], "Zapinka")
        self.assertNotIn("Seria", lens_filtered)
        self.assertNotIn("Seria", clip_filtered)

    def test_accepted_fls_series_stays_on_non_lens_accessories(self) -> None:
        normalizer = build_feature_value_normalizer({})

        filtered = filter_features_for_catalog(
            build_features(
                pd.Series({"attr_typ": "Wspornik montażowy", "attr_seria": "FLS"}),
                False,
                feature_value_normalizer=normalizer,
            ),
            normalizer,
            [],
            {"name": "Wspornik montażowy do FL STADER FLS"},
        )

        self.assertEqual(filtered["Typ produktu"], "Wspornik montażowy")
        self.assertEqual(filtered["Seria"], "FLS")

    def test_accessories_do_not_export_compatible_fixture_power_as_own_power(self) -> None:
        normalizer = build_feature_value_normalizer({})

        filtered = filter_features_for_catalog(
            build_features(
                pd.Series({"attr_typ": "Siatka ochronna", "attr_moc": "240W"}),
                False,
                feature_value_normalizer=normalizer,
            ),
            normalizer,
            [],
            {"name": "Siatka ochronna do FL AGOR HI GRID 240W"},
        )

        self.assertEqual(filtered["Typ produktu"], "Siatka ochronna")
        self.assertNotIn("Moc [W]", filtered)

    def test_panel_frame_does_not_export_dimension_code_as_power(self) -> None:
        normalizer = build_feature_value_normalizer({})

        filtered = filter_features_for_catalog(
            build_features(
                pd.Series(
                    {
                        "attr_typ": "Ramka",
                        "attr_seria": "ADTR",
                        "Moc [W]": "60",
                        "attr_wymiary": "60x60 cm",
                    }
                ),
                False,
                feature_value_normalizer=normalizer,
            ),
            normalizer,
            [],
            {"name": "Ramka do paneli ADTR SKY 60x60 cm"},
        )

        self.assertEqual(filtered["Typ produktu"], "Ramka")
        self.assertNotIn("Moc [W]", filtered)
        self.assertEqual(filtered["Wymiary [mm]"], "600x600")

    def test_latest_review_color_producer_values_keep_base_color_clean(self) -> None:
        normalizer = build_feature_value_normalizer(
            {
                "colors": [
                    {"color": "Biały", "count": 10},
                    {"color": "Czarny", "count": 10},
                ],
            }
        )

        white = filter_features_for_catalog(
            build_features(
                pd.Series({"attr_kolor": "Biały mat"}),
                False,
                feature_value_normalizer=normalizer,
            ),
            normalizer,
            [],
        )
        black = filter_features_for_catalog(
            build_features(
                pd.Series({"attr_kolor": "Czarny mat"}),
                False,
                feature_value_normalizer=normalizer,
            ),
            normalizer,
            [],
        )

        self.assertEqual(white["Kolor"], "Biały")
        self.assertEqual(white["Kolor producenta"], "Biały mat")
        self.assertEqual(black["Kolor"], "Czarny")
        self.assertEqual(black["Kolor producenta"], "Czarny mat")

    def test_phlox_c_type_is_normalized_after_comparable_key_cleanup(self) -> None:
        normalizer = build_feature_value_normalizer({})
        filtered = filter_features_for_catalog(
            build_features(
                pd.Series({"attr_typ": "PHLOX C"}),
                False,
                feature_value_normalizer=normalizer,
            ),
            normalizer,
            [],
        )

        self.assertEqual(filtered["Typ produktu"], "Oprawa sufitowa")

    def test_feature_review_defaults_to_xlsx_and_can_be_opened_by_pandas(self) -> None:
        self.assertTrue(default_feature_review_output("output/import.csv").endswith(".xlsx"))
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "feature_review.xlsx"
            write_feature_review_file(
                [
                    {
                        "sku": "TEST",
                        "ean": "5900000000000",
                        "name": "Produkt testowy",
                        "feature_name": "Nowy atrybut",
                        "feature_value": "Tak",
                        "reason": "UNKNOWN_ATTRIBUTE",
                    }
                ],
                path,
            )

            df = pd.read_excel(path)

        self.assertEqual(df.loc[0, "sku"], "TEST")
        self.assertEqual(df.loc[0, "reason"], "UNKNOWN_ATTRIBUTE")


if __name__ == "__main__":
    unittest.main()
