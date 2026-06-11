from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from export_to_baselinker_csv import (  # noqa: E402
    build_baselinker_rows,
    build_feature_value_normalizer,
    build_features,
    default_feature_review_output,
    filter_features_for_catalog,
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
                {"attribute": "Wysokość", "top_values": "112 mm (4); 60 mm (2)"},
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
        self.assertEqual(features["Wymiary"], "60x60 cm")
        self.assertEqual(features["Kształt"], "Okrągły")
        self.assertEqual(features["Kolor"], "Biały / Czarny")
        self.assertEqual(features["Barwa światła"], "Neutralna")

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

        self.assertEqual(filtered["Kąt świecenia"], "90")
        self.assertEqual(filtered["Wysokość"], "112 mm")
        self.assertEqual(filtered["Moc [W]"], "17.5")

    def test_parametr_columns_use_the_same_normalization(self) -> None:
        row = pd.Series(
            {
                "Parametr: Stopień ochrony IP": "IP65",
                "Parametr: moc": "20 W",
            }
        )

        features = build_features(row, False, feature_value_normalizer=self.normalizer)

        self.assertEqual(features["Stopień ochrony [IP]"], "IP 65")
        self.assertEqual(features["Moc [W]"], "20")

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
        self.assertEqual(filtered["Wymiary"], "3x1 mm2")
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
        self.assertEqual(filtered["Kolory"], "Biały / Brązowy")

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
        self.assertEqual(filtered["Kąt świecenia"], "125")
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
