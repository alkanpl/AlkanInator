from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import pandas as pd  # noqa: E402

from enrich_from_parameters import (  # noqa: E402
    build_parameter_index,
    normalize_parameter_value,
    parameter_attribute_to_internal,
)


class EnrichFromParametersTests(unittest.TestCase):
    def test_maps_new_lighting_attributes(self) -> None:
        cases = {
            "UGR": "ugr",
            "Współczynnik oddawania barw Ra": "cri",
            "Trwałość [h]": "trwalosc",
            "Możliwość łączenia przelotowego opraw": "laczenie_przelotowe",
            "Możliwość współpracy ze ściemniaczem": "sciemnianie",
            "Źródło światła": "zrodlo_swiatla",
            "Źródło światła w komplecie": "zrodlo_w_komplecie",
            "Zintegrowane źródło światła LED": "zrodlo_w_komplecie",
        }
        for source, expected in cases.items():
            with self.subTest(source=source):
                self.assertEqual(parameter_attribute_to_internal(source), expected)

    def test_lookalike_parameters_are_not_mapped(self) -> None:
        cases = [
            "Źródło światła zasilane lub niezasilane napięciem sieciowym",
            "Maksymalna moc opraw połączonych przelotowo [? W]",
            "Pobór mocy silnika [W]",
            "Moc transmisji [mW]",
            "Klasa energetyczna statecznika",
            "Współczynnik mocy",
        ]
        for source in cases:
            with self.subTest(source=source):
                self.assertEqual(parameter_attribute_to_internal(source), "")

    def test_maps_energy_class_and_mounting_place(self) -> None:
        self.assertEqual(parameter_attribute_to_internal("Klasa efektywności energetycznej"), "klasa_energetyczna")
        self.assertEqual(parameter_attribute_to_internal("Miejsce montażu"), "sposob_montazu")

    def test_normalizes_energy_class_and_mounting_values(self) -> None:
        self.assertEqual(normalize_parameter_value("f", "klasa_energetyczna", "Klasa efektywności energetycznej"), "F")
        self.assertEqual(normalize_parameter_value("nie dotyczy", "klasa_energetyczna", "Klasa efektywności energetycznej"), "")
        self.assertEqual(
            normalize_parameter_value(
                "do nadbudowania na ścianie,do nadbudowania na suficie", "sposob_montazu", "Miejsce montażu"
            ),
            "Natynkowy",
        )
        self.assertEqual(
            normalize_parameter_value("do wbudowania w sufit", "sposob_montazu", "Miejsce montażu"),
            "Podtynkowy",
        )
        self.assertEqual(
            normalize_parameter_value(
                "do nadbudowania na suficie,do wbudowania w sufit (sufit kasetonowy)",
                "sposob_montazu",
                "Miejsce montażu",
            ),
            "",
        )

    def test_max_bulb_power_is_separated_from_fixture_power(self) -> None:
        self.assertEqual(normalize_parameter_value("max 60", "moc_max_zrodla", "Moc maksymalna [W]"), "60W")
        self.assertEqual(normalize_parameter_value("2 x max 40", "moc_max_zrodla", "Moc maksymalna [W]"), "40W")
        self.assertEqual(normalize_parameter_value("2 x 36", "moc_max_zrodla", "Moc maksymalna [W]"), "36W")

    def test_integrated_led_max_power_is_fixture_power(self) -> None:
        params = pd.DataFrame(
            [
                {"Kod Kanlux": "1", "Nazwa": "ANTO", "ID atrybutu": "1", "Nazwa atrybutu": "Moc maksymalna [W]", "Wartość": "max 37"},
                {"Kod Kanlux": "1", "Nazwa": "ANTO", "ID atrybutu": "2", "Nazwa atrybutu": "Zintegrowane źródło światła LED", "Wartość": "1"},
                {"Kod Kanlux": "2", "Nazwa": "STONO", "ID atrybutu": "1", "Nazwa atrybutu": "Moc maksymalna [W]", "Wartość": "max 25"},
                {"Kod Kanlux": "2", "Nazwa": "STONO", "ID atrybutu": "2", "Nazwa atrybutu": "Trzonek", "Wartość": "E27"},
                {"Kod Kanlux": "3", "Nazwa": "DICHT", "ID atrybutu": "1", "Nazwa atrybutu": "Moc maksymalna [W]", "Wartość": "2 x 36"},
            ]
        )
        columns = {"sku": "Kod Kanlux", "name": "Nazwa", "attribute_id": "ID atrybutu", "attribute_name": "Nazwa atrybutu", "value": "Wartość"}

        index, _ = build_parameter_index(params, columns)

        self.assertEqual(index["1"]["moc"]["value"], "37W")
        self.assertNotIn("moc_max_zrodla", index["1"])
        self.assertEqual(index["2"]["moc_max_zrodla"]["value"], "25W")
        self.assertNotIn("moc", index["2"])
        self.assertEqual(index["3"]["moc_max_zrodla"]["value"], "36W")
        self.assertEqual(index["3"]["liczba_gniazd"]["value"], "2")

    def test_normalizes_new_lighting_values(self) -> None:
        self.assertEqual(normalize_parameter_value("< 19", "ugr", "UGR"), "<19")
        self.assertEqual(normalize_parameter_value("80", "cri", "Współczynnik oddawania barw Ra"), "80")
        self.assertEqual(normalize_parameter_value("50 000", "trwalosc", "Trwałość [h]"), "50000")
        self.assertEqual(
            normalize_parameter_value("1", "laczenie_przelotowe", "Możliwość łączenia przelotowego opraw"),
            "Tak",
        )
        self.assertEqual(
            normalize_parameter_value("DALI", "sciemnianie", "Możliwość współpracy ze ściemniaczem"),
            "DALI",
        )
        self.assertEqual(
            normalize_parameter_value("tak", "zrodlo_w_komplecie", "Źródło światła w komplecie"),
            "Tak",
        )
        self.assertEqual(
            normalize_parameter_value("1", "zrodlo_w_komplecie", "Zintegrowane źródło światła LED"),
            "Tak",
        )


if __name__ == "__main__":
    unittest.main()
