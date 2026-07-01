"""Testy czytelnych rozroznikow wariantow (Faza B) - zamiana surowych kodow
Kanlux na fraze po polsku, oraz krotnosci mocy w oprawach hermetycznych."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

TESTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TESTS_DIR.parent / "src"))

from optimize_titles import (  # noqa: E402
    hermetic_bulb_count,
    panel_format_from_supplier_name,
    readable_duplicate_variant,
)


def variant(nazwa: str, typ: str = "", **extra) -> str:
    row = pd.Series({"Nazwa Kanlux": nazwa, "attr_typ": typ, **extra})
    return readable_duplicate_variant(row)


class ReadableVariantTests(unittest.TestCase):
    def test_soczewka_zwraca_moc_docelowej_oprawy(self) -> None:
        self.assertEqual(variant("HBPHS LENS 200W 60D", "Soczewka", **{"Moc [W]": "200"}), "200W")
        self.assertEqual(variant("HBPHS LENS 150W 120D", "Soczewka", **{"Moc [W]": "150"}), "150W")

    def test_klips_format_panela_i_ilosc(self) -> None:
        self.assertEqual(
            variant("CLIPS PANEL 60-62", "Klips", attr_ilosc_sztuk="4 szt."),
            "do paneli 60x60/62x62 cm 4 szt.",
        )
        self.assertEqual(variant("CLIPS PANEL 120", "Klips"), "do paneli 120x30 cm")

    def test_dicht_klosz_i_odblysnik(self) -> None:
        self.assertEqual(variant("DICHT 4LED NP 236/PS", "oprawa hermetyczna LED"), "klosz PS")
        self.assertEqual(
            variant("DICHT 4LED PI 236/PC", "oprawa hermetyczna LED"), "z odbłyśnikiem klosz PC"
        )

    def test_milo_rodzaj_siatki(self) -> None:
        self.assertEqual(variant("MILO 7040T", "Oprawa kanałowa"), "siatka metalowa")
        self.assertEqual(variant("MILO 7040T/P", "Oprawa kanałowa"), "siatka plastikowa")

    def test_jasmin_c_to_wiszacy(self) -> None:
        self.assertEqual(variant("JASMIN C 370-W/M", "plafoniera drewniana"), "wiszący")
        self.assertEqual(variant("JASMIN 370-W/M", "plafoniera drewniana"), "")


class PanelFormatTests(unittest.TestCase):
    def test_rozmiary(self) -> None:
        self.assertEqual(panel_format_from_supplier_name("CLIPS PANEL 60-62"), "60x60/62x62 cm")
        self.assertEqual(panel_format_from_supplier_name("SPN BRAVO S/P 12030"), "120x30 cm")
        self.assertEqual(panel_format_from_supplier_name("CLIPS PANEL 120"), "120x30 cm")
        self.assertEqual(panel_format_from_supplier_name("COS INNEGO"), "")


class HermeticBulbCountTests(unittest.TestCase):
    def test_mah_plus_krotnosc_z_kodu(self) -> None:
        # MAH PLUS-258 = 2x58W; -158 = pojedyncza (1).
        values2 = {"typ": "oprawa hermetyczna na T8", "nazwa_kanlux": "MAH PLUS-258/4LED/PC", "liczba_zrodel": ""}
        values1 = {"typ": "oprawa hermetyczna na T8", "nazwa_kanlux": "MAH PLUS-158/4LED/PC", "liczba_zrodel": ""}
        self.assertEqual(hermetic_bulb_count(values2, "58W"), 2)
        self.assertEqual(hermetic_bulb_count(values1, "58W"), 0)

    def test_liczba_zrodel_ma_priorytet(self) -> None:
        values = {"typ": "oprawa hermetyczna LED", "nazwa_kanlux": "", "liczba_zrodel": "2"}
        self.assertEqual(hermetic_bulb_count(values, "36W"), 2)

    def test_nie_dotyczy_nie_hermetycznych(self) -> None:
        values = {"typ": "plafon led", "nazwa_kanlux": "X-258", "liczba_zrodel": ""}
        self.assertEqual(hermetic_bulb_count(values, "58W"), 0)


if __name__ == "__main__":
    unittest.main()
