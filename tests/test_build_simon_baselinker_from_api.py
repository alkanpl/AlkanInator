import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from build_simon_baselinker_from_api import (  # noqa: E402
    SimonRow,
    build_features,
    build_name,
    check_against_api,
    load_names,
    product_key,
    split_symbol,
)


class SymbolTests(unittest.TestCase):
    def test_split_symbol_with_transparent_flap(self):
        self.assertEqual(split_symbol("NGZ1BCZ.01/38A"), ("NGZ1BCZ.01", "38", True))
        self.assertEqual(product_key("NGZ1BCZ.01/38A"), "NGZ1BCZ.01/..A")

    def test_split_symbol_keeps_variant_marker(self):
        self.assertEqual(split_symbol("NW6/2.01/X/11"), ("NW6/2.01/X", "11", False))

    def test_symbol_without_color(self):
        self.assertEqual(split_symbol("NUV2"), ("NUV2", "", False))
        self.assertEqual(split_symbol("NTC-03"), ("NTC-03", "", False))


class NameTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = load_names()

    def test_white_inflects_by_gender(self):
        self.assertEqual(build_name("NR1/11", self.config), "Ramka 1-krotna biała Simon 24 NR1/11")
        self.assertEqual(build_name("NW1.01/11", self.config), "Łącznik jednobiegunowy 10 AX Push In biały Simon 24 NW1.01/11")
        self.assertEqual(build_name("N51.01/11", self.config), "Gniazdo komputerowe pojedyncze RJ45 kat. 5e białe Simon 24 N51.01/11")

    def test_matte_colors_do_not_inflect(self):
        self.assertEqual(build_name("NR1/38", self.config), "Ramka 1-krotna grafitowy mat Simon 24 NR1/38")

    def test_transparent_flap_variant_has_own_stem(self):
        self.assertIn("z klapką transparentną", build_name("NGZ1BCZ.01/49A", self.config))
        self.assertNotIn("transparent", build_name("NGZ1BCZ.01/11", self.config))

    def test_basic_seal_is_not_named_simon_24(self):
        name = build_name("NU1", self.config)
        self.assertEqual(name, "Uszczelka IP44 do ramki 1-krotnej Simon Basic Standard NU1")

    def test_every_name_ends_with_symbol(self):
        for symbol in ["NUV2", "NTC-03", "NSC/49", "NW7/2.01/X/11", "NGZ1BCU.01/34A"]:
            self.assertTrue(build_name(symbol, self.config).endswith(symbol), symbol)


class FeatureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = load_names()

    def test_frame_has_material_and_no_mounting(self):
        spec = self.config["products"]["NR1"]
        features = build_features("NR1/34", spec, {"Stopień ochrony IP": "IP20/IP44"}, self.config)
        self.assertEqual(features["Kolor"], "Złoty")
        self.assertEqual(features["Kolor producenta"], "Szampański")
        self.assertEqual(features["Materiał"], "Tworzywo sztuczne")
        self.assertNotIn("Sposób montażu", features)
        self.assertEqual(features["Stopień ochrony [IP]"], "IP20/IP44")

    def test_simon_go_gets_wifi_and_ip_format(self):
        spec = self.config["products"]["NEW1W.01"]
        features = build_features("NEW1W.01/49", spec, {"Stopień ochrony IP": "IP20"}, self.config)
        self.assertEqual(features["Typ łączności"], "Wi-Fi")
        self.assertEqual(features["Stopień ochrony [IP]"], "IP 20")
        self.assertEqual(features["Sposób montażu"], "Podtynkowy")
        self.assertEqual(features["Seria"], "Simon 24")


class ApiCheckTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = load_names()

    def _row(self, symbol):
        return SimonRow(lp="1", sku=f"{symbol}/KON", symbol=symbol, ekonom_name="")

    def test_push_in_mismatch_is_flagged(self):
        row = self._row("NW7.01/11")
        row.api_series, row.api_color = "Simon 24", "biały"
        product = {"symbol": "NW7.01/11", "name": "Łącznik krzyżowy"}
        check_against_api(row, self.config["products"]["NW7.01"], product, {"montaż": "mocowanie Push In"}, self.config)
        self.assertEqual(row.status, "DO_SPRAWDZENIA")

    def test_color_mismatch_is_flagged(self):
        row = self._row("NR1/38")
        row.api_series, row.api_color = "Simon 24", "czarny mat"
        check_against_api(row, self.config["products"]["NR1"], {"symbol": "NR1/38", "name": "Ramka"}, {}, self.config)
        self.assertEqual(row.status, "DO_SPRAWDZENIA")

    def test_missing_product_is_flagged(self):
        row = self._row("NDF")
        check_against_api(row, self.config["products"]["NDF"], None, {}, self.config)
        self.assertEqual(row.status, "DO_SPRAWDZENIA")


if __name__ == "__main__":
    unittest.main()
