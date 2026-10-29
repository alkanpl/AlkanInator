import sys
import unittest
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from fill_product_weights import max_height_cm, packing_rule, pieces_per_package, shipping_weight  # noqa: E402

PACKING = yaml.safe_load((ROOT / "dictionaries" / "panel_shipping_packing.yaml").read_text(encoding="utf-8"))


class PanelPackingTest(unittest.TestCase):
    def per_package(self, title, sku="X/AAA"):
        return packing_rule(title, sku, PACKING)[1]["per_package"]

    def test_brand_rules(self):
        self.assertEqual(self.per_package("Panel LED 60x60 34W RC065B LEDINAIRE PHILIPS 911401885080"), 8)
        self.assertEqual(self.per_package("Panel LED 60×60 40W RC007B PILA IP20 neutralny 911401879584"), 8)
        self.assertEqual(self.per_package("Panel LED 60x60 600 E 36W biały Ledvance 4099854705052"), 12)
        self.assertEqual(self.per_package("Panel LED 600 E 36W", "4099854705052/LDV"), 12)
        self.assertEqual(self.per_package("Panel LED BRISBANE 60x60 36W Kobi Premium 001559"), 10)
        # "PILA" tylko jako osobne slowo
        self.assertEqual(self.per_package("Panel LED 60x60 PILAR 36W"), 10)

    def test_shipping_weight_splits_after_per_package(self):
        self.assertEqual(shipping_weight(8, PACKING), 3.6)
        self.assertEqual(shipping_weight(12, PACKING), 2.4)
        self.assertEqual(shipping_weight(10, PACKING), 2.88)
        for n in (8, 10, 12):
            self.assertEqual(pieces_per_package(shipping_weight(n, PACKING), 30), n)

    def test_pieces_threshold_is_exclusive(self):
        self.assertEqual(pieces_per_package(2.5, 30), 11)  # 12 x 2,5 = 30,0 -> juz przedzial 30-60
        self.assertEqual(pieces_per_package(61 * 61 * 6 / 5000, 30), 6)

    def test_max_height_keeps_dimensional_weight_below_limit(self):
        h = max_height_cm(61, 61, 8, 5000, 30)
        self.assertEqual(h, 5.0)
        self.assertEqual(pieces_per_package(61 * 61 * h / 5000, 30), 8)


if __name__ == "__main__":
    unittest.main()
