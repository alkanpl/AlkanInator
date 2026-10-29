import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fill_panel_dimensions import parse_dims_cell, parse_single_mm, parse_title  # noqa: E402


class FillPanelDimensionsTest(unittest.TestCase):
    def test_title_footprint_units(self):
        self.assertEqual(parse_title("Panel LED BLINGO 595x595mm 34W")["footprint"], [595, 595])
        self.assertEqual(parse_title("Panel LED BRISBANE 30x120 36W")["footprint"], [1200, 300])
        self.assertEqual(parse_title("Panel LED mini 29x4,5 20W")["footprint"], [290, 45])

    def test_title_round_and_height(self):
        self.assertEqual(parse_title("Panel LED mini okrągły 8,4 cm 3W")["diameter"], 84)
        self.assertEqual(parse_title("Ramka do panelu LED 63mm 30x120 klik")["height"], 63)
        self.assertEqual(parse_title("SURFACE MOUNT KIT H70 VALUE")["height"], 70)

    def test_cell_is_mm_unless_matches_title_in_cm(self):
        self.assertEqual(parse_dims_cell("100 x 100", [100, 100]), [100, 100])
        self.assertEqual(parse_dims_cell("60x60", [600, 600]), [600, 600])
        self.assertEqual(parse_dims_cell("595 x 595 x 31"), [595, 595, 31])

    def test_single_value_units(self):
        self.assertEqual(parse_single_mm("60 cm"), 600)
        self.assertEqual(parse_single_mm("38mm"), 38)


if __name__ == "__main__":
    unittest.main()
