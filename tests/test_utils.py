"""Testy walidacji EAN/GTIN (suma kontrolna dla 8/12/13/14 cyfr)."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from utils import valid_ean  # noqa: E402


class TestValidEan(unittest.TestCase):
    def test_valid_codes_across_lengths(self) -> None:
        for code in ["73513537", "036000291452", "4006381333931", "10614141000415"]:
            self.assertTrue(valid_ean(code), code)

    def test_rejects_bad_check_digit(self) -> None:
        self.assertFalse(valid_ean("73513530"))  # EAN-8
        self.assertFalse(valid_ean("5901234567896"))  # EAN-13

    def test_rejects_wrong_length(self) -> None:
        self.assertFalse(valid_ean("1234567"))
        self.assertFalse(valid_ean("123456789012345"))

    def test_rejects_all_same_digit(self) -> None:
        self.assertFalse(valid_ean("0000000000000"))

    def test_strips_non_digits_before_check(self) -> None:
        self.assertTrue(valid_ean(" 4006381333931 "))


if __name__ == "__main__":
    unittest.main()
