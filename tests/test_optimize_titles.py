"""Testy charakteryzujace dla generowania tytulow SEO (caly rdzen pipeline'u).

Golden-file: fixture ``title_snapshot.json`` zawiera surowe wiersze produktow
oraz oczekiwany ``new_title`` po przejsciu przez extract -> role -> optimize.
Test odbudowuje DataFrame z fixture i przepuszcza go przez ten sam rdzen, wiec
NIE wymaga pliku XLSX. Po swiadomej zmianie regul/configow zregeneruj snapshot:

    py tests/generate_snapshots.py
"""

from __future__ import annotations

import io
import json
import sys
import unittest
from pathlib import Path

import pandas as pd

TESTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TESTS_DIR))
sys.path.insert(0, str(TESTS_DIR.parent / "src"))

from generate_snapshots import build_pipeline_titles  # noqa: E402

FIXTURE = TESTS_DIR / "fixtures" / "title_snapshot.json"


class TestTitleSnapshot(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        with io.open(FIXTURE, encoding="utf-8") as handle:
            cls.cases = json.load(handle)
        source_df = pd.DataFrame([case["source"] for case in cls.cases])
        cls.optimized = build_pipeline_titles(source_df).reset_index(drop=True)

    def test_snapshot_not_empty(self) -> None:
        self.assertGreaterEqual(len(self.cases), 30)

    def test_new_titles_match_snapshot(self) -> None:
        mismatches = []
        for position, case in enumerate(self.cases):
            expected = case["expected"]["new_title"]
            got = str(self.optimized.at[position, "new_title"]).strip()
            if got != expected:
                mismatches.append((case["expected"].get("new_title", ""), expected, got))
        if mismatches:
            lines = [
                f"\n  EXPECTED: {exp}\n  GOT:      {got}" for _, exp, got in mismatches
            ]
            self.fail(
                f"{len(mismatches)} tytul(ow) sie zmienilo. Jesli to zamierzone, "
                f"uruchom: py tests/generate_snapshots.py" + "".join(lines)
            )

    def test_product_roles_match_snapshot(self) -> None:
        for position, case in enumerate(self.cases):
            expected = case["expected"]["product_role"]
            got = str(self.optimized.at[position, "product_role"]).strip()
            self.assertEqual(got, expected, f"rola w wierszu {position}")

    def test_every_row_has_non_empty_title(self) -> None:
        for position in range(len(self.cases)):
            title = str(self.optimized.at[position, "new_title"]).strip()
            self.assertTrue(title, f"pusty new_title w wierszu {position}")


if __name__ == "__main__":
    unittest.main()
