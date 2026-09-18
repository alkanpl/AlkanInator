from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from restore_woo_description_scripts import decide, has_valid_ld_json, normalize_without_scripts  # noqa: E402

OURS = '<p>Opis</p>\n<script type="application/ld+json">\n{"@type": "FAQPage"}\n</script>'
STRIPPED = '<p>Opis</p>\n\n{"@type": "FAQPage"}\n'


class RestoreScriptsTest(unittest.TestCase):
    def test_stripped_content_matches_ours(self) -> None:
        self.assertEqual(normalize_without_scripts(STRIPPED), normalize_without_scripts(OURS))
        self.assertEqual(decide(STRIPPED, OURS), "przywrocono")

    def test_manual_edit_in_shop_is_not_overwritten(self) -> None:
        self.assertEqual(decide('<p>Opis poprawiony recznie</p>{"@type": "FAQPage"}', OURS), "tresc w Woo rozni sie od naszej - NIE nadpisano")

    def test_already_correct_is_left_alone(self) -> None:
        self.assertEqual(decide(OURS, OURS), "Woo ma juz poprawny script - bez zmian")

    def test_invalid_json_ld_is_rejected(self) -> None:
        broken = '<p>Opis</p><script type="application/ld+json">{nie json}</script>'
        self.assertFalse(has_valid_ld_json(broken))
        self.assertEqual(decide("<p>Opis</p>{nie json}", broken), "nasz opis nie ma poprawnego JSON-LD")


if __name__ == "__main__":
    unittest.main()
