from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fill_variant_tags_from_siblings import (  # noqa: E402
    choose_tags,
    color_code,
    fixed_sibling_tags,
    decide_tags,
    split_tags,
    tag_slug,
)
from fix_simon_attributes_to_shop_model import load_dictionary  # noqa: E402


def row(sku: str, tags: str | None, status: str = "publish") -> dict:
    return {"Sku": sku, "Title": sku, "Status": status, "Product Tags": tags}


class HelpersTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()

    def test_color_code(self) -> None:
        self.assertEqual(color_code("TR1/142/KON", self.config), "142")
        self.assertEqual(color_code("DW7.01/X/42/KON", self.config), "42")
        self.assertIsNone(color_code("G22/KON", self.config))

    def test_tag_slug_matches_wordpress_style(self) -> None:
        self.assertEqual(tag_slug("Łącznik schodowy pojedynczy 10 A 55 simon"), "lacznik-schodowy-pojedynczy-10-a-55-simon")

    def test_split_tags(self) -> None:
        self.assertEqual(split_tags("A|B"), ["A", "B"])
        self.assertEqual(split_tags(None), [])


class ChooseTagsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()

    def test_unanimous(self) -> None:
        tags, source, _ = choose_tags([["Ramka X"], ["Ramka X"]], "TRX", self.config)
        self.assertEqual((tags, source), (["Ramka X"], "jednoznaczny tag rodzenstwa"))

    def test_event_tag_is_not_copied(self) -> None:
        tags, source, _ = choose_tags([["ElektroTargi Alkan 2026", "Ramka X"], ["Ramka X"]], "TRX", self.config)
        self.assertEqual(tags, ["Ramka X"])
        self.assertEqual(source, "jednoznaczny tag rodzenstwa")

    def test_majority_wins_over_stray_tag(self) -> None:
        tags, source, _ = choose_tags([["Ramka 1"], ["Ramka 2"], ["Ramka 2"], []], "TRX", self.config)
        self.assertEqual((tags, source), (["Ramka 2"], "wiekszosc rodzenstwa"))

    def test_even_split_needs_decision(self) -> None:
        tags, source, _ = choose_tags([["A"], ["B"]], "NIEZNANY", self.config)
        self.assertEqual(tags, [])
        self.assertIn("wymaga decyzji", source)

    def test_override_from_dictionary(self) -> None:
        tags, source, _ = choose_tags([["A"], ["B"]], "T62E.01", self.config)
        self.assertEqual(tags, ["Gniazdo komputerowe podwójne RJ45 kategoria 6 ekranowane 55 simon"])
        self.assertIn("overrides", source)

    def test_siblings_without_tags(self) -> None:
        tags, source, _ = choose_tags([[], []], "TKW1", self.config)
        self.assertEqual(tags, [])
        self.assertEqual(source, "rodzenstwo nie ma tagow wariantowych")


class DecideTagsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_dictionary()

    def test_new_color_gets_family_tag_and_variant_with_x_is_separate(self) -> None:
        rows = [
            row("TW6.01/111/KON", "Schodowy"),
            row("TW6.01/149/KON", "Schodowy"),
            row("TW6.01/X/111/KON", "Schodowy bez piktogramu"),
            row("TW6.01/142/KON", None, "draft"),
            row("TW6.01/X/142/KON", None, "draft"),
        ]
        decisions, anomalies = decide_tags(rows, {"142"}, self.config, {})
        by_sku = {item.sku: item for item in decisions}
        self.assertEqual(by_sku["TW6.01/142/KON"].tags, ["Schodowy"])
        self.assertEqual(by_sku["TW6.01/X/142/KON"].tags, ["Schodowy bez piktogramu"])
        self.assertEqual(anomalies, [])

    def test_reports_sibling_anomalies(self) -> None:
        rows = [
            row("TR2/111/KON", "Ramka 1"),
            row("TR2/114/KON", "Ramka 2"),
            row("TR2/116/KON", "Ramka 2"),
            row("TR2/149/KON", None),
            row("TR2/142/KON", None, "draft"),
        ]
        decisions, anomalies = decide_tags(rows, {"142"}, self.config, {})
        self.assertEqual(decisions[0].tags, ["Ramka 2"])
        self.assertEqual({item["sku"] for item in anomalies}, {"TR2/111/KON", "TR2/149/KON"})

    def test_fix_keeps_event_tag_and_replaces_wrong_variant_tag(self) -> None:
        rows = [
            row("TR2/111/KON", "ElektroTargi Alkan 2026|Ramka 1"),
            row("TR2/114/KON", "Ramka 2"),
            row("TR2/116/KON", "Ramka 2"),
            row("TR2/142/KON", None, "draft"),
        ]
        _, anomalies = decide_tags(rows, {"142"}, self.config, {})
        self.assertEqual(len(anomalies), 1)
        self.assertEqual(fixed_sibling_tags(anomalies[0]), "ElektroTargi Alkan 2026|Ramka 2")

    def test_no_siblings(self) -> None:
        decisions, _ = decide_tags([row("DSC/42/KON", None, "draft")], {"42"}, self.config, {})
        self.assertEqual(decisions[0].tags, [])
        self.assertEqual(decisions[0].source, "brak rodzenstwa w pliku")


if __name__ == "__main__":
    unittest.main()
