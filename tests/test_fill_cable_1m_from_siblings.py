from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fill_cable_1m_from_siblings import (  # noqa: E402
    ATTR_PREFIX,
    content_for_1m,
    decide_row,
    find_siblings,
    fix_title,
    is_new_row,
    mirror_taxonomy_columns,
    select_new_rows,
    title_derived,
    voltage_from_title,
    weight_per_meter,
)

ATTRS = [ATTR_PREFIX + name for name in ("Długość", "Producent", "Kolor", "Dane producenta", "Napięcie [V]", "Typ izolacji", "Ekranowany", "Żyła ochronna")]


def row(row_id: str, sku: str, title: str, tags: str = "", ship: str = "", **attrs: str) -> dict:
    data = {"id": row_id, "Sku": sku, "Title": title, "Product Tags": tags, "Shipping Class": ship, "hwp_product_gtin": "590"}
    data.update({ATTR_PREFIX + k.replace("_", " "): v for k, v in attrs.items()})
    return data


def sibling(length: str, tags: str = "kampania|PopUpKablePaczkomat|Przewód YDY") -> dict:
    return row(f"s{length}", f"590/EKB/{length}m", f"Przewód YDY {length}m", tags, "Kable i przewody do 15 kg",
               Producent="Polski Producent", Kolor="Biały", **{"Długość": f"{length} m", "Dane producenta": "ELEKTROKABEL, Turek"})


def changes_by_field(changes) -> dict:
    return {c.field.replace(ATTR_PREFIX, ""): c for c in changes}


class HelpersTest(unittest.TestCase):
    def test_voltage_takes_higher_value_of_pair(self) -> None:
        self.assertEqual(voltage_from_title("Przewód YDY żo 3×1 450/750V biały 1m"), "750")
        self.assertEqual(voltage_from_title("Przewód instalacyjny YDYpżo 450/750 4×1,5 biały 1m"), "750")
        self.assertEqual(voltage_from_title("H05VV-F OWY żo 4×0,5 300/500V biały 1m"), "500")
        self.assertEqual(voltage_from_title("H05VV-F OWY 5×6 500 V biały 1m"), "500")
        self.assertEqual(voltage_from_title("Przewód bez napięcia 3x1,5"), "")

    def test_title_derived_attributes(self) -> None:
        self.assertEqual(
            title_derived("Przewód instalacyjny YDYżo 450/750 5×1,5 biały 1m"),
            {"Typ izolacji": "PVC", "Ekranowany": "Nie", "Napięcie [V]": "750", "Żyła ochronna": "Tak"},
        )
        self.assertNotIn("Żyła ochronna", title_derived("H05VV-F OWY 2×2,5 500 V czarny 1m"))

    def test_fix_title_replaces_producer_with_length(self) -> None:
        self.assertEqual(fix_title("Przewód OWY 4x1,5 500 V biały Elektrokabel")[0], "Przewód OWY 4x1,5 500 V biały 1m")
        self.assertEqual(fix_title("Przewód YDY 5x1 750V biały 1m"), ("Przewód YDY 5x1 750V biały 1m", []))

    def test_fix_title_normalizes_cross_section_voltage_and_color(self) -> None:
        self.assertEqual(fix_title("Przewód instalacyjny YDYpżo 450/750 4×1,5 biały 1m")[0],
                         "Przewód instalacyjny YDYpżo 450/750V 4x1,5 biały 1m")
        self.assertEqual(fix_title("Przewód YDY 5X1 750V biały 1m")[0], "Przewód YDY 5x1 750V biały 1m")
        self.assertEqual(fix_title("Przewód płaski YDYp żo 2×1 750V 1m", "Biały")[0], "Przewód płaski YDYp żo 2x1 750V biały 1m")
        self.assertEqual(fix_title("Przewód H05VV-F OWY żo 4×0,5 300/500V biały 1m", "Biały")[0],
                         "Przewód H05VV-F OWY żo 4x0,5 300/500V biały 1m")

    def test_siblings_match_by_base_sku_only(self) -> None:
        new = row("n", "590/EKB", "Przewód YDY 1m", "Odcinek 1m")
        other = row("x", "5901/EKB/25m", "inny")
        rows = [new, sibling("25"), sibling("100"), other]
        self.assertTrue(is_new_row(new))
        self.assertEqual([s["id"] for s in find_siblings(new, rows)], ["s25", "s100"])


class DecideRowTest(unittest.TestCase):
    def test_fills_empty_row_from_siblings_and_title(self) -> None:
        new = row("n", "590/EKB", "Przewód YDY żo 3×1 450/750V biały 1m", "Domyślny dla grupy|Odcinek 1m")
        siblings = [sibling("25"), sibling("50"), sibling("100", "DISABLE|kampania|PopUpKablePaczkomat|Przewód YDY")]
        result = changes_by_field(decide_row(new, siblings, ATTRS))

        self.assertEqual(result["Długość"].new, "1 m")
        self.assertEqual(result["Producent"].new, "Polski Producent")
        self.assertEqual(result["Kolor"].new, "Biały")
        self.assertEqual(result["Dane producenta"].new, "ELEKTROKABEL, Turek")
        self.assertEqual(result["Napięcie [V]"].source, "tytul")
        self.assertEqual(result["Shipping Class"].new, "Kable i przewody do 15 kg")
        # DISABLE jest tylko na 100 m, wiec nie trafia do odcinka.
        self.assertEqual(result["Product Tags"].new, "Domyślny dla grupy|Odcinek 1m|kampania|PopUpKablePaczkomat|Przewód YDY")

    def test_producer_is_always_polish_and_wrong_insulation_is_fixed(self) -> None:
        new = row("n", "590/EKB", "Przewód YDY żo 2×1,5 450/750V biały 1m", "Odcinek 1m",
                  Producent="Elektrokabel", **{"Długość": "1m", "Typ izolacji": "X - polietylen"})
        result = changes_by_field(decide_row(new, [sibling("25")], ATTRS))
        self.assertEqual(result["Producent"].new, "Polski Producent")
        self.assertEqual(result["Długość"].new, "1 m")
        self.assertEqual((result["Typ izolacji"].new, result["Typ izolacji"].status), ("PVC", "WARNING"))

    def test_group_tag_follows_majority_of_siblings(self) -> None:
        new = row("n", "590/EKB", "Przewód YDY żo 3×1 450/750V biały 1m", "Odcinek 1m")
        siblings = [
            sibling("25", "kampania|PopUpKablePaczkomat|Przewód YDY"),
            sibling("50", "PopUpKablePaczkomat|Przewód YDY"),
            sibling("100", "DISABLE|kampania|PopUpKablePaczkomat|Przewód YDYżo"),
        ]
        result = changes_by_field(decide_row(new, siblings, ATTRS))
        self.assertEqual(result["Product Tags"].new, "Odcinek 1m|kampania|PopUpKablePaczkomat|Przewód YDY")

    def test_row_without_siblings_is_reported(self) -> None:
        new = row("n", "590/EKB", "Przewód YDY 1m", "Odcinek 1m")
        [change] = decide_row(new, [], ATTRS)
        self.assertEqual(change.status, "DO_SPRAWDZENIA")


class ExportWithoutTagsTest(unittest.TestCase):
    """Eksport z kolumnami pa_ i bez tagow (2026-10-13): tagi rodzenstwa z --tags-source."""

    def setUp(self) -> None:
        self.new = row("n", "590/EKB", "Przewód YDY żo 3x2,5 450/750V biały 1m")
        self.new.update({"Weight": "0", "Content": " ", "Status": "draft", "Product Tags": None})
        self.siblings = []
        for length, weight in (("25", "3.1"), ("50", "6.2"), ("100", "12.4")):
            s = sibling(length)
            s.update({"Weight": weight, "Status": "publish", ATTR_PREFIX + "Liczba żył": "3", ATTR_PREFIX + "Przekrój żyły [mm]": "2,5",
                      "Content": f"<h2>Co wyróżnia przewód YDY biały {length}m?</h2>\n<ul>\n <li>Długość: 450 m</li>\n <li>Przekrój 2,5 mm²</li>\n</ul>",
                      "Attribute Value (pa_kolor)": "Biały", "Attribute In Variations (pa_kolor)": "yes",
                      "Attribute Is Visible (pa_kolor)": "yes", "Attribute Is Taxonomy (pa_kolor)": "yes"})
            self.siblings.append(s)

    def test_new_rows_are_drafts_with_base_sku_and_siblings(self) -> None:
        lonely = dict(row("x", "8PYDY", "Przewód YDY"), Status="draft")
        published = dict(row("p", "591/EKB", "Przewód"), Status="publish")
        rows = [self.new, lonely, published, *self.siblings]
        self.assertEqual([r["id"] for r in select_new_rows(rows, has_tags=False)], ["n"])

    def test_weight_content_tags_and_gtin(self) -> None:
        for s in self.siblings:
            s["hwp_product_gtin"] = "5907702810393"
        self.new["hwp_product_gtin"] = None
        result = changes_by_field(decide_row(self.new, self.siblings, ATTRS))
        self.assertEqual(result["Weight"].new, "0.124")
        self.assertEqual(result["hwp_product_gtin"].new, "5907702810393")
        self.assertEqual(result["Content"].source, "rodzenstwo 100 m")
        self.assertEqual(result["Content"].new, "<h2>Co wyróżnia przewód YDY biały?</h2>\n<ul>\n <li>Przekrój 2,5 mm²</li>\n</ul>")
        self.assertTrue(result["Product Tags"].new.startswith("Domyślny dla grupy|Odcinek 1m|kampania"))

    def test_weight_below_copper_mass_is_not_written(self) -> None:
        for s, weight in zip(self.siblings, ("0.175", "0.35", "0.7")):
            s["Weight"] = weight  # 0,007 kg/m przy 3x2,5 mm2 miedzi (0,067 kg/m) - niemozliwe
        changes = decide_row(self.new, self.siblings, ATTRS)
        self.assertNotIn("Weight", changes_by_field(changes))
        self.assertTrue(any(c.status == "DO_SPRAWDZENIA" and c.note.startswith("Weight") for c in changes))

    def test_weight_per_meter_flags_inconsistent_siblings(self) -> None:
        self.assertEqual(weight_per_meter(self.siblings), ("0.124", True))
        self.siblings[0]["Weight"] = "5"
        self.assertFalse(weight_per_meter(self.siblings)[1])

    def test_content_for_1m_keeps_other_measures(self) -> None:
        self.assertEqual(content_for_1m("YDY 3x2,5 mm² 100 m, prąd 25 mA")[0], "YDY 3x2,5 mm², prąd 25 mA")

    def test_taxonomy_columns_follow_attribute_changes(self) -> None:
        changes = decide_row(self.new, self.siblings, ATTRS)
        mirrored = changes_by_field(mirror_taxonomy_columns(changes, self.new, self.siblings, {"Kolor": "pa_kolor"}))
        self.assertEqual(mirrored["Attribute Value (pa_kolor)"].new, "Biały")
        self.assertEqual(mirrored["Attribute In Variations (pa_kolor)"].new, "yes")


if __name__ == "__main__":
    unittest.main()
