from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fill_cable_attributes_from_references import (  # noqa: E402
    ATTR_PREFIX,
    decide,
    length_from,
    load_rules,
    voltage_from_title,
)

NAMES = ("Typ", "Kolor", "Liczba żył", "Przekrój żyły [mm]", "Napięcie [V]", "Typ izolacji", "Ekranowany",
         "Żyła ochronna", "Producent", "Długość", "Dane producenta")


def row(row_id: str, sku: str, title: str, **attrs: str) -> dict:
    data = {"id": row_id, "Sku": sku, "Title": title}
    data.update({ATTR_PREFIX + name: None for name in NAMES})
    data.update({ATTR_PREFIX + k.replace("_", " "): v for k, v in attrs.items()})
    return data


def by_attr(changes) -> dict:
    return {(c.row_id, c.attribute): c for c in changes}


class HelpersTest(unittest.TestCase):
    def test_voltage_from_title(self) -> None:
        self.assertEqual(voltage_from_title("Przewód YDYżo 450/750V 3x4 biały 1m"), "750")
        self.assertEqual(voltage_from_title("H05VV-F OWY żo 3x1,5 300/500V"), "500")
        self.assertEqual(voltage_from_title("Kabel YKY 5x2,5 0,6/1kV"), "1000")
        self.assertEqual(voltage_from_title("OWY 5x1 500 V czarny"), "500")
        self.assertEqual(voltage_from_title("Przewód YDYp 3x1,5"), "")

    def test_length_from_sku_or_title_end(self) -> None:
        self.assertEqual(length_from("590/EKB/100m", "Przewód 100m"), "100")
        self.assertEqual(length_from("590/EKB", "Przewód YDY 2x2,5 450/750V biały 1m"), "1")
        self.assertEqual(length_from("713531/HLK", "Kabel H1Z2Z2-K 1x6mm2 1m czarny 713531"), "1")
        self.assertEqual(length_from("PZ/ANT", "Przewód YDY 4X10"), "")


class DecideTest(unittest.TestCase):
    rules = load_rules()

    def test_new_1m_row_takes_siblings_title_and_type(self) -> None:
        new = row("n", "590/EKB", "Przewód instalacyjny YDYżo 450/750V 4x1,5 biały 1m")
        sibling = row("s", "590/EKB/100m", "Przewód YDYżo 450/750 4x1,5 biały 100m", Typ="YDY", Kolor="Biały",
                      **{"Liczba żył": "4", "Przekrój żyły [mm]": "1,5", "Producent": "Polski Producent", "Długość": "100 m"})
        result = by_attr(decide([new, sibling], self.rules)[0])
        self.assertEqual(result[("n", "Typ")].new, "YDY")
        self.assertEqual(result[("n", "Liczba żył")].source, "rodzenstwo")
        self.assertEqual(result[("n", "Napięcie [V]")].new, "750")
        self.assertEqual(result[("n", "Typ izolacji")].new, "PVC")
        self.assertEqual(result[("n", "Ekranowany")].new, "Nie")
        self.assertEqual(result[("n", "Żyła ochronna")].new, "Tak")
        self.assertEqual(result[("n", "Długość")].new, "1 m")
        self.assertEqual(result[("n", "Producent")].new, "Polski Producent")
        # rodzenstwo 100 m dostaje wartosci wyliczone z wlasnego tytulu i typu
        self.assertEqual(result[("s", "Żyła ochronna")].new, "Tak")

    def test_existing_value_is_kept_and_conflict_reported(self) -> None:
        product = row("p", "8PYDYP215", "Przewód płaski YDYp żo 2×1,5 750V", Typ="YDYp", Ekranowany="Tak")
        changes, conflicts = decide([product], self.rules)
        self.assertNotIn(("p", "Ekranowany"), by_attr(changes))
        self.assertEqual([(c.attribute, c.old, c.new) for c in conflicts], [("Ekranowany", "Tak", "Nie")])

    def test_lgy_voltage_by_cross_section(self) -> None:
        thin = row("a", "L1", "Przewód LGY 0,75 czarny", Typ="LGY", **{"Przekrój żyły [mm]": "0,75"})
        thick = row("b", "L2", "Przewód LGY 2,5 czarny", Typ="LGY", **{"Przekrój żyły [mm]": "2,5"})
        result = by_attr(decide([thin, thick], self.rules)[0])
        self.assertEqual(result[("a", "Napięcie [V]")].new, "500")
        self.assertEqual(result[("b", "Napięcie [V]")].new, "750")

    def test_color_from_title_only_when_empty(self) -> None:
        pv = row("p", "713569/HLK", "Kabel fotowoltaiczny H1Z2Z2-K 1x6mm2 1m czerwony 713569", Typ="Fotowoltaiczne")
        black = row("c", "X", "Przewód LGY 1,5 czerwony", Typ="LGY", Kolor="Czarny")
        result = by_attr(decide([pv, black], self.rules)[0])
        self.assertEqual(result[("p", "Kolor")].new, "Czerwony")
        self.assertNotIn(("c", "Kolor"), result)

    def test_rows_without_cable_type_are_left_alone(self) -> None:
        lan = row("l", "TI0006", "Kabel U/UTP 4x2x0,5 kat. 5e 1m TI0006")
        self.assertEqual(decide([lan], self.rules), ([], []))


if __name__ == "__main__":
    unittest.main()
