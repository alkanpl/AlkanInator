import sys
import unittest
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from build_simon_woo_relations import (  # noqa: E402
    Product,
    build_relations,
    group_tag,
    image_file_name,
    load_variant_colors,
    symbol_parts,
    variant_groups,
)

RULES = yaml.safe_load((ROOT / "dictionaries" / "simon_woo_relations.yaml").read_text(encoding="utf-8"))


def product(index: int, sku: str, title: str) -> Product:
    base, color = symbol_parts(sku)
    return Product(row=index + 2, id=str(1000 + index), sku=sku, title=title, base=base, color=color)


def sample() -> list[Product]:
    titles = [
        ("NW1.01/11/KON", "Łącznik jednobiegunowy 10 AX Push In biały Simon 24 NW1.01/11"),
        ("NW1.01/49/KON", "Łącznik jednobiegunowy 10 AX Push In czarny mat Simon 24 NW1.01/49"),
        ("NW1.01/42/KON", "Łącznik jednobiegunowy 10 AX Push In kaszmirowy Simon 24 NW1.01/42"),
        ("NGZ1CZ.01/11/KON", "Gniazdo wtyczkowe pojedyncze z uziemieniem i przesłonami 16 A Push In białe Simon 24 NGZ1CZ.01/11"),
        ("NGZ1BCZ.01/11/KON", "Gniazdo wtyczkowe pojedyncze IP44 z uszczelką i przesłonami 16 A Push In białe Simon 24 NGZ1BCZ.01/11"),
        ("NGZ1BCZ.01/49A/KON", "Gniazdo wtyczkowe pojedyncze IP44 z uszczelką i przesłonami z klapką transparentną 16 A Push In czarny mat Simon 24 NGZ1BCZ.01/49A"),
        ("NR1/11/KON", "Ramka 1-krotna biała Simon 24 NR1/11"),
        ("NR1/49/KON", "Ramka 1-krotna czarny mat Simon 24 NR1/49"),
        ("NR2/11/KON", "Ramka 2-krotna biała Simon 24 NR2/11"),
        ("NR2/49/KON", "Ramka 2-krotna czarny mat Simon 24 NR2/49"),
        ("NUV2/KON", "Zestaw uszczelek IP44 do ramki 2-krotnej w układzie pionowym Simon 24 NUV2"),
        ("NSC/11/KON", "Puszka natynkowa pojedyncza składana 40 mm biała Simon 24 NSC/11"),
        ("NEZ1W.01/49/KON", "Sterownik rolet, żaluzji i markiz 2x5 A Simon GO Wi-Fi czarny mat Simon 24 NEZ1W.01/49"),
    ]
    products = [product(i, sku, title) for i, (sku, title) in enumerate(titles)]
    build_relations(products, RULES)
    return products


class SymbolTest(unittest.TestCase):
    def test_symbol_parts(self):
        self.assertEqual(symbol_parts("NW6/2.01/X/11/KON"), ("NW6/2.01/X", "11"))
        self.assertEqual(symbol_parts("NGZ1BCZ.01/49A/KON"), ("NGZ1BCZ.01", "49A"))
        self.assertEqual(symbol_parts("NUV2/KON"), ("NUV2", ""))

    def test_group_tag_strips_color_current_push_in_and_flap(self):
        self.assertEqual(
            group_tag("Gniazdo wtyczkowe pojedyncze IP44 z uszczelką i przesłonami z klapką transparentną 16 A Push In czarny mat Simon 24 NGZ1BCZ.01/49A", RULES),
            "Gniazdo wtyczkowe pojedyncze IP44 z uszczelką i przesłonami 24 simon",
        )
        self.assertEqual(
            group_tag("Sterownik rolet, żaluzji i markiz 2x5 A Simon GO Wi-Fi czarny mat Simon 24 NEZ1W.01/49", RULES),
            "Sterownik rolet, żaluzji i markiz Simon GO Wi-Fi 24 simon",
        )

    def test_image_file_name_keeps_code(self):
        p = product(0, "NEW1W.01/49/KON", "Sterownik oświetleniowy 16 A z pomiarem energii Simon GO Wi-Fi czarny mat Simon 24 NEW1W.01/49")
        name = image_file_name(p, 0)
        self.assertTrue(name.endswith("-new1w-01-49.jpg"))
        self.assertNotIn("--", name)


class RelationsTest(unittest.TestCase):
    def setUp(self):
        self.by_sku = {p.sku: p for p in sample()}

    def test_colors_share_one_tag(self):
        tags = {self.by_sku[s].tag for s in ("NW1.01/11/KON", "NW1.01/49/KON", "NW1.01/42/KON")}
        self.assertEqual(tags, {"Łącznik jednobiegunowy 24 simon"})
        self.assertEqual(self.by_sku["NGZ1BCZ.01/11/KON"].tag, self.by_sku["NGZ1BCZ.01/49A/KON"].tag)

    def test_module_up_sells_are_frames_of_same_color(self):
        self.assertEqual(self.by_sku["NW1.01/49/KON"].up_sells, ["NR1/49/KON", "NR2/49/KON"])
        # Klapka transparentna (49A) = kolor 49.
        self.assertEqual(self.by_sku["NGZ1BCZ.01/49A/KON"].up_sells, ["NR1/49/KON", "NR2/49/KON"])

    def test_missing_frame_color_is_reported(self):
        cashmere = self.by_sku["NW1.01/42/KON"]
        self.assertEqual(cashmere.up_sells, [])
        self.assertTrue(cashmere.notes)

    def test_cross_sells_skip_own_group_and_up_sells(self):
        switch = self.by_sku["NW1.01/11/KON"]
        self.assertIn("NGZ1CZ.01/11/KON", switch.cross_sells)
        self.assertNotIn("NW1.01/49/KON", switch.cross_sells)
        self.assertFalse(set(switch.cross_sells) & set(switch.up_sells))

    def test_compatible_is_mirrored_and_any_color_for_seals(self):
        self.assertEqual(self.by_sku["NUV2/KON"].compatible, ["NR2/11/KON", "NR2/49/KON"])
        self.assertIn("NUV2/KON", self.by_sku["NR2/49/KON"].compatible)
        self.assertIn("NSC/11/KON", self.by_sku["NR1/11/KON"].compatible)
        self.assertNotIn("NSC/11/KON", self.by_sku["NR1/49/KON"].compatible)


class VariantGroupsTest(unittest.TestCase):
    def setUp(self):
        products = sample()
        rows = variant_groups(products, RULES, load_variant_colors(RULES))
        self.by_tag = {row[0]: row for row in rows}

    def test_group_names_wpc_group_and_attribute(self):
        row = self.by_tag["Łącznik jednobiegunowy 24 simon"]
        self.assertEqual(row[3:5], ["Kolor - Simon", "Kolor"])
        self.assertEqual(row[5], "NW1.01/11/KON: Biały | NW1.01/49/KON: Czarny | NW1.01/42/KON: Kaszmir")
        self.assertEqual(row[6], "OK")

    def test_flap_version_uses_plain_color(self):
        row = next(r for r in self.by_tag.values() if "NGZ1BCZ.01/49A/KON" in r[8])
        self.assertIn("NGZ1BCZ.01/49A/KON: Czarny", row[5])
        self.assertEqual(row[6], "OK")

    def test_product_without_color_has_no_wpc_group(self):
        row = next(r for r in self.by_tag.values() if r[8] == "NUV2/KON")
        self.assertEqual(row[3:5], ["-", "-"])

    def test_repeated_value_needs_extra_axis(self):
        products = [
            product(0, "NW1.01/49/KON", "Łącznik jednobiegunowy 10 AX czarny mat Simon 24 NW1.01/49"),
            product(1, "NW1.01/49A/KON", "Łącznik jednobiegunowy 10 AX czarny mat Simon 24 NW1.01/49A"),
        ]
        build_relations(products, RULES)
        row = variant_groups(products, RULES, load_variant_colors(RULES))[0]
        self.assertEqual(row[6], "DO_SPRAWDZENIA")
        self.assertIn("dodatkowej osi", row[7])


if __name__ == "__main__":
    unittest.main()
