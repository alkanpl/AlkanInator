from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from build_simon_subseries_woo_update import (  # noqa: E402
    build_title,
    load_rules,
    process_row,
    resolve_subseries,
)
from simon_api import flatten_category_tree, normalize_symbol  # noqa: E402

RULES = load_rules()


def api_product(symbol: str, seria: str, article_id: int = 1) -> dict:
    return {"id": str(article_id), "symbol": symbol, "params": [{"name": "Seria", "value": seria}]}


LINE_PATH = [["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon 55",
              "Ramki Simon 55 Line + Akcesoria (Wieszak, Podstawka...)", "Ramki Simon 55 Line w kolorze białym"]]
DUO_PATH = [["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon 55", "Ramki Simon 55 Duo",
             "Ramki Simon 55 Duo w kolorze czarnym"]]
MECHANISM_PATH = [["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon 54",
                   "Łączniki i przyciski IP20/IP44", "Łączniki jednobiegunowe"]]
KEY_PATHS = [["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon 55", "Łączniki elektroniczne",
              "Klawisze 1/2 do produktów elektronicznych"],
             ["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon GO - osprzęt sterowany smartfonem",
              "Sterowniki Simon 55 GO", "Klawisze 1/2 do produktów elektronicznych"]]
GASKET_PATHS = [["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon 55",
                 "Ramki Simon 55 Line + Akcesoria (Wieszak, Podstawka...)", "Uszczelki IP44 do ramek Simon 55 Line"],
                ["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon 55", "Ramki Simon 55 Duo",
                 "Uszczelki IP44 do ramek Simon 55 Duo"]]
BOX_PATH = [["KONTAKT-SIMON", "Osprzęt elektroinstalacyjny", "Simon 55", "Puszki natynkowe do ramek Line i Duo",
             "Puszki natynkowe 1-, 2- i 3-krotne do ramek Line i DUO"]]


class ResolveSubseriesTest(unittest.TestCase):
    def test_api_series_names_subseries_directly(self) -> None:
        result = resolve_subseries("Simon 55 Nature", [], RULES)
        self.assertEqual((result.series, result.subseries, result.status), ("Simon 55", "Nature", "OK"))
        result = resolve_subseries("Simon 400, \nFlash 60", [], RULES)
        self.assertEqual((result.series, result.subseries), ("Simon 400", "Flash"))

    def test_line_and_duo_come_from_category_tree(self) -> None:
        self.assertEqual(resolve_subseries("Simon 55", LINE_PATH, RULES).subseries, "Line")
        self.assertEqual(resolve_subseries("Simon 55", DUO_PATH, RULES).subseries, "Duo")

    def test_mechanisms_get_no_subseries(self) -> None:
        result = resolve_subseries("Simon 54", MECHANISM_PATH, RULES)
        self.assertEqual((result.series, result.subseries, result.status), ("Simon 54", None, "OK"))

    def test_keys_in_two_categories_do_not_become_go(self) -> None:
        result = resolve_subseries("Simon 55", KEY_PATHS, RULES)
        self.assertIsNone(result.subseries)
        self.assertEqual(result.series, "Simon 55")

    def test_multi_subseries_accessories_get_all_values_joined(self) -> None:
        for paths, expected in ((GASKET_PATHS, "Line|Duo|Nature"), (BOX_PATH, "Line|Duo")):
            result = resolve_subseries("Simon 55", paths, RULES)
            self.assertEqual(result.subseries, expected)
            self.assertEqual(result.status, "WARNING")
            self.assertFalse(result.touch_title)

    def test_simon_connect_keeps_series_and_title(self) -> None:
        result = resolve_subseries("K45", [], RULES)
        self.assertEqual((result.series, result.subseries), ("Simon Connect", "K45"))
        self.assertFalse(result.series_full)
        self.assertFalse(result.touch_title)

    def test_unknown_api_series_does_not_override_shop_series(self) -> None:
        result = resolve_subseries("Akcesoria do wielu serii", [], RULES)
        self.assertIsNone(result.series)


class BuildTitleTest(unittest.TestCase):
    def test_replaces_series_mention_and_strips_uppercase_token(self) -> None:
        title, note = build_title("Ramka 1-krotna LINE biały mat Simon 55 TR1/111", "Simon 55", "Line", "TR1/111/KON", RULES)
        self.assertEqual(title, "Ramka 1-krotna biały mat Simon 55 Line TR1/111")
        self.assertEqual(note, "")
        title, _ = build_title("Ramka 2-krotna DUO: Czarna, podstawa złota Simon 55 TRD2.166/149", "Simon 55", "Duo", "TRD2.166/149/KON", RULES)
        self.assertEqual(title, "Ramka 2-krotna: Czarna, podstawa złota Simon 55 Duo TRD2.166/149")

    def test_inserts_before_trailing_code_when_series_missing(self) -> None:
        title, note = build_title("Sterownik przyciskowy, WiFi, czarny mat TEW1W.01/149", "Simon 55", "GO", "TEW1W.01/149/KON", RULES)
        self.assertEqual(title, "Sterownik przyciskowy, WiFi, czarny mat Simon 55 GO TEW1W.01/149")
        self.assertEqual(note, "")

    def test_simon_go_mention_becomes_simon_55_go(self) -> None:
        title, _ = build_title("Kontroler Simon GO, Wi-Fi, czarny mat TEK1W.01/149", "Simon 55", "GO", "TEK1W.01/149/KON", RULES)
        self.assertEqual(title, "Kontroler Simon 55 GO, Wi-Fi, czarny mat TEK1W.01/149")

    def test_existing_full_name_is_left_alone(self) -> None:
        original = "Ramka pojedyncza biała DR1/11 Simon 54 Premium"
        self.assertEqual(build_title(original, "Simon 54", "Premium", "DR1/11/KON", RULES)[0], original)

    def test_series_correction_drops_old_subseries_word(self) -> None:
        title, _ = build_title("Łącznik świecznikowy kremowy DW5.01/41 Simon 54 Premium", "Simon 54", None, "DW5.01/41/KON", RULES)
        self.assertEqual(title, "Łącznik świecznikowy kremowy DW5.01/41 Simon 54")
        title, _ = build_title("Wyłącznik pojedynczy SIMON 54 kremowy DW1.01/41", "Simon 54", None, "DW1.01/41/KON", RULES)
        self.assertEqual(title, "Wyłącznik pojedynczy SIMON 54 kremowy DW1.01/41")

    def test_appends_with_warning_when_code_not_at_end(self) -> None:
        title, note = build_title("Ramka biała Line edycja specjalna", "Simon 55", "Line", "TRX/111/KON", RULES)
        self.assertEqual(title, "Ramka biała Line edycja specjalna Simon 55 Line")
        self.assertIn("na koncu", note)


class ProcessRowTest(unittest.TestCase):
    def row(self, sku: str, title: str, series: str | None) -> dict:
        return {"id": "1", "SKU": sku, "Title": title, "Atrybut Produktu: Seria": series}

    def test_frame_gets_full_series_subseries_and_title(self) -> None:
        result = process_row(self.row("TR1/111/KON", "Ramka 1-krotna LINE biały mat Simon 55 TR1/111", "Simon 55"),
                             api_product("TR1/111", "Simon 55"), LINE_PATH, RULES)
        self.assertEqual(result.series_after, "Simon 55 Line")
        self.assertEqual(result.subseries, "Line")
        self.assertEqual(result.title_after, "Ramka 1-krotna biały mat Simon 55 Line TR1/111")
        self.assertEqual(result.status, "OK")
        self.assertEqual(result.changes, ["Seria", "Podseria", "Title"])

    def test_mechanism_with_premium_in_woo_is_corrected_with_warning(self) -> None:
        result = process_row(self.row("DW5.01/41/KON", "Łącznik świecznikowy kremowy DW5.01/41 Simon 54 Premium", "Simon 54 Premium"),
                             api_product("DW5.01/41", "Simon 54"), MECHANISM_PATH, RULES)
        self.assertEqual(result.series_after, "Simon 54")
        self.assertEqual(result.subseries, "")
        self.assertEqual(result.title_after, "Łącznik świecznikowy kremowy DW5.01/41 Simon 54")
        self.assertEqual(result.status, "WARNING")

    def test_keep_woo_series_only_reports(self) -> None:
        result = process_row(self.row("DW5.01/41/KON", "Łącznik świecznikowy kremowy DW5.01/41 Simon 54 Premium", "Simon 54 Premium"),
                             api_product("DW5.01/41", "Simon 54"), MECHANISM_PATH, RULES, keep_woo_series=True)
        self.assertEqual(result.series_after, "Simon 54 Premium")
        self.assertEqual(result.title_after, "Łącznik świecznikowy kremowy DW5.01/41 Simon 54 Premium")
        self.assertEqual(result.status, "WARNING")

    def test_multi_subseries_box_gets_joined_attributes_and_untouched_title(self) -> None:
        title = "Puszka natynkowa 3-krotna do ramek LINE/DUO: Czarny mat Simon 55 TPN3/149"
        result = process_row(self.row("TPN3/149/KON", title, "Simon 55"),
                             api_product("TPN3/149", "Simon 55"), BOX_PATH, RULES)
        self.assertEqual(result.subseries, "Line|Duo")
        self.assertEqual(result.series_after, "Simon 55 Line|Simon 55 Duo")
        self.assertEqual(result.title_after, title)
        self.assertEqual(result.status, "WARNING")
        self.assertEqual(result.changes, ["Seria", "Podseria"])

    def test_missing_product_is_untouched(self) -> None:
        result = process_row(self.row("CGZ1.01/11/CR1/11", "Gniazdo z ramką SIMON 10", "Simon 10"),
                             {"symbol": "CGZ1.01/11/CR1/11", "_missing": True}, [], RULES)
        self.assertEqual(result.status, "DO_SPRAWDZENIA")
        self.assertEqual(result.changes, [])
        self.assertEqual(result.series_after, "Simon 10")

    def test_connect_gets_subseries_without_touching_series_or_title(self) -> None:
        result = process_row(self.row("K02/9/KON", "Gniazdo K45 Simon Connect K02/9", "Simon Connect"),
                             api_product("K02/9", "K45"), [], RULES)
        self.assertEqual(result.subseries, "K45")
        self.assertEqual(result.series_after, "Simon Connect")
        self.assertEqual(result.title_after, "Gniazdo K45 Simon Connect K02/9")


class CategoryTreeTest(unittest.TestCase):
    def test_flatten_handles_dict_and_list_children(self) -> None:
        tree = {"id": 1, "name": "ROOT", "subCategories": {"1": {"id": 2, "name": "A", "articleIdList": [10],
                "subCategories": [{"id": 3, "name": "B", "articleIdList": [10, "11"]}]}}}
        paths = flatten_category_tree(tree)
        self.assertEqual(paths[10], [["ROOT", "A"], ["ROOT", "A", "B"]])
        self.assertEqual(paths[11], [["ROOT", "A", "B"]])

    def test_normalize_symbol_strips_shop_suffix(self) -> None:
        self.assertEqual(normalize_symbol("tr1/111/KON "), "TR1/111")
        self.assertEqual(normalize_symbol("CGZ1.01/11/CR1/11"), "CGZ1.01/11/CR1/11")


if __name__ == "__main__":
    unittest.main()
