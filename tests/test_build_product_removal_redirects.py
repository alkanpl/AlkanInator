from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from build_product_removal_redirects import (  # noqa: E402
    CHECK,
    GROUP_NAME,
    OK,
    WARNING,
    Category,
    Redirect,
    categories_from_records,
    match_category_paths,
    parse_product_categories,
    plan_removal,
    write_redirects_csv,
    write_redirects_json,
)

PLAFONY = "Oświetlenie > Oświetlenie wewnętrzne > Plafony LED"
PANELE = "Oświetlenie > Oświetlenie wewnętrzne > Panele LED"
HEADER = ["id", "Title", "Sku", "Status", "Product Type", "Permalink", "Product categories",
          "Atrybut produktu: Moc [W]", "Atrybut produktu: Kolor", "Atrybut produktu: Symbol", "Do usunięcia", "Zamiennik"]


def cats(panele_count: int | None = 5) -> dict[int, Category]:
    return {
        1: Category(1, 0, "Oświetlenie", "oswietlenie", 50),
        2: Category(2, 1, "Oświetlenie wewnętrzne", "oswietlenie-wewnetrzne", 30),
        3: Category(3, 2, "Plafony LED", "plafony-led", 10),
        4: Category(4, 2, "Panele LED", "panele-led", panele_count),
    }


def row(row_id: str, sku: str, slug: str, category: str = PLAFONY, delete: str = "", replacement: str = "",
        power: str = "12", color: str = "Biały", status: str = "publish", kind: str = "simple", title: str = "") -> dict:
    return {
        "id": row_id, "Title": title or f"Plafon LED {power}W {color} {sku}", "Sku": sku, "Status": status,
        "Product Type": kind, "Permalink": f"https://alkan.pl/sklep/{slug}/", "Product categories": category,
        "Atrybut produktu: Moc [W]": power, "Atrybut produktu: Kolor": color, "Atrybut produktu: Symbol": sku,
        "Do usunięcia": delete, "Zamiennik": replacement,
    }


def redirect(rule_id: int, source: str, target: str) -> Redirect:
    target_path = target.replace("https://alkan.pl", "").rstrip("/") if target.startswith("https://alkan.pl") else None
    return Redirect(rule_id, source, source.rstrip("/"), target, target_path, 1)


def plan(rows: list[dict], redirects: list[Redirect] | None = None, categories: dict | None = None):
    return plan_removal(HEADER, rows, categories or cats(), redirects or [])


class ReplacementTest(unittest.TestCase):
    def test_replacement_by_sku_redirects_to_product(self) -> None:
        result = plan([row("10", "P-A", "plafon-a", delete="tak", replacement="P-B"), row("11", "P-B", "plafon-b")])
        d = result.decisions[0]
        self.assertEqual((d.source_path, d.target_path, d.target_kind), ("/sklep/plafon-a", "/sklep/plafon-b", "zamiennik"))
        self.assertEqual(d.status, OK)
        self.assertTrue(d.add_redirect and d.set_draft)
        self.assertEqual(result.suggestions, [])

    def test_replacement_marked_for_deletion_passes_its_target_on(self) -> None:
        result = plan([
            row("10", "P-A", "plafon-a", delete="x", replacement="P-B"),
            row("11", "P-B", "plafon-b", delete="x", replacement="P-C"),
            row("12", "P-C", "plafon-c"),
        ])
        self.assertEqual([d.target_path for d in result.decisions], ["/sklep/plafon-c", "/sklep/plafon-c"])

    def test_ambiguous_sku_blocks_product(self) -> None:
        result = plan([
            row("10", "P-A", "plafon-a", delete="tak", replacement="P-B"),
            row("11", "P-B", "plafon-b"),
            row("12", "P-B", "plafon-b-2"),
        ])
        d = result.decisions[0]
        self.assertEqual(d.status, CHECK)
        self.assertFalse(d.add_redirect or d.set_draft)
        self.assertIn("niejednoznaczne", d.notes[0])

    def test_unpublished_replacement_blocks_product(self) -> None:
        result = plan([row("10", "P-A", "plafon-a", delete="tak", replacement="11"), row("11", "P-B", "plafon-b", status="draft")])
        self.assertEqual(result.decisions[0].status, CHECK)

    def test_replacement_loop_is_detected(self) -> None:
        result = plan([
            row("10", "P-A", "plafon-a", delete="tak", replacement="P-B"),
            row("11", "P-B", "plafon-b", delete="tak", replacement="P-A"),
        ])
        self.assertTrue(all(d.status == CHECK for d in result.decisions))


class CategoryTest(unittest.TestCase):
    def test_no_replacement_uses_deepest_category_and_suggests_similar(self) -> None:
        result = plan([
            row("10", "P-A", "plafon-a", delete="tak"),
            row("11", "P-B", "plafon-b"),
            row("12", "P-C", "plafon-c", color="Czarny"),
            row("13", "P-E", "plafon-e", power="18", color="Czarny"),
            row("14", "P-D", "panel-d", category=PANELE),
        ])
        d = result.decisions[0]
        self.assertEqual(d.target_path, "/kategoria-produktu/oswietlenie/oswietlenie-wewnetrzne/plafony-led")
        self.assertEqual(d.target_kind, "kategoria")
        # P-E rozni sie wszystkimi parametrami, P-D jest w innej kategorii - nie sa propozycjami.
        self.assertEqual([s.sku for s in result.suggestions], ["P-B", "P-C"])
        self.assertGreater(result.suggestions[0].score, result.suggestions[1].score)
        self.assertEqual(result.suggestions[1].differences, "Kolor: Biały -> Czarny")

    def test_category_left_empty_falls_back_to_parent(self) -> None:
        result = plan([row("20", "PN-1", "panel-1", category=PANELE, delete="tak"), row("11", "P-B", "plafon-b")],
                      categories=cats(panele_count=1))
        d = result.decisions[0]
        self.assertEqual(d.target_path, "/kategoria-produktu/oswietlenie/oswietlenie-wewnetrzne")
        self.assertEqual(d.status, WARNING)

    def test_unknown_category_without_count_only_warns(self) -> None:
        result = plan([row("20", "PN-1", "panel-1", category=PANELE, delete="tak")], categories=cats(panele_count=None))
        d = result.decisions[0]
        self.assertEqual(d.target_path, "/kategoria-produktu/oswietlenie/oswietlenie-wewnetrzne/panele-led")
        self.assertEqual(d.status, WARNING)


class ExistingRedirectsTest(unittest.TestCase):
    def test_target_follows_existing_redirect_chain(self) -> None:
        rules = [
            redirect(5, "/kategoria-produktu/oswietlenie/oswietlenie-wewnetrzne/plafony-led/", "https://alkan.pl/kategoria-produktu/plafony/"),
            redirect(6, "/kategoria-produktu/plafony/", "https://alkan.pl/kategoria-produktu/oswietlenie/"),
        ]
        result = plan([row("10", "P-A", "plafon-a", delete="tak"), row("11", "P-B", "plafon-b")], rules)
        self.assertEqual(result.decisions[0].target_path, "/kategoria-produktu/oswietlenie")
        self.assertEqual(len(result.chains), 1)

    def test_source_already_redirected_is_not_duplicated(self) -> None:
        rules = [redirect(7, "/sklep/plafon-a/", "https://alkan.pl/sklep/inny/")]
        result = plan([row("10", "P-A", "plafon-a", delete="tak", replacement="P-B"), row("11", "P-B", "plafon-b")], rules)
        d = result.decisions[0]
        self.assertEqual(d.status, WARNING)
        self.assertFalse(d.add_redirect)
        self.assertTrue(d.set_draft)

    def test_existing_rules_pointing_to_deleted_product_are_listed(self) -> None:
        rules = [redirect(8, "/sklep/stary-plafon/", "https://alkan.pl/sklep/plafon-a/")]
        result = plan([row("10", "P-A", "plafon-a", delete="tak", replacement="P-B"), row("11", "P-B", "plafon-b")], rules)
        self.assertEqual([(r.id, d.target_path) for r, d in result.updates], [(8, "/sklep/plafon-b")])
        self.assertEqual(result.decisions[0].status, WARNING)


class InputTest(unittest.TestCase):
    def test_variation_is_blocked(self) -> None:
        result = plan([row("30", "V-1", "wariant", delete="tak", kind="variation")])
        self.assertEqual(result.decisions[0].status, CHECK)
        self.assertFalse(result.decisions[0].set_draft)

    def test_missing_delete_column_stops(self) -> None:
        with self.assertRaises(SystemExit):
            plan_removal(["id", "Title"], [{"id": "1", "Title": "x"}], cats(), [])

    def test_category_paths_and_name_matching(self) -> None:
        paths = parse_product_categories("Oświetlenie &gt; Oświetlenie wewnętrzne > Plafony LED|Panele LED")
        self.assertEqual(paths[1], ["Panele LED"])
        self.assertEqual(match_category_paths(paths, cats()), ([3, 4], []))

    def test_categories_from_export_with_parent_slug(self) -> None:
        records = [
            {"Term ID": "1", "Term Name": "Oświetlenie", "Term Slug": "oswietlenie", "Parent Slug": ""},
            {"Term ID": "2", "Term Name": "Plafony &amp; lampy", "Term Slug": "plafony", "Parent Slug": "oswietlenie"},
        ]
        result = categories_from_records(records)
        self.assertEqual((result[2].parent_id, result[2].name), (1, "Plafony & lampy"))


class OutputTest(unittest.TestCase):
    def test_csv_and_json_for_redirection(self) -> None:
        result = plan([row("10", "P-A", "plafon-a", delete="tak", replacement="P-B"), row("11", "P-B", "plafon-b")])
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "r.csv"
            self.assertEqual(write_redirects_csv(csv_path, result.decisions, "https://alkan.pl"), 1)
            raw = csv_path.read_bytes()
            self.assertTrue(raw.startswith(b"source,target,regex,code"))
            self.assertIn(b"/sklep/plafon-a/,https://alkan.pl/sklep/plafon-b/,0,301", raw)
            json_path = Path(tmp) / "r.json"
            write_redirects_json(json_path, result.decisions, "https://alkan.pl")
            data = json.loads(json_path.read_text(encoding="utf-8"))
            self.assertEqual(data["groups"][0]["name"], GROUP_NAME)
            self.assertEqual(data["redirects"][0]["match_url"], "/sklep/plafon-a")


if __name__ == "__main__":
    unittest.main()
