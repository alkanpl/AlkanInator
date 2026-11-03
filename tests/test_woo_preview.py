from __future__ import annotations

import http.client
import json
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from unittest import TestCase

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from woo_api.client import WooBridgeError
from woo_api.workbook import read_products_xlsx, write_products_xlsx
from woo_preview.cards import (
    Catalog,
    PreviewContext,
    autop,
    build_views,
    change_html,
    changes_signature,
    load_variant_groups,
    related_html,
    render_card,
    tag_slug,
    variants_html,
)
from woo_preview.server import STALE_PREVIEW, App, PreviewServer, list_profiles, safe_file_name


COLUMNS = ["id", "Title", "SKU", "Content", "Image URL", "Atrybut Produktu: Kolor", "Post Modified Date"]


def row(product_id: int, sku: str, title: str, color: str = "Biały") -> dict[str, object]:
    return {"id": product_id, "Title": title, "SKU": sku, "Content": "<p>Opis</p>", "Image URL": "https://x/a.jpg", "Atrybut Produktu: Kolor": color}


class FakeShop:
    """Sklep, w ktorym produkty maja tytul "Gniazdo" i kolor "Biały"; wszystko inne to zmiana."""

    def __init__(self) -> None:
        self.titles: dict[int, str] = {}
        self.calls: list[dict[str, object]] = []
        self.reverts: list[tuple[str, bool]] = []
        self.groups: list[dict[str, object]] | None = None  # None = wtyczka sprzed 2.1 (404)
        self.tagged: dict[str, list[dict[str, str]]] = {}

    def shipping_preview(self, products):
        self.calls.append({"shipping": products})
        if self.groups is None:
            raise WooBridgeError("HTTP 404: rest_no_route", status_code=404)
        standard = [{"title": "Paczkomat InPost", "price": 16.99, "cod_price": None, "image_url": "", "border_col": "", "is_local_pickup": False},
                    {"title": "Kurier DPD", "price": 19.99, "cod_price": 24.99, "image_url": "", "border_col": "", "is_local_pickup": False}]
        out = []
        for item in products:
            before = {"qty": 1, "shipping_class": "Standard", "weight": "0.05", "methods": standard}
            after = None
            if item["values"]:
                after = {"qty": 1, "shipping_class": item["values"].get("Shipping Class", "Standard"), "weight": "3.6", "methods": standard[1:]}
            out.append({"id": item["id"], "before": before, "after": after, "warnings": []})
        return out

    def variant_groups(self):
        if self.groups is None:
            raise WooBridgeError("HTTP 404: rest_no_route", status_code=404)
        return self.groups

    def __enter__(self) -> "FakeShop":
        return self

    def __exit__(self, *_: object) -> None:
        pass

    def sync(self, rows, *, dry_run=True, allow_new_terms=False, allow_destructive=False, columns=None, changeset=""):
        self.calls.append({"rows": [dict(r) for r in rows], "dry_run": dry_run, "columns": columns})
        products = []
        for item in rows:
            product_id = int(item["id"])
            changes = []
            for column, current in (("Title", self.titles.get(product_id, "Gniazdo")), ("Atrybut Produktu: Kolor", "Biały"), ("Shipping Class", "Standard")):
                if column in item and (not columns or column in columns) and item[column] != current:
                    changes.append({"column": column, "old": current, "new": item[column]})
            errors = ["Shipping Class: nie ma klasy"] if item.get("SKU") == "ZLE" else []
            status = "error" if errors else ("changed" if changes else "unchanged")
            plan_hash = "hash-" + json.dumps(changes, sort_keys=True)
            if status == "changed" and not dry_run:
                status = "applied" if item.get("_plan_hash") == plan_hash else "error"
            products.append({"row": item["_row"], "id": product_id, "sku": item["SKU"], "title": item.get("Title", ""), "status": status, "plan_hash": plan_hash, "changes": changes, "errors": errors, "warnings": []})
        summary = {
            "rows": len(products),
            "changed": sum(p["status"] in ("changed", "applied") for p in products),
            "applied": sum(p["status"] == "applied" for p in products),
            "errors": sum(bool(p["errors"]) for p in products),
            "warnings": 0,
            "changes": sum(len(p["changes"]) for p in products),
        }
        return {"changeset": changeset, "summary": summary, "products": products, "new_terms": [], "written_columns": ["Title"], "read_only_columns": [], "ignored_columns": [], "max_products": 500, "writes_enabled": True}

    def iter_product_pages(self, per_page=200, **filters):
        self.calls.append({"pull": filters})
        rows = [{"id": "90", "SKU": sku, "Title": f"Powiazany {sku}"} for sku in filters.get("sku", [])]
        rows += [item for tag in filters.get("tag", []) for item in self.tagged.get(tag, [])]
        yield {"rows": rows, "total_pages": 1}

    def revert(self, changeset, *, dry_run=True):
        self.reverts.append((changeset, dry_run))
        return {"changeset": changeset, "summary": {"changed": 1, "applied": 0 if dry_run else 1, "errors": 0}, "products": []}


class CardsTests(TestCase):
    def dry(self) -> dict[str, object]:
        return {
            "products": [
                {"row": 2, "id": 1, "sku": "A", "status": "unchanged", "changes": [], "errors": [], "warnings": []},
                {"row": 3, "id": 2, "sku": "B", "status": "changed", "changes": [{"column": "Title", "old": "Stary", "new": "Nowy <b>"}], "errors": [], "warnings": []},
                {"row": 4, "id": 3, "sku": "C", "status": "error", "changes": [], "errors": ["zle"], "warnings": []},
            ]
        }

    def rows(self) -> list[dict[str, object]]:
        return [
            {"_row": 2, "id": "1", "SKU": "A", "Title": "Bez zmian"},
            {"_row": 3, "id": "2", "SKU": "B", "Title": "Nowy <b>", "Atrybut Produktu: Producent": "Ospel"},
            {"_row": 4, "id": "3", "SKU": "C", "Title": "Blad"},
        ]

    def test_views_skip_unchanged_put_errors_first_and_restore_old_values(self) -> None:
        views = build_views(self.rows(), self.dry())
        self.assertEqual([view.sku for view in views], ["C", "B"])
        changed = views[1]
        self.assertEqual(changed.before["Title"], "Stary")
        self.assertEqual(changed.after["Title"], "Nowy <b>")
        self.assertEqual(changed.before["Atrybut Produktu: Producent"], "Ospel")

    def test_card_escapes_values_and_marks_changed_parts(self) -> None:
        view = build_views(self.rows(), self.dry())[1]
        after = render_card(view, "after", theme_base="/theme")
        self.assertIn("Nowy &lt;b&gt;", after)
        self.assertIn('page-name product-title ai-changed', after)
        self.assertIn('href="/theme/style.css"', after)
        before = render_card(view, "before", theme_base=None)
        self.assertIn(">Stary</h1>", before)
        self.assertNotIn("/theme/", before)

    def test_signature_lists_only_products_to_write(self) -> None:
        self.assertEqual(changes_signature(self.dry()), frozenset({("id:2", "Title", "Stary", "Nowy <b>")}))

    def test_change_html_lists_and_cleared_fields(self) -> None:
        html = change_html("Product Tags", "A|B", "B|C")
        self.assertIn("<ins>+ C</ins>", html)
        self.assertIn("<del>− A</del>", html)
        self.assertIn("kolejność", change_html("Up-Sells", "A|B", "B|A"))
        self.assertIn("wyczyszczone", change_html("Content", "<p>x</p>", ""))
        self.assertIn("<ins>nowy</ins>", change_html("Title", "Gniazdo stare", "Gniazdo nowy"))

    def test_autop_wraps_plain_text_but_keeps_blocks(self) -> None:
        self.assertEqual(autop("Linia 1\nLinia 2\n\n<h3>X</h3>"), "<p>Linia 1<br>\nLinia 2</p>\n<h3>X</h3>")


class RelationsTests(TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.groups_path = Path(self.tmp.name) / "groups.tsv"
        self.groups_path.write_text(
            "Kolor producenta - Ospel\tpublish\tproduct_tag\tgniazdo-antenowe-ospel-impresja,inny-tag\t5\tKolor producenta\timg:Kolor producenta\tsw:\tdd:\n"
            "Stara grupa\tdraft\tproduct_tag\tgniazdo-antenowe-ospel-impresja\t0\tKolor\timg:\tsw:\tdd:\n",
            encoding="utf-8",
        )

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def product(self, sku: str, color: str, tags: str = "Gniazdo antenowe Ospel Impresja") -> dict[str, str]:
        return {"SKU": sku, "id": sku[-1], "Title": f"Gniazdo {color}", "Product Tags": tags, "Atrybut Produktu: Kolor Producenta": color, "Image URL": f"https://x/{sku}.jpg"}

    def test_tag_slug_like_wordpress(self) -> None:
        self.assertEqual(tag_slug("Łącznik świecznikowy 55 simon"), "lacznik-swiecznikowy-55-simon")
        self.assertEqual(tag_slug("Gniazdo antenowe typu F pojedyncze Ospel Impresja"), "gniazdo-antenowe-typu-f-pojedyncze-ospel-impresja")

    def test_variant_switcher_shows_axis_and_flags_duplicates(self) -> None:
        groups = load_variant_groups(self.groups_path)
        self.assertEqual([g.name for g in groups["gniazdo-antenowe-ospel-impresja"]], ["Kolor producenta - Ospel"])
        own = self.product("GPA-1", "Biały")
        catalog = Catalog([own, self.product("GPA-2", "Srebro"), self.product("GPA-3", "Biały")])
        html = variants_html(own, catalog, PreviewContext(catalog, catalog, groups), set())
        self.assertIn("Kolor producenta - Ospel", html)
        self.assertIn('wpclv-term-image wpclv-term-image--with-label wpclv-term active', html)
        self.assertIn("Srebro", html)
        self.assertIn("DO_SPRAWDZENIA: ta sama wartość osi co GPA-3", html)

    def test_tag_without_group_is_reported(self) -> None:
        own = self.product("A-1", "Biały", tags="Bez grupy")
        catalog = Catalog([own, self.product("A-2", "Czarny", tags="Bez grupy")])
        html = variants_html(own, catalog, PreviewContext(catalog, catalog, load_variant_groups(self.groups_path)), set())
        self.assertIn("nie należy do żadnej grupy wariantów", html)

    def test_related_sections_mark_added_removed_and_unknown(self) -> None:
        catalog = Catalog([self.product("R-1", "Biały"), self.product("R-2", "Czarny")])
        after = {"Up-Sells": "R-1|ZZZ", "_alkan_compatible_product_skus": "R-2"}
        before = {"Up-Sells": "R-1|R-2", "_alkan_compatible_product_skus": "R-2"}
        html = related_html(after, before, catalog, {"Up-Sells"})
        self.assertIn("Produkty kompatybilne", html)
        self.assertIn("Polecamy również", html)
        self.assertIn("suggestion-card ai-added ai-missing", html)  # ZZZ: dodany, nieznany
        self.assertIn("suggestion-card ai-removed", html)  # R-2 usuniety z Up-Sells
        self.assertNotIn("Cross-Sells", html)


class AppFlowTests(TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.config = self.root / "config.json"
        self.config.write_text(
            json.dumps({
                "default": "local",
                "profiles": {
                    "local": {"url": "https://alkanlocal.test:8443", "token": "SEKRET-LOCAL", "verify_ssl": False},
                    "prod": {"url": "https://alkan.pl", "token": "SEKRET-PROD", "production": True},
                },
            }),
            encoding="utf-8",
        )
        self.shop = FakeShop()
        self.questions: list[str] = []
        self.answer = True

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def make_app(self) -> App:
        def factory(profile, log):
            return self.shop, {"url": "https://alkan.pl" if profile == "prod" else "https://alkanlocal.test:8443", "production": profile == "prod"}

        def confirm(text, title):
            self.questions.append(text)
            return self.answer

        return App(
            theme_dir=None,
            config_path=self.config,
            work_dir=self.root / "work",
            report_dir=self.root / "reports",
            client_factory=factory,
            confirm_production=confirm,
            now=lambda: datetime(2026, 11, 2, 12, 0, 0),
        )

    def xlsx(self, rows: list[dict[str, object]]) -> bytes:
        path = self.root / "plik.xlsx"
        write_products_xlsx(path, COLUMNS, rows)
        return path.read_bytes()

    def preview(self, app: App, profile: str, rows: list[dict[str, object]]) -> None:
        app.start_preview(profile, "Mój plik.xlsx", self.xlsx(rows))
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)

    def test_profiles_never_expose_tokens(self) -> None:
        profiles = list_profiles(self.config)
        self.assertEqual([(p["name"], p["production"]) for p in profiles], [("local", False), ("prod", True)])
        self.assertNotIn("SEKRET", json.dumps(profiles))
        self.assertNotIn("SEKRET", json.dumps(self.make_app().state()))

    def test_local_apply_sends_only_changed_products_and_can_revert(self) -> None:
        app = self.make_app()
        self.preview(app, "local", [row(1, "A", "Gniazdo"), row(2, "B", "Gniazdo nowe")])
        self.assertEqual(app.session.to_apply, 1)
        self.assertEqual(app.session.file_name, "Mój plik.xlsx")
        self.shop.calls.clear()

        app.start_apply()
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)
        self.assertEqual([len(call["rows"]) for call in self.shop.calls], [1, 1])  # dry-run i zapis tylko produktu B
        self.assertEqual([call["dry_run"] for call in self.shop.calls], [True, False])
        self.assertTrue(app.session.changeset.startswith("ai-20261102-120000-Moj_plik"))
        self.assertEqual(self.questions, [])  # sklep testowy: potwierdzenie w przegladarce

        app.start_revert()
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)
        self.assertEqual(self.shop.reverts, [(app.session.changeset, False)])

    def test_production_apply_waits_for_windows_confirmation(self) -> None:
        app = self.make_app()
        self.preview(app, "prod", [row(2, "B", "Gniazdo nowe")])
        self.answer = False
        self.shop.calls.clear()
        app.start_apply()
        app.wait()
        self.assertEqual(app.job.state, "error")
        self.assertIn("nie potwierdzono", app.job.message)
        self.assertEqual(len(self.questions), 1)
        self.assertIn("PRODUKCJI", self.questions[0])
        self.assertFalse(any(not call["dry_run"] for call in self.shop.calls))

        self.answer = True
        app.start_apply()
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)
        self.assertTrue(any(not call["dry_run"] for call in self.shop.calls))

    def test_apply_stops_when_shop_changed_since_preview(self) -> None:
        app = self.make_app()
        self.preview(app, "prod", [row(2, "B", "Gniazdo nowe")])
        self.shop.titles[2] = "Gniazdo nowe"  # ktos zapisal to samo w sklepie - zmiana znika
        app.start_apply()
        app.wait()
        self.assertEqual(app.job.message, "Brak zmian do zapisania - sklep ma już te wartości.")

        self.shop.titles[2] = "Inny tytul"
        self.preview(app, "prod", [row(2, "B", "Gniazdo nowe", color="Czarny")])
        self.shop.titles[2] = "Gniazdo"  # sklep zmienil sie po podgladzie: zapis bylby inny niz na ekranie
        self.questions.clear()
        app.start_apply()
        app.wait()
        self.assertEqual(app.job.message, STALE_PREVIEW)
        self.assertEqual(self.questions, [])

    def test_preview_fetches_only_related_products_missing_from_file(self) -> None:
        app = self.make_app()
        changed = dict(row(2, "B", "Gniazdo nowe"), **{"Up-Sells": "A|SPOZA"})
        columns = COLUMNS + ["Up-Sells"]
        path = self.root / "plik.xlsx"
        write_products_xlsx(path, columns, [row(1, "A", "Gniazdo"), changed])
        app.start_preview("local", "plik.xlsx", path.read_bytes())
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)
        pulls = [call["pull"] for call in self.shop.calls if "pull" in call]
        self.assertEqual(pulls, [{"sku": ["SPOZA"]}])
        self.assertEqual(app.session.related_fetched, 1)
        card = app.card(0, "after")
        self.assertIn("Powiazany SPOZA", card)
        self.assertIn("Polecamy również", card)

    def tagged_file(self) -> bytes:
        columns = COLUMNS + ["Product Tags", "Atrybut Produktu: Kolor Producenta"]
        changed = dict(row(2, "B", "Gniazdo nowe"), **{"Product Tags": "Gniazdo antenowe Ospel", "Atrybut Produktu: Kolor Producenta": "Biały"})
        path = self.root / "plik.xlsx"
        write_products_xlsx(path, columns, [changed])
        return path.read_bytes()

    def test_live_variant_groups_fetch_siblings_by_tag(self) -> None:
        self.shop.groups = [
            {"name": "Ospel - kolor", "status": "publish", "source": "product_tag", "terms": ["gniazdo-antenowe-ospel"], "terms_all": False, "attributes": ["Kolor producenta"], "images": [], "swatches": [], "dropdown": []},
            {"name": "Szkic", "status": "draft", "source": "product_tag", "terms": ["gniazdo-antenowe-ospel"], "terms_all": False, "attributes": ["Kolor"], "images": [], "swatches": [], "dropdown": []},
        ]
        self.shop.tagged["Gniazdo antenowe Ospel"] = [
            {"id": "2", "SKU": "B", "Title": "Gniazdo nowe", "Product Tags": "Gniazdo antenowe Ospel"},  # juz w pliku - pomijany
            {"id": "7", "SKU": "B-18", "Title": "Gniazdo srebrne", "Product Tags": "Gniazdo antenowe Ospel", "Atrybut Produktu: Kolor Producenta": "Srebrny"},
        ]
        app = self.make_app()
        app.start_preview("local", "plik.xlsx", self.tagged_file())
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)
        self.assertEqual([c["pull"] for c in self.shop.calls if "pull" in c], [{"tag": ["Gniazdo antenowe Ospel"]}])
        self.assertTrue(app.session.variant_groups["live"])
        self.assertEqual(app.session.variant_groups["count"], 1)
        self.assertEqual(app.session.related_fetched, 1)
        card = app.card(0, "after")
        self.assertIn("Ospel - kolor", card)
        self.assertIn("Srebrny", card)
        self.assertNotIn("Szkic", card)

    def test_old_plugin_uses_snapshot_and_never_filters_by_tag(self) -> None:
        app = self.make_app()
        app.start_preview("local", "plik.xlsx", self.tagged_file())
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)
        self.assertFalse(app.session.variant_groups["live"])
        self.assertFalse(any("tag" in c["pull"] for c in self.shop.calls if "pull" in c))
        self.assertIn("sprzed 2.1", "\n".join(app.job.log))

    def shipping_file(self) -> bytes:
        path = self.root / "wysylka.xlsx"
        write_products_xlsx(path, COLUMNS + ["Shipping Class", "Stock", "Backorders", "Wysyłka w ciągu"], [
            dict(row(2, "B", "Gniazdo"), **{"Shipping Class": "Gabaryt", "Stock": "0", "Backorders": "notify", "Wysyłka w ciągu": "do 7 dni"}),
        ])
        return path.read_bytes()

    def test_shipping_methods_calculated_once_when_card_opens(self) -> None:
        self.shop.groups = []
        app = self.make_app()
        app.start_preview("local", "wysylka.xlsx", self.shipping_file())
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)
        self.assertFalse(any("shipping" in c for c in self.shop.calls))  # nic przed otwarciem karty

        after = app.card(0, "after")
        before = app.card(0, "before")
        shipping_calls = [c["shipping"] for c in self.shop.calls if "shipping" in c]
        self.assertEqual(shipping_calls, [[{"id": 2, "values": {"Shipping Class": "Gabaryt"}}]])
        self.assertIn("znikają: Paczkomat InPost", after)
        self.assertIn("ai-ship-method ai-removed", after)
        self.assertIn("Paczkomat InPost", before)
        self.assertNotIn("znikają", before)
        self.assertIn("Produkt na zamówienie", after)
        self.assertIn("pobranie 24,99 zł", after)

    def test_old_plugin_shows_note_instead_of_methods(self) -> None:
        app = self.make_app()
        app.start_preview("local", "wysylka.xlsx", self.shipping_file())
        app.wait()
        card = app.card(0, "after")
        app.card(0, "before")
        self.assertIn("wtyczką alkan-inator-bridge 2.1", card)
        self.assertIn("Klasa wysyłkowa: <strong>Gabaryt</strong>", card)
        self.assertEqual(len([c for c in self.shop.calls if "shipping" in c]), 1)

    def test_file_from_other_shop_stops_before_dry_run(self) -> None:
        columns = COLUMNS + ["Permalink"]
        path = self.root / "z_produkcji.xlsx"
        write_products_xlsx(path, columns, [dict(row(2, "B", "Gniazdo nowe"), Permalink="https://alkan.pl/sklep/gniazdo/")])
        app = self.make_app()
        app.start_preview("local", "z_produkcji.xlsx", path.read_bytes())
        app.wait()
        self.assertEqual(app.job.state, "error")
        self.assertIn("pobrano ze sklepu alkan.pl", app.job.message)
        self.assertIn("alkanlocal.test", app.job.message)
        self.assertEqual(self.shop.calls, [])

        app.start_preview("prod", "z_produkcji.xlsx", path.read_bytes())
        app.wait()
        self.assertEqual(app.job.state, "done", app.job.message)

    def test_errors_block_apply_and_columns_limit_preview(self) -> None:
        app = self.make_app()
        self.preview(app, "local", [row(2, "B", "Gniazdo nowe", color="Czarny"), row(3, "ZLE", "Gniazdo")])
        with self.assertRaises(ValueError):
            app.start_apply()
        self.assertEqual(app.session.all_columns, ["Title", "Atrybut Produktu: Kolor"])

        app.start_recalc(["Title"])
        app.wait()
        self.assertEqual(app.session.columns, ["Title"])
        self.assertEqual({c["column"] for v in app.session.views for c in v.changes}, {"Title"})

    def test_upload_copy_keeps_original_and_rejects_non_xlsx(self) -> None:
        app = self.make_app()
        with self.assertRaises(ValueError):
            app.start_preview("local", "x.xlsx", b"to nie xlsx")
        with self.assertRaises(ValueError):
            app.start_preview("nieznany", "x.xlsx", self.xlsx([row(1, "A", "Gniazdo")]))
        self.preview(app, "local", [row(2, "B", "Gniazdo nowe")])
        self.assertTrue(app.session.upload.is_relative_to(self.root / "work"))
        headers, rows = read_products_xlsx(app.session.upload)
        self.assertEqual(rows[0]["Title"], "Gniazdo nowe")
        self.assertEqual(safe_file_name("Łączniki 2026 (kopia).xlsx"), "Laczniki_2026_kopia_.xlsx")


class ServerTests(TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        config = Path(self.tmp.name) / "config.json"
        config.write_text(json.dumps({"url": "https://alkanlocal.test:8443", "token": "SEKRET"}), encoding="utf-8")
        self.app = App(theme_dir=None, config_path=config, work_dir=Path(self.tmp.name) / "work")
        self.server = PreviewServer(self.app)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self) -> None:
        self.server.shutdown()
        self.server.server_close()
        self.tmp.cleanup()

    def get(self, path: str, *, host: str | None = None, key: str | None = None) -> tuple[int, str]:
        connection = http.client.HTTPConnection("127.0.0.1", self.app.port, timeout=10)
        headers = {"Host": host or f"127.0.0.1:{self.app.port}"}
        if key:
            headers["X-Key"] = key
        connection.request("GET", path, headers=headers)
        response = connection.getresponse()
        body = response.read().decode("utf-8")
        connection.close()
        return response.status, body

    def test_api_requires_key_and_local_host(self) -> None:
        self.assertEqual(self.get("/api/state")[0], 403)
        self.assertEqual(self.get("/api/state", key="zly")[0], 403)
        self.assertEqual(self.get("/api/state", key=self.app.key, host="evil.example:80")[0], 403)
        status, body = self.get("/api/state", key=self.app.key)
        self.assertEqual(status, 200)
        self.assertNotIn("SEKRET", body)
        self.assertEqual(json.loads(body)["profiles"][0]["name"], "local")

    def test_page_and_static_files(self) -> None:
        self.assertEqual(self.get("/")[0], 403)
        status, body = self.get(f"/?k={self.app.key}")
        self.assertEqual(status, 200)
        self.assertIn("Wgraj zmiany do sklepu", body)
        self.assertEqual(self.get("/static/app.js")[0], 200)
        self.assertEqual(self.get("/static/../server.py")[0], 404)
