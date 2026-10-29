from __future__ import annotations

import csv
import json
import sys
import base64
import tempfile
from argparse import Namespace
from io import BytesIO
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from baselinker_api import images as images_module
from baselinker_api.export import ExportOptions, export_inventory_csv
from baselinker_api.images import build_images_payload, external_image_value
from baselinker_api.input_data import load_product_records
from baselinker_api.sync import SyncOptions, build_category_maps, sync_inventory
from baselinker_api_cli import _token, _token_from_file


class BaseLinkerTokenTests(TestCase):
    def test_reads_plain_token_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "token.txt"
            path.write_text("sekretny-token\n", encoding="utf-8")
            self.assertEqual(_token_from_file(path), "sekretny-token")

    def test_reads_token_from_json_config(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "config.json"
            path.write_text(json.dumps({"token": "token-z-json"}), encoding="utf-8")
            args = Namespace(token_file=str(path), token_env="NIEISTNIEJACA_ZMIENNA_TESTOWA")
            self.assertEqual(_token(args), "token-z-json")


class FakeClient:
    def __init__(self) -> None:
        self.added_products: list[dict[str, object]] = []
        self.added_categories: list[dict[str, object]] = []
        self.added_manufacturers: list[str] = []

    def get_products_list(self, inventory_id: int, **kwargs):
        if kwargs.get("page", 1) > 1:
            return []
        return [{"id": 101, "name": "Stara nazwa", "sku": "A1/SIM", "ean": "5901234123457"}]

    def get_products_data(self, inventory_id: int, product_ids: list[int]):
        if 101 not in product_ids:
            return {}
        return {
            "101": {
                "sku": "A1/SIM",
                "ean": "5901234123457",
                "manufacturer_id": 5,
                "category_id": 10,
                "text_fields": {
                    "name|pl": "Stara nazwa",
                    "description|pl": "Stary opis",
                    "features|pl": {"Seria": "Simon 54", "Kolor": "biały"},
                },
                "images": {},
            }
        }

    def get_categories(self, inventory_id: int):
        return [
            {"category_id": 1, "name": "Osprzęt", "parent_id": 0},
            {"category_id": 10, "name": "Simon 54", "parent_id": 1},
        ]

    def get_manufacturers(self):
        return [{"manufacturer_id": 5, "name": "Simon"}]

    def add_category(self, parameters: dict[str, object]):
        self.added_categories.append(parameters)
        return {"status": "SUCCESS", "category_id": 20 + len(self.added_categories)}

    def add_manufacturer(self, name: str):
        self.added_manufacturers.append(name)
        return {"status": "SUCCESS", "manufacturer_id": 30 + len(self.added_manufacturers)}

    def add_product(self, parameters: dict[str, object]):
        self.added_products.append(parameters)
        return {"status": "SUCCESS", "product_id": parameters.get("product_id"), "warnings": {}}


def _write_input(path: Path, rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(
            stream,
            fieldnames=[
                "product_id",
                "name",
                "sku",
                "ean",
                "manufacturer_name",
                "category",
                "description",
                "features",
                "images_urls",
            ],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerows(rows)


def _valid_row(**changes: str) -> dict[str, str]:
    row = {
        "product_id": "101",
        "name": "Łącznik kaszmirowy Kontakt-Simon Simon 54 A1",
        "sku": "A1/SIM",
        "ean": "5901234123457",
        "manufacturer_name": "Simon",
        "category": "Osprzęt > Simon 54",
        "description": "<p>Nowy opis produktu.</p>",
        "features": '{"Kolor":"kaszmirowy","Napięcie":"250 V"}',
        "images_urls": "https://example.com/A1.jpg",
    }
    row.update(changes)
    return row


class BaseLinkerInputTests(TestCase):
    def test_loads_canonical_csv_and_features(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "input.csv"
            _write_input(path, [_valid_row()])
            records = load_product_records(path)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].name, "Łącznik kaszmirowy Kontakt-Simon Simon 54 A1")
        self.assertEqual(records[0].features, {"Kolor": "kaszmirowy", "Napięcie": "250 V"})
        self.assertEqual(records[0].images_urls, ["https://example.com/A1.jpg"])

    def test_marks_duplicate_sku_and_invalid_features(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "input.csv"
            _write_input(path, [_valid_row(product_id="", features="{"), _valid_row(product_id="", ean="")])
            records = load_product_records(path)
        self.assertTrue(any("Powtórzona wartość sku" in error for error in records[0].errors))
        self.assertTrue(any("Niepoprawny JSON" in error for error in records[0].errors))


def _png_bytes(mode: str = "RGBA", size: tuple[int, int] = (500, 300)) -> bytes:
    """PNG dostawcy: przezroczyste tlo z kolorowym prostokatem posrodku."""
    if mode == "P":
        image = Image.new("P", size, 0)
        image.putpalette([0, 0, 0, 200, 30, 30] + [0] * 762)
        image.paste(1, (size[0] // 4, size[1] // 4, size[0] * 3 // 4, size[1] * 3 // 4))
        image.info["transparency"] = 0
    else:
        image = Image.new("RGBA", size, (0, 0, 0, 0))
        image.paste((200, 30, 30, 255), (size[0] // 4, size[1] // 4, size[0] * 3 // 4, size[1] * 3 // 4))
    buffer = BytesIO()
    image.save(buffer, format="PNG", transparency=0) if mode == "P" else image.save(buffer, format="PNG")
    return buffer.getvalue()


def _decode(value: str) -> Image.Image:
    return Image.open(BytesIO(base64.b64decode(value[len("data:"):])))


class BaseLinkerImageTests(TestCase):
    def test_external_urls_require_https_and_skip_base_cdn(self) -> None:
        self.assertIsNotNone(external_image_value("https://example.com/a.jpg")[0])
        self.assertIsNone(external_image_value("http://example.com/a.jpg")[0])
        self.assertIsNone(external_image_value("https://upload.cdn.baselinker.com/products/a.jpg")[0])

    def test_local_image_is_encoded_under_api_limit(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            image_path = root / "a.png"
            Image.new("RGBA", (50, 80), (10, 20, 30, 128)).save(image_path)
            payload, warnings = build_images_payload([], ["a.png"], input_dir=root)
        self.assertEqual(list(payload), ["0"])
        self.assertTrue(payload["0"].startswith("data:"))
        self.assertLessEqual(len(payload["0"]) - 5, 2_000_000)
        self.assertFalse(any("Nie można" in warning for warning in warnings))

    def test_remote_image_is_downloaded_and_padded_to_1000_on_white(self) -> None:
        for mode in ("RGBA", "P"):
            with self.subTest(mode=mode), patch.object(images_module, "fetch_image_bytes", return_value=_png_bytes(mode)):
                payload, warnings = build_images_payload(["https://example.com/a.png"], [], input_dir=Path("."))
            self.assertEqual(warnings, [])
            image = _decode(payload["0"])
            self.assertEqual(image.size, (1000, 1000))
            self.assertEqual(image.format, "JPEG")
            # (50, 250) lezy w obszarze samego zdjecia (500x300 -> 1000x600 od y=200), w przezroczystym tle.
            for corner in ((2, 2), (997, 997), (50, 250)):
                self.assertTrue(all(channel >= 250 for channel in image.getpixel(corner)), (mode, corner))

    def test_raw_urls_option_keeps_original_address(self) -> None:
        with patch.object(images_module, "fetch_image_bytes", side_effect=AssertionError("bez pobierania")):
            payload, _ = build_images_payload(["https://example.com/a.png"], [], input_dir=Path("."), raw_urls=True)
        self.assertEqual(payload, {"0": "url:https://example.com/a.png"})

    def test_failed_download_is_reported_not_sent_raw(self) -> None:
        with patch.object(images_module, "fetch_image_bytes", side_effect=OSError("timeout")):
            payload, warnings = build_images_payload(["https://example.com/a.png"], [], input_dir=Path("."))
        self.assertEqual(payload, {})
        self.assertTrue(any("Nie można pobrać" in warning for warning in warnings))


class BaseLinkerMappingTests(TestCase):
    def test_category_full_paths_are_resolved_and_duplicate_leaves_rejected(self) -> None:
        mapping, duplicates = build_category_maps(
            [
                {"category_id": 1, "name": "A", "parent_id": 0},
                {"category_id": 2, "name": "B", "parent_id": 0},
                {"category_id": 3, "name": "Ramki", "parent_id": 1},
                {"category_id": 4, "name": "Ramki", "parent_id": 2},
            ]
        )
        self.assertEqual(mapping["a > ramki"], 3)
        self.assertEqual(mapping["b > ramki"], 4)
        self.assertIn("ramki", duplicates)
        self.assertNotIn("ramki", mapping)


class BaseLinkerSyncTests(TestCase):
    def setUp(self) -> None:
        fetch = patch.object(images_module, "fetch_image_bytes", return_value=_png_bytes())
        fetch.start()
        self.addCleanup(fetch.stop)

    def test_auto_match_uses_sku_before_stale_woocommerce_product_id(self) -> None:
        client = FakeClient()
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.csv"
            _write_input(source, [_valid_row(product_id="999999")])
            result = sync_inventory(client, source, SyncOptions(inventory_id=1), report_dir=root / "report")
        self.assertEqual(result.audit_rows[0].product_id, "101")
        self.assertEqual(result.audit_rows[0].status, "READY")

    def test_dry_run_builds_full_patch_and_reports_without_writing(self) -> None:
        client = FakeClient()
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.csv"
            report = root / "report"
            _write_input(source, [_valid_row()])
            result = sync_inventory(client, source, SyncOptions(inventory_id=1), report_dir=report)
            summary = json.loads((report / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(result.audit_rows[0].status, "READY")
        self.assertEqual(client.added_products, [])
        self.assertIn("name", result.audit_rows[0].changed_fields)
        self.assertIn("description", result.audit_rows[0].changed_fields)
        self.assertIn("features", result.audit_rows[0].changed_fields)
        self.assertIn("images", result.audit_rows[0].changed_fields)
        self.assertEqual(summary["mode"], "DRY_RUN")

    def test_apply_sends_name_description_merged_features_and_images(self) -> None:
        client = FakeClient()
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.csv"
            _write_input(source, [_valid_row()])
            result = sync_inventory(
                client,
                source,
                SyncOptions(inventory_id=1, apply=True),
                report_dir=root / "report",
            )
        self.assertEqual(result.audit_rows[0].status, "SUCCESS")
        self.assertEqual(len(client.added_products), 1)
        payload = client.added_products[0]
        text_fields = payload["text_fields"]
        self.assertEqual(text_fields["name|pl"], _valid_row()["name"])
        self.assertEqual(text_fields["features|pl"]["Seria"], "Simon 54")
        self.assertEqual(text_fields["features|pl"]["Kolor"], "kaszmirowy")
        self.assertEqual(_decode(payload["images"]["0"]).size, (1000, 1000))

    def test_batch_is_blocked_when_any_input_row_is_invalid(self) -> None:
        client = FakeClient()
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.csv"
            _write_input(source, [_valid_row(), _valid_row(product_id="", sku="", ean="")])
            result = sync_inventory(
                client,
                source,
                SyncOptions(inventory_id=1, apply=True),
                report_dir=root / "report",
            )
        self.assertTrue(result.blocked)
        self.assertEqual(client.added_products, [])
        self.assertIn("BLOCKED_BATCH", {row.status for row in result.audit_rows})
        self.assertIn("INVALID", {row.status for row in result.audit_rows})

    def test_missing_category_and_manufacturer_can_be_created_explicitly(self) -> None:
        client = FakeClient()
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.csv"
            _write_input(source, [_valid_row(category="Osprzęt > Simon 55", manufacturer_name="Nowa marka")])
            result = sync_inventory(
                client,
                source,
                SyncOptions(
                    inventory_id=1,
                    apply=True,
                    create_missing_categories=True,
                    create_missing_manufacturers=True,
                ),
                report_dir=root / "report",
            )
        self.assertEqual(result.audit_rows[0].status, "SUCCESS")
        self.assertEqual(client.added_categories[-1]["name"], "Simon 55")
        self.assertEqual(client.added_manufacturers, ["Nowa marka"])
        self.assertIn("category_id", client.added_products[0])
        self.assertIn("manufacturer_id", client.added_products[0])

    def test_new_category_goes_under_unique_parent_path_despite_duplicate_root(self) -> None:
        # Base ma kilka korzeni "Osprzęt"; pelna sciezka rodzica jest jednak unikalna,
        # wiec nowa kategoria ma trafic pod nia, bez zakladania kolejnego korzenia.
        client = FakeClient()
        client.get_categories = lambda inventory_id: [
            {"category_id": 1, "name": "Osprzęt", "parent_id": 0},
            {"category_id": 2, "name": "Osprzęt", "parent_id": 0},
            {"category_id": 10, "name": "Simon 54", "parent_id": 1},
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.csv"
            _write_input(source, [_valid_row(category="Osprzęt > Simon 54 > Ramki")])
            result = sync_inventory(
                client,
                source,
                SyncOptions(inventory_id=1, apply=True, create_missing_categories=True),
                report_dir=root / "report",
            )
        self.assertEqual(result.audit_rows[0].status, "SUCCESS")
        self.assertEqual(client.added_categories, [{"inventory_id": 1, "name": "Ramki", "parent_id": 10}])

    def test_category_under_ambiguous_root_is_not_created(self) -> None:
        client = FakeClient()
        client.get_categories = lambda inventory_id: [
            {"category_id": 1, "name": "Osprzęt", "parent_id": 0},
            {"category_id": 2, "name": "Osprzęt", "parent_id": 0},
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "input.csv"
            _write_input(source, [_valid_row(category="Osprzęt > Simon 24")])
            with self.assertRaises(RuntimeError):
                sync_inventory(
                    client,
                    source,
                    SyncOptions(inventory_id=1, apply=True, create_missing_categories=True),
                    report_dir=root / "report",
                )
        self.assertEqual(client.added_categories, [])
        self.assertEqual(client.added_products, [])


class BaseLinkerExportTests(TestCase):
    def test_export_writes_canonical_csv_with_full_category_path(self) -> None:
        client = FakeClient()
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "export.csv"
            count = export_inventory_csv(client, output, ExportOptions(inventory_id=1))
            with output.open("r", encoding="utf-8-sig", newline="") as stream:
                rows = list(csv.DictReader(stream, delimiter=";"))
        self.assertEqual(count, 1)
        self.assertEqual(rows[0]["category"], "Osprzęt > Simon 54")
        self.assertEqual(rows[0]["manufacturer_name"], "Simon")
        self.assertEqual(json.loads(rows[0]["features"])["Seria"], "Simon 54")
