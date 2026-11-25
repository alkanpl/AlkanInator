from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from unittest import TestCase

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from woo_api.client import WooBridgeClient, WooBridgeError
from woo_api.customers import COLUMNS, pull_customers, write_customers


def customer(number: int) -> dict[str, object]:
    return {
        "id": number,
        "name": f"Jan Kowalski {number}",
        "username": "",
        "email": f"jan{number}@example.com",
        "date_registered": None,
        "date_last_active": "2026-11-20 10:00:00",
        "orders_count": 2,
        "total_spend": 123.45,
        "avg_order_value": 61.73,
        "country": "PL",
        "city": "Łódź",
        "state": "",
        "postcode": "90-001",
    }


class PagedClient(WooBridgeClient):
    """Sklep z 5 klientami, 2 na stronę; zapamiętuje pobrane strony."""

    def __init__(self, error: WooBridgeError | None = None) -> None:
        self.pages: list[int] = []
        self.error = error

    def customers_page(self, page: int = 1, per_page: int = 1000) -> dict[str, object]:
        if self.error:
            raise self.error
        self.pages.append(page)
        everyone = [customer(n) for n in range(1, 6)]
        return {
            "total": 5,
            "total_pages": 3,
            "page": page,
            "per_page": 2,
            "customers": everyone[(page - 1) * 2 : page * 2],
        }


class CustomersTest(TestCase):
    def test_pulls_every_page_into_xlsx(self) -> None:
        client = PagedClient()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "klienci.xlsx"
            count = pull_customers(client, path, per_page=2)
            sheet = load_workbook(path).active
            rows = list(sheet.iter_rows(values_only=True))

        self.assertEqual(count, 5)
        self.assertEqual(client.pages, [1, 2, 3])
        self.assertEqual(list(rows[0]), [label for _, label in COLUMNS])
        self.assertEqual(len(rows), 6)
        self.assertEqual(rows[1][0], "Jan Kowalski 1")
        self.assertEqual(rows[1][9], "Łódź")

    def test_csv_uses_semicolon_and_bom_for_polish_excel(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = write_customers([customer(1)], Path(tmp) / "klienci.csv")
            raw = path.read_bytes()

        self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))
        lines = raw.decode("utf-8-sig").splitlines()
        self.assertEqual(lines[0].split(";")[0], "Imię i nazwisko")
        self.assertIn("Łódź", lines[1].split(";"))
        # Brak daty rejestracji (gość) to pusta komórka, nie "None".
        self.assertEqual(lines[1].split(";")[3], "")

    def test_old_plugin_gets_clear_message(self) -> None:
        client = PagedClient(WooBridgeError("HTTP 404: brak", status_code=404))
        with tempfile.TemporaryDirectory() as tmp, self.assertRaises(WooBridgeError) as caught:
            pull_customers(client, Path(tmp) / "klienci.xlsx")

        self.assertIn("2.2", str(caught.exception))

    def test_switched_off_reading_is_reported_as_is(self) -> None:
        error = WooBridgeError("HTTP 403: Odczyt klientów przez API jest wyłączony", status_code=403)
        with tempfile.TemporaryDirectory() as tmp, self.assertRaises(WooBridgeError) as caught:
            pull_customers(PagedClient(error), Path(tmp) / "klienci.xlsx")

        self.assertIn("wyłączony", str(caught.exception))
