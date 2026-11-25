"""Lista klientów ze sklepu - te same dane co eksport w WooCommerce > Klienci.

Wtyczka alkan-inator-bridge 2.2+, odczyt włączany w panelu (Alkan > AlkanInator API > Odczyt klientów).
To dane osobowe: plik trafia do output/ (poza gitem) i nie powinien wychodzić poza firmę.
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Callable

from openpyxl import Workbook

from woo_api.client import WooBridgeClient, WooBridgeError

# Klucz z API -> nagłówek kolumny (jak w eksporcie z zakładki Klienci).
COLUMNS: list[tuple[str, str]] = [
    ("name", "Imię i nazwisko"),
    ("username", "Nazwa użytkownika"),
    ("date_last_active", "Ostatnia aktywność"),
    ("date_registered", "Data rejestracji"),
    ("email", "E-mail"),
    ("orders_count", "Zamówienia"),
    ("total_spend", "Łączne wydatki"),
    ("avg_order_value", "Śr. wartość zamówienia"),
    ("country", "Kraj / region"),
    ("city", "Miasto"),
    ("state", "Region"),
    ("postcode", "Kod pocztowy"),
]


def pull_customers(
    client: WooBridgeClient, output: str | Path, *, per_page: int = 1000, log: Callable[[str], None] | None = None
) -> int:
    """Pobiera wszystkich klientów strona po stronie i zapisuje plik; zwraca liczbę klientów."""
    log = log or (lambda _: None)
    customers: list[dict[str, Any]] = []
    try:
        for data in client.iter_customer_pages(per_page=per_page):
            customers.extend(data.get("customers") or [])
            log(f"Klienci: {len(customers)} z {data.get('total', '?')}")
    except WooBridgeError as exc:
        if exc.status_code == 404:
            raise WooBridgeError("Pobieranie klientów wymaga wtyczki alkan-inator-bridge 2.2 w tym sklepie.") from exc
        raise
    write_customers(customers, output)
    return len(customers)


def write_customers(customers: list[dict[str, Any]], path: str | Path) -> Path:
    """XLSX (domyślnie) albo CSV ze średnikiem i BOM, żeby polski Excel otworzył go poprawnie."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    header = [label for _, label in COLUMNS]
    rows = [[_value(customer.get(key)) for key, _ in COLUMNS] for customer in customers]

    if path.suffix.lower() == ".csv":
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(header)
            writer.writerows(rows)
        return path

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Klienci"
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A2"
    workbook.save(path)
    return path


def _value(value: Any) -> Any:
    if value is None:
        return ""
    return value
