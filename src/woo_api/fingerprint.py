"""Odciski wierszy z chwili pobrania: push wysyła do sklepu tylko wiersze zmienione w pliku.

pull dopisuje ukryty arkusz "_alkan": wiersze meta (#site, #pulled_at, #columns) i dla każdego
produktu klucz (id:123 albo sku:X) ze skrótem wartości kolumn zapisywalnych. push liczy skrót od
nowa; wiersz z tym samym skrótem nie trafia do dry-runu ani zapisu - przyspiesza to dry-run
i nie cofa zmian zrobionych w sklepie po pobraniu w produktach, których nikt w pliku nie ruszał.
Plik bez arkusza (WP All Export, starszy pull) wysyła się w całości, tak jak wcześniej.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook

from woo_api.workbook import ROW_NUMBER_KEY


SHEET = "_alkan"


@dataclass
class Fingerprints:
    site: str
    pulled_at: str
    columns: list[str]
    rows: dict[str, str] = field(default_factory=dict)


def row_key(row: dict[str, Any]) -> str:
    """Ten sam klucz co prepare_rows w push: id, a bez id - SKU."""
    product_id = row.get("id")
    return f"id:{product_id}" if product_id not in (None, "") else f"sku:{row.get('SKU', '')}"


def _text(value: Any) -> str:
    # Tak jak wartość wróci z pliku: write_products_xlsx zapisuje str(), pusta komórka czyta się jako "".
    return "" if value is None else str(value)


def row_fingerprint(row: dict[str, Any], columns: list[str]) -> str:
    payload = json.dumps([_text(row.get(column)) for column in columns], ensure_ascii=False)
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:20]


def fingerprint_columns(columns: Iterable[str], skip: Iterable[str]) -> list[str]:
    skipped = set(skip)
    return [column for column in columns if column not in skipped]


def fingerprint_sheet(rows: list[dict[str, Any]], columns: list[str], *, site: str, pulled_at: str) -> list[list[Any]]:
    """Zawartość ukrytego arkusza dla write_products_xlsx(hidden_sheets=...)."""
    sheet: list[list[Any]] = [
        ["#site", site],
        ["#pulled_at", pulled_at],
        ["#columns", json.dumps(columns, ensure_ascii=False)],
    ]
    sheet.extend([row_key(row), row_fingerprint(row, columns)] for row in rows)
    return sheet


def read_fingerprints(path: str | Path) -> Fingerprints | None:
    workbook = load_workbook(Path(path), read_only=True, data_only=True)
    try:
        if SHEET not in workbook.sheetnames:
            return None
        meta: dict[str, str] = {}
        rows: dict[str, str] = {}
        for key, value, *_ in workbook[SHEET].iter_rows(values_only=True):
            if key is None:
                continue
            if str(key).startswith("#"):
                meta[str(key)] = _text(value)
            else:
                rows[str(key)] = _text(value)
    finally:
        workbook.close()
    try:
        columns = json.loads(meta.get("#columns", ""))
    except json.JSONDecodeError:
        return None
    return Fingerprints(site=meta.get("#site", ""), pulled_at=meta.get("#pulled_at", ""), columns=columns, rows=rows)


def changed_rows(prints: Fingerprints, headers: list[str], rows: list[dict[str, Any]], skip: Iterable[str]) -> set[int]:
    """Numery wierszy do wysłania: bez odcisku (dopisane), z innym skrótem albo z wartością w nowej kolumnie.

    Pusta komórka w kolumnie dodanej po pobraniu nie liczy się jako zmiana - inaczej każda nowa
    kolumna czyściłaby pole we wszystkich produktach pliku (do tego służy push --all-rows).
    """
    known = set(prints.columns) | set(skip)
    new_columns = [column for column in headers if column not in known]
    changed: set[int] = set()
    for row in rows:
        stored = prints.rows.get(row_key(row))
        if (
            stored is None
            or row_fingerprint(row, prints.columns) != stored
            or any(_text(row.get(column)).strip() for column in new_columns)
        ):
            changed.add(row[ROW_NUMBER_KEY])
    return changed
