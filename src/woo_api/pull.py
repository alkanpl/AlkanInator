from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from woo_api.client import WooBridgeClient
from woo_api.workbook import merge_columns, read_products_xlsx, write_products_xlsx


def read_sku_file(path: str | Path) -> list[str]:
    """SKU z kolumny "SKU" pliku XLSX albo po jednym w linii pliku TXT/CSV."""
    path = Path(path)
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        headers, rows = read_products_xlsx(path)
        if "SKU" not in headers:
            raise ValueError(f"{path}: brak kolumny SKU")
        values = [row.get("SKU", "") for row in rows]
    else:
        values = [line.split(";")[0].split(",")[0] for line in path.read_text(encoding="utf-8-sig").splitlines()]
    skus = [value.strip() for value in values if value and value.strip() and value.strip().upper() != "SKU"]
    return list(dict.fromkeys(skus))


def pull_products(
    client: WooBridgeClient,
    output: str | Path,
    *,
    filters: dict[str, Any],
    per_page: int = 200,
    log: Callable[[str], None] = print,
) -> tuple[int, list[str]]:
    rows: list[dict[str, Any]] = []
    column_lists: list[list[str]] = []
    for page in client.iter_product_pages(per_page=per_page, **filters):
        rows.extend(page.get("rows", []))
        column_lists.append(page.get("columns", []))
        log(f"Strona {page.get('page')}/{max(1, int(page.get('total_pages') or 0))}: {len(rows)}/{page.get('total')} produktów")
    warnings = write_products_xlsx(output, merge_columns(column_lists), rows)
    return len(rows), warnings
