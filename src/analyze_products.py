from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, enrich_config_with_catalog_knowledge, load_catalog_knowledge
from extract_attributes import ATTRIBUTE_COLUMNS
from utils import detect_column, ensure_dir, is_blank, load_yaml, read_products, valid_ean


def detect_product_columns(df: pd.DataFrame, global_config: dict[str, Any] | None = None) -> dict[str, str | None]:
    config = global_config or load_yaml("configs/global.yaml")
    return {
        "title": detect_column(df, config.get("default_title_column_candidates", [])),
        "sku": detect_column(df, config.get("default_sku_column_candidates", [])),
        "ean": detect_column(df, config.get("default_ean_column_candidates", [])),
        "producer": detect_column(df, config.get("default_producer_column_candidates", [])),
        "category": detect_column(df, config.get("default_category_column_candidates", [])),
    }


def analyze_dataframe(df: pd.DataFrame, columns: dict[str, str | None], config: dict[str, Any]) -> dict[str, Any]:
    title_column = columns.get("title")
    sku_column = columns.get("sku")
    ean_column = columns.get("ean")
    producer_column = columns.get("producer")
    category_column = columns.get("category")
    max_title_length = int(config.get("max_title_length", 110))
    min_title_length = int(config.get("min_title_length", 40))

    empty_fields = {column: int(df[column].map(is_blank).sum()) for column in df.columns}
    title_lengths = df[title_column].astype(str).map(len) if title_column else pd.Series(dtype=int)

    summary: dict[str, Any] = {
        "rows": int(len(df)),
        "columns": list(df.columns),
        "detected_columns": columns,
        "empty_fields": empty_fields,
        "categories_count": int(df[category_column].nunique()) if category_column else 0,
        "producers_count": int(df[producer_column].nunique()) if producer_column else 0,
        "titles_too_short": int((title_lengths < min_title_length).sum()) if title_column else 0,
        "titles_too_long": int((title_lengths > max_title_length).sum()) if title_column else 0,
        "products_without_producer": int(df[producer_column].map(is_blank).sum()) if producer_column else int(len(df)),
        "products_without_ean": int(df[ean_column].map(is_blank).sum()) if ean_column else int(len(df)),
        "duplicate_sku": int(df[sku_column].duplicated().sum()) if sku_column else 0,
        "duplicate_ean": int(df[df[ean_column].astype(str).str.strip() != ""][ean_column].duplicated().sum()) if ean_column else 0,
        "invalid_ean": int((~df[ean_column].map(valid_ean) & (df[ean_column].astype(str).str.strip() != "")).sum()) if ean_column else 0,
    }

    return summary


def build_missing_attributes(df: pd.DataFrame, columns: dict[str, str | None], config: dict[str, Any]) -> pd.DataFrame:
    sku_column = columns.get("sku")
    title_column = columns.get("title")
    rows: list[dict[str, Any]] = []

    for _, row in df.iterrows():
        role = str(row.get("product_role", "main") or "main")
        required_fields = (config.get("required_fields_by_role") or {}).get(role, config.get("required_fields") or [])
        missing = []
        for field in required_fields:
            if field == "moc" and not is_blank(row.get("attr_gwint", row.get("gwint", ""))):
                continue
            value = row.get(f"attr_{field}", row.get(field, ""))
            if is_blank(value):
                missing.append(field)
        if missing:
            rows.append(
                {
                    "sku": row.get(sku_column, "") if sku_column else "",
                    "title": row.get(title_column, "") if title_column else "",
                    "missing_attributes": ";".join(missing),
                }
            )

    return pd.DataFrame(rows, columns=["sku", "title", "missing_attributes"])


def build_category_summary(df: pd.DataFrame, columns: dict[str, str | None]) -> pd.DataFrame:
    category_column = columns.get("category")
    if not category_column:
        return pd.DataFrame(columns=["category", "products_count"])
    return (
        df.groupby(category_column, dropna=False)
        .size()
        .reset_index(name="products_count")
        .rename(columns={category_column: "category"})
        .sort_values("products_count", ascending=False)
    )


def write_quality_report(summary: dict[str, Any], output_path: str | Path) -> None:
    path = Path(output_path)
    ensure_dir(path.parent)
    detected = summary["detected_columns"]
    empty_rows = "\n".join(
        f"<tr><td>{column}</td><td>{count}</td></tr>" for column, count in summary["empty_fields"].items()
    )
    html = f"""<!doctype html>
<html lang="pl">
<head>
  <meta charset="utf-8">
  <title>Quality report</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 32px; color: #1f2933; }}
    table {{ border-collapse: collapse; width: 100%; margin: 16px 0; }}
    th, td {{ border: 1px solid #d9e2ec; padding: 8px; text-align: left; }}
    th {{ background: #f0f4f8; }}
    code {{ background: #f0f4f8; padding: 2px 4px; }}
  </style>
</head>
<body>
  <h1>Raport jakosci danych</h1>
  <h2>Podsumowanie</h2>
  <ul>
    <li>Liczba produktow: {summary["rows"]}</li>
    <li>Liczba kolumn: {len(summary["columns"])}</li>
    <li>Liczba kategorii: {summary["categories_count"]}</li>
    <li>Liczba producentow: {summary["producers_count"]}</li>
    <li>Produkty bez producenta: {summary["products_without_producer"]}</li>
    <li>Produkty bez EAN: {summary["products_without_ean"]}</li>
    <li>Niepoprawne EAN: {summary["invalid_ean"]}</li>
    <li>Duplikaty SKU: {summary["duplicate_sku"]}</li>
    <li>Duplikaty EAN: {summary["duplicate_ean"]}</li>
    <li>Tytuly za krotkie: {summary["titles_too_short"]}</li>
    <li>Tytuly za dlugie: {summary["titles_too_long"]}</li>
  </ul>
  <h2>Wykryte kolumny</h2>
  <table>
    <tr><th>Rola</th><th>Kolumna</th></tr>
    <tr><td>title</td><td><code>{detected.get("title") or ""}</code></td></tr>
    <tr><td>sku</td><td><code>{detected.get("sku") or ""}</code></td></tr>
    <tr><td>ean</td><td><code>{detected.get("ean") or ""}</code></td></tr>
    <tr><td>producer</td><td><code>{detected.get("producer") or ""}</code></td></tr>
    <tr><td>category</td><td><code>{detected.get("category") or ""}</code></td></tr>
  </table>
  <h2>Puste pola</h2>
  <table>
    <tr><th>Kolumna</th><th>Puste wartosci</th></tr>
    {empty_rows}
  </table>
</body>
</html>
"""
    path.write_text(html, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analizuje plik produktow.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet")
    parser.add_argument("--config", default="configs/categories/oprawy-sufitowe.yaml")
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE_PATH)
    parser.add_argument("--reports-dir", default="reports")
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    config = enrich_config_with_catalog_knowledge(load_yaml(args.config), load_catalog_knowledge(args.catalog_knowledge))
    columns = detect_product_columns(df)
    summary = analyze_dataframe(df, columns, config)
    reports_dir = ensure_dir(args.reports_dir)
    write_quality_report(summary, reports_dir / "quality_report.html")
    build_category_summary(df, columns).to_csv(reports_dir / "category_summary.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(columns=["sku", "title", "missing_attributes"]).to_csv(
        reports_dir / "missing_attributes.csv", index=False, encoding="utf-8-sig"
    )


if __name__ == "__main__":
    main()
