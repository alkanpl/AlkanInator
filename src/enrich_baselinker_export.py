from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, load_catalog_knowledge
from export_to_baselinker_csv import build_manufacturer_data_by_producer
from recommend_categories import (
    DEFAULT_OVERRIDES_PATH,
    DEFAULT_TAXONOMY_PATH,
    CategoryRecommender,
    detect_recommendation_columns,
)
from utils import compact_spaces, load_yaml, read_products


DEFAULT_MANUFACTURER_DATA_ATTRIBUTE = "Dane producenta"
BASELINKER_CATEGORY_COLUMN = "category"
CAT_PENDANT = "Oświetlenie > Oświetlenie wewnętrzne > Lampy wiszące i żyrandole"
CAT_SURFACE = "Oświetlenie > Oświetlenie wewnętrzne > Oprawy natynkowe"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Dopisuje kategorie i dane producenta do gotowego CSV importowego Baselinkera."
    )
    parser.add_argument("--input", required=True, help="Gotowy CSV Baselinkera.")
    parser.add_argument("--output", required=True, help="Docelowy CSV Baselinkera.")
    parser.add_argument("--delimiter", default=";", help="Separator CSV. Domyslnie srednik.")
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE_PATH)
    parser.add_argument("--taxonomy", default=DEFAULT_TAXONOMY_PATH)
    parser.add_argument("--overrides", default=DEFAULT_OVERRIDES_PATH)
    parser.add_argument("--category-reference", default="", help="Opcjonalny CSV/XLSX z zaakceptowanymi kategoriami po SKU.")
    parser.add_argument("--category-reference-sheet", default=None, help="Arkusz pliku referencyjnego XLSX.")
    parser.add_argument("--reference-sku-column", default="Kod")
    parser.add_argument("--reference-category-column", default="proponowana_kategoria_1")
    parser.add_argument("--corrections-report", default="", help="CSV z lista automatycznych korekt kategorii.")
    parser.add_argument("--top-n", type=int, default=3)
    args = parser.parse_args()

    rows, fieldnames = read_csv_rows(args.input, args.delimiter)
    catalog_knowledge = load_catalog_knowledge(args.catalog_knowledge)
    taxonomy = load_yaml(args.taxonomy) if args.taxonomy else {}
    overrides = load_yaml(args.overrides) if args.overrides else {}
    recommender = CategoryRecommender(catalog_knowledge, taxonomy, overrides)
    manufacturer_data = build_manufacturer_data_by_producer(catalog_knowledge)
    category_reference = load_category_reference(
        args.category_reference,
        args.category_reference_sheet,
        args.reference_sku_column,
        args.reference_category_column,
    )

    enriched, stats, corrections = enrich_rows(rows, recommender, manufacturer_data, args.top_n, category_reference)
    output_fieldnames = with_category_field(fieldnames)
    write_csv_rows(args.output, enriched, output_fieldnames, args.delimiter)
    corrections_report = args.corrections_report or str(Path("reports") / f"{Path(args.output).stem}_category_corrections.csv")
    write_corrections_report(corrections_report, corrections)

    print(f"OK: wczytano {len(rows)} produktow")
    print(f"OK: dopisano kategorie: {stats['categories_added']}")
    print(f"OK: poprawiono kategorie regulami: {stats['categories_corrected']}")
    print(f"OK: dopisano dane producenta: {stats['manufacturer_data_added']}")
    print(f"OK: zapisano {args.output}")
    print(f"OK: raport korekt {corrections_report}")


def read_csv_rows(path: str | Path, delimiter: str) -> tuple[list[dict[str, str]], list[str]]:
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter=delimiter)
        rows = [{str(key): str(value or "") for key, value in row.items()} for row in reader]
        return rows, list(reader.fieldnames or [])


def write_csv_rows(path: str | Path, rows: list[dict[str, str]], fieldnames: list[str], delimiter: str) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(rows)


def write_corrections_report(path: str | Path, rows: list[dict[str, str]]) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ["sku", "name", "old_category", "new_category", "reason"]
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def with_category_field(fieldnames: list[str]) -> list[str]:
    if BASELINKER_CATEGORY_COLUMN in fieldnames:
        return fieldnames
    result: list[str] = []
    inserted = False
    for field in fieldnames:
        result.append(field)
        if field == "manufacturer_name":
            result.append(BASELINKER_CATEGORY_COLUMN)
            inserted = True
    if not inserted:
        result.append(BASELINKER_CATEGORY_COLUMN)
    return result


def enrich_rows(
    rows: list[dict[str, str]],
    recommender: CategoryRecommender,
    manufacturer_data_by_producer: dict[str, str],
    top_n: int,
    category_reference: dict[str, str] | None = None,
) -> tuple[list[dict[str, str]], dict[str, int], list[dict[str, str]]]:
    recommendation_df = build_recommendation_dataframe(rows)
    columns = detect_recommendation_columns(recommendation_df)
    recommendations = [recommender.recommend(row, columns, top_n) for _, row in recommendation_df.iterrows()]

    result: list[dict[str, str]] = []
    categories_added = 0
    categories_corrected = 0
    manufacturer_data_added = 0
    corrections: list[dict[str, str]] = []

    for row, row_recommendations in zip(rows, recommendations):
        enriched = dict(row)
        if not compact_spaces(enriched.get(BASELINKER_CATEGORY_COLUMN, "")):
            reference_category = (category_reference or {}).get(normalize_sku(enriched.get("sku", "")), "")
            recommended_category = row_recommendations[0]["category"] if row_recommendations else ""
            category = reference_category or recommended_category
            if category:
                enriched[BASELINKER_CATEGORY_COLUMN] = category
                categories_added += 1

        features = parse_features(enriched.get("features", ""))
        if features is not None:
            old_category = compact_spaces(enriched.get(BASELINKER_CATEGORY_COLUMN, ""))
            corrected_category, correction_reason = corrected_category_for_kanlux(enriched, features, old_category)
            if corrected_category and corrected_category != old_category:
                enriched[BASELINKER_CATEGORY_COLUMN] = corrected_category
                categories_corrected += 1
                corrections.append(
                    {
                        "sku": enriched.get("sku", ""),
                        "name": enriched.get("name", ""),
                        "old_category": old_category,
                        "new_category": corrected_category,
                        "reason": correction_reason,
                    }
                )

            producer = (
                compact_spaces(enriched.get("manufacturer_name", ""))
                or compact_spaces(str(features.get("Producent", "")))
                or compact_spaces(str(features.get("Marka", "")))
            )
            if not compact_spaces(str(features.get(DEFAULT_MANUFACTURER_DATA_ATTRIBUTE, ""))):
                data = manufacturer_data_by_producer.get(normalize_producer_key(producer))
                if data:
                    features[DEFAULT_MANUFACTURER_DATA_ATTRIBUTE] = data
                    enriched["features"] = json.dumps(features, ensure_ascii=False, separators=(",", ":"))
                    manufacturer_data_added += 1

        result.append(enriched)

    return result, {
        "categories_added": categories_added,
        "categories_corrected": categories_corrected,
        "manufacturer_data_added": manufacturer_data_added,
    }, corrections


def corrected_category_for_kanlux(row: dict[str, str], features: dict[str, Any], current_category: str) -> tuple[str, str]:
    name = compact_spaces(row.get("name", ""))
    feature_text = " ".join(compact_spaces(str(features.get(key, ""))) for key in ["Typ produktu", "Seria", "Model"])
    haystack = f"{name} {feature_text}".lower()

    if "lampa wisząca" in haystack or "lampa wiszaca" in haystack:
        return CAT_PENDANT, "lampa_wiszaca"

    if "oprawa kanałowa" in haystack or "oprawa kanalowa" in haystack:
        return CAT_SURFACE, "oprawa_kanalowa"

    surface_series = {"aqilo", "bord", "blurro", "gord", "phlox", "riti", "sani", "stobi", "toleo"}
    series = compact_spaces(str(features.get("Seria", ""))).lower()
    is_surface_point = (
        "oprawa sufitowa natynkowa" in haystack
        or ("oprawa sufitowa" in haystack and "gu10" in haystack)
        or ("oprawa sufitowa" in haystack and series in surface_series)
    )
    if is_surface_point:
        return CAT_SURFACE, "oprawa_sufitowa_natynkowa_lub_gu10"

    return current_category, ""


def build_recommendation_dataframe(rows: list[dict[str, str]]) -> pd.DataFrame:
    normalized_rows: list[dict[str, Any]] = []
    for row in rows:
        features = parse_features(row.get("features", "")) or {}
        normalized_rows.append(
            {
                **features,
                "name": row.get("name", ""),
                "sku": row.get("sku", ""),
                "manufacturer_name": row.get("manufacturer_name", ""),
                "category": row.get(BASELINKER_CATEGORY_COLUMN, ""),
            }
        )
    return pd.DataFrame(normalized_rows)


def parse_features(value: str) -> dict[str, Any] | None:
    text = compact_spaces(value)
    if not text:
        return {}
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return None
    return data if isinstance(data, dict) else {}


def load_category_reference(
    path: str,
    sheet_name: str | int | None,
    sku_column: str,
    category_column: str,
) -> dict[str, str]:
    if not path:
        return {}
    df = read_products(path, sheet_name=sheet_name)
    if sku_column not in df.columns or category_column not in df.columns:
        raise SystemExit(f"Brak kolumn referencyjnych: {sku_column}, {category_column}")

    result: dict[str, str] = {}
    for _, row in df.iterrows():
        sku = normalize_sku(row.get(sku_column, ""))
        category = compact_spaces(str(row.get(category_column, "")))
        if sku and category:
            result[sku] = category
    return result


def normalize_sku(value: Any) -> str:
    sku = compact_spaces(str(value or ""))
    if "/" in sku:
        sku = sku.rsplit("/", 1)[0]
    if sku.endswith(".0"):
        sku = sku[:-2]
    return sku.upper()


def normalize_producer_key(value: str) -> str:
    from utils import normalize_header

    return normalize_header(value)


if __name__ == "__main__":
    main()
