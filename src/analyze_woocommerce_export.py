from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

from utils import compact_spaces, ensure_dir, read_products, slugify, valid_ean


ATTRIBUTE_NAME_PREFIX = "Nazwa atrybutu "
ATTRIBUTE_VALUE_PREFIX = "Wartości atrybutu "
ATTRIBUTE_VISIBLE_SUFFIX = " widoczny"
ATTRIBUTE_GLOBAL_SUFFIX = ": globalny"

CORE_COLUMNS = [
    "ID",
    "Rodzaj",
    "SKU",
    "GTIN, UPC, EAN lub ISBN",
    "Nazwa",
    "Opublikowano",
    "Widoczność w katalogu",
    "Krótki opis",
    "Opis",
    "W magazynie?",
    "Stan magazynowy",
    "Cena",
    "Kategorie",
    "Tagi",
    "Obrazki",
    "Marki",
]

IMPORTANT_ATTRIBUTES = {
    "Producent",
    "Dane producenta",
    "Kolor",
    "Kolor producenta",
    "Seria",
    "Typ produktu",
    "Podtyp produktu",
    "Rodzaj",
    "Stopień ochrony [IP]",
    "Sposób montażu",
    "Napięcie [V]",
    "Prąd znamionowy [A]",
    "Moc [W]",
    "Temperatura barwowa [K]",
    "Strumień świetlny [lm]",
    "Barwa światła",
    "Trzonek",
    "Materiał",
    "Uziemienie",
    "Krotność",
}

EAN_COLUMNS = [
    "GTIN, UPC, EAN lub ISBN",
    "Pola: hwp_product_gtin",
    "Pola: hwp_var_gtin",
    "Pola: _gees_ean",
    "Pola: _wpm_gtin_code",
]


def split_list(value: Any, separators: str = r",|\|") -> list[str]:
    parts = re.split(separators, str(value or ""))
    return [compact_spaces(part) for part in parts if compact_spaces(part)]


def split_attribute_values(value: Any) -> list[str]:
    text = str(value or "")
    if not compact_spaces(text):
        return []
    return split_list(text, separators=r",|\||>")


def split_category_paths(value: Any) -> list[str]:
    return split_list(value, separators=r",|\|")


def category_parts(path: str) -> list[str]:
    return [compact_spaces(part) for part in str(path or "").split(">") if compact_spaces(part)]


def category_leaf(path: str) -> str:
    parts = category_parts(path)
    return parts[-1] if parts else ""


def find_attribute_indexes(columns: list[str]) -> list[int]:
    indexes: list[int] = []
    for column in columns:
        match = re.fullmatch(r"Nazwa atrybutu (\d+)", column)
        if match:
            indexes.append(int(match.group(1)))
    return sorted(indexes)


def row_attributes(row: pd.Series, indexes: list[int]) -> list[dict[str, Any]]:
    attributes: list[dict[str, Any]] = []
    for index in indexes:
        name = compact_spaces(row.get(f"{ATTRIBUTE_NAME_PREFIX}{index}", ""))
        raw_value = compact_spaces(row.get(f"{ATTRIBUTE_VALUE_PREFIX}{index}", ""))
        if not name or not raw_value:
            continue
        attributes.append(
            {
                "index": index,
                "name": name,
                "raw_value": raw_value,
                "values": split_attribute_values(raw_value),
                "visible": compact_spaces(row.get(f"Atrybut {index}{ATTRIBUTE_VISIBLE_SUFFIX}", "")),
                "global": compact_spaces(row.get(f"Atrybut {index}{ATTRIBUTE_GLOBAL_SUFFIX}", "")),
            }
        )
    return attributes


def value_for_attribute(attributes: list[dict[str, Any]], name: str) -> str:
    for attribute in attributes:
        if attribute["name"] == name:
            return attribute["raw_value"]
    return ""


def first_row_value(row: pd.Series, columns: list[str]) -> str:
    for column in columns:
        value = compact_spaces(row.get(column, ""))
        if value:
            return value
    return ""


def first_ean(row: pd.Series) -> str:
    raw_value = first_row_value(row, EAN_COLUMNS)
    for value in split_list(raw_value, separators=r"\||,|;"):
        digits = re.sub(r"\D", "", value)
        if digits:
            return digits
    return ""


def text_word_counts(values: list[str], limit: int = 250) -> list[dict[str, Any]]:
    stop_words = {
        "oraz",
        "dla",
        "jest",
        "przy",
        "nad",
        "pod",
        "bez",
        "z",
        "w",
        "i",
        "na",
        "do",
        "od",
        "po",
        "ze",
        "się",
        "sie",
        "ten",
        "ta",
        "to",
    }
    counter: Counter[str] = Counter()
    for value in values:
        for token in re.findall(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż0-9+-]{3,}", str(value).lower()):
            if token not in stop_words:
                counter[token] += 1
    return [{"word": word, "count": count} for word, count in counter.most_common(limit)]


def counter_rows(counter: Counter[str], key_name: str, limit: int | None = None) -> list[dict[str, Any]]:
    items = counter.most_common(limit)
    return [{key_name: key, "count": count} for key, count in items]


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str] | None = None) -> None:
    ensure_dir(path.parent)
    if not rows:
        names = fieldnames or []
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=names)
            if names:
                writer.writeheader()
        return
    names = fieldnames or list(rows[0].keys())
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=names)
        writer.writeheader()
        writer.writerows(rows)


def build_analysis(df: pd.DataFrame) -> dict[str, Any]:
    attribute_indexes = find_attribute_indexes(list(df.columns))

    product_types: Counter[str] = Counter()
    visibility: Counter[str] = Counter()
    stock_status: Counter[str] = Counter()
    categories: Counter[str] = Counter()
    category_leaves: Counter[str] = Counter()
    category_depths: Counter[str] = Counter()
    category_edges: Counter[tuple[str, str]] = Counter()
    brands: Counter[str] = Counter()
    tags: Counter[str] = Counter()
    attribute_names: Counter[str] = Counter()
    attribute_values: dict[str, Counter[str]] = defaultdict(Counter)
    attributes_by_category: dict[str, Counter[str]] = defaultdict(Counter)
    attribute_values_by_category: dict[tuple[str, str], Counter[str]] = defaultdict(Counter)
    producers: Counter[str] = Counter()
    series: Counter[str] = Counter()
    types: Counter[str] = Counter()
    subtypes: Counter[str] = Counter()
    colors: Counter[str] = Counter()
    ip_values: Counter[str] = Counter()
    title_words: list[str] = []
    description_words: list[str] = []
    rows_without_category: list[dict[str, str]] = []
    rows_without_sku: list[dict[str, str]] = []
    rows_without_ean: list[dict[str, str]] = []
    invalid_eans: list[dict[str, str]] = []
    duplicate_skus: Counter[str] = Counter()
    duplicate_eans: Counter[str] = Counter()
    column_nonempty: Counter[str] = Counter()
    product_examples_by_category: dict[str, list[dict[str, str]]] = defaultdict(list)

    sku_counter: Counter[str] = Counter(compact_spaces(value) for value in df.get("SKU", []) if compact_spaces(value))
    ean_values = []
    for _, row in df.iterrows():
        ean = first_ean(row)
        if ean:
            ean_values.append(ean)
    ean_counter: Counter[str] = Counter(ean_values)

    for sku, count in sku_counter.items():
        if count > 1:
            duplicate_skus[sku] = count
    for ean, count in ean_counter.items():
        if count > 1:
            duplicate_eans[ean] = count

    for _, row in df.iterrows():
        row_id = compact_spaces(row.get("ID", ""))
        sku = compact_spaces(row.get("SKU", ""))
        title = compact_spaces(row.get("Nazwa", ""))
        ean = first_ean(row)
        row_stub = {"id": row_id, "sku": sku, "title": title}

        for column, value in row.items():
            if compact_spaces(value):
                column_nonempty[str(column)] += 1

        product_types[compact_spaces(row.get("Rodzaj", ""))] += 1
        visibility[compact_spaces(row.get("Widoczność w katalogu", ""))] += 1
        stock_status[compact_spaces(row.get("W magazynie?", ""))] += 1
        title_words.append(title)
        description_words.append(compact_spaces(row.get("Opis", "")))

        if not sku:
            rows_without_sku.append(row_stub)
        if not ean:
            rows_without_ean.append(row_stub)
        elif not valid_ean(ean):
            invalid_eans.append({**row_stub, "ean": ean})

        paths = split_category_paths(row.get("Kategorie", ""))
        if not paths:
            rows_without_category.append(row_stub)
        for path in paths:
            parts = category_parts(path)
            categories[path] += 1
            category_depths[str(len(parts))] += 1
            leaf = parts[-1] if parts else ""
            if leaf:
                category_leaves[leaf] += 1
                if len(product_examples_by_category[path]) < 5:
                    product_examples_by_category[path].append(row_stub)
            for parent, child in zip(parts, parts[1:]):
                category_edges[(parent, child)] += 1

        for brand in split_list(row.get("Marki", "")):
            brands[brand] += 1
        for tag in split_list(row.get("Tagi", "")):
            tags[tag] += 1

        attributes = row_attributes(row, attribute_indexes)
        producer = value_for_attribute(attributes, "Producent") or compact_spaces(row.get("Marki", ""))
        if producer:
            producers[producer] += 1
        for value, counter_name in [
            (value_for_attribute(attributes, "Seria"), series),
            (value_for_attribute(attributes, "Typ produktu"), types),
            (value_for_attribute(attributes, "Podtyp produktu"), subtypes),
            (value_for_attribute(attributes, "Kolor"), colors),
            (value_for_attribute(attributes, "Stopień ochrony [IP]"), ip_values),
        ]:
            for item in split_attribute_values(value):
                counter_name[item] += 1

        for attribute in attributes:
            name = attribute["name"]
            attribute_names[name] += 1
            values = attribute["values"] or [attribute["raw_value"]]
            for value in values:
                attribute_values[name][value] += 1
            for path in paths:
                attributes_by_category[path][name] += 1
                for value in values:
                    attribute_values_by_category[(path, name)][value] += 1

    attr_rows: list[dict[str, Any]] = []
    for name, count in attribute_names.most_common():
        values_counter = attribute_values[name]
        attr_rows.append(
            {
                "attribute": name,
                "products_count": count,
                "unique_values": len(values_counter),
                "top_values": "; ".join(f"{value} ({value_count})" for value, value_count in values_counter.most_common(12)),
            }
        )

    attr_value_rows = [
        {"attribute": name, "value": value, "count": count}
        for name, counter in sorted(attribute_values.items())
        for value, count in counter.most_common(250)
    ]

    attr_by_category_rows = [
        {
            "category": category,
            "attribute": name,
            "products_count": count,
            "coverage_percent": round(count / categories[category] * 100, 2) if categories[category] else 0,
        }
        for category, counter in sorted(attributes_by_category.items())
        for name, count in counter.most_common()
    ]

    top_attr_value_by_category_rows = [
        {
            "category": category,
            "attribute": name,
            "value": value,
            "count": count,
        }
        for (category, name), counter in sorted(attribute_values_by_category.items())
        if name in IMPORTANT_ATTRIBUTES
        for value, count in counter.most_common(30)
    ]

    category_tree_rows = [
        {"parent": parent, "child": child, "count": count}
        for (parent, child), count in category_edges.most_common()
    ]

    quality_summary = {
        "without_sku": len(rows_without_sku),
        "without_ean": len(rows_without_ean),
        "invalid_ean": len(invalid_eans),
        "without_category": len(rows_without_category),
        "duplicate_sku_values": len(duplicate_skus),
        "duplicate_ean_values": len(duplicate_eans),
    }

    return {
        "summary": {
            "rows": int(len(df)),
            "columns": int(len(df.columns)),
            "attribute_slots": len(attribute_indexes),
            "attribute_names": len(attribute_names),
            "category_paths": len(categories),
            "category_leaves": len(category_leaves),
            "brands": len(brands),
            "tags": len(tags),
            **quality_summary,
        },
        "product_types": counter_rows(product_types, "type"),
        "visibility": counter_rows(visibility, "visibility"),
        "stock_status": counter_rows(stock_status, "in_stock"),
        "categories": counter_rows(categories, "category"),
        "category_leaves": counter_rows(category_leaves, "category_leaf"),
        "category_depths": counter_rows(category_depths, "depth"),
        "category_tree": category_tree_rows,
        "brands": counter_rows(brands, "brand"),
        "tags": counter_rows(tags, "tag"),
        "attributes": attr_rows,
        "attribute_values": attr_value_rows,
        "attributes_by_category": attr_by_category_rows,
        "top_attribute_values_by_category": top_attr_value_by_category_rows,
        "producers": counter_rows(producers, "producer"),
        "series": counter_rows(series, "series"),
        "types": counter_rows(types, "type"),
        "subtypes": counter_rows(subtypes, "subtype"),
        "colors": counter_rows(colors, "color"),
        "ip_values": counter_rows(ip_values, "ip"),
        "title_words": text_word_counts(title_words),
        "description_words": text_word_counts(description_words),
        "quality": {
            "summary": quality_summary,
            "without_sku": rows_without_sku[:500],
            "without_ean": rows_without_ean[:500],
            "invalid_ean": invalid_eans[:500],
            "without_category": rows_without_category[:500],
            "duplicate_skus": counter_rows(duplicate_skus, "sku"),
            "duplicate_eans": counter_rows(duplicate_eans, "ean"),
        },
        "columns_profile": [
            {
                "column": column,
                "nonempty": column_nonempty[column],
                "empty": int(len(df) - column_nonempty[column]),
                "coverage_percent": round(column_nonempty[column] / len(df) * 100, 2) if len(df) else 0,
                "kind": "core"
                if column in CORE_COLUMNS
                else "attribute"
                if column.startswith(("Nazwa atrybutu ", "Wartości atrybutu ", "Atrybut "))
                else "meta"
                if column.startswith("Pola: ")
                else "other",
            }
            for column in df.columns
        ],
        "examples_by_category": dict(product_examples_by_category),
    }


def write_html_summary(path: Path, analysis: dict[str, Any]) -> None:
    summary = analysis["summary"]
    top_categories = "\n".join(
        f"<tr><td>{row['category']}</td><td>{row['count']}</td></tr>" for row in analysis["categories"][:25]
    )
    top_attributes = "\n".join(
        f"<tr><td>{row['attribute']}</td><td>{row['products_count']}</td><td>{row['unique_values']}</td><td>{row['top_values']}</td></tr>"
        for row in analysis["attributes"][:25]
    )
    quality = analysis["quality"]["summary"]
    html = f"""<!doctype html>
<html lang="pl">
<head>
  <meta charset="utf-8">
  <title>WooCommerce catalog analysis</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 32px; color: #1f2933; }}
    table {{ border-collapse: collapse; width: 100%; margin: 16px 0 28px; }}
    th, td {{ border: 1px solid #d9e2ec; padding: 8px; text-align: left; vertical-align: top; }}
    th {{ background: #f0f4f8; }}
    code {{ background: #f0f4f8; padding: 2px 4px; }}
  </style>
</head>
<body>
  <h1>Analiza katalogu WooCommerce</h1>
  <h2>Podsumowanie</h2>
  <ul>
    <li>Produkty: {summary['rows']}</li>
    <li>Kolumny: {summary['columns']}</li>
    <li>Unikalne nazwy atrybutow: {summary['attribute_names']}</li>
    <li>Sciezki kategorii: {summary['category_paths']}</li>
    <li>Lisci kategorii: {summary['category_leaves']}</li>
    <li>Marki: {summary['brands']}</li>
    <li>Tagi: {summary['tags']}</li>
  </ul>
  <h2>Jakosc danych</h2>
  <ul>
    <li>Bez SKU: {quality['without_sku']}</li>
    <li>Bez EAN: {quality['without_ean']}</li>
    <li>Niepoprawny EAN: {quality['invalid_ean']}</li>
    <li>Bez kategorii: {quality['without_category']}</li>
    <li>Powtorzone SKU: {quality['duplicate_sku_values']}</li>
    <li>Powtorzone EAN: {quality['duplicate_ean_values']}</li>
  </ul>
  <h2>Najwieksze kategorie</h2>
  <table><tr><th>Kategoria</th><th>Produkty</th></tr>{top_categories}</table>
  <h2>Najczestsze atrybuty</h2>
  <table><tr><th>Atrybut</th><th>Produkty</th><th>Unikalne wartosci</th><th>Najczestsze wartosci</th></tr>{top_attributes}</table>
</body>
</html>
"""
    ensure_dir(path.parent)
    path.write_text(html, encoding="utf-8")


def write_outputs(analysis: dict[str, Any], reports_dir: Path, dictionary_path: Path) -> None:
    ensure_dir(reports_dir)
    ensure_dir(dictionary_path.parent)

    (reports_dir / "summary.json").write_text(
        json.dumps(analysis["summary"], ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_html_summary(reports_dir / "summary.html", analysis)

    csv_outputs = {
        "categories.csv": analysis["categories"],
        "category_leaves.csv": analysis["category_leaves"],
        "category_depths.csv": analysis["category_depths"],
        "category_tree.csv": analysis["category_tree"],
        "brands.csv": analysis["brands"],
        "tags.csv": analysis["tags"],
        "attributes.csv": analysis["attributes"],
        "attribute_values.csv": analysis["attribute_values"],
        "attributes_by_category.csv": analysis["attributes_by_category"],
        "top_attribute_values_by_category.csv": analysis["top_attribute_values_by_category"],
        "producers.csv": analysis["producers"],
        "series.csv": analysis["series"],
        "types.csv": analysis["types"],
        "subtypes.csv": analysis["subtypes"],
        "colors.csv": analysis["colors"],
        "ip_values.csv": analysis["ip_values"],
        "title_words.csv": analysis["title_words"],
        "description_words.csv": analysis["description_words"],
        "columns_profile.csv": analysis["columns_profile"],
        "duplicate_skus.csv": analysis["quality"]["duplicate_skus"],
        "duplicate_eans.csv": analysis["quality"]["duplicate_eans"],
        "invalid_eans.csv": analysis["quality"]["invalid_ean"],
        "products_without_ean.csv": analysis["quality"]["without_ean"],
        "products_without_sku.csv": analysis["quality"]["without_sku"],
        "products_without_category.csv": analysis["quality"]["without_category"],
    }
    for filename, rows in csv_outputs.items():
        write_csv(reports_dir / filename, rows)

    dictionary = {
        "summary": analysis["summary"],
        "top_categories": analysis["categories"][:150],
        "category_tree": analysis["category_tree"],
        "top_attributes": analysis["attributes"],
        "top_attribute_values_by_category": analysis["top_attribute_values_by_category"],
        "producers": analysis["producers"],
        "series": analysis["series"][:500],
        "types": analysis["types"][:500],
        "subtypes": analysis["subtypes"][:500],
        "colors": analysis["colors"][:300],
        "ip_values": analysis["ip_values"],
    }
    dictionary_path.write_text(yaml.safe_dump(dictionary, allow_unicode=True, sort_keys=False), encoding="utf-8")

    category_examples_dir = ensure_dir(reports_dir / "category_examples")
    for category, examples in analysis["examples_by_category"].items():
        if not examples:
            continue
        filename = slugify(category)[:120] or "category"
        write_csv(category_examples_dir / f"{filename}.csv", examples)


def main() -> None:
    parser = argparse.ArgumentParser(description="Wyciaga wiedze z eksportu produktow WooCommerce CSV.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--reports-dir", default="reports/woocommerce_catalog")
    parser.add_argument("--dictionary", default="dictionaries/woocommerce_catalog_knowledge.yaml")
    args = parser.parse_args()

    df = read_products(args.input)
    analysis = build_analysis(df)
    write_outputs(analysis, Path(args.reports_dir), Path(args.dictionary))

    print(f"OK: produkty: {analysis['summary']['rows']}")
    print(f"OK: raporty: {args.reports_dir}")
    print(f"OK: slownik: {args.dictionary}")
    print(json.dumps(analysis["summary"], ensure_ascii=False))


if __name__ == "__main__":
    main()
