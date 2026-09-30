"""Lista kandydatow slow kluczowych per typ produktu - wejscie do Keyword Plannera.

Czyta produkty (XLSX/CSV), przypisuje kazdy do typu po poczatku tytulu wedlug
slownika kandydatow (np. `dictionaries/incobex_keyword_candidates.yaml`) i
zapisuje do katalogu raportow:

- `keyword_candidates.xlsx` - arkusze "Typy produktow", "Kandydaci",
  "Nieprzypisane",
- `keyword_candidates.csv` - kolumny product_type, group, keyword
  (utf-8-sig, dla Excela),
- `planner_seed_list.txt` - unikalne frazy, jedna na linie, do wklejenia w
  Plannerze ("Sprawdz liczbe wyszukiwan i prognozy").

Plik wejsciowy nie jest modyfikowany.
"""

from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, detect_column, ensure_dir, load_yaml, normalize_header, read_products


TITLE_COLUMN_CANDIDATES = ["Title", "Nazwa", "Nazwa produktu", "name", "title"]
SKU_COLUMN_CANDIDATES = ["SKU", "sku", "Kod", "Kod Producenta"]
KEYWORD_GROUPS = ["katalogowa", "synonimy", "przeznaczenie", "marka_kod"]
UNASSIGNED = "nieprzypisane"


@dataclass
class ProductType:
    key: str
    name: str
    prefixes: list[str]
    keywords: dict[str, list[str]]
    products: list[dict[str, str]] = field(default_factory=list)


def load_product_types(config: dict[str, Any]) -> list[ProductType]:
    types: list[ProductType] = []
    for key, spec in (config.get("product_types") or {}).items():
        spec = spec or {}
        prefixes = [normalize_header(value) for value in (spec.get("match") or {}).get("starts_with") or []]
        keywords = {
            group: [compact_spaces(str(value)).lower() for value in (spec.get("keywords") or {}).get(group) or []]
            for group in KEYWORD_GROUPS
        }
        types.append(ProductType(key, str(spec.get("name") or key), [p for p in prefixes if p], keywords))
    return types


def assign_type(title: str, types: list[ProductType]) -> ProductType | None:
    """Typ z najdluzszym prefiksem pasujacym do poczatku tytulu (bez polskich znakow)."""
    normalized = normalize_header(title)
    best: tuple[int, ProductType] | None = None
    for product_type in types:
        for prefix in product_type.prefixes:
            if (normalized == prefix or normalized.startswith(prefix + " ")) and (best is None or len(prefix) > best[0]):
                best = (len(prefix), product_type)
    return best[1] if best else None


def assign_products(df: pd.DataFrame, types: list[ProductType]) -> list[dict[str, str]]:
    title_column = detect_column(df, TITLE_COLUMN_CANDIDATES)
    if not title_column:
        raise SystemExit(f"Brak kolumny z tytulem. Szukano: {', '.join(TITLE_COLUMN_CANDIDATES)}")
    sku_column = detect_column(df, SKU_COLUMN_CANDIDATES)
    unassigned: list[dict[str, str]] = []
    for _, row in df.iterrows():
        title = compact_spaces(str(row[title_column]))
        if not title:
            continue
        product = {"SKU": str(row[sku_column]) if sku_column else "", "Title": title}
        product_type = assign_type(title, types)
        if product_type is None:
            unassigned.append(product)
        else:
            product_type.products.append(product)
    return unassigned


def candidate_rows(types: list[ProductType]) -> list[dict[str, str]]:
    """Kandydaci per typ; fraza powtorzona w innym typie zostaje przy pierwszym."""
    rows: list[dict[str, str]] = []
    seen: set[str] = set()
    for product_type in types:
        if not product_type.products:
            continue
        for group in KEYWORD_GROUPS:
            for keyword in product_type.keywords[group]:
                key = normalize_header(keyword)
                if not key or key in seen:
                    continue
                seen.add(key)
                rows.append({"product_type": product_type.name, "group": group, "keyword": keyword})
    return rows


def type_summary_rows(types: list[ProductType]) -> list[dict[str, Any]]:
    rows = []
    for product_type in sorted(types, key=lambda item: -len(item.products)):
        rows.append(
            {
                "product_type": product_type.name,
                "products": len(product_type.products),
                "keyword_candidates": sum(len(values) for values in product_type.keywords.values()),
                "example_titles": " | ".join(p["Title"] for p in product_type.products[:3]),
            }
        )
    return rows


def write_reports(
    reports_dir: Path,
    types: list[ProductType],
    candidates: list[dict[str, str]],
    unassigned: list[dict[str, str]],
) -> dict[str, Path]:
    ensure_dir(reports_dir)
    paths = {
        "xlsx": reports_dir / "keyword_candidates.xlsx",
        "csv": reports_dir / "keyword_candidates.csv",
        "seed": reports_dir / "planner_seed_list.txt",
    }
    products = [
        {"product_type": product_type.name, **product}
        for product_type in types
        for product in product_type.products
    ]
    with pd.ExcelWriter(paths["xlsx"]) as writer:
        pd.DataFrame(type_summary_rows(types)).to_excel(writer, sheet_name="Typy produktow", index=False)
        pd.DataFrame(candidates, columns=["product_type", "group", "keyword"]).to_excel(
            writer, sheet_name="Kandydaci", index=False
        )
        pd.DataFrame(products, columns=["product_type", "SKU", "Title"]).to_excel(
            writer, sheet_name="Produkty", index=False
        )
        pd.DataFrame(unassigned, columns=["SKU", "Title"]).to_excel(writer, sheet_name="Nieprzypisane", index=False)
    with paths["csv"].open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["product_type", "group", "keyword"])
        writer.writeheader()
        writer.writerows(candidates)
    paths["seed"].write_text("\n".join(row["keyword"] for row in candidates) + "\n", encoding="utf-8")
    return paths


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", required=True, help="Plik XLSX/CSV z produktami (kolumna Title).")
    parser.add_argument("--sheet", help="Arkusz XLSX.")
    parser.add_argument("--dictionary", required=True, help="YAML z typami i kandydatami fraz.")
    parser.add_argument("--reports-dir", required=True, help="Katalog raportow.")
    args = parser.parse_args()

    types = load_product_types(load_yaml(args.dictionary))
    unassigned = assign_products(read_products(args.input, args.sheet), types)
    candidates = candidate_rows(types)
    paths = write_reports(Path(args.reports_dir), types, candidates, unassigned)

    assigned = sum(len(product_type.products) for product_type in types)
    print(f"Produkty przypisane: {assigned}, nieprzypisane: {len(unassigned)}")
    print(f"Typy z produktami: {sum(1 for t in types if t.products)}/{len(types)}, frazy do Plannera: {len(candidates)}")
    for path in paths.values():
        print(f"Zapisano: {path}")


if __name__ == "__main__":
    main()
