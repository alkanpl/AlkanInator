"""Kandydaci slow kluczowych per typ produktu i ich wolumeny z Keyword Plannera.

Czyta produkty (XLSX/CSV), przypisuje kazdy do typu po poczatku tytulu wedlug
slownika kandydatow (np. `dictionaries/incobex_keyword_candidates.yaml`) i
zapisuje do katalogu raportow:

- `keyword_candidates.xlsx` - arkusze "Typy produktow", "Kandydaci", "Produkty",
  "Nieprzypisane" (+ "Planner poza lista" przy `--planner-stats`),
- `keyword_candidates.csv` - kandydaci (utf-8-sig, dla Excela),
- `planner_seed_list.csv` - unikalne frazy w formacie szablonu Keyword Plannera
  (jedna kolumna `Keyword`, utf-8 BEZ BOM - BOM psuje naglowek przy wgrywaniu),
  do wgrania w "Sprawdz liczbe wyszukiwan i prognozy".

Krok 2 (`--planner-stats`, mozna podac kilka plikow): eksporty "Plan historical
metrics" i "Keyword ideas" z Plannera (UTF-16, tabulatory, 2 linie tytulu nad
naglowkiem) dopisuja do kandydatow
wolumen, konkurencje i stawki, a w "Typy produktow" wskazuje fraze glowna
(najwyzszy wolumen, bez fraz marki/kodu i fraz oznaczonych jako `ambiguous`).
Frazy z Plannera spoza listy kandydatow ida do arkusza "Planner poza lista" -
te przejrzane i odrzucone (`rejected_ideas` w YAML) z powodem, pozostale z
pustym powodem do przejrzenia.
Konto bez aktywnych kampanii dostaje wolumeny zaokraglone do przedzialow
(50 = 10-100, 500 = 100-1 tys., 5000 = 1-10 tys.) - kolumna `searches_range`.

Pliki wejsciowe nie sa modyfikowane.
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
PRIMARY_GROUPS = ["katalogowa", "synonimy", "przeznaczenie"]
PLANNER_KEYWORD_HEADER = "Keyword"
SECONDARY_LIMIT = 5
CANDIDATE_COLUMNS = ["product_type", "group", "keyword"]
STATS_COLUMNS = [
    "avg_monthly_searches",
    "searches_range",
    "competition",
    "competition_index",
    "bid_low_pln",
    "bid_high_pln",
    "yoy_change",
    "status",
    "note",
]
# Wartosci, ktorymi Planner zastepuje przedzialy na kontach bez aktywnych kampanii.
BUCKET_RANGES = {50: "10-100", 500: "100-1 tys.", 5000: "1-10 tys.", 50000: "10-100 tys.", 500000: "100 tys.-1 mln"}


@dataclass
class ProductType:
    key: str
    name: str
    prefixes: list[str]
    keywords: dict[str, list[str]]
    ambiguous: dict[str, str] = field(default_factory=dict)
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
        ambiguous = {normalize_header(k): str(v) for k, v in (spec.get("ambiguous") or {}).items()}
        types.append(ProductType(key, str(spec.get("name") or key), [p for p in prefixes if p], keywords, ambiguous))
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


def candidate_rows(types: list[ProductType]) -> list[dict[str, Any]]:
    """Kandydaci per typ; fraza powtorzona w innym typie zostaje przy pierwszym."""
    rows: list[dict[str, Any]] = []
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


def _parse_number(value: str) -> float | None:
    text = compact_spaces(str(value or "")).replace(" ", "").replace(" ", "").replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def load_planner_stats(path: str | Path) -> dict[str, dict[str, Any]]:
    """Eksport Keyword Plannera -> {znormalizowana fraza: metryki}.

    Obsluguje natywny eksport (UTF-16 + tabulatory) i CSV zapisany w UTF-8.
    Naglowek jest szukany po pierwszej komorce `Keyword` - linie tytulu i
    wiersze podsumowania (pusta fraza) sa pomijane.
    """
    raw = Path(path).read_bytes()
    encoding = "utf-16" if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
    lines = raw.decode(encoding).splitlines()
    delimiter = "\t" if any("\t" in line for line in lines[:5]) else ","
    rows = list(csv.reader(lines, delimiter=delimiter))
    header_index = next((i for i, row in enumerate(rows) if row and row[0].strip() == PLANNER_KEYWORD_HEADER), None)
    if header_index is None:
        raise SystemExit(f"{path}: brak naglowka 'Keyword' - to nie jest eksport 'Plan historical metrics'.")
    header = [cell.strip() for cell in rows[header_index]]
    stats: dict[str, dict[str, Any]] = {}
    for row in rows[header_index + 1 :]:
        record = dict(zip(header, row))
        keyword = compact_spaces(record.get("Keyword", ""))
        if not keyword:
            continue
        searches = _parse_number(record.get("Avg. monthly searches", ""))
        stats[normalize_header(keyword)] = {
            "planner_keyword": keyword,
            "avg_monthly_searches": int(searches) if searches is not None else None,
            "competition": record.get("Competition", "").strip(),
            "competition_index": _parse_number(record.get("Competition (indexed value)", "")),
            "bid_low_pln": _parse_number(record.get("Top of page bid (low range)", "")),
            "bid_high_pln": _parse_number(record.get("Top of page bid (high range)", "")),
            "yoy_change": record.get("YoY change", "").strip(),
        }
    return stats


def merge_planner_stats(paths: list[str | Path]) -> dict[str, dict[str, Any]]:
    """Laczy kilka eksportow; fraza z wolumenem wygrywa z pusta z innego pliku."""
    merged: dict[str, dict[str, Any]] = {}
    for path in paths:
        for key, metric in load_planner_stats(path).items():
            current = merged.get(key)
            if current is None or (current["avg_monthly_searches"] is None and metric["avg_monthly_searches"] is not None):
                merged[key] = metric
    return merged


def attach_planner_stats(
    candidates: list[dict[str, Any]],
    types: list[ProductType],
    stats: dict[str, dict[str, Any]],
    rejected: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """Dopisuje metryki do kandydatow; zwraca frazy z Plannera spoza listy kandydatow.

    Frazy z `rejected` (znormalizowana fraza -> powod) dostaja powod odrzucenia,
    pozostale maja pusty powod, czyli czekaja na przejrzenie.
    """
    rejected = rejected or {}
    ambiguous_by_type = {product_type.name: product_type.ambiguous for product_type in types}
    used: set[str] = set()
    for row in candidates:
        key = normalize_header(row["keyword"])
        metric = stats.get(key)
        note = ambiguous_by_type.get(row["product_type"], {}).get(key, "")
        searches = metric["avg_monthly_searches"] if metric else None
        row.update(
            {
                "avg_monthly_searches": searches if searches is not None else "",
                "searches_range": BUCKET_RANGES.get(searches, "") if searches else "",
                "competition": metric["competition"] if metric else "",
                "competition_index": metric["competition_index"] if metric and metric["competition_index"] is not None else "",
                "bid_low_pln": metric["bid_low_pln"] if metric and metric["bid_low_pln"] is not None else "",
                "bid_high_pln": metric["bid_high_pln"] if metric and metric["bid_high_pln"] is not None else "",
                "yoy_change": metric["yoy_change"] if metric else "",
                "note": note,
            }
        )
        if metric is None:
            row["status"] = "BRAK_W_EKSPORCIE"
        elif note:
            row["status"] = "NIEJEDNOZNACZNA"
        elif not searches:
            row["status"] = "BRAK_WOLUMENU"
        else:
            row["status"] = "OK"
        if metric:
            used.add(key)
    extra = [
        {
            "keyword": metric["planner_keyword"],
            "avg_monthly_searches": metric["avg_monthly_searches"] or "",
            "searches_range": BUCKET_RANGES.get(metric["avg_monthly_searches"] or 0, ""),
            "rejected_reason": rejected.get(key, ""),
        }
        for key, metric in stats.items()
        if key not in used
    ]
    # najpierw do przejrzenia, potem odrzucone; w grupie malejaco po wolumenie
    return sorted(extra, key=lambda row: (bool(row["rejected_reason"]), -(row["avg_monthly_searches"] or 0)))


def load_rejected_ideas(config: dict[str, Any]) -> dict[str, str]:
    return {normalize_header(k): str(v) for k, v in (config.get("rejected_ideas") or {}).items()}


def _format_keyword(row: dict[str, Any]) -> str:
    label = row.get("searches_range") or row.get("avg_monthly_searches")
    return f"{row['keyword']} ({label})" if label != "" else row["keyword"]


def type_summary_rows(types: list[ProductType], candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    with_stats = bool(candidates) and "status" in candidates[0]
    rows = []
    for product_type in sorted(types, key=lambda item: -len(item.products)):
        summary: dict[str, Any] = {
            "product_type": product_type.name,
            "products": len(product_type.products),
            "keyword_candidates": sum(1 for row in candidates if row["product_type"] == product_type.name),
        }
        if with_stats:
            own = [row for row in candidates if row["product_type"] == product_type.name]
            ranked = sorted(
                (row for row in own if row["status"] == "OK" and row["group"] in PRIMARY_GROUPS),
                key=lambda row: -int(row["avg_monthly_searches"]),
            )  # sortowanie stabilne: przy remisie wygrywa kolejnosc grup (katalogowa pierwsza)
            summary.update(
                {
                    "primary_keyword": ranked[0]["keyword"] if ranked else "",
                    "primary_searches_range": (ranked[0]["searches_range"] or ranked[0]["avg_monthly_searches"]) if ranked else "",
                    "secondary_keywords": " | ".join(_format_keyword(row) for row in ranked[1 : 1 + SECONDARY_LIMIT]),
                    "with_volume": sum(1 for row in own if row["status"] == "OK"),
                    "ambiguous": " | ".join(_format_keyword(row) for row in own if row["status"] == "NIEJEDNOZNACZNA"),
                }
            )
        summary["example_titles"] = " | ".join(p["Title"] for p in product_type.products[:3])
        rows.append(summary)
    return rows


def write_reports(
    reports_dir: Path,
    types: list[ProductType],
    candidates: list[dict[str, Any]],
    unassigned: list[dict[str, str]],
    planner_extra: list[dict[str, Any]] | None = None,
) -> dict[str, Path]:
    ensure_dir(reports_dir)
    paths = {
        "xlsx": reports_dir / "keyword_candidates.xlsx",
        "csv": reports_dir / "keyword_candidates.csv",
        "seed": reports_dir / "planner_seed_list.csv",
    }
    columns = CANDIDATE_COLUMNS + (STATS_COLUMNS if candidates and "status" in candidates[0] else [])
    products = [
        {"product_type": product_type.name, **product}
        for product_type in types
        for product in product_type.products
    ]
    with pd.ExcelWriter(paths["xlsx"]) as writer:
        pd.DataFrame(type_summary_rows(types, candidates)).to_excel(writer, sheet_name="Typy produktow", index=False)
        pd.DataFrame(candidates, columns=columns).to_excel(writer, sheet_name="Kandydaci", index=False)
        pd.DataFrame(products, columns=["product_type", "SKU", "Title"]).to_excel(
            writer, sheet_name="Produkty", index=False
        )
        pd.DataFrame(unassigned, columns=["SKU", "Title"]).to_excel(writer, sheet_name="Nieprzypisane", index=False)
        if planner_extra is not None:
            pd.DataFrame(
                planner_extra, columns=["keyword", "avg_monthly_searches", "searches_range", "rejected_reason"]
            ).to_excel(
                writer, sheet_name="Planner poza lista", index=False
            )
    with paths["csv"].open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(candidates)
    with paths["seed"].open("w", encoding="utf-8", newline="") as handle:
        seed_writer = csv.writer(handle, lineterminator="\n")
        seed_writer.writerow([PLANNER_KEYWORD_HEADER])
        seed_writer.writerows([row["keyword"]] for row in candidates)
    return paths


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", required=True, help="Plik XLSX/CSV z produktami (kolumna Title).")
    parser.add_argument("--sheet", help="Arkusz XLSX.")
    parser.add_argument("--dictionary", required=True, help="YAML z typami i kandydatami fraz.")
    parser.add_argument("--reports-dir", required=True, help="Katalog raportow.")
    parser.add_argument(
        "--planner-stats", nargs="+", help="Eksporty z Keyword Plannera (Plan historical metrics / Keyword ideas)."
    )
    args = parser.parse_args()

    config = load_yaml(args.dictionary)
    types = load_product_types(config)
    unassigned = assign_products(read_products(args.input, args.sheet), types)
    candidates = candidate_rows(types)
    planner_extra = None
    if args.planner_stats:
        planner_extra = attach_planner_stats(
            candidates, types, merge_planner_stats(args.planner_stats), load_rejected_ideas(config)
        )
    paths = write_reports(Path(args.reports_dir), types, candidates, unassigned, planner_extra)

    assigned = sum(len(product_type.products) for product_type in types)
    print(f"Produkty przypisane: {assigned}, nieprzypisane: {len(unassigned)}")
    print(f"Typy z produktami: {sum(1 for t in types if t.products)}/{len(types)}, frazy do Plannera: {len(candidates)}")
    if planner_extra is not None:
        statuses: dict[str, int] = {}
        for row in candidates:
            statuses[row["status"]] = statuses.get(row["status"], 0) + 1
        print("Statusy fraz: " + ", ".join(f"{k}={v}" for k, v in sorted(statuses.items())))
        to_review = sum(1 for row in planner_extra if not row["rejected_reason"])
        print(f"Frazy z Plannera spoza listy: {len(planner_extra)} (do przejrzenia: {to_review})")
    for path in paths.values():
        print(f"Zapisano: {path}")


if __name__ == "__main__":
    main()
