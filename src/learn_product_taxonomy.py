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

from utils import compact_spaces, detect_column, ensure_dir, normalize_header, read_products


KNOWN_SUFFIX_PRODUCERS = {
    "KAN": "Kanlux",
    "KANL": "Kanlux",
    "KRL": "Karlik",
    "OSP": "Ospel",
    "KON": "Kontakt-Simon",
    "HAG": "Hager",
    "EDO": "Edo",
    "EPS": "Elektro-Plast",
    "FIF": "F&F",
    "GTV": "GTV",
    "LDV": "LEDVANCE",
    "LED": "LEDVANCE",
    "LV": "LEDVANCE",
    "BMK": "Bemko",
    "BEM": "Bemko",
    "SIM": "Kontakt-Simon",
    "PAW": "Pawbol",
    "POL": "Polmark",
    "ETI": "ETI",
    "NKT": "NKT",
    "WAG": "WAGO",
    "ORN": "Orno",
    "EAT": "Eaton",
    "SCH": "Schneider Electric",
    "LEG": "Legrand",
}

KNOWN_PRODUCER_NAMES = {
    "Kanlux",
    "Karlik",
    "Ospel",
    "Kontakt Simon",
    "Kontakt-Simon",
    "Hager",
    "Berker",
    "Edo",
    "Elektro-Plast",
    "F&F",
    "GTV",
    "LEDVANCE",
    "Bemko",
    "Pawbol",
    "Polmark",
    "ETI",
    "ELKO-BIS",
    "EPN",
    "EPO",
    "EPS",
    "FELO",
    "Helukabel",
    "NKT",
    "WAGO",
    "Orno",
    "Eaton",
    "Lange",
    "Schneider Electric",
    "Legrand",
    "Morek",
    "PCE",
    "Plastrol",
    "Rawlplug",
    "Simet",
    "Steinel",
    "V-TAC",
}

TYPE_STOP_WORDS = {
    "biały",
    "biała",
    "białe",
    "bialy",
    "biala",
    "czarny",
    "czarna",
    "czarne",
    "szary",
    "szara",
    "popielaty",
    "brąz",
    "braz",
    "antracyt",
    "chrom",
    "satyna",
    "ip20",
    "ip44",
    "ip54",
    "ip65",
    "ip66",
    "led",
    "gu10",
    "e27",
    "e14",
}


def detect_columns(df: pd.DataFrame) -> dict[str, str | None]:
    return {
        "title": detect_column(df, ["Title", "Nazwa", "Nazwa produktu", "name"]),
        "sku": detect_column(df, ["Sku", "SKU", "Kod", "Indeks", "PRODUCT_CODE", "TOWARIDX"]),
        "ean": detect_column(df, ["hwp_product_gtin", "EAN", "GTIN"]),
        "category": detect_column(df, ["Kategorie produktów", "Kategorie", "category", "categories"]),
        "producer": detect_column(df, ["PRODUCER", "Producent", "manufacturer_name", "manufacturer", "Marka"]),
        "status": detect_column(df, ["Status"]),
        "product_type": detect_column(df, ["Product Type", "Typ produktu"]),
    }


def split_categories(value: Any) -> list[str]:
    result: list[str] = []
    for path in str(value or "").split("|"):
        path = compact_spaces(path)
        if path:
            result.append(path)
    return result


def category_leaf(path: str) -> str:
    parts = [compact_spaces(part) for part in str(path).split(">") if compact_spaces(part)]
    return parts[-1] if parts else ""


def normalize_sku_suffix(value: Any) -> str:
    sku = compact_spaces(str(value or "")).upper()
    if "/" in sku:
        return re.sub(r"[^A-Z0-9]+", "", sku.rsplit("/", 1)[-1])
    match = re.match(r"([A-Z]+)\d+", sku)
    return match.group(1) if match else ""


def infer_producer(title: str, sku: str, categories: list[str], known_producers: set[str]) -> tuple[str, str]:
    suffix = normalize_sku_suffix(sku)
    if suffix in KNOWN_SUFFIX_PRODUCERS:
        return KNOWN_SUFFIX_PRODUCERS[suffix], f"sku_suffix:{suffix}"

    text_norm = normalize_header(" ".join([title, *categories]))
    for producer in sorted(known_producers, key=len, reverse=True):
        producer_norm = normalize_header(producer)
        if producer_norm and re.search(rf"\b{re.escape(producer_norm)}\b", text_norm):
            return producer, "title_known"
    return "", ""


def title_tokens(title: str) -> list[str]:
    return re.findall(r"[A-Za-zŁŚŻŹĆŃÓĘĄłśżźćńóęą0-9+/-]{2,}", compact_spaces(title))


def infer_type_from_title(title: str) -> str:
    tokens = title_tokens(title)
    useful: list[str] = []
    for token in tokens[:6]:
        normalized = normalize_header(token)
        if not normalized or normalized in TYPE_STOP_WORDS:
            continue
        if re.fullmatch(r"\d+[a-z]*", normalized):
            continue
        useful.append(token)
        if len(useful) >= 3:
            break
    if not useful:
        return ""
    if len(useful) >= 2 and normalize_header(useful[1]) in {"podwojne", "pojedyncze", "schodowy", "krzyzowy"}:
        return " ".join(useful[:2])
    return useful[0] if len(useful) == 1 else " ".join(useful[:2])


def normalize_explicit_producer(value: Any) -> str:
    producer = compact_spaces(str(value or ""))
    if not producer:
        return ""
    aliases = {
        "BEMKO": "Bemko",
        "ELKO-BIS": "ELKO-BIS",
        "EPN": "EPN",
        "EPO": "EPO",
        "EPS": "EPS",
        "EATON": "Eaton",
        "ETI": "ETI",
        "F&F": "F&F",
        "GTV": "GTV",
        "HAGER": "Hager",
        "HELUKABEL": "Helukabel",
        "KANLUX": "Kanlux",
        "LANGE ŁUKASZUK": "Lange",
        "LEDVANCE": "LEDVANCE",
        "MOREK": "Morek",
        "PCE": "PCE",
        "PLASTROL": "Plastrol",
        "RAWLPLUG": "Rawlplug",
        "SIMET": "Simet",
    }
    return aliases.get(producer.upper(), producer)


def build_taxonomy(df: pd.DataFrame, columns: dict[str, str | None]) -> dict[str, Any]:
    title_col = columns["title"]
    sku_col = columns["sku"]
    category_col = columns["category"]
    producer_col = columns.get("producer")
    if not title_col:
        raise SystemExit("Nie wykryto kolumny z nazwą produktu.")

    known_producers = set(KNOWN_SUFFIX_PRODUCERS.values()) | KNOWN_PRODUCER_NAMES
    producers: Counter[str] = Counter()
    producer_sources: Counter[str] = Counter()
    sku_suffixes: Counter[str] = Counter()
    categories: Counter[str] = Counter()
    category_leaves: Counter[str] = Counter()
    types: Counter[str] = Counter()
    type_by_category: dict[str, Counter[str]] = defaultdict(Counter)
    examples: dict[str, list[dict[str, str]]] = defaultdict(list)

    for _, row in df.iterrows():
        title = compact_spaces(str(row.get(title_col, "")))
        sku = compact_spaces(str(row.get(sku_col, ""))) if sku_col else ""
        suffix = normalize_sku_suffix(sku)
        if suffix:
            sku_suffixes[suffix] += 1
        paths = split_categories(row.get(category_col, "")) if category_col else []
        explicit_producer = normalize_explicit_producer(row.get(producer_col, "")) if producer_col else ""
        if explicit_producer:
            producer = explicit_producer
            source = f"sku_suffix:{suffix}" if suffix else "producer_column"
        else:
            producer, source = infer_producer(title, sku, paths, known_producers)
        if producer:
            producers[producer] += 1
            producer_sources[f"{producer}|{source}"] += 1
            known_producers.add(producer)
            title_producer, title_source = infer_producer(title, "", paths, known_producers)
            if (
                suffix
                and title_source == "title_known"
                and title_producer
                and normalize_header(title_producer) != normalize_header(producer)
            ):
                producers[title_producer] += 1
                producer_sources[f"{title_producer}|sku_suffix:{suffix}"] += 1
                known_producers.add(title_producer)

        for path in paths:
            categories[path] += 1
            leaf = category_leaf(path)
            if leaf:
                category_leaves[leaf] += 1

        product_type = infer_type_from_title(title)
        if product_type:
            types[product_type] += 1
            for path in paths:
                leaf = category_leaf(path)
                if leaf:
                    type_by_category[leaf][product_type] += 1
            if len(examples[product_type]) < 5:
                examples[product_type].append(
                    {
                        "title": title,
                        "sku": sku,
                        "producer": producer,
                        "category": paths[0] if paths else "",
                    }
                )

    return {
        "summary": {
            "rows": int(len(df)),
            "producers": len(producers),
            "category_paths": len(categories),
            "category_leaves": len(category_leaves),
            "product_types": len(types),
        },
        "producers": [{"name": name, "count": count} for name, count in producers.most_common()],
        "producer_sources": [
            {"producer": key.split("|", 1)[0], "source": key.split("|", 1)[1], "count": count}
            for key, count in producer_sources.most_common()
        ],
        "sku_suffixes": [{"suffix": name, "count": count} for name, count in sku_suffixes.most_common()],
        "categories": [{"path": name, "count": count} for name, count in categories.most_common()],
        "category_leaves": [{"name": name, "count": count} for name, count in category_leaves.most_common()],
        "product_types": [{"type": name, "count": count} for name, count in types.most_common()],
        "type_by_category": [
            {
                "category": category,
                "type": product_type,
                "count": count,
            }
            for category, counter in sorted(type_by_category.items())
            for product_type, count in counter.most_common(10)
        ],
        "examples": dict(examples),
    }


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    ensure_dir(path.parent)
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Buduje słowniki typów, producentów i kategorii z eksportu produktów.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet", default=0)
    parser.add_argument("--output", default="dictionaries/learned_product_taxonomy.yaml")
    parser.add_argument("--reports-dir", default="reports/product_taxonomy")
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    columns = detect_columns(df)
    taxonomy = build_taxonomy(df, columns)

    output = Path(args.output)
    ensure_dir(output.parent)
    output.write_text(yaml.safe_dump(taxonomy, allow_unicode=True, sort_keys=False), encoding="utf-8")

    reports_dir = ensure_dir(args.reports_dir)
    (reports_dir / "detected_columns.json").write_text(json.dumps(columns, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(reports_dir / "producers.csv", taxonomy["producers"])
    write_csv(reports_dir / "producer_sources.csv", taxonomy["producer_sources"])
    write_csv(reports_dir / "sku_suffixes.csv", taxonomy["sku_suffixes"])
    write_csv(reports_dir / "categories.csv", taxonomy["categories"])
    write_csv(reports_dir / "category_leaves.csv", taxonomy["category_leaves"])
    write_csv(reports_dir / "product_types.csv", taxonomy["product_types"])
    write_csv(reports_dir / "type_by_category.csv", taxonomy["type_by_category"])

    print(f"OK: zapisano {output}")
    print(f"OK: zapisano raporty w {reports_dir}")
    print(json.dumps(taxonomy["summary"], ensure_ascii=False))


if __name__ == "__main__":
    main()
