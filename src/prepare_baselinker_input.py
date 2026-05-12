from __future__ import annotations

import argparse
import csv
import json
import re
from html import unescape
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, ensure_dir, is_blank, write_products


BASELINKER_FIELD_ALIASES = {
    "product_id": ["product_id", "id", "produkt_id"],
    "name": ["name", "nazwa", "product_name"],
    "sku": ["sku", "symbol", "kod"],
    "ean": ["ean", "gtin", "ean13"],
    "manufacturer_name": ["manufacturer_name", "manufacturer", "producer", "producent", "marka"],
    "description": ["description", "opis"],
    "features": ["features", "parametry", "cechy"],
    "images_urls": ["images_urls", "images", "zdjecia", "obrazki"],
}

FEATURE_OUTPUT_MAP = {
    "Rodzaj": ["Rodzaj"],
    "Typ": ["Rodzaj produktu", "Typ", "Typ gniazda", "Rodzaj lampy", "Zrodlo swiatla", "Źródło światła"],
    "Seria": ["Seria", "Kolekcja"],
    "Rodzina": ["Seria", "Kolekcja"],
    "Model": ["Model"],
    "Kolor": ["Kolor", "Kolor dominujacy", "Kolor dominujący", "Kolor szkla", "Kolor szkła"],
    "Kolor obudowy": ["Kolor", "Kolor dominujacy", "Kolor dominujący", "Kolor obudowy"],
    "Moc [W]": ["Moc", "Moc znamionowa", "Moc zrodla swiatla", "Moc źródła światła"],
    "Barwa CCT [K]": ["Temperatura barwowa"],
    "Barwa - kategoria": ["Barwa swiatla", "Barwa światła"],
    "Klasa IP": ["Stopien ochrony IP", "Stopień ochrony IP"],
    "Strumień [lm]": ["Jasnosc", "Jasność", "Strumien swietlny", "Strumień świetlny"],
    "Gwint": ["Rodzaj gwintu", "Gwint", "Trzonek"],
    "Napięcie [V]": ["Napiecie (V)", "Napięcie (V)", "Napiecie pracy", "Napięcie pracy", "Zasilanie"],
    "Materiał": ["Material", "Materiał", "Material dominujacy", "Materiał dominujący"],
    "Kształt": ["Ksztalt", "Kształt", "Ksztalt oprawy", "Kształt oprawy"],
    "Średnica": ["Srednica", "Średnica", "Srednica/szerokosc", "Średnica/szerokość"],
    "Szerokość": ["Szerokosc", "Szerokość", "Szerokosc produktu", "Szerokość produktu"],
    "Wysokość": ["Wysokosc", "Wysokość", "Wysokosc produktu", "Wysokość produktu"],
    "Długość": ["Dlugosc", "Długość", "Dlugosc/wysokosc", "Długość/wysokość"],
    "Uziemienie": ["Uziemienie"],
    "Liczba sztuk": ["Liczba sztuk", "Liczba elementow w zestawie", "Liczba elementów w zestawie"],
}

BASE_OUTPUT_COLUMNS = [
    "Baselinker ID",
    "Kod",
    "SKU",
    "EAN",
    "Producent",
    "Kategoria",
    "Nazwa",
    "Nazwa B2C / SEO",
    "Opis",
    "Opis HTML",
    "Obrazki",
]


def main() -> None:
    parser = argparse.ArgumentParser(description="Przygotowuje eksport Baselinker CSV pod pipeline AlkanInator.")
    parser.add_argument("--input", required=True, help="CSV z Baselinkera.")
    parser.add_argument("--output", default="output/baselinker_prepared.csv", help="Docelowy plik CSV/XLSX.")
    parser.add_argument("--default-category", default="", help="Kategoria wpisywana, gdy CSV nie ma kategorii.")
    parser.add_argument("--delimiter", default="", help="Separator CSV. Domyslnie wykrywany automatycznie.")
    parser.add_argument("--encoding", default="utf-8-sig", help="Kodowanie pliku wejsciowego.")
    parser.add_argument("--reports-dir", default="reports/baselinker_prepare", help="Katalog raportow.")
    args = parser.parse_args()

    rows, headers = read_baselinker_rows(args.input, args.encoding, args.delimiter)
    prepared, report = prepare_rows(rows, headers, args.default_category)
    write_products(prepared, args.output)
    write_report(report, args.reports_dir)

    print(f"OK: wczytano {len(rows)} produktow")
    print(f"OK: zapisano {args.output}")
    print(f"OK: raport w {args.reports_dir}")


def read_baselinker_rows(path: str | Path, encoding: str, delimiter: str = "") -> tuple[list[dict[str, str]], list[str]]:
    input_path = Path(path)
    with input_path.open("r", encoding=encoding, newline="") as handle:
        sample = handle.read(8192)
        handle.seek(0)
        dialect = csv.excel()
        dialect.delimiter = delimiter or detect_delimiter_from_header(sample)
        if not dialect.delimiter:
            dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
        reader = csv.DictReader(handle, dialect=dialect)
        rows = [{str(key): str(value or "") for key, value in row.items()} for row in reader]
        return rows, list(reader.fieldnames or [])


def detect_delimiter_from_header(sample: str) -> str:
    header = sample.splitlines()[0] if sample else ""
    counts = {candidate: header.count(candidate) for candidate in [",", ";", "\t", "|"]}
    delimiter, count = max(counts.items(), key=lambda item: item[1])
    return delimiter if count else ""


def prepare_rows(rows: list[dict[str, str]], headers: list[str], default_category: str) -> tuple[pd.DataFrame, dict[str, Any]]:
    aliases = resolve_aliases(headers)
    prepared_rows: list[dict[str, Any]] = []
    feature_keys: set[str] = set()
    malformed_features = 0
    without_sku = 0
    without_ean = 0

    for row in rows:
        features, features_ok = parse_features(row.get(aliases.get("features", ""), ""))
        if not features_ok:
            malformed_features += 1
        feature_keys.update(features.keys())

        sku = first_present(row.get(aliases.get("sku", ""), ""), features.get("Kod producenta", ""))
        ean = first_present(row.get(aliases.get("ean", ""), ""), features.get("EAN (GTIN)", ""))
        producer = first_present(row.get(aliases.get("manufacturer_name", ""), ""), features.get("Marka", ""))
        title = first_present(row.get(aliases.get("name", ""), ""), features.get("Nazwa", ""))
        category = first_present(features.get("Kategoria", ""), row.get("category", ""), default_category)

        if is_blank(sku):
            without_sku += 1
        if is_blank(ean):
            without_ean += 1

        output = {
            "Baselinker ID": row.get(aliases.get("product_id", ""), ""),
            "Kod": sku,
            "SKU": sku,
            "EAN": only_digits(ean),
            "Producent": normalize_producer(producer),
            "Kategoria": category,
            "Nazwa": title,
            "Nazwa B2C / SEO": title,
            "Opis": html_to_text(row.get(aliases.get("description", ""), "")),
            "Opis HTML": row.get(aliases.get("description", ""), ""),
            "Obrazki": row.get(aliases.get("images_urls", ""), ""),
        }

        for output_column, source_keys in FEATURE_OUTPUT_MAP.items():
            output[output_column] = normalize_feature_value(first_present(*(features.get(key, "") for key in source_keys)))

        for key, value in features.items():
            output[f"Parametr: {key}"] = value

        prepared_rows.append(output)

    df = pd.DataFrame(prepared_rows)
    ordered_columns = [column for column in [*BASE_OUTPUT_COLUMNS, *FEATURE_OUTPUT_MAP.keys()] if column in df.columns]
    parameter_columns = sorted(column for column in df.columns if column.startswith("Parametr: "))
    other_columns = [column for column in df.columns if column not in ordered_columns and column not in parameter_columns]
    df = df[[*ordered_columns, *other_columns, *parameter_columns]]

    report = {
        "rows": len(rows),
        "input_columns": headers,
        "detected_columns": aliases,
        "feature_keys": sorted(feature_keys),
        "feature_keys_count": len(feature_keys),
        "malformed_features": malformed_features,
        "without_sku": without_sku,
        "without_ean": without_ean,
        "output_columns": list(df.columns),
    }
    return df, report


def resolve_aliases(headers: list[str]) -> dict[str, str]:
    normalized_headers = {normalize_key(header): header for header in headers}
    result: dict[str, str] = {}
    for target, candidates in BASELINKER_FIELD_ALIASES.items():
        for candidate in candidates:
            match = normalized_headers.get(normalize_key(candidate))
            if match:
                result[target] = match
                break
    return result


def parse_features(value: str) -> tuple[dict[str, str], bool]:
    value = compact_spaces(value)
    if not value:
        return {}, True
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError:
        return {}, False
    if not isinstance(parsed, dict):
        return {}, False
    return {str(key): compact_spaces(str(val)) for key, val in parsed.items() if not is_blank(val)}, True


def html_to_text(value: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", str(value or ""))
    text = re.sub(r"(?i)<br\s*/?>|</p>|</li>|</h[1-6]>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    return compact_spaces(unescape(text))


def first_present(*values: Any) -> str:
    for value in values:
        if not is_blank(value):
            return compact_spaces(str(value))
    return ""


def normalize_key(value: str) -> str:
    replacements = str.maketrans(
        "ąćęłńóśźżĄĆĘŁŃÓŚŹŻ",
        "acelnoszzACELNOSZZ",
    )
    value = str(value).translate(replacements).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return compact_spaces(value)


def normalize_feature_value(value: str) -> str:
    value = compact_spaces(value)
    value = re.sub(r"\s*\|\s*", ", ", value)
    return value


def normalize_producer(value: str) -> str:
    value = compact_spaces(value)
    known = {
        "OSPEL": "Ospel",
        "KONTAKT-SIMON": "Kontakt Simon",
        "KONTAKT SIMON": "Kontakt Simon",
        "HAGER": "Hager",
        "BERKER": "Berker",
        "KANLUX": "Kanlux",
    }
    return known.get(value.upper(), value)


def only_digits(value: str) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    return digits or compact_spaces(value)


def write_report(report: dict[str, Any], reports_dir: str | Path) -> None:
    path = ensure_dir(reports_dir)
    (path / "summary.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    pd.DataFrame({"feature_key": report["feature_keys"]}).to_csv(
        path / "feature_keys.csv",
        index=False,
        encoding="utf-8-sig",
    )


if __name__ == "__main__":
    main()
