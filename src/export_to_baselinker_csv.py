from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, is_blank, normalize_header, read_products


BASELINKER_COLUMNS = [
    "product_id",
    "name",
    "sku",
    "ean",
    "manufacturer_name",
    "description",
    "features",
    "images_urls",
]

FEATURE_MAP = {
    "Producent": ["attr_producent", "Producent", "Marka"],
    "EAN (GTIN)": ["EAN", "ean"],
    "Kod producenta": ["Kod Producenta", "Kod", "SKU", "sku"],
    "Typ produktu": ["attr_typ", "Typ", "Rodzaj produktu"],
    "Seria": ["attr_seria", "Rodzina"],
    "Model": ["attr_model", "Nazwa Kanlux"],
    "Moc [W]": ["attr_moc", "Moc [W]", "Moc"],
    "Napięcie [V]": ["attr_napiecie", "Napięcie [V]", "Napięcie (V)"],
    "Temperatura barwowa [K]": ["attr_barwa", "Barwa CCT [K]", "Temperatura barwowa"],
    "Barwa światła": ["Barwa - kategoria", "attr_barwa_zakres", "Barwa światła"],
    "Strumień świetlny [lm]": ["attr_strumien", "Strumień [lm]", "Strumień świetlny [lm]", "Jasność"],
    "Stopień ochrony [IP]": ["attr_ip", "Klasa IP", "Stopień ochrony [IP]", "Stopień ochrony IP"],
    "Klasa ochronności": ["attr_ik", "Klasa ochronności"],
    "Kąt świecenia": ["attr_kat_swiecenia", "Kąt", "Kąt świecenia"],
    "Kolor": ["attr_kolor", "Kolor obudowy"],
    "Trzonek": ["attr_gwint", "Gwint", "Trzonek", "Rodzaj gwintu"],
    "Materiał": ["attr_material", "Materiał"],
    "Kształt": ["attr_ksztalt", "Kształt"],
    "Wymiary": ["attr_wymiary", "Wymiary"],
    "Długość": ["attr_dlugosc", "Długość"],
    "Szerokość": ["attr_szerokosc", "Szerokość"],
    "Wysokość": ["attr_wysokosc", "Wysokość"],
    "Średnica": ["attr_srednica", "Średnica"],
    "Czujnik ruchu": ["attr_czujnik", "Czujnik ruchu"],
    "Liczba sztuk": ["attr_ilosc_sztuk"],
    "Gwarancja": ["attr_gwarancja"],
    "Sterowanie": ["attr_sterowanie"],
    "Pasuje do": ["attr_pasuje_do"],
    "Klasa efektywności energetycznej": ["Klasa EEi", "Klasa efektywności energetycznej"],
}

FEATURE_NAME_ALIASES = {
    "marka": "Producent",
    "producent": "Producent",
    "rodzaj produktu": "Typ produktu",
    "typ": "Typ produktu",
    "typ produktu": "Typ produktu",
    "moc": "Moc [W]",
    "moc w": "Moc [W]",
    "moc znamionowa": "Moc [W]",
    "napiecie": "Napięcie [V]",
    "napięcie": "Napięcie [V]",
    "napiecie v": "Napięcie [V]",
    "napięcie v": "Napięcie [V]",
    "temperatura barwowa": "Temperatura barwowa [K]",
    "temperatura barwowa k": "Temperatura barwowa [K]",
    "barwa cct k": "Temperatura barwowa [K]",
    "jasnosc": "Strumień świetlny [lm]",
    "jasność": "Strumień świetlny [lm]",
    "strumien lm": "Strumień świetlny [lm]",
    "strumień lm": "Strumień świetlny [lm]",
    "strumien swietlny lm": "Strumień świetlny [lm]",
    "strumień świetlny lm": "Strumień świetlny [lm]",
    "stopien ochrony ip": "Stopień ochrony [IP]",
    "stopień ochrony ip": "Stopień ochrony [IP]",
    "klasa ip": "Stopień ochrony [IP]",
    "rodzaj gwintu": "Trzonek",
    "gwint": "Trzonek",
    "trzonek": "Trzonek",
    "barwa swiatla": "Barwa światła",
    "barwa światła": "Barwa światła",
    "material": "Materiał",
    "materiał": "Materiał",
    "kolor obudowy": "Kolor",
}


def main() -> None:
    parser = argparse.ArgumentParser(description="Eksportuje wynik AlkanInatora do CSV importowego Baselinkera.")
    parser.add_argument("--input", required=True, help="CSV/XLSX po AlkanInatorze.")
    parser.add_argument("--sheet", help="Arkusz XLSX, jesli input ma wiele arkuszy.")
    parser.add_argument("--output", default="output/baselinker_import.csv", help="Docelowy CSV dla Baselinkera.")
    parser.add_argument("--name-column", default="", help="Wymusza kolumne nazwy. Domyslnie: new_title, Nazwa, Nazwa B2C / SEO final.")
    parser.add_argument("--include-empty-features", action="store_true", help="Zapisuje puste parametry w JSON features.")
    parser.add_argument("--links-file", default="", help="XLSX/CSV z linkami do zdjec, np. input/Linki.xlsx.")
    parser.add_argument("--links-sheet", default=None, help="Arkusz w pliku linkow. Domyslnie pierwszy arkusz.")
    parser.add_argument("--image-key-column", default="Kod Kanlux", help="Kolumna kodu produktu w pliku linkow.")
    parser.add_argument("--main-image-column", default="Zdjecie glowne", help="Kolumna glownego zdjecia w pliku linkow.")
    parser.add_argument(
        "--sku-format",
        choices=["plain", "producer_suffix"],
        default="producer_suffix",
        help="Format SKU w eksporcie. producer_suffix daje np. 22606/KAN.",
    )
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    image_matches = 0
    if args.links_file:
        df, image_matches = assign_main_images(
            df,
            args.links_file,
            args.links_sheet,
            args.image_key_column,
            args.main_image_column,
        )
    exported = build_baselinker_rows(df, args.name_column, args.include_empty_features, args.sku_format)
    write_baselinker_csv(exported, args.output)

    print(f"OK: wczytano {len(df)} produktow")
    if args.links_file:
        print(f"OK: przypisano zdjecia glowne: {image_matches}")
    print(f"OK: zapisano {args.output}")


def build_baselinker_rows(
    df: pd.DataFrame,
    name_column: str = "",
    include_empty_features: bool = False,
    sku_format: str = "producer_suffix",
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for _, row in df.iterrows():
        name = first_value(row, [name_column] if name_column else [])
        if not name:
            name = first_value(row, ["new_title", "Nazwa B2C / SEO final", "Nazwa", "Nazwa B2C / SEO", "Nazwa Kanlux"])

        sku = first_value(row, ["SKU", "Kod Producenta", "Kod", "sku"])
        ean = only_digits(first_value(row, ["EAN", "ean"]))
        producer = first_value(row, ["attr_producent", "Producent", "manufacturer_name", "Marka"])
        export_sku = format_sku(sku, producer, sku_format)
        description = first_value(row, ["Opis HTML", "description", "Opis"])
        images = first_value(row, ["images_urls", "Zdjęcie URL", "Obrazki", "Zdjecia", "Zdjęcia"])

        features = build_features(row, include_empty_features)
        rows.append(
            {
                "product_id": first_value(row, ["product_id", "Baselinker ID"]),
                "name": name,
                "sku": export_sku,
                "ean": ean,
                "manufacturer_name": producer,
                "description": description,
                "features": json.dumps(features, ensure_ascii=False, separators=(",", ":")),
                "images_urls": images,
            }
        )
    return rows


def assign_main_images(
    df: pd.DataFrame,
    links_file: str | Path,
    links_sheet: str | int | None,
    key_column: str,
    main_image_column: str,
) -> tuple[pd.DataFrame, int]:
    links = read_products(links_file, sheet_name=links_sheet)
    link_key_column = find_column(links, key_column) or find_column(links, "Kod")
    link_image_column = find_column(links, main_image_column) or find_column(links, "Zdjecie glowne")
    if not link_key_column or not link_image_column:
        raise SystemExit("Nie znaleziono kolumn kodu lub zdjecia glownego w pliku linkow.")

    image_by_code = {
        normalize_code(row.get(link_key_column, "")): compact_spaces(str(row.get(link_image_column, "")))
        for _, row in links.iterrows()
        if normalize_code(row.get(link_key_column, "")) and compact_spaces(str(row.get(link_image_column, "")))
    }

    result = df.copy()
    matched = 0
    images: list[str] = []
    for _, row in result.iterrows():
        code = normalize_code(first_value(row, ["Kod Producenta", "Kod", "SKU", "sku"]))
        image = image_by_code.get(code, "")
        if image:
            matched += 1
        images.append(image)
    result["images_urls"] = images
    return result, matched


def find_column(df: pd.DataFrame, expected: str) -> str | None:
    normalized_expected = normalize_header(expected)
    for column in df.columns:
        if normalize_header(str(column)) == normalized_expected:
            return str(column)
    for column in df.columns:
        normalized = normalize_header(str(column))
        if normalized_expected and normalized_expected in normalized:
            return str(column)
    return None


def normalize_code(value: Any) -> str:
    value = compact_spaces(str(value or ""))
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    return value


def build_features(row: pd.Series, include_empty: bool) -> dict[str, str]:
    features: dict[str, str] = {}
    for target, source_columns in FEATURE_MAP.items():
        value = normalize_feature_value(first_value(row, source_columns), target)
        if value or include_empty:
            features[target] = value

    for column, value in row.items():
        column_name = str(column)
        if not column_name.startswith("Parametr: "):
            continue
        raw_feature_name = column_name.replace("Parametr: ", "", 1)
        feature_name = normalize_feature_name(raw_feature_name)
        feature_value = normalize_feature_value(str(value), feature_name)
        if feature_value or include_empty:
            features.setdefault(feature_name, feature_value)
    return features


def write_baselinker_csv(rows: list[dict[str, str]], output: str | Path) -> None:
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=BASELINKER_COLUMNS, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(rows)


def first_value(row: pd.Series, columns: list[str]) -> str:
    for column in columns:
        if not column or column not in row:
            continue
        value = row.get(column, "")
        if not is_blank(value):
            return compact_spaces(str(value))
    return ""


def normalize_feature_name(value: str) -> str:
    value = compact_spaces(value)
    normalized = normalize_header(value)
    return FEATURE_NAME_ALIASES.get(normalized, value)


def normalize_feature_value(value: str, feature_name: str = "") -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    value = re.sub(r"\s*\|\s*", "|", value)

    if feature_name == "Stopień ochrony [IP]":
        return normalize_ip_value(value)
    if feature_name == "Moc [W]":
        return strip_unit(value, "W")
    if feature_name == "Temperatura barwowa [K]":
        return normalize_temperature_value(value)
    if feature_name == "Strumień świetlny [lm]":
        return strip_unit(value, "lm")
    if feature_name == "Napięcie [V]":
        return normalize_voltage_value(value)
    if feature_name == "Barwa światła":
        return normalize_light_color_value(value)
    if feature_name == "Trzonek":
        return normalize_socket_value(value)
    if feature_name == "Kolor":
        return normalize_color_value(value)
    if feature_name == "Materiał":
        return normalize_material_value(value)
    return value


def normalize_ip_value(value: str) -> str:
    normalized = value.upper().replace(" ", "")
    return re.sub(r"^IP(\d+)$", r"IP \1", normalized)


def strip_unit(value: str, unit: str) -> str:
    value = normalize_decimal_separator(value)
    value = re.sub(rf"\s*{re.escape(unit)}\b", "", value, flags=re.IGNORECASE)
    return compact_spaces(value)


def normalize_temperature_value(value: str) -> str:
    value = value.replace(" ", "")
    if re.fullmatch(r"[0-9/]+K?", value, flags=re.IGNORECASE):
        return value.replace("K", "").replace("k", "")
    return re.sub(r"(?<=\d)\s*K\b", "", value, flags=re.IGNORECASE)


def normalize_voltage_value(value: str) -> str:
    if value == "MLS":
        return ""
    value = re.sub(r"\s*AC\b", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s*DC\b", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s*V\b", "", value, flags=re.IGNORECASE)
    return compact_spaces(normalize_decimal_separator(value))


def normalize_light_color_value(value: str) -> str:
    mapping = {
        "Neutralna (3500-4500K)": "Neutralna",
        "Zimna (≥5000K)": "Zimna",
        "Ciepła (≤3000K)": "Ciepła",
        "Zmienna CCT": "Zmienna",
    }
    return mapping.get(value, value)


def normalize_socket_value(value: str) -> str:
    normalized = value.upper().replace(" ", "")
    match = re.fullmatch(r"\d+X(.+)", normalized)
    return match.group(1) if match else normalized


def normalize_color_value(value: str) -> str:
    aliases = {
        "bialy": "Biały",
        "biały": "Biały",
    }
    return aliases.get(value, capitalize_first(value))


def normalize_material_value(value: str) -> str:
    aliases = {
        "1": "",
        "PC": "Tworzywo sztuczne",
        "PMMA": "Tworzywo sztuczne",
        "PS": "Tworzywo sztuczne",
        "PP": "Polipropylen",
        "tworzywo sztuczne": "Tworzywo sztuczne",
        "szkło": "Szkło",
        "szkło hartowane": "Szkło",
        "drewniana": "Drewno",
        "stop aluminium": "Metal",
        "stal": "Metal",
        "stal nierdzewna": "Metal",
    }
    return aliases.get(value, capitalize_first(value))


def normalize_decimal_separator(value: str) -> str:
    return re.sub(r"(?<=\d)\.(?=\d)", ",", value)


def capitalize_first(value: str) -> str:
    return value[:1].upper() + value[1:] if value else value


def format_sku(sku: str, producer: str, sku_format: str) -> str:
    sku = compact_spaces(sku)
    if sku_format == "plain" or not sku:
        return sku
    if "/" in sku:
        return sku
    suffix = producer_suffix(producer)
    return f"{sku}/{suffix}" if suffix else sku


def producer_suffix(producer: str) -> str:
    letters = re.findall(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", str(producer or ""))
    return "".join(letters[:3]).upper()


def only_digits(value: Any) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    return digits or compact_spaces(str(value or ""))


if __name__ == "__main__":
    main()
