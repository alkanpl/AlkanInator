from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, first_present, load_yaml, read_products, strip_accents, write_products


ATTRIBUTE_COLUMNS = [
    "typ",
    "rodzaj",
    "moc",
    "napiecie",
    "barwa",
    "strumien",
    "ip",
    "ik",
    "kat_swiecenia",
    "kolor",
    "producent",
    "seria",
    "model",
    "wymiary",
    "ksztalt",
    "dlugosc",
    "szerokosc",
    "wysokosc",
    "czujnik",
    "gwint",
    "srednica",
    "material",
]


def extract_attributes_from_text(text: str, config: dict[str, Any] | None = None, producer_hint: str = "") -> dict[str, Any]:
    config = config or {}
    title = compact_spaces(text)
    title_ascii = strip_accents(title).lower()
    attributes: dict[str, str] = {}
    sources: dict[str, str] = {}
    confidence: dict[str, float] = {}

    patterns = {
        "moc": r"(?<![A-Z0-9])(\d{1,4}(?:[,.]\d+)?)\s*W(?![a-z])",
        "napiecie": r"(?<![A-Z0-9])(\d{2,4})\s*V\b",
        "barwa": r"(?<![A-Z0-9])([23645]\d{3})\s*K\b",
        "ip": r"\bIP\s*([0-9]{2})\b",
        "ik": r"\bIK\s*([0-9]{2})\b",
        "kat_swiecenia": r"(?<![A-Z0-9])(\d{2,3})\s*(?:°|st\.?|stopni)\b",
        "gwint": r"\b([1-9]x?\s*E(?:14|27|40))\b",
    }

    for attr, pattern in patterns.items():
        match = re.search(pattern, title, flags=re.IGNORECASE)
        if not match:
            continue
        value = match.group(1).replace(",", ".")
        if attr == "moc":
            value = f"{value}W"
        elif attr == "napiecie":
            value = f"{value}V"
        elif attr == "barwa":
            value = f"{value}K"
        elif attr == "strumien":
            value = f"{value}lm"
        elif attr == "ip":
            value = f"IP{value}"
        elif attr == "ik":
            value = f"IK{value}"
        elif attr == "kat_swiecenia":
            value = f"{value}°"
        elif attr == "gwint":
            value = normalize_socket(value)
        attributes[attr] = value
        sources[attr] = "title"
        confidence[attr] = 0.98

    luminous_flux = extract_luminous_flux(title)
    if luminous_flux:
        attributes["strumien"] = luminous_flux
        sources["strumien"] = "title"
        confidence["strumien"] = 0.95

    if "moc" not in attributes:
        power = extract_power_from_special_formats(title)
        if power:
            attributes["moc"] = power
            sources["moc"] = "title"
            confidence["moc"] = 0.85

    dimensions = extract_dimensions(title)
    for attr, value in dimensions.items():
        attributes[attr] = value
        sources[attr] = "title"
        confidence[attr] = 0.88

    shape = extract_shape(title_ascii)
    if shape:
        attributes["ksztalt"] = shape
        sources["ksztalt"] = "title"
        confidence["ksztalt"] = 0.88

    diameter = extract_diameter(title)
    if diameter:
        attributes["srednica"] = diameter
        sources["srednica"] = "title"
        confidence["srednica"] = 0.86

    material = extract_material(title_ascii)
    if material:
        attributes["material"] = material
        sources["material"] = "title"
        confidence["material"] = 0.82

    color = extract_color(title_ascii, config)
    if color:
        attributes["kolor"] = color
        sources["kolor"] = "title"
        confidence["kolor"] = 0.90

    producer = extract_producer(title, config, producer_hint)
    if producer:
        attributes["producent"] = producer
        sources["producent"] = "producer_column" if producer_hint else "title"
        confidence["producent"] = 0.99 if producer_hint else 0.90

    series, model = extract_series_and_model(title, attributes, config)
    if series:
        attributes["seria"] = series
        sources["seria"] = "title"
        confidence["seria"] = 0.65
    if model:
        attributes["model"] = model
        sources["model"] = "title"
        confidence["model"] = 0.60

    for key, value in (config.get("default_values") or {}).items():
        if key not in attributes:
            attributes[key] = str(value)
            sources[key] = "config_default"
            confidence[key] = 0.70

    return {"attributes": attributes, "sources": sources, "confidence": confidence}


def extract_color(title_ascii: str, config: dict[str, Any]) -> str:
    colors = config.get("known_colors") or {}
    for normalized, variants in colors.items():
        for variant in variants:
            variant_ascii = strip_accents(str(variant)).lower()
            if re.search(rf"\b{re.escape(variant_ascii)}\b", title_ascii):
                return str(normalized)
    return ""


def extract_power_from_special_formats(title: str) -> str:
    compact_title = compact_spaces(title)
    match = re.search(r"\b(\d{1,3}(?:[,.]\d+)?)W(?:CCT|DIM|NW|WW|CW)\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{normalize_number(match.group(1))}W"
    match = re.search(r"\b(\d{1,3})W\d{1,3}NW\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\bLED[-\s]?(\d{1,3})(?:[-A-Z]|$)", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\b(\d{1,3})W[DS]\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\bTW(\d{1,3})NW\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\bDRV(?:M)?(\d{1,3})W(?![a-z])", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\b(\d{1,2})\s*x\s*max\s*(\d{1,3})\b", compact_title, flags=re.IGNORECASE)
    if match:
        count = int(match.group(1))
        power = int(match.group(2))
        return f"{count}x{power}W"
    return ""


def extract_luminous_flux(title: str) -> str:
    compact_title = compact_spaces(title)
    list_match = re.search(
        r"\b(\d+(?:[,.]\d+)?(?:\s*/\s*\d+(?:[,.]\d+)?)+)\s*(?:lm|lumenow|lumenów)\b",
        compact_title,
        flags=re.IGNORECASE,
    )
    if list_match:
        values = [float(value.replace(",", ".")) for value in re.findall(r"\d+(?:[,.]\d+)?", list_match.group(1))]
        if not values:
            return ""
        min_value = format_number(min(values))
        max_value = format_number(max(values))
        return f"{min_value}-{max_value}lm" if min_value != max_value else f"{max_value}lm"

    match = re.search(r"(?<![A-Z0-9])(\d+(?:[,.]\d+)?)\s*(?:lm|lumenow|lumenów)\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{normalize_number(match.group(1))}lm"
    return ""


def extract_dimensions(title: str) -> dict[str, str]:
    dimensions: dict[str, str] = {}
    compact_title = compact_spaces(title)
    length = re.search(r"\b(?:dł\.?|dl\.?|długość|dlugosc)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    width = re.search(r"\b(?:szer\.?|szerokość|szerokosc)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    height = re.search(r"\b(?:wys\.?|wysokość|wysokosc)\s*(\d+(?:[,.]\d+)?)\s*(?:cm|mm)\b", compact_title, flags=re.IGNORECASE)
    if length:
        dimensions["dlugosc"] = normalize_number(length.group(1)) + "cm"
    if width:
        dimensions["szerokosc"] = normalize_number(width.group(1)) + "cm"
    if height:
        dimensions["wysokosc"] = normalize_number(height.group(1)) + "cm"
    if "dlugosc" in dimensions and "szerokosc" in dimensions:
        dimensions["wymiary"] = f"{dimensions['dlugosc']} x {dimensions['szerokosc']}"
    return dimensions


def extract_diameter(title: str) -> str:
    compact_title = compact_spaces(title)
    match = re.search(r"(?:Ø|fi|średnica|srednica)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    if match:
        return normalize_number(match.group(1)) + "cm"
    match = re.search(r"\b(?:okrągła|okragla|okrągły|okragly)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    if match:
        return normalize_number(match.group(1)) + "cm"
    return ""


def extract_material(title_ascii: str) -> str:
    if re.search(r"\bdrewnian", title_ascii):
        return "drewniana"
    if re.search(r"\balumini", title_ascii):
        return "aluminium"
    if re.search(r"\bstal", title_ascii):
        return "stal"
    return ""


def extract_shape(title_ascii: str) -> str:
    if re.search(r"\bprostokatny\b", title_ascii):
        return "prostokątna"
    if re.search(r"\bkwadratowy\b", title_ascii):
        return "kwadratowa"
    if re.search(r"\bokragl[ay]\b", title_ascii):
        return "okrągła"
    return ""


def normalize_number(value: str) -> str:
    value = value.replace(",", ".")
    if value.endswith(".0"):
        value = value[:-2]
    return value


def format_number(value: float) -> str:
    return str(int(value)) if value.is_integer() else str(value).rstrip("0").rstrip(".")


def normalize_socket(value: str) -> str:
    value = re.sub(r"\s+", "", value.upper())
    value = value.replace("X", "x")
    return value


def extract_producer(title: str, config: dict[str, Any], producer_hint: str = "") -> str:
    if producer_hint:
        return compact_spaces(producer_hint)
    title_ascii = strip_accents(title).lower()
    for producer in config.get("known_producers") or []:
        producer_ascii = strip_accents(str(producer)).lower()
        if re.search(rf"\b{re.escape(producer_ascii)}\b", title_ascii):
            return str(producer)
    return ""


def extract_series_and_model(title: str, attributes: dict[str, str], config: dict[str, Any]) -> tuple[str, str]:
    cleaned = title
    for value in sorted(attributes.values(), key=len, reverse=True):
        cleaned = re.sub(re.escape(str(value)), " ", cleaned, flags=re.IGNORECASE)
    technical_patterns = [
        r"\b\d+(?:[,.]\d+)?\s*W\b",
        r"\b\d{2,4}\s*V\b",
        r"\b[23645]\d{3}\s*K\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:lm|lumenow|lumenów)\b",
        r"\bIP\s*\d{2}\b",
        r"\bIK\s*\d{2}\b",
    ]
    for pattern in technical_patterns:
        cleaned = re.sub(pattern, " ", cleaned, flags=re.IGNORECASE)
    stop_words = {strip_accents(word).lower() for word in config.get("series_stop_words") or []}
    tokens = re.findall(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż0-9-]{3,}", cleaned)
    useful = [token for token in tokens if strip_accents(token).lower() not in stop_words]
    uppercase_tokens = [token for token in useful if token.isupper() and not token.isdigit()]
    titlecase_tokens = [token for token in useful if token[:1].isupper() and not any(char.isdigit() for char in token)]
    series = " ".join((uppercase_tokens or titlecase_tokens)[:2])
    model = first_present(
        *(token for token in useful if any(char.isdigit() for char in token) and not token.isdigit())
    )
    return compact_spaces(series), compact_spaces(model)


def extract_attributes_for_dataframe(
    df: pd.DataFrame,
    title_column: str,
    config: dict[str, Any],
    producer_column: str | None = None,
) -> pd.DataFrame:
    result = df.copy()
    extracted_rows: list[dict[str, Any]] = []

    for _, row in result.iterrows():
        producer_hint = row.get(producer_column, "") if producer_column else ""
        extracted = extract_attributes_from_text(str(row.get(title_column, "")), config, str(producer_hint))
        attrs = extracted["attributes"]
        merge_attributes_from_columns(attrs, extracted["sources"], extracted["confidence"], row, config)
        for attr in ATTRIBUTE_COLUMNS:
            result.at[row.name, f"attr_{attr}"] = attrs.get(attr, "")
        result.at[row.name, "attribute_sources"] = json.dumps(extracted["sources"], ensure_ascii=False, sort_keys=True)
        result.at[row.name, "attribute_confidence"] = json.dumps(extracted["confidence"], ensure_ascii=False, sort_keys=True)
        extracted_rows.append(attrs)

    return result


def merge_attributes_from_columns(
    attrs: dict[str, str],
    sources: dict[str, str],
    confidence: dict[str, float],
    row: pd.Series,
    config: dict[str, Any],
) -> None:
    mappings = config.get("input_attribute_columns") or {}
    for attr, column in mappings.items():
        if should_keep_existing_attribute(str(attr), attrs.get(str(attr), "")):
            continue
        if column not in row or not str(row.get(column, "")).strip():
            continue
        value = normalize_attribute_value(str(row.get(column, "")).strip(), str(attr))
        if not value:
            continue
        attrs[str(attr)] = value
        sources[str(attr)] = f"column:{column}"
        confidence[str(attr)] = 0.95
    apply_attribute_overrides(attrs, sources, confidence, row, config)


def should_keep_existing_attribute(attr: str, value: str) -> bool:
    if not value:
        return False
    if attr == "strumien" and "-" in str(value):
        return True
    return False


def normalize_attribute_value(value: str, attr: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    if attr == "moc" and re.fullmatch(r"\d+(?:[,.]\d+)?", value):
        return f"{value.replace(',', '.')}W"
    if attr == "barwa" and re.fullmatch(r"\d{3,5}", value):
        return f"{value}K"
    if attr == "strumien" and re.fullmatch(r"\d+(?:[,.]\d+)?", value):
        return f"{value.replace(',', '.')}lm"
    if attr == "ip" and re.fullmatch(r"\d{2}", value):
        return f"IP{value}"
    if attr == "gwint":
        return normalize_socket(value)
    if attr == "czujnik":
        lowered = strip_accents(value).lower()
        if lowered in {"tak", "yes", "1", "true"}:
            return "z czujnikiem"
        if lowered in {"nie", "no", "0", "false"}:
            return ""
        if "czuj" in lowered or "pir" in lowered or "mikrofal" in lowered:
            return "z czujnikiem"
    return value


def apply_attribute_overrides(
    attrs: dict[str, str],
    sources: dict[str, str],
    confidence: dict[str, float],
    row: pd.Series,
    config: dict[str, Any],
) -> None:
    for rule in config.get("attribute_overrides") or []:
        attr = str(rule.get("attribute", ""))
        if not attr:
            continue
        current_value = str(attrs.get(attr, ""))
        category_column = str(rule.get("category_column", "Kategoria"))
        category = str(row.get(category_column, ""))
        if not contains_text(category, str(rule.get("category_contains", ""))):
            continue
        if not contains_text(current_value, str(rule.get("value_contains", ""))):
            continue
        attrs[attr] = str(rule.get("value", current_value))
        sources[attr] = f"{sources.get(attr, 'unknown')}+override"
        confidence[attr] = min(float(confidence.get(attr, 0.95)), 0.90)


def contains_text(value: str, needle: str) -> bool:
    if not needle:
        return True
    return strip_accents(needle).lower() in strip_accents(value).lower()


def main() -> None:
    parser = argparse.ArgumentParser(description="Wyciaga atrybuty z tytulow produktow.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet")
    parser.add_argument("--output", required=True)
    parser.add_argument("--title-column", required=True)
    parser.add_argument("--producer-column")
    parser.add_argument("--config", default="configs/categories/oprawy-sufitowe.yaml")
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    config = load_yaml(args.config)
    result = extract_attributes_for_dataframe(df, args.title_column, config, args.producer_column)
    write_products(result, Path(args.output))


if __name__ == "__main__":
    main()
