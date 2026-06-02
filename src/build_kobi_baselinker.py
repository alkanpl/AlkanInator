from __future__ import annotations

import argparse
import csv
import json
import re
import xml.etree.ElementTree as ET
from html import unescape
from pathlib import Path
from typing import Any

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, load_catalog_knowledge
from export_to_baselinker_csv import build_manufacturer_data_by_producer, manufacturer_data_for_producer
from generate_product_descriptions import build_description_html
from recommend_categories import (
    DEFAULT_OVERRIDES_PATH,
    DEFAULT_TAXONOMY_PATH,
    CategoryRecommender,
    detect_recommendation_columns,
)
from utils import load_yaml, strip_accents


BASELINKER_COLUMNS = [
    "product_id",
    "name",
    "sku",
    "ean",
    "manufacturer_name",
    "category",
    "description",
    "features",
    "images_urls",
]

FEATURE_NAME_MAP = {
    "EAN": "EAN (GTIN)",
    "Kod_producenta": "Kod producenta",
    "Producent odpowiedzialny": "Dane producenta",
    "Podmiot odpowiedzialny": "Dane producenta",
    "Marka": "Marka",
    "Moc [W]": "Moc [W]",
    "Napięcie zasilania": "Napięcie [V]",
    "Temperatura barwowa |K|": "Temperatura barwowa [K]",
    "Strumień świetlny |lm|": "Strumień świetlny [lm]",
    "IP": "Stopień ochrony [IP]",
    "IK": "Stopień ochrony [IK]",
    "Kąt świecenia [°]": "Kąt świecenia [°]",
    "Kolor produktu": "Kolor",
    "Klasa ochronności PPE": "Klasa ochronności",
    "Klasa źródła w oprawie": "Klasa źródła w oprawie",
    "EEI": "EEI",
    "Gwarancja w latach": "Gwarancja",
    "Możliwość sterowania": "Możliwość sterowania",
    "Sposób montażu": "Sposób montażu",
    "Typ mocowania": "Typ mocowania",
    "Typ czujnika": "Typ czujnika",
    "Typ baterii": "Typ baterii",
    "Trzonek": "Trzonek",
    "Wskaźnik RA": "Wskaźnik RA",
    "Żywotność w godz.": "Żywotność [h]",
    "Liczba cykli": "Liczba cykli",
    "Atesty": "Atesty",
    "Eprel": "EPREL",
    "Prąd |A|": "Prąd [A]",
    "Zakres temperatury otoczenia": "Zakres temperatury otoczenia",
    "Producent GPSR": "Dane producenta",
}

TRAILING_BRANDS = [
    "Kobi Professional",
    "Kobi Premium",
    "Kobi Pro",
    "KOBI",
    "Kobi",
    "LED2B RED",
    "LED2B",
]

KOBI_MANUFACTURER_ALIASES = {
    "KOBI",
    "KOBI LIGHT",
    "KOBI PREMIUM",
    "KOBI PRO",
    "KOBI PROFESSIONAL",
    "LED2B",
    "LED2B RED",
    "RED",
}

COLOR_FORMS = {
    "biały": "biały",
    "bialy": "biały",
    "biała": "biały",
    "biala": "biały",
    "białe": "biały",
    "czarny": "czarny",
    "czarna": "czarny",
    "czarne": "czarny",
    "szary": "szary",
    "szara": "szary",
    "chrom": "chrom",
}

COLOR_INFLECTIONS = {
    "biały": {"masculine": "biały", "feminine": "biała", "neuter": "białe"},
    "czarny": {"masculine": "czarny", "feminine": "czarna", "neuter": "czarne"},
    "szary": {"masculine": "szary", "feminine": "szara", "neuter": "szare"},
    "grafitowy": {"masculine": "grafitowy", "feminine": "grafitowa", "neuter": "grafitowe"},
    "srebrny": {"masculine": "srebrny", "feminine": "srebrna", "neuter": "srebrne"},
    "złoty": {"masculine": "złoty", "feminine": "złota", "neuter": "złote"},
}

SHAPE_INFLECTIONS = {
    "kwadratowy": {"masculine": "kwadratowy", "feminine": "kwadratowa", "neuter": "kwadratowe"},
    "prostokątny": {"masculine": "prostokątny", "feminine": "prostokątna", "neuter": "prostokątne"},
    "okrągły": {"masculine": "okrągły", "feminine": "okrągła", "neuter": "okrągłe"},
}


def main() -> None:
    parser = argparse.ArgumentParser(description="Buduje nazwy i import Baselinker dla listy Kobi.")
    parser.add_argument("--input", default="input/Kobi_27_05_2026.xlsx", help="Plik XLSX z lista produktow.")
    parser.add_argument("--xml", default="input/ceneo_kobi.xml", help="XML Ceneo/Kobi z atrybutami.")
    parser.add_argument("--sheet", default=None, help="Arkusz XLSX. Domyslnie pierwszy.")
    parser.add_argument(
        "--output-xlsx",
        default="output/kobi_27_05_2026_uzupelnione.xlsx",
        help="Pelny plik kontrolny XLSX.",
    )
    parser.add_argument(
        "--output-csv",
        default="output/baselinker_import_kobi_27_05_2026.csv",
        help="CSV importowy Baselinker.",
    )
    parser.add_argument(
        "--report",
        default="reports/kobi_27_05_2026_summary.json",
        help="Raport JSON z dopasowania.",
    )
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE_PATH)
    parser.add_argument("--taxonomy", default=DEFAULT_TAXONOMY_PATH)
    parser.add_argument("--overrides", default=DEFAULT_OVERRIDES_PATH)
    args = parser.parse_args()

    source_df = pd.read_excel(args.input, sheet_name=args.sheet)
    if isinstance(source_df, dict):
        source_df = next(iter(source_df.values()))
    source_df = source_df[source_df["Kod"].notna()].copy()

    offers = read_xml_offers(args.xml)
    catalog_knowledge = load_catalog_knowledge(args.catalog_knowledge)
    taxonomy = load_yaml(args.taxonomy) if args.taxonomy else {}
    overrides = load_yaml(args.overrides) if args.overrides else {}
    recommender = CategoryRecommender(catalog_knowledge, taxonomy, overrides)
    allowed_feature_specs = build_allowed_feature_specs(catalog_knowledge)
    manufacturer_data_by_producer = build_manufacturer_data_by_producer(catalog_knowledge)
    enriched_df, baselinker_rows, report = build_outputs(
        source_df,
        offers,
        recommender,
        allowed_feature_specs,
        manufacturer_data_by_producer,
    )

    write_xlsx(enriched_df, args.output_xlsx)
    write_baselinker_csv(baselinker_rows, args.output_csv)
    write_report(report, args.report)

    print(f"OK: produkty w XLSX: {len(source_df)}")
    print(f"OK: dopasowano XML: {report['matched_xml']} / {report['rows']}")
    print(f"OK: zapisano {args.output_xlsx}")
    print(f"OK: zapisano {args.output_csv}")
    print(f"OK: raport {args.report}")


def read_xml_offers(path: str | Path) -> dict[str, dict[str, Any]]:
    offers: dict[str, dict[str, Any]] = {}
    for _, elem in ET.iterparse(path, events=("end",)):
        if elem.tag != "o":
            continue

        attrs = {
            compact_spaces(a.attrib.get("name", "")): compact_spaces(a.text or "")
            for a in elem.findall("./attrs/a")
        }
        features = {
            compact_spaces(feature.attrib.get("name", "")): compact_spaces(feature.text or "")
            for feature in elem.findall("./features/feature")
        }
        code = normalize_code(attrs.get("Kod_producenta", ""))
        if code:
            offers[code] = {
                "product_id": compact_spaces(elem.attrib.get("id", "")),
                "url": compact_spaces(elem.attrib.get("url", "")),
                "price": compact_spaces(elem.attrib.get("price", "")),
                "price_net": compact_spaces(elem.attrib.get("price_net", "")),
                "avail": compact_spaces(elem.attrib.get("avail", "")),
                "availability": compact_spaces(elem.attrib.get("dostepnosc", "")),
                "category": compact_spaces(elem.findtext("cat") or ""),
                "name": compact_spaces(elem.findtext("name") or ""),
                "description_html": compact_html(elem.findtext("desc") or ""),
                "description_text": html_to_text(elem.findtext("desc") or ""),
                "images": [
                    compact_spaces(image.attrib.get("url", ""))
                    for image in elem.findall("./imgs/*")
                    if compact_spaces(image.attrib.get("url", ""))
                ],
                "attrs": attrs,
                "features": features,
            }
        elem.clear()
    return offers


def build_allowed_feature_specs(catalog_knowledge: dict[str, Any]) -> dict[str, dict[str, Any]]:
    specs: dict[str, dict[str, Any]] = {}
    for row in catalog_knowledge.get("top_attributes") or []:
        attribute = compact_spaces(row.get("attribute", ""))
        if attribute:
            spec = specs.setdefault(normalize_catalog_key(attribute), {"name": attribute, "values": {}})
            for value in parse_top_attribute_values(row.get("top_values", "")):
                spec["values"].setdefault(normalize_catalog_key(value), value)

    for row in catalog_knowledge.get("top_attribute_values_by_category") or []:
        attribute = compact_spaces(row.get("attribute", ""))
        value = compact_spaces(row.get("value", ""))
        if not attribute or not value:
            continue
        spec = specs.setdefault(normalize_catalog_key(attribute), {"name": attribute, "values": {}})
        spec["values"].setdefault(normalize_catalog_key(value), value)
    return specs


def parse_top_attribute_values(value: str) -> list[str]:
    values: list[str] = []
    for part in str(value or "").split(";"):
        cleaned = re.sub(r"\s+\(\d+\)\s*$", "", compact_spaces(part))
        if cleaned:
            values.append(cleaned)
    return values


def normalize_catalog_key(value: str) -> str:
    value = strip_accents(compact_spaces(value)).lower()
    value = re.sub(r"\\+$", "", value).strip()
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def filter_baselinker_features(
    features: dict[str, str],
    allowed_feature_specs: dict[str, dict[str, Any]],
) -> tuple[dict[str, str], list[str]]:
    filtered: dict[str, str] = {}
    removed: list[str] = []
    for key, value in features.items():
        value = compact_spaces(value)
        if not value:
            continue
        spec = allowed_feature_specs.get(normalize_catalog_key(key))
        if not spec:
            removed.append(key)
            continue
        canonical_value = find_allowed_feature_value(key, value, spec)
        if not canonical_value:
            removed.append(key)
            continue
        filtered[str(spec["name"])] = canonical_value
    removed = sorted(dict.fromkeys(removed))
    return filtered, removed


def find_allowed_feature_value(key: str, value: str, spec: dict[str, Any]) -> str:
    values = spec.get("values", {})
    if normalize_catalog_key(key) == "dane producenta":
        canonical_parts: list[str] = []
        for part in split_manufacturer_data_parts(value):
            canonical = values.get(normalize_catalog_key(part))
            if not canonical:
                return ""
            canonical_parts.append(str(canonical).strip("\\").strip())
        if len(canonical_parts) > 1:
            return "; ".join(canonical_parts)

    for candidate in feature_value_candidates(key, value):
        canonical = values.get(normalize_catalog_key(candidate))
        if canonical:
            return str(canonical)
    return ""


def split_manufacturer_data_parts(value: str) -> list[str]:
    return [
        part.strip("\\").strip()
        for part in re.split(r"\s*(?:;|\||>)\s*", compact_spaces(value))
        if part.strip("\\").strip()
    ]


def feature_value_candidates(key: str, value: str) -> list[str]:
    attr = normalize_catalog_key(key)
    raw = compact_spaces(value)
    candidates = [raw]
    no_unit_space = normalize_unit_spacing(raw)
    candidates.append(no_unit_space)
    candidates.append(raw.upper())
    candidates.append(raw.title())

    numeric = normalize_numeric_value(raw)
    if numeric:
        candidates.append(numeric)

    if "stopien ochrony" in attr:
        ip_match = re.search(r"\bIP\s*-?\s*(\d{2})\b", raw, flags=re.IGNORECASE)
        if ip_match:
            candidates.extend([f"IP {ip_match.group(1)}", f"IP{ip_match.group(1)}", f"IP-{ip_match.group(1)}"])

    if attr in {"strumien swietlny lm", "strumien swietlny", "jasnosc"}:
        lumen_match = re.search(r"\d+(?:[.,]\d+)?", raw)
        if lumen_match:
            candidates.append(lumen_match.group(0).replace(",", "."))

    if attr in {"temperatura barwowa k", "temperatura barwowa"}:
        cct_numbers = re.findall(r"\d{4}", raw)
        if cct_numbers:
            candidates.extend(cct_numbers)
            if len(cct_numbers) > 1:
                candidates.append("/".join(cct_numbers))
                candidates.append("-".join(cct_numbers))

    if attr in {"kat swiecenia", "kat swiecenia"} or "kat swiecenia" in attr:
        angle = re.search(r"\d+(?:[.,]\d+)?", raw)
        if angle:
            candidates.append(angle.group(0).replace(",", "."))

    if attr in {"moc w", "moc"}:
        if numeric:
            candidates.append(numeric.rstrip("0").rstrip("."))

    if attr in {"napiecie v", "napiecie", "zasilanie"}:
        voltage = re.search(r"\d+(?:[.,]\d+)?", raw)
        if voltage:
            candidates.append(voltage.group(0).replace(",", "."))

    if attr == "ksztalt":
        shape = normalize_catalog_key(raw)
        shape_map = {
            "kwadratowy": "Kwadrat",
            "kwadratowa": "Kwadrat",
            "kwadratowe": "Kwadrat",
            "kwadrat": "Kwadrat",
            "prostokatny": "Prostokątny",
            "prostokatna": "Prostokątny",
            "prostokatne": "Prostokątny",
            "prostokat": "Prostokątny",
            "okragly": "Okrągły",
            "okragla": "Okrągły",
            "okragle": "Okrągły",
            "okrag": "Okrągły",
        }
        mapped = shape_map.get(shape)
        if mapped:
            candidates.append(mapped)

    if attr == "kolor":
        color = normalize_color(raw)
        color_map = {
            "biały": "Biały",
            "czarny": "Czarny",
            "szary": "Szary",
            "grafitowy": "Grafitowy",
            "srebrny": "Srebrny",
            "złoty": "Złoty",
        }
        mapped = color_map.get(color)
        if mapped:
            candidates.append(mapped)

    if attr == "gwarancja":
        warranty = re.search(r"\d+", raw)
        if warranty:
            number = warranty.group(0)
            candidates.extend([f"{number} Lata", f"{number} lata", f"{number} Miesiące", f"{number} miesiące"])

    return list(dict.fromkeys(candidate for candidate in candidates if compact_spaces(candidate)))


def normalize_numeric_value(value: str) -> str:
    match = re.search(r"\d+(?:[.,]\d+)?", compact_spaces(value))
    if not match:
        return ""
    number = match.group(0).replace(",", ".")
    return re.sub(r"\.0$", "", number)


def build_outputs(
    source_df: pd.DataFrame,
    offers: dict[str, dict[str, Any]],
    recommender: CategoryRecommender,
    allowed_feature_specs: dict[str, dict[str, Any]],
    manufacturer_data_by_producer: dict[str, str],
) -> tuple[pd.DataFrame, list[dict[str, str]], dict[str, Any]]:
    enriched_rows: list[dict[str, Any]] = []
    baselinker_rows: list[dict[str, str]] = []
    missing_codes: list[dict[str, str]] = []
    removed_feature_names: dict[str, int] = {}

    for _, source_row in source_df.iterrows():
        row = {str(column): source_row.get(column, "") for column in source_df.columns}
        code = normalize_code(first_present(row.get("Kod"), row.get("Kod towaru")))
        offer = offers.get(code, {})
        attrs = offer.get("attrs", {})
        features = offer.get("features", {})
        original_name = first_present(row.get("Nazwa producenta"), offer.get("name"))
        brand = normalize_brand(first_present(features.get("Marka"), infer_brand_from_title(original_name), row.get("Producent")))
        producer = normalize_manufacturer(first_present(attrs.get("Producent"), row.get("Producent"), brand))
        ean = only_digits(attrs.get("EAN", ""))
        sku = first_present(row.get("Kod towaru"), code)
        generated_name_without_code = build_product_name(original_name, features, brand, "")
        generated_name_without_code = enforce_catalog_name_style(generated_name_without_code, row, features)
        generated_name = append_code_to_name(generated_name_without_code, code)
        images = "|".join(offer.get("images", []))
        description_text = offer.get("description_text", "")
        feature_payload = build_feature_payload(row, offer, brand, manufacturer_data_by_producer)
        category, recommendations = recommend_store_category(row, generated_name, feature_payload, recommender)
        feature_payload["Kategoria import"] = category
        baselinker_feature_payload, removed_features = filter_baselinker_features(feature_payload, allowed_feature_specs)
        for removed_feature in removed_features:
            removed_feature_names[removed_feature] = removed_feature_names.get(removed_feature, 0) + 1
        feature_payload_json = json.dumps(baselinker_feature_payload, ensure_ascii=False, separators=(",", ":"))

        match_status = "matched_xml" if offer else "missing_xml"
        if not offer:
            missing_codes.append({"code": code, "name": original_name})

        enriched = {
            **row,
            "Kod producenta": code,
            "Dopasowanie XML": match_status,
            "Nazwa SEO bez kodu": generated_name_without_code,
            "Nazwa wygenerowana": generated_name,
            "Nazwa oryginalna XML": offer.get("name", ""),
            "Baselinker product_id": offer.get("product_id", ""),
            "SKU": sku,
            "EAN": ean,
            "Producent import": producer,
            "Marka import": brand,
            "Kategoria import": category,
            "proponowana_kategoria_1": recommendations[0]["category"] if recommendations else "",
            "category_score_1": recommendations[0]["score"] if recommendations else "",
            "category_reasons_1": recommendations[0]["reasons"] if recommendations else "",
            "category_recommendation_status": "OK" if recommendations else "FALLBACK",
            "Kategoria XML": offer.get("category", ""),
            "Cena brutto XML": offer.get("price", ""),
            "Cena netto XML": offer.get("price_net", ""),
            "Stan XML": offer.get("avail", ""),
            "Dostępność XML": offer.get("availability", ""),
            "URL produktu": offer.get("url", ""),
            "Opis producenta": description_text,
            "Opis HTML producenta": offer.get("description_html", ""),
            "Obrazki": images,
            "name": generated_name,
            "sku": sku,
            "manufacturer_name": producer,
            "category": category,
            "features": feature_payload_json,
            "features_json": feature_payload_json,
            "features_opisowe_json": json.dumps(feature_payload, ensure_ascii=False, separators=(",", ":")),
            "Pominięte parametry Baselinker": ", ".join(removed_features),
        }
        for key, value in feature_payload.items():
            enriched[f"Parametr: {key}"] = value
        for key, value in baselinker_feature_payload.items():
            enriched[f"Baselinker parametr: {key}"] = value
        description_row = {**enriched, "features": json.dumps(feature_payload, ensure_ascii=False, separators=(",", ":"))}
        generated_description = build_description_html(pd.Series(description_row), title_column="name")
        enriched["Opis"] = html_to_text(generated_description)
        enriched["Opis HTML"] = generated_description
        enriched_rows.append(enriched)

        baselinker_rows.append(
            {
                "product_id": offer.get("product_id", ""),
                "name": generated_name,
                "sku": sku,
                "ean": ean,
                "manufacturer_name": producer,
                "category": category,
                "description": generated_description,
                "features": feature_payload_json,
                "images_urls": images,
            }
        )

    enriched_df = pd.DataFrame(enriched_rows)
    if "Nazwa SEO bez kodu" in enriched_df.columns:
        title_keys = enriched_df["Nazwa SEO bez kodu"].map(normalize_for_compare)
        duplicate_counts = title_keys.value_counts()
        enriched_df["Duplikat nazwy SEO bez kodu"] = title_keys.map(
            lambda key: f"TAK ({int(duplicate_counts.get(key, 0))})" if key and int(duplicate_counts.get(key, 0)) > 1 else ""
        )
    report = {
        "rows": len(enriched_rows),
        "matched_xml": len(enriched_rows) - len(missing_codes),
        "missing_xml": len(missing_codes),
        "missing_codes": missing_codes,
        "duplicate_seo_titles_without_code": int((enriched_df.get("Duplikat nazwy SEO bez kodu", "") != "").sum()),
        "removed_baselinker_features": removed_feature_names,
        "output_columns": list(enriched_df.columns),
    }
    return enriched_df, baselinker_rows, report


def build_product_name(original_name: str, features: dict[str, str], brand: str, code: str = "") -> str:
    title_brand = infer_brand_from_title(original_name)
    if title_brand and (not brand or brand == "Kobi"):
        brand = title_brand
    base = strip_trailing_brand(compact_spaces(original_name))
    specs: list[str] = []

    add_spec_if_missing_kind(specs, format_power(features.get("Moc [W]", "")), base, r"\b\d+(?:[./,-]\d+)*\s*W\b")
    add_spec_if_missing_kind(
        specs,
        format_cct(features.get("Temperatura barwowa |K|", ""), base),
        base,
        r"\b(?:\d{4}\s*K|[23]CCT)\b",
    )
    add_spec_if_missing_kind(specs, format_lumens(features.get("Strumień świetlny |lm|", "")), base, r"\blm\b")
    add_spec_if_missing_kind(specs, format_ip(features.get("IP", "")), base, r"\bIP\s*\d{2}\b")

    if should_include_angle(base):
        add_spec_if_missing_kind(specs, format_angle(features.get("Kąt świecenia [°]", "")), base, r"\d+\s*[x/,-]?\s*\d*\s*°")

    add_spec_if_missing_kind(specs, format_socket(features.get("Trzonek", "")), base, r"\b(?:E27|E14|GU10|G9|GX53|T8|G13)\b")
    add_spec_if_missing_kind(specs, format_sensor(features.get("Typ czujnika", "")), base, r"\bczuj")
    add_spec_if_missing_kind(
        specs,
        format_shape_for_title(infer_shape(original_name, features), base),
        base,
        r"\b(?:kwadrat\w*|prostok[ąa]t\w*|okr[ąa]g\w*)\b",
    )
    add_spec_if_missing_kind(
        specs,
        format_color_for_title(features.get("Kolor produktu", ""), base),
        base,
        r"\b(?:biały|bialy|biała|biala|czarny|czarna|szary|szara|chrom)\b",
    )

    parts = [base, *specs]
    if brand and not contains_token(" ".join(parts), brand):
        parts.append(brand)

    return append_code_to_name(final_cleanup(" ".join(part for part in parts if part)), code)


def enforce_catalog_name_style(name: str, source_row: dict[str, Any], features: dict[str, str]) -> str:
    name = compact_spaces(name)
    source_category = normalize_for_compare(compact_spaces(source_row.get("Kategoria", "")))
    source_subcategory = normalize_for_compare(compact_spaces(source_row.get("Podkategoria", "")))

    if "highbay" in source_category or "oswietleniehal" in source_category:
        name = normalize_high_bay_name(name)
    elif "oprawyprzemyslowe" in source_category or "uliczne" in source_subcategory:
        name = normalize_street_light_name(name)
    elif "oprawydomowe" in source_category:
        name = normalize_garden_post_name(name)
    elif "oswietleniesolarne" in source_category:
        name = normalize_solar_lighting_name(name)
    elif "oprawypyloszczelne" in source_category or "hermetyczne" in source_category:
        name = normalize_hermetic_name(name)
    elif "plafonieryioprawysufitowe" in source_category:
        name = normalize_plafon_name(name, features)

    return final_cleanup(name)


def normalize_high_bay_name(name: str) -> str:
    name = re.sub(r"(?i)^High\s+Bay\s+LED\s+", "Oprawa High Bay LED ", name)
    name = re.sub(r"(?i)^Oprawa\s+LED\s+HIGH\s+BAY\s+", "Oprawa High Bay LED ", name)
    name = re.sub(r"(?i)^Uchwyt\s+do\s+LED\s+", "Uchwyt do oprawy High Bay LED ", name)
    name = re.sub(r"(?i)^Uniwersalny\s+uchwyt\s+do\s+High\s+Bay\s+", "Uchwyt uniwersalny do oprawy High Bay ", name)
    name = re.sub(r"(?i)^Osłona\s+do\s+LED\s+", "Osłona do oprawy High Bay LED ", name)
    name = re.sub(r"(?i)^Osłona\s+do\s+NEO\s+", "Osłona do oprawy High Bay LED NEO ", name)
    return name


def normalize_street_light_name(name: str) -> str:
    name = re.sub(r"(?i)^Oprawa\s+drogowa\s+Solar\s+LED\s+", "Oprawa uliczna solarna LED ", name)
    name = re.sub(r"(?i)^Oprawa\s+drogowa\s+LED\s+", "Oprawa uliczna LED ", name)
    return name


def normalize_garden_post_name(name: str) -> str:
    name = re.sub(r"(?i)^Słup\s+", "Słupek ogrodowy ", name)
    name = re.sub(r"(?i)^Moduł\s+BASE\s+", "Moduł do słupka ogrodowego BASE ", name)
    name = re.sub(r"(?i)^Przedłużka\s+BASE\s+", "Przedłużka do słupka ogrodowego BASE ", name)
    return name


def normalize_solar_lighting_name(name: str) -> str:
    name = normalize_street_light_name(name)
    name = re.sub(r"(?i)^Oprawa\s+parkowa\s+HYBRID\s+LED\s+", "Lampa ogrodowa parkowa LED HYBRID ", name)
    return name


def normalize_hermetic_name(name: str) -> str:
    name = re.sub(r"(?i)^Oprawa\s+linowa\s+LED\s+", "Oprawa hermetyczna LED ", name)
    name = re.sub(r"(?i)^Oprawa\s+liniowa\s+hermetyczna\s+LED\s+", "Oprawa hermetyczna LED ", name)
    name = re.sub(r"(?i)^Oprawa\s+liniowa\s+LED\s+", "Oprawa hermetyczna LED ", name)
    name = re.sub(
        r"(?i)^Zestaw\s+HERMETIC\s+(.+?)\s*\+\s*(\d+x\s*)?LED\s+T8\s+",
        lambda match: (
            f"Lampa hermetyczna HERMETIC {match.group(1)} "
            f"{'z ' + match.group(2).strip().rstrip('xX') + ' świetlówkami' if match.group(2) else 'ze świetlówką'} LED T8 "
        ),
        name,
    )
    return name


def normalize_plafon_name(name: str, features: dict[str, str]) -> str:
    has_integrated_led = (
        compact_spaces(features.get("Moc [W]", ""))
        or compact_spaces(features.get("Strumień świetlny |lm|", ""))
        or bool(re.match(r"(?i)^Plafon\s+LED\s+", name))
    )
    if has_integrated_led:
        name = re.sub(r"(?i)^Plafoniera\s+LED\s+", "Plafon LED ", name)
    return name


def build_feature_payload(
    source_row: dict[str, Any],
    offer: dict[str, Any],
    brand: str,
    manufacturer_data_by_producer: dict[str, str],
) -> dict[str, str]:
    payload: dict[str, str] = {}
    attrs = offer.get("attrs", {})
    features = offer.get("features", {})

    payload["Producent"] = normalize_manufacturer(first_present(attrs.get("Producent"), source_row.get("Producent")))
    if brand:
        payload["Marka"] = brand
    product_type = infer_product_type(first_present(source_row.get("Nazwa producenta"), offer.get("name")))
    if product_type:
        payload["Typ produktu"] = product_type

    for key in ["EAN", "Kod_producenta", "Producent odpowiedzialny", "Podmiot odpowiedzialny"]:
        value = normalize_feature_value(attrs.get(key, ""))
        if value:
            target = FEATURE_NAME_MAP.get(key, key)
            if target == "Dane producenta" and payload.get(target):
                continue
            payload[target] = value

    catalog_manufacturer_data = manufacturer_data_for_producer(
        manufacturer_data_by_producer,
        payload.get("Producent", ""),
        attrs.get("Producent", ""),
        source_row.get("Producent", ""),
        features.get("Marka", ""),
        brand,
    )
    if catalog_manufacturer_data:
        payload["Dane producenta"] = catalog_manufacturer_data

    if offer.get("category"):
        payload["Kategoria dostawcy"] = offer["category"]
    if offer.get("url"):
        payload["URL dostawcy"] = offer["url"]

    for key, value in features.items():
        clean_value = normalize_feature_value(value)
        if not clean_value:
            continue
        payload[FEATURE_NAME_MAP.get(key, key)] = clean_value

    add_title_derived_features(payload, first_present(source_row.get("Nazwa producenta"), offer.get("name")))
    return {key: value for key, value in payload.items() if value}


def recommend_store_category(
    source_row: dict[str, Any],
    generated_name: str,
    features: dict[str, str],
    recommender: CategoryRecommender,
) -> tuple[str, list[dict[str, Any]]]:
    recommendation_row = {
        **features,
        "name": generated_name,
        "sku": first_present(source_row.get("Kod towaru"), source_row.get("Kod")),
        "manufacturer_name": normalize_manufacturer(first_present(source_row.get("Producent"), features.get("Producent"))),
        "Kategoria": compact_spaces(source_row.get("Kategoria", "")),
        "Podkategoria": compact_spaces(source_row.get("Podkategoria", "")),
        "Seria": compact_spaces(source_row.get("Marka / Rodzina", "")),
    }
    recommendation_df = pd.DataFrame([recommendation_row])
    recommendations = recommender.recommend(
        recommendation_df.iloc[0],
        detect_recommendation_columns(recommendation_df),
        3,
    )
    guarded_category = guarded_store_category(source_row, generated_name)
    if guarded_category:
        return guarded_category, recommendations
    recommended = recommendations[0]["category"] if recommendations else ""
    return recommended or fallback_store_category(source_row, generated_name), recommendations


def guarded_store_category(source_row: dict[str, Any], generated_name: str) -> str:
    source_category = normalize_for_compare(compact_spaces(source_row.get("Kategoria", "")))
    source_subcategory = normalize_for_compare(compact_spaces(source_row.get("Podkategoria", "")))
    title = normalize_for_compare(generated_name)

    if "paneleled" in source_category:
        return "Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Panele LED sufitowe"
    if "oprawypyloszczelne" in source_category or "hermetyczne" in source_category:
        return "Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Lampy hermetyczne"
    if "highbay" in source_category or "oswietleniehal" in source_category:
        return "Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Oprawy High-bay LED"
    if "naswietlaczeled" in source_category:
        return "Oświetlenie > Oświetlenie zewnętrzne > Naświetlacze LED"
    if "plafonieryioprawysufitowe" in source_category:
        return "Oświetlenie > Oświetlenie wewnętrzne > Plafony LED"
    if "oprawyprzemyslowe" in source_category or "uliczne" in source_subcategory or "oprawadrogowa" in title:
        return "Oświetlenie > Oświetlenie zewnętrzne > Lampy uliczne"
    if "zrodlaswiatla" in source_category or "swietlowki" in source_subcategory:
        return "Żarówki > Świetlówki LED"
    if "oprawydomowe" in source_category:
        return "Oświetlenie > Oświetlenie zewnętrzne > Lampy ogrodowe > Słupki ogrodowe"
    if "oswietleniesolarne" in source_category:
        if "oprawadrogowa" in title or "ulicz" in source_subcategory:
            return "Oświetlenie > Oświetlenie zewnętrzne > Lampy uliczne"
        if "naswietlacz" in title:
            return "Oświetlenie > Oświetlenie zewnętrzne > Naświetlacze LED"
        return "Oświetlenie > Oświetlenie zewnętrzne > Lampy ogrodowe"
    return ""


def fallback_store_category(source_row: dict[str, Any], generated_name: str) -> str:
    source_category = normalize_for_compare(compact_spaces(source_row.get("Kategoria", "")))
    source_subcategory = normalize_for_compare(compact_spaces(source_row.get("Podkategoria", "")))
    title = normalize_for_compare(generated_name)
    haystack = f"{source_category} {source_subcategory} {title}"

    if "oprawadrogowa" in title or "uliczne" in source_subcategory:
        return "Oświetlenie > Oświetlenie zewnętrzne > Lampy uliczne"
    if "slup" in title or "slupek" in title or "base" in title:
        return "Oświetlenie > Oświetlenie zewnętrzne > Lampy ogrodowe > Słupki ogrodowe"
    if "swietlowka" in title or "zrodlaswiatla" in source_category:
        return "Żarówki > Świetlówki LED"
    if "panel" in haystack:
        return "Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Panele LED sufitowe"
    if "hermetycz" in haystack or "pyloszczel" in haystack:
        return "Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Lampy hermetyczne"
    if "highbay" in haystack:
        return "Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Oprawy High-bay LED"
    if "naswietlacz" in haystack:
        return "Oświetlenie > Oświetlenie zewnętrzne > Naświetlacze LED"
    if "plafon" in haystack or "oprawysufitowe" in haystack:
        return "Oświetlenie > Oświetlenie wewnętrzne > Plafony LED"
    return "Oświetlenie"


def add_title_derived_features(payload: dict[str, str], title: str) -> None:
    if title and "Kształt" not in payload:
        shape = infer_shape(title, payload)
        if shape:
            payload["Kształt"] = format_shape_for_title(shape, title)
    if title and "Wymiary" not in payload:
        dimensions = find_dimensions(title)
        if dimensions:
            payload["Wymiary"] = dimensions
    if title and "UGR" not in payload:
        ugr = re.search(r"\bUGR\s*<?\s*(\d+)", title, flags=re.IGNORECASE)
        if ugr:
            payload["UGR"] = f"<{ugr.group(1)}" if "<" in ugr.group(0) else ugr.group(1)


def build_category_path(row: dict[str, Any]) -> str:
    category = compact_spaces(row.get("Kategoria", ""))
    subcategory = compact_spaces(row.get("Podkategoria", ""))
    if category and subcategory:
        return f"{category}/{subcategory}"
    return category or subcategory


def infer_product_type(title: str) -> str:
    text = normalize_for_compare(title)
    ordered_types = [
        ("zestawhermetic", "Lampa hermetyczna ze świetlówką LED T8"),
        ("oprawadrogowasolar", "Oprawa drogowa solarna LED"),
        ("oprawadrogowa", "Oprawa drogowa LED"),
        ("naswietlaczsolar", "Naświetlacz solarny LED"),
        ("naswietlacz", "Naświetlacz LED"),
        ("highbay", "Oprawa High Bay LED"),
        ("oprawahermetyczna", "Oprawa hermetyczna LED"),
        ("oprawaliniowahermetyczna", "Oprawa hermetyczna LED"),
        ("oprawaliniowa", "Oprawa liniowa LED"),
        ("panelled", "Panel LED"),
        ("plafoniera", "Plafon LED"),
        ("plafon", "Plafon LED"),
        ("swietlowka", "Świetlówka LED"),
        ("zasilacz", "Zasilacz LED"),
        ("ramka", "Ramka"),
        ("uchwyt", "Uchwyt"),
        ("linka", "Linka"),
        ("slup", "Słupek ogrodowy"),
        ("przedluzka", "Przedłużka"),
        ("modul", "Moduł"),
    ]
    for needle, product_type in ordered_types:
        if needle in text:
            return product_type
    return ""


def write_xlsx(df: pd.DataFrame, path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Kobi import")
        worksheet = writer.sheets["Kobi import"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column_cells in worksheet.columns:
            header = str(column_cells[0].value or "")
            width = min(max(len(header) + 2, 12), 60)
            worksheet.column_dimensions[column_cells[0].column_letter].width = width


def write_baselinker_csv(rows: list[dict[str, str]], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=BASELINKER_COLUMNS, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(rows)


def write_report(report: dict[str, Any], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


def strip_trailing_brand(value: str) -> str:
    result = compact_spaces(value)
    result = re.sub(r"\s*-\s*nowy model\s*$", "", result, flags=re.IGNORECASE)
    for brand in TRAILING_BRANDS:
        result = re.sub(rf"\s+{re.escape(brand)}\s*$", "", result, flags=re.IGNORECASE)
    return compact_spaces(result)


def infer_brand_from_title(value: str) -> str:
    text = compact_spaces(value)
    for brand in TRAILING_BRANDS:
        if re.search(rf"\s+{re.escape(brand)}(?:\s*-\s*nowy model)?\s*$", text, flags=re.IGNORECASE):
            return normalize_brand(brand)
    return ""


def add_spec(specs: list[str], value: str, existing_text: str) -> None:
    value = compact_spaces(value)
    if not value or contains_token(existing_text, value):
        return
    if any(contains_token(existing, value) for existing in specs):
        return
    specs.append(value)


def add_spec_if_missing_kind(specs: list[str], value: str, existing_text: str, existing_pattern: str) -> None:
    value = compact_spaces(value)
    if not value:
        return
    combined = " ".join([existing_text, *specs])
    if re.search(existing_pattern, combined, flags=re.IGNORECASE):
        return
    add_spec(specs, value, existing_text)


def contains_token(text: str, token: str) -> bool:
    if not text or not token:
        return False
    return normalize_for_compare(token) in normalize_for_compare(text)


def normalize_for_compare(value: str) -> str:
    value = compact_spaces(value).lower()
    replacements = {
        "ą": "a",
        "ć": "c",
        "ę": "e",
        "ł": "l",
        "ń": "n",
        "ó": "o",
        "ś": "s",
        "ż": "z",
        "ź": "z",
        "°": "",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = re.sub(r"(?<=\d)\s+(?=[wkv]|lm|cm|mm)", "", value)
    return re.sub(r"[^a-z0-9<]+", "", value)


def format_power(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    value = normalize_number_list(value)
    return value if re.search(r"w$", value, flags=re.IGNORECASE) else f"{value}W"


def format_cct(value: str, base: str) -> str:
    if re.search(r"\b[23]CCT\b", base, flags=re.IGNORECASE):
        return ""
    value = compact_spaces(value)
    if not value:
        return ""
    numbers = re.findall(r"\d{4}", value)
    if len(numbers) >= 2:
        return f"{min(numbers)}-{max(numbers)}K"
    if numbers:
        return f"{numbers[0]}K"
    return value if value.upper().endswith("K") else f"{value}K"


def format_lumens(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    value = normalize_number_list(value)
    return value if re.search(r"lm$", value, flags=re.IGNORECASE) else f"{value}lm"


def format_ip(value: str) -> str:
    value = compact_spaces(value).upper()
    if not value:
        return ""
    match = re.search(r"IP\s*\d{2}", value)
    return match.group(0).replace(" ", "") if match else value


def format_angle(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    value = normalize_number_list(value)
    return value if "°" in value else f"{value}°"


def format_socket(value: str) -> str:
    value = compact_spaces(value).upper()
    return value if value and value != "NIE DOTYCZY" else ""


def format_sensor(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    if re.search(r"czuj", value, flags=re.IGNORECASE):
        return value
    return f"czujnik {value}"


def should_include_angle(base: str) -> bool:
    return bool(re.search(r"\b(high bay|naświetlacz|oprawa drogowa|uliczna)\b", base, flags=re.IGNORECASE))


def normalize_color(value: str) -> str:
    value = compact_spaces(value).lower()
    return COLOR_FORMS.get(value, value)


def format_color_for_title(value: str, title_context: str) -> str:
    color = normalize_color(value)
    if not color:
        return ""
    gender = infer_title_gender(title_context)
    return COLOR_INFLECTIONS.get(color, {}).get(gender, color)


def infer_shape(title: str, features: dict[str, str] | None = None) -> str:
    text = compact_spaces(title)
    normalized = strip_accents(text).lower()
    if re.search(r"\bokr[ąa]g\w*|\bround\b", normalized):
        return "okrągły"
    if re.search(r"\bkwadrat\w*|\bsquare\b", normalized):
        return "kwadratowy"
    if re.search(r"\bprostok[ąa]t\w*|\brectang", normalized):
        return "prostokątny"

    shape_sensitive = re.search(
        r"\b(panel|plafon|downlight|oprawa\s+sufitowa|oprawa\s+podtynkowa)\b",
        normalized,
    )
    if not shape_sensitive:
        return ""

    dimensions_source = " ".join(
        [
            text,
            compact_spaces((features or {}).get("Wymiary", "")),
            compact_spaces((features or {}).get("Wymiar", "")),
        ]
    )
    match = re.search(r"\b(\d{2,4})\s*x\s*(\d{2,4})(?:\s*x\s*\d{1,4})?", dimensions_source, flags=re.IGNORECASE)
    if not match:
        return ""
    first = int(match.group(1))
    second = int(match.group(2))
    if first == second:
        return "kwadratowy"
    return "prostokątny"


def format_shape_for_title(shape: str, title_context: str) -> str:
    shape = compact_spaces(shape).lower()
    if not shape:
        return ""
    gender = infer_title_gender(title_context)
    return SHAPE_INFLECTIONS.get(shape, {}).get(gender, shape)


def infer_title_gender(title_context: str) -> str:
    text = strip_accents(compact_spaces(title_context)).lower()
    feminine_terms = [
        "oprawa",
        "lampa",
        "swietlowka",
        "przedluzka",
        "ramka",
        "linka",
        "oslona",
        "girlanda",
    ]
    masculine_terms = [
        "panel",
        "plafon",
        "naswietlacz",
        "zasilacz",
        "uchwyt",
        "czujnik",
        "slupek",
        "modul",
        "klips",
    ]
    neuter_terms = ["gniazdo"]
    for term in feminine_terms:
        if re.search(rf"\b{re.escape(term)}\b", text):
            return "feminine"
    for term in masculine_terms:
        if re.search(rf"\b{re.escape(term)}\b", text):
            return "masculine"
    for term in neuter_terms:
        if re.search(rf"\b{re.escape(term)}\b", text):
            return "neuter"
    return "masculine"


def normalize_brand(value: str) -> str:
    value = compact_spaces(value)
    normalized = normalize_alias_key(value)
    aliases = {
        "KOBI PROFESSIONAL": "Kobi Pro",
        "KOBI LIGHT": "Kobi",
        "KOBI": "Kobi",
        "LED2B RED": "Kobi",
        "RED": "Kobi",
    }
    return aliases.get(normalized, value)


def normalize_manufacturer(value: str) -> str:
    value = compact_spaces(value)
    if not value or normalize_alias_key(value) in KOBI_MANUFACTURER_ALIASES:
        return "Kobi"
    return value


def normalize_alias_key(value: str) -> str:
    value = compact_spaces(value).upper()
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    return compact_spaces(value)


def normalize_feature_value(value: Any) -> str:
    value = compact_spaces(value)
    value = re.sub(r"\s*,\s*", ", ", value)
    value = re.sub(r"\s*\|\s*", "|", value)
    return value


def normalize_number_list(value: str) -> str:
    value = compact_spaces(value).replace(",", ".")
    value = re.sub(r"\s*/\s*", "/", value)
    value = re.sub(r"\s*-\s*", "-", value)
    value = re.sub(r"\.0\b", "", value)
    return value


def append_code_to_name(name: str, code: str) -> str:
    name = compact_spaces(name)
    code = normalize_code(code)
    if not code or contains_token(name, code):
        return name
    brand = infer_brand_from_title(name)
    if brand and re.search(rf"\b{re.escape(brand)}\s*$", name, flags=re.IGNORECASE):
        return compact_spaces(re.sub(rf"\b{re.escape(brand)}\s*$", f"{brand} {code}", name, flags=re.IGNORECASE))
    return compact_spaces(f"{name} {code}")


def find_dimensions(value: str) -> str:
    patterns = [
        r"\b\d{2,4}\s*x\s*\d{2,4}(?:\s*x\s*\d{1,4})?\s*(?:mm|cm)?\b",
        r"\b\d{2,4}\s*cm\b",
        r"\b(?:fi|ø)\s*\d{1,4}\s*mm\b",
    ]
    matches: list[str] = []
    for pattern in patterns:
        matches.extend(re.findall(pattern, value, flags=re.IGNORECASE))
    cleaned = [compact_spaces(match).replace(" x ", "x") for match in matches]
    return ", ".join(dict.fromkeys(cleaned))


def compact_html(value: str) -> str:
    value = str(value or "").replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def html_to_text(value: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", str(value or ""))
    text = re.sub(r"(?i)<br\s*/?>|</p>|</li>|</h[1-6]>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    return compact_spaces(unescape(text))


def first_present(*values: Any) -> str:
    for value in values:
        if value is None:
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        text = compact_spaces(value)
        if text:
            return text
    return ""


def normalize_code(value: Any) -> str:
    text = compact_spaces(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    if "/" in text:
        text = text.split("/", 1)[0]
    digits = re.sub(r"\D", "", text)
    return digits.zfill(6) if digits else ""


def only_digits(value: Any) -> str:
    return re.sub(r"\D", "", compact_spaces(value))


def compact_spaces(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def final_cleanup(value: str) -> str:
    value = compact_spaces(value)
    value = re.sub(r"\s+([,/])", r"\1", value)
    value = normalize_unit_spacing(value)
    value = re.sub(r"\bIP\s+(\d{2})\b", r"IP\1", value, flags=re.IGNORECASE)
    value = normalize_color_adjectives(value)
    value = re.sub(r"(?i)\bLED LED\b", "LED", value)
    value = re.sub(r"\s{2,}", " ", value)
    return value.strip()


def normalize_unit_spacing(value: str) -> str:
    units = "lm/W|mAh|mm2|mm²|mA|kA|Wh|Wp|lm|mm|cm|Hz|W|V|K|A|J|m|°"
    return re.sub(rf"(?<=\d)\s+(?=(?:{units})\b)", "", value, flags=re.IGNORECASE)


def normalize_color_adjectives(value: str) -> str:
    gender = infer_title_gender(value)
    replacements = [
        ("biały", r"\b(?:biały|biała|białe|bialy|biala|biale)\b"),
        ("czarny", r"\b(?:czarny|czarna|czarne)\b"),
        ("szary", r"\b(?:szary|szara|szare)\b"),
        ("grafitowy", r"\b(?:grafitowy|grafitowa|grafitowe)\b"),
        ("srebrny", r"\b(?:srebrny|srebrna|srebrne)\b"),
        ("złoty", r"\b(?:złoty|złota|złote|zloty|zlota|zlote)\b"),
    ]
    for canonical, pattern in replacements:
        replacement = COLOR_INFLECTIONS.get(canonical, {}).get(gender, canonical)
        value = re.sub(pattern, replacement, value, flags=re.IGNORECASE)
    return value


if __name__ == "__main__":
    main()
