from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, enrich_config_with_catalog_knowledge, load_catalog_knowledge
from title_anatomy import (
    DEFAULT_TITLE_ANATOMY_PATH,
    build_title_from_anatomy,
    load_title_anatomy_config,
    normalize_product_type_key,
)
from utils import compact_spaces, is_blank, load_yaml, read_products, strip_accents, write_products


def build_title(row: pd.Series, config: dict[str, Any]) -> tuple[str, list[str]]:
    _, title, warnings, _ = build_title_parts(row, config)
    return title, warnings


def build_title_parts(
    row: pd.Series,
    config: dict[str, Any],
    anatomy_config: dict[str, Any] | None = None,
) -> tuple[str, str, list[str], dict[str, Any]]:
    product_role = str(row.get("product_role", "main") or "main")
    template, matched_rule = resolve_title_template_and_rule(row, config, product_role)
    max_length = int(config.get("max_title_length", 110))
    forbidden = [str(word) for word in config.get("forbidden_title_words") or []]
    values: dict[str, str] = {}

    for key, value in row.items():
        if str(key).startswith("attr_"):
            values[str(key)[5:]] = "" if is_blank(value) else str(value)
    values["old_title"] = str(row.get("old_title", ""))
    values["nazwa_kanlux"] = first_non_blank(row.get("Nazwa Kanlux", ""), row.get("Nazwa producenta", ""))
    values["moc_zakres"] = first_non_blank(row.get("Moc - zakres", ""), row.get("attr_moc_zakres", ""))
    source_title_parts: list[str] = []
    for column in ["Nazwa B2C / SEO", "Nazwa", "Nazwa Kanlux", "name", "old_title"]:
        value = compact_spaces(str(row.get(column, "")))
        if value and value not in source_title_parts:
            source_title_parts.append(value)
    values["source_title"] = " ".join(source_title_parts)
    values["product_role"] = product_role
    values["accessory_type"] = "" if is_blank(row.get("accessory_type", "")) else str(row.get("accessory_type", ""))
    values["sku"] = resolve_sku_value(row, config)
    values.update(build_inflected_values(values, config))

    anatomy_result = build_title_from_anatomy(row, anatomy_config or {}, values, max_length=max_length)
    if anatomy_result is not None:
        metadata = {
            "title_source": anatomy_result.source,
            "title_anatomy_rule": anatomy_result.rule_key,
            "title_anatomy_used_attributes": ",".join(anatomy_result.used_attributes),
            "title_anatomy_skipped_attributes": ",".join(anatomy_result.skipped_attributes),
            "title_anatomy_product_type_key": normalize_product_type_key(values.get("typ", row.get("attr_typ", row.get("Typ", "")))),
            "title_duplicate_disambiguation": "",
        }
        # Te same normalizacje SEO co dla legacy (np. Plafoniera -> Plafon).
        seo_without_code = uppercase_first_letter(normalize_seo_title_terms(anatomy_result.title_without_code))
        final_title = uppercase_first_letter(normalize_seo_title_terms(anatomy_result.title))
        return seo_without_code, final_title, anatomy_result.warnings, metadata

    title = template
    for field in re.findall(r"{([^{}]+)}", template):
        title = title.replace("{" + field + "}", values.get(field, ""))

    for word in forbidden:
        title = re.sub(rf"\b{re.escape(word)}\b", " ", title, flags=re.IGNORECASE)

    seo_title_without_code = uppercase_first_letter(normalize_seo_title_terms(compact_spaces(title)))
    title = seo_title_without_code
    if config.get("append_sku_after_producer"):
        title = append_sku_after_producer(title, values.get("producent", ""), values.get("sku", ""))
    title = uppercase_first_letter(title)
    warnings: list[str] = []
    if len(title) > max_length:
        warnings.append(f"title_too_long:{len(title)}>{max_length}")

    # Priorytety required_fields: regula > rola > globalny
    if matched_rule is not None and matched_rule.get("required_fields") is not None:
        required_fields = list(matched_rule["required_fields"])
    else:
        required_fields = (config.get("required_fields_by_role") or {}).get(product_role, config.get("required_fields") or [])

    for required in required_fields:
        if required == "moc":
            # Gwint (np. GU10) zastepuje moc dla opraw bez wlasnego LED
            if not is_blank(values.get("gwint", "")):
                continue
            # Strumien zastepuje moc dla produktow solarnych i CCT bez watazu
            if not is_blank(values.get("strumien", "")):
                continue
        if is_blank(values.get(required, "")):
            warnings.append(f"missing_required:{required}")
    if is_blank(title):
        warnings.append("empty_title")

    title_source = "manual_template" if matched_rule is not None and matched_rule.get("source") == "manual_template" else "legacy_template"
    metadata = {
        "title_source": title_source,
        "title_anatomy_rule": "",
        "title_anatomy_used_attributes": "",
        "title_anatomy_skipped_attributes": "",
        "title_anatomy_product_type_key": normalize_product_type_key(values.get("typ", row.get("attr_typ", row.get("Typ", "")))),
        "title_duplicate_disambiguation": "",
    }
    if title_source == "legacy_template":
        warnings.append("title_anatomy_missing_accepted_rule")

    return seo_title_without_code, title, warnings, metadata


def normalize_seo_title_terms(title: str) -> str:
    title = re.sub(r"\bPlafoniera\b", "Plafon", title, flags=re.IGNORECASE)
    title = re.sub(r"\bPlafon\s+drewniana\b", "Plafon drewniany", title, flags=re.IGNORECASE)
    return compact_spaces(title)


def normalize_title_uniqueness_key(title: str) -> str:
    value = strip_accents(compact_spaces(title)).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return compact_spaces(value)


def duplicate_disambiguation_attributes(config: dict[str, Any]) -> list[str]:
    configured = config.get("duplicate_title_disambiguation_attributes")
    if configured:
        return [str(item) for item in configured]
    return [
        "wariant",
        "czujnik",
        "dlugosc_przewodu",
        "wysokosc",
        "material",
        "moc",
        "srednica",
        "wymiary",
        "kat_swiecenia",
        "model",
    ]


def resolve_duplicate_title_differences(
    df: pd.DataFrame,
    generated_rows: list[tuple[int, str, str, list[str], str, dict[str, Any], str]],
    config: dict[str, Any],
) -> list[tuple[int, str, str, list[str], str, dict[str, Any], str]]:
    rows = [
        {
            "index": index,
            "seo_title_without_code": seo_title_without_code,
            "new_title": new_title,
            "warnings": warnings,
            "title_uniqueness_key": title_uniqueness_key,
            "metadata": metadata,
            "final_title_uniqueness_key": final_title_uniqueness_key,
        }
        for index, seo_title_without_code, new_title, warnings, title_uniqueness_key, metadata, final_title_uniqueness_key in generated_rows
    ]
    rows_by_key: dict[str, list[dict[str, Any]]] = {}
    for item in rows:
        key = str(item["title_uniqueness_key"])
        if key:
            rows_by_key.setdefault(key, []).append(item)

    for group in rows_by_key.values():
        if len(group) <= 1:
            continue
        apply_duplicate_disambiguation(df, group, config)

    return [
        (
            int(item["index"]),
            str(item["seo_title_without_code"]),
            str(item["new_title"]),
            list(item["warnings"]),
            normalize_title_uniqueness_key(str(item["seo_title_without_code"])),
            dict(item["metadata"]),
            normalize_title_uniqueness_key(str(item["new_title"])),
        )
        for item in rows
    ]


def apply_duplicate_disambiguation(df: pd.DataFrame, group: list[dict[str, Any]], config: dict[str, Any]) -> None:
    for attribute in duplicate_disambiguation_attributes(config):
        display_values = {
            item["index"]: duplicate_disambiguation_value(df.loc[item["index"]], attribute)
            for item in group
        }
        nonblank_values = {value for value in display_values.values() if value}
        all_values = {value or "__BLANK__" for value in display_values.values()}
        if not nonblank_values or len(all_values) <= 1:
            continue

        proposed_keys: set[str] = set()
        proposals: dict[int, tuple[str, str]] = {}
        for item in group:
            index = int(item["index"])
            value = display_values[index]
            seo_title = str(item["seo_title_without_code"])
            new_title = str(item["new_title"])
            if value and not title_contains_disambiguator(seo_title, value):
                seo_title = insert_disambiguator_before_producer(seo_title, value, df.loc[index])
                new_title = insert_disambiguator_before_producer(new_title, value, df.loc[index])
            key = normalize_title_uniqueness_key(seo_title)
            if key in proposed_keys:
                break
            proposed_keys.add(key)
            proposals[index] = (seo_title, new_title)
        else:
            if len(proposed_keys) == len(group):
                for item in group:
                    index = int(item["index"])
                    item["seo_title_without_code"], item["new_title"] = proposals[index]
                    item["metadata"] = {**item["metadata"], "title_duplicate_disambiguation": attribute}
                return


def duplicate_disambiguation_value(row: pd.Series, attribute: str) -> str:
    if attribute == "wariant":
        return supplier_variant_for_duplicate(row)
    value = first_non_blank(row.get(f"attr_{attribute}", ""), row.get(attribute, ""))
    value = compact_spaces(str(value))
    if is_blank(value):
        return ""
    if attribute == "dlugosc_przewodu":
        return f"przewód {value}" if not normalize_text(value).startswith("przewod") else value
    if attribute == "wysokosc":
        return f"wys. {value}" if not normalize_text(value).startswith("wys") else value
    return value


def supplier_variant_for_duplicate(row: pd.Series) -> str:
    supplier_name = first_non_blank(row.get("Nazwa Kanlux", ""))
    if is_blank(supplier_name):
        return ""
    value = compact_spaces(str(supplier_name))
    series = first_non_blank(row.get("attr_seria", ""), row.get("Rodzina", ""))
    if series:
        value = re.sub(rf"^{re.escape(str(series))}\b", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bLED\b", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\b\d+(?:[,.]\d+)?\s*W(?:[-/]?(?:NW|WW|CW))?\b", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\b(?:NW|WW|CW)\b", " ", value, flags=re.IGNORECASE)
    return compact_spaces(value)


def title_contains_disambiguator(title: str, value: str) -> bool:
    return normalize_title_uniqueness_key(value) in normalize_title_uniqueness_key(title)


def insert_disambiguator_before_producer(title: str, value: str, row: pd.Series) -> str:
    value = compact_spaces(value)
    if not value:
        return title
    producer = first_non_blank(row.get("attr_producent", ""), row.get("Producent", ""), row.get("producer", ""), "Kanlux")
    sku = resolve_sku_value(row, {"sku_columns": ["Kod", "sku", "SKU"]})
    if producer:
        pattern = rf"\s+({re.escape(producer)})(\s+{re.escape(sku)})?$" if sku else rf"\s+({re.escape(producer)})$"
        replacement = rf" {value} \1\2" if sku else rf" {value} \1"
        updated = re.sub(pattern, replacement, title, count=1, flags=re.IGNORECASE)
        if updated != title:
            return compact_spaces(updated)
    return compact_spaces(f"{title} {value}")


def uppercase_first_letter(value: str) -> str:
    for index, char in enumerate(value):
        if char.isalpha():
            return value[:index] + char.upper() + value[index + 1 :]
    return value


def resolve_sku_value(row: pd.Series, config: dict[str, Any]) -> str:
    for column in config.get("sku_columns") or ["sku", "Kod"]:
        if column in row and not is_blank(row.get(column, "")):
            value = str(row.get(column, "")).strip()
            return value[:-2] if re.fullmatch(r"\d+\.0", value) else value
    return ""


def append_sku_after_producer(title: str, producer: str, sku: str) -> str:
    if is_blank(sku):
        return title
    sku = str(sku).strip()
    if re.search(rf"(?<!\d){re.escape(sku)}(?!\d)", title):
        return title
    if producer and re.search(rf"\b{re.escape(producer)}\b", title, flags=re.IGNORECASE):
        return compact_spaces(re.sub(rf"\b{re.escape(producer)}\b", f"{producer} {sku}", title, count=1, flags=re.IGNORECASE))
    return compact_spaces(f"{title} {sku}")


def resolve_title_template(row: pd.Series, config: dict[str, Any], product_role: str) -> str:
    template, _ = resolve_title_template_and_rule(row, config, product_role)
    return template


def resolve_title_template_and_rule(row: pd.Series, config: dict[str, Any], product_role: str) -> tuple[str, dict | None]:
    sku = resolve_sku_value(row, config)
    manual_templates = config.get("manual_title_templates_by_sku") or {}
    if sku and sku in manual_templates:
        return str(manual_templates[sku]), {"required_fields": [], "source": "manual_template"}
    if product_role != "accessory":
        category = str(row.get("Kategoria", row.get("category", "")))
        product_type = first_non_blank(row.get("attr_typ", ""), row.get("Typ", ""))
        title = str(row.get("old_title", ""))
        for rule in config.get("title_strategy_rules") or []:
            if strategy_rule_matches(rule, category, product_type, title):
                return str(rule.get("template", config.get("title_template", "{old_title}"))), rule
        return str(config.get("title_template", "{old_title}")), None
    accessory_type = str(row.get("accessory_type", ""))
    templates = config.get("accessory_title_templates") or {}
    return str(templates.get(accessory_type, config.get("accessory_title_template", config.get("title_template", "{old_title}")))), None


def first_non_blank(*values: Any) -> str:
    for value in values:
        if not is_blank(value):
            return str(value)
    return ""


def normalize_text(value: str) -> str:
    replacements = str.maketrans(
        "\u0105\u0107\u0119\u0142\u0144\u00f3\u015b\u017a\u017c\u0104\u0106\u0118\u0141\u0143\u00d3\u015a\u0179\u017b",
        "acelnoszzACELNOSZZ"
    )
    return str(value).translate(replacements).lower()


def strategy_rule_matches(rule: dict[str, Any], category: str, product_type: str, title: str) -> bool:
    checks = [
        (rule.get("category_contains"), category),
        (rule.get("type_contains"), product_type),
        (rule.get("title_contains"), title),
    ]
    for needle, haystack in checks:
        if needle and normalize_text(str(needle)) not in normalize_text(str(haystack)):
            return False
    return bool(rule.get("category_contains") or rule.get("type_contains") or rule.get("title_contains"))


def build_inflected_values(values: dict[str, str], config: dict[str, Any]) -> dict[str, str]:
    inflections = config.get("inflections") or {}
    result: dict[str, str] = {}
    for target, mapping in inflections.items():
        if target == "kolor_feminine":
            color = values.get("kolor", "")
            result[target] = inflect_value(color, mapping) if isinstance(mapping, dict) else color
        if target.startswith("ksztalt_"):
            shape = values.get("ksztalt", "")
            result[target] = str(mapping.get(shape, shape)) if isinstance(mapping, dict) else shape
        if target == "kolor_neuter":
            color = values.get("kolor", "")
            result[target] = inflect_value(color, mapping) if isinstance(mapping, dict) else color
    # Tlumaczenie barwy na slowo kluczowe SEO (np. 4000K -> neutralna)
    barwa = values.get("barwa", "")
    if barwa:
        translations = config.get("barwa_translations") or {}
        result["barwa_slownie"] = translations.get(barwa, "")
        result["barwa_swiatla"] = f"światło {result['barwa_slownie']}" if result["barwa_slownie"] else ""
        result["barwa_neuter"] = build_neuter_light_color(result["barwa_slownie"])
    else:
        result["barwa_slownie"] = ""
        result["barwa_swiatla"] = ""
        result["barwa_neuter"] = ""
    # Barwa "pelna" do tytulu: zakres CCT ma priorytet (np. 3000-4000K),
    # w przeciwnym razie punkt + slowo SEO (np. "4000K neutralne").
    barwa_zakres = values.get("barwa_zakres", "")
    if barwa_zakres:
        # Zakres CCT pokazujemy osobnym atrybutem (wczesniej w tytule); nie
        # dublujemy go punktem barwy.
        result["barwa_full"] = ""
    elif barwa:
        result["barwa_full"] = compact_spaces(f"{barwa} {result['barwa_neuter']}")
    else:
        result["barwa_full"] = ""
    pasuje_do = values.get("pasuje_do", "")
    series = values.get("seria", "")
    if pasuje_do and series:
        cleaned_target = re.sub(rf"\b{re.escape(series)}\b", " ", pasuje_do, flags=re.IGNORECASE)
        result["pasuje_do_clean"] = compact_spaces(cleaned_target)
    else:
        result["pasuje_do_clean"] = pasuje_do
    eco = values.get("eco", "")
    if eco and "eco" not in normalize_text(values.get("seria", "")):
        result["eco_clean"] = eco
    else:
        result["eco_clean"] = ""
    old_title_normalized = normalize_text(values.get("old_title", ""))
    result["cct"] = "CCT" if "cct" in old_title_normalized or values.get("barwa_zakres", "") else ""
    result["dim"] = "DIM" if "cctdim" in old_title_normalized or "sciemn" in old_title_normalized else ""
    result["rgb"] = "RGB" if "rgb" in old_title_normalized else ""
    typ_normalized = normalize_text(values.get("typ", ""))
    source_title_normalized = normalize_text(" ".join([values.get("old_title", ""), values.get("source_title", "")]))
    result["punktowa"] = "punktowa" if "punktow" in typ_normalized or "punktow" in source_title_normalized else ""
    supplier_family = supplier_model_family(values)
    if supplier_family:
        result["seria"] = supplier_family
    title_power_range = normalize_title_power_range(values.get("moc_zakres", ""))
    if title_power_range:
        result["moc"] = title_power_range
    elif not values.get("moc", "") and "panel" in typ_normalized:
        panel_max_power = normalize_title_power(values.get("moc_max_zrodla", ""))
        if panel_max_power:
            result["moc"] = panel_max_power
    result["optyka"] = agor_beam_type(values)
    series_clean = re.sub(r"(?<![-\w])LED(?![-\w])", " ", result.get("seria", values.get("seria", "")), flags=re.IGNORECASE)
    result["seria_clean"] = compact_spaces(series_clean) or values.get("seria", "")
    # Sposob montazu panelu wyprowadzony z typu produktu (uniwersalny = natynkowy,
    # zwieszany = zwieszany, pozostale panele = podtynkowy).
    if "panel" in typ_normalized:
        if "aio" in typ_normalized:
            result["montaz"] = "podtynkowy"
        elif "uniwersal" in typ_normalized:
            result["montaz"] = "natynkowy"
        elif "zwieszan" in typ_normalized:
            result["montaz"] = "zwieszany"
        else:
            result["montaz"] = "podtynkowy"
    else:
        result["montaz"] = ""
    return result


def normalize_title_power_range(value: str) -> str:
    value = compact_spaces(str(value or ""))
    if not value:
        return ""
    match = re.fullmatch(
        r"(\d+(?:[,.]\d+)?)\s*(?:W\s*)?(?:-|/|–|—)\s*(\d+(?:[,.]\d+)?)\s*W?",
        value,
        flags=re.IGNORECASE,
    )
    if not match:
        return ""
    left, right = (part.replace(",", ".") for part in match.groups())
    return f"{left}-{right}W"


def normalize_title_power(value: str) -> str:
    value = compact_spaces(str(value or ""))
    if not value:
        return ""
    value = re.sub(r"^(?:max\.?|maks(?:ymalnie)?|do)\s+", "", value, flags=re.IGNORECASE)
    match = re.fullmatch(r"(\d+(?:[,.]\d+)?)\s*W?", value, flags=re.IGNORECASE)
    if not match:
        return ""
    return f"{match.group(1).replace(',', '.')}W"


def supplier_model_family(values: dict[str, str]) -> str:
    producer_name = compact_spaces(values.get("nazwa_kanlux", ""))
    if not producer_name:
        return ""
    normalized = normalize_text(producer_name)
    if not normalized:
        return ""
    tokens = producer_name.split()
    collected: list[str] = []
    for token in tokens:
        cleaned = compact_spaces(token.strip(",;"))
        if not cleaned:
            continue
        prefix = supplier_model_prefix_before_spec(cleaned)
        if prefix:
            collected.append(prefix)
            break
        if supplier_model_token_is_spec(cleaned):
            break
        collected.append(cleaned)
        if len(collected) >= 4:
            break
    candidate = compact_spaces(" ".join(collected))
    candidate = re.sub(r"\bAGOR/A\b", "AGOR", candidate, flags=re.IGNORECASE)
    candidate = re.sub(r"\s+", " ", candidate).strip()
    series = compact_spaces(values.get("seria", ""))
    if not candidate or not series:
        return candidate
    if not supplier_model_candidate_is_supported(candidate, series):
        return ""
    if normalize_text(candidate) == normalize_text(series):
        return ""
    if normalize_text(candidate).startswith(normalize_text(series)):
        return candidate
    return ""


def supplier_model_candidate_is_supported(candidate: str, series: str) -> bool:
    candidate_key = normalize_text(candidate)
    series_key = normalize_text(series)
    if not candidate_key.startswith(series_key):
        return False
    suffix = compact_spaces(candidate[len(series):].strip(" -/"))
    if not suffix:
        return False
    supported_suffixes = {
        "sky",
        "sky-v",
        "pt",
        "pro",
        "hi",
        "aio",
    }
    suffix_key = normalize_text(suffix).replace(" ", "-")
    return suffix_key in supported_suffixes


def supplier_model_prefix_before_spec(token: str) -> str:
    match = re.fullmatch(r"([A-Z]{1,4})\d+(?:[,.]\d+)?(?:-|/)\d+(?:[,.]\d+)?W", token, flags=re.IGNORECASE)
    return match.group(1).upper() if match else ""


def supplier_model_token_is_spec(token: str) -> bool:
    return bool(
        re.fullmatch(r"\d+(?:[,.]\d+)?\s*W(?:[-/]?(?:NW|WW|CW))?", token, flags=re.IGNORECASE)
        or re.fullmatch(r"[A-Z]{1,3}\d+(?:[,.]\d+)?W(?:[-/]?(?:NW|WW|CW))?", token, flags=re.IGNORECASE)
        or bool(re.search(r"\d", token))
        or re.fullmatch(r"\d+(?:[,.]\d+)?(?:-|/)\d+(?:[,.]\d+)?W?", token, flags=re.IGNORECASE)
        or re.fullmatch(r"\d{3,5}(?:-[A-Z]+)?", token, flags=re.IGNORECASE)
        or re.fullmatch(r"\d+(?:[,.]\d+)?MM", token, flags=re.IGNORECASE)
        or re.fullmatch(r"\d+CCT", token, flags=re.IGNORECASE)
        or re.fullmatch(r"(?:NW|WW|CW|CCT|RGB|W|B|SR|G-BR)", token, flags=re.IGNORECASE)
    )


def agor_beam_type(values: dict[str, str]) -> str:
    haystack = " ".join(
        [
            values.get("nazwa_kanlux", ""),
            values.get("source_title", ""),
            values.get("seria", ""),
        ]
    )
    normalized = normalize_text(haystack)
    if "agor" not in normalized:
        return ""
    if re.search(r"\bAGOR/A\b", haystack, flags=re.IGNORECASE):
        return "asymetryczny"
    angle = compact_spaces(values.get("kat_swiecenia", ""))
    if angle:
        angle_number = re.sub(r"[^\d,.]", "", angle).replace(",", ".")
        if angle_number in {"125", "40"}:
            return "asymetryczny"
        if angle_number in {"110", "90"}:
            return "symetryczny"
    if re.search(r"\bFL\s+AGOR\b", haystack, flags=re.IGNORECASE):
        return "symetryczny"
    return ""


def inflect_value(value: str, mapping: dict[str, Any]) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    direct = mapping.get(value)
    if direct:
        return str(direct)
    if "/" not in value:
        return str(mapping.get(value, value))
    parts = [compact_spaces(part) for part in value.split("/") if compact_spaces(part)]
    return " / ".join(str(mapping.get(part, part)) for part in parts)


def build_neuter_light_color(value: str) -> str:
    mapping = {
        "neutralna": "neutralne",
        "ciepłobiała": "ciepłobiałe",
        "chłodnobiała": "chłodnobiałe",
        "cieplobiala": "ciepłobiałe",
        "chlodnobiala": "chłodnobiałe",
    }
    return mapping.get(value, value)


def optimize_titles_for_dataframe(
    df: pd.DataFrame,
    title_column: str,
    config: dict[str, Any],
    anatomy_config: dict[str, Any] | None = None,
) -> pd.DataFrame:
    result = df.copy()
    result["old_title"] = result[title_column].astype(str)
    generated_rows: list[tuple[int, str, str, list[str], str, dict[str, Any], str]] = []

    for index, row in result.iterrows():
        seo_title_without_code, new_title, warnings, metadata = build_title_parts(row, config, anatomy_config)
        title_uniqueness_key = normalize_title_uniqueness_key(seo_title_without_code)
        final_title_uniqueness_key = normalize_title_uniqueness_key(new_title)
        generated_rows.append((index, seo_title_without_code, new_title, warnings, title_uniqueness_key, metadata, final_title_uniqueness_key))

    generated_rows = resolve_duplicate_title_differences(result, generated_rows, config)

    duplicate_counts = pd.Series(
        [item[4] for item in generated_rows if item[4]],
        dtype="object",
    ).value_counts()
    final_duplicate_counts = pd.Series(
        [item[6] for item in generated_rows if item[6]],
        dtype="object",
    ).value_counts()

    for index, seo_title_without_code, new_title, warnings, title_uniqueness_key, metadata, final_title_uniqueness_key in generated_rows:
        old_title = str(result.at[index, "old_title"])
        if title_uniqueness_key and int(duplicate_counts.get(title_uniqueness_key, 0)) > 1:
            warnings = [*warnings, f"duplicate_seo_title_without_code:{int(duplicate_counts[title_uniqueness_key])}"]
        if final_title_uniqueness_key and int(final_duplicate_counts.get(final_title_uniqueness_key, 0)) > 1:
            warnings = [*warnings, f"duplicate_final_title:{int(final_duplicate_counts[final_title_uniqueness_key])}"]
        changed_fields = ["title"] if new_title != old_title else []
        result.at[index, "seo_title_without_code"] = seo_title_without_code
        result.at[index, "title_uniqueness_key"] = title_uniqueness_key
        result.at[index, "final_title_uniqueness_key"] = final_title_uniqueness_key
        result.at[index, "new_title"] = new_title
        result.at[index, "title_status"] = "WARNING" if warnings else "OK"
        result.at[index, "title_length"] = len(new_title)
        result.at[index, "changed_fields"] = ";".join(changed_fields)
        result.at[index, "warnings"] = ";".join(warnings)
        for key, value in metadata.items():
            result.at[index, key] = value

    return result


def build_changes_report(df: pd.DataFrame) -> pd.DataFrame:
    report_columns = [
        "sku",
        "ean",
        "producer",
        "category",
        "old_title",
        "seo_title_without_code",
        "new_title",
        "title_uniqueness_key",
        "final_title_uniqueness_key",
        "title_source",
        "title_anatomy_rule",
        "title_anatomy_used_attributes",
        "title_anatomy_skipped_attributes",
        "title_duplicate_disambiguation",
        "title_status",
        "title_length",
        "new_attributes",
        "source",
        "confidence",
        "category_fit_status",
        "category_fit_reasons",
        "product_role",
        "accessory_type",
        "product_role_reasons",
        "changed_fields",
        "warnings",
    ]
    report = pd.DataFrame()
    for column in report_columns:
        report[column] = df[column] if column in df.columns else ""

    attribute_columns = [column for column in df.columns if column.startswith("attr_")]
    report["new_attributes"] = df[attribute_columns].apply(
        lambda row: json.dumps(
            {col[5:]: val for col, val in row.items() if not is_blank(val)},
            ensure_ascii=False,
            sort_keys=True,
        ),
        axis=1,
    )
    report["source"] = df["attribute_sources"] if "attribute_sources" in df.columns else ""
    report["confidence"] = df["attribute_confidence"] if "attribute_confidence" in df.columns else ""
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Generuje nowe tytuly SEO.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet")
    parser.add_argument("--output", required=True)
    parser.add_argument("--title-column", required=True)
    parser.add_argument("--config", default="configs/categories/oprawy-sufitowe.yaml")
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE_PATH)
    parser.add_argument("--title-anatomy", default=DEFAULT_TITLE_ANATOMY_PATH)
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    config = enrich_config_with_catalog_knowledge(load_yaml(args.config), load_catalog_knowledge(args.catalog_knowledge))
    result = optimize_titles_for_dataframe(df, args.title_column, config, load_title_anatomy_config(args.title_anatomy))
    write_products(result, Path(args.output))


if __name__ == "__main__":
    main()
