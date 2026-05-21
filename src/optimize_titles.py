from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, enrich_config_with_catalog_knowledge, load_catalog_knowledge
from utils import compact_spaces, is_blank, load_yaml, read_products, write_products


def build_title(row: pd.Series, config: dict[str, Any]) -> tuple[str, list[str]]:
    product_role = str(row.get("product_role", "main") or "main")
    template, matched_rule = resolve_title_template_and_rule(row, config, product_role)
    max_length = int(config.get("max_title_length", 110))
    forbidden = [str(word) for word in config.get("forbidden_title_words") or []]
    values: dict[str, str] = {}

    for key, value in row.items():
        if str(key).startswith("attr_"):
            values[str(key)[5:]] = "" if is_blank(value) else str(value)
    values["old_title"] = str(row.get("old_title", ""))
    values["product_role"] = product_role
    values["accessory_type"] = "" if is_blank(row.get("accessory_type", "")) else str(row.get("accessory_type", ""))
    values["sku"] = resolve_sku_value(row, config)
    values.update(build_inflected_values(values, config))

    title = template
    for field in re.findall(r"{([^{}]+)}", template):
        title = title.replace("{" + field + "}", values.get(field, ""))

    for word in forbidden:
        title = re.sub(rf"\b{re.escape(word)}\b", " ", title, flags=re.IGNORECASE)

    title = compact_spaces(title)
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

    return title, warnings


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
        return str(manual_templates[sku]), {"required_fields": []}
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
    series_clean = re.sub(r"(?<![-\w])LED(?![-\w])", " ", values.get("seria", ""), flags=re.IGNORECASE)
    result["seria_clean"] = compact_spaces(series_clean) or values.get("seria", "")
    return result


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


def optimize_titles_for_dataframe(df: pd.DataFrame, title_column: str, config: dict[str, Any]) -> pd.DataFrame:
    result = df.copy()
    result["old_title"] = result[title_column].astype(str)

    for index, row in result.iterrows():
        new_title, warnings = build_title(row, config)
        old_title = str(row.get("old_title", ""))
        changed_fields = ["title"] if new_title != old_title else []
        result.at[index, "new_title"] = new_title
        result.at[index, "title_status"] = "WARNING" if warnings else "OK"
        result.at[index, "title_length"] = len(new_title)
        result.at[index, "changed_fields"] = ";".join(changed_fields)
        result.at[index, "warnings"] = ";".join(warnings)

    return result


def build_changes_report(df: pd.DataFrame) -> pd.DataFrame:
    report_columns = [
        "sku",
        "ean",
        "producer",
        "category",
        "old_title",
        "new_title",
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
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    config = enrich_config_with_catalog_knowledge(load_yaml(args.config), load_catalog_knowledge(args.catalog_knowledge))
    result = optimize_titles_for_dataframe(df, args.title_column, config)
    write_products(result, Path(args.output))


if __name__ == "__main__":
    main()
