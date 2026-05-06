from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, is_blank, load_yaml, read_products, write_products


def build_title(row: pd.Series, config: dict[str, Any]) -> tuple[str, list[str]]:
    product_role = str(row.get("product_role", "main") or "main")
    template = resolve_title_template(row, config, product_role)
    max_length = int(config.get("max_title_length", 110))
    forbidden = [str(word) for word in config.get("forbidden_title_words") or []]
    values: dict[str, str] = {}

    for key, value in row.items():
        if str(key).startswith("attr_"):
            values[str(key)[5:]] = "" if is_blank(value) else str(value)
    values["old_title"] = str(row.get("old_title", ""))
    values["product_role"] = product_role
    values["accessory_type"] = "" if is_blank(row.get("accessory_type", "")) else str(row.get("accessory_type", ""))
    values.update(build_inflected_values(values, config))

    title = template
    for field in re.findall(r"{([^{}]+)}", template):
        title = title.replace("{" + field + "}", values.get(field, ""))

    for word in forbidden:
        title = re.sub(rf"\b{re.escape(word)}\b", " ", title, flags=re.IGNORECASE)

    title = compact_spaces(title)
    title = uppercase_first_letter(title)
    warnings: list[str] = []
    if len(title) > max_length:
        warnings.append(f"title_too_long:{len(title)}>{max_length}")
    required_fields = (config.get("required_fields_by_role") or {}).get(product_role, config.get("required_fields") or [])
    for required in required_fields:
        if required == "moc" and not is_blank(values.get("gwint", "")):
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


def resolve_title_template(row: pd.Series, config: dict[str, Any], product_role: str) -> str:
    if product_role != "accessory":
        category = str(row.get("Kategoria", row.get("category", "")))
        product_type = str(row.get("Typ", ""))
        title = str(row.get("old_title", ""))
        for rule in config.get("title_strategy_rules") or []:
            if strategy_rule_matches(rule, category, product_type, title):
                return str(rule.get("template", config.get("title_template", "{old_title}")))
        return str(config.get("title_template", "{old_title}"))
    accessory_type = str(row.get("accessory_type", ""))
    templates = config.get("accessory_title_templates") or {}
    return str(templates.get(accessory_type, config.get("accessory_title_template", config.get("title_template", "{old_title}"))))


def normalize_text(value: str) -> str:
    replacements = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")
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
            result[target] = str(mapping.get(color, color)) if isinstance(mapping, dict) else color
    return result


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
            {column[5:]: value for column, value in row.items() if not is_blank(value)},
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
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    config = load_yaml(args.config)
    result = optimize_titles_for_dataframe(df, args.title_column, config)
    write_products(result, Path(args.output))


if __name__ == "__main__":
    main()
