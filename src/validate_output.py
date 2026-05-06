from __future__ import annotations

import argparse
import re
from typing import Any

import pandas as pd

from utils import compact_spaces, is_blank, load_yaml, read_products, strip_accents, write_products


def classify_product_roles(df: pd.DataFrame, config: dict[str, Any]) -> pd.DataFrame:
    classification = config.get("product_role_classification") or {}
    name_columns = classification.get("name_columns") or []
    default_role = str(classification.get("default_role", "main"))
    result = df.copy()

    for index, row in result.iterrows():
        searchable_text = build_searchable_text(row, name_columns)
        role = default_role
        accessory_type = ""
        reasons: list[str] = []

        for rule in classification.get("accessory_rules") or []:
            matches = find_matching_terms(searchable_text, rule.get("terms") or [])
            if not matches:
                continue
            role = "accessory"
            accessory_type = str(rule.get("accessory_type", "Akcesorium"))
            reasons.append(f"accessory_terms:{','.join(matches)}")
            break

        clear_accessory_attributes(result, index, accessory_type, config)

        result.at[index, "product_role"] = role
        result.at[index, "accessory_type"] = accessory_type
        result.at[index, "product_role_reasons"] = ";".join(reasons)

    return result


def clear_accessory_attributes(result: pd.DataFrame, index: int, accessory_type: str, config: dict[str, Any]) -> None:
    if not accessory_type:
        return
    policy = config.get("accessory_attribute_policy") or {}
    clear_map = policy.get("clear_attributes") or {}
    for attr in clear_map.get(accessory_type, []):
        column = f"attr_{attr}"
        if column in result.columns:
            result.at[index, column] = ""


def validate_category_fit(df: pd.DataFrame, config: dict[str, Any]) -> tuple[pd.DataFrame, pd.DataFrame]:
    validation = config.get("category_validation") or {}
    rules = validation.get("rules") or []
    category_column = validation.get("category_column", "category")
    name_columns = validation.get("name_columns") or []
    result = df.copy()
    suspicious_rows: list[dict[str, Any]] = []

    for index, row in result.iterrows():
        status = "OK"
        reasons: list[str] = []
        category = str(row.get(category_column, row.get("category", "")))
        searchable_text = build_searchable_text(row, name_columns)

        for rule in rules:
            if not category_matches(category, str(rule.get("match_category_contains", ""))):
                continue

            disallowed = find_matching_terms(searchable_text, rule.get("disallowed_terms") or [])
            if disallowed:
                status = "DO_SPRAWDZENIA"
                reasons.append(f"category_disallowed_terms:{','.join(disallowed)}")

            expected_terms = rule.get("expected_any_terms") or []
            if expected_terms and not find_matching_terms(searchable_text, expected_terms):
                status = "DO_SPRAWDZENIA"
                reasons.append("category_expected_terms_missing")

            missing_attrs = [
                attr for attr in rule.get("suspicious_if_missing_attributes") or [] if is_blank(row.get(f"attr_{attr}", ""))
            ]
            threshold = int(rule.get("suspicious_missing_threshold", len(missing_attrs) + 1))
            if len(missing_attrs) >= threshold:
                status = "DO_SPRAWDZENIA"
                reasons.append(f"category_many_missing_attributes:{','.join(missing_attrs)}")

        result.at[index, "category_fit_status"] = status
        result.at[index, "category_fit_reasons"] = ";".join(reasons)
        if status != "OK":
            suspicious_rows.append(
                {
                    "sku": row.get("sku", row.get("Kod", "")),
                    "ean": row.get("ean", row.get("EAN", "")),
                    "category": category,
                    "title": row.get("old_title", row.get("Nazwa B2C / SEO", row.get("Nazwa Kanlux", ""))),
                    "new_title": row.get("new_title", ""),
                    "product_role": row.get("product_role", ""),
                    "accessory_type": row.get("accessory_type", ""),
                    "category_fit_status": status,
                    "category_fit_reasons": ";".join(reasons),
                }
            )

    suspicious = pd.DataFrame(
        suspicious_rows,
        columns=[
            "sku",
            "ean",
            "category",
            "title",
            "new_title",
            "product_role",
            "accessory_type",
            "category_fit_status",
            "category_fit_reasons",
        ],
    )
    return result, suspicious


def build_searchable_text(row: pd.Series, name_columns: list[str]) -> str:
    parts = []
    for column in name_columns:
        if column in row and not is_blank(row.get(column, "")):
            parts.append(str(row.get(column, "")))
    parts.extend(str(row.get(column, "")) for column in ["old_title", "new_title"] if column in row)
    return strip_accents(compact_spaces(" ".join(parts))).lower()


def category_matches(category: str, needle: str) -> bool:
    if not needle:
        return False
    return strip_accents(needle).lower() in strip_accents(category).lower()


def find_matching_terms(text: str, terms: list[str]) -> list[str]:
    matches = []
    for term in terms:
        normalized = strip_accents(str(term)).lower()
        if " " not in normalized and "-" not in normalized and normalized in text:
            matches.append(str(term))
        elif re.search(rf"(?<![a-z0-9]){re.escape(normalized)}(?![a-z0-9])", text):
            matches.append(str(term))
    return matches


def main() -> None:
    parser = argparse.ArgumentParser(description="Waliduje zgodnosc produktow z kategoria.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet")
    parser.add_argument("--config", default="configs/categories/kanlux-oswietlenie.yaml")
    parser.add_argument("--output")
    args = parser.parse_args()
    df = read_products(args.input, sheet_name=args.sheet)
    config = load_yaml(args.config)
    validated, suspicious = validate_category_fit(df, config)
    if args.output:
        write_products(validated, args.output)
    print(f"OK: plik odczytany, liczba wierszy: {len(df)}")
    print(f"Podejrzane dopasowanie kategorii: {len(suspicious)}")


if __name__ == "__main__":
    main()
