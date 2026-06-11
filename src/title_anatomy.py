from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, ensure_dir, is_blank, load_yaml, normalize_header, strip_accents


DEFAULT_TITLE_ANATOMY_PATH = "configs/title_anatomy.yaml"


@dataclass(frozen=True)
class AnatomyTitleResult:
    title_without_code: str
    title: str
    source: str
    rule_key: str
    warnings: list[str]
    used_attributes: list[str]
    skipped_attributes: list[str]


def load_title_anatomy_config(path: str | Path | None = DEFAULT_TITLE_ANATOMY_PATH) -> dict[str, Any]:
    return load_yaml(path) if path else {}


def build_title_from_anatomy(
    row: pd.Series,
    anatomy_config: dict[str, Any],
    values: dict[str, str],
    max_length: int = 120,
) -> AnatomyTitleResult | None:
    if not anatomy_config or anatomy_config.get("enabled") is False:
        return None

    product_type = product_type_for_row(row, values, anatomy_config)
    normalized_type = normalize_product_type_key(product_type)
    rule = rule_for_product_type(normalized_type, anatomy_config)
    if not rule:
        return None

    warnings: list[str] = []
    if not rule_is_accepted(rule, anatomy_config):
        return None

    # Oprawy na zrodlo wymienne (gwint GU10/E14/E27) nie sa zintegrowanym LED -
    # wyekstrahowana "moc" to maks. moc zarowki, nie wlasna. Taka regula oddaje
    # produkt do legacy (np. "Oprawa sufitowa / Plafon ... E27"), bez "LED" i bez
    # mylacej mocy.
    if rule.get("skip_if_socket") and not is_blank(values.get("gwint", "")):
        return None

    # Slowo kluczowe: jawny primary_keyword z danych > kanoniczne `title_keyword`
    # reguly > surowy typ produktu (ostatecznosc, sygnalizowana ostrzezeniem).
    first_part = explicit_primary_keyword(row, anatomy_config)
    if not first_part:
        rule_keyword = normalize_woo_like_value(rule.get("title_keyword", ""))
        if rule_keyword:
            first_part = rule_keyword
        else:
            first_part = product_type
            warnings.append("title_anatomy_missing_primary_keyword")
    # Nadpisanie slowa kluczowego dla nietypowych serii (np. AVAR to nie panel,
    # tylko "Ramka oswietleniowa LED"). Ma priorytet nad title_keyword reguly.
    override = series_keyword_override(values, anatomy_config)
    if override:
        first_part = override

    used_attributes: list[str] = []
    skipped_attributes: list[str] = []
    parts = [first_part]
    seen_keys: set[str] = set()
    for value in parts:
        remember_seen_value(seen_keys, value)

    # `required_title_attributes` definiuje KOLEJNOSC i sklad tytulu (atrybut
    # trafia do tytulu, jesli jest obecny). `warn_if_missing` zaweza, ktore z
    # nich naprawde musza wystapic (ostrzezenie przy braku). Gdy go nie ma,
    # ostrzegamy dla wszystkich required - zachowanie wsteczne.
    warn_if_missing = rule.get("warn_if_missing")
    for attribute in list(rule.get("required_title_attributes") or []):
        value = anatomy_attribute_value(attribute, values, anatomy_config)
        if should_skip_attribute(attribute, value, values, rule, seen_keys, anatomy_config):
            skipped_attributes.append(attribute)
            continue
        if is_blank(value):
            if warn_if_missing is None or attribute in warn_if_missing:
                warnings.append(f"title_anatomy_missing_required:{attribute}")
            continue
        parts.append(value)
        used_attributes.append(attribute)
        remember_seen_value(seen_keys, value)

    for attribute in list(rule.get("optional_title_attributes") or []):
        value = anatomy_attribute_value(attribute, values, anatomy_config)
        if should_skip_attribute(attribute, value, values, rule, seen_keys, anatomy_config) or is_blank(value):
            continue
        parts.append(value)
        used_attributes.append(attribute)
        remember_seen_value(seen_keys, value)

    title_without_code = uppercase_first_letter(compact_title_parts(parts))
    producer = producer_for_row(row, values, anatomy_config)
    sku = sku_for_row(row, anatomy_config)
    title = append_producer_and_code(title_without_code, producer, sku)
    if len(title) > max_length:
        warnings.append(f"title_too_long:{len(title)}>{max_length}")
    if not title:
        warnings.append("empty_title")

    return AnatomyTitleResult(
        title_without_code=title_without_code,
        title=uppercase_first_letter(title),
        source="anatomy_rule",
        rule_key=normalized_type,
        warnings=warnings,
        used_attributes=used_attributes,
        skipped_attributes=skipped_attributes,
    )


def product_type_for_row(row: pd.Series, values: dict[str, str], anatomy_config: dict[str, Any]) -> str:
    candidates = [
        values.get("typ", ""),
        row.get("attr_typ", ""),
        row.get("Typ produktu", ""),
        row.get("Typ", ""),
        row.get("Rodzaj produktu", ""),
    ]
    for value in candidates:
        if not is_blank(value):
            return compact_spaces(str(value))
    return ""


def series_keyword_override(values: dict[str, str], anatomy_config: dict[str, Any]) -> str:
    """Slowo kluczowe nadpisane per seria (config: series_title_keyword).

    Dla serii, ktore mimo typu sa czyms innym (np. AVAR = ramka oswietleniowa).
    """
    _, keyword = series_keyword_override_match(values, anatomy_config)
    return keyword


def series_keyword_override_match(values: dict[str, str], anatomy_config: dict[str, Any]) -> tuple[str, str]:
    overrides = anatomy_config.get("series_title_keyword") or {}
    if not overrides:
        return "", ""
    # Marker moze byc w serii albo w samej nazwie produktu (np. SLR = solar).
    tokens = set(comparable_value_key(f"{values.get('seria', '')} {values.get('old_title', '')}").split())
    for series_name, keyword in overrides.items():
        marker = comparable_value_key(series_name)
        if marker in tokens:
            return marker, normalize_woo_like_value(keyword)
    return "", ""


def explicit_primary_keyword(row: pd.Series, anatomy_config: dict[str, Any]) -> str:
    """Jawne slowo kluczowe SEO z danych (kolumna primary_keyword).

    Tylko realny primary_keyword - bez fallbacku na typ produktu. Gdy go brak,
    slowo kluczowe pochodzi z kanonicznego `title_keyword` reguly (patrz
    build_title_from_anatomy).
    """
    source = str(anatomy_config.get("default_primary_keyword_source") or "primary_keyword")
    primary = first_row_value(row, [source, "primary_keyword"])
    return normalize_woo_like_value(primary) if primary else ""


def anatomy_attribute_value(attribute: str, values: dict[str, str], anatomy_config: dict[str, Any] | None = None) -> str:
    if is_avar_frame(values, anatomy_config or {}):
        if attribute == "montaz":
            return feminine_mounting(values.get("montaz", ""))
        if attribute == "ksztalt_masculine":
            return normalize_woo_like_value(values.get("ksztalt_feminine", "")) or normalize_woo_like_value(values.get(attribute, ""))
    return normalize_woo_like_value(values.get(attribute, ""))


def rule_for_product_type(normalized_type: str, anatomy_config: dict[str, Any]) -> dict[str, Any] | None:
    if not normalized_type:
        return None
    canonical_type = resolve_type_alias(normalized_type, anatomy_config)
    for rule in anatomy_config.get("rules") or []:
        if not isinstance(rule, dict):
            continue
        rule_key = normalize_product_type_key(rule.get("normalized_product_type", ""))
        if rule_key == canonical_type:
            return rule
    return None


def resolve_type_alias(normalized_type: str, anatomy_config: dict[str, Any]) -> str:
    """Mapuje warianty nazwy typu na typ kanoniczny, ktory ma regule anatomii.

    Aliasy mieszkaja w configu (`product_type_aliases`), wiec dopasowanie do
    regul mozna rozszerzac bez zmian w kodzie.
    """
    aliases = {
        normalize_product_type_key(key): normalize_product_type_key(value)
        for key, value in (anatomy_config.get("product_type_aliases") or {}).items()
    }
    return aliases.get(normalized_type, normalized_type)


def rule_is_accepted(rule: dict[str, Any], anatomy_config: dict[str, Any]) -> bool:
    if bool(rule.get("manual_review_required")):
        return False
    accepted = {normalize_header(value) for value in anatomy_config.get("accepted_status_values") or ["approved"]}
    status = normalize_header(str(rule.get("status", "")))
    return status in accepted


def should_skip_attribute(
    attribute: str,
    value: str,
    values: dict[str, str],
    rule: dict[str, Any],
    seen_keys: set[str],
    anatomy_config: dict[str, Any] | None = None,
) -> bool:
    if is_blank(value):
        return False
    if attribute in {"moc", "lm_w"} and is_solar_product(values, anatomy_config or {}):
        return True
    value_key = comparable_value_key(value)
    if not value_key or value_key in seen_keys:
        return True
    excluded = rule.get("excluded_if_redundant") or {}
    redundant_sources = excluded.get(attribute) or []
    for source in redundant_sources:
        source_value = values.get(str(source), "")
        if source_value and value_key and value_key in comparable_value_key(source_value):
            return True
    if attribute.startswith("kolor"):
        return color_redundant_with_light(value, values)
    return False


def is_avar_frame(values: dict[str, str], anatomy_config: dict[str, Any]) -> bool:
    marker, keyword = series_keyword_override_match(values, anatomy_config)
    return marker == "avar" or "ramka oswietleniowa" in comparable_value_key(keyword)


def is_solar_product(values: dict[str, str], anatomy_config: dict[str, Any]) -> bool:
    marker, keyword = series_keyword_override_match(values, anatomy_config)
    haystack = comparable_value_key(
        " ".join(
            [
                marker,
                keyword,
                values.get("seria", ""),
                values.get("old_title", ""),
                values.get("typ", ""),
            ]
        )
    )
    return bool(re.search(r"(?<!\w)(?:slr|solar\w*|solarn\w*)(?!\w)", haystack))


def feminine_mounting(value: str) -> str:
    mapping = {
        "podtynkowy": "podtynkowa",
        "natynkowy": "natynkowa",
        "zwieszany": "zwieszana",
    }
    normalized = normalize_woo_like_value(value)
    return mapping.get(comparable_value_key(normalized), normalized)


def color_redundant_with_light(color: str, values: dict[str, str]) -> bool:
    color_key = comparable_value_key(color)
    if not color_key:
        return False
    light_text = " ".join(
        comparable_value_key(values.get(key, ""))
        for key in ["barwa", "barwa_swiatla", "barwa_slownie", "barwa_neuter", "barwa_zakres"]
    )
    # Do not treat CCT color temperature names as housing colors.
    housing_colors = {
        "bialy",
        "biala",
        "biale",
        "czarny",
        "czarna",
        "czarne",
        "szary",
        "szara",
        "srebrny",
        "srebrna",
        "grafitowy",
        "grafitowa",
        "bezowy",
        "bezowa",
        "brazowy",
        "brazowa",
        "zloty",
        "zlota",
    }
    return color_key in housing_colors and re.search(rf"(?<!\w){re.escape(color_key)}(?!\w)", light_text) is not None


def compact_title_parts(parts: list[str]) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for part in parts:
        cleaned = normalize_woo_like_value(part)
        key = comparable_value_key(cleaned)
        if not cleaned or key in seen:
            continue
        result.append(cleaned)
        remember_seen_value(seen, cleaned)
    return compact_spaces(" ".join(result))


def append_producer_and_code(title: str, producer: str, sku: str) -> str:
    title_with_producer = title
    if producer and not re.search(rf"\b{re.escape(producer)}\b", title_with_producer, flags=re.IGNORECASE):
        title_with_producer = compact_spaces(f"{title_with_producer} {producer}")
    if is_blank(sku):
        return title_with_producer
    sku = compact_spaces(str(sku))
    if re.search(rf"(?<!\w){re.escape(sku)}(?!\w)", title_with_producer):
        return title_with_producer
    return compact_spaces(f"{title_with_producer} {sku}")


def sku_for_row(row: pd.Series, anatomy_config: dict[str, Any]) -> str:
    return normalize_sku(first_row_value(row, list(anatomy_config.get("sku_fields") or [])))


def producer_for_row(row: pd.Series, values: dict[str, str], anatomy_config: dict[str, Any]) -> str:
    return first_present(
        values.get("producent", ""),
        first_row_value(row, list(anatomy_config.get("producer_fields") or [])),
    )


def first_row_value(row: pd.Series, columns: list[str]) -> str:
    for column in columns:
        if column in row and not is_blank(row.get(column, "")):
            return compact_spaces(str(row.get(column, "")))
    return ""


def first_present(*values: Any) -> str:
    for value in values:
        if not is_blank(value):
            return compact_spaces(str(value))
    return ""


def normalize_sku(value: Any) -> str:
    value = compact_spaces(str(value or ""))
    return value[:-2] if re.fullmatch(r"\d+\.0", value) else value


def normalize_product_type_key(value: Any) -> str:
    return normalize_header(str(value or ""))


def normalize_woo_like_value(value: Any) -> str:
    value = compact_spaces(str(value or ""))
    if not value:
        return ""
    value = re.sub(r"\s*\|\s*", " / ", value)
    value = re.sub(r"\s*/\s*", " / ", value)
    value = re.sub(r"\bip\s*(\d{2})\b", r"IP\1", value, flags=re.IGNORECASE)
    value = re.sub(r"(?<=\d)\s+([WK])\b", r"\1", value, flags=re.IGNORECASE)
    value = re.sub(r"(?<=\d)\s+lm\b", "lm", value, flags=re.IGNORECASE)
    # Jednostka skutecznosci ma byc zwarta: "120lm / W" -> "120lm/W".
    value = re.sub(r"lm\s*/\s*W\b", "lm/W", value, flags=re.IGNORECASE)
    return compact_spaces(value)


def comparable_value_key(value: Any) -> str:
    value = strip_accents(compact_spaces(str(value or ""))).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return compact_spaces(value)


def remember_seen_value(seen: set[str], value: Any) -> None:
    key = comparable_value_key(value)
    if key:
        seen.add(key)


def uppercase_first_letter(value: str) -> str:
    for index, char in enumerate(value):
        if char.isalpha():
            return value[:index] + char.upper() + value[index + 1 :]
    return value


def build_title_type_review(df: pd.DataFrame, anatomy_config: dict[str, Any], max_examples: int = 3) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    if df.empty:
        return pd.DataFrame(columns=title_type_review_columns())

    for product_type, group in grouped_by_product_type(df):
        normalized = normalize_product_type_key(product_type)
        rule = rule_for_product_type(normalized, anatomy_config)
        first = group.iloc[0]
        examples = [compact_spaces(str(value)) for value in group.get("new_title", group.get("old_title", pd.Series(dtype=str))).head(max_examples)]
        primary = first_row_value(first, ["primary_keyword"]) or product_type
        rows.append(
            {
                "normalized_product_type": normalized,
                "product_type": product_type,
                "products_count": len(group),
                "review_status": str(rule.get("status", "missing_rule")) if rule else "missing_rule",
                "manual_review_required": bool(rule.get("manual_review_required", not rule_is_accepted(rule, anatomy_config))) if rule else True,
                "primary_keyword": primary,
                "required_title_attributes": ",".join(rule.get("required_title_attributes") or []) if rule else "",
                "optional_title_attributes": ",".join(rule.get("optional_title_attributes") or []) if rule else "",
                "example_title_1": examples[0] if len(examples) > 0 else "",
                "example_title_2": examples[1] if len(examples) > 1 else "",
                "example_title_3": examples[2] if len(examples) > 2 else "",
            }
        )
    return pd.DataFrame(rows, columns=title_type_review_columns())


def grouped_by_product_type(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    type_values = []
    for _, row in df.iterrows():
        product_type = first_row_value(row, ["attr_typ", "Typ produktu", "Typ", "Rodzaj produktu"])
        type_values.append(product_type or "__MISSING__")
    temp = df.copy()
    temp["__title_anatomy_type"] = type_values
    return [(str(key), group.drop(columns=["__title_anatomy_type"])) for key, group in temp.groupby("__title_anatomy_type", sort=False)]


def title_type_review_columns() -> list[str]:
    return [
        "normalized_product_type",
        "product_type",
        "products_count",
        "review_status",
        "manual_review_required",
        "primary_keyword",
        "required_title_attributes",
        "optional_title_attributes",
        "example_title_1",
        "example_title_2",
        "example_title_3",
    ]


def duplicate_titles_report(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for column in ["title_uniqueness_key", "final_title_uniqueness_key"]:
        if column not in df.columns:
            continue
        counts = df[column].astype(str).map(compact_spaces).value_counts()
        duplicate_keys = {key for key, count in counts.items() if key and count > 1}
        for index, row in df[df[column].isin(duplicate_keys)].iterrows():
            rows.append(
                {
                    "duplicate_type": column,
                    "duplicate_key": row.get(column, ""),
                    "row_index": int(index) + 2,
                    "sku": first_row_value(row, ["sku", "SKU", "Kod", "Kod Producenta"]),
                    "new_title": row.get("new_title", ""),
                    "seo_title_without_code": row.get("seo_title_without_code", ""),
                    "product_type": first_row_value(row, ["attr_typ", "Typ produktu", "Typ", "Rodzaj produktu"]),
                }
            )
    return pd.DataFrame(rows, columns=duplicate_titles_columns())


def duplicate_titles_columns() -> list[str]:
    return [
        "duplicate_type",
        "duplicate_key",
        "row_index",
        "sku",
        "new_title",
        "seo_title_without_code",
        "product_type",
    ]


def title_anatomy_warnings_columns() -> list[str]:
    return ["row_index", "sku", "product_type", "warning", "new_title"]


def title_anatomy_warnings_report(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    if "warnings" not in df.columns:
        return pd.DataFrame(columns=title_anatomy_warnings_columns())
    for index, row in df.iterrows():
        warnings = [item for item in str(row.get("warnings", "")).split(";") if item.startswith("title_anatomy") or item.startswith("duplicate_")]
        for warning in warnings:
            rows.append(
                {
                    "row_index": int(index) + 2,
                    "sku": first_row_value(row, ["sku", "SKU", "Kod", "Kod Producenta"]),
                    "product_type": first_row_value(row, ["attr_typ", "Typ produktu", "Typ", "Rodzaj produktu"]),
                    "warning": warning,
                    "new_title": row.get("new_title", ""),
                }
            )
    return pd.DataFrame(rows, columns=title_anatomy_warnings_columns())


def add_review_columns(df: pd.DataFrame, proposal_column: str) -> pd.DataFrame:
    """Dokleja kolumny do recznego przegladu:

    - ``proponowana zmiana`` - aktualny wynik programu (zrodlo: proposal_column),
    - ``oczekiwany wynik`` - pusta kolumna do wpisania docelowej wartosci.
    """
    result = df.copy()
    result["proponowana zmiana"] = result[proposal_column] if proposal_column in result.columns else ""
    result["oczekiwany wynik"] = ""
    return result


def write_title_anatomy_reports(df: pd.DataFrame, reports_dir: str | Path, anatomy_config: dict[str, Any]) -> None:
    reports = ensure_dir(Path(reports_dir) / "title_anatomy")
    sheets = {
        "Przeglad typow": add_review_columns(build_title_type_review(df, anatomy_config), "example_title_1"),
        "Duplikaty": add_review_columns(duplicate_titles_report(df), "new_title"),
        "Ostrzezenia": add_review_columns(title_anatomy_warnings_report(df), "new_title"),
    }
    output_path = reports / "title_anatomy_review.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
