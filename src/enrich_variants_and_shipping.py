from __future__ import annotations

import argparse
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, NamedTuple

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, load_catalog_knowledge
from recommend_categories import (
    DEFAULT_OVERRIDES_PATH,
    DEFAULT_TAXONOMY_PATH,
    CategoryRecommender,
    normalize_category_path,
)
from utils import compact_spaces, detect_column, ensure_dir, is_blank, load_yaml, normalize_header, read_products, write_products


DEFAULT_HISTORY_PATH = "archived_input_files/wszystko.csv"
DEFAULT_SHIPPING_RULES_PATH = "dictionaries/shipping_class_rules.yaml"

TITLE_COLUMNS = [
    "Title",
    "Nazwa",
    "Nazwa produktu",
    "name",
    "new_title",
    "Nazwa B2C / SEO final",
    "Nazwa B2C / SEO",
]
SKU_COLUMNS = ["Sku", "SKU", "sku", "Kod", "Kod Producenta", "PRODUCT_CODE", "TOWARIDX"]
TAG_COLUMNS = ["Product Tags", "Tagi", "tags", "product_tags"]
SHIPPING_COLUMNS = ["Shipping Class", "Klasa wysyłkowa", "Klasa wysylkowa", "shipping_class"]
CATEGORY_COLUMNS = ["Kategorie", "Kategoria", "category", "categories"]
WEIGHT_COLUMNS = ["Waga (kg)", "weight", "Weight"]
LENGTH_COLUMNS = ["Długość (cm)", "Dlugosc (cm)", "Length", "length"]
MIN_VARIANT_GROUP_SIZE = 2

PRODUCER_SUFFIX_RE = re.compile(r"^[A-Z]{2,6}[A-Z0-9]?$")
MODEL_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z][A-Z0-9-]{2,})(?![A-Za-z0-9])")
MODEL_STOP_WORDS = {"led", "cct", "rgb", "rgbw", "e27", "e14", "gu10", "ip"}
ATTRIBUTE_PREFIXES = ["Atrybut produktu:", "Parametr:", "attr_"]
IGNORED_VARIANT_ATTRIBUTES = {
    "dane producenta",
    "producent",
    "marka",
    "symbol",
    "kod",
    "ean",
    "gtin",
    "wariant",
}

TITLE_STOP_WORDS = {
    "do",
    "z",
    "ze",
    "i",
    "oraz",
    "na",
    "dla",
    "kolor",
    "barwa",
    "neutralne",
    "neutralna",
    "ciepla",
    "cieply",
    "zimna",
    "zimne",
    "biala",
    "bialy",
    "biale",
    "czarna",
    "czarny",
    "czarne",
    "szara",
    "szary",
    "grafitowa",
    "grafitowy",
    "srebrna",
    "srebrny",
    "brazowa",
    "brazowy",
    "bezowa",
    "bezowy",
    "mat",
    "matowa",
    "okragla",
    "okragly",
    "kwadratowa",
    "kwadratowy",
    "prostokatna",
    "prostokatny",
    "mikrofalowym",
    "czujnikiem",
}

SPEC_PATTERNS = [
    r"\bip\s*\d{2}\b",
    r"\b\d+(?:[,.]\d+)?\s*(?:w|lm|k|v|a|cm|mm|m|hz|mah)\b",
    r"\b\d+\s*[-x]\s*\d+(?:\s*[-x]\s*\d+)?\s*(?:mm|cm|m)?\b",
    r"\b\d{3,5}\s*[-/]\s*\d{3,5}\s*k\b",
    r"\b\d+p\+?n?\b",
    r"\b\d+r/\d+m\b",
]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Dopisuje tagi wariantow i rekomenduje klase wysylkowa dla pliku produktowego."
    )
    parser.add_argument("--input", required=True, help="CSV/XLSX z produktami.")
    parser.add_argument("--sheet", help="Arkusz XLSX.")
    parser.add_argument("--output", default="output/variants_and_shipping.xlsx", help="Plik wynikowy CSV/XLSX.")
    parser.add_argument("--history", default=DEFAULT_HISTORY_PATH, help="Eksport WooCommerce z historycznymi klasami wysylkowymi.")
    parser.add_argument("--shipping-rules", default=DEFAULT_SHIPPING_RULES_PATH)
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE_PATH)
    parser.add_argument("--taxonomy", default=DEFAULT_TAXONOMY_PATH)
    parser.add_argument("--overrides", default=DEFAULT_OVERRIDES_PATH)
    parser.add_argument(
        "--overwrite-shipping",
        action="store_true",
        help="Nadpisuje istniejaca klase wysylkowa. Domyslnie uzupelnia tylko puste komorki.",
    )
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    catalog_knowledge = load_catalog_knowledge(args.catalog_knowledge)
    taxonomy = load_yaml(args.taxonomy) if args.taxonomy else {}
    overrides = load_yaml(args.overrides) if args.overrides else {}
    rules = load_yaml(args.shipping_rules) if args.shipping_rules else {}

    category_recommender = CategoryRecommender(catalog_knowledge, taxonomy, overrides)
    shipping_recommender = ShippingClassRecommender.from_history(args.history, rules)

    result = enrich_dataframe(
        df,
        category_recommender=category_recommender,
        shipping_recommender=shipping_recommender,
        overwrite_shipping=args.overwrite_shipping,
    )
    write_products(result, args.output)
    report_path = write_report(result, args.output)

    print(f"OK: wczytano {len(df)} produktow")
    print(f"OK: zapisano {args.output}")
    print(f"OK: raport {report_path}")


def enrich_dataframe(
    df: pd.DataFrame,
    category_recommender: CategoryRecommender,
    shipping_recommender: "ShippingClassRecommender",
    overwrite_shipping: bool = False,
) -> pd.DataFrame:
    columns = detect_columns(df)
    result = add_variant_tags(df, columns)
    result = add_shipping_classes(
        result,
        columns=detect_columns(result),
        category_recommender=category_recommender,
        shipping_recommender=shipping_recommender,
        overwrite_shipping=overwrite_shipping,
    )
    return result


def detect_columns(df: pd.DataFrame) -> dict[str, str | None]:
    return {
        "title": detect_column(df, TITLE_COLUMNS),
        "sku": detect_column(df, SKU_COLUMNS),
        "tags": detect_column(df, TAG_COLUMNS),
        "shipping": detect_column(df, SHIPPING_COLUMNS),
        "category": detect_column(df, CATEGORY_COLUMNS),
        "weight": detect_column(df, WEIGHT_COLUMNS),
        "length": detect_column(df, LENGTH_COLUMNS),
    }


def add_variant_tags(df: pd.DataFrame, columns: dict[str, str | None]) -> pd.DataFrame:
    result = df.copy()
    title_col = columns.get("title")
    sku_col = columns.get("sku")
    tags_col = columns.get("tags")
    if not tags_col:
        tags_col = "Product Tags"
        result[tags_col] = ""

    family_keys = []
    family_labels = []
    for _, row in result.iterrows():
        title = first_value(row, [title_col])
        sku = first_value(row, [sku_col])
        family_key, family_label = product_family(title, sku)
        family_keys.append(family_key)
        family_labels.append(family_label)

    groups: dict[str, list[int]] = defaultdict(list)
    for index, key in enumerate(family_keys):
        if key:
            groups[key].append(index)

    group_sizes = {key: len(indices) for key, indices in groups.items()}
    attribute_columns = detect_variant_attribute_columns(result)
    group_infos = {
        key: analyze_variant_group(result, indices, attribute_columns)
        for key, indices in groups.items()
        if len(indices) >= MIN_VARIANT_GROUP_SIZE
    }
    tag_by_key = {
        key: f"Wariant - {family_labels[indices[0]]}"
        for key, indices in groups.items()
        if key in group_infos and group_infos[key].qualifies
    }

    variant_tags: list[str] = []
    variant_link_tags: list[str] = []
    variant_reasons: list[str] = []
    variant_differentiating_attributes: list[str] = []
    variant_attribute_summaries: list[str] = []
    for row_index, key in enumerate(family_keys):
        tag = tag_by_key.get(key, "")
        variant_tags.append(tag)
        variant_link_tags.append(tag)
        info = group_infos.get(key)
        if info:
            variant_differentiating_attributes.append(", ".join(info.differentiating_attributes))
            variant_attribute_summaries.append(info.attribute_summary)
        else:
            variant_differentiating_attributes.append("")
            variant_attribute_summaries.append("")
        if tag:
            result.at[result.index[row_index], tags_col] = append_tag(result.iloc[row_index].get(tags_col, ""), tag)
            variant_reasons.append("wspolny model/rodzina produktu i roznice w atrybutach")
        elif info and not info.qualifies:
            variant_reasons.append(info.reject_reason)
        else:
            variant_reasons.append("pojedynczy produkt albo zbyt slaby sygnal wariantu")

    result["variant_group_key"] = family_keys
    result["variant_group_size"] = [group_sizes.get(key, 0) if key else 0 for key in family_keys]
    result["variant_tag"] = variant_tags
    result["variant_link_tag"] = variant_link_tags
    result["variant_differentiating_attributes"] = variant_differentiating_attributes
    result["variant_attribute_summary"] = variant_attribute_summaries
    result["variant_reason"] = variant_reasons
    return result


class VariantGroupInfo(NamedTuple):
    qualifies: bool
    differentiating_attributes: list[str]
    attribute_summary: str
    reject_reason: str


def detect_variant_attribute_columns(df: pd.DataFrame) -> list[str]:
    columns: list[str] = []
    for column in df.columns:
        text = str(column)
        attr_name = clean_attribute_name(text)
        if not attr_name:
            continue
        if normalize_header(attr_name) in IGNORED_VARIANT_ATTRIBUTES:
            continue
        if any(text.startswith(prefix) for prefix in ATTRIBUTE_PREFIXES):
            columns.append(column)
    return columns


def clean_attribute_name(column: str) -> str:
    text = compact_spaces(str(column))
    for prefix in ATTRIBUTE_PREFIXES:
        if text.startswith(prefix):
            return compact_spaces(text[len(prefix) :])
    return ""


def analyze_variant_group(df: pd.DataFrame, indices: list[int], attribute_columns: list[str]) -> VariantGroupInfo:
    differentiating: list[str] = []
    summary_parts: list[str] = []
    for column in attribute_columns:
        values = []
        for index in indices:
            value = normalize_attribute_value(df.iloc[index].get(column, ""))
            if value:
                values.append(value)
        unique_values = unique(values)
        if len(unique_values) <= 1:
            continue
        attr_name = clean_attribute_name(str(column))
        differentiating.append(attr_name)
        summary_parts.append(f"{attr_name}: {' | '.join(unique_values[:8])}")

    if differentiating:
        return VariantGroupInfo(True, differentiating, "; ".join(summary_parts[:12]), "")
    return VariantGroupInfo(
        False,
        [],
        "",
        "odrzucono wariant: wspolna rodzina, ale brak realnych roznic w atrybutach",
    )


def normalize_attribute_value(value: Any) -> str:
    if is_blank(value):
        return ""
    return compact_spaces(str(value)).strip(" ;,")


def product_family(title: str, sku: str) -> tuple[str, str]:
    title_key, label = title_family(title)
    sku_key = sku_family(sku)
    if title_key:
        return title_key, label
    if sku_key:
        return sku_key, sku_key.upper()
    return "", ""


def title_family(title: str) -> tuple[str, str]:
    clean = normalize_title_for_family(title)
    tokens = meaningful_tokens(clean)
    if len(tokens) < 2:
        return "", ""

    models = [
        normalize_header(match.group(1))
        for match in MODEL_RE.finditer(strip_known_codes(title))
        if normalize_header(match.group(1)) not in MODEL_STOP_WORDS
    ]
    model = next(iter(models), "")
    if model:
        model_index = tokens.index(model) if model in tokens else min(len(tokens), 3)
        prefix = tokens[: max(model_index, 1)]
        nearby_models = [item for item in models if item in tokens[model_index : model_index + 5]][:3]
        selected = [*prefix[:4], *nearby_models]
        if "cct" in tokens[: model_index + 1] and "cct" not in selected:
            selected.insert(min(len(selected), 2), "cct")
    else:
        selected = tokens[:5]

    selected = [token for token in selected if token and token not in TITLE_STOP_WORDS]
    if len(selected) < 2:
        return "", ""
    key = "title:" + "-".join(selected[:6])
    label = " ".join(token.upper() if len(token) <= 4 and token.isalnum() else token.capitalize() for token in selected[:6])
    return key, label


def normalize_title_for_family(title: str) -> str:
    value = normalize_header(strip_known_codes(title))
    for pattern in SPEC_PATTERNS:
        value = re.sub(pattern, " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\b(?:kanlux|kobi|hager|ospel|berker|karlik|kontakt simon|simon|edo|philips|ledvance)\b", " ", value)
    return compact_spaces(value)


def strip_known_codes(title: str) -> str:
    value = str(title or "")
    value = re.sub(r"\b\d{5,14}\b\s*$", " ", value)
    return value


def meaningful_tokens(value: str) -> list[str]:
    tokens = re.findall(r"[a-z0-9]{2,}", value)
    result: list[str] = []
    for token in tokens:
        if token in TITLE_STOP_WORDS:
            continue
        if token.isdigit():
            continue
        result.append(token)
    return result


def sku_family(sku: str) -> str:
    base = sku_without_producer_suffix(sku)
    normalized = normalize_header(base)
    if not normalized:
        return ""
    parts = normalized.split()
    if len(parts) >= 2 and len(parts[-1]) <= 4:
        return "sku:" + "-".join(parts[:-1])
    match = re.match(r"^([a-z]+[a-z0-9-]*?)[- ]?\d{1,4}[a-z]?$", normalized)
    if match and len(match.group(1)) >= 3:
        return "sku:" + match.group(1)
    return ""


def sku_without_producer_suffix(sku: str) -> str:
    text = compact_spaces(str(sku or "")).upper()
    if "/" not in text:
        return text
    head, tail = text.rsplit("/", 1)
    if PRODUCER_SUFFIX_RE.fullmatch(tail):
        return head
    return text


def append_tag(current: Any, tag: str) -> str:
    tags = split_tags(current)
    normalized = {normalize_header(item) for item in tags}
    if normalize_header(tag) not in normalized:
        tags.append(tag)
    return ", ".join(tags)


def split_tags(value: Any) -> list[str]:
    text = compact_spaces(str(value or ""))
    if not text:
        return []
    parts = re.split(r"\s*,\s*|\s*;\s*|\s*\|\s*", text)
    return [compact_spaces(part) for part in parts if compact_spaces(part)]


def add_shipping_classes(
    df: pd.DataFrame,
    columns: dict[str, str | None],
    category_recommender: CategoryRecommender,
    shipping_recommender: "ShippingClassRecommender",
    overwrite_shipping: bool,
) -> pd.DataFrame:
    result = df.copy()
    shipping_col = columns.get("shipping")
    if not shipping_col:
        shipping_col = "Shipping Class"
        result[shipping_col] = ""

    category_columns = {
        "title": columns.get("title"),
        "type": None,
        "producer": None,
        "series": None,
        "sku": columns.get("sku"),
        "category": columns.get("category"),
    }

    previous_values = []
    predicted_values = []
    confidence_values = []
    reason_values = []
    guessed_categories = []
    for index, row in result.iterrows():
        recommendations = category_recommender.recommend(row, category_columns, top_n=1)
        guessed_category = recommendations[0]["category"] if recommendations else ""
        prediction = shipping_recommender.recommend(row, columns, guessed_category)

        previous = compact_spaces(str(row.get(shipping_col, "")))
        previous_values.append(previous)
        predicted_values.append(prediction.shipping_class)
        confidence_values.append(prediction.confidence)
        reason_values.append(prediction.reason)
        guessed_categories.append(guessed_category)

        if overwrite_shipping or is_blank(previous):
            result.at[index, shipping_col] = prediction.shipping_class

    result["shipping_class_previous"] = previous_values
    result["shipping_class_predicted"] = predicted_values
    result["shipping_class_confidence"] = confidence_values
    result["shipping_class_reason"] = reason_values
    result["shipping_category_guess"] = guessed_categories
    return result


class ShippingPrediction:
    def __init__(self, shipping_class: str, confidence: str, reason: str) -> None:
        self.shipping_class = shipping_class
        self.confidence = confidence
        self.reason = reason


class ShippingClassRecommender:
    def __init__(self, rules: dict[str, Any], category_stats: dict[str, Counter[str]]) -> None:
        self.default_class = str(rules.get("default_shipping_class") or "Standard")
        self.rules = sorted(rules.get("rules") or [], key=lambda item: int(item.get("priority", 0)), reverse=True)
        self.category_stats = category_stats

    @classmethod
    def from_history(cls, history_path: str | Path, rules: dict[str, Any]) -> "ShippingClassRecommender":
        path = Path(history_path)
        if not path.is_absolute():
            path = Path(__file__).resolve().parents[1] / path
        if not path.exists():
            return cls(rules, {})
        history = read_products(path)
        columns = detect_columns(history)
        shipping_col = columns.get("shipping")
        category_col = columns.get("category")
        category_stats: dict[str, Counter[str]] = defaultdict(Counter)
        if shipping_col and category_col:
            for _, row in history.iterrows():
                shipping_class = compact_spaces(str(row.get(shipping_col, "")))
                if not shipping_class:
                    continue
                for category in split_categories(row.get(category_col, "")):
                    category_stats[normalize_header(category)][shipping_class] += 1
                    category_stats[normalize_header(category_leaf(category))][shipping_class] += 1
        return cls(rules, category_stats)

    def recommend(self, row: pd.Series, columns: dict[str, str | None], guessed_category: str = "") -> ShippingPrediction:
        title = first_value(row, [columns.get("title")])
        category = first_value(row, [columns.get("category")])
        sku = first_value(row, [columns.get("sku")])
        weight = parse_decimal(first_value(row, [columns.get("weight")]))
        length = parse_decimal(first_value(row, [columns.get("length")]))

        title_text = normalize_header(" ".join([title, sku]))
        category_text = normalize_header(" ".join([category, guessed_category]))
        combined_text = compact_spaces(f"{title_text} {category_text}")

        rule_prediction = self.recommend_by_rules(combined_text, category_text, weight, length)
        if rule_prediction:
            return rule_prediction

        category_prediction = self.recommend_by_category(category, guessed_category)
        if category_prediction:
            return category_prediction

        return ShippingPrediction(self.default_class, "niska", "domyslna klasa, brak silniejszych sygnalow")

    def recommend_by_rules(
        self,
        combined_text: str,
        category_text: str,
        weight: float | None,
        length: float | None,
    ) -> ShippingPrediction | None:
        for rule in self.rules:
            matches: list[str] = []
            min_weight = parse_decimal(rule.get("min_weight_kg", ""))
            if min_weight is not None and weight is not None and weight >= min_weight:
                matches.append(f"waga >= {min_weight:g} kg")
            min_length = parse_decimal(rule.get("min_length_cm", ""))
            if min_length is not None and length is not None and length >= min_length:
                matches.append(f"dlugosc >= {min_length:g} cm")
            title_terms = matching_terms(combined_text, rule.get("title_terms_any") or [])
            category_terms = matching_terms(category_text, rule.get("category_terms_any") or [])
            matches.extend(title_terms)
            matches.extend(category_terms)
            if not matches:
                continue
            reason = compact_spaces(str(rule.get("reason") or "regula klasy wysylkowej"))
            return ShippingPrediction(str(rule.get("class")), "wysoka", f"{reason}: {', '.join(matches[:4])}")
        return None

    def recommend_by_category(self, category: str, guessed_category: str) -> ShippingPrediction | None:
        categories = [(item, "plik") for item in split_categories(category)]
        if guessed_category:
            categories.append((guessed_category, "zgadnieta kategoria"))
        best: tuple[str, int, int, str, str] | None = None
        for category_value, source in categories:
            key = normalize_header(category_value)
            stats = self.category_stats.get(key) or self.category_stats.get(normalize_header(category_leaf(category_value)))
            if not stats:
                continue
            shipping_class, count = stats.most_common(1)[0]
            total = sum(stats.values())
            if total <= 0:
                continue
            if best is None or count > best[1]:
                best = (shipping_class, count, total, category_value, source)
        if not best:
            return None
        shipping_class, count, total, category_value, source = best
        ratio = count / total
        if total >= 3 and ratio >= 0.65:
            confidence = "wysoka" if ratio >= 0.85 else "srednia"
            return ShippingPrediction(
                shipping_class,
                confidence,
                f"historia kategorii ({source}): {category_value} ({count}/{total})",
            )
        if source == "plik" and shipping_class != self.default_class and total >= 3 and ratio >= 0.5:
            return ShippingPrediction(
                shipping_class,
                "srednia",
                f"historia kategorii ({source}): {category_value} ({count}/{total})",
            )
        return None


def split_categories(value: Any) -> list[str]:
    text = compact_spaces(str(value or ""))
    if not text:
        return []
    return [normalize_category_path(part) for part in re.split(r"\s*,\s*", text) if compact_spaces(part)]


def category_leaf(path: str) -> str:
    parts = [compact_spaces(part) for part in normalize_category_path(path).split(">") if compact_spaces(part)]
    return parts[-1] if parts else ""


def matching_terms(text: str, terms: list[str]) -> list[str]:
    matches = []
    for term in terms:
        normalized = normalize_header(str(term))
        if not normalized:
            continue
        if re.search(rf"(?<![a-z0-9]){re.escape(normalized)}(?![a-z0-9])", text):
            matches.append(str(term))
    return matches


def parse_decimal(value: Any) -> float | None:
    if is_blank(value):
        return None
    text = str(value).replace(",", ".")
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


def first_value(row: pd.Series, columns: list[str | None]) -> str:
    for column in columns:
        if column and column in row and not is_blank(row.get(column, "")):
            return compact_spaces(str(row.get(column, "")))
    return ""


def write_report(result: pd.DataFrame, output: str | Path) -> Path:
    output_path = Path(output)
    reports_dir = ensure_dir(Path("reports") / output_path.stem)
    summary_path = reports_dir / "variants_and_shipping_summary.csv"
    report_path = reports_dir / "variant_tag_report.xlsx"
    summary_rows = [
        {"metric": "rows", "value": len(result)},
        {"metric": "rows_with_variant_tag", "value": int((result["variant_tag"] != "").sum())},
        {"metric": "variant_groups", "value": int(result.loc[result["variant_tag"] != "", "variant_tag"].nunique())},
        {
            "metric": "rejected_candidate_groups_without_attribute_differences",
            "value": int(count_rejected_variant_groups(result)),
        },
    ]
    for shipping_class, count in result["shipping_class_predicted"].value_counts(dropna=False).items():
        summary_rows.append({"metric": f"predicted_shipping:{shipping_class}", "value": int(count)})
    summary = pd.DataFrame(summary_rows)
    summary.to_csv(summary_path, index=False, encoding="utf-8-sig")

    groups = build_variant_group_report(result)
    items = build_variant_item_report(result)
    with pd.ExcelWriter(report_path) as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        groups.to_excel(writer, sheet_name="variant_groups", index=False)
        items.to_excel(writer, sheet_name="variant_items", index=False)
    return report_path


def count_rejected_variant_groups(result: pd.DataFrame) -> int:
    rejected = result[
        (result["variant_group_size"].fillna(0).astype(str) != "0")
        & (result["variant_tag"] == "")
        & result["variant_reason"].str.contains("brak realnych roznic", na=False)
    ]
    return rejected["variant_group_key"].nunique()


def build_variant_group_report(result: pd.DataFrame) -> pd.DataFrame:
    columns = detect_columns(result)
    title_col = columns.get("title") or "Title"
    sku_col = columns.get("sku") or "Sku"
    tags_col = columns.get("tags") or "Product Tags"
    rows: list[dict[str, Any]] = []
    candidates = result[result["variant_group_size"].fillna(0).astype(str).str.isnumeric()].copy()
    candidates = candidates[candidates["variant_group_size"].astype(int) >= MIN_VARIANT_GROUP_SIZE]
    for group_key, group in candidates.groupby("variant_group_key", sort=True):
        first = group.iloc[0]
        variant_tag = compact_spaces(str(first.get("variant_link_tag", "")))
        status = "TAGGED" if variant_tag else "REJECTED"
        rows.append(
            {
                "status": status,
                "variant_group_key": group_key,
                "variant_link_tag": variant_tag,
                "tag_to_add": variant_tag,
                "group_size": len(group),
                "differentiating_attributes": first.get("variant_differentiating_attributes", ""),
                "attribute_value_summary": clip(first.get("variant_attribute_summary", ""), 1200),
                "reason": first.get("variant_reason", ""),
                "skus": clip(", ".join(str(value) for value in group.get(sku_col, [])), 1200),
                "titles_sample": clip(" | ".join(str(value) for value in group.get(title_col, []).head(8)), 1200),
                "tags_after_sample": clip(" | ".join(unique([str(value) for value in group.get(tags_col, []).head(8)])), 1200),
            }
        )
    return pd.DataFrame(rows)


def build_variant_item_report(result: pd.DataFrame) -> pd.DataFrame:
    columns = detect_columns(result)
    title_col = columns.get("title") or "Title"
    sku_col = columns.get("sku") or "Sku"
    tags_col = columns.get("tags") or "Product Tags"
    item_columns = [
        "variant_link_tag",
        "variant_group_key",
        "variant_group_size",
        sku_col,
        title_col,
        tags_col,
        "variant_differentiating_attributes",
        "variant_attribute_summary",
        "variant_reason",
        "shipping_class_predicted",
        "shipping_class_reason",
    ]
    existing = [column for column in item_columns if column in result.columns]
    candidates = result[result["variant_group_size"].fillna(0).astype(str).str.isnumeric()].copy()
    candidates = candidates[candidates["variant_group_size"].astype(int) >= MIN_VARIANT_GROUP_SIZE]
    return candidates[existing].rename(
        columns={
            sku_col: "sku",
            title_col: "title",
            tags_col: "tags_after",
        }
    )


def unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = compact_spaces(str(value))
        key = normalize_header(text)
        if text and key not in seen:
            seen.add(key)
            result.append(text)
    return result


def clip(value: Any, max_length: int) -> str:
    text = compact_spaces(str(value or ""))
    if len(text) <= max_length:
        return text
    return text[: max_length - 3] + "..."


if __name__ == "__main__":
    main()
