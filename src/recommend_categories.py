from __future__ import annotations

import argparse
import math
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, load_catalog_knowledge
from utils import compact_spaces, detect_column, ensure_dir, is_blank, load_yaml, normalize_header, read_products, write_products


DEFAULT_TAXONOMY_PATH = "dictionaries/learned_product_taxonomy.yaml"
DEFAULT_OVERRIDES_PATH = "dictionaries/category_recommendation_overrides.yaml"
MIN_RECOMMENDATION_SCORE = 12.0

TITLE_COLUMNS = [
    "new_title",
    "name",
    "Nazwa",
    "Nazwa produktu",
    "Nazwa B2C / SEO final",
    "Nazwa B2C / SEO",
    "gazetki_new_title",
    "gazetki_Nazwa produktu",
]
TYPE_COLUMNS = ["Typ produktu", "Typ", "attr_typ", "gazetki_attr_typ", "product_type"]
PRODUCER_COLUMNS = ["Producent", "manufacturer_name", "Marka", "attr_producent", "gazetki_attr_producent"]
SERIES_COLUMNS = ["Seria", "attr_seria", "gazetki_attr_seria"]
SKU_COLUMNS = ["SKU", "sku", "Kod", "Kod Producenta", "PRODUCT_CODE", "TOWARIDX"]
CATEGORY_COLUMNS = ["Kategoria", "category", "Kategorie", "categories"]

ATTRIBUTE_CATEGORY_WEIGHTS = {
    "Typ produktu": 24.0,
    "Rodzaj produktu": 20.0,
    "Podtyp produktu": 18.0,
    "Seria": 14.0,
    "Producent": 8.0,
    "Marka": 8.0,
}

STOP_WORDS = {
    "do",
    "i",
    "oraz",
    "z",
    "ze",
    "w",
    "na",
    "dla",
    "typ",
    "produkt",
    "produkty",
    "kolor",
    "bialy",
    "biala",
    "czarny",
    "czarna",
    "szary",
    "szara",
    "led",
    "ip",
}

KEYWORD_CATEGORY_HINTS = [
    (["peszel", "rura gietka", "rura giętka", "karbowana", "uv"], ["peszle", "rury odporne na uv"], 22.0),
    (["rura", "peszel", "karbowana"], ["peszle", "rury", "systemy prowadzenia kabli"], 12.0),
    (["przedluzacz bebnowy", "przedłużacz bębnowy", "zwijacz"], ["przedluzacze na bebnie", "przedłużacze na bębnie"], 22.0),
    (["przedluzacz", "przedlużacz", "przedłużacz", "listwa przepieciowa", "listwa przeciwprzepieciowa"], ["przedluzacze", "przedłużacze"], 12.0),
    (["gniazdo hermetyczne"], ["gniazdka i laczniki", "gniazdka i łączniki"], 24.0),
    (["gniazdo", "wtyk", "combo", "cee"], ["gniazda", "osprzet silowy", "osprzęt siłowy"], 12.0),
    (["rozdzielnica natynkowa"], ["rozdzielnice natynkowe"], 24.0),
    (["rozdzielnica przenosna", "rozdzielnica przenośna", "opole"], ["rozdzielnie elektryczne"], 24.0),
    (["obudowa rozdzielnicy", "erp18"], ["skrzynki rozdzielcze"], 24.0),
    (["rozdzielnica", "obudowa rozdzielnicy"], ["rozdzielnice"], 12.0),
    (["puszka do pustych scian", "puszka do pustych ścian"], ["puszki podtynkowe"], 24.0),
    (["puszka"], ["puszki"], 12.0),
    (["zlaczka", "złączka", "blok dystrybucyjny", "hlak", "otl"], ["zlaczki", "złączki", "bloki rozdzielcze"], 12.0),
    (["ogranicznik przepiec", "ogranicznik przepięć", "spd"], ["ochronniki przepieciowe", "ograniczniki", "aparatura modulowa"], 16.0),
    (["kzs", "rcbo", "roznicowopradowy", "różnicowoprądowy"], ["wylaczniki roznicowopradowe", "wyłączniki różnicowoprądowe"], 26.0),
    (["wylacznik", "wyłącznik", "rozlacznik", "rozłącznik"], ["wylaczniki", "wyłączniki", "aparatura modulowa"], 12.0),
    (["solarflex", "h1z2z2", "fotowoltaiczny", "mc4"], ["fotowoltaika", "akcesoria do fotowoltaiki"], 26.0),
    (["panel led", "backlight", "bls", "blo"], ["panele led", "panele"], 18.0),
    (["oprawa hermetyczna", "damp proof", "balwir"], ["lampy hermetyczne", "hermetyczne"], 22.0),
    (["naswietlacz", "naświetlacz"], ["naswietlacze", "naświetlacze"], 18.0),
    (["plafoniera", "plafon"], ["plafony"], 18.0),
    (["latarka"], ["latarki"], 18.0),
    (["lampa ogrodowa", "stono", "reka"], ["lampy ogrodowe"], 18.0),
    (["zarowka", "żarówka"], ["zarowki", "żarówki"], 18.0),
    (["reflektor roboczy"], ["oswietlenie", "oświetlenie"], 18.0),
    (["oprawa", "lampa"], ["oprawy", "lampy"], 10.0),
    (["kabel", "przewod", "przewód"], ["kable", "przewody"], 10.0),
    (["silikon", "tasma ostrzegawcza", "taśma ostrzegawcza", "wkretak", "wkrętak", "szczypce", "zaciskarka", "osadzak"], ["narzedzia", "narzędzia"], 18.0),
    (["spodnie robocze", "kurtka robocza", "bluza robocza", "odziez robocza", "odzież robocza"], ["narzedzia", "narzędzia"], 14.0),
    (["uziom", "odgrom", "zlacze krzyzowe", "złącze krzyżowe"], ["ochrona odgromowa"], 18.0),
]

SOURCE_CATEGORY_HINTS = [
    (["panele led"], ["Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Panele LED sufitowe"], 42.0),
    (
        ["oprawy pyloszczelne", "oprawy pyłoszczelne", "hermetyczne"],
        ["Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Lampy hermetyczne"],
        42.0,
    ),
    (["high bay", "oswietlenie hal", "oświetlenie hal"], ["Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Oprawy High-bay LED"], 42.0),
    (["naswietlacze led", "naświetlacze led"], ["Oświetlenie > Oświetlenie zewnętrzne > Naświetlacze LED"], 42.0),
    (["plafoniery", "oprawy sufitowe"], ["Oświetlenie > Oświetlenie wewnętrzne > Plafony LED"], 42.0),
]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Proponuje kategorie produktow na podstawie nazwy, typu produktu i wiedzy katalogowej."
    )
    parser.add_argument("--input", required=True, help="CSV/XLSX z produktami.")
    parser.add_argument("--sheet", help="Arkusz XLSX.")
    parser.add_argument("--output", default="output/category_recommendations.xlsx", help="Plik wynikowy CSV/XLSX.")
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE_PATH)
    parser.add_argument("--taxonomy", default=DEFAULT_TAXONOMY_PATH)
    parser.add_argument("--overrides", default=DEFAULT_OVERRIDES_PATH)
    parser.add_argument("--top-n", type=int, default=3, help="Liczba propozycji kategorii na produkt.")
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    catalog_knowledge = load_catalog_knowledge(args.catalog_knowledge)
    taxonomy = load_yaml(args.taxonomy) if args.taxonomy else {}
    overrides = load_yaml(args.overrides) if args.overrides else {}
    recommender = CategoryRecommender(catalog_knowledge, taxonomy, overrides)
    result = recommend_for_dataframe(df, recommender, args.top_n)
    write_products(result, Path(args.output))
    write_summary(result, args.output)

    print(f"OK: wczytano {len(df)} produktow")
    print(f"OK: zapisano {args.output}")


class CategoryRecommender:
    def __init__(self, catalog_knowledge: dict[str, Any], taxonomy: dict[str, Any], overrides: dict[str, Any] | None = None) -> None:
        self.category_counts = build_category_counts(catalog_knowledge, taxonomy)
        self.leaf_to_paths = build_leaf_to_paths(self.category_counts)
        self.category_tokens = {path: token_set(path) for path in self.category_counts}
        self.attribute_values = build_attribute_value_index(catalog_knowledge)
        self.type_by_category = build_type_by_category_index(taxonomy, self.leaf_to_paths)
        self.sku_category_overrides = build_sku_category_overrides(overrides or {})

    def recommend(self, row: pd.Series, columns: dict[str, str | None], top_n: int) -> list[dict[str, Any]]:
        title = first_value(row, [columns.get("title")])
        product_type = first_value(row, [columns.get("type")])
        producer = first_value(row, [columns.get("producer")])
        series = first_value(row, [columns.get("series")])
        source_category = first_value(row, [columns.get("category")])
        sku = normalize_sku(first_value(row, [columns.get("sku")]))

        query = compact_spaces(" ".join([title, product_type, producer, series]))
        query_norm = normalize_text(query)
        query_tokens = token_set(query)
        scores: dict[str, float] = defaultdict(float)
        reasons: dict[str, list[str]] = defaultdict(list)

        self.add_lexical_signals(query_norm, query_tokens, scores, reasons)
        self.add_attribute_signals(product_type, producer, series, scores, reasons)
        self.add_taxonomy_type_signals(product_type, query_tokens, scores, reasons)
        self.add_keyword_hints(query_norm, query_tokens, scores, reasons)
        self.add_source_category_hints(source_category, scores, reasons)
        self.add_sku_override(sku, scores, reasons)

        ranked = [
            item
            for item in sorted(scores.items(), key=lambda item: (item[1], self.category_counts.get(item[0], 0)), reverse=True)
            if item[1] >= MIN_RECOMMENDATION_SCORE
        ]
        recommendations: list[dict[str, Any]] = []
        for category, score in ranked[:top_n]:
            recommendations.append(
                {
                    "category": category,
                    "score": round(score, 2),
                    "confidence": confidence_label(score),
                    "reasons": "; ".join(unique(reasons[category])[:5]),
                }
            )
        return recommendations

    def add_lexical_signals(
        self,
        query_norm: str,
        query_tokens: set[str],
        scores: dict[str, float],
        reasons: dict[str, list[str]],
    ) -> None:
        for category, category_tokens in self.category_tokens.items():
            overlap = query_tokens & category_tokens
            if not overlap:
                continue
            leaf = category_leaf(category)
            leaf_norm = normalize_text(leaf)
            score = len(overlap) * 2.0 + popularity_boost(self.category_counts.get(category, 0))
            if leaf_norm and leaf_norm in query_norm:
                score += 16.0
                reasons[category].append(f"nazwa zawiera lisc kategorii: {leaf}")
            elif len(overlap) >= 2:
                score += 5.0
                reasons[category].append(f"wspolne slowa z kategoria: {', '.join(sorted(overlap)[:4])}")
            else:
                reasons[category].append(f"slowo wspolne z kategoria: {next(iter(overlap))}")
            scores[category] += score

    def add_attribute_signals(
        self,
        product_type: str,
        producer: str,
        series: str,
        scores: dict[str, float],
        reasons: dict[str, list[str]],
    ) -> None:
        values = [
            ("Typ produktu", product_type),
            ("Rodzaj produktu", product_type),
            ("Podtyp produktu", product_type),
            ("Producent", producer),
            ("Marka", producer),
            ("Seria", series),
        ]
        for attribute, value in values:
            value_key = normalize_text(value)
            if not value_key:
                continue
            for category, count in self.attribute_values.get((attribute, value_key), []):
                weight = ATTRIBUTE_CATEGORY_WEIGHTS.get(attribute, 8.0)
                scores[category] += weight + popularity_boost(count)
                reasons[category].append(f"{attribute}: {value}")

    def add_taxonomy_type_signals(
        self,
        product_type: str,
        query_tokens: set[str],
        scores: dict[str, float],
        reasons: dict[str, list[str]],
    ) -> None:
        product_type_key = normalize_text(product_type)
        for learned_type, rows in self.type_by_category.items():
            learned_tokens = token_set(learned_type)
            if not learned_tokens:
                continue
            exact = product_type_key and (product_type_key == learned_type or product_type_key in learned_type or learned_type in product_type_key)
            overlap = query_tokens & learned_tokens
            if not exact and len(overlap) < 2:
                continue
            for category, count in rows:
                if exact:
                    scores[category] += 30.0 + popularity_boost(count)
                    reasons[category].append(f"nauczony typ: {learned_type}")
                else:
                    scores[category] += 10.0 + popularity_boost(count)
                    reasons[category].append(f"typ podobny: {learned_type}")

    def add_keyword_hints(
        self,
        query_norm: str,
        query_tokens: set[str],
        scores: dict[str, float],
        reasons: dict[str, list[str]],
    ) -> None:
        for needles, category_needles, weight in KEYWORD_CATEGORY_HINTS:
            if not any(contains_normalized_phrase(query_norm, query_tokens, needle) for needle in needles):
                continue
            for category in self.category_counts:
                category_norm = normalize_text(category)
                category_tokens = self.category_tokens.get(category) or token_set(category)
                if any(contains_normalized_phrase(category_norm, category_tokens, needle) for needle in category_needles):
                    scores[category] += weight + popularity_boost(self.category_counts.get(category, 0))
                    reasons[category].append("regula slow kluczowych")

    def add_source_category_hints(
        self,
        source_category: str,
        scores: dict[str, float],
        reasons: dict[str, list[str]],
    ) -> None:
        source_norm = normalize_text(source_category)
        source_tokens = token_set(source_category)
        if not source_norm:
            return
        for needles, categories, weight in SOURCE_CATEGORY_HINTS:
            if not any(contains_normalized_phrase(source_norm, source_tokens, needle) for needle in needles):
                continue
            for category in categories:
                self.category_counts.setdefault(category, 1)
                self.category_tokens.setdefault(category, token_set(category))
                scores[category] += weight + popularity_boost(self.category_counts.get(category, 0))
                reasons[category].append(f"kategoria z pliku: {source_category}")

    def add_sku_override(
        self,
        sku: str,
        scores: dict[str, float],
        reasons: dict[str, list[str]],
    ) -> None:
        category = self.sku_category_overrides.get(sku)
        if not category:
            return
        self.category_counts.setdefault(category, 1)
        self.category_tokens.setdefault(category, token_set(category))
        scores[category] += 120.0
        reasons[category].append("zaakceptowana korekta po SKU")


def recommend_for_dataframe(df: pd.DataFrame, recommender: CategoryRecommender, top_n: int) -> pd.DataFrame:
    columns = detect_recommendation_columns(df)
    result = df.copy()
    result["category_recommendation_title_column"] = columns.get("title") or ""
    result["category_recommendation_type_column"] = columns.get("type") or ""

    recommendation_rows = [recommender.recommend(row, columns, top_n) for _, row in result.iterrows()]
    for index in range(top_n):
        result[f"proponowana_kategoria_{index + 1}"] = [
            recommendations[index]["category"] if index < len(recommendations) else "" for recommendations in recommendation_rows
        ]
        result[f"category_score_{index + 1}"] = [
            recommendations[index]["score"] if index < len(recommendations) else "" for recommendations in recommendation_rows
        ]
        result[f"category_confidence_{index + 1}"] = [
            recommendations[index]["confidence"] if index < len(recommendations) else "" for recommendations in recommendation_rows
        ]
        result[f"category_reasons_{index + 1}"] = [
            recommendations[index]["reasons"] if index < len(recommendations) else "" for recommendations in recommendation_rows
        ]
    result["category_recommendation_status"] = [
        "OK" if recommendations else "BRAK_PROPOZYCJI" for recommendations in recommendation_rows
    ]
    return result


def detect_recommendation_columns(df: pd.DataFrame) -> dict[str, str | None]:
    return {
        "title": detect_column(df, TITLE_COLUMNS),
        "type": detect_column(df, TYPE_COLUMNS),
        "producer": detect_column(df, PRODUCER_COLUMNS),
        "series": detect_column(df, SERIES_COLUMNS),
        "sku": detect_column(df, SKU_COLUMNS),
        "category": detect_column(df, CATEGORY_COLUMNS),
    }


def build_category_counts(catalog_knowledge: dict[str, Any], taxonomy: dict[str, Any]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for row in catalog_knowledge.get("top_categories") or []:
        add_category_count(counts, row.get("category", ""), row.get("count", 0))
    for row in taxonomy.get("categories") or []:
        add_category_count(counts, row.get("path", ""), row.get("count", 0))
    return counts


def add_category_count(counts: dict[str, int], category: Any, count: Any) -> None:
    path = normalize_category_path(category)
    if not path:
        return
    counts[path] = max(counts.get(path, 0), int(count or 0))


def build_leaf_to_paths(category_counts: dict[str, int]) -> dict[str, list[str]]:
    result: dict[str, list[str]] = defaultdict(list)
    for path in category_counts:
        result[normalize_text(category_leaf(path))].append(path)
    return result


def build_attribute_value_index(catalog_knowledge: dict[str, Any]) -> dict[tuple[str, str], list[tuple[str, int]]]:
    result: dict[tuple[str, str], list[tuple[str, int]]] = defaultdict(list)
    for row in catalog_knowledge.get("top_attribute_values_by_category") or []:
        attribute = compact_spaces(str(row.get("attribute", "")))
        if attribute not in ATTRIBUTE_CATEGORY_WEIGHTS:
            continue
        category = normalize_category_path(row.get("category", ""))
        value = normalize_text(row.get("value", ""))
        if not category or not value:
            continue
        result[(attribute, value)].append((category, int(row.get("count") or 0)))
    return result


def build_type_by_category_index(
    taxonomy: dict[str, Any],
    leaf_to_paths: dict[str, list[str]],
) -> dict[str, list[tuple[str, int]]]:
    result: dict[str, list[tuple[str, int]]] = defaultdict(list)
    for row in taxonomy.get("type_by_category") or []:
        learned_type = normalize_text(row.get("type", ""))
        leaf = normalize_text(row.get("category", ""))
        if not learned_type or not leaf:
            continue
        paths = leaf_to_paths.get(leaf) or [normalize_category_path(row.get("category", ""))]
        for path in paths:
            if path:
                result[learned_type].append((path, int(row.get("count") or 0)))
    return result


def build_sku_category_overrides(overrides: dict[str, Any]) -> dict[str, str]:
    rows = overrides.get("sku_category_overrides") or {}
    if not isinstance(rows, dict):
        return {}
    result: dict[str, str] = {}
    for sku, category in rows.items():
        normalized_sku = normalize_sku(str(sku))
        normalized_category = normalize_category_path(category)
        if normalized_sku and normalized_category:
            result[normalized_sku] = normalized_category
    return result


def normalize_sku(value: Any) -> str:
    sku = compact_spaces(str(value or ""))
    if "/" in sku:
        sku = sku.rsplit("/", 1)[0]
    if re.fullmatch(r"\d+\.0", sku):
        sku = sku[:-2]
    return sku.upper()


def normalize_category_path(value: Any) -> str:
    text = compact_spaces(str(value or ""))
    text = re.sub(r"\s*>\s*", " > ", text)
    return compact_spaces(text)


def category_leaf(path: str) -> str:
    parts = [compact_spaces(part) for part in normalize_category_path(path).split(">") if compact_spaces(part)]
    return parts[-1] if parts else ""


def token_set(value: Any) -> set[str]:
    normalized = normalize_text(value)
    tokens = set(re.findall(r"[a-z0-9]{2,}", normalized))
    return {token for token in tokens if token not in STOP_WORDS and not token.isdigit()}


def contains_normalized_phrase(haystack_norm: str, haystack_tokens: set[str], needle: Any) -> bool:
    needle_norm = normalize_text(needle)
    if not needle_norm:
        return False
    if " " not in needle_norm:
        return needle_norm in haystack_tokens
    return re.search(rf"(^| ){re.escape(needle_norm)}($| )", haystack_norm) is not None


def normalize_text(value: Any) -> str:
    return normalize_header(str(value or ""))


def first_value(row: pd.Series, columns: list[str | None]) -> str:
    for column in columns:
        if column and column in row and not is_blank(row.get(column, "")):
            return compact_spaces(str(row.get(column, "")))
    return ""


def popularity_boost(count: int) -> float:
    return min(6.0, math.log10(max(int(count), 1) + 1))


def confidence_label(score: float) -> str:
    if score >= 45:
        return "wysoka"
    if score >= 25:
        return "srednia"
    return "niska"


def unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = value.lower()
        if value and key not in seen:
            seen.add(key)
            result.append(value)
    return result


def write_summary(result: pd.DataFrame, output: str | Path) -> None:
    output_path = Path(output)
    reports_dir = ensure_dir(Path("reports") / output_path.stem)
    summary = (
        result["proponowana_kategoria_1"]
        .value_counts(dropna=False)
        .rename_axis("category")
        .reset_index(name="recommended_products")
    )
    summary.to_csv(reports_dir / "category_recommendation_summary.csv", index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    main()
