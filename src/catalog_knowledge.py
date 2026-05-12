from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from utils import load_yaml, strip_accents


DEFAULT_CATALOG_KNOWLEDGE_PATH = "dictionaries/woocommerce_catalog_knowledge.yaml"

ATTRIBUTE_FILTER_MAP = {
    "Producent": "producent",
    "Kolor": "kolor",
    "Kolor producenta": "kolor",
    "Seria": "seria",
    "Typ produktu": "typ",
    "Podtyp produktu": "rodzaj",
    "Rodzaj": "rodzaj",
    "Stopien ochrony [IP]": "ip",
    "Stopień ochrony [IP]": "ip",
    "Sposob montazu": "sposob_montazu",
    "Sposób montażu": "sposob_montazu",
    "Napiecie [V]": "napiecie",
    "Napięcie [V]": "napiecie",
    "Moc [W]": "moc",
    "Temperatura barwowa [K]": "barwa",
    "Strumien swietlny [lm]": "strumien",
    "Strumień świetlny [lm]": "strumien",
    "Barwa swiatla": "barwa",
    "Barwa światła": "barwa",
    "Trzonek": "gwint",
    "Material": "material",
    "Materiał": "material",
}


def load_catalog_knowledge(path: str | Path | None = DEFAULT_CATALOG_KNOWLEDGE_PATH) -> dict[str, Any]:
    if not path:
        return {}
    return load_yaml(path)


def enrich_config_with_catalog_knowledge(config: dict[str, Any], knowledge: dict[str, Any]) -> dict[str, Any]:
    if not knowledge:
        return config

    enriched = deepcopy(config)
    enriched["catalog_knowledge_summary"] = knowledge.get("summary", {})
    enriched["known_category_paths"] = _values(knowledge.get("top_categories"), "category", 500)
    enriched["known_category_tree"] = knowledge.get("category_tree") or []
    enriched["known_series"] = _values(knowledge.get("series"), "series", 300)
    enriched["known_product_types"] = _values(knowledge.get("types"), "type", 300)
    enriched["known_product_subtypes"] = _values(knowledge.get("subtypes"), "subtype", 300)
    enriched["known_ip_values"] = _values(knowledge.get("ip_values"), "ip", 100)

    enriched["known_producers"] = _merge_list(
        enriched.get("known_producers") or [],
        _values(knowledge.get("producers"), "producer", 300),
    )
    enriched["known_colors"] = _merge_known_colors(enriched.get("known_colors") or {}, knowledge.get("colors") or [])
    enriched["recommended_filters"] = _merge_list(
        enriched.get("recommended_filters") or [],
        _recommended_filters_from_catalog(knowledge),
    )
    return enriched


def _values(rows: Any, key: str, limit: int) -> list[str]:
    values: list[str] = []
    for row in list(rows or [])[:limit]:
        value = str(row.get(key, "")).strip() if isinstance(row, dict) else ""
        if value:
            values.append(value)
    return values


def _merge_list(current: list[Any], learned: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in [*current, *learned]:
        text = str(value).strip()
        key = strip_accents(text).lower()
        if not text or key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def _merge_known_colors(current: dict[str, Any], learned_rows: list[dict[str, Any]]) -> dict[str, list[str]]:
    colors: dict[str, list[str]] = {
        str(name): [str(value) for value in variants or []]
        for name, variants in current.items()
    }
    for row in learned_rows:
        color = str(row.get("color", "")).strip()
        if not color:
            continue
        key = strip_accents(color).lower()
        variants = colors.setdefault(key, [])
        for variant in {color, color.lower(), key}:
            if variant and variant not in variants:
                variants.append(variant)
    return colors


def _recommended_filters_from_catalog(knowledge: dict[str, Any]) -> list[str]:
    filters: list[str] = []
    for row in knowledge.get("top_attributes") or []:
        attribute = str(row.get("attribute", "")).strip()
        mapped = ATTRIBUTE_FILTER_MAP.get(attribute) or ATTRIBUTE_FILTER_MAP.get(strip_accents(attribute))
        if mapped:
            filters.append(mapped)
    return filters
