"""Zapis wiedzy o sklepie z /knowledge do słowników AlkanInatora.

Pliki TSV mają ten sam układ co zrzuty z bazy (scripts/dump_woo_*.py), więc
istniejący kod (shop_attribute_values.py, build_product_removal_redirects.py)
czyta je bez zmian.
"""

from __future__ import annotations

import csv
import json
from datetime import date
from pathlib import Path
from typing import Any
from urllib.parse import urlparse


ATTRIBUTE_COLUMNS = ["attribute", "taxonomy", "term", "slug", "count", "parent_term"]
CATEGORY_COLUMNS = ["term_id", "parent_id", "name", "slug", "count"]
SHOP_TERM_COLUMNS = ["taxonomy", "term_id", "name", "slug", "count"]
GENERATOR = "py src/woo_api_cli.py knowledge"


def _write_tsv(path: Path, title: str, columns: list[str], rows: list[list[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as stream:
        stream.write(f"# {title}, zrzut {date.today().isoformat()}.\n")
        stream.write(f"# Generuje: {GENERATOR} - nie edytowac recznie.\n")
        writer = csv.writer(stream, delimiter="\t", lineterminator="\n")
        writer.writerow(columns)
        writer.writerows(rows)


def attribute_rows(knowledge: dict[str, Any]) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for attribute in sorted(knowledge.get("attributes", []), key=lambda item: item["label"].casefold()):
        terms = sorted(attribute.get("terms", []), key=lambda item: item["name"].casefold())
        if not terms:
            rows.append([attribute["label"], attribute["taxonomy"], "", "", 0, ""])
        for term in terms:
            rows.append([attribute["label"], attribute["taxonomy"], term["name"], term["slug"], term["count"], term.get("parent_term", "")])
    return rows


def category_rows(knowledge: dict[str, Any]) -> list[list[Any]]:
    categories = sorted(knowledge.get("categories", []), key=lambda item: (item["parent_id"], item["name"].casefold()))
    return [[item["term_id"], item["parent_id"], item["name"], item["slug"], item["count"]] for item in categories]


def shop_term_rows(knowledge: dict[str, Any]) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for key, taxonomy in (("shipping_classes", "product_shipping_class"), ("shipping_within", "shipping_within"), ("tags", "product_tag")):
        for term in sorted(knowledge.get(key, []), key=lambda item: item["name"].casefold()):
            rows.append([taxonomy, term["term_id"], term["name"], term["slug"], term["count"]])
    return rows


def write_knowledge(knowledge: dict[str, Any], dictionaries_dir: Path, cache_path: Path) -> dict[str, Path]:
    source = urlparse(knowledge.get("site_url", "")).netloc or "sklep"
    paths = {
        "attributes": dictionaries_dir / "woo_attribute_terms_snapshot.tsv",
        "categories": dictionaries_dir / "woo_product_categories_snapshot.tsv",
        "shop_terms": dictionaries_dir / "woo_shop_terms_snapshot.tsv",
        "full": cache_path,
    }
    _write_tsv(paths["attributes"], f"Terminy atrybutow Woo z API sklepu {source}", ATTRIBUTE_COLUMNS, attribute_rows(knowledge))
    _write_tsv(paths["categories"], f"Kategorie produktow Woo z API sklepu {source}", CATEGORY_COLUMNS, category_rows(knowledge))
    _write_tsv(
        paths["shop_terms"],
        f"Klasy wysylkowe, terminy 'Wysylka w ciagu' i tagi Woo z API sklepu {source}",
        SHOP_TERM_COLUMNS,
        shop_term_rows(knowledge),
    )
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(json.dumps(knowledge, ensure_ascii=False, indent=1), encoding="utf-8")
    return paths
