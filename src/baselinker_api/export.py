from __future__ import annotations

import csv
import json
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable

from baselinker_api.client import BaseLinkerClient, RateLimiter, call_with_retry


EXPORT_COLUMNS = [
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


@dataclass
class ExportOptions:
    inventory_id: int
    include_variants: bool = False
    category_id: int | None = None
    manufacturer_id: int | None = None
    sku: str = ""
    ean: str = ""
    language: str = "pl"
    requests_per_minute: int = 100
    max_retries: int = 5
    backoff_base_seconds: float = 1.5


def _int(value: Any) -> int | None:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _api(fn, *, limiter: RateLimiter, options: ExportOptions, log: Callable[[str], None]):
    return call_with_retry(
        fn,
        limiter=limiter,
        max_retries=options.max_retries,
        backoff_base_seconds=options.backoff_base_seconds,
        log=log,
    )


def _category_paths(categories: list[dict[str, Any]]) -> dict[int, str]:
    by_id = {
        category_id: category
        for category in categories
        if (category_id := _int(category.get("category_id", category.get("id")))) is not None
    }
    cache: dict[int, str] = {}

    def build(category_id: int, visiting: set[int] | None = None) -> str:
        if category_id in cache:
            return cache[category_id]
        category = by_id.get(category_id, {})
        name = " ".join(str(category.get("name", "")).split())
        parent_id = _int(category.get("parent_id")) or 0
        active = set(visiting or ())
        if parent_id in by_id and category_id not in active:
            active.add(category_id)
            parent = build(parent_id, active)
            name = f"{parent} > {name}" if parent else name
        cache[category_id] = name
        return name

    return {category_id: build(category_id) for category_id in by_id}


def _manufacturer_names(manufacturers: list[dict[str, Any]]) -> dict[int, str]:
    output: dict[int, str] = {}
    for manufacturer in manufacturers:
        manufacturer_id = _int(manufacturer.get("manufacturer_id", manufacturer.get("id")))
        if manufacturer_id is not None:
            output[manufacturer_id] = str(manufacturer.get("name", manufacturer.get("manufacturer_name", "")))
    return output


def _text_value(text_fields: dict[str, Any], field_name: str, language: str) -> Any:
    for key in (f"{field_name}|{language}" if language else "", field_name):
        if key and key in text_fields:
            return text_fields.get(key)
    return ""


def _image_urls(images: Any) -> str:
    if not isinstance(images, dict):
        return ""
    values: list[tuple[int, str]] = []
    for position, raw_url in images.items():
        if "|" in str(position):
            continue
        try:
            numeric_position = int(position)
        except (TypeError, ValueError):
            continue
        url = str(raw_url or "").strip()
        if url:
            values.append((numeric_position, url))
    return "|".join(url for _, url in sorted(values))


def iter_inventory_export(
    client: BaseLinkerClient,
    options: ExportOptions,
    *,
    log: Callable[[str], None] | None = None,
) -> Iterable[dict[str, str]]:
    logger = log or (lambda _: None)
    limiter = RateLimiter(options.requests_per_minute)
    categories = _api(
        lambda: client.get_categories(options.inventory_id), limiter=limiter, options=options, log=logger
    )
    manufacturers = _api(client.get_manufacturers, limiter=limiter, options=options, log=logger)
    category_paths = _category_paths(categories)
    manufacturer_names = _manufacturer_names(manufacturers)

    page = 1
    while True:
        products = _api(
            lambda page=page: client.get_products_list(
                options.inventory_id,
                page=page,
                include_variants=options.include_variants,
                filter_category_id=options.category_id,
                filter_sku=options.sku or None,
                filter_ean=options.ean or None,
            ),
            limiter=limiter,
            options=options,
            log=logger,
        )
        if not products:
            break
        product_ids = [
            product_id
            for product in products
            if (product_id := _int(product.get("id", product.get("product_id")))) is not None
        ]
        details: dict[str, dict[str, Any]] = {}
        for index in range(0, len(product_ids), 50):
            batch = product_ids[index : index + 50]
            details.update(
                _api(
                    lambda batch=batch: client.get_products_data(options.inventory_id, batch),
                    limiter=limiter,
                    options=options,
                    log=logger,
                )
            )
        for product in products:
            product_id = _int(product.get("id", product.get("product_id")))
            if product_id is None:
                continue
            detail = details.get(str(product_id), {})
            manufacturer_id = _int(detail.get("manufacturer_id"))
            if options.manufacturer_id is not None and manufacturer_id != options.manufacturer_id:
                continue
            category_id = _int(detail.get("category_id"))
            raw_text_fields = detail.get("text_fields", {})
            text_fields = raw_text_fields if isinstance(raw_text_fields, dict) else {}
            features = _text_value(text_fields, "features", options.language)
            if isinstance(features, dict):
                features_text = json.dumps(features, ensure_ascii=False, separators=(",", ":"))
            elif features:
                features_text = str(features)
            else:
                features_text = "{}"
            yield {
                "product_id": str(product_id),
                "name": str(_text_value(text_fields, "name", options.language) or product.get("name", "")),
                "sku": str(detail.get("sku", product.get("sku", "")) or ""),
                "ean": str(detail.get("ean", product.get("ean", "")) or ""),
                "manufacturer_name": manufacturer_names.get(manufacturer_id or -1, ""),
                "category": category_paths.get(category_id or -1, ""),
                "description": str(_text_value(text_fields, "description", options.language) or ""),
                "features": features_text,
                "images_urls": _image_urls(detail.get("images")),
            }
        logger(f"Eksport BaseLinkera: strona {page}, rekordów {len(products)}.")
        page += 1


def export_inventory_csv(
    client: BaseLinkerClient,
    output_path: str | Path,
    options: ExportOptions,
    *,
    log: Callable[[str], None] | None = None,
) -> int:
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(prefix=destination.name + ".", suffix=".tmp", dir=destination.parent)
    count = 0
    try:
        with os.fdopen(handle, "w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=EXPORT_COLUMNS, delimiter=";")
            writer.writeheader()
            for row in iter_inventory_export(client, options, log=log):
                writer.writerow(row)
                count += 1
        os.replace(temp_name, destination)
    except Exception:
        Path(temp_name).unlink(missing_ok=True)
        raise
    return count
