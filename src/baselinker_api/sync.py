from __future__ import annotations

import csv
import json
import os
import re
import tempfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from baselinker_api.client import BaseLinkerClient, RateLimiter, call_with_retry
from baselinker_api.images import build_images_payload
from baselinker_api.input_data import ProductRecord, load_product_records


LogFn = Callable[[str], None]
DEFAULT_FIELDS = {"name", "manufacturer", "category", "description", "features", "images"}


@dataclass
class SyncOptions:
    inventory_id: int
    apply: bool = False
    fields: set[str] = field(default_factory=lambda: set(DEFAULT_FIELDS))
    match_by: str = "auto"
    include_variants: bool = False
    merge_features: bool = True
    update_identifiers: bool = False
    create_missing_categories: bool = False
    create_missing_manufacturers: bool = False
    allow_partial: bool = False
    allow_http_images: bool = False
    language: str = "pl"
    requests_per_minute: int = 100
    max_retries: int = 5
    backoff_base_seconds: float = 1.5


@dataclass
class AuditRow:
    row_number: int
    product_id: str
    sku: str
    ean: str
    status: str
    changed_fields: str = ""
    warnings: str = ""
    error: str = ""
    before_json: str = ""
    after_json: str = ""
    payload: dict[str, Any] | None = field(default=None, repr=False)


@dataclass
class SyncResult:
    report_dir: Path
    audit_rows: list[AuditRow]
    applied: bool
    blocked: bool


def _api(
    fn: Callable[[], Any],
    *,
    limiter: RateLimiter,
    options: SyncOptions,
    log: LogFn,
) -> Any:
    return call_with_retry(
        fn,
        limiter=limiter,
        max_retries=options.max_retries,
        backoff_base_seconds=options.backoff_base_seconds,
        log=log,
    )


def _int(value: Any) -> int | None:
    try:
        text = str(value).strip()
        return int(float(text)) if text else None
    except (TypeError, ValueError):
        return None


def _norm_sku(value: str) -> str:
    return " ".join(str(value or "").split()).casefold()


def _norm_ean(value: str) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _norm_path(value: str) -> str:
    return " > ".join(" ".join(part.split()) for part in str(value or "").split(">") if part.strip()).casefold()


def _norm_name(value: str) -> str:
    return " ".join(str(value or "").split()).casefold()


def _batches(values: list[int], size: int = 50) -> list[list[int]]:
    return [values[index : index + size] for index in range(0, len(values), size)]


def _fetch_product_index(
    client: BaseLinkerClient,
    options: SyncOptions,
    limiter: RateLimiter,
    log: LogFn,
) -> tuple[dict[str, int], set[str], dict[str, int], set[str], set[int]]:
    sku_map: dict[str, int] = {}
    ean_map: dict[str, int] = {}
    duplicate_skus: set[str] = set()
    duplicate_eans: set[str] = set()
    product_ids: set[int] = set()
    page = 1
    while True:
        products = _api(
            lambda page=page: client.get_products_list(
                options.inventory_id,
                page=page,
                include_variants=options.include_variants,
            ),
            limiter=limiter,
            options=options,
            log=log,
        )
        if not products:
            break
        new_ids = 0
        for product in products:
            product_id = _int(product.get("id", product.get("product_id")))
            if product_id is None:
                continue
            if product_id not in product_ids:
                new_ids += 1
            product_ids.add(product_id)
            for raw, mapping, duplicates, normalizer in (
                (product.get("sku", ""), sku_map, duplicate_skus, _norm_sku),
                (product.get("ean", ""), ean_map, duplicate_eans, _norm_ean),
            ):
                normalized = normalizer(str(raw or ""))
                if not normalized:
                    continue
                previous = mapping.get(normalized)
                if previous is not None and previous != product_id:
                    duplicates.add(normalized)
                else:
                    mapping[normalized] = product_id
        log(f"BaseLinker: indeks produktów, strona {page}, rekordów {len(products)}.")
        if new_ids == 0:
            break
        page += 1
    for duplicate in duplicate_skus:
        sku_map.pop(duplicate, None)
    for duplicate in duplicate_eans:
        ean_map.pop(duplicate, None)
    return sku_map, duplicate_skus, ean_map, duplicate_eans, product_ids


def _fetch_details(
    client: BaseLinkerClient,
    product_ids: list[int],
    options: SyncOptions,
    limiter: RateLimiter,
    log: LogFn,
) -> dict[str, dict[str, Any]]:
    output: dict[str, dict[str, Any]] = {}
    for batch in _batches(sorted(set(product_ids))):
        output.update(
            _api(
                lambda batch=batch: client.get_products_data(options.inventory_id, batch),
                limiter=limiter,
                options=options,
                log=log,
            )
        )
    return output


def build_category_maps(categories: list[dict[str, Any]]) -> tuple[dict[str, int], set[str]]:
    by_id = {
        category_id: category
        for category in categories
        if (category_id := _int(category.get("category_id", category.get("id")))) is not None
    }
    cache: dict[int, str] = {}

    def path_for(category_id: int, visiting: set[int] | None = None) -> str:
        if category_id in cache:
            return cache[category_id]
        category = by_id.get(category_id, {})
        name = " ".join(str(category.get("name", "")).split())
        parent_id = _int(category.get("parent_id")) or 0
        visiting = set(visiting or ())
        if parent_id and parent_id in by_id and category_id not in visiting:
            visiting.add(category_id)
            parent = path_for(parent_id, visiting)
            value = f"{parent} > {name}" if parent else name
        else:
            value = name
        cache[category_id] = value
        return value

    mapping: dict[str, int] = {}
    duplicates: set[str] = set()
    for category_id, category in by_id.items():
        for candidate in (path_for(category_id), str(category.get("name", ""))):
            normalized = _norm_path(candidate)
            if not normalized:
                continue
            if normalized in mapping and mapping[normalized] != category_id:
                duplicates.add(normalized)
            else:
                mapping[normalized] = category_id
    for duplicate in duplicates:
        mapping.pop(duplicate, None)
    return mapping, duplicates


def build_manufacturer_maps(manufacturers: list[dict[str, Any]]) -> tuple[dict[str, int], set[str]]:
    mapping: dict[str, int] = {}
    duplicates: set[str] = set()
    for manufacturer in manufacturers:
        manufacturer_id = _int(manufacturer.get("manufacturer_id", manufacturer.get("id")))
        normalized = _norm_name(str(manufacturer.get("name", manufacturer.get("manufacturer_name", ""))))
        if manufacturer_id is None or not normalized:
            continue
        if normalized in mapping and mapping[normalized] != manufacturer_id:
            duplicates.add(normalized)
        else:
            mapping[normalized] = manufacturer_id
    for duplicate in duplicates:
        mapping.pop(duplicate, None)
    return mapping, duplicates


def _resolve_product_id(
    record: ProductRecord,
    *,
    options: SyncOptions,
    sku_map: dict[str, int],
    duplicate_skus: set[str],
    ean_map: dict[str, int],
    duplicate_eans: set[str],
    product_ids: set[int],
) -> tuple[int | None, str | None]:
    # W historycznych plikach AlkanInatora product_id bywa ID WooCommerce.
    # SKU jest stabilnym kluczem katalogu BaseLinkera i ma pierwszeństwo.
    modes = [options.match_by] if options.match_by != "auto" else ["sku", "ean", "product_id"]
    for mode in modes:
        if mode == "product_id" and record.product_id:
            product_id = _int(record.product_id)
            if product_id is None or product_id not in product_ids:
                return None, f"product_id {record.product_id!r} nie istnieje w wybranym katalogu."
            return product_id, None
        if mode == "sku" and record.sku:
            key = _norm_sku(record.sku)
            if key in duplicate_skus:
                return None, f"SKU {record.sku!r} jest nieunikalne w katalogu BaseLinkera."
            product_id = sku_map.get(key)
            return (product_id, None) if product_id is not None else (None, f"Nie znaleziono SKU {record.sku!r}.")
        if mode == "ean" and record.ean:
            key = _norm_ean(record.ean)
            if key in duplicate_eans:
                return None, f"EAN {record.ean!r} jest nieunikalny w katalogu BaseLinkera."
            product_id = ean_map.get(key)
            return (product_id, None) if product_id is not None else (None, f"Nie znaleziono EAN {record.ean!r}.")
    return None, f"Brak wartości dla wybranego sposobu dopasowania: {options.match_by}."


def _text_key(field_name: str, text_fields: dict[str, Any], language: str) -> str:
    localized = f"{field_name}|{language}" if language else field_name
    if localized in text_fields:
        return localized
    if field_name in text_fields:
        return field_name
    return localized


def _features(value: Any) -> tuple[dict[str, str], str | None]:
    if value is None or value == "":
        return {}, None
    parsed = value
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError:
            return {}, "Istniejące features w BaseLinkerze nie są poprawnym JSON-em."
    if not isinstance(parsed, dict):
        return {}, "Istniejące features w BaseLinkerze nie są obiektem."
    return {str(key): str(raw) for key, raw in parsed.items() if str(key).strip() and str(raw).strip()}, None


def _merge_features(existing: dict[str, str], incoming: dict[str, str]) -> dict[str, str]:
    output = dict(existing)
    keys = {_norm_name(key): key for key in output}
    for key, value in incoming.items():
        normalized = _norm_name(key)
        actual = keys.get(normalized, key)
        output[actual] = value
        keys[normalized] = actual
    return output


def _ensure_categories(
    client: BaseLinkerClient,
    requested: list[str],
    mapping: dict[str, int],
    options: SyncOptions,
    limiter: RateLimiter,
    log: LogFn,
) -> None:
    for raw_path in sorted(set(requested), key=lambda value: (_norm_path(value).count(">"), _norm_path(value))):
        if _int(raw_path) is not None:
            continue
        parent_id = 0
        parts: list[str] = []
        for part in (" ".join(item.split()) for item in raw_path.split(">") if item.strip()):
            parts.append(part)
            path = " > ".join(parts)
            normalized = _norm_path(path)
            existing = mapping.get(normalized)
            if existing is not None:
                parent_id = existing
                continue
            response = _api(
                lambda part=part, parent_id=parent_id: client.add_category(
                    {"inventory_id": options.inventory_id, "name": part, "parent_id": parent_id}
                ),
                limiter=limiter,
                options=options,
                log=log,
            )
            category_id = _int(response.get("category_id"))
            if category_id is None:
                raise RuntimeError(f"BaseLinker nie zwrócił ID utworzonej kategorii: {path}")
            mapping[normalized] = category_id
            parent_id = category_id
            log(f"Utworzono kategorię: {path} (ID {category_id}).")


def _ensure_manufacturers(
    client: BaseLinkerClient,
    requested: list[str],
    mapping: dict[str, int],
    options: SyncOptions,
    limiter: RateLimiter,
    log: LogFn,
) -> None:
    for name in sorted(set(requested), key=str.casefold):
        normalized = _norm_name(name)
        if not normalized or normalized in mapping:
            continue
        response = _api(
            lambda name=name: client.add_manufacturer(name),
            limiter=limiter,
            options=options,
            log=log,
        )
        manufacturer_id = _int(response.get("manufacturer_id"))
        if manufacturer_id is None:
            raise RuntimeError(f"BaseLinker nie zwrócił ID utworzonego producenta: {name}")
        mapping[normalized] = manufacturer_id
        log(f"Utworzono producenta: {name} (ID {manufacturer_id}).")


def _plan_row(
    record: ProductRecord,
    product_id: int | None,
    existing: dict[str, Any],
    *,
    options: SyncOptions,
    category_map: dict[str, int],
    category_ids: set[int],
    category_duplicates: set[str],
    manufacturer_map: dict[str, int],
    manufacturer_duplicates: set[str],
    input_dir: Path,
    resolution_error: str | None = None,
) -> AuditRow:
    errors = list(record.errors)
    warnings = list(record.warnings)
    if resolution_error:
        errors.append(resolution_error)
    if product_id is None:
        return AuditRow(
            row_number=record.row_number,
            product_id="",
            sku=record.sku,
            ean=record.ean,
            status="INVALID",
            warnings=" | ".join(warnings),
            error=" | ".join(errors),
        )

    existing_sku = str(existing.get("sku", "") or "")
    existing_ean = str(existing.get("ean", "") or "")
    if record.sku and existing_sku and _norm_sku(record.sku) != _norm_sku(existing_sku):
        errors.append(f"Niezgodność SKU: wejście={record.sku!r}, BaseLinker={existing_sku!r}.")
    if record.ean and existing_ean and _norm_ean(record.ean) != _norm_ean(existing_ean):
        errors.append(f"Niezgodność EAN: wejście={record.ean!r}, BaseLinker={existing_ean!r}.")

    parameters: dict[str, Any] = {"inventory_id": options.inventory_id, "product_id": product_id}
    before: dict[str, Any] = {}
    after: dict[str, Any] = {}
    changed: list[str] = []
    text_fields_raw = existing.get("text_fields", {})
    text_fields = text_fields_raw if isinstance(text_fields_raw, dict) else {}
    new_text_fields: dict[str, Any] = {}

    if "name" in options.fields and record.name:
        if len(record.name) > 200:
            errors.append("Nazwa przekracza limit 200 znaków BaseLinkera.")
        else:
            key = _text_key("name", text_fields, options.language)
            old = str(text_fields.get(key, ""))
            if old != record.name:
                new_text_fields[key] = record.name
                before["name"] = old
                after["name"] = record.name
                changed.append("name")

    if "description" in options.fields and record.description:
        key = _text_key("description", text_fields, options.language)
        old = str(text_fields.get(key, ""))
        if old != record.description:
            new_text_fields[key] = record.description
            before["description"] = old
            after["description"] = record.description
            changed.append("description")

    if "features" in options.fields and record.features:
        key = _text_key("features", text_fields, options.language)
        old_features, feature_error = _features(text_fields.get(key))
        if feature_error and options.merge_features:
            errors.append(feature_error + " Merge został zatrzymany, aby nie usunąć cech.")
        else:
            new_features = _merge_features(old_features, record.features) if options.merge_features else dict(record.features)
            if new_features != old_features:
                new_text_fields[key] = new_features
                before["features"] = old_features
                after["features"] = new_features
                changed.append("features")

    if new_text_fields:
        parameters["text_fields"] = new_text_fields

    if "manufacturer" in options.fields and record.manufacturer_name:
        normalized = _norm_name(record.manufacturer_name)
        if normalized in manufacturer_duplicates:
            errors.append(f"Producent jest nieunikalny w BaseLinkerze: {record.manufacturer_name}")
        else:
            manufacturer_id = manufacturer_map.get(normalized)
            if manufacturer_id is None:
                if options.create_missing_manufacturers and not options.apply:
                    changed.append("manufacturer(create)")
                    after["manufacturer_name"] = record.manufacturer_name
                else:
                    errors.append(f"Brak producenta w BaseLinkerze: {record.manufacturer_name}")
            elif _int(existing.get("manufacturer_id")) != manufacturer_id:
                parameters["manufacturer_id"] = manufacturer_id
                before["manufacturer_id"] = _int(existing.get("manufacturer_id"))
                after["manufacturer_id"] = manufacturer_id
                changed.append("manufacturer")

    if "category" in options.fields and record.category:
        category_id = _int(record.category)
        normalized = _norm_path(record.category)
        if category_id is not None and category_id not in category_ids:
            errors.append(f"category_id {category_id} nie istnieje w wybranym katalogu.")
            category_id = None
        elif category_id is None and normalized in category_duplicates:
            errors.append(f"Kategoria jest nieunikalna; podaj pełną ścieżkę: {record.category}")
        elif category_id is None:
            category_id = category_map.get(normalized)
            if category_id is None:
                if options.create_missing_categories and not options.apply:
                    changed.append("category(create)")
                    after["category"] = record.category
                else:
                    errors.append(f"Brak kategorii w BaseLinkerze: {record.category}")
        if category_id is not None and _int(existing.get("category_id")) != category_id:
            parameters["category_id"] = category_id
            before["category_id"] = _int(existing.get("category_id"))
            after["category_id"] = category_id
            changed.append("category")

    if "images" in options.fields and (record.images_urls or record.images_paths):
        images, image_warnings = build_images_payload(
            record.images_urls,
            record.images_paths,
            input_dir=input_dir,
            allow_http=options.allow_http_images,
        )
        warnings.extend(image_warnings)
        if images:
            parameters["images"] = images
            before["images_count"] = len(existing.get("images", {})) if isinstance(existing.get("images"), dict) else 0
            after["images_count"] = len(images)
            changed.append("images")

    if options.update_identifiers:
        for field_name, value in (("sku", record.sku), ("ean", record.ean)):
            if value and str(existing.get(field_name, "")) != value:
                parameters[field_name] = value
                before[field_name] = str(existing.get(field_name, ""))
                after[field_name] = value
                changed.append(field_name)

    status = "INVALID" if errors else ("READY" if changed else "NO_CHANGES")
    return AuditRow(
        row_number=record.row_number,
        product_id=str(product_id),
        sku=record.sku,
        ean=record.ean,
        status=status,
        changed_fields=",".join(changed),
        warnings=" | ".join(warnings),
        error=" | ".join(errors),
        before_json=json.dumps(before, ensure_ascii=False, separators=(",", ":")),
        after_json=json.dumps(after, ensure_ascii=False, separators=(",", ":")),
        payload=parameters if changed and not errors else None,
    )


def _atomic_write_text(path: Path, text: str, *, encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(handle, "w", encoding=encoding, newline="") as stream:
            stream.write(text)
        os.replace(temp_name, path)
    except Exception:
        Path(temp_name).unlink(missing_ok=True)
        raise


def _write_audit(path: Path, rows: list[AuditRow]) -> None:
    from io import StringIO

    output = StringIO(newline="")
    fieldnames = [
        "row_number",
        "product_id",
        "sku",
        "ean",
        "status",
        "changed_fields",
        "warnings",
        "error",
        "before_json",
        "after_json",
    ]
    writer = csv.DictWriter(output, fieldnames=fieldnames, delimiter=";")
    writer.writeheader()
    for row in rows:
        writer.writerow({key: getattr(row, key) for key in fieldnames})
    _atomic_write_text(path, output.getvalue(), encoding="utf-8-sig")


def default_report_dir() -> Path:
    return Path("reports") / "baselinker_api_sync" / datetime.now().strftime("%Y%m%d_%H%M%S")


def sync_inventory(
    client: BaseLinkerClient,
    input_path: str | Path,
    options: SyncOptions,
    *,
    report_dir: str | Path | None = None,
    log: LogFn | None = None,
) -> SyncResult:
    logger = log or (lambda _: None)
    source_path = Path(input_path)
    destination = Path(report_dir) if report_dir else default_report_dir()
    destination.mkdir(parents=True, exist_ok=True)
    records = load_product_records(source_path)
    limiter = RateLimiter(options.requests_per_minute)

    requires_catalog_index = options.match_by in {"sku", "ean"} or (
        options.match_by == "auto" and any(record.sku or record.ean for record in records)
    )
    if requires_catalog_index:
        logger(f"Wczytano {len(records)} rekordów. Pobieram indeks katalogu BaseLinkera.")
        sku_map, duplicate_skus, ean_map, duplicate_eans, catalog_product_ids = _fetch_product_index(
            client, options, limiter, logger
        )
    else:
        logger(f"Wczytano {len(records)} rekordów. Dopasowanie po product_id nie wymaga pełnego indeksu katalogu.")
        sku_map, duplicate_skus, ean_map, duplicate_eans = {}, set(), {}, set()
        catalog_product_ids = {
            product_id
            for record in records
            if (product_id := _int(record.product_id)) is not None
        }
    resolved: list[tuple[ProductRecord, int | None, str | None]] = []
    for record in records:
        product_id, error = _resolve_product_id(
            record,
            options=options,
            sku_map=sku_map,
            duplicate_skus=duplicate_skus,
            ean_map=ean_map,
            duplicate_eans=duplicate_eans,
            product_ids=catalog_product_ids,
        )
        resolved.append((record, product_id, error))

    details = _fetch_details(
        client,
        [product_id for _, product_id, _ in resolved if product_id is not None],
        options,
        limiter,
        logger,
    )
    resolved = [
        (
            record,
            product_id,
            resolution_error
            or (
                f"Nie udało się pobrać danych product_id {product_id}."
                if product_id is not None and str(product_id) not in details
                else None
            ),
        )
        for record, product_id, resolution_error in resolved
    ]
    _atomic_write_text(
        destination / "backup_before.json",
        json.dumps(details, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    categories = _api(
        lambda: client.get_categories(options.inventory_id), limiter=limiter, options=options, log=logger
    )
    manufacturers = _api(client.get_manufacturers, limiter=limiter, options=options, log=logger)
    _atomic_write_text(
        destination / "catalog_metadata_before.json",
        json.dumps({"categories": categories, "manufacturers": manufacturers}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    category_map, category_duplicates = build_category_maps(categories)
    category_ids = {
        category_id
        for category in categories
        if (category_id := _int(category.get("category_id", category.get("id")))) is not None
    }
    manufacturer_map, manufacturer_duplicates = build_manufacturer_maps(manufacturers)

    eligible_records: list[ProductRecord] = []
    for record, _, resolution_error in resolved:
        invalid = bool(record.errors or resolution_error)
        invalid = invalid or (len(record.name) > 200 and "name" in options.fields)
        if "category" in options.fields and record.category:
            numeric_category = _int(record.category)
            normalized_category = _norm_path(record.category)
            invalid = invalid or (
                numeric_category is not None and numeric_category not in category_ids
            )
            invalid = invalid or (
                numeric_category is None
                and (
                    normalized_category in category_duplicates
                    or (
                        normalized_category not in category_map
                        and not options.create_missing_categories
                    )
                )
            )
        if "manufacturer" in options.fields and record.manufacturer_name:
            normalized_manufacturer = _norm_name(record.manufacturer_name)
            invalid = invalid or normalized_manufacturer in manufacturer_duplicates
            invalid = invalid or (
                normalized_manufacturer not in manufacturer_map
                and not options.create_missing_manufacturers
            )
        if not invalid:
            eligible_records.append(record)

    preliminary_blockers = len(eligible_records) != len(records)
    if options.apply and (not preliminary_blockers or options.allow_partial):
        if options.create_missing_categories:
            _ensure_categories(
                client,
                [record.category for record in eligible_records if record.category],
                category_map,
                options,
                limiter,
                logger,
            )
        if options.create_missing_manufacturers:
            _ensure_manufacturers(
                client,
                [record.manufacturer_name for record in eligible_records if record.manufacturer_name],
                manufacturer_map,
                options,
                limiter,
                logger,
            )

    audit_rows = [
        _plan_row(
            record,
            product_id,
            details.get(str(product_id), {}) if product_id is not None else {},
            options=options,
            category_map=category_map,
            category_ids=category_ids,
            category_duplicates=category_duplicates,
            manufacturer_map=manufacturer_map,
            manufacturer_duplicates=manufacturer_duplicates,
            input_dir=source_path.parent,
            resolution_error=resolution_error,
        )
        for record, product_id, resolution_error in resolved
    ]
    blocked = any(row.status == "INVALID" for row in audit_rows) and not options.allow_partial

    if options.apply and blocked:
        for row in audit_rows:
            if row.status == "READY":
                row.status = "BLOCKED_BATCH"
                row.error = "Nie wysłano: inne rekordy mają błędy. Użyj --allow-partial dopiero po kontroli raportu."
    elif options.apply:
        for row in audit_rows:
            if row.status != "READY" or row.payload is None:
                continue
            try:
                response = _api(
                    lambda payload=row.payload: client.add_product(payload),
                    limiter=limiter,
                    options=options,
                    log=logger,
                )
                warnings = response.get("warnings")
                if warnings:
                    serialized = json.dumps(warnings, ensure_ascii=False, separators=(",", ":"))
                    row.warnings = " | ".join(value for value in (row.warnings, serialized) if value)
                row.status = "SUCCESS"
            except Exception as exc:
                row.status = "ERROR"
                row.error = str(exc)

    _write_audit(destination / "audit.csv", audit_rows)
    counts: dict[str, int] = {}
    for row in audit_rows:
        counts[row.status] = counts.get(row.status, 0) + 1
    summary = {
        "input": str(source_path.resolve()),
        "inventory_id": options.inventory_id,
        "mode": "APPLY" if options.apply else "DRY_RUN",
        "blocked": blocked,
        "row_count": len(records),
        "status_counts": counts,
        "backup": str((destination / "backup_before.json").resolve()),
        "catalog_metadata_backup": str((destination / "catalog_metadata_before.json").resolve()),
        "audit": str((destination / "audit.csv").resolve()),
        "options": {
            key: sorted(value) if isinstance(value, set) else value
            for key, value in asdict(options).items()
            if key != "apply"
        },
    }
    _atomic_write_text(destination / "summary.json", json.dumps(summary, ensure_ascii=False, indent=2))
    logger(f"Raport synchronizacji: {destination}")
    return SyncResult(report_dir=destination, audit_rows=audit_rows, applied=options.apply and not blocked, blocked=blocked)
