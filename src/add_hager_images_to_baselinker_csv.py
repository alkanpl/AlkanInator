from __future__ import annotations

import argparse
import csv
import hashlib
import json
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "output" / "Hager_Baselinker_poza_Woo_57_produkty_2026-08-07.csv"
DEFAULT_XML = ROOT / "input" / "BMEcat_PL 17.04.2026.xml"
DEFAULT_OUTPUT = (
    ROOT / "output" / "Hager_Baselinker_poza_Woo_57_produkty_ZDJECIA_2026-08-11.csv"
)
DEFAULT_REPORT = (
    ROOT / "reports" / "Hager_Baselinker_poza_Woo_57_produkty_ZDJECIA_2026-08-11.json"
)
EXPECTED_FIELDS = [
    "product_id",
    "name",
    "sku",
    "manufacturer_name",
    "category",
    "features",
]
IMAGE_CODES = ("MD01", "MD25")


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def supplier_code(sku: str) -> str:
    value = sku.strip()
    return value[:-4] if value.upper().endswith("/HAG") else value


def read_source(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream, delimiter=";")
        fields = reader.fieldnames or []
        rows = list(reader)
    if fields != EXPECTED_FIELDS:
        raise ValueError(f"Unexpected source fields: {fields}")
    if len(rows) != 57:
        raise ValueError(f"Expected 57 products, got {len(rows)}")
    skus = [row["sku"] for row in rows]
    if len(set(skus)) != len(skus):
        raise ValueError("Duplicate SKU values in source CSV")
    for row in rows:
        json.loads(row["features"])
    return fields, rows


def extract_images(
    xml_path: Path, needed_codes: set[str]
) -> tuple[dict[str, list[str]], dict[str, str]]:
    direct: dict[str, list[str]] = {}
    predecessor_to_successor: dict[str, str] = {}
    successor_images: dict[str, list[str]] = {}

    for _event, element in ET.iterparse(xml_path, events=("end",)):
        if local_name(element.tag) != "PRODUCT":
            continue

        product_code = next(
            (
                (node.text or "").strip()
                for node in element.iter()
                if local_name(node.tag) == "SUPPLIER_PID"
            ),
            "",
        )
        predecessors = {
            (node.text or "").strip()
            for node in element.iter()
            if local_name(node.tag) == "UDX.EDXF.PREDECESSOR_PID"
            and (node.text or "").strip() in needed_codes
        }

        image_groups = {code: [] for code in IMAGE_CODES}
        if product_code in needed_codes or predecessors:
            for mime in element.iter():
                if local_name(mime.tag) != "UDX.EDXF.MIME":
                    continue
                values = {
                    local_name(child.tag): (child.text or "").strip()
                    for child in mime
                }
                mime_code = values.get("UDX.EDXF.MIME_CODE", "")
                source = values.get("UDX.EDXF.MIME_SOURCE", "")
                if (
                    mime_code in image_groups
                    and source
                    and source not in image_groups[mime_code]
                ):
                    image_groups[mime_code].append(source)

        ordered_images = image_groups["MD01"] + image_groups["MD25"]
        if product_code in needed_codes:
            if product_code in direct:
                raise ValueError(f"Duplicate SUPPLIER_PID in XML: {product_code}")
            direct[product_code] = ordered_images
        if predecessors and ordered_images:
            successor_images[product_code] = ordered_images
            for predecessor in predecessors:
                if predecessor in predecessor_to_successor:
                    raise ValueError(f"Multiple successors for predecessor: {predecessor}")
                predecessor_to_successor[predecessor] = product_code

        element.clear()

    fallback_used: dict[str, str] = {}
    for predecessor, successor in predecessor_to_successor.items():
        if not direct.get(predecessor):
            direct[predecessor] = successor_images[successor]
            fallback_used[predecessor] = successor
    return direct, fallback_used


def build_csv(
    input_path: Path, xml_path: Path, output_path: Path, report_path: Path
) -> dict:
    fields, rows = read_source(input_path)
    codes = {supplier_code(row["sku"]) for row in rows}
    images_by_code, predecessor_to_successor = extract_images(xml_path, codes)

    output_rows: list[dict[str, str]] = []
    row_report: list[dict[str, object]] = []
    for source_row in rows:
        code = supplier_code(source_row["sku"])
        urls = images_by_code.get(code, [])
        successor = predecessor_to_successor.get(code)
        if successor:
            status = "FALLBACK_SUCCESSOR"
        elif urls:
            status = "DIRECT"
        else:
            status = "MISSING"

        output_row = dict(source_row)
        output_row["images_urls"] = "|".join(urls)
        output_rows.append(output_row)
        row_report.append(
            {
                "sku": source_row["sku"],
                "supplier_pid": code,
                "status": status,
                "successor_pid": successor,
                "image_count": len(urls),
                "images_urls": urls,
            }
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(
            stream,
            fieldnames=fields + ["images_urls"],
            delimiter=";",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(output_rows)

    with output_path.open("r", encoding="utf-8-sig", newline="") as stream:
        validation_reader = csv.DictReader(stream, delimiter=";")
        validation_fields = validation_reader.fieldnames or []
        validation_rows = list(validation_reader)
    source_unchanged = all(
        all(before[field] == after[field] for field in fields)
        for before, after in zip(rows, validation_rows, strict=True)
    )
    valid_features = all(
        isinstance(json.loads(row["features"]), dict) for row in validation_rows
    )
    validation = {
        "row_count": len(validation_rows),
        "fieldnames": validation_fields,
        "source_fields_unchanged": source_unchanged,
        "unique_sku_count": len({row["sku"] for row in validation_rows}),
        "valid_features_json_count": sum(
            isinstance(json.loads(row["features"]), dict) for row in validation_rows
        ),
        "utf8_bom": output_path.read_bytes().startswith(b"\xef\xbb\xbf"),
    }
    validation["passed"] = all(
        (
            validation["row_count"] == 57,
            validation_fields == fields + ["images_urls"],
            source_unchanged,
            validation["unique_sku_count"] == 57,
            valid_features,
            validation["utf8_bom"],
        )
    )
    if not validation["passed"]:
        raise AssertionError(json.dumps(validation, ensure_ascii=False, indent=2))

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "input_file": str(input_path),
        "input_sha256": sha256(input_path),
        "xml_file": str(xml_path),
        "xml_sha256": sha256(xml_path),
        "output_file": str(output_path),
        "output_sha256": sha256(output_path),
        "product_count": len(rows),
        "direct_image_matches": sum(r["status"] == "DIRECT" for r in row_report),
        "successor_fallback_matches": sum(
            r["status"] == "FALLBACK_SUCCESSOR" for r in row_report
        ),
        "products_without_images": sum(r["status"] == "MISSING" for r in row_report),
        "total_image_urls": sum(int(r["image_count"]) for r in row_report),
        "rows": row_report,
        "validation": validation,
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return report


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Add Hager BMEcat MD01/MD25 image URLs to a BaseLinker CSV"
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--xml", type=Path, default=DEFAULT_XML)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()
    report = build_csv(args.input, args.xml, args.output, args.report)
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
