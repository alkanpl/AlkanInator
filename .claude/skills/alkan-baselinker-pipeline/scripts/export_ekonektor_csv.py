from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path


FIELDS = (
    "TOWARIDX",
    "ACTIVE",
    "PRODUCT_CODE",
    "NAME",
    "TAX_ID",
    "UNIT",
    "PRODUCER",
    "CATEGORY",
)

TRANSLITERATION = str.maketrans(
    {
        "\u00b2": "2",
        "\u00b3": "3",
        "\u2264": "<=",
        "\u2265": ">=",
        "\u00d7": "x",
        "\u2212": "-",
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
    }
)


def normalize(value: object) -> str:
    text = str(value or "").replace("\u00a0", " ").translate(TRANSLITERATION)
    return re.sub(r"\s+", " ", text).strip()


def ensure_cp1250(value: str, field: str, row_number: int) -> None:
    try:
        value.encode("cp1250", errors="strict")
    except UnicodeEncodeError as exc:
        character = value[exc.start : exc.end]
        codepoints = " ".join(f"U+{ord(char):04X}" for char in character)
        raise ValueError(
            f"Row {row_number}, field {field}: unsupported cp1250 character {codepoints}"
        ) from exc


def build_rows(
    input_path: Path,
    input_encoding: str,
    code_column: str,
    name_column: str,
    unit: str,
    producer: str,
    category: str,
) -> list[dict[str, str]]:
    with input_path.open("r", encoding=input_encoding, newline="") as stream:
        reader = csv.DictReader(stream, delimiter=";")
        source_fields = reader.fieldnames or []
        missing = [field for field in (code_column, name_column) if field not in source_fields]
        if missing:
            raise ValueError(f"Missing source columns: {missing}; found: {source_fields}")

        rows: list[dict[str, str]] = []
        for source_row_number, source in enumerate(reader, start=2):
            if not any(normalize(value) for value in source.values()):
                continue
            code = normalize(source.get(code_column))
            name = normalize(source.get(name_column))
            target = {
                "TOWARIDX": code,
                "ACTIVE": "1",
                "PRODUCT_CODE": code,
                "NAME": name,
                "TAX_ID": "1",
                "UNIT": normalize(unit),
                "PRODUCER": normalize(producer),
                "CATEGORY": normalize(category),
            }
            empty = [field for field in FIELDS if not target[field]]
            if empty:
                raise ValueError(f"Source row {source_row_number}: empty target fields {empty}")
            for field, value in target.items():
                ensure_cp1250(value, field, source_row_number)
            rows.append(target)
    return rows


def write_and_validate(output_path: Path, rows: list[dict[str, str]]) -> dict:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="cp1250", errors="strict", newline="") as stream:
        writer = csv.DictWriter(
            stream,
            fieldnames=FIELDS,
            delimiter=";",
            lineterminator="\r\n",
            quoting=csv.QUOTE_MINIMAL,
        )
        writer.writeheader()
        writer.writerows(rows)

    raw = output_path.read_bytes()
    if raw.startswith((b"\xef\xbb\xbf", b"\xff\xfe", b"\xfe\xff")):
        raise AssertionError("ANSI output must not contain a BOM")
    decoded = raw.decode("cp1250", errors="strict")
    parsed = list(csv.DictReader(decoded.splitlines(), delimiter=";"))
    parsed_fields = tuple(parsed[0].keys()) if parsed else tuple()
    codes = [row["PRODUCT_CODE"] for row in parsed]
    validation = {
        "headers": list(parsed_fields),
        "rows": len(parsed),
        "all_fields_nonempty": all(all(row[field] for field in FIELDS) for row in parsed),
        "active_all_1": all(row["ACTIVE"] == "1" for row in parsed),
        "tax_id_all_1": all(row["TAX_ID"] == "1" for row in parsed),
        "codes_match": all(row["TOWARIDX"] == row["PRODUCT_CODE"] for row in parsed),
        "codes_unique": len(codes) == len(set(codes)),
        "cp1250_without_bom": True,
    }
    validation["passed"] = all(
        (
            parsed_fields == FIELDS,
            validation["rows"] == len(rows),
            validation["all_fields_nonempty"],
            validation["active_all_1"],
            validation["tax_id_all_1"],
            validation["codes_match"],
            validation["codes_unique"],
        )
    )
    if not validation["passed"]:
        raise AssertionError(json.dumps(validation, ensure_ascii=False, indent=2))
    return validation


def main() -> None:
    parser = argparse.ArgumentParser(description="Export an eKonektor mass-product ANSI CSV")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--input-encoding", default="utf-8-sig")
    parser.add_argument("--code-column", default="sku")
    parser.add_argument("--name-column", default="name")
    parser.add_argument("--unit", default="szt")
    parser.add_argument("--producer", required=True)
    parser.add_argument("--category", default="Bez kategorii")
    args = parser.parse_args()

    rows = build_rows(
        input_path=args.input,
        input_encoding=args.input_encoding,
        code_column=args.code_column,
        name_column=args.name_column,
        unit=args.unit,
        producer=args.producer,
        category=args.category,
    )
    validation = write_and_validate(args.output, rows)
    print(
        json.dumps(
            {
                "input": str(args.input),
                "output": str(args.output),
                "unit": args.unit,
                "producer": args.producer,
                "category": args.category,
                "validation": validation,
            },
            ensure_ascii=True,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
