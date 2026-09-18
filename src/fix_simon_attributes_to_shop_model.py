"""Poprawia atrybuty filtrowe Kontakt Simon w CSV Baselinkera do modelu sklepu.

Zrodlem prawdy jest ``dictionaries/kontakt_simon_shop_structure.yaml``:
mapa symbol bazowy -> atrybuty rodzenstwa (inne kolory tego samego produktu
w sklepie), a gdy rodzenstwa nie ma - reguly po nazwie. Skrypt rusza tylko
atrybuty z listy ``managed_attributes``; reszta cech (wymiary, GPSR, ETIM)
zostaje nietknieta. Plik wejsciowy nie jest nadpisywany.

Uzycie:
    py src/fix_simon_attributes_to_shop_model.py \
        --input output/Simon_54_55_kaszmir_98_Baselinker.csv \
        --output output/Simon_54_55_kaszmir_98_Baselinker_TYPY_POPRAWIONE.csv \
        --control output/Simon_54_55_kaszmir_98_Baselinker_TYPY_kontrola.xlsx \
        --report reports/simon_54_55_kaszmir_98_typy_report.json
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DICTIONARY = ROOT / "dictionaries" / "kontakt_simon_shop_structure.yaml"

TYP = "Typ produktu"
PODTYP = "Podtyp produktu"
KROTNOSC = "Krotność"
MULTI_SEPARATOR = ", "


@dataclass
class FixResult:
    sku: str
    name: str
    source: str
    before: dict[str, str]
    after: dict[str, str]
    changes: dict[str, tuple[str | None, str | None]] = field(default_factory=dict)
    problems: list[str] = field(default_factory=list)


def load_dictionary(path: Path = DEFAULT_DICTIONARY) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        return yaml.safe_load(stream)


def symbol_base(sku: str, config: dict[str, Any]) -> str:
    settings = config.get("symbol_base", {})
    value = (sku or "").strip()
    suffix = settings.get("strip_supplier_suffix")
    if suffix and value.endswith(suffix):
        value = value[: -len(suffix)]
    regex = settings.get("color_suffix_regex")
    if regex:
        value = re.sub(regex, "", value)
    return value


def _normalize_value(attribute: str, value: str | None, config: dict[str, Any]) -> str | None:
    if value is None:
        return None
    table = (config.get("value_normalization") or {}).get(attribute) or {}
    stripped = value.strip()
    if stripped in table:
        return table[stripped]
    lowered = stripped.lower()
    for source, target in table.items():
        if str(source).lower() == lowered:
            return target
    return stripped


def _resolve_sibling_value(raw: Any, current: str | None) -> str:
    if isinstance(raw, dict) and "niejednoznaczne" in raw:
        options = [str(item) for item in raw["niejednoznaczne"]]
        if current and current in options:
            return current
        return options[0]
    return str(raw)


def _match_name_rule(name: str, config: dict[str, Any]) -> dict[str, Any] | None:
    lowered = (name or "").strip().lower()
    for rule in config.get("name_rules") or []:
        if re.search(rule["match"], lowered):
            return rule
    return None


def _krotnosc_from_name(name: str) -> str | None:
    match = re.search(r"(\d)\s*-\s*krotn", name or "", flags=re.IGNORECASE)
    return f"{match.group(1)}x" if match else None


def fix_features(sku: str, name: str, features: dict[str, str], config: dict[str, Any]) -> FixResult:
    managed: list[str] = list(config.get("managed_attributes") or [])
    before = {key: value for key, value in features.items() if key in managed}
    target: dict[str, str | None] = {}

    entry = (config.get("symbol_map") or {}).get(symbol_base(sku, config))
    sibling_attrs: dict[str, Any] = dict((entry or {}).get("atrybuty") or {})
    source = "brak rodzenstwa"

    if entry:
        source = f"rodzenstwo w sklepie ({entry.get('rodzenstwo', '?')} szt.)"
        for key in managed:
            if key in sibling_attrs:
                target[key] = _resolve_sibling_value(sibling_attrs[key], features.get(key))
            else:
                target[key] = None

    if not entry or TYP not in sibling_attrs:
        rule = _match_name_rule(name, config)
        if rule is not None:
            source = f"{source} + regula nazwy '{rule['match']}'" if entry else f"regula nazwy '{rule['match']}'"
            target[TYP] = rule.get("typ")
            target[PODTYP] = rule.get("podtyp")
            if rule.get("typ") == "Ramki":
                target.setdefault(KROTNOSC, None)
                if not target.get(KROTNOSC):
                    target[KROTNOSC] = _krotnosc_from_name(name) or features.get(KROTNOSC)
        elif not entry:
            target[TYP] = features.get(TYP)
            target[PODTYP] = features.get(PODTYP)
        for key in managed:
            if key in target:
                continue
            target[key] = _normalize_value(key, features.get(key), config)

    if target.get(TYP) != "Ramki":
        for key in config.get("managed_only_for_frames") or []:
            if key in features:
                target[key] = features[key]
            elif key in target and target[key] is not None and key not in sibling_attrs:
                target[key] = None

    after_all: dict[str, str] = {}
    inserted = False
    new_keys = [key for key in managed if target.get(key) and key not in features]
    for key, value in features.items():
        if key in managed:
            resolved = target.get(key)
            if resolved:
                after_all[key] = resolved
            continue
        after_all[key] = value
        if not inserted and key == "Seria":
            for new_key in new_keys:
                after_all[new_key] = target[new_key]  # type: ignore[assignment]
            inserted = True
    if not inserted:
        for new_key in new_keys:
            after_all[new_key] = target[new_key]  # type: ignore[assignment]

    after = {key: value for key, value in after_all.items() if key in managed}
    changes = {
        key: (before.get(key), after.get(key))
        for key in managed
        if before.get(key) != after.get(key)
    }
    result = FixResult(sku=sku, name=name, source=source, before=before, after=after_all, changes=changes)
    result.problems = validate_features(after, config)
    return result


def validate_features(managed_values: dict[str, str], config: dict[str, Any]) -> list[str]:
    problems: list[str] = []
    types: dict[str, Any] = (config.get("filter_model") or {}).get("types") or {}
    typ = managed_values.get(TYP)
    podtyp = managed_values.get(PODTYP)
    if typ is None:
        if podtyp:
            problems.append("Podtyp bez Typu")
        return problems
    if typ not in types:
        problems.append(f"Typ spoza slownika sklepu: {typ}")
        return problems
    allowed = list(types[typ].get("podtypy") or [])
    if typ == "Ramki":
        if podtyp:
            problems.append("Ramki nie maja Podtypu")
        if managed_values.get(KROTNOSC) not in (config["filter_model"]["value_formats"].get(KROTNOSC) or []):
            problems.append("Ramka bez poprawnej Krotnosci")
        return problems
    if not podtyp:
        problems.append(f"Brak Podtypu dla typu {typ}")
        return problems
    for part in podtyp.split(MULTI_SEPARATOR):
        if part not in allowed:
            problems.append(f"Podtyp spoza slownika dla {typ}: {part}")
    return problems


def read_baselinker_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream, delimiter=";")
        rows = list(reader)
        return list(reader.fieldnames or []), rows


def write_baselinker_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames, delimiter=";", lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(rows)


def _simplify_attribute_name(name: str) -> str:
    lowered = re.sub(r"\[.*?\]", "", name).strip().lower()
    return re.sub(r"[^a-z0-9ąćęłńóśźż]+", "", lowered)


def allowed_attributes(config: dict[str, Any]) -> set[str]:
    return set(config.get("woo_attributes") or []) | set(config.get("approved_extra_attributes") or [])


def suggest_attribute_mapping(name: str, allowed: set[str]) -> str | None:
    """Podpowiada istniejacy atrybut Woo o tej samej nazwie bez jednostki/znakow."""
    key = _simplify_attribute_name(name)
    if not key:
        return None
    for candidate in sorted(allowed):
        if _simplify_attribute_name(candidate) == key:
            return candidate
    return None


def strip_unknown_attributes(
    result: FixResult, allowed: set[str], unknown: dict[str, dict[str, Any]]
) -> None:
    """Usuwa z cech klucze spoza Woo i zapisuje je do slownika ``unknown`` (raport)."""
    kept: dict[str, str] = {}
    for key, value in result.after.items():
        if key in allowed:
            kept[key] = value
            continue
        entry = unknown.setdefault(key, {"produktow": 0, "przyklady": [], "sku": []})
        entry["produktow"] += 1
        if value not in entry["przyklady"] and len(entry["przyklady"]) < 3:
            entry["przyklady"].append(value)
        entry["sku"].append(result.sku)
        result.changes[key] = (value, None)
    result.after = kept


def process_rows(
    rows: list[dict[str, str]], config: dict[str, Any], keep_unknown: bool = False
) -> tuple[list[FixResult], dict[str, dict[str, Any]]]:
    results: list[FixResult] = []
    unknown: dict[str, dict[str, Any]] = {}
    allowed = allowed_attributes(config)
    for row in rows:
        features = json.loads(row.get("features") or "{}")
        result = fix_features(row.get("sku", ""), row.get("name", ""), features, config)
        if not keep_unknown:
            strip_unknown_attributes(result, allowed, unknown)
        row["features"] = json.dumps(result.after, ensure_ascii=False)
        results.append(result)
    for key, entry in unknown.items():
        entry["sugerowany_atrybut_woo"] = suggest_attribute_mapping(key, allowed)
    return results, unknown


def write_control_workbook(
    path: Path, results: list[FixResult], managed: list[str], unknown: dict[str, dict[str, Any]] | None = None
) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Kontrola"
    if unknown:
        approval = workbook.create_sheet("Atrybuty do zatwierdzenia")
        approval.append(["Decyzja (TAK/NIE)", "Atrybut z katalogu", "Produktow", "Sugerowany atrybut Woo", "Przyklady wartosci", "SKU"])
        for key, entry in sorted(unknown.items(), key=lambda item: -item[1]["produktow"]):
            approval.append(["", key, entry["produktow"], entry.get("sugerowany_atrybut_woo") or "", " | ".join(entry["przyklady"]), ", ".join(entry["sku"][:12])])
        for cell in approval[1]:
            cell.font = Font(bold=True)
        for column, width in zip("ABCDEF", [16, 40, 10, 30, 60, 60]):
            approval.column_dimensions[column].width = width
        approval.freeze_panes = "B2"
    header = ["SKU", "Nazwa", "Zrodlo decyzji", "Problemy"]
    for key in managed:
        header += [f"{key} PRZED", f"{key} PO"]
    sheet.append(header)
    changed_fill = PatternFill("solid", fgColor="FFF2CC")
    problem_fill = PatternFill("solid", fgColor="F8CBAD")
    for result in results:
        row = [result.sku, result.name, result.source, "; ".join(result.problems)]
        for key in managed:
            row += [result.before.get(key, ""), result.after.get(key, "")]
        sheet.append(row)
        current = sheet.max_row
        if result.problems:
            sheet.cell(current, 4).fill = problem_fill
        for index, key in enumerate(managed):
            if key in result.changes:
                sheet.cell(current, 6 + index * 2).fill = changed_fill
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "C2"
    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 60
    sheet.column_dimensions["C"].width = 34
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def build_report(
    results: list[FixResult], input_path: Path, output_path: Path, unknown: dict[str, dict[str, Any]] | None = None
) -> dict[str, Any]:
    change_counts: dict[str, int] = {}
    for result in results:
        for key in result.changes:
            change_counts[key] = change_counts.get(key, 0) + 1
    return {
        "input": str(input_path),
        "output": str(output_path),
        "rows": len(results),
        "rows_changed": sum(1 for result in results if result.changes),
        "rows_with_problems": sum(1 for result in results if result.problems),
        "atrybuty_do_zatwierdzenia": {
            key: {k: v for k, v in entry.items() if k != "sku"}
            for key, entry in sorted((unknown or {}).items(), key=lambda item: -item[1]["produktow"])
        },
        "changes_by_attribute": dict(sorted(change_counts.items(), key=lambda item: -item[1])),
        "sources": {
            source: sum(1 for result in results if result.source == source)
            for source in sorted({result.source for result in results})
        },
        "problems": [
            {"sku": result.sku, "name": result.name, "problems": result.problems}
            for result in results
            if result.problems
        ],
        "details": [
            {
                "sku": result.sku,
                "name": result.name,
                "source": result.source,
                "changes": {key: {"przed": old, "po": new} for key, (old, new) in result.changes.items()},
            }
            for result in results
            if result.changes
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", required=True, type=Path, help="CSV Baselinkera (product_id;name;sku;...;features).")
    parser.add_argument("--output", required=True, type=Path, help="Poprawiony CSV (nowy plik).")
    parser.add_argument("--control", type=Path, help="Skoroszyt kontrolny przed/po.")
    parser.add_argument("--report", type=Path, help="Raport JSON.")
    parser.add_argument("--dictionary", type=Path, default=DEFAULT_DICTIONARY)
    parser.add_argument(
        "--keep-unknown-attributes",
        action="store_true",
        help="Nie usuwa cech spoza listy atrybutow Woo (tylko do podgladu; domyslnie sa usuwane i raportowane).",
    )
    args = parser.parse_args(argv)

    if args.output.resolve() == args.input.resolve():
        parser.error("Plik wyjsciowy nie moze nadpisywac wejsciowego.")

    config = load_dictionary(args.dictionary)
    fieldnames, rows = read_baselinker_csv(args.input)
    results, unknown = process_rows(rows, config, keep_unknown=args.keep_unknown_attributes)
    write_baselinker_csv(args.output, fieldnames, rows)
    report = build_report(results, args.input, args.output, unknown)
    if args.control:
        write_control_workbook(args.control, results, list(config.get("managed_attributes") or []), unknown)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        f"Wiersze: {report['rows']}, zmienione: {report['rows_changed']}, "
        f"z problemami: {report['rows_with_problems']}, atrybutow spoza Woo do zatwierdzenia: {len(unknown)} -> {args.output}"
    )
    return 1 if report["rows_with_problems"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
