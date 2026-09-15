from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

from utils import compact_spaces, ensure_dir, normalize_header


DEFAULT_INPUT = "archived_input_files/input_2026-08-05/Poprawione atrybuty.xlsx"
DEFAULT_CATALOG_KNOWLEDGE = "dictionaries/woocommerce_catalog_knowledge.yaml"
DEFAULT_OUTPUT = "dictionaries/lighting_attribute_knowledge.yaml"
DEFAULT_REPORTS_DIR = "reports/attribute_knowledge_import"

ROLE_HEADERS = {
    "filters": "atrybuty do filtrowania",
    "required_description": "atrybuty obowiazkowe do opisu",
    "optional_description": "atrybuty opcjonalne do opisu",
}

# Prefer the established Woo attribute names (shop uses unit-less dimension
# names like "Długość"; units live in the values, e.g. "600mm").
PREFERRED_CATALOG_ALIASES = {
    "barwa": "Barwa światła",
    "dane produceta": "Dane producenta",
    "dlugosc": "Długość",
    "dlugosc mm": "Długość",
    "ilosc zrodel swiatla": "Liczba źródeł światła",
    "ilosc zrodel swiatla w komplecie": "Liczba źródeł światła",
    "kat swiecenia": "Kąt świecenia [°]",
    "klasa efektywnosci energetycznej": "Klasa energetyczna",
    "liczba zrodel swiatla": "Liczba źródeł światła",
    "moc": "Moc [W]",
    "napiecie": "Napięcie [V]",
    "napiecie znamionowe": "Napięcie [V]",
    "skutecznosc swietlna": "Skuteczność świetlna [lm/W]",
    "srednica": "Średnica",
    "srednica mm": "Średnica",
    "szerokosc": "Szerokość",
    "szerokosc mm": "Szerokość",
    "stopien ochrony": "Stopień ochrony [IP]",
    "stopien ochrony ip": "Stopień ochrony [IP]",
    "strumien swietlny": "Strumień świetlny [lm]",
    "temperatura": "Temperatura barwowa [K]",
    "temperatura barwowa": "Temperatura barwowa [K]",
    "trwalosc": "Trwałość [h]",
    "trwalosc h": "Trwałość [h]",
    "typ latarki": "Typ produktu",
    "wydajnosc swietlna": "Skuteczność świetlna [lm/W]",
    "wymiar": "Wymiary [mm]",
    "wymiary": "Wymiary [mm]",
    "wyskosc": "Wysokość",
    "wysokosc": "Wysokość",
    "wysokosc mm": "Wysokość",
    "zrodlo swiatla w komplecie tak nie": "Źródło światła w komplecie",
}

SEMANTIC_CATALOG_ALIASES = {
    "bateria": "Rodzaj baterii",
    # W oswietleniu krotnosc trzonkow to liczba zrodel swiatla, nie "Liczba gniazd"
    # (ta w sklepie dotyczy gniazdek elektrycznych).
    "liczba gniazd": "Liczba źródeł światła",
    "ilosc gniazd": "Liczba źródeł światła",
    "maksymalna moc zrodla swiatla": "Maksymalna moc źródła światła",
    "mozliwosc laczenia przelotowego opraw": "Możliwość łączenia",
    "optyka typ swiecenia": "Optyka",
}

NEW_ATTRIBUTE_ALIASES = {
    "czas podtrzymania czas pracy awaryjnej": "Czas pracy awaryjnej",
    "czas swiecenia": "Czas świecenia",
    "kierunek swiecenia": "Kierunek świecenia",
    "klasa odpornosci ik": "Stopień odporności [IK]",
    "mozliwosc pracy ze sciemniaczem": "Współpraca ze ściemniaczem",
    "mozliwosc wspolpracy z czujnikiem ruchu": "Współpraca z czujnikiem ruchu",
    "mozliwosc wspolpracy ze sciemniaczem": "Współpraca ze ściemniaczem",
    "stopien ik": "Stopień odporności [IK]",
    "test": "Tryb testu awaryjnego",
    "ugr": "Wskaźnik olśnienia [UGR]",
    "waga": "Waga",
    "zasieg": "Zasięg",
    "zrodlo swiatla": "Źródło światła",
    "zrodlo swiatla w komplecie": "Źródło światła w komplecie",
}

SYSTEM_FIELD_ALIASES = {
    "dostepnosc": "Dostępność",
    "dostepnosc w magazynie": "Dostępność",
}

# Reczne uzupelnienia macierzy (decyzje sklepu spoza pliku "Poprawione atrybuty"):
# oprawy hermetyczne na swietlowki maja zachowywac trzonek i liczbe zrodel swiatla.
CATEGORY_EXTRA_ATTRIBUTES = {
    "oprawy hermetyczne": {
        "optional_description": ["Trzonek", "Liczba źródeł światła"],
    },
}

# Atrybuty zdegradowane z wymaganych do opcjonalnych (decyzja sklepu):
# UGR producent podaje tylko dla czesci rodzin, wiec nie moze byc obowiazkowy.
REQUIRED_DEMOTED_TO_OPTIONAL = {"Wskaźnik olśnienia [UGR]"}


def apply_category_extra_attributes(sections: list[dict[str, Any]]) -> None:
    demoted = {normalize_header(name) for name in REQUIRED_DEMOTED_TO_OPTIONAL}
    for section in sections:
        moved = [name for name in section.get("required_description", []) if normalize_header(name) in demoted]
        if moved:
            section["required_description"] = [
                name for name in section["required_description"] if normalize_header(name) not in demoted
            ]
            section["optional_description"] = unique([*section.get("optional_description", []), *moved])
        extras = CATEGORY_EXTRA_ATTRIBUTES.get(normalize_header(str(section.get("category", ""))))
        if not extras:
            continue
        for role, names in extras.items():
            section[role] = unique([*section.get(role, []), *names])


def strip_parenthetical_notes(value: str) -> str:
    text = compact_spaces(value)
    previous = None
    while text != previous:
        previous = text
        text = re.sub(r"\s*\([^()]*\)\s*", " ", text)
        text = compact_spaces(text)
    return text


def split_outside_parentheses(value: str) -> list[str]:
    text = compact_spaces(value)
    if not text:
        return []
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    for char in text:
        if char == "(":
            depth += 1
        elif char == ")" and depth:
            depth -= 1
        if char == "," and depth == 0:
            part = compact_spaces("".join(current))
            if part:
                parts.append(part)
            current = []
            continue
        current.append(char)
    last = compact_spaces("".join(current))
    if last:
        parts.append(last)
    return parts


def unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = normalize_header(value)
        if value and key not in seen:
            seen.add(key)
            result.append(value)
    return result


def is_header_cell(value: Any) -> bool:
    normalized = normalize_header(str(value or ""))
    return normalized in ROLE_HEADERS.values() or "atrybuty do tytulu" in normalized or "kolejnosc atrybutow" in normalized


def first_nonempty(values: list[Any]) -> str:
    for value in values:
        text = compact_spaces(str(value or ""))
        if text:
            return text
    return ""


def find_family(df: pd.DataFrame) -> str:
    for _, row in df.iterrows():
        value = first_nonempty(row.tolist())
        if value and not is_header_cell(value):
            return value
    return ""


def find_category_with_row(df: pd.DataFrame, header_row: int, filter_column: int) -> tuple[str, int]:
    for row_index in range(header_row - 1, -1, -1):
        row = df.iloc[row_index].tolist()
        preferred = compact_spaces(str(row[filter_column] or "")) if filter_column < len(row) else ""
        candidate = preferred or first_nonempty(row)
        if candidate and not is_header_cell(candidate) and not normalize_header(candidate).startswith("do wszystkich"):
            return candidate, row_index
    return "", -1


def role_columns(
    header_values: list[Any],
    previous_values: list[Any],
    df: pd.DataFrame,
    start_row: int,
    end_row: int,
) -> tuple[dict[str, int], str]:
    columns: dict[str, int] = {}
    order_context = ""
    combined_headers = []
    for column, value in enumerate(header_values):
        current = str(value or "")
        previous = str(previous_values[column] or "") if column < len(previous_values) else ""
        combined_headers.append(current if is_header_cell(current) else previous)
    for column, value in enumerate(combined_headers):
        normalized = normalize_header(str(value or ""))
        for role, expected in ROLE_HEADERS.items():
            if normalized == expected:
                columns[role] = column
        if "atrybuty do tytulu" in normalized:
            columns["ordered_attributes"] = column
            order_context = "title"
        elif "kolejnosc atrybutow" in normalized and "opisie" in normalized:
            columns["ordered_attributes"] = column
            order_context = "description"

    if "ordered_attributes" not in columns and "optional_description" in columns:
        candidate = columns["optional_description"] + 1
        if candidate < df.shape[1]:
            values = [
                compact_spaces(str(df.iat[row_index, candidate] or ""))
                for row_index in range(start_row, end_row)
            ]
            if any(values):
                columns["ordered_attributes"] = candidate
                order_context = "unspecified"
    return columns, order_context


def is_instruction(value: str) -> bool:
    normalized = normalize_header(value)
    return normalized.startswith("zalezne od typu produktu")


def parse_sheet(df: pd.DataFrame, sheet_name: str) -> list[dict[str, Any]]:
    family = find_family(df)
    sections: list[dict[str, Any]] = []
    header_rows = [
        int(row_index)
        for row_index, row in df.iterrows()
        if ROLE_HEADERS["filters"] in [normalize_header(str(value or "")) for value in row.tolist()]
    ]
    for header_position, header_row in enumerate(header_rows):
        row = df.iloc[header_row]
        normalized_cells = [normalize_header(str(value or "")) for value in row.tolist()]
        filter_column = normalized_cells.index(ROLE_HEADERS["filters"])
        category, _ = find_category_with_row(df, header_row, filter_column)
        previous_values = df.iloc[header_row - 1].tolist() if header_row else []
        split_header = any(is_header_cell(value) for value in previous_values)
        start_row = header_row if split_header else header_row + 1
        end_row = len(df)
        if header_position + 1 < len(header_rows):
            next_header_row = header_rows[header_position + 1]
            _, next_category_row = find_category_with_row(df, next_header_row, filter_column)
            end_row = next_category_row if next_category_row > header_row else next_header_row
        columns, order_context = role_columns(row.tolist(), previous_values, df, start_row, end_row)
        section: dict[str, Any] = {
            "family": family,
            "category": category,
            "sheet": sheet_name,
            "source_row": int(header_row + 1),
            "order_context": order_context,
            "filters": [],
            "required_description": [],
            "optional_description": [],
            "ordered_attributes": [],
            "notes": [],
        }
        role_column_indexes = set(columns.values())
        for row_index in range(start_row, end_row):
            for role, column in columns.items():
                raw_value = compact_spaces(str(df.iat[row_index, column] or ""))
                if not raw_value or is_header_cell(raw_value):
                    continue
                if is_instruction(raw_value):
                    section["notes"].append(raw_value)
                    continue
                section[role].extend(split_outside_parentheses(raw_value))
            for column in range(df.shape[1]):
                if column in role_column_indexes:
                    continue
                note = compact_spaces(str(df.iat[row_index, column] or ""))
                if note:
                    section["notes"].append(note)
        for role in ("filters", "required_description", "optional_description", "ordered_attributes", "notes"):
            section[role] = unique(section[role])
        sections.append(section)
    return sections


def load_catalog_attributes(path: Path) -> tuple[dict[str, str], set[str]]:
    knowledge = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    best_by_normalized: dict[str, tuple[int, str]] = {}
    canonical_names: set[str] = set()
    for row in knowledge.get("top_attributes") or []:
        name = compact_spaces(str(row.get("attribute", "")))
        if not name:
            continue
        canonical_names.add(name)
        normalized = normalize_header(name)
        count = int(row.get("products_count") or 0)
        current = best_by_normalized.get(normalized)
        if current is None or count > current[0]:
            best_by_normalized[normalized] = (count, name)
    return {key: value[1] for key, value in best_by_normalized.items()}, canonical_names


def normalized_candidates(raw_name: str) -> list[str]:
    raw = normalize_header(raw_name)
    stripped = normalize_header(strip_parenthetical_notes(raw_name))
    candidates = [raw, stripped]
    if raw.startswith("zrodlo swiatla"):
        candidates.append("zrodlo swiatla")
    return unique(candidates)


def canonicalize_attribute(
    raw_name: str,
    catalog_by_normalized: dict[str, str],
    catalog_names: set[str],
) -> dict[str, Any]:
    candidates = normalized_candidates(raw_name)
    for candidate in candidates:
        if candidate in SYSTEM_FIELD_ALIASES:
            return mapping_row(raw_name, SYSTEM_FIELD_ALIASES[candidate], "SYSTEM_FIELD", 1.0, "OK")
    for candidate in candidates:
        if candidate in PREFERRED_CATALOG_ALIASES:
            canonical = PREFERRED_CATALOG_ALIASES[candidate]
            return mapping_row(raw_name, canonical, "CATALOG_ALIAS", 0.99, "OK")
    for candidate in candidates:
        if candidate in SEMANTIC_CATALOG_ALIASES:
            canonical = SEMANTIC_CATALOG_ALIASES[candidate]
            return mapping_row(raw_name, canonical, "CATALOG_SEMANTIC_ALIAS", 0.96, "OK")
    for candidate in candidates:
        if candidate in catalog_by_normalized:
            canonical = catalog_by_normalized[candidate]
            match_type = "CATALOG_EXACT" if normalize_header(raw_name) == candidate else "CATALOG_NORMALIZED"
            return mapping_row(raw_name, canonical, match_type, 1.0 if match_type == "CATALOG_EXACT" else 0.98, "OK")
    for candidate in candidates:
        if candidate in NEW_ATTRIBUTE_ALIASES:
            canonical = NEW_ATTRIBUTE_ALIASES[candidate]
            return mapping_row(raw_name, canonical, "NEW_ATTRIBUTE", 0.95, "OK")

    canonical = strip_parenthetical_notes(raw_name) or compact_spaces(raw_name)
    status = "OK" if canonical in catalog_names else "DO_SPRAWDZENIA"
    return mapping_row(raw_name, canonical, "UNRESOLVED", 0.6, status)


def mapping_row(raw_name: str, canonical_name: str, match_type: str, confidence: float, status: str) -> dict[str, Any]:
    return {
        "source_name": compact_spaces(raw_name),
        "canonical_name": canonical_name,
        "match_type": match_type,
        "source": DEFAULT_INPUT,
        "confidence": confidence,
        "status": status,
    }


def canonicalize_sections(
    sections: list[dict[str, Any]],
    catalog_by_normalized: dict[str, str],
    catalog_names: set[str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    mappings: dict[str, dict[str, Any]] = {}
    roles = ("filters", "required_description", "optional_description", "ordered_attributes")
    result: list[dict[str, Any]] = []
    for section in sections:
        canonical_section = dict(section)
        for role in roles:
            canonical_values: list[str] = []
            for raw_name in section[role]:
                key = normalize_header(raw_name)
                mapping = mappings.setdefault(
                    key,
                    canonicalize_attribute(raw_name, catalog_by_normalized, catalog_names),
                )
                canonical_values.append(mapping["canonical_name"])
            canonical_section[role] = unique(canonical_values)
        result.append(canonical_section)
    return result, sorted(mappings.values(), key=lambda row: normalize_header(row["source_name"]))


def build_attribute_usage(sections: list[dict[str, Any]]) -> dict[str, list[str]]:
    usage: dict[str, set[str]] = defaultdict(set)
    for section in sections:
        category = section["category"]
        for role in ("filters", "required_description", "optional_description", "ordered_attributes"):
            for attribute in section[role]:
                usage[attribute].add(category)
    return {attribute: sorted(categories) for attribute, categories in usage.items()}


def build_knowledge(
    input_path: Path,
    catalog_path: Path,
    sections: list[dict[str, Any]],
    mappings: list[dict[str, Any]],
    catalog_names: set[str],
) -> dict[str, Any]:
    usage = build_attribute_usage(sections)
    new_names = sorted(
        {
            row["canonical_name"]
            for row in mappings
            if row["match_type"] == "NEW_ATTRIBUTE" or (
                row["match_type"] == "UNRESOLVED" and row["canonical_name"] not in catalog_names
            )
        },
        key=normalize_header,
    )
    system_fields = sorted(
        {row["canonical_name"] for row in mappings if row["match_type"] == "SYSTEM_FIELD"},
        key=normalize_header,
    )
    unresolved = [row for row in mappings if row["status"] == "DO_SPRAWDZENIA"]
    return {
        "version": 1,
        "source": {
            "file": input_path.as_posix(),
            "sha256": hashlib.sha256(input_path.read_bytes()).hexdigest(),
            "imported_at": date.today().isoformat(),
            "catalog_knowledge": catalog_path.as_posix(),
        },
        "summary": {
            "categories": len(sections),
            "source_attribute_names": len(mappings),
            "catalog_matches": sum(1 for row in mappings if row["match_type"].startswith("CATALOG_")),
            "new_attributes": len(new_names),
            "system_fields": len(system_fields),
            "to_review": len(unresolved),
        },
        "attribute_mappings": mappings,
        "new_attributes": [
            {
                "attribute": name,
                "categories": usage.get(name, []),
                "source": input_path.as_posix(),
                "confidence": min(
                    row["confidence"] for row in mappings if row["canonical_name"] == name
                ),
                "status": (
                    "DO_SPRAWDZENIA"
                    if any(row["canonical_name"] == name and row["status"] == "DO_SPRAWDZENIA" for row in mappings)
                    else "OK"
                ),
            }
            for name in new_names
        ],
        "system_fields": system_fields,
        "categories": sections,
    }


def write_reports(reports_dir: Path, knowledge: dict[str, Any]) -> None:
    ensure_dir(reports_dir)
    mappings = knowledge["attribute_mappings"]
    with (reports_dir / "attribute_mappings.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(mappings[0].keys()) if mappings else [])
        if mappings:
            writer.writeheader()
            writer.writerows(mappings)

    summary = knowledge["summary"]
    lines = [
        "# Import wiedzy o atrybutach oświetlenia",
        "",
        f"- Kategorie: {summary['categories']}",
        f"- Nazwy źródłowe: {summary['source_attribute_names']}",
        f"- Dopasowane do catalog knowledge: {summary['catalog_matches']}",
        f"- Nowe atrybuty: {summary['new_attributes']}",
        f"- Pola systemowe: {summary['system_fields']}",
        f"- Do sprawdzenia: {summary['to_review']}",
        "",
        "## Nowe atrybuty",
        "",
    ]
    for row in knowledge["new_attributes"]:
        lines.append(f"- {row['attribute']} ({len(row['categories'])} kategorii, {row['status']})")
    lines.extend(["", "## Do sprawdzenia", ""])
    review_rows = [row for row in mappings if row["status"] == "DO_SPRAWDZENIA"]
    if review_rows:
        for row in review_rows:
            lines.append(f"- {row['source_name']} -> {row['canonical_name']}")
    else:
        lines.append("- Brak")
    (reports_dir / "summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def import_knowledge(input_path: Path, catalog_path: Path) -> dict[str, Any]:
    workbook = pd.ExcelFile(input_path)
    raw_sections: list[dict[str, Any]] = []
    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(input_path, sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
        raw_sections.extend(parse_sheet(df, sheet_name))
    catalog_by_normalized, catalog_names = load_catalog_attributes(catalog_path)
    sections, mappings = canonicalize_sections(raw_sections, catalog_by_normalized, catalog_names)
    apply_category_extra_attributes(sections)
    for mapping in mappings:
        mapping["source"] = input_path.as_posix()
    return build_knowledge(input_path, catalog_path, sections, mappings, catalog_names)


def main() -> None:
    parser = argparse.ArgumentParser(description="Importuje macierz atrybutow oswietlenia do lokalnej wiedzy YAML.")
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--reports-dir", default=DEFAULT_REPORTS_DIR)
    args = parser.parse_args()

    input_path = Path(args.input)
    catalog_path = Path(args.catalog_knowledge)
    output_path = Path(args.output)
    knowledge = import_knowledge(input_path, catalog_path)
    ensure_dir(output_path.parent)
    output_path.write_text(
        yaml.safe_dump(knowledge, allow_unicode=True, sort_keys=False, width=120),
        encoding="utf-8",
    )
    write_reports(Path(args.reports_dir), knowledge)
    print(json.dumps(knowledge["summary"], ensure_ascii=True, sort_keys=True))


if __name__ == "__main__":
    main()
