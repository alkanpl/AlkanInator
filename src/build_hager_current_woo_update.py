from __future__ import annotations

from copy import copy
import hashlib
import json
from pathlib import Path
from typing import Any

import yaml
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "input" / "Aparatura-Modulowa-Hager.xlsx"
FINAL = ROOT / "output" / "Hager_nazwy_po_decyzjach_2026-08-07.xlsx"
CONFIG = ROOT / "configs" / "hager_woocommerce_implementation.yaml"
OUTPUT = ROOT / "output" / "Aparatura-Modulowa-Hager_Woo_nazwy_atrybuty_2026-08-12.xlsx"
REPORT = ROOT / "reports" / "hager_current_woo_update_2026-08-12" / "summary.json"

PREFIX_CURRENT = "Atrybut produktu: "
PREFIX_FINAL = "Atrybut Produktu: "


def normalized(value: Any) -> str:
    return str(value).strip().casefold() if value is not None else ""


def clean_value(target: str, value: Any) -> Any:
    if value in (None, ""):
        return None
    if target == "Kategoria użytkowania":
        return str(value).replace(" / ", "/")
    if target == "Rodzaj sieci":
        return str(value).replace("TN-S / TT", "TN-S, TT")
    if target == "Rozmiar wkładki bezpiecznikowej":
        return str(value).replace("10x38", "10×38").replace("22x58", "22×58")
    if target == "Technologia ogranicznika przepięć":
        text = str(value).strip()
        if text.casefold() == "mov":
            return "MOV (warystorowa)"
        if "kombin" in text.casefold() or "iskier" in text.casefold():
            return "Kombinowana (iskiernikowa)"
    return value


def source_value(final_row, final_index, sources):
    for source in sources:
        header = PREFIX_FINAL + source
        index = final_index.get(normalized(header))
        if index is None:
            continue
        value = final_row[index]
        if value not in (None, ""):
            return value
    return None


def copy_header_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.protection = copy(source_cell.protection)
        target_cell.number_format = source_cell.number_format


def build():
    input_hash_before = hashlib.sha256(INPUT.read_bytes()).hexdigest()
    config = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))

    final_workbook = load_workbook(FINAL, read_only=True, data_only=True)
    final_sheet = final_workbook["Worksheet"]
    final_rows = list(final_sheet.iter_rows(values_only=True))
    final_headers = list(final_rows[0])
    final_index = {normalized(header): index for index, header in enumerate(final_headers) if header is not None}
    final_sku_index = final_index[normalized("SKU")]
    final_id_index = final_index[normalized("id")]
    final_title_index = final_index[normalized("Title")]
    final_by_sku = {row[final_sku_index]: row for row in final_rows[1:] if row[final_sku_index]}
    old_skus = {sku for sku, row in final_by_sku.items() if row[final_id_index] not in (None, "")}

    workbook = load_workbook(INPUT)
    sheet = workbook["Worksheet"]
    current_headers = [cell.value for cell in sheet[1]]
    current_index = {normalized(header): index for index, header in enumerate(current_headers) if header is not None}
    sku_index = current_index[normalized("Sku")]
    title_index = current_index[normalized("Title")]

    current_skus = [sheet.cell(row, sku_index + 1).value for row in range(2, sheet.max_row + 1)]
    missing_in_final = sorted(sku for sku in current_skus if sku not in final_by_sku)
    missing_in_current = sorted(sku for sku in final_by_sku if sku not in set(current_skus))
    if missing_in_final or missing_in_current:
        raise RuntimeError(f"Niezgodność SKU: brak w final={missing_in_final}; brak w aktualnym={missing_in_current}")

    # Tylko stare produkty dostają przygotowane wcześniej nazwy. Nowo dodane pozostają bez zmian.
    title_changes = 0
    preserved_new_titles = 0
    for row_number in range(2, sheet.max_row + 1):
        sku = sheet.cell(row_number, sku_index + 1).value
        if sku in old_skus:
            final_title = final_by_sku[sku][final_title_index]
            if sheet.cell(row_number, title_index + 1).value != final_title:
                sheet.cell(row_number, title_index + 1).value = final_title
                title_changes += 1
        else:
            preserved_new_titles += 1

    # Końcowe atrybuty: 17 nowych globalnych oraz 2 wcześniej przemianowane pola istniejące.
    attribute_specs = []
    for item in config["attribute_actions"]:
        attribute_specs.append({
            "target": item["target_name"],
            "sources": item["source_attributes"],
            "kind": item["action"],
        })

    # Istniejące pola, które mają zostać zachowane i uzupełnione po ostatnich decyzjach.
    attribute_specs.extend([
        {"target": "Liczba faz", "sources": ["Liczba Faz"], "kind": "UZUPEŁNIJ ISTNIEJĄCY"},
        {"target": "Typ wyłącznika", "sources": ["Typ Wyłącznika"], "kind": "UZUPEŁNIJ ISTNIEJĄCY"},
        {"target": "Napięcie [V]", "sources": ["Napięcie [V]"], "kind": "UZUPEŁNIJ ISTNIEJĄCY"},
    ])

    # Deduplicate specs by target name, preserving the config order.
    deduplicated_specs = []
    seen_targets = set()
    for spec in attribute_specs:
        key = normalized(spec["target"])
        if key in seen_targets:
            continue
        seen_targets.add(key)
        deduplicated_specs.append(spec)

    new_columns = []
    populated_counts = {}
    changed_cells = 0
    source_style_cell = sheet.cell(1, sheet.max_column)
    for spec in deduplicated_specs:
        target = spec["target"]
        header = PREFIX_CURRENT + target
        key = normalized(header)
        column_index = current_index.get(key)
        if column_index is None:
            column_index = sheet.max_column
            target_cell = sheet.cell(1, column_index + 1, header)
            copy_header_style(source_style_cell, target_cell)
            target_cell.alignment = Alignment(wrap_text=True, vertical="center")
            sheet.column_dimensions[target_cell.column_letter].width = 30
            current_index[key] = column_index
            new_columns.append(target)

        nonempty = 0
        for row_number in range(2, sheet.max_row + 1):
            sku = sheet.cell(row_number, sku_index + 1).value
            value = source_value(final_by_sku[sku], final_index, spec["sources"])
            value = clean_value(target, value)
            if value not in (None, ""):
                nonempty += 1
            cell = sheet.cell(row_number, column_index + 1)
            if cell.value != value:
                cell.value = value
                changed_cells += 1
        populated_counts[target] = nonempty

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(sheet.max_row, sheet.max_column).coordinate}"
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 80
    sheet.row_dimensions[1].height = max(sheet.row_dimensions[1].height or 15, 42)
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="2F75B5")
        cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    main_widths = {"A": 12, "B": 76, "C": 18, "D": 20, "E": 14, "F": 18, "G": 12, "H": 10}
    for column, width in main_widths.items():
        sheet.column_dimensions[column].width = width
    for column_number in range(9, sheet.max_column + 1):
        column_letter = sheet.cell(1, column_number).column_letter
        if not sheet.column_dimensions[column_letter].width or sheet.column_dimensions[column_letter].width < 20:
            sheet.column_dimensions[column_letter].width = 28
    for row_number in range(2, sheet.max_row + 1):
        sheet.cell(row_number, sku_index + 1).number_format = "@"
        gtin_index = current_index.get(normalized("hwp_product_gtin"))
        if gtin_index is not None:
            sheet.cell(row_number, gtin_index + 1).number_format = "0"

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT)
    input_hash_after = hashlib.sha256(INPUT.read_bytes()).hexdigest()
    if input_hash_before != input_hash_after:
        raise RuntimeError("Plik wejściowy został nieoczekiwanie zmieniony")

    # Validation from the saved artifact.
    check_workbook = load_workbook(OUTPUT, read_only=True, data_only=True)
    check_sheet = check_workbook["Worksheet"]
    check_rows = list(check_sheet.iter_rows(values_only=True))
    check_headers = list(check_rows[0])
    check_index = {normalized(header): index for index, header in enumerate(check_headers) if header is not None}
    check_sku_index = check_index[normalized("Sku")]
    check_title_index = check_index[normalized("Title")]
    check_id_index = check_index[normalized("id")]

    errors = []
    if len(check_rows) - 1 != 153:
        errors.append(f"Nieprawidłowa liczba produktów: {len(check_rows) - 1}")
    if len({row[check_sku_index] for row in check_rows[1:]}) != 153:
        errors.append("Wykryto duplikaty SKU")
    if any(row[check_id_index] in (None, "") for row in check_rows[1:]):
        errors.append("Wykryto produkt bez ID Woo")
    old_title_mismatches = []
    new_title_changes = []
    original_workbook = load_workbook(INPUT, read_only=True, data_only=True)
    original_rows = list(original_workbook["Worksheet"].iter_rows(values_only=True))
    original_headers = list(original_rows[0])
    original_index = {normalized(header): index for index, header in enumerate(original_headers) if header is not None}
    original_titles = {row[original_index[normalized("Sku")]]: row[original_index[normalized("Title")]] for row in original_rows[1:]}
    for row in check_rows[1:]:
        sku = row[check_sku_index]
        title = row[check_title_index]
        if sku in old_skus and title != final_by_sku[sku][final_title_index]:
            old_title_mismatches.append(sku)
        if sku not in old_skus and title != original_titles[sku]:
            new_title_changes.append(sku)
    if old_title_mismatches:
        errors.append(f"Niepoprawne nazwy starych produktów: {old_title_mismatches}")
    if new_title_changes:
        errors.append(f"Zmieniono nazwy nowych produktów: {new_title_changes}")

    required_targets = [spec["target"] for spec in deduplicated_specs]
    missing_attribute_columns = [
        target for target in required_targets
        if normalized(PREFIX_CURRENT + target) not in check_index
    ]
    if missing_attribute_columns:
        errors.append(f"Brak kolumn atrybutów: {missing_attribute_columns}")

    report = {
        "input": str(INPUT),
        "output": str(OUTPUT),
        "input_sha256": input_hash_before,
        "input_preserved": input_hash_before == input_hash_after,
        "products": len(check_rows) - 1,
        "sku_matches": len(current_skus),
        "woo_ids_present": sum(row[check_id_index] not in (None, "") for row in check_rows[1:]),
        "old_products_renamed": title_changes,
        "new_product_titles_preserved": preserved_new_titles,
        "new_attribute_columns_added": new_columns,
        "new_attribute_column_count": len(new_columns),
        "target_attribute_count": len(deduplicated_specs),
        "attribute_nonempty_counts": populated_counts,
        "attribute_cells_changed": changed_cells,
        "errors": errors,
        "status": "OK" if not errors else "ERROR",
    }
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if errors:
        raise RuntimeError("; ".join(errors))
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    build()
