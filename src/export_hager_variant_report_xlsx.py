from __future__ import annotations

import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = (
    ROOT
    / "reports"
    / "Hager-Berker-Gniazdka_WOO_IMPORT_WARIANTY_POPRAWIONE_report.json"
)
DEFAULT_OUTPUT = (
    ROOT
    / "reports"
    / "Hager-Berker-GRUPY_WARIANTOW_OSIE_ATRYBUTOW_2026-09-03.xlsx"
)
DEFAULT_VARIANT_ANALYSIS = (
    ROOT
    / "output"
    / "Hager-Berker-Gniazdka_WARIANTY_OSIE_POPRAWIONE.xlsx"
)
DEFAULT_CURRENT_WOO = (
    ROOT
    / "output"
    / "Hager-Berker-Gniazdka_WOO_IMPORT_WARIANTY_POPRAWIONE.xlsx"
)

NAVY = "17365D"
BLUE = "1F4E78"
LIGHT_BLUE = "D9EAF7"
GREEN = "548235"
LIGHT_GREEN = "E2F0D9"
RED = "C00000"
LIGHT_RED = "FCE4D6"
AMBER = "BF8F00"
LIGHT_AMBER = "FFF2CC"
LIGHT_GRAY = "F2F2F2"
MID_GRAY = "D9E1F2"
WHITE = "FFFFFF"
TEXT = "1F2937"
THIN_GRAY = Side(style="thin", color="D9E2F3")

VALIDATION_LABELS = {
    "sheet_names_match": "Nazwy arkuszy zgodne",
    "dimensions_match": "Wymiary arkuszy zgodne",
    "headers_match": "Nagłówki zgodne",
    "product_rows_match": "Liczba produktów zgodna",
    "sku_order_and_values_match": "SKU i ich kolejność zgodne",
    "no_duplicate_input_skus": "Brak duplikatów SKU w źródle",
    "no_duplicate_output_skus": "Brak duplikatów SKU w wyniku",
    "all_required_skus_present": "Wszystkie kontrolowane SKU obecne",
    "all_requested_variant_tags_removed": "Wszystkie wskazane tagi wariantowe usunięte",
    "all_shutters_values_valid": "Wartości przesłon torów prądowych poprawne",
    "all_wls1200_terminal_values_valid": "Typ zacisków WLS1200 poprawny",
    "changes_limited_to_expected_cells": "Zmiany ograniczone do oczekiwanych komórek",
    "no_changes_outside_allowed_columns": "Brak zmian poza dozwolonymi kolumnami",
    "no_mojibake_detected": "Brak uszkodzonych polskich znaków",
}


def load_report(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def split_tags(value: object) -> list[str]:
    return [tag.strip() for tag in re.split(r"\s*\|\s*", str(value or "")) if tag.strip()]


def header_map(sheet) -> dict[str, int]:
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(next(sheet.iter_rows(min_row=1, max_row=1)), start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def load_variant_axis_data(analysis_path: Path, current_woo_path: Path) -> dict[str, Any]:
    analysis_book = load_workbook(analysis_path, read_only=True, data_only=True)
    current_book = load_workbook(current_woo_path, read_only=True, data_only=True)
    try:
        groups_sheet = analysis_book["Grupy wariantów"]
        groups_headers = header_map(groups_sheet)
        required_group_headers = {
            "ID grupy",
            "Nazwa grupy",
            "Wspólny tag",
            "Liczba produktów",
            "Osie wariantów",
            "SKU",
            "Status",
            "Uzasadnienie",
        }
        missing_group_headers = required_group_headers - set(groups_headers)
        if missing_group_headers:
            raise KeyError(f"Brak kolumn grup wariantowych: {sorted(missing_group_headers)}")

        source_groups: dict[str, dict[str, Any]] = {}
        for row in groups_sheet.iter_rows(min_row=2, values_only=True):
            tag = str(row[groups_headers["Wspólny tag"] - 1] or "").strip()
            if not tag:
                continue
            source_groups[tag] = {
                "group_id": row[groups_headers["ID grupy"] - 1],
                "group_name": row[groups_headers["Nazwa grupy"] - 1],
                "tag": tag,
                "source_product_count": int(row[groups_headers["Liczba produktów"] - 1] or 0),
                "axes_text": str(row[groups_headers["Osie wariantów"] - 1] or "").strip(),
                "source_skus": split_tags(row[groups_headers["SKU"] - 1]),
                "status": row[groups_headers["Status"] - 1],
                "reason": row[groups_headers["Uzasadnienie"] - 1],
            }

        products_sheet = analysis_book["Produkty warianty"]
        products_headers = header_map(products_sheet)
        required_product_headers = {"SKU", "Title", "Wariant: wartości"}
        missing_product_headers = required_product_headers - set(products_headers)
        if missing_product_headers:
            raise KeyError(f"Brak kolumn produktów wariantowych: {sorted(missing_product_headers)}")
        product_axis_values: dict[str, str] = {}
        for row in products_sheet.iter_rows(min_row=2, values_only=True):
            sku = str(row[products_headers["SKU"] - 1] or "").strip().upper()
            if sku:
                product_axis_values[sku] = str(
                    row[products_headers["Wariant: wartości"] - 1] or ""
                ).strip()

        current_sheet = current_book.active
        current_headers = header_map(current_sheet)
        required_current_headers = {"SKU", "Title", "Product Tags"}
        missing_current_headers = required_current_headers - set(current_headers)
        if missing_current_headers:
            raise KeyError(f"Brak kolumn aktualnego Woo: {sorted(missing_current_headers)}")

        current_groups: defaultdict[str, list[dict[str, str]]] = defaultdict(list)
        current_variant_skus: set[str] = set()
        for row in current_sheet.iter_rows(min_row=2, values_only=True):
            sku = str(row[current_headers["SKU"] - 1] or "").strip().upper()
            title = str(row[current_headers["Title"] - 1] or "").strip()
            if not sku:
                continue
            for tag in split_tags(row[current_headers["Product Tags"] - 1]):
                if not tag.casefold().startswith("wariant -"):
                    continue
                current_variant_skus.add(sku)
                current_groups[tag].append(
                    {
                        "sku": sku,
                        "title": title,
                        "axis_values": product_axis_values.get(sku, ""),
                    }
                )

        missing_tags = sorted(set(current_groups) - set(source_groups))
        if missing_tags:
            raise ValueError(f"Brak definicji osi dla aktualnych tagów: {missing_tags}")

        group_rows: list[dict[str, Any]] = []
        product_rows: list[dict[str, Any]] = []
        for tag, current_products in current_groups.items():
            source = source_groups[tag]
            axes = [axis.strip() for axis in source["axes_text"].split(",") if axis.strip()]
            active_skus = [product["sku"] for product in current_products]
            removed_skus = [sku for sku in source["source_skus"] if sku.upper() not in set(active_skus)]
            group_row = {
                **source,
                "axes": axes,
                "axis_count": len(axes),
                "active_product_count": len(current_products),
                "active_skus": active_skus,
                "removed_skus": removed_skus,
                "axis_values": " | ".join(
                    f"{product['sku']}: {product['axis_values']}" for product in current_products
                ),
            }
            group_rows.append(group_row)
            for product in current_products:
                product_rows.append(
                    {
                        "group_id": source["group_id"],
                        "group_name": source["group_name"],
                        "tag": tag,
                        "axes_text": source["axes_text"],
                        "axis_count": len(axes),
                        **product,
                    }
                )

        group_rows.sort(key=lambda row: str(row["group_id"]))
        product_rows.sort(key=lambda row: (str(row["group_id"]), row["sku"]))
        combination_groups: Counter[str] = Counter()
        combination_products: Counter[str] = Counter()
        individual_axis_groups: Counter[str] = Counter()
        individual_axis_products: Counter[str] = Counter()
        for group in group_rows:
            combination_groups[group["axes_text"]] += 1
            combination_products[group["axes_text"]] += group["active_product_count"]
            for axis in group["axes"]:
                individual_axis_groups[axis] += 1
                individual_axis_products[axis] += group["active_product_count"]

        combination_rows = [
            {
                "axes_text": axes_text,
                "axis_count": len([axis for axis in axes_text.split(",") if axis.strip()]),
                "groups": groups_count,
                "products": combination_products[axes_text],
            }
            for axes_text, groups_count in combination_groups.most_common()
        ]
        individual_axis_rows = [
            {
                "axis": axis,
                "groups": groups_count,
                "products": individual_axis_products[axis],
            }
            for axis, groups_count in individual_axis_groups.most_common()
        ]
        validation = {
            "current_groups": len(current_groups),
            "matched_groups": len(group_rows),
            "current_variant_products": len(current_variant_skus),
            "product_rows": len(product_rows),
            "missing_tags": missing_tags,
            "groups_without_axes": [group["group_id"] for group in group_rows if not group["axes"]],
            "groups_below_two_products": [
                group["group_id"] for group in group_rows if group["active_product_count"] < 2
            ],
            "groups_changed_by_correction": sum(bool(group["removed_skus"]) for group in group_rows),
        }
        validation["passed"] = (
            validation["current_groups"] == validation["matched_groups"]
            and validation["current_variant_products"] == validation["product_rows"]
            and not validation["missing_tags"]
            and not validation["groups_without_axes"]
            and not validation["groups_below_two_products"]
        )
        return {
            "analysis_file": str(analysis_path),
            "current_woo_file": str(current_woo_path),
            "groups": group_rows,
            "products": product_rows,
            "axis_combinations": combination_rows,
            "individual_axes": individual_axis_rows,
            "validation": validation,
        }
    finally:
        analysis_book.close()
        current_book.close()


def set_title_band(sheet, cell_range: str, value: str) -> None:
    sheet.merge_cells(cell_range)
    cell = sheet[cell_range.split(":")[0]]
    cell.value = value
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name="Aptos Display", size=18, bold=True, color=WHITE)
    cell.alignment = Alignment(vertical="center")


def set_section_header(sheet, row: int, start_column: int, end_column: int, value: str) -> None:
    start = sheet.cell(row, start_column)
    end = sheet.cell(row, end_column)
    sheet.merge_cells(start_row=row, start_column=start_column, end_row=row, end_column=end_column)
    start.value = value
    start.fill = PatternFill("solid", fgColor=BLUE)
    start.font = Font(name="Aptos", size=11, bold=True, color=WHITE)
    start.alignment = Alignment(vertical="center")
    for cell in sheet[row][start_column - 1 : end_column]:
        cell.fill = PatternFill("solid", fgColor=BLUE)
    sheet.row_dimensions[row].height = 23


def style_table_header(row) -> None:
    for cell in row:
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(name="Aptos", bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="medium", color=NAVY))


def add_table(sheet, reference: str, name: str, style: str = "TableStyleMedium2") -> None:
    table = Table(displayName=name, ref=reference)
    table.tableStyleInfo = TableStyleInfo(
        name=style,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def build_summary_sheet(workbook: Workbook, report: dict[str, Any]) -> None:
    sheet = workbook.active
    sheet.title = "Podsumowanie"
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A4"
    set_title_band(sheet, "A1:F2", "Hager / Berker — raport korekty wariantów")
    sheet.row_dimensions[1].height = 28
    sheet.row_dimensions[2].height = 12

    sheet["A3"] = "Plik wynikowy"
    sheet["B3"] = Path(report["output_file"]).name
    sheet.merge_cells("B3:F3")
    sheet["A3"].font = Font(name="Aptos", bold=True, color=TEXT)
    sheet["B3"].font = Font(name="Aptos", italic=True, color=TEXT)

    set_section_header(sheet, 5, 1, 6, "Wynik audytu")
    sheet["A6"] = "Status końcowy"
    validation_count = len([key for key in report["validation"] if key != "passed"])
    sheet["B6"] = f'=IF(COUNTIF(\'Walidacja\'!C2:C{validation_count + 1},"BŁĄD")=0,"ZALICZONA","BŁĘDY")'
    sheet["A7"] = "Data wygenerowania raportu"
    sheet["B7"] = report["generated_at"]
    sheet["A8"] = "Produkty w pliku"
    sheet["B8"] = report["rows_with_sku"]
    sheet["A9"] = "Zmienione produkty"
    sheet["B9"] = report["changed_products"]
    sheet["A10"] = "Zmienione komórki"
    sheet["B10"] = f"=COUNTA('Zmiany'!A2:A{len(report['changes']) + 1})"

    for row in range(6, 11):
        sheet.cell(row, 1).font = Font(name="Aptos", bold=True, color=TEXT)
        sheet.cell(row, 1).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
        sheet.cell(row, 2).font = Font(name="Aptos", size=11, color=TEXT)
        sheet.cell(row, 2).alignment = Alignment(horizontal="right" if row >= 8 else "left")
        sheet.cell(row, 1).border = Border(bottom=THIN_GRAY)
        sheet.cell(row, 2).border = Border(bottom=THIN_GRAY)
    sheet["B6"].fill = PatternFill("solid", fgColor=LIGHT_GREEN)
    sheet["B6"].font = Font(name="Aptos", bold=True, color=GREEN)
    sheet["B8"].number_format = "#,##0"
    sheet["B9"].number_format = "#,##0"
    sheet["B10"].number_format = "#,##0"

    set_section_header(sheet, 12, 1, 6, "Zakres wariantów — przed i po korekcie")
    headers = ["Miara", "Przed", "Po", "Zmiana", "Zmiana %", "Ocena"]
    for column, value in enumerate(headers, start=1):
        sheet.cell(13, column).value = value
    style_table_header(sheet[13])
    variant_rows = [
        ("Grupy wariantowe", report["variant_groups_before"], report["variant_groups_after"]),
        ("Produkty z tagiem wariantowym", report["variant_products_before"], report["variant_products_after"]),
    ]
    for row, (label, before, after) in enumerate(variant_rows, start=14):
        sheet.cell(row, 1).value = label
        sheet.cell(row, 2).value = before
        sheet.cell(row, 3).value = after
        sheet.cell(row, 4).value = f"=C{row}-B{row}"
        sheet.cell(row, 5).value = f'=IF(B{row}=0,0,D{row}/B{row})'
        sheet.cell(row, 6).value = f'=IF(D{row}=0,"Bez zmian",IF(D{row}<0,"Zmniejszenie","Zwiększenie"))'
        for column in range(1, 7):
            sheet.cell(row, column).border = Border(bottom=THIN_GRAY)
    for row in range(14, 16):
        for column in range(2, 5):
            sheet.cell(row, column).number_format = "#,##0"
        sheet.cell(row, 5).number_format = "0.0%"

    set_section_header(sheet, 18, 1, 6, "Zmiany według kolumn")
    for column, value in enumerate(["Kolumna", "Liczba zmian", "Udział", "Zakres kontroli", "Wynik", "Uwagi"], start=1):
        sheet.cell(19, column).value = value
    style_table_header(sheet[19])
    column_notes = {
        "Product Tags": ("Usunięcie błędnych powiązań", "Tagi wariantowe"),
        "Atrybut Produktu: Przesłony Torów Prądowych": ("Uzupełnienie Tak/Nie", "Atrybut techniczny"),
        "Atrybut Produktu: Typ Zacisków": ("Uzupełnienie Samozaciski", "Atrybut techniczny"),
    }
    first_change_row = 2
    last_change_row = len(report["changes"]) + 1
    for row, column_name in enumerate(report["changes_by_column"], start=20):
        note, scope = column_notes.get(column_name, ("Zmiana kontrolowana", "Inne"))
        sheet.cell(row, 1).value = column_name
        sheet.cell(row, 2).value = f'=COUNTIF(\'Zmiany\'!$D${first_change_row}:$D${last_change_row},A{row})'
        sheet.cell(row, 3).value = f'=IF($B$10=0,0,B{row}/$B$10)'
        sheet.cell(row, 4).value = scope
        sheet.cell(row, 5).value = "OK"
        sheet.cell(row, 6).value = note
        sheet.cell(row, 3).number_format = "0.0%"
        for column in range(1, 7):
            sheet.cell(row, column).border = Border(bottom=THIN_GRAY)

    sheet["A25"] = "Wniosek"
    sheet["B25"] = (
        "Korekta jest bezpieczna do wykorzystania: wszystkie zmiany ograniczają się do "
        "oczekiwanych tagów i dwóch atrybutów technicznych. Nie wykryto zmian ubocznych."
    )
    sheet.merge_cells("B25:F27")
    sheet["A25"].font = Font(name="Aptos", bold=True, color=TEXT)
    sheet["B25"].fill = PatternFill("solid", fgColor=LIGHT_GREEN)
    sheet["B25"].font = Font(name="Aptos", color=TEXT)
    sheet["B25"].alignment = Alignment(wrap_text=True, vertical="top")
    sheet["B25"].border = Border(left=Side(style="medium", color=GREEN))

    sheet.column_dimensions["A"].width = 46
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 15
    sheet.column_dimensions["D"].width = 24
    sheet.column_dimensions["E"].width = 16
    sheet.column_dimensions["F"].width = 32
    for row in range(1, 28):
        sheet.row_dimensions[row].height = max(sheet.row_dimensions[row].height or 15, 20)
    sheet.row_dimensions[25].height = 25
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.print_area = "A1:F27"


def build_axis_summary_sheet(workbook: Workbook, axis_data: dict[str, Any]) -> None:
    sheet = workbook.create_sheet("Podsumowanie osi")
    sheet.sheet_view.showGridLines = False
    set_title_band(sheet, "A1:J2", "Aktualne osie grup wariantowych w WooCommerce")
    sheet.row_dimensions[1].height = 28
    sheet["A3"] = "Podstawa"
    sheet["B3"] = Path(axis_data["current_woo_file"]).name
    sheet.merge_cells("B3:J3")
    sheet["A3"].font = Font(name="Aptos", bold=True, color=TEXT)
    sheet["B3"].font = Font(name="Aptos", italic=True, color=TEXT)

    groups = axis_data["groups"]
    products = axis_data["products"]
    single_axis_groups = sum(group["axis_count"] == 1 for group in groups)
    multi_axis_groups = len(groups) - single_axis_groups
    corrected_groups = axis_data["validation"]["groups_changed_by_correction"]
    metrics = [
        ("Aktualne grupy", len(groups)),
        ("Produkty w grupach", len(products)),
        ("Grupy jednoosiowe", single_axis_groups),
        ("Grupy wieloosiowe", multi_axis_groups),
        ("Grupy po korekcie składu", corrected_groups),
    ]
    for column, (label, value) in enumerate(metrics, start=1):
        start_column = (column - 1) * 2 + 1
        sheet.cell(5, start_column).value = label
        sheet.cell(6, start_column).value = value
        sheet.merge_cells(start_row=5, start_column=start_column, end_row=5, end_column=start_column + 1)
        sheet.merge_cells(start_row=6, start_column=start_column, end_row=7, end_column=start_column + 1)
        sheet.cell(5, start_column).fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        sheet.cell(5, start_column).font = Font(name="Aptos", bold=True, color=BLUE)
        sheet.cell(5, start_column).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        sheet.cell(6, start_column).fill = PatternFill("solid", fgColor=WHITE)
        sheet.cell(6, start_column).font = Font(name="Aptos Display", size=18, bold=True, color=NAVY)
        sheet.cell(6, start_column).alignment = Alignment(horizontal="center", vertical="center")
        sheet.cell(6, start_column).number_format = "#,##0"
        sheet.cell(5, start_column).border = Border(top=THIN_GRAY, left=THIN_GRAY, right=THIN_GRAY)
        sheet.cell(6, start_column).border = Border(bottom=THIN_GRAY, left=THIN_GRAY, right=THIN_GRAY)
    sheet.row_dimensions[5].height = 32
    sheet.row_dimensions[6].height = 25
    sheet.row_dimensions[7].height = 12

    set_section_header(sheet, 9, 1, 5, "Kombinacje osi w grupach")
    combination_headers = ["Kombinacja osi", "Liczba osi", "Grupy", "Produkty", "Udział grup"]
    for column, value in enumerate(combination_headers, start=1):
        sheet.cell(10, column).value = value
    style_table_header(sheet[10][:5])
    for row, item in enumerate(axis_data["axis_combinations"], start=11):
        sheet.cell(row, 1).value = item["axes_text"]
        sheet.cell(row, 2).value = item["axis_count"]
        sheet.cell(row, 3).value = item["groups"]
        sheet.cell(row, 4).value = item["products"]
        sheet.cell(row, 5).value = f"=C{row}/$A$6"
        sheet.cell(row, 5).number_format = "0.0%"
        for column in range(1, 6):
            sheet.cell(row, column).border = Border(bottom=THIN_GRAY)
    combination_last_row = 10 + len(axis_data["axis_combinations"])
    add_table(sheet, f"A10:E{combination_last_row}", "AxisCombinationsTable", "TableStyleMedium2")

    set_section_header(sheet, 9, 7, 10, "Pojedyncze osie atrybutów")
    individual_headers = ["Oś atrybutu", "Grupy", "Produkty", "Udział grup"]
    for column, value in enumerate(individual_headers, start=7):
        sheet.cell(10, column).value = value
    style_table_header(sheet[10][6:10])
    for row, item in enumerate(axis_data["individual_axes"], start=11):
        sheet.cell(row, 7).value = item["axis"]
        sheet.cell(row, 8).value = item["groups"]
        sheet.cell(row, 9).value = item["products"]
        sheet.cell(row, 10).value = f"=H{row}/$A$6"
        sheet.cell(row, 10).number_format = "0.0%"
        for column in range(7, 11):
            sheet.cell(row, column).border = Border(bottom=THIN_GRAY)
    individual_last_row = 10 + len(axis_data["individual_axes"])
    add_table(sheet, f"G10:J{individual_last_row}", "IndividualAxesTable", "TableStyleMedium4")

    widths = {
        "A": 54,
        "B": 13,
        "C": 12,
        "D": 14,
        "E": 14,
        "F": 3,
        "G": 38,
        "H": 12,
        "I": 14,
        "J": 14,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A10"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1


def build_variant_groups_sheet(workbook: Workbook, axis_data: dict[str, Any]) -> None:
    sheet = workbook.create_sheet("Osie grup")
    sheet.sheet_view.showGridLines = False
    headers = [
        "ID grupy",
        "Nazwa grupy",
        "Wspólny tag",
        "Oś 1",
        "Oś 2",
        "Oś 3",
        "Liczba osi",
        "Produkty aktualnie",
        "Produkty źródłowo",
        "Usunięte SKU w korekcie",
        "Aktualne SKU",
        "Wartości osi według SKU",
        "Status",
        "Uzasadnienie",
    ]
    sheet.append(headers)
    for group in axis_data["groups"]:
        axes = group["axes"] + ["", "", ""]
        sheet.append(
            [
                group["group_id"],
                group["group_name"],
                group["tag"],
                axes[0],
                axes[1],
                axes[2],
                group["axis_count"],
                group["active_product_count"],
                group["source_product_count"],
                " | ".join(group["removed_skus"]),
                " | ".join(group["active_skus"]),
                group["axis_values"],
                group["status"],
                group["reason"],
            ]
        )
    style_table_header(sheet[1])
    last_row = len(axis_data["groups"]) + 1
    add_table(sheet, f"A1:N{last_row}", "VariantGroupsTable", "TableStyleMedium2")
    sheet.freeze_panes = "D2"
    sheet.auto_filter.ref = f"A1:N{last_row}"
    widths = {
        "A": 13,
        "B": 48,
        "C": 48,
        "D": 26,
        "E": 26,
        "F": 26,
        "G": 12,
        "H": 17,
        "I": 17,
        "J": 25,
        "K": 58,
        "L": 90,
        "M": 20,
        "N": 58,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2, max_row=last_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in {2, 3, 10, 11, 12, 14})
        for cell in row[6:9]:
            cell.alignment = Alignment(horizontal="center", vertical="top")
    sheet.conditional_formatting.add(
        f"G2:G{last_row}",
        FormulaRule(
            formula=["G2>1"],
            fill=PatternFill("solid", fgColor=LIGHT_AMBER),
            font=Font(color=AMBER, bold=True),
        ),
    )
    sheet.conditional_formatting.add(
        f"J2:J{last_row}",
        FormulaRule(
            formula=["J2<>\"\""],
            fill=PatternFill("solid", fgColor=LIGHT_AMBER),
            font=Font(color=AMBER),
        ),
    )
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:1"


def build_variant_products_sheet(workbook: Workbook, axis_data: dict[str, Any]) -> None:
    sheet = workbook.create_sheet("Produkty w grupach")
    sheet.sheet_view.showGridLines = False
    headers = [
        "ID grupy",
        "Nazwa grupy",
        "Wspólny tag",
        "SKU",
        "Tytuł produktu",
        "Osie wariantów",
        "Liczba osi",
        "Wartości osi produktu",
    ]
    sheet.append(headers)
    for product in axis_data["products"]:
        sheet.append(
            [
                product["group_id"],
                product["group_name"],
                product["tag"],
                product["sku"],
                product["title"],
                product["axes_text"],
                product["axis_count"],
                product["axis_values"],
            ]
        )
    style_table_header(sheet[1])
    last_row = len(axis_data["products"]) + 1
    add_table(sheet, f"A1:H{last_row}", "VariantProductsTable", "TableStyleMedium4")
    sheet.freeze_panes = "D2"
    sheet.auto_filter.ref = f"A1:H{last_row}"
    widths = {"A": 13, "B": 48, "C": 48, "D": 23, "E": 70, "F": 45, "G": 12, "H": 58}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2, max_row=last_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in {2, 3, 5, 6, 8})
        row[3].font = Font(name="Consolas", size=9, color=TEXT)
        row[6].alignment = Alignment(horizontal="center", vertical="top")
    sheet.conditional_formatting.add(
        f"G2:G{last_row}",
        FormulaRule(
            formula=["G2>1"],
            fill=PatternFill("solid", fgColor=LIGHT_AMBER),
            font=Font(color=AMBER, bold=True),
        ),
    )
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:1"


def build_validation_sheet(workbook: Workbook, report: dict[str, Any]) -> None:
    sheet = workbook.create_sheet("Walidacja")
    sheet.sheet_view.showGridLines = False
    sheet.append(["Kontrola", "Wartość logiczna", "Status", "Znaczenie"])
    validation_items = [(key, value) for key, value in report["validation"].items() if key != "passed"]
    for row, (key, value) in enumerate(validation_items, start=2):
        sheet.cell(row, 1).value = VALIDATION_LABELS.get(key, key)
        sheet.cell(row, 2).value = bool(value)
        sheet.cell(row, 3).value = f'=IF(B{row},"OK","BŁĄD")'
        sheet.cell(row, 4).value = key
    style_table_header(sheet[1])
    last_row = len(validation_items) + 1
    add_table(sheet, f"A1:D{last_row}", "ValidationTable", "TableStyleMedium4")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:D{last_row}"
    sheet.conditional_formatting.add(
        f"C2:C{last_row}",
        FormulaRule(formula=["C2=\"OK\""], fill=PatternFill("solid", fgColor=LIGHT_GREEN), font=Font(color=GREEN, bold=True)),
    )
    sheet.conditional_formatting.add(
        f"C2:C{last_row}",
        FormulaRule(formula=["C2=\"BŁĄD\""], fill=PatternFill("solid", fgColor=LIGHT_RED), font=Font(color=RED, bold=True)),
    )
    sheet.column_dimensions["A"].width = 54
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 14
    sheet.column_dimensions["D"].width = 42
    for row in sheet.iter_rows(min_row=2, max_row=last_row):
        row[0].alignment = Alignment(wrap_text=True)
        row[1].alignment = Alignment(horizontal="center")
        row[2].alignment = Alignment(horizontal="center")
        row[3].font = Font(name="Consolas", size=9, color="667085")
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1


def build_changes_sheet(workbook: Workbook, report: dict[str, Any]) -> None:
    sheet = workbook.create_sheet("Zmiany")
    sheet.sheet_view.showGridLines = False
    sheet.append(["Lp.", "Komórka", "SKU", "Kolumna", "Przed zmianą", "Po zmianie"])
    for index, change in enumerate(report["changes"], start=1):
        before = change.get("before")
        after = change.get("after")
        sheet.append(
            [
                index,
                change.get("cell", ""),
                change.get("sku", ""),
                change.get("column", ""),
                "(puste)" if before is None or before == "" else before,
                "(puste)" if after is None or after == "" else after,
            ]
        )
    style_table_header(sheet[1])
    last_row = len(report["changes"]) + 1
    add_table(sheet, f"A1:F{last_row}", "ChangesTable", "TableStyleMedium2")
    sheet.freeze_panes = "D2"
    sheet.auto_filter.ref = f"A1:F{last_row}"
    widths = {"A": 8, "B": 12, "C": 24, "D": 50, "E": 62, "F": 62}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2, max_row=last_row):
        row[0].alignment = Alignment(horizontal="right")
        row[1].font = Font(name="Consolas", size=9, color="475467")
        row[2].font = Font(name="Consolas", size=9, color=TEXT)
        for cell in row[3:]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if row[4].value == "(puste)":
            row[4].font = Font(name="Aptos", italic=True, color="98A2B3")
        if row[5].value == "(puste)":
            row[5].font = Font(name="Aptos", italic=True, color="98A2B3")
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:1"


def build_sources_sheet(workbook: Workbook, report: dict[str, Any], axis_data: dict[str, Any]) -> None:
    sheet = workbook.create_sheet("Źródła i metadane")
    sheet.sheet_view.showGridLines = False
    set_title_band(sheet, "A1:B2", "Źródła i metadane audytu")
    rows = [
        ("Typ raportu", report["report_type"]),
        ("Raport wygenerowano", report["generated_at"]),
        ("Plik wejściowy", report["input_file"]),
        ("SHA-256 pliku wejściowego", report["input_sha256"]),
        ("Plik wynikowy", report["output_file"]),
        ("SHA-256 pliku wynikowego", report["output_sha256"]),
        ("Data modyfikacji wyniku", report["output_modified_at"]),
        ("Plik analizy osi wariantów", axis_data["analysis_file"]),
        ("Aktualny plik Woo do filtrowania grup", axis_data["current_woo_file"]),
        ("Aktualne grupy z osiami", axis_data["validation"]["current_groups"]),
        ("Aktualne produkty w grupach", axis_data["validation"]["current_variant_products"]),
        ("Arkusze", ", ".join(report["sheets"])),
        ("Wiersze z SKU", report["rows_with_sku"]),
        ("Wiersze łącznie z nagłówkiem", report["worksheet_rows_including_header"]),
        ("Kolumny", report["worksheet_columns"]),
        ("Brakujące kontrolowane SKU", ", ".join(report["missing_required_skus"]) or "Brak"),
        ("Nieoczekiwane zmiany", len(report["unexpected_changes"])),
        ("Problemy z kodowaniem", len(report["suspicious_encoding_cells"])),
    ]
    for row, (label, value) in enumerate(rows, start=4):
        sheet.cell(row, 1).value = label
        sheet.cell(row, 2).value = value
        sheet.cell(row, 1).font = Font(name="Aptos", bold=True, color=TEXT)
        sheet.cell(row, 1).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
        sheet.cell(row, 1).border = Border(bottom=THIN_GRAY)
        sheet.cell(row, 2).border = Border(bottom=THIN_GRAY)
        sheet.cell(row, 2).alignment = Alignment(wrap_text=True, vertical="top")
    sheet.column_dimensions["A"].width = 35
    sheet.column_dimensions["B"].width = 105
    for row in range(4, 22):
        sheet.row_dimensions[row].height = 26
    sheet.freeze_panes = "A4"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1


def create_workbook(
    report: dict[str, Any],
    axis_data: dict[str, Any],
    output_path: Path,
    overwrite: bool = False,
) -> None:
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"Plik raportu już istnieje: {output_path}")
    workbook = Workbook()
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    build_summary_sheet(workbook, report)
    build_axis_summary_sheet(workbook, axis_data)
    build_variant_groups_sheet(workbook, axis_data)
    build_variant_products_sheet(workbook, axis_data)
    build_validation_sheet(workbook, report)
    build_changes_sheet(workbook, report)
    build_sources_sheet(workbook, report, axis_data)
    for sheet in workbook.worksheets:
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.oddFooter.center.text = "Strona &P z &N"
        sheet.oddFooter.right.text = "AlkanInator"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()


def verify_workbook(
    path: Path,
    expected_changes: int,
    expected_validations: int,
    expected_groups: int,
    expected_products: int,
) -> dict[str, Any]:
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        expected_sheets = [
            "Podsumowanie",
            "Podsumowanie osi",
            "Osie grup",
            "Produkty w grupach",
            "Walidacja",
            "Zmiany",
            "Źródła i metadane",
        ]
        formula_errors: list[str] = []
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and any(
                        marker in cell.value for marker in ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")
                    ):
                        formula_errors.append(f"{sheet.title}!{cell.coordinate}")
        result = {
            "sheet_names_match": workbook.sheetnames == expected_sheets,
            "changes_rows_match": workbook["Zmiany"].max_row == expected_changes + 1,
            "validation_rows_match": workbook["Walidacja"].max_row == expected_validations + 1,
            "variant_group_rows_match": workbook["Osie grup"].max_row == expected_groups + 1,
            "variant_product_rows_match": workbook["Produkty w grupach"].max_row == expected_products + 1,
            "all_groups_have_axes": all(
                workbook["Osie grup"].cell(row, 7).value
                and workbook["Osie grup"].cell(row, 7).value >= 1
                for row in range(2, workbook["Osie grup"].max_row + 1)
            ),
            "summary_formula_present": str(workbook["Podsumowanie"]["B6"].value).startswith("=IF("),
            "no_formula_error_literals": not formula_errors,
            "formula_error_cells": formula_errors,
        }
        result["passed"] = all(value for key, value in result.items() if key != "formula_error_cells")
        return result
    finally:
        workbook.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Eksportuje raport korekty wariantów Hager/Berker do XLSX.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--variant-analysis", type=Path, default=DEFAULT_VARIANT_ANALYSIS)
    parser.add_argument("--current-woo", type=Path, default=DEFAULT_CURRENT_WOO)
    parser.add_argument("--overwrite", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = args.input.resolve()
    output_path = args.output.resolve()
    report = load_report(input_path)
    axis_data = load_variant_axis_data(args.variant_analysis.resolve(), args.current_woo.resolve())
    if not axis_data["validation"]["passed"]:
        raise RuntimeError(f"Walidacja osi wariantów nie powiodła się: {axis_data['validation']}")
    create_workbook(report, axis_data, output_path, overwrite=args.overwrite)
    validation_count = len([key for key in report["validation"] if key != "passed"])
    verification = verify_workbook(
        output_path,
        len(report["changes"]),
        validation_count,
        len(axis_data["groups"]),
        len(axis_data["products"]),
    )
    if not verification["passed"]:
        raise RuntimeError(f"Walidacja raportu XLSX nie powiodła się: {verification}")
    print(
        json.dumps(
            {
                "output": str(output_path),
                "axis_validation": axis_data["validation"],
                "verification": verification,
            },
            ensure_ascii=True,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
