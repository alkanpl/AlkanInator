from __future__ import annotations

import argparse

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, enrich_config_with_catalog_knowledge, load_catalog_knowledge
from analyze_products import detect_product_columns
from analyze_keywords import add_keyword_pipeline_arguments
from enrich_from_parameters import (
    build_parameter_index,
    detect_parameter_columns,
    enrich_from_parameters,
    write_reports as write_parameter_reports,
)
from pipeline_core import finalize_and_export, write_input_analysis_reports
from title_anatomy import DEFAULT_TITLE_ANATOMY_PATH, load_title_anatomy_config
from utils import ensure_dir, load_yaml, read_products


def main() -> None:
    parser = argparse.ArgumentParser(description="Generuje nazwy SEO po uzupelnieniu atrybutow z Parametry.xlsx.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet")
    parser.add_argument("--parameters", required=True)
    parser.add_argument("--parameters-sheet", default=0)
    parser.add_argument("--config", default="configs/categories/kanlux-oswietlenie.yaml")
    parser.add_argument(
        "--catalog-knowledge",
        default=DEFAULT_CATALOG_KNOWLEDGE_PATH,
        help="Sciezka do slownika wiedzy z eksportu WooCommerce. Uzyj pustej wartosci, aby pominac.",
    )
    parser.add_argument("--output", default="output/products_optimized_with_parameters.xlsx")
    parser.add_argument("--reports-dir", default="reports/optimized_with_parameters")
    parser.add_argument("--title-anatomy", default=DEFAULT_TITLE_ANATOMY_PATH, help="Globalny config anatomii tytulow.")
    add_keyword_pipeline_arguments(parser)
    args = parser.parse_args()

    config = enrich_config_with_catalog_knowledge(load_yaml(args.config), load_catalog_knowledge(args.catalog_knowledge))
    anatomy_config = load_title_anatomy_config(args.title_anatomy)
    reports_dir = ensure_dir(args.reports_dir)
    parameter_reports_dir = ensure_dir(reports_dir / "parameters")

    products = read_products(args.input, sheet_name=args.sheet)
    params = read_products(args.parameters, sheet_name=args.parameters_sheet)
    columns = detect_product_columns(products)
    title_column = columns.get("title")
    if not title_column:
        raise SystemExit("Nie wykryto kolumny tytulu/nazwy produktu.")

    write_input_analysis_reports(products, columns, config, reports_dir)

    parameter_columns = detect_parameter_columns(params)
    parameter_index, unmapped = build_parameter_index(params, parameter_columns)
    enriched_df, parameter_reports = enrich_from_parameters(products, columns, config, parameter_index)
    write_parameter_reports(parameter_reports, unmapped, products, str(columns["sku"]), parameter_index, parameter_reports_dir)

    finalize_and_export(enriched_df, columns, config, anatomy_config, args, reports_dir)

    print(f"OK: zapisano {args.output}")
    print(f"OK: zapisano raporty w {reports_dir}")
    print(f"Dopasowane produkty z parametrami: {len(parameter_reports['matches'])}")
    print(f"Uzupelnione atrybuty: {len(parameter_reports['filled'])}")
    print(f"Konflikty parametrow: {len(parameter_reports['conflicts'])}")


if __name__ == "__main__":
    main()
