from __future__ import annotations

import argparse
from pathlib import Path

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, enrich_config_with_catalog_knowledge, load_catalog_knowledge
from analyze_products import (
    analyze_dataframe,
    build_category_summary,
    build_missing_attributes,
    detect_product_columns,
    write_quality_report,
)
from enrich_from_parameters import (
    build_parameter_index,
    detect_parameter_columns,
    enrich_from_parameters,
    write_reports as write_parameter_reports,
)
from export_results import export_outputs
from optimize_titles import build_changes_report, optimize_titles_for_dataframe
from title_strategy import build_title_strategy_report
from utils import ensure_dir, load_yaml, read_products
from validate_output import classify_product_roles, validate_category_fit


def add_standard_report_columns(df, columns):
    result = df.copy()
    mapping = {
        "sku": columns.get("sku"),
        "ean": columns.get("ean"),
        "producer": columns.get("producer"),
        "category": columns.get("category"),
    }
    for target, source in mapping.items():
        result[target] = result[source] if source and source in result.columns else ""
    return result


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
    args = parser.parse_args()

    config = enrich_config_with_catalog_knowledge(load_yaml(args.config), load_catalog_knowledge(args.catalog_knowledge))
    reports_dir = ensure_dir(args.reports_dir)
    parameter_reports_dir = ensure_dir(reports_dir / "parameters")

    products = read_products(args.input, sheet_name=args.sheet)
    params = read_products(args.parameters, sheet_name=args.parameters_sheet)
    columns = detect_product_columns(products)
    title_column = columns.get("title")
    if not title_column:
        raise SystemExit("Nie wykryto kolumny tytulu/nazwy produktu.")

    analyzed_summary = analyze_dataframe(products, columns, config)
    write_quality_report(analyzed_summary, reports_dir / "quality_report.html")
    build_category_summary(products, columns).to_csv(reports_dir / "category_summary.csv", index=False, encoding="utf-8-sig")

    parameter_columns = detect_parameter_columns(params)
    parameter_index, unmapped = build_parameter_index(params, parameter_columns)
    enriched_df, parameter_reports = enrich_from_parameters(products, columns, config, parameter_index)
    write_parameter_reports(parameter_reports, unmapped, products, str(columns["sku"]), parameter_index, parameter_reports_dir)

    classified_df = classify_product_roles(enriched_df, config)
    optimized_df = optimize_titles_for_dataframe(classified_df, "__working_title", config)
    optimized_df, suspicious_category_df = validate_category_fit(optimized_df, config)

    missing_df = build_missing_attributes(optimized_df, columns, config)
    missing_df.to_csv(reports_dir / "missing_attributes.csv", index=False, encoding="utf-8-sig")
    suspicious_category_df.to_csv(reports_dir / "suspicious_category_fit.csv", index=False, encoding="utf-8-sig")

    changes_df = build_changes_report(add_standard_report_columns(optimized_df, columns))
    title_strategy_df, title_strategy_examples_df = build_title_strategy_report(optimized_df)
    title_strategy_df.to_csv(reports_dir / "title_strategy_report.csv", index=False, encoding="utf-8-sig")
    title_strategy_examples_df.to_csv(reports_dir / "title_strategy_examples.csv", index=False, encoding="utf-8-sig")
    export_outputs(optimized_df, changes_df, args.output, reports_dir)

    print(f"OK: zapisano {args.output}")
    print(f"OK: zapisano raporty w {reports_dir}")
    print(f"Dopasowane produkty z parametrami: {len(parameter_reports['matches'])}")
    print(f"Uzupelnione atrybuty: {len(parameter_reports['filled'])}")
    print(f"Konflikty parametrow: {len(parameter_reports['conflicts'])}")


if __name__ == "__main__":
    main()
