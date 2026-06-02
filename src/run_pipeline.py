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
from analyze_keywords import add_keyword_pipeline_arguments, maybe_enrich_pipeline_with_keywords
from export_results import export_outputs
from extract_attributes import extract_attributes_for_dataframe
from optimize_titles import build_changes_report, optimize_titles_for_dataframe
from title_anatomy import DEFAULT_TITLE_ANATOMY_PATH, load_title_anatomy_config, write_title_anatomy_reports
from title_strategy import build_title_strategy_report
from utils import ensure_dir, load_yaml, read_products, slugify
from validate_output import classify_product_roles, validate_category_fit


def resolve_config(category: str, config_path: str | None) -> Path:
    if config_path:
        return Path(config_path)
    return Path("configs") / "categories" / f"{slugify(category)}.yaml"


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


def add_working_title_column(df, title_column: str, config: dict) -> tuple[object, str]:
    result = df.copy()
    fallback_columns = [column for column in config.get("title_fallback_columns", []) if column in result.columns]
    working = result[title_column].astype(str)
    for column in fallback_columns:
        fallback = result[column].astype(str)
        working = working.where(working.str.strip().ne(""), fallback)
    result["__working_title"] = working
    return result, "__working_title"


def main() -> None:
    parser = argparse.ArgumentParser(description="Uruchamia MVP pipeline dla produktow.")
    parser.add_argument("--input", required=True, help="Plik CSV/XLSX z produktami.")
    parser.add_argument("--sheet", help="Nazwa arkusza XLSX, np. '7. Wszystkie SKU (master)'.")
    parser.add_argument("--category", default="Oprawy sufitowe", help="Nazwa kategorii do wyboru configu.")
    parser.add_argument("--config", help="Sciezka do configu YAML kategorii.")
    parser.add_argument(
        "--catalog-knowledge",
        default=DEFAULT_CATALOG_KNOWLEDGE_PATH,
        help="Sciezka do slownika wiedzy z eksportu WooCommerce. Uzyj pustej wartosci, aby pominac.",
    )
    parser.add_argument("--output", default="output/products_optimized.xlsx", help="Docelowy plik XLSX/CSV.")
    parser.add_argument("--reports-dir", default="reports", help="Katalog raportow.")
    parser.add_argument("--title-anatomy", default=DEFAULT_TITLE_ANATOMY_PATH, help="Globalny config anatomii tytulow.")
    add_keyword_pipeline_arguments(parser)
    args = parser.parse_args()

    config_path = resolve_config(args.category, args.config)
    config = load_yaml(config_path)
    if not config:
        raise SystemExit(f"Brak configu kategorii: {config_path}")
    config = enrich_config_with_catalog_knowledge(config, load_catalog_knowledge(args.catalog_knowledge))
    anatomy_config = load_title_anatomy_config(args.title_anatomy)

    reports_dir = ensure_dir(args.reports_dir)
    df = read_products(args.input, sheet_name=args.sheet)
    columns = detect_product_columns(df)
    title_column = columns.get("title")
    if not title_column:
        raise SystemExit("Nie wykryto kolumny tytulu/nazwy produktu. Podaj kolumne o nazwie np. title, name, nazwa.")

    analyzed_summary = analyze_dataframe(df, columns, config)
    write_quality_report(analyzed_summary, reports_dir / "quality_report.html")
    build_category_summary(df, columns).to_csv(reports_dir / "category_summary.csv", index=False, encoding="utf-8-sig")

    working_df, working_title_column = add_working_title_column(df, title_column, config)
    extracted_df = extract_attributes_for_dataframe(working_df, working_title_column, config, columns.get("producer"))
    classified_df = classify_product_roles(extracted_df, config)
    classified_df = maybe_enrich_pipeline_with_keywords(classified_df, args, reports_dir)
    optimized_df = optimize_titles_for_dataframe(classified_df, working_title_column, config, anatomy_config)
    optimized_df, suspicious_category_df = validate_category_fit(optimized_df, config)
    write_title_anatomy_reports(optimized_df, reports_dir, anatomy_config)
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


if __name__ == "__main__":
    main()
