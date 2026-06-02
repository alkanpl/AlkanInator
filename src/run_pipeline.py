from __future__ import annotations

import argparse
from pathlib import Path

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, enrich_config_with_catalog_knowledge, load_catalog_knowledge
from analyze_products import detect_product_columns
from analyze_keywords import add_keyword_pipeline_arguments
from extract_attributes import extract_attributes_for_dataframe
from pipeline_core import finalize_and_export, write_input_analysis_reports
from title_anatomy import DEFAULT_TITLE_ANATOMY_PATH, load_title_anatomy_config
from utils import ensure_dir, load_yaml, read_products, slugify


def resolve_config(category: str, config_path: str | None) -> Path:
    if config_path:
        return Path(config_path)
    return Path("configs") / "categories" / f"{slugify(category)}.yaml"


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

    write_input_analysis_reports(df, columns, config, reports_dir)

    working_df, working_title_column = add_working_title_column(df, title_column, config)
    extracted_df = extract_attributes_for_dataframe(working_df, working_title_column, config, columns.get("producer"))
    finalize_and_export(extracted_df, columns, config, anatomy_config, args, reports_dir, working_title_column)

    print(f"OK: zapisano {args.output}")
    print(f"OK: zapisano raporty w {reports_dir}")


if __name__ == "__main__":
    main()
