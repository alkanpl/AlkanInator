"""Wspolny rdzen pipeline'u uzywany przez run_pipeline.py i
run_pipeline_with_parameters.py.

Oba skrypty roznia sie tylko sposobem przygotowania kolumny roboczej tytulu
(``__working_title``): jeden ekstrahuje atrybuty z tytulu, drugi uzupelnia je z
pliku parametrow. Cala reszta - analiza wejscia, klasyfikacja rol, slowa
kluczowe, generowanie tytulow, walidacja kategorii i eksport raportow - jest
identyczna i mieszka tutaj, zeby nie utrzymywac dwoch kopii.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from analyze_keywords import maybe_enrich_pipeline_with_keywords
from analyze_products import (
    analyze_dataframe,
    build_category_summary,
    build_missing_attributes,
    write_quality_report,
)
from export_results import export_outputs
from optimize_titles import build_changes_report, optimize_titles_for_dataframe
from title_anatomy import write_title_anatomy_reports
from title_strategy import build_title_strategy_report
from validate_output import classify_product_roles, validate_category_fit


def add_standard_report_columns(df: pd.DataFrame, columns: dict[str, str | None]) -> pd.DataFrame:
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


def write_input_analysis_reports(
    df: pd.DataFrame,
    columns: dict[str, str | None],
    config: dict[str, Any],
    reports_dir: Path,
) -> None:
    summary = analyze_dataframe(df, columns, config)
    write_quality_report(summary, reports_dir / "quality_report.html")
    build_category_summary(df, columns).to_csv(
        reports_dir / "category_summary.csv", index=False, encoding="utf-8-sig"
    )


def finalize_and_export(
    working_df: pd.DataFrame,
    columns: dict[str, str | None],
    config: dict[str, Any],
    anatomy_config: dict[str, Any],
    args: Any,
    reports_dir: Path,
    working_title_column: str = "__working_title",
) -> pd.DataFrame:
    classified_df = classify_product_roles(working_df, config)
    classified_df = maybe_enrich_pipeline_with_keywords(classified_df, args, reports_dir)
    optimized_df = optimize_titles_for_dataframe(
        classified_df, working_title_column, config, anatomy_config
    )
    optimized_df, suspicious_category_df = validate_category_fit(optimized_df, config)
    write_title_anatomy_reports(optimized_df, reports_dir, anatomy_config)

    build_missing_attributes(optimized_df, columns, config).to_csv(
        reports_dir / "missing_attributes.csv", index=False, encoding="utf-8-sig"
    )
    suspicious_category_df.to_csv(
        reports_dir / "suspicious_category_fit.csv", index=False, encoding="utf-8-sig"
    )

    changes_df = build_changes_report(add_standard_report_columns(optimized_df, columns))
    title_strategy_df, title_strategy_examples_df = build_title_strategy_report(optimized_df)
    title_strategy_df.to_csv(
        reports_dir / "title_strategy_report.csv", index=False, encoding="utf-8-sig"
    )
    title_strategy_examples_df.to_csv(
        reports_dir / "title_strategy_examples.csv", index=False, encoding="utf-8-sig"
    )
    export_outputs(optimized_df, changes_df, args.output, reports_dir)
    return optimized_df
