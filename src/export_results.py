from __future__ import annotations

from pathlib import Path

import pandas as pd

from utils import ensure_dir, write_products


def export_outputs(products_df: pd.DataFrame, changes_df: pd.DataFrame, output_path: str | Path, reports_dir: str | Path) -> None:
    output = Path(output_path)
    reports = ensure_dir(reports_dir)
    write_products(products_df, output)
    write_products(products_df, output.with_suffix(".csv"))
    write_products(changes_df, reports / "changes_report.xlsx")
    write_products(changes_df, reports / "changes_report.csv")
