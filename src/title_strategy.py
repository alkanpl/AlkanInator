from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd

from utils import ensure_dir, read_products


def build_title_strategy_report(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    group_columns = ["Kategoria", "Typ", "product_role", "accessory_type"]
    for column in group_columns:
        if column not in df.columns:
            df[column] = ""

    rows: list[dict[str, Any]] = []
    examples: list[dict[str, Any]] = []

    grouped = df.groupby(group_columns, dropna=False)
    for key, group in grouped:
        category, product_type, role, accessory_type = key
        statuses = group["title_status"].value_counts().to_dict() if "title_status" in group.columns else {}
        category_statuses = (
            group["category_fit_status"].value_counts().to_dict() if "category_fit_status" in group.columns else {}
        )
        sample = group.head(3)
        rows.append(
            {
                "category": category,
                "type": product_type,
                "product_role": role,
                "accessory_type": accessory_type,
                "products_count": len(group),
                "title_ok": statuses.get("OK", 0),
                "title_warning": statuses.get("WARNING", 0),
                "category_ok": category_statuses.get("OK", 0),
                "category_to_check": category_statuses.get("DO_SPRAWDZENIA", 0),
                "recommended_strategy": recommend_strategy(str(category), str(product_type), str(role), str(accessory_type)),
                "sample_old_titles": " || ".join(sample.get("old_title", pd.Series(dtype=str)).astype(str).tolist()),
                "sample_new_titles": " || ".join(sample.get("new_title", pd.Series(dtype=str)).astype(str).tolist()),
            }
        )

        for _, item in sample.iterrows():
            examples.append(
                {
                    "category": category,
                    "type": product_type,
                    "product_role": role,
                    "accessory_type": accessory_type,
                    "sku": item.get("Kod", item.get("sku", "")),
                    "ean": item.get("EAN", item.get("ean", "")),
                    "old_title": item.get("old_title", ""),
                    "new_title": item.get("new_title", ""),
                    "title_status": item.get("title_status", ""),
                    "category_fit_status": item.get("category_fit_status", ""),
                    "warnings": item.get("warnings", ""),
                    "category_fit_reasons": item.get("category_fit_reasons", ""),
                }
            )

    summary = pd.DataFrame(rows).sort_values(["products_count", "category", "type"], ascending=[False, True, True])
    examples_df = pd.DataFrame(examples)
    return summary, examples_df


def recommend_strategy(category: str, product_type: str, role: str, accessory_type: str) -> str:
    text = f"{category} {product_type} {accessory_type}".lower()
    if role == "accessory":
        return f"accessory:{accessory_type or 'Akcesorium'}"
    if "panel" in text:
        return "main:Panel LED + seria + moc + strumien + barwa + ksztalt/wymiary + kolor + producent"
    if "naswietlacz" in normalize_ascii(text) or "naświetlacz" in text:
        return "main:Naświetlacz LED + seria + moc + strumien + barwa + ip + czujnik + kolor + producent"
    if "high bay" in text:
        return "main:Oprawa High Bay LED + seria + moc + strumien + barwa + ip + producent"
    if "hermetycz" in normalize_ascii(text) or "pyloszczel" in normalize_ascii(text):
        return "main:Oprawa hermetyczna LED + seria + moc + strumien + barwa + ip + kolor + producent"
    if "plafon" in text:
        return "main:Plafon LED + seria + moc + barwa + ip + kolor/material + czujnik + producent"
    return "main:typ + seria + moc + parametry zakupowe + producent"


def normalize_ascii(value: str) -> str:
    return (
        value.replace("ą", "a")
        .replace("ć", "c")
        .replace("ę", "e")
        .replace("ł", "l")
        .replace("ń", "n")
        .replace("ó", "o")
        .replace("ś", "s")
        .replace("ź", "z")
        .replace("ż", "z")
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Buduje raport strategii tytulow SEO per typ produktu.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--reports-dir", default="reports")
    args = parser.parse_args()

    df = read_products(args.input)
    reports_dir = ensure_dir(args.reports_dir)
    summary, examples = build_title_strategy_report(df)
    summary.to_csv(Path(reports_dir) / "title_strategy_report.csv", index=False, encoding="utf-8-sig")
    examples.to_csv(Path(reports_dir) / "title_strategy_examples.csv", index=False, encoding="utf-8-sig")
    print(f"OK: zapisano {Path(reports_dir) / 'title_strategy_report.csv'}")
    print(f"OK: zapisano {Path(reports_dir) / 'title_strategy_examples.csv'}")


if __name__ == "__main__":
    main()
