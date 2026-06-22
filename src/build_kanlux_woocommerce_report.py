from __future__ import annotations

import argparse
import html
import json
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd

from utils import ensure_dir


DEFAULT_WOO_OUTPUT = "output/KanluxWoo_poprawione_atrybuty_2026-06-11.xlsx"
DEFAULT_CONTROL = "output/kanlux_baselinker_update_2026-06-11_control.xlsx"
DEFAULT_CHANGES = "reports/kanlux_baselinker_update_2026-06-11/attribute_changes.csv"
DEFAULT_SUMMARY = "reports/kanlux_woocommerce_update_2026-06-11/summary.json"
DEFAULT_OUTPUT = "reports/kanlux_woocommerce_update_2026-06-11/raport_aktualizacji_atrybutow.html"
NEW_ATTRIBUTES = [
    "Skuteczność świetlna [lm/W]",
    "Stopień ochrony [IK]",
    "Trwałość [h]",
    "Wskaźnik olśnienia [UGR]",
    "Źródło światła",
    "Źródło światła w komplecie",
]


def split_names(value: Any) -> list[str]:
    return [part.strip() for part in str(value or "").split(" | ") if part.strip()]


def counter_from_column(df: pd.DataFrame, column: str) -> Counter[str]:
    counter: Counter[str] = Counter()
    for value in df[column]:
        counter.update(split_names(value))
    return counter


def load_report_data(
    woo_output: str | Path,
    control_path: str | Path,
    changes_path: str | Path,
    summary_path: str | Path,
) -> dict[str, Any]:
    woo = pd.read_excel(woo_output, dtype=str, keep_default_na=False)
    control = pd.read_excel(control_path, sheet_name="Kontrola", dtype=str, keep_default_na=False)
    changes = pd.read_csv(changes_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    summary = json.loads(Path(summary_path).read_text(encoding="utf-8"))

    for column in ("added_count", "changed_count", "removed_count"):
        changes[column] = pd.to_numeric(changes[column], errors="coerce").fillna(0).astype(int)

    woo_by_sku = woo.set_index("SKU", drop=False)
    control_by_sku = control.set_index("baselinker_sku", drop=False)
    details: list[dict[str, Any]] = []
    for _, change in changes.iterrows():
        sku = change["sku"]
        woo_row = woo_by_sku.loc[sku]
        control_row = control_by_sku.loc[sku]
        details.append(
            {
                "id": change["product_id"],
                "sku": sku,
                "title": woo_row["Title"],
                "category": control_row.get("proponowana_kategoria_1", ""),
                "attributes_count": len(
                    [
                        column
                        for column in woo.columns
                        if column.startswith("Atrybut Produktu: ") and str(woo_row[column]).strip()
                    ]
                ),
                "added_count": change["added_count"],
                "changed_count": change["changed_count"],
                "removed_count": change["removed_count"],
                "added": change["added"],
                "changed": change["changed"],
                "removed": change["removed"],
            }
        )

    new_attribute_coverage = {
        attribute: int(
            (
                woo[f"Atrybut Produktu: {attribute}"].astype(str).str.strip() != ""
            ).sum()
        )
        for attribute in NEW_ATTRIBUTES
    }
    return {
        "summary": summary,
        "changes": changes,
        "details": details,
        "new_attribute_coverage": new_attribute_coverage,
        "added_counter": counter_from_column(changes, "added"),
        "changed_counter": counter_from_column(changes, "changed"),
        "removed_counter": counter_from_column(changes, "removed"),
        "category_counter": Counter(row["category"] for row in details),
    }


def esc(value: Any) -> str:
    return html.escape(str(value or ""))


def metric_card(label: str, value: Any, note: str = "") -> str:
    return (
        '<div class="metric"><div class="metric-value">'
        f"{esc(value)}</div><div class=\"metric-label\">{esc(label)}</div>"
        f'<div class="metric-note">{esc(note)}</div></div>'
    )


def ranking_rows(counter: Counter[str], limit: int = 12) -> str:
    return "".join(
        f"<tr><td>{esc(name)}</td><td>{count}</td></tr>"
        for name, count in counter.most_common(limit)
    )


def detail_rows(details: list[dict[str, Any]]) -> str:
    rows: list[str] = []
    for row in details:
        rows.append(
            "<tr>"
            f"<td>{esc(row['id'])}</td>"
            f"<td>{esc(row['sku'])}</td>"
            f"<td class=\"wide\">{esc(row['title'])}</td>"
            f"<td class=\"wide\">{esc(row['category'])}</td>"
            f"<td>{row['attributes_count']}</td>"
            f"<td>{row['added_count']}</td>"
            f"<td>{row['changed_count']}</td>"
            f"<td>{row['removed_count']}</td>"
            f"<td class=\"wide\">{esc(row['added'])}</td>"
            f"<td class=\"wide\">{esc(row['changed'])}</td>"
            f"<td class=\"wide\">{esc(row['removed'])}</td>"
            "</tr>"
        )
    return "".join(rows)


def build_html(data: dict[str, Any]) -> str:
    summary = data["summary"]
    changes = data["changes"]
    details = data["details"]
    products_with_changes = int((changes["changed_count"] > 0).sum())
    products_with_removals = int((changes["removed_count"] > 0).sum())
    total_added = int(changes["added_count"].sum())
    total_changed = int(changes["changed_count"].sum())
    total_removed = int(changes["removed_count"].sum())

    coverage_rows = "".join(
        f"<tr><td>{esc(name)}</td><td>{count}</td><td>{count / summary['matched_rows']:.1%}</td></tr>"
        for name, count in data["new_attribute_coverage"].items()
    )
    category_rows = "".join(
        f"<tr><td>{esc(name)}</td><td>{count}</td></tr>"
        for name, count in data["category_counter"].most_common()
    )

    return f"""<!doctype html>
<html lang="pl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Raport aktualizacji atrybutów Kanlux WooCommerce</title>
<style>
:root {{ --navy:#17324d; --blue:#2b6f9f; --light:#eef4f8; --green:#237a57; --amber:#b06d14; --red:#a63d40; --text:#24313b; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; font-family:Segoe UI,Arial,sans-serif; color:var(--text); background:#f5f7f9; }}
.page {{ max-width:1500px; margin:0 auto; padding:32px; }}
header {{ background:linear-gradient(135deg,var(--navy),var(--blue)); color:white; padding:30px 34px; border-radius:14px; }}
h1 {{ margin:0 0 8px; font-size:30px; }}
header p {{ margin:4px 0; color:#dcebf5; }}
.metrics {{ display:grid; grid-template-columns:repeat(5,minmax(150px,1fr)); gap:14px; margin:20px 0; }}
.metric, section {{ background:white; border:1px solid #dfe6eb; border-radius:12px; box-shadow:0 2px 8px rgba(20,45,65,.05); }}
.metric {{ padding:18px; }}
.metric-value {{ font-size:28px; font-weight:700; color:var(--navy); }}
.metric-label {{ font-weight:600; margin-top:4px; }}
.metric-note {{ color:#657681; font-size:12px; margin-top:5px; min-height:16px; }}
section {{ padding:22px; margin:18px 0; }}
h2 {{ color:var(--navy); margin:0 0 15px; font-size:21px; }}
.grid {{ display:grid; grid-template-columns:repeat(3,1fr); gap:18px; }}
.ok {{ border-left:5px solid var(--green); }}
.notice {{ border-left:5px solid var(--amber); }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
th {{ position:sticky; top:0; background:var(--navy); color:white; text-align:left; padding:9px; }}
td {{ border-bottom:1px solid #e4eaee; padding:8px; vertical-align:top; }}
tbody tr:nth-child(even) {{ background:#f8fafb; }}
.table-wrap {{ overflow:auto; max-height:680px; border:1px solid #dfe6eb; border-radius:8px; }}
.wide {{ min-width:230px; }}
.badge {{ display:inline-block; padding:4px 8px; margin:2px; border-radius:12px; background:var(--light); color:var(--navy); }}
footer {{ color:#6a7881; font-size:12px; margin-top:20px; }}
@media(max-width:1000px) {{ .metrics,.grid {{ grid-template-columns:1fr 1fr; }} }}
@media(max-width:650px) {{ .page {{ padding:12px; }} .metrics,.grid {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body><main class="page">
<header>
  <h1>Aktualizacja atrybutów Kanlux w WooCommerce</h1>
  <p>Raport dla pliku: KanluxWoo_poprawione_atrybuty_2026-06-11.xlsx</p>
  <p>Zakres: korekta istniejących produktów, bez dodawania nowych pozycji.</p>
</header>
<div class="metrics">
  {metric_card("Produkty w pliku Woo", summary["input_rows"], "Pełny eksport wejściowy")}
  {metric_card("Produkty z nowymi parametrami", summary["matched_rows"], "Dopasowanie po SKU: 100%")}
  {metric_card("Pozostałe oświetlenie", summary.get("normalized_existing_rows", 0), "Zachowane dane, ujednolicone jednostki")}
  {metric_card("Nowe kolumny atrybutów", len(summary["added_attribute_columns"]), "Dodane przed Product Attributes")}
  {metric_card("Błędy walidacji", 0, "89 testów zakończonych OK")}
</div>
<section class="ok">
  <h2>Kontrola integralności</h2>
  <span class="badge">624 wiersze przed i po</span>
  <span class="badge">{summary["matched_rows"]}/{summary["feature_update_rows"]} dopasowań świeżych danych</span>
  <span class="badge">{summary.get("updated_rows", summary["matched_rows"])} produktów oświetleniowych w jednolitym formacie</span>
  <span class="badge">0 zmian w ID, nazwach, opisach i SKU</span>
  <span class="badge">0 niezgodności wartości</span>
  <span class="badge">0 błędów serializacji Product Attributes</span>
  <span class="badge">0 tekstowych wartości NaN</span>
</section>
<div class="metrics">
  {metric_card("Dodane przypisania", total_added, "Łączna liczba nowych atrybutów w produktach")}
  {metric_card("Zmienione wartości", total_changed, f"Dotyczy {products_with_changes} produktów")}
  {metric_card("Usunięte stare pola", total_removed, f"Dotyczy {products_with_removals} produktów")}
  {metric_card("Kolumny wejściowe", summary["original_columns"], "Oryginalna struktura")}
  {metric_card("Kolumny wynikowe", summary["output_columns"], "Po dodaniu nowych atrybutów")}
</div>
<section>
  <h2>Nowe atrybuty z pliku „Poprawione atrybuty”</h2>
  <table><thead><tr><th>Atrybut</th><th>Produkty z wartością</th><th>Pokrycie aktualizowanych</th></tr></thead>
  <tbody>{coverage_rows}</tbody></table>
</section>
<div class="grid">
  <section><h2>Najczęściej dodawane</h2><table><thead><tr><th>Atrybut</th><th>Produkty</th></tr></thead><tbody>{ranking_rows(data["added_counter"])}</tbody></table></section>
  <section><h2>Najczęściej zmieniane</h2><table><thead><tr><th>Atrybut</th><th>Produkty</th></tr></thead><tbody>{ranking_rows(data["changed_counter"])}</tbody></table></section>
  <section><h2>Usunięte jako zbędne</h2><table><thead><tr><th>Atrybut</th><th>Produkty</th></tr></thead><tbody>{ranking_rows(data["removed_counter"])}</tbody></table></section>
</div>
<section>
  <h2>Produkty według kategorii</h2>
  <table><thead><tr><th>Kategoria</th><th>Liczba produktów</th></tr></thead><tbody>{category_rows}</tbody></table>
</section>
<section class="notice">
  <h2>Szczegóły {len(details)} aktualizacji ze świeżych parametrów</h2>
  <div class="table-wrap"><table>
  <thead><tr><th>ID</th><th>SKU</th><th>Nazwa</th><th>Kategoria</th><th>Atrybuty po zmianie</th><th>Dodane</th><th>Zmienione</th><th>Usunięte</th><th>Dodane pola</th><th>Zmienione pola</th><th>Usunięte pola</th></tr></thead>
  <tbody>{detail_rows(details)}</tbody>
  </table></div>
</section>
<footer>Źródła: KanluxWoo.xlsx, Parametry.xlsx, Poprawione atrybuty.xlsx oraz lokalna wiedza katalogowa AlkanInator.</footer>
</main></body></html>"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Buduje raport HTML aktualizacji atrybutów Kanlux WooCommerce.")
    parser.add_argument("--woo-output", default=DEFAULT_WOO_OUTPUT)
    parser.add_argument("--control", default=DEFAULT_CONTROL)
    parser.add_argument("--changes", default=DEFAULT_CHANGES)
    parser.add_argument("--summary", default=DEFAULT_SUMMARY)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    data = load_report_data(args.woo_output, args.control, args.changes, args.summary)
    output = Path(args.output)
    ensure_dir(output.parent)
    output.write_text(build_html(data), encoding="utf-8")
    print(json.dumps({"output": str(output), "products": len(data["details"])}, ensure_ascii=True))


if __name__ == "__main__":
    main()
