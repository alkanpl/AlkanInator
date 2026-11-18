"""Raport terminow atrybutow Woo do sprzatniecia w sklepie (tylko odczyt zrzutu).

Zrzut ``dictionaries/woo_attribute_terms_snapshot.tsv`` idzie od 2026-11-18 z
produkcji (``py src/woo_api_cli.py --profile prod knowledge``), a produkcja ma
terminy po bledach importow. ``shop_attribute_values.load_vocabulary`` pomija je
przy kontroli wartosci; ten skrypt pokazuje je w jednym skoroszycie:

- "Smieci (pomijane)" - terminy, ktore kontrola wartosci odrzuca (klucze
  ``value_map``, ``shop_junk_terms``, sklejki z przecinkiem), z licznikiem produktow,
- "Bez produktow" - wszystkie terminy z licznikiem 0 (kandydaci do usuniecia w Woo,
  do decyzji uzytkownika - nic nie jest usuwane automatycznie).

Uzycie:
    py scripts/report_shop_junk_terms.py
    py scripts/report_shop_junk_terms.py --snapshot cache/woo_api/prod/woo_attribute_terms_snapshot.tsv
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from shop_attribute_values import DEFAULT_RULES, DEFAULT_SNAPSHOT, load_rules, load_vocabulary  # noqa: E402


def unused_terms(snapshot: Path) -> list[dict[str, str]]:
    with snapshot.open("r", encoding="utf-8", newline="") as stream:
        rows = csv.DictReader((line for line in stream if not line.startswith("#")), delimiter="\t")
        return [row for row in rows if row.get("term") and (row.get("count") or "").strip() == "0"]


def write_report(path: Path, junk: list[dict[str, str]], unused: list[dict[str, str]]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Smieci (pomijane)"
    sheet.append(["Atrybut", "Termin", "Produktow w sklepie", "Powod"])
    for item in sorted(junk, key=lambda item: (item["atrybut"], item["termin"])):
        sheet.append([item["atrybut"], item["termin"], int(item["produktow"] or 0), item["powod"]])
    unused_sheet = workbook.create_sheet("Bez produktow")
    unused_sheet.append(["Atrybut", "Taksonomia", "Termin", "Slug", "Powiazanie filtra"])
    for row in unused:
        unused_sheet.append([row["attribute"], row["taxonomy"], row["term"], row["slug"], row.get("parent_term") or ""])
    for current, widths in ((sheet, (28, 60, 18, 50)), (unused_sheet, (28, 28, 60, 40, 24))):
        for cell in current[1]:
            cell.font = Font(bold=True)
        for column, width in zip("ABCDE", widths):
            current.column_dimensions[column].width = width
        current.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--snapshot", type=Path, default=DEFAULT_SNAPSHOT)
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES)
    parser.add_argument("--output", type=Path,
                        default=ROOT / "reports" / "woo_api" / f"terminy_do_sprzatniecia_{date.today():%Y-%m-%d}.xlsx")
    args = parser.parse_args()

    vocabulary = load_vocabulary(args.snapshot, load_rules(args.rules))
    unused = unused_terms(args.snapshot)
    write_report(args.output, vocabulary.junk, unused)
    in_use = sum(1 for item in vocabulary.junk if (item["produktow"] or "0") != "0")
    print(f"Smieci pomijane przy kontroli: {len(vocabulary.junk)} (w tym przypiete do produktow: {in_use}); "
          f"terminy bez produktow: {len(unused)}")
    print(f"Raport: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
