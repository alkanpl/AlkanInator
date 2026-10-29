"""Zrzuca slownik wartosci atrybutow Woo (terminy pa_*) z bazy sklepu do TSV.

Wynik (``dictionaries/woo_attribute_terms_snapshot.tsv``) jest zrodlem prawdy
dla ``src/shop_attribute_values.py``: kazda wartosc atrybutu spoza tej listy
utworzy w Woo NOWY termin. Kolumna ``parent_term`` to powiazanie filtra
(``ext_pf_cross_parent_term_id``), np. Podtyp "Antenowe" -> Typ "Gniazda".

Baza lokalnej kopii (AlkanLocal) bywa starsza niz sklep produkcyjny - terminy
dodane pozniej uzupelnia sie przy kontroli eksportem sklepu (``--shop-export``).

Uzycie:
    py scripts/dump_woo_attribute_terms.py
    py scripts/dump_woo_attribute_terms.py --database AlkanLocal --output dictionaries/woo_attribute_terms_snapshot.tsv
"""

from __future__ import annotations

import argparse
import csv
import glob
import subprocess
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "dictionaries" / "woo_attribute_terms_snapshot.tsv"
MYSQL_GLOB = "C:/laragon/bin/mysql/*/bin/mysql.exe"

QUERY = """
SELECT a.attribute_label, CONCAT('pa_', a.attribute_name), IFNULL(t.name, ''), IFNULL(t.slug, ''),
       IFNULL(tt.count, 0), IFNULL(p.name, '')
FROM wp_woocommerce_attribute_taxonomies a
LEFT JOIN wp_term_taxonomy tt ON tt.taxonomy = CONCAT('pa_', a.attribute_name)
LEFT JOIN wp_terms t ON t.term_id = tt.term_id
LEFT JOIN wp_termmeta tm ON tm.term_id = t.term_id AND tm.meta_key = 'ext_pf_cross_parent_term_id'
LEFT JOIN wp_terms p ON p.term_id = tm.meta_value
ORDER BY a.attribute_label, t.name
"""

COLUMNS = ["attribute", "taxonomy", "term", "slug", "count", "parent_term"]


def find_mysql() -> str:
    candidates = sorted(glob.glob(MYSQL_GLOB))
    if not candidates:
        raise SystemExit(f"Nie znaleziono klienta mysql ({MYSQL_GLOB}); podaj --mysql")
    return candidates[0]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--mysql", default=None, help="Sciezka do mysql.exe (domyslnie Laragon)")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--user", default="root")
    parser.add_argument("--password", default="")
    parser.add_argument("--database", default="AlkanLocal")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args(argv)

    command = [args.mysql or find_mysql(), f"-u{args.user}", f"-h{args.host}", "--default-character-set=utf8mb4",
               "--batch", "--raw", "--skip-column-names", args.database, "-e", QUERY]
    if args.password:
        command.insert(2, f"-p{args.password}")
    completed = subprocess.run(command, capture_output=True, check=True)
    lines = completed.stdout.decode("utf-8").splitlines()
    rows = [line.split("\t") for line in lines if line.strip()]

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8", newline="") as stream:
        stream.write(f"# Terminy atrybutow Woo z bazy {args.database} ({args.host}), zrzut {date.today().isoformat()}.\n")
        stream.write("# Generuje: py scripts/dump_woo_attribute_terms.py - nie edytowac recznie.\n")
        writer = csv.writer(stream, delimiter="\t", lineterminator="\n")
        writer.writerow(COLUMNS)
        writer.writerows(rows)
    attributes = {row[0] for row in rows}
    terms = sum(1 for row in rows if len(row) > 2 and row[2])
    print(f"Atrybutow: {len(attributes)}, terminow: {terms} -> {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
