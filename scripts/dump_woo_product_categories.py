"""Zrzuca drzewo kategorii produktow Woo (product_cat) z bazy do TSV.

Wynik (``dictionaries/woo_product_categories_snapshot.tsv``) sluzy
``src/build_product_removal_redirects.py`` do budowy adresow kategorii
(``/kategoria-produktu/<slug-rodzica>/<slug>/``) przy przekierowaniach
usuwanych produktow.

Baza lokalnej kopii (AlkanLocal) bywa starsza niz sklep produkcyjny - 2026-10-12
brakowalo w niej 29 z 97 kategorii, na ktore produkcja juz przekierowuje.
Przy usuwaniu produktow lepiej podac eksport kategorii z produkcji
(``--categories``), a ten zrzut traktowac jako zapas.

Uzycie:
    py scripts/dump_woo_product_categories.py
    py scripts/dump_woo_product_categories.py --database AlkanLocal --output dictionaries/woo_product_categories_snapshot.tsv
"""

from __future__ import annotations

import argparse
import csv
import glob
import html
import subprocess
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "dictionaries" / "woo_product_categories_snapshot.tsv"
MYSQL_GLOB = "C:/laragon/bin/mysql/*/bin/mysql.exe"

QUERY = """
SELECT t.term_id, tt.parent, t.name, t.slug, tt.count
FROM wp_term_taxonomy tt
JOIN wp_terms t ON t.term_id = tt.term_id
WHERE tt.taxonomy = 'product_cat'
ORDER BY tt.parent, t.name
"""

COLUMNS = ["term_id", "parent_id", "name", "slug", "count"]


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
    rows = [line.split("\t") for line in completed.stdout.decode("utf-8").splitlines() if line.strip()]
    for row in rows:
        row[2] = html.unescape(row[2])

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8", newline="") as stream:
        stream.write(f"# Kategorie produktow Woo z bazy {args.database} ({args.host}), zrzut {date.today().isoformat()}.\n")
        stream.write("# Generuje: py scripts/dump_woo_product_categories.py - nie edytowac recznie.\n")
        writer = csv.writer(stream, delimiter="\t", lineterminator="\n")
        writer.writerow(COLUMNS)
        writer.writerows(rows)
    print(f"Kategorii: {len(rows)} -> {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
