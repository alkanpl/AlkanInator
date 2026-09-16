"""Sklada plik aktualizacji Baselinkera: poprawione nazwy + gotowe opisy HTML.

Bierze kanoniczny CSV wyslany 2026-09-15, podmienia kolumne `description` na
recznie napisane opisy z katalogu HTML i nanosi punktowe poprawki nazw
produktow. Plik wejsciowy NIE jest nadpisywany - wynik idzie do nowego,
datowanego pliku.

Poprawki nazw sa jawna lista, a nie regula: kazda zmiana nazwy produktu jest
widoczna dla klienta, wiec musi byc wskazana wprost i uzasadniona polem
producenta.
"""
from __future__ import annotations

import argparse
import csv
import io
import pathlib
import re

# Nazwy rozjezdzajace sie z danymi producenta. Klucz: SKU.
# Wartosc: (fragment w nazwie, poprawka, pole producenta uzasadniajace zmiane).
NAME_FIXES = {
    "TESL1W.01/142/KON": ("2–250 W", "2–200 W", "Moc obciążenia: 2-200W"),
    "TOS4K14.01/142/KON": ("4200 K", "4000 K", "Temp. barw.: 4000° K"),
}


def slug(sku: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", sku).strip("_")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Kanoniczny CSV Baselinkera.")
    parser.add_argument("--html-dir", required=True, help="Katalog z gotowymi opisami.")
    parser.add_argument("--output", required=True)
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    raw = pathlib.Path(args.input).read_bytes().decode("utf-8-sig")
    rows = list(csv.DictReader(io.StringIO(raw), delimiter=";"))
    html_dir = pathlib.Path(args.html_dir)

    log: list[str] = [f"ZRODLO: {args.input}", f"OPISY: {args.html_dir}", ""]
    replaced = 0
    missing: list[str] = []
    renamed: list[str] = []

    for row in rows:
        sku = row["sku"]
        path = html_dir / f"{slug(sku)}.html"
        if path.exists():
            row["description"] = path.read_text(encoding="utf-8").strip()
            replaced += 1
        else:
            missing.append(sku)

        fix = NAME_FIXES.get(sku)
        if fix:
            old, new, source = fix
            if old not in row["name"]:
                raise SystemExit(f"W nazwie {sku} nie ma fragmentu '{old}' - sprawdz plik wejsciowy.")
            before = row["name"]
            row["name"] = row["name"].replace(old, new)
            renamed.append(sku)
            log.append(f"NAZWA POPRAWIONA: {sku}")
            log.append(f"   bylo : {before}")
            log.append(f"   jest : {row['name']}")
            log.append(f"   pole producenta: {source}")
            log.append("")

    out = pathlib.Path(args.output)
    with out.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    log.append(f"WIERSZY: {len(rows)}")
    log.append(f"OPISOW PODMIENIONYCH: {replaced}")
    log.append(f"NAZW POPRAWIONYCH: {len(renamed)} ({', '.join(renamed)})")
    log.append(f"BEZ PLIKU OPISU: {len(missing)}" + (f" -> {missing}" if missing else ""))
    pathlib.Path(args.report).write_text("\n".join(log), encoding="utf-8")

    print(f"OK: {out} ({len(rows)} wierszy, {replaced} opisow, {len(renamed)} nazw)")
    if missing:
        print(f"UWAGA: {len(missing)} produktow bez pliku opisu - nie zostaly zmienione.")


if __name__ == "__main__":
    main()
