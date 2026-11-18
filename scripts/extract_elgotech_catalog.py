"""Katalog Elgotech (PDF ze strony producenta) -> wiersze tabel kod / EAN / dlugosc / przekroj / kolor.

Wynik `cache/elgotech/katalog_wiersze.json` czyta `src/build_ekonom_names.py` (sekcja `elgotech`, `catalog_rows`):
katalog uzupelnia przekroj i kolor, ktorych krotkie nazwy w Ekonomie nie podaja, i wylapuje kolor sprzeczny z nazwa.

    py scripts/extract_elgotech_catalog.py            # pobiera PDF, gdy go nie ma w cache
    py scripts/extract_elgotech_catalog.py --refresh  # pobiera PDF od nowa (nowe wydanie katalogu)

Wiersz tabeli w tekscie PDF: "[liczba gniazd] [dlugosc] KOD EAN reszta", np.
"4 1,5m PSF3-401 5902694013517 3x1 mm2 szary", "10m PK-1010 5902694053568 2x1 mm2 OMY", "R-30 5902694650606 biały".
"""
from __future__ import annotations

import argparse
import json
import re
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CATALOG_URL = "https://foto.elgotech.pl/katalog/katalog_elgotech_qr.pdf"
CACHE = ROOT / "cache" / "elgotech"
ROW_RE = re.compile(r"(\S+)\s+(590\d{10})\b(.*)$")


def parse_rows(text: str) -> dict[str, list[dict[str, str]]]:
    rows: dict[str, list[dict[str, str]]] = {}
    for line in text.splitlines():
        m = ROW_RE.search(line)
        if not m:
            continue
        code, ean, rest = m.group(1), m.group(2), m.group(3).strip()
        pre = line[:m.start(1)].split()
        length = next((p for p in pre if re.fullmatch(r"\d+(,\d+)?m", p)), "")
        sockets = next((p for p in pre if re.fullmatch(r"\d", p)), "")
        rows.setdefault(code, []).append({"ean": ean, "len": length, "gn": sockets, "rest": rest})
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--refresh", action="store_true", help="pobierz PDF od nowa")
    args = parser.parse_args()
    CACHE.mkdir(parents=True, exist_ok=True)
    pdf = CACHE / "katalog_elgotech_qr.pdf"
    if args.refresh or not pdf.exists():
        req = urllib.request.Request(CATALOG_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=300) as resp:
            pdf.write_bytes(resp.read())
    import pypdf

    text = "\n".join(page.extract_text() or "" for page in pypdf.PdfReader(str(pdf)).pages)
    rows = parse_rows(text)
    out = CACHE / "katalog_wiersze.json"
    out.write_text(json.dumps(rows, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"Kody w katalogu: {len(rows)} -> {out}")


if __name__ == "__main__":
    main()
