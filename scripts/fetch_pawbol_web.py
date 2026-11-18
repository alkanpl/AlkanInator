"""Pobiera katalog produktow z pawbol.pl (REST API WordPressa) - zrodlo danych dla nazw do Ekonoma.

pawbol.pl to WordPress z typem wpisu `produkt` (`/wp-json/wp/v2/produkt`, ~1900 produktow). Tytul to czysta
nazwa producenta ("Rozdzielnica R-BOX 150: 2x250V, 250V, IP44"), tresc zaczyna sie komentarzem z kodem
(`<!-- S-BOX 036C#MARK -->`) i niesie pelny opis (czasem z IK, IP, dlugoscia), a taksonomia `grupa-produktow`
podaje grupe ("Puszki S-BOX z dlawikami"). Kod producenta = `Kod producenta` w Ekonomie (`Kod towaru` bez `/PAW`).

Wynik: `cache/pawbol_web/produkty.json` = {kod: {"code", "name", "description", "groups": [...], "link"}} i
`cache/pawbol_web/grupy.json` (id -> nazwa i rodzic). `build_ekonom_names.py` czyta to przez `params_web_loader:
pawbol`. Jeden przebieg to ~20 zapytan po 100 produktow.

Drugie zrodlo (wiecej kodow niz strona): cennik PDF z pawbol.pl/strefa-klienta/cennik (`--cennik-url`), zapisywany
do `cache/pawbol_web/pdf/` i rozkladany na `cache/pawbol_web/cennik_2026.json` = {kod: {"name", "group"}}. Rekord w
tekscie PDF: "<lp> <kod>", nazwa, "<cena> zł", grupa... "Produkt wycofany" zostaje w pliku - loader traktuje go jak brak.

    py scripts/fetch_pawbol_web.py
    py scripts/fetch_pawbol_web.py --skip-web   # tylko cennik
"""
from __future__ import annotations

import argparse
import html
import json
import re
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / "cache" / "pawbol_web"
API = "https://www.pawbol.pl/wp-json/wp/v2/"
HEADERS = {"User-Agent": "Mozilla/5.0 (AlkanInator)"}


def get(url: str) -> tuple[list | dict, dict[str, str]]:
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8")), dict(resp.headers)


def clean(text: str) -> str:
    return " ".join(html.unescape(re.sub(r"<[^>]+>", " ", text)).split())


def fetch_all(endpoint: str, fields: str, pause: float) -> list[dict]:
    out, page, pages = [], 1, 1
    while page <= pages:
        data, headers = get(f"{API}{endpoint}?per_page=100&page={page}&_fields={fields}")
        pages = int(headers.get("X-WP-TotalPages") or headers.get("x-wp-totalpages") or 1)
        out += data
        print(f"{endpoint}: strona {page}/{pages}")
        page += 1
        time.sleep(pause)
    return out


CENNIK_URL = "https://www.pawbol.pl/wp-content/uploads/2024/12/Cennik-PAWBOL-2026.pdf"


def parse_price_list(text: str) -> dict[str, dict[str, str]]:
    """Tekst cennika PDF -> {kod: {"name", "group"}}; rekord: "<lp> <kod>", nazwa, "<cena> zł", grupa."""
    lines = [line.strip() for line in text.splitlines()]
    out: dict[str, dict[str, str]] = {}
    i = 0
    while i < len(lines) - 3:
        m = re.match(r"^(\d+) (.+)$", lines[i])
        if m and lines[i + 2].endswith("zł"):
            out[m.group(2).strip()] = {"name": lines[i + 1], "group": lines[i + 3]}
            i += 3
        i += 1
    return out


def fetch_price_list(url: str) -> int:
    import fitz  # PyMuPDF - tylko dla cennika

    pdf_dir = CACHE / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = pdf_dir / url.rsplit("/", 1)[-1]
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=120) as resp:
        pdf_path.write_bytes(resp.read())
    text = "".join(page.get_text() for page in fitz.open(pdf_path))
    (CACHE / "cennik_2026.txt").write_text(text, encoding="utf-8")
    records = parse_price_list(text)
    (CACHE / "cennik_2026.json").write_text(json.dumps(records, ensure_ascii=False, indent=0), encoding="utf-8")
    return len(records)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pause", type=float, default=1.0, help="przerwa miedzy zapytaniami [s]")
    parser.add_argument("--cennik-url", default=CENNIK_URL, help="cennik PDF Pawbol (pusty = bez cennika)")
    parser.add_argument("--skip-web", action="store_true", help="tylko cennik, bez REST API strony")
    args = parser.parse_args()
    CACHE.mkdir(parents=True, exist_ok=True)
    if args.cennik_url:
        print(f"cennik: {fetch_price_list(args.cennik_url)} kodow")
    if args.skip_web:
        return

    groups = {}
    for g in fetch_all("grupa-produktow", "id,name,parent,slug", args.pause):
        groups[g["id"]] = {"name": html.unescape(g["name"]), "parent": g["parent"], "slug": g["slug"]}
    (CACHE / "grupy.json").write_text(json.dumps(groups, ensure_ascii=False, indent=1), encoding="utf-8")

    products: dict[str, dict] = {}
    duplicates = 0
    for p in fetch_all("produkt", "id,title,content,link,grupa-produktow", args.pause):
        raw = p["content"]["rendered"]
        # "<!-- H.7015B -25 -->", "<!-- S-BOX 036C#MARK -->" - kod w Ekonomie bez spacji przed wariantem
        # opakowania ("H.7015B-25"), ze spacja w nazwie serii ("S-BOX 036C")
        m = re.search(r"<!--\s*(.*?)\s*-->", raw)
        if not m:
            continue
        code = re.sub(r"\s+(?=-)", "", " ".join(m.group(1).replace("#MARK", "").split()))
        rec = {"code": code, "name": clean(p["title"]["rendered"]), "description": clean(raw),
               "groups": [groups.get(g, {}).get("name", "") for g in p.get("grupa-produktow", [])],
               "link": p["link"]}
        if code in products:
            duplicates += 1
            continue
        products[code] = rec
    (CACHE / "produkty.json").write_text(json.dumps(products, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"produkty z kodem: {len(products)}, powtorzone kody: {duplicates}, grupy: {len(groups)}")


if __name__ == "__main__":
    main()
