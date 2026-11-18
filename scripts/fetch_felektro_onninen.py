"""Karty produktow F-Elektro z onninen.pl - zrodlo IP i wersji dla nazw do Ekonoma.

F-Elektro nie ma wlasnej strony z katalogiem (f-elektro.pl nie istnieje). Hurtownia Onninen ma karty wszystkich
produktow marki F-ELEKTRO, a adres karty to tytul producenta z kodem na koncu, np.
`/produkt/F-ELEKTRO-rozdzielnica-lago-10m-gniazda-32a-5p-2x230v-ip65-F3-0322,626419`. Tytul rozni sie od nazwy
w Ekonomie tam, gdzie Ekonom zgubil ceche (wersje IP65 rozdzielnic i przedluzaczy, przelacznik RB-300).

Dwa kroki:
1. mapa strony (robots.txt jej nie blokuje): `sitemap_products_*.xml` -> adresy F-ELEKTRO -> `adresy.json`
   (kod -> adres i slug; IP z tytulu czyta generator nazw);
2. `--codes-from-report <skoroszyt kontrolny>`: karty kodow z arkusza "Powtorzenia" (identyczne nazwy w Ekonomie)
   -> `karty.json` (EAN, tytul, cena, waga, IP z opisu). Opis na karcie Onninen jest generowany - generator bierze z
   niego tylko IP i tylko, gdy opis podaje jedno IP.

Uzycie (z katalogu repo):
    py scripts/fetch_felektro_onninen.py --sitemap
    py scripts/fetch_felektro_onninen.py --codes-from-report reports/Ekonom-nazwy-F-Elektro_kontrola_<data>.xlsx
"""
from __future__ import annotations

import argparse
import html
import json
import re
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / "cache" / "onninen_felektro"
SITEMAP = "https://onninen.pl/sitemap.xml"
HEADERS = {"User-Agent": "Mozilla/5.0 (AlkanInator; nazwy do Ekonoma)"}


def code_from_slug(slug: str) -> str:
    """Kod producenta z konca sluga: "...-ip65-F3-0322" -> "F3-0322", "...-74185-4" -> "74185-4" (kropka = myslnik)."""
    m = re.search(r"-((?:F\d-\d{4})|(?:\d{2}-\d{4})|(?:\d{4,6}(?:-\d{1,3})?[A-Z]?)|(?:\d{4,5}FR))$", slug, re.I)
    return m.group(1).upper() if m else ""


def norm_code(code: str) -> str:
    """Kod z Ekonoma w zapisie sluga: "F3.0322" -> "F3-0322", "74185.4" -> "74185-4"."""
    return code.strip().upper().replace(".", "-")


def fetch_sitemap(session: requests.Session) -> dict[str, dict[str, str]]:
    index = session.get(SITEMAP, timeout=60).text
    parts = re.findall(r"<loc>(https://onninen\.pl/sitemap_products_\d+\.xml)</loc>", index)
    found: dict[str, dict[str, str]] = {}
    for i, url in enumerate(parts, 1):
        text = session.get(url, timeout=120).text
        for loc in re.findall(r"<loc>([^<]*/produkt/F-ELEKTRO-[^<]*)</loc>", text):
            slug = loc.rsplit("/", 1)[1].split(",")[0]
            code = code_from_slug(slug)
            if code:
                found.setdefault(code, {"url": loc, "slug": slug})
        print(f"  mapa {i}/{len(parts)}: {len(found)} kart F-Elektro", end="\r")
        time.sleep(0.5)
    print()
    return found


def page_record(session: requests.Session, url: str) -> dict[str, object]:
    text = session.get(url, timeout=60).text
    body = re.sub(r"<script.*?</script>|<style.*?</style>", " ", text, flags=re.S)
    lines = [ln.strip() for ln in html.unescape(re.sub(r"<[^>]+>", "\n", body)).split("\n") if ln.strip()]
    title = re.sub(r"\s+-\s+Hurtownia Onninen$", "", lines[0]) if lines else ""
    ean = re.search(r"EAN:\s*(?:<[^>]+>\s*)*(\d{13})", text)
    price = next((ln for ln in lines if re.fullmatch(r"\d+,\d{2} zł", ln)), "")
    dims = next((ln for ln in lines if re.fullmatch(r"[\d,]+ m / [\d,]+ m / [\d,]+ m / [\d,]+ kg / [\d,]+ m3", ln)), "")
    # opis karty: od ostatniego wystapienia tytulu (naglowek opisu) do zalacznikow ("D_<kod>_manual1")
    start = max((i for i, ln in enumerate(lines) if ln == title), default=len(lines))
    desc_lines = []
    for ln in lines[start + 1:]:
        if ln.startswith("D_") or ln.startswith("Oświadczenie"):
            break
        desc_lines.append(ln)
    desc = " ".join(desc_lines)
    ips = sorted(set(re.findall(r"\bIP\s?(\d{2})\b", desc)))
    return {"url": url, "title": title, "ean": ean.group(1) if ean else "", "price": price, "dims": dims,
            "ip_opis": [f"IP{x}" for x in ips]}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--sitemap", action="store_true", help="pobierz od nowa adresy kart z mapy strony")
    parser.add_argument("--codes-from-report", type=Path, help="skoroszyt kontrolny - kody z arkusza Powtorzenia")
    parser.add_argument("--codes", nargs="*", default=[], help="dodatkowe kody (np. F7.0143)")
    parser.add_argument("--refresh", action="store_true", help="pobierz ponownie karty juz zapisane")
    args = parser.parse_args()

    CACHE.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update(HEADERS)
    addr_path, cards_path = CACHE / "adresy.json", CACHE / "karty.json"
    if args.sitemap or not addr_path.is_file():
        addresses = fetch_sitemap(session)
        addr_path.write_text(json.dumps(addresses, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"Adresy kart: {len(addresses)} -> {addr_path}")
    addresses = json.loads(addr_path.read_text(encoding="utf-8"))

    codes = list(args.codes)
    if args.codes_from_report:
        import openpyxl

        ws = openpyxl.load_workbook(args.codes_from_report, read_only=True)["Powtorzenia"]
        rows = list(ws.values)
        col = rows[0].index("Kod towaru")
        codes += [str(r[col]).rsplit("/", 1)[0] for r in rows[1:] if r[col]]
    cards = json.loads(cards_path.read_text(encoding="utf-8")) if cards_path.is_file() else {}
    todo = [c for c in dict.fromkeys(codes) if args.refresh or norm_code(c) not in cards]
    missing = []
    for i, code in enumerate(todo, 1):
        rec = addresses.get(norm_code(code))
        if not rec:
            missing.append(code)
            continue
        try:
            cards[norm_code(code)] = page_record(session, rec["url"])
        except requests.RequestException as exc:
            print(f"\n  {code}: {exc}")
            continue
        print(f"  karta {i}/{len(todo)}: {code}", end="\r")
        if i % 25 == 0:
            cards_path.write_text(json.dumps(cards, ensure_ascii=False, indent=1), encoding="utf-8")
        time.sleep(1.0)
    cards_path.write_text(json.dumps(cards, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nKarty: {len(cards)} w {cards_path}; bez karty w Onninen: {len(missing)} {missing[:20]}")


if __name__ == "__main__":
    main()
