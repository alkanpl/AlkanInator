"""Pobiera z bemko.pl karte produktu po kodzie Bemko - cache dla nazw do Ekonoma.

Wyszukiwarka `https://bemko.pl/pl/wyniki-wyszukiwania/1?search=<kod>` zwraca kafelki "NAZWA (KOD)" z linkiem
do karty; bierzemy tylko kafelek z dokladnie tym kodem (szukanie "C37-PLA-350-240-3K" zwraca tez "-MS").
Karta `/pl/produkt/<slug>` ma naglowek z nazwa, "Kod:", "EAN:" i tabele "Specyfikacja techniczna"
(Moc w trybie wlaczenia (Pon), Strumien swietlny, Barwa swiatla, Stopien szczelnosci, Material, Przekroj...).
Kod nieznany wyszukiwarce = produktu nie ma na stronie (wycofany).

Wynik: `cache/bemko_web/<kod>.json` = {"code", "found", "name", "url", "ean", "params": {atrybut: wartosc}}.
`build_ekonom_names.py` czyta cache (`web_params_cache` w sekcji `bemko`). Kod z `Kod producenta` w CSV
Ekonoma (C37-PLA-350-240-3K). Kazdy kod pytamy raz (`--refresh` pyta ponownie).

    py scripts/fetch_bemko_web.py --input "input/Nazwy do ekonom/Bemko/Bemko.csv"
    py scripts/fetch_bemko_web.py --code C37-PLA-450-360-3K E00-PM22-100
"""
from __future__ import annotations

import argparse
import csv
import html
import json
import re
import ssl
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / "cache" / "bemko_web"
SEARCH = "https://bemko.pl/pl/wyniki-wyszukiwania/1?search={}"
# bemko.pl ma niepelny lancuch certyfikatu (curl: blad 60) - czytamy publiczne karty bez weryfikacji
CONTEXT = ssl._create_unverified_context()


def clean(text: str) -> str:
    return " ".join(html.unescape(re.sub(r"<[^>]+>", " ", text)).split())


def get(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=40, context=CONTEXT) as resp:
        return resp.read().decode("utf-8", "replace")


def norm_code(code: str) -> str:
    return re.sub(r"\s+", "", code).upper()


def find_card(raw: str, code: str) -> tuple[str, str] | None:
    """Kafelek wynikow wyszukiwania z dokladnie tym kodem: (url karty, nazwa)."""
    for url, body in re.findall(r'<a[^>]+href="(https://bemko\.pl/pl/produkt/[^"]+)"[^>]*>(.*?)</a>', raw, flags=re.S):
        m = re.search(r"<h6[^>]*>(.*?)</h6>", body, flags=re.S)
        text = clean(m.group(1) if m else body)
        m = re.match(r"(.*)\(([^()]+)\)\s*$", text)
        if m and norm_code(m.group(2)) == norm_code(code):
            return url, m.group(1).strip()
    return None


def parse_page(raw: str) -> dict:
    title = re.search(r"<h1>(.*?)</h1>", raw, flags=re.S)
    code = re.search(r"Kod:\s*</span>\s*<span[^>]*>(.*?)</span>", raw, flags=re.S)
    ean = re.search(r"EAN:\s*</span>\s*<span[^>]*>(.*?)</span>", raw, flags=re.S)
    params: dict[str, str] = {}
    for label, value in re.findall(r"<th[^>]*>([^<]*)</th>\s*<td[^>]*>(.*?)</td>", raw, flags=re.S):
        key = clean(label)
        if key and key not in params:
            params[key] = clean(value)
    return {"name": clean(title.group(1)) if title else "", "page_code": clean(code.group(1)) if code else "",
            "ean": clean(ean.group(1)) if ean else "", "params": params}


def fetch(code: str) -> dict:
    raw = get(SEARCH.format(urllib.parse.quote(code.strip())))
    if "Wyniki wyszukiwania dla frazy" not in raw:
        # strona bledu / ochrona przed botami - to nie jest "brak produktu", kod zostaje bez cache
        raise RuntimeError("niepelna strona wynikow")
    card = find_card(raw, code)
    if not card:
        return {"code": code, "found": False, "name": "", "url": "", "ean": "", "params": {}}
    url, card_name = card
    rec = parse_page(get(url))
    return {"code": code, "found": True, "name": rec["name"] or card_name, "url": url, "ean": rec["ean"],
            "page_code": rec["page_code"], "params": rec["params"]}


def codes_from_csv(path: Path) -> list[str]:
    with path.open(encoding="cp1250", newline="") as f:
        return [(r.get("Kod producenta") or "").strip() for r in csv.DictReader(f, delimiter=";")]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", type=Path, help="CSV Ekonoma Bemko - kody z kolumny Kod producenta")
    parser.add_argument("--code", nargs="*", default=[])
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument("--delay", type=float, default=2.0)
    parser.add_argument("--workers", type=int, default=1, help="rownolegle zapytania - bemko.pl blokuje IP po ~250 zapytaniach w 5 watkach")
    args = parser.parse_args()

    codes = list(dict.fromkeys(c for c in args.code + (codes_from_csv(args.input) if args.input else []) if c))
    CACHE.mkdir(parents=True, exist_ok=True)

    def one(code: str) -> dict | None:
        path = CACHE / f"{re.sub(r'[^A-Za-z0-9._-]', '_', code)}.json"
        if path.exists() and not args.refresh:
            return json.loads(path.read_text(encoding="utf-8"))
        try:
            rec = fetch(code)
        except Exception as exc:  # siec / zmiana strony - kod zostaje bez cache, ponowimy nastepnym razem
            print(f"{code}: blad {exc}")
            return None
        path.write_text(json.dumps(rec, ensure_ascii=False, indent=1), encoding="utf-8")
        time.sleep(args.delay)
        return rec

    found = missing = errors = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        for rec in pool.map(one, codes):
            if rec is None:
                errors += 1
            else:
                found += rec["found"]
                missing += not rec["found"]
    print(f"Kody: {len(codes)}  na bemko.pl: {found}  brak na stronie: {missing}  bledy: {errors}")


if __name__ == "__main__":
    main()
