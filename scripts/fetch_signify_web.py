"""Pobiera z lighting.philips.pl karty produktow Signify (Philips / PILA) po 12NC - cache dla nazw do Ekonoma.

`Kod producenta` w zestawieniach Ekonoma Signify to 12NC (929003691502). Karta
`/content/signify-multibrand/pl/pl/product/product-detail-page.<12NC>_EU.html` ma:

- JSON-LD `Product`: nazwa z parametrami ("Philips MASTER, LEDbulb, A60, 60 W, E27, 2700 K, 840 lm, CRI 80 ..."),
  opis, kategoria ("Lampy i tuby LED > Zarowki LED > MASTER UltraEfficient LED bulb"), `gtin13`,
- specyfikacje w parach `specification__list-header` / `specification__list-description`
  (Podstawa-nasadka, Strumien swietlny, Moc, Temperatura barwowa, Kod IP ...),
- `pageLoadInfo` z nazwa katalogowa ("MAS LEDBulb D 4-60W E27 827 A60 CL G UE").

Strona 404 = produktu nie ma w aktualnej ofercie Signify (wycofany) - zostaje z danymi z nazwy. Serwer pod
obciazeniem oddaje czasem 404 dla istniejacej karty - po pierwszym przebiegu uruchom `--retry-missing`.

Wynik: `cache/signify_web/<12NC>.json` = {"code", "found", "name", "catalog_name", "title", "category",
"gtin", "url", "params": {atrybut: wartosc}}. Kazdy kod pytamy raz (`--refresh` pyta ponownie).

    py scripts/fetch_signify_web.py --input "archived_input_files/input_2026-11-10/EkonomNazwy/Signify"
    py scripts/fetch_signify_web.py --code 929003691502 911401881902
"""
from __future__ import annotations

import argparse
import csv
import html
import json
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / "cache" / "signify_web"
PRODUCT_PAGE = "https://www.lighting.philips.pl/content/signify-multibrand/pl/pl/product/product-detail-page.{}_EU.html"


def get(url: str) -> tuple[int, str]:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return resp.status, resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return 404, ""
        raise


def clean(raw: str) -> str:
    return " ".join(html.unescape(re.sub(r"<[^>]+>", " ", raw)).replace(" ", " ").split())


def parse_params(raw: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for name, value in re.findall(
            r'<dt class="specification__list-header[^"]*">(.*?)</dt>\s*'
            r'<dd class="specification__list-description-wrapper">(.*?)</dd>', raw, flags=re.S):
        key, val = clean(name), clean(value)
        if key and val and key not in out:
            out[key] = val
    return out


def parse_page(code: str, raw: str) -> dict:
    rec: dict = {"code": code, "found": True, "name": "", "catalog_name": "", "title": "", "category": "",
                 "gtin": "", "url": "", "params": parse_params(raw)}
    m = re.search(r"<title>(.*?)</title>", raw, flags=re.S)
    rec["title"] = clean(m.group(1)) if m else ""
    m = re.search(r'<script type="application/ld\+json">\s*(\{"@context":"https://schema.org/","@type":"Product".*?)'
                  r"\s*</script>", raw, flags=re.S)
    if m:
        try:
            ld = json.loads(m.group(1))
            rec.update(name=ld.get("name", ""), category=ld.get("category", ""), gtin=ld.get("gtin13", ""),
                       url=ld.get("url", ""), description=ld.get("description", ""))
        except json.JSONDecodeError:
            pass
    m = re.search(r'\\x22Name\\x22:\\x22(.*?)\\x22', raw)
    if m:
        # lancuch JS: \x22, \u002D, \/ - po zamianie \xNN na \u00NN to poprawny lancuch JSON
        try:
            rec["catalog_name"] = json.loads('"' + re.sub(r"\\x([0-9a-fA-F]{2})", r"\\u00\1", m.group(1)) + '"')
        except json.JSONDecodeError:
            pass
    rec["catalog_name"] = rec["params"].get("Nazwa produktu na zamówieniu") or rec["catalog_name"]
    return rec


def fetch(code: str) -> dict:
    status, raw = get(PRODUCT_PAGE.format(code))
    if status == 404 or "specification__list" not in raw:
        return {"code": code, "found": False, "params": {}}
    return parse_page(code, raw)


def codes_from_inputs(path: Path) -> list[str]:
    files = sorted(path.glob("*.csv")) if path.is_dir() else [path]
    codes = []
    for f in files:
        with f.open(encoding="cp1250", newline="") as fh:
            for row in csv.DictReader(fh, delimiter=";"):
                code = (row.get("Kod producenta") or "").strip()
                if re.fullmatch(r"\d{12}", code):
                    codes.append(code)
    return codes


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", type=Path, help="CSV z Ekonoma albo katalog z plikami CSV (kolumna Kod producenta)")
    parser.add_argument("--code", nargs="*", default=[])
    parser.add_argument("--refresh", action="store_true")
    # serwer Signify pod obciazeniem oddaje czasem 404 dla istniejacej karty (CoreLine Malaga 910925865340) -
    # drugi przebieg pyta jeszcze raz tylko o kody bez karty
    parser.add_argument("--retry-missing", action="store_true", help="zapytaj ponownie o kody zapisane jako brak karty")
    parser.add_argument("--delay", type=float, default=0.3)
    parser.add_argument("--workers", type=int, default=4, help="rownolegle zapytania (nie przesadzac - serwer producenta)")
    args = parser.parse_args()

    codes = list(dict.fromkeys(args.code + (codes_from_inputs(args.input) if args.input else [])))
    CACHE.mkdir(parents=True, exist_ok=True)

    def one(code: str) -> dict | None:
        path = CACHE / f"{code}.json"
        if path.exists() and not args.refresh:
            rec = json.loads(path.read_text(encoding="utf-8"))
            if rec["found"] or not args.retry_missing:
                return rec
        try:
            rec = fetch(code)
        except Exception as exc:  # siec / zmiana strony - kod zostaje bez cache, ponowimy nastepnym razem
            print(f"{code}: blad {exc}")
            return None
        path.write_text(json.dumps(rec, ensure_ascii=False, indent=1), encoding="utf-8")
        time.sleep(args.delay)
        return rec

    from concurrent.futures import ThreadPoolExecutor

    found = missing = errors = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        for n, rec in enumerate(pool.map(one, codes), 1):
            if rec is None:
                errors += 1
            else:
                found += rec["found"]
                missing += not rec["found"]
            if n % 200 == 0:
                print(f"{n}/{len(codes)}", flush=True)
    print(f"Kody: {len(codes)}  karta na lighting.philips.pl: {found}  brak karty: {missing}  bledy: {errors}")


if __name__ == "__main__":
    main()
