"""Pobiera z gtv.com.pl dane techniczne produktu po kodzie GTV - cache dla nazw do Ekonoma.

Karta `https://gtv.com.pl/produkt/<kod>/` (kod z `Kod towaru` bez `/GTV`, np. LD-BERGA50W-30) ma tabele
"Dane techniczne produktu" (`p-product__attributes-label` / `-values`): Temperatura barwowa swiatla [K],
Strumien swietlny [lm], Stopien ochrony IP... 404 = produktu nie ma na stronie GTV (wycofany).

Wynik: `cache/gtv_web/<kod>.json` = {"code", "found", "name", "params": {atrybut: wartosc}}.
`build_ekonom_names.py` czyta cache (`params_web_cache` w sekcji `gtv`) i uzupelnia tylko braki -
parametry z nazwy w Ekonomie maja pierwszenstwo. Kazdy kod pytamy raz (`--refresh` pyta ponownie).

    py scripts/fetch_gtv_web.py --report reports/Ekonom-nazwy-Gtv_kontrola_2026-10-14.xlsx
    py scripts/fetch_gtv_web.py --code LD-BERGA50W-30 LD-NO24CB-00
"""
from __future__ import annotations

import argparse
import html
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / "cache" / "gtv_web"
PRODUCT_PAGE = "https://gtv.com.pl/produkt/{}/"


def clean(text: str) -> str:
    return " ".join(html.unescape(re.sub(r"<[^>]+>", " ", text)).split())


def parse_page(raw: str) -> tuple[str, dict[str, str]]:
    title = re.search(r"<h1[^>]*>(.*?)</h1>", raw, flags=re.S)
    params = {}
    for label, value in re.findall(r'class="p-product__attributes-label">(.*?)</th>\s*'
                                   r'<td class="p-product__attributes-values">(.*?)</td>', raw, flags=re.S):
        key = clean(label)
        if key and key not in params:
            params[key] = clean(value)
    return (clean(title.group(1)) if title else ""), params


def fetch(code: str) -> dict:
    req = urllib.request.Request(PRODUCT_PAGE.format(urllib.parse.quote(code)), headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=40) as resp:
            raw = resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return {"code": code, "found": False, "name": "", "params": {}}
        raise
    name, params = parse_page(raw)
    return {"code": code, "found": True, "name": name, "params": params}


def codes_from_report(path: Path, statuses: set[str]) -> list[str]:
    """Kody oswietlenia/elektryki GTV (bez narzedzi Hogert HT...) z problemem w skoroszycie kontrolnym."""
    wb = openpyxl.load_workbook(path, read_only=True)
    rows = list(wb["Nazwy"].iter_rows(values_only=True))
    head = rows[0]
    i_code, i_status = head.index("Kod towaru"), head.index("Status")
    out = []
    for r in rows[1:]:
        code = str(r[i_code] or "").split("/")[0].strip()
        if code and not code.upper().startswith("HT") and r[i_status] in statuses:
            out.append(code)
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--report", type=Path, help="skoroszyt kontrolny GTV - kody z problemem")
    parser.add_argument("--status", nargs="*", default=["DO_SPRAWDZENIA", "WARNING"])
    parser.add_argument("--code", nargs="*", default=[])
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument("--delay", type=float, default=0.3)
    parser.add_argument("--workers", type=int, default=4, help="rownolegle zapytania (nie przesadzac - serwer producenta)")
    args = parser.parse_args()

    codes = list(dict.fromkeys(args.code + (codes_from_report(args.report, set(args.status)) if args.report else [])))
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
    print(f"Kody: {len(codes)}  na gtv.com.pl: {found}  brak na stronie: {missing}  bledy: {errors}")


if __name__ == "__main__":
    main()
