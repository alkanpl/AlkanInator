"""Pobiera z kanlux.com parametry produktu po kodzie Kanlux - cache dla nazw do Ekonoma.

Dla kodow, ktorych nie ma w pobranych plikach `Parametry*.xlsx`:

- podpowiedzi wyszukiwarki `/pl/ajax/products/autocomplete/<kod>` - JSON z `sku`, `seourl`, nazwa i typ;
  pusta lista albo brak dokladnie tego `sku` = produktu nie ma w aktualnej ofercie Kanlux,
- karta `/pl/produkt/<sku>/<seourl>` - specyfikacja w parach `params-name` / `params-value`, te same
  nazwy atrybutow co w plikach Parametry (Moc znamionowa [W], Calkowity znamionowy strumien...).

Wynik: `cache/kanlux_web/<kod>.json` = {"code", "found", "name", "params": {atrybut: wartosc}}.
`build_ekonom_names.py --params-dir` czyta ten cache jak plik Parametry (pliki xlsx wygrywaja).
Kazdy kod pytamy raz (`--refresh` pyta ponownie).

    py scripts/fetch_kanlux_web.py --report reports/Ekonom-nazwy-Kanlux_kontrola_2026-10-14.xlsx
    py scripts/fetch_kanlux_web.py --code 37246 37242
"""
from __future__ import annotations

import argparse
import html
import json
import re
import time
import urllib.parse
import urllib.request
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / "cache" / "kanlux_web"
AUTOCOMPLETE = "https://www.kanlux.com/pl/ajax/products/autocomplete/{}"
PRODUCT_PAGE = "https://www.kanlux.com/pl/produkt/{}/{}"


def get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "X-Requested-With": "XMLHttpRequest"})
    with urllib.request.urlopen(req, timeout=40) as resp:
        return resp.read()


def parse_params(raw: str) -> dict[str, str]:
    """Pary specyfikacji; ten sam atrybut dwa razy (Φuse "806" i "w kuli (360°)") - wygrywa wartosc z liczba."""
    out: dict[str, str] = {}
    for name, value in re.findall(r'<span class="params-name">(.*?)</span>\s*<span class="params-value">(.*?)</span>',
                                  raw, flags=re.S):
        key = " ".join(html.unescape(re.sub(r"<[^>]+>", " ", name)).split())
        val = " ".join(html.unescape(re.sub(r"<[^>]+>", " ", value)).split())
        if not key:
            continue
        old = out.get(key)
        if old is None or (not re.search(r"\d", old) and re.search(r"\d", val)):
            out[key] = val
    return out


def fetch(code: str) -> dict:
    hits = json.loads(get(AUTOCOMPLETE.format(urllib.parse.quote(code))))
    hit = next((h for h in hits if str(h.get("sku")) == code), None)
    if not hit:
        return {"code": code, "found": False, "name": "", "params": {}}
    raw = get(PRODUCT_PAGE.format(hit["sku"], urllib.parse.quote(hit["seourl"]))).decode("utf-8", "replace")
    name = f"{hit.get('name', '')} {hit.get('subcategoryDescription', '')}".strip()
    return {"code": code, "found": True, "name": name, "family": hit.get("familyName", ""),
            "params": parse_params(raw)}


def codes_from_report(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    rows = list(wb["Nazwy"].iter_rows(values_only=True))
    head = rows[0]
    i_code, i_note = head.index("Kod towaru"), head.index("Uwagi")
    return [str(r[i_code]).split("/")[0].strip().lstrip("0") for r in rows[1:]
            if r[i_code] and "brak w plikach Parametry" in (r[i_note] or "")]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--report", type=Path, help="skoroszyt kontrolny Kanlux - kody bez parametrow")
    parser.add_argument("--code", nargs="*", default=[])
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument("--delay", type=float, default=0.3)
    parser.add_argument("--workers", type=int, default=4, help="rownolegle zapytania (nie przesadzac - serwer producenta)")
    args = parser.parse_args()

    codes = list(dict.fromkeys(args.code + (codes_from_report(args.report) if args.report else [])))
    CACHE.mkdir(parents=True, exist_ok=True)

    def one(code: str) -> dict | None:
        path = CACHE / f"{code}.json"
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
    print(f"Kody: {len(codes)}  w ofercie kanlux.com: {found}  brak w ofercie: {missing}  bledy: {errors}")


if __name__ == "__main__":
    main()
