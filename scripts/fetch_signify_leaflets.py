"""Karty produktu PDF Signify (docs.signify.com) dla 12NC, ktorych nie ma juz na lighting.philips.pl.

Strona produktu znika po wycofaniu z oferty, ale karta PDF zostaje pod stalym adresem:
`https://www.docs.signify.com/assets/Leaflet/Philips_PROF/FP/<jezyk>/<12NC>_EU.<jezyk>.PROF.FP.pdf`
(probujemy pl_PL, potem en_GB / en_AA - nie kazdy jezyk istnieje dla kazdego kodu). Karta ma te same pola co strona
(Zuzycie energii / Power Consumption, Strumien swietlny / Luminous Flux...), a pierwsze linie to nazwa katalogowa,
12NC i podsumowanie "<rodzina>, 168 W, D355 mm, 20500 lm, 4000 K, ...".

Wynik trafia do `cache/signify_web/<12NC>.json` w ukladzie karty strony (`found: true`, `source: "pdf"`), z polskimi
nazwami pol - `build_ekonom_names.py` czyta go jak karte strony. Tekst PDF: `cache/signify_leaflets/<12NC>.txt`.
Pytamy tylko o kody zapisane przez `fetch_signify_web.py` jako brak karty (`--refresh` pyta ponownie o karty PDF).

    py scripts/fetch_signify_leaflets.py
"""
from __future__ import annotations

import argparse
import json
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WEB_CACHE = ROOT / "cache" / "signify_web"
TEXT_CACHE = ROOT / "cache" / "signify_leaflets"
LEAFLET = "https://www.docs.signify.com/assets/Leaflet/Philips_PROF/FP/{loc}/{code}_EU.{loc}.PROF.FP.pdf"
LOCALES = ["pl_PL", "en_GB", "en_AA"]

# pole karty strony (polska nazwa) -> etykieta w karcie PDF po polsku i po angielsku
FIELDS = {
    "Zużycie energii": r"Zużycie energii|Power Consumption",
    "Strumień Świetlny": r"Strumień Świetlny|Strumień świetlny|Luminous Flux",
    "Skorelowana temperatura barwowa (Nom)": r"Skorelowana temperatura barwowa \(Nom\)|Correlated Colou?r Temperature \(Nom\)",
    "Podstawa-nasadka": r"Podstawa-nasadka|Cap-Base",
    "Kod stopnia ochrony": r"Kod stopnia ochrony|IP code",
    "Kolor Korpusu": r"Kolor Korpusu|Housing colou?r",
    "Interfejs sterownika": r"Interfejs sterownika|Control interface",
    "Całkowita długość": r"Całkowita długość|Overall length",
    "Całkowita szerokość": r"Całkowita szerokość|Overall width",
    "Całkowita średnica": r"Całkowita średnica|Overall diameter",
    "Zunifikowany wskaźnik ograniczenia olśnienia CEN": r"Zunifikowany wskaźnik ograniczenia olśnienia CEN|Unified glare rating CEN",
    "Kształt bańki": r"Kształt bańki|Bulb shape",
    "Kąt rozsyłu światła (Nom)": r"Kąt rozsyłu światła \(Nom\)|Beam angle \(Nom\)",
    "Z możliwością przyciemniania": r"Z możliwością przyciemniania|Dimmable",
    "Napięcie (Nom)": r"Napięcie \(Nom\)|Voltage \(Nom\)",
    "Barwa źródła światła": r"Barwa źródła światła|Light source colou?r",
}
# wartosc za etykieta: liczba z jednostka albo krotki tekst
VALUES = {
    "Zużycie energii": r"((?:[\d.,]+\s?W\s*)+)",
    "Strumień Świetlny": r"((?:[\d.,]+(?: \d{3})*\s?lm\s*)+)",
    "Skorelowana temperatura barwowa (Nom)": r"((?:\d{4}\s?K\s*)+|-)",
    "Podstawa-nasadka": r"(\S+)",
    "Kod stopnia ochrony": r"(IP\d\d(?:/\d\d)?)",
    "Kolor Korpusu": r"([A-Za-zÀ-žąćęłńóśźżĄĆĘŁŃÓŚŹŻ-]+)",
    "Interfejs sterownika": r"(DALI|1-10 ?V|-|\S+)",
    "Całkowita długość": r"([\d.,]+(?: \d{3})?\s?mm)",
    "Całkowita szerokość": r"([\d.,]+(?: \d{3})?\s?mm)",
    "Całkowita średnica": r"([\d.,]+(?: \d{3})?\s?mm)",
    "Zunifikowany wskaźnik ograniczenia olśnienia CEN": r"(\d{2}|Not applicable)",
    "Kształt bańki": r"(\S+)",
    "Kąt rozsyłu światła (Nom)": r"(\d+\s?°)",
    "Z możliwością przyciemniania": r"(Tak|Nie|Yes|No)",
    "Napięcie (Nom)": r"([\d.,-]+\s?V)",
    "Barwa źródła światła": r"(RGBW|[^,]{0,40}?)(?=\s+[A-ZŁŚŻ][a-ząćęłńóśźż]|\s*$)",
}
EN_TO_PL = {"Yes": "Tak", "No": "Nie", "White": "Biel", "Grey": "Szary", "Gray": "Szary", "Black": "Czarny",
            "Aluminium": "Aluminium", "Aluminum": "Aluminium", "Dark": "Ciemnoszary"}


def get(url: str) -> bytes | None:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
            return data if data[:4] == b"%PDF" else None
    except urllib.error.HTTPError as exc:
        if exc.code in (403, 404):
            return None
        raise


def pdf_text(data: bytes) -> str:
    import fitz  # PyMuPDF

    doc = fitz.open(stream=data, filetype="pdf")
    return "\n".join(page.get_text() for page in doc)


def parse_leaflet(code: str, text: str, url: str, loc: str) -> dict:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    flat = " ".join(lines)
    if loc != "pl_PL":
        flat = re.sub(r"(?<=\d),(?=\d{3}\b)", "", flat)  # "5,300 lm" -> "5300 lm" (angielski separator tysiecy)
    params: dict[str, str] = {}
    for key, label in FIELDS.items():
        m = re.search(rf"(?i:{label})\s+{VALUES[key]}", flat)
        if m:
            value = " ".join(m.group(1).split())
            params[key] = EN_TO_PL.get(value, value)
    i = next((n for n, ln in enumerate(lines) if ln == code), None)
    catalog = " ".join(lines[:i]) if i else (lines[0] if lines else "")
    summary = lines[i + 1] if i is not None and i + 1 < len(lines) else ""
    family = summary.split(",")[0].strip()
    name = " ".join(lines[i + 1:i + 3]) if i is not None else ""  # "<rodzina>, 168 W, D355 mm, 20500 lm..."
    return {"code": code, "found": True, "source": "pdf", "name": name, "catalog_name": catalog, "title": "",
            "category": family, "gtin": "", "url": url, "params": params}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--code", nargs="*", default=[])
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument("--delay", type=float, default=0.3)
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--reparse", action="store_true", help="tylko ponowny odczyt zapisanych tekstow PDF, bez sieci")
    args = parser.parse_args()

    if args.reparse:
        for f in sorted(WEB_CACHE.glob("*.json")):
            rec = json.loads(f.read_text(encoding="utf-8"))
            txt = TEXT_CACHE / f"{rec['code']}.txt"
            if rec.get("source") == "pdf" and txt.exists():
                loc = re.search(r"/FP/(\w+)/", rec["url"]).group(1)
                new = parse_leaflet(rec["code"], txt.read_text(encoding="utf-8"), rec["url"], loc)
                f.write_text(json.dumps(new, ensure_ascii=False, indent=1), encoding="utf-8")
        print("przeczytane ponownie")
        return

    TEXT_CACHE.mkdir(parents=True, exist_ok=True)
    codes = args.code or sorted(
        json.loads(f.read_text(encoding="utf-8"))["code"] for f in WEB_CACHE.glob("*.json")
        if not json.loads(f.read_text(encoding="utf-8"))["found"]
        or (args.refresh and json.loads(f.read_text(encoding="utf-8")).get("source") == "pdf"))

    def one(code: str) -> str:
        for loc in LOCALES:
            url = LEAFLET.format(loc=loc, code=code)
            try:
                data = get(url)
            except Exception as exc:  # siec - kod zostaje bez karty PDF, ponowimy nastepnym razem
                print(f"{code}: blad {exc}")
                return "blad"
            time.sleep(args.delay)
            if data:
                text = pdf_text(data)
                (TEXT_CACHE / f"{code}.txt").write_text(text, encoding="utf-8")
                rec = parse_leaflet(code, text, url, loc)
                (WEB_CACHE / f"{code}.json").write_text(json.dumps(rec, ensure_ascii=False, indent=1), encoding="utf-8")
                return loc
        return "brak"

    from collections import Counter
    from concurrent.futures import ThreadPoolExecutor

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        stats = Counter(pool.map(one, codes))
    print(f"Kody bez strony: {len(codes)}  karty PDF: {dict(stats)}")


if __name__ == "__main__":
    main()
