"""Sklada wejscie do briefow opisow dla produktow Kontakt Simon z eksportu Woo.

Po co: eksport WP All Export z alkan.pl ma tylko tytul i SKU, a brief opisu
potrzebuje faktow. Skrypt laczy:
  - tytul i SKU z eksportu Woo (nazwa w opisie = nazwa w sklepie),
  - atrybuty sklepu z CSV Baselinkera (Typ/Podtyp, Kolor, IP...),
  - parametry producenta z cache API e-katalogu (`cache/simon_api/`),
  - fakty o zgodnosci z katalogu serii (`--catalog-facts`) jako "Pasuje do".

Wynik to CSV w ukladzie Baselinkera z kolumna `Fraza kluczowa` - wejscie dla
`generate_product_descriptions.py --mode codex_brief`. Skrypt nie pisze prozy.
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import pathlib
import re
import sys

import openpyxl
import yaml

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "src"))

from simon_api import normalize_symbol  # noqa: E402

COLOR_SEGMENT = re.compile(r"^(AB)?\d{2}A?$")

# Parametr API -> etykieta faktu. Pominiete: opakowania, producent, ETIM, EAN,
# status, jednostki - nie opisuja produktu.
API_FACTS = {
    "Nazwa produktu/Rodzaj": "Rodzaj (producent)",
    "Kolor": "Kolor wg producenta",
    "RAL (zbliżony)": "RAL (zbliżony)",
    "Prąd znamionowy": "Prąd znamionowy",
    "Prąd znamionowy świetlówkowy": "Prąd znamionowy",
    "Napięcie znamionowe": "Napięcie znamionowe",
    "Rodzaj podłączenia / Typ zacisku": "Typ zacisków",
    "Stopień ochrony IP": "Stopień ochrony IP (producent)",
    "Uwagi": "Uwagi producenta",
    "Nadruk/Piktogram": "Nadruk/Piktogram",
    "Rodzaj materiału": "Materiał",
    "Wykończenie powierzchni": "Wykończenie powierzchni",
    "Zabezpieczenie powierzchni": "Zabezpieczenie powierzchni",
    "Sposób mocowania": "Sposób mocowania",
    "montaż": "Mocowanie",
    "Wysokość produktu": "Wysokość",
    "Szerokość produktu": "Szerokość",
    "Głębokość produktu": "Głębokość",
    "Głębokość wbudowania": "Głębokość wbudowania",
    "Głębokość wewnętrzna / pod osprzęt": "Głębokość wewnętrzna",
    "Metoda eksploatacji/Sposób działania": "Sposób działania",
    "Ilość klawiszy / dźwigni": "Liczba klawiszy",
    "Układ połączeń": "Układ połączeń",
    "Typ łącznika / przycisku": "Typ łącznika / przycisku",
    "Liczba biegunów": "Liczba biegunów",
    "Styki - działanie": "Styki",
    "Okienko kontrolki świetlnej": "Okienko kontrolki świetlnej",
    "Sterowanie mechaniczne": "Sterowanie mechaniczne",
    "Liczba modulów": "Liczba modułów",
    "Liczba jednostek": "Liczba gniazd",
    "Klapka uchylna": "Klapka uchylna",
    "Kolor klapki": "Kolor klapki",
    "Pokrywa gniazda": "Pokrywa gniazda",
    "Przesłony otworów prądowych": "Przesłony torów prądowych",
    "Klasa ochronności": "Klasa ochronności",
    "Ochrona przeciwkurzowa": "Ochrona przeciwkurzowa",
    "Kierunek wyjścia RJ": "Kierunek wyjścia RJ",
    "Ekranowany RJ": "Ekranowany RJ",
    "Model": "Model",
    "Liczba wyjść": "Liczba wyjść",
    "Kierunek montażu": "Kierunek montażu",
    "Adapter na": "Adapter na",
    "Rodzaj akcesoriów": "Rodzaj akcesoriów",
    "Standard HDMI": "Standard HDMI",
    "USB": "Złącza USB",
    "Moc ładowania": "Moc ładowania",
    "Maksymalna moc ładowania": "Maksymalna moc ładowania",
    "Zużycie energii": "Pobór mocy w czuwaniu",
    "Standard komunikacji": "Standard komunikacji",
    "Rodzaj transmisji": "Rodzaj transmisji",
    "Częstotliwość transmisji": "Częstotliwość transmisji",
    "kompatybilny z": "Kompatybilny z",
    "Sterowanie z aplikacji": "Aplikacja",
    "Maksymalna moc częstotliwości radiowej": "Maksymalna moc radiowa",
    "Maksymalne obciążenie": "Maksymalne obciążenie",
    "Tryb pracy": "Tryb pracy",
    "Pamięć poziomu jasności": "Pamięć poziomu jasności",
    "Moc obciążenia": "Moc obciążenia",
    "Napięcie zasilania": "Napięcie zasilania",
    "Do źródeł światła": "Do źródeł światła",
}
# Zakres czestotliwosci w API ma zla jednostke ("5--2400Hz"), wlasciwy zakres
# podaje katalog w "Pasuje do".
SKIPPED_API = {"Zakres częstotliwości"}

# Fraza kluczowa = doslowny prefiks tytulu. Domyslnie dwa pierwsze slowa,
# wyjatki dla nazw, w ktorych dwa slowa nie sa nazwa produktu.
KEYWORD_WORDS = {
    "Adapter": 1,
    "Ściemniacz": 3,
    "Pokrywa": 3,
}


def symbol_parts(sku: str) -> tuple[str, str]:
    """('NW6/2.01/X', '11') z 'NW6/2.01/X/11/KON'; kolor pusty dla NUV2."""
    symbol = normalize_symbol(sku)
    head, _, tail = symbol.rpartition("/")
    if head and COLOR_SEGMENT.match(tail):
        return head, tail
    return symbol, ""


def keyword_for(title: str) -> str:
    words = title.split()
    count = KEYWORD_WORDS.get(words[0], 2)
    return " ".join(words[:count]).rstrip(",;:")


def clean_value(value: str) -> str:
    value = re.sub(r",\s*\n\s*", ", ", str(value or ""))
    return re.sub(r"\s+", " ", value).strip()


def mm_value(param: dict) -> str:
    unit = param.get("rawUnit")
    raw = param.get("rawValue")
    if unit in {"mm", "cm"} and raw not in (None, ""):
        number = float(raw) * (10 if unit == "cm" else 1)
        return f"{number:.1f}".rstrip("0").rstrip(".").replace(".", ",") + " mm"
    return clean_value(param.get("value", ""))


def api_facts(product: dict) -> dict[str, str]:
    facts: dict[str, str] = {}
    for param in product.get("params") or []:
        name = param.get("name", "")
        if name in SKIPPED_API or name not in API_FACTS:
            continue
        label = API_FACTS[name]
        if label in {"Wysokość", "Szerokość", "Głębokość", "Głębokość wbudowania", "Głębokość wewnętrzna"}:
            value = mm_value(param)
        else:
            value = clean_value(param.get("value", ""))
        if value and label not in facts:
            facts[label] = value
    return facts


def compatibility(base: str, product: dict, catalog: dict) -> str:
    params = {p.get("name"): clean_value(p.get("value", "")) for p in product.get("params") or []}
    parts: list[str] = []
    if params.get("Typ produktu") == "moduł":
        parts.append(catalog["module"])
    if params.get("montaż", "").lower() == "mocowanie push in":
        parts.append(catalog["push_in"])
    if "uszczelkę ramki" in params.get("Uwagi", ""):
        parts.append(catalog["ip44_seal"])
    if base in catalog["by_symbol"]:
        parts.append(catalog["by_symbol"][base])
    return " ".join(" ".join(part.split()) for part in parts)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--woo-export", required=True, help="XLSX z WP All Export (id, Title, Sku).")
    parser.add_argument("--baselinker-csv", required=True, help="CSV Baselinkera z atrybutami sklepu (kolumna features).")
    parser.add_argument("--catalog-facts", required=True, help="YAML z faktami katalogu serii.")
    parser.add_argument("--api-cache", default="cache/simon_api/articles")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    raw = pathlib.Path(args.baselinker_csv).read_bytes().decode("utf-8-sig")
    base_rows = {row["sku"]: row for row in csv.DictReader(io.StringIO(raw), delimiter=";")}
    catalog = yaml.safe_load(pathlib.Path(args.catalog_facts).read_text(encoding="utf-8"))
    cache = pathlib.Path(args.api_cache)

    sheet = openpyxl.load_workbook(args.woo_export, read_only=True).worksheets[0]
    rows = sheet.iter_rows(values_only=True)
    header = [str(c or "") for c in next(rows)]
    out_rows: list[dict[str, str]] = []
    problems: list[str] = []
    for values in rows:
        record = dict(zip(header, values))
        sku = str(record.get("Sku") or "").strip()
        title = str(record.get("Title") or "").strip()
        if not sku:
            continue
        base, _color = symbol_parts(sku)
        cache_file = cache / (normalize_symbol(sku).replace("/", "_") + ".json")
        if not cache_file.exists():
            problems.append(f"{sku}: brak w cache API ({cache_file})")
            continue
        product = json.loads(cache_file.read_text(encoding="utf-8"))
        shop = base_rows.get(sku)
        if shop is None:
            problems.append(f"{sku}: brak w CSV Baselinkera")
            continue
        features = {k: v for k, v in json.loads(shop["features"]).items() if k != "Dane producenta"}
        for label, value in api_facts(product).items():
            features.setdefault(label, value)
        out_rows.append({
            "product_id": "",
            "name": title,
            "sku": sku,
            "ean": shop.get("ean") or product.get("ean", ""),
            "manufacturer_name": shop.get("manufacturer_name", ""),
            "category": shop.get("category", ""),
            "description": "",
            "features": json.dumps(features, ensure_ascii=False),
            "images_urls": "",
            "Pasuje do": compatibility(base, product, catalog),
            "Fraza kluczowa": keyword_for(title),
        })

    out = pathlib.Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(out_rows[0].keys()), delimiter=";")
        writer.writeheader()
        writer.writerows(out_rows)
    print(f"Zapisano {len(out_rows)} wierszy: {out}")
    for problem in problems:
        print("PROBLEM:", problem)
    if problems:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
