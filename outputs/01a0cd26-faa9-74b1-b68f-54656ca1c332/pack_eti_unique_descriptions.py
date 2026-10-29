"""Pack individually authored ETI copy into the BaseLinker import workbook.

The TSV holds the prose for each SKU. This script only adds HTML markup and
copies confirmed technical values; it does not write product prose.
"""

from __future__ import annotations

import html
import json
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment


HERE = Path(__file__).resolve().parent
SOURCE = HERE / "ETI_modulowka_IMPORT_BaseLinker.xlsx"
COPY = HERE / "eti_unique_copy.tsv"
TARGET = HERE / "ETI_modulowka_IMPORT_BaseLinker_opisy_unikalne_FAQ.xlsx"

KEYWORDS = {
    "Wyłącznik nadprądowy MCB": "Wyłącznik nadprądowy",
    "Ogranicznik mocy": "Ogranicznik mocy",
    "Wyłącznik różnicowoprądowy RCCB": "Wyłącznik różnicowoprądowy",
    "Wyłącznik różnicowo-nadprądowy RCBO": "Wyłącznik różnicowo-nadprądowy",
    "Rozłącznik izolacyjny": "Rozłącznik izolacyjny",
    "Przełącznik trójpołożeniowy": "Przełącznik trójpołożeniowy",
    "Sygnalizator obecności napięcia": "Sygnalizator obecności napięcia",
    "Stycznik modułowy": "Stycznik modułowy",
    "Rozłącznik bezpiecznikowy": "Rozłącznik bezpiecznikowy",
    "Wstawka redukcyjna": "Wstawka redukcyjna",
    "Podstawa bezpiecznikowa DC": "Podstawa bezpiecznikowa DC",
}

UNITS = {
    "Prąd znamionowy [A]": " A",
    "Znamionowa zdolność zwarciowa Icn [kA]": " kA",
    "Prąd różnicowy IΔn [mA]": " mA",
    "Warunkowy prąd zwarciowy Inc [kA]": " kA",
    "Napięcie znamionowe AC [V]": " V",
    "Napięcie znamionowe DC [V]": " V",
    "Napięcie cewki [V]": " V",
}

LABELS = {
    "Symbol": "Symbol",
    "Charakterystyka wyzwalania": "Charakterystyka wyzwalania",
    "Prąd znamionowy [A]": "Prąd znamionowy",
    "Układ biegunów": "Układ biegunów",
    "Znamionowa zdolność zwarciowa Icn [kA]": "Znamionowa zdolność zwarciowa Icn",
    "Prąd różnicowy IΔn [mA]": "Prąd różnicowy IΔn",
    "Typ prądu różnicowego": "Typ prądu różnicowego",
    "Warunkowy prąd zwarciowy Inc [kA]": "Warunkowy prąd zwarciowy Inc",
    "Napięcie znamionowe AC [V]": "Napięcie znamionowe AC",
    "Napięcie znamionowe DC [V]": "Napięcie znamionowe DC",
    "Napięcie cewki [V]": "Napięcie cewki",
    "Konfiguracja styków": "Konfiguracja styków",
    "Rodzaj wkładki": "Rodzaj wkładki",
    "Rozmiar": "Rozmiar wkładki",
    "Barwa lampek": "Barwa lampek",
    "Ilość lampek": "Liczba lampek",
    "Pozycje przełącznika": "Pozycje przełącznika",
    "Kompatybilność": "Kompatybilność",
    "Rodzaj akcesoria": "Rodzaj akcesoria",
    "Liczba modułów": "Liczba modułów",
    "Stopień ochrony [IP]": "Stopień ochrony IP",
    "Kategoria użytkowania": "Kategoria użytkowania",
}

KEY_FEATURES = {
    "Wyłącznik nadprądowy MCB": ["Prąd znamionowy [A]", "Układ biegunów"],
    "Ogranicznik mocy": ["Prąd znamionowy [A]", "Układ biegunów"],
    "Wyłącznik różnicowoprądowy RCCB": ["Prąd różnicowy IΔn [mA]", "Układ biegunów"],
    "Wyłącznik różnicowo-nadprądowy RCBO": ["Prąd różnicowy IΔn [mA]", "Typ prądu różnicowego"],
    "Rozłącznik izolacyjny": ["Prąd znamionowy [A]", "Układ biegunów"],
    "Przełącznik trójpołożeniowy": ["Pozycje przełącznika", "Układ biegunów"],
    "Sygnalizator obecności napięcia": ["Barwa lampek", "Ilość lampek"],
    "Stycznik modułowy": ["Konfiguracja styków", "Napięcie cewki [V]"],
    "Rozłącznik bezpiecznikowy": ["Rozmiar", "Układ biegunów"],
    "Wstawka redukcyjna": ["Kompatybilność", "Prąd znamionowy [A]"],
    "Podstawa bezpiecznikowa DC": ["Rozmiar", "Napięcie znamionowe DC [V]"],
}

H2_QUESTIONS = {
    "Wyłącznik nadprądowy MCB": "jak dobrać ten wariant do obwodu?",
    "Ogranicznik mocy": "które parametry sprawdzić?",
    "Wyłącznik różnicowoprądowy RCCB": "co sprawdzić przy doborze?",
    "Wyłącznik różnicowo-nadprądowy RCBO": "jak odczytać obie funkcje ochronne?",
    "Rozłącznik izolacyjny": "jak dobrać liczbę biegunów?",
    "Przełącznik trójpołożeniowy": "jak odczytać pozycje I-0-II?",
    "Sygnalizator obecności napięcia": "jaką sygnalizację zapewnia?",
    "Stycznik modułowy": "jak sprawdzić styki i cewkę?",
    "Rozłącznik bezpiecznikowy": "jak dobrać wkładkę?",
    "Wstawka redukcyjna": "z czym jest zgodna?",
    "Podstawa bezpiecznikowa DC": "jaką wkładkę dobrać?",
}


def read_copy() -> dict[str, tuple[str, str, str, str]]:
    entries = {}
    for line_number, line in enumerate(COPY.read_text(encoding="utf-8").splitlines(), 1):
        if not line or line.startswith("#"):
            continue
        parts = line.split("|")
        if len(parts) != 5:
            raise ValueError(f"Wrong field count on line {line_number}")
        code, *sections = [part.strip() for part in parts]
        if code in entries or not all(sections):
            raise ValueError(f"Duplicate or empty copy for {code}")
        entries[code] = tuple(sections)
    return entries


def faq_for_product(code: str, features: dict[str, str]) -> list[tuple[str, str]]:
    """Three SKU-specific purchase questions using confirmed feature values."""
    family = features["Typ produktu"]
    symbol = features["Symbol"]
    current = features.get("Prąd znamionowy [A]", "")
    poles = features.get("Układ biegunów", "")
    if family == "Wyłącznik nadprądowy MCB":
        curve = features["Charakterystyka wyzwalania"]
        icn = features["Znamionowa zdolność zwarciowa Icn [kA]"]
        return [
            (f"Jaką charakterystykę i prąd ma {symbol}?", f"Model {symbol} ma charakterystykę {curve} i prąd znamionowy {current} A."),
            (f"Jaki układ biegunów ma {symbol}?", f"W tym wykonaniu potwierdzono układ {poles}."),
            (f"Ile wynosi Icn dla kodu {code}?", f"Znamionowa zdolność zwarciowa Icn wynosi {icn} kA."),
        ]
    if family == "Ogranicznik mocy":
        return [
            (f"Jaki prąd znamionowy ma {symbol}?", f"Prąd znamionowy tego wykonania to {current} A."),
            (f"Ile biegunów ma model {symbol}?", f"Potwierdzony układ biegunów to {poles}."),
            (f"Jak rozpoznać wariant {symbol}?", f"Jego symbol to {symbol}, a kod producenta {code}."),
        ]
    if family == "Wyłącznik różnicowoprądowy RCCB":
        rcd = features["Prąd różnicowy IΔn [mA]"]
        rcd_type = features["Typ prądu różnicowego"]
        inc = features["Warunkowy prąd zwarciowy Inc [kA]"]
        return [
            (f"Jakie prąd i czułość ma {symbol}?", f"Prąd znamionowy wynosi {current} A, a prąd różnicowy IΔn {rcd} mA."),
            (f"Jaki typ i układ biegunów ma kod {code}?", f"To wykonanie typu {rcd_type} w układzie {poles}."),
            (f"Czy wartość {inc} kA oznacza samodzielną zdolność wyłączania zwarcia?", f"Nie. {inc} kA to warunkowy prąd zwarciowy Inc; dobór współpracującego zabezpieczenia nadprądowego wymaga osobnego sprawdzenia."),
        ]
    if family == "Wyłącznik różnicowo-nadprądowy RCBO":
        curve = features["Charakterystyka wyzwalania"]
        rcd = features["Prąd różnicowy IΔn [mA]"]
        rcd_type = features["Typ prądu różnicowego"]
        icn = features["Znamionowa zdolność zwarciowa Icn [kA]"]
        return [
            (f"Jaką charakterystykę i prąd ma {symbol}?", f"Część nadprądowa ma charakterystykę {curve} i prąd znamionowy {current} A."),
            (f"Jaki typ i czułość różnicową ma kod {code}?", f"Potwierdzono typ {rcd_type} oraz prąd różnicowy IΔn {rcd} mA."),
            (f"Jakie są układ biegunów i Icn modelu {symbol}?", f"Układ biegunów to {poles}, a znamionowa zdolność zwarciowa Icn wynosi {icn} kA."),
        ]
    if family == "Rozłącznik izolacyjny":
        voltage = features.get("Napięcie znamionowe AC [V]", "")
        third = (f"Potwierdzone napięcie znamionowe AC wynosi {voltage} V." if voltage else
                 f"Pełny symbol to {symbol}, a kod producenta to {code}.")
        return [
            (f"Jaki prąd ma {symbol}?", f"Prąd znamionowy tego rozłącznika wynosi {current} A."),
            (f"Ile torów można rozłączyć modelem {symbol}?", f"Model ma układ {poles}; liczbę torów trzeba dopasować do schematu instalacji."),
            (f"Jakie dodatkowe dane potwierdzono dla kodu {code}?", third),
        ]
    if family == "Przełącznik trójpołożeniowy":
        positions = features["Pozycje przełącznika"]
        voltage = features["Napięcie znamionowe AC [V]"]
        return [
            (f"Jakie pozycje ma {symbol}?", f"Dostępne pozycje przełącznika są oznaczone {positions}."),
            (f"Jaki prąd i układ biegunów ma kod {code}?", f"Prąd znamionowy wynosi {current} A, a układ biegunów to {poles}."),
            (f"Jakie napięcie AC potwierdzono dla {symbol}?", f"Napięcie znamionowe AC wynosi {voltage} V."),
        ]
    if family == "Sygnalizator obecności napięcia":
        color = features["Barwa lampek"]
        count = features["Ilość lampek"]
        voltage = features["Napięcie znamionowe AC [V]"]
        return [
            (f"Ile lampek ma {symbol} i w jakich barwach?", f"Model ma {count} lampki; potwierdzona barwa to {color.lower()}."),
            (f"Na jakie napięcie jest przeznaczony kod {code}?", f"Napięcie znamionowe wynosi {voltage} V AC."),
            (f"Czy wskazanie modelu {symbol} zastępuje pomiar braku napięcia?", f"Nie. Wskazanie {symbol} nie zastępuje potwierdzenia braku napięcia odpowiednim przyrządem przed pracą przy instalacji."),
        ]
    if family == "Stycznik modułowy":
        coil = features["Napięcie cewki [V]"]
        contacts = features["Konfiguracja styków"]
        modules = features["Liczba modułów"]
        return [
            (f"Jakiego napięcia cewki wymaga {symbol}?", f"Napięcie cewki tego wykonania wynosi {coil} V."),
            (f"Jakie styki ma kod {code}?", f"Potwierdzona konfiguracja styków to {contacts}."),
            (f"Jaki prąd i szerokość ma {symbol}?", f"Prąd znamionowy wynosi {current} A, a aparat zajmuje {modules} {'moduł' if modules == '1' else 'moduły' if modules in {'2', '3', '4'} else 'modułów'}."),
        ]
    if family == "Rozłącznik bezpiecznikowy":
        size = features["Rozmiar"]
        fuse = features["Rodzaj wkładki"]
        voltage = features.get("Napięcie znamionowe AC [V]", "")
        third = (f"Potwierdzone napięcie znamionowe AC wynosi {voltage} V." if voltage else
                 f"Symbol aparatu to {symbol}, a kod producenta {code}.")
        return [
            (f"Jaki format wkładki pasuje do {symbol}?", f"Potwierdzony rodzaj wkładki to {fuse}, w rozmiarze {size}."),
            (f"Jaki prąd i układ biegunów ma kod {code}?", f"Aparat ma prąd znamionowy {current} A i układ {poles}."),
            (f"Jakie inne dane potwierdzono dla {symbol}?", third),
        ]
    if family == "Wstawka redukcyjna":
        compatibility = features["Kompatybilność"]
        return [
            (f"Z czym jest zgodna wstawka {symbol}?", f"Potwierdzona kompatybilność to {compatibility}."),
            (f"Jaki prąd ma kod {code}?", f"Prąd znamionowy tej wstawki wynosi {current} A."),
            ("Jak odróżnić ten wariant przy zamówieniu?", f"Sprawdź symbol {symbol} i kod producenta {code}."),
        ]
    if family == "Podstawa bezpiecznikowa DC":
        size = features["Rozmiar"]
        fuse = features["Rodzaj wkładki"]
        voltage = features["Napięcie znamionowe DC [V]"]
        return [
            (f"Jakiej wkładki wymaga {symbol}?", f"Potwierdzono wkładkę {fuse} w rozmiarze {size}."),
            (f"Jaki układ i prąd ma kod {code}?", f"Podstawa ma układ {poles} oraz prąd znamionowy {current} A."),
            (f"Jakie napięcie DC ma {symbol}?", f"Napięcie znamionowe DC wynosi {voltage} V."),
        ]
    raise ValueError(f"No FAQ plan for {family}")


def render(name: str, code: str, features: dict[str, str], sections: tuple[str, str, str, str]) -> str:
    family = features["Typ produktu"]
    keyword = KEYWORDS[family]
    intro, body, benefit, close = sections
    if not name.startswith(keyword) or not body.startswith(keyword):
        raise ValueError(f"SEO phrase missing at opening for {code}")
    h2 = f"{keyword} {features['Symbol']} – {H2_QUESTIONS[family]}"
    esc = html.escape

    feature_bullets = [benefit]
    for key in KEY_FEATURES[family]:
        raw = str(features.get(key, "")).strip()
        if not raw:
            raise ValueError(f"Missing feature {key} for {code}")
        feature_bullets.append(f"{LABELS[key]}: {raw}{UNITS.get(key, '')}.")

    specs = [("Producent", "ETI"), ("Kod producenta", code)]
    for key, label in LABELS.items():
        raw = str(features.get(key, "")).strip()
        if raw and keyword.casefold() not in raw.casefold():
            specs.append((label, raw + UNITS.get(key, "")))

    faq = faq_for_product(code, features)
    schema = {
        "@context": "https://schema.org",
        "@type": "FAQPage",
        "mainEntity": [
            {"@type": "Question", "name": question,
             "acceptedAnswer": {"@type": "Answer", "text": answer}}
            for question, answer in faq
        ],
    }
    schema_json = json.dumps(schema, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    result = "\n".join([
        f"<p><strong>{esc(name)}</strong> {esc(intro)}</p>",
        f"<h2>{esc(h2)}</h2>",
        f"<p>{esc(body)}</p>",
        "<h3>Najważniejsze cechy</h3>",
        "<ul>" + "".join(f"<li>{esc(item)}</li>" for item in feature_bullets) + "</ul>",
        "<h3>Specyfikacja techniczna</h3>",
        "<ul>" + "".join(f"<li><strong>{esc(label)}:</strong> {esc(raw)}</li>" for label, raw in specs) + "</ul>",
        f"<p>{esc(close)}</p>",
        "<h3>FAQ – pytania i odpowiedzi</h3>",
        "\n".join(f"<h4>{esc(question)}</h4>\n<p>{esc(answer)}</p>" for question, answer in faq),
        f'<script type="application/ld+json">{schema_json}</script>',
    ])
    visible_html = result.split('<script type="application/ld+json">', 1)[0]
    if visible_html.casefold().count(keyword.casefold()) != 3:
        raise ValueError(f"SEO phrase repeated too often for {code}")
    return result


def main() -> None:
    copy = read_copy()
    workbook = load_workbook(SOURCE)
    sheet = workbook["Import"]
    columns = {cell.value: cell.column for cell in sheet[1]}
    workbook_codes = set()
    lengths = []
    for row_index in range(2, sheet.max_row + 1):
        sku = str(sheet.cell(row_index, columns["sku"]).value)
        code = sku.split("/")[0]
        workbook_codes.add(code)
        if code not in copy:
            raise ValueError(f"No individually written copy for {code}")
        name = str(sheet.cell(row_index, columns["name"]).value)
        features = json.loads(sheet.cell(row_index, columns["features"]).value)
        if "Seria" in features:
            raise ValueError(f"Unwanted Seria feature for {code}")
        description = render(name, code, features, copy[code])
        cell = sheet.cell(row_index, columns["description"], description)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        visible_html = description.split('<script type="application/ld+json">', 1)[0]
        visible = " ".join(re.sub(r"<[^>]*>", " ", visible_html).split())
        lengths.append(len(visible))
    if workbook_codes != set(copy) or len(workbook_codes) != 127:
        raise ValueError("Source and authored SKU sets differ")
    workbook.save(TARGET)
    print(f"Saved {TARGET}")
    print(f"products={len(workbook_codes)}; min_chars={min(lengths)}; max_chars={max(lengths)}; below_1200={sum(n < 1200 for n in lengths)}")


if __name__ == "__main__":
    main()
