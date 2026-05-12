from __future__ import annotations

import argparse
import hashlib
import json
import re
from html import escape
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, is_blank, normalize_header, read_products, write_products


DEFAULT_DESCRIPTION_COLUMN = "description_html"
DEFAULT_MIN_CHARS_NO_SPACES = 1500

FORBIDDEN_PHRASES = [
    "Dzięki temu klient",
    "zamiast porównywać",
    "Najważniejsze parametry tego wariantu",
    "W praktyce oznacza to",
    "ułatwia porównanie tego wariantu",
    "opis pomaga",
]

ATTRIBUTE_SOURCE_COLUMNS = {
    "Producent": ["Producent", "Marka", "manufacturer_name", "attr_producent"],
    "EAN": ["EAN", "ean", "EAN (GTIN)"],
    "Kod producenta": ["Kod", "SKU", "sku", "Kod producenta"],
    "Typ produktu": ["Typ produktu", "Rodzaj produktu", "Typ", "attr_typ"],
    "Seria": ["Seria", "Rodzina", "attr_seria"],
    "Model": ["Model", "Nazwa Kanlux", "attr_model"],
    "Moc [W]": ["Moc [W]", "Moc", "attr_moc"],
    "Napięcie [V]": ["Napięcie [V]", "Napięcie (V)", "attr_napiecie"],
    "Temperatura barwowa [K]": ["Temperatura barwowa [K]", "Temperatura barwowa", "Barwa CCT [K]", "attr_barwa"],
    "Barwa światła": ["Barwa światła", "Barwa - kategoria", "attr_barwa_zakres"],
    "Strumień świetlny [lm]": ["Strumień świetlny [lm]", "Strumień [lm]", "Jasność", "attr_strumien"],
    "Skuteczność świetlna": ["Skuteczność świetlna", "attr_lm_w"],
    "Stopień ochrony [IP]": ["Stopień ochrony [IP]", "Stopień ochrony IP", "Klasa IP", "attr_ip"],
    "Klasa ochronności": ["Klasa ochronności", "attr_ik"],
    "Kąt świecenia": ["Kąt świecenia", "Kąt", "attr_kat_swiecenia"],
    "Trzonek": ["Trzonek", "Rodzaj gwintu", "Gwint", "attr_gwint"],
    "Kolor": ["Kolor", "Kolor obudowy", "attr_kolor"],
    "Materiał": ["Materiał", "attr_material"],
    "Kształt": ["Kształt", "attr_ksztalt"],
    "Wymiary": ["Wymiary", "attr_wymiary"],
    "Długość": ["Długość", "attr_dlugosc"],
    "Szerokość": ["Szerokość", "attr_szerokosc"],
    "Wysokość": ["Wysokość", "attr_wysokosc"],
    "Średnica": ["Średnica", "attr_srednica"],
    "Czujnik ruchu": ["Czujnik ruchu", "Czujnik", "attr_czujnik"],
    "Liczba sztuk": ["Liczba sztuk", "attr_ilosc_sztuk"],
    "Gwarancja": ["Gwarancja", "attr_gwarancja"],
    "Sterowanie": ["Sterowanie", "attr_sterowanie"],
    "Pasuje do": ["Pasuje do", "attr_pasuje_do"],
}

TECHNICAL_ORDER = [
    "Producent",
    "Seria",
    "Model",
    "Kod producenta",
    "EAN",
    "Typ produktu",
    "Moc [W]",
    "Napięcie [V]",
    "Temperatura barwowa [K]",
    "Barwa światła",
    "Strumień świetlny [lm]",
    "Skuteczność świetlna",
    "Stopień ochrony [IP]",
    "Klasa ochronności",
    "Kąt świecenia",
    "Trzonek",
    "Czujnik ruchu",
    "Kolor",
    "Materiał",
    "Kształt",
    "Wymiary",
    "Długość",
    "Szerokość",
    "Wysokość",
    "Średnica",
    "Pasuje do",
    "Sterowanie",
    "Liczba sztuk",
    "Gwarancja",
]


def main() -> None:
    parser = argparse.ArgumentParser(description="Generuje naturalne opisy HTML SEO na podstawie nazw i atrybutow produktow.")
    parser.add_argument("--input", required=True, help="CSV/XLSX z produktami po AlkanInatorze.")
    parser.add_argument("--sheet", help="Arkusz XLSX.")
    parser.add_argument("--output", required=True, help="Plik wynikowy CSV/XLSX.")
    parser.add_argument("--description-column", default=DEFAULT_DESCRIPTION_COLUMN)
    parser.add_argument("--title-column", default="", help="Wymusza kolumne z nazwa produktu.")
    parser.add_argument("--min-words", type=int, default=0, help="Zgodnosc wsteczna; preferuj --min-chars-no-spaces.")
    parser.add_argument("--min-chars-no-spaces", type=int, default=DEFAULT_MIN_CHARS_NO_SPACES)
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    result = add_descriptions(df, args.description_column, args.title_column, args.min_chars_no_spaces)
    write_products(result, Path(args.output))

    print(f"OK: wczytano {len(df)} produktow")
    print(f"OK: zapisano {args.output}")


def add_descriptions(
    df: pd.DataFrame,
    description_column: str = DEFAULT_DESCRIPTION_COLUMN,
    title_column: str = "",
    min_chars_no_spaces: int = DEFAULT_MIN_CHARS_NO_SPACES,
) -> pd.DataFrame:
    result = df.copy()
    result[description_column] = [
        build_description_html(row, title_column, min_chars_no_spaces) for _, row in result.iterrows()
    ]
    return result


def build_description_html(
    row: pd.Series,
    title_column: str = "",
    min_chars_no_spaces: int = DEFAULT_MIN_CHARS_NO_SPACES,
) -> str:
    title = product_title(row, title_column)
    attrs = collect_attributes(row)
    family = classify_family(title, attrs)
    seed = stable_seed(first_present(row.get("sku", ""), row.get("SKU", ""), row.get("Kod", ""), title))
    keyword = keyword_from_title(title, attrs)

    intro = variant(INTRO_BUILDERS[family], seed)(keyword, attrs)
    benefits = choose_benefits(keyword, attrs, family, seed)
    advice = variant(ADVICE_BUILDERS[family], seed // 7)(attrs)
    technical_rows = build_technical_rows(attrs)

    parts = [
        f"<p>{escape(intro)}</p>",
        "",
        f"<h2>{escape(keyword)} - czym wyróżnia się ten produkt?</h2>",
        "",
        "<ul>",
        *[f"  <li>- {escape(item)}</li>" for item in benefits],
        "</ul>",
        "",
        f"<p>{escape(advice)}</p>",
        "",
        "<h3>Specyfikacja techniczna</h3>",
        "",
        "<ul>",
        *[f"  <li>- {escape(label)}: {escape(value)}</li>" for label, value in technical_rows],
        "</ul>",
    ]
    html = "\n".join(parts)
    html = ensure_minimum_length(html, keyword, attrs, family, seed, min_chars_no_spaces)
    validate_description(html)
    return html


def product_title(row: pd.Series, title_column: str = "") -> str:
    return first_present(
        row.get(title_column, "") if title_column else "",
        row.get("new_title", ""),
        row.get("name", ""),
        row.get("Nazwa B2C / SEO final", ""),
        row.get("Nazwa", ""),
        row.get("Nazwa B2C / SEO", ""),
        row.get("Nazwa Kanlux", ""),
        "Produkt",
    )


def keyword_from_title(title: str, attrs: dict[str, str]) -> str:
    producer = attrs.get("Producent", "")
    code = attrs.get("Kod producenta", "")
    keyword = compact_spaces(title)
    if code:
        keyword = re.sub(rf"\s+{re.escape(code)}\s*$", "", keyword, flags=re.IGNORECASE)
    if producer:
        keyword = re.sub(rf"\s+{re.escape(producer)}\s*$", "", keyword, flags=re.IGNORECASE)
    return compact_spaces(keyword)


def collect_attributes(row: pd.Series) -> dict[str, str]:
    attrs = read_features_json(row)
    for target, columns in ATTRIBUTE_SOURCE_COLUMNS.items():
        if attrs.get(target):
            attrs[target] = normalize_attribute_value(target, attrs[target])
            continue
        value = first_present(*(row.get(column, "") for column in columns))
        if value:
            attrs[target] = normalize_attribute_value(target, value)
    return {key: value for key, value in attrs.items() if value}


def read_features_json(row: pd.Series) -> dict[str, str]:
    raw = first_present(row.get("features", ""), row.get("Parametry", ""))
    if not raw:
        return {}
    try:
        data = json.loads(raw)
    except (TypeError, ValueError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {
        normalize_feature_name(str(key)): normalize_attribute_value(normalize_feature_name(str(key)), str(value))
        for key, value in data.items()
        if not is_blank(value)
    }


def normalize_feature_name(value: str) -> str:
    normalized = normalize_header(value)
    aliases = {
        "marka": "Producent",
        "rodzaj produktu": "Typ produktu",
        "moc": "Moc [W]",
        "napiecie v": "Napięcie [V]",
        "napięcie v": "Napięcie [V]",
        "temperatura barwowa": "Temperatura barwowa [K]",
        "jasnosc": "Strumień świetlny [lm]",
        "jasność": "Strumień świetlny [lm]",
        "stopien ochrony ip": "Stopień ochrony [IP]",
        "stopień ochrony ip": "Stopień ochrony [IP]",
        "rodzaj gwintu": "Trzonek",
        "barwa swiatla": "Barwa światła",
        "barwa światła": "Barwa światła",
    }
    return aliases.get(normalized, value)


def normalize_attribute_value(attr: str, value: Any) -> str:
    value = compact_spaces(str(value))
    if not value or value in {"nan", "None", "1", "MLS"}:
        return ""
    if attr == "Stopień ochrony [IP]":
        return re.sub(r"^IP(\d+)$", r"IP \1", value.upper().replace(" ", ""))
    if attr == "Moc [W]":
        return strip_unit(value, "W")
    if attr == "Temperatura barwowa [K]":
        value = value.replace(" ", "")
        if re.fullmatch(r"[0-9/]+K?", value, flags=re.IGNORECASE):
            return value.replace("K", "").replace("k", "")
        return re.sub(r"(?<=\d)\s*K\b", "", value, flags=re.IGNORECASE)
    if attr == "Strumień świetlny [lm]":
        return strip_unit(value, "lm")
    if attr == "Napięcie [V]":
        value = re.sub(r"\s*AC\b", "", value, flags=re.IGNORECASE)
        value = re.sub(r"\s*DC\b", "", value, flags=re.IGNORECASE)
        return strip_unit(value, "V")
    if attr == "Barwa światła":
        return {
            "Neutralna (3500-4500K)": "Neutralna",
            "Zimna (≥5000K)": "Zimna",
            "Ciepła (≤3000K)": "Ciepła",
            "Zmienna CCT": "Zmienna",
        }.get(value, value)
    if attr == "Trzonek":
        normalized = value.upper().replace(" ", "")
        match = re.fullmatch(r"\d+X(.+)", normalized)
        return match.group(1) if match else normalized
    if attr == "Kolor":
        return {"bialy": "Biały", "biały": "Biały"}.get(value, capitalize_first(value))
    if attr == "Materiał":
        return {
            "PC": "Tworzywo sztuczne",
            "PMMA": "Tworzywo sztuczne",
            "PS": "Tworzywo sztuczne",
            "PP": "Polipropylen",
            "tworzywo sztuczne": "Tworzywo sztuczne",
            "szkło": "Szkło",
            "szkło hartowane": "Szkło",
            "drewniana": "Drewno",
            "stop aluminium": "Metal",
            "stal": "Metal",
            "stal nierdzewna": "Metal",
        }.get(value, capitalize_first(value))
    return value


def strip_unit(value: str, unit: str) -> str:
    value = re.sub(r"(?<=\d)\.(?=\d)", ",", value)
    value = re.sub(rf"\s*{re.escape(unit)}\b", "", value, flags=re.IGNORECASE)
    return compact_spaces(value)


def classify_family(title: str, attrs: dict[str, str]) -> str:
    text = normalize_header(" ".join([title, attrs.get("Typ produktu", ""), attrs.get("Seria", "")]))
    if any(token in text for token in ["klosz", "zapinka", "wspornik", "siatka", "konektor", "pilot", "uchwyt", "soczewka", "klips", "linka", "czujnik", "lacznik", "łącznik"]):
        return "akcesorium"
    if "ramka" in text or "adtr" in text:
        return "ramka"
    if "zasilacz" in text or "driver" in text:
        return "zasilacz"
    if "high bay" in text or "hb ufo" in text:
        return "high_bay"
    if "naswietlacz" in text or "naświetlacz" in title.lower():
        return "naswietlacz"
    if "plafon" in text:
        return "plafon"
    if "panel" in text or "blingo" in text or "barev" in text:
        return "panel"
    if "hermetycz" in text or "dicht" in text or "mah" in text or "tp strong" in text:
        return "hermetyczna"
    if "oprawa sufitowa" in text or "oprawa punktowa" in text or "lampa wiszaca" in text or "lampa wisząca" in title.lower():
        return "oprawa"
    return "ogolny"


def stable_seed(value: str) -> int:
    digest = hashlib.sha1(value.encode("utf-8")).hexdigest()
    return int(digest[:8], 16)


def variant(items: list[Any], seed: int) -> Any:
    return items[seed % len(items)]


def choose_benefits(keyword: str, attrs: dict[str, str], family: str, seed: int) -> list[str]:
    candidates = BENEFIT_BUILDERS[family](keyword, attrs)
    start = seed % len(candidates)
    ordered = candidates[start:] + candidates[:start]
    return ordered[:4]


def dim(attrs: dict[str, str]) -> str:
    if attrs.get("Wymiary"):
        return attrs["Wymiary"]
    parts = []
    for label in ["Długość", "Szerokość", "Wysokość", "Średnica"]:
        if attrs.get(label):
            parts.append(f"{label.lower()} {attrs[label]}")
    return ", ".join(parts)


def parameter_list(attrs: dict[str, str]) -> str:
    return natural_join([attrs.get("Moc [W]", ""), attrs.get("Strumień świetlny [lm]", ""), attrs.get("Temperatura barwowa [K]", ""), attrs.get("Stopień ochrony [IP]", "")])


def light_sentence(attrs: dict[str, str], place: str) -> str:
    power = attrs.get("Moc [W]", "")
    flux = attrs.get("Strumień świetlny [lm]", "")
    temp = attrs.get("Temperatura barwowa [K]", "")
    parts = []
    if power and flux:
        parts.append(f"Moc {power}W i strumień {flux}lm pomagają doświetlić {place}")
    elif power:
        parts.append(f"Moc {power}W pozwala dobrać oprawę do wielkości i funkcji przestrzeni")
    elif flux:
        parts.append(f"Strumień {flux}lm ułatwia zaplanowanie ilości światła w pomieszczeniu")
    if temp:
        parts.append(f"Barwa {temp}K daje neutralne, czytelne światło" if temp == "4000" else f"Barwa {temp}K pozwala dopasować odbiór światła do zastosowania")
    return ". ".join(parts) + ("." if parts else "")


def ip_sentence(attrs: dict[str, str], context: str) -> str:
    ip = attrs.get("Stopień ochrony [IP]", "")
    return f"Stopień ochrony {ip} ma znaczenie {context}." if ip else ""


def producer_text(attrs: dict[str, str]) -> str:
    producer = attrs.get("Producent", "")
    return f" marki {producer}" if producer else ""


def intro_high_bay_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to oprawa przemysłowa LED{producer_text(attrs)} do hal, magazynów i wysokich pomieszczeń technicznych. {light_sentence(attrs, 'stanowiska pracy, regały i ciągi komunikacyjne')} {ip_sentence(attrs, 'w przestrzeniach, gdzie oprawa pracuje w pyle, wilgoci lub przy intensywnej eksploatacji')}"


def intro_high_bay_b(keyword: str, attrs: dict[str, str]) -> str:
    angle = attrs.get("Kąt świecenia", "")
    angle_text = f" Kąt świecenia {angle} pomaga prowadzić światło z większej wysokości." if angle else ""
    return f"{keyword} sprawdzi się tam, gdzie zwykła oprawa sufitowa nie daje wystarczającego zasięgu. {light_sentence(attrs, 'dużą powierzchnię roboczą')}{angle_text} {ip_sentence(attrs, 'przy montażu w wymagającym obiekcie przemysłowym')}"


def intro_high_bay_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest przeznaczona do mocnego, roboczego oświetlenia z dużej wysokości. {light_sentence(attrs, 'hale produkcyjne, magazyny i strefy techniczne')} W takim zastosowaniu liczy się stabilny strumień światła, odporna obudowa i czytelność przestrzeni przez wiele godzin pracy."


def intro_panel_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to panel LED{producer_text(attrs)} do biur, recepcji, sal sprzedaży i innych wnętrz wymagających równego światła. {light_sentence(attrs, 'stanowiska pracy i ciągi komunikacyjne')} Format {dim(attrs) or 'panelu'} warto dopasować do układu sufitu."


def intro_panel_b(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} dobrze zastępuje starsze oprawy świetlówkowe, gdy potrzebne jest spokojne światło na większej powierzchni. {light_sentence(attrs, 'pomieszczenie bez mocno punktowego efektu')} Konstrukcja panelu ułatwia utrzymanie uporządkowanego wyglądu sufitu."


def intro_panel_c(keyword: str, attrs: dict[str, str]) -> str:
    ip = attrs.get("Stopień ochrony [IP]", "")
    protection = f" Stopień ochrony {ip} warto uwzględnić przy wyborze miejsca montażu." if ip else ""
    return f"{keyword} jest praktycznym wyborem do wnętrz, w których liczy się komfort pracy i jednolite oświetlenie. {light_sentence(attrs, 'biuro, lokal usługowy albo korytarz')}{protection}"


def intro_hermetic_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to oprawa do garaży, warsztatów, zapleczy i korytarzy technicznych. {light_sentence(attrs, 'przestrzeń roboczą lub komunikacyjną')} {ip_sentence(attrs, 'w miejscach narażonych na kurz, wilgoć i codzienne zabrudzenia')}"


def intro_hermetic_b(keyword: str, attrs: dict[str, str]) -> str:
    material = attrs.get("Materiał", "")
    material_text = f" Obudowa z materiału {material} ułatwia codzienne użytkowanie w technicznych warunkach." if material else ""
    return f"{keyword} sprawdzi się tam, gdzie oświetlenie ma być proste, odporne i przewidywalne. {light_sentence(attrs, 'garaż, piwnicę, warsztat lub zaplecze')}{material_text}"


def intro_hermetic_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest przeznaczona do pomieszczeń, w których oprawa pracuje bliżej kurzu, wilgoci i intensywnej eksploatacji niż typowe oświetlenie dekoracyjne. {light_sentence(attrs, 'stanowiska robocze i przejścia techniczne')} {ip_sentence(attrs, 'przy doborze oprawy do takich warunków')}"


def intro_plafon_a(keyword: str, attrs: dict[str, str]) -> str:
    sensor = " z czujnikiem ruchu" if attrs.get("Czujnik ruchu") else ""
    return f"{keyword} to plafon{sensor}{producer_text(attrs)} do korytarzy, wejść, klatek schodowych i pomieszczeń pomocniczych. {light_sentence(attrs, 'przejścia oraz niewielkie strefy użytkowe')} {ip_sentence(attrs, 'przy montażu w miejscu bardziej narażonym na wilgoć lub zabrudzenia')}"


def intro_plafon_b(keyword: str, attrs: dict[str, str]) -> str:
    sensor = " Czujnik ruchu ogranicza potrzebę ręcznego włączania światła." if attrs.get("Czujnik ruchu") else ""
    return f"{keyword} sprawdzi się w miejscach, gdzie oświetlenie ma działać funkcjonalnie i bez zbędnej obsługi.{sensor} {light_sentence(attrs, 'wejście, korytarz lub pomieszczenie gospodarcze')}"


def intro_plafon_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest praktyczną oprawą do przestrzeni przejściowych, w których światło powinno włączać się szybko i równomiernie obejmować najważniejszą strefę. {light_sentence(attrs, 'codzienne przejścia i wejścia')} Wymiary {dim(attrs) or 'z danych technicznych'} pomagają ocenić, czy oprawa nie będzie zbyt masywna."


def intro_naswietlacz_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to naświetlacz LED{producer_text(attrs)} do elewacji, podjazdów, bram, placów i stref dostaw. {light_sentence(attrs, 'wybraną strefę na zewnątrz budynku')} {ip_sentence(attrs, 'przy pracy w deszczu, pyle i zmiennych warunkach pogodowych')}"


def intro_naswietlacz_b(keyword: str, attrs: dict[str, str]) -> str:
    angle = attrs.get("Kąt świecenia", "")
    angle_text = f" Kąt {angle} pomaga skierować światło na konkretną część posesji." if angle else ""
    return f"{keyword} pozwala doświetlić miejsce, w którym po zmroku potrzebna jest dobra widoczność. {light_sentence(attrs, 'podjazd, wejście, bramę lub bok budynku')}{angle_text}"


def intro_naswietlacz_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest dobrym wyborem, gdy światło ma wspierać orientację i pracę na zewnątrz, a nie pełnić wyłącznie funkcję dekoracyjną. {light_sentence(attrs, 'otwartą przestrzeń lub elewację')} {ip_sentence(attrs, 'w instalacji zewnętrznej')}"


def intro_ramka_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to element montażowy{producer_text(attrs)} przeznaczony do estetycznego osadzenia panelu LED. Format {dim(attrs) or 'podany w specyfikacji'} pozwala dopasować ramkę do oprawy i zachować równe wykończenie sufitu. To akcesorium montażowe, więc nie należy oceniać go przez parametry świetlne."


def intro_ramka_b(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} porządkuje montaż panelu LED w miejscu, gdzie potrzebne jest stabilne i estetyczne wykończenie. Wymiary {dim(attrs) or 'z danych technicznych'} warto sprawdzić razem z formatem panelu. Neutralny kolor ułatwia dopasowanie ramki do jasnego sufitu."


def intro_ramka_c(keyword: str, attrs: dict[str, str]) -> str:
    material = attrs.get("Materiał", "")
    material_text = f" Materiał {material} wspiera stabilne osadzenie oprawy." if material else ""
    return f"{keyword} pomaga zamontować panel LED poza typowym sufitem kasetonowym lub w miejscu wymagającym dodatkowego obramowania.{material_text} Format {dim(attrs) or 'ramki'} decyduje o zgodności z konkretną oprawą."


def intro_zasilacz_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to element zasilający przeznaczony do współpracy z kompatybilną oprawą. Moc {attrs.get('Moc [W]', 'z danych technicznych')}W trzeba dobrać do konkretnego panelu lub systemu. Przy takim produkcie najważniejsza jest zgodność elektryczna i miejsce montażu."


def intro_zasilacz_b(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} warto traktować jako część konkretnego zestawu oświetleniowego, a nie uniwersalny dodatek do dowolnej oprawy. Napięcie {attrs.get('Napięcie [V]', 'z danych technicznych')} oraz moc {attrs.get('Moc [W]', 'z danych technicznych')}W pomagają potwierdzić kompatybilność przed montażem."


def intro_zasilacz_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} odpowiada za zasilanie dopasowanej oprawy LED. Wymiary {dim(attrs) or 'podane w specyfikacji'} ułatwiają zaplanowanie miejsca w suficie lub obudowie. Dobór takiego elementu powinien wynikać z parametrów panelu, a nie wyłącznie z nazwy serii."


def intro_ogolny_a(keyword: str, attrs: dict[str, str]) -> str:
    params = parameter_list(attrs)
    return f"{keyword} to produkt{producer_text(attrs)} przeznaczony do instalacji, w której liczy się zgodność parametrów i pewne dopasowanie do miejsca montażu. {f'Parametry {params} pomagają ocenić zakres zastosowania.' if params else 'Dane techniczne warto porównać z wymaganiami konkretnej instalacji.'}"


def intro_oprawa_a(keyword: str, attrs: dict[str, str]) -> str:
    light = light_sentence(attrs, "wybraną strefę pomieszczenia")
    light_text = f" {light}" if light else ""
    return f"{keyword} to oprawa{producer_text(attrs)} do oświetlenia wnętrz, w których liczy się prosty montaż i dopasowanie do aranżacji.{light_text} Warto sprawdzić typ źródła światła, kolor i wymiary przed zakupem."


def intro_oprawa_b(keyword: str, attrs: dict[str, str]) -> str:
    socket = f" Trzonek {attrs['Trzonek']} pozwala dobrać kompatybilne źródło światła." if attrs.get("Trzonek") else ""
    return f"{keyword} sprawdzi się jako funkcjonalny punkt światła w pomieszczeniu mieszkalnym, usługowym lub komunikacyjnym.{socket} Kolor {attrs.get('Kolor', 'oprawy')} warto dopasować do sufitu, ścian i pozostałych elementów wyposażenia."


def intro_oprawa_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest wyborem do instalacji, w której ważne jest połączenie wyglądu oprawy z praktycznym światłem. Wymiary {dim(attrs) or 'z danych technicznych'} pomagają ocenić proporcje produktu w pomieszczeniu, a specyfikacja ułatwia dobranie właściwego źródła światła."


def intro_akcesorium_a(keyword: str, attrs: dict[str, str]) -> str:
    target = attrs.get("Pasuje do", "")
    target_text = ""
    if target:
        target_text = f" i jest przeznaczone {target}" if normalize_header(target).startswith("do ") else f" i jest przeznaczone do {target}"
    return f"{keyword} to akcesorium{producer_text(attrs)} do uzupełnienia lub serwisowania instalacji oświetleniowej{target_text}. W takim produkcie najważniejsze jest dopasowanie do właściwej oprawy, serii lub miejsca montażu, a nie parametry świetlne."


def intro_akcesorium_b(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} pomaga utrzymać kompletność zestawu oświetleniowego albo dopasować oprawę do konkretnego sposobu montażu. Przed zakupem warto porównać serię, kod produktu i wymiary z elementem, do którego akcesorium ma zostać zastosowane."


def intro_akcesorium_c(keyword: str, attrs: dict[str, str]) -> str:
    material = attrs.get("Materiał", "")
    material_text = f" Materiał {material} ma znaczenie dla trwałości i stabilności elementu." if material else ""
    return f"{keyword} jest elementem pomocniczym do instalacji oświetleniowej, przy którym kluczowa jest zgodność z konkretną oprawą lub systemem.{material_text} Warto sprawdzić dane techniczne przed montażem, zwłaszcza gdy produkt ma zastąpić zużyty albo brakujący element."


INTRO_BUILDERS = {
    "high_bay": [intro_high_bay_a, intro_high_bay_b, intro_high_bay_c],
    "panel": [intro_panel_a, intro_panel_b, intro_panel_c],
    "hermetyczna": [intro_hermetic_a, intro_hermetic_b, intro_hermetic_c],
    "plafon": [intro_plafon_a, intro_plafon_b, intro_plafon_c],
    "naswietlacz": [intro_naswietlacz_a, intro_naswietlacz_b, intro_naswietlacz_c],
    "ramka": [intro_ramka_a, intro_ramka_b, intro_ramka_c],
    "zasilacz": [intro_zasilacz_a, intro_zasilacz_b, intro_zasilacz_c],
    "akcesorium": [intro_akcesorium_a, intro_akcesorium_b, intro_akcesorium_c],
    "oprawa": [intro_oprawa_a, intro_oprawa_b, intro_oprawa_c],
    "ogolny": [intro_ogolny_a],
}


def benefit_common(keyword: str, attrs: dict[str, str]) -> list[str]:
    items = []
    if attrs.get("Stopień ochrony [IP]"):
        items.append(f"Stopień ochrony {attrs['Stopień ochrony [IP]']} pomaga dobrać produkt do warunków pracy w miejscu montażu.")
    if attrs.get("Kolor"):
        items.append(f"Kolor {attrs['Kolor']} ułatwia dopasowanie produktu do sufitu, ściany lub pozostałych elementów instalacji.")
    if attrs.get("Materiał"):
        items.append(f"Materiał {attrs['Materiał']} wpływa na trwałość i sposób użytkowania w codziennej eksploatacji.")
    if attrs.get("Gwarancja"):
        items.append(f"Gwarancja {attrs['Gwarancja']} jest dodatkową informacją przy wyborze produktu do częstej pracy.")
    return items


def benefits_high_bay(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} daje mocne światło robocze do wysokich pomieszczeń.",
        flux_benefit(attrs, "pomaga doświetlić stanowiska pracy i alejki magazynowe"),
        angle_benefit(attrs, "ułatwia zaplanowanie rozstawu opraw nad powierzchnią roboczą"),
        color_temp_benefit(attrs, "sprzyja czytelności detali podczas pracy"),
        *benefit_common(keyword, attrs),
    ]


def benefits_panel(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} daje szerokie, równomierne światło bez mocno punktowego efektu.",
        f"Moc {attrs.get('Moc [W]', 'panelu')}W pomaga dobrać jasność do funkcji pomieszczenia.",
        f"Format {dim(attrs) or 'panelu'} ułatwia dopasowanie do sufitu modułowego lub planowanego układu opraw.",
        color_temp_benefit(attrs, "dobrze pasuje do pracy biurowej, obsługi klienta i nauki"),
        *benefit_common(keyword, attrs),
    ]


def benefits_hermetic(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} zapewnia stabilne oświetlenie robocze w pomieszczeniach technicznych i gospodarczych.",
        power_benefit(attrs, "pomaga dobrać oprawę do garażu, warsztatu lub zaplecza"),
        f"Długość {attrs.get('Długość', dim(attrs) or 'z danych technicznych')} ułatwia planowanie rozmieszczenia punktów świetlnych.",
        f"Trwała konstrukcja sprawdza się w miejscach, w których oprawa jest narażona na codzienne zabrudzenia.",
        *benefit_common(keyword, attrs),
    ]


def benefits_plafon(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} zapewnia funkcjonalne światło w miejscach przejściowych i pomocniczych.",
        "Czujnik ruchu zwiększa wygodę w korytarzu, przy wejściu lub na klatce schodowej." if attrs.get("Czujnik ruchu") else "Prosta forma dobrze pasuje do korytarzy, wejść i pomieszczeń pomocniczych.",
        f"Wymiary {dim(attrs) or 'z danych technicznych'} pomagają ocenić, czy oprawa będzie dobrze wyglądała w miejscu montażu.",
        f"Trzonek {attrs['Trzonek']} pozwala dobrać kompatybilne źródło światła." if attrs.get("Trzonek") else "Konstrukcja oprawy ułatwia codzienne użytkowanie bez skomplikowanej obsługi.",
        *benefit_common(keyword, attrs),
    ]


def benefits_naswietlacz(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} skutecznie doświetla wybraną strefę na zewnątrz budynku.",
        flux_benefit(attrs, "pomaga prowadzić światło na podjazd, elewację lub bramę"),
        angle_benefit(attrs, "ułatwia skierowanie światła na konkretny fragment przestrzeni"),
        color_temp_benefit(attrs, "daje praktyczne światło do orientacji i pracy po zmroku"),
        *benefit_common(keyword, attrs),
    ]


def benefits_ramka(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} porządkuje montaż panelu LED i pozwala uzyskać estetyczne wykończenie sufitu.",
        f"Wymiary {dim(attrs) or 'ramki'} pomagają dobrać element do konkretnego formatu panelu.",
        "Biały kolor dobrze łączy się z jasnymi sufitami i nie odciąga uwagi od oświetlenia." if attrs.get("Kolor", "").lower() == "biały" else f"Kolor {attrs.get('Kolor', 'ramki')} warto dopasować do sufitu i oprawy.",
        f"Materiał {attrs.get('Materiał', 'konstrukcyjny')} poprawia stabilność osadzenia panelu.",
        "Ramka ułatwia montaż tam, gdzie panel nie trafia bezpośrednio do sufitu kasetonowego.",
    ]


def benefits_zasilacz(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} jest elementem potrzebnym do poprawnej pracy kompatybilnej oprawy LED.",
        f"Moc {attrs.get('Moc [W]', 'z danych technicznych')}W należy dobrać do konkretnego wariantu panelu lub oprawy.",
        f"Napięcie {attrs.get('Napięcie [V]', 'z danych technicznych')} pomaga potwierdzić zgodność elektryczną przed montażem.",
        f"Wymiary {dim(attrs) or 'z danych technicznych'} ułatwiają zaplanowanie miejsca na zasilacz.",
        "Dobry dobór elementu ogranicza ryzyko problemów z kompatybilnością zestawu.",
    ]


def benefits_oprawa(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} pozwala zbudować funkcjonalny punkt światła w wybranym pomieszczeniu.",
        f"Trzonek {attrs['Trzonek']} ułatwia dobór kompatybilnego źródła światła." if attrs.get("Trzonek") else "Specyfikacja oprawy pomaga dobrać właściwe źródło światła.",
        f"Kolor {attrs.get('Kolor', 'oprawy')} można dopasować do sufitu, ścian lub pozostałych elementów wyposażenia.",
        f"Wymiary {dim(attrs) or 'z danych technicznych'} warto porównać z miejscem montażu przed zakupem.",
        *benefit_common(keyword, attrs),
    ]


def benefits_akcesorium(keyword: str, attrs: dict[str, str]) -> list[str]:
    return [
        f"{keyword} pomaga uzupełnić instalację oświetleniową bez wymiany całej oprawy.",
        f"Zgodność z serią {attrs.get('Seria', 'produktu')} ułatwia dobór elementu do istniejącego zestawu.",
        f"Kod producenta {attrs.get('Kod producenta', 'z danych technicznych')} pozwala szybko sprawdzić, czy wybrany element odpowiada potrzebnej części.",
        f"Wymiary {dim(attrs) or 'z danych technicznych'} warto porównać z miejscem montażu przed zakupem.",
        *benefit_common(keyword, attrs),
    ]


def power_benefit(attrs: dict[str, str], text: str) -> str:
    power = attrs.get("Moc [W]", "")
    if power:
        return f"Moc {power}W {text}."
    return f"Dane techniczne pomagają ocenić produkt po porównaniu z miejscem montażu."


def flux_benefit(attrs: dict[str, str], text: str) -> str:
    flux = attrs.get("Strumień świetlny [lm]", "")
    if flux:
        return f"Strumień {flux}lm {text}."
    return f"Parametry świetlne {text} po dopasowaniu oprawy do miejsca montażu."


def angle_benefit(attrs: dict[str, str], text: str) -> str:
    angle = attrs.get("Kąt świecenia", "")
    if angle:
        return f"Kąt {angle} {text}."
    return f"Optyka oprawy {text}."


def color_temp_benefit(attrs: dict[str, str], text: str) -> str:
    temperature = attrs.get("Temperatura barwowa [K]", "")
    if temperature:
        return f"Barwa {temperature}K {text}."
    if attrs.get("Barwa światła"):
        return f"Barwa {attrs['Barwa światła'].lower()} {text}."
    return f"Dobrze dobrana barwa światła {text}."


BENEFIT_BUILDERS = {
    "high_bay": benefits_high_bay,
    "panel": benefits_panel,
    "hermetyczna": benefits_hermetic,
    "plafon": benefits_plafon,
    "naswietlacz": benefits_naswietlacz,
    "ramka": benefits_ramka,
    "zasilacz": benefits_zasilacz,
    "akcesorium": benefits_akcesorium,
    "oprawa": benefits_oprawa,
    "ogolny": lambda keyword, attrs: [
        f"{keyword} pomaga dopasować produkt do konkretnego miejsca montażu.",
        *benefit_common(keyword, attrs),
        "Specyfikacja techniczna ułatwia sprawdzenie zgodności z instalacją.",
    ],
}


def advice_high_bay_a(attrs: dict[str, str]) -> str:
    return f"W oświetleniu halowym znaczenie ma wysokość montażu, rozstaw opraw i kierunek świecenia. Seria {attrs.get('Seria', 'High Bay')} sprawdzi się w obiektach, gdzie światło pracuje długo i musi zachować powtarzalny efekt na całej powierzchni roboczej."


def advice_high_bay_b(attrs: dict[str, str]) -> str:
    return f"Przy wyborze warto zestawić strumień świetlny, kąt świecenia i stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')} z realnymi warunkami obiektu. Inne potrzeby ma stanowisko pracy, a inne szeroka alejka magazynowa."


def advice_panel_a(attrs: dict[str, str]) -> str:
    return f"Panel LED najlepiej sprawdza się tam, gdzie potrzebne jest spokojne światło na większej powierzchni. Przed zakupem warto sprawdzić format {dim(attrs) or 'panelu'}, oczekiwaną jasność i barwę światła, bo te parametry najmocniej wpływają na komfort pracy."


def advice_panel_b(attrs: dict[str, str]) -> str:
    return f"Seria {attrs.get('Seria', 'Kanlux')} może zastąpić starsze oprawy świetlówkowe albo uzupełnić nowy układ oświetlenia w suficie modułowym. Przy modernizacji warto porównać wymiary panelu z istniejącym sufitem i zaplanować liczbę punktów."


def advice_hermetic_a(attrs: dict[str, str]) -> str:
    return f"W oprawach technicznych liczy się nie tylko jasność, ale też odporność obudowy i łatwe utrzymanie czystości. Stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')} warto dobrać do poziomu kurzu, wilgoci i intensywności użytkowania."


def advice_hermetic_b(attrs: dict[str, str]) -> str:
    return f"Przed montażem warto porównać długość, sposób zasilania i materiał obudowy z warunkami konkretnego pomieszczenia. W garażu, warsztacie lub zapleczu oprawa ma przede wszystkim działać pewnie i nie wymagać częstej obsługi."


def advice_plafon_a(attrs: dict[str, str]) -> str:
    return f"Plafon z czujnikiem sprawdza się w miejscach, gdzie światło ma włączać się tylko wtedy, gdy ktoś faktycznie korzysta z przestrzeni. Przy montażu warto uwzględnić zasięg czujnika, średnicę oprawy i stopień ochrony {attrs.get('Stopień ochrony [IP]', 'z danych technicznych')}."


def advice_plafon_b(attrs: dict[str, str]) -> str:
    return "W korytarzach i przejściach ważne jest proste działanie oraz równomierne oświetlenie najczęściej używanej strefy. Dlatego warto dobrać oprawę do wysokości montażu, wielkości pomieszczenia i oczekiwanego sposobu sterowania."


def advice_naswietlacz_a(attrs: dict[str, str]) -> str:
    return f"Naświetlacz powinien być dobrany do szerokości strefy, wysokości montażu i kierunku, w którym ma świecić. W praktyce największe znaczenie mają strumień świetlny, kąt świecenia i stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')}."


def advice_naswietlacz_b(attrs: dict[str, str]) -> str:
    return "Przy oświetleniu zewnętrznym lepiej dobrać światło do konkretnej strefy niż kierować je przypadkowo na całą przestrzeń. Pozwala to poprawić widoczność przy wejściu, bramie lub podjeździe bez nadmiernego rozproszenia światła."


def advice_ramka_a(attrs: dict[str, str]) -> str:
    return f"Przy takim akcesorium najważniejsze jest zgodne dopasowanie do panelu i sposób wykończenia sufitu. Wysokość oraz format {dim(attrs) or 'podane w specyfikacji'} warto sprawdzić przed zakupem razem z wymiarem oprawy."


def advice_ramka_b(attrs: dict[str, str]) -> str:
    return "Dobrze dobrana ramka ułatwia późniejszy dostęp serwisowy i pozwala zachować spójny wygląd kilku paneli zamontowanych w jednym pomieszczeniu. To szczególnie ważne w biurach, lokalach usługowych i korytarzach."


def advice_zasilacz_a(attrs: dict[str, str]) -> str:
    return "Zasilacz warto traktować jako część konkretnego systemu, a nie uniwersalny dodatek do dowolnej oprawy. Przed montażem trzeba porównać moc, napięcie i typ współpracującego panelu."


def advice_zasilacz_b(attrs: dict[str, str]) -> str:
    return "Dobrze dobrany element zasilający ułatwia późniejszą obsługę instalacji i ogranicza ryzyko problemów z kompatybilnością. Znaczenie ma także miejsce montażu oraz dostęp do elementu w razie serwisu."


def advice_ogolny_a(attrs: dict[str, str]) -> str:
    return "Przed zakupem warto porównać dane techniczne z miejscem montażu i oczekiwanym sposobem użytkowania. Taki dobór ogranicza ryzyko zakupu produktu, który pasuje nazwą, ale nie pasuje parametrami do instalacji."


def advice_oprawa_a(attrs: dict[str, str]) -> str:
    return "Przy oprawach sufitowych i punktowych warto zacząć od miejsca montażu, oczekiwanego efektu światła oraz typu źródła. Dopiero potem dobrze porównać kolor, wymiary i sposób mocowania z wystrojem pomieszczenia."


def advice_oprawa_b(attrs: dict[str, str]) -> str:
    return "Jeśli oprawa ma być widoczna we wnętrzu, znaczenie ma nie tylko parametr techniczny, ale też proporcja produktu do sufitu i pozostałych elementów wyposażenia. W praktyce warto sprawdzić wymiary, trzonek i dopuszczalną moc źródła światła."


def advice_akcesorium_a(attrs: dict[str, str]) -> str:
    return "Przy akcesoriach najważniejsza jest kompatybilność z konkretną oprawą, serią lub sposobem montażu. Warto sprawdzić kod producenta, nazwę serii i wymiary, ponieważ podobnie wyglądające elementy nie zawsze są zamienne."


def advice_akcesorium_b(attrs: dict[str, str]) -> str:
    return "Taki element najlepiej dobierać do istniejącej instalacji, a nie wyłącznie do ogólnej kategorii produktu. Jeśli akcesorium ma zastąpić uszkodzoną część, pomocne będzie porównanie oznaczeń na starej części z danymi w specyfikacji."


ADVICE_BUILDERS = {
    "high_bay": [advice_high_bay_a, advice_high_bay_b],
    "panel": [advice_panel_a, advice_panel_b],
    "hermetyczna": [advice_hermetic_a, advice_hermetic_b],
    "plafon": [advice_plafon_a, advice_plafon_b],
    "naswietlacz": [advice_naswietlacz_a, advice_naswietlacz_b],
    "ramka": [advice_ramka_a, advice_ramka_b],
    "zasilacz": [advice_zasilacz_a, advice_zasilacz_b],
    "akcesorium": [advice_akcesorium_a, advice_akcesorium_b],
    "oprawa": [advice_oprawa_a, advice_oprawa_b],
    "ogolny": [advice_ogolny_a],
}


def build_technical_rows(attrs: dict[str, str]) -> list[tuple[str, str]]:
    rows = []
    for key in TECHNICAL_ORDER:
        value = attrs.get(key, "")
        if value:
            rows.append((key, value))
    return rows


def ensure_minimum_length(
    html: str,
    keyword: str,
    attrs: dict[str, str],
    family: str,
    seed: int,
    min_chars_no_spaces: int,
) -> str:
    additions = [
        extra_application(keyword, attrs, family),
        extra_selection_advice(attrs, family),
        extra_mounting_note(attrs, family),
    ]
    index = seed % len(additions)
    used = 0
    while chars_without_spaces(strip_html(html)) < min_chars_no_spaces and used < len(additions):
        text = additions[(index + used) % len(additions)]
        html = html.replace("<h3>Specyfikacja techniczna</h3>", f"<p>{escape(text)}</p>\n\n<h3>Specyfikacja techniczna</h3>", 1)
        used += 1
    if chars_without_spaces(strip_html(html)) < min_chars_no_spaces:
        fallback = (
            "Opis warto odczytywać razem ze specyfikacją techniczną, ponieważ to ona potwierdza zgodność produktu "
            "z miejscem montażu, sposobem zasilania i oczekiwanym efektem użytkowym."
        )
        html = html.replace("<h3>Specyfikacja techniczna</h3>", f"<p>{escape(fallback)}</p>\n\n<h3>Specyfikacja techniczna</h3>", 1)
    if chars_without_spaces(strip_html(html)) < min_chars_no_spaces:
        final_note = (
            "Przy produktach technicznych warto zachować zgodność z oznaczeniem producenta, ponieważ podobne nazwy "
            "mogą dotyczyć różnych wymiarów, sposobów mocowania lub parametrów pracy."
        )
        html = html.replace("<h3>Specyfikacja techniczna</h3>", f"<p>{escape(final_note)}</p>\n\n<h3>Specyfikacja techniczna</h3>", 1)
    return html


def extra_application(keyword: str, attrs: dict[str, str], family: str) -> str:
    if family == "ramka":
        return f"{keyword} jest szczególnie przydatna wtedy, gdy panel LED ma zostać zamontowany poza standardowym sufitem modułowym. Ramka pomaga utrzymać równą linię montażu i ogranicza wrażenie przypadkowego dołożenia oprawy."
    if family == "zasilacz":
        return "W przypadku zasilacza kluczowa jest zgodność z oprawą, a nie wygląd elementu. Warto zachować dostęp do miejsca montażu, bo ułatwia to diagnostykę i ewentualną wymianę w przyszłości."
    if family == "akcesorium":
        return f"{keyword} warto dobrać do konkretnej oprawy, a nie tylko do ogólnej kategorii produktu. Taki element może decydować o stabilności montażu, ochronie oprawy albo wygodzie późniejszego serwisu."
    if family == "oprawa":
        return "W oprawach widocznych we wnętrzu warto połączyć wymagania techniczne z wyglądem produktu. Znaczenie ma miejsce montażu, typ źródła światła, kolor obudowy i to, czy oprawa ma być dyskretna, czy wyraźnie widoczna."
    if family == "high_bay":
        return "Wysokie pomieszczenia wymagają opraw, które zachowują użyteczny strumień światła po dotarciu do powierzchni roboczej. Dlatego przy takim produkcie liczy się nie tylko moc, ale też optyka i odporność obudowy."
    if family == "panel":
        return "W pomieszczeniach biurowych i usługowych równomierność światła często jest ważniejsza niż dekoracyjna forma oprawy. Panel pomaga uzyskać spokojny efekt wizualny i ograniczyć kontrasty na suficie."
    if family == "plafon":
        return "W przestrzeniach przejściowych oprawa powinna być prosta w obsłudze i odporna na częste krótkie cykle pracy. Dlatego znaczenie ma zarówno sposób sterowania, jak i dopasowanie średnicy do pomieszczenia."
    if family == "naswietlacz":
        return "Na zewnątrz warto ustawić oprawę tak, aby oświetlała realnie używaną strefę, a nie przypadkowo świeciła w okna lub poza teren. Dobrze dobrany kąt i strumień światła poprawiają komfort po zmroku."
    return "Dane techniczne warto traktować jako podstawę doboru, ale decyzję najlepiej odnieść do realnego miejsca montażu, sposobu użytkowania i warunków pracy."


def extra_selection_advice(attrs: dict[str, str], family: str) -> str:
    if family in {"panel", "high_bay", "hermetyczna", "naswietlacz"}:
        return f"Jeżeli produkt ma pracować przez wiele godzin dziennie, warto zwrócić uwagę na strumień świetlny, barwę światła oraz stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')}. Te parametry najmocniej wpływają na wygodę użytkowania i trwałość instalacji."
    if family == "ramka":
        return "Przed zakupem ramki najlepiej porównać jej format z wymiarem panelu oraz sprawdzić wysokość profilu. Dzięki temu montaż będzie wyglądał spójnie także przy kilku oprawach w jednym pomieszczeniu."
    if family == "zasilacz":
        return "Przed zakupem warto sprawdzić moc, napięcie i serię oprawy. Zasilacz o podobnej nazwie nie zawsze oznacza pełną zgodność z konkretnym panelem."
    if family == "akcesorium":
        return "Przy częściach pomocniczych najlepiej porównać serię, kod producenta, kształt oraz sposób mocowania. To ogranicza ryzyko zakupu elementu, który wygląda podobnie, ale nie pasuje do danej oprawy."
    if family == "oprawa":
        return "Przed wyborem dobrze sprawdzić, czy oprawa pasuje do planowanego źródła światła i czy jej wymiary są odpowiednie do danego sufitu lub ściany. Ma to znaczenie zarówno dla montażu, jak i końcowego efektu wizualnego."
    return "Przy wyborze dobrze sprawdzić nie tylko pojedynczy parametr, ale cały zestaw danych technicznych, bo dopiero wtedy widać, czy produkt pasuje do instalacji."


def extra_mounting_note(attrs: dict[str, str], family: str) -> str:
    dimensions = dim(attrs)
    if dimensions:
        return f"Wymiary {dimensions} pomagają zaplanować miejsce montażu i uniknąć kolizji z innymi elementami instalacji. Ma to znaczenie szczególnie wtedy, gdy oprawa trafia do istniejącej zabudowy."
    return "Przed montażem warto sprawdzić miejsce instalacji, sposób zasilania i dostęp serwisowy. Takie podejście ogranicza ryzyko późniejszych poprawek."


def validate_description(html: str) -> None:
    for phrase in FORBIDDEN_PHRASES:
        if phrase in html:
            raise ValueError(f"Opis zawiera zakazana fraze: {phrase}")


def natural_join(values: list[str]) -> str:
    cleaned = [value for value in unique([compact_spaces(value) for value in values]) if value]
    if len(cleaned) <= 1:
        return "".join(cleaned)
    if len(cleaned) == 2:
        return f"{cleaned[0]} i {cleaned[1]}"
    return f"{', '.join(cleaned[:-1])} i {cleaned[-1]}"


def first_present(*values: Any) -> str:
    for value in values:
        if not is_blank(value):
            return compact_spaces(str(value))
    return ""


def unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = value.lower()
        if key not in seen:
            seen.add(key)
            result.append(value)
    return result


def strip_html(value: str) -> str:
    return re.sub(r"<[^>]+>", " ", value)


def chars_without_spaces(value: str) -> int:
    return len(re.sub(r"\s+", "", value))


def capitalize_first(value: str) -> str:
    return value[:1].upper() + value[1:] if value else value


if __name__ == "__main__":
    main()
