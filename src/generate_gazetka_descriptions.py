from __future__ import annotations

import argparse
import re
from html import escape
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, is_blank, read_products, strip_accents, write_products


DESCRIPTION_COLUMN = "Opis HTML"
MIN_CHARS_NO_SPACES = 1500


TECHNICAL_ROWS = [
    ("Producent", "attr_producent"),
    ("Kod producenta", "Kod Producenta"),
    ("EAN", "EAN"),
    ("Typ produktu", "attr_typ"),
    ("Seria", "attr_seria"),
    ("Model", "attr_model"),
    ("Moc", "attr_moc"),
    ("Napięcie", "attr_napiecie"),
    ("Prąd znamionowy", "attr_prad"),
    ("Czułość", "attr_czulosc"),
    ("Zdolność zwarciowa", "attr_zwarciowa_zdolnosc"),
    ("Temperatura barwowa", "attr_barwa"),
    ("Strumień świetlny", "attr_strumien"),
    ("Skuteczność świetlna", "attr_lm_w"),
    ("Stopień ochrony", "attr_ip"),
    ("Klasa odporności", "attr_ik"),
    ("Kąt świecenia", "attr_kat_swiecenia"),
    ("Trzonek", "attr_gwint"),
    ("Czujnik", "attr_czujnik"),
    ("Kolor", "attr_kolor"),
    ("Materiał", "attr_material"),
    ("Wymiary", "attr_wymiary"),
    ("Długość", "attr_dlugosc"),
    ("Szerokość", "attr_szerokosc"),
    ("Wysokość", "attr_wysokosc"),
    ("Średnica", "attr_srednica"),
    ("Liczba biegunów", "attr_liczba_biegunow"),
    ("Liczba modułów", "attr_liczba_modulow"),
    ("Długość przewodu", "attr_dlugosc_przewodu"),
    ("Pojemność", "attr_pojemnosc"),
    ("Gwarancja", "attr_gwarancja"),
]


def main() -> None:
    parser = argparse.ArgumentParser(description="Generuje opisy HTML dla pliku Gazetka produkty.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--title-column", default="new_title")
    parser.add_argument("--description-column", default=DESCRIPTION_COLUMN)
    args = parser.parse_args()

    df = read_products(args.input)
    result = add_descriptions(df, args.title_column, args.description_column)
    write_products(result, Path(args.output))
    print(f"OK: wczytano {len(df)} produktow")
    print(f"OK: zapisano {args.output}")


def add_descriptions(df: pd.DataFrame, title_column: str, description_column: str) -> pd.DataFrame:
    result = df.copy()
    result[description_column] = [build_description(row, title_column) for _, row in result.iterrows()]
    return result


def build_description(row: pd.Series, title_column: str) -> str:
    title = value(row, title_column) or value(row, "new_title") or value(row, "Nazwa produktu")
    producer = value(row, "attr_producent") or value(row, "Producent")
    family = classify_product(title, row)
    keyword = description_keyword(title, row)
    intro = build_intro(title, producer, family, row)
    benefits = build_benefits(family, row)
    benefits = compact_list([first_keyword_benefit(keyword, family, row), *benefits])[:6]
    usage = build_usage_note(family, row)
    technical = technical_rows(row)

    parts = [
        intro_with_strong_title(title, intro),
        "",
        f"<h2>Dlaczego warto wybrać {escape(keyword)}?</h2>",
        "",
        "<ul>",
        *[f"  <li>{escape(item)}</li>" for item in benefits],
        "</ul>",
        "",
        f"<p>{escape(usage)}</p>",
        "",
        f"<p>{escape(build_context_note(family, row))}</p>",
        "",
        "<h3>Specyfikacja techniczna</h3>",
        "",
        "<ul>",
        *[f"  <li><strong>{escape(label)}:</strong> {escape(item)}</li>" for label, item in technical],
        "</ul>",
    ]
    return ensure_minimum_length("\n".join(parts), family, row)


def intro_with_strong_title(title: str, intro: str) -> str:
    if intro.startswith(title):
        return f"<p><strong>{escape(title)}</strong>{escape(intro[len(title):])}</p>"
    return f"<p><strong>{escape(title)}</strong> {escape(intro)}</p>"


def description_keyword(title: str, row: pd.Series) -> str:
    keyword = title
    for suffix in [value(row, "Kod Producenta"), value(row, "attr_producent"), value(row, "Producent")]:
        if suffix:
            keyword = re.sub(rf"\s+{re.escape(suffix)}\s*$", "", keyword, flags=re.IGNORECASE)
    return compact_spaces(keyword) or title


def first_keyword_benefit(keyword: str, family: str, row: pd.Series) -> str:
    if family == "lighting":
        return f"{keyword} daje konkretny zestaw parametrów świetlnych, który można dopasować do miejsca montażu i oczekiwanego efektu."
    if family == "distribution_box":
        return f"{keyword} porządkuje montaż aparatury modułowej i ułatwia późniejszą obsługę instalacji."
    if family == "protection":
        return f"{keyword} wspiera ochronę obwodów i pozwala dobrać aparat do parametrów instalacji."
    if family == "socket":
        return f"{keyword} ułatwia podłączenie kilku urządzeń w jednym miejscu bez dokładania przypadkowych adapterów."
    if family == "installation":
        return f"{keyword} pomaga prowadzić instalację w uporządkowany sposób i dopasować osprzęt do podłoża."
    if family == "connector":
        return f"{keyword} pozwala wykonać czytelne połączenie lub rozdział przewodów w instalacji."
    if family == "tools":
        return f"{keyword} ułatwia wykonanie prac instalacyjnych, serwisowych lub montażowych."
    if family == "earthing":
        return f"{keyword} pomaga wykonać trwały element instalacji odgromowej albo uziemiającej."
    return f"{keyword} pomaga dobrać produkt do konkretnego zastosowania i parametrów instalacji."


def build_context_note(family: str, row: pd.Series) -> str:
    if family == "lighting":
        return "W praktyce warto dobrać oprawę nie tylko po mocy, ale też po strumieniu świetlnym, barwie, kącie świecenia i odporności obudowy. Dzięki temu produkt będzie pasował do realnego miejsca montażu: elewacji, garażu, warsztatu, ogrodu, korytarza albo stanowiska roboczego."
    if family == "distribution_box":
        return "Przy obudowach i rozdzielnicach ważne jest zostawienie miejsca na aparaturę, przewody oraz późniejszy dostęp serwisowy. Zbyt mała obudowa utrudnia uporządkowanie instalacji, dlatego przed zakupem warto sprawdzić liczbę modułów, wymiary i sposób montażu."
    if family == "protection":
        return "Aparatura zabezpieczająca powinna być dobierana do projektu instalacji, przewidywanego obciążenia oraz układu sieci. Parametry takie jak prąd znamionowy, liczba biegunów, charakterystyka i zdolność zwarciowa mają bezpośredni wpływ na poprawne działanie zabezpieczenia."
    if family == "socket":
        return "W listwach, gniazdach i przedłużaczach liczy się nie tylko liczba punktów zasilania, ale też długość przewodu, typ gniazd, obecność zabezpieczeń i wygodny dostęp do portów. Dobrze dobrany osprzęt pozwala uporządkować zasilanie przy biurku, stanowisku roboczym albo w domowym warsztacie."
    if family == "installation":
        return "Elementy prowadzenia instalacji powinny być dopasowane do średnicy przewodów, sposobu montażu i środowiska pracy. Przy rurach i puszkach znaczenie ma materiał, odporność mechaniczna, kolor oraz możliwość użycia wewnątrz lub na zewnątrz budynku."
    if family == "connector":
        return "Przy złączkach i blokach rozdzielczych najważniejsze jest dopasowanie przekroju przewodów, materiału żył oraz dopuszczalnego napięcia i prądu. Warto też zwrócić uwagę na liczbę torów, kolorystykę oraz sposób montażu w rozdzielnicy."
    if family == "tools":
        return "W narzędziach instalacyjnych liczy się zakres pracy, pewny chwyt i odporność na intensywne użytkowanie. Przy narzędziach VDE oraz sprzęcie akumulatorowym szczególnie ważne jest porównanie oznaczeń producenta z typem wykonywanych prac."
    if family == "earthing":
        return "W instalacjach odgromowych i uziemiających element musi być dopasowany do średnicy drutu, sposobu prowadzenia oraz miejsca montażu. Trwałość połączenia i stabilność mocowania są ważniejsze niż sama zgodność ogólnej kategorii produktu."
    return "Przed zakupem warto sprawdzić nie tylko nazwę produktu, ale też kod producenta i parametry techniczne. Podobne produkty mogą różnić się wymiarem, sposobem montażu, materiałem albo zakresem dopuszczalnej pracy."


def ensure_minimum_length(html: str, family: str, row: pd.Series) -> str:
    result = html
    for note in [
        build_selection_note(family, row),
        build_mounting_note(family, row),
        build_difference_note(family, row),
        build_final_note(family, row),
    ]:
        if chars_without_spaces(strip_html(result)) >= MIN_CHARS_NO_SPACES:
            break
        result = result.replace(
            "<h3>Specyfikacja techniczna</h3>",
            f"<p>{escape(note)}</p>\n\n<h3>Specyfikacja techniczna</h3>",
            1,
        )
    return result


def build_final_note(family: str, row: pd.Series) -> str:
    producer = value(row, "attr_producent") or value(row, "Producent")
    producer_text = f" producenta {producer}" if producer else ""
    if family == "protection":
        return f"W razie wymiany aparatu warto porównać oznaczenia ze starego elementu z kodem i parametrami nowego produktu{producer_text}. Pozwala to zachować zgodność techniczną bez opierania decyzji wyłącznie na podobnej nazwie."
    if family == "installation":
        return f"Przy zakupie elementów instalacyjnych dobrze jest sprawdzić także opakowanie, długość odcinka i sposób prowadzenia przewodów. Te informacje często decydują o wygodzie pracy bardziej niż sama nazwa produktu{producer_text}."
    return f"Jeżeli produkt ma trafić do istniejącej instalacji, przed zakupem warto porównać jego kod, wymiary i parametry z elementem wymienianym lub z dokumentacją techniczną{producer_text}."


def build_selection_note(family: str, row: pd.Series) -> str:
    if family == "lighting":
        return "Jeżeli produkt ma pracować długo lub w trudniejszych warunkach, warto zwrócić uwagę na parametry eksploatacyjne, a nie tylko na nazwę serii. Wariant o podobnej mocy może dawać inny strumień świetlny, mieć inną barwę albo inny stopień ochrony."
    if family == "socket":
        return "Przy osprzęcie zasilającym dobrze jest od razu zaplanować liczbę podłączanych urządzeń i miejsce ustawienia produktu. To pomaga uniknąć sytuacji, w której przewód jest za krótki albo liczba gniazd okazuje się niewystarczająca."
    if family == "tools":
        return "Dobierając narzędzie, warto porównać zakres pracy z faktycznymi zadaniami na budowie lub w serwisie. Inny zestaw sprawdzi się przy drobnej obsłudze rozdzielnicy, a inny przy pracy z przewodami o większych przekrojach."
    if family == "connector":
        return "W połączeniach przewodów nie warto kierować się wyłącznie wyglądem elementu. Kluczowe jest to, czy złączka pasuje do przekroju, rodzaju przewodu oraz warunków pracy w danej rozdzielnicy."
    return "Najbezpieczniej dobierać produkt po komplecie danych: kodzie producenta, wymiarach, materiale i parametrach pracy. To ogranicza ryzyko zamiany z podobnym wariantem."


def build_mounting_note(family: str, row: pd.Series) -> str:
    if family == "tools":
        return "Przy narzędziach warto sprawdzić nie tylko pojedynczy parametr, ale też wygodę pracy, zakres zastosowania i sposób przechowywania. To szczególnie ważne, gdy produkt ma być używany często, a nie tylko okazjonalnie."
    if family == "socket":
        return "Przed ustawieniem lub podłączeniem warto sprawdzić miejsce pracy, dostęp do gniazd oraz sposób prowadzenia przewodu. Dobrze dobrane położenie osprzętu poprawia wygodę i ogranicza plątaninę kabli."
    dimensions = value(row, "attr_wymiary") or value(row, "attr_srednica") or value(row, "attr_dlugosc")
    if dimensions:
        return f"Wymiar {dimensions} warto porównać z miejscem montażu przed zakupem. Ma to znaczenie szczególnie przy wymianie istniejącego elementu albo przy montażu w ograniczonej przestrzeni."
    if family == "protection":
        return "Przed montażem aparatu należy sprawdzić zgodność z rozdzielnicą, przekrojami przewodów i wymaganiami zabezpieczanego obwodu. Pozwala to zachować czytelny układ oraz poprawną pracę instalacji."
    return "Przed montażem dobrze jest sprawdzić miejsce instalacji, sposób mocowania i dostęp serwisowy. Takie podejście zmniejsza ryzyko poprawek po rozpoczęciu prac."


def build_difference_note(family: str, row: pd.Series) -> str:
    code = value(row, "Kod Producenta")
    code_text = f" Kod producenta {code} pomaga odróżnić ten wariant od podobnych produktów." if code else ""
    if family == "lighting":
        return "Podobne oprawy mogą różnić się strumieniem świetlnym, temperaturą barwową, kątem świecenia albo odpornością na wilgoć i pył." + code_text
    if family == "installation":
        return "W elementach instalacyjnych drobna różnica średnicy, koloru lub materiału może oznaczać inne zastosowanie w praktyce." + code_text
    if family == "connector":
        return "W złączkach i blokach rozdzielczych różnice w przekroju przewodu oraz liczbie torów mają bezpośrednie znaczenie dla doboru." + code_text
    return "Warto porównać produkt z alternatywnymi wariantami tej samej kategorii, ponieważ różnice często dotyczą szczegółów technicznych, a nie samej nazwy." + code_text


def classify_product(title: str, row: pd.Series) -> str:
    text = normalized(" ".join([title, value(row, "attr_typ"), value(row, "Opis")]))
    if any(term in text for term in ["panel led", "plafoniera", "oprawa", "naswietlacz", "lampa ogrodowa", "latarka", "zarowka led"]):
        return "lighting"
    if any(term in text for term in ["rozdzielnica", "obudowa rozdzielnicy"]):
        return "distribution_box"
    if any(term in text for term in ["wyłącznik", "wylacznik", "rcbo", "kzs", "ogranicznik przepiec", "ogranicznik przepięć", "spd"]):
        return "protection"
    if any(term in text for term in ["gniazdo", "przedluzacz", "przedłużacz", "listwa przepieciowa", "listwa przepięciowa"]):
        return "socket"
    if any(term in text for term in ["puszka", "rura elektroinstalacyjna"]):
        return "installation"
    if any(term in text for term in ["zlaczka", "złączka", "blok dystrybucyjny", "hlak", "otl"]):
        return "connector"
    if any(term in text for term in ["zaciskarka", "osadzak", "wkrętak", "wkretak", "szczyp", "kurtka", "spodnie", "zestaw"]):
        return "tools"
    if any(term in text for term in ["uziom", "odgrom", "złącze krzyżowe", "zlacze krzyzowe"]):
        return "earthing"
    return "general"


def build_intro(title: str, producer: str, family: str, row: pd.Series) -> str:
    brand = f" marki {producer}" if producer else ""
    if family == "lighting":
        params = join_values([with_unit(row, "attr_moc"), with_unit(row, "attr_strumien"), with_unit(row, "attr_barwa"), value(row, "attr_ip")])
        detail = f" Wyróżniają go parametry {params}, które pomagają ocenić dopasowanie do miejsca pracy." if params else ""
        return f"{title} to produkt oświetleniowy{brand} do zastosowań domowych, technicznych lub zewnętrznych, zależnie od miejsca montażu.{detail}"
    if family == "distribution_box":
        return f"{title} to element do uporządkowania i zabezpieczenia aparatury modułowej w instalacji elektrycznej. Konstrukcja{brand} ułatwia montaż osprzętu oraz późniejszą obsługę rozdzielnicy."
    if family == "protection":
        return f"{title} służy do ochrony lub rozłączania obwodów w instalacji elektrycznej. Przy takim produkcie kluczowe są parametry znamionowe, typ aparatu i zgodność z układem instalacji."
    if family == "socket":
        return f"{title} to osprzęt zasilający{brand} do codziennego korzystania z instalacji elektrycznej. Produkt warto dobrać do obciążenia, miejsca pracy oraz wymaganego stopnia ochrony."
    if family == "installation":
        return f"{title} to element instalacyjny{brand} przeznaczony do prowadzenia lub osadzania przewodów i osprzętu. Dobre dopasowanie wymiarów ułatwia szybki, czysty montaż."
    if family == "connector":
        return f"{title} to element łączeniowy{brand} do rozdziału lub połączenia przewodów w instalacji. Warto zwrócić uwagę na przekrój przewodów, liczbę torów i napięcie pracy."
    if family == "tools":
        return f"{title} to narzędzie lub zestaw narzędzi dla instalatorów i prac technicznych. Dobór warto oprzeć na zakresie pracy, ergonomii oraz parametrach podanych przez producenta."
    if family == "earthing":
        return f"{title} to element instalacji uziemiającej lub odgromowej. Produkt pomaga wykonać trwałe połączenie mechaniczne i elektryczne w systemie ochronnym."
    return f"{title} to produkt techniczny{brand}, którego dobór warto oprzeć na danych z karty produktu i realnych warunkach montażu."


def build_benefits(family: str, row: pd.Series) -> list[str]:
    common = [
        code_benefit(row),
        producer_benefit(row),
    ]
    if family == "lighting":
        return compact_list([
            lighting_power_benefit(row),
            ip_benefit(row, "do warunków montażu"),
            sensor_benefit(row),
            color_or_dimension_benefit(row),
            *common,
        ])[:5]
    if family == "distribution_box":
        return compact_list([
            modules_benefit(row),
            ip_benefit(row, "do miejsca montażu rozdzielnicy"),
            material_benefit(row),
            dimension_benefit(row),
            *common,
        ])[:5]
    if family == "protection":
        return compact_list([
            current_benefit(row),
            protection_benefit(row),
            ip_benefit(row, "do warunków pracy aparatu"),
            "Opisane parametry ułatwiają sprawdzenie zgodności z projektem instalacji.",
            *common,
        ])[:5]
    if family == "socket":
        return compact_list([
            socket_layout_benefit(row),
            current_benefit(row),
            voltage_benefit(row),
            ip_benefit(row, "do miejsca użytkowania osprzętu"),
            cable_benefit(row),
            *common,
        ])[:5]
    if family == "installation":
        return compact_list([
            dimension_benefit(row),
            material_benefit(row),
            ip_benefit(row, "do sposobu zabudowy"),
            "Produkt pomaga utrzymać porządek i powtarzalność wykonania instalacji.",
            *common,
        ])[:5]
    if family == "connector":
        return compact_list([
            cross_section_benefit(row),
            voltage_benefit(row),
            current_benefit(row),
            color_or_dimension_benefit(row),
            *common,
        ])[:5]
    if family == "tools":
        return compact_list([
            range_benefit(row),
            material_benefit(row),
            "Parametry narzędzia pomagają dobrać je do rzeczywistego zakresu prac.",
            "Zestaw lub narzędzie ogranicza potrzebę kompletowania wielu elementów osobno.",
            *common,
        ])[:5]
    if family == "earthing":
        return compact_list([
            dimension_benefit(row),
            material_benefit(row),
            "Produkt jest przeznaczony do prac, w których ważna jest trwałość połączenia.",
            *common,
        ])[:5]
    return compact_list([
        dimension_benefit(row),
        material_benefit(row),
        ip_benefit(row, "do miejsca użytkowania"),
        *common,
    ])[:5]


def build_usage_note(family: str, row: pd.Series) -> str:
    if family == "lighting":
        return "Przed zakupem warto porównać moc, strumień świetlny, barwę światła i stopień ochrony z miejscem montażu. Inne parametry będą ważne przy oświetleniu zewnętrznym, a inne przy panelu, plafonie lub oprawie technicznej."
    if family == "distribution_box":
        return "Przy rozdzielnicach najważniejsze jest dopasowanie liczby modułów, wymiarów i sposobu montażu do planowanej aparatury. Pozwala to zostawić miejsce na czytelne prowadzenie przewodów i późniejszą obsługę."
    if family == "protection":
        return "Aparaturę zabezpieczającą należy dobierać zgodnie z projektem i parametrami instalacji. Szczególnie ważne są prąd znamionowy, charakterystyka, liczba biegunów oraz wymagany poziom ochrony."
    if family == "socket":
        return "W osprzęcie zasilającym warto sprawdzić obciążalność, napięcie, typ gniazd i warunki użytkowania. Ma to znaczenie zarówno dla wygody, jak i dla bezpiecznego podłączenia urządzeń."
    if family == "installation":
        return "Elementy instalacyjne najlepiej dobierać do typu podłoża, średnicy przewodów i sposobu prowadzenia instalacji. Zgodność wymiarów ogranicza ryzyko poprawek podczas montażu."
    if family == "connector":
        return "Przed montażem złączki lub bloku dystrybucyjnego trzeba sprawdzić przekrój i materiał przewodów oraz dopuszczalne napięcie i prąd pracy. To decyduje o poprawnym dopasowaniu elementu."
    if family == "tools":
        return "Przy narzędziach dla elektryków liczy się zakres pracy, wygoda chwytu i zgodność z zastosowaniem. Warto sprawdzić oznaczenia producenta, szczególnie przy narzędziach VDE i sprzęcie akumulatorowym."
    if family == "earthing":
        return "W instalacjach uziemiających i odgromowych liczy się trwałość mechaniczna oraz poprawne połączenie elementów. Przed montażem należy dobrać produkt do przekroju, materiału i sposobu prowadzenia instalacji."
    return "Przed wyborem warto zestawić nazwę, kod producenta i parametry techniczne z wymaganiami instalacji. Pozwala to uniknąć pomyłki między podobnymi wariantami."


def technical_rows(row: pd.Series) -> list[tuple[str, str]]:
    rows: list[tuple[str, str]] = []
    for label, column in TECHNICAL_ROWS:
        item = value(row, column)
        if item:
            rows.append((label, item))
    return rows


def lighting_power_benefit(row: pd.Series) -> str:
    power = with_unit(row, "attr_moc")
    flux = with_unit(row, "attr_strumien")
    color = with_unit(row, "attr_barwa")
    parts = []
    if power:
        parts.append(f"moc {power}")
    if flux:
        parts.append(f"strumień świetlny {flux}")
    if color:
        parts.append(f"temperaturę barwową {color}")
    return f"Parametry świetlne obejmują {join_values(parts)}; dzięki temu łatwiej ocenić jasność i charakter światła przed zakupem." if parts else ""


def ip_benefit(row: pd.Series, context: str) -> str:
    ip = value(row, "attr_ip")
    return f"Stopień ochrony {ip} pomaga dobrać produkt {context}." if ip else ""


def sensor_benefit(row: pd.Series) -> str:
    sensor = value(row, "attr_czujnik")
    if not sensor:
        return ""
    cleaned = re.sub(r"^z\s+", "", sensor, flags=re.IGNORECASE)
    cleaned = re.sub(r"^czujnikiem\b", "czujnik", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace("mikrofalowym", "mikrofalowy")
    return f"{cleaned.capitalize()} ogranicza potrzebę ręcznego włączania światła i poprawia wygodę użytkowania."


def socket_layout_benefit(row: pd.Series) -> str:
    title = value(row, "new_title")
    sockets = first_match(title, r"\b\d+\s*gniazd")
    usb = "USB-A/USB" if re.search(r"USB", title, flags=re.IGNORECASE) else ""
    cable = first_match(title, r"\b\d+\+?\d*\s*m\b")
    parts = []
    if sockets:
        parts.append(sockets)
    if usb:
        parts.append(usb)
    if cable:
        parts.append(f"przewód {cable}")
    return f"Układ {join_values(parts)} ułatwia zasilanie kilku urządzeń z jednego punktu." if parts else ""


def color_or_dimension_benefit(row: pd.Series) -> str:
    color = value(row, "attr_kolor")
    dimensions = value(row, "attr_wymiary") or value(row, "attr_srednica")
    if color and dimensions:
        return f"Kolor {color} i wymiar {dimensions} ułatwiają dopasowanie produktu do miejsca montażu."
    if color:
        return f"Kolor {color} pozwala dopasować produkt do pozostałych elementów instalacji."
    if dimensions:
        return f"Wymiar {dimensions} pomaga ocenić miejsce potrzebne do montażu."
    return ""


def dimension_benefit(row: pd.Series) -> str:
    dimensions = value(row, "attr_wymiary") or join_values([value(row, "attr_dlugosc"), value(row, "attr_szerokosc"), value(row, "attr_wysokosc")])
    return f"Wymiary {dimensions} pomagają sprawdzić dopasowanie do miejsca montażu." if dimensions else ""


def material_benefit(row: pd.Series) -> str:
    material = value(row, "attr_material")
    return f"Materiał: {material}, wpływa na trwałość i sposób użytkowania produktu." if material else ""


def modules_benefit(row: pd.Series) -> str:
    modules = value(row, "attr_liczba_modulow")
    rows_count = re.search(r"(\d+)x12|(\d+) rzę", value(row, "new_title"), flags=re.IGNORECASE)
    row_text = f" oraz układ {rows_count.group(0)}" if rows_count else ""
    return f"Liczba modułów {modules}{row_text} ułatwia zaplanowanie aparatury w obudowie." if modules else ""


def current_benefit(row: pd.Series) -> str:
    current = value(row, "attr_prad") or first_match(value(row, "new_title"), r"\b\d+\s*A\b")
    if current and re.fullmatch(r"\d+(?:[,.]\d+)?", current):
        current = f"{current}A"
    return f"Prąd znamionowy {current} ułatwia sprawdzenie zgodności z obwodem." if current else ""


def voltage_benefit(row: pd.Series) -> str:
    voltage = value(row, "attr_napiecie") or first_match(value(row, "new_title"), r"\b\d{3,4}\s*V(?:AC|DC)?\b")
    return f"Napięcie pracy {voltage} należy porównać z parametrami instalacji." if voltage else ""


def protection_benefit(row: pd.Series) -> str:
    sensitivity = value(row, "attr_czulosc") or first_match(value(row, "new_title"), r"\b\d+\s*mA\b")
    short = value(row, "attr_zwarciowa_zdolnosc") or first_match(value(row, "new_title"), r"\b\d+\s*kA\b")
    details = join_values([f"czułość {sensitivity}" if sensitivity else "", f"zdolność zwarciowa {short}" if short else ""])
    return f"Parametry ochronne obejmują {details}." if details else ""


def cable_benefit(row: pd.Series) -> str:
    cable = value(row, "attr_dlugosc_przewodu") or first_match(value(row, "new_title"), r"\b\d+\s*m\b")
    return f"Długość przewodu {cable} pozwala dopasować produkt do stanowiska pracy." if cable else ""


def cross_section_benefit(row: pd.Series) -> str:
    section = value(row, "attr_przekroj") or first_match(value(row, "new_title"), r"\b[\d,]+[-–]\d+\s*mm2\b|\b\d+mm2\+\d+mm2\b")
    return f"Zakres przekrojów {section} pomaga dobrać element do przewodów w instalacji." if section else ""


def range_benefit(row: pd.Series) -> str:
    title = value(row, "new_title")
    found = first_match(title, r"\b\d+[-–]\d+\s*mm2\b|\b\d+\s*kN\b|\b\d+\s*J\b|\b\d+[-–]\d+\s*mm\b")
    return f"Zakres pracy {found} pozwala ocenić, do jakich zadań produkt będzie odpowiedni." if found else ""


def code_benefit(row: pd.Series) -> str:
    code = value(row, "Kod Producenta")
    return f"Kod producenta {code} ułatwia jednoznaczną identyfikację wariantu." if code else ""


def producer_benefit(row: pd.Series) -> str:
    producer = value(row, "attr_producent") or value(row, "Producent")
    return f"Producent: {producer}." if producer else ""


def with_unit(row: pd.Series, column: str) -> str:
    item = value(row, column)
    if not item:
        return ""
    unit_by_column = {
        "attr_moc": "W",
        "attr_strumien": "lm",
        "attr_barwa": "K",
        "attr_lm_w": "lm/W",
    }
    unit = unit_by_column.get(column, "")
    return item if not unit or re.search(rf"{re.escape(unit)}\b", item, flags=re.IGNORECASE) else f"{item}{unit}"


def first_match(text: str, pattern: str) -> str:
    match = re.search(pattern, text, flags=re.IGNORECASE)
    return compact_spaces(match.group(0)) if match else ""


def value(row: pd.Series, column: str) -> str:
    if not column or column not in row:
        return ""
    item = row.get(column, "")
    return "" if is_blank(item) else compact_spaces(str(item))


def join_values(values: list[str]) -> str:
    cleaned = compact_list(values)
    if len(cleaned) <= 1:
        return "".join(cleaned)
    if len(cleaned) == 2:
        return f"{cleaned[0]} i {cleaned[1]}"
    return f"{', '.join(cleaned[:-1])} i {cleaned[-1]}"


def compact_list(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for item in values:
        item = compact_spaces(item)
        key = item.lower()
        if item and key not in seen:
            result.append(item)
            seen.add(key)
    return result


def normalized(text: str) -> str:
    return strip_accents(text).lower()


def strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", " ", text)


def chars_without_spaces(text: str) -> int:
    return len(re.sub(r"\s+", "", text))


if __name__ == "__main__":
    main()
