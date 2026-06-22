from __future__ import annotations

import argparse
import hashlib
import json
import re
from html import escape
from pathlib import Path
from typing import Any

import pandas as pd

import export_to_baselinker_csv as baselinker
from utils import compact_spaces, is_blank, normalize_header, read_products, write_products


DEFAULT_DESCRIPTION_COLUMN = "description_html"
DEFAULT_MIN_CHARS_NO_SPACES = 1500
DEFAULT_CODEX_BRIEF_OUTPUT = "reports/descriptions/description_codex_brief.xlsx"
DEFAULT_SEO_KNOWLEDGE_PATH = "dictionaries/seo_description_knowledge.yaml"
CODEX_BRIEF_SHEETS = [
    "Produkt",
    "Product facts",
    "Compatibility facts",
    "Rejected facts",
    "Prompt dla Codex",
    "Opis wygenerowany",
    "Walidacja",
]
CODEX_MIN_CHARS_NO_SPACES = 1500
CODEX_SPARSE_FACTS_MIN_CHARS_NO_SPACES = 900
CODEX_MEDIUM_FACTS_MIN_CHARS_NO_SPACES = 1200
POLISH_DIACRITICS = "ąćęłńóśźżĄĆĘŁŃÓŚŹŻ"
MOJIBAKE_MARKERS = ["�", "Ä", "Å", "Ĺ", "Â"]
ASCII_POLISH_MARKERS = [
    r"\bocenic\b",
    r"\bkatem\b",
    r"\bpor[oó]wnac\b",
    r"\bporown",
    r"\boswiet",
    r"\bswiatl",
    r"\bDzieki\b",
    r"\bktore\b",
    r"\bwazne\b",
    r"\blatwiej\b",
    r"\bspojny\b",
    r"\bpomylki\b",
    r"\buzytk",
    r"\brozwiaz",
]
META_DESCRIPTION_PATTERNS = [
    r"\btaki spos[oó]b opisu\b",
    r"\btaki zestaw danych\b",
    r"\bod razu porz[aą]dkuje wyb[oó]r\b",
    r"\bm[oó]wimy o elemencie\b",
    r"\bdla klienta oznacza to\b",
    r"\bprostszy wyb[oó]r\b",
    r"\bmniejsze ryzyko pomy[lł]ki\b",
    r"\bnie udaje\b",
    r"\bnie zmienia parametr[oó]w oprawy\b",
    r"\bsamodzielnego [źz]r[oó]d[lł]a [śs]wiat[lł]a\b",
    r"\bsam typ produktu\b",
    r"\bjasno wskazuje\b",
    r"\bkomponent wspieraj[aą]cy funkcj[eę]\b",
    r"\botrzymujesz akcesorium\b",
    r"\bprzygotowane do pracy\b",
    r"\bw obr[eę]bie sp[oó]jnej rodziny produktowej\b",
    r"\bsp[oó]jna rodzina produktowa\b",
    r"\bca[lł]ego zestawu\b",
    r"\bopisany przez kluczowe parametry\b",
    r"\bpozwala szybko oceni[cć]\b",
    r"\bdoboru do projektu\b",
    r"\bpozycjami w tej samej grupie produktowej\b",
    r"\bprzy wyborze warto patrze[cć]\b",
    r"\bnaj[lł]atwiej rozr[oó][zż]niaj[aą] modele\b",
    r"\bopis pozostaje\b",
    r"\bdanych pochodz[aą]cych z briefu\b",
    r"\bbriefu\b",
    r"\bsklep chce\b",
    r"\bw praktyce pomaga to wtedy\b",
    r"\btak przygotowany opis\b",
    r"\bpokaza[cć] model\b",
    r"\bbez nadmiarowych deklaracji\b",
    r"\bbez nadmiaru og[oó]lnik[oó]w\b",
    r"\bklient szuka rozwi[aą]zania\b",
    r"\bjasnego opisu cech\b",
    r"\bopis zgodno[śs]ci\b",
    r"\bporz[aą]dkuje dob[oó]r\b",
    r"\bporz[aą]dkuje informacje\b",
    r"\btechniczny profil wariantu\b",
    r"\bzachowa[cć] porz[aą]dek przy zam[oó]wieniu\b",
]
ROBOTIC_LANGUAGE_PATTERNS = [
    r"\bzostal[ay] przygotowan[ay] do\b",
    r"\bwspiera (?:zastosowanie|prace|doswietlenie|neutralny wyglad|czytelne|jasne)\b",
    r"\bporzadkuj\w* (?:parametry|wybor|identyfikacje)\b",
    r"\buzupelnia\w* (?:zestaw|profil)\b",
    r"\bzestaw (?:informacji|parametrow) (?:istotnych|waznych)\b",
    r"\bprofil produktu\b",
    r"\bdobrze wpisuje sie\b",
    r"\bulatwia (?:przypisanie|uwzglednienie|ocene|planowanie)\b",
    r"\bpomaga\w* utrzymac spojnosc\b",
    r"\bspojnosc (?:calego systemu|parametrow|wizualna)\b",
    r"\bten model bedzie praktyczny\b",
    r"\blatwa identyfikacja wlasciwego wariantu\b",
    r"\bbez siegania po niepotwierdzone zalozenia\b",
    r"\bprzewidywalnych parametrach\b",
    r"\bspokojny wyglad sufitu\b",
    r"\bw ktorych liczy sie\b",
    r"\btam gdzie wazna jest\b",
]

FORBIDDEN_PHRASES = [
    "Dzięki temu klient",
    "zamiast porównywać",
    "Najważniejsze parametry tego wariantu",
    "W praktyce oznacza to",
    "ułatwia porównanie tego wariantu",
    "opis pomaga",
    "Taki sposób opisu",
    "Dla klienta oznacza to",
    "nie udaje samodzielnego źródła światła",
    "Sam typ produktu jasno wskazuje",
    "komponent wspierający funkcję",
    "otrzymujesz akcesorium",
    "spójnej rodziny produktowej",
    "opisany przez kluczowe parametry",
    "Taki zestaw danych",
    "danych pochodzących z briefu",
    "sklep chce",
    "Tak przygotowany opis",
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
    parser.add_argument("--input", default="", help="CSV/XLSX z produktami po AlkanInatorze.")
    parser.add_argument("--sheet", help="Arkusz XLSX.")
    parser.add_argument("--output", default="", help="Plik wynikowy CSV/XLSX dla trybu generate.")
    parser.add_argument("--mode", choices=["generate", "codex_brief", "write_codex_review", "validate_codex_review"], default="generate")
    parser.add_argument("--sku", default="", help="SKU/Kod produktu dla trybu codex_brief, np. 33482/KAN.")
    parser.add_argument("--brief-input", default="", help="XLSX brief do złożenia review w trybie write_codex_review.")
    parser.add_argument("--brief-output", default=DEFAULT_CODEX_BRIEF_OUTPUT, help="XLSX z briefem dla Codex.")
    parser.add_argument("--review-input", default="", help="XLSX review do walidacji w trybie validate_codex_review.")
    parser.add_argument("--review-output", default="", help="XLSX review do zapisania w trybie write_codex_review.")
    parser.add_argument("--description-html", default="", help="Gotowy HTML opisu dla trybu write_codex_review.")
    parser.add_argument("--description-file", default="", help="Plik UTF-8 z gotowym HTML opisu dla trybu write_codex_review.")
    parser.add_argument("--catalog-knowledge", default=baselinker.DEFAULT_CATALOG_KNOWLEDGE_PATH)
    parser.add_argument("--seo-knowledge", default=DEFAULT_SEO_KNOWLEDGE_PATH)
    parser.add_argument("--description-column", default=DEFAULT_DESCRIPTION_COLUMN)
    parser.add_argument("--title-column", default="", help="Wymusza kolumne z nazwa produktu.")
    parser.add_argument("--min-words", type=int, default=0, help="Zgodnosc wsteczna; preferuj --min-chars-no-spaces.")
    parser.add_argument("--min-chars-no-spaces", type=int, default=DEFAULT_MIN_CHARS_NO_SPACES)
    args = parser.parse_args()

    if args.mode == "validate_codex_review":
        if not args.review_input:
            raise SystemExit("Tryb validate_codex_review wymaga --review-input.")
        checks = validate_codex_review_file(args.review_input)
        for check in checks:
            print(f"{check['status']}: {check['check']} - {check['details']}")
        if any(check["status"] in {"ERROR", "WARNING"} for check in checks):
            raise SystemExit("Review wymaga poprawy przed oddaniem.")
        print("OK: review przeszedł walidację.")
        return

    if args.mode == "write_codex_review":
        if not args.brief_input:
            raise SystemExit("Tryb write_codex_review wymaga --brief-input.")
        if not args.review_output:
            raise SystemExit("Tryb write_codex_review wymaga --review-output.")
        description_html = read_description_input(args.description_html, args.description_file)
        checks = write_codex_review_from_description(args.brief_input, args.review_output, description_html)
        for check in checks:
            print(f"{check['status']}: {check['check']} - {check['details']}")
        if any(check["status"] in {"ERROR", "WARNING"} for check in checks):
            raise SystemExit("Review zapisany, ale wymaga poprawy przed oddaniem.")
        print(f"OK: zapisano poprawny review Codex: {args.review_output}")
        return

    if not args.input:
        raise SystemExit(f"Tryb {args.mode} wymaga --input.")
    df = read_products(args.input, sheet_name=args.sheet)
    if args.mode == "codex_brief":
        if not args.sku:
            raise SystemExit("Tryb codex_brief wymaga --sku.")
        catalog_knowledge = baselinker.load_yaml(args.catalog_knowledge) if args.catalog_knowledge else {}
        seo_knowledge = baselinker.load_yaml(args.seo_knowledge) if args.seo_knowledge else {}
        brief = build_codex_brief_for_sku(
            df,
            args.sku,
            args.title_column,
            catalog_knowledge,
            seo_knowledge,
        )
        write_codex_brief_report(brief, args.brief_output)
        print(f"OK: zapisano brief Codex: {args.brief_output}")
        return

    if not args.output:
        raise SystemExit("Tryb generate wymaga --output.")
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


def build_codex_brief_for_sku(
    df: pd.DataFrame,
    sku: str,
    title_column: str = "",
    catalog_knowledge: dict[str, Any] | None = None,
    seo_knowledge: dict[str, Any] | None = None,
) -> dict[str, Any]:
    row = find_product_row_by_sku(df, sku)
    normalizer = baselinker.build_feature_value_normalizer(catalog_knowledge or {})
    raw_features = baselinker.build_features(row, False, feature_value_normalizer=normalizer)
    product_name = product_title(row, title_column)
    context = {
        "sku": sku_for_row(row),
        "ean": first_present(row.get("EAN", ""), row.get("ean", "")),
        "name": product_name,
    }
    review_rows: list[dict[str, str]] = []
    product_features = baselinker.filter_features_for_catalog(raw_features, normalizer, review_rows, context)
    apply_power_range_for_description(row, product_features)
    product_facts = build_product_facts(row, product_features)
    compatibility_facts = build_compatibility_facts(row, raw_features)
    rejected_facts = build_rejected_facts(raw_features, product_features, row)
    seo_knowledge = seo_knowledge or load_default_seo_knowledge()
    seo_keyword = find_seo_keyword(row, seo_knowledge)
    prompt = build_codex_description_prompt(
        product_name,
        product_facts,
        compatibility_facts,
        rejected_facts,
        seo_keyword,
        seo_knowledge,
    )
    validation = validate_codex_description("", product_facts, rejected_facts, prompt_only=True)
    return {
        "product": {
            "sku": sku_for_row(row),
            "requested_sku": sku,
            "ean": first_present(row.get("EAN", ""), row.get("ean", "")),
            "name": product_name,
            "seo_keyword": seo_keyword,
            "category": first_present(row.get("proponowana_kategoria_1", ""), row.get("category", ""), row.get("Kategoria", "")),
            "source_row_index": int(row.name) if isinstance(row.name, int) else str(row.name),
        },
        "product_facts": product_facts,
        "compatibility_facts": compatibility_facts,
        "rejected_facts": rejected_facts,
        "prompt": prompt,
        "generated_description": "",
        "validation": validation,
    }


def find_product_row_by_sku(df: pd.DataFrame, sku: str) -> pd.Series:
    requested = normalize_sku_for_match(sku)
    for _, row in df.iterrows():
        candidates = [
            row.get("sku", ""),
            row.get("SKU", ""),
            row.get("Kod", ""),
            row.get("Kod producenta", ""),
            row.get("Kod Producenta", ""),
        ]
        if any(normalize_sku_for_match(value) == requested for value in candidates if not is_blank(value)):
            return row
    raise SystemExit(f"Nie znaleziono produktu dla SKU/Kod: {sku}")


def normalize_sku_for_match(value: Any) -> str:
    text = compact_spaces(str(value or ""))
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text.split("/", 1)[0].strip().upper()


def sku_for_row(row: pd.Series) -> str:
    return first_present(row.get("sku", ""), row.get("SKU", ""), row.get("Kod", ""), row.get("Kod producenta", ""))


def build_product_facts(row: pd.Series, product_features: dict[str, str]) -> list[dict[str, str]]:
    facts: list[dict[str, str]] = []
    for label, value in {
        "Kod producenta": first_present(row.get("Kod", ""), row.get("Kod producenta", ""), row.get("sku", "")),
        "EAN": first_present(row.get("EAN", ""), row.get("ean", "")),
    }.items():
        if value:
            facts.append({"fact": label, "value": value, "source": "product_identity"})
    for label, value in product_features.items():
        if label == "Dane producenta":
            continue
        facts.append({"fact": label, "value": value, "source": "filtered_baselinker_features"})
    return dedupe_fact_rows(facts)


def build_compatibility_facts(row: pd.Series, raw_features: dict[str, str]) -> list[dict[str, str]]:
    facts: list[dict[str, str]] = []
    compatibility = first_present(
        row.get("attr_pasuje_do", ""),
        row.get("Pasuje do", ""),
        raw_features.get("Zastosowanie", ""),
    )
    if compatibility:
        facts.append({
            "fact": "Pasuje do",
            "value": compatibility,
            "source": "compatibility_source",
        })
    return dedupe_fact_rows(facts)


def build_rejected_facts(
    raw_features: dict[str, str],
    product_features: dict[str, str],
    row: pd.Series,
) -> list[dict[str, str]]:
    rejected: list[dict[str, str]] = []
    system_facts = {"EAN (GTIN)", "Kod producenta", "Dane producenta"}
    product_type = product_features.get("Typ produktu", raw_features.get("Typ produktu", ""))
    for label, value in raw_features.items():
        if label in system_facts or not value:
            continue
        if label in product_features and product_features[label] == value:
            continue
        reason = "not_in_filtered_baselinker_features"
        if label == "Moc [W]" and product_type in baselinker.ACCESSORY_PRODUCT_TYPES_WITHOUT_OWN_POWER:
            reason = "accessory_compatible_fixture_power_not_product_power"
        rejected.append({
            "fact": label,
            "value": value,
            "source": "raw_feature_rejected_by_filter",
            "reason": reason,
        })
    return dedupe_fact_rows(rejected)


def dedupe_fact_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for row in rows:
        key = (row.get("fact", ""), row.get("value", ""), row.get("source", ""))
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def load_default_seo_knowledge() -> dict[str, Any]:
    path = Path(DEFAULT_SEO_KNOWLEDGE_PATH)
    if not path.exists():
        return {}
    return baselinker.load_yaml(path)


def find_seo_keyword(row: pd.Series, seo_knowledge: dict[str, Any]) -> str:
    columns = (
        seo_knowledge.get("product_description", {})
        .get("input", {})
        .get("seo_keyword_columns", [])
    )
    return first_present(*(row.get(column, "") for column in columns))


def build_codex_description_prompt(
    product_name: str,
    product_facts: list[dict[str, str]],
    compatibility_facts: list[dict[str, str]],
    rejected_facts: list[dict[str, str]],
    seo_keyword: str = "",
    seo_knowledge: dict[str, Any] | None = None,
) -> str:
    seo_rules = (seo_knowledge or load_default_seo_knowledge()).get("product_description", {})
    minimum_chars = int(seo_rules.get("length", {}).get("minimum_chars_with_spaces", 1200))
    maximum_keyword_occurrences = int(seo_rules.get("seo_keyword", {}).get("maximum_occurrences", 3))
    keyword_instruction = (
        f'- Fraza kluczowa dla tego produktu to: "{seo_keyword}".'
        if seo_keyword
        else "- Fraza kluczowa nie została dostarczona; nie wymyślaj jej i stosuj naturalną frazę produktową."
    )
    return "\n".join([
        "Napisz ręcznie unikalny opis HTML SEO dla jednego produktu do sklepu Alkan.",
        "",
        "Sposób pracy:",
        "- Przed pisaniem przeczytaj wzorzec: prompts/examples/good_description_arot.html. Naśladuj poziom konkretności i strukturę, ale nie jego fakty.",
        "- Ten opis tworzysz jako osobne zadanie. Najpierw przeczytaj fakty tego produktu, wybierz jego główny motyw i dopiero potem napisz tekst.",
        "- Nie twórz skryptu, pętli, słownika akapitów ani szablonu generującego treść dla wielu produktów.",
        "- Python może wyłącznie odczytać XLSX, zapisać napisany przez Ciebie HTML i uruchomić walidator.",
        "- Jeśli przetwarzasz paczkę, zakończ opis, walidację i poprawki jednego SKU przed otwarciem następnego briefu.",
        "",
        "SEO i konstrukcja opisu:",
        keyword_instruction,
        f"- Opis musi mieć co najmniej {minimum_chars} znaków ze spacjami, o ile pozwala na to liczba potwierdzonych faktów.",
        f"- Fraza kluczowa może wystąpić maksymalnie {maximum_keyword_occurrences} razy w całym opisie.",
        "- Gdy fraza jest dostarczona: rozpocznij nią wstęp, umieść ją naturalnie w <h2> i rozpocznij nią pierwszy akapit bezpośrednio po <h2>.",
        "- Nagłówki <h3> nie mogą zawierać frazy kluczowej.",
        "- Preferuj <h2> w formie naturalnego pytania o konkretny produkt, jeśli taka forma pasuje do treści.",
        "- Pierwszy akapit rozpocznij od <strong>pełnej nazwy produktu</strong>, a następnie od razu wyjaśnij jego zastosowanie i najważniejsze potwierdzone cechy.",
        "- Umieść naturalną frazę produktową w <h2>; nagłówek ma mówić o konkretnym produkcie lub jego przewadze, a nie brzmieć 'Najważniejsze cechy'.",
        "- Po części wprowadzającej dodaj sekcję <h3>Najważniejsze zalety</h3> z korzyściami wynikającymi z faktów.",
        "- Dodaj <h3>Specyfikacja techniczna</h3>; nazwy parametrów w liście wyróżnij tagiem <strong>.",
        "- Bezpośrednio po liście specyfikacji dodaj końcowy akapit <p>. Ma podsumować potwierdzone zalety tego wariantu i naturalnie zachęcić do zakupu albo wskazać praktyczny powód wyboru.",
        "- Końcowy akapit jest obowiązkowy i musi być ostatnim widocznym elementem opisu. Nie kończ opisu na </ul> specyfikacji.",
        "- Sekcję o zastosowaniu dodaj tylko wtedy, gdy można ją napisać rzetelnie na podstawie typu produktu i faktów.",
        "- Pisz językiem używanym przez człowieka kupującego produkt: najpierw zastosowanie, potem cechy i wynikające z nich korzyści, na końcu parametry.",
        "- Każdą ważną cechę rozwiń według modelu: cecha -> zaleta -> praktyczna korzyść dla użytkownika.",
        "- Akapity mają być krótkie i zawierać najwyżej 3-4 zdania; unikaj ścian tekstu.",
        "- Lista najważniejszych cech lub korzyści powinna mieć 3-6 punktów.",
        "- Nie używaj pustych superlatywów, takich jak: rewelacyjny, innowacyjny, najlepszy, niesamowity.",
        "- Produkty różniące się tylko kolorem lub wariantem muszą mieć odmienną treść; nie przestawiaj tych samych zdań.",
        "- Zmieniaj kompozycję, argumentację i słownictwo zależnie od produktu. Stałe mogą być wyłącznie zasady HTML i nazwa sekcji specyfikacji.",
        "- Pisz tak, jak do klienta w sklepie: krótkie, konkretne zdania, codzienne słownictwo i jasna odpowiedź, po co dana cecha jest przydatna.",
        "- Preferuj zdania bezpośrednie: 'Obudowa ma klasę IP65 i jest chroniona przed pyłem oraz strugami wody' zamiast 'IP65 wspiera zastosowanie w wymagającym otoczeniu'.",
        "- Najpierw nazwij cechę, potem podaj jej praktyczny skutek. Nie opisuj procesu analizowania, dobierania ani klasyfikowania produktu.",
        "- Kod, EAN, seria, napięcie i gwarancja mogą zostać tylko w specyfikacji, jeśli nie da się z nich zbudować naturalnego i użytecznego zdania.",
        "",
        "Zasady bezwzględne:",
        "- Używaj tylko faktów z sekcji PRODUCT_FACTS jako cech produktu.",
        "- Fakty z COMPATIBILITY_FACTS możesz opisać tylko jako kompatybilność albo dopasowanie do innej oprawy/części.",
        "- Nie używaj REJECTED_FACTS jako cech produktu.",
        "- Jeśli REJECTED_FACTS zawiera moc, nie pisz, że opisywany produkt ma tę moc.",
        "- Nie wymyślaj parametrów, wymiarów, certyfikatów, zastosowań ani obietnic producenta.",
        "- Zachowaj dłuższy styl SEO, ale bez pustych zdań typu 'z danych technicznych'.",
        "- Pisz z polskimi znakami. Opis z '?' zamiast polskich liter albo mojibake typu 'Ä', 'Å', 'Ĺ' jest błędny.",
        "- Celuj w 1500-2200 znaków bez spacji przy bogatych danych. Przy skromnych faktach napisz krócej i konkretniej zamiast sztucznie rozciągać tekst.",
        "- Nie kopiuj poprzedniego wyniku z pliku review. Jeśli wcześniejszy opis był krótki albo miał uszkodzone polskie znaki, napisz nowy opis od zera.",
        "- Review XLSX zapisuj w głównym repo: C:\\Users\\Handlowiec\\Desktop\\AlkanInator\\reports\\descriptions. Nie zapisuj finalnego pliku w .codex\\worktrees.",
        "- Nie pisz o samym opisie ani o procesie wyboru. Zakazane są zdania typu 'Taki sposób opisu...', 'dla klienta oznacza to...', 'mówimy o elemencie...'.",
        "- Nie tłumacz, czym produkt nie jest. Zamiast 'nie udaje źródła światła' po prostu opisz go jako akcesorium/siatkę ochronną i podaj kompatybilność.",
        "- Nie używaj abstrakcyjnych wypełniaczy typu 'sam typ produktu jasno wskazuje', 'komponent wspierający funkcję', 'otrzymujesz akcesorium', 'spójna rodzina produktowa'.",
        "- Jeśli faktów jest mało, możesz dodać krótki, neutralny akapit o typie produktu: czym zwykle jest taka część, gdzie się ją stosuje i na co zwrócić uwagę przy doborze.",
        "- Ogólny akapit o typie produktu nie może dopisywać niepotwierdzonych parametrów konkretnego modelu.",
        "- Dobre rozwinięcie: 'Siatka ochronna jest akcesorium stosowanym przy oprawach, gdy potrzebna jest dodatkowa fizyczna osłona przed przypadkowym kontaktem lub uderzeniem. Przy doborze warto sprawdzić zgodność z konkretną serią i modelem oprawy.'",
        "- Lepiej napisać krótszy, konkretny opis niż dociągać długość pustymi zdaniami.",
        "- Nie pisz o briefie, opisie, sklepie, kliencie ani procesie porównywania produktów. Zakazane: 'opisany przez kluczowe parametry', 'taki zestaw danych', 'danych pochodzących z briefu', 'sklep chce', 'klient szuka rozwiązania', 'tak przygotowany opis'.",
        "- Nie używaj polskich słów bez znaków diakrytycznych. Zakazane są formy typu 'ocenic', 'katem', 'porownac', 'Dzieki', 'swiatla', 'oswietlenia', 'ktore'.",
        "- Pełna nazwa w otwarciu służy SEO, ale zdanie po niej ma być naturalne. Nie używaj automatycznej formuły 'marki ... z serii ... to ...' w każdym opisie.",
        "- Nie nadużywaj słów 'porządkuje', 'wariant', 'profil', 'dobór', 'pozwala ocenić' ani zdań o porównywaniu modeli.",
        "- Nie używaj urzędowo-technicznych konstrukcji: 'wspiera zastosowanie', 'uzupełnia zestaw informacji', 'profil produktu', 'ułatwia przypisanie', 'porządkuje parametry', 'dobrze wpisuje się'.",
        "- Nie pisz, że seria produktu 'pomaga zachować spójność', jeśli brief nie potwierdza istnienia pasujących wizualnie produktów.",
        "- Nie zamieniaj parametrów w pozorne korzyści. Napięcie 220-240 V nie 'ułatwia dopasowania', a kod producenta nie jest zaletą użytkową.",
        "- Unikaj zdań zbudowanych według schematu 'cecha wspiera...', 'parametr ułatwia...', 'seria pomaga...'. Napisz wprost, co produkt robi.",
        "- Przykładowy opis AROT jest wzorcem jakości i struktury, nie źródłem faktów. Nie kopiuj z niego IP, temperatur, odporności, materiału ani zastosowań do innego produktu.",
        "",
        "Wymagany format HTML:",
        "- 3-5 akapitów <p>; pierwszy zaczyna się od <strong>pełnej nazwy produktu</strong>",
        "- co najmniej jedna konkretna sekcja <h2>",
        "- sekcja <h3>Najważniejsze zalety</h3> i lista <ul><li> z konkretnymi korzyściami",
        "- sekcja <h3>Specyfikacja techniczna</h3> tylko z PRODUCT_FACTS; etykiety parametrów w <strong>",
        "- po liście specyfikacji obowiązkowy końcowy <p> z podsumowaniem i naturalną zachętą do zakupu",
        "- elementy listy bez myślnika po tagu, czyli <li>tekst</li>, nie <li>- tekst</li>",
        "",
        "Walidacja gotowego review przed oddaniem:",
        "Najpierw złóż review helperem, który istnieje w repo:",
        'py C:\\Users\\Handlowiec\\Desktop\\AlkanInator\\src\\codex_description_agent.py --brief-input "<brief.xlsx>" --description-file "<opis.html>" --review-output "<review.xlsx>"',
        'py src\\generate_product_descriptions.py --mode validate_codex_review --review-input "<review.xlsx>"',
        "Jeśli ta komenda zwróci ERROR albo WARNING, popraw opis i uruchom ją ponownie.",
        "",
        f"PRODUCT_NAME: {product_name}",
        f"SEO_KEYWORD: {seo_keyword or 'brak'}",
        "",
        "PRODUCT_FACTS:",
        facts_as_bullets(product_facts),
        "",
        "COMPATIBILITY_FACTS:",
        facts_as_bullets(compatibility_facts),
        "",
        "REJECTED_FACTS:",
        facts_as_bullets(rejected_facts),
    ])


def facts_as_bullets(facts: list[dict[str, str]]) -> str:
    if not facts:
        return "- brak"
    return "\n".join(f"- {item.get('fact', '')}: {item.get('value', '')}" for item in facts)


def validate_codex_description(
    description_html: str,
    product_facts: list[dict[str, str]],
    rejected_facts: list[dict[str, str]],
    prompt_only: bool = False,
    product_name: str = "",
    seo_keyword: str = "",
    seo_knowledge: dict[str, Any] | None = None,
) -> list[dict[str, str]]:
    if prompt_only and not description_html:
        return [{"status": "WAITING_FOR_CODEX", "check": "description_present", "details": "Brief gotowy, opis nie został jeszcze wygenerowany."}]
    checks: list[dict[str, str]] = []
    text = strip_html(description_html)
    if not compact_spaces(text):
        checks.append({"status": "ERROR", "check": "description_present", "details": "Brak opisu do walidacji."})
        return checks
    checks.append({"status": "OK", "check": "description_present", "details": "Opis jest obecny."})
    seo_rules = (seo_knowledge or load_default_seo_knowledge()).get("product_description", {})
    minimum_chars_with_spaces = int(seo_rules.get("length", {}).get("minimum_chars_with_spaces", 1200))
    if len(text) < minimum_chars_with_spaces:
        checks.append({
            "status": "WARNING",
            "check": "description_length_with_spaces",
            "details": f"Opis ma {len(text)} znaków ze spacjami; wymagane minimum to {minimum_chars_with_spaces}.",
        })
    else:
        checks.append({
            "status": "OK",
            "check": "description_length_with_spaces",
            "details": "Opis spełnia minimalną długość liczoną ze spacjami.",
        })
    chars_no_spaces = len(re.sub(r"\s+", "", text))
    minimum_chars = codex_minimum_chars(product_facts)
    if chars_no_spaces < minimum_chars:
        checks.append({
            "status": "WARNING",
            "check": "description_length",
            "details": f"Opis ma {chars_no_spaces} znaków bez spacji, wymagane minimum dla tego briefu to {minimum_chars}.",
        })
    else:
        checks.append({"status": "OK", "check": "description_length", "details": "Opis spełnia minimalną długość."})
    if not any(char in text for char in POLISH_DIACRITICS):
        checks.append({
            "status": "ERROR",
            "check": "polish_diacritics",
            "details": "Opis nie zawiera polskich znaków; prawdopodobnie został zapisany bez ogonków.",
        })
    if has_encoding_damage(description_html):
        checks.append({
            "status": "ERROR",
            "check": "encoding_integrity",
            "details": "Opis zawiera znaki sugerujące uszkodzone kodowanie albo zamianę polskich liter na '?'.",
        })
    ascii_polish_marker = find_ascii_polish_marker(text)
    if ascii_polish_marker:
        checks.append({
            "status": "ERROR",
            "check": "missing_polish_diacritics",
            "details": f"Opis zawiera polskie słowo zapisane bez znaków diakrytycznych: {ascii_polish_marker}",
        })
    if re.search(r"<li[^>]*>\s*[-–—]\s*", description_html, flags=re.IGNORECASE):
        checks.append({
            "status": "WARNING",
            "check": "html_list_format",
            "details": "Elementy listy zaczynają się od myślnika; użyj <li>tekst</li> bez dodatkowego '-'.",
        })
    paragraph_count = len(re.findall(r"<p\b", description_html, flags=re.IGNORECASE))
    if paragraph_count < 3 or paragraph_count > 5:
        checks.append({
            "status": "WARNING",
            "check": "html_paragraph_count",
            "details": f"Opis ma {paragraph_count} akapitów <p>; zalecane 3-5.",
        })
    if not re.search(r"<h2\b", description_html, flags=re.IGNORECASE):
        checks.append({"status": "WARNING", "check": "html_h2", "details": "Opis nie zawiera sekcji <h2>."})
    validate_seo_opening(checks, description_html, product_name)
    validate_seo_headings(checks, description_html, product_facts)
    if seo_keyword:
        validate_seo_keyword_placement(checks, description_html, seo_keyword, seo_rules)
    if not re.search(r"<h3\b[^>]*>\s*Specyfikacja techniczna\s*</h3>", description_html, flags=re.IGNORECASE):
        checks.append({"status": "WARNING", "check": "html_specification_h3", "details": "Brak nagłówka <h3>Specyfikacja techniczna</h3>."})
    if not re.search(r"<h3\b[^>]*>\s*Najważniejsze zalety\s*</h3>", description_html, flags=re.IGNORECASE):
        checks.append({"status": "WARNING", "check": "html_benefits_h3", "details": "Brak sekcji <h3>Najważniejsze zalety</h3>."})
    if re.search(r"<h3\b[^>]*>\s*Specyfikacja techniczna\s*</h3>", description_html, flags=re.IGNORECASE):
        specification_html = re.split(
            r"<h3\b[^>]*>\s*Specyfikacja techniczna\s*</h3>",
            description_html,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[1]
        if not re.search(r"<li[^>]*>\s*<strong\b", specification_html, flags=re.IGNORECASE):
            checks.append({
                "status": "WARNING",
                "check": "html_specification_labels",
                "details": "Etykiety parametrów w specyfikacji powinny być wyróżnione tagiem <strong>.",
            })
        validate_closing_summary(checks, description_html)
    meta_phrase = find_meta_description_phrase(text)
    if meta_phrase:
        checks.append({
            "status": "ERROR",
            "check": "meta_description_language",
            "details": f"Opis zawiera meta-komentarz zamiast opisu produktu: {meta_phrase}",
        })
    for rejected in rejected_facts:
        fact = rejected.get("fact", "")
        value = rejected.get("value", "")
        if rejected_fact_used_as_product_claim(text, fact, value):
            checks.append({
                "status": "ERROR",
                "check": "rejected_fact_not_used",
                "details": f"Opis używa odrzuconego faktu jako cechy produktu: {fact}={value}",
            })
    for phrase in ["z danych technicznych", "opis warto odczytywać", "dane techniczne pomagają"]:
        if phrase in text.lower():
            checks.append({"status": "WARNING", "check": "generic_phrase", "details": f"Opis zawiera generyczną frazę: {phrase}"})
    repeated_word = find_overused_word(text)
    if repeated_word:
        checks.append({
            "status": "WARNING",
            "check": "repetitive_language",
            "details": f"Opis nadużywa słowa lub konstrukcji '{repeated_word}' i brzmi szablonowo.",
        })
    robotic_phrase = find_robotic_language_phrase(text)
    if robotic_phrase:
        checks.append({
            "status": "WARNING",
            "check": "robotic_language",
            "details": f"Opis zawiera urzędowo-robotyczną konstrukcję: {robotic_phrase}",
        })
    if not any(check["status"] in {"ERROR", "WARNING"} for check in checks):
        checks.append({"status": "OK", "check": "rejected_facts", "details": "Nie wykryto użycia odrzuconych faktów."})
    return checks


def codex_minimum_chars(product_facts: list[dict[str, str]]) -> int:
    meaningful_facts = [
        row for row in product_facts
        if compact_spaces(str(row.get("value", "")))
        and row.get("fact") not in {"EAN", "Kod producenta", "Producent"}
    ]
    if len(meaningful_facts) <= 4:
        return CODEX_SPARSE_FACTS_MIN_CHARS_NO_SPACES
    if len(meaningful_facts) <= 8:
        return CODEX_MEDIUM_FACTS_MIN_CHARS_NO_SPACES
    return CODEX_MIN_CHARS_NO_SPACES


def validate_seo_opening(checks: list[dict[str, str]], description_html: str, product_name: str) -> None:
    opening_match = re.search(
        r"<p\b[^>]*>\s*<strong\b[^>]*>(.*?)</strong>",
        description_html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not opening_match:
        checks.append({
            "status": "WARNING",
            "check": "seo_opening",
            "details": "Pierwszy akapit powinien zaczynać się od pełnej nazwy produktu w tagu <strong>.",
        })
        return
    if product_name:
        opening_name = normalize_header(strip_html(opening_match.group(1)))
        expected_name = normalize_header(product_name)
        if opening_name != expected_name:
            checks.append({
                "status": "WARNING",
                "check": "seo_product_name",
                "details": "Wyróżniona nazwa w otwarciu nie jest pełną nazwą produktu z briefu.",
            })


def validate_seo_headings(
    checks: list[dict[str, str]],
    description_html: str,
    product_facts: list[dict[str, str]],
) -> None:
    headings = [
        compact_spaces(strip_html(value))
        for value in re.findall(r"<h2\b[^>]*>(.*?)</h2>", description_html, flags=re.IGNORECASE | re.DOTALL)
    ]
    if not headings:
        return
    generic_headings = {"najwazniejsze cechy", "zastosowanie i dopasowanie", "informacje o produkcie"}
    if any(normalize_header(heading) in generic_headings for heading in headings):
        checks.append({
            "status": "WARNING",
            "check": "seo_h2_specificity",
            "details": "Nagłówek <h2> jest generyczny; powinien zawierać naturalną frazę produktową.",
        })
    product_type = next(
        (compact_spaces(str(row.get("value", ""))) for row in product_facts if row.get("fact") == "Typ produktu"),
        "",
    )
    if product_type and not any(normalize_header(product_type) in normalize_header(heading) for heading in headings):
        checks.append({
            "status": "WARNING",
            "check": "seo_h2_product_type",
            "details": f"Nagłówek <h2> nie zawiera typu produktu '{product_type}'.",
        })


def validate_seo_keyword_placement(
    checks: list[dict[str, str]],
    description_html: str,
    seo_keyword: str,
    seo_rules: dict[str, Any],
) -> None:
    normalized_keyword = normalize_header(seo_keyword)
    normalized_text = normalize_header(strip_html(description_html))
    maximum = int(seo_rules.get("seo_keyword", {}).get("maximum_occurrences", 3))
    occurrences = len(re.findall(rf"(?<!\w){re.escape(normalized_keyword)}(?!\w)", normalized_text))
    if occurrences > maximum:
        checks.append({
            "status": "WARNING",
            "check": "seo_keyword_occurrences",
            "details": f"Fraza kluczowa '{seo_keyword}' występuje {occurrences} razy; maksimum to {maximum}.",
        })

    opening = re.search(r"<p\b[^>]*>(.*?)</p>", description_html, flags=re.IGNORECASE | re.DOTALL)
    opening_text = normalize_header(strip_html(opening.group(1))) if opening else ""
    if not opening_text.startswith(normalized_keyword):
        checks.append({
            "status": "WARNING",
            "check": "seo_keyword_opening",
            "details": f"Pierwszy akapit powinien zaczynać się od frazy kluczowej '{seo_keyword}'.",
        })

    h2_values = re.findall(r"<h2\b[^>]*>(.*?)</h2>", description_html, flags=re.IGNORECASE | re.DOTALL)
    if not any(normalized_keyword in normalize_header(strip_html(value)) for value in h2_values):
        checks.append({
            "status": "WARNING",
            "check": "seo_keyword_h2",
            "details": f"Nagłówek <h2> powinien zawierać frazę kluczową '{seo_keyword}'.",
        })

    after_h2 = re.search(
        r"</h2>\s*<p\b[^>]*>(.*?)</p>",
        description_html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    after_h2_text = normalize_header(strip_html(after_h2.group(1))) if after_h2 else ""
    if not after_h2_text.startswith(normalized_keyword):
        checks.append({
            "status": "WARNING",
            "check": "seo_keyword_after_h2",
            "details": f"Pierwszy akapit po <h2> powinien zaczynać się od frazy kluczowej '{seo_keyword}'.",
        })

    h3_values = re.findall(r"<h3\b[^>]*>(.*?)</h3>", description_html, flags=re.IGNORECASE | re.DOTALL)
    if any(normalized_keyword in normalize_header(strip_html(value)) for value in h3_values):
        checks.append({
            "status": "WARNING",
            "check": "seo_keyword_h3",
            "details": "Nagłówki <h3> nie powinny zawierać frazy kluczowej.",
        })


def validate_closing_summary(checks: list[dict[str, str]], description_html: str) -> None:
    summary_match = re.search(
        r"<h3\b[^>]*>\s*Specyfikacja techniczna\s*</h3>"
        r".*?</ul>\s*<p\b[^>]*>(.*?)</p>\s*(?:<!--.*?-->\s*)?$",
        description_html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not summary_match or not compact_spaces(strip_html(summary_match.group(1))):
        checks.append({
            "status": "WARNING",
            "check": "html_closing_summary",
            "details": (
                "Po liście specyfikacji brakuje końcowego akapitu <p> z podsumowaniem "
                "potwierdzonych zalet i naturalną zachętą do zakupu."
            ),
        })


def find_overused_word(text: str) -> str:
    normalized = normalize_header(text)
    thresholds = {
        "porzadkuje": 2,
        "wariant": 5,
        "ulatwia": 5,
        "pomaga": 5,
        "pozwala": 5,
    }
    for word, maximum in thresholds.items():
        if len(re.findall(rf"\b{word}\w*\b", normalized)) > maximum:
            return word
    return ""


def find_robotic_language_phrase(text: str) -> str:
    normalized = normalize_header(text)
    for pattern in ROBOTIC_LANGUAGE_PATTERNS:
        match = re.search(pattern, normalized)
        if match:
            return match.group(0)
    return ""


def has_encoding_damage(description_html: str) -> bool:
    if any(marker in description_html for marker in MOJIBAKE_MARKERS):
        return True
    text = strip_html(description_html)
    if "?" not in text:
        return False
    return bool(re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]\?|\?[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", text))


def find_meta_description_phrase(text: str) -> str:
    normalized = normalize_header(text)
    for pattern in META_DESCRIPTION_PATTERNS:
        match = re.search(pattern, normalized)
        if match:
            return match.group(0)
    return ""


def find_ascii_polish_marker(text: str) -> str:
    for pattern in ASCII_POLISH_MARKERS:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(0)
    return ""


def validate_codex_review_file(review_input: str | Path) -> list[dict[str, str]]:
    path = Path(review_input)
    if not path.exists():
        raise SystemExit(f"Nie znaleziono pliku review: {path}")
    with pd.ExcelFile(path) as xls:
        missing_sheets = [sheet for sheet in ["Produkt", "Product facts", "Rejected facts", "Opis wygenerowany"] if sheet not in xls.sheet_names]
        if missing_sheets:
            raise SystemExit(f"Review nie ma wymaganych arkuszy: {', '.join(missing_sheets)}")
        products = sheet_records(xls, "Produkt")
        product_facts = sheet_records(xls, "Product facts")
        rejected_facts = sheet_records(xls, "Rejected facts")
        descriptions = sheet_records(xls, "Opis wygenerowany")
    description_html = ""
    if descriptions:
        description_html = str(descriptions[0].get("description_html", "") or "")
    product_name = str(products[0].get("name", "") or "") if products else ""
    seo_keyword = str(products[0].get("seo_keyword", "") or "") if products else ""
    return validate_codex_description(
        description_html,
        product_facts,
        rejected_facts,
        product_name=product_name,
        seo_keyword=seo_keyword,
    )


def write_codex_review_from_description(
    brief_input: str | Path,
    review_output: str | Path,
    description_html: str,
) -> list[dict[str, str]]:
    brief = read_codex_brief_report(brief_input)
    brief["generated_description"] = description_html
    brief["validation"] = validate_codex_description(
        description_html,
        brief["product_facts"],
        brief["rejected_facts"],
        product_name=str(brief.get("product", {}).get("name", "") or ""),
        seo_keyword=str(brief.get("product", {}).get("seo_keyword", "") or ""),
    )
    write_codex_brief_report(brief, review_output)
    return brief["validation"]


def read_codex_brief_report(brief_input: str | Path) -> dict[str, Any]:
    path = Path(brief_input)
    if not path.exists():
        raise SystemExit(f"Nie znaleziono briefu: {path}")
    with pd.ExcelFile(path) as xls:
        missing_sheets = [sheet for sheet in CODEX_BRIEF_SHEETS if sheet not in xls.sheet_names]
        if missing_sheets:
            raise SystemExit(f"Brief nie ma wymaganych arkuszy: {', '.join(missing_sheets)}")
        product_rows = sheet_records(xls, "Produkt")
        product_facts = sheet_records(xls, "Product facts")
        compatibility_facts = sheet_records(xls, "Compatibility facts")
        rejected_facts = sheet_records(xls, "Rejected facts")
        prompt_rows = sheet_records(xls, "Prompt dla Codex")
    prompt = ""
    if prompt_rows:
        prompt = next(iter(prompt_rows[0].values()), "")
    return {
        "product": product_rows[0] if product_rows else {},
        "product_facts": product_facts,
        "compatibility_facts": compatibility_facts,
        "rejected_facts": rejected_facts,
        "prompt": prompt,
        "generated_description": "",
        "validation": [],
    }


def read_description_input(description_html: str, description_file: str) -> str:
    if description_html and description_file:
        raise SystemExit("Podaj tylko jedno: --description-html albo --description-file.")
    if description_file:
        path = Path(description_file)
        if not path.exists():
            raise SystemExit(f"Nie znaleziono pliku opisu: {path}")
        return path.read_text(encoding="utf-8")
    if description_html:
        return description_html
    raise SystemExit("Tryb write_codex_review wymaga --description-html albo --description-file.")


def sheet_records(xls: pd.ExcelFile, sheet_name: str) -> list[dict[str, str]]:
    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")
    return [{str(key): str(value) for key, value in row.items()} for row in df.to_dict(orient="records")]


def rejected_fact_used_as_product_claim(text: str, fact: str, value: str) -> bool:
    if not value:
        return False
    normalized = normalize_header(text)
    normalized_value = normalize_header(str(value))
    if fact == "Moc [W]":
        number = re.escape(normalized_value.replace("w", "").strip())
        return bool(re.search(rf"\b(?:moc|mocy|ma moc|o mocy)\s+{number}\s*w?\b", normalized))
    if fact and normalize_header(f"{fact}: {value}") in normalized:
        return True
    return False


def write_codex_brief_report(brief: dict[str, Any], output: str | Path) -> None:
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame([brief["product"]]).to_excel(writer, sheet_name="Produkt", index=False)
        pd.DataFrame(brief["product_facts"], columns=["fact", "value", "source"]).to_excel(writer, sheet_name="Product facts", index=False)
        pd.DataFrame(brief["compatibility_facts"], columns=["fact", "value", "source"]).to_excel(writer, sheet_name="Compatibility facts", index=False)
        pd.DataFrame(brief["rejected_facts"], columns=["fact", "value", "source", "reason"]).to_excel(writer, sheet_name="Rejected facts", index=False)
        pd.DataFrame([{"Prompt dla agenta Codex": brief["prompt"]}]).to_excel(writer, sheet_name="Prompt dla Codex", index=False)
        pd.DataFrame([{"description_html": brief["generated_description"]}]).to_excel(writer, sheet_name="Opis wygenerowany", index=False)
        pd.DataFrame(brief["validation"], columns=["status", "check", "details"]).to_excel(writer, sheet_name="Walidacja", index=False)
        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for column in worksheet.columns:
                width = max(len(str(cell.value or "")) for cell in column[:40])
                worksheet.column_dimensions[column[0].column_letter].width = min(max(width + 2, 14), 90)


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
    apply_power_range_for_description(row, attrs)
    return {key: value for key, value in attrs.items() if value}


def apply_power_range_for_description(row: pd.Series, attrs: dict[str, str]) -> None:
    power_range = normalized_power_range(first_present(row.get("Moc - zakres", ""), row.get("attr_moc_zakres", "")))
    if power_range:
        attrs["Moc [W]"] = power_range


def normalized_power_range(value: str) -> str:
    value = compact_spaces(str(value or ""))
    match = re.fullmatch(
        r"(\d+(?:[,.]\d+)?)\s*(?:W\s*)?(?:-|/|–|—)\s*(\d+(?:[,.]\d+)?)\s*W?",
        value,
        flags=re.IGNORECASE,
    )
    if not match:
        return ""
    left, right = (part.replace(",", ".") for part in match.groups())
    return f"{left}-{right}W"


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
