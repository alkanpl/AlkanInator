from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
import re
from html import escape
from pathlib import Path
from typing import Any

import pandas as pd

import export_to_baselinker_csv as baselinker
from utils import compact_spaces, ensure_dir, is_blank, normalize_header, read_products, write_products


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
    "Skuteczność świetlna [lm/W]",
    "Skuteczność świetlna",
    "Maksymalna moc źródła światła",
    "Źródło światła",
    "Trzonek",
    "Liczba źródeł światła",
    "Stopień ochrony [IP]",
    "Stopień odporności [IK]",
    "Klasa ochronności",
    "Klasa energetyczna",
    "Wskaźnik oddawania barw",
    "Wskaźnik olśnienia [UGR]",
    "Kąt świecenia [°]",
    "Kąt świecenia",
    "Czujnik ruchu",
    "Współpraca ze ściemniaczem",
    "Zasilanie",
    "Klosz",
    "Kolor",
    "Materiał",
    "Kształt",
    "Sposób montażu",
    "Wymiary",
    "Długość",
    "Szerokość",
    "Wysokość",
    "Średnica",
    "Trwałość [h]",
    "Pasuje do",
    "Sterowanie",
    "Gwarancja",
]

# Dane logistyczne - nie naleza do specyfikacji technicznej opisu (wg wytycznych SEO).
LOGISTIC_SKIP_FEATURES = {
    "Liczba sztuk",
    "Opakowanie",
    "Ilość w opakowaniu",
    "Ilość sztuk w opakowaniu jednostkowym",
    "Dane producenta",
    "Typ produktu",  # typ jest juz fraza kluczowa w prozie - nie dublujemy go w specyfikacji
}


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
    parser.add_argument("--sample", type=int, default=0, help="Tryb probki: wygeneruj N reprezentatywnych opisow (po rodzinach) do przegladu.")
    parser.add_argument("--sample-html", default="", help="Sciezka podgladu HTML probki.")
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

    if args.sample:
        df = select_sample(df, args.sample, args.title_column)
    result = add_descriptions(df, args.description_column, args.title_column, args.min_chars_no_spaces)

    if args.sample:
        write_sample_report(result, args.title_column, args.description_column, Path(args.output), args.sample_html)
        print(f"OK: probka {len(result)} opisow -> {args.output}" + (f" + {args.sample_html}" if args.sample_html else ""))
        return

    write_products(result, Path(args.output))
    print(f"OK: wczytano {len(df)} produktow")
    print(f"OK: zapisano {args.output}")


def select_sample(df: pd.DataFrame, n: int, title_column: str = "") -> pd.DataFrame:
    """Reprezentatywna probka: rownomiernie po rodzinach produktow (classify_family)."""
    by_family: dict[str, list[int]] = {}
    for index, row in df.iterrows():
        family = classify_family(product_title(row, title_column), collect_attributes(row))
        by_family.setdefault(family, []).append(index)
    chosen: list[int] = []
    families = list(by_family)
    position = 0
    while len(chosen) < n and any(by_family.values()):
        family = families[position % len(families)]
        if by_family[family]:
            chosen.append(by_family[family].pop(0))
        position += 1
        if position > n * len(families) + len(families):
            break
    return df.loc[chosen].copy()


def validate_generated_description(html: str, keyword: str, min_chars_no_spaces: int) -> list[str]:
    issues: list[str] = []
    main = html.split("<hr>")[0]
    main_text = strip_html(main)
    if chars_without_spaces(main_text) < min_chars_no_spaces:
        issues.append(f"opis krotszy niz {min_chars_no_spaces} znakow bez spacji")
    keyword_count = main_text.lower().count(keyword.lower())
    if keyword_count != 3:
        issues.append(f"fraza wystepuje {keyword_count}x (ma byc 3)")
    if not re.search(r"<h2>[^<]*\?\s*</h2>", main):
        issues.append("H2 nie jest pytaniem")
    after_h2 = re.search(r"</h2>\s*<p\b[^>]*>(.*?)</p>", main, flags=re.S | re.I)
    after_h2_text = strip_html(after_h2.group(1)).lower() if after_h2 else ""
    if not after_h2_text.startswith(keyword.lower()):
        issues.append("akapit po H2 nie zaczyna sie fraza kluczowa")
    for section in ["Najważniejsze cechy", "Specyfikacja techniczna"]:
        if section not in main:
            issues.append(f"brak sekcji: {section}")
    if "<hr>" not in html:
        issues.append("brak sekcji FAQ (<hr>)")
    for answer in re.findall(r"<br>\s*(.*?)</p>", html, flags=re.S):
        length = len(strip_html(answer).strip())
        if not 200 <= length <= 300:
            issues.append(f"odpowiedz FAQ ma {length} znakow (200-300)")
    for phrase in FORBIDDEN_PHRASES:
        if phrase in html:
            issues.append(f"zakazana fraza: {phrase}")
    issues.extend(algorithmic_style_issues(html))
    return issues


def write_sample_report(
    result: pd.DataFrame,
    title_column: str,
    description_column: str,
    output: Path,
    sample_html: str = "",
) -> None:
    rows: list[dict[str, str]] = []
    html_blocks: list[str] = []
    for _, row in result.iterrows():
        html = row[description_column]
        attrs = collect_attributes(row)
        name = product_title(row, title_column)
        keyword = build_seo_keyword(attrs, name)
        family = classify_family(name, attrs)
        issues = validate_generated_description(html, keyword, DEFAULT_MIN_CHARS_NO_SPACES)
        warnings = review_warnings(html, name, attrs, family)
        sku = first_present(row.get("sku", ""), row.get("SKU", ""), row.get("Kod", ""))
        rows.append({
            "sku": sku,
            "nazwa": name,
            "rodzina": family,
            "fraza_kluczowa": keyword,
            "walidacja": "OK" if not issues else " | ".join(issues),
            "uwagi": "" if not warnings else " | ".join(warnings),
            "opis_html": html,
        })
        uwagi_line = f" | uwagi: {escape(' | '.join(warnings))}" if warnings else ""
        html_blocks.append(f"<article><h1>{escape(sku)} — {escape(name)}</h1>\n<p><em>rodzina: {escape(family)} | walidacja: {escape('OK' if not issues else ' | '.join(issues))}{uwagi_line}</em></p>\n{html}\n</article>\n<hr><hr>")
    ensure_dir(output.parent)
    pd.DataFrame(rows, columns=["sku", "nazwa", "rodzina", "fraza_kluczowa", "walidacja", "uwagi", "opis_html"]).to_excel(output, index=False)
    if sample_html:
        preview = "<!DOCTYPE html>\n<html lang='pl'><head><meta charset='utf-8'><title>Probka opisow</title></head><body>\n" + "\n".join(html_blocks) + "\n</body></html>"
        Path(sample_html).write_text(preview, encoding="utf-8")


def description_signature(html: str) -> str:
    """Sygnatura frazowania (bez liczb) z prozy: wstep + akapit po H2 + naglowek H3 +
    akapit rozwijajacy + zamkniecie. Warianty tej samej serii (rozniace sie tylko
    wartosciami) daja te sama sygnature, wiec straznik wymusi inne sformulowania."""
    main = html.split("<hr>")[0]
    paragraphs = re.findall(r"<p>(.*?)</p>", main, flags=re.S)
    # Wstep to najbardziej widoczne zdanie - pilnujemy, zeby nie powtarzal sie miedzy produktami.
    intro = strip_html(paragraphs[0]) if paragraphs else ""
    return re.sub(r"\d+", "", normalize_header(intro))


def _dedupe_sentence_key(sentence: str) -> str:
    """Klucz znaczeniowy zdania: bez liczb, bez znakow, z odcieta KONCOWKA (2 ostatnie slowa).
    Dzieki temu '...do warunkow montazu.' i '...do warunkow pracy.' maja ten sam klucz."""
    norm = re.sub(r"[^a-z ]", " ", re.sub(r"\d+", "", normalize_header(strip_html(sentence))))
    words = [word for word in norm.split() if word]
    if len(words) < 5:
        return ""  # za krotkie zdania zostawiamy w spokoju (brak ryzyka, brak wartosci z dedupu)
    return " ".join(words[:-2])


def dedupe_prose_sentences(html: str, keyword: str, once_only_phrases: "list[str] | None" = None) -> str:
    """Usuwa z prozy (<p>) zdania o tym samym znaczeniu - nawet gdy roznia sie tylko koncowka
    lub liczbami (np. to samo zdanie o IP65 w dwoch akapitach). Dodatkowo `once_only_phrases`
    (dlugie formuly, np. pelne znaczenie IP) moga pasc w glownej tresci tylko RAZ - kolejne
    zdanie je zawierajace jest usuwane w calosci. Zdania z fraza kluczowa sa chronione
    (nie psujemy licznika frazy = 3). Listy, naglowki i FAQ pozostaja nietkniete."""
    seen: set[str] = set()
    seen_phrases: set[str] = set()
    kw_norm = normalize_header(keyword)
    once = [normalize_header(phrase) for phrase in (once_only_phrases or []) if phrase]

    def process(match: "re.Match[str]") -> str:
        # Dzielimy na zdania tylko gdy po kropce nastepuje WIELKA litera - inaczej skroty typu
        # "np." czy "śr." rozbijaja zdanie na fragmenty.
        sentences = re.split(r"(?<=[.!?])\s+(?=[A-ZĄĆĘŁŃÓŚŹŻ])", match.group(1).strip())
        kept: list[str] = []
        for raw in sentences:
            sentence = raw.strip()
            if not sentence:
                continue
            norm_sentence = normalize_header(strip_html(sentence))
            contains_kw = bool(kw_norm) and kw_norm in norm_sentence
            # Dlugie formuly (np. pelne znaczenie IP) tylko raz w glownej tresci.
            phrase_drop = False
            for phrase in once:
                if phrase and phrase in norm_sentence:
                    if phrase in seen_phrases and not contains_kw:
                        phrase_drop = True
                    seen_phrases.add(phrase)
                    break
            if phrase_drop:
                continue
            key = _dedupe_sentence_key(sentence)
            if key and key in seen and not contains_kw:
                continue
            if key:
                seen.add(key)
            kept.append(sentence)
        return f"<p>{' '.join(kept)}</p>" if kept else ""

    main, sep, faq = html.partition("<hr>")
    main = re.sub(r"<p>(.*?)</p>", process, main, flags=re.S)
    main = re.sub(r"\n{3,}", "\n\n", main)
    return main + sep + faq


def add_descriptions(
    df: pd.DataFrame,
    description_column: str = DEFAULT_DESCRIPTION_COLUMN,
    title_column: str = "",
    min_chars_no_spaces: int = DEFAULT_MIN_CHARS_NO_SPACES,
) -> pd.DataFrame:
    result = df.copy()
    descriptions: list[str] = []
    seen_signatures: set[str] = set()
    for _, row in result.iterrows():
        html = build_description_html(row, title_column, min_chars_no_spaces)
        offset = 0
        while description_signature(html) in seen_signatures and offset < 6:
            offset += 1
            html = build_description_html(row, title_column, min_chars_no_spaces, seed_offset=offset)
        seen_signatures.add(description_signature(html))
        descriptions.append(html)
    result[description_column] = descriptions
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
        "- Po części wprowadzającej dodaj sekcję <h3>Najważniejsze cechy</h3> z korzyściami wynikającymi z faktów.",
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
        "- sekcja <h3>Najważniejsze cechy</h3> i lista <ul><li> z konkretnymi korzyściami",
        "- sekcja <h3>Specyfikacja techniczna</h3> tylko z PRODUCT_FACTS; etykiety parametrów w <strong>",
        "- po liście specyfikacji obowiązkowy końcowy <p> z podsumowaniem i naturalną zachętą do zakupu",
        "- elementy listy zalet zaczynaj od myślnika po tagu, np. <li>- <strong>Cecha</strong> - korzyść.</li>",
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
        checks.append({"status": "WARNING", "check": "html_specification_h3", "details": "Brak nagłówka <h3>Specyfikacja techniczna:</h3>."})
    if not re.search(r"<h3\b[^>]*>\s*Najważniejsze cechy\s*</h3>", description_html, flags=re.IGNORECASE):
        checks.append({"status": "WARNING", "check": "html_benefits_h3", "details": "Brak sekcji <h3>Najważniejsze cechy</h3>."})
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
    seed_offset: int = 0,
) -> str:
    title = product_title(row, title_column)
    attrs = collect_attributes(row)
    family = classify_family(title, attrs)
    seed = stable_seed(first_present(row.get("sku", ""), row.get("SKU", ""), row.get("Kod", ""), title)) + seed_offset
    keyword = build_seo_keyword(attrs, title)
    kw_cap = uppercase_first_letter(keyword)
    kw_mid = lower_first_word(keyword)
    semantic_facts = derive_semantic_facts(attrs, family)
    description_plan = select_description_plan(family, semantic_facts, attrs)

    intro, intro_used_roles = realize_intro_from_plan(kw_cap, attrs, family, semantic_facts, description_plan, seed)
    intro = compact_spaces(f"{intro} {family_intro_role(family, seed)}")
    h2 = variant(H2_QUESTION_BUILDERS, seed // 5)(kw_mid)
    after_h2 = realize_after_h2_from_plan(kw_cap, attrs, family, semantic_facts, description_plan, intro_used_roles, seed)
    h3_dev = description_plan.h3
    advice = realize_development_from_plan(attrs, family, semantic_facts, description_plan, seed)
    benefits = build_decision_highlights(description_plan, semantic_facts, attrs, family)
    technical_rows = build_technical_rows(attrs)
    closing = variant(CLOSING_BUILDERS.get(family, CLOSING_BUILDERS["ogolny"]), seed // 11)(kw_mid, attrs, family)

    parts = [
        f"<p>{escape(intro)}</p>",
        "",
        f"<h2>{escape(h2)}</h2>",
        "",
        f"<p>{escape(after_h2)}</p>",
        "",
        f"<h3>{escape(h3_dev)}</h3>",
        "",
        f"<p>{escape(advice)}</p>",
        "",
        "<h3>Najważniejsze cechy</h3>",
        "",
        "<ul>",
        *[render_benefit_item(item) for item in benefits],
        "</ul>",
        "",
        "<h3>Specyfikacja techniczna</h3>",
        "",
        "<ul>",
        *[f"  <li><strong>{escape(label)}:</strong> {escape(value)}</li>" for label, value in technical_rows],
        "</ul>",
    ]
    html = "\n".join(parts)
    html = f"{html}\n\n<p>{escape(closing)}</p>"
    ip_value = attrs.get("Stopień ochrony [IP]", "")
    once_only = [ip_meaning(ip_value)] if ip_value else []
    html = dedupe_prose_sentences(html, keyword, once_only_phrases=once_only)
    html = ensure_minimum_length(html, keyword, attrs, family, seed, min_chars_no_spaces)
    html = f"{html}\n\n<hr>\n\n{build_faq(kw_cap, attrs, family, seed)}"
    validate_description(html)
    assert_semantic_safety(html, attrs, family)
    return html


def build_seo_keyword(attrs: dict[str, str], title: str) -> str:
    """Fraza kluczowa = naturalna nazwa produktu: typ + seria (np. 'Naświetlacz LED IQ-LED FL').

    Pelna nazwa ze specyfikacja (moc, lm, kolor) zostaje w tytule i specyfikacji; do
    3-krotnego powtorzenia w tekiscie uzywamy krotszej, naturalnie brzmiacej frazy.
    """
    typ = compact_spaces(attrs.get("Typ produktu", ""))
    seria = compact_spaces(attrs.get("Seria", ""))
    keyword = compact_spaces(" ".join(part for part in [typ, seria] if part))
    return keyword or keyword_from_title(title, attrs)


def lower_first_word(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return value
    return value[0].lower() + value[1:]


def uppercase_first_letter(value: str) -> str:
    value = compact_spaces(value)
    for index, char in enumerate(value):
        if char.isalpha():
            return value[:index] + char.upper() + value[index + 1:]
    return value


def has_feature(attrs: dict[str, str], key: str) -> bool:
    """True tylko gdy atrybut realnie wystepuje. Po rundzie v20 wartosci Tak/Nie sa wpisywane
    wprost, wiec 'Nie'/'Brak' jest truthy - to powodowalo zmyslanie cech (np. czujnik ruchu).
    """
    return normalize_header(attrs.get(key, "")) not in {"", "nie", "brak", "n/d", "nd", "0"}


LIGHT_FAMILIES = {"high_bay", "panel", "hermetyczna", "plafon", "naswietlacz", "oprawa", "ogrodowa", "ogolny"}


def is_light_product(family: str) -> bool:
    return family in LIGHT_FAMILIES


def is_channel_luminaire(attrs: dict[str, str]) -> bool:
    text = normalize_header(" ".join([attrs.get("Typ produktu", ""), attrs.get("Seria", ""), attrs.get("__working_title", "")]))
    return "oprawa kanalowa" in text or "oprawa kana" in text


def subject_for(family: str) -> str:
    return {
        "ramka": "Ta ramka",
        "zasilacz": "Ten zasilacz",
        "akcesorium": "To akcesorium",
        "panel": "Ten panel",
        "plafon": "Ten plafon",
        "naswietlacz": "Ten naświetlacz",
        "oprawa": "Ta oprawa",
        "hermetyczna": "Ta oprawa",
        "high_bay": "Ta oprawa",
        "ogrodowa": "Ta lampa",
    }.get(family, "Ten produkt")


def ip_meaning(ip: str) -> str:
    digits = re.sub(r"\D", "", ip)
    meanings = {
        "20": "ochronę odpowiednią do suchych wnętrz, bez narażenia na wodę i podwyższoną wilgotność",
        "23": "ochronę odpowiednią do suchych wnętrz oraz przed kroplami wody padającymi pod kątem",
        "40": "ochronę przed ciałami obcymi większymi niż 1 mm, bez ochrony przed wodą",
        "44": "ochronę przed bryzgami wody z każdej strony oraz ciałami obcymi większymi niż 1 mm",
        "54": "ochronę przed pyłem i zachlapaniem wodą",
        "65": "pełną pyłoszczelność i odporność na strugę wody, co pozwala na montaż na zewnątrz",
        "66": "pełną pyłoszczelność i odporność na silne strugi wody",
        "67": "pełną pyłoszczelność i ochronę przy krótkotrwałym zanurzeniu w wodzie",
    }
    return meanings.get(digits, "warunki podane w klasie szczelności")


def ip_resistance_line(ip: str) -> str:
    """Pelne, poprawne merytorycznie zdanie o klasie IP. WAZNE: IP20/IP40 (druga cyfra 0/1) to
    SUCHE WNETRZA - nie wolno pisac o 'pyloszczelnosci' ani 'odpornosci na wilgoc'."""
    d = re.sub(r"\D", "", ip)
    if d in {"65", "66", "67"}:
        return f"Klasa {ip} oznacza wysoką odporność na pył i wodę, ale miejsce montażu nadal powinno być zgodne z instrukcją producenta."
    if d == "54":
        return f"Klasa {ip} oznacza ochronę przed pyłem i zachlapaniem wodą z każdej strony."
    if d == "44":
        return f"Klasa {ip} oznacza ochronę przed bryzgami wody i drobnymi ciałami obcymi, dlatego produkt można rozważyć w miejscach narażonych na lekkie zachlapanie."
    if d and d[-1] in "01":  # druga cyfra 0/1 = brak/znikoma ochrona przed woda -> suche wnetrza
        return f"Klasa {ip} oznacza ochronę odpowiednią do suchych wnętrz. Produktu nie należy montować w miejscach narażonych na wodę, zachlapanie lub podwyższoną wilgotność."
    return f"Klasa {ip} oznacza {ip_meaning(ip)}."


def environment_usage_sentences(ip: str, family: str = "") -> list[str]:
    """Holistyczne zdania o odpornosci. WAZNE: IP mowi o ODPORNOSCI, nie o rekomendowanym
    miejscu - dlatego rekomendacje 'na zewnatrz/taras/ogrod' TYLKO dla rodzin zewnetrznych
    (ogrodowa/naswietlacz). High Bay/hermetyczna z IP65 = 'znosi trudne warunki', nie 'ogrod'."""
    d = re.sub(r"\D", "", ip)
    outdoor = family in {"ogrodowa", "naswietlacz"}
    if d in {"65", "66", "67"}:
        if outdoor:
            return [
                "Sprawdzi się również na zewnątrz – poradzi sobie z deszczem, pyłem i wilgocią.",
                "Nadaje się do montażu na dworze, także w miejscu narażonym na deszcz i kurz.",
                "Nadaje się do pracy na zewnątrz, bo nie boi się deszczu ani zapylenia.",
                "Odporna obudowa pozwala na pracę na zewnątrz przez cały rok, niezależnie od pogody.",
                "Bez obaw sprawdzi się pod gołym niebem – poradzi sobie z deszczem i zapyleniem.",
                "Nada się też na taras, elewację czy do ogrodu, bo dobrze znosi warunki atmosferyczne.",
            ]
        return [
            f"Wysoka szczelność {ip} pozwala pracować w kurzu, wilgoci i trudnych warunkach.",
            f"Klasa {ip} oznacza dużą odporność na pył i wodę, przydatną tam, gdzie warunki bywają wymagające.",
            "Odporna, szczelna obudowa dobrze znosi kurz, wilgoć i intensywną eksploatację.",
            f"Dzięki klasie {ip} nie boi się zapylenia ani zachlapania w miejscu pracy.",
        ]
    if d == "54":
        return [
            "Poradzi sobie z kurzem i zachlapaniem, więc sprawdzi się też w pomieszczeniach gospodarczych i pod zadaszeniem na zewnątrz.",
            "Zniesie pył i bryzgi wody, dlatego nada się również do garażu, piwnicy czy zadaszonego tarasu.",
            "Odporność na pył i zachlapanie pozwala na montaż w warsztacie, piwnicy czy pod wiatą.",
            "Dobrze sprawdza się tam, gdzie bywa kurz i wilgoć – w pomieszczeniach gospodarczych i pod zadaszeniem.",
        ]
    if d == "44":
        return [
            "Zniesie bryzgi wody, dlatego może sprawdzić się w wybranych miejscach łazienki lub pod zadaszeniem, jeśli montaż jest zgodny z klasą IP44 i instrukcją.",
            "Poradzi sobie z bryzgami wody, więc bywa odpowiedni do wilgotniejszych pomieszczeń – o ile miejsce montażu odpowiada klasie IP44.",
            "Odporność na bryzgi wody pozwala rozważyć montaż w łazience lub pralni, w strefie zgodnej z instrukcją producenta.",
            "Nada się także do wilgotniejszych miejsc, bo dobrze znosi pojedyncze zachlapania, przy montażu zgodnym z warunkami klasy IP44.",
        ]
    return [
        "Nadaje się do suchych wnętrz – pokoi, biur czy korytarzy, z dala od wody i wilgoci.",
        "Najlepiej sprawdzi się w suchych pomieszczeniach, bez narażenia na wodę i większą wilgoć.",
        "To rozwiązanie do suchych wnętrz, z dala od zachlapania i podwyższonej wilgotności.",
        "Przeznaczony jest do suchych pomieszczeń, gdzie nie ma kontaktu z wodą.",
        "Najlepiej czuje się w typowych, suchych wnętrzach mieszkalnych i biurowych.",
    ]


def ip_short(ip: str) -> str:
    """Skrocone znaczenie IP do PUNKTOW listy - pelna formula (ip_meaning) zostaje w prozie,
    zeby nie powtarzac tego samego zdania dwa razy w jednym opisie."""
    digits = re.sub(r"\D", "", ip)
    short = {
        "20": "do suchych wnętrz",
        "23": "do suchych wnętrz",
        "40": "do suchych wnętrz",
        "44": "ochrona przed bryzgami wody",
        "54": "ochrona przed pyłem i zachlapaniem",
        "65": "montaż także na zewnątrz",
        "66": "montaż także na zewnątrz",
        "67": "montaż także na zewnątrz",
    }
    return short.get(digits, "warunki pracy wg klasy szczelności")


_COLOR_VARIABLE = [
    "światło o regulowanej barwie – od ciepłej do chłodnej zależnie od potrzeby",
    "barwę światła, którą ustawisz od ciepłej do chłodnej według nastroju",
    "możliwość płynnej zmiany barwy – cieplejszej wieczorem, chłodniejszej do pracy",
]
_COLOR_WARM = [
    "ciepłe, przytulne światło, które ociepla wnętrze i buduje nastrój",
    "ciepłe światło, które tworzy przytulny, kameralny klimat",
    "miękkie, ciepłe światło sprzyjające odpoczynkowi i relaksowi",
]
_COLOR_COOL = [
    "chłodne, wyraziste światło, dobre tam, gdzie liczy się dobra widoczność",
    "chłodne, energetyczne światło poprawiające koncentrację i widoczność",
    "zimne, klarowne światło, które dobrze sprawdza się w pracy precyzyjnej",
]
_COLOR_NEUTRAL = [
    "neutralne, czytelne światło, przy którym wygodnie widać szczegóły",
    "neutralne światło zbliżone do dziennego, komfortowe dla oczu",
    "naturalne, czytelne światło dobre do pracy i codziennych czynności",
]


def color_temp_effect(value: str, seed: int = 0) -> str:
    """Z temperatury/barwy -> opis EFEKTU swiatla (calosc), a nie 'barwa = X K'."""
    n = normalize_header(value)
    if re.search(r"\d{3,}\s*\D+\s*\d{3,}", value) or "zmien" in n:
        return pick_variant(_COLOR_VARIABLE, seed)
    digits = re.sub(r"\D", "", value)
    if digits.isdigit() and digits:
        t = int(digits)
        if t <= 3300:
            return pick_variant(_COLOR_WARM, seed)
        if t >= 5000:
            return pick_variant(_COLOR_COOL, seed)
        return pick_variant(_COLOR_NEUTRAL, seed)
    if "ciep" in n:
        return pick_variant(_COLOR_WARM, seed)
    if "zimn" in n or "chlod" in n:
        return pick_variant(_COLOR_COOL, seed)
    return pick_variant(_COLOR_NEUTRAL, seed)


def color_locative(value: str) -> str:
    """Odmiana koloru po 'w kolorze ...': bialy->bialym, niebieski->niebieskim, szary->szarym.
    Obsluguje zlozenia 'bialy / czarny' -> 'bialym / czarnym'."""
    def decline(word: str) -> str:
        word = word.strip().lower()
        if word.endswith("y"):
            return word[:-1] + "ym"
        if word.endswith("i"):
            return word[:-1] + "im"
        return word
    parts = [decline(p) for p in re.split(r"\s*/\s*", value) if p.strip()]
    return " / ".join(parts)


# Korzysci koloru (do listy "Najwazniejsze cechy" i prozy) - rozne, by sie nie powtarzaly.
COLOR_FIT_BENEFITS = [
    "łatwiej dopasować do otoczenia",
    "spokojnie komponuje się z aranżacją wnętrza",
    "nie rzuca się w oczy po zamontowaniu",
    "dobrze wpisuje się w sufit i ściany",
    "łatwo zestawić z wystrojem pomieszczenia",
    "neutralnie wygląda w różnych aranżacjach",
]
COLOR_FIT_BENEFITS_OUTDOOR = [
    "łatwo dopasować do elewacji, tarasu lub aranżacji ogrodu",
    "dobrze komponuje się z elewacją i otoczeniem domu",
    "nie rzuca się w oczy na tle elewacji i zieleni",
    "spokojnie wpisuje się w architekturę ogrodu i posesji",
    "łatwo zestawić z elewacją, ogrodzeniem czy tarasem",
]


def color_fit_benefit(family: str, seed: int) -> str:
    pool = COLOR_FIT_BENEFITS_OUTDOOR if family in {"ogrodowa", "naswietlacz"} else COLOR_FIT_BENEFITS
    return pick_variant(pool, seed)


FAMILY_INTRO_ROLE = {
    "ogrodowa": [
        "Dobrze rozmieszczone światło w ogrodzie podnosi bezpieczeństwo wieczornego dojścia i podkreśla aranżację posesji.",
        "Oświetlenie wokół domu ułatwia poruszanie się po zmroku i nadaje otoczeniu przytulny charakter.",
        "Taka lampa pełni dwie role naraz: poprawia widoczność i buduje nastrój wokół budynku.",
    ],
    "panel": [
        "Równe, spokojne światło na większej powierzchni ułatwia pracę i ogranicza męczące kontrasty na suficie.",
        "Panel dobrze zastępuje starsze oprawy świetlówkowe tam, gdzie liczy się jednolite oświetlenie.",
        "Jednolite światło na całej płaszczyźnie sprzyja koncentracji i komfortowi pracy.",
    ],
    "plafon": [
        "Prosta forma i szybkie włączanie sprawiają, że oprawa dobrze odnajduje się w często używanych przejściach.",
        "W codziennie używanych pomieszczeniach liczy się światło, które działa od razu i nie wymaga obsługi.",
        "Niewielka oprawa sufitowa porządkuje oświetlenie wejść, korytarzy i pomieszczeń pomocniczych.",
    ],
    "high_bay": [
        "Z dużej wysokości liczy się mocne, równomierne światło, które nie zostawia ciemnych stref na powierzchni roboczej.",
        "W wysokich halach najważniejsza jest oprawa utrzymująca użyteczny strumień światła po dotarciu do podłogi.",
        "Dobre oświetlenie hali przekłada się na komfort i bezpieczeństwo pracy przez wiele godzin.",
    ],
    "hermetyczna": [
        "W trudnych warunkach najważniejsze jest światło, które działa pewnie mimo kurzu, wilgoci i częstej eksploatacji.",
        "Oprawa do pomieszczeń technicznych ma przede wszystkim świecić niezawodnie i nie wymagać częstej obsługi.",
        "W garażu czy warsztacie liczą się odporna obudowa i przewidywalne, mocne światło.",
    ],
    "naswietlacz": [
        "Skierowane światło na zewnątrz poprawia orientację i bezpieczeństwo po zmroku, nie oświetlając przy tym przypadkowych miejsc.",
        "Dobrze ustawiony naświetlacz doświetla dokładnie tę strefę, która tego potrzebuje.",
        "Oświetlenie zewnętrzne zwiększa poczucie bezpieczeństwa i ułatwia korzystanie z terenu po zmroku.",
    ],
    "oprawa": [
        "Dobrze dobrana oprawa łączy funkcję światła z wyglądem, który pasuje do wnętrza.",
        "We wnętrzu liczy się zarówno jakość światła, jak i to, jak oprawa prezentuje się po montażu.",
        "Oprawa tworzy konkretny punkt światła i wpływa na odbiór całego pomieszczenia.",
    ],
    "ramka": [
        "Estetyczne obramowanie porządkuje montaż i sprawia, że panel wygląda jak zaplanowany element wykończenia.",
        "Ramka pozwala zamontować panel tam, gdzie nie ma typowego sufitu kasetonowego, bez utraty estetyki.",
        "Dobrze dobrana ramka ułatwia montaż i pozwala zachować spójny wygląd kilku opraw w pomieszczeniu.",
    ],
    "zasilacz": [
        "Dobrze dobrany zasilacz to podstawa stabilnej i bezawaryjnej pracy całej oprawy.",
        "To element, od którego zależy, czy oprawa LED pracuje pewnie i zgodnie z przeznaczeniem.",
        "Zgodny zasilacz ogranicza ryzyko migotania światła i przedwczesnej awarii.",
    ],
    "akcesorium": [
        "Drobny, ale właściwie dobrany element decyduje o tym, czy oprawa działa pewnie i wygląda kompletnie.",
        "Dobrze dopasowane akcesorium pozwala utrzymać instalację w pełnej sprawności bez wymiany całej oprawy.",
        "To część, przy której najważniejsza jest zgodność z konkretną oprawą lub serią.",
    ],
    "ogolny": [
        "Dobrze dobrane oświetlenie poprawia komfort i wygląd przestrzeni, w której pracuje.",
        "Właściwa oprawa łączy użyteczne światło z dopasowaniem do miejsca montażu.",
        "Liczy się nie tylko sam strumień światła, ale też dopasowanie oprawy do funkcji pomieszczenia.",
    ],
}


def family_intro_role(family: str, seed: int) -> str:
    pool = FAMILY_INTRO_ROLE.get(family, FAMILY_INTRO_ROLE["ogolny"])
    return pick_variant(pool, seed // 3)


# "Czym jest" - kategoria produktu (po nazwie: "{nazwa} {definicja}.").
FAMILY_DEFINITION = {
    "panel": ["to energooszczędna oprawa sufitowa", "to płaska oprawa dająca równe, rozproszone światło", "to nowoczesna oprawa do sufitów modułowych i podwieszanych"],
    "plafon": ["to natynkowa oprawa sufitowa", "to prosta oprawa do montażu na suficie lub ścianie", "to oprawa sufitowa do wnętrz użytkowych"],
    "oprawa": ["to oprawa oświetleniowa do wnętrz", "to dekoracyjno-użytkowa oprawa do pomieszczeń", "to oprawa do montażu na suficie lub ścianie"],
    "high_bay": ["to mocna oprawa do wysokich pomieszczeń", "to przemysłowa oprawa typu high bay do hal i magazynów", "to oprawa do montażu na dużej wysokości"],
    "hermetyczna": ["to pyło- i bryzgoszczelna oprawa liniowa", "to odporna oprawa techniczna", "to hermetyczna oprawa do trudnych warunków"],
    "naswietlacz": ["to zewnętrzny reflektor LED", "to mocny reflektor do oświetlenia terenu", "to oprawa do doświetlania elewacji i placów"],
    "ogrodowa": ["to zewnętrzna oprawa do oświetlenia ogrodu i otoczenia domu", "to dekoracyjna lampa do ogrodu i tarasu", "to oprawa do oświetlenia ścieżek, zieleni i wejścia"],
    "ramka": ["to element montażowy do paneli LED", "to obramowanie ułatwiające montaż panelu na suficie", "to ramka do natynkowego osadzenia panelu"],
    "zasilacz": ["to element zasilający do opraw i paneli LED", "to zasilacz do instalacji oświetleniowej LED", "to podzespół do kompletacji oświetlenia LED"],
    "akcesorium": ["to element uzupełniający instalację oświetleniową", "to część do opraw i systemów oświetleniowych", "to akcesorium do montażu lub serwisu oświetlenia"],
    "ogolny": ["to oprawa oświetleniowa", "to element instalacji oświetleniowej", "to produkt z zakresu oświetlenia"],
}


def family_definition(family: str, seed: int) -> str:
    return pick_variant(FAMILY_DEFINITION.get(family, FAMILY_DEFINITION["ogolny"]), seed // 2)


# Zdania ZASTOSOWANIA per rodzina (ogolne prawdy, grounded) do akapitu rozwijajacego.
FAMILY_USE = {
    "panel": [
        "Najlepiej rozłożyć takie panele równomiernie nad strefą pracy, aby światło było jednolite na całej powierzchni.",
        "Sprawdza się w biurach, salach, gabinetach i punktach obsługi, gdzie liczy się komfort patrzenia.",
        "W otwartej przestrzeni biurowej pomaga ograniczyć cienie i nierówności oświetlenia.",
        "Dobrze zastępuje starsze, rastrowe oprawy świetlówkowe podczas modernizacji.",
        "Świetnie czuje się tam, gdzie światło ma być obecne, ale nie narzucające się.",
    ],
    "plafon": [
        "Najczęściej trafia do korytarzy, przedpokojów, klatek schodowych i pomieszczeń gospodarczych.",
        "Sprawdza się tam, gdzie światło ma włączać się szybko i obejmować całą strefę przejścia.",
        "Dobrze odnajduje się w miejscach często odwiedzanych, ale na krótko.",
        "To praktyczne źródło światła do wnętrz, w których nie potrzeba dekoracyjnej oprawy.",
        "Pasuje do niskich i typowych sufitów, gdzie liczy się prostota i równe światło.",
    ],
    "oprawa": [
        "Buduje konkretny punkt światła w salonie, sypialni, przedpokoju lub lokalu usługowym.",
        "Sprawdza się tam, gdzie oprawa ma być elementem aranżacji, a nie tylko źródłem światła.",
        "Dobrze komponuje się z meblami i wykończeniem, jeśli dobierze się ją do stylu wnętrza.",
        "To wybór do wnętrz, w których liczy się i światło, i wygląd po zamontowaniu.",
        "Pasuje do codziennego użytkowania w domu oraz w lokalach usługowych.",
    ],
    "high_bay": [
        "Montuje się je wysoko nad halą produkcyjną, magazynem lub strefą kompletacji.",
        "Sprawdza się tam, gdzie zwykłe oprawy sufitowe nie dają już wystarczającego zasięgu światła.",
        "Dobrze rozłożone, eliminuje ciemne strefy między regałami i ciągami komunikacyjnymi.",
        "To oprawa do intensywnej, wielogodzinnej pracy obiektu.",
        "Pasuje do wnętrz o dużej kubaturze, gdzie liczy się mocne, równe światło.",
    ],
    "hermetyczna": [
        "Najczęściej trafia do garaży, warsztatów, piwnic, hal i pomieszczeń gospodarczych.",
        "Sprawdza się tam, gdzie oprawa musi znosić kurz, wilgoć i codzienne zabrudzenia.",
        "Szczelna obudowa ułatwia utrzymanie czystości i ogranicza wnikanie pyłu.",
        "To praktyczne, odporne źródło światła do warunków technicznych.",
        "Pasuje do miejsc, w których światło ma działać pewnie mimo trudnego otoczenia.",
    ],
    "naswietlacz": [
        "Kieruje strumień światła na elewację, podjazd, bramę, plac lub strefę dostaw.",
        "Sprawdza się tam, gdzie po zmroku potrzebna jest dobra widoczność na konkretnym obszarze.",
        "Dobrze ustawiony, oświetla używaną strefę, nie świecąc przypadkowo poza teren.",
        "To praktyczne oświetlenie zewnętrzne, które poprawia bezpieczeństwo po zmroku.",
        "Pasuje wszędzie tam, gdzie trzeba doświetlić wybrany fragment otoczenia budynku.",
    ],
    "ogrodowa": [
        "Oświetla miejsca, z których korzystasz na co dzień: furtkę, ścieżkę, wejście lub taras.",
        "Sprawdza się jako element małej architektury, łącząc funkcję z dekoracją.",
        "Dobrze rozmieszczona, podkreśla zieleń, rabaty i detale wokół domu.",
        "To oświetlenie, które poprawia bezpieczeństwo wieczornego dojścia i buduje nastrój.",
        "Pasuje do ogrodu, tarasu i strefy wypoczynku wokół budynku.",
    ],
    "ramka": [
        "Przydaje się, gdy panel ma trafić na sufit pełny, bez kasetonu.",
        "Pozwala osadzić panel równo i estetycznie, jak zaplanowany element wykończenia.",
        "Sprawdza się przy modernizacji, gdy zmienia się sposób montażu oświetlenia.",
        "Porządkuje wygląd kilku paneli osadzonych w jednym pomieszczeniu.",
        "Pasuje do biur, lokali i korytarzy, gdzie liczy się spójne wykończenie sufitu.",
    ],
    "zasilacz": [
        "Stanowi część konkretnego zestawu oświetleniowego, a nie uniwersalny dodatek.",
        "Sprawdza się tam, gdzie liczy się zgodność elektryczna i stabilna praca oprawy.",
        "Dobrze dobrany, ogranicza ryzyko migotania i przedwczesnej awarii.",
        "To element, od którego zależy poprawne i bezpieczne działanie oświetlenia LED.",
        "Pasuje do instalacji, w której trzeba uzupełnić lub wymienić źródło zasilania.",
    ],
    "akcesorium": [
        "Najczęściej dobiera się je do istniejącej już oprawy lub serii.",
        "Sprawdza się przy uzupełnianiu, serwisie lub odświeżeniu instalacji.",
        "Dobrze dopasowane, pozwala uniknąć wymiany całej oprawy.",
        "To drobny element, przy którym najważniejsza jest zgodność z konkretnym produktem.",
        "Pasuje tam, gdzie trzeba przywrócić kompletność lub sprawność zestawu.",
    ],
    "ogolny": [
        "Sprawdza się jako praktyczny punkt światła w typowej instalacji.",
        "Dobrze dobrana, łączy użyteczne światło z dopasowaniem do miejsca.",
        "To rozwiązanie do codziennego oświetlenia wnętrz i przejść.",
        "Pasuje wszędzie tam, gdzie liczy się prosty montaż i przewidywalne światło.",
        "Najlepiej dobrać ją do funkcji pomieszczenia i oczekiwanego efektu.",
    ],
}

# Ogolne, bezpieczne zdania o jakosci/wykonaniu (do akapitu rozwijajacego).
DEV_QUALITY_COMMON = [
    "Solidne wykonanie i dobre materiały przekładają się na trwałość w codziennym użytkowaniu.",
    "Prosta, przemyślana konstrukcja ułatwia montaż i późniejszą eksploatację.",
    "Stonowana forma sprawia, że produkt łatwo wpisuje się w różne aranżacje.",
    "To rozwiązanie pomyślane tak, aby działało pewnie i nie sprawiało kłopotu na co dzień.",
    "Estetyczne wykończenie idzie tu w parze z praktycznym podejściem do użytkowania.",
    "Dobre dopasowanie do miejsca montażu przekłada się na komfort korzystania przez lata.",
    "Przemyślane proporcje sprawiają, że produkt dobrze wygląda po zamontowaniu.",
    "Liczy się tu nie efektowny opis, lecz to, że produkt po prostu spełnia swoje zadanie.",
    "Staranne wykonanie przekłada się na spokojną, bezproblemową pracę przez długi czas.",
    "Czysta, nieprzeładowana forma dobrze znosi próbę czasu i mody.",
    "Produkt stawia na funkcjonalność, bez zbędnych dodatków podnoszących cenę.",
    "Wygodny montaż i prosta obsługa to atuty, które docenia się na co dzień.",
    "Trwałe materiały sprawiają, że oprawa zachowuje wygląd i sprawność przez lata.",
    "Czytelne dane techniczne ułatwiają dopasowanie oprawy do miejsca montażu.",
]


def dev_quality_sentence(seed: int) -> str:
    # niezalezny hash, zeby nie korelowac ze zdaniem zastosowania (i uniknac kolizji modulo)
    return pick_variant(DEV_QUALITY_COMMON, stable_seed(f"quality|{seed}"))


def concrete_benefit_sentence(attrs: dict[str, str], family: str, seed: int) -> str:
    """Zdanie oparte na KONKRETNYM atrybucie (IP/IK/lm/W/czujnik/sciemnianie/solar/trwalosc/
    trzonek/barwa/wymiary). Family-aware: IP mowi o odpornosci, nie o rekomendowanym miejscu."""
    cands: list[str] = []
    lightp = is_light_product(family)
    outdoor = family in {"ogrodowa", "naswietlacz"}
    ip = attrs.get("Stopień ochrony [IP]", "")
    ik = attrs.get("Stopień odporności [IK]", "")
    # WAZNE: nie zaczynac zdan od nazwy parametru (walidator karze >=4 takich zdan) - liczby
    # w srodku zdania, lead czasownikowy/podmiotowy, zeby proza nie brzmiala jak tabela.
    if ip and ik:
        cands.append(f"Dzięki klasie {ip} i odporności mechanicznej {ik} oprawa znosi kurz, wilgoć i przypadkowe uderzenia.")
    elif ip:
        d = re.sub(r"\D", "", ip)
        if d in {"65", "66", "67"}:
            cands.append(
                f"Dzięki klasie {ip} można montować oprawę na zewnątrz, w deszczu i przy zapyleniu." if outdoor
                else f"Dzięki klasie {ip} oprawa pracuje w kurzu i wilgoci bez utraty niezawodności."
            )
        elif d == "54":
            cands.append(f"Dzięki klasie {ip} oprawa znosi pył i zachlapanie w trudniejszych warunkach.")
        elif d == "44":
            cands.append(f"Dzięki klasie {ip} oprawa jest zabezpieczona przed bryzgami wody.")
        else:
            cands.append(f"Produkt jest przeznaczony do suchych wnętrz (klasa {ip}), bez kontaktu z wodą.")
    lm = attrs.get("Strumień świetlny [lm]", "")
    watt = attrs.get("Moc [W]", "")
    if lightp and lm and watt:
        cands.append(f"Oprawa daje {lm} lm przy {watt} W, co ułatwia ocenę, czy jasność będzie odpowiednia dla wybranej strefy.")
    elif lightp and lm:
        cands.append(f"Przy strumieniu {lm} lm łatwo oszacować, ile światła da oprawa w docelowym miejscu.")
    if lightp and has_feature(attrs, "Czujnik ruchu"):
        cands.append("Wbudowany czujnik ruchu włącza światło automatycznie i ogranicza zużycie energii.")
    if has_feature(attrs, "Współpraca ze ściemniaczem"):
        cands.append(
            "Współpraca ze ściemniaczem pozwala płynnie regulować jasność podłączonej oprawy." if family == "zasilacz"
            else "Współpraca ze ściemniaczem pozwala płynnie regulować jasność światła."
        )
    if normalize_header(attrs.get("Zasilanie", "")) == "solarne":
        cands.append("Zasilanie solarne nie wymaga prowadzenia przewodu, co upraszcza montaż w terenie.")
    durability = attrs.get("Trwałość [h]", "")
    if durability:
        cands.append(f"Deklarowane {durability} h pracy ogranicza częstość i koszt wymian.")
    # Trzonek celowo NIE jest tu kandydatem: dobor zarowki omawia juz akapit po H2 (fakt
    # lamp_choice) oraz lista cech - powtorzenie E27 w sasiednim zdaniu brzmialo redundantnie.
    temp = attrs.get("Temperatura barwowa [K]", "")
    if lightp and temp and "-" not in temp:
        cands.append(f"Światło w barwie {temp} K daje {color_temp_effect(temp, seed)}.")
    dim = dim_compact(attrs)
    if dim and not dim.startswith("średnica"):
        if family == "zasilacz":
            cands.append(f"Znając wymiary {dim}, łatwo zaplanujesz miejsce na zasilacz i dostęp serwisowy.")
        elif family == "ramka":
            cands.append(f"Format {dim} pomaga dopasować ramkę do panelu i zachować estetyczne obramowanie.")
        else:
            cands.append(f"Znając wymiary {dim}, prościej zaplanujesz miejsce montażu.")
    if not cands:
        return ""
    return pick_variant(cands, seed)


INDOOR_LIGHT_FAMILIES = {"high_bay", "panel", "hermetyczna", "plafon", "oprawa", "ogolny"}


def _int_or_none(value: str) -> "int | None":
    digits = re.sub(r"\D", "", value.split("-")[0].split("/")[0])
    return int(digits) if digits else None


def advisory_reasoning_sentence(attrs: dict[str, str], family: str, seed: int) -> str:
    """Zdanie DORADCZE (pkt 8): z wartosci atrybutu wyprowadza wniosek dla kupujacego
    (duzy strumien -> wieksza przestrzen; chlodna barwa -> praca/technika; ciepla -> dom;
    IK -> narazenie na uderzenia). Family-aware, lead nie-parametrowy, wartosc w srodku zdania."""
    cands: list[str] = []
    lightp = is_light_product(family)
    outdoor = family in {"ogrodowa", "naswietlacz"}
    lm = _int_or_none(attrs.get("Strumień świetlny [lm]", "")) if lightp else None
    if lm:
        if lm >= 8000:
            cands.append(
                f"Tak duży strumień {lm} lm oświetli większy fragment terenu bez stawiania kolejnych słupów." if outdoor
                else f"Tak duży strumień {lm} lm rozświetli obszerne lub wysokie wnętrze bez dostawiania kolejnych opraw."
            )
        elif lm >= 3000:
            cands.append(f"Strumień na poziomie {lm} lm wystarcza do oświetlenia typowego pomieszczenia użytkowego jedną oprawą.")
        else:
            cands.append(f"Umiarkowany strumień {lm} lm sprawdzi się jako doświetlenie punktowe lub uzupełnienie istniejącego oświetlenia.")
    temp_raw = attrs.get("Temperatura barwowa [K]", "")
    temp = _int_or_none(temp_raw) if lightp else None
    home = family in {"plafon", "oprawa", "panel", "ogolny"}
    industrial = family in {"high_bay", "hermetyczna"}
    if temp and "-" not in temp_raw and "/" not in temp_raw:
        if temp >= 5000 and family in INDOOR_LIGHT_FAMILIES:
            cands.append(f"Chłodna barwa {temp} K sprzyja koncentracji, dlatego dobrze pracuje w biurach, halach, garażach i pomieszczeniach technicznych.")
        elif temp <= 3000 and home:
            cands.append(f"Ciepła barwa {temp} K buduje przytulny nastrój, więc najlepiej odnajdzie się w częściach mieszkalnych.")
        elif 3500 <= temp <= 4500 and industrial:
            cands.append(f"Neutralna barwa {temp} K daje czytelne, robocze światło do pracy w hali, magazynie lub warsztacie.")
        elif 3500 <= temp <= 4500 and home:
            cands.append(f"Neutralna barwa {temp} K jest uniwersalna i pasuje zarówno do domu, jak i przestrzeni użytkowych.")
    ik = attrs.get("Stopień odporności [IK]", "")
    if ik:
        cands.append(f"Podwyższona odporność mechaniczna {ik} przydaje się tam, gdzie oprawa jest narażona na uderzenia — w halach, garażach czy na zapleczu.")
    if not cands:
        return ""
    return pick_variant(cands, seed)


# Synonimy "sprawdzi sie" - pasuja do tej samej skladni (po nich nastepuje "w/przy/na ...").
SUITS_PHRASES = [
    "sprawdzi się",
    "dobrze odnajdzie się",
    "dobrze poradzi sobie",
    "znajdzie zastosowanie",
    "spełni swoją rolę",
    "będzie dobrym wyborem",
    "świetnie się sprawdza",
    "przyda się",
    "dobrze się sprawdza",
    "zda egzamin",
]


def suits_phrase(seed: int) -> str:
    return pick_variant(SUITS_PHRASES, seed)


def light_purpose(family: str) -> str:
    return {
        "panel": "do pracy biurowej i codziennych czynności",
        "oprawa": "do codziennego użytkowania w pomieszczeniu",
        "plafon": "w przejściach, wejściach i pomieszczeniach pomocniczych",
        "high_bay": "na stanowiskach i w strefach roboczych",
        "hermetyczna": "w pomieszczeniach gospodarczych i technicznych",
        "naswietlacz": "na zewnątrz budynku po zmroku",
        "ogrodowa": "w ogrodzie i przy domu po zmroku",
    }.get(family, "w danej przestrzeni")


def installation_holistic(installation: SemanticFact | None, seed: int = 0) -> str:
    """Montaz/wymiary jako CECHA produktu (np. 'Montuje sie natynkowo, a niewielka srednica
    pozwala wpasowac go w ciasne miejsce'), zamiast 'porownaj X z miejscem montazu'."""
    if not installation:
        return ""
    raw = compact_spaces(installation.value)
    mtype = ""
    m = re.search(r"(natynkow\w*|podtynkow\w*|wpuszczan\w*|zwieszan\w*|nabudow\w*|wisz\w*|przykręc\w*)", raw, re.I)
    if m:
        mtype = m.group(1).lower()
    dim = ""
    d = re.search(r"(śred\w*\s*[\d.,]+\s*mm|[\d.,]+(?:\s*[x×]\s*[\d.,]+){1,2}\s*mm|[\d.,]+\s*mm)", raw, re.I)
    if d:
        dim = d.group(0)
    mount_pool = {
        "natynkowy": ["Montuje się natynkowo", "Montaż natynkowy nie wymaga kucia ani wnęki w suficie", "Mocuje się wprost na suficie lub ścianie"],
        "podtynkowy": ["Montuje się podtynkowo", "Montaż podtynkowy chowa oprawę równo z powierzchnią"],
    }.get(mtype)
    if mount_pool:
        mount_clause = pick_variant(mount_pool, seed)
    elif mtype:
        mount_clause = f"Przewidziano montaż {mtype}"
    else:
        mount_clause = ""
    if dim and "śred" in dim.lower():
        val = re.sub(r"śred\w*\s*", "", dim, flags=re.I).strip()
        dim_clause = pick_variant([
            f"niewielka średnica {val} ułatwia montaż nawet w ograniczonej przestrzeni",
            f"średnica {val} ułatwia dobór miejsca montażu",
            f"przy średnicy {val} produkt nie zdominuje przestrzeni",
        ], seed)
    elif dim:
        dim_clause = pick_variant([
            f"wymiary {dim} ułatwiają wpasowanie w wybrane miejsce",
            f"format {dim} łatwo dopasować do dostępnej przestrzeni",
            f"przy wymiarach {dim} bez trudu zaplanujesz montaż",
        ], seed)
    else:
        dim_clause = ""
    if mount_clause and dim_clause:
        return pick_variant([f"{mount_clause}, a {dim_clause}.", f"{mount_clause}; {dim_clause}."], seed)
    if mount_clause:
        return pick_variant([f"{mount_clause}, więc instalacja jest prosta.", f"{mount_clause} bez skomplikowanego osprzętu.", f"{mount_clause} – bez dodatkowych przeróbek."], seed)
    if dim_clause:
        return pick_variant([f"{dim_clause[0].upper()}{dim_clause[1:]}.", "Kompaktowe wymiary ułatwiają wpasowanie w wybrane miejsce.", "Niewielki format ułatwia zaplanowanie miejsca montażu."], seed)
    return ""


def dim_compact(attrs: dict[str, str]) -> str:
    if attrs.get("Wymiary"):
        return compact_spaces(attrs["Wymiary"])
    diameter = first_present(attrs.get("Średnica", ""), attrs.get("Srednica", ""))
    if diameter:
        value = re.sub(r"\s*mm\b", "", compact_spaces(diameter), flags=re.IGNORECASE)
        return f"średnica {value} mm"
    parts = []
    for label in ["Długość", "Szerokość", "Wysokość"]:
        value = attrs.get(label, "")
        if value:
            parts.append(re.sub(r"\s*mm\b", "", compact_spaces(value), flags=re.IGNORECASE))
    return f"{' x '.join(parts)} mm" if parts else ""


_HARVESTED_BANK_CACHE: "dict[str, list[str]] | None" = None
HARVESTED_BANK_PATH = Path(__file__).resolve().parents[1] / "dictionaries" / "seo_phrase_bank_harvested.yaml"


def load_harvested_bank() -> "dict[str, list[str]]":
    """Bank naturalnych fraz z eksportu Woo. Do PROZY (flavour), nie do faktow. Filtrujemy
    frazy z liczbami - one moglyby przeczyc atrybutom produktu (np. inna barwa/IP)."""
    global _HARVESTED_BANK_CACHE
    if _HARVESTED_BANK_CACHE is not None:
        return _HARVESTED_BANK_CACHE
    bank: dict[str, list[str]] = {}
    try:
        import yaml
        data = yaml.safe_load(HARVESTED_BANK_PATH.read_text(encoding="utf-8")) or {}
        for cat, items in (data.get("phrases", {}) or {}).items():
            bank[cat] = [compact_spaces(s) for s in (items or []) if s and is_safe_flavor(s)]
    except Exception:
        bank = {}
    _HARVESTED_BANK_CACHE = bank
    return bank


# Frazy z banku to FLAVOUR - nie moga niesc specyfiki, ktora moze przeczyc produktowi
# (kolor, obca nazwa produktu, material, niszowa funkcja, liczby). Zostaja tylko abstrakcyjne.
_UNSAFE_FLAVOR = re.compile(
    r"(\d"
    r"|biał|czarn|szar|grafit|srebrn|złot|miedzian|brązow|beżow|lawendow|niebiesk|granatow|zielon|czerwon|żółt|popielat|chrom|nikl|nikiel|antracyt|kremow|różow|fioletow"
    r"|rozdzielnic|gniazd|przełączni|przełącz|łączni|\btaśm|szynoprzewód|szyno|reflektor|żyrandol|kinkiet|abażur|\bklosz|przedłużacz|listw|puszk|peszel|opask|\bwtyk|bateri|akumulator|fotowolt|panel\b|panele\b|paneli\b|lampk|świetlówk|żarówk|halogen"
    r"|strefa czasow|zegar|programator|\bpilot|aplikacj|wi-?fi|bluetooth|zdaln|\bsmart|tuya|zigbee|zmierzch|fotokomórk|timer|sterown"
    r"|\bABS\b|\bPVC\b|\bPCV\b|stali\b|\bstal\b|alumini|poliwęglan|\bszkł|\bszkl|drewn|mosiądz|żeliw|\bPC\b"
    r"|czujnik|detektor|wykryw|sieci komputer|teleinform|telefoni|telekom|\bISDN|światłowod|sygnaliz|nadajnik|odbiornik|kontroler|moduł\b|\bprzew[oó]d|\bkabl"
    r"|tworzyw|duroplast|poliamid|plastik|\bmetal|miedzi|\bcynk|ceramik|kamien|beton|gumow"
    r"|kolor|odcień|odcieni|barwie|\bseri|kolekcj|rodzin[yie]|\bmodel|\bwersj|\blinii|\bdąb|\bdębow|orzech|wenge|sosn|mahon"
    r"|czas działania|zakres pracy|opóźnieni|interwał|harmonogram|nastaw|programow|tryb pracy|tryby\b|temperatur otoczen|zakres temperatur"
    r"|podtynkow|natynkow|wpuszczan|zwieszan|nabudow|\bzacisk|pierścień|śrubow|gwintow|urządzeń elektr|montaż jest|montażu jest|\bnóżk|kołnierz)",
    re.I,
)
# Zdania z zaimkiem odnoszacym sie do wczesniejszego kontekstu nie moga stac samodzielnie.
_CONTEXT_DEPENDENT = re.compile(r"^(dzięki (temu|niemu|niej|nim|nią|nimi)|temu |dlatego |z tego (powodu|względu)|oprócz tego|ponadto|poza tym|w ten sposób|to (rozwiązanie|sprawia|oznacza|pozwala)|jego |jej |ich )", re.I)
# Urwane zdania konczace sie skrotem ("...np.", "...m.in.") sa niekompletne.
_DANGLING = re.compile(r"\b(np|m\.in|tzn|tj|itd|wg|ok|tzw)\.$", re.I)
# Slowa pisane KAPITALIKAMI to zwykle marki/serie/skroty (DECO, ISDN) - case-sensitive, bez LED/UV/IP.
_CAPS_TOKEN = re.compile(r"\b(?!LED\b|UV\b|IP\d|RGB\b|USB\b)[A-ZĄĆĘŁŃÓŚŹŻ]{3,}\b")
# Wyraz z Wielkiej Litery W SRODKU zdania = niemal zawsze marka/nazwa wlasna (Elektroplast).
_MIDCAP_TOKEN = re.compile(r"(?<=[a-ząćęłńóśźż] )[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]{2,}")
# Obce nazwy produktow / odwolania do producenta - flavour ma byc o NASZYM produkcie.
_FOREIGN_PRODUCT = re.compile(r"(producent|poczwórn|potrójn|podwójn|jednobiegunow|dwubiegunow|ramka\b|gniazdo\b|włącznik|wyłącznik|przedłużacz)", re.I)


def is_safe_flavor(s: str) -> bool:
    return (not _UNSAFE_FLAVOR.search(s) and not _CONTEXT_DEPENDENT.match(s) and not _DANGLING.search(s)
            and not _CAPS_TOKEN.search(s) and not _MIDCAP_TOKEN.search(s) and not _FOREIGN_PRODUCT.search(s))


# Tylko najbezpieczniejsze kategorie (design/jakosc/zastosowanie) - "funkcje_korzysci" niesie
# najwiecej obcego kontekstu (timery, czujniki, sieci), dlatego do flavour go nie bierzemy.
FAMILY_FLAVOR_CATS = {
    "panel": ["zastosowanie_wnetrza", "design_estetyka", "jakosc_trwalosc"],
    "plafon": ["zastosowanie_wnetrza", "design_estetyka", "jakosc_trwalosc"],
    "oprawa": ["design_estetyka", "zastosowanie_wnetrza", "jakosc_trwalosc"],
    "ogolny": ["design_estetyka", "jakosc_trwalosc", "zastosowanie_wnetrza"],
    "high_bay": ["jakosc_trwalosc", "zastosowanie_wnetrza", "design_estetyka"],
    "hermetyczna": ["jakosc_trwalosc", "zastosowanie_wnetrza", "design_estetyka"],
    "naswietlacz": ["zastosowanie_zewnetrze", "jakosc_trwalosc", "design_estetyka"],
    "ogrodowa": ["zastosowanie_zewnetrze", "design_estetyka", "jakosc_trwalosc"],
    "ramka": ["design_estetyka", "jakosc_trwalosc"],
    "zasilacz": ["jakosc_trwalosc", "design_estetyka"],
    "akcesorium": ["jakosc_trwalosc", "design_estetyka"],
}
# Frazy o swieceniu nie pasuja do produktow, ktore nie sa zrodlem swiatla (ramka/zasilacz/akcesorium).
_NONLIGHT_BAN = re.compile(r"(świat[łl]|oświetl|barw|jasnoś|doświetl|reflektor|lamp|żarów|świec)", re.I)


def bank_flavor_sentences(family: str, seed: int, n: int = 2) -> list[str]:
    bank = load_harvested_bank()
    cats = FAMILY_FLAVOR_CATS.get(family, ["funkcje_korzysci", "design_estetyka", "jakosc_trwalosc"])
    light = is_light_product(family)
    pool: list[str] = []
    for cat in cats:
        for sentence in bank.get(cat, []):
            if not light and _NONLIGHT_BAN.search(sentence):
                continue
            pool.append(sentence)
    if not pool:
        return []
    start = (seed * 13) % len(pool)
    ordered = pool[start:] + pool[:start]
    out: list[str] = []
    out_words: list[set[str]] = []
    for sentence in ordered:
        if sentence in out:
            continue
        words = {w for w in normalize_header(sentence).split() if len(w) > 3}
        # odrzuc zdania o tym samym poczatku lub duzym pokryciu tresci (bliskie duplikaty)
        first = normalize_header(sentence).split()[:1]
        if out and (first == normalize_header(out[-1]).split()[:1]):
            continue
        if any(words and ow and len(words & ow) / min(len(words), len(ow)) > 0.5 for ow in out_words):
            continue
        out.append(sentence)
        out_words.append(words)
        if len(out) >= n:
            break
    return out


def value_with_unit(value: str, unit: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    return re.sub(rf"\s*{re.escape(unit)}\b", "", value, flags=re.IGNORECASE).strip() + f" {unit}"


@dataclass(frozen=True)
class SemanticFact:
    role: str
    label: str
    value: str
    meaning: str
    salience: float = 0.5
    prose_allowed: bool = True


@dataclass(frozen=True)
class DescriptionPlan:
    angle: str
    situation: str
    h3: str
    buyer_goal: str


def first_fact(facts: list[SemanticFact], *roles: str) -> SemanticFact | None:
    wanted = set(roles)
    return next((fact for fact in facts if fact.role in wanted), None)


def facts_for(facts: list[SemanticFact], *roles: str) -> list[SemanticFact]:
    wanted = set(roles)
    return [fact for fact in facts if fact.role in wanted]


def derive_semantic_facts(attrs: dict[str, str], family: str) -> list[SemanticFact]:
    facts: list[SemanticFact] = []
    light = is_light_product(family)
    power = value_with_unit(attrs.get("Moc [W]", ""), "W")
    flux = value_with_unit(attrs.get("Strumień świetlny [lm]", ""), "lm")
    temp = value_with_unit(attrs.get("Temperatura barwowa [K]", ""), "K")
    # Fakty o SWIETLE tylko dla opraw swiecacych - zasilacz/ramka/akcesorium nie "oswietlaja".
    if light and power and flux:
        facts.append(SemanticFact("light_effect", "światło", f"{power}, {flux}", f"łączy moc {power} ze strumieniem {flux}", 0.9))
    elif light and flux:
        facts.append(SemanticFact("light_effect", "strumień świetlny", flux, f"określa ilość światła dostępnego w oprawie", 0.85))
    elif light and power:
        facts.append(SemanticFact("light_effect", "moc", power, f"pomaga dobrać oprawę do wielkości oświetlanej strefy", 0.65))
    if light and temp:
        if attrs.get("Temperatura barwowa [K]", "") == "4000":
            meaning = "daje neutralne, czytelne światło do codziennego użytkowania"
        elif attrs.get("Temperatura barwowa [K]", "") in {"2700", "3000"}:
            meaning = "daje cieplejszy, spokojniejszy odbiór światła"
        elif "-" in attrs.get("Temperatura barwowa [K]", ""):
            meaning = "pozwala dopasować charakter światła do sytuacji"
        else:
            meaning = "wpływa na odbiór światła w miejscu montażu"
        facts.append(SemanticFact("light_tone", "barwa światła", temp, meaning, 0.8))
    elif light and attrs.get("Barwa światła"):
        facts.append(SemanticFact("light_tone", "barwa światła", attrs["Barwa światła"], f"określa charakter światła w oprawie", 0.65))

    if light and has_feature(attrs, "Czujnik ruchu"):
        facts.append(SemanticFact("control", "czujnik ruchu", attrs.get("Czujnik ruchu", "tak"), "uruchamia światło wtedy, gdy ktoś korzysta z przestrzeni", 0.95))
    if has_feature(attrs, "Współpraca ze ściemniaczem"):
        facts.append(SemanticFact("control", "ściemnianie", attrs.get("Współpraca ze ściemniaczem", "tak"), "umożliwia regulację jasności", 0.75))

    ip = attrs.get("Stopień ochrony [IP]", "")
    if ip:
        facts.append(SemanticFact("environment", "klasa ochrony", ip, ip_meaning(ip), 0.9))
    ik = attrs.get("Stopień odporności [IK]", "")
    if ik:
        facts.append(SemanticFact("environment", "odporność mechaniczna", ik, "zwiększa odporność obudowy na uderzenia", 0.75))

    source = attrs.get("Źródło światła", "")
    socket = attrs.get("Trzonek", "")
    if socket:
        meaning = "pozwala dobrać żarówkę LED pod oczekiwaną barwę, jasność i późniejszą wymianę"
        if normalize_header(source).startswith("niez"):
            meaning = "zostawia wybór żarówki po stronie użytkownika, więc łatwiej dopasować efekt do miejsca"
        facts.append(SemanticFact("lamp_choice", "gwint", socket, meaning, 0.85))
    elif source and not normalize_header(source).startswith("niez"):
        facts.append(SemanticFact("lamp_choice", "źródło światła", source, "upraszcza użytkowanie, bo źródło jest częścią oprawy", 0.7))

    mount = attrs.get("Sposób montażu", "")
    dimensions = dim_compact(attrs)
    if mount or dimensions:
        value = natural_join([mount, dimensions])
        if value:
            facts.append(SemanticFact("installation", "montaż", value, "decyduje o dopasowaniu oprawy do przygotowanego miejsca", 0.75))
    color = attrs.get("Kolor", "")
    if color and family != "zasilacz":  # zasilacz jest ukryty - kolor nie jest argumentem
        facts.append(SemanticFact("visual_fit", "kolor", color, "pomaga dopasować widoczną oprawę do sufitu, ściany lub elewacji", 0.55))
    material = prose_material(attrs)
    if material:
        facts.append(SemanticFact("build", "materiał", material, "wpływa na obudowę i codzienne użytkowanie produktu", 0.55))
    if attrs.get("Seria") and family in {"akcesorium", "ramka", "zasilacz"}:
        facts.append(SemanticFact("compatibility", "seria", attrs["Seria"], "zawęża dobór do właściwej rodziny produktów", 0.85))
    if attrs.get("Pasuje do"):
        facts.append(SemanticFact("compatibility", "pasuje do", attrs["Pasuje do"], "określa zgodność z konkretną oprawą lub systemem", 0.95))

    return sorted(facts, key=lambda fact: fact.salience, reverse=True)


def select_description_plan(family: str, facts: list[SemanticFact], attrs: dict[str, str]) -> DescriptionPlan:
    has_control = bool(first_fact(facts, "control"))
    has_environment = bool(first_fact(facts, "environment"))
    if family == "plafon":
        if has_control:
            return DescriptionPlan("automatic_passage_light", "przy wejściu, w korytarzu albo w pomieszczeniu pomocniczym", "Automatyczne światło w codziennych przejściach", "światło ma włączać się samo i od razu doświetlać używaną strefę")
        return DescriptionPlan("even_passage_light", "w korytarzu, przy wejściu albo w pomieszczeniu pomocniczym", "Równomierne światło do przejść i wejść", "oprawa ma świecić spokojnie, bez dominowania nad pomieszczeniem")
    if family == "panel":
        return DescriptionPlan("flat_work_light", "w biurze, recepcji, sali sprzedaży lub korytarzu", "Równe światło na większej powierzchni", "ważna jest równomierność i przewidywalny montaż w suficie")
    if family == "hermetyczna":
        return DescriptionPlan("technical_resistant_light", "w garażu, warsztacie, magazynie lub na zapleczu", "Oprawa do pomieszczeń roboczych", "światło ma działać pewnie mimo kurzu, wilgoci lub częstej eksploatacji")
    if family == "naswietlacz":
        return DescriptionPlan("outdoor_zone_light", "przy elewacji, podjeździe, bramie albo placu", "Światło skierowane na konkretną strefę", "trzeba oświetlić wybrany fragment terenu, a nie całą przestrzeń przypadkowo")
    if family == "high_bay":
        return DescriptionPlan("high_mount_work_light", "w hali, magazynie lub wysokim pomieszczeniu technicznym", "Światło robocze z dużej wysokości", "oprawa ma równomiernie doświetlić powierzchnię z większej odległości")
    if family == "ogrodowa":
        return DescriptionPlan("garden_safety_mood", "przy ścieżce, wejściu, tarasie lub w ogrodzie", "Światło przy domu i w ogrodzie", "oświetlenie ma poprawić widoczność po zmroku i pasować do otoczenia")
    if family == "oprawa" and is_channel_luminaire(attrs):
        return DescriptionPlan("channel_utility_light", "w korytarzu, przejściu, pomieszczeniu technicznym albo pod osłoniętą elewacją", "Proste światło do przejść i zaplecza", "oprawa ma być trwała, łatwa w serwisowaniu i dopasowana do miejsca")
    if family == "oprawa":
        return DescriptionPlan("visible_point_light", "w pomieszczeniu mieszkalnym, usługowym albo komunikacyjnym", "Dopasowanie oprawy do miejsca montażu", "liczy się efekt światła, wygląd oprawy i zgodność ze źródłem światła")
    if family == "akcesorium":
        return DescriptionPlan("replacement_part", "przy uzupełnieniu lub wymianie elementu w istniejącej oprawie", "Dopasowanie do oprawy i serii", "najważniejsze jest potwierdzenie zgodności z częścią, która ma zostać wymieniona")
    if family == "ramka":
        return DescriptionPlan("frame_mounting", "przy montażu panelu poza sufitem kasetonowym", "Ramka do estetycznego osadzenia panelu", "panel ma być zamontowany równo i wyglądać jak zaplanowany element sufitu")
    if family == "zasilacz":
        return DescriptionPlan("power_compatibility", "przy kompletacji panelu lub oprawy LED", "Zasilanie dobrane do oprawy", "najważniejsza jest zgodność elektryczna i miejsce montażu zasilacza")
    if has_environment:
        return DescriptionPlan("resistant_general_light", "w wybranym pomieszczeniu lub strefie technicznej", "Dobór do warunków pracy", "produkt ma pasować do warunków miejsca montażu")
    return DescriptionPlan("general_fit", "w miejscu dopasowanym do jego funkcji", "Dopasowanie produktu do instalacji", "produkt ma pasować do planowanego miejsca i sposobu użytkowania")


def fact_phrase(fact: SemanticFact) -> str:
    if fact.role == "environment":
        return f"klasa {fact.value} daje {fact.meaning}"
    if fact.role == "control":
        return f"{fact.label} {fact.meaning}"
    if fact.role == "light_tone":
        return f"{fact.value} {fact.meaning}"
    if fact.role == "light_effect":
        return fact.meaning
    if fact.role == "lamp_choice":
        return f"{fact.label} {fact.value} {fact.meaning}"
    if fact.role == "installation":
        return f"{fact.value} {fact.meaning}"
    if fact.role == "visual_fit":
        return f"kolor {fact.value.lower()} {fact.meaning}"
    if fact.role == "compatibility":
        return f"{fact.label} {fact.value} {fact.meaning}"
    return f"{fact.label} {fact.value} {fact.meaning}".strip()


def aggregate_fact_sentence(facts: list[SemanticFact], roles: tuple[str, ...], prefix: str = "") -> str:
    selected = facts_for(facts, *roles)
    if not selected:
        return ""
    phrases = [fact_phrase(fact) for fact in selected[:3]]
    sentence = natural_join(phrases)
    return compact_spaces(f"{prefix}{sentence}.")


GOAL_REFRAME_PATTERNS = [
    "{keyword} sprawdzi się tam, gdzie {goal}.",
    "{keyword} warto rozważyć w miejscu, w którym {goal}.",
    "{keyword} dobrze pasuje do sytuacji, w której {goal}.",
    "{keyword} ma sens szczególnie tam, gdzie {goal}.",
    "{keyword} przydaje się wtedy, gdy {goal}.",
    "{keyword} najlepiej oceniać przez pryzmat miejsca, w którym {goal}.",
    "{keyword} pasuje do instalacji, gdzie {goal}.",
]


def goal_reframe_sentence(keyword: str, goal: str, seed_text: str = "") -> str:
    seed = stable_seed(f"{keyword}|{goal}|{seed_text}")
    return GOAL_REFRAME_PATTERNS[seed % len(GOAL_REFRAME_PATTERNS)].format(keyword=keyword, goal=goal)


def pick_variant(pool: list[str], seed: int) -> str:
    return pool[seed % len(pool)] if pool else ""


def light_context_sentence(facts: list[SemanticFact], seed: int = 0, family: str = "") -> str:
    """Holistycznie: JAKIE swiatlo daje produkt (efekt), a nie 'moc/strumien/barwa = wartosci'."""
    effect = first_fact(facts, "light_effect")
    tone = first_fact(facts, "light_tone")
    if tone:
        color = color_temp_effect(tone.value, seed)
        return pick_variant([
            f"Świeci {color}.",
            f"Daje {color}.",
            f"Zapewnia {color}.",
            f"Tworzy {color}.",
            f"Emituje {color}.",
        ], seed)
    if effect:
        return pick_variant([
            f"Zapewnia tyle światła, ile potrzeba {light_purpose(family)}.",
            f"Świeci wystarczająco jasno, by wygodnie oświetlić przestrzeń {light_purpose(family)}.",
            f"Daje jasne, równe światło odpowiednie {light_purpose(family)}.",
            f"Oświetla przestrzeń na tyle dobrze, by wygodnie z niej korzystać {light_purpose(family)}.",
        ], seed)
    return ""


def environment_context_sentence(environment: SemanticFact | None, place: str = "miejsca montażu", seed: int = 0, family: str = "") -> str:
    """Holistycznie: odpornosc produktu. IP != rekomendacja miejsca (family-aware)."""
    if not environment or environment.label != "klasa ochrony":
        return ""
    return pick_variant(environment_usage_sentences(environment.value, family), seed)


def installation_context_sentence(installation: SemanticFact | None, seed: int = 0) -> str:
    if not installation:
        return ""
    return installation_holistic(installation, seed)


def realize_intro_from_plan(keyword: str, attrs: dict[str, str], family: str, facts: list[SemanticFact], plan: DescriptionPlan, seed: int = 0) -> "tuple[str, set[str]]":
    """Zwraca (tekst wstepu, zbior ZUZYTYCH rol faktow) - dzieki temu akapit po H2 moze
    siegnac po INNE fakty i nie powtarzac tych samych zdan."""
    producer = producer_text(attrs, seed)
    control = first_fact(facts, "control")
    tone = first_fact(facts, "light_tone")
    effect = first_fact(facts, "light_effect")
    environment = first_fact(facts, "environment")
    lamp_choice = first_fact(facts, "lamp_choice")

    used: set[str] = set()
    # Zd.1: CZYM JEST (kategoria). Zd.2: DO CZEGO (gdzie + cel). Fraza kluczowa tylko w zd.1.
    definition = f"{keyword}{producer} {family_definition(family, seed)}."
    suits = f"{uppercase_first_letter(suits_phrase(seed))} {plan.situation}, gdzie {plan.buyer_goal}."
    parts = [definition, suits]

    if family == "plafon":
        bits = []
        if control:
            bits.append(control.meaning)
            used.add("control")
        if tone:
            bits.append(tone.meaning)
            used.add("light_tone")
        elif effect:
            bits.append("zapewnia jasne, czytelne światło do codziennego użytkowania")
            used.add("light_effect")
        if bits:
            parts.append(f"{natural_join(bits).capitalize()}.")
        if environment:
            parts.append(environment_context_sentence(environment, seed=seed, family=family))
            used.add("environment")
    elif family in {"panel", "hermetyczna", "naswietlacz", "high_bay", "ogrodowa"}:
        light = light_context_sentence(facts, seed, family)
        env = environment_context_sentence(environment, seed=seed, family=family)
        if light:
            parts.append(light)
            used |= {"light_effect", "light_tone"}
        if env:
            parts.append(env)
            used.add("environment")
    elif family == "akcesorium":
        compat = first_fact(facts, "compatibility")
        visual = first_fact(facts, "visual_fit")
        if compat:
            parts.append(f"Najbezpieczniej trzymać się oznaczenia {compat.value}, bo podobne elementy mogą różnić się mocowaniem albo wymiarem.")
            used.add("compatibility")
        if visual:
            parts.append(f"Wariant w kolorze {color_locative(visual.value)} łatwiej zestawić z obudową oprawy.")
            used.add("visual_fit")
    else:  # oprawa, ogolny, zasilacz, ramka
        if lamp_choice:
            parts.append(f"Popularny gwint {lamp_choice.value} zostawia swobodę wyboru żarówki i efektu światła.")
            used.add("lamp_choice")
        if environment:
            parts.append(environment_context_sentence(environment, seed=seed, family=family))
            used.add("environment")
        inst = first_fact(facts, "installation")
        if inst:
            parts.append(installation_context_sentence(inst, seed))
            used.add("installation")
    return compact_spaces(" ".join(p for p in parts if p)), used


AFTER_H2_OPENERS = [
    "{keyword} wyróżnia się w codziennym użytkowaniu kilkoma konkretnymi cechami.",
    "{keyword} najlepiej pokazuje swoją wartość w konkretnych parametrach użytkowych.",
    "{keyword} odsłania swoje zalety od strony montażu i codziennej eksploatacji.",
    "{keyword} łączy kilka rozwiązań, które widać dopiero w codziennym użytkowaniu.",
    "{keyword} ma kilka cech, które warto poznać jeszcze przed zakupem.",
    "{keyword} zwraca uwagę kilkoma praktycznymi rozwiązaniami.",
    "{keyword} pokazuje swoje zalety przede wszystkim w codziennej pracy.",
    "{keyword} broni się przede wszystkim konkretnymi, użytkowymi cechami.",
    "{keyword} to kilka przemyślanych rozwiązań zamkniętych w jednym produkcie.",
    "{keyword} realnie sprawdza się dopiero w codziennym użytkowaniu.",
]

# Kolejnosc faktow drugorzednych w akapicie po H2 (po fakty, ktorych wstep nie zuzyl).
AFTER_H2_SECONDARY_ROLES = ["control", "lamp_choice", "light_tone", "light_effect", "installation", "visual_fit", "compatibility", "build"]


def render_secondary_fact(fact: SemanticFact, seed: int = 0, family: str = "") -> str:
    """Kazdy fakt jako CECHA/ZACHOWANIE produktu (calosc), nie 'parametr = definicja'."""
    role = fact.role
    if role == "control":
        if "sciem" in normalize_header(fact.label):
            if family == "zasilacz":
                return pick_variant([
                    "Obsługuje ściemnianie, więc pozwala płynnie regulować jasność podłączonej oprawy.",
                    "Współpraca ze ściemniaczem umożliwia sterowanie natężeniem światła zasilanej oprawy.",
                ], seed)
            return pick_variant([
                "Jasność reguluje się ściemniaczem, więc łatwo dopasujesz światło do pory dnia.",
                "Współpracuje ze ściemniaczem – natężenie światła ustawisz pod nastrój i sytuację.",
                "Natężenie światła płynnie przyciemnisz, gdy potrzebujesz spokojniejszego klimatu.",
                "Ściemnianie pozwala zejść z jasnością wieczorem i podkręcić ją do pracy.",
                "Światło ureguluje ściemniacz, więc jedna oprawa obsłuży różne sytuacje.",
            ], seed)
        return pick_variant([
            "Światło włącza się samo, gdy ktoś się pojawi, i gaśnie po chwili – wygodnie tam, gdzie nie chce się szukać włącznika.",
            "Reaguje na ruch: zapala się automatycznie i nie pali się niepotrzebnie, gdy nikogo nie ma.",
            "Zapala się samoczynnie, gdy wykryje obecność, więc nie musisz szukać kontaktu po ciemku.",
            "Czujnik włącza światło tylko wtedy, gdy ktoś jest w pobliżu, co ogranicza zużycie energii.",
            "Oprawa sama reaguje na ruch – świeci, gdy trzeba, i wygasza światło, gdy nikogo nie ma.",
        ], seed)
    if role == "lamp_choice":
        if "gwint" in normalize_header(fact.label) or "trzon" in normalize_header(fact.label):
            return pick_variant([
                f"Źródło światła dobierasz samodzielnie – pasują żarówki {fact.value}, więc łatwo zmienisz barwę i jasność.",
                f"Pod oprawę dobierzesz własną żarówkę {fact.value}, dopasowaną do efektu, jakiego oczekujesz.",
                f"Żarówkę {fact.value} kupujesz osobno, więc barwę i moc światła ustawiasz po swojemu.",
                f"Popularny standard {fact.value} daje swobodę wyboru źródła i łatwą wymianę po latach.",
                f"Sam decydujesz o świetle – wystarczy żarówka {fact.value} o wybranej barwie i mocy.",
            ], seed)
        return pick_variant([
            "Ma wbudowane źródło LED, więc działa od razu i nie wymaga wymiany żarówek.",
            "Źródło LED jest częścią oprawy – nie trzeba kupować ani wymieniać żarówek.",
            "Diody są zintegrowane, więc oprawa świeci od pierwszego uruchomienia, bez dokupowania żarówek.",
            "Zintegrowany LED upraszcza użytkowanie – nie ma żarówek do wkręcania i wymiany.",
            "Wbudowane diody pracują latami i oszczędzają kłopotu z wymianą źródła.",
        ], seed)
    if role == "light_tone":
        color = color_temp_effect(fact.value, seed)
        return pick_variant([f"Daje {color}.", f"Świeci, dając {color}.", f"Tworzy {color}.", f"Zapewnia {color}."], seed)
    if role == "light_effect":
        return pick_variant([
            f"Zapewnia tyle światła, ile potrzeba {light_purpose(family)}.",
            "Świeci wystarczająco jasno do codziennego użytkowania.",
            f"Daje jasne, równe światło odpowiednie {light_purpose(family)}.",
            "Oświetla przestrzeń na tyle dobrze, by wygodnie korzystać z niej po zmroku.",
        ], seed)
    if role == "installation":
        return installation_holistic(fact, seed)
    if role == "visual_fit":
        if family in {"ogrodowa", "naswietlacz"}:
            return pick_variant([
                f"Kolor {fact.value.lower()} dobrze komponuje się z elewacją, tarasem i otoczeniem domu.",
                f"W kolorze {color_locative(fact.value)} oprawa nie rzuca się w oczy na tle elewacji i zieleni.",
                f"Kolor {fact.value.lower()} łatwo dopasować do architektury ogrodu i posesji.",
            ], seed)
        return pick_variant([
            f"Obudowa w kolorze {color_locative(fact.value)} dyskretnie wpisuje się w sufit i ścianę.",
            f"Wykończenie w kolorze {color_locative(fact.value)} łatwo zestawić z aranżacją wnętrza.",
            f"Kolor {fact.value.lower()} sprawia, że oprawa nie rzuca się w oczy po montażu.",
            f"W kolorze {color_locative(fact.value)} oprawa spokojnie komponuje się z otoczeniem.",
            f"Kolor {fact.value.lower()} pozwala wtopić oprawę w wystrój, zamiast go zaburzać.",
        ], seed)
    if role == "build":
        ml = normalize_header(fact.value)
        if "szk" in ml and is_light_product(family):
            return pick_variant([
                "Szklany klosz wpływa na wygląd oprawy i sposób rozproszenia światła.",
                "Szkło nadaje oprawie estetyczny charakter i łagodnie rozprasza światło.",
            ], seed)
        if any(w in ml for w in ("metal", "stal", "alumin", "żeliw", "stalow")):
            return pick_variant([
                f"Obudowa {material_genitive_phrase(fact.value)} jest wytrzymała i dobrze znosi codzienne użytkowanie.",
                f"Metalowa konstrukcja zwiększa trwałość i odporność oprawy.",
            ], seed)
        return pick_variant([
            f"Obudowa {material_genitive_phrase(fact.value)} jest lekka i odporna na codzienne użytkowanie.",
            f"Materiał {fact.value} ułatwia utrzymanie oprawy w czystości i wpływa na jej trwałość.",
        ], seed)
    return f"{fact.label.capitalize()} {fact.value} {fact.meaning}."


def realize_after_h2_from_plan(keyword: str, attrs: dict[str, str], family: str, facts: list[SemanticFact], plan: DescriptionPlan, used_roles: "set[str] | None" = None, seed: int = 0) -> str:
    """Akapit po H2 zaczyna sie fraza kluczowa (wymog SEO), ale NIE powtarza celu zakupowego
    ze wstepu - wnosi fakty, ktorych wstep nie zuzyl (`used_roles`). Klasa ochrony (IP) zostaje
    we wstepie, wiec tu jej nie ma."""
    used = set(used_roles or ())
    opener = AFTER_H2_OPENERS[(seed or stable_seed(f"{keyword}|{family}")) % len(AFTER_H2_OPENERS)].format(keyword=keyword)
    sentences = [opener]
    for role in AFTER_H2_SECONDARY_ROLES:
        if role in used:
            continue
        fact = first_fact(facts, role)
        if not fact:
            continue
        sentences.append(render_secondary_fact(fact, seed, family))
        if len(sentences) >= 3:
            break
    if len(sentences) == 1:
        sentences.append(f"{subject_for(family)} warto dobrać do realnych warunków montażu, kierując się parametrami ze specyfikacji oraz miejscem pracy.")
    return compact_spaces(" ".join(sentences))


def realize_development_from_plan(attrs: dict[str, str], family: str, facts: list[SemanticFact], plan: DescriptionPlan, seed: int = 0) -> str:
    """Akapit rozwijajacy: GROUNDED pule (ogolne prawdy o rodzinie + jakosc) - inspirowane stylem
    banku, ale bez verbatim, wiec zero ryzyka obcego kontekstu. 1 zdanie zastosowania + 1 jakosci."""
    use_pool = FAMILY_USE.get(family, FAMILY_USE["ogolny"])
    use = pick_variant(use_pool, seed)
    # 2. zdanie: KONKRET z danych. Dla polowy produktow (wg seeda) wersja DORADCZA (pkt 8:
    # wniosek z parametru), dla reszty opisowa. Obie sa grounded i niosa liczbe/kotwice.
    if seed % 2 == 0:
        quality = advisory_reasoning_sentence(attrs, family, seed) or concrete_benefit_sentence(attrs, family, seed)
    else:
        quality = concrete_benefit_sentence(attrs, family, seed) or advisory_reasoning_sentence(attrs, family, seed)
    quality = quality or dev_quality_sentence(seed)
    return compact_spaces(f"{use} {quality}")


def build_decision_highlights(plan: DescriptionPlan, facts: list[SemanticFact], attrs: dict[str, str], family: str) -> list[str]:
    items: list[str] = []
    control = first_fact(facts, "control")
    environment = first_fact(facts, "environment")
    tone = first_fact(facts, "light_tone")
    effect = first_fact(facts, "light_effect")
    lamp_choice = first_fact(facts, "lamp_choice")
    installation = first_fact(facts, "installation")
    visual = first_fact(facts, "visual_fit")
    compat = first_fact(facts, "compatibility")

    if family == "plafon":
        if control:
            items.append("Automatyczne światło w przejściu - oprawa włącza się wtedy, gdy ktoś korzysta z korytarza, wejścia lub zaplecza.")
        else:
            items.append("Spokojne światło na co dzień - plafon pasuje do przejść, wejść i pomieszczeń pomocniczych.")
        if tone:
            items.append(f"Czytelny odbiór przestrzeni - {tone.value} {tone.meaning}.")
    elif family == "akcesorium":
        items.append("Wymiana bez całej oprawy - element pozwala uzupełnić lub odświeżyć istniejący zestaw.")
        if compat:
            items.append(f"Dobór do właściwej serii - {compat.value} ogranicza ryzyko pomyłki przy podobnych częściach.")
    elif family in {"panel", "hermetyczna", "naswietlacz", "high_bay", "ogrodowa", "oprawa", "ogolny"}:
        if family == "panel":
            items.append("Równe światło do pracy - panel sprawdzi się tam, gdzie liczy się spokojne oświetlenie większej powierzchni.")
        elif family == "hermetyczna":
            items.append("Do zaplecza i pomieszczeń roboczych - oprawa jest dobierana do miejsc, w których światło ma działać mimo kurzu, wilgoci lub częstej eksploatacji.")
        elif family == "naswietlacz":
            items.append("Doświetlenie konkretnej strefy - naświetlacz pomaga objąć światłem podjazd, elewację, bramę albo plac.")
        elif family == "high_bay":
            items.append("Praca z dużej wysokości - oprawa jest dobierana do hal, magazynów i wysokich pomieszczeń technicznych.")
        elif family == "ogrodowa":
            items.append("Światło wokół domu - lampa pomaga poprawić widoczność przy ścieżce, wejściu, tarasie lub w ogrodzie.")
        elif family == "oprawa" and is_channel_luminaire(attrs):
            items.append("Do przejść i zaplecza - prosta oprawa sprawdza się tam, gdzie światło ma działać pewnie i bez częstej obsługi.")
        else:
            items.append(f"Dopasowanie do miejsca montażu - wariant warto zestawić z tym, jak i gdzie ma być używany.")

    if effect:
        items.append(f"Jasność dobrana do potrzeb - parametry {effect.value} pomagają dobrać oprawę do wielkości oświetlanej strefy.")
    if lamp_choice:
        if lamp_choice.label == "gwint":
            items.append(f"Wybór źródła światła - gwint {lamp_choice.value} pozwala dopasować żarówkę do oczekiwanego efektu.")
        else:
            items.append("Zintegrowane źródło LED - oprawa nie wymaga osobnego doboru żarówki do pierwszego uruchomienia.")
    if environment:
        short = ip_short(environment.value) if environment.label == "klasa ochrony" else environment.meaning
        items.append(f"Odporność dobrana do warunków - klasa {environment.value}, {short}.")
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), attrs.get("Seria", "")))
    if installation:
        mount = attrs.get("Sposób montażu", "")
        dim = dim_compact(attrs)
        sufit_ok = family in {"panel", "plafon", "oprawa", "ogolny"}  # tylko te realnie na suficie
        if dim.startswith("średnica"):
            n = first_present(*re.findall(r"\d+", dim)) or ""
            items.append(pick_variant([
                f"Średnica {n} mm - z góry wiadomo, ile miejsca zajmie produkt po montażu.",
                f"Okrągły kształt o średnicy {n} mm - czytelnie się prezentuje i równomiernie rozkłada światło." if is_light_product(family) else f"Średnica {n} mm - kompaktowy rozmiar ułatwia dopasowanie do miejsca.",
                f"Średnica {n} mm - łatwo ocenić, czy rozmiar pasuje do wybranego miejsca.",
            ], s))
        elif dim:
            fmt = "pasuje do typowych sufitów i ułatwia rozplanowanie opraw" if sufit_ok else "ułatwia zaplanowanie montażu w wybranym miejscu"
            items.append(pick_variant([
                f"Wymiary {dim} - od razu wiadomo, ile miejsca zajmie produkt po montażu.",
                f"Format {dim} - {fmt}.",
                f"Wymiary {dim} - pomagają sprawdzić, czy produkt zmieści się w zaplanowanym miejscu.",
            ], s))
        elif mount:
            items.append(pick_variant([
                f"Montaż {mount.lower()} - prosta instalacja bez dodatkowych przeróbek.",
                f"Montaż {mount.lower()} - oprawę zamontujesz szybko, bez specjalistycznego osprzętu.",
            ], s))
    if visual and family not in {"akcesorium"}:
        outdoor = family in {"ogrodowa", "naswietlacz"}
        vlabel = pick_variant(["Spójny wygląd", "Dyskretne wykończenie", "Estetyka po montażu", "Dopasowanie do otoczenia" if outdoor else "Dopasowanie do wnętrza"], s)
        items.append(f"{vlabel} - kolor {visual.value.lower()} {color_fit_benefit(family, s)}.")
    if len(items) < 3 and attrs.get("Gwarancja"):
        items.append(f"Zakup z informacją o ochronie - producent obejmuje produkt gwarancją {attrs['Gwarancja']}.")
    if len(items) < 3:
        items.append(f"Praktyczny wybór do instalacji - najważniejsze dane warto porównać z miejscem montażu przed zakupem.")
    return items[:6]


# Tylko formy, w ktorych fraza pozostaje w mianowniku (bez odmiany) - patrz zasada uzytkownika.
H2_QUESTION_BUILDERS = [
    lambda kw: f"Czym wyróżnia się {kw}?",
    lambda kw: f"Jakie cechy posiada {kw}?",
    lambda kw: f"Co oferuje {kw}?",
    lambda kw: f"Jakie korzyści daje {kw}?",
    lambda kw: f"Co warto wiedzieć, gdy interesuje Cię {kw}?",
    lambda kw: f"Czym charakteryzuje się {kw}?",
    lambda kw: f"Co potrafi {kw}?",
    lambda kw: f"Jak w praktyce sprawdza się {kw}?",
    lambda kw: f"Na czym polega zaleta, jaką daje {kw}?",
    lambda kw: f"Jak wypada {kw} w codziennym użytkowaniu?",
]

H3_DEV_BY_FAMILY = {
    "ogrodowa": [
        "Oświetlenie ogrodu bez prowadzenia przewodów",
        "Światło przy wejściu, ścieżce i tarasie",
        "Zewnętrzne światło do codziennego dojścia",
    ],
    "panel": [
        "Równe światło do biura i ciągów komunikacyjnych",
        "Panel do sufitu modułowego i modernizacji oświetlenia",
        "Światło robocze bez mocno punktowego efektu",
    ],
    "ramka": [
        "Montaż panelu poza sufitem kasetonowym",
        "Ramka do estetycznego osadzenia panelu",
        "Dopasowanie formatu ramki do oprawy",
    ],
    "zasilacz": [
        "Zasilanie i kompletacja paneli",
        "Dobór zasilacza do oprawy",
        "Zgodność elektryczna przed montażem",
    ],
    "hermetyczna": [
        "Oświetlenie techniczne do garażu, warsztatu i zaplecza",
        "Oprawa do pomieszczeń roboczych i korytarzy technicznych",
        "Światło do miejsc narażonych na kurz i zabrudzenia",
    ],
    "high_bay": [
        "Mocne światło do hal i magazynów",
        "Oprawa do wysokich pomieszczeń technicznych",
        "Światło robocze na dużej wysokości",
    ],
    "plafon": [
        "Światło do korytarzy, wejść i pomieszczeń pomocniczych",
        "Plafon do przestrzeni przejściowych",
        "Równomierne światło w codziennym użytkowaniu",
    ],
    "naswietlacz": [
        "Doświetlenie elewacji, podjazdu i placu",
        "Światło na zewnątrz budynku",
        "Widoczność przy bramie, wejściu i strefie dostaw",
    ],
    "oprawa": [
        "Dopasowanie oprawy do wnętrza i montażu",
        "Światło w pomieszczeniu mieszkalnym lub usługowym",
        "Praktyczny punkt światła w aranżacji",
    ],
    "akcesorium": [
        "Dopasowanie do oprawy i serii",
        "Kompletacja i serwis instalacji oświetleniowej",
        "Element pomocniczy do konkretnej oprawy",
    ],
    "ogolny": [
        "Parametry istotne przy montażu",
        "Dopasowanie produktu do instalacji",
        "Co sprawdzić przed zakupem",
    ],
}


def after_place(family: str) -> str:
    """BIERNIK - do 'doswietlic/oswietlac X'."""
    return {
        "high_bay": "halę, magazyn lub strefę techniczną",
        "panel": "biuro, recepcję lub salę sprzedaży",
        "hermetyczna": "garaż, warsztat lub zaplecze",
        "plafon": "korytarz, wejście lub pomieszczenie pomocnicze",
        "naswietlacz": "podjazd, elewację lub plac",
        "oprawa": "salon, przedpokój lub pomieszczenie usługowe",
        "ogrodowa": "ścieżkę, taras lub wejście do domu",
    }.get(family, "pomieszczenie, korytarz lub strefę techniczną")


def places_nominative(family: str) -> str:
    """MIANOWNIK - po 'takich miejscach/stref jak ...' i jako podmiot."""
    return {
        "high_bay": "hala, magazyn lub strefa techniczna",
        "panel": "biuro, recepcja lub sala sprzedaży",
        "hermetyczna": "garaż, warsztat lub zaplecze",
        "plafon": "korytarz, wejście lub pomieszczenie pomocnicze",
        "naswietlacz": "podjazd, elewacja lub plac",
        "oprawa": "salon, przedpokój lub pomieszczenie usługowe",
        "ogrodowa": "ścieżka, taras lub wejście do domu",
    }.get(family, "pomieszczenie, korytarz lub strefa techniczna")


def secondary_fact_sentences(attrs: dict[str, str], family: str) -> list[str]:
    """Zdania faktograficzne do akapitu po H2 - INNE niz we wstepie (bez moc/strumien/barwa/IP,
    ktore pokrywa wstep), tylko z potwierdzonych atrybutow, bez frazy kluczowej."""
    facts: list[str] = []
    efficacy = attrs.get("Skuteczność świetlna [lm/W]", "") or attrs.get("Skuteczność świetlna", "")
    if efficacy:
        facts.append(f"Skuteczność świetlna {efficacy} lm/W oznacza dobry stosunek światła do pobieranej energii.")
    angle = attrs.get("Kąt świecenia [°]", "") or attrs.get("Kąt świecenia", "")
    if angle:
        facts.append(f"Kąt świecenia {angle}° pomaga skierować światło tam, gdzie jest potrzebne.")
    cri = attrs.get("Wskaźnik oddawania barw", "")
    if cri:
        facts.append(f"Wskaźnik oddawania barw Ra {cri} pozwala wiernie ocenić kolory w oświetlonej przestrzeni.")
    ugr = attrs.get("Wskaźnik olśnienia [UGR]", "")
    if ugr:
        facts.append(f"Współczynnik olśnienia UGR {ugr} ogranicza efekt oślepienia i zmęczenie wzroku.")
    ik = attrs.get("Stopień odporności [IK]", "")
    if ik:
        facts.append(f"Odporność mechaniczna {ik} zwiększa wytrzymałość na uderzenia.")
    socket = attrs.get("Trzonek", "")
    if socket:
        facts.append(f"Dzięki popularnemu gwintowi {socket} można dobrać żarówkę LED o takiej barwie i jasności, jaka najlepiej pasuje do sposobu użycia oprawy.")
    source = attrs.get("Źródło światła", "")
    if source and normalize_header(source).startswith("niez"):
        facts.append("Brak zintegrowanego źródła pozwala samodzielnie dobrać żarówkę i dopasować barwę do potrzeb.")
    elif source:
        facts.append("Zintegrowane źródło LED nie wymaga wymiany żarówki i upraszcza eksploatację.")
    power_supply = attrs.get("Zasilanie", "")
    if power_supply and normalize_header(power_supply) not in {"sieciowe"}:
        facts.append(f"Zasilanie {power_supply.lower()} rozszerza możliwości montażu tam, gdzie trudno o stałe podłączenie.")
    sensor = attrs.get("Czujnik ruchu", "")
    if sensor and sensor.lower() not in {"nie"}:
        facts.append("Czujnik ruchu włącza światło automatycznie i ogranicza zużycie energii.")
    durability = attrs.get("Trwałość [h]", "")
    if durability:
        facts.append(f"Deklarowana trwałość {durability} h przekłada się na długą eksploatację bez wymiany.")
    return facts


def nonlight_fact_sentences(attrs: dict[str, str], family: str) -> list[str]:
    facts: list[str] = []
    dimensions = dim_compact(attrs)
    if family == "ramka":
        if dimensions:
            facts.append(f"Format {dimensions} pomaga potwierdzić zgodność z panelem przed montażem.")
        if attrs.get("Kolor"):
            facts.append(f"Kolor {attrs['Kolor'].lower()} warto zestawić z sufitem i widoczną częścią oprawy.")
        if attrs.get("Sposób montażu"):
            facts.append(f"Montaż {attrs['Sposób montażu'].lower()} sprawdza się wtedy, gdy panel nie trafia bezpośrednio do sufitu kasetonowego.")
        elif "natynk" in normalize_header(attrs.get("Typ produktu", "")):
            facts.append("Montaż natynkowy pozwala osadzić panel poza standardowym sufitem modułowym.")
    elif family == "zasilacz":
        if attrs.get("Moc [W]") or attrs.get("Napięcie [V]"):
            parts = natural_join([value_with_unit(attrs.get("Moc [W]", ""), "W"), attrs.get("Napięcie [V]", "")])
            facts.append(f"Parametry {parts} służą do potwierdzenia zgodności z konkretnym panelem lub oprawą.")
        if has_feature(attrs, "Współpraca ze ściemniaczem"):
            facts.append("Obsługa ściemniania pozwala skompletować zasilanie z instalacją regulacji jasności.")
        if attrs.get("Stopień ochrony [IP]"):
            facts.append(f"Stopień ochrony {attrs['Stopień ochrony [IP]']} oznacza {ip_meaning(attrs['Stopień ochrony [IP]'])}.")
    else:
        target = attrs.get("Pasuje do", "")
        if target:
            facts.append(f"Kompatybilność {target.lower()} jest najważniejszą informacją przy doborze elementu.")
        if has_feature(attrs, "Czujnik ruchu") or "czujnik" in normalize_header(attrs.get("Typ produktu", "")):
            facts.append("Czujnik automatyzuje pracę oprawy i ogranicza potrzebę ręcznego sterowania światłem.")
        if attrs.get("Seria"):
            facts.append(f"W praktyce najbezpieczniej trzymać się serii {attrs['Seria']}, bo podobne elementy mogą różnić się mocowaniem albo wymiarem.")
        if dimensions:
            facts.append(f"Rozmiar {dimensions} warto porównać z częścią, która jest już zamontowana w oprawie.")
    if not facts:
        facts.append("Najważniejsze jest potwierdzenie zgodności z konkretną oprawą, serią albo sposobem montażu.")
    return facts


def build_after_h2(keyword: str, attrs: dict[str, str], family: str) -> str:
    facts = secondary_fact_sentences(attrs, family) if is_light_product(family) else nonlight_fact_sentences(attrs, family)
    if is_light_product(family) and not facts:
        light = light_sentence(attrs, after_place(family))
        if light:
            facts = [light.rstrip(". ") + "."]
    if family == "oprawa" and is_channel_luminaire(attrs):
        color = attrs.get("Kolor", "")
        dimensions = dim_compact(attrs)
        color_text = f" Kolor {color.lower()} pozwala spokojnie zestawić ją z jasną ścianą, sufitem albo elementami instalacji technicznej." if color else ""
        dim_text = f" Wymiary {dimensions} ułatwiają ocenę, czy model zmieści się w wąskim przejściu, wnęce lub nad drzwiami." if dimensions else ""
        return compact_spaces(
            f"{keyword} będzie dobrym wyborem tam, gdzie oprawa ma być widoczna, ale nie powinna dominować nad otoczeniem.{color_text}{dim_text} Taki charakter sprawdza się szczególnie w ciągach komunikacyjnych, na zapleczu i przy wejściach."
        )
    opener = {
        "panel": "daje równe światło na większej powierzchni i dobrze odnajduje się w suficie modułowym.",
        "plafon": "sprawdza się w codziennych przejściach, gdzie potrzebne jest proste i równomierne oświetlenie.",
        "hermetyczna": "jest przeznaczona do pomieszczeń roboczych, w których oprawa ma znosić kurz, zabrudzenia i częstą eksploatację.",
        "high_bay": "pracuje z dużej wysokości, dlatego liczy się w nim nie tylko moc, ale też sposób rozprowadzenia światła.",
        "naswietlacz": "pozwala objąć światłem konkretną strefę na zewnątrz budynku, na przykład podjazd, elewację albo bramę.",
        "ogrodowa": "łączy funkcję użytkową z dekoracyjną i pomaga prowadzić światło przy ścieżce, wejściu albo tarasie.",
        "oprawa": "tworzy punkt światła w miejscu, w którym ważne są zarówno kierunek świecenia, jak i wygląd oprawy.",
        "ramka": "ułatwia estetyczny montaż panelu LED tam, gdzie nie ma typowego sufitu kasetonowego.",
        "zasilacz": "odpowiada za poprawne zasilanie oprawy, dlatego najważniejsza jest zgodność mocy, napięcia i sposobu sterowania.",
        "akcesorium": "to część dobierana do konkretnej oprawy, dlatego przed zakupem najlepiej porównać serię, kolor i sposób mocowania z elementem, który ma zostać wymieniony lub uzupełniony.",
        "ogolny": "sprawdza się jako praktyczny punkt światła w instalacji, w której liczy się prosty montaż i możliwość doboru źródła światła.",
    }.get(family, "sprawdza się w instalacji, w której potrzebne jest proste i przewidywalne światło.")
    body = " ".join(facts[:3])
    return compact_spaces(f"{keyword} {opener} {body}".strip())


def guarantee_note(attrs: dict[str, str]) -> str:
    return f" Producent obejmuje produkt gwarancją {attrs['Gwarancja']}." if attrs.get("Gwarancja") else ""


def closing_light_general(keyword: str, attrs: dict[str, str], family: str) -> str:
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), keyword))
    return pick_variant([
        f"To dobry wybór do instalacji, w której oprawa ma dawać przewidywalne światło i pracować w warunkach zgodnych z klasą ochrony.{guarantee_note(attrs)} Zamów wariant pasujący do planowanego montażu.",
        f"Jeśli zależy Ci na pewnym, równym świetle bez niespodzianek, ten model spełni swoje zadanie.{guarantee_note(attrs)} Dopasuj wariant do miejsca, w którym ma pracować.",
        f"Sprawdzi się wszędzie tam, gdzie oświetlenie ma po prostu działać niezawodnie na co dzień.{guarantee_note(attrs)} Zamów wariant dopasowany do swojej instalacji.",
    ], s)


def closing_light_places(keyword: str, attrs: dict[str, str], family: str) -> str:
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), keyword))
    return pick_variant([
        f"Jeśli oprawa ma oświetlać {after_place(family)}, wybierz wariant o parametrach dopasowanych do tej strefy. Dzięki temu światło będzie użyteczne, a montaż łatwiejszy do zaplanowania.",
        f"Do takich miejsc jak {places_nominative(family)} warto dobrać oprawę pod konkretne potrzeby – wtedy światło naprawdę służy, a nie tylko świeci.",
        f"Przy oświetleniu takich stref jak {places_nominative(family)} liczy się dopasowanie do miejsca; ten model daje na to dobrą podstawę.",
    ], s)


def closing_panel(keyword: str, attrs: dict[str, str], family: str) -> str:
    dimensions = dim_compact(attrs)
    dim_text = f" Format {dimensions} ułatwia porównanie z układem sufitu." if dimensions else ""
    return f"Ten wariant sprawdzi się przy modernizacji oświetlenia biura, recepcji lub korytarza, gdy liczy się równe światło i przewidywalny montaż.{dim_text} Wybierz wariant do sufitu i układu opraw zgodnych ze specyfikacją."


def closing_outdoor(keyword: str, attrs: dict[str, str], family: str) -> str:
    ip = attrs.get("Stopień ochrony [IP]", "")
    ip_text = f" Klasa {ip} dopuszcza {ip_short(ip)}." if ip else ""
    return f"Ten model pasuje do stref zewnętrznych, w których światło ma poprawić widoczność po zmroku.{ip_text} Zamów go, jeśli odpowiada warunkom montażu przy wejściu, ścieżce, tarasie albo elewacji."


def closing_frame(keyword: str, attrs: dict[str, str], family: str) -> str:
    dimensions = dim_compact(attrs) or "format podany w specyfikacji"
    return f"Ramka będzie właściwym wyborem, gdy panel ma być zamontowany poza sufitem kasetonowym albo wymaga estetycznego obramowania. Sprawdź {dimensions} i zamów element pasujący do konkretnej oprawy."


def closing_power_supply(keyword: str, attrs: dict[str, str], family: str) -> str:
    power = value_with_unit(attrs.get("Moc [W]", ""), "W") or "moc ze specyfikacji"
    return f"Zasilacz dobierz do konkretnej oprawy, a nie tylko do nazwy serii. Jeśli {power} oraz napięcie zgadzają się z wymaganiami instalacji, zamów ten element jako uzupełnienie zestawu."


def closing_accessory(keyword: str, attrs: dict[str, str], family: str) -> str:
    target = attrs.get("Pasuje do", "") or attrs.get("Seria", "") or "wskazanej oprawy"
    return f"Akcesorium wybierz wtedy, gdy zgadza się z elementem: {target}. Porównaj kod, serię i wymiary, a następnie zamów część do właściwej oprawy."


def closing_light_quality(keyword: str, attrs: dict[str, str], family: str) -> str:
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), keyword))
    return pick_variant([
        f"Solidne wykonanie i przemyślana konstrukcja sprawiają, że oprawa posłuży przez lata.{guarantee_note(attrs)} Dopasuj wariant do miejsca montażu i sprawdź pełną specyfikację.",
        f"Staranne wykonanie sprawia, że ten model długo zachowa sprawność i wygląd.{guarantee_note(attrs)} Dopasuj wariant do miejsca, w którym ma pracować.",
        f"To oprawa pomyślana z myślą o trwałości i spokojnej, wieloletniej pracy.{guarantee_note(attrs)} Dobierz ją do swojej instalacji i sprawdź dane techniczne.",
    ], s)


def closing_light_value(keyword: str, attrs: dict[str, str], family: str) -> str:
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), keyword))
    return pick_variant([
        "Jeśli szukasz oświetlenia, które połączy funkcjonalność z prostym montażem, ten model będzie trafnym wyborem. Sprawdź parametry i dobierz wariant do swojej przestrzeni.",
        "To rozsądny wybór, gdy zależy Ci na zrównoważeniu jakości światła, trwałości i łatwego montażu. Porównaj dane techniczne i dobierz wariant do siebie.",
        "Dla osób ceniących praktyczne, bezproblemowe oświetlenie ten model będzie sensowną propozycją. Zobacz specyfikację i dopasuj go do miejsca.",
        "Gdy potrzebujesz światła, które po prostu działa i nie sprawia kłopotu, ten model się sprawdzi. Sprawdź dane i dobierz wariant pod swoje potrzeby.",
        "To wybór dla tych, którzy wolą prostotę i pewność działania od zbędnych dodatków. Zobacz parametry i dopasuj oprawę do miejsca montażu.",
    ], s)


def closing_light_fit(keyword: str, attrs: dict[str, str], family: str) -> str:
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), keyword))
    return pick_variant([
        "Dobierz wariant do charakteru pomieszczenia i sposobu montażu, a oprawa odwdzięczy się czytelnym, wygodnym światłem na co dzień.",
        "Dopasowana do miejsca i funkcji oprawa zapewni komfortowe, użyteczne światło na co dzień.",
        "Wybrana pod konkretne pomieszczenie, da światło, które realnie ułatwia korzystanie z przestrzeni.",
    ], s)


CLOSING_CTA = [
    "Sprawdź pełną specyfikację i dobierz wariant do swojej instalacji.",
    "Porównaj dane techniczne i zamów model dopasowany do miejsca montażu.",
    "Zobacz parametry i wybierz wariant pod swoje potrzeby.",
    "Dopasuj wariant do miejsca pracy i zamów go razem z potrzebnym osprzętem.",
]


def concrete_cta(attrs: dict[str, str], family: str, seed: int, avoid: str = "") -> str:
    """CTA z KONKRETEM z produktu (pkt 7): format/moc/sterowanie/IP zamiast ogolnika. Kandydat,
    ktorego wartosc juz padla w zdaniu wiodacym (`avoid`), jest odrzucany - bez powtorzen tej
    samej danej w sasiadujacych zdaniach (pkt 5). Trzonek celowo pomijamy - dobor zarowki
    omawia juz akapit rozwijajacy i lista cech."""
    d = dim_compact(attrs)
    moc = attrs.get("Moc [W]", "")
    ip = attrs.get("Stopień ochrony [IP]", "")
    a = normalize_header(avoid)
    cands: list[tuple[str, str]] = []  # (wartosc do sprawdzenia z 'avoid', tekst CTA)
    if family in {"panel", "ramka"} and d and not d.startswith("średnica"):
        cands.append((d, f"Przed zakupem sprawdź format {d} i upewnij się, że pasuje do planowanego sposobu montażu."))
    if family == "zasilacz" and moc:
        ster = " oraz typ sterowania" if has_feature(attrs, "Współpraca ze ściemniaczem") else ""
        cands.append((moc, f"Przy wyborze porównaj moc {moc} W{ster} z wymaganiami konkretnej oprawy."))
    if ip and family in {"ogrodowa", "naswietlacz"}:
        cands.append((ip, f"Przed montażem na zewnątrz sprawdź, czy klasa {ip} odpowiada warunkom w wybranym miejscu."))
    if d and not d.startswith("średnica"):
        cands.append((d, f"Zanim zamówisz, zestaw wymiary {d} z miejscem, w którym produkt ma pracować."))
    available = [(value, text) for value, text in cands if normalize_header(str(value)) not in a]
    pool = available or cands
    return pick_variant([text for _, text in pool], seed) if pool else pick_variant(CLOSING_CTA, seed)


def closing_data(keyword: str, attrs: dict[str, str], family: str) -> str:
    """Zamkniecie oparte na KONKRECIE z danych + CTA (tez konkretne, pkt 7)."""
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), keyword))
    concrete = concrete_benefit_sentence(attrs, family, s + 5)
    lead = concrete or dev_quality_sentence(s)
    return f"{lead} {concrete_cta(attrs, family, s, avoid=lead)}"


def closing_advice(keyword: str, attrs: dict[str, str], family: str) -> str:
    """Zamkniecie DORADCZE (pkt 8): wniosek z parametru + CTA z konkretem. Fallback na konkret,
    gdy brak danych do doradztwa - dzieki temu zawsze konczymy konkretem, nie ogolnikiem."""
    s = stable_seed(first_present(attrs.get("Kod producenta", ""), keyword))
    advice = advisory_reasoning_sentence(attrs, family, s + 3)
    lead = advice or concrete_benefit_sentence(attrs, family, s + 5) or dev_quality_sentence(s)
    return f"{lead} {concrete_cta(attrs, family, s, avoid=lead)}"


CLOSING_BUILDERS = {  # closing_data (konkret+CTA) wazony 2x, zeby czesciej konczyc konkretem (pkt 1)
    "high_bay": [closing_data, closing_advice, closing_light_places, closing_data],
    "panel": [closing_data, closing_advice, closing_panel, closing_data],
    "hermetyczna": [closing_data, closing_advice, closing_light_places, closing_data],
    "plafon": [closing_data, closing_advice, closing_light_places, closing_data],
    "naswietlacz": [closing_data, closing_advice, closing_outdoor, closing_data],
    "oprawa": [closing_data, closing_advice, closing_light_places, closing_data],
    "ogrodowa": [closing_data, closing_advice, closing_outdoor, closing_data],
    "ramka": [closing_data, closing_frame],
    "zasilacz": [closing_data, closing_power_supply],
    "akcesorium": [closing_data, closing_accessory],
    "ogolny": [closing_data, closing_light_value],
}


def faq_light_subject(family: str) -> "tuple[str, str]":
    """(mianownik, dopelniacz 'do ...') dla naturalnego jezyka FAQ."""
    return {
        "panel": ("Panel", "tego panelu"),
        "plafon": ("Plafon", "tego plafonu"),
        "naswietlacz": ("Naświetlacz", "tego naświetlacza"),
        "ogrodowa": ("Lampa", "tej lampy"),
        "high_bay": ("Oprawa", "tej oprawy"),
        "hermetyczna": ("Oprawa", "tej oprawy"),
        "oprawa": ("Oprawa", "tej oprawy"),
    }.get(family, ("Produkt", "tego produktu"))


def faq_material_comment(material: str, outdoor: bool) -> str:
    """Ludzki komentarz do materialu obudowy - zalezny od typu tworzywa, liczby (l.mnoga dla
    'ABS i PE') i miejsca pracy."""
    m = normalize_header(material)
    if "szk" in m:
        return "Szklany element ładnie rozprasza światło i nie żółknie z czasem, warto tylko zachować ostrożność przy montażu."
    if "metal" in m or "aluminium" in m or "stal" in m:
        return "To solidny, trwały materiał, który dobrze odprowadza ciepło i długo zachowuje wygląd."
    plural = material_is_plural(material)
    if outdoor:
        return (
            "To lekkie tworzywa odporne na wilgoć i zmienne warunki pogodowe, dlatego dobrze sprawdzają się na zewnątrz." if plural
            else "To lekkie tworzywo odporne na wilgoć i zmienne warunki pogodowe, dlatego dobrze sprawdza się na zewnątrz."
        )
    return (
        "To lekkie, odporne tworzywa, które dobrze znoszą codzienne użytkowanie i nie wymagają szczególnej pielęgnacji." if plural
        else "To lekkie, odporne tworzywo, które dobrze znosi codzienne użytkowanie i nie wymaga szczególnej pielęgnacji."
    )


def faq_barwa_comment(temp: str, family: str = "") -> str:
    """Ludzki opis charakteru swiatla wg temperatury barwowej. Neutralna 4000 K bez frazy
    'pasuje niemal wszedzie' (za szeroka) - konkretniej o zastosowaniu."""
    n = _int_or_none(temp)
    if n and n <= 3300:
        return "To ciepłe, przyjemne światło, które buduje przytulny nastrój i wieczorem nie męczy oczu."
    if n and n >= 5000:
        return "To chłodne, wyraźne światło, które ułatwia pracę tam, gdzie liczy się dobra widoczność."
    if family == "plafon":
        return "Taka barwa dobrze sprawdzi się w korytarzu, łazience, pomieszczeniu technicznym lub strefie wejściowej."
    return "To neutralne światło, które dobrze sprawdza się tam, gdzie liczy się dobra widoczność, bez wyraźnie ciepłego ani chłodnego efektu."


def faq_candidates(kw_cap: str, attrs: dict[str, str], family: str, seed: int = 0) -> list[tuple[str, str]]:
    """Kandydaci na pytania FAQ - tylko z potwierdzonych faktow + generyczne uzupelnienia.
    Ton LUDZKI: bezposrednia odpowiedz na starcie, druga osoba, wyjasnienie 'dlaczego'."""
    cands: list[tuple[str, str]] = []
    subj, gen = faq_light_subject(family)
    is_solar = normalize_header(attrs.get("Zasilanie", "")) == "solarne"
    montaz = attrs.get("Sposób montażu", "")
    if montaz:
        przezn = "przeznaczona" if subj in {"Oprawa", "Lampa"} else "przeznaczony"
        cands.append((
            "Jak montuje się ten produkt?",
            f"{subj} jest {przezn} do montażu {mount_genitive(montaz)}. Zanim zaczniesz, sprawdź wymiary oraz miejsce mocowania — dzięki temu element dobrze przylgnie do podłoża i nie będzie kolidował z pozostałymi częściami instalacji.",
        ))
    if not is_light_product(family):
        dimensions = dim_compact(attrs)
        target = attrs.get("Pasuje do", "")
        if family == "ramka":
            cands.append((
                "Do jakich paneli pasuje ta ramka?",
                f"Ramkę dobiera się przede wszystkim do formatu panelu LED. {('W tym modelu format to ' + dimensions + ', dlatego porównaj go z wymiarem oprawy przed zakupem.') if dimensions else 'Przed zakupem porównaj wymiary ramki z wymiarem panelu oraz sposobem jego osadzenia.'}",
            ))
            cands.append((
                "Kiedy potrzebna jest ramka natynkowa?",
                "Ramka natynkowa przydaje się wtedy, gdy panel LED ma być zamontowany poza sufitem kasetonowym albo w miejscu bez przygotowanego modułu pod oprawę. Pozwala estetycznie osadzić panel na widocznej powierzchni sufitu.",
            ))
        elif family == "zasilacz":
            cands.append((
                "Do jakich opraw pasuje ten zasilacz?",
                f"Zasilacz dobiera się do konkretnej oprawy po mocy, napięciu i serii. {('Ten model ma moc ' + value_with_unit(attrs.get('Moc [W]', ''), 'W') + ', więc porównaj ją z wymaganiami panelu.') if attrs.get('Moc [W]') else 'Porównaj parametry z dokumentacją oprawy przed montażem.'}",
            ))
            if has_feature(attrs, "Współpraca ze ściemniaczem"):
                cands.append((
                    "Czy zasilacz obsługuje ściemnianie?",
                    "Tak, atrybut współpracy ze ściemniaczem potwierdza możliwość użycia w instalacji z regulacją jasności. Przed zakupem sprawdź typ sterowania oraz zgodność z oprawą, do której zasilacz ma zostać podłączony.",
                ))
            cands.append((
                "Gdzie zamontować zasilacz?",
                "Zasilacz najlepiej umieścić w miejscu z dostępem serwisowym i zgodnym z jego stopniem ochrony. Nie należy traktować go jak elementu dekoracyjnego; ważniejsze są zgodność elektryczna, chłodzenie i możliwość późniejszej kontroli połączeń.",
            ))
        else:
            if "czujnik" in normalize_header(attrs.get("Typ produktu", "")) or has_feature(attrs, "Czujnik ruchu"):
                cands.append((
                    "Z jaką oprawą współpracuje czujnik?",
                    f"Czujnik należy dobrać do zgodnej oprawy lub serii. {('W danych produktu wskazano kompatybilność: ' + target + '.') if target else 'Przed zakupem sprawdź serię, kod producenta i sposób podłączenia.'}",
                ))
                cands.append((
                    "Jak czujnik automatyzuje pracę oprawy?",
                    "Czujnik reaguje na ruch i pozwala uruchamiać światło bez ręcznego włączania. Taki element jest przydatny w przejściach, strefach technicznych i miejscach, gdzie oprawa ma działać tylko wtedy, gdy ktoś korzysta z przestrzeni.",
                ))
            else:
                cands.append((
                    "Jak sprawdzić kompatybilność akcesorium?",
                    f"Najpierw porównaj serię, kod producenta oraz wymiary z oprawą, do której element ma pasować. {('W danych produktu wskazano: ' + target + '.') if target else 'Podobny wygląd nie zawsze oznacza zamienność części.'}",
                ))
                cands.append((
                    "Kiedy warto wymienić taki element?",
                    "Akcesorium dobiera się wtedy, gdy trzeba uzupełnić zestaw, odtworzyć brakującą część albo dopasować oprawę do konkretnego montażu. Najważniejsze jest potwierdzenie zgodności z modelem oprawy.",
                ))
        if attrs.get("Stopień ochrony [IP]"):
            ip = attrs["Stopień ochrony [IP]"]
            cands.append((
                "Co oznacza stopień ochrony tego elementu?",
                f"Stopień ochrony {ip} oznacza {ip_meaning(ip)}. Porównaj tę informację z miejscem montażu, zwłaszcza jeśli element ma pracować poza suchym wnętrzem albo w pobliżu wilgoci.",
            ))
        if len(cands) < 3:
            cands.append((
                "Co sprawdzić przed zakupem?",
                "Przed zakupem porównaj nazwę serii, kod producenta, wymiary i sposób montażu z oprawą, do której element ma pasować. Przy akcesoriach zgodność jest ważniejsza niż sam wygląd części.",
            ))
        return cands

    outdoor = family in {"ogrodowa", "naswietlacz"}
    if attrs.get("Trzonek"):
        trzonek = attrs["Trzonek"]
        cands.append((
            "Jakie źródło światła pasuje do tego produktu?",
            pick_variant([
                f"Do {gen} pasuje żarówka z trzonkiem {trzonek}. Możesz więc samodzielnie dobrać jej moc, jasność i barwę światła — cieplejszą do nastrojowego oświetlenia albo neutralną, jeśli zależy Ci głównie na dobrej widoczności.",
                f"{subj} współpracuje z żarówkami o trzonku {trzonek}, więc źródło światła dobierasz sam. Wybierz cieplejszą barwę, gdy chcesz przytulny klimat, albo chłodniejszą, kiedy potrzebujesz mocnego, wyraźnego światła do pracy.",
            ], seed),
        ))
    elif attrs.get("Źródło światła"):
        integrated = "niez" not in normalize_header(attrs["Źródło światła"])
        if integrated:
            cands.append((
                "Czy trzeba kupować żarówki do tego produktu?",
                f"Nie. {subj} ma wbudowane źródło LED, więc nie kupujesz ani nie wymieniasz żarówek — świeci od razu po podłączeniu. Dzięki temu przez cały czas masz równe, stabilne światło i mniej rzeczy do pilnowania na co dzień.",
            ))
        else:
            cands.append((
                "Czy źródło światła jest wbudowane?",
                f"Nie, żarówkę dobierasz samodzielnie do {gen}. To wygodne, bo możesz dopasować moc i barwę światła do konkretnego pomieszczenia, a po latach po prostu wymienić samo źródło, bez wymiany całej oprawy.",
            ))
    temp = attrs.get("Temperatura barwowa [K]", "")
    if temp and "-" not in temp and "/" not in temp:
        barwa = f" ({attrs['Barwa światła'].lower()})" if attrs.get("Barwa światła") else ""
        cands.append((
            "Jaką barwę światła daje ten produkt?",
            f"{subj} świeci światłem o temperaturze barwowej {temp} K{barwa}. {faq_barwa_comment(temp, family)} Dla spójnego efektu warto dobrać podobną barwę do pozostałych lamp w pomieszczeniu.",
        ))
    ip = attrs.get("Stopień ochrony [IP]", "")
    if ip and outdoor:
        digits = re.sub(r"\D", "", ip)
        outdoor_ok = digits in {"44", "54", "65", "66", "67"}
        answer = (
            f"Tak. Klasa {ip} pozwala na montaż na zewnątrz, pod warunkiem prawidłowego podłączenia i zastosowania zgodnie z instrukcją. Sprawdzi się przy wejściu, ścieżce, tarasie albo w ogrodzie, jeśli instalacja zostanie wykonana prawidłowo."
            if outdoor_ok
            else f"Ostrożnie — nie zakładaj montażu na zewnątrz tylko po nazwie. {ip_resistance_line(ip)} Te warunki trzeba najpierw zestawić z konkretnym miejscem, w którym produkt ma pracować."
        )
        cands.append(("Czy ten produkt nadaje się na zewnątrz?", answer))
    elif ip:
        cands.append((
            "W jakich warunkach można montować ten produkt?",
            f"{ip_resistance_line(ip)} Zestaw tę informację z miejscem montażu, bo inne wymagania ma suchy pokój, a inne łazienka, garaż czy strefa narażona na zachlapanie.",
        ))
    sensor = attrs.get("Czujnik ruchu", "")
    if sensor and sensor.lower() not in {"nie"}:
        gdzie = "przy furtce, wejściu, garażu, ścieżce lub tarasie" if outdoor else "w przejściach, na klatce czy w pomieszczeniu gospodarczym"
        # Produkt solarny nie ma "rachunkow" - pracuje z akumulatora, wiec mowimy o energii.
        oszczednosc = "pomaga oszczędzać energię zgromadzoną w akumulatorze" if is_solar else "pomaga ograniczyć zużycie energii i rachunki"
        cands.append((
            "Jak działa czujnik ruchu w tym produkcie?",
            f"Wbudowany czujnik sam włącza światło, gdy ktoś się zbliży, i gasi je po chwili bez ruchu. Dzięki temu nie szukasz włącznika po ciemku, a światło nie świeci niepotrzebnie — to wygodne {gdzie} i {oszczednosc}.",
        ))
    if attrs.get("Gwarancja") or attrs.get("Trwałość [h]"):
        if attrs.get("Trwałość [h]"):
            g = f"Producent deklaruje trwałość na poziomie {attrs['Trwałość [h]']} h" + (f", a produkt obejmuje gwarancją {attrs['Gwarancja']}. " if attrs.get("Gwarancja") else ". ")
        else:
            g = f"Produkt objęto gwarancją {attrs['Gwarancja']}. "
        cands.append((
            "Jak długo posłuży ten produkt?",
            f"{g}Przy typowym użytkowaniu może to oznaczać wiele lat pracy bez częstych wymian, o ile oprawa działa w odpowiednich warunkach i przy właściwym zasilaniu. Realna żywotność zależy bowiem od warunków, w jakich produkt pracuje.",
        ))
    if prose_material(attrs):
        made = f"{subj} została wykonana" if subj in {"Oprawa", "Lampa"} else f"{subj} został wykonany"
        cands.append((
            "Z jakiego materiału wykonano produkt?",
            f"{made} {material_genitive_phrase(prose_material(attrs))}. {faq_material_comment(prose_material(attrs), outdoor)} Materiał obudowy warto brać pod uwagę przy wyborze, bo wpływa na trwałość i codzienny komfort użytkowania.",
        ))
    # Generyczne uzupelnienia (zawsze dostepne, oparte na typie/zastosowaniu)
    cands.append((
        "Gdzie najlepiej sprawdzi się ten produkt?",
        f"Najlepiej odnajdzie się w takich miejscach jak {places_nominative(family)}. Przy wyborze kieruj się przede wszystkim strumieniem i barwą światła oraz sposobem montażu, żeby oświetlenie realnie pasowało do funkcji danej przestrzeni.",
    ))
    cands.append((
        "Jak dobrać parametry do miejsca montażu?",
        "Zacznij od wysokości montażu, wielkości strefy i tego, jakie światło chcesz uzyskać. Potem zestaw moc, strumień, barwę oraz wymiary z warunkami w miejscu pracy — wtedy najłatwiej trafić w odpowiedni wariant.",
    ))
    return cands


FAQ_FILLERS = [
    " Pełne dane znajdziesz w specyfikacji technicznej powyżej.",
    " W razie wątpliwości porównaj parametry z wymaganiami swojej instalacji.",
    " Szczegóły techniczne zebraliśmy w tabeli specyfikacji.",
    " Przy wyborze warto kierować się przeznaczeniem i miejscem montażu.",
    " Jeśli masz wątpliwości, sprawdź zgodność z resztą instalacji.",
]


def fit_faq_answer(text: str, used_fillers: "set[str] | None" = None, low: int = 200, high: int = 300) -> str:
    text = compact_spaces(text)
    used = used_fillers if used_fillers is not None else set()
    if len(text) > high:
        cut = text[:high]
        boundary = cut.rfind(". ")
        text = (cut[: boundary + 1] if boundary > low else cut).strip()
    for filler in FAQ_FILLERS:
        if len(text) >= low:
            break
        if filler in used:  # kazdy dopelniacz max 1x na cale FAQ - zeby nie brzmialo mechanicznie
            continue
        if len(text) + len(filler) <= high:
            text += filler
            used.add(filler)
    return text


def build_faq(kw_cap: str, attrs: dict[str, str], family: str, seed: int) -> str:
    candidates = faq_candidates(kw_cap, attrs, family, seed)
    start = seed % len(candidates)
    ordered = candidates[start:] + candidates[:start]
    chosen: list[tuple[str, str]] = []
    seen_questions: set[str] = set()
    used_fillers: set[str] = set()
    for question, answer in ordered:
        if question in seen_questions:
            continue
        seen_questions.add(question)
        chosen.append((question, fit_faq_answer(answer, used_fillers)))
        if len(chosen) == 3:
            break
    items = [
        f"<p><strong>{index}. {escape(question)}</strong><br>\n{escape(answer)}</p>"
        for index, (question, answer) in enumerate(chosen, start=1)
    ]
    return "\n\n".join(items)


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
    # Samodzielny czujnik to akcesorium; "z czujnikiem ruchu" w nazwie oprawy - nie.
    if text.startswith("czujnik"):
        return "akcesorium"
    if any(token in text for token in ["klosz", "zapinka", "wspornik", "siatka", "konektor", "pilot", "uchwyt", "soczewka", "klips", "linka", "lacznik", "łącznik"]):
        return "akcesorium"
    # Oswietlenie zewnetrzne/ogrodowe - lampy ogrodowe, kule, slupki, kinkiety i oprawy elewacyjne.
    if any(token in text for token in ["ogrodow", "elewacyjn", "slupek", "słupek", "kinkiet", "kula"]):
        return "ogrodowa"
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
    if (
        "oprawa sufitowa" in text
        or "oprawa punktowa" in text
        or "oprawa kanalowa" in text
        or "oprawa kanałowa" in title.lower()
        or "lampa wiszaca" in text
        or "lampa wisząca" in title.lower()
    ):
        return "oprawa"
    return "ogolny"


def stable_seed(value: str) -> int:
    digest = hashlib.sha1(value.encode("utf-8")).hexdigest()
    return int(digest[:8], 16)


def variant(items: list[Any], seed: int) -> Any:
    return items[seed % len(items)]


def choose_benefits(keyword: str, attrs: dict[str, str], family: str, seed: int) -> list[str]:
    # Buildery zwracaja konkretne punkty cecha->korzysc w kolejnosci waznosci - nie rotujemy
    # seedem (to wypychalo najmocniejsze cechy poza liste), bierzemy pierwsze 6.
    return BENEFIT_BUILDERS[family](keyword, attrs)[:6]


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
        parts.append(f"Moc {value_with_unit(power, 'W')} i strumień {value_with_unit(flux, 'lm')} zapewniają światło odpowiednie na {place}")
    elif power:
        parts.append(f"Moc {value_with_unit(power, 'W')} sprawdzi się tam, gdzie potrzebny jest wyraźny punkt światła")
    elif flux:
        parts.append(f"Strumień {value_with_unit(flux, 'lm')} daje jasność przydatną w codziennym użytkowaniu pomieszczenia")
    if temp:
        temp_text = value_with_unit(temp, "K")
        parts.append(f"Barwa {temp_text} daje neutralne, czytelne światło" if temp == "4000" else f"Barwa {temp_text} pozwala dopasować odbiór światła do zastosowania")
    return ". ".join(parts) + ("." if parts else "")


def ip_sentence(attrs: dict[str, str], context: str) -> str:
    ip = attrs.get("Stopień ochrony [IP]", "")
    return f"Stopień ochrony {ip} oznacza {ip_meaning(ip)} i ma znaczenie {context}." if ip else ""


def channel_intro_text(keyword: str, attrs: dict[str, str]) -> str:
    producer = producer_text(attrs)
    socket = attrs.get("Trzonek", "")
    ip = attrs.get("Stopień ochrony [IP]", "")
    socket_text = (
        f" Dzięki popularnemu gwintowi {socket} można dobrać do niej żarówkę LED o odpowiedniej mocy, barwie i jasności."
        if socket
        else ""
    )
    ip_text = (
        f" Stopień ochrony {ip} zapewnia {ip_meaning(ip)}, dlatego oprawa nadaje się do miejsc, w których liczy się proste, trwałe i łatwe w serwisowaniu oświetlenie."
        if ip
        else " To dobry wybór tam, gdzie oświetlenie ma być proste, trwałe i łatwe w późniejszym serwisowaniu."
    )
    return f"{keyword}{producer} dobrze sprawdzi się w korytarzach, przejściach, pomieszczeniach technicznych oraz pod osłoniętą elewacją.{socket_text}{ip_text}"


PRODUCER_FORMS = [
    " marki {p}", " od {p}", " produkcji {p}", " z oferty {p}", " z portfolio {p}",
    " w ofercie {p}", " w wykonaniu {p}", " z katalogu {p}", " marki {p}", " od {p}",
]


def producer_text(attrs: dict[str, str], seed: int = 0) -> str:
    producer = attrs.get("Producent", "")
    return pick_variant(PRODUCER_FORMS, seed).format(p=producer) if producer else ""


def intro_high_bay_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to oprawa przemysłowa LED{producer_text(attrs)} do hal, magazynów i wysokich pomieszczeń technicznych. {light_sentence(attrs, 'stanowiska pracy, regały i ciągi komunikacyjne')} {ip_sentence(attrs, 'w przestrzeniach, gdzie oprawa pracuje w pyle, wilgoci lub przy intensywnej eksploatacji')}"


def intro_high_bay_b(keyword: str, attrs: dict[str, str]) -> str:
    angle = attrs.get("Kąt świecenia", "")
    angle_text = f" Kąt świecenia {angle} pomaga prowadzić światło z większej wysokości." if angle else ""
    return f"{keyword} sprawdzi się tam, gdzie zwykła oprawa sufitowa nie daje wystarczającego zasięgu. {light_sentence(attrs, 'dużą powierzchnię roboczą')}{angle_text} {ip_sentence(attrs, 'przy montażu w wymagającym obiekcie przemysłowym')}"


def intro_high_bay_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest przeznaczona do mocnego, roboczego oświetlenia z dużej wysokości. {light_sentence(attrs, 'hale produkcyjne, magazyny i strefy techniczne')} W takim zastosowaniu liczy się stabilny strumień światła, odporna obudowa i czytelność przestrzeni przez wiele godzin pracy."


def intro_panel_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to panel LED{producer_text(attrs)} do biur, recepcji, sal sprzedaży i innych wnętrz wymagających równego światła. {light_sentence(attrs, 'stanowiska pracy i ciągi komunikacyjne')} Format {dim_compact(attrs) or 'panelu'} warto dopasować do układu sufitu."


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
    material_text = f" Obudowa {material_genitive_phrase(material)} ułatwia codzienne użytkowanie w technicznych warunkach." if material else ""
    return f"{keyword} sprawdzi się tam, gdzie oświetlenie ma być proste, odporne i przewidywalne. {light_sentence(attrs, 'garaż, piwnicę, warsztat lub zaplecze')}{material_text}"


def intro_hermetic_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest przeznaczona do pomieszczeń, w których oprawa pracuje bliżej kurzu, wilgoci i intensywnej eksploatacji niż typowe oświetlenie dekoracyjne. {light_sentence(attrs, 'stanowiska robocze i przejścia techniczne')} {ip_sentence(attrs, 'przy doborze oprawy do takich warunków')}"


def intro_plafon_a(keyword: str, attrs: dict[str, str]) -> str:
    sensor = " z czujnikiem ruchu" if has_feature(attrs, "Czujnik ruchu") else ""
    return f"{keyword} to plafon{sensor}{producer_text(attrs)} do korytarzy, wejść, klatek schodowych i pomieszczeń pomocniczych. {light_sentence(attrs, 'przejścia oraz niewielkie strefy użytkowe')} {ip_sentence(attrs, 'przy montażu w miejscu bardziej narażonym na wilgoć lub zabrudzenia')}"


def intro_plafon_b(keyword: str, attrs: dict[str, str]) -> str:
    sensor = " Czujnik ruchu ogranicza potrzebę ręcznego włączania światła." if has_feature(attrs, "Czujnik ruchu") else ""
    return f"{keyword} sprawdzi się w miejscach, gdzie oświetlenie ma działać funkcjonalnie i bez zbędnej obsługi.{sensor} {light_sentence(attrs, 'wejście, korytarz lub pomieszczenie gospodarcze')}"


def intro_plafon_c(keyword: str, attrs: dict[str, str]) -> str:
    dimensions = dim_compact(attrs)
    dim_text = f" Wymiary {dimensions} pozwalają ocenić, czy oprawa nie będzie zbyt masywna." if dimensions else ""
    return f"{keyword} jest praktyczną oprawą do przestrzeni przejściowych, w których światło powinno włączać się szybko i równomiernie obejmować najważniejszą strefę. {light_sentence(attrs, 'codzienne przejścia i wejścia')}{dim_text}"


def intro_naswietlacz_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to naświetlacz LED{producer_text(attrs)} do elewacji, podjazdów, bram, placów i stref dostaw. {light_sentence(attrs, 'wybraną strefę na zewnątrz budynku')} {ip_sentence(attrs, 'przy pracy w deszczu, pyle i zmiennych warunkach pogodowych')}"


def intro_naswietlacz_b(keyword: str, attrs: dict[str, str]) -> str:
    angle = attrs.get("Kąt świecenia", "")
    angle_text = f" Kąt {angle} pomaga skierować światło na konkretną część posesji." if angle else ""
    return f"{keyword} pozwala doświetlić miejsce, w którym po zmroku potrzebna jest dobra widoczność. {light_sentence(attrs, 'podjazd, wejście, bramę lub bok budynku')}{angle_text}"


def intro_naswietlacz_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} jest dobrym wyborem, gdy światło ma wspierać orientację i pracę na zewnątrz, a nie pełnić wyłącznie funkcję dekoracyjną. {light_sentence(attrs, 'otwartą przestrzeń lub elewację')} {ip_sentence(attrs, 'w instalacji zewnętrznej')}"


def intro_ramka_a(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} to element montażowy{producer_text(attrs)} przeznaczony do estetycznego osadzenia panelu LED. Format {dim_compact(attrs) or 'podany w specyfikacji'} pozwala dopasować ramkę do oprawy i zachować równe wykończenie sufitu."


def intro_ramka_b(keyword: str, attrs: dict[str, str]) -> str:
    dimensions = dim_compact(attrs)
    dim_text = f" Wymiary {dimensions} warto sprawdzić razem z formatem panelu." if dimensions else ""
    return f"{keyword} pomaga zamontować panel LED w miejscu, gdzie potrzebne jest stabilne i estetyczne wykończenie.{dim_text} Neutralny kolor ułatwia dopasowanie ramki do jasnego sufitu."


def intro_ramka_c(keyword: str, attrs: dict[str, str]) -> str:
    material = attrs.get("Materiał", "")
    material_text = f" Materiał {material} odpowiada za stabilne osadzenie oprawy." if material else ""
    return f"{keyword} pomaga zamontować panel LED poza typowym sufitem kasetonowym lub w miejscu wymagającym dodatkowego obramowania.{material_text} Format {dim_compact(attrs) or 'ramki'} decyduje o zgodności z konkretną oprawą."


def intro_zasilacz_a(keyword: str, attrs: dict[str, str]) -> str:
    power = value_with_unit(attrs.get("Moc [W]", ""), "W")
    power_text = f" Moc {power} trzeba zestawić z wymaganiami konkretnego panelu lub systemu." if power else ""
    return f"{keyword} to element zasilający przeznaczony do współpracy z kompatybilną oprawą.{power_text} Przy takim produkcie najważniejsza jest zgodność elektryczna i miejsce montażu."


def intro_zasilacz_b(keyword: str, attrs: dict[str, str]) -> str:
    values = natural_join([attrs.get("Napięcie [V]", ""), value_with_unit(attrs.get("Moc [W]", ""), "W")])
    value_text = f" Parametry {values} pomagają potwierdzić kompatybilność przed montażem." if values else ""
    return f"{keyword} warto traktować jako część konkretnego zestawu oświetleniowego, a nie uniwersalny dodatek do dowolnej oprawy.{value_text}"


def intro_zasilacz_c(keyword: str, attrs: dict[str, str]) -> str:
    return f"{keyword} odpowiada za zasilanie dopasowanej oprawy LED. Wymiary {dim_compact(attrs) or 'podane w specyfikacji'} ułatwiają zaplanowanie miejsca w suficie lub obudowie. Dobór takiego elementu powinien wynikać z parametrów panelu, a nie wyłącznie z nazwy serii."


def prose_material(attrs: dict[str, str]) -> str:
    """Material w prozie: 'ABS|PE' -> 'ABS i PE' (w specyfikacji separator '|' zostaje)."""
    return compact_spaces(attrs.get("Materiał", "").replace("|", " i "))


# Dopelniacz bez przyimka (przyimek z/ze dobierany osobno) - pozwala odmieniac tez zlozenia
# "Metal i Tworzywo sztuczne" -> "z metalu i tworzywa sztucznego".
_MATERIAL_GEN_BARE = {
    "szklo": "szkła",
    "metal": "metalu",
    "tworzywo sztuczne": "tworzywa sztucznego",
    "tworzywo": "tworzywa",
    "polipropylen": "polipropylenu",
    "poliweglan": "poliwęglanu",
    "drewno": "drewna",
    "aluminium": "aluminium",
    "stal": "stali",
}


def _material_is_codes(mat: str) -> bool:
    tokens = [t for t in re.split(r"\s+i\s+|\s+", mat.strip()) if t]
    return bool(tokens) and all(re.fullmatch(r"[A-Z]{2,}\d?", t) for t in tokens)


def material_is_plural(mat: str) -> bool:
    return len(re.split(r"\s+i\s+", compact_spaces(mat))) > 1


def _material_prep(bare: str) -> str:
    return "ze" if re.match(r"(sz|s[bcćdfghklmnprstwzż]|z)", bare) else "z"


def material_genitive_phrase(mat: str) -> str:
    """Zwraca poprawny dopelniacz z przyimkiem: 'Szklo' -> 'ze szkla', 'Metal' -> 'z metalu',
    'Metal i Tworzywo sztuczne' -> 'z metalu i tworzywa sztucznego', 'ABS i PE' -> 'z tworzyw
    ABS i PE'. Nieznane wartosci: bezpieczna apozycja 'z materialu X'."""
    mat = compact_spaces(mat)
    if _material_is_codes(mat):
        return f"z tworzyw {mat}" if material_is_plural(mat) else f"z tworzywa {mat}"
    parts = [p for p in re.split(r"\s+i\s+", mat) if p]
    bares: list[str] = []
    for part in parts:
        key = normalize_header(part)
        if key in _MATERIAL_GEN_BARE:
            bares.append(_MATERIAL_GEN_BARE[key])
        elif _material_is_codes(part):
            bares.append(f"tworzywa {part}")
        else:
            return f"z materiału {mat}"  # nieznany skladnik -> bezpieczny fallback
    return f"{_material_prep(bares[0])} {' i '.join(bares)}"


def mount_genitive(value: str) -> str:
    """Sposob montazu (przymiotnik) w dopelniaczu: 'Natynkowy' -> 'natynkowego' (do 'do montazu X').
    Dziala dla wszystkich wartosci -y/-i (podtynkowy, zwieszany, wpuszczany, nascienny)."""
    def dec(w: str) -> str:
        if re.fullmatch(r"[A-Za-ząćęłńóśźż]+", w) and w[-1].lower() in "yi":
            return w[:-1] + "ego"
        return w
    return " ".join(dec(w) for w in compact_spaces(value).split()).lower()


def render_benefit_item(item: str) -> str:
    """Punkt listy: 'cecha - korzysc' -> pogrubiona cecha; inaczej zwykly punkt z myslnikiem."""
    item = compact_spaces(item)
    if " - " in item:
        lead, rest = item.split(" - ", 1)
        return f"  <li>- <strong>{escape(lead)}</strong> - {escape(rest)}</li>"
    return f"  <li>- {escape(item)}</li>"


GARDEN_PLACES = "ścieżek, tarasów, ogrodów, wejść, rabat i stref wypoczynkowych"


def intro_ogrodowa_a(keyword: str, attrs: dict[str, str]) -> str:
    ip = attrs.get("Stopień ochrony [IP]", "")
    ip_text = f" Stopień ochrony {ip} pozwala na pracę na zewnątrz, w deszczu i przy zapyleniu." if ip else ""
    shape = " o ozdobnym, okrągłym kształcie" if normalize_header(attrs.get("Kształt", "")) == "okragly" else ""
    return f"{keyword} to zewnętrzna oprawa{shape}{producer_text(attrs)} do oświetlenia {GARDEN_PLACES}.{ip_text} Sprawdza się tam, gdzie po zmroku liczą się zarówno bezpieczne dojście, jak i dekoracyjny charakter światła."


def intro_ogrodowa_b(keyword: str, attrs: dict[str, str]) -> str:
    socket = f" Gniazdo {attrs['Trzonek']} pozwala samodzielnie dobrać żarówkę i barwę światła do klimatu ogrodu." if attrs.get("Trzonek") else ""
    return f"{keyword} doświetla ogród, taras i dojście do domu, łącząc funkcję użytkową z estetyką małej architektury.{socket} To wygodny wybór, gdy oświetlenie ma podkreślać aranżację przestrzeni wokół budynku."


def intro_ogrodowa_c(keyword: str, attrs: dict[str, str]) -> str:
    material = prose_material(attrs)
    material_text = f" Obudowa {material_genitive_phrase(material)} dobrze znosi zmienne warunki atmosferyczne." if material else ""
    return f"{keyword} jest przeznaczona do montażu na zewnątrz, gdzie oprawa pracuje w deszczu, wilgoci i zmiennych temperaturach.{material_text} Dobrze odnajdzie się przy ścieżkach, wejściach, rabatach oraz w strefie wypoczynku w ogrodzie."


def intro_ogolny_a(keyword: str, attrs: dict[str, str]) -> str:
    socket = f" Wymienny trzonek {attrs['Trzonek']} daje swobodę wyboru żarówki, więc łatwo dobrać cieplejsze światło do domu albo neutralne do strefy roboczej." if attrs.get("Trzonek") else ""
    ip = f" Stopień ochrony {attrs['Stopień ochrony [IP]']} oznacza {ip_meaning(attrs['Stopień ochrony [IP]'])}." if attrs.get("Stopień ochrony [IP]") else ""
    light = light_sentence(attrs, "korytarz, pomieszczenie techniczne lub strefę przejściową")
    body = f" {light}" if light else ""
    return f"{keyword}{producer_text(attrs)} to funkcjonalna oprawa do miejsc, w których liczy się proste światło i łatwy montaż.{body}{socket}{ip}"


def intro_oprawa_a(keyword: str, attrs: dict[str, str]) -> str:
    if is_channel_luminaire(attrs):
        return channel_intro_text(keyword, attrs)
    light = light_sentence(attrs, "korytarz, wnękę, przejście albo strefę użytkową")
    light_text = f" {light}" if light else ""
    source = f" Trzonek {attrs['Trzonek']} pozwala zastosować żarówkę o wybranej barwie i mocy." if attrs.get("Trzonek") else ""
    return f"{keyword} to oprawa{producer_text(attrs)} do oświetlenia wnętrz, w których liczy się prosty montaż i czysty efekt na suficie lub ścianie.{light_text}{source}"


def intro_oprawa_b(keyword: str, attrs: dict[str, str]) -> str:
    if is_channel_luminaire(attrs):
        return channel_intro_text(keyword, attrs)
    socket = f" Trzonek {attrs['Trzonek']} pozwala dobrać kompatybilne źródło światła." if attrs.get("Trzonek") else ""
    return f"{keyword} sprawdzi się jako funkcjonalny punkt światła w pomieszczeniu mieszkalnym, usługowym lub komunikacyjnym.{socket} Kolor {attrs.get('Kolor', 'oprawy')} warto dopasować do sufitu, ścian i pozostałych elementów wyposażenia."


def intro_oprawa_c(keyword: str, attrs: dict[str, str]) -> str:
    if is_channel_luminaire(attrs):
        return channel_intro_text(keyword, attrs)
    dimensions = dim_compact(attrs)
    dim_text = f" Wymiary {dimensions} pomagają przewidzieć, jak oprawa ułoży się w planowanym miejscu montażu." if dimensions else ""
    return f"{keyword} pasuje do instalacji, w której oprawa ma być widoczna, ale nie dominować nad aranżacją. Daje bazę pod praktyczne światło punktowe i pozwala dopasować efekt do wybranej żarówki.{dim_text}"


def intro_akcesorium_a(keyword: str, attrs: dict[str, str]) -> str:
    target = attrs.get("Pasuje do", "")
    target_text = ""
    if target:
        target_text = f" i jest przeznaczone {target}" if normalize_header(target).startswith("do ") else f" i jest przeznaczone do {target}"
    return f"{keyword} to akcesorium{producer_text(attrs)} do uzupełnienia lub serwisowania instalacji oświetleniowej{target_text}. W takim produkcie najważniejsze jest dopasowanie do właściwej oprawy, serii lub miejsca montażu, a nie parametry świetlne."


def intro_akcesorium_b(keyword: str, attrs: dict[str, str]) -> str:
    color = f" Wariant w kolorze {color_locative(attrs['Kolor'])} łatwiej dopasować do obudowy i miejsca montażu." if attrs.get("Kolor") else ""
    return f"{keyword} sprawdzi się przy uzupełnieniu lub wymianie elementu w istniejącej oprawie. Najważniejsze jest tu dopasowanie do właściwej serii, wymiaru i sposobu mocowania, bo podobne części nie zawsze są zamienne.{color}"


def intro_akcesorium_c(keyword: str, attrs: dict[str, str]) -> str:
    material = attrs.get("Materiał", "")
    material_text = f" Materiał {material} ma znaczenie dla trwałości i stabilności elementu." if material else ""
    return f"{keyword} jest elementem pomocniczym do instalacji oświetleniowej, przy którym kluczowa jest zgodność z konkretną oprawą lub systemem.{material_text} Przed montażem porównaj serię, kod i wymiary, zwłaszcza gdy produkt ma zastąpić zużyty albo brakujący element."


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
    "ogrodowa": [intro_ogrodowa_a, intro_ogrodowa_b, intro_ogrodowa_c],
    "ogolny": [intro_ogolny_a],
}


def benefit_place(family: str) -> str:
    return {
        "high_bay": "halę produkcyjną, magazyn lub strefę techniczną",
        "panel": "biuro, recepcję lub salę sprzedaży",
        "hermetyczna": "garaż, warsztat lub zaplecze",
        "plafon": "korytarz, wejście lub pomieszczenie pomocnicze",
        "naswietlacz": "elewację, podjazd lub plac",
        "oprawa": "wybrane wnętrze",
    }.get(family, "pomieszczenie, korytarz lub strefę techniczną")


# Priorytet punktow listy zalet per rodzina - najpierw to, co najmocniej definiuje produkt.
DEFAULT_BENEFIT_ORDER = [
    "power_flux", "efficacy", "color", "cri", "ugr", "angle", "sensor", "source",
    "dimmer", "ip", "ik", "durability", "energy", "montaz", "dim", "material", "kolor", "gwarancja",
]
FAMILY_BENEFIT_ORDER = {
    "high_bay": ["power_flux", "angle", "efficacy", "ip", "durability", "source", "ik", "color", "dim", "material", "energy", "gwarancja"],
    "panel": ["power_flux", "dim", "color", "cri", "ugr", "efficacy", "dimmer", "source", "ip", "durability", "montaz", "energy", "material", "kolor", "gwarancja"],
    "hermetyczna": ["power_flux", "ip", "durability", "dim", "source", "ik", "color", "material", "efficacy", "montaz", "kolor", "gwarancja"],
    "plafon": ["power_flux", "sensor", "color", "montaz", "dim", "source", "ip", "durability", "efficacy", "material", "kolor", "gwarancja"],
    "naswietlacz": ["power_flux", "angle", "ip", "efficacy", "color", "durability", "source", "ik", "montaz", "material", "kolor", "gwarancja"],
    "oprawa": ["source", "power_flux", "color", "dim", "montaz", "ip", "material", "kolor", "durability", "gwarancja"],
}


def benefits_luminaire(keyword: str, attrs: dict[str, str], family: str) -> list[str]:
    """Konkretne punkty cecha->korzysc z potwierdzonych atrybutow, w kolejnosci waznosci
    zaleznej od rodziny. Zero generycznych zdan typu 'Ten model zapewnia...'."""
    place = benefit_place(family)
    raw: dict[str, str] = {}

    power = attrs.get("Moc [W]", "")
    flux = attrs.get("Strumień świetlny [lm]", "")
    if power and flux:
        raw["power_flux"] = f"Moc {power} W i strumień {flux} lm - dają światło wystarczające, by równomiernie doświetlić {place}."
    elif flux:
        raw["power_flux"] = f"Strumień {flux} lm - ułatwia zaplanowanie ilości światła potrzebnej na {place}."
    elif power:
        raw["power_flux"] = f"Moc {power} W - pozwala dobrać oprawę do wielkości i funkcji przestrzeni."

    efficacy = attrs.get("Skuteczność świetlna [lm/W]", "") or attrs.get("Skuteczność świetlna", "")
    if efficacy:
        raw["efficacy"] = f"Skuteczność {efficacy} lm/W - oznacza niskie zużycie energii przy danej jasności i niższe koszty świecenia."

    temp = attrs.get("Temperatura barwowa [K]", "")
    barwa = attrs.get("Barwa światła", "")
    if temp and "-" not in temp:
        if temp == "4000":
            klimat = "neutralne, czytelne światło do pracy i codziennych czynności"
        elif temp in {"2700", "3000"}:
            klimat = "ciepłe, przytulne światło budujące komfortowy nastrój"
        else:
            klimat = "chłodne, wyraziste światło tam, gdzie liczy się dobra widoczność"
        raw["color"] = f"Barwa {temp} K - daje {klimat}."
    elif normalize_header(barwa) == "zmienna" or "-" in temp:
        raw["color"] = "Zmienna barwa światła - pozwala wybrać cieplejsze lub chłodniejsze światło zależnie od pory i potrzeby."
    elif barwa:
        raw["color"] = f"Barwa {barwa.lower()} - pozwala dopasować charakter światła do funkcji pomieszczenia."

    cri = attrs.get("Wskaźnik oddawania barw", "")
    if cri:
        raw["cri"] = f"Oddawanie barw Ra {cri} - sprawia, że kolory przedmiotów i wykończeń wyglądają naturalnie."

    ugr = attrs.get("Wskaźnik olśnienia [UGR]", "")
    if ugr:
        raw["ugr"] = f"Wskaźnik olśnienia UGR {ugr} - ogranicza efekt oślepienia i zmęczenie wzroku przy dłuższej pracy."

    angle = attrs.get("Kąt świecenia [°]", "") or attrs.get("Kąt świecenia", "")
    if angle:
        raw["angle"] = f"Kąt świecenia {angle}° - kieruje światło tam, gdzie jest potrzebne, bez rozpraszania na boki."

    if has_feature(attrs, "Czujnik ruchu"):
        raw["sensor"] = "Czujnik ruchu - automatycznie włącza światło, gdy ktoś korzysta z przestrzeni, co podnosi wygodę i ogranicza zużycie energii."

    source = attrs.get("Źródło światła", "")
    trzonek = attrs.get("Trzonek", "")
    if normalize_header(source).startswith("niez") and trzonek:
        raw["source"] = f"Popularny gwint {trzonek} - pozwala dobrać żarówkę LED do potrzebnego efektu, od cieplejszego światła po wyraźniejsze oświetlenie użytkowe."
    elif source and not normalize_header(source).startswith("niez"):
        raw["source"] = "Zintegrowane źródło LED - nie wymaga wymiany żarówki i ogranicza koszty obsługi przez lata."
    elif trzonek:
        raw["source"] = f"Gwint {trzonek} - daje wybór źródła światła pod kątem barwy, jasności i późniejszej wymiany."

    if has_feature(attrs, "Współpraca ze ściemniaczem"):
        raw["dimmer"] = "Współpraca ze ściemniaczem - umożliwia regulację jasności i dopasowanie światła do pory dnia oraz sytuacji."

    ip = attrs.get("Stopień ochrony [IP]", "")
    if ip:
        raw["ip"] = f"Klasa {ip} - {ip_short(ip)}, więc łatwiej dobrać oprawę do warunków montażu."

    ik = attrs.get("Stopień odporności [IK]", "")
    if ik:
        raw["ik"] = f"Odporność mechaniczna {ik} - zwiększa wytrzymałość na uderzenia w miejscach o intensywnej eksploatacji."

    durability = attrs.get("Trwałość [h]", "")
    if durability:
        raw["durability"] = f"Trwałość {durability} h - przekłada się na długą pracę bez wymiany i niższe koszty utrzymania."

    energy = attrs.get("Klasa energetyczna", "")
    if energy:
        raw["energy"] = f"Klasa energetyczna {energy} - ułatwia ocenę zużycia energii i kosztów świecenia."

    montaz = attrs.get("Sposób montażu", "")
    if montaz:
        raw["montaz"] = f"Montaż {montaz.lower()} - upraszcza instalację i dopasowanie oprawy do podłoża."

    d = dim_compact(attrs)
    if d:
        raw["dim"] = f"Wymiary {d} - pomagają ocenić proporcje oprawy i zaplanować miejsce montażu."

    if prose_material(attrs):
        raw["material"] = f"Obudowa {material_genitive_phrase(prose_material(attrs))} - wpływa na trwałość i odporność w codziennym użytkowaniu."

    if attrs.get("Kolor"):
        raw["kolor"] = f"Kolor {attrs['Kolor'].lower()} - ułatwia dopasowanie oprawy do wnętrza i stylu aranżacji."

    if attrs.get("Gwarancja"):
        raw["gwarancja"] = f"Gwarancja {attrs['Gwarancja']} - potwierdza trwałość i ułatwia decyzję o zakupie."

    order = FAMILY_BENEFIT_ORDER.get(family, DEFAULT_BENEFIT_ORDER)
    items = [raw[key] for key in order if key in raw]
    # Awaryjnie - gdy produkt ma bardzo malo atrybutow.
    if len(items) < 3:
        if attrs.get("Seria"):
            items.append(f"Seria {attrs['Seria']} - ułatwia dobór produktu zgodnego z resztą instalacji.")
        if attrs.get("Producent"):
            items.append(f"Producent {attrs['Producent']} - to rozpoznawalny dostawca osprzętu oświetleniowego.")
    return items


def benefit_common(keyword: str, attrs: dict[str, str]) -> list[str]:
    items = []
    if attrs.get("Stopień ochrony [IP]"):
        items.append(f"Stopień ochrony {attrs['Stopień ochrony [IP]']} - {ip_short(attrs['Stopień ochrony [IP]'])}.")
    if attrs.get("Kolor"):
        items.append(f"Kolor {attrs['Kolor'].lower()} - ułatwia dopasowanie oprawy do otoczenia i stylu aranżacji.")
    if prose_material(attrs):
        items.append(f"Materiał {prose_material(attrs)} - wpływa na trwałość i odporność w codziennym użytkowaniu.")
    if attrs.get("Gwarancja"):
        items.append(f"Gwarancja {attrs['Gwarancja']} - potwierdza trwałość i ułatwia decyzję o zakupie.")
    return items


def benefits_ramka(keyword: str, attrs: dict[str, str]) -> list[str]:
    items = [
        f"Format {dim_compact(attrs) or 'ramki'} - pozwala dobrać element do konkretnego rozmiaru panelu LED.",
        "Biały kolor - dobrze łączy się z jasnymi sufitami i nie odciąga uwagi od oświetlenia." if attrs.get("Kolor", "").lower() == "biały" else f"Kolor {attrs.get('Kolor', 'ramki').lower()} - warto dopasować do sufitu i oprawy.",
    ]
    if prose_material(attrs):
        items.append(f"Materiał {prose_material(attrs)} - poprawia stabilność osadzenia panelu.")
    items.append("Montaż natynkowy - ułatwia osadzenie panelu tam, gdzie nie trafia on bezpośrednio do sufitu kasetonowego.")
    if attrs.get("Seria"):
        items.append(f"Zgodność z serią {attrs['Seria']} - upraszcza dobór ramki do posiadanego panelu.")
    return items


def benefits_zasilacz(keyword: str, attrs: dict[str, str]) -> list[str]:
    items = []
    if attrs.get("Moc [W]"):
        items.append(f"Moc {value_with_unit(attrs['Moc [W]'], 'W')} - pozwala dopasować zasilacz do konkretnego wariantu panelu lub oprawy.")
    if attrs.get("Napięcie [V]"):
        items.append(f"Napięcie {attrs['Napięcie [V]']} - pomaga potwierdzić zgodność elektryczną przed montażem.")
    if dim_compact(attrs):
        items.append(f"Wymiary {dim_compact(attrs)} - ułatwiają zaplanowanie miejsca na zasilacz w instalacji.")
    if attrs.get("Stopień ochrony [IP]"):
        items.append(f"Stopień ochrony {attrs['Stopień ochrony [IP]']} - {ip_short(attrs['Stopień ochrony [IP]'])}.")
    if attrs.get("Seria"):
        items.append(f"Zgodność z serią {attrs['Seria']} - ogranicza ryzyko problemów z kompatybilnością zestawu.")
    if len(items) < 3:
        items.append("Dobór do konkretnej oprawy - jest ważniejszy niż wygląd, dlatego warto kierować się danymi technicznymi.")
    return items


def benefits_akcesorium(keyword: str, attrs: dict[str, str]) -> list[str]:
    items = []
    if attrs.get("Seria"):
        items.append(f"Seria {attrs['Seria']} - zawęża wybór do elementów pasujących do tej samej rodziny opraw.")
    if attrs.get("Kod producenta"):
        items.append(f"Kod producenta {attrs['Kod producenta']} - pozwala szybko potwierdzić, że to właśnie potrzebna część.")
    if dim_compact(attrs):
        items.append(f"Wymiary {dim_compact(attrs)} - warto porównać z miejscem montażu przed zakupem.")
    if prose_material(attrs):
        items.append(f"Materiał {prose_material(attrs)} - wpływa na trwałość i stabilność elementu.")
    if attrs.get("Kolor"):
        items.append(f"Kolor {attrs['Kolor'].lower()} - pozwala zachować spokojny wygląd oprawy po wymianie lub uzupełnieniu części.")
    if len(items) < 3:
        items.append("Wymiana samego elementu - pozwala odświeżyć lub uzupełnić oprawę bez kupowania całego nowego kompletu.")
    return items


def benefits_ogrodowa(keyword: str, attrs: dict[str, str]) -> list[str]:
    """Punkty w kolejnosci waznosci - najpierw cechy wyrozniajace produkt (solar, czujnik,
    zmienna barwa, zrodlo, moc), potem cechy ogolne. choose_benefits bierze pierwsze 6."""
    items: list[str] = []
    if normalize_header(attrs.get("Zasilanie", "")) == "solarne":
        items.append("Zasilanie solarne - nie wymaga prowadzenia przewodu zasilającego, więc lampę zamontujesz przy furtce, ścieżce, altanie, garażu, ogrodzeniu lub na tarasie.")
    if attrs.get("Czujnik ruchu") and attrs["Czujnik ruchu"].lower() not in {"nie"}:
        items.append("Czujnik ruchu - włącza światło automatycznie po wykryciu obecności i wygasza je po chwili, co poprawia bezpieczeństwo dojścia i ogranicza zużycie energii.")
    if normalize_header(attrs.get("Barwa światła", "")) == "zmienna" or "-" in attrs.get("Temperatura barwowa [K]", ""):
        zakres = f" {attrs['Temperatura barwowa [K]']} K" if "-" in attrs.get("Temperatura barwowa [K]", "") else ""
        items.append(f"Zmienna barwa światła{zakres} - pozwala wybrać cieplejsze, nastrojowe światło do ogrodu lub chłodniejsze, neutralne do lepszej widoczności przy wejściu i dojściu.")
    if normalize_header(attrs.get("Źródło światła", "")).startswith("niez"):
        items.append("Brak zintegrowanego źródła - pozwala dobrać cieplejsze światło do nastrojowego ogrodu lub neutralne do wejścia i dojścia.")
    elif attrs.get("Źródło światła"):
        items.append("Zintegrowane źródło LED - nie wymaga wymiany żarówki i ogranicza obsługę przez lata pracy na zewnątrz.")
    if attrs.get("Stopień ochrony [IP]"):
        items.append(f"Stopień ochrony {attrs['Stopień ochrony [IP]']} - {ip_short(attrs['Stopień ochrony [IP]'])}.")
    if attrs.get("Moc [W]"):
        items.append(f"Moc {attrs['Moc [W]']} W - zapewnia oszczędne, ale wystarczające światło do doświetlenia dojścia i strefy wokół domu.")
    if attrs.get("Trzonek"):
        items.append(f"Trzonek {attrs['Trzonek']} - daje swobodę doboru żarówki LED pod kątem mocy, barwy i efektu dekoracyjnego w ogrodzie.")
    if normalize_header(attrs.get("Kształt", "")) == "okragly":
        items.append("Okrągły, dekoracyjny kształt - daje miękkie, rozproszone światło i zdobi przestrzeń także za dnia.")
    if prose_material(attrs):
        items.append(f"Obudowa {material_genitive_phrase(prose_material(attrs))} - jest odporna na warunki zewnętrzne i łatwa w utrzymaniu.")
    if attrs.get("Kolor"):
        items.append(f"Kolor {attrs['Kolor'].lower()} - łatwo komponuje się z elewacją, tarasem i aranżacją ogrodu.")
    if attrs.get("Gwarancja"):
        items.append(f"Gwarancja {attrs['Gwarancja']} - potwierdza trwałość produktu w codziennym użytkowaniu na zewnątrz.")
    if len(items) < 3:
        items.append("Oświetlenie zewnętrzne - poprawia bezpieczeństwo wieczornego dojścia i podkreśla aranżację posesji.")
    return items


BENEFIT_BUILDERS = {
    "high_bay": lambda keyword, attrs: benefits_luminaire(keyword, attrs, "high_bay"),
    "panel": lambda keyword, attrs: benefits_luminaire(keyword, attrs, "panel"),
    "hermetyczna": lambda keyword, attrs: benefits_luminaire(keyword, attrs, "hermetyczna"),
    "plafon": lambda keyword, attrs: benefits_luminaire(keyword, attrs, "plafon"),
    "naswietlacz": lambda keyword, attrs: benefits_luminaire(keyword, attrs, "naswietlacz"),
    "ramka": benefits_ramka,
    "zasilacz": benefits_zasilacz,
    "akcesorium": benefits_akcesorium,
    "oprawa": lambda keyword, attrs: benefits_luminaire(keyword, attrs, "oprawa"),
    "ogrodowa": benefits_ogrodowa,
    "ogolny": lambda keyword, attrs: benefits_luminaire(keyword, attrs, "ogolny"),
}


def advice_high_bay_a(attrs: dict[str, str]) -> str:
    return f"W oświetleniu halowym znaczenie ma wysokość montażu, rozstaw opraw i kierunek świecenia. Seria {attrs.get('Seria', 'High Bay')} sprawdzi się w obiektach, gdzie światło pracuje długo i musi zachować powtarzalny efekt na całej powierzchni roboczej."


def advice_high_bay_b(attrs: dict[str, str]) -> str:
    return f"Przy wyborze warto zestawić strumień świetlny, kąt świecenia i stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')} z realnymi warunkami obiektu. Inne potrzeby ma stanowisko pracy, a inne szeroka alejka magazynowa."


def advice_panel_a(attrs: dict[str, str]) -> str:
    return f"Panel LED najlepiej sprawdza się tam, gdzie potrzebne jest spokojne światło na większej powierzchni. Przed zakupem warto sprawdzić format {dim_compact(attrs) or 'panelu'}, oczekiwaną jasność i barwę światła, bo te parametry najmocniej wpływają na komfort pracy."


def advice_panel_b(attrs: dict[str, str]) -> str:
    return f"Seria {attrs.get('Seria', 'Kanlux')} może zastąpić starsze oprawy świetlówkowe albo uzupełnić nowy układ oświetlenia w suficie modułowym. Przy modernizacji warto porównać wymiary panelu z istniejącym sufitem i zaplanować liczbę punktów."


def advice_hermetic_a(attrs: dict[str, str]) -> str:
    ip = attrs.get("Stopień ochrony [IP]", "")
    ip_text = f" Stopień ochrony {ip} oznacza {ip_meaning(ip)}." if ip else ""
    return f"W oprawach technicznych liczy się nie tylko jasność, ale też odporność obudowy i łatwe utrzymanie czystości.{ip_text} Warunki pracy warto porównać z miejscem montażu."


def advice_hermetic_b(attrs: dict[str, str]) -> str:
    return f"Przed montażem warto porównać długość, sposób zasilania i materiał obudowy z warunkami konkretnego pomieszczenia. W garażu, warsztacie lub zapleczu oprawa ma przede wszystkim działać pewnie i nie wymagać częstej obsługi."


def advice_plafon_a(attrs: dict[str, str]) -> str:
    if has_feature(attrs, "Czujnik ruchu"):
        ip = f" i stopień ochrony {attrs['Stopień ochrony [IP]']}" if attrs.get("Stopień ochrony [IP]") else ""
        return f"Plafon z czujnikiem sprawdza się w miejscach, gdzie światło ma włączać się tylko wtedy, gdy ktoś faktycznie korzysta z przestrzeni. Przy montażu warto uwzględnić zasięg czujnika, średnicę oprawy{ip}."
    ip = f" oraz stopień ochrony {attrs['Stopień ochrony [IP]']}" if attrs.get("Stopień ochrony [IP]") else ""
    return f"Plafon dobrze sprawdza się w korytarzach, wejściach i pomieszczeniach pomocniczych, gdzie liczy się równomierne światło na całej powierzchni. Przy montażu warto uwzględnić średnicę oprawy{ip}."


def advice_plafon_b(attrs: dict[str, str]) -> str:
    return "W korytarzach i przejściach ważne jest proste działanie oraz równomierne oświetlenie najczęściej używanej strefy. Dlatego warto dobrać oprawę do wysokości montażu, wielkości pomieszczenia i oczekiwanego sposobu sterowania."


def advice_naswietlacz_a(attrs: dict[str, str]) -> str:
    return f"Naświetlacz powinien być dobrany do szerokości strefy, wysokości montażu i kierunku, w którym ma świecić. W praktyce największe znaczenie mają strumień świetlny, kąt świecenia i stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')}."


def advice_naswietlacz_b(attrs: dict[str, str]) -> str:
    return "Przy oświetleniu zewnętrznym lepiej dobrać światło do konkretnej strefy niż kierować je przypadkowo na całą przestrzeń. Pozwala to poprawić widoczność przy wejściu, bramie lub podjeździe bez nadmiernego rozproszenia światła."


def advice_ramka_a(attrs: dict[str, str]) -> str:
    return f"Przy takim akcesorium najważniejsze jest zgodne dopasowanie do panelu i sposób wykończenia sufitu. Wysokość oraz format {dim_compact(attrs) or 'podane w specyfikacji'} warto sprawdzić przed zakupem razem z wymiarem oprawy."


def advice_ramka_b(attrs: dict[str, str]) -> str:
    return "Dobrze dobrana ramka ułatwia późniejszy dostęp serwisowy i pozwala zachować spójny wygląd kilku paneli zamontowanych w jednym pomieszczeniu. To szczególnie ważne w biurach, lokalach usługowych i korytarzach."


def advice_zasilacz_a(attrs: dict[str, str]) -> str:
    return "Zasilacz warto traktować jako część konkretnego systemu, a nie uniwersalny dodatek do dowolnej oprawy. Przed montażem trzeba porównać moc, napięcie i typ współpracującego panelu."


def advice_zasilacz_b(attrs: dict[str, str]) -> str:
    return "Dobrze dobrany element zasilający ułatwia późniejszą obsługę instalacji i ogranicza ryzyko problemów z kompatybilnością. Znaczenie ma także miejsce montażu oraz dostęp do elementu w razie serwisu."


def advice_ogolny_a(attrs: dict[str, str]) -> str:
    socket = f" Trzonek {attrs['Trzonek']} pozwala dobrać żarówkę do oczekiwanej barwy i jasności." if attrs.get("Trzonek") else ""
    ip = f" Stopień ochrony {attrs['Stopień ochrony [IP]']} wskazuje, w jakich warunkach oprawa pracuje bezpiecznie." if attrs.get("Stopień ochrony [IP]") else ""
    return f"O praktycznej wartości tej oprawy decyduje połączenie jej parametrów z miejscem, w którym ma świecić.{socket}{ip}".strip()


def advice_ogrodowa_a(attrs: dict[str, str]) -> str:
    return "W ogrodzie światło najlepiej rozplanować wzdłuż ścieżek, przy wejściu i w strefie wypoczynku, aby poprawić bezpieczeństwo po zmroku i podkreślić aranżację posesji. Oprawa zewnętrzna powinna być odporna na wilgoć i dobrze komponować się z otoczeniem także za dnia."


def advice_ogrodowa_b(attrs: dict[str, str]) -> str:
    socket = f" Dzięki gniazdu {attrs['Trzonek']} dobierzesz cieplejszą barwę do nastrojowego ogrodu albo neutralną do dojścia i wejścia." if attrs.get("Trzonek") else ""
    return f"Lampę ogrodową warto dopasować do charakteru przestrzeni - inny efekt sprawdzi się przy nowoczesnej elewacji, a inny w zielonym, naturalnym ogrodzie.{socket}"


def advice_oprawa_a(attrs: dict[str, str]) -> str:
    if is_channel_luminaire(attrs):
        return "Oprawę kanałową warto dobrać do miejsca, w którym światło ma przede wszystkim ułatwiać przejście, wejście albo obsługę zaplecza. Liczy się odporność obudowy, prosty dostęp do żarówki i zgodność z planowanym sposobem montażu."
    return "Przy oprawach sufitowych i punktowych warto zacząć od miejsca montażu, oczekiwanego efektu światła oraz typu źródła. Dopiero potem dobrze porównać kolor, wymiary i sposób mocowania z wystrojem pomieszczenia."


def advice_oprawa_b(attrs: dict[str, str]) -> str:
    if is_channel_luminaire(attrs):
        return "W ciągach komunikacyjnych i pomieszczeniach technicznych lepiej sprawdza się oprawa prosta, szczelna i łatwa do utrzymania w czystości. Taki wybór ogranicza problem późniejszego serwisu i dobrze pasuje do miejsc, w których światło ma przede wszystkim działać pewnie na co dzień."
    return "Jeśli oprawa ma być widoczna we wnętrzu, znaczenie ma proporcja produktu do sufitu, kolor obudowy i rodzaj źródła światła. Wymiary, trzonek oraz dopuszczalną moc warto zestawić z konkretnym miejscem montażu."


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
    "ogrodowa": [advice_ogrodowa_a, advice_ogrodowa_b],
    "ogolny": [advice_ogolny_a],
}


def build_technical_rows(attrs: dict[str, str]) -> list[tuple[str, str]]:
    rows = []
    for key in TECHNICAL_ORDER:
        if key in LOGISTIC_SKIP_FEATURES:
            continue
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
            "Przed montażem najlepiej zestawić produkt z konkretnym miejscem pracy: wysokością montażu, dostępną przestrzenią "
            "oraz oczekiwanym sposobem użytkowania. Dzięki temu łatwiej wybrać wariant, który będzie pasował nie tylko nazwą, "
            "ale też codziennym działaniem."
        )
        html = html.replace("<h3>Specyfikacja techniczna</h3>", f"<p>{escape(fallback)}</p>\n\n<h3>Specyfikacja techniczna</h3>", 1)
    if chars_without_spaces(strip_html(html)) < min_chars_no_spaces:
        final_note = (
            "Jeżeli produkt ma zastąpić wcześniejszy element, porównaj jego wymiary, sposób montażu i oznaczenie producenta "
            "z częścią, która jest już w instalacji. To ogranicza ryzyko wyboru podobnego, ale niepasującego wariantu."
        )
        html = html.replace("<h3>Specyfikacja techniczna</h3>", f"<p>{escape(final_note)}</p>\n\n<h3>Specyfikacja techniczna</h3>", 1)
    return html


def extra_application(keyword: str, attrs: dict[str, str], family: str) -> str:
    if family == "ramka":
        return "Ta ramka jest szczególnie przydatna wtedy, gdy panel LED ma zostać zamontowany poza standardowym sufitem modułowym. Pomaga utrzymać równą linię montażu i ogranicza wrażenie przypadkowego dołożenia oprawy."
    if family == "zasilacz":
        return "W przypadku zasilacza kluczowa jest zgodność z oprawą, a nie wygląd elementu. Warto zachować dostęp do miejsca montażu, bo ułatwia to diagnostykę i ewentualną wymianę w przyszłości."
    if family == "akcesorium":
        return "To akcesorium warto dobrać do konkretnej oprawy, a nie tylko do ogólnej kategorii produktu. Taki element może decydować o stabilności montażu, ochronie oprawy albo wygodzie późniejszego serwisu."
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
    if family == "ogrodowa":
        return "Oświetlenie ogrodu pełni dwie role: poprawia bezpieczeństwo wieczornego dojścia i buduje nastrój wokół domu. Dobrze rozmieszczone lampy podkreślają ścieżki, taras, rabaty oraz detale małej architektury, a odporna obudowa pozwala im pracować przez cały rok."
    return "W tej oprawie najważniejsze jest dopasowanie jej parametrów do miejsca i sposobu użytkowania, bo to one decydują o realnym efekcie świetlnym i wygodzie codziennej eksploatacji."


def extra_selection_advice(attrs: dict[str, str], family: str) -> str:
    if family == "ogrodowa":
        if normalize_header(attrs.get("Barwa światła", "")) == "zmienna" or "-" in attrs.get("Temperatura barwowa [K]", ""):
            zakres = f" {attrs['Temperatura barwowa [K]']} K" if "-" in attrs.get("Temperatura barwowa [K]", "") else ""
            return f"Zmienna barwa światła{zakres} pozwala dopasować klimat oświetlenia do pory i potrzeby: cieplejsze, nastrojowe światło sprawdzi się w strefie wypoczynku w ogrodzie, a chłodniejsze, neutralne poprawi widoczność przy wejściu i dojściu do domu."
        return "W ogrodzie światło najlepiej rozplanować wzdłuż ścieżek, przy wejściu i w strefie wypoczynku, aby po zmroku poprawić bezpieczeństwo dojścia i podkreślić aranżację posesji oraz detale małej architektury."
    if family in {"panel", "high_bay", "hermetyczna", "naswietlacz"}:
        return f"Jeżeli produkt ma pracować przez wiele godzin dziennie, warto zwrócić uwagę na strumień świetlny, barwę światła oraz stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')}. Te parametry najmocniej wpływają na wygodę użytkowania i trwałość instalacji."
    if family == "plafon":
        return "W korytarzach, wejściach i pomieszczeniach pomocniczych liczy się równomierne światło oraz prosty montaż. Przy wyborze warto porównać średnicę oprawy, barwę światła i sposób montażu z miejscem instalacji."
    if family == "ramka":
        return "Przed zakupem ramki najlepiej porównać jej format z wymiarem panelu oraz sprawdzić wysokość profilu. Dzięki temu montaż będzie wyglądał spójnie także przy kilku oprawach w jednym pomieszczeniu."
    if family == "zasilacz":
        return "Przed zakupem warto sprawdzić moc, napięcie i serię oprawy. Zasilacz o podobnej nazwie nie zawsze oznacza pełną zgodność z konkretnym panelem."
    if family == "akcesorium":
        return "Przy częściach pomocniczych najlepiej porównać serię, kod producenta, kształt oraz sposób mocowania. To ogranicza ryzyko zakupu elementu, który wygląda podobnie, ale nie pasuje do danej oprawy."
    if family == "oprawa":
        return "Przed wyborem dobrze sprawdzić, czy oprawa pasuje do planowanego źródła światła i czy jej wymiary są odpowiednie do danego sufitu lub ściany. Ma to znaczenie zarówno dla montażu, jak i końcowego efektu wizualnego."
    return f"Najważniejsze przy doborze są jasność, barwa światła oraz stopień ochrony {attrs.get('Stopień ochrony [IP]', 'IP')} - to one decydują, czy oprawa sprawdzi się w danym miejscu i jak długo posłuży."


def extra_mounting_note(attrs: dict[str, str], family: str) -> str:
    if family in {"ogrodowa", "naswietlacz"} and normalize_header(attrs.get("Zasilanie", "")) == "solarne":
        return "Zasilanie solarne nie wymaga prowadzenia przewodu zasilającego ani podłączenia do sieci, dlatego lampę zamontujesz w wygodnym miejscu - przy furtce, ścieżce, altanie, garażu, wejściu bocznym, ogrodzeniu lub na tarasie - tam, gdzie poprowadzenie kabla byłoby kłopotliwe."
    dimensions = dim_compact(attrs)
    if dimensions:
        return pick_variant([
            f"Wymiary {dimensions} pomagają zaplanować miejsce montażu i uniknąć kolizji z innymi elementami instalacji. Ma to znaczenie szczególnie wtedy, gdy oprawa trafia do istniejącej zabudowy.",
            f"Znając wymiary {dimensions}, łatwo ocenisz, czy produkt zmieści się w przewidzianym miejscu montażu.",
            f"Format {dimensions} warto mieć z tyłu głowy przy planowaniu rozmieszczenia, zwłaszcza w gotowej już zabudowie.",
        ], stable_seed(first_present(attrs.get("Kod producenta", ""), dimensions)))
    if family in {"ogrodowa", "naswietlacz"}:
        return "Montaż na zewnątrz warto zaplanować tak, aby światło realnie obejmowało używaną strefę - wejście, ścieżkę, taras lub elewację - oraz aby zachować łatwy dostęp do oprawy na potrzeby konserwacji."
    return "Przed montażem warto sprawdzić miejsce instalacji, sposób zasilania i dostęp serwisowy. Takie podejście ogranicza ryzyko późniejszych poprawek."


# --- Pkt 7: kontrola zgodnosci tytul <-> specyfikacja -----------------------------------
# Cecha obecna TYLKO w nazwie (bez pokrycia w atrybutach) jest NIEPEWNA. Generator i tak
# opiera sie wylacznie na `attrs` (has_feature), wiec nie buduje na niej korzysci - ale
# rozjazd nazwa/specyfikacja zglaszamy do recznego przegladu (kolumna "uwagi").
TITLE_FEATURE_MARKERS = [
    ("czujnik", "Czujnik ruchu", "czujnik ruchu w nazwie, brak w specyfikacji"),
    ("sciemnia", "Współpraca ze ściemniaczem", " sciemnianie w nazwie, brak w specyfikacji"),
    ("dali", "Współpraca ze ściemniaczem", "DALI w nazwie, brak potwierdzenia w specyfikacji"),
]


def title_spec_consistency(title: str, attrs: dict[str, str], family: str) -> list[str]:
    warnings: list[str] = []
    t = normalize_header(title)
    for marker, attr, msg in TITLE_FEATURE_MARKERS:
        if marker in t and not has_feature(attrs, attr):
            warnings.append(msg.strip())
    if "solarn" in t and normalize_header(attrs.get("Zasilanie", "")) != "solarne":
        warnings.append("zasilanie solarne w nazwie, brak w specyfikacji")
    # Sprzecznosc klasy IP: nazwa vs specyfikacja.
    m_ip = re.search(r"\bip\s?(\d{2})\b", t)
    if m_ip:
        spec_ip = re.sub(r"\D", "", attrs.get("Stopień ochrony [IP]", ""))
        if not spec_ip:
            warnings.append(f"IP{m_ip.group(1)} w nazwie, brak IP w specyfikacji")
        elif spec_ip != m_ip.group(1):
            warnings.append(f"IP w nazwie (IP{m_ip.group(1)}) != IP w specyfikacji (IP {spec_ip})")
    # Sprzecznosc mocy: porownujemy tylko, gdy obie wartosci sa liczba calkowita (bez zakresow).
    m_w = re.search(r"\b(\d{1,4})\s?w\b", t)
    spec_w_raw = attrs.get("Moc [W]", "")
    if m_w and re.fullmatch(r"\d+", spec_w_raw) and m_w.group(1).lstrip("0") != spec_w_raw.lstrip("0"):
        warnings.append(f"moc w nazwie ({m_w.group(1)} W) != moc w specyfikacji ({spec_w_raw} W)")
    # Sprzecznosc temperatury barwowej.
    m_k = re.search(r"\b(\d{4})\s?k\b", t)
    spec_k_raw = attrs.get("Temperatura barwowa [K]", "")
    if m_k and re.fullmatch(r"\d+", spec_k_raw) and m_k.group(1) != spec_k_raw:
        warnings.append(f"barwa w nazwie ({m_k.group(1)} K) != specyfikacja ({spec_k_raw} K)")
    return warnings


# --- Pkt 10: warstwa walidacji SEMANTYCZNEJ po wygenerowaniu tekstu ----------------------
# Zdanie jest KONKRETNE, gdy niesie dane (moc/lm/K/IP/IK/trzonek) ALBO konkretna kotwice:
# miejsce zastosowania, material, cecha. Sam rzeczownik produktu (oprawa/panel) NIE liczy sie -
# inaczej kazde zdanie byloby "konkretne". Miara sluzy jako tripwire ogolnikowosci (pkt 9/10).
CONCRETE_TOKEN = re.compile(
    r"(\d+\s?(?:w|lm|k|mm|cm|h|v)\b|\bip\s?\d|\bik\s?\d|\be27\b|\bgu10\b|\bg9\b|\bgx\d|\bdali\b|\d+\s?x\s?\d+"
    r"|hal[ai]|magazyn|biur|recepcj|korytarz|gara|warsztat|taras|ogrod|ogród|elewacj|wejści|wejsci"
    r"|ścieżk|sciezk|podjazd|sufi|ścian|scian|łazienk|lazienk|kuchni|salon|sypialni|zaplecz|przejśc|przejsc"
    r"|słupk|slupek|cokoł|aluminium|stal|szkł|szkl|tworzyw|poliw|klosz|radiator|czujnik|ściemnia|sciemnia"
    r"|trzonek|solarn|żarów|zarow|driver)",
    re.IGNORECASE,
)
# Frazy, ktore twierdza, ze SAM produkt emituje/rozklada swiatlo lub reaguje na ruch -
# nie wolno ich uzyc dla zasilacza/ramki/akcesorium (pkt 3) ani gdy brak czujnika.
LIGHT_EMIT_PHRASES = ["oswietla przestrzen", "rozlozy swiatlo", "rozklada swiatlo", "daje swiatlo", "swieci"]
SENSOR_PHRASES = ["reaguje na ruch", "czujnik ruchu", "wykrywa obecnosc", "zapala sie samoczynnie", "wlacza swiatlo automatycznie", "gdy ktos"]


def prose_raw_text(html: str) -> str:
    main = html.split("<hr>")[0]
    paragraphs = re.findall(r"<p\b[^>]*>(.*?)</p>", main, flags=re.S | re.I)
    return compact_spaces(" ".join(strip_html(p) for p in paragraphs))


def prose_text_normalized(html: str) -> str:
    # normalize_header usuwa kropki, wiec granice zdan gina - do liczenia zdan uzyj prose_raw_text.
    return normalize_header(prose_raw_text(html))


def semantic_validation_issues(html: str, attrs: dict[str, str], family: str) -> list[str]:
    """Sprawdza WYGENEROWANY tekst pod katem zgodnosci znaczeniowej z kategoria i atrybutami:
    zasilacz/ramka nie jak oprawa, brak czujnika = brak zdan o czujniku, miejsce zgodne z
    kategoria, dostateczna konkretnosc. Zwraca miekkie uwagi (nie blokuje generacji)."""
    issues: list[str] = []
    prose = prose_text_normalized(html)
    # 1. Bezpiecznik kategorii - produkt nieswiecacy nie moze byc opisany jak oprawa.
    if not is_light_product(family):
        hit = [p for p in LIGHT_EMIT_PHRASES if p in prose]
        if hit:
            issues.append(f"kategoria: {family} opisany jak oprawa ({hit[0]})")
    # 2. Czujnik - zdanie o czujniku bez atrybutu potwierdzajacego.
    if not has_feature(attrs, "Czujnik ruchu"):
        hit = [p for p in SENSOR_PHRASES if p in prose]
        if hit:
            issues.append(f"czujnik: tekst sugeruje czujnik, atrybut go nie potwierdza ({hit[0]})")
    # 3. Miejsce zgodne z kategoria.
    outdoor = family in {"ogrodowa", "naswietlacz"}
    if outdoor and ("na suficie" in prose or "wpisuje sie w sufit" in prose):
        issues.append("miejsce: produkt zewnetrzny opisany przez sufit")
    if family in INDOOR_LIGHT_FAMILIES and any(p in prose for p in ["na taras", "do ogrodu", "w ogrodzie", "na elewacj"]):
        issues.append("miejsce: produkt wewnetrzny opisany przez ogrod/elewacje")
    # 4. Konkretnosc - udzial zdan prozy z co najmniej jednym konkretem (zdania z surowego
    #    tekstu, bo normalize_header usuwa kropki i zlepia zdania w jedno).
    sentences = [s for s in re.split(r"(?<=[.!?])\s+", prose_raw_text(html)) if len(s.split()) >= 4]
    if sentences:
        concrete = sum(1 for s in sentences if CONCRETE_TOKEN.search(s))
        ratio = concrete / len(sentences)
        if ratio < 0.30:  # tripwire ogolnikowosci: flaguje tylko wyraznie najslabsze opisy
            issues.append(f"niska konkretnosc prozy ({concrete}/{len(sentences)} zdan)")
    # 5. GRAMATYKA/MERYT - regresje odmiany i tresci (naprawione u zrodla, tu straznik):
    raw = prose_text_normalized(html)
    if re.search(r"\bz (szklo|metal|tworzywo|drewno|stal|aluminium)\b", raw):
        issues.append("gramatyka: nieodmieniony material po 'z'")
    if re.search(r"w (wariancie|sposob) \w+(owy|any|nny|ony)\b", raw):
        issues.append("gramatyka: nieodmieniony sposob montazu")
    ip_digits = re.sub(r"\D", "", attrs.get("Stopień ochrony [IP]", ""))
    if ip_digits and ip_digits[-1] in "01" and re.search(r"(pyloszczel|odporno[sś]c na wod|strug[ei] wody|odporn\w* na wilgo)", raw):
        issues.append(f"meryt: IP{ip_digits} (suche wnetrza) opisany jak odporny na wode/pyl")
    return issues


def review_warnings(html: str, title: str, attrs: dict[str, str], family: str) -> list[str]:
    """Laczy uwagi do recznego przegladu (pkt 7 + pkt 10) - do kolumny 'uwagi' w raporcie."""
    return title_spec_consistency(title, attrs, family) + semantic_validation_issues(html, attrs, family)


def assert_semantic_safety(html: str, attrs: dict[str, str], family: str) -> None:
    """Twardy straznik dla sygnatur dawnych bledow: sciemniacz renderowany jak czujnik oraz
    zasilacz/ramka opisany jak oprawa. Te przypadki to zawsze regresja - przerywamy glosno
    zamiast wyslac zly opis. Miekkie uwagi (miejsce, konkretnosc) zostaja w review_warnings."""
    prose = prose_text_normalized(html)
    if not is_light_product(family):
        for phrase in LIGHT_EMIT_PHRASES:
            if phrase in prose:
                raise ValueError(f"Regresja: {family} opisany jak oprawa ({phrase})")
    if not has_feature(attrs, "Czujnik ruchu"):
        for phrase in ["reaguje na ruch", "czujnik ruchu"]:
            if phrase in prose:
                raise ValueError(f"Regresja: opis sugeruje czujnik ruchu bez atrybutu ({phrase})")


def validate_description(html: str) -> None:
    for phrase in FORBIDDEN_PHRASES:
        if phrase in html:
            raise ValueError(f"Opis zawiera zakazana fraze: {phrase}")


TECHNICAL_LEAD_PATTERNS = [
    r"^moc\b",
    r"^strumie[nń]\b",
    r"^barwa\b",
    r"^temperatura barwowa\b",
    r"^trzonek\b",
    r"^gwint\b",
    r"^stopie[nń] ochrony\b",
    r"^klasa ip\b",
    r"^wymiary\b",
    r"^kolor\b",
    r"^materia[lł]\b",
    r"^seria\b",
    r"^kod producenta\b",
    r"^napi[eę]cie\b",
]


def starts_like_parameter(text: str) -> bool:
    normalized = normalize_header(text)
    return any(re.search(pattern, normalized) for pattern in TECHNICAL_LEAD_PATTERNS)


def algorithmic_style_issues(html: str) -> list[str]:
    issues: list[str] = []
    main = html.split("<hr>")[0]
    prose_paragraphs = re.findall(r"<p\b[^>]*>(.*?)</p>", main, flags=re.S | re.I)
    text = " ".join(strip_html(paragraph) for paragraph in prose_paragraphs)
    sentences = [compact_spaces(sentence) for sentence in re.split(r"(?<=[.!?])\s+", text) if compact_spaces(sentence)]
    parameter_starts = [sentence for sentence in sentences if starts_like_parameter(sentence)]
    if len(parameter_starts) >= 4:
        issues.append(f"za duzo zdan zaczyna sie od parametru ({len(parameter_starts)})")

    list_items = [strip_html(item) for item in re.findall(r"<li\b[^>]*>(.*?)</li>", main, flags=re.S | re.I)]
    benefit_items = []
    if "Najważniejsze cechy" in main:
        before_spec = main.split("Specyfikacja techniczna")[0]
        benefit_items = [strip_html(item) for item in re.findall(r"<li\b[^>]*>(.*?)</li>", before_spec, flags=re.S | re.I)]
    technical_benefits = [item for item in benefit_items if starts_like_parameter(item.lstrip("-–— ").split(" - ", 1)[0])]
    if len(technical_benefits) >= 2:
        issues.append("lista cech zaczyna sie od parametrow zamiast decyzji kupujacego")
    if len(list_items) >= 3:
        feature_arrow = [item for item in list_items if re.search(r"\b(?:pozwala|ułatwia|oznacza|pomaga)\b", normalize_header(item))]
        if len(feature_arrow) > max(2, len(list_items) // 2):
            issues.append("za duzo punktow ma schemat cecha -> korzysc")
    return issues


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
