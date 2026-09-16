"""Zbiera frazy z poprawionych opisow do banku z zakresem rodziny produktowej.

Po co: generator czerpie proze z `dictionaries/seo_phrase_bank_harvested.yaml`,
ktory jest WSPOLNY dla calego katalogu. Dlatego przy wczytywaniu przepuszcza go
bardzo ostry filtr `is_safe_flavor` - odrzuca zdania ze slowami `gniazd`,
`laczni`, `puszk`, `czujnik`, `podtynkow`, `zacisk`. Bank globalny nie wie, do
jakiego produktu fraza trafi, wiec musi zakladac najgorsze.

Bank z zakresem rodziny to wie. Zdanie o montazu podtynkowym jest bezpieczne
dla rodziny `gniazdo_zasilajace` i nigdy nie trafi na plafon, bo jest zapisane
pod konkretna rodzina i domena. Dzieki temu mozna zbierac slownictwo branzowe,
ktore w banku globalnym musialoby zostac wyciete.

Czego skrypt NIE robi: nie pisze prozy. Wylacznie wycina zdania z gotowych,
zatwierdzonych opisow, odrzuca te nienadajace sie do ponownego uzycia i
grupuje reszte.

Wynik idzie domyslnie do katalogu roboczego w `reports/` - do przejrzenia.
Dopiero `--merge-into` dopisuje frazy do slownika w `dictionaries/`.
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import html as html_mod
import pathlib
import re
import sys

import yaml

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "src"))

from generate_product_descriptions import (  # noqa: E402
    classify_family,
    collect_attributes,
    family_entry,
    is_light_product,
    product_title,
)
from utils import compact_spaces, read_products  # noqa: E402

# --- Co dyskwalifikuje zdanie z ponownego uzycia -----------------------------

# Liczba = parametr konkretnego produktu. Takie zdanie przy innym SKU klamie.
_HAS_DIGIT = re.compile(r"\d")
# Kod produktu / marka pisana kapitalikami (TGZ1CZ, BRAVO, ISDN). LED/IP/USB zostaja.
_CODE_OR_BRAND = re.compile(r"\b(?!LED\b|UV\b|RGB\b|USB\b|AC\b|DC\b|PC\b)[A-ZĄĆĘŁŃÓŚŹŻ]{2,}\b")
# Wielka litera w srodku zdania to niemal zawsze nazwa wlasna.
_MIDCAP = re.compile(r"(?<=[a-ząćęłńóśźż] )[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]{2,}")
# Zdanie odwolujace sie do poprzedniego nie moze stac samodzielnie.
_CONTEXT_DEPENDENT = re.compile(
    r"^(dzięki (temu|niemu|niej|nim|nią|nimi)|temu |dlatego |z tego (powodu|względu)"
    r"|oprócz tego|ponadto|poza tym|w ten sposób|to (rozwiązanie|sprawia|oznacza|pozwala)"
    r"|jego |jej |ich |taki |taka |takie |ten |ta |to )",
    re.I,
)
# Urwany skrot na koncu = zdanie niekompletne.
_DANGLING = re.compile(r"\b(np|m\.in|tzn|tj|itd|wg|ok|tzw)\.$", re.I)
# Zwrot do klienta i meta-komentarz nie nadaja sie do banku.
_ADDRESSED = re.compile(r"\b(wybierz|sprawdź|kupujesz|zamów|Twoj|Twój|Twoja|Twoje)\b", re.I)
# Kolory obudowy roznia sie miedzy wariantami tego samego produktu.
_COLOR = re.compile(
    r"(biał|czarn|szar|grafit|srebrn|złot|miedzian|brązow|beżow|kaszmir|niebiesk"
    r"|zielon|czerwon|żółt|antracyt|kremow|różow|fioletow)",
    re.I,
)

MIN_CHARS = 45
MAX_CHARS = 240

# Zdanie, ktore NAZYWA swoj typ produktu, jest prawdziwe tylko w swojej rodzinie.
# "Lacznik schodowy pracuje w ukladzie bistabilnym" nie moze trafic na gniazdo,
# a "Gniazdo tej wersji ma ruchome przeslony" nie moze trafic na lacznik.
# Takie zdania zapisujemy wylacznie w `by_family`, nigdy w szerszym `by_domain`.
_PRODUCT_NOUN = re.compile(
    r"(gniazd|łączni|łącznik|przycisk|ramk|klawisz|puszk|zaślepk|pokryw|ściemniacz"
    r"|czujnik|sterownik|ładowark|oprawa|oprawy|plafon|panel|naświetlacz|żarówk"
    r"|szynoprzewod|taśm|peszel|koryt|rozdzielnic|wyłączni)",
    re.I,
)
# Slownictwo swietlne nie moze trafic do domen nieoswietleniowych.
# Uwaga na odmiane: "punktu swietlnego" i "kontrolke swietlna" nie lapie ani
# `swiat[ll]`, ani `oswietl` - potrzebny jest rdzen `swietl`.
_LIGHT_WORDS = re.compile(r"(świat[łl]|świetl|jasnoś|strumie|barw|lamp|żarów|świec)", re.I)


def portable_across_family(sentence: str, domain: str) -> bool:
    """Czy zdanie wolno wrzucic do szerszej puli domeny, a nie tylko rodziny."""
    if _PRODUCT_NOUN.search(sentence):
        return False
    if domain != "oswietlenie" and _LIGHT_WORDS.search(sentence):
        return False
    return True

# --- Kategorie fraz ---------------------------------------------------------
# Nazwy zgodne z kategoriami w seo_phrase_bank_harvested.yaml, zeby oba banki
# dalo sie czytac tym samym kodem.
CATEGORY_RULES: list[tuple[str, re.Pattern[str]]] = [
    ("montaz_uzytkowanie", re.compile(
        r"(montaż|montuje|zamontow|puszk|zacisk|przewod|instalac|wkręt|pazurk"
        r"|szybkozłącz|podłącz|mocow|wpina|osadz)", re.I)),
    ("swiatlo_barwa", re.compile(
        r"(świat[łl]|oświetl|barw|jasnoś|strumie|doświetl|olśnien)", re.I)),
    ("energo_eko", re.compile(r"(energi|zużyci|oszczęd|pobór|klasa energet)", re.I)),
    ("jakosc_trwalosc", re.compile(
        r"(trwał|wytrzym|odporn|jakoś|solidn|bezhalogen|materiał|eksploatac)", re.I)),
    ("design_estetyka", re.compile(
        r"(wyglą|estetyk|matow|powierzchni|wykończen|aranżac|wzornic|spójn|dyskretn)", re.I)),
    ("zastosowanie_wnetrza", re.compile(
        r"(pokoj|pokój|salon|sypialn|kuchn|łazien|korytarz|biur|mieszkan|wnętrz|pomieszczen)", re.I)),
    ("zastosowanie_zewnetrze", re.compile(
        r"(na zewnątrz|elewac|ogrod|taras|garaż|zewnętrzn|wilgot|bryzg)", re.I)),
    ("funkcje_korzysci", re.compile(
        r"(pozwala|umożliwia|obsług|sterow|funkcj|działa|reguluj|zapewnia)", re.I)),
]


def sentence_category(sentence: str) -> str:
    for name, pattern in CATEGORY_RULES:
        if pattern.search(sentence):
            return name
    return "ogolne"


def rejection_reason(sentence: str) -> str:
    """Zwraca powod odrzucenia albo pusty string, gdy zdanie nadaje sie do banku."""
    if not MIN_CHARS <= len(sentence) <= MAX_CHARS:
        return "dlugosc"
    if _HAS_DIGIT.search(sentence):
        return "liczba (parametr konkretnego produktu)"
    if _CODE_OR_BRAND.search(sentence):
        return "kod produktu albo marka"
    if _MIDCAP.search(sentence):
        return "nazwa wlasna w srodku zdania"
    if _CONTEXT_DEPENDENT.match(sentence):
        return "zdanie zalezne od kontekstu"
    if _DANGLING.search(sentence):
        return "urwany skrot na koncu"
    if _ADDRESSED.search(sentence):
        return "bezposredni zwrot do klienta"
    if _COLOR.search(sentence):
        return "kolor obudowy"
    return ""


def body_sentences(markup: str) -> list[str]:
    """Zdania z prozy opisu.

    Bierzemy WYLACZNIE akapity <p>, i to bez pierwszego:

    - naglowki <h2>/<h3> to pytania o konkretny model, nie material na bank;
    - punkty <li> to pary cecha-korzysc, czyli warstwa argumentacyjna prawdziwa
      tylko dla jednego produktu - tego bank z zasady nie potrafi skalowac;
    - pierwszy akapit zaczyna sie od pelnej nazwy produktu;
    - sekcja FAQ ma inny rejestr, a specyfikacja to pary etykieta-wartosc.
    """
    body = markup.split("<hr>")[0]
    paragraphs = re.findall(r"<p\b[^>]*>(.*?)</p>", body, flags=re.I | re.S)[1:]
    sentences: list[str] = []
    for paragraph in paragraphs:
        text = html_mod.unescape(re.sub(r"<[^>]+>", " ", paragraph))
        sentences.extend(compact_spaces(s) for s in re.split(r"(?<=[.!?])\s+", text))
    return sentences


def harvest_file(markup: str, family: str) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
    kept: list[tuple[str, str]] = []
    dropped: list[tuple[str, str]] = []
    for sentence in body_sentences(markup):
        reason = rejection_reason(sentence)
        if reason:
            if len(sentence) >= MIN_CHARS:
                dropped.append((sentence, reason))
            continue
        kept.append((sentence, sentence_category(sentence)))
    return kept, dropped


def product_name_from_markup(markup: str) -> str:
    """Pelna nazwa produktu z otwarcia opisu.

    Kontrakt opisu wymaga, zeby pierwszy akapit zaczynal sie od
    <strong>pelnej nazwy produktu</strong>, wiec to wiarygodne zrodlo nazwy
    dla plikow HTML, ktore nie niosa zadnych metadanych.
    """
    match = re.search(r"<p\b[^>]*>\s*<strong\b[^>]*>(.*?)</strong>", markup, flags=re.I | re.S)
    if not match:
        return ""
    return compact_spaces(html_mod.unescape(re.sub(r"<[^>]+>", " ", match.group(1))))


def load_sources(args: argparse.Namespace) -> list[tuple[str, str, str]]:
    """Zwraca liste (etykieta, nazwa produktu, html opisu)."""
    items: list[tuple[str, str, str]] = []
    if args.html_dir:
        for path in sorted(pathlib.Path(args.html_dir).glob("*.html")):
            markup = path.read_text(encoding="utf-8")
            items.append((path.stem, product_name_from_markup(markup) or path.stem, markup))
    if args.input:
        df = read_products(args.input, sheet_name=args.sheet)
        column = args.description_column
        if column not in df.columns:
            raise SystemExit(f"Brak kolumny '{column}'. Dostepne: {list(df.columns)[:20]}")
        for _, row in df.iterrows():
            markup = str(row.get(column, "") or "")
            if markup.strip():
                items.append((str(row.get("sku", "")), product_title(row, args.title_column), markup))
    if not items:
        raise SystemExit("Brak zrodel: podaj --html-dir albo --input.")
    return items


def resolve_family(label: str, title: str, args: argparse.Namespace) -> str:
    if args.family:
        return args.family
    return classify_family(title, {"__working_title": title})


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--html-dir", default="", help="Katalog z gotowymi plikami .html opisow.")
    parser.add_argument("--input", default="", help="XLSX/CSV z kolumna opisu.")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--description-column", default="description")
    parser.add_argument("--title-column", default="")
    parser.add_argument("--family", default="", help="Wymusza rodzine dla calej paczki.")
    parser.add_argument("--output", required=True, help="Plik YAML z zebranymi frazami (do przegladu).")
    parser.add_argument("--rejected-report", default="", help="Opcjonalny raport odrzuconych zdan.")
    parser.add_argument(
        "--merge-into",
        default="",
        help="Slownik do uzupelnienia, np. dictionaries/seo_phrase_bank_by_family.yaml. "
             "Bez tej flagi skrypt niczego nie dopisuje do dictionaries/.",
    )
    args = parser.parse_args()

    sources = load_sources(args)
    by_family: dict[str, dict[str, list[str]]] = collections.defaultdict(lambda: collections.defaultdict(list))
    by_domain: dict[str, dict[str, list[str]]] = collections.defaultdict(lambda: collections.defaultdict(list))
    dropped_all: list[tuple[str, str, str]] = []
    families_seen: collections.Counter[str] = collections.Counter()
    kept_count = 0

    for label, title, markup in sources:
        family = resolve_family(label, title, args)
        families_seen[family] += 1
        entry = family_entry(family)
        domain = str(entry.get("domain", "nieznana"))
        kept, dropped = harvest_file(markup, family)
        for sentence, category in kept:
            if sentence not in by_family[family][category]:
                by_family[family][category].append(sentence)
            if portable_across_family(sentence, domain) and sentence not in by_domain[domain][category]:
                by_domain[domain][category].append(sentence)
            kept_count += 1
        dropped_all.extend((label, s, r) for s, r in dropped)

    payload = {
        "version": 1,
        "metadata": {
            "generated_at": dt.date.today().isoformat(),
            "sources": len(sources),
            "sentences_kept": kept_count,
            "sentences_rejected": len(dropped_all),
            "families": dict(families_seen),
            "note": (
                "Frazy maja zakres rodziny i domeny produktowej. Nie wolno ich "
                "splaszczac do wspolnej puli - slownictwo osprzetu wyciekloby na oprawy."
            ),
        },
        "by_family": {f: dict(c) for f, c in sorted(by_family.items())},
        "by_domain": {d: dict(c) for d, c in sorted(by_domain.items())},
    }

    output = pathlib.Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False, width=100), encoding="utf-8")
    print(f"OK: zebrano {kept_count} fraz z {len(sources)} opisow -> {output}")
    print(f"    odrzucono {len(dropped_all)} zdan jako nienadajace sie do ponownego uzycia")
    print(f"    rodziny: {dict(families_seen)}")

    if args.rejected_report:
        lines = ["ODRZUCONE ZDANIA (nie nadaja sie do banku)", ""]
        for label, sentence, reason in dropped_all:
            lines.append(f"[{reason}] {label}")
            lines.append(f"    {sentence}")
        pathlib.Path(args.rejected_report).write_text("\n".join(lines), encoding="utf-8")
        print(f"    raport odrzuconych -> {args.rejected_report}")

    if args.merge_into:
        merge_into_bank(pathlib.Path(args.merge_into), payload)


def merge_into_bank(path: pathlib.Path, payload: dict) -> None:
    existing = {}
    if path.exists():
        existing = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    added = 0
    for scope in ("by_family", "by_domain"):
        target = existing.setdefault(scope, {})
        for key, categories in payload[scope].items():
            bucket = target.setdefault(key, {})
            for category, sentences in categories.items():
                current = bucket.setdefault(category, [])
                for sentence in sentences:
                    if sentence not in current:
                        current.append(sentence)
                        added += 1
    existing["version"] = 1
    meta = existing.setdefault("metadata", {})
    meta["last_merged_at"] = payload["metadata"]["generated_at"]
    meta["note"] = payload["metadata"]["note"]
    path.write_text(yaml.safe_dump(existing, allow_unicode=True, sort_keys=False, width=100), encoding="utf-8")
    print(f"    dopisano {added} nowych fraz do {path}")


if __name__ == "__main__":
    main()
