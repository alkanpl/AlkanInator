from __future__ import annotations

import argparse
import csv
import json
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd

from utils import compact_spaces, is_blank, load_yaml, normalize_header, read_products


DEFAULT_SKU_SUFFIX_KNOWLEDGE_PATH = "dictionaries/learned_product_taxonomy.yaml"
DEFAULT_CATALOG_KNOWLEDGE_PATH = "dictionaries/woocommerce_catalog_knowledge.yaml"
DEFAULT_SUPPLIER_GLOBAL_KNOWLEDGE_PATH = "dictionaries/supplier_global_knowledge.yaml"
DEFAULT_ATTRIBUTE_KNOWLEDGE_PATH = "dictionaries/lighting_attribute_knowledge.yaml"
FEATURE_REVIEW_COLUMNS = ["sku", "ean", "name", "feature_name", "feature_value", "reason"]
EXCLUDED_BASELINKER_FEATURES = {"EAN (GTIN)", "Kod producenta"}
MANUALLY_APPROVED_FEATURE_VALUES = {
    "Typ produktu": {
        "Czujnik ruchu",
        "Klips",
        "Konektor",
        "Lampa wisząca",
        "Linka do podwieszenia",
        "Łącznik",
        "Naświetlacz LED",
        "Oprawa High Bay",
        "Oprawa hermetyczna",
        "Oprawa kanałowa",
        "Oprawa sufitowa",
        "Oprawa sufitowa punktowa",
        "Ramka do paneli",
        "Panel LED",
        "Pilot do oprawy",
        "Plafon",
        "Ramka",
        "Soczewka",
        "Siatka ochronna",
        "Uchwyt",
        "Wspornik montażowy",
        "Zapinka",
        "Zasilacz",
    },
    "Seria": {
        "ACETE",
        "ADTR",
        "AGOR",
        "AGZAR",
        "ALIN",
        "AMAZI",
        "ANBAR",
        "ANTEM",
        "ANTO",
        "AQILO",
        "ARVOS",
        "AVAR",
        "AZPO",
        "BAREV",
        "BENO",
        "BLINGO",
        "BLURRO",
        "BONSA",
        "BORD",
        "BRAVO",
        "CARVO",
        "CORSO",
        "DABA",
        "DABER",
        "DICHT",
        "DUNO",
        "ENELO",
        "ERTON",
        "EXATE",
        "FL",
        "FLS",
        "FOGLER",
        "FT1200",
        "FT1500",
        "FTD1200",
        "FTD1500",
        "FTHL1500",
        "FTHT1500",
        "GALOBA",
        "GORD",
        "GRUN",
        "HBPA",
        "HBPHS",
        "HERMI",
        "IQ-LED",
        "IPER",
        "INES",
        "JASMIN",
        "JURBA",
        "LUMKO",
        "MAH",
        "MARC",
        "MILO",
        "NIFU",
        "PHLOX",
        "PIRES",
        "PLAFMIN",
        "PLAVE",
        "RIFA",
        "RITI",
        "S",
        "SANGA",
        "SANI",
        "SANSO",
        "SOLN",
        "STATO",
        "STIVI",
        "STOBI",
        "TP",
        "TOLEO",
        "TOLU",
        "TUNA",
        "TURA",
        "TURK",
        "TYBIA",
        "VARSO",
        "VAND",
    },
    "Kolor producenta": {
        "Biały mat",
        "Czarny mat",
        "Dąb sonoma",
        "Nikiel satynowy",
        "STAL INOX",
        "Wenge",
        "Złoto-brązowy",
        "Złoty dąb",
    },
    "Zastosowanie": {"Do paneli"},
}
DROP_UNKNOWN_FEATURES = {"Model", "Sterowanie"}
NON_BASE_COLORS_AS_PRODUCER_COLOR = {"Biały mat", "Czarny mat", "Dąb sonoma", "Nikiel satynowy", "STAL INOX", "Wenge", "Złoty dąb"}
ACCESSORY_PRODUCT_TYPES_WITHOUT_OWN_POWER = {
    "Klips",
    "Klosz",
    "Linka do podwieszenia",
    "Pilot do oprawy",
    "Ramka",
    "Ramka do paneli",
    "Siatka ochronna",
    "Soczewka",
    "Uchwyt",
    "Wspornik montażowy",
    "Zapinka",
    "Łącznik",
}


def default_supplier_feature_policy() -> dict[str, Any]:
    return {
        "excluded_features": set(EXCLUDED_BASELINKER_FEATURES),
        "drop_unknown_features": set(DROP_UNKNOWN_FEATURES),
        "non_base_colors_as_producer_color": set(NON_BASE_COLORS_AS_PRODUCER_COLOR),
        "accessory_product_types_without_own_power": set(ACCESSORY_PRODUCT_TYPES_WITHOUT_OWN_POWER),
        "approved_feature_values": {
            feature_name: set(values)
            for feature_name, values in MANUALLY_APPROVED_FEATURE_VALUES.items()
        },
    }


def load_supplier_feature_policy(path: str | Path | None = DEFAULT_SUPPLIER_GLOBAL_KNOWLEDGE_PATH) -> dict[str, Any]:
    policy = default_supplier_feature_policy()
    knowledge = load_yaml(path) if path else {}
    supplier = ((knowledge.get("suppliers") or {}).get("kanlux") or {}) if isinstance(knowledge, dict) else {}
    if not supplier:
        return policy

    accepted = supplier.get("accepted_feature_values") or {}
    merge_approved_values(policy, "Typ produktu", accepted.get("product_types") or [])
    merge_approved_values(policy, "Seria", accepted.get("series") or [])
    merge_approved_values(policy, "Kolor producenta", accepted.get("producer_colors") or [])
    merge_approved_values(policy, "Zastosowanie", accepted.get("applications") or [])

    feature_rules = supplier.get("feature_export_rules") or {}
    policy["excluded_features"].update(str(value) for value in feature_rules.get("excluded_features") or [])
    policy["drop_unknown_features"].update(str(value) for value in feature_rules.get("drop_unknown_features") or [])
    policy["non_base_colors_as_producer_color"].update(str(value) for value in accepted.get("producer_colors") or [])

    accessory_rules = supplier.get("accessory_rules") or {}
    policy["accessory_product_types_without_own_power"].update(
        str(value) for value in accessory_rules.get("product_types_without_own_power") or []
    )
    return policy


def merge_approved_values(policy: dict[str, Any], feature_name: str, values: list[Any]) -> None:
    approved = policy["approved_feature_values"].setdefault(feature_name, set())
    approved.update(str(value) for value in values if compact_spaces(str(value or "")))

BASELINKER_COLUMNS = [
    "product_id",
    "name",
    "sku",
    "ean",
    "manufacturer_name",
    "category",
    "description",
    "features",
    "images_urls",
]

FEATURE_MAP = {
    "Producent": ["attr_producent", "Producent", "manufacturer_name", "producer", "Marka"],
    "Dane producenta": ["Dane producenta", "Dane Producenta", "manufacturer_data", "producer_data"],
    "EAN (GTIN)": ["EAN", "ean"],
    "Kod producenta": ["Kod Producenta", "Kod", "SKU", "sku"],
    "Typ produktu": ["attr_typ", "Typ", "Rodzaj produktu"],
    "Seria": ["attr_seria", "Rodzina"],
    "Model": ["attr_model", "Nazwa Kanlux"],
    "Moc [W]": ["attr_moc", "Moc [W]", "Moc"],
    "Napięcie [V]": ["attr_napiecie", "Napięcie [V]", "Napięcie (V)"],
    "Temperatura barwowa [K]": ["attr_barwa", "Barwa CCT [K]", "Temperatura barwowa"],
    "Barwa światła": ["Barwa - kategoria", "attr_barwa_zakres", "Barwa światła"],
    "Strumień świetlny [lm]": ["attr_strumien", "Strumień [lm]", "Strumień świetlny [lm]", "Jasność"],
    "Stopień ochrony [IP]": ["attr_ip", "Klasa IP", "Stopień ochrony [IP]", "Stopień ochrony IP"],
    "Stopień ochrony [IK]": ["attr_ik", "Klasa ochronności"],
    "Kąt świecenia [°]": ["attr_kat_swiecenia", "Kąt", "Kąt świecenia"],
    "Kolor": ["attr_kolor", "Kolor obudowy"],
    "Kolor producenta": ["Kolor producenta", "attr_kolor_producenta", "attr_kolor", "Kolor obudowy"],
    "Trzonek": ["attr_gwint", "Gwint", "Trzonek", "Rodzaj gwintu"],
    "Materiał": ["attr_material", "Materiał"],
    "Kształt": ["attr_ksztalt", "Kształt"],
    "Wymiary [mm]": ["attr_wymiary", "Wymiary"],
    "Długość": ["attr_dlugosc", "Długość [mm]", "Długość"],
    "Szerokość": ["attr_szerokosc", "Szerokość [mm]", "Szerokość"],
    "Wysokość": ["attr_wysokosc", "Wysokość [mm]", "Wysokość"],
    "Średnica": ["attr_srednica", "Średnica [mm]", "Średnica"],
    "Głębokość": ["attr_glebokosc", "Głębokość [mm]", "Głębokość"],
    "Czujnik ruchu": ["attr_czujnik", "Czujnik ruchu"],
    "Opakowanie": ["attr_ilosc_sztuk"],
    "Gwarancja": ["attr_gwarancja"],
    "Skuteczność świetlna [lm/W]": ["attr_lm_w"],
    "Wskaźnik oddawania barw": ["attr_cri"],
    "Trwałość [h]": ["attr_trwalosc"],
    "Wskaźnik olśnienia [UGR]": ["attr_ugr", "UGR<19"],
    "Możliwość łączenia": ["attr_laczenie_przelotowe"],
    "Współpraca ze ściemniaczem": ["attr_sciemnianie", "Ściemnialne"],
    "Źródło światła": ["attr_zrodlo_swiatla"],
    "Źródło światła w komplecie": ["attr_zrodlo_w_komplecie"],
    "Sterowanie": ["attr_sterowanie"],
    "Zastosowanie": ["attr_pasuje_do"],
    "Klasa energetyczna": ["attr_klasa_energetyczna", "Klasa EEi", "Klasa efektywności energetycznej"],
    "Sposób montażu": ["attr_sposob_montazu", "Sposób montażu"],
    "Maksymalna moc źródła światła": ["attr_moc_max_zrodla"],
    "Liczba źródeł światła": ["attr_liczba_gniazd", "Liczba źródeł światła", "Liczba gniazd"],
    "Zasilanie": ["attr_zasilanie", "Zasilanie"],
}

FEATURE_NAME_ALIASES = {
    "marka": "Producent",
    "producent": "Producent",
    "dane producenta": "Dane producenta",
    "manufacturer data": "Dane producenta",
    "producer data": "Dane producenta",
    "rodzaj produktu": "Typ produktu",
    "typ": "Typ produktu",
    "typ produktu": "Typ produktu",
    "moc": "Moc [W]",
    "moc w": "Moc [W]",
    "moc znamionowa": "Moc [W]",
    "napiecie": "Napięcie [V]",
    "napięcie": "Napięcie [V]",
    "napiecie v": "Napięcie [V]",
    "napięcie v": "Napięcie [V]",
    "temperatura barwowa": "Temperatura barwowa [K]",
    "temperatura barwowa k": "Temperatura barwowa [K]",
    "barwa cct k": "Temperatura barwowa [K]",
    "jasnosc": "Strumień świetlny [lm]",
    "jasność": "Strumień świetlny [lm]",
    "strumien lm": "Strumień świetlny [lm]",
    "strumień lm": "Strumień świetlny [lm]",
    "strumien swietlny lm": "Strumień świetlny [lm]",
    "strumień świetlny lm": "Strumień świetlny [lm]",
    "stopien ochrony ip": "Stopień ochrony [IP]",
    "stopień ochrony ip": "Stopień ochrony [IP]",
    "klasa ip": "Stopień ochrony [IP]",
    "kat swiecenia": "Kąt świecenia [°]",
    "kąt świecenia": "Kąt świecenia [°]",
    "kat swiecenia °": "Kąt świecenia [°]",
    "kąt świecenia °": "Kąt świecenia [°]",
    "kat swiecenia deg": "Kąt świecenia [°]",
    "rodzaj gwintu": "Trzonek",
    "gwint": "Trzonek",
    "trzonek": "Trzonek",
    "barwa swiatla": "Barwa światła",
    "barwa światła": "Barwa światła",
    "material": "Materiał",
    "materiał": "Materiał",
    "kolor obudowy": "Kolor",
    "kolor producenta": "Kolor producenta",
    "ksztalt": "Kształt",
    "kształt": "Kształt",
    "ksztalt oprawy": "Kształt",
    "kształt oprawy": "Kształt",
    "model": "Model",
    "klasa ochronnosci": "Stopień ochrony [IK]",
    "klasa ochronności": "Stopień ochrony [IK]",
    "stopien ochrony ik": "Stopień ochrony [IK]",
    "stopień ochrony ik": "Stopień ochrony [IK]",
    "liczba sztuk": "Opakowanie",
    "ilosc sztuk": "Opakowanie",
    "ilość sztuk": "Opakowanie",
    "pasuje do": "Zastosowanie",
    "wymiary": "Wymiary [mm]",
    "wymiary mm": "Wymiary [mm]",
    "dlugosc": "Długość",
    "długość": "Długość",
    "dlugosc mm": "Długość",
    "długość mm": "Długość",
    "szerokosc": "Szerokość",
    "szerokość": "Szerokość",
    "szerokosc mm": "Szerokość",
    "szerokość mm": "Szerokość",
    "wysokosc": "Wysokość",
    "wysokość": "Wysokość",
    "wysokosc mm": "Wysokość",
    "wysokość mm": "Wysokość",
    "srednica": "Średnica",
    "średnica": "Średnica",
    "srednica mm": "Średnica",
    "średnica mm": "Średnica",
    "glebokosc": "Głębokość",
    "głębokość": "Głębokość",
    "glebokosc mm": "Głębokość",
    "głębokość mm": "Głębokość",
    "sposob montazu": "Sposób montażu",
    "sposób montażu": "Sposób montażu",
    "miejsce montazu": "Sposób montażu",
    "klasa energetyczna": "Klasa energetyczna",
    "liczba gniazd": "Liczba źródeł światła",
    "ilosc gniazd": "Liczba źródeł światła",
    "liczba zrodel swiatla": "Liczba źródeł światła",
    "moc maksymalna zarowki": "Maksymalna moc źródła światła",
    "maksymalna moc zrodla swiatla": "Maksymalna moc źródła światła",
    "skutecznosc swietlna": "Skuteczność świetlna [lm/W]",
    "skuteczność świetlna": "Skuteczność świetlna [lm/W]",
    "skutecznosc swietlna lm w": "Skuteczność świetlna [lm/W]",
    "skuteczność świetlna lm w": "Skuteczność świetlna [lm/W]",
    "trwalosc": "Trwałość [h]",
    "trwałość": "Trwałość [h]",
    "trwalosc h": "Trwałość [h]",
    "trwałość h": "Trwałość [h]",
    "zrodlo swiatla w komplecie": "Źródło światła w komplecie",
    "źródło światła w komplecie": "Źródło światła w komplecie",
    "klasa efektywnosci energetycznej": "Klasa energetyczna",
}


def main() -> None:
    parser = argparse.ArgumentParser(description="Eksportuje wynik AlkanInatora do CSV importowego Baselinkera.")
    parser.add_argument("--input", required=True, help="CSV/XLSX po AlkanInatorze.")
    parser.add_argument("--sheet", help="Arkusz XLSX, jesli input ma wiele arkuszy.")
    parser.add_argument("--output", default="output/baselinker_import.csv", help="Docelowy CSV dla Baselinkera.")
    parser.add_argument("--name-column", default="", help="Wymusza kolumne nazwy. Domyslnie: new_title, Nazwa, Nazwa B2C / SEO final.")
    parser.add_argument(
        "--category-column",
        default="",
        help="Wymusza kolumne kategorii. Domyslnie: proponowana_kategoria_1, Kategoria, category.",
    )
    parser.add_argument("--include-empty-features", action="store_true", help="Zapisuje puste parametry w JSON features.")
    parser.add_argument("--links-file", default="", help="XLSX/CSV z linkami do zdjec, np. input/Linki.xlsx.")
    parser.add_argument("--links-sheet", default=None, help="Arkusz w pliku linkow. Domyslnie pierwszy arkusz.")
    parser.add_argument("--image-key-column", default="Kod Kanlux", help="Kolumna kodu produktu w pliku linkow.")
    parser.add_argument("--main-image-column", default="Zdjecie glowne", help="Kolumna glownego zdjecia w pliku linkow.")
    parser.add_argument(
        "--sku-format",
        choices=["plain", "producer_suffix"],
        default="producer_suffix",
        help="Format SKU w eksporcie. producer_suffix daje np. 22606/KAN.",
    )
    parser.add_argument(
        "--sku-suffix-knowledge",
        default=DEFAULT_SKU_SUFFIX_KNOWLEDGE_PATH,
        help="Slownik z nauczonymi suffixami SKU producentow.",
    )
    parser.add_argument(
        "--catalog-knowledge",
        default=DEFAULT_CATALOG_KNOWLEDGE_PATH,
        help="Slownik wiedzy katalogowej uzywany pomocniczo przy dopasowaniu suffixow.",
    )
    parser.add_argument(
        "--feature-review-output",
        default="",
        help="CSV z atrybutami/wartosciami, ktore wymagaja akceptacji przed dodaniem do Woo.",
    )
    parser.add_argument(
        "--supplier-knowledge",
        default=DEFAULT_SUPPLIER_GLOBAL_KNOWLEDGE_PATH,
        help="Globalny slownik wiedzy o dostawcach, m.in. zaakceptowane wartosci Kanlux.",
    )
    parser.add_argument(
        "--attribute-knowledge",
        default=DEFAULT_ATTRIBUTE_KNOWLEDGE_PATH,
        help="Reguly doboru i kolejnosci atrybutow dla kategorii oswietleniowych.",
    )
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    image_matches = 0
    if args.links_file:
        df, image_matches = assign_main_images(
            df,
            args.links_file,
            args.links_sheet,
            args.image_key_column,
            args.main_image_column,
        )
    catalog_knowledge = load_yaml(args.catalog_knowledge) if args.catalog_knowledge else {}
    producer_suffixes = load_producer_suffixes(args.sku_suffix_knowledge, args.catalog_knowledge)
    manufacturer_data_by_producer = build_manufacturer_data_by_producer(catalog_knowledge)
    feature_value_normalizer = build_feature_value_normalizer(catalog_knowledge)
    feature_policy = load_supplier_feature_policy(args.supplier_knowledge)
    attribute_knowledge = load_yaml(args.attribute_knowledge) if args.attribute_knowledge else {}
    feature_review_rows: list[dict[str, str]] = []
    exported = build_baselinker_rows(
        df,
        args.name_column,
        args.category_column,
        args.include_empty_features,
        args.sku_format,
        producer_suffixes,
        manufacturer_data_by_producer,
        feature_value_normalizer,
        feature_review_rows,
        feature_policy,
        attribute_knowledge,
    )
    write_baselinker_csv(exported, args.output)
    feature_review_output = args.feature_review_output or default_feature_review_output(args.output)
    write_feature_review_file(feature_review_rows, feature_review_output)

    print(f"OK: wczytano {len(df)} produktow")
    if args.links_file:
        print(f"OK: przypisano zdjecia glowne: {image_matches}")
    print(f"OK: zapisano {args.output}")
    print(f"OK: zapisano raport atrybutow do akceptacji: {feature_review_output}")


def build_baselinker_rows(
    df: pd.DataFrame,
    name_column: str = "",
    category_column: str = "",
    include_empty_features: bool = False,
    sku_format: str = "producer_suffix",
    producer_suffixes: dict[str, str] | None = None,
    manufacturer_data_by_producer: dict[str, str] | None = None,
    feature_value_normalizer: dict[str, dict[str, str]] | None = None,
    feature_review_rows: list[dict[str, str]] | None = None,
    feature_policy: dict[str, Any] | None = None,
    attribute_knowledge: dict[str, Any] | None = None,
) -> list[dict[str, str]]:
    if sku_format == "producer_suffix" and producer_suffixes is None:
        producer_suffixes = load_producer_suffixes()

    rows: list[dict[str, str]] = []
    for _, row in df.iterrows():
        name = first_value(row, [name_column] if name_column else [])
        if not name:
            name = first_value(row, ["new_title", "Nazwa B2C / SEO final", "Nazwa", "Nazwa B2C / SEO", "Nazwa Kanlux"])

        sku = first_value(row, ["SKU", "Kod Producenta", "Kod", "sku"])
        ean = only_digits(first_value(row, ["EAN", "ean"]))
        producer = first_value(row, ["attr_producent", "Producent", "manufacturer_name", "Marka"])
        export_sku = format_sku(sku, producer, sku_format, producer_suffixes)
        category = first_value(
            row,
            [category_column] if category_column else [
                "proponowana_kategoria_1",
                "Kategoria",
                "category",
                "Kategorie",
                "gazetki_Kategoria",
            ],
        )
        description = first_value(row, ["Opis HTML", "description_html", "description", "Opis"])
        images = first_value(row, ["images_urls", "Zdjęcie URL", "Obrazki", "Zdjecia", "Zdjęcia"])

        features = build_features(
            row,
            include_empty_features,
            manufacturer_data_by_producer,
            producer,
            feature_value_normalizer,
        )
        features = apply_category_attribute_knowledge(features, row, category, attribute_knowledge)
        features = filter_features_for_catalog(
            features,
            feature_value_normalizer,
            feature_review_rows,
            {
                "sku": export_sku,
                "ean": ean,
                "name": name,
            },
            feature_policy,
        )
        rows.append(
            {
                "product_id": first_value(row, ["product_id", "Baselinker ID"]),
                "name": name,
                "sku": export_sku,
                "ean": ean,
                "manufacturer_name": producer,
                "category": category,
                "description": description,
                "features": json.dumps(features, ensure_ascii=False, separators=(",", ":")),
                "images_urls": images,
            }
        )
    return rows


def apply_category_attribute_knowledge(
    features: dict[str, str],
    row: pd.Series,
    store_category: str,
    attribute_knowledge: dict[str, Any] | None,
) -> dict[str, str]:
    if not attribute_knowledge:
        return features
    rule = match_attribute_category_rule(row, store_category, attribute_knowledge, features)
    if not rule:
        return features

    # Reguly-placeholdery (bez listy atrybutow do opisu) nie przycinaja eksportu.
    has_attribute_lists = any(
        rule.get(role) for role in ("ordered_attributes", "required_description", "optional_description")
    )
    if not has_attribute_lists:
        return features
    ordered = category_attribute_order(rule)
    if not ordered:
        return features
    allowed_keys = {normalize_header(name) for name in ordered}
    # Atrybuty spoza macierzy, ktore maja zostac w eksporcie zawsze, gdy istnieja
    # (m.in. "Klasa energetyczna", "Kolor producenta" i "Napiecie [V]" dodane na zyczenie sklepu).
    requested = [
        "Producent",
        "Dane producenta",
        "Typ produktu",
        "Klasa energetyczna",
        "Kolor producenta",
        "Liczba źródeł światła",
        "Maksymalna moc źródła światła",
        "Napięcie [V]",
        *ordered,
    ]
    by_normalized = {normalize_header(name): (name, value) for name, value in features.items()}
    result: dict[str, str] = {}
    seen: set[str] = set()
    for requested_name in requested:
        normalized = normalize_header(requested_name)
        if normalized in seen:
            continue
        seen.add(normalized)
        item = by_normalized.get(normalized)
        if item:
            result[item[0]] = item[1]
    for name, value in features.items():
        normalized = normalize_header(name)
        if normalized in allowed_keys and name not in result:
            result[name] = value
    return result


# Atrybutowe schematy kategorii dotycza typu produktu, nie kategorii sklepu:
# akcesoria nie wchodza w schemat oprawy, a np. "Oprawa kanalowa" to nie plafon.
ACCESSORY_PRODUCT_TYPE_TERMS = (
    "ramka",
    "zasilacz",
    "siatka",
    "soczewka",
    "klips",
    "linka",
    "wspornik",
    "klosz",
    "zapinka",
    "czujnik",
    "konektor",
    "pilot",
    "uchwyt",
    "lacznik",
    "kabel",
    "oprawka",
    "adapter",
)

PRODUCT_TYPE_RULE_ALIASES = [
    # Oprawy kanalowe i sufitowe (punktowe, na GU10/E27) ida schematem opraw natynkowych.
    ("oprawa kanalowa", "oprawy natynkowe"),
    ("oprawa sufitowa", "oprawy natynkowe"),
    ("plafon", "plafony"),
    ("oprawa hermetyczna", "oprawy hermetyczne"),
    ("naswietlacz", "naswietlacze led"),
    ("panel", "panele led"),
    ("high bay", "high bay"),
    ("lampa wiszaca", "lampy wiszace zyrandole"),
    ("zyrandol", "lampy wiszace zyrandole"),
    ("oprawa natynkowa", "oprawy natynkowe"),
    ("oprawa podtynkowa", "oprawy podtynkowe"),
    ("downlight", "oprawy downlight"),
    ("lampa ogrodowa", "oprawy ogrodowe"),
    ("oprawa elewacyjna", "lampy elewacyjne kinkiety zewnetrzne"),
    ("kinkiet", "lampy elewacyjne kinkiety zewnetrzne"),
    ("lampa uliczna", "lampy uliczne"),
    ("latarka", "latarki"),
]


def is_accessory_product_type(product_type: str) -> bool:
    normalized = normalize_header(product_type)
    return bool(normalized) and any(term in normalized for term in ACCESSORY_PRODUCT_TYPE_TERMS)


def find_rule_by_category_name(attribute_knowledge: dict[str, Any], target: str) -> dict[str, Any] | None:
    for rule in attribute_knowledge.get("categories") or []:
        text = normalize_header(str(rule.get("category", "")))
        if text and all(part in text for part in target.split()):
            return rule
    return None


def match_attribute_category_rule(
    row: pd.Series,
    store_category: str,
    attribute_knowledge: dict[str, Any],
    features: dict[str, str] | None = None,
) -> dict[str, Any] | None:
    product_type = ""
    if features:
        product_type = compact_spaces(str(features.get("Typ produktu", "")))
    if not product_type:
        product_type = first_value(row, ["attr_typ", "Typ"])
    normalized_type = normalize_header(product_type)
    if normalized_type:
        if is_accessory_product_type(normalized_type):
            return None
        for needle, target in PRODUCT_TYPE_RULE_ALIASES:
            if needle in normalized_type:
                return find_rule_by_category_name(attribute_knowledge, target) if target else None

    source_category = first_value(row, ["Kategoria", "category", "Kategorie"])
    store_text = normalize_header(store_category)
    source_text = normalize_header(source_category)
    aliases = [
        ("oprawy natynkowe", "oprawy natynkowe"),
        ("oprawy podtynkowe", "oprawy podtynkowe"),
        ("panele led", "panele led"),
        ("high bay", "high bay"),
        ("pyloszczelne hermetyczne", "oprawy hermetyczne"),
        ("lampy hermetyczne", "oprawy hermetyczne"),
        ("naswietlacze", "naswietlacze led"),
        ("plafoniery oprawy sufitowe", "plafony"),
        ("plafony led", "plafony"),
        ("kinkiety zewnetrzne", "lampy elewacyjne kinkiety zewnetrzne"),
        ("lampy ogrodowe", "oprawy ogrodowe"),
    ]
    target = match_category_alias(store_text, aliases)
    if not target:
        target = match_category_alias(source_text, aliases)
    if not target:
        return None
    for rule in attribute_knowledge.get("categories") or []:
        if normalize_header(str(rule.get("category", ""))) == target:
            return rule
    return None


def match_category_alias(text: str, aliases: list[tuple[str, str]]) -> str:
    for needle, category in aliases:
        if all(part in text for part in needle.split()):
            return category
    return ""


def category_attribute_order(rule: dict[str, Any]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for role in ("ordered_attributes", "required_description", "filters", "optional_description"):
        for name in rule.get(role) or []:
            normalized = normalize_header(str(name))
            if normalized and normalized not in seen:
                seen.add(normalized)
                result.append(str(name))
    return result


def assign_main_images(
    df: pd.DataFrame,
    links_file: str | Path,
    links_sheet: str | int | None,
    key_column: str,
    main_image_column: str,
) -> tuple[pd.DataFrame, int]:
    links = read_products(links_file, sheet_name=links_sheet)
    link_key_column = find_column(links, key_column) or find_column(links, "Kod")
    link_image_column = find_column(links, main_image_column) or find_column(links, "Zdjecie glowne")
    if not link_key_column or not link_image_column:
        raise SystemExit("Nie znaleziono kolumn kodu lub zdjecia glownego w pliku linkow.")

    image_by_code = {
        normalize_code(row.get(link_key_column, "")): compact_spaces(str(row.get(link_image_column, "")))
        for _, row in links.iterrows()
        if normalize_code(row.get(link_key_column, "")) and compact_spaces(str(row.get(link_image_column, "")))
    }

    result = df.copy()
    matched = 0
    images: list[str] = []
    for _, row in result.iterrows():
        code = normalize_code(first_value(row, ["Kod Producenta", "Kod", "SKU", "sku"]))
        image = image_by_code.get(code, "")
        if image:
            matched += 1
        images.append(image)
    result["images_urls"] = images
    return result, matched


def find_column(df: pd.DataFrame, expected: str) -> str | None:
    normalized_expected = normalize_header(expected)
    for column in df.columns:
        if normalize_header(str(column)) == normalized_expected:
            return str(column)
    for column in df.columns:
        normalized = normalize_header(str(column))
        if normalized_expected and normalized_expected in normalized:
            return str(column)
    return None


def normalize_code(value: Any) -> str:
    value = compact_spaces(str(value or ""))
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    return value


def build_features(
    row: pd.Series,
    include_empty: bool,
    manufacturer_data_by_producer: dict[str, str] | None = None,
    producer: str = "",
    feature_value_normalizer: dict[str, dict[str, str]] | None = None,
) -> dict[str, str]:
    features: dict[str, str] = {}
    for target, source_columns in FEATURE_MAP.items():
        value = normalize_feature_value(first_value(row, source_columns), target)
        value = normalize_to_woo_feature_value(target, value, feature_value_normalizer)
        if value or include_empty:
            features[target] = value

    for column, value in row.items():
        column_name = str(column)
        if not column_name.startswith("Parametr: "):
            continue
        if pd.isna(value):
            continue
        raw_feature_name = column_name.replace("Parametr: ", "", 1)
        feature_name = normalize_feature_name(raw_feature_name)
        feature_value = normalize_feature_value(str(value), feature_name)
        feature_value = normalize_to_woo_feature_value(feature_name, feature_value, feature_value_normalizer)
        if feature_value or include_empty:
            features.setdefault(feature_name, feature_value)
    align_product_type_with_pipeline_title(features, row)
    reclassify_max_bulb_power_feature(features)
    drop_fixture_power_for_products_without_light_source(features)
    promote_panel_max_power_feature(features)
    derive_socket_count_feature(features, row)
    derive_light_source_feature(features)
    drop_socket_features_for_integrated_source(features)
    derive_light_color_from_cct(features)
    derive_power_supply_feature(features, row)
    derive_shape_from_dimensions(features)
    add_manufacturer_data_feature(features, manufacturer_data_by_producer, producer)
    remove_redundant_producer_color(features)
    return features


def dimension_in_mm(value: str) -> float | None:
    match = re.fullmatch(r"(\d+(?:[,.]\d+)?)\s*(?:mm)?", compact_spaces(value))
    return float(match.group(1).replace(",", ".")) if match else None


def derive_shape_from_dimensions(features: dict[str, str]) -> None:
    """Ksztalt z wymiarow: srednica -> Okragly, D == S -> Kwadrat, D != S -> Prostokatny."""
    if features.get("Kształt"):
        return
    if features.get("Średnica"):
        features["Kształt"] = "Okrągły"
        return
    length = dimension_in_mm(features.get("Długość", ""))
    width = dimension_in_mm(features.get("Szerokość", ""))
    if length is None or width is None:
        match = re.fullmatch(
            r"(\d+(?:[,.]\d+)?)x(\d+(?:[,.]\d+)?)",
            compact_spaces(features.get("Wymiary [mm]", "")).replace(" ", ""),
        )
        if not match:
            return
        length = float(match.group(1).replace(",", "."))
        width = float(match.group(2).replace(",", "."))
    features["Kształt"] = "Kwadrat" if abs(length - width) < 0.001 else "Prostokątny"


# Frazy typu na poczatku tytulu (najdluzsze najpierw) - tytuly z pipeline'u
# nazw sa kotwica klasyfikacji typu produktu.
TITLE_TYPE_PHRASES = [
    ("oprawa sufitowa punktowa", "Oprawa sufitowa punktowa"),
    ("oprawa hermetyczna", "Oprawa hermetyczna"),
    ("oprawa sufitowa", "Oprawa sufitowa"),
    ("oprawa kanalowa", "Oprawa kanałowa"),
    ("oprawa natynkowa", "Oprawa natynkowa"),
    ("oprawa podtynkowa", "Oprawa podtynkowa"),
    ("oprawa najazdowa", "Oprawa najazdowa"),
    ("panel led", "Panel LED"),
    ("plafon", "Plafon"),
    ("naswietlacz", "Naświetlacz LED"),
    ("lampa wiszaca", "Lampa wisząca"),
    ("lampa ogrodowa", "Lampa ogrodowa"),
    ("lampa uliczna", "Lampa uliczna"),
    ("ramka do paneli", "Ramka do paneli"),
    ("ramka", "Ramka"),
    ("zasilacz", "Zasilacz"),
]


def product_type_from_title(title: str) -> str:
    normalized = normalize_header(title)
    for phrase, type_name in TITLE_TYPE_PHRASES:
        if normalized.startswith(phrase):
            return type_name
    return ""


def product_types_agree(type_a: str, type_b: str) -> bool:
    a, b = normalize_header(type_a), normalize_header(type_b)
    return bool(a) and bool(b) and (a.startswith(b) or b.startswith(a))


def align_product_type_with_pipeline_title(features: dict[str, str], row: pd.Series) -> None:
    """Typ z nazwy wygenerowanej przez pipeline wygrywa nad kolumna "Typ" producenta.

    Producent potrafi nazwac oprawe sufitowa "plafoniera LED" - klasyfikacja
    z naszego pipeline'u nazw jest wiarygodniejsza.
    """
    pipeline_title = first_value(row, ["pipeline_title"])
    title_type = product_type_from_title(pipeline_title)
    if not title_type:
        return
    current = compact_spaces(str(features.get("Typ produktu", "")))
    if not current or not product_types_agree(current, title_type):
        features["Typ produktu"] = title_type


def normalize_light_source_value(value: str) -> str:
    """Zrodlo swiatla ma tylko dwie wartosci: Zintegrowane albo Nie zintegrowane."""
    normalized = comparable_feature_value(value)
    if not normalized:
        return ""
    if normalized.startswith("nie zintegrowan"):
        return "Nie zintegrowane"
    if "zintegrowan" in normalized or normalized in {"led", "led smd", "cob"}:
        return "Zintegrowane"
    replaceable_terms = (
        "wymienn", "t8", "gls", "cfl", "zarowk", "swietlowk",
        "par16", "par20", "par30", "par38", "mr16",
        "gu10", "e27", "e14", "g13", "gx53", "g9",
    )
    if any(term in normalized for term in replaceable_terms):
        return "Nie zintegrowane"
    return value


def derive_light_source_feature(features: dict[str, str]) -> None:
    if features.get("Źródło światła"):
        return
    if features.get("Źródło światła w komplecie") == "Tak":
        features["Źródło światła"] = "Zintegrowane"
    elif features.get("Trzonek") or features.get("Maksymalna moc źródła światła"):
        features["Źródło światła"] = "Nie zintegrowane"


def drop_socket_features_for_integrated_source(features: dict[str, str]) -> None:
    """Zintegrowane zrodlo nie ma trzonka ani liczby wymiennych zrodel."""
    if features.get("Źródło światła") != "Zintegrowane":
        return
    features.pop("Trzonek", None)
    features.pop("Liczba źródeł światła", None)


def derive_light_color_from_cct(features: dict[str, str]) -> None:
    """Barwa swiatla wynika z temperatury barwowej: Ciepla / Neutralna / Zimna / Zmienna."""
    if features.get("Barwa światła"):
        return
    cct = compact_spaces(features.get("Temperatura barwowa [K]", ""))
    numbers = [int(number) for number in re.findall(r"\d{4}", cct)]
    if not numbers:
        return
    if len(set(numbers)) > 1:
        features["Barwa światła"] = "Zmienna"
    elif numbers[0] < 3300:
        features["Barwa światła"] = "Ciepła"
    elif numbers[0] <= 5300:
        features["Barwa światła"] = "Neutralna"
    else:
        features["Barwa światła"] = "Zimna"


def derive_power_supply_feature(features: dict[str, str], row: pd.Series) -> None:
    """Zasilanie mowi, jak zasilany jest produkt: Solarne / Akumulator / Sieciowe."""
    if features.get("Zasilanie"):
        return
    name = first_value(
        row,
        ["new_title", "Nazwa B2C / SEO final", "Nazwa", "Nazwa Kanlux", "__working_title", "Title"],
    )
    text = comparable_feature_value(f"{name} {features.get('Typ produktu', '')}")
    voltage = comparable_feature_value(features.get("Napięcie [V]", ""))
    if "solarn" in text:
        features["Zasilanie"] = "Solarne"
    elif "akumulator" in text or "accu" in voltage:
        features["Zasilanie"] = "Akumulator"
    elif voltage.startswith("220-240"):
        features["Zasilanie"] = "Sieciowe"


def reclassify_max_bulb_power_feature(features: dict[str, str]) -> None:
    """Zapisy "max 20" / "3 x max 20" / "10 LED" w mocy to maksymalna moc zrodla, nie moc oprawy."""
    value = compact_spaces(features.get("Moc [W]", ""))
    match = re.fullmatch(
        r"(?:(\d{1,2})\s*x\s*)?(?:max\.?\s*)?(\d+(?:[,.]\d+)?)\s*W?\s*LED",
        value,
        flags=re.IGNORECASE,
    ) or re.fullmatch(
        r"(?:(\d{1,2})\s*x\s*)?max\.?\s*(\d+(?:[,.]\d+)?)\s*W?",
        value,
        flags=re.IGNORECASE,
    )
    if not match:
        return
    features.pop("Moc [W]", None)
    features.setdefault("Maksymalna moc źródła światła", match.group(2).replace(",", "."))
    if match.group(1):
        features.setdefault("Liczba źródeł światła", match.group(1))


def drop_fixture_power_for_products_without_light_source(features: dict[str, str]) -> None:
    """Produkt bez zrodla swiatla w srodku nie ma wlasnej mocy - zostaje tylko maksymalna moc zrodla."""
    if features.get("Źródło światła w komplecie") == "Tak":
        return
    has_max_source_power = bool(features.get("Maksymalna moc źródła światła"))
    replaceable_source = comparable_feature_value(features.get("Źródło światła", "")) in {"wymienne", "nie zintegrowane"}
    if has_max_source_power or replaceable_source:
        features.pop("Moc [W]", None)


def promote_panel_max_power_feature(features: dict[str, str]) -> None:
    """Panele bez zasilacza moga miec moc producenta tylko jako "Moc maksymalna [W]"."""
    product_type = comparable_feature_value(features.get("Typ produktu", ""))
    if product_type != "panel led" or features.get("Moc [W]"):
        return
    maximum_power = normalize_power_value(features.get("Maksymalna moc źródła światła", ""))
    if maximum_power:
        features["Moc [W]"] = maximum_power


def hermetic_fluorescent_title(title: str, features: dict[str, str]) -> str:
    """Oprawy hermetyczne z trzonkiem to oprawy na swietlowki - tytul ma to mowic wprost.

    Reguly tytulow dotycza typu produktu, nie kategorii: akcesoria (np. czujnik ruchu
    do oprawy hermetycznej) nie ida tym schematem.
    """
    if not features.get("Trzonek"):
        return title
    lowered = title.lower()
    if not lowered.startswith("oprawa hermetyczna"):
        return title
    if "świetlówk" in lowered or "swietlowk" in lowered:
        return title
    return re.sub(r"\bhermetyczna LED\b", "hermetyczna do świetlówek LED", title)


def derive_socket_count_feature(features: dict[str, str], row: pd.Series) -> None:
    """Rozbija zapisy typu "3xGU10": trzonek zostaje w Trzonek, krotnosc w Liczba zrodel swiatla."""
    if features.get("Liczba źródeł światła"):
        return
    socket = features.get("Trzonek", "")
    if not socket:
        return
    raw_socket = first_value(row, ["attr_gwint", "Gwint", "Trzonek", "Rodzaj gwintu"])
    match = re.match(r"^\s*(\d{1,2})\s*[xX]", raw_socket)
    if not match:
        name = first_value(
            row,
            ["new_title", "Nazwa B2C / SEO final", "Nazwa", "Nazwa Kanlux", "__working_title", "Title"],
        )
        match = re.search(rf"\b(\d{{1,2}})\s*[xX]\s*{re.escape(socket)}\b", name, flags=re.IGNORECASE)
    if match:
        features["Liczba źródeł światła"] = match.group(1)


def split_multi_color_feature(features: dict[str, str]) -> None:
    """Pelny kolor producenta zostaje w "Kolor producenta"; "Kolor" dostaje pierwszy kolor."""
    value = compact_spaces(features.get("Kolor", ""))
    parts = split_color_parts(value)
    if len(parts) > 1:
        features["Kolor"] = parts[0]
        features["Kolor producenta"] = " / ".join(parts)
        features.pop("Kolory", None)


def split_color_parts(value: str) -> list[str]:
    value = normalize_compound_color_value(value)
    parts = [compact_spaces(part) for part in re.split(r"\s*[/|]\s*", value) if compact_spaces(part)]
    if len(parts) > 1:
        return parts
    hyphen_parts = [compact_spaces(part) for part in re.split(r"\s*-\s*", value) if compact_spaces(part)]
    if len(hyphen_parts) == 2 and all(is_base_color_part(part) for part in hyphen_parts):
        return [normalize_producer_color_part(part) for part in hyphen_parts]
    return [value] if value else []


def is_base_color_part(value: str) -> bool:
    return normalize_color_adjective(value).lower() in BASE_COLOR_ALIASES


def add_manufacturer_data_feature(
    features: dict[str, str],
    manufacturer_data_by_producer: dict[str, str] | None,
    producer: str = "",
) -> None:
    if compact_spaces(features.get("Dane producenta", "")):
        features["Dane producenta"] = normalize_manufacturer_data_value(features["Dane producenta"])
        return
    if not manufacturer_data_by_producer:
        return

    data = manufacturer_data_for_producer(manufacturer_data_by_producer, producer, features.get("Producent", ""))
    if data:
        features["Dane producenta"] = data


def manufacturer_data_for_producer(
    manufacturer_data_by_producer: dict[str, str] | None,
    *producer_names: Any,
) -> str:
    if not manufacturer_data_by_producer:
        return ""
    for producer_name in producer_names:
        normalized = normalize_header(str(producer_name or ""))
        if not normalized:
            continue
        data = manufacturer_data_by_producer.get(normalized)
        if data:
            return data
    return ""


def write_baselinker_csv(rows: list[dict[str, str]], output: str | Path) -> None:
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=BASELINKER_COLUMNS, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(rows)


def default_feature_review_output(output: str | Path) -> str:
    output_path = Path(output)
    return str(output_path.with_name(f"{output_path.stem}_feature_review_required.xlsx"))


def write_feature_review_file(rows: list[dict[str, str]], output: str | Path) -> None:
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    unique_rows: list[dict[str, str]] = []
    seen: set[tuple[str, str, str, str]] = set()
    for row in rows:
        key = (
            row.get("feature_name", ""),
            row.get("feature_value", ""),
            row.get("reason", ""),
            row.get("sku", ""),
        )
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(row)
    if output_path.suffix.lower() in {".xlsx", ".xlsm"}:
        write_feature_review_xlsx(unique_rows, output_path)
        return
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FEATURE_REVIEW_COLUMNS)
        writer.writeheader()
        writer.writerows(unique_rows)


def write_feature_review_xlsx(rows: list[dict[str, str]], output: str | Path) -> None:
    df = pd.DataFrame(rows, columns=FEATURE_REVIEW_COLUMNS)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Do akceptacji", index=False)
        worksheet = writer.sheets["Do akceptacji"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        widths = {
            "A": 16,
            "B": 16,
            "C": 72,
            "D": 28,
            "E": 36,
            "F": 20,
        }
        for column, width in widths.items():
            worksheet.column_dimensions[column].width = width
        for cell in worksheet[1]:
            cell.style = "Headline 4"


def first_value(row: pd.Series, columns: list[str]) -> str:
    for column in columns:
        if not column or column not in row:
            continue
        value = row.get(column, "")
        if pd.isna(value):
            continue
        if not is_blank(value):
            return compact_spaces(str(value))
    return ""


def normalize_feature_name(value: str) -> str:
    value = compact_spaces(value)
    normalized = normalize_header(value)
    return FEATURE_NAME_ALIASES.get(normalized, value)


def filter_features_for_catalog(
    features: dict[str, str],
    feature_value_normalizer: dict[str, dict[str, str]] | None,
    review_rows: list[dict[str, str]] | None = None,
    context: dict[str, str] | None = None,
    feature_policy: dict[str, Any] | None = None,
) -> dict[str, str]:
    filtered: dict[str, str] = {}
    if feature_value_normalizer is None:
        feature_value_normalizer = {}
    policy = feature_policy or default_supplier_feature_policy()
    context = context or {}
    if features.get("Typ produktu"):
        context = {**context, "typ_produktu": features.get("Typ produktu", "")}
    for feature_name, value in features.items():
        value = compact_spaces(str(value or ""))
        if not value:
            continue
        if feature_name in policy["excluded_features"]:
            continue
        if drop_feature_for_product_context(feature_name, value, context, policy):
            continue
        if feature_name == "Kolor":
            if add_color_features(filtered, value, feature_value_normalizer, policy):
                continue
        if feature_name in policy["drop_unknown_features"]:
            continue
        known_values = feature_value_normalizer.get(feature_name)
        if manually_approved_feature_value(feature_name, value, policy):
            filtered[feature_name] = value
            continue
        if valid_flexible_feature_value(feature_name, value):
            filtered[feature_name] = value
            continue
        if known_values is None:
            remember_feature_review(review_rows, context, feature_name, value, "UNKNOWN_ATTRIBUTE")
            continue
        if not feature_value_is_known(value, known_values):
            remember_feature_review(review_rows, context, feature_name, value, "UNKNOWN_VALUE")
            continue
        filtered[feature_name] = value
    return filtered


def add_color_features(
    filtered: dict[str, str],
    value: str,
    feature_value_normalizer: dict[str, dict[str, str]],
    feature_policy: dict[str, Any] | None = None,
) -> bool:
    policy = feature_policy or default_supplier_feature_policy()
    if value in policy["non_base_colors_as_producer_color"]:
        producer_values = feature_value_normalizer.get("Kolor producenta", {})
        filtered["Kolor"] = "Brązowy" if value in {"Dąb sonoma", "Wenge"} else base_color_from_producer_color(value)
        filtered["Kolor producenta"] = canonical_or_original_feature_value(value, producer_values)
        return True
    parts = split_color_parts(value)
    if len(parts) > 1:
        color_values = feature_value_normalizer.get("Kolor", {})
        producer_values = feature_value_normalizer.get("Kolor producenta", {})
        filtered["Kolor"] = canonical_or_original_feature_value(parts[0], color_values)
        filtered["Kolor producenta"] = " / ".join(
            canonical_or_original_feature_value(part, producer_values) for part in parts
        )
        return True
    return False


def base_color_from_producer_color(value: str) -> str:
    normalized = comparable_feature_value(value)
    if normalized in {"dab sonoma", "wenge", "zloto dab", "zloty dab", "złoty dab"}:
        return "Brązowy"
    if normalized in {"nikiel satynowy", "stal inox", "inox"}:
        return "Srebrny"
    if normalized in {"bialy mat", "biały mat"}:
        return "Biały"
    if normalized == "czarny mat":
        return "Czarny"
    return value


def canonical_or_original_feature_value(value: str, known_values: dict[str, str]) -> str:
    for key in feature_value_lookup_keys(value):
        canonical = known_values.get(key)
        if canonical:
            return canonical
    return value


def drop_feature_for_product_context(
    feature_name: str,
    value: str,
    context: dict[str, str],
    feature_policy: dict[str, Any] | None = None,
) -> bool:
    policy = feature_policy or default_supplier_feature_policy()
    name_key = comparable_feature_value(context.get("name", ""))
    product_type = context.get("typ_produktu", "")
    if feature_name == "Moc [W]" and product_type in policy["accessory_product_types_without_own_power"]:
        return True
    if feature_name == "Wymiary [mm]" and re.search(r"\bmm2\b", value, flags=re.IGNORECASE):
        return True
    if feature_name == "Seria" and ("oprawa high bay" in name_key or "high bay" in name_key):
        return True
    if feature_name == "Kąt świecenia [°]" and (
        "oprawa hermetyczna" in name_key
        or "panel led" in name_key
        or "panel " in name_key
        or "plafon" in name_key
    ):
        return True
    if feature_name == "Zastosowanie" and not manually_approved_feature_value(feature_name, value, policy):
        return True
    if feature_name == "Kolor" and value in policy["non_base_colors_as_producer_color"]:
        return False
    if feature_name == "Seria" and (
        "klips" in name_key
        or "linka" in name_key
        or "klosz" in name_key
        or "soczewka" in name_key
        or "zapinka" in name_key
    ):
        return True
    return False


def manually_approved_feature_value(
    feature_name: str,
    value: str,
    feature_policy: dict[str, Any] | None = None,
) -> bool:
    policy = feature_policy or default_supplier_feature_policy()
    approved_values = policy["approved_feature_values"].get(feature_name, set())
    return value in approved_values


def valid_flexible_feature_value(feature_name: str, value: str) -> bool:
    value = compact_spaces(str(value or ""))
    if not value:
        return False
    if feature_name == "Wymiary [mm]":
        return bool(
            re.fullmatch(r"\d+(?:[,.]\d+)?(?:x\d+(?:[,.]\d+)?){0,2}", value)
            or valid_dimension_like_value(value)
        )
    if feature_name in {"Długość", "Szerokość", "Wysokość", "Średnica", "Głębokość"}:
        return bool(re.fullmatch(r"\d+(?:[,.]\d+)?(?:x\d+(?:[,.]\d+)?){0,2}(?:\s*(?:mm|cm|m))?", value, flags=re.IGNORECASE))
    if feature_name in {"Moc [W]", "Strumień świetlny [lm]", "Maksymalna moc źródła światła"}:
        return valid_number_or_range(value)
    if feature_name == "Klasa energetyczna":
        return bool(re.fullmatch(r"[A-G]\+{0,3}", value))
    if feature_name == "Liczba źródeł światła":
        return bool(re.fullmatch(r"\d{1,2}", value))
    if feature_name == "Źródło światła":
        return value in {"Zintegrowane", "Nie zintegrowane"}
    if feature_name == "Temperatura barwowa [K]":
        return bool(re.fullmatch(r"\d+(?:[/-]\d+){0,3}|RGB", value, flags=re.IGNORECASE))
    if feature_name == "Napięcie [V]":
        return bool(re.fullmatch(r"\d+(?:[,.]\d+)?(?:-\d+(?:[,.]\d+)?){0,1}(?:\s(?:AC|DC))?", value))
    if feature_name == "Gwarancja":
        return bool(re.fullmatch(r"\d+\s*(?:lat|lata|rok|roku|miesiecy|miesięcy|mies\.?)", value, flags=re.IGNORECASE))
    if feature_name == "Stopień ochrony [IK]":
        return bool(re.fullmatch(r"IK\s*\d{2}", value, flags=re.IGNORECASE))
    if feature_name == "Kąt świecenia [°]":
        return bool(re.fullmatch(r"\d+(?:[,.]\d+)?\s*°?", value))
    if feature_name == "Skuteczność świetlna [lm/W]":
        return bool(re.fullmatch(r"\d+(?:[,.]\d+)?\s*(?:lm/W)?", value, flags=re.IGNORECASE))
    if feature_name == "Wskaźnik oddawania barw":
        return bool(re.fullmatch(r"\d{2,3}", value))
    if feature_name == "Trwałość [h]":
        return bool(re.fullmatch(r"\d+", value))
    if feature_name == "Wskaźnik olśnienia [UGR]":
        return bool(re.fullmatch(r"(?:<|≤)?\d{1,2}", value))
    if feature_name == "Współpraca ze ściemniaczem":
        return value in {"Tak", "Nie", "DALI", "1-10V", "0-10V"}
    if feature_name == "Źródło światła w komplecie":
        return value in {"Tak", "Nie"}
    if feature_name == "Opakowanie":
        return bool(re.fullmatch(r"\d+\s*(?:szt\.?|sztuk|sztuki)", value, flags=re.IGNORECASE))
    return False


def valid_dimension_like_value(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"\d+(?:[,.]\d+)?(?:x\d+(?:[,.]\d+)?){0,2}\s*(?:mm|cm|m)",
            value,
            flags=re.IGNORECASE,
        )
        or bool(re.fullmatch(r"\d+(?:x\d+){1,2}\s*mm2", value, flags=re.IGNORECASE))
    )


def valid_number_or_range(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:[,.]\d+)?(?:-\d+(?:[,.]\d+)?){0,1}", value))


def feature_value_is_known(value: str, known_values: dict[str, str]) -> bool:
    if "|" in value:
        return all(feature_value_is_known(part, known_values) for part in value.split("|") if compact_spaces(part))
    for key in feature_value_lookup_keys(value):
        if key in known_values:
            return True
    return False


def remember_feature_review(
    review_rows: list[dict[str, str]] | None,
    context: dict[str, str],
    feature_name: str,
    feature_value: str,
    reason: str,
) -> None:
    if review_rows is None:
        return
    review_rows.append(
        {
            "sku": context.get("sku", ""),
            "ean": context.get("ean", ""),
            "name": context.get("name", ""),
            "feature_name": feature_name,
            "feature_value": feature_value,
            "reason": reason,
        }
    )


def build_feature_value_normalizer(catalog_knowledge: dict[str, Any]) -> dict[str, dict[str, str]]:
    normalizer: dict[str, dict[str, tuple[str, int]]] = {}

    def remember(feature_name: str, value: Any, count: int = 0) -> None:
        feature_name = normalize_feature_name(str(feature_name or ""))
        value_text = clean_catalog_feature_value(value)
        if not feature_name or not value_text:
            return
        feature_values = normalizer.setdefault(feature_name, {})
        for key in feature_value_lookup_keys(value_text):
            previous = feature_values.get(key)
            if not previous or count > previous[1]:
                feature_values[key] = (value_text, count)

    for row in catalog_knowledge.get("colors") or []:
        if isinstance(row, dict):
            remember("Kolor", row.get("color", ""), int(row.get("count") or 0))
            remember("Kolor producenta", row.get("color", ""), int(row.get("count") or 0))

    for row in catalog_knowledge.get("ip_values") or []:
        if isinstance(row, dict):
            remember("Stopień ochrony [IP]", row.get("ip", ""), int(row.get("count") or 0))

    for row in catalog_knowledge.get("top_attributes") or []:
        if not isinstance(row, dict):
            continue
        feature_name = row.get("attribute", "")
        for value, count in parse_top_attribute_values(row.get("top_values", "")):
            remember(feature_name, value, count)

    for row in catalog_knowledge.get("top_attribute_values_by_category") or []:
        if not isinstance(row, dict):
            continue
        remember(row.get("attribute", ""), row.get("value", ""), int(row.get("count") or 0))

    for data in build_manufacturer_data_by_producer(catalog_knowledge).values():
        remember("Dane producenta", data, 0)

    return {
        feature_name: {key: value for key, (value, _) in values.items()}
        for feature_name, values in normalizer.items()
    }


def parse_top_attribute_values(value: Any) -> list[tuple[str, int]]:
    text = compact_spaces(str(value or ""))
    if not text:
        return []
    result: list[tuple[str, int]] = []
    for part in re.split(r"\s*;\s*", text):
        part = compact_spaces(part)
        if not part:
            continue
        match = re.fullmatch(r"(.+?)\s+\((\d+)\)", part)
        if match:
            result.append((match.group(1), int(match.group(2))))
        else:
            result.append((part, 0))
    return result


def normalize_to_woo_feature_value(
    feature_name: str,
    value: str,
    feature_value_normalizer: dict[str, dict[str, str]] | None = None,
) -> str:
    value = compact_spaces(str(value or ""))
    if not value:
        return ""

    value = normalize_feature_value_before_lookup(feature_name, value)
    if not feature_value_normalizer:
        return value

    known_values = feature_value_normalizer.get(normalize_feature_name(feature_name), {})
    for key in feature_value_lookup_keys(value):
        canonical = known_values.get(key)
        if canonical:
            return canonical
    if feature_name == "Stopień ochrony [IP]":
        split_ip = normalize_compound_ip_to_multiple_values(value, known_values)
        if split_ip:
            return split_ip
    return value


def normalize_feature_value_before_lookup(feature_name: str, value: str) -> str:
    if feature_name == "Typ produktu":
        return normalize_product_type_feature_value(value)
    if feature_name == "Seria":
        return normalize_series_feature_value(value)
    if feature_name == "Wymiary [mm]":
        return normalize_dimensions_to_mm(value)
    if feature_name in {"Długość", "Szerokość", "Wysokość", "Średnica", "Głębokość"}:
        return normalize_dimension_to_mm_with_unit(value)
    if feature_name == "Moc [W]":
        return normalize_power_value(value)
    if feature_name == "Strumień świetlny [lm]":
        return normalize_luminous_flux_value(value)
    if feature_name == "Kształt":
        return normalize_shape_value(value)
    if feature_name == "Barwa światła":
        return normalize_light_color_value(value)
    if feature_name == "Kąt świecenia [°]":
        return normalize_angle_value(value)
    if feature_name == "Skuteczność świetlna [lm/W]":
        return normalize_luminous_efficacy_value(value)
    if feature_name in {"Kolor", "Kolor producenta"}:
        return normalize_feature_color_value(feature_name, value)
    if feature_name == "Czujnik ruchu":
        return normalize_motion_sensor_value(value)
    if feature_name == "Zastosowanie":
        return normalize_application_value(value)
    if feature_name == "Opakowanie":
        return normalize_package_value(value)
    if feature_name == "Stopień ochrony [IK]":
        return normalize_ik_value(value)
    if feature_name == "Źródło światła w komplecie":
        normalized = normalize_header(value)
        return "Tak" if normalized in {"1", "tak", "yes", "true"} else "Nie" if normalized in {"0", "nie", "no", "false"} else ""
    return value


def normalize_product_type_feature_value(value: str) -> str:
    normalized = comparable_feature_value(value)
    if "high bay" in normalized:
        return "Oprawa High Bay"
    if normalized == "mah":
        return "Czujnik ruchu"
    if normalized in {"iq-led"}:
        return "Konektor"
    if "naswietlacz" in normalized:
        return "Naświetlacz LED"
    if normalized in {"adtr", "adtr pt"}:
        return "Ramka do paneli"
    if normalized == "siatka ochronna":
        return "Siatka ochronna"
    if normalized == "linka":
        return "Linka do podwieszenia"
    if normalized == "wspornik montazowy":
        return "Wspornik montażowy"
    if normalized in {"oprawa punktowa bord", "oprawa punktowa bord one", "oprawa sufitowa natynkowa"}:
        return "Oprawa sufitowa"
    if normalized in {"oprawa hermetyczna led", "oprawa hermetyczna led duzej mocy", "oprawa hermetyczna led pro", "oprawa hermetyczna na t8", "oprawa stropowa al"}:
        return "Oprawa hermetyczna"
    if normalized in {"pilot"}:
        return "Pilot do oprawy"
    if normalized in {"uchwyt montazu natynkowego"}:
        return "Uchwyt"
    if normalized in {"uchwyt"}:
        return "Klips"
    if normalized in {"phlox c", "phloxc"}:
        return "Oprawa sufitowa"
    if normalized in {"toleo"}:
        # TOLEO DTL to natynkowa oprawa punktowa (downlight tube), nie lampa wiszaca.
        return "Oprawa sufitowa"
    if "panel led" in normalized:
        return "Panel LED"
    if "plafon" in normalized:
        return "Plafon"
    if normalized == "zasilacz":
        return "Zasilacz"
    return capitalize_first(value)


def normalize_series_feature_value(value: str) -> str:
    normalized = comparable_feature_value(value)
    if normalized.startswith("al55"):
        return "ALIN"
    if normalized.startswith("blingo"):
        return "BLINGO"
    if normalized in {"varso led", "varso"}:
        return "VARSO"
    if normalized in {"barev bl", "barev eco"}:
        return "BAREV"
    if normalized in {"iq-led fl", "iq-led"}:
        return "IQ-LED"
    if normalized.startswith("fl agor"):
        return "AGOR"
    if normalized.startswith("dicht"):
        return "DICHT"
    if normalized.startswith("adtr"):
        return "ADTR"
    if normalized.startswith("fogler"):
        return "FOGLER"
    if normalized.startswith("mah"):
        return "MAH"
    if normalized.startswith("tp"):
        return "TP"
    if normalized.startswith("bord"):
        return "BORD"
    if normalized.startswith("pires"):
        return "PIRES"
    return value


def clean_catalog_feature_value(value: Any) -> str:
    return compact_spaces(str(value or "").replace("\\", "").strip())


def feature_value_lookup_keys(value: str) -> list[str]:
    key = comparable_feature_value(value)
    if not key:
        return []
    candidates = {key}
    candidates.add(key.replace(" ", ""))
    if "°" in key:
        candidates.add(key.replace("°", ""))
    if "," in key:
        candidates.add(key.replace(",", "."))
    if "." in key:
        candidates.add(key.replace(".", ","))
    return [candidate for candidate in candidates if candidate]


def comparable_feature_value(value: str) -> str:
    value = clean_catalog_feature_value(value)
    value = value.replace("×", "x").lower()
    value = "".join(
        char for char in unicodedata.normalize("NFKD", value)
        if not unicodedata.combining(char)
    )
    value = re.sub(r"\s*/\s*", "/", value)
    value = re.sub(r"\s*x\s*", "x", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_feature_value(value: str, feature_name: str = "") -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    value = re.sub(r"\s*\|\s*", "|", value)
    if feature_name == "Dane producenta":
        return normalize_manufacturer_data_value(value)

    if feature_name == "Stopień ochrony [IP]":
        return normalize_ip_value(value)
    if feature_name == "Moc [W]":
        return normalize_power_value(value)
    if feature_name == "Temperatura barwowa [K]":
        return normalize_temperature_value(value)
    if feature_name == "Strumień świetlny [lm]":
        return strip_unit(value, "lm")
    if feature_name == "Napięcie [V]":
        return normalize_voltage_value(value)
    if feature_name == "Barwa światła":
        return normalize_light_color_value(value)
    if feature_name == "Trzonek":
        return normalize_socket_value(value)
    if feature_name == "Kąt świecenia [°]":
        return normalize_angle_value(value)
    if feature_name == "Wymiary [mm]":
        return normalize_dimensions_to_mm(value)
    if feature_name in {"Długość", "Szerokość", "Wysokość", "Średnica", "Głębokość"}:
        return normalize_dimension_to_mm_with_unit(value)
    if feature_name == "Maksymalna moc źródła światła":
        value = re.sub(r"^(?:max\.?|maks(?:ymalnie)?|do)\s+", "", value, flags=re.IGNORECASE)
        return strip_unit(value, "W")
    if feature_name == "Klasa energetyczna":
        return value.upper() if re.fullmatch(r"[a-gA-G]\+{0,3}", value.strip()) else value
    if feature_name == "Liczba źródeł światła":
        return value if re.fullmatch(r"\d{1,2}", value) else ""
    if feature_name == "Źródło światła":
        return normalize_light_source_value(value)
    if feature_name == "Skuteczność świetlna [lm/W]":
        return normalize_luminous_efficacy_value(value)
    if feature_name == "Wskaźnik olśnienia [UGR]":
        return compact_spaces(value).replace(" ", "").replace("<=", "≤")
    if feature_name == "Kolor":
        return normalize_color_value(value)
    if feature_name == "Kolor producenta":
        return normalize_producer_color_value(value)
    if feature_name == "Materiał":
        return normalize_material_value(value)
    if feature_name == "Źródło światła w komplecie":
        normalized = normalize_header(value)
        return "Tak" if normalized in {"1", "tak", "yes", "true"} else "Nie" if normalized in {"0", "nie", "no", "false"} else ""
    return value


def normalize_ip_value(value: str) -> str:
    normalized = value.upper().replace(" ", "")
    if normalized == "IP12":
        return "IP 20"
    return re.sub(r"IP(\d+)", r"IP \1", normalized)


def normalize_compound_ip_to_multiple_values(value: str, known_values: dict[str, str]) -> str:
    parts = [compact_spaces(part) for part in re.split(r"\s*/\s*", value) if compact_spaces(part)]
    if len(parts) < 2:
        return ""
    canonical_parts: list[str] = []
    for part in parts:
        canonical = ""
        for key in feature_value_lookup_keys(part):
            canonical = known_values.get(key, "")
            if canonical:
                break
        if not canonical:
            return ""
        canonical_parts.append(canonical)
    return "|".join(canonical_parts)


def strip_unit(value: str, unit: str) -> str:
    value = normalize_decimal_separator(value)
    value = re.sub(rf"\s*{re.escape(unit)}\b", "", value, flags=re.IGNORECASE)
    return compact_spaces(value)


def normalize_power_value(value: str) -> str:
    value = compact_spaces(normalize_decimal_separator(value))
    without_prefix = re.sub(r"^(?:max\.?|maks(?:ymalnie)?|do)\s+", "", value, flags=re.IGNORECASE)
    range_match = re.fullmatch(
        r"(\d+(?:[,.]\d+)?)\s*(?:W\s*)?(?:-|/|–|—)\s*(\d+(?:[,.]\d+)?)\s*W?",
        without_prefix,
        flags=re.IGNORECASE,
    )
    if range_match:
        numbers = [float(group.replace(",", ".")) for group in range_match.groups()]
        maximum = max(numbers)
        return str(int(maximum)) if maximum.is_integer() else str(maximum).rstrip("0").rstrip(".")
    if re.match(r"^\d", without_prefix):
        return strip_unit(without_prefix, "W")
    return value


def normalize_luminous_flux_value(value: str) -> str:
    value = strip_unit(value, "lm")
    return compact_spaces(value)


def normalize_luminous_efficacy_value(value: str) -> str:
    value = normalize_decimal_separator(value)
    value = re.sub(r"\s*lm\s*/\s*W\b", "", value, flags=re.IGNORECASE)
    return compact_spaces(value)


def normalize_dimensions_to_mm(value: str) -> str:
    normalized = normalize_dimension_value(value)
    match = re.fullmatch(
        r"(\d+(?:[,.]\d+)?(?:x\d+(?:[,.]\d+)?){0,2})\s*(mm|cm|m)",
        normalized,
        flags=re.IGNORECASE,
    )
    if not match:
        return normalized
    factor = {"mm": 1.0, "cm": 10.0, "m": 1000.0}[match.group(2).lower()]
    values = [float(part.replace(",", ".")) * factor for part in match.group(1).split("x")]
    formatted = [str(int(number)) if number.is_integer() else str(number).rstrip("0").rstrip(".") for number in values]
    return " x ".join(formatted)


def normalize_dimension_value(value: str) -> str:
    value = normalize_decimal_separator(value).replace("×", "x")
    value = re.sub(r"\s*x\s*", "x", value, flags=re.IGNORECASE)
    value = re.sub(r"\s*mm2\b", " mm2", value, flags=re.IGNORECASE)
    value = re.sub(r"\s*(mm|cm|m)\b", r" \1", value, flags=re.IGNORECASE)
    return compact_spaces(value)


def normalize_dimension_to_mm_with_unit(value: str) -> str:
    # Pojedyncze wymiary sklep zapisuje w mm z jednostka po spacji, np. "600 mm"
    # (tak wygladaja wartosci "Wysokosc"/"Szerokosc" w Woo catalog knowledge).
    normalized = normalize_dimension_value(value)
    match = re.fullmatch(r"(\d+(?:[,.]\d+)?)\s*(mm|cm|m)", normalized, flags=re.IGNORECASE)
    if not match:
        return normalized
    factor = {"mm": 1.0, "cm": 10.0, "m": 1000.0}[match.group(2).lower()]
    number = float(match.group(1).replace(",", ".")) * factor
    formatted = str(int(number)) if number.is_integer() else str(number).rstrip("0").rstrip(".")
    return f"{formatted} mm"


def normalize_shape_value(value: str) -> str:
    aliases = {
        "okragla": "Okrągły",
        "okrągła": "Okrągły",
        "okragly": "Okrągły",
        "okrągły": "Okrągły",
        "kwadratowa": "Kwadrat",
        "kwadratowy": "Kwadrat",
        "kwadrat": "Kwadrat",
        "prostokatna": "Prostokątny",
        "prostokątna": "Prostokątny",
        "prostokatny": "Prostokątny",
        "prostokątny": "Prostokątny",
    }
    return aliases.get(value.lower(), capitalize_first(value))


def normalize_temperature_value(value: str) -> str:
    value = value.replace(" ", "")
    if re.fullmatch(r"[0-9/]+K?", value, flags=re.IGNORECASE):
        return value.replace("K", "").replace("k", "")
    return re.sub(r"(?<=\d)\s*K\b", "", value, flags=re.IGNORECASE)


def normalize_voltage_value(value: str) -> str:
    # Sklep chce napiecie z rodzajem pradu, np. "220-240 AC".
    if value == "MLS":
        return ""
    current_type = ""
    if re.search(r"\bAC\b", value, flags=re.IGNORECASE):
        current_type = "AC"
    elif re.search(r"\bDC\b", value, flags=re.IGNORECASE):
        current_type = "DC"
    value = re.sub(r"\s*(?:AC|DC|V)\b", "", value, flags=re.IGNORECASE)
    value = compact_spaces(normalize_decimal_separator(value))
    if not value:
        return ""
    if not current_type and re.fullmatch(r"(?:220-240|230|220|240|110-240)", value):
        current_type = "AC"
    return f"{value} {current_type}" if current_type else value


def normalize_light_color_value(value: str) -> str:
    normalized = comparable_feature_value(value)
    mapping = {
        "Neutralna (3500-4500K)": "Neutralna",
        "Zimna (≥5000K)": "Zimna",
        "Ciepła (≤3000K)": "Ciepła",
        "Zmienna CCT": "Zmienna",
        "neutralna": "Neutralna",
        "neutralne": "Neutralna",
        "ciepla": "Ciepła",
        "cieple": "Ciepła",
        "cieplobiala": "Ciepła",
        "cieplobiale": "Ciepła",
        "zimna": "Zimna",
        "zimne": "Zimna",
        "chlodna": "Zimna",
        "chlodne": "Zimna",
        "chlodnobiala": "Zimna",
        "chlodnobiale": "Zimna",
        "zmienna": "Zmienna",
        "zmienna cct": "Zmienna",
    }
    return mapping.get(value, mapping.get(normalized, value))


def normalize_socket_value(value: str) -> str:
    normalized = value.upper().replace(" ", "")
    match = re.fullmatch(r"\d+X(.+)", normalized)
    return match.group(1) if match else normalized


def normalize_angle_value(value: str) -> str:
    return compact_spaces(str(value or "").replace("°", ""))


def normalize_color_value(value: str) -> str:
    aliases = {
        "bialy": "Biały",
        "biały": "Biały",
    }
    return aliases.get(value, normalize_compound_color_value(value))


def normalize_feature_color_value(feature_name: str, value: str) -> str:
    normalized = comparable_feature_value(value)
    if feature_name == "Kolor" and normalized in {"stal inox", "inox"}:
        return "Srebrny"
    if feature_name == "Kolor producenta" and normalized in {"stal inox", "inox"}:
        return "STAL INOX"
    if feature_name == "Kolor" and normalized in {"bialy mat", "biały mat"}:
        return "Biały"
    if feature_name == "Kolor" and normalized == "czarny mat":
        return "Czarny"
    if feature_name == "Kolor" and normalized in {"zloto-brazowy", "złoto-brazowy"}:
        return "Złoty"
    if feature_name == "Kolor producenta" and normalized in {"zloto-brazowy", "złoto-brazowy"}:
        return "Złoto-brązowy"
    return normalize_compound_color_value(value)


def normalize_compound_color_value(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    value = re.sub(r"\s*\|\s*", " / ", value)
    value = re.sub(r"\s*/\s*", " / ", value)
    parts = [compact_spaces(part) for part in value.split(" / ") if compact_spaces(part)]
    if len(parts) > 1:
        return " / ".join(normalize_producer_color_part(part) for part in parts)
    return normalize_producer_color_part(value)


def normalize_producer_color_value(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    value = re.sub(r"\s*\|\s*", " / ", value)
    value = re.sub(r"\s*/\s*", " / ", value)
    parts = [compact_spaces(part) for part in value.split(" / ") if compact_spaces(part)]
    if len(parts) > 1:
        return " / ".join(normalize_producer_color_part(part) for part in parts)
    return normalize_producer_color_part(value)


BASE_COLOR_ALIASES = {
    "bialy": "Biały",
    "biały": "Biały",
    "czarny": "Czarny",
    "szary": "Szary",
    "srebrny": "Srebrny",
    "grafitowy": "Grafitowy",
    "brazowy": "Brązowy",
    "brązowy": "Brązowy",
    "bezowy": "Beżowy",
    "beżowy": "Beżowy",
    "zloty": "Złoty",
    "złoty": "Złoty",
}


def normalize_producer_color_part(value: str) -> str:
    aliases = {
        **BASE_COLOR_ALIASES,
        "bialy mat": "Biały mat",
        "biały mat": "Biały mat",
        "czarny mat": "Czarny mat",
        "nikiel satynowy": "Nikiel satynowy",
        "inox": "STAL INOX",
        "zloty dab": "Złoty dąb",
        "złoty dąb": "Złoty dąb",
        "dab sonoma": "Dąb sonoma",
        "dąb sonoma": "Dąb sonoma",
        "wenge": "Wenge",
        "drewno": "Drewno",
    }
    normalized = normalize_color_adjective(value).lower()
    return aliases.get(normalized, aliases.get(value.lower(), capitalize_first(value)))


def normalize_color_adjective(value: str) -> str:
    normalized = comparable_feature_value(value)
    adjective_forms = {
        "biala": "bialy",
        "biale": "bialy",
        "bialo": "bialy",
        "czarna": "czarny",
        "czarne": "czarny",
        "czarno": "czarny",
        "szara": "szary",
        "szare": "szary",
        "srebrna": "srebrny",
        "srebrne": "srebrny",
        "grafitowa": "grafitowy",
        "grafitowe": "grafitowy",
        "brazowa": "brazowy",
        "brazowe": "brazowy",
        "bezowa": "bezowy",
        "bezowe": "bezowy",
        "zlota": "zloty",
        "zlote": "zloty",
    }
    return adjective_forms.get(normalized, normalized)


def remove_redundant_producer_color(features: dict[str, str]) -> None:
    return


def colors_are_equivalent(color: str, producer_color: str) -> bool:
    color_values = comparable_color_values(color)
    producer_color_values = comparable_color_values(producer_color)
    return bool(color_values) and color_values == producer_color_values


def comparable_color_values(value: str) -> set[str]:
    base_colors = {
        "bialy": "Biały",
        "biały": "Biały",
        "czarny": "Czarny",
        "szary": "Szary",
        "srebrny": "Srebrny",
        "zloty": "Złoty",
        "złoty": "Złoty",
        "brazowy": "Brązowy",
        "brązowy": "Brązowy",
        "grafitowy": "Grafitowy",
        "bezowy": "Beżowy",
        "beżowy": "Beżowy",
        "kaszmir": "Kaszmir",
        "zielony": "Zielony",
        "czerwony": "Czerwony",
        "pomaranczowy": "Pomarańczowy",
        "pomarańczowy": "Pomarańczowy",
        "antracyt": "Antracyt",
        "niebieski": "Niebieski",
        "transparentny": "Transparentny",
        "zolty": "Żółty",
        "żółty": "Żółty",
        "rozowy": "Różowy",
        "różowy": "Różowy",
        "fioletowy": "Fioletowy",
    }
    result: set[str] = set()
    for part in re.split(r"\s*(?:\||/)\s*", str(value or "")):
        normalized = part.lower().strip()
        if normalized in base_colors:
            result.add(base_colors[normalized])
    return result


def normalize_material_value(value: str) -> str:
    aliases = {
        "1": "",
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
    }
    # Kilka materialow w jednym polu to osobne wartosci atrybutu.
    parts = [compact_spaces(part) for part in re.split(r"\s*[,|]\s*", value) if compact_spaces(part)]
    if len(parts) > 1:
        normalized_parts = [
            aliases.get(part, aliases.get(part.lower(), capitalize_first(part))) for part in parts
        ]
        return "|".join(unique_values(normalized_parts))
    return aliases.get(value, capitalize_first(value))


def unique_values(values: list[str]) -> list[str]:
    result: list[str] = []
    for value in values:
        if value and value not in result:
            result.append(value)
    return result


def normalize_motion_sensor_value(value: str) -> str:
    normalized = comparable_feature_value(value)
    if "czujnik" in normalized:
        return "Tak"
    if normalized in {"tak", "yes", "1"}:
        return "Tak"
    if normalized in {"nie", "no", "0", "brak"}:
        return "Nie"
    return capitalize_first(value)


def normalize_application_value(value: str) -> str:
    normalized = comparable_feature_value(value)
    if normalized == "do paneli":
        return "Do paneli"
    return capitalize_first(value)


def normalize_package_value(value: str) -> str:
    value = compact_spaces(value.lower())
    match = re.fullmatch(r"(\d+)\s*(?:szt\.?|sztuk|sztuki)", value, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)} szt."
    return compact_spaces(value)


def normalize_ik_value(value: str) -> str:
    value = compact_spaces(value.upper().replace(" ", ""))
    return re.sub(r"^IK(\d+)$", r"IK \1", value)


def normalize_manufacturer_data_value(value: str) -> str:
    # Sklep rozdziela czlony danych producenta przecinkami, nie srednikami.
    parts = manufacturer_data_parts(value)
    return ", ".join(parts) if parts else compact_spaces(value.replace("\\", ""))


def manufacturer_data_parts(value: Any) -> list[str]:
    text = compact_spaces(str(value or ""))
    if not text:
        return []
    raw_parts = re.split(r"\s*(?:;|\||,|>)\s*", text)
    parts: list[str] = []
    for part in raw_parts:
        cleaned = compact_spaces(part).strip("\\").strip()
        if cleaned:
            parts.append(cleaned)
    return parts


def build_manufacturer_data_by_producer(catalog_knowledge: dict[str, Any]) -> dict[str, str]:
    best: dict[str, tuple[str, int, int]] = {}
    rows = catalog_knowledge.get("top_attribute_values_by_category") or []
    current_category = ""
    current_count: int | None = None
    current_parts: list[str] = []
    producer_counts_by_category: dict[tuple[str, int], list[str]] = {}
    data_candidates: list[tuple[str, int, str, str, int]] = []

    def remember(producer: str, data: str, count: int, parts_count: int) -> None:
        producer_key = normalize_header(producer)
        if not producer_key:
            return
        previous = best.get(producer_key)
        candidate = (data, count, parts_count)
        if not previous or (candidate[1], candidate[2]) > (previous[1], previous[2]):
            best[producer_key] = candidate

    def flush() -> None:
        nonlocal current_count, current_parts
        if len(current_parts) < 2 or current_count is None:
            current_parts = []
            current_count = None
            return
        producer_key = normalize_header(current_parts[0])
        if not producer_key:
            current_parts = []
            current_count = None
            return
        data = ", ".join(current_parts)
        remember(current_parts[0], data, current_count, len(current_parts))
        data_candidates.append((current_category, current_count, current_parts[0], data, len(current_parts)))
        current_parts = []
        current_count = None

    for row in rows:
        if not isinstance(row, dict):
            continue
        attribute = compact_spaces(str(row.get("attribute", "")))
        category = compact_spaces(str(row.get("category", "")))
        count = int(row.get("count") or 0)
        if normalize_header(attribute) == "producent":
            producer = compact_spaces(str(row.get("value", "")))
            if producer:
                producer_counts_by_category.setdefault((category, count), []).append(producer)
        if normalize_header(attribute) != "dane producenta":
            flush()
            current_category = category
            continue

        if current_parts and (category != current_category or count != current_count):
            flush()
        current_category = category
        current_count = count
        current_parts.extend(manufacturer_data_parts(row.get("value", "")))
        if any("@" in part for part in current_parts):
            flush()
    flush()

    for category, count, data_producer, data, parts_count in data_candidates:
        data_producer_key = normalize_header(data_producer)
        for producer in producer_counts_by_category.get((category, count), []):
            producer_key = normalize_header(producer)
            if producer_key == data_producer_key or data_producer_key.startswith(f"{producer_key} "):
                remember(producer, data, count, parts_count)

    return {producer: data for producer, (data, _, __) in best.items()}


def normalize_decimal_separator(value: str) -> str:
    return re.sub(r"(?<=\d)\.(?=\d)", ",", value)


def capitalize_first(value: str) -> str:
    return value[:1].upper() + value[1:] if value else value


def format_sku(
    sku: str,
    producer: str,
    sku_format: str,
    producer_suffixes: dict[str, str] | None = None,
) -> str:
    sku = compact_spaces(sku)
    if sku_format == "plain" or not sku:
        return sku
    if "/" in sku:
        return sku
    suffix = producer_suffix(producer, producer_suffixes)
    return f"{sku}/{suffix}" if suffix else sku


def producer_suffix(producer: str, producer_suffixes: dict[str, str] | None = None) -> str:
    normalized_producer = normalize_producer_key(producer)
    if producer_suffixes and normalized_producer:
        suffix = producer_suffixes.get(normalized_producer)
        if suffix:
            return suffix
    letters = re.findall(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", str(producer or ""))
    return "".join(letters[:3]).upper()


def load_producer_suffixes(
    suffix_knowledge_path: str | Path | None = DEFAULT_SKU_SUFFIX_KNOWLEDGE_PATH,
    catalog_knowledge_path: str | Path | None = DEFAULT_CATALOG_KNOWLEDGE_PATH,
) -> dict[str, str]:
    knowledge = load_yaml(suffix_knowledge_path) if suffix_knowledge_path else {}
    catalog_knowledge = load_yaml(catalog_knowledge_path) if catalog_knowledge_path else {}

    suffix_counts: dict[str, dict[str, int]] = {}
    collect_suffixes_from_sources(knowledge, suffix_counts)
    collect_suffixes_from_examples(knowledge, suffix_counts, minimum_count=5)
    collect_suffixes_from_titles(knowledge, catalog_knowledge, suffix_counts, minimum_count=5)

    return {
        producer_key: best_suffix(counts)
        for producer_key, counts in suffix_counts.items()
        if best_suffix(counts)
    }


def collect_suffixes_from_sources(knowledge: dict[str, Any], suffix_counts: dict[str, dict[str, int]]) -> None:
    for row in knowledge.get("producer_sources") or []:
        if not isinstance(row, dict):
            continue
        source = str(row.get("source", ""))
        match = re.fullmatch(r"sku_suffix:([A-Za-z0-9]+)", source)
        if not match:
            continue
        add_suffix_count(
            suffix_counts,
            str(row.get("producer", "")),
            match.group(1),
            int(row.get("count") or 1),
        )


def collect_suffixes_from_examples(
    knowledge: dict[str, Any],
    suffix_counts: dict[str, dict[str, int]],
    minimum_count: int,
) -> None:
    grouped_counts: dict[str, dict[str, int]] = {}
    for row in iter_knowledge_examples(knowledge):
        add_suffix_count(grouped_counts, row.get("producer", ""), sku_suffix_from_value(row.get("sku", "")), 1)

    for producer_key, counts in grouped_counts.items():
        suffix, count = best_suffix_with_count(counts)
        if suffix and count >= minimum_count:
            add_suffix_count(suffix_counts, producer_key, suffix, count)


def collect_suffixes_from_titles(
    knowledge: dict[str, Any],
    catalog_knowledge: dict[str, Any],
    suffix_counts: dict[str, dict[str, int]],
    minimum_count: int,
) -> None:
    producers = known_catalog_producers(catalog_knowledge)
    if not producers:
        return

    grouped_counts: dict[str, dict[str, int]] = {}
    for row in iter_knowledge_examples(knowledge):
        title_key = normalize_producer_key(row.get("title", ""))
        suffix = sku_suffix_from_value(row.get("sku", ""))
        if not title_key or not suffix:
            continue
        for producer, producer_key in producers:
            if re.search(rf"(?<!\S){re.escape(producer_key)}(?!\S)", title_key):
                add_suffix_count(grouped_counts, producer, suffix, 1)

    for producer_key, counts in grouped_counts.items():
        suffix, count = best_suffix_with_count(counts)
        if suffix and count >= minimum_count:
            add_suffix_count(suffix_counts, producer_key, suffix, count)


def iter_knowledge_examples(knowledge: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    examples = knowledge.get("examples") or {}
    if not isinstance(examples, dict):
        return rows
    for example_rows in examples.values():
        if not isinstance(example_rows, list):
            continue
        rows.extend(row for row in example_rows if isinstance(row, dict))
    return rows


def known_catalog_producers(catalog_knowledge: dict[str, Any]) -> list[tuple[str, str]]:
    producers: list[tuple[str, str]] = []
    for row in catalog_knowledge.get("producers") or []:
        producer = str(row.get("producer", "")).strip() if isinstance(row, dict) else ""
        producer_key = normalize_producer_key(producer)
        if len(producer_key) >= 3:
            producers.append((producer, producer_key))
    return producers


def add_suffix_count(counts: dict[str, dict[str, int]], producer: Any, suffix: Any, count: int) -> None:
    producer_key = normalize_producer_key(producer)
    suffix_text = compact_spaces(str(suffix or "")).upper()
    if not producer_key or not suffix_text:
        return
    producer_counts = counts.setdefault(producer_key, {})
    producer_counts[suffix_text] = producer_counts.get(suffix_text, 0) + count


def sku_suffix_from_value(value: Any) -> str:
    sku = compact_spaces(str(value or ""))
    if "/" not in sku:
        return ""
    return sku.rsplit("/", 1)[1].strip().upper()


def best_suffix(counts: dict[str, int]) -> str:
    return best_suffix_with_count(counts)[0]


def best_suffix_with_count(counts: dict[str, int]) -> tuple[str, int]:
    if not counts:
        return "", 0
    return max(counts.items(), key=lambda item: (item[1], item[0]))


def normalize_producer_key(value: Any) -> str:
    return normalize_header(str(value or ""))


def only_digits(value: Any) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    return digits or compact_spaces(str(value or ""))


if __name__ == "__main__":
    main()
