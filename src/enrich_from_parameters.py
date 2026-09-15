from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

from analyze_products import detect_product_columns
from extract_attributes import ATTRIBUTE_COLUMNS, extract_attributes_for_dataframe, normalize_attribute_value
from utils import compact_spaces, detect_column, ensure_dir, is_blank, load_yaml, normalize_header, read_products, write_products


PARAMETER_ATTRIBUTE_COLUMNS = ATTRIBUTE_COLUMNS + [
    "barwa_zakres",
    "cri",
    "klosz",
    "klasa_energetyczna",
    "laczenie_przelotowe",
    "liczba_gniazd",
    "lm_w",
    "eco",
    "moc_max_zrodla",
    "pasuje_do",
    "sciemnianie",
    "sposob_montazu",
    "sterowanie",
    "trwalosc",
    "ugr",
    "zrodlo_swiatla",
    "zrodlo_w_komplecie",
]

# Parametry producenta, ktore wygladaja jak znane atrybuty, ale nimi nie sa
# (np. "MLS" przy "napieciem sieciowym" to typ zrodla, nie napiecie).
IGNORED_PARAMETER_NAME_PARTS = (
    "wspolczynnik mocy",
    "barwa swiatla",
    "materialem termoizolacyjnym",
    "napieciem sieciowym",
    "maksymalna moc opraw",
    "pobor mocy",
    "moc transmisji",
    "moc maksymalna va",
    "klasa energetyczna statecznika",
)


def normalize_sku(value: Any) -> str:
    value = compact_spaces(str(value))
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    return value


def detect_parameter_columns(df: pd.DataFrame) -> dict[str, str | None]:
    return {
        "sku": detect_column(df, ["Kod Kanlux", "kod", "sku", "indeks"]),
        "name": detect_column(df, ["Nazwa", "nazwa produktu", "name"]),
        "attribute_id": detect_column(df, ["ID atrybutu", "id atrybutu", "attribute id"]),
        "attribute_name": detect_column(df, ["Nazwa atrybutu", "atrybut", "attribute"]),
        "value": detect_column(df, ["Wartość", "wartosc", "value"]),
    }


def parameter_attribute_to_internal(attribute_name: str) -> str:
    name = normalize_header(attribute_name)
    if any(part in name for part in IGNORED_PARAMETER_NAME_PARTS):
        return ""
    exact_rules = {
        "ugr": "ugr",
        "wspolczynnik oddawania barw ra": "cri",
        "trwalosc h": "trwalosc",
        "klasa efektywnosci energetycznej": "klasa_energetyczna",
        "miejsce montazu": "sposob_montazu",
        "mozliwosc laczenia przelotowego opraw": "laczenie_przelotowe",
        "mozliwosc wspolpracy ze sciemniaczem": "sciemnianie",
        "zrodlo swiatla": "zrodlo_swiatla",
        "zrodlo swiatla w komplecie": "zrodlo_w_komplecie",
        "zintegrowane zrodlo swiatla led": "zrodlo_w_komplecie",
        "ilosc sztuk w opakowaniu jednostkowym": "ilosc_sztuk",
        "typ klosza": "klosz",
    }
    if name in exact_rules:
        return exact_rules[name]
    rules = [
        ("barwa_zakres", ["zakres temperatury barwowej", "regulowana temperatura barwowa"]),
        ("lm_w", ["skutecznosc swietlna", "lm w"]),
        ("strumien", ["uzyteczny strumien swietlny", "strumien swietlny", "strumien"]),
        ("moc", ["moc w trybie wlaczenia", "moc znamionowa", "moc maksymalna", "moc"]),
        ("barwa", ["temperatura barwowa"]),
        ("ip", ["stopien ip", "klasa ip"]),
        ("ik", ["stopien ik", "klasa ik"]),
        ("kat_swiecenia", ["kat swiecenia", "kat rozsyłu", "kat rozsylu"]),
        ("gwint", ["trzonek", "gwint"]),
        ("srednica", ["srednica"]),
        ("dlugosc_przewodu", ["dlugosc przewodu"]),
        ("dlugosc", ["dlugosc"]),
        ("szerokosc", ["szerokosc"]),
        ("wysokosc", ["wysokosc"]),
        ("kolor", ["kolor"]),
        ("material", ["material"]),
        ("napiecie", ["napiecie znamionowe", "napiecie"]),
        ("czujnik", ["czujnik ruchu", "czujnik"]),
    ]
    for internal, needles in rules:
        if any(needle in name for needle in needles):
            return internal
    return ""


def normalize_parameter_value(value: str, attr: str, parameter_name: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    name = normalize_header(parameter_name)

    if attr in {"dlugosc", "szerokosc", "wysokosc", "srednica"}:
        unit = "mm" if "mm" in name else "cm" if "cm" in name else ""
        match = re.search(r"\d+(?:[,.]\d+)?", value)
        if match and unit:
            return f"{match.group(0).replace(',', '.')}{unit}"
    if attr == "dlugosc_przewodu":
        unit = "m" if re.search(r"\bm\b", name) else "mm" if "mm" in name else "cm" if "cm" in name else ""
        match = re.search(r"\d+(?:[,.]\d+)?", value)
        if match and unit:
            return f"{match.group(0).replace(',', '.')}{unit}"
    if attr == "ilosc_sztuk":
        match = re.search(r"\d+", value)
        return f"{match.group(0)} szt." if match else ""
    if attr == "lm_w":
        match = re.search(r"\d+(?:[,.]\d+)?", value)
        return f"{match.group(0).replace(',', '.')}lm/W" if match else ""
    if attr == "ugr":
        match = re.search(r"(?:<|≤|<=)?\s*\d{1,2}", value)
        return re.sub(r"\s+", "", match.group(0)).replace("<=", "≤") if match else ""
    if attr == "cri":
        match = re.search(r"\d{2,3}", value)
        return match.group(0) if match else ""
    if attr == "trwalosc":
        match = re.search(r"\d+", value.replace(" ", ""))
        return match.group(0) if match else ""
    if attr == "laczenie_przelotowe":
        lowered = normalize_header(value)
        return "Tak" if lowered in {"1", "tak", "yes", "true"} else "Nie" if lowered in {"0", "nie", "no", "false"} else ""
    if attr == "sciemnianie":
        lowered = normalize_header(value)
        if lowered in {"nie", "no", "0", "false"}:
            return "Nie"
        # Kazda forma wspolpracy (Tak, DALI, 1-10V, 0-10V...) to po prostu "Tak".
        return "Tak"
    if attr == "klosz":
        # "pryzmatyczny,waskostrumieniowy" -> "pryzmatyczny, waskostrumieniowy".
        return compact_spaces(re.sub(r"\s*,\s*", ", ", value))
    if attr == "zrodlo_swiatla":
        return value
    if attr == "zrodlo_w_komplecie":
        lowered = normalize_header(value)
        if lowered in {"1", "tak", "yes", "true"}:
            return "Tak"
        if lowered in {"0", "nie", "no", "false"}:
            return "Nie"
        return ""
    if attr == "barwa_zakres":
        values = [int(v) for v in re.findall(r"[23645]\d{3}", value)]
        return f"{min(values)}-{max(values)}K" if len(set(values)) >= 2 else ""
    if attr == "barwa":
        values = [int(v) for v in re.findall(r"[23645]\d{3}", value)]
        if len(set(values)) >= 2:
            return f"{min(values)}-{max(values)}K"
    if attr == "klasa_energetyczna":
        match = re.fullmatch(r"[A-G]\+{0,3}", value.strip().upper())
        return match.group(0) if match else ""
    if attr == "sposob_montazu":
        lowered = normalize_header(value)
        surface = "do nadbudowania" in lowered or "na podlozu" in lowered
        recessed = "do wbudowania" in lowered
        if surface and not recessed:
            return "Natynkowy"
        if recessed and not surface:
            return "Podtynkowy"
        return ""
    if attr == "moc_max_zrodla":
        match = re.search(r"(\d+(?:[,.]\d+)?)\s*W?\s*$", value, flags=re.IGNORECASE)
        return f"{match.group(1).replace(',', '.')}W" if match else ""
    if attr == "moc":
        max_match = re.search(r"max\s*(\d+(?:[,.]\d+)?)\s*W?", value, flags=re.IGNORECASE)
        if max_match:
            return f"max {max_match.group(1).replace(',', '.')}W"
        range_match = re.fullmatch(r"\s*(\d+(?:[,.]\d+)?)\s*/\s*(\d+(?:[,.]\d+)?)\s*", value)
        if range_match:
            return f"{range_match.group(1).replace(',', '.')}-{range_match.group(2).replace(',', '.')}W"
    if attr == "ip":
        values = re.findall(r"\d{2}", value)
        if len(values) >= 2:
            return "/".join(f"IP{item}" for item in values[:2])
    if attr == "czujnik":
        lowered = normalize_header(value)
        if lowered in {"nie", "no", "0", "false"}:
            return ""
        if "zmierzch" in lowered:
            return "z czujnikiem zmierzchu"
        if "mikrofal" in lowered:
            return "z czujnikiem mikrofalowym"
        if "pir" in lowered:
            return "z czujnikiem ruchu PIR"
        if lowered in {"tak", "yes", "1", "true"} or "czuj" in lowered or "sensor" in lowered:
            return "z czujnikiem ruchu"
    if attr == "ik":
        match = re.search(r"\d{1,2}", value)
        if match:
            return f"IK{int(match.group(0)):02d}"
    if attr == "strumien":
        values = [float(v.replace(",", ".")) for v in re.findall(r"\d+(?:[,.]\d+)?", value)]
        if values:
            min_value = min(values)
            max_value = max(values)
            def fmt(number: float) -> str:
                return str(int(number)) if number.is_integer() else str(number).rstrip("0").rstrip(".")
            return f"{fmt(min_value)}-{fmt(max_value)}lm" if min_value != max_value else f"{fmt(max_value)}lm"

    return normalize_attribute_value(value, attr)


# Parametry, dla ktorych kolumny master/ekstrakcja maja pierwszenstwo nad Parametrami
# producenta:
#  - "material": Parametry podaja material klosza (np. szklo) zamiast materialu obudowy
#    (np. drewno przy plafonach drewnianych);
#  - "moc": Parametry podaja moc przelaczana/wielowartosciowa (np. "52 / 44", "11-19"),
#    ktora nie jest poprawna wartoscia atrybutu "Moc [W]" (zakres idzie osobno do
#    "Moc - zakres"); zostaje pojedyncza moc z master/tytulu.
PARAMETER_OVERRIDE_EXCEPTIONS = {"material", "moc"}


def should_parameter_override(attr: str, current_value: Any, parameter_value: str, source: str) -> bool:
    """Parametry producenta sa zrodlem prawdy - wygrywaja z kolumnami master i ekstrakcja
    z tytulu (poza PARAMETER_OVERRIDE_EXCEPTIONS). Reczne uzupelnienia uzytkownika sa
    nakladane pozniej (build_kanlux_baselinker_update) i nadal wygrywaja z Parametrami."""
    current = compact_spaces(str(current_value))
    if is_blank(current) or not parameter_value:
        return False
    if attr in PARAMETER_OVERRIDE_EXCEPTIONS:
        return False
    return True


def build_dimensions_from_parts(length: Any, width: Any) -> str:
    if is_blank(length) or is_blank(width):
        return ""
    length_value = compact_spaces(str(length))
    width_value = compact_spaces(str(width))
    length_match = re.fullmatch(r"(\d+(?:[,.]\d+)?)(mm|cm)", length_value, flags=re.IGNORECASE)
    width_match = re.fullmatch(r"(\d+(?:[,.]\d+)?)(mm|cm)", width_value, flags=re.IGNORECASE)
    if length_match and width_match and length_match.group(2).lower() == width_match.group(2).lower():
        length_number = length_match.group(1).replace(",", ".")
        width_number = width_match.group(1).replace(",", ".")
        return f"{length_number}x{width_number}{length_match.group(2).lower()}"
    return f"{length_value} x {width_value}"


def infer_shape_from_dimensions(dimensions: Any = "", length: Any = "", width: Any = "") -> str:
    length_value = ""
    width_value = ""

    if not is_blank(length) and not is_blank(width):
        length_value = compact_spaces(str(length))
        width_value = compact_spaces(str(width))
    elif not is_blank(dimensions):
        text = compact_spaces(str(dimensions)).replace("×", "x").lower()
        match = re.search(r"(\d+(?:[,.]\d+)?)\s*(?:cm|mm)?\s*x\s*(\d+(?:[,.]\d+)?)", text)
        if match:
            length_value = match.group(1)
            width_value = match.group(2)

    if not length_value or not width_value:
        return ""
    length_match = re.search(r"\d+(?:[,.]\d+)?", length_value)
    width_match = re.search(r"\d+(?:[,.]\d+)?", width_value)
    if not length_match or not width_match:
        return ""
    length_number = float(length_match.group(0).replace(",", "."))
    width_number = float(width_match.group(0).replace(",", "."))
    return "kwadratowa" if abs(length_number - width_number) < 0.001 else "prostokątna"


def synthesize_composite_attributes(enriched: pd.DataFrame, reports: dict[str, pd.DataFrame], sku_column: str) -> pd.DataFrame:
    result = enriched.copy()
    synthesized: list[dict[str, Any]] = []
    for index, row in result.iterrows():
        sku = normalize_sku(row.get(sku_column, ""))
        try:
            sources = json.loads(row.get("attribute_sources", "{}") or "{}")
        except json.JSONDecodeError:
            sources = {}
        try:
            confidence = json.loads(row.get("attribute_confidence", "{}") or "{}")
        except json.JSONDecodeError:
            confidence = {}

        dimensions = row.get("attr_wymiary", "")
        if is_blank(dimensions):
            dimensions = build_dimensions_from_parts(row.get("attr_dlugosc", ""), row.get("attr_szerokosc", ""))
            if dimensions:
                result.at[index, "attr_wymiary"] = dimensions
                synthesized.append(
                    {
                        "sku": sku,
                        "attribute": "wymiary",
                        "old_value": "",
                        "new_value": dimensions,
                        "parameter_attribute": "Długość + Szerokość",
                        "parameter_raw_value": f"{row.get('attr_dlugosc', '')} + {row.get('attr_szerokosc', '')}",
                    }
                )
                sources["wymiary"] = "parameters:dlugosc+szerokosc"
                confidence["wymiary"] = 0.94

        if is_blank(row.get("attr_ksztalt", "")):
            shape = infer_shape_from_dimensions(dimensions, row.get("attr_dlugosc", ""), row.get("attr_szerokosc", ""))
            if shape:
                result.at[index, "attr_ksztalt"] = shape
                synthesized.append(
                    {
                        "sku": sku,
                        "attribute": "ksztalt",
                        "old_value": "",
                        "new_value": shape,
                        "parameter_attribute": "Wymiary",
                        "parameter_raw_value": dimensions,
                    }
                )
                sources["ksztalt"] = "dimensions"
                confidence["ksztalt"] = 0.82

        result.at[index, "attribute_sources"] = json.dumps(sources, ensure_ascii=False, sort_keys=True)
        result.at[index, "attribute_confidence"] = json.dumps(confidence, ensure_ascii=False, sort_keys=True)

    if synthesized:
        synthesized_df = pd.DataFrame(
            synthesized,
            columns=["sku", "attribute", "old_value", "new_value", "parameter_attribute", "parameter_raw_value"],
        )
        reports["filled"] = pd.concat([reports["filled"], synthesized_df], ignore_index=True)
    return result


def build_parameter_index(params: pd.DataFrame, columns: dict[str, str | None]) -> tuple[dict[str, dict[str, dict[str, str]]], pd.DataFrame]:
    required = ["sku", "attribute_name", "value"]
    missing_columns = [key for key in required if not columns.get(key)]
    if missing_columns:
        raise SystemExit(f"Brak kolumn w pliku parametrow: {', '.join(missing_columns)}")

    sku_column = str(columns["sku"])
    name_column = columns.get("name")
    attr_column = str(columns["attribute_name"])
    value_column = str(columns["value"])

    index: dict[str, dict[str, dict[str, str]]] = {}
    unmapped_rows: list[dict[str, Any]] = []

    for _, row in params.iterrows():
        sku = normalize_sku(row.get(sku_column, ""))
        parameter_name = compact_spaces(str(row.get(attr_column, "")))
        raw_value = compact_spaces(str(row.get(value_column, "")))
        if not sku or not parameter_name or not raw_value:
            continue

        internal = parameter_attribute_to_internal(parameter_name)
        socket_count = ""
        if internal == "moc" and "maksymalna" in normalize_header(parameter_name) and (
            re.search(r"\bmax\b", raw_value, flags=re.IGNORECASE)
            or re.fullmatch(r"\d+(?:[,.]\d+)?\s*x\s*\d+(?:[,.]\d+)?\s*W?", raw_value, flags=re.IGNORECASE)
        ):
            # "max 60" / "2 x max 40" / "2 x 36" przy oprawie na wymienne zrodlo
            # to maksymalna moc zrodla swiatla, nie moc oprawy.
            internal = "moc_max_zrodla"
            count_match = re.match(r"^\s*(\d{1,2})\s*x\b", raw_value, flags=re.IGNORECASE)
            if count_match:
                socket_count = count_match.group(1)
        if not internal:
            unmapped_rows.append(
                {
                    "parameter_attribute": parameter_name,
                    "value": raw_value,
                    "count": 1,
                }
            )
            continue

        normalized_value = normalize_parameter_value(raw_value, internal, parameter_name)
        if not normalized_value:
            continue

        product = index.setdefault(sku, {})
        existing = product.get(internal)
        explicit_included_source = normalize_header(parameter_name) == "zrodlo swiatla w komplecie"
        existing_is_inferred_source = existing and normalize_header(existing["parameter_attribute"]) == "zintegrowane zrodlo swiatla led"
        if internal not in product or (explicit_included_source and existing_is_inferred_source):
            product[internal] = {
                "value": normalized_value,
                "parameter_attribute": parameter_name,
                "raw_value": raw_value,
                "parameter_product_name": compact_spaces(str(row.get(name_column, ""))) if name_column else "",
            }
        if socket_count and "liczba_gniazd" not in product:
            product["liczba_gniazd"] = {
                "value": socket_count,
                "parameter_attribute": parameter_name,
                "raw_value": raw_value,
                "parameter_product_name": compact_spaces(str(row.get(name_column, ""))) if name_column else "",
            }

    reroute_integrated_led_max_power(index)

    unmapped = pd.DataFrame(unmapped_rows)
    if not unmapped.empty:
        unmapped = (
            unmapped.groupby(["parameter_attribute", "value"], dropna=False)
            .size()
            .reset_index(name="count")
            .sort_values("count", ascending=False)
        )
    else:
        unmapped = pd.DataFrame(columns=["parameter_attribute", "value", "count"])
    return index, unmapped


def reroute_integrated_led_max_power(index: dict[str, dict[str, dict[str, str]]]) -> None:
    """Przy zintegrowanym LED bez trzonka "max NN W" to moc oprawy, nie moc zarowki."""
    for product in index.values():
        moc_max = product.get("moc_max_zrodla")
        if not moc_max or "gwint" in product:
            continue
        integrated = product.get("zrodlo_w_komplecie")
        if (
            integrated
            and normalize_header(integrated["parameter_attribute"]) == "zintegrowane zrodlo swiatla led"
            and integrated["value"] == "Tak"
        ):
            product.setdefault("moc", moc_max)
            del product["moc_max_zrodla"]


def add_working_title_column(df: pd.DataFrame, title_column: str, config: dict[str, Any]) -> tuple[pd.DataFrame, str]:
    result = df.copy()
    fallback_columns = [column for column in config.get("title_fallback_columns", []) if column in result.columns]
    working = result[title_column].astype(str)
    for column in fallback_columns:
        fallback = result[column].astype(str)
        working = working.where(working.str.strip().ne(""), fallback)
    result["__working_title"] = working
    return result, "__working_title"


def enrich_from_parameters(
    products: pd.DataFrame,
    product_columns: dict[str, str | None],
    config: dict[str, Any],
    parameter_index: dict[str, dict[str, dict[str, str]]],
) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    sku_column = product_columns.get("sku")
    title_column = product_columns.get("title")
    if not sku_column:
        raise SystemExit("Nie wykryto kolumny SKU w pliku produktow.")
    if not title_column:
        raise SystemExit("Nie wykryto kolumny nazwy/tytulu w pliku produktow.")

    working_df, working_title_column = add_working_title_column(products, title_column, config)
    enriched = extract_attributes_for_dataframe(working_df, working_title_column, config, product_columns.get("producer"))

    matches: list[dict[str, Any]] = []
    filled: list[dict[str, Any]] = []
    conflicts: list[dict[str, Any]] = []
    unmatched: list[dict[str, Any]] = []

    for index, row in enriched.iterrows():
        sku = normalize_sku(row.get(sku_column, ""))
        params = parameter_index.get(sku)
        if not params:
            unmatched.append(
                {
                    "sku": sku,
                    "title": row.get(title_column, ""),
                    "category": row.get(product_columns.get("category") or "category", ""),
                }
            )
            enriched.at[index, "parameter_match_status"] = "NO_MATCH"
            continue

        fill_count = 0
        conflict_count = 0
        try:
            attribute_sources = json.loads(row.get("attribute_sources", "{}") or "{}")
        except json.JSONDecodeError:
            attribute_sources = {}
        try:
            attribute_confidence = json.loads(row.get("attribute_confidence", "{}") or "{}")
        except json.JSONDecodeError:
            attribute_confidence = {}
        for attr, payload in params.items():
            if attr not in PARAMETER_ATTRIBUTE_COLUMNS:
                continue
            column = f"attr_{attr}"
            current_value = row.get(column, "")
            parameter_value = payload["value"]
            if is_blank(current_value):
                enriched.at[index, column] = parameter_value
                attribute_sources[attr] = f"parameter:{payload['parameter_attribute']}"
                attribute_confidence[attr] = 0.97
                fill_count += 1
                filled.append(
                    {
                        "sku": sku,
                        "attribute": attr,
                        "old_value": "",
                        "new_value": parameter_value,
                        "parameter_attribute": payload["parameter_attribute"],
                        "parameter_raw_value": payload["raw_value"],
                    }
                )
                continue
            current_source = str(attribute_sources.get(attr, ""))
            if should_parameter_override(attr, current_value, parameter_value, current_source):
                enriched.at[index, column] = parameter_value
                attribute_sources[attr] = f"parameter_override:{payload['parameter_attribute']}"
                attribute_confidence[attr] = 0.98
                fill_count += 1
                filled.append(
                    {
                        "sku": sku,
                        "attribute": attr,
                        "old_value": current_value,
                        "new_value": parameter_value,
                        "parameter_attribute": payload["parameter_attribute"],
                        "parameter_raw_value": payload["raw_value"],
                    }
                )
                continue
            if compact_spaces(str(current_value)).lower() != compact_spaces(str(parameter_value)).lower():
                conflict_count += 1
                conflicts.append(
                    {
                        "sku": sku,
                        "attribute": attr,
                        "current_value": current_value,
                        "parameter_value": parameter_value,
                        "parameter_attribute": payload["parameter_attribute"],
                        "parameter_raw_value": payload["raw_value"],
                    }
                )

        if "czujnik" not in params and not is_blank(enriched.at[index, "attr_czujnik"] if "attr_czujnik" in enriched.columns else ""):
            current_source = str(attribute_sources.get("czujnik", ""))
            if current_source.startswith(("column:", "title", "row", "text_columns")):
                old_value = enriched.at[index, "attr_czujnik"]
                enriched.at[index, "attr_czujnik"] = ""
                attribute_sources["czujnik"] = "parameter_negative:no_sensor_attribute"
                attribute_confidence["czujnik"] = 0.98
                filled.append(
                    {
                        "sku": sku,
                        "attribute": "czujnik",
                        "old_value": old_value,
                        "new_value": "",
                        "parameter_attribute": "brak parametru czujnika",
                        "parameter_raw_value": "",
                    }
                )

        enriched.at[index, "attribute_sources"] = json.dumps(attribute_sources, ensure_ascii=False, sort_keys=True)
        enriched.at[index, "attribute_confidence"] = json.dumps(attribute_confidence, ensure_ascii=False, sort_keys=True)
        enriched.at[index, "parameter_match_status"] = "EXACT_SKU"
        enriched.at[index, "parameter_attributes"] = json.dumps(
            {attr: payload["value"] for attr, payload in sorted(params.items())},
            ensure_ascii=False,
            sort_keys=True,
        )
        matches.append(
            {
                "sku": sku,
                "title": row.get(title_column, ""),
                "parameter_product_name": next(iter(params.values())).get("parameter_product_name", ""),
                "matched_attributes": len(params),
                "filled_attributes": fill_count,
                "conflicts": conflict_count,
            }
        )

    reports = {
        "matches": pd.DataFrame(
            matches,
            columns=["sku", "title", "parameter_product_name", "matched_attributes", "filled_attributes", "conflicts"],
        ),
        "filled": pd.DataFrame(
            filled,
            columns=["sku", "attribute", "old_value", "new_value", "parameter_attribute", "parameter_raw_value"],
        ),
        "conflicts": pd.DataFrame(
            conflicts,
            columns=["sku", "attribute", "current_value", "parameter_value", "parameter_attribute", "parameter_raw_value"],
        ),
        "unmatched": pd.DataFrame(unmatched, columns=["sku", "title", "category"]),
    }
    enriched = synthesize_composite_attributes(enriched, reports, sku_column)
    return enriched, reports


def write_reports(
    reports: dict[str, pd.DataFrame],
    unmapped: pd.DataFrame,
    products: pd.DataFrame,
    product_sku_column: str,
    parameter_index: dict[str, dict[str, dict[str, str]]],
    reports_dir: Path,
) -> None:
    ensure_dir(reports_dir)
    for name, df in reports.items():
        df.to_csv(reports_dir / f"parameter_{name}.csv", index=False, encoding="utf-8-sig")
    unmapped.to_csv(reports_dir / "parameter_unmapped_attributes.csv", index=False, encoding="utf-8-sig")

    product_skus = {normalize_sku(value) for value in products[product_sku_column].astype(str)}
    parameter_skus = set(parameter_index)
    outside_rows = []
    for sku in sorted(parameter_skus - product_skus):
        attrs = parameter_index[sku]
        first = next(iter(attrs.values())) if attrs else {}
        outside_rows.append(
            {
                "sku": sku,
                "parameter_product_name": first.get("parameter_product_name", ""),
                "mapped_attributes": len(attrs),
            }
        )
    pd.DataFrame(outside_rows, columns=["sku", "parameter_product_name", "mapped_attributes"]).to_csv(
        reports_dir / "parameter_products_outside_input.csv", index=False, encoding="utf-8-sig"
    )

    summary = pd.DataFrame(
        [
            {
                "products_count": len(products),
                "parameter_products_count": len(parameter_skus),
                "matched_products": len(reports["matches"]),
                "unmatched_products": len(reports["unmatched"]),
                "filled_attributes": len(reports["filled"]),
                "conflicts": len(reports["conflicts"]),
                "parameter_products_outside_input": len(parameter_skus - product_skus),
            }
        ]
    )
    summary.to_csv(reports_dir / "parameter_summary.csv", index=False, encoding="utf-8-sig")


def main() -> None:
    parser = argparse.ArgumentParser(description="Dopasowuje plik parametrow Kanlux do produktow i uzupelnia brakujace atrybuty.")
    parser.add_argument("--input", required=True, help="Plik produktow, np. archived_input_files/input_2026-08-05/Alkan_Kanlux_pelne_rodziny.xlsx")
    parser.add_argument("--sheet", help="Arkusz produktow, np. 7. Wszystkie SKU (master)")
    parser.add_argument("--parameters", required=True, help="Plik Parametry.xlsx")
    parser.add_argument("--parameters-sheet", default=0)
    parser.add_argument("--config", default="configs/categories/kanlux-oswietlenie.yaml")
    parser.add_argument("--output", default="output/products_with_parameter_attributes.xlsx")
    parser.add_argument("--reports-dir", default="reports/parameter_enrichment")
    args = parser.parse_args()

    products = read_products(args.input, sheet_name=args.sheet)
    params = read_products(args.parameters, sheet_name=args.parameters_sheet)
    config = load_yaml(args.config)

    product_columns = detect_product_columns(products)
    parameter_columns = detect_parameter_columns(params)
    parameter_index, unmapped = build_parameter_index(params, parameter_columns)

    enriched, reports = enrich_from_parameters(products, product_columns, config, parameter_index)
    write_products(enriched, args.output)
    write_reports(reports, unmapped, products, str(product_columns["sku"]), parameter_index, Path(args.reports_dir))

    print(f"OK: zapisano {args.output}")
    print(f"OK: zapisano raporty w {args.reports_dir}")
    print(f"Dopasowane produkty: {len(reports['matches'])}")
    print(f"Uzupelnione atrybuty: {len(reports['filled'])}")
    print(f"Konflikty: {len(reports['conflicts'])}")


if __name__ == "__main__":
    main()
