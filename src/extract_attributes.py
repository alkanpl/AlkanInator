from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

from catalog_knowledge import DEFAULT_CATALOG_KNOWLEDGE_PATH, enrich_config_with_catalog_knowledge, load_catalog_knowledge
from utils import compact_spaces, first_present, load_yaml, read_products, strip_accents, write_products


ATTRIBUTE_COLUMNS = [
    "typ",
    "rodzaj",
    "moc",
    "napiecie",
    "barwa",
    "barwa_zakres",
    "strumien",
    "lm_w",
    "ip",
    "ik",
    "kat_swiecenia",
    "kolor",
    "producent",
    "seria",
    "model",
    "wymiary",
    "ksztalt",
    "dlugosc",
    "szerokosc",
    "wysokosc",
    "czujnik",
    "ilosc_sztuk",
    "eco",
    "gwarancja",
    "gwint",
    "srednica",
    "material",
    "pasuje_do",
    "sterowanie",
]


def extract_attributes_from_text(text: str, config: dict[str, Any] | None = None, producer_hint: str = "") -> dict[str, Any]:
    config = config or {}
    title = compact_spaces(text)
    title_ascii = strip_accents(title).lower()
    attributes: dict[str, str] = {}
    sources: dict[str, str] = {}
    confidence: dict[str, float] = {}

    patterns = {
        "moc": r"(?<![A-Z0-9])(\d{1,4}(?:[,.]\d+)?)\s*W(?![a-z])",
        "napiecie": r"(?<![A-Z0-9])(\d{2,4})\s*V\b",
        "barwa": r"(?<![A-Z0-9])([23645]\d{3})\s*K\b",
        "lm_w": r"(?<![A-Z0-9])(\d{2,3}(?:[,.]\d+)?)\s*lm\s*/\s*W\b",
        "ip": r"\bIP\s*([0-9]{2})\b",
        "ik": r"\bIK\s*([0-9]{2})\b",
        "kat_swiecenia": r"(?<![A-Z0-9])(\d{2,3})\s*(?:°|st\.?|stopni)\b",
        "gwint": r"\b([1-9][Xx]?\s*(?:E(?:14|27|40)|GU(?:5\.3|10)|MR16))\b",
    }

    for attr, pattern in patterns.items():
        match = re.search(pattern, title, flags=re.IGNORECASE)
        if not match:
            continue
        value = match.group(1).replace(",", ".")
        if attr == "moc":
            value = f"{value}W"
        elif attr == "napiecie":
            value = f"{value}V"
        elif attr == "barwa":
            value = f"{value}K"
        elif attr == "lm_w":
            value = f"{normalize_number(value)}lm/W"
        elif attr == "strumien":
            value = f"{value}lm"
        elif attr == "ip":
            value = f"IP{value}"
        elif attr == "ik":
            value = f"IK{value}"
        elif attr == "kat_swiecenia":
            value = f"{value}°"
        elif attr == "gwint":
            value = normalize_socket(value)
        attributes[attr] = value
        sources[attr] = "title"
        confidence[attr] = 0.98

    luminous_flux = extract_luminous_flux(title)
    if luminous_flux:
        attributes["strumien"] = luminous_flux
        sources["strumien"] = "title"
        confidence["strumien"] = 0.95

    color_temperature_range = extract_color_temperature_range(title)
    if color_temperature_range:
        attributes["barwa_zakres"] = color_temperature_range
        sources["barwa_zakres"] = "title"
        confidence["barwa_zakres"] = 0.88

    if "kat_swiecenia" not in attributes:
        angle_match = re.search(r"\b(?:k\u0105t|kat)\s*(\d{2,3})\s*(?:\u00b0|st\.?|stopni|D)?\b", title, flags=re.IGNORECASE)
        if angle_match:
            attributes["kat_swiecenia"] = f"{angle_match.group(1)}\u00b0"
            sources["kat_swiecenia"] = "title"
            confidence["kat_swiecenia"] = 0.92

    if "moc" not in attributes:
        power = extract_power_from_special_formats(title)
        if power:
            attributes["moc"] = power
            sources["moc"] = "title"
            confidence["moc"] = 0.85

    # Barwa z sufiksu Kanlux: NW=4000K, WW=3000K, CW=6500K (np. 30NW, 12-NW, 120NW, 125LM120NW)
    if "barwa" not in attributes:
        nw_match = (
            re.search(r"\b\d{1,3}[-\u2013]?(NW|WW|CW)\b", title, flags=re.IGNORECASE)
            or re.search(r"(?<=[A-Z])\d{1,3}(NW|WW|CW)\b", title)
        )
        if nw_match:
            nw_barwa = {"NW": "4000K", "WW": "3000K", "CW": "6500K"}
            barwa_val = nw_barwa.get(nw_match.group(1).upper(), "")
            if barwa_val:
                attributes["barwa"] = barwa_val
                sources["barwa"] = "title"
                confidence["barwa"] = 0.80

    dimensions = extract_dimensions(title)
    for attr, value in dimensions.items():
        attributes[attr] = value
        sources[attr] = "title"
        confidence[attr] = 0.88

    connector_size = extract_connector_size(title)
    if connector_size:
        attributes["wymiary"] = connector_size
        sources["wymiary"] = "title"
        confidence["wymiary"] = 0.90

    # Format panelu z kodu Kanlux (6060 → 60x60 cm)
    if "wymiary" not in attributes:
        panel_fmt = extract_panel_format(title)
        if panel_fmt:
            attributes["wymiary"] = panel_fmt
            sources["wymiary"] = "title"
            confidence["wymiary"] = 0.82

    shape = extract_shape(title_ascii)
    if shape:
        attributes["ksztalt"] = shape
        sources["ksztalt"] = "title"
        confidence["ksztalt"] = 0.88

    diameter = extract_diameter(title)
    if diameter:
        attributes["srednica"] = diameter
        sources["srednica"] = "title"
        confidence["srednica"] = 0.86

    sensor = extract_sensor(title)
    if sensor:
        attributes["czujnik"] = sensor
        sources["czujnik"] = "title"
        confidence["czujnik"] = 0.84

    quantity = extract_package_quantity(title)
    if quantity:
        attributes["ilosc_sztuk"] = quantity
        sources["ilosc_sztuk"] = "title"
        confidence["ilosc_sztuk"] = 0.92

    warranty = extract_warranty(title)
    if warranty:
        attributes["gwarancja"] = warranty
        sources["gwarancja"] = "title"
        confidence["gwarancja"] = 0.90

    if re.search(r"\bECO\b", title, flags=re.IGNORECASE):
        attributes["eco"] = "ECO"
        sources["eco"] = "title"
        confidence["eco"] = 0.90

    material = extract_material(title_ascii)
    if material:
        attributes["material"] = material
        sources["material"] = "title"
        confidence["material"] = 0.82

    accessory_target = extract_accessory_target(title)
    if accessory_target:
        attributes["pasuje_do"] = accessory_target
        sources["pasuje_do"] = "title"
        confidence["pasuje_do"] = 0.78

    driver_control = extract_driver_control(title)
    if driver_control:
        attributes["sterowanie"] = driver_control
        sources["sterowanie"] = "title"
        confidence["sterowanie"] = 0.86

    color = extract_color(title_ascii, config)
    if color:
        attributes["kolor"] = color
        sources["kolor"] = "title"
        confidence["kolor"] = 0.90

    # Kanlux: pojedyncza litera koloru na koncu kodu modelu (np. "DABER CCT W" -> bialy, "PHLOX GU10 B" -> czarny)
    if "kolor" not in attributes:
        suffix_match = re.search(r"\s([WB])$", title.strip(), re.IGNORECASE)
        if suffix_match:
            suffix = suffix_match.group(1).upper()
            kanlux_color_suffix = {"W": "bialy", "B": "czarny", "G": "grafitowy"}
            mapped = kanlux_color_suffix.get(suffix, "")
            if mapped:
                attributes["kolor"] = mapped
                sources["kolor"] = "title_suffix"
                confidence["kolor"] = 0.75

    producer = extract_producer(title, config, producer_hint)
    if producer:
        attributes["producent"] = producer
        sources["producent"] = "producer_column" if producer_hint else "title"
        confidence["producent"] = 0.99 if producer_hint else 0.90

    series, model = extract_series_and_model(title, attributes, config)
    if series:
        attributes["seria"] = series
        sources["seria"] = "title"
        confidence["seria"] = 0.65
    if model:
        attributes["model"] = model
        sources["model"] = "title"
        confidence["model"] = 0.60

    for key, value in (config.get("default_values") or {}).items():
        if key not in attributes:
            attributes[key] = str(value)
            sources[key] = "config_default"
            confidence[key] = 0.70

    return {"attributes": attributes, "sources": sources, "confidence": confidence}


def extract_color(title_ascii: str, config: dict[str, Any]) -> str:
    colors = config.get("known_colors") or {}
    for normalized, variants in colors.items():
        for variant in variants:
            variant_ascii = strip_accents(str(variant)).lower()
            if re.search(rf"\b{re.escape(variant_ascii)}\b", title_ascii):
                return str(normalized)
    return ""


def extract_power_from_special_formats(title: str) -> str:
    compact_title = compact_spaces(title)
    match = re.search(r"\b(\d{1,3}(?:[,.]\d+)?)W(?:CCT|DIM|NW|WW|CW)\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{normalize_number(match.group(1))}W"
    match = re.search(r"\b(\d{1,3})W\d{1,3}NW\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    # Kanlux: 12-NW, 18-NW. Nie lap formatow paneli typu 60NW/120NW jako mocy.
    match = re.search(r"\b(\d{1,3})[-\u2013]?(NW|WW|CW)\b", compact_title, flags=re.IGNORECASE)
    if match and int(match.group(1)) <= 50 and not re.search(r"\b(?:\d{2,3}LM\s*)?\d{2,3}(?:NW|WW|CW)\b", compact_title, flags=re.IGNORECASE):
        return f"{match.group(1)}W"
    # Kanlux kod zlozony. Pomijamy 125LM120NW, bo LM oznacza skutecznosc, a 120 format panelu.
    match = re.search(r"(?<=[A-Z])(\d{1,3})(NW|WW|CW)\b", compact_title)
    if match and int(match.group(1)) <= 50 and not re.search(r"\d{2,3}LM\d{2,3}(?:NW|WW|CW)\b", compact_title, flags=re.IGNORECASE):
        return f"{match.group(1)}W"
    # Kanlux kod modelu: SRM40W, TUN48W, U34W
    match = re.search(r"\b[A-Z]{2,}(\d{1,3})W(?!\s*[a-z])", compact_title)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\bLED[-\s]?(\d{1,3})(?:[-A-Z]|$)", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\b(\d{1,3})W[DS]\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\bTW(\d{1,3})NW\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\bDRV(?:M)?(\d{1,3})W(?![a-z])", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}W"
    match = re.search(r"\b(\d{1,2})\s*x\s*max\s*(\d{1,3})\b", compact_title, flags=re.IGNORECASE)
    if match:
        count = int(match.group(1))
        power = int(match.group(2))
        return f"{count}x{power}W"
    return ""


def extract_panel_format(title: str) -> str:
    """Wyciaga format panelu z kodu Kanlux (np. 6060 -> 60x60 cm) lub z tekstu (60x60)."""
    compact_title = compact_spaces(title)
    formats = {
        "60120": "60x120 cm",
        "30120": "30x120 cm",
        "12030": "120x30 cm",
        "6060": "60x60 cm",
        "3060": "30x60 cm",
        "2060": "20x60 cm",
        "3030": "30x30 cm",
        "1260": "12x60 cm",
    }
    for code, label in formats.items():
        if re.search(rf"\b{code}\b", compact_title):
            return label
    match = re.search(r"\b(\d{2,3})\s*[xX\u00d7]\s*(\d{2,3})(?:\s*cm)?\b", compact_title)
    if match:
        a, b = match.group(1), match.group(2)
        if 10 <= int(a) <= 200 and 10 <= int(b) <= 200:
            return f"{a}x{b} cm"
    return ""


def extract_luminous_flux(title: str) -> str:
    compact_title = compact_spaces(title)
    if re.search(r"\b\d{2,3}LM\b", compact_title, flags=re.IGNORECASE):
        return ""
    max_match = re.search(r"\bmax\s*(\d+(?:[,.]\d+)?)\s*(?:lm|lumenow|lumen\u00f3w)\b", compact_title, flags=re.IGNORECASE)
    if max_match:
        return f"{normalize_number(max_match.group(1))}lm"

    list_match = re.search(
        r"\b(\d+(?:[,.]\d+)?(?:\s*/\s*\d+(?:[,.]\d+)?)+)\s*(?:lm|lumenow|lumen\u00f3w)\b",
        compact_title,
        flags=re.IGNORECASE,
    )
    if list_match:
        values = [float(value.replace(",", ".")) for value in re.findall(r"\d+(?:[,.]\d+)?", list_match.group(1))]
        if not values:
            return ""
        min_value = format_number(min(values))
        max_value = format_number(max(values))
        return f"{min_value}-{max_value}lm" if min_value != max_value else f"{max_value}lm"

    match = re.search(r"(?<![A-Z0-9])(\d+(?:[,.]\d+)?)\s*(?:lm|lumenow|lumen\u00f3w)\b", compact_title, flags=re.IGNORECASE)
    if match:
        return f"{normalize_number(match.group(1))}lm"
    return ""


def extract_color_temperature_range(title: str) -> str:
    compact_title = compact_spaces(title)
    if not re.search(r"\bCCT\b|\bDIM\b|barwa", compact_title, flags=re.IGNORECASE):
        return ""
    values = [int(value) for value in re.findall(r"\b([23645]\d{3})(?=\s*K|\s*[-/])", compact_title, flags=re.IGNORECASE)]
    values.extend(int(value) for value in re.findall(r"(?<=[-/])\s*([23645]\d{3})\s*K\b", compact_title, flags=re.IGNORECASE))
    if len(set(values)) < 2:
        return ""
    return f"{min(values)}-{max(values)}K"


def extract_dimensions(title: str) -> dict[str, str]:
    dimensions: dict[str, str] = {}
    compact_title = compact_spaces(title)
    length = re.search(r"\b(?:d\u0142\.?|dl\.?|d\u0142ugo\u015b\u0107|dlugosc)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    length_mm = re.search(r"\b(?:d\u0142\.?|dl\.?|d\u0142ugo\u015b\u0107|dlugosc)?\s*(\d{3,4})\s*mm\b", compact_title, flags=re.IGNORECASE)
    width = re.search(r"\b(?:szer\.?|szeroko\u015b\u0107|szerokosc)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    height = re.search(r"\b(?:wys\.?|wysoko\u015b\u0107|wysokosc)\s*(\d+(?:[,.]\d+)?)\s*(?:cm|mm)\b", compact_title, flags=re.IGNORECASE)
    cm_pair = re.search(r"\b(\d+(?:[,.]\d+)?)\s*cm\s*[xX\u00d7/]\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title)
    if length:
        dimensions["dlugosc"] = normalize_number(length.group(1)) + "cm"
    elif length_mm:
        dimensions["dlugosc"] = normalize_number(length_mm.group(1)) + "mm"
    elif cm_pair:
        dimensions["dlugosc"] = normalize_number(cm_pair.group(1)) + "cm"
        dimensions["szerokosc"] = normalize_number(cm_pair.group(2)) + "cm"
    if width:
        dimensions["szerokosc"] = normalize_number(width.group(1)) + "cm"
    if height:
        dimensions["wysokosc"] = normalize_number(height.group(1)) + "cm"
    if "dlugosc" in dimensions and "szerokosc" in dimensions:
        dimensions["wymiary"] = format_dimensions_pair(dimensions["dlugosc"], dimensions["szerokosc"])
    return dimensions


def format_dimensions_pair(length: str, width: str) -> str:
    length_match = re.fullmatch(r"(\d+(?:[,.]\d+)?)(mm|cm)", compact_spaces(length), flags=re.IGNORECASE)
    width_match = re.fullmatch(r"(\d+(?:[,.]\d+)?)(mm|cm)", compact_spaces(width), flags=re.IGNORECASE)
    if length_match and width_match and length_match.group(2).lower() == width_match.group(2).lower():
        return f"{normalize_number(length_match.group(1))}x{normalize_number(width_match.group(1))}{length_match.group(2).lower()}"
    return f"{length} x {width}"


def extract_connector_size(title: str) -> str:
    match = re.search(r"\b(\d+\s*x\s*\d+(?:[,.]\d+)?)\s*MM2\b", title, flags=re.IGNORECASE)
    if not match:
        return ""
    return re.sub(r"\s+", "", match.group(1).replace(",", ".")) + "mm2"


def extract_diameter(title: str) -> str:
    compact_title = compact_spaces(title)
    match = re.search(r"(?:\u00d8|fi|\u015brednicy|srednica)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    if match:
        return normalize_number(match.group(1)) + "cm"
    match = re.search(r"(?:\u00d8|fi)\s*(\d+(?:[,.]\d+)?)\b", compact_title, flags=re.IGNORECASE)
    if match:
        value = float(match.group(1).replace(",", "."))
        if 10 <= value <= 100:
            return normalize_number(match.group(1)) + "cm"
    match = re.search(r"\b(?:okr\u0105g\u0142a|okragla|okr\u0105g\u0142y|okragly)\s*(\d+(?:[,.]\d+)?)\s*cm\b", compact_title, flags=re.IGNORECASE)
    if match:
        return normalize_number(match.group(1)) + "cm"
    return ""


def extract_sensor(title: str) -> str:
    title_ascii = strip_accents(title).lower()
    if re.search(r"\bmikrofal", title_ascii):
        return "z czujnikiem mikrofalowym"
    if re.search(r"\bpir\b", title_ascii):
        return "z czujnikiem ruchu PIR"
    if re.search(r"\bzmierzch", title_ascii):
        return "z czujnikiem zmierzchu"
    if re.search(r"\b(?:czujnik ruchu|sensor)\b", title_ascii):
        return "z czujnikiem ruchu"
    if re.search(r"\bczujnik\b", title_ascii):
        return "z czujnikiem ruchu"
    if re.search(r"\b[A-Z0-9]+(?:-[A-Z0-9]+)*-SE[A-Z]?\b", title):
        return "z czujnikiem ruchu"
    return ""


def extract_package_quantity(title: str) -> str:
    match = re.search(r"\b(?:kpl\.?|komplet)?\s*(\d+)\s*szt\.?\b", title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)} szt."
    match = re.search(r"\bkpl\.?\s*(\d+)\b", title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)} szt."
    return ""


def extract_warranty(title: str) -> str:
    match = re.search(r"\bGwarancj[ai]\s*(?:\[lat\])?\s*(?:lat)?\s*(\d+)\b", title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)} lat"
    match = re.search(r"\b(\d+)\s*lat\s*Gwarancji\b", title, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)} lat"
    return ""


def extract_material(title_ascii: str) -> str:
    if re.search(r"\bdrewnian", title_ascii):
        return "drewniana"
    if re.search(r"\balumini", title_ascii):
        return "aluminium"
    if re.search(r"\bstal", title_ascii):
        return "stal"
    return ""


def extract_accessory_target(title: str) -> str:
    compact_title = compact_spaces(title)
    normalized = strip_accents(compact_title).lower()
    if not re.search(r"\bdo\b", normalized):
        return ""

    if re.search(r"\bdo\s+paneli\b", normalized):
        return "do paneli"

    patterns = [
        r"\bdo\s+oprawy\s+(.+?)(?:\s+Gwarancja|\s+kolor|\s+kat|\s+k\u0105t|\s+LENS|\s+SENSOR|\s+REMOTE|\s+BRACKET|\s+kpl|\s*$)",
        r"\bdo\s+opraw\s+(.+?)(?:\s*\(|\s+Gwarancja|\s+kolor|\s+kpl|\s*$)",
        r"\bdo\s+(.+?)(?:\s+Gwarancja|\s+kolor|\s+kat|\s+k\u0105t|\s+LENS|\s+SENSOR|\s+REMOTE|\s+BRACKET|\s+CON|\s+kpl|\s*$)",
    ]
    for pattern in patterns:
        match = re.search(pattern, compact_title, flags=re.IGNORECASE)
        if not match:
            continue
        target = clean_accessory_target(match.group(1))
        if target:
            prefix = "do oprawy" if "oprawy" in pattern else "do opraw" if "opraw" in pattern else "do"
            return compact_spaces(f"{prefix} {target}")
    return ""


def clean_accessory_target(value: str) -> str:
    value = re.sub(r"\bkod\s+[\d,\s]+\b", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\b(?:LED|LENS|SENSOR|REMOTE|BRACKET|DRIVER|DRV|ON-OFF)\b", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\b\d+(?:[,.]\d+)?\s*(?:W|st\.?|stopni|D)\b", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\b\d{2,4}\b", " ", value)
    tokens = re.findall(r"[A-Za-z\u0104\u0106\u0118\u0141\u0143\u00d3\u015a\u0179\u017b\u0105\u0107\u0119\u0142\u0144\u00f3\u015b\u017a\u017c0-9-]{2,}", value)
    stop_words = {"kpl", "szt", "lat", "gwarancja", "kolor", "barwa", "zlozona", "zlozony"}
    useful = [token for token in tokens if strip_accents(token).lower() not in stop_words]
    return compact_spaces(" ".join(useful[:4]))


def extract_driver_control(title: str) -> str:
    compact_title = compact_spaces(title)
    normalized = strip_accents(compact_title).lower()
    if not re.search(r"\b(?:zasilacz|driver|drv|drvm|dali|on-off|1-10v)\b", normalized):
        return ""

    parts: list[str] = []
    if re.search(r"\bDRVM?\b|\bDRVM?\d", compact_title, flags=re.IGNORECASE):
        drv_match = re.search(r"\b(DRVM?)\b|\b(DRVM?)(?=\d)", compact_title, flags=re.IGNORECASE)
        if drv_match:
            parts.append(drv_match.group(1) or drv_match.group(2))
    elif re.search(r"\bDRIVER\b", compact_title, flags=re.IGNORECASE):
        parts.append("DRIVER")

    if re.search(r"\bON[-\s]?OFF\b", compact_title, flags=re.IGNORECASE):
        parts.append("ON-OFF")
    if re.search(r"\bDALI\b", compact_title, flags=re.IGNORECASE):
        parts.append("DALI")
    if re.search(r"1[-\s]?10\s*(?:V|mA)\b", compact_title, flags=re.IGNORECASE):
        parts.append("1-10V")

    result: list[str] = []
    for part in parts:
        normalized_part = part.upper()
        if normalized_part not in result:
            result.append(normalized_part)
    return " ".join(result)


def extract_shape(title_ascii: str) -> str:
    if re.search(r"\bprostokatny\b", title_ascii):
        return "prostok\u0105tna"
    if re.search(r"\bkwadratowy\b", title_ascii):
        return "kwadratowa"
    if re.search(r"\bokragl[ay]\b", title_ascii):
        return "okr\u0105g\u0142a"
    return ""


def normalize_number(value: str) -> str:
    value = value.replace(",", ".")
    if value.endswith(".0"):
        value = value[:-2]
    return value


def format_number(value: float) -> str:
    return str(int(value)) if value.is_integer() else str(value).rstrip("0").rstrip(".")


def normalize_socket(value: str) -> str:
    value = re.sub(r"\s+", "", value.upper())
    value = value.replace("X", "x")
    return value


def extract_producer(title: str, config: dict[str, Any], producer_hint: str = "") -> str:
    if producer_hint:
        return compact_spaces(producer_hint)
    title_ascii = strip_accents(title).lower()
    for producer in config.get("known_producers") or []:
        producer_ascii = strip_accents(str(producer)).lower()
        if re.search(rf"\b{re.escape(producer_ascii)}\b", title_ascii):
            return str(producer)
    return ""


def extract_series_and_model(title: str, attributes: dict[str, str], config: dict[str, Any]) -> tuple[str, str]:
    cleaned = title
    for value in sorted(attributes.values(), key=len, reverse=True):
        cleaned = re.sub(re.escape(str(value)), " ", cleaned, flags=re.IGNORECASE)
    technical_patterns = [
        r"\b\d+(?:[,.]\d+)?\s*W\b",
        r"\b\d{2,4}\s*V\b",
        r"\b[23645]\d{3}\s*K\b",
        r"\b\d+(?:[,.]\d+)?\s*(?:lm|lumenow|lumen\u00f3w)\b",
        r"\b\d{2,3}(?:[,.]\d+)?\s*lm\s*/\s*W\b",
        r"\bIP\s*\d{2}\b",
        r"\bIK\s*\d{2}\b",
    ]
    for pattern in technical_patterns:
        cleaned = re.sub(pattern, " ", cleaned, flags=re.IGNORECASE)
    stop_words = {strip_accents(word).lower() for word in config.get("series_stop_words") or []}
    tokens = re.findall(r"[A-Za-z\u0104\u0106\u0118\u0141\u0143\u00d3\u015a\u0179\u017b\u0105\u0107\u0119\u0142\u0144\u00f3\u015b\u017a\u017c0-9-]{3,}", cleaned)
    useful = [token for token in tokens if strip_accents(token).lower() not in stop_words]
    uppercase_tokens = [token for token in useful if token.isupper() and not token.isdigit()]
    titlecase_tokens = [token for token in useful if token[:1].isupper() and not any(char.isdigit() for char in token)]
    series = first_present(find_known_series(title, config), " ".join((uppercase_tokens or titlecase_tokens)[:2]))
    model = first_present(
        *(token for token in useful if any(char.isdigit() for char in token) and not token.isdigit())
    )
    return compact_spaces(series), compact_spaces(model)


def find_known_series(title: str, config: dict[str, Any]) -> str:
    title_ascii = strip_accents(title).lower()
    for series in sorted(config.get("known_series") or [], key=lambda value: len(str(value)), reverse=True):
        normalized = strip_accents(str(series)).lower()
        if len(normalized) < 3:
            continue
        if re.search(rf"(?<![a-z0-9]){re.escape(normalized)}(?![a-z0-9])", title_ascii):
            return str(series)
    return ""


def extract_attributes_for_dataframe(
    df: pd.DataFrame,
    title_column: str,
    config: dict[str, Any],
    producer_column: str | None = None,
) -> pd.DataFrame:
    result = df.copy()
    extracted_rows: list[dict[str, Any]] = []

    for _, row in result.iterrows():
        producer_hint = row.get(producer_column, "") if producer_column else ""
        extracted = extract_attributes_from_text(str(row.get(title_column, "")), config, str(producer_hint))
        attrs = extracted["attributes"]
        merge_attributes_from_columns(attrs, extracted["sources"], extracted["confidence"], row, config)
        sensor_from_row = extract_sensor_from_row(row, config)
        if sensor_from_row:
            attrs["czujnik"] = sensor_from_row
            extracted["sources"]["czujnik"] = "row"
            extracted["confidence"]["czujnik"] = 0.88
        for attr in ATTRIBUTE_COLUMNS:
            result.at[row.name, f"attr_{attr}"] = attrs.get(attr, "")
        result.at[row.name, "attribute_sources"] = json.dumps(extracted["sources"], ensure_ascii=False, sort_keys=True)
        result.at[row.name, "attribute_confidence"] = json.dumps(extracted["confidence"], ensure_ascii=False, sort_keys=True)
        extracted_rows.append(attrs)

    return result


def merge_attributes_from_columns(
    attrs: dict[str, str],
    sources: dict[str, str],
    confidence: dict[str, float],
    row: pd.Series,
    config: dict[str, Any],
) -> None:
    mappings = config.get("input_attribute_columns") or {}
    for attr, column in mappings.items():
        if should_keep_existing_attribute(str(attr), attrs.get(str(attr), "")):
            continue
        if column not in row or not str(row.get(column, "")).strip():
            continue
        value = normalize_attribute_value(str(row.get(column, "")).strip(), str(attr))
        if not value:
            continue
        attrs[str(attr)] = value
        sources[str(attr)] = f"column:{column}"
        confidence[str(attr)] = 0.95
        if str(attr) == "wymiary" and "dlugosc" not in attrs:
            for dimension_attr, dimension_value in extract_dimensions(value).items():
                if dimension_attr not in attrs:
                    attrs[dimension_attr] = dimension_value
                    sources[dimension_attr] = f"column:{column}"
                    confidence[dimension_attr] = 0.90
    apply_attribute_overrides(attrs, sources, confidence, row, config)


def should_keep_existing_attribute(attr: str, value: str) -> bool:
    if not value:
        return False
    if attr == "strumien":
        return True
    return False


def normalize_attribute_value(value: str, attr: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    if attr == "moc" and re.fullmatch(r"\d+(?:[,.]\d+)?", value):
        return f"{value.replace(',', '.')}W"
    if attr == "moc":
        max_match = re.fullmatch(r"max\s*(\d+(?:[,.]\d+)?)\s*W?", value, flags=re.IGNORECASE)
        if max_match:
            return f"max {normalize_number(max_match.group(1))}W"
        range_match = re.fullmatch(r"(\d+(?:[,.]\d+)?)\s*/\s*(\d+(?:[,.]\d+)?)", value)
        if range_match:
            return f"{normalize_number(range_match.group(1))}-{normalize_number(range_match.group(2))}W"
    if attr == "barwa" and re.fullmatch(r"\d{3,5}", value):
        return f"{value}K"
    if attr == "barwa":
        values = [int(v) for v in re.findall(r"[23645]\d{3}", value)]
        if len(set(values)) >= 2:
            return f"{min(values)}-{max(values)}K"
    if attr == "strumien" and re.fullmatch(r"\d+(?:[,.]\d+)?", value):
        return f"{value.replace(',', '.')}lm"
    if attr == "wymiary":
        return normalize_dimensions_value(value)
    if attr == "lm_w":
        match = re.search(r"(\d{2,3}(?:[,.]\d+)?)\s*(?:lm\s*/\s*W|lm/W)?", value, flags=re.IGNORECASE)
        if match:
            return f"{normalize_number(match.group(1))}lm/W"
    if attr == "ip" and re.fullmatch(r"\d{2}", value):
        return f"IP{value}"
    if attr == "kat_swiecenia":
        match = re.search(r"(\d{2,3})", value)
        if match:
            return f"{match.group(1)}\u00b0"
    if attr == "gwint":
        return normalize_socket(value)
    if attr == "czujnik":
        lowered = strip_accents(value).lower()
        if lowered in {"tak", "yes", "1", "true", "z czujnikiem"}:
            return "z czujnikiem ruchu"
        if lowered in {"nie", "no", "0", "false"}:
            return ""
        if "zmierzch" in lowered:
            return "z czujnikiem zmierzchu"
        if "czuj" in lowered or "pir" in lowered or "mikrofal" in lowered:
            return extract_sensor(value) or "z czujnikiem ruchu"
    return value


def extract_sensor_from_row(row: pd.Series, config: dict[str, Any]) -> str:
    columns = config.get("sensor_detection_columns") or ["Typ", "Nazwa B2C / SEO", "Nazwa Kanlux", "__working_title"]
    text = " ".join(str(row.get(column, "")) for column in columns if column in row and not pd.isna(row.get(column, "")))
    return extract_sensor(text)


def normalize_dimensions_value(value: str) -> str:
    value = compact_spaces(value)
    if not value:
        return ""
    if re.search(r"\b(?:cm|mm)\b", value, flags=re.IGNORECASE):
        return value
    match = re.fullmatch(r"(\d+(?:[,.]\d+)?)\s*[xX\u00d7]\s*(\d+(?:[,.]\d+)?)", value)
    if match:
        return f"{normalize_number(match.group(1))}x{normalize_number(match.group(2))}cm"
    match = re.fullmatch(r"(?:\u00d8|fi)?\s*(\d+(?:[,.]\d+)?)", value, flags=re.IGNORECASE)
    if match:
        return f"{normalize_number(match.group(1))}cm"
    return value


def apply_attribute_overrides(
    attrs: dict[str, str],
    sources: dict[str, str],
    confidence: dict[str, float],
    row: pd.Series,
    config: dict[str, Any],
) -> None:
    for rule in config.get("attribute_overrides") or []:
        attr = str(rule.get("attribute", ""))
        if not attr:
            continue
        current_value = str(attrs.get(attr, ""))
        category_column = str(rule.get("category_column", "Kategoria"))
        category = str(row.get(category_column, ""))
        if not contains_text(category, str(rule.get("category_contains", ""))):
            continue
        if not contains_text(current_value, str(rule.get("value_contains", ""))):
            continue
        attrs[attr] = str(rule.get("value", current_value))
        sources[attr] = f"{sources.get(attr, 'unknown')}+override"
        confidence[attr] = min(float(confidence.get(attr, 0.95)), 0.90)


def contains_text(value: str, needle: str) -> bool:
    if not needle:
        return True
    return strip_accents(needle).lower() in strip_accents(value).lower()


def main() -> None:
    parser = argparse.ArgumentParser(description="Wyciaga atrybuty z tytulow produktow.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet")
    parser.add_argument("--output", required=True)
    parser.add_argument("--title-column", required=True)
    parser.add_argument("--producer-column")
    parser.add_argument("--config", default="configs/categories/oprawy-sufitowe.yaml")
    parser.add_argument("--catalog-knowledge", default=DEFAULT_CATALOG_KNOWLEDGE_PATH)
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    config = enrich_config_with_catalog_knowledge(load_yaml(args.config), load_catalog_knowledge(args.catalog_knowledge))
    result = extract_attributes_for_dataframe(df, args.title_column, config, args.producer_column)
    write_products(result, Path(args.output))


if __name__ == "__main__":
    main()
