"""Uzupelnia nowe odcinki 1 m kabli na podstawie konfekcji 25/50/100 m.

Odcinek 1 m ma SKU bazowe (np. ``5907702810553/EKB``), a jego rodzenstwo to
ten sam SKU z dopiskiem dlugosci (``.../25m``, ``.../50m``, ``.../100m``) i z tym
samym EAN. Z rodzenstwa kopiowane sa atrybuty, klasa wysylkowa i tagi obecne
u wiekszosci konfekcji (tag grupy wariantow, kampanie, tagi paczkomatowe).
``DISABLE`` (wylaczony krazek 100 m) i tagi eventowe nie sa kopiowane nigdy.

Decyzje sklepu:
- ``Producent`` kabli to zawsze ``Polski Producent`` (nie NKT/Elektrokabel/Zamel),
- ``Dlugosc`` zapisujemy jak u rodzenstwa, ze spacja: ``1 m``,
- atrybuty, ktorych rodzenstwo nie ma (napiecie, izolacja, ekran, zyla
  ochronna), sa uzupelniane z tytulu tylko gdy sa puste.

Wejscie to pelny eksport WP All Export (XLSX). Wynik jest pelna kopia pliku -
zmieniaja sie wylacznie wiersze odcinkow 1 m. Skoroszyt kontrolny zawiera kazda
zmiane z ``source`` i ``status``.

Eksport bez kolumn ``Product Tags`` / ``hwp_product_gtin`` (szablon z kolumnami
``Attribute Value (pa_...)``, 2026-10-13) wymaga ``--tags-source`` - starszego
eksportu z tagami i EAN rodzenstwa. Nowe odcinki to wtedy szkice z SKU bazowym,
ktore maja konfekcje, a ich tagi zaczynaja sie od ``Domyslny dla grupy|Odcinek 1m``
(jak w partii z 2026-10-07). Plik wynikowy zawiera wtedy TYLKO nowe odcinki:
dopisane kolumny tagow i EAN bylyby puste dla reszty produktow, a WP All Import
wyczyscilby im tagi. Ponadto, gdy odcinek ich nie ma:
- kolumny ``Attribute Value (pa_...)`` i flagi atrybutu ida za ``Atrybut produktu:``,
- waga = waga konfekcji / dlugosc (na metr),
- opis = opis najdluzszej konfekcji (100 m zostaje w sklepie) bez dlugosci.

Uzycie:
    py src/fill_cable_1m_from_siblings.py --input "input/kable + 1m.xlsx"
    py src/fill_cable_1m_from_siblings.py --input input/Wszystkie-kable-przeglad-konfekcji.xlsx \\
        --tags-source "input/kable + 1m.xlsx"
"""

from __future__ import annotations

import argparse
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from shop_attribute_values import (
    DEFAULT_RULES as DEFAULT_VALUE_RULES,
    DEFAULT_SNAPSHOT,
    ValueCheck,
    check_features,
    extend_from_export,
    load_rules as load_value_rules,
    load_vocabulary,
    write_check_sheets,
)

ROOT = Path(__file__).resolve().parents[1]
TAG_SEPARATOR = "|"
SKU_COLUMN = "Sku"
TITLE_COLUMN = "Title"
TAGS_COLUMN = "Product Tags"
GTIN_COLUMN = "hwp_product_gtin"
SHIPPING_COLUMN = "Shipping Class"
WEIGHT_COLUMN = "Weight"
CONTENT_COLUMN = "Content"
ATTR_PREFIX = "Atrybut produktu: "
PA_NAME = "Attribute Name ({})"
PA_VALUE = "Attribute Value ({})"
PA_FLAGS = ("Attribute In Variations ({})", "Attribute Is Visible ({})", "Attribute Is Taxonomy ({})")
NEW_ROW_TAG = "Odcinek 1m"
# Tagi nowego odcinka, gdy eksport nie ma kolumny tagow (tak oznaczono partie z 2026-10-07).
NEW_ROW_BASE_TAGS = ["Domyślny dla grupy", NEW_ROW_TAG]
# Opis krotszy niz tyle znakow traktujemy jako pusty (np. sama spacja z edytora).
MIN_CONTENT = 20
# DISABLE wylacza krazek 100 m, ElektroTargi to tag eventowy - zadnego nie dziedziczymy.
IGNORED_TAGS = {"DISABLE", "ElektroTargi Alkan 2026"}

PRODUCER = "Polski Producent"
LENGTH_VALUE = "1 m"
LENGTH_ATTR = "Długość"
PRODUCER_ATTR = "Producent"
# Kable YDY/YDYp/OWY/OMY maja izolacje i powloke z PVC i nie sa ekranowane.
PVC_TYPES = ("YDY", "OWY", "OMY")

SIBLING_SKU = re.compile(r"^(?P<base>.+)/(?P<length>\d+)\s*m$", re.IGNORECASE)
ZO_PATTERN = re.compile(r"żo", re.IGNORECASE)
COLOR_PATTERN = re.compile(r"biał|czarn|szar|niebiesk|brąz|czerwon|zielon|żółt|pomarańcz", re.IGNORECASE)
VOLTAGE_PATTERN = re.compile(r"(\d{3})(?:\s*/\s*(\d{3}))?\s*V?\b")


@dataclass
class Change:
    row_id: str
    sku: str
    title: str
    field: str
    old: str
    new: str
    source: str
    status: str
    note: str = ""


def text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def split_tags(value: Any) -> list[str]:
    return [part.strip() for part in text(value).split(TAG_SEPARATOR) if part.strip()]


def is_new_row(row: dict[str, Any]) -> bool:
    return NEW_ROW_TAG in split_tags(row.get(TAGS_COLUMN))


def select_new_rows(rows: list[dict[str, Any]], has_tags: bool) -> list[dict[str, Any]]:
    """Po tagu ``Odcinek 1m``, a bez kolumny tagow: szkic z SKU bazowym, ktory ma konfekcje."""
    if has_tags:
        return [r for r in rows if is_new_row(r)]
    return [
        r for r in rows
        if text(r.get("Status")) == "draft" and not SIBLING_SKU.match(text(r.get(SKU_COLUMN))) and find_siblings(r, rows)
    ]


def sibling_length(sibling: dict[str, Any]) -> int:
    match = SIBLING_SKU.match(text(sibling.get(SKU_COLUMN)))
    return int(match.group("length")) if match else 0


def weight_per_meter(siblings: list[dict[str, Any]]) -> tuple[str, bool]:
    """Waga 1 m z konfekcji; druga wartosc mowi, czy konfekcje sa zgodne (do 2%)."""
    values = []
    for sibling in siblings:
        try:
            weight = float(text(sibling.get(WEIGHT_COLUMN)) or 0)
        except ValueError:
            continue
        length = sibling_length(sibling)
        if weight > 0 and length:
            values.append(weight / length)
    if not values:
        return "", True
    values.sort()
    median = values[len(values) // 2]
    consistent = all(abs(v - median) <= 0.02 * median for v in values)
    return f"{median:.3f}".rstrip("0").rstrip("."), consistent


COPPER_KG_PER_MM2_M = 0.0089  # gestosc miedzi 8,9 g/cm3 -> 1 mm2 x 1 m = 8,9 g


def copper_kg_per_meter(cores: str, cross_section: str) -> float:
    """Dolna granica wagi 1 m: sama miedz zyl. Waga ponizej tej wartosci jest bledna."""
    try:
        return int(cores) * float(cross_section.replace(",", ".")) * COPPER_KG_PER_MM2_M
    except ValueError:
        return 0.0


# Punkt specyfikacji z dlugoscia krazka/bebna nie pasuje do odcinka 1 m.
LENGTH_ITEM = re.compile(r"[ \t]*<li>(?:(?!</li>).)*Długość:(?:(?!</li>).)*</li>\n?", re.IGNORECASE | re.DOTALL)
LENGTH_TOKEN = re.compile(r"\s(?:25|50|100)\s?m\b")


def content_for_1m(content: str) -> tuple[str, list[str]]:
    """Opis konfekcji bez jej dlugosci ("... bialy 100m?" -> "... bialy?")."""
    notes = []
    new, count = LENGTH_ITEM.subn("", content)
    if count:
        notes.append(f"usuniety punkt 'Długość' ({count})")
    new, count = LENGTH_TOKEN.subn("", new)
    if count:
        notes.append(f"usunieta dlugosc konfekcji ({count}x)")
    return new, notes


def find_siblings(row: dict[str, Any], rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    base = text(row.get(SKU_COLUMN)).lower()
    found = []
    for other in rows:
        match = SIBLING_SKU.match(text(other.get(SKU_COLUMN)))
        if match and match.group("base").lower() == base:
            found.append(other)
    return found


def majority(values: list[str]) -> tuple[str, bool]:
    """Najczestsza niepusta wartosc i informacja, czy rodzenstwo jest zgodne."""
    filled = [v for v in values if v]
    if not filled:
        return "", True
    counts = Counter(filled)
    value, _ = counts.most_common(1)[0]
    return value, len(counts) == 1


def common_tags(siblings: list[dict[str, Any]]) -> list[str]:
    """Tagi obecne u wiekszosci rodzenstwa, w kolejnosci pierwszego wystapienia.

    Wiekszosc, a nie czesc wspolna: tag grupy bywa niespojny w samym rodzenstwie
    (25/50 m "Przewod YDY", 100 m "Przewod YDYzo") i wtedy wygrywa grupa wiekszosci.
    """
    counts: Counter[str] = Counter()
    order: list[str] = []
    for sibling in siblings:
        for tag in dict.fromkeys(split_tags(sibling.get(TAGS_COLUMN))):
            if tag not in counts:
                order.append(tag)
            counts[tag] += 1
    return [tag for tag in order if tag not in IGNORED_TAGS and counts[tag] * 2 > len(siblings)]


def cable_type_in_title(title: str) -> str:
    upper = title.upper()
    for cable_type in PVC_TYPES:
        if re.search(rf"\b{cable_type}", upper):
            return cable_type
    return ""


def voltage_from_title(title: str) -> str:
    """'450/750' -> 750, '300/500V' -> 500, '500 V' -> 500 (wyzsza wartosc pary)."""
    for match in VOLTAGE_PATTERN.finditer(title):
        first, second = match.group(1), match.group(2)
        if not second and not re.search(rf"{first}\s*V", title):
            continue
        value = max(int(first), int(second or first))
        if value in (300, 500, 750):
            return str(value)
    return ""


def title_derived(title: str) -> dict[str, str]:
    values: dict[str, str] = {}
    if cable_type_in_title(title):
        values["Typ izolacji"] = "PVC"
        values["Ekranowany"] = "Nie"
    voltage = voltage_from_title(title)
    if voltage:
        values["Napięcie [V]"] = voltage
    if ZO_PATTERN.search(title):
        values["Żyła ochronna"] = "Tak"
    return values


def fix_title(title: str, color: str = "") -> tuple[str, list[str]]:
    """Ujednolica tytul odcinka do zapisu uzywanego w sklepie.

    - przekroj zapisany zwyklym 'x' ("3x1,5"), jak u rodzenstwa i w wyszukiwarce,
      a nie '×' ani 'X',
    - napiecie ma jednostke ("450/750V", nie "450/750"),
    - kolor z atrybutu, gdy tytul go nie ma,
    - na koncu dlugosc '1m', a nie nazwa producenta.
    """
    notes: list[str] = []
    new = re.sub(r"\s+", " ", title).strip()

    stripped = re.sub(r"\s+Elektrokabel$", "", new)
    if stripped != new:
        new = stripped
        notes.append("usunieta nazwa producenta")
    if not re.search(r"\b1\s*m$", new):
        new = f"{new} 1m"
        notes.append("dopisana dlugosc 1m")

    crossed = re.sub(r"(\d)\s*[×X]\s*(\d)", r"\1x\2", new)
    if crossed != new:
        new = crossed
        notes.append("przekroj zapisany 'x'")

    with_unit = re.sub(r"\b(\d{3}/\d{3})(?!\s*V)\b", r"\1V", new)
    if with_unit != new:
        new = with_unit
        notes.append("napiecie z jednostka V")

    if color and "|" not in color and not COLOR_PATTERN.search(new):
        new = re.sub(r"\s+1m$", f" {color.lower()} 1m", new)
        notes.append(f"dopisany kolor '{color.lower()}'")
    return new, notes


def title_warnings(title: str) -> list[str]:
    """Niezgodnosci, ktorych nie poprawiamy automatycznie - do sprawdzenia u dostawcy."""
    warnings = []
    if re.search(r"\bOWY\b", title) and re.search(r"\dx0,5\b", title.replace("×", "x")):
        warnings.append("OWY (H05VV-F) normowo zaczyna sie od 0,75 mm2; 0,5 mm2 to zwykle H03VV-F OMY - sprawdz typ u dostawcy")
    return warnings


def decide_row(row: dict[str, Any], siblings: list[dict[str, Any]], attr_columns: list[str]) -> list[Change]:
    row_id = text(row.get("id"))
    sku = text(row.get(SKU_COLUMN))
    title = text(row.get(TITLE_COLUMN))
    changes: list[Change] = []

    def add(field: str, new: str, source: str, status: str, note: str = "") -> None:
        old = text(row.get(field))
        if old != new:
            changes.append(Change(row_id, sku, title, field, old, new, source, status, note))

    if not siblings:
        changes.append(Change(row_id, sku, title, "-", "", "", "brak", "DO_SPRAWDZENIA", "brak konfekcji 25/50/100 m"))
        return changes

    color_column = ATTR_PREFIX + "Kolor"
    color = text(row.get(color_column)) or majority([text(s.get(color_column)) for s in siblings])[0]
    new_title, notes = fix_title(title, color)
    if new_title != title:
        add(TITLE_COLUMN, new_title, "regula", "OK", ", ".join(notes))
    for warning in title_warnings(title):
        changes.append(Change(row_id, sku, title, "uwaga", "", "", "regula", "DO_SPRAWDZENIA", warning))

    derived = title_derived(title)
    for column in attr_columns:
        name = column[len(ATTR_PREFIX):]
        own = text(row.get(column))
        if name == LENGTH_ATTR:
            add(column, LENGTH_VALUE, "regula", "OK")
            continue
        if name == PRODUCER_ATTR:
            add(column, PRODUCER, "decyzja uzytkownika", "OK")
            continue
        value, unanimous = majority([text(s.get(column)) for s in siblings])
        if value:
            status = "OK" if unanimous else "WARNING"
            note = "" if unanimous else "rodzenstwo niezgodne - wybrana wartosc wiekszosci"
            if own and own != value:
                status, note = "WARNING", "wartosc odcinka rozna od rodzenstwa - wygrywa rodzenstwo"
            add(column, value, "rodzenstwo", status, note)
            continue
        if name in derived:
            if not own:
                add(column, derived[name], "tytul", "OK")
            elif name == "Typ izolacji" and own != derived[name]:
                add(column, derived[name], "tytul", "WARNING", f"{cable_type_in_title(title)} ma izolacje PVC")

    shipping, unanimous = majority([text(s.get(SHIPPING_COLUMN)) for s in siblings])
    if shipping and not text(row.get(SHIPPING_COLUMN)):
        add(SHIPPING_COLUMN, shipping, "rodzenstwo", "OK" if unanimous else "WARNING")

    gtin, unanimous = majority([text(s.get(GTIN_COLUMN)) for s in siblings])
    if gtin and not text(row.get(GTIN_COLUMN)):
        add(GTIN_COLUMN, gtin, "rodzenstwo", "OK" if unanimous else "WARNING")

    if WEIGHT_COLUMN in row and not float(text(row.get(WEIGHT_COLUMN)) or 0):
        weight, consistent = weight_per_meter(siblings)
        def attr(name: str) -> str:
            column = ATTR_PREFIX + name
            return text(row.get(column)) or majority([text(s.get(column)) for s in siblings])[0]
        copper = copper_kg_per_meter(attr("Liczba żył"), attr("Przekrój żyły [mm]"))
        if weight and copper and float(weight) < copper:
            changes.append(Change(row_id, sku, title, "uwaga", "", "", "rodzenstwo", "DO_SPRAWDZENIA",
                                  f"Weight: waga konfekcji daje {weight} kg/m, a sama miedz zyl to {copper:.3f} kg/m - "
                                  "waga rodzenstwa bledna, nie wpisano"))
        elif weight:
            add(WEIGHT_COLUMN, weight, "rodzenstwo", "OK" if consistent else "WARNING",
                "waga konfekcji / dlugosc" + ("" if consistent else " - konfekcje niezgodne, mediana"))

    if CONTENT_COLUMN in row and len(text(row.get(CONTENT_COLUMN))) < MIN_CONTENT:
        donors = [s for s in siblings if len(text(s.get(CONTENT_COLUMN))) >= MIN_CONTENT]
        if donors:
            donor = max(donors, key=sibling_length)
            content, notes = content_for_1m(text(donor.get(CONTENT_COLUMN)))
            add(CONTENT_COLUMN, content, f"rodzenstwo {sibling_length(donor)} m", "OK", ", ".join(notes))

    # Bez kolumny tagow (None) nie wiemy, co odcinek ma w sklepie - dostaje tagi partii 2026-10-07.
    own_tags = split_tags(row.get(TAGS_COLUMN)) if row.get(TAGS_COLUMN) is not None else list(NEW_ROW_BASE_TAGS)
    merged = own_tags + [tag for tag in common_tags(siblings) if tag not in own_tags]
    add(TAGS_COLUMN, TAG_SEPARATOR.join(merged), "rodzenstwo", "OK", "tagi wspolne dla 25/50/100 m")
    for tag in split_tags_minority(siblings):
        changes.append(Change(row_id, sku, title, "uwaga", "", "", "rodzenstwo", "WARNING",
                              f"tag '{tag}' ma tylko czesc konfekcji - pominiety (wiekszosc)"))
    return changes


def split_tags_minority(siblings: list[dict[str, Any]]) -> list[str]:
    """Tagi obecne u czesci rodzenstwa, ale nie u wiekszosci - do raportu."""
    counts: Counter[str] = Counter(tag for s in siblings for tag in dict.fromkeys(split_tags(s.get(TAGS_COLUMN))))
    return [tag for tag, n in counts.items() if tag not in IGNORED_TAGS and n * 2 <= len(siblings)]


def mirror_taxonomy_columns(changes: list[Change], row: dict[str, Any], siblings: list[dict[str, Any]],
                            slug_by_name: dict[str, str]) -> list[Change]:
    """Zmiany ``Atrybut produktu: X`` powtorzone w ``Attribute Value (pa_x)`` z flagami jak u rodzenstwa."""
    extra: list[Change] = []
    for change in changes:
        if not change.field.startswith(ATTR_PREFIX):
            continue
        slug = slug_by_name.get(change.field[len(ATTR_PREFIX):])
        if not slug:
            continue
        targets = [(PA_VALUE.format(slug), change.new)]
        for flag in PA_FLAGS:
            column = flag.format(slug)
            value = majority([text(s.get(column)) for s in siblings if text(s.get(PA_VALUE.format(slug)))])[0] or "yes"
            targets.append((column, value))
        for column, new in targets:
            old = text(row.get(column))
            if old != new:
                extra.append(Change(change.row_id, change.sku, change.title, column, old, new, change.source, change.status, change.note))
    return extra


def duplicate_gtin_rows(row: dict[str, Any], siblings: list[dict[str, Any]], rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    gtin = text(row.get(GTIN_COLUMN))
    if not gtin:
        return []
    skip = {id(row)} | {id(s) for s in siblings}
    return [o for o in rows if id(o) not in skip and text(o.get(GTIN_COLUMN)) == gtin]


def check_shop_values(changes: list[Change], header: list[str], rows: list[dict[str, Any]],
                      new_rows: list[dict[str, Any]], snapshot: Path = DEFAULT_SNAPSHOT,
                      rules_path: Path = DEFAULT_VALUE_RULES) -> list[ValueCheck]:
    """Raport: czy wartosci wpisane odcinkom 1 m tworza w Woo nowe terminy (slownik = zrzut + reszta eksportu)."""
    vocabulary, rules = load_vocabulary(snapshot), load_value_rules(rules_path)
    new_ids = {id(row) for row in new_rows}
    extend_from_export(vocabulary, header, [row for row in rows if id(row) not in new_ids], rules)
    per_row: dict[tuple[str, str, str], dict[str, str]] = {}
    for change in changes:
        if change.field.startswith(ATTR_PREFIX) and change.new:
            per_row.setdefault((change.row_id, change.sku, change.title), {})[change.field[len(ATTR_PREFIX):]] = change.new
    checks: list[ValueCheck] = []
    for (_, sku, title), features in per_row.items():
        checks.extend(check_features(sku, title, features, vocabulary, rules, keep_new=True)[1])
    return checks


def read_rows(path: Path) -> tuple[list[str], list[dict[str, Any]]]:
    from openpyxl import load_workbook

    sheet = load_workbook(path, read_only=True).worksheets[0]
    values = list(sheet.iter_rows(values_only=True))
    header = [text(h) for h in values[0]]
    return header, [dict(zip(header, r)) for r in values[1:]]


def load_tag_source(path: Path) -> dict[str, dict[str, Any]]:
    """id -> tagi i EAN ze starszego eksportu (dla eksportu bez tych kolumn)."""
    _, rows = read_rows(path)
    return {text(r.get("id")): {TAGS_COLUMN: r.get(TAGS_COLUMN), GTIN_COLUMN: r.get(GTIN_COLUMN)} for r in rows}


def taxonomy_slugs(header: list[str], rows: list[dict[str, Any]]) -> dict[str, str]:
    """Nazwa atrybutu ('Długość') -> slug kolumn 'Attribute ... (pa_dlugosc)'."""
    slugs: dict[str, str] = {}
    for column in header:
        match = re.fullmatch(r"Attribute Name \((?P<slug>.+)\)", column)
        if not match:
            continue
        name = next((text(r.get(column)) for r in rows if text(r.get(column))), "")
        if name:
            slugs.setdefault(name, match.group("slug"))
    return slugs


def write_new_rows(output_path: Path, header: list[str], new_rows: list[dict[str, Any]], changes: list[Change]) -> None:
    """Sam naglowek i nowe odcinki - plik importu bez ryzyka dla reszty produktow."""
    from openpyxl import Workbook

    by_row: dict[str, dict[str, str]] = {}
    for change in changes:
        if change.field in header:
            value: Any = float(change.new) if change.field == WEIGHT_COLUMN else change.new
            by_row.setdefault(change.row_id, {})[change.field] = value
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(header)
    for row in new_rows:
        updates = by_row.get(text(row.get("id")), {})
        sheet.append([updates.get(h, row.get(h)) for h in header])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def write_output(input_path: Path, output_path: Path, header: list[str], changes: list[Change]) -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(input_path)
    sheet = workbook.worksheets[0]
    col = {name: idx + 1 for idx, name in enumerate(header)}
    row_by_id = {text(sheet.cell(r, col["id"]).value): r for r in range(2, sheet.max_row + 1)}
    for change in changes:
        if change.field in col:
            sheet.cell(row_by_id[change.row_id], col[change.field]).value = change.new
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def write_control(path: Path, changes: list[Change], duplicates: list[dict[str, str]],
                  value_checks: list[ValueCheck] | None = None) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    fills = {"WARNING": "FFF2CC", "DO_SPRAWDZENIA": "F8CBAD"}
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Zmiany"
    sheet.append(["id", "SKU", "Tytul", "Pole", "Bylo", "Jest", "source", "status", "Uwagi"])
    for c in changes:
        sheet.append([c.row_id, c.sku, c.title, c.field.replace(ATTR_PREFIX, ""), c.old, c.new, c.source, c.status, c.note])
        if c.status in fills:
            for cell in sheet[sheet.max_row]:
                cell.fill = PatternFill("solid", fgColor=fills[c.status])
    dup = workbook.create_sheet("Ten sam EAN")
    dup.append(["id odcinka 1 m", "SKU odcinka", "id innego produktu", "Status", "SKU", "Tytul"])
    for d in duplicates:
        dup.append(list(d.values()))
    if value_checks:
        write_check_sheets(workbook, value_checks)
    for ws in workbook.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        for column in ws.columns:
            width = max(len(text(cell.value)) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(max(width + 2, 8), 70)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def main(argv: list[str] | None = None) -> int:
    today = date.today().isoformat()
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", type=Path, default=ROOT / "output" / f"kable_1m_UZUPELNIONE_{today}.xlsx")
    parser.add_argument("--control", type=Path, default=ROOT / "reports" / f"kable_1m_kontrola_{today}.xlsx")
    parser.add_argument("--shop-terms", type=Path, default=DEFAULT_SNAPSHOT,
                        help="Zrzut terminow Woo (scripts/dump_woo_attribute_terms.py)")
    parser.add_argument("--tags-source", type=Path,
                        help="Starszy eksport z 'Product Tags' i 'hwp_product_gtin', gdy --input ich nie ma")
    args = parser.parse_args(argv)

    if args.output.resolve() == args.input.resolve():
        raise SystemExit("Plik wynikowy nie moze nadpisac pliku wejsciowego.")

    header, rows = read_rows(args.input)
    has_tags = TAGS_COLUMN in header
    added_columns = [c for c in (TAGS_COLUMN, GTIN_COLUMN) if c not in header]
    if added_columns:
        if not args.tags_source:
            raise SystemExit(f"Eksport nie ma kolumn {added_columns} - podaj --tags-source ze starszym eksportem.")
        source = load_tag_source(args.tags_source)
        for row in rows:
            for column in added_columns:
                row[column] = source.get(text(row.get("id")), {}).get(column)
        header = header + added_columns
    attr_columns = [h for h in header if h.startswith(ATTR_PREFIX)]
    slug_by_name = taxonomy_slugs(header, rows)
    changes: list[Change] = []
    duplicates: list[dict[str, str]] = []
    new_rows = select_new_rows(rows, has_tags)
    for row in new_rows:
        siblings = find_siblings(row, rows)
        row_changes = decide_row(row, siblings, attr_columns)
        changes.extend(row_changes + mirror_taxonomy_columns(row_changes, row, siblings, slug_by_name))
        for other in duplicate_gtin_rows(row, siblings, rows):
            duplicates.append({
                "id": text(row.get("id")), "sku": text(row.get(SKU_COLUMN)), "other_id": text(other.get("id")),
                "status": text(other.get("Status")), "other_sku": text(other.get(SKU_COLUMN)), "title": text(other.get(TITLE_COLUMN)),
            })

    if added_columns:
        write_new_rows(args.output, header, new_rows, changes)
    else:
        write_output(args.input, args.output, header, changes)
    value_checks = check_shop_values(changes, header, rows, new_rows, snapshot=args.shop_terms)
    write_control(args.control, changes, duplicates, value_checks)
    statuses = Counter(c.status for c in changes)
    new_values = sorted({(c.attribute, c.value) for c in value_checks if c.status == "DO_SPRAWDZENIA"})
    if new_values:
        print(f"UWAGA: {len(new_values)} wartosci utworzy w Woo nowe terminy (arkusz 'Nowe wartosci'): "
              + "; ".join(f"{a}={v}" for a, v in new_values[:10]))
    print(f"Odcinkow 1 m: {len(new_rows)}, zmian: {len(changes)} {dict(statuses)}, ten sam EAN: {len(duplicates)}")
    print(f"Wynik: {args.output}")
    print(f"Kontrola: {args.control}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
