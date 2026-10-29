"""Uzupelnia wymiary w eksporcie Woo paneli LED i akcesoriow do paneli.

Zrodla (od najsilniejszego):
  1. Parametry Kanlux (archived_input_files/Parametry.xlsx) - Dlugosc/Szerokosc/
     Wysokosc/Srednica [mm] po kodzie Kanlux (zrodlo prawdy dla Kanlux).
  2. Istniejace komorki eksportu (Wymiary [mm], Wysokosc, Dlugosc).
  3. Nazwa produktu ("595x595mm", "60x60", "30x120", "17 cm" dla okraglych,
     "63MM", "H70").

Format jak w pipeline (export_to_baselinker_csv): "Wymiary [mm]" = "600 x 600"
(dluzszy bok pierwszy, bez jednostki), pojedyncze wymiary "600 mm", przecinek
dziesietny. Produkty okragle dostaja "Srednica" (istniejacy atrybut Woo),
a "Wymiary [mm]" zostaje puste - tak jak w danych Kanlux.
Woo Length/Width/Height (cm) sa wypelniane tylko gdy sa puste/0.
Akcesoria (linki, klipsy, uchwyty, zasilacze bez parametrow) nie dostaja
wymiarow z frazy "do paneli 60x60" - to format panelu, nie produktu.

Plik wejsciowy nie jest nadpisywany.
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import re
from pathlib import Path

import openpyxl
import yaml
from openpyxl.styles import PatternFill

COL_DIM = "Atrybut produktu: Wymiary [mm]"
COL_H = "Atrybut produktu: Wysokość"
COL_L = "Atrybut produktu: Długość"
COL_D = "Atrybut produktu: Średnica"
WOO = ("Length", "Width", "Height")
ACCESSORY_RE = re.compile(r"^(Linka|Klips|Uchwyt|Zasilacz)\b", re.I)
NUM = r"(\d+(?:[.,]\d+)?)"


def num(text) -> float | None:
    if text is None or text == "":
        return None
    try:
        return float(str(text).replace(",", "."))
    except ValueError:
        return None


def fmt(n: float) -> str:
    s = f"{n:.2f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def cm(mm: float):
    v = round(mm / 10, 2)
    return int(v) if float(v).is_integer() else v


def to_mm(value: float, unit: str | None, all_values: list[float]) -> float:
    unit = (unit or "").lower()
    if unit == "cm":
        return value * 10
    if unit == "m":
        return value * 1000
    if unit == "mm":
        return value
    # brak jednostki: nazwy paneli podaja "60x60", "30x120", "10x10" w cm
    return value * 10 if max(all_values) <= 200 else value


def parse_single_mm(text) -> float | None:
    if text is None:
        return None
    m = re.fullmatch(r"\s*" + NUM + r"\s*(mm|cm|m)?\s*", str(text), re.I)
    if not m:
        return None
    v = num(m.group(1))
    unit = (m.group(2) or "mm").lower()
    return v * {"mm": 1, "cm": 10, "m": 1000}[unit]


def parse_dims_cell(text, title_footprint: list[float] | None = None) -> list[float] | None:
    """Kolumna jest w mm; wyjatek: wartosc bez jednostki rowna wymiarowi z nazwy
    zapisanemu w cm (Kobi "60x60", "30x120") jest traktowana jako cm."""
    if not text:
        return None
    t = str(text).replace("×", "x").replace("X", "x")
    m = re.fullmatch(r"\s*" + NUM + r"\s*x\s*" + NUM + r"(?:\s*x\s*" + NUM + r")?\s*(mm|cm)?\s*", t, re.I)
    if not m:
        return None
    vals = [num(g) for g in m.groups()[:3] if g]
    unit = (m.group(4) or "").lower()
    if not unit:
        unit = "mm"
        if title_footprint and sorted((v * 10 for v in vals[:2]), reverse=True) == title_footprint:
            unit = "cm"
    return [v * (10 if unit == "cm" else 1) for v in vals]


def parse_title(title: str) -> dict:
    out: dict = {}
    t = title.replace("×", "x")
    # wymiar AxB z opcjonalna jednostka (ignoruj "60x60/62x62" w akcesoriach - obsluga wyzej)
    m = re.search(r"(?<![\w/])" + NUM + r"\s*[xX]\s*" + NUM + r"\s*(mm|cm)?(?![\w/])", t)
    if m:
        vals = [num(m.group(1)), num(m.group(2))]
        out["footprint"] = sorted((to_mm(v, m.group(3), vals) for v in vals), reverse=True)
    if re.search(r"okrąg", t, re.I):
        m = re.search(r"(?<![\dx,])" + NUM + r"\s*(cm|mm)\b", t, re.I)
        if m:
            out["diameter"] = to_mm(num(m.group(1)), m.group(2), [num(m.group(1))])
    m = re.search(r"(?<![x\d,])(\d{2,3})\s*mm\b(?!\s*x)", t, re.I)
    if m and "footprint" in out and not re.search(r"\d\s*x\s*\d+\s*mm", t[: m.end()][-12:], re.I):
        out["height"] = float(m.group(1))
    m = re.search(r"\bH(\d{2,3})\b", t)
    if m:
        out["height"] = float(m.group(1))
    return out


def load_kanlux_params(path: Path, codes: set[str]) -> dict:
    ws = openpyxl.load_workbook(path, read_only=True).active
    data: dict = collections.defaultdict(dict)
    for r in ws.iter_rows(min_row=2, values_only=True):
        c = str(r[0])
        if c in codes and r[3] in {"Długość [mm]", "Szerokość [mm]", "Wysokość [mm]", "Średnica [mm]"}:
            data[c][r[3]] = num(r[4])
    return data


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", default="input/Panele-Led.xlsx")
    ap.add_argument("--kanlux-params", default="archived_input_files/Parametry.xlsx")
    ap.add_argument("--verified", default="dictionaries/panel_dimensions_verified.yaml",
                    help="wymiary potwierdzone w kartach producentow (zrodlo nadrzedne)")
    today = dt.date.today().isoformat()
    ap.add_argument("--output", default=f"output/Panele-Led_WYMIARY_UZUPELNIONE_{today}.xlsx")
    ap.add_argument("--report", default=f"reports/Panele-Led_WYMIARY_KONTROLA_{today}.xlsx")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.input)
    ws = wb.active
    header = [c.value for c in ws[1]]
    if COL_D not in header:
        ws.cell(row=1, column=len(header) + 1, value=COL_D)
        header.append(COL_D)
    col = {name: header.index(name) + 1 for name in header}

    kan_codes = {str(ws.cell(r, col["Sku"]).value).split("/")[0]
                 for r in range(2, ws.max_row + 1)
                 if str(ws.cell(r, col["Sku"]).value).endswith("/KAN")}
    params = load_kanlux_params(Path(args.kanlux_params), kan_codes)
    verified_path = Path(args.verified)
    verified = (yaml.safe_load(verified_path.read_text(encoding="utf-8")) or {}).get("products", {}) if verified_path.exists() else {}

    changed_fill = PatternFill("solid", fgColor="FFF2CC")
    report = []
    for r in range(2, ws.max_row + 1):
        get = lambda name: ws.cell(r, col[name]).value
        title, sku = str(get("Title") or ""), str(get("Sku") or "")
        code = sku.split("/")[0]
        accessory = bool(ACCESSORY_RE.match(title))
        notes, sources = [], []
        status = "OK"

        old = {k: get(k) for k in (COL_DIM, COL_H, COL_L, COL_D) + WOO}
        tparse = {} if accessory else parse_title(title)
        cell_dims = parse_dims_cell(old[COL_DIM], tparse.get("footprint"))
        round_shape = bool(re.search(r"okrąg", title, re.I))

        footprint = diameter = height = length = None
        from_params = False
        p = params.get(code) if sku.endswith("/KAN") else None
        if p:
            if p.get("Długość [mm]") and p.get("Szerokość [mm]"):
                footprint = sorted([p["Długość [mm]"], p["Szerokość [mm]"]], reverse=True)
                from_params = True
            elif p.get("Długość [mm]"):
                length = p["Długość [mm]"]
            if p.get("Średnica [mm]") and round_shape:
                diameter = p["Średnica [mm]"]
            height = p.get("Wysokość [mm]")
            if footprint or diameter or height or length:
                sources.append("Parametry Kanlux")

        v = verified.get(sku)
        if v:
            from_params = True
            footprint = sorted(v["footprint"], reverse=True) if v.get("footprint") else footprint
            diameter = v.get("diameter", diameter)
            height = v.get("height", height)
            sources.append(f"karta producenta: {v['source']}")
            status = v.get("status", "OK")
            if v.get("note"):
                notes.append(v["note"])

        if not footprint and not diameter and cell_dims and not accessory:
            footprint = sorted(cell_dims[:2], reverse=True)
            if len(cell_dims) == 3 and not height:
                height = cell_dims[2]
            sources.append("komórka Wymiary")
        if not footprint and not diameter and tparse.get("footprint") and not round_shape:
            footprint = tparse["footprint"]
            sources.append("nazwa")
            status = "WARNING"
            notes.append("wymiar nominalny z nazwy")
        if round_shape and not diameter:
            d = tparse.get("diameter") or parse_single_mm(old[COL_D]) or parse_single_mm(old[COL_L])
            if d:
                diameter = d
                sources.append("nazwa/Długość (średnica)")
            elif tparse.get("footprint"):
                diameter = tparse["footprint"][0]
                sources.append("nazwa (średnica z AxA)")
        if not height:
            h = parse_single_mm(old[COL_H])
            if h:
                height = h
                sources.append("komórka Wysokość")
            elif tparse.get("height"):
                height = tparse["height"]
                sources.append("nazwa (wysokość)")
                status = "WARNING"
                notes.append("wysokość odczytana z nazwy")
        if not length:
            length = footprint[0] if footprint else (parse_single_mm(old[COL_L]) if not round_shape or old[COL_L] else None)

        new = dict(old)
        if round_shape and diameter:
            new[COL_D] = f"{fmt(diameter)} mm"
            if old[COL_DIM] and cell_dims:
                notes.append(f"okrągły - Wymiary '{old[COL_DIM]}' zastąpione Średnicą")
                new[COL_DIM] = None
            if old[COL_L]:
                new[COL_L] = f"{fmt(parse_single_mm(old[COL_L]) or diameter)} mm"
        elif footprint:
            new[COL_DIM] = f"{fmt(footprint[0])} x {fmt(footprint[1])}"
        if not (round_shape and diameter) and length:
            new[COL_L] = f"{fmt(length)} mm"
        if height:
            new[COL_H] = f"{fmt(height)} mm"

        if accessory and cell_dims and not from_params and max(cell_dims) >= 290:
            status = "DO_SPRAWDZENIA"
            notes.append(f"akcesorium ma Wymiary '{old[COL_DIM]}' - to format panelu, nie produktu; zostawione")

        # Woo Length/Width/Height (cm) tylko gdy puste/0
        plan = None
        if footprint:
            plan = (footprint[0], footprint[1])
        elif round_shape and diameter:
            plan = (diameter, diameter)
        woo_existing = [num(old[k]) or 0 for k in WOO]
        if plan and not any(woo_existing):
            new["Length"], new["Width"] = cm(plan[0]), cm(plan[1])
            if height:
                new["Height"] = cm(height)
        elif plan and any(woo_existing):
            L, W = woo_existing[0], woo_existing[1]
            if abs(max(L, W) * 10 - plan[0]) > 30 or abs(min(L, W) * 10 - plan[1]) > 30:
                notes.append(f"Woo L/W/H {L}/{W}/{woo_existing[2]} cm nie pasuje do wymiaru - nie zmieniane")
                if status == "OK":
                    status = "WARNING"

        # zmiany istniejacych wartosci (poza sama normalizacja formatu)
        def comparable(v):
            d = parse_dims_cell(v, tparse.get("footprint"))
            if d:
                return tuple(sorted(d[:2], reverse=True))
            s = parse_single_mm(v)
            return s
        for k in (COL_DIM, COL_H, COL_L):
            if old[k] and new[k] and comparable(old[k]) != comparable(new[k]):
                notes.append(f"{k.split(': ')[1]}: '{old[k]}' -> '{new[k]}'")
                if status == "OK" and not v:
                    status = "WARNING"
        if not height and not accessory:
            notes.append("brak wysokości w danych")
            if status == "OK":
                status = "DO_SPRAWDZENIA"
        if not footprint and not diameter and not accessory:
            notes.append("brak wymiaru")
            status = "DO_SPRAWDZENIA"

        for k, v in new.items():
            if v != old[k]:
                c = ws.cell(r, col[k])
                c.value = v
                c.fill = changed_fill
        report.append({
            "id": get("id"), "Sku": sku, "Title": title,
            **{f"stare {k.split(': ')[-1]}": old[k] for k in (COL_DIM, COL_H, COL_L)},
            **{f"nowe {k.split(': ')[-1]}": new[k] for k in (COL_DIM, COL_H, COL_L, COL_D)},
            "Woo L/W/H przed": "/".join(str(old[k]) for k in WOO),
            "Woo L/W/H po": "/".join(str(new[k]) for k in WOO),
            "source": ", ".join(dict.fromkeys(sources)) or "-",
            "confidence": {"OK": 0.95, "WARNING": 0.8, "DO_SPRAWDZENIA": 0.5}[status],
            "status": status, "uwagi": "; ".join(notes),
        })

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)
    rb = openpyxl.Workbook()
    rs = rb.active
    rs.title = "Wymiary"
    rs.append(list(report[0].keys()))
    for row in report:
        rs.append(list(row.values()))
    rs.freeze_panes = "D2"
    rs.auto_filter.ref = rs.dimensions
    Path(args.report).parent.mkdir(parents=True, exist_ok=True)
    rb.save(args.report)
    cnt = collections.Counter(x["status"] for x in report)
    print(f"rows={len(report)} {dict(cnt)} -> {args.output} | {args.report}")


if __name__ == "__main__":
    main()
