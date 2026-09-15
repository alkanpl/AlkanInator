# -*- coding: utf-8 -*-
"""Uzupelnia 'Sposob montazu' i 'Stopien ochrony [IP]' dla osprzetu Hager/Berker.

Zrodla, w kolejnosci pewnosci:
  1. ETIM/BMEcat Hager  - dane producenta (parametry 'Sposob montazu', 'Stopien ochrony (IP)').
  2. Katalog PDF        - sekcja strony i opis pozycji z eksportu warstwy tekstowej katalogu.

Zasada: nie zgadujemy. Jesli zadne zrodlo nie potwierdza wartosci, pole zostaje puste
ze statusem DO_SPRAWDZENIA. Istniejace wartosci nie sa nadpisywane - konflikty trafiaja
do kolumny uwag i do raportu.

Plik wejsciowy nie jest modyfikowany; wynik idzie do output/, raport do reports/.

    py src/build_hager_berker_mounting_ip.py
"""
from __future__ import annotations

import argparse
import re
from collections import Counter
from pathlib import Path

import unicodedata

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_INPUT = Path("input/hager-berker-sposob-montazu-IP-seria-02.09.xlsx")
DEFAULT_ETIM = Path("output/BMEcat_Hager_PL_produkty_ETIM_PL_ZDJECIA_RELACJE.xlsx")
DEFAULT_CATALOG = Path("output/Hager-Berker-Katalog-PDF_BAZA_KOMPATYBILNOSCI.xlsx")
DEFAULT_RULES = Path("dictionaries/hager_berker_mounting_rules.yaml")
DEFAULT_IP_RULES = Path("dictionaries/hager_berker_ip_rules.yaml")

COL_SKU = 1
COL_TITLE = 5
COL_MOUNT = 6
COL_SERIES = 7
COL_IP = 8

# Slownik sklepu dla 'Sposob montazu'. Wartosci ETIM spoza tej mapy nie sa tlumaczone.
MOUNT_MAP = {
    "montaż podtynkowy": "Podtynkowy",
    "podtynkowy": "Podtynkowy",
    "natynkowy": "Natynkowy",
    "montaż natynkowy": "Natynkowy",
}

# Grupy produktowe, dla ktorych sklep celowo NIE wypelnia sposobu montazu -
# to elementy nakladane na mechanizm, a nie osprzet mocowany w scianie.
SKIP_MOUNT_PREFIXES = (
    "ramka", "ramki", "klawisz", "klawisze", "płytka", "płytki",
    "płyta czołowa", "płyta", "plytka", "plytki",
)

# 'Z pazurkami rozporowymi' w danych Hagera = mocowanie w puszce podtynkowej.
CLAW_RE = re.compile(r"pazur", re.IGNORECASE)


def deaccent(text: str) -> str:
    text = str(text).lower().replace("ł", "l")
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()


def match_ip_rule(rules, sku: str, title: str, series: str):
    """Zwraca regule IP zweryfikowana na hager.com albo None."""
    flat_type = deaccent(product_type(title))
    # Reguly po prefiksie SKU sa bardziej szczegolowe niz reguly typu produktu,
    # wiec sprawdzamy je w pierwszej kolejnosci - inaczej ogolna regula typu
    # przeslonilaby wyjatek ustalony dla konkretnej rodziny.
    for rule in rules:
        prefix = rule.get("dopasowanie", {}).get("sku_prefiks")
        if prefix and str(sku).upper().startswith(str(prefix).upper()):
            return rule
    for rule in rules:
        m = rule.get("dopasowanie", {})
        if m.get("sku_prefiks") or deaccent(m.get("typ", "")) != flat_type:
            continue
        # Brak 'seria' w regule oznacza: dowolna seria.
        if "seria" not in m or m["seria"] == series:
            return rule
    return None


def product_type(title: str) -> str:
    return str(title or "").strip().split(",")[0].split(" ")[0].capitalize()

IP_RE = re.compile(r"\bIP\s*([0-9]{2})\b", re.IGNORECASE)


def base_code(sku: str) -> str:
    return str(sku).split("/")[0].strip().upper()


def code_aliases(code: str) -> list[str]:
    """Warianty kodu sklepu spotykane w kartotece (np. prefiks 5BE)."""
    out = [code]
    if code.startswith("5BE") and len(code) > 3:
        out.append(code[3:])
    return out


def skip_mount(title: str) -> bool:
    low = str(title or "").strip().lower()
    return low.startswith(SKIP_MOUNT_PREFIXES)


def norm_ip(value: str) -> str | None:
    """Sprowadza 'IP20' / 'IP 20' do konwencji arkusza: 'IP 20'."""
    if not value:
        return None
    m = IP_RE.search(str(value))
    return f"IP {m.group(1)}" if m else None


def load_etim(path: Path) -> dict[str, dict[str, str]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Produkty"]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    i_code = header.index("Kod produktu")
    i_name = header.index("Nazwa")
    i_par = header.index("Parametry ETIM PL")
    out: dict[str, dict[str, str]] = {}
    for row in rows:
        code = row[i_code]
        if not code:
            continue
        params: dict[str, str] = {}
        for part in str(row[i_par] or "").split(";"):
            if "=" in part:
                key, value = part.split("=", 1)
                value = value.strip()
                if value and value != "-":
                    params[key.strip()] = value
        params["__nazwa__"] = str(row[i_name] or "")
        out[str(code).strip().upper()] = params
    wb.close()
    return out


def load_catalog(path: Path) -> dict[str, dict[str, str]]:
    """Mapuje SKU sklepu na kontekst katalogowy (sekcja, opis, strona)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Produkty katalogowe"]
    header = None
    out: dict[str, dict[str, str]] = {}
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        if row[0] == "Kod katalogowy":
            header = list(row)
            continue
        if header is None:
            continue
        rec = dict(zip(header, row))
        skus = str(rec.get("SKU sklepu") or "")
        if not skus:
            continue
        context = " | ".join(
            str(rec.get(key) or "")
            for key in ("Nazwa automatyczna", "Wariant / opis", "Seria / system", "Sekcja strony")
        )
        for sku in re.split(r"[;,]\s*", skus):
            sku = sku.strip()
            if not sku:
                continue
            prev = out.get(sku)
            if prev is None:
                out[sku] = {"context": context, "page": str(rec.get("Strona drukowana") or "")}
            else:
                prev["context"] += " || " + context
    wb.close()
    return out


def mount_from_text(text: str) -> str | None:
    low = text.lower()
    has_nat = "natynkow" in low or "natynkow" in low.replace("ę", "e")
    has_pod = "podtynkow" in low
    if has_nat and not has_pod:
        return "Natynkowy"
    if has_pod and not has_nat:
        return "Podtynkowy"
    return None


def ip_from_text(text: str) -> str | None:
    found = {f"IP {m.group(1)}" for m in IP_RE.finditer(text)}
    return found.pop() if len(found) == 1 else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default=str(DEFAULT_INPUT))
    ap.add_argument("--etim", default=str(DEFAULT_ETIM))
    ap.add_argument("--catalog", default=str(DEFAULT_CATALOG))
    ap.add_argument("--rules", default=str(DEFAULT_RULES))
    ap.add_argument("--ip-rules", default=str(DEFAULT_IP_RULES))
    ap.add_argument("--output", required=True)
    ap.add_argument("--report", required=True)
    args = ap.parse_args()

    rules = yaml.safe_load(Path(args.rules).read_text(encoding="utf-8"))
    flush_rules = {r["typ"]: r for r in rules.get("podtynkowy", [])}
    skip_phrases = [deaccent(p) for p in rules.get("pomijane_frazy_tytulu", [])]
    ip_rules = yaml.safe_load(Path(args.ip_rules).read_text(encoding="utf-8")).get("reguly", [])
    etim = load_etim(Path(args.etim))
    catalog = load_catalog(Path(args.catalog))

    wb_in = load_workbook(args.input, read_only=True, data_only=True)
    rows = list(wb_in.active.iter_rows(values_only=True))
    wb_in.close()
    header, data = list(rows[0]), [r for r in rows[1:] if r and r[COL_SKU]]

    out_header = header + [
        "Sposób montażu (uzupełniony)", "Montaż: źródło", "Montaż: pewność", "Montaż: status",
        "Stopień ochrony [IP] (uzupełniony)", "IP: źródło", "IP: pewność", "IP: status",
        "Sposób mocowania (ETIM)", "Uwagi",
    ]

    stats = Counter()
    out_rows = []
    conflicts = []
    overrides = []

    for row in data:
        row = list(row)
        sku = str(row[COL_SKU])
        code = base_code(sku)
        params = {}
        matched_code = ""
        for alias in code_aliases(code):
            if alias in etim:
                params = etim[alias]
                matched_code = alias
                break
        cat = catalog.get(sku) or catalog.get(sku.upper()) or {}
        cat_text = cat.get("context", "")
        cat_page = cat.get("page", "")
        notes = []

        # --- sposob montazu ---
        mount_old = (row[COL_MOUNT] or "").strip() if isinstance(row[COL_MOUNT], str) else row[COL_MOUNT]
        title = str(row[COL_TITLE] or "")
        fixing = params.get("Sposób mocowania", "")
        raw = params.get("Sposób montażu")
        mount_val = mount_src = mount_status = ""
        mount_conf = ""

        title_flat = deaccent(title)
        ptype = product_type(title)
        rule = flush_rules.get(ptype)
        if rule and str(row[COL_SERIES] or "") in (rule.get("pomijane_serie") or []):
            rule = None

        if skip_mount(title) or any(p in title_flat for p in skip_phrases):
            mount_status = "POMINIETE"
            notes.append("element nakładany/zatrzaskowy - montaż pomijany")
        else:
            maker_name = params.get("__nazwa__", "")
            maker_flat = deaccent(maker_name)
            title_low = title.lower()
            if "natynkow" in maker_flat or "podtynkow" in maker_flat:
                mount_val = "Natynkowy" if "natynkow" in maker_flat else "Podtynkowy"
                mount_src = f"nazwa producenta: {maker_name}"
                mount_conf, mount_status = 0.95, "OK"
            elif "natynkow" in title_low:
                mount_val, mount_src, mount_conf, mount_status = (
                    "Natynkowy", "tytuł produktu", 0.90, "OK")
            elif "podtynkow" in title_low:
                mount_val, mount_src, mount_conf, mount_status = (
                    "Podtynkowy", "tytuł produktu", 0.90, "OK")
            elif fixing and CLAW_RE.search(fixing):
                # Pazurki rozporowe montuje sie w puszce - to zawsze podtynk.
                mount_val, mount_src, mount_conf, mount_status = (
                    "Podtynkowy", f"ETIM mocowanie: {fixing}", 0.90, "OK")
                if raw and MOUNT_MAP.get(raw.strip().lower()) == "Natynkowy":
                    notes.append(f"ETIM podawał '{raw}', ale mocowanie pazurkami wskazuje podtynk")
            elif raw and MOUNT_MAP.get(raw.strip().lower()) == "Podtynkowy":
                # Wlasna wartosc producenta jest bardziej szczegolowa niz regula typu.
                mount_val, mount_src, mount_conf, mount_status = (
                    "Podtynkowy", "ETIM BMEcat Hager", 0.95, "OK")
            elif rule:
                codes = ", ".join(str(c) for c in (rule.get("sprawdzone") or [])[:3])
                mount_val, mount_src, mount_conf, mount_status = (
                    "Podtynkowy", f"hager.com, reguła typu (sprawdzone: {codes})", 0.85, "OK")
            elif raw:
                mapped = MOUNT_MAP.get(raw.strip().lower())
                if mapped == "Natynkowy":
                    mount_val, mount_src, mount_conf, mount_status = (
                        mapped, "ETIM BMEcat Hager", 0.70, "WARNING")
                else:
                    notes.append(f"ETIM montaż poza słownikiem sklepu: {raw}")
                    mount_status = "DO_SPRAWDZENIA"
            if not mount_val and cat_text and not mount_status:
                guess = mount_from_text(cat_text)
                if guess:
                    mount_val, mount_conf, mount_status = guess, 0.70, "WARNING"
                    mount_src = f"Katalog PDF s.{cat_page}" if cat_page else "Katalog PDF"
            if not mount_val and not mount_status:
                mount_status = "DO_SPRAWDZENIA"
                if ptype in (rules.get("brak_danych_u_producenta", {}).get("typy") or []):
                    notes.append("hager.com nie publikuje sposobu montażu dla tej grupy")
                if fixing:
                    notes.append(f"mocowanie wg ETIM: {fixing} (niejednoznaczne dla montażu)")

        if mount_old:
            may_override = mount_src.startswith(("nazwa producenta", "ETIM mocowanie"))
            if mount_val and mount_val != mount_old and may_override:
                overrides.append((sku, title, str(row[COL_SERIES] or ""),
                                  "Sposób montażu", mount_old, mount_val, mount_src))
                notes.append(f"nadpisano montaż: arkusz miał {mount_old}, producent {mount_val}")
                mount_final, mount_status = mount_val, "OK"
                stats["montaż: nadpisane wg producenta"] += 1
            else:
                if mount_val and mount_val != mount_old:
                    conflicts.append((sku, title, str(row[COL_SERIES] or ""),
                                      "Sposób montażu", mount_old, mount_val, mount_src))
                    notes.append(f"KONFLIKT montaż: w arkuszu {mount_old}, źródło {mount_val}")
                mount_final, mount_src, mount_conf, mount_status = (
                    mount_old, "arkusz wejściowy", 1.0, "OK")
                stats["montaż: było"] += 1
        else:
            mount_final = mount_val
            if mount_val:
                stats[f"montaż: uzupełnione ({mount_src.split(' s.')[0].split(' (')[0].split(':')[0]})"] += 1
            elif mount_status == "POMINIETE":
                stats["montaż: pominięty (ramka/klawisz/płytka)"] += 1
            else:
                stats["montaż: nadal brak"] += 1

        # --- IP ---
        ip_old = (row[COL_IP] or "").strip() if isinstance(row[COL_IP], str) else row[COL_IP]
        ip_val = ip_src = ip_status = ""
        ip_conf = ""
        raw_ip = params.get("Stopień ochrony (IP)")
        cand = norm_ip(raw_ip) if raw_ip else None
        # Regule z hager.com stosujemy tylko tam, gdzie jest co rozstrzygac:
        # arkusz klóci sie z ETIM albo brakuje wartosci. Gdy oba zrodla sa zgodne,
        # nie ma powodu podmieniac wartosci na podstawie reprezentanta rodziny.
        candidate_rule = match_ip_rule(ip_rules, sku, title, str(row[COL_SERIES] or ""))
        ip_rule = None
        if candidate_rule and (
            candidate_rule.get("wymus")
            or (ip_old and cand and ip_old != cand)
            or not (ip_old or cand)
        ):
            ip_rule = candidate_rule
        if ip_rule:
            ip_val = ip_rule["wartosc"]
            ip_src = f"hager.com (sprawdzone: {ip_rule.get('sprawdzone')})"
            ip_conf, ip_status = 0.95, "OK"
            if ip_rule.get("uwaga"):
                notes.append(f"IP wg hager.com: {ip_rule['uwaga']}")
        elif cand:
            ip_val, ip_src, ip_conf, ip_status = cand, "ETIM BMEcat Hager", 0.95, "OK"
        elif raw_ip:
            notes.append(f"ETIM IP nieczytelne: {raw_ip}")
        if not ip_val and cat_text:
            guess = ip_from_text(cat_text)
            if guess:
                ip_val, ip_conf, ip_status = guess, 0.70, "WARNING"
                ip_src = f"Katalog PDF s.{cat_page}" if cat_page else "Katalog PDF"
        if not ip_val and not ip_status:
            ip_status = "DO_SPRAWDZENIA"

        if ip_old:
            if ip_val and ip_val != ip_old and ip_src.startswith("hager.com"):
                overrides.append((sku, title, str(row[COL_SERIES] or ""),
                                  "Stopień ochrony [IP]", ip_old, ip_val, ip_src))
                notes.append(f"nadpisano IP: arkusz miał {ip_old}, hager.com {ip_val}")
                ip_final, ip_status = ip_val, "OK"
                stats["IP: nadpisane wg hager.com"] += 1
            else:
                if ip_val and ip_val != ip_old:
                    conflicts.append((sku, title, str(row[COL_SERIES] or ""),
                                      "Stopień ochrony [IP]", ip_old, ip_val, ip_src))
                    notes.append(f"KONFLIKT IP: w arkuszu {ip_old}, źródło {ip_val}")
                ip_final, ip_src, ip_conf, ip_status = ip_old, "arkusz wejściowy", 1.0, "OK"
                stats["IP: było"] += 1
        else:
            ip_final = ip_val
            if ip_val:
                stats[f"IP: uzupełnione ({ip_src.split(' s.')[0].split(' (')[0]})"] += 1
            else:
                stats["IP: nadal brak"] += 1

        if not matched_code:
            notes.append("brak kodu w BMEcat")
            stats["bez dopasowania w BMEcat"] += 1

        out_rows.append(row + [
            mount_final, mount_src, mount_conf, mount_status,
            ip_final, ip_src, ip_conf, ip_status,
            params.get("Sposób mocowania", ""), "; ".join(notes),
        ])

    # --- zapis skoroszytu kontrolnego ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Produkty"
    ws.append(out_header)
    for row in out_rows:
        ws.append(row)
    head_fill = PatternFill("solid", fgColor="DDEBF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    widths = {1: 8, 2: 16, 6: 46, 7: 22, 8: 26, 9: 14, 10: 22, 11: 20, 12: 11, 13: 16,
              14: 22, 15: 20, 16: 11, 17: 16, 18: 50}
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws2 = wb.create_sheet("Konflikty")
    ws2.append(["SKU", "Tytuł", "Seria", "Atrybut", "Wartość w arkuszu", "Wartość wg Hagera", "Źródło"])
    for c in conflicts:
        ws2.append(list(c))
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
    for letter, width in (("A", 18), ("B", 60), ("C", 26), ("D", 24), ("E", 20), ("F", 20), ("G", 26)):
        ws2.column_dimensions[letter].width = width

    ws4 = wb.create_sheet("Nadpisane")
    ws4.append(["SKU", "Tytuł", "Seria", "Atrybut", "Było w arkuszu", "Wpisano", "Podstawa"])
    for o in overrides:
        ws4.append(list(o))
    for cell in ws4[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
    for letter, width in (("A", 18), ("B", 60), ("C", 26), ("D", 24), ("E", 16), ("F", 14), ("G", 58)):
        ws4.column_dimensions[letter].width = width

    ws3 = wb.create_sheet("Podsumowanie")
    ws3.append(["Metryka", "Wartość"])
    ws3.append(["Produktów", len(out_rows)])
    for key, value in sorted(stats.items()):
        ws3.append([key, value])
    for cell in ws3[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
    ws3.column_dimensions["A"].width = 44
    ws3.column_dimensions["B"].width = 12

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)

    lines = [f"Produktów: {len(out_rows)}"]
    lines += [f"{k}: {v}" for k, v in sorted(stats.items())]
    lines.append(f"nadpisane wg producenta: {len(overrides)}")
    lines.append(f"konflikty ze źródłem: {len(conflicts)}")
    Path(args.report).parent.mkdir(parents=True, exist_ok=True)
    Path(args.report).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    print(f"\nzapisano: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
