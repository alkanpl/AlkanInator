"""Propozycja zmian regul Flexible Shipping (format alkan-shipping-methods-v1).

Czyta eksport metod wysylki, wprowadza zmiany dla paneli LED (kategoria 379)
i rur karbowanych (kategoria 4773) i zapisuje plik do importu na stanowisku
testowym + liste zmian. Eksport wejsciowy nie jest nadpisywany.

Zmiany (analiza 2026-10-05):
 1. DHL (34, 35): klasa Gabaryt (2126) z kategorii Panele LED (379) -> anuluj
    (karton 60x60 ma druga strone >60 cm, panel 120x30 >120 cm = NST w DHL).
 2. DHL gabaryt (45, 46): kategoria 379 -> anuluj.
 3. Gabaryt GLS/DPD/DHL (41-46): przedzial 60-90 kg wagi gabarytowej to
    3 paczki - dotad kosztowal tyle co 2. Nowa cena = cena 30-60 + doplata
    za kolejna paczke (roznica cen 1->2 paczki tej metody); przy zamowieniu
    od 499 zl cena jak dotad za 30-60 (pierwsza paczka gratis).
 4. Przesylka gabarytowa (24): regula "waga <= 28 kg -> anuluj" blokowala
    metode takze przy wadze gabarytowej >= 90 kg (towar lekki, ale duzy).
    Anulowanie tylko gdy waga <= 28 kg I waga gabarytowa < 90 kg.
 5. Paleta (29): rury karbowane (4773) z tagiem PopUpOsobisty (7583) albo
    Arot50+ (7624) -> 219 zl (dotad mialy wylacznie odbior osobisty).
 6. Klasa Gabaryt liczy paczki takze po WADZE (nie tylko po wadze gabarytowej):
    waga >= 30 kg -> zwykli kurierzy anulowani, gabaryt 30-60 / 60-90 jak dla
    wagi gabarytowej, >= 90 kg -> Przesylka gabarytowa 219 zl. Panele dostaja
    wage wysylkowa wedlug producenta (dictionaries/panel_shipping_packing.yaml,
    decyzja uzytkownika 2026-10-06: paczka = 2 kartony zbiorcze): Philips/
    Signify/PILA 8 szt. (3,6 kg; 9 szt. = 32,4 kg -> druga paczka), Ledvance
    12 szt. (2,4 kg), pozostali 10 szt. (2,88 kg).
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import json
from pathlib import Path

CAT_PANELS = "379"
CAT_PIPES = "4773"
CLS_GABARYT = "2126"
TAG_POPUP_OSOBISTY = "7583"
TAG_AROT50 = "7624"

# (metoda: (cena 30-60 <499, cena 30-60 >=499, doplata za kolejna paczke))
BANDS = {
    41: ("32.80", "19.90", 12.90),   # GLS gabaryt: 19.90 -> 32.80
    42: ("40.99", "25.90", 15.09),   # GLS pobranie gabaryt: 25.90 -> 40.99
    43: ("36", "19.99", 16.01),      # DPD gabaryt: 19.99 -> 36
    44: ("46.80", "25.99", 20.81),   # DPD pobranie gabaryt: 25.99 -> 46.80
    45: ("44.80", "24.90", 19.90),   # DHL gabaryt: 24.90 -> 44.80
    46: ("55.60", "30.90", 24.70),   # DHL pobranie gabaryt: 30.90 -> 55.60
}


def cond_range(cid, lo="", hi=""):
    return {"condition_id": cid, "operator": "is", "min": str(lo), "max": str(hi)}


def cond_any(cid, ids, op="any"):
    return {"condition_id": cid, "operator": op, cid: list(ids)}


def rule(conds, cost="", action="none"):
    return {"conditions": conds, "cost_per_order": cost, "additional_costs": [], "special_action": action}


def is_band(r, lo, hi):
    return any(c["condition_id"] == "dimensional_weight" and c.get("min") == lo and c.get("max") == hi
               for c in r["conditions"])


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", default="input/alkan-shipping-methods-2026-10-05-065701.json")
    today = dt.date.today().isoformat()
    ap.add_argument("--output", default=f"output/alkan-shipping-methods-PROPOZYCJA-{today}.json")
    ap.add_argument("--changes", default=f"reports/alkan-shipping-methods-ZMIANY-{today}.md")
    args = ap.parse_args()

    raw = Path(args.input).read_text(encoding="utf-8-sig")
    data = json.loads(raw)
    orig = copy.deepcopy(data)
    zone = next(z for z in data["zones"] if z["zone_name"] == "Polska")
    methods = {m["instance_id"]: m for m in zone["methods"]}
    log = []

    def rules_of(iid):
        return json.loads(methods[iid]["settings"]["method_rules"])

    def save(iid, rules):
        methods[iid]["settings"]["method_rules"] = json.dumps(rules, indent=4, ensure_ascii=False)

    # 1. DHL zwykly i pobranie: panele w klasie Gabaryt -> anuluj
    for iid in (34, 35):
        rules = rules_of(iid)
        rules.insert(0, rule([cond_any("shipping_class", [CLS_GABARYT]),
                              cond_any("product_category", [CAT_PANELS])], action="cancel"))
        save(iid, rules)
        log.append(f"#{iid} {methods[iid]['title']}: + anuluj gdy klasa Gabaryt (2126) i kategoria Panele LED (379)")

    # 2. DHL gabaryt: panele -> anuluj
    for iid in (45, 46):
        rules = rules_of(iid)
        rules.insert(0, rule([cond_any("product_category", [CAT_PANELS])], action="cancel"))
        save(iid, rules)
        log.append(f"#{iid} {methods[iid]['title']}: + anuluj dla kategorii Panele LED (379)")

    # 3. przedzial 60-90 = 3 paczki
    for iid, (p2, p2_free, step) in BANDS.items():
        rules = rules_of(iid)
        old = [r for r in rules if is_band(r, "60", "90")]
        assert len(old) == 1, (iid, old)
        idx = rules.index(old[0])
        p3 = f"{float(p2) + step:.2f}"
        new = [
            rule([cond_range("dimensional_weight", 60, 90), cond_any("shipping_class", [CLS_GABARYT]),
                  cond_range("value", 0, "498.99")], p3),
            rule([cond_range("dimensional_weight", 60, 90), cond_any("shipping_class", [CLS_GABARYT]),
                  cond_range("value", 499, "")], p2),
        ]
        rules[idx:idx + 1] = new
        save(iid, rules)
        log.append(f"#{iid} {methods[iid]['title']}: waga gabarytowa 60-90 kg: {old[0]['cost_per_order']} zl -> "
                   f"{p3} zl (<499 zl) / {p2} zl (>=499 zl)")

    # 4. Przesylka gabarytowa: nie anuluj lekkiego, ale duzego towaru
    rules = rules_of(24)
    hit = [r for r in rules if r["special_action"] == "cancel" and len(r["conditions"]) == 1
           and r["conditions"][0]["condition_id"] == "weight" and r["conditions"][0].get("max") == "28"]
    assert len(hit) == 1
    hit[0]["conditions"].append(cond_range("dimensional_weight", "", "89.99"))
    save(24, rules)
    log.append("#24 Przesyłka gabarytowa: anulowanie 'waga <= 28 kg' tylko gdy także waga gabarytowa < 90 kg")

    # 5. Paleta dla duzych krazkow rur
    rules = rules_of(29)
    for tag in (TAG_POPUP_OSOBISTY, TAG_AROT50):
        rules.append(rule([cond_any("product_category", [CAT_PIPES]), cond_any("product_tag", [tag])], "219"))
    save(29, rules)
    log.append("#29 Paleta: + 219 zł dla rur karbowanych (4773) z tagiem PopUpOsobisty (7583) lub Arot50+ (7624)")

    # 6. Gabaryt liczony takze po wadze
    for iid in (30, 31, 32, 33, 34, 35):
        rules = rules_of(iid)
        rules.insert(0, rule([cond_range("weight", 30, ""), cond_any("shipping_class", [CLS_GABARYT])], action="cancel"))
        save(iid, rules)
        log.append(f"#{iid} {methods[iid]['title']}: + anuluj gdy klasa Gabaryt i waga >= 30 kg")
    for iid, (p2, p2_free, step) in BANDS.items():
        rules = rules_of(iid)
        low = [r for r in rules if r["special_action"] == "cancel" and is_band(r, "0", "30")]
        assert len(low) == 1, (iid, low)
        low[0]["conditions"].append(cond_range("weight", "", "29.99"))
        rules.append(rule([cond_range("weight", 90, ""), cond_any("shipping_class", [CLS_GABARYT])], action="cancel"))
        p3 = f"{float(p2) + step:.2f}"
        for lo, hi, small, big in ((30, 60, p2, p2_free), (60, 90, p3, p2)):
            rules.append(rule([cond_range("weight", lo, hi), cond_any("shipping_class", [CLS_GABARYT]),
                               cond_range("value", 0, "498.99")], small))
            rules.append(rule([cond_range("weight", lo, hi), cond_any("shipping_class", [CLS_GABARYT]),
                               cond_range("value", 499, "")], big))
        save(iid, rules)
        log.append(f"#{iid} {methods[iid]['title']}: progi 30-60 / 60-90 kg liczone też po wadze (klasa Gabaryt); "
                   f"anulowanie 'poniżej 30 kg' tylko gdy obie wagi < 30; waga >= 90 kg -> anuluj")
    rules = rules_of(24)
    rules.append(rule([cond_range("weight", 90, 500), cond_any("shipping_class", [CLS_GABARYT])], "219"))
    save(24, rules)
    log.append("#24 Przesyłka gabarytowa: + 219 zł gdy klasa Gabaryt i waga >= 90 kg")

    data["created_at_gmt"] = dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    out = json.dumps(data, indent=4, ensure_ascii=False)
    Path(args.output).write_text(out, encoding="utf-8-sig")

    # kontrola: zmienione tylko zaplanowane metody
    changed = [m["instance_id"] for m, o in zip(zone["methods"], next(z for z in orig["zones"] if z["zone_name"] == "Polska")["methods"]) if m != o]
    md = [f"# Propozycja reguł wysyłki – {today}", "", f"Źródło: `{args.input}`", f"Wynik: `{args.output}`",
          f"Zmienione metody (instance_id): {changed}", "", "## Zmiany", ""] + [f"- {x}" for x in log]
    Path(args.changes).write_text("\n".join(md) + "\n", encoding="utf-8")
    print("\n".join(log)); print("changed:", changed)


if __name__ == "__main__":
    main()
