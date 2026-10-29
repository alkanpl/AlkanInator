from __future__ import annotations

import csv
import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
WORK_DIR = ROOT / "outputs" / "zlaczki_seo_20261009"
OUTPUT = WORK_DIR / "BASELINKER_zlaczki_Simet_Morek_20261009.csv"
AUDIT_OUTPUT = WORK_DIR / "baselinker_import_audit.json"
DESCRIPTION_DIR = ROOT / "reports" / "descriptions" / "zlaczki_20261009"

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


def write_baselinker_csv(rows: list[dict[str, str]], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=BASELINKER_COLUMNS, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(rows)


SIMET_DATA = {
    "43814204": {
        "ean": "5907813210891",
        "image": "https://www.simet.com.pl/wp-content/uploads/43814204-ZSE1-1.0-6.0z-zlaczka-skretna-simet-1.jpg",
        "page": "https://www.simet.com.pl/produkty/zlaczka-skretna-uniwersalna-do-przewodow-1-6-mm%C2%B2-5/",
        "category": "Puszki i złączki > Złączki skręcane",
        "features": {
            "Rodzaj": "Złączka skrętna",
            "Model": "ZSE1-1.0/6.0z",
            "Kolor": "Żółty",
            "Liczba torów": "1",
            "Przekrój znamionowy [mm²]": "1-6",
            "Napięcie [V]": "400",
            "Prąd znamionowy [A]": "41",
            "Materiał": "Poliamid PA 66",
            "Temperatura pracy [°C]": "-25 do +125",
            "Wymiary [mm]": "19 x 27",
        },
    },
    "43814207": {
        "ean": "5907813213830",
        "image": "https://www.simet.com.pl/wp-content/uploads/43814207-ZSE1-1.0-6.0d-zlaczka-skretna-simet-1.jpg",
        "page": "https://www.simet.com.pl/produkty/zlaczka-skretna-uniwersalna-do-przewodow-1-6-mm%C2%B2-2/",
        "category": "Puszki i złączki > Złączki skręcane",
        "features": {
            "Rodzaj": "Złączka skrętna",
            "Model": "ZSE1-1.0/6.0d",
            "Kolor": "Czarny",
            "Liczba torów": "1",
            "Przekrój znamionowy [mm²]": "1-6",
            "Napięcie [V]": "400",
            "Prąd znamionowy [A]": "41",
            "Materiał": "Poliamid PA 66",
            "Temperatura pracy [°C]": "-25 do +125",
            "Wymiary [mm]": "19 x 27",
        },
    },
    "21412108": {
        "ean": "5907813207730",
        "image": "https://www.simet.com.pl/wp-content/uploads/6-t1168.webp",
        "page": "https://www.simet.com.pl/produkty/listwa-zaciskowa-gwintowa-termoplastyczna-6-mm%C2%B2-12-torowa-55c/",
        "category": "Puszki i złączki > Złączki skręcane",
        "features": {
            "Rodzaj": "Listwa zaciskowa",
            "Model": "LTF12-6.0",
            "Kolor": "Pomarańczowy",
            "Liczba torów": "12",
            "Przekrój znamionowy drut [mm²]": "2,5-6",
            "Przekrój znamionowy linka [mm²]": "1,5-4",
            "Napięcie [V]": "450",
            "Prąd znamionowy [A]": "41",
            "Materiał": "Polietylen + mosiądz",
            "Temperatura pracy [°C]": "-25 do +55",
            "Wymiary [mm]": "118 x 19 x 15",
        },
    },
    "21512108": {
        "ean": "5907813220012",
        "image": "https://www.simet.com.pl/wp-content/uploads/10-t1167.webp",
        "page": "https://www.simet.com.pl/produkty/listwa-zaciskowa-gwintowa-termoplastyczna-10-mm%C2%B2-12-torowa-55c/",
        "category": "Puszki i złączki > Złączki skręcane",
        "features": {
            "Rodzaj": "Listwa zaciskowa",
            "Model": "LTF12-10.0",
            "Kolor": "Pomarańczowy",
            "Liczba torów": "12",
            "Przekrój znamionowy drut [mm²]": "4-10",
            "Przekrój znamionowy linka [mm²]": "2,5-6",
            "Napięcie [V]": "750",
            "Prąd znamionowy [A]": "57",
            "Materiał": "Polietylen + mosiądz",
            "Temperatura pracy [°C]": "-25 do +55",
            "Wymiary [mm]": "141 x 25 x 20",
        },
    },
    "21612108": {
        "ean": "5907813207815",
        "image": "https://www.simet.com.pl/wp-content/uploads/Listwa-zaciskowa-gwintowa-termoplastyczna-LTF12-16.0.jpg",
        "page": "https://www.simet.com.pl/produkty/listwa-zaciskowa-gwintowa-termoplastyczna-16-mm%C2%B2-12-torowa-55c/",
        "category": "Puszki i złączki > Złączki skręcane",
        "features": {
            "Rodzaj": "Listwa zaciskowa",
            "Model": "LTF12-16.0",
            "Kolor": "Pomarańczowy",
            "Liczba torów": "12",
            "Przekrój znamionowy drut [mm²]": "6-16",
            "Przekrój znamionowy linka [mm²]": "4-10",
            "Napięcie [V]": "750",
            "Prąd znamionowy [A]": "76",
            "Materiał": "Polietylen + mosiądz",
            "Temperatura pracy [°C]": "-25 do +55",
            "Wymiary [mm]": "153 x 27 x 22",
        },
    },
    "21712008": {
        "ean": "5907813217272",
        "image": "https://www.simet.com.pl/wp-content/uploads/21712008-LTF12-25-listwa-12-torowa-simet-25mm2-1.jpg",
        "page": "https://www.simet.com.pl/produkty/listwa-zaciskowa-gwintowa-termoplastyczna-25-mm%C2%B2-12-torowa-85c/",
        "category": "Puszki i złączki > Złączki skręcane",
        "features": {
            "Rodzaj": "Listwa zaciskowa",
            "Model": "LTF12-25",
            "Kolor": "Pomarańczowy",
            "Liczba torów": "12",
            "Przekrój znamionowy drut [mm²]": "10-25",
            "Przekrój znamionowy linka [mm²]": "6-16",
            "Napięcie [V]": "750",
            "Prąd znamionowy [A]": "81",
            "Materiał": "Polietylen + mosiądz",
            "Temperatura pracy [°C]": "-40 do +85",
            "Wymiary [mm]": "212 x 33 x 25",
        },
    },
}

MANUFACTURER_DATA = {
    "Simet": "SIMET, Aleja Jana Pawła II 33 58-506 Jelenia Góra, sekretariat@simet.com.pl",
    "Morek": "MOREK, ul. Grzybowska 87 00-844 Warszawa, morekpl@morek.eu",
}


def gtin_valid(value: str) -> bool:
    if not value.isdigit() or len(value) not in {8, 12, 13, 14}:
        return False
    total = sum(int(digit) * (3 if index % 2 == 0 else 1) for index, digit in enumerate(value[-2::-1]))
    check = (10 - total % 10) % 10
    return check == int(value[-1])


def format_dimension(value: str) -> str:
    number = float(value) * 1000
    return f"{number:g}".replace(".", ",")


def morek_features(item: dict[str, str]) -> tuple[str, dict[str, str]]:
    name = item["name"]
    features: dict[str, str] = {
        "Producent": "Morek",
        "Dane producenta": MANUFACTURER_DATA["Morek"],
    }
    dimensions = [item.get("lengthM", ""), item.get("widthM", ""), item.get("heightM", "")]
    if all(dimensions):
        features["Wymiary [mm]"] = " x ".join(format_dimension(value) for value in dimensions)

    otl = re.fullmatch(
        r"Złączka\s+(OTL\S+)\s+(szary|niebieski|czerwony|żółto-zielony)\s+"
        r"(\d+)xAl/Cu\s+([\d,]+-[\d,]+)mm²\s+1000V AC/1500V DC\s+Zacisk uniwersalny",
        re.sub(r"\s+", " ", name).strip(),
        flags=re.IGNORECASE,
    )
    if otl:
        model, color, connection_groups, cross_section = otl.groups()
        holes_match = re.search(r"\((\d+)\s+otwor", item.get("group", ""), flags=re.IGNORECASE)
        features.update({
            "Rodzaj": "Bloki rozdzielcze",
            "Model": model,
            "Kolor": color.capitalize(),
            "Materiał przewodu": "Al/Cu",
            "Liczba grup przyłączeniowych": connection_groups,
            "Przekrój znamionowy [mm²]": cross_section,
            "Napięcie [V]": "1000 V AC / 1500 V DC",
            "Stopień ochrony [IP]": "IP20",
            "Sposób montażu": "Szyna DIN lub śruby",
        })
        if holes_match:
            features["Liczba otworów"] = holes_match.group(1)
        return "Puszki i złączki > Bloki rozdzielcze", features

    code = item["code"]
    list_data = {
        "MAD1007G15": {
            "Rodzaj": "Listwa zaciskowa uziemiająca",
            "Model": "PE7H",
            "Kolor": "Zielony",
            "Liczba torów": "7",
            "Przekrój znamionowy [mm²]": "1,5-16",
            "Napięcie [V]": "690",
            "Prąd znamionowy [A]": "63",
            "Stopień ochrony [IP]": "IP20",
            "Sposób montażu": "Szyna DIN",
        },
        "MAD1007N15": {
            "Rodzaj": "Listwa zaciskowa neutralna",
            "Model": "N7H",
            "Kolor": "Niebieski",
            "Liczba torów": "7",
            "Przekrój znamionowy [mm²]": "1,5-16",
            "Napięcie [V]": "690",
            "Prąd znamionowy [A]": "63",
            "Stopień ochrony [IP]": "IP20",
            "Sposób montażu": "Szyna DIN",
        },
        "MAD0012N15": {
            "Rodzaj": "Listwa zaciskowa neutralna",
            "Model": "N12",
            "Kolor": "Niebieski",
            "Liczba torów": "12",
            "Przekrój znamionowy [mm²]": "1,5-16",
            "Napięcie [V]": "690",
            "Sposób montażu": "Szyna TH35",
        },
        "MAD0052E21": {
            "Rodzaj": "Listwa zaciskowa N/PE",
            "Model": "BB52",
            "Liczba torów": "52",
            "Przekrój znamionowy [mm²]": "16",
            "Materiał": "Mosiądz",
        },
    }
    if code not in list_data:
        raise ValueError(f"Brak reguły atrybutów Morek dla {code}")
    features.update(list_data[code])
    return "Puszki i złączki > Listwy zaciskowe na szynę", features


def main() -> None:
    titles = json.loads((WORK_DIR / "seo_titles_for_import.json").read_text(encoding="utf-8"))
    sources = json.loads((WORK_DIR / "source_records.json").read_text(encoding="utf-8"))
    morek_rows = json.loads((WORK_DIR / "baselinker_source_data.json").read_text(encoding="utf-8"))

    title_by_sku = {str(row["SKU"]): str(row["Tytuł SEO"]) for row in titles}
    source_by_sku = {str(row["sku"]): row for row in sources}
    morek_by_code = {str(row["code"]): row for row in morek_rows}

    rows: list[dict[str, str]] = []
    audit: list[dict[str, object]] = []
    for sku, title in title_by_sku.items():
        source = source_by_sku[sku]
        code = re.sub(r"/(SIM|MOR)$", "", sku, flags=re.IGNORECASE)
        producer = str(source["supplier"])

        if producer == "Simet":
            data = SIMET_DATA[code]
            features = {
                "Producent": "Simet",
                "Dane producenta": MANUFACTURER_DATA["Simet"],
                **data["features"],
            }
            ean = data["ean"]
            category = data["category"]
            image = data["image"]
            attribute_source = data["page"]
            image_source = data["page"]
        else:
            item = morek_by_code[code]
            category, features = morek_features(item)
            ean = str(item["ean"])
            image = str(item["imageUrl"]).replace("http://", "https://", 1)
            attribute_source = f"MOREK - CENNIK.xlsx, CENNIK BAZOWY, wiersz {item['sourceRow']}"
            image_source = attribute_source

        if not gtin_valid(ean):
            raise ValueError(f"Niepoprawny GTIN: {sku} = {ean}")
        if not image.startswith("https://"):
            raise ValueError(f"Niepoprawny URL zdjęcia: {sku} = {image}")
        description_path = DESCRIPTION_DIR / f"{sku.replace('/', '_')}.html"
        description = description_path.read_text(encoding="utf-8").strip()

        rows.append({
            "product_id": "",
            "name": title,
            "sku": sku,
            "ean": ean,
            "manufacturer_name": producer,
            "category": category,
            "description": description,
            "features": json.dumps(features, ensure_ascii=False, separators=(",", ":")),
            "images_urls": image,
        })
        audit.append({
            "sku": sku,
            "status": "OK",
            "confidence": 1.0,
            "title_source": "ZLACZKI_tytuly_SEO.xlsx",
            "attribute_source": attribute_source,
            "image_source": image_source,
            "attributes": [
                {"name": name, "value": value, "source": attribute_source, "confidence": 1.0, "status": "OK"}
                for name, value in features.items()
            ],
        })

    if len(rows) != 37:
        raise ValueError(f"Oczekiwano 37 wierszy, otrzymano {len(rows)}")
    if len({row["sku"] for row in rows}) != 37 or len({row["ean"] for row in rows}) != 37:
        raise ValueError("SKU lub EAN nie są unikalne")
    if any(not row["images_urls"] for row in rows):
        raise ValueError("Brakuje URL zdjęcia")
    for row in rows:
        json.loads(row["features"])

    write_baselinker_csv(rows, OUTPUT)
    AUDIT_OUTPUT.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "output": str(OUTPUT),
        "audit": str(AUDIT_OUTPUT),
        "rows": len(rows),
        "simet": sum(row["manufacturer_name"] == "Simet" for row in rows),
        "morek": sum(row["manufacturer_name"] == "Morek" for row in rows),
        "images": sum(bool(row["images_urls"]) for row in rows),
        "valid_eans": sum(gtin_valid(row["ean"]) for row in rows),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
