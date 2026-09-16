"""Raportuje, jak taksonomia rodzin klasyfikuje produkty z pliku wejsciowego.

Sluzy do pilnowania zasady: kazdy nowy typ produktu dostaje wlasna rodzine.
Produkty w rodzinie `nieznana` to sygnal, ze brakuje klasy w
dictionaries/product_family_taxonomy.yaml.

Konsola Windows uzywa cp1250, dlatego raport idzie do pliku w utf-8.
"""
from __future__ import annotations

import argparse
import collections
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "src"))

from generate_product_descriptions import (  # noqa: E402
    classify_family,
    collect_attributes,
    is_light_product,
    load_family_taxonomy,
    product_title,
    subject_for,
)
from utils import read_products  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True)
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--report", required=True)
    parser.add_argument("--title-column", default="")
    args = parser.parse_args()

    df = read_products(args.input, sheet_name=args.sheet)
    taxonomy = load_family_taxonomy()
    unknown_name = taxonomy.get("defaults", {}).get("unknown_family", "nieznana")

    counts: collections.Counter[str] = collections.Counter()
    samples: dict[str, list[str]] = collections.defaultdict(list)
    for _, row in df.iterrows():
        title = product_title(row, args.title_column)
        family = classify_family(title, collect_attributes(row))
        counts[family] += 1
        if len(samples[family]) < 5:
            samples[family].append(title[:80])

    total = sum(counts.values())
    unknown = counts.get(unknown_name, 0)
    lines = [
        f"PLIK: {args.input}",
        f"PRODUKTOW: {total}",
        f"RODZIN UZYTYCH: {len(counts)}",
        f"NIEROZPOZNANE ({unknown_name}): {unknown} ({100 * unknown / max(total, 1):.1f}%)",
        "",
        "ROZKLAD:",
    ]
    for family, count in counts.most_common():
        flag = "  <-- BRAK KLASY, DODAJ DO TAKSONOMII" if family == unknown_name else ""
        lines.append(f"  {count:5}  {family:28} swiatlo={str(is_light_product(family)):5} "
                     f'podmiot="{subject_for(family)}"{flag}')
        for sample in samples[family]:
            lines.append(f"            {sample}")

    pathlib.Path(args.report).write_text("\n".join(lines), encoding="utf-8")
    print(f"OK: raport -> {args.report}; nierozpoznanych {unknown}/{total}")
    if unknown:
        print("UWAGA: czesc produktow nie ma wlasnej klasy - zglos to uzytkownikowi.")


if __name__ == "__main__":
    main()
