"""Odswieza cala wiedze o sklepie z produkcji alkan.pl (decyzja uzytkownika 2026-11-18).

Kroki (kazdy przez istniejace narzedzia - API sklepu tylko przez ``woo_api_cli.py``):
1. ``woo_api_cli.py --profile prod knowledge`` -> ``dictionaries/woo_*_snapshot.tsv``
   (terminy atrybutow, kategorie, terminy sklepu, grupy wariantow) i
   ``cache/woo_api/knowledge.json``; kopia w ``cache/woo_api/prod/``.
2. Lista ``woo_attributes`` w ``dictionaries/kontakt_simon_shop_structure.yaml``
   = etykiety atrybutow globalnych z ``knowledge.json``.
3. Raport smieci i terminow bez produktow (``scripts/report_shop_junk_terms.py``).
4. Z pelnym katalogiem (``--pull`` albo ``--pull-file``):
   ``analyze_woocommerce_export.py`` -> ``dictionaries/woocommerce_catalog_knowledge.yaml``
   i ``learn_product_taxonomy.py`` -> ``dictionaries/learned_product_taxonomy.yaml``.

Pelny pull produkcji to ok. 25 zapytan po 500 produktow (tylko odczyt) - najlepiej
poza godzinami ruchu. Po odswiezeniu uruchom testy; zmiana wiedzy o katalogu
moze wymagac ``py tests/generate_snapshots.py`` (swiadoma zmiana slownikow).

Uzycie:
    py scripts/refresh_shop_knowledge.py --pull
    py scripts/refresh_shop_knowledge.py --pull-file cache/woo_api/prod/katalog_prod_2026-11-18.xlsx --skip-knowledge
    py scripts/refresh_shop_knowledge.py            # bez katalogu: zrzuty, atrybuty, raport smieci
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DICTIONARIES = ROOT / "dictionaries"
KNOWLEDGE_JSON = ROOT / "cache" / "woo_api" / "knowledge.json"
PROD_CACHE = ROOT / "cache" / "woo_api" / "prod"
SIMON_STRUCTURE = DICTIONARIES / "kontakt_simon_shop_structure.yaml"
SNAPSHOTS = (
    "woo_attribute_terms_snapshot.tsv",
    "woo_product_categories_snapshot.tsv",
    "woo_shop_terms_snapshot.tsv",
    "woo_variant_groups_snapshot.tsv",
)
ATTRIBUTES_BLOCK = re.compile(r"(?m)^woo_attributes:\r?\n(?:  - [^\r\n]*\r?\n)+")


def run(args: list[str]) -> None:
    print(">", " ".join(args), flush=True)
    subprocess.run([sys.executable, *args], cwd=ROOT, check=True)


def refresh_snapshots(profile: str) -> None:
    run(["src/woo_api_cli.py", "--profile", profile, "knowledge",
         "--dictionaries-dir", str(DICTIONARIES), "--cache", str(KNOWLEDGE_JSON)])
    PROD_CACHE.mkdir(parents=True, exist_ok=True)
    for name in SNAPSHOTS:
        source = DICTIONARIES / name
        if source.is_file():
            shutil.copy2(source, PROD_CACHE / name)
    shutil.copy2(KNOWLEDGE_JSON, PROD_CACHE / "knowledge.json")


def shop_attribute_labels(knowledge_path: Path) -> list[str]:
    knowledge = json.loads(knowledge_path.read_text(encoding="utf-8"))
    return sorted({attribute["label"] for attribute in knowledge.get("attributes") or [] if attribute.get("label")})


def update_simon_attribute_list(knowledge_path: Path, structure_path: Path = SIMON_STRUCTURE) -> tuple[list[str], list[str]]:
    """Podmienia liste ``woo_attributes`` (reszta pliku bez zmian). Zwraca (dodane, usuniete)."""
    labels = shop_attribute_labels(knowledge_path)
    text = structure_path.read_text(encoding="utf-8")
    match = ATTRIBUTES_BLOCK.search(text)
    if not match:
        raise SystemExit(f"Brak listy woo_attributes w {structure_path}")
    old = [line[len('  - '):].strip().strip('"') for line in match.group(0).splitlines()[1:]]
    block = "woo_attributes:\n" + "".join(f'  - "{label}"\n' for label in labels)
    text = text[:match.start()] + block + text[match.end():]
    text = re.sub(r'(?m)^updated_at: "[^"]*"', f'updated_at: "{date.today():%Y-%m-%d}"', text, count=1)
    structure_path.write_text(text, encoding="utf-8")
    return [label for label in labels if label not in old], [label for label in old if label not in labels]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--profile", default="prod", help="Profil sklepu z config.json (domyslnie prod).")
    parser.add_argument("--skip-knowledge", action="store_true", help="Bez knowledge - uzyj obecnych zrzutow i knowledge.json.")
    parser.add_argument("--pull", action="store_true", help="Pobierz pelny katalog (pull --all --per-page 500).")
    parser.add_argument("--pull-file", type=Path, help="Gotowy plik z pull --all (zamiast --pull).")
    args = parser.parse_args()

    if not args.skip_knowledge:
        refresh_snapshots(args.profile)
    added, removed = update_simon_attribute_list(KNOWLEDGE_JSON)
    print(f"woo_attributes: {len(shop_attribute_labels(KNOWLEDGE_JSON))} etykiet; dodane: {len(added)}; usuniete: {len(removed)}"
          " (szczegoly: git diff dictionaries/kontakt_simon_shop_structure.yaml)")
    run(["scripts/report_shop_junk_terms.py"])

    pull_file = args.pull_file
    if args.pull:
        pull_file = PROD_CACHE / f"katalog_{args.profile}_{date.today():%Y-%m-%d}.xlsx"
        run(["src/woo_api_cli.py", "--profile", args.profile, "pull", "--all", "--per-page", "500", "--output", str(pull_file)])
    if pull_file:
        run(["src/analyze_woocommerce_export.py", "--input", str(pull_file), "--woo-knowledge", str(KNOWLEDGE_JSON)])
        run(["src/learn_product_taxonomy.py", "--input", str(pull_file)])
    else:
        print("Bez --pull/--pull-file: woocommerce_catalog_knowledge.yaml i learned_product_taxonomy.yaml bez zmian.")
    print('Dalej: py -m unittest discover -s tests -p "test_*.py"')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
