"""Regeneruje golden-file snapshoty dla testow charakteryzujacych.

Snapshoty utrwalaja AKTUALNE zachowanie ekstrakcji atrybutow i generowania
tytulow na realnych danych z pliku master Kanlux. Testy w
``tests/test_extract_attributes.py`` i ``tests/test_optimize_titles.py``
porownuja kod z tymi plikami, wiec kazda niezamierzona zmiana wyniku jest
od razu widoczna.

Uruchamiaj TYLKO wtedy, gdy swiadomie akceptujesz nowe wyniki jako poprawne:

    py tests/generate_snapshots.py

Wymaga obecnosci pliku wejsciowego
``archived_input_files/input_2026-08-05/Alkan_Kanlux_pelne_rodziny.xlsx``.
Same testy go nie wymagaja - czytaja gotowe fixtures z ``tests/fixtures/``.
"""

from __future__ import annotations

import io
import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from analyze_products import detect_product_columns  # noqa: E402
from catalog_knowledge import (  # noqa: E402
    enrich_config_with_catalog_knowledge,
    load_catalog_knowledge,
)
from extract_attributes import (  # noqa: E402
    extract_attributes_from_text,
    extract_attributes_for_dataframe,
)
from optimize_titles import optimize_titles_for_dataframe  # noqa: E402
from title_anatomy import load_title_anatomy_config  # noqa: E402
from utils import load_yaml  # noqa: E402
from validate_output import classify_product_roles  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
INPUT = (
    ROOT
    / "archived_input_files"
    / "input_2026-08-05"
    / "Alkan_Kanlux_pelne_rodziny.xlsx"
)
SHEET = "7. Wszystkie SKU (master)"
SAMPLE_SIZE = 40


def diverse_titles(df: pd.DataFrame, columns: list[str]) -> list[str]:
    seen: set[str] = set()
    titles: list[str] = []
    for column in columns:
        if column not in df.columns:
            continue
        for value in df[column].tolist():
            key = str(value).strip()
            if key and key not in seen:
                seen.add(key)
                titles.append(key)
    step = max(1, len(titles) // SAMPLE_SIZE)
    return titles[::step][:SAMPLE_SIZE]


def write_json(path: Path, payload: object) -> None:
    with io.open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def generate_extraction_snapshot(df: pd.DataFrame) -> int:
    titles = diverse_titles(df, ["Nazwa B2C / SEO", "Nazwa Kanlux"])
    cases = [
        {"title": title, "attributes": extract_attributes_from_text(title)["attributes"]}
        for title in titles
    ]
    write_json(FIXTURES / "extraction_snapshot.json", cases)
    return len(cases)


def build_pipeline_titles(source_df: pd.DataFrame) -> pd.DataFrame:
    """Odtwarza rdzen pipeline'u tytulow (extract -> role -> optimize).

    Uzywane zarowno przez generator snapshotu, jak i przez test, dzieki czemu
    test nie potrzebuje pliku XLSX - wystarczaja surowe wiersze z fixture.
    """
    config = enrich_config_with_catalog_knowledge(
        load_yaml("configs/categories/kanlux-oswietlenie.yaml"),
        load_catalog_knowledge(),
    )
    anatomy = load_title_anatomy_config()
    columns = detect_product_columns(source_df)
    title_column = columns["title"]

    # Te same zasady budowy tytulu roboczego co add_working_title_column w
    # run_pipeline.py: fallback na kolejne kolumny, gdy glowny tytul jest pusty.
    working = source_df.copy()
    working_title = working[title_column].astype(str)
    fallback_columns = [c for c in config.get("title_fallback_columns", []) if c in working.columns]
    for column in fallback_columns:
        working_title = working_title.where(
            working_title.str.strip().ne(""), working[column].astype(str)
        )
    working["__working_title"] = working_title
    extracted = extract_attributes_for_dataframe(
        working, "__working_title", config, columns.get("producer")
    )
    classified = classify_product_roles(extracted, config)
    return optimize_titles_for_dataframe(classified, "__working_title", config, anatomy)


def generate_title_snapshot(df: pd.DataFrame) -> int:
    source_columns = list(df.columns)
    sample = df.iloc[:: max(1, len(df) // SAMPLE_SIZE)].head(SAMPLE_SIZE).copy()
    optimized = build_pipeline_titles(sample[source_columns])

    cases = []
    for (_, source_row), (_, result_row) in zip(
        sample[source_columns].iterrows(), optimized.iterrows()
    ):
        cases.append(
            {
                "source": {key: str(value) for key, value in source_row.items()},
                "expected": {
                    "new_title": str(result_row.get("new_title", "")).strip(),
                    "product_role": str(result_row.get("product_role", "")).strip(),
                    "title_status": str(result_row.get("title_status", "")).strip(),
                },
            }
        )
    write_json(FIXTURES / "title_snapshot.json", cases)
    return len(cases)


def main() -> None:
    if not INPUT.exists():
        raise SystemExit(f"Brak pliku wejsciowego: {INPUT}")
    FIXTURES.mkdir(parents=True, exist_ok=True)
    df = pd.read_excel(INPUT, sheet_name=SHEET, dtype=str, keep_default_na=False)
    extraction = generate_extraction_snapshot(df)
    titles = generate_title_snapshot(df)
    print(f"OK: extraction_snapshot.json ({extraction} cases)")
    print(f"OK: title_snapshot.json ({titles} cases)")


if __name__ == "__main__":
    main()
