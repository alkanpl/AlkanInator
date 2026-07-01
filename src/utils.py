from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
import yaml


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def load_yaml(path: str | Path) -> dict[str, Any]:
    yaml_path = Path(path)
    if not yaml_path.is_absolute():
        yaml_path = PROJECT_ROOT / yaml_path
    if not yaml_path.exists():
        return {}
    with yaml_path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def ensure_dir(path: str | Path) -> Path:
    directory = Path(path)
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def read_products(path: str | Path, sheet_name: str | int | None = None) -> pd.DataFrame:
    product_path = Path(path)
    suffix = product_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(product_path, dtype=str, keep_default_na=False, sep=None, engine="python")
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(product_path, sheet_name=sheet_name or 0, dtype=str, keep_default_na=False)
    raise ValueError(f"Nieobslugiwany format pliku: {suffix}. Uzyj CSV albo XLSX.")


def write_products(df: pd.DataFrame, path: str | Path) -> None:
    output_path = Path(path)
    ensure_dir(output_path.parent)
    suffix = output_path.suffix.lower()
    if suffix == ".csv":
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
        return
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        df.to_excel(output_path, index=False)
        return
    raise ValueError(f"Nieobslugiwany format eksportu: {suffix}. Uzyj CSV albo XLSX.")


def normalize_header(value: str) -> str:
    value = strip_accents(str(value)).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


# Rdzenie nazw kolorow obudowy - do porownan niezaleznych od formy gramatycznej
# (np. "szary"/"szara"/"szare" -> "szar"), uzywane przy wykrywaniu konfliktu koloru
# miedzy tytulem a atrybutem.
COLOR_STEMS = [
    "bial", "czarn", "szar", "srebrn", "grafit", "zlot", "bezow", "brazow",
    "chrom", "antracyt", "nikl", "satyn", "miedz", "zielon", "niebiesk", "czerwon",
]


def color_stem(value: str) -> str:
    """Rdzen pierwszego rozpoznanego koloru obudowy w tekscie ("Szary" -> "szar")."""
    normalized = normalize_header(value)
    for stem in COLOR_STEMS:
        if stem in normalized:
            return stem
    return ""


def strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    without_combining = "".join(char for char in normalized if not unicodedata.combining(char))
    return without_combining.translate(str.maketrans({"ł": "l", "Ł": "L"}))


def detect_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized_columns = {normalize_header(column): column for column in df.columns}
    for candidate in candidates:
        match = normalized_columns.get(normalize_header(candidate))
        if match:
            return match
    for normalized, original in normalized_columns.items():
        if any(normalize_header(candidate) in normalized for candidate in candidates):
            return original
    return None


def compact_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def is_blank(value: Any) -> bool:
    return value is None or compact_spaces(str(value)) == ""


def valid_ean(value: Any) -> bool:
    digits = re.sub(r"\D", "", str(value))
    if len(digits) not in {8, 12, 13, 14}:
        return False
    if len(set(digits)) == 1:
        return False
    return _gtin_check_digit_valid(digits)


def _gtin_check_digit_valid(digits: str) -> bool:
    # Suma kontrolna GTIN (EAN-8/UPC-12/EAN-13/GTIN-14): wagi 3 i 1 od prawej.
    checksum = sum(
        (3 if index % 2 == 0 else 1) * int(digit)
        for index, digit in enumerate(reversed(digits[:-1]))
    )
    expected = (10 - checksum % 10) % 10
    return expected == int(digits[-1])


def slugify(value: str) -> str:
    value = strip_accents(value).lower()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    return value.strip("-")


def first_present(*values: Any) -> str:
    for value in values:
        if not is_blank(value):
            return compact_spaces(str(value))
    return ""
