from __future__ import annotations

import argparse

from utils import read_products


def main() -> None:
    parser = argparse.ArgumentParser(description="Placeholder rekomendacji filtrow. Modul bedzie rozwijany po MVP.")
    parser.add_argument("--input", required=True)
    args = parser.parse_args()
    df = read_products(args.input)
    print(f"OK: plik odczytany, liczba kolumn: {len(df.columns)}")


if __name__ == "__main__":
    main()
