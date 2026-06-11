from __future__ import annotations

import argparse
import re
from pathlib import Path

from generate_product_descriptions import read_description_input, write_codex_review_from_description


DEFAULT_REPORTS_DIR = "reports/descriptions"


def discover_latest_brief_batch(reports_dir: str | Path, sku: str = "") -> list[Path]:
    root = Path(reports_dir)
    if not root.exists():
        raise SystemExit(f"Nie znaleziono katalogu raportów: {root}")
    briefs = [
        path for path in root.rglob("description_codex_brief_*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not briefs:
        raise SystemExit(f"Nie znaleziono briefów Codex w: {root}")

    if sku:
        normalized_sku = normalize_sku(sku)
        matching = [path for path in briefs if normalize_sku(brief_sku(path)) == normalized_sku]
        if not matching:
            raise SystemExit(f"Nie znaleziono briefu dla SKU: {sku}")
        return [max(matching, key=lambda path: path.stat().st_mtime_ns)]

    batches: dict[Path, list[Path]] = {}
    for path in briefs:
        batches.setdefault(path.parent, []).append(path)
    latest_directory = max(
        batches,
        key=lambda directory: max(path.stat().st_mtime_ns for path in batches[directory]),
    )
    return sorted(batches[latest_directory], key=brief_sort_key)


def brief_sku(path: Path) -> str:
    return path.stem.removeprefix("description_codex_brief_")


def normalize_sku(value: str) -> str:
    return str(value).split("/", 1)[0].strip().upper()


def brief_sort_key(path: Path) -> tuple[int, int | str]:
    sku = brief_sku(path)
    if re.fullmatch(r"\d+", sku):
        return (0, int(sku))
    return (1, sku.lower())


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Składa review XLSX opisu Codex na podstawie briefu i gotowego HTML."
    )
    parser.add_argument("--list-briefs", action="store_true", help="Wypisuje najnowszy zestaw briefów do wykonania.")
    parser.add_argument("--reports-dir", default=DEFAULT_REPORTS_DIR, help="Katalog przeszukiwany rekursywnie po użyciu --list-briefs.")
    parser.add_argument("--sku", default="", help="Opcjonalnie ogranicza --list-briefs do najnowszego briefu danego SKU.")
    parser.add_argument("--brief-input", default="", help="XLSX brief z arkuszami Product facts/Rejected facts.")
    parser.add_argument("--review-output", default="", help="Docelowy XLSX review w głównym repo.")
    parser.add_argument("--description-html", default="", help="Gotowy HTML opisu. Dla długiego HTML preferuj --description-file.")
    parser.add_argument("--description-file", default="", help="Plik UTF-8 z gotowym HTML opisu.")
    args = parser.parse_args()

    if args.list_briefs:
        briefs = discover_latest_brief_batch(args.reports_dir, args.sku)
        print(f"BRIEF_BATCH_DIR={briefs[0].parent}")
        print(f"BRIEF_COUNT={len(briefs)}")
        for path in briefs:
            print(path)
        return

    if not args.brief_input:
        raise SystemExit("Podaj --brief-input albo użyj --list-briefs.")
    if not args.review_output:
        raise SystemExit("Podaj --review-output.")
    description_html = read_description_input(args.description_html, args.description_file)
    checks = write_codex_review_from_description(args.brief_input, args.review_output, description_html)
    for check in checks:
        print(f"{check['status']}: {check['check']} - {check['details']}")
    if any(check["status"] in {"ERROR", "WARNING"} for check in checks):
        raise SystemExit("Review zapisany, ale wymaga poprawy przed oddaniem.")
    print(f"OK: zapisano poprawny review Codex: {Path(args.review_output)}")


if __name__ == "__main__":
    main()
