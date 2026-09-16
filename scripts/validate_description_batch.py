"""Sklada review XLSX i uruchamia walidator dla podanych slugow opisow.

Narzedzie pomocnicze: nie tworzy tresci opisow, tylko uruchamia istniejace
skrypty repo i zbiera wynik. Konsola Windows uzywa cp1250, wiec raport
zapisywany jest do pliku w utf-8.
"""
from __future__ import annotations

import argparse
import pathlib
import subprocess
import sys


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-dir", required=True, help="Katalog z podfolderami briefs/html/reviews.")
    parser.add_argument("--report", required=True)
    parser.add_argument("--slugs", nargs="*", help="Sluga bez rozszerzenia. Domyslnie wszystkie pliki z html/.")
    args = parser.parse_args()

    base = pathlib.Path(args.base_dir)
    slugs = args.slugs or sorted(p.stem for p in (base / "html").glob("*.html"))

    log: list[str] = []
    clean = 0
    for slug in slugs:
        brief = base / "briefs" / f"description_codex_brief_{slug}.xlsx"
        html = base / "html" / f"{slug}.html"
        review = base / "reviews" / f"review_{slug}.xlsx"
        if not brief.exists():
            log.append(f"--- {slug}: BRAK BRIEFU {brief.name}")
            continue
        r1 = subprocess.run(
            [sys.executable, "src/codex_description_agent.py", "--brief-input", str(brief),
             "--description-file", str(html), "--review-output", str(review)],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        r2 = subprocess.run(
            [sys.executable, "src/generate_product_descriptions.py", "--mode", "validate_codex_review",
             "--review-input", str(review)],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        out = r1.stdout + r1.stderr + r2.stdout + r2.stderr
        problems = sorted({l.strip() for l in out.splitlines() if "ERROR" in l or "WARNING" in l})
        if problems:
            log.append(f"--- {slug}: {len(problems)} problemow")
            log += ["      " + l for l in problems]
        else:
            clean += 1
            log.append(f"--- {slug}: OK")

    log.append("")
    log.append(f"PODSUMOWANIE: {clean}/{len(slugs)} bez uwag")
    pathlib.Path(args.report).write_text("\n".join(log), encoding="utf-8")
    print(f"OK: raport -> {args.report} ({clean}/{len(slugs)} bez uwag)")


if __name__ == "__main__":
    main()
