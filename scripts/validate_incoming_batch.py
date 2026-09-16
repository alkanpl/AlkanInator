"""Waliduje przychodzaca paczke opisow: review XLSX + walidator + kontrola dodatkowa.

Uzywane po kazdym powrocie opisow pisanych poza repo. Poza standardowym
walidatorem sprawdza rzeczy, ktorych on nie lapie, a ktore juz sie zdarzaly:
  - slowa z kuchni projektu ("brief", "material zrodlowy") w tekscie dla klienta,
  - stary format punktu listy z myslnikiem po <li>.

Raport idzie do pliku w utf-8 - konsola Windows uzywa cp1250.
"""
from __future__ import annotations

import argparse
import pathlib
import re
import subprocess
import sys

# Slownictwo wewnetrzne, ktore nigdy nie moze trafic do opisu dla klienta.
_INTERNAL_WORDS = re.compile(r"\b(brief\w*|materia[łl] [źz]r[óo]d[łl]owy|dane wej[śs]ciowe)\b", re.I)
_OLD_BULLET = re.compile(r"<li>\s*-\s")


def extra_checks(markup: str) -> list[str]:
    issues: list[str] = []
    text = re.sub(r"<[^>]+>", " ", markup)
    for match in sorted({m.group(0) for m in _INTERNAL_WORDS.finditer(text)}):
        issues.append(f"WARNING: internal_vocabulary - slowo z kuchni projektu w tekscie: '{match}'")
    if _OLD_BULLET.search(markup):
        issues.append("WARNING: bullet_format - punkt listy zaczyna sie myslnikiem po <li>")
    return issues


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-dir", required=True)
    parser.add_argument("--html-dir", required=True, help="Katalog z przychodzacymi opisami.")
    parser.add_argument("--reviews-dir", required=True)
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    base = pathlib.Path(args.base_dir)
    html_dir = pathlib.Path(args.html_dir)
    reviews_dir = pathlib.Path(args.reviews_dir)
    reviews_dir.mkdir(parents=True, exist_ok=True)

    log: list[str] = []
    clean = 0
    slugs = sorted(p.stem for p in html_dir.glob("*.html"))
    for slug in slugs:
        brief = base / "briefs" / f"description_codex_brief_{slug}.xlsx"
        html = html_dir / f"{slug}.html"
        review = reviews_dir / f"review_{slug}.xlsx"
        if not brief.exists():
            log.append(f"--- {slug}: BRAK BRIEFU")
            continue
        first = subprocess.run(
            [sys.executable, "src/codex_description_agent.py", "--brief-input", str(brief),
             "--description-file", str(html), "--review-output", str(review)],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        second = subprocess.run(
            [sys.executable, "src/generate_product_descriptions.py", "--mode", "validate_codex_review",
             "--review-input", str(review)],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        output = first.stdout + first.stderr + second.stdout + second.stderr
        problems = sorted({line.strip() for line in output.splitlines()
                           if "ERROR" in line or "WARNING" in line})
        problems += extra_checks(html.read_text(encoding="utf-8"))
        if problems:
            log.append(f"--- {slug}: {len(problems)}")
            log += ["      " + p for p in problems]
        else:
            clean += 1
            log.append(f"--- {slug}: OK")

    log.append("")
    log.append(f"PODSUMOWANIE: {clean}/{len(slugs)} bez uwag")
    pathlib.Path(args.report).write_text("\n".join(log), encoding="utf-8")
    print(f"OK: raport -> {args.report} ({clean}/{len(slugs)} bez uwag)")


if __name__ == "__main__":
    main()
