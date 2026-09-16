"""Rozbija zbiorczy podglad HTML opisow na pojedyncze pliki per SKU.

Sluzy do przyjmowania opisow poprawionych poza repo (np. w zewnetrznym czacie),
ktore wracaja jako jeden plik podgladu zlozony z blokow:

    <div class="item"><div class="sku">SLUG</div> ...opis... </div>

Skrypt niczego nie pisze ani nie poprawia - tylko wycina zawartosc kazdego
bloku, usuwa elementy samego podgladu (naglowek SKU, stopka .meta) i zapisuje
opis do <output-dir>/<SLUG>.html w utf-8.

Oryginalny plik wejsciowy nie jest modyfikowany.
"""
from __future__ import annotations

import argparse
import pathlib
import re

ITEM_RE = re.compile(r'<div class="item">(.*?)</div>\s*(?=<div class="item">|</div>\s*$|$)', re.S)
SKU_RE = re.compile(r'<div class="sku">\s*(.*?)\s*</div>', re.S)
META_RE = re.compile(r'<div class="meta">.*?</div>\s*', re.S)


def split_items(markup: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    for chunk in markup.split('<div class="item">')[1:]:
        sku_match = SKU_RE.search(chunk)
        if not sku_match:
            continue
        slug = sku_match.group(1).strip()
        body = SKU_RE.sub("", chunk, count=1)
        body = META_RE.sub("", body)
        # Domknij blok na ostatnim elemencie opisu, odcinajac ogon podgladu.
        for tail in ("</script>", "</p>", "</ul>"):
            end = body.rfind(tail)
            if end != -1:
                body = body[: end + len(tail)]
                break
        items.append((slug, body.strip() + "\n"))
    return items


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Zbiorczy plik podgladu HTML.")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--report", default="", help="Opcjonalny raport co zostalo zapisane.")
    args = parser.parse_args()

    markup = pathlib.Path(args.input).read_text(encoding="utf-8")
    items = split_items(markup)
    if not items:
        raise SystemExit("Nie znaleziono blokow <div class=\"item\"> z <div class=\"sku\">.")

    out_dir = pathlib.Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    lines = [f"ZRODLO: {args.input}", f"OPISOW: {len(items)}", ""]
    for slug, body in items:
        path = out_dir / f"{slug}.html"
        path.write_text(body, encoding="utf-8")
        visible = len(re.sub(r"<[^>]+>", "", body.split("<hr>")[0]).strip())
        lines.append(f"  {slug:24} {len(body):6} B  znakow widocznych w opisie: {visible}")
        print(f"OK: {path}")

    if args.report:
        pathlib.Path(args.report).write_text("\n".join(lines), encoding="utf-8")
    print(f"Zapisano {len(items)} opisow -> {out_dir}")


if __name__ == "__main__":
    main()
