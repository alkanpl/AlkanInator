"""Sklada podglad HTML gotowych opisow do przegladu w przegladarce.

Narzedzie pomocnicze: sklada istniejace pliki HTML w jeden dokument.
Nie generuje tresci opisow.
"""
from __future__ import annotations

import argparse
import html as html_mod
import pathlib
import re

PAGE_HEAD = """<!doctype html>
<meta charset="utf-8">
<title>Podglad opisow - Kontakt Simon kaszmir</title>
<style>
 body{font:15px/1.6 system-ui,Segoe UI,sans-serif;margin:0;background:#f6f6f4;color:#1c1c1a}
 .wrap{max-width:860px;margin:0 auto;padding:24px 16px}
 .item{background:#fff;border:1px solid #ddd;border-radius:8px;padding:20px 24px;margin:0 0 28px}
 .sku{font:13px ui-monospace,Consolas,monospace;color:#666;border-bottom:1px solid #eee;
      padding-bottom:8px;margin-bottom:14px}
 .meta{font-size:12px;color:#777;margin-top:16px;border-top:1px solid #eee;padding-top:8px}
 h1{font-size:22px} h2{font-size:19px;margin-top:22px} h3{font-size:16px;margin-top:18px}
 ul{padding-left:22px} li{margin:3px 0}
</style>
<div class="wrap">
<h1>Podglad opisow - Kontakt Simon, kaszmir</h1>
"""


def visible_chars(markup: str) -> int:
    text = re.sub(r"<[^>]+>", "", markup)
    return len(text.strip())


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--html-dir", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    files = sorted(pathlib.Path(args.html_dir).glob("*.html"))
    parts = [PAGE_HEAD, f"<p>Opisow w podgladzie: <strong>{len(files)}</strong></p>"]
    for path in files:
        body = path.read_text(encoding="utf-8")
        chars = visible_chars(body)
        parts.append('<div class="item">')
        parts.append(f'<div class="sku">{html_mod.escape(path.stem)}</div>')
        parts.append(body)
        parts.append(f'<div class="meta">Znakow widocznych ze spacjami: {chars}</div>')
        parts.append("</div>")
    parts.append("</div>")

    pathlib.Path(args.output).write_text("\n".join(parts), encoding="utf-8")
    print(f"OK: podglad -> {args.output} ({len(files)} opisow)")


if __name__ == "__main__":
    main()
