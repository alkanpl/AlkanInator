"""Doklada napisany recznie blok FAQ + JSON-LD do gotowego pliku opisu.

To NIE jest generator tresci: pytania i odpowiedzi przychodza gotowe z JSON-a
wejsciowego, napisane recznie dla konkretnego SKU. Skrypt tylko skleja je w
ustalony w repo format (<hr>, naglowek, 3 pary <p>, schema FAQPage) i pilnuje,
zeby tresc widoczna byla identyczna z JSON-LD.
"""
from __future__ import annotations

import argparse
import json
import pathlib
import re

HEADING = "Najczęściej zadawane pytania"


def build_block(faq: list[dict[str, str]]) -> str:
    lines = ["<hr>", "", f"<h3>{HEADING}</h3>", ""]
    for index, item in enumerate(faq, 1):
        lines.append(f'<p><strong>{index}. {item["q"]}</strong><br>')
        lines.append(f'{item["a"]}</p>')
        lines.append("")
    schema = {
        "@context": "https://schema.org",
        "@type": "FAQPage",
        "mainEntity": [
            {
                "@type": "Question",
                "name": item["q"],
                "acceptedAnswer": {"@type": "Answer", "text": item["a"]},
            }
            for item in faq
        ],
    }
    lines.append('<script type="application/ld+json">')
    lines.append(json.dumps(schema, ensure_ascii=False, indent=2))
    lines.append("</script>")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--html", required=True, help="Plik opisu do uzupelnienia.")
    parser.add_argument("--faq-json", required=True, help="JSON: lista {q, a} napisana recznie.")
    args = parser.parse_args()

    faq = json.loads(pathlib.Path(args.faq_json).read_text(encoding="utf-8"))
    if len(faq) != 3:
        raise SystemExit(f"FAQ ma miec 3 pozycje, jest {len(faq)}")
    for item in faq:
        length = len(item["a"].strip())
        if not 200 <= length <= 300:
            raise SystemExit(f"Odpowiedz ma {length} znakow (wymagane 200-300): {item['q']}")

    path = pathlib.Path(args.html)
    text = path.read_text(encoding="utf-8")
    text = re.sub(r"\n<hr>.*$", "\n", text, flags=re.DOTALL).rstrip()
    path.write_text(text + "\n\n" + build_block(faq) + "\n", encoding="utf-8")
    print(f"OK: FAQ dopisane do {path.name}")


if __name__ == "__main__":
    main()
