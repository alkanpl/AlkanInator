#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dwukierunkowa kontrola spojnosci instrukcji Codex <-> Claude.

Uruchom na starcie kazdego nowego czatu (Codex i Claude tak samo):

    py scripts\\check_agent_sync.py             (Windows)
    python3 scripts/check_agent_sync.py         (sesja Cowork / Linux)

Porownuje:
  1. skille Codexa (C:\\Users\\Handlowiec\\.codex\\skills) z kopia w .claude/skills,
     w OBIE strony - wykrywa zarowno nowy skill Codexa, jak i nowy skill Claude'a;
  2. naglowki sekcji AGENTS.md i CLAUDE.md, w obie strony;
  3. prompts/alkan_product_description_skill.md ze skillem alkan-product-description.

Kody wynikow:
  OK             zgodne
  ROZNI SIE      ten sam plik/sekcja, inna tresc - porownaj i scal
  TYLKO U CODEXA nowy lub usuniety po stronie Claude'a - skopiuj do .claude/skills
  TYLKO U CLAUDE nowy po stronie Claude'a - skopiuj do .codex/skills

Kod wyjscia 1, gdy cokolwiek wymaga uwagi.
"""
from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

CODEX_CANDIDATES = [
    Path(r"C:\Users\Handlowiec\.codex\skills"),
    Path.home() / ".codex" / "skills",
    Path.home() / "mnt" / "skills",   # Cowork: podpiety C:\Users\Handlowiec\.codex\skills
]


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def find_codex_skills(explicit: str | None) -> Path | None:
    if explicit:
        p = Path(explicit)
        return p if p.is_dir() else None
    for cand in CODEX_CANDIDATES:
        if cand.is_dir() and any(cand.glob("*/SKILL.md")):
            return cand
    return None


def rel_files(root: Path) -> dict[str, Path]:
    out: dict[str, Path] = {}
    if not root.is_dir():
        return out
    for f in root.rglob("*"):
        if not f.is_file():
            continue
        rel = f.relative_to(root).as_posix()
        if rel.startswith(".system/") or "/agents/" in rel or rel.endswith(".marker"):
            continue
        out[rel] = f
    return out


def headings(path: Path) -> list[str]:
    return [
        line.strip()
        for line in path.read_text(encoding="utf-8", errors="replace").splitlines()
        if line.startswith("## ")
    ]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--codex-skills", help="katalog skilli Codexa")
    ap.add_argument("--mirror", default=str(REPO / ".claude" / "skills"),
                    help="katalog kopii dla Claude'a")
    args = ap.parse_args()

    problems = 0
    mirror = Path(args.mirror)
    codex = find_codex_skills(args.codex_skills)

    print("== Skille: .codex/skills <-> .claude/skills ==")
    if codex is None:
        print("BRAK DOSTEPU: nie znaleziono katalogu skilli Codexa.")
        print("  Windows: C:\\Users\\Handlowiec\\.codex\\skills")
        print("  Cowork:  popros o dostep przez device_request_folder_access")
        problems += 1
    else:
        print(f"codex:  {codex}")
        print(f"claude: {mirror}")
        src = rel_files(codex)
        dst = rel_files(mirror)
        for rel in sorted(set(src) | set(dst)):
            if rel not in dst:
                print(f"TYLKO U CODEXA  {rel}  -> skopiuj do .claude/skills/{rel}")
                problems += 1
            elif rel not in src:
                print(f"TYLKO U CLAUDE  {rel}  -> skopiuj do .codex/skills/{rel}")
                problems += 1
            elif sha(src[rel]) != sha(dst[rel]):
                print(f"ROZNI SIE       {rel}  -> porownaj obie wersje i scal")
                problems += 1
            else:
                print(f"OK              {rel}")

    print()
    print("== Instrukcje repo: AGENTS.md <-> CLAUDE.md ==")
    agents = REPO / "AGENTS.md"
    claude_md = REPO / "CLAUDE.md"
    if not agents.is_file() or not claude_md.is_file():
        print("BRAK     AGENTS.md albo CLAUDE.md")
        problems += 1
    else:
        a_text = agents.read_text(encoding="utf-8", errors="replace")
        c_text = claude_md.read_text(encoding="utf-8", errors="replace")
        # Sekcje wlasne kazdego agenta - nie wymagaja odpowiednika po drugiej stronie.
        own = {"## Skille projektowe (kopia zywych skilli Codexa)"}
        missing_in_claude = [h for h in headings(agents) if h not in c_text and h not in own]
        missing_in_agents = [h for h in headings(claude_md) if h not in a_text and h not in own]
        for h in missing_in_claude:
            print(f"TYLKO W AGENTS.md  {h}  -> rozwaz dopisanie do CLAUDE.md")
            problems += 1
        for h in missing_in_agents:
            print(f"TYLKO W CLAUDE.md  {h}  -> rozwaz dopisanie do AGENTS.md")
            problems += 1
        if not missing_in_claude and not missing_in_agents:
            print("OK       oba pliki maja te same sekcje")

    skill_prompt = REPO / "prompts" / "alkan_product_description_skill.md"
    mirrored = mirror / "alkan-product-description" / "SKILL.md"
    if skill_prompt.is_file() and mirrored.is_file():
        if sha(skill_prompt) != sha(mirrored):
            print("ROZNI SIE  prompts/alkan_product_description_skill.md != skill alkan-product-description")
            problems += 1
        else:
            print("OK       prompts/alkan_product_description_skill.md zgodny ze skillem")

    print()
    print("WYNIK: wszystko aktualne" if problems == 0
          else f"WYNIK: {problems} pozycji do przejrzenia - zaktualizuj druga strone i powiadom uzytkownika")
    return 0 if problems == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
