---
name: ver-upd
description: Summarize important Git changes for version/update notes. Use when the user writes "ver-upd", asks for "najwazniejsze zmiany", "zmiany od ostatniej wersji", "miedzy mainem a wersja", release notes, changelog summary, or a concise business-readable summary of Git differences.
---

# Ver Upd

## Workflow

1. Determine the comparison range explicitly.
   - If the user gives a range, use it.
   - If not, compare the latest version tag to `main`: `git tag --sort=-creatordate`, then `git diff <latest-tag>..main`.
   - If `main` and the current branch are the same and there is no tag context, say that directly and check the working tree separately.

2. Gather evidence with Git before summarizing:
   - `git status --short`
   - `git branch --show-current`
   - `git log --oneline --decorate <range>`
   - `git diff --stat <range>`
   - `git diff --name-status <range>`
   - Read targeted diffs for files with meaningful changes.

3. Produce a concise Polish summary focused on user-visible or operational impact.
   Prefer categories such as:
   - Koszyk i checkout
   - Filtrowanie / lista produktów
   - Elektrotargi / landing
   - Style i UX
   - Techniczne / porządki

4. Mention the exact range used, for example `v1.3.4..main`.
   If there are uncommitted changes, list them separately as "lokalne niecommitowane".

5. Avoid dumping raw diffs unless the user asks. Keep the first answer high-signal: the most important changes, then touched files/commits only if helpful.
