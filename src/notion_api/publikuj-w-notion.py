#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Publikacja dokumentów AlkanInatora (changelog, instrukcja) na stronach Notion.

Każdy dokument z mapy DOKUMENTY w całości zastępuje treść przypisanej strony
(PATCH /v1/pages/{id}/markdown, replace_content), a tytuł strony dostaje
nagłówek H1 dokumentu z numerem wersji („… v1.14”). Publikowana jest treść
z commita (git show), nigdy niezacommitowana praca. Linki między dokumentami
z mapy prowadzą do ich stron w Notion, pozostałe linki względne — na GitHub.

Po każdym commicie uruchamia go hook .githooks/post-commit (włączenie:
git config core.hooksPath .githooks). Ręcznie:
    py -I src/notion_api/publikuj-w-notion.py [--commit SHA] [--wszystkie] [--podglad]

Token wewnętrznej integracji Notion: zmienna NOTION_TOKEN albo {"token": "..."}
w src/notion_api/config.json (ignorowany przez git). Integrację trzeba dodać do
strony (••• → Połączenia). Wymaga: Python 3.9+, bez dodatkowych pakietów.
"""
import argparse
import datetime
import json
import os
import posixpath
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Dokument w repo → identyfikator strony Notion (z adresu strony, ostatnie 32 znaki).
# Pusty identyfikator = strona jeszcze nie założona, dokument jest pomijany.
DOKUMENTY = {
    "docs/changelog.md": "40f0af326a0880f5bfcaec9f92332d40",
    "docs/instrukcja.md": "3e20af326a08800596c5fc070d1a502f",
}

# Linki względne prowadzą do plików repo, których w Notion nie ma — kierujemy je na GitHub.
# Wyjątek: dokumenty z mapy DOKUMENTY, które linkują do siebie nawzajem, dostają adres strony Notion.
GITHUB = "https://github.com/alkanpl/AlkanInator/blob/main/"
STRONA_NOTION = "https://app.notion.com/p/"

API = "https://api.notion.com/v1"
NOTION_VERSION = "2026-03-11"
PLIK_TOKENU = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# Języki bloków kodu bez odpowiednika w Notion — pusty fence daje „plain text”.
JEZYKI_KODU = {"text": "", "properties": ""}

# Znaki, które enhanced markdown Notion traktuje specjalnie, a w dokumencie są zwykłym tekstem
# (np. „> 0”, „$” uruchamiałby matematykę inline). Znaki składni markdown (* [ ] `) zostają.
ZNAKI_DO_ESCAPE = "<>{}$^~"

# Polskie skróty, które Notion bierze za adresy i linkuje („m.in.” → http://m.in, .in to domena Indii).
# Niewidoczny „word joiner” po kropce rozbija adres, a tekst na stronie wygląda tak samo.
SKROTY_JAK_ADRES = re.compile(r"\b(m\.in|n\.p\.m|p\.n\.e)\.")
WORD_JOINER = "⁠"

# Szerokości kolumn tabel w px. Notion bierze <col width> dosłownie i nie rozciąga tabeli do strony,
# a dokumentacja jest zawsze w trybie „Full width” — ok. 1815 px treści na ekranie autora
# (standardowa strona ma 708 px). Mały zapas, żeby nie pojawiło się przewijanie w poziomie.
SZEROKOSC_STRONY = 1800
# Przybliżenie szerokości znaku przy domyślnej czcionce tabeli plus odstępy wewnątrz komórki
PIKSELE_NA_ZNAK = 8
MARGINES_KOMORKI = 24
MIN_KOLUMNY = 60
MAKS_MINIMUM = 240
KROK_KOLUMNY = 4

WERSJA = re.compile(r"\*\*Wersja:\s*([0-9][0-9.]*)\*\*")
LINK = re.compile(r"\[([^\]]*)\]\(([^)\s]+)\)")
KOD_INLINE = re.compile(r"(`[^`]*`)")
CYTAT = re.compile(r"^\s*(>\s?)*")
SEPARATOR_TABELI = re.compile(r"^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$")


def log(komunikat, plik=sys.stdout):
    print("[%s] %s" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), komunikat), file=plik, flush=True)


def git(*argumenty):
    return subprocess.run(["git", "-C", REPO] + list(argumenty), check=True,
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE).stdout.decode("utf-8")


def zmienione_pliki(sha):
    # --root: pierwszy commit repo też ma listę plików
    wynik = git("diff-tree", "--no-commit-id", "--name-only", "-r", "--root", "-z", sha)
    return [p for p in wynik.split("\0") if p]


# ---------------------------------------------------------------------------
# Przekształcenie markdown repo → enhanced markdown Notion
# ---------------------------------------------------------------------------

def przepisz_link(dopasowanie, sciezka_dokumentu):
    tekst, cel = dopasowanie.group(1), dopasowanie.group(2)
    if cel.startswith("#"):
        # Notion nie obsługuje kotwic do nagłówków — zostaje sam tekst
        return tekst
    if re.match(r"^[a-z][a-z0-9+.-]*:", cel, re.IGNORECASE):
        return dopasowanie.group(0)
    plik, _, fragment = cel.partition("#")
    sciezka = posixpath.normpath(posixpath.join(posixpath.dirname(sciezka_dokumentu), plik))
    if DOKUMENTY.get(sciezka):
        # Kotwica nagłówka z repo nie zadziała na stronie Notion — link prowadzi na początek strony
        return "[%s](%s%s)" % (tekst, STRONA_NOTION, DOKUMENTY[sciezka])
    return "[%s](%s%s%s)" % (tekst, GITHUB, sciezka, "#" + fragment if fragment else "")


def przeksztalc_tekst(linia, sciezka_dokumentu):
    # „>” na początku linii to cytat markdown, a nie znak do escape
    cytat = CYTAT.match(linia).group(0)
    linia = LINK.sub(lambda d: przepisz_link(d, sciezka_dokumentu), linia[len(cytat):])
    # Escape tylko poza `inline code` — zawartość kodu Notion bierze dosłownie
    czesci = KOD_INLINE.split(linia)
    for i in range(0, len(czesci), 2):
        czesci[i] = "".join("\\" + z if z in ZNAKI_DO_ESCAPE else z for z in czesci[i])
        czesci[i] = SKROTY_JAK_ADRES.sub(lambda d: d.group(0).replace(".", "." + WORD_JOINER), czesci[i])
    return cytat + "".join(czesci)


def komorki(wiersz):
    """Dzieli wiersz tabeli pipe na komórki, pomijając | wewnątrz `kodu` i \\|."""
    komorki, biezaca, w_kodzie, poprzedni = [], "", False, ""
    for znak in wiersz.strip():
        if znak == "`":
            w_kodzie = not w_kodzie
        if znak == "|" and not w_kodzie and poprzedni != "\\":
            komorki.append(biezaca)
            biezaca = ""
        else:
            biezaca += znak
        poprzedni = znak
    komorki.append(biezaca)
    # Skrajne | dają puste komórki na początku i końcu
    if komorki and not komorki[0].strip():
        komorki = komorki[1:]
    if komorki and not komorki[-1].strip():
        komorki = komorki[:-1]
    return [k.strip() for k in komorki]


def widoczny_tekst(komorka):
    # Długość liczona po tym, co widać w Notion: bez adresów linków i znaczników markdown
    komorka = LINK.sub(lambda d: d.group(1), komorka)
    return komorka.replace("**", "").replace("`", "").replace("\\|", "|")


def szerokosci_kolumn(wiersze_komorek):
    """Szerokości kolumn w px, sumujące się do SZEROKOSC_STRONY.

    Bez <colgroup> Notion dobiera szerokości po swojemu i kolumna z opisem bywa tak wąska jak
    kolumna z numerem wersji. Każda kolumna startuje od szerokości najdłuższego słowa, a resztę
    strony rozdajemy po KROK_KOLUMNY tam, gdzie najbardziej zmniejsza to szacowaną wysokość
    tabeli. Wysokość wiersza wyznacza komórka z największą liczbą linii, więc krótkie kolumny
    obok długiego opisu mogą się zawinąć — i tak nie podwyższą wiersza.
    """
    n = max(len(w) for w in wiersze_komorek)
    teksty = [[widoczny_tekst(w[i]) if i < len(w) else "" for i in range(n)] for w in wiersze_komorek]
    dlugosci = [[len(t) for t in wiersz] for wiersz in teksty]
    # Najdłuższe słowo nie powinno się łamać, ale nazwy plików potrafią mieć 50+ znaków — stąd limit
    minimum = [max(MIN_KOLUMNY, min(MAKS_MINIMUM, max((len(s) for wiersz in teksty for s in wiersz[i].split()), default=0)
                                    * PIKSELE_NA_ZNAK + MARGINES_KOMORKI)) for i in range(n)]

    def wysokosc(szerokosci):
        # Ułamkowa liczba linii zamiast zaokrąglonej w górę, żeby każdy krok coś zmieniał
        return sum(max(max(1.0, d * PIKSELE_NA_ZNAK / (s - MARGINES_KOMORKI)) for d, s in zip(wiersz, szerokosci))
                   for wiersz in dlugosci)

    def gestosc(i, szerokosci):
        return sum(wiersz[i] for wiersz in dlugosci) / szerokosci[i]

    szerokosci = list(minimum)
    while sum(szerokosci) + KROK_KOLUMNY <= SZEROKOSC_STRONY:
        kandydaci = [szerokosci[:i] + [szerokosci[i] + KROK_KOLUMNY] + szerokosci[i + 1:] for i in range(n)]
        # Przy remisie (wszystko już w jednej linii) poszerzamy kolumnę z największą ilością tekstu na piksel
        najlepsza = min(range(n), key=lambda i: (wysokosc(kandydaci[i]), -gestosc(i, szerokosci)))
        szerokosci = kandydaci[najlepsza]

    # Dopasowanie do dokładnie SZEROKOSC_STRONY — także gdy same minima się nie mieszczą
    wynik = [int(s * SZEROKOSC_STRONY / sum(szerokosci)) for s in szerokosci]
    wynik[wynik.index(max(wynik))] += SZEROKOSC_STRONY - sum(wynik)
    return wynik


def tabela_html(wiersze, sciezka_dokumentu):
    # Tabele pipe zamieniamy na formę <table> udokumentowaną w enhanced markdown —
    # dla tabel pipe Notion nie gwarantuje parsowania na wejściu.
    # fit-page-width sam nie rozciąga tabeli na stronie „Full width” (sprawdzone 2026-10-20) —
    # szerokość wyznaczają <col width>; atrybut zostaje dla zgodności z opcją „Fit to page width”.
    naglowek = len(wiersze) > 1 and SEPARATOR_TABELI.match(wiersze[1])
    dane = [wiersze[0]] + wiersze[2:] if naglowek else wiersze
    wiersze_komorek = [komorki(w) for w in dane]
    wynik = ['<table fit-page-width="true" header-row="%s">' % ("true" if naglowek else "false"), "\t<colgroup>"]
    wynik += ['\t\t<col width="%d">' % s for s in szerokosci_kolumn(wiersze_komorek)]
    wynik.append("\t</colgroup>")
    for komorki_wiersza in wiersze_komorek:
        wynik.append("\t<tr>")
        for komorka in komorki_wiersza:
            wynik.append("\t\t<td>%s</td>" % przeksztalc_tekst(komorka, sciezka_dokumentu))
        wynik.append("\t</tr>")
    wynik.append("</table>")
    return wynik


def przygotuj(sciezka_dokumentu, tresc, sha, data_commita):
    """Zwraca (tytuł strony, treść w enhanced markdown)."""
    linie = tresc.splitlines()
    indeks_h1 = next((i for i, l in enumerate(linie) if l.startswith("# ")), None)
    if indeks_h1 is None:
        raise ValueError("brak nagłówka H1 w %s" % sciezka_dokumentu)
    tytul = linie[indeks_h1][2:].strip()
    wersja = WERSJA.search(tresc)
    if wersja:
        tytul += " v" + wersja.group(1)
    linie = linie[:indeks_h1] + linie[indeks_h1 + 1:]

    wynik = [
        "*Strona generowana automatycznie z repozytorium AlkanInator ([%s](%s%s), commit `%s`, %s). "
        "Ręczne zmiany w Notion zostaną nadpisane przy kolejnym commicie.*"
        % (sciezka_dokumentu, GITHUB, sciezka_dokumentu, sha[:7], data_commita),
        "",
    ]
    w_bloku_kodu = False
    tabela = []
    for linia in linie:
        if tabela and not linia.lstrip().startswith("|"):
            wynik += tabela_html(tabela, sciezka_dokumentu)
            tabela = []
        if linia.lstrip().startswith("```"):
            if not w_bloku_kodu:
                jezyk = linia.lstrip()[3:].strip()
                linia = "```" + JEZYKI_KODU.get(jezyk, jezyk)
            w_bloku_kodu = not w_bloku_kodu
            wynik.append(linia)
        elif w_bloku_kodu:
            wynik.append(linia)
        elif linia.lstrip().startswith("|"):
            tabela.append(linia)
        else:
            wynik.append(przeksztalc_tekst(linia, sciezka_dokumentu))
    if tabela:
        wynik += tabela_html(tabela, sciezka_dokumentu)
    return tytul, "\n".join(wynik).strip() + "\n"


# ---------------------------------------------------------------------------
# API Notion
# ---------------------------------------------------------------------------

class BladNotion(Exception):
    def __init__(self, status, tresc):
        super().__init__("HTTP %s: %s" % (status, tresc))
        self.status = status
        self.tresc = tresc


def wczytaj_token():
    token = os.environ.get("NOTION_TOKEN", "").strip()
    if token or not os.path.isfile(PLIK_TOKENU):
        return token
    with open(PLIK_TOKENU, encoding="utf-8-sig") as plik:
        return str(json.load(plik).get("token") or "").strip()


def notion(metoda, sciezka, token, dane=None):
    zadanie = urllib.request.Request(
        API + sciezka if sciezka.startswith("/") else sciezka,
        method=metoda,
        data=json.dumps(dane).encode("utf-8") if dane is not None else None,
        headers={
            "Authorization": "Bearer " + token,
            "Notion-Version": NOTION_VERSION,
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(zadanie, timeout=180) as odpowiedz:
            return json.loads(odpowiedz.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as blad:
        raise BladNotion(blad.code, blad.read().decode("utf-8", "replace"))


def czekaj_na_zadanie(zadanie, token):
    # Zadanie asynchroniczne: queued/running/retrying → odpytujemy, aż będzie succeeded albo failed
    koniec = time.time() + 600
    while zadanie.get("status") not in ("succeeded", "failed"):
        if time.time() > koniec:
            raise RuntimeError("zadanie %s nie zakończyło się w 10 minut" % zadanie.get("id"))
        time.sleep(max(2, zadanie.get("poll_after_seconds") or 2))
        zadanie = notion("GET", zadanie["status_url"], token)
    if zadanie["status"] == "failed":
        raise BladNotion(zadanie.get("error", {}).get("status", "?"), json.dumps(zadanie.get("error"), ensure_ascii=False))
    return zadanie.get("result", {})


def zastap_tresc(id_strony, markdown, token):
    dane = {"type": "replace_content", "replace_content": {"new_str": markdown}}
    try:
        return notion("PATCH", "/pages/%s/markdown" % id_strony, token, dane)
    except BladNotion as blad:
        # Treść za duża na jedno wywołanie synchroniczne — ponowienie w trybie asynchronicznym
        if blad.status != 400 or "too large" not in blad.tresc.lower():
            raise
    log("treść za duża na jedno wywołanie, ponawiam asynchronicznie")
    dane["allow_async"] = True
    return czekaj_na_zadanie(notion("PATCH", "/pages/%s/markdown" % id_strony, token, dane), token)


def publikuj(id_strony, tytul, markdown, token):
    strona = notion("GET", "/pages/%s" % id_strony, token)
    # Strona leży w wiki (bazie), więc nazwa kolumny tytułu zależy od jej schematu
    klucz_tytulu = next(k for k, w in strona["properties"].items() if w.get("type") == "title")

    # Bez allow_deleting_content: jeśli strona ma podstrony lub bazy, Notion odmówi zamiast je skasować
    wynik = zastap_tresc(id_strony, markdown, token)
    for ostrzezenie in wynik.get("warnings") or []:
        log("ostrzeżenie Notion: %s — %s" % (ostrzezenie.get("code"), ostrzezenie.get("message")), sys.stderr)

    # Tytuł dopiero po udanej treści, żeby numer wersji nie wyprzedził treści strony
    notion("PATCH", "/pages/%s" % id_strony, token,
           {"properties": {klucz_tytulu: {"title": [{"type": "text", "text": {"content": tytul}}]}}})


def main():
    # Konsola Windows (cp1250) nie ma znaków typu „→” — zamiast błędu wypisujemy „?”
    for strumien in (sys.stdout, sys.stderr):
        strumien.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="Publikuje dokumenty repo na stronach wiki Notion")
    ap.add_argument("--commit", default="HEAD", help="commit, z którego brana jest treść (domyślnie HEAD)")
    ap.add_argument("--wszystkie", action="store_true", help="publikuje wszystkie zmapowane dokumenty, nie tylko zmienione w commicie")
    ap.add_argument("--podglad", action="store_true", help="wypisuje tytuł i treść bez wysyłania do Notion")
    args = ap.parse_args()

    sha = git("rev-parse", args.commit).strip()
    zmienione = set(zmienione_pliki(sha))
    dokumenty = [d for d in DOKUMENTY if args.wszystkie or d in zmienione]
    if not args.podglad:
        for sciezka in [d for d in dokumenty if not DOKUMENTY[d]]:
            log("pomijam %s: brak strony Notion w mapie DOKUMENTY" % sciezka, sys.stderr)
        dokumenty = [d for d in dokumenty if DOKUMENTY[d]]
    if not dokumenty:
        return 0

    token = None
    if not args.podglad:
        token = wczytaj_token()
        if not token:
            log("brak tokenu Notion: ustaw NOTION_TOKEN albo klucz token w %s "
                "(wewnętrzna integracja Notion dodana do strony przez ••• → Połączenia)" % PLIK_TOKENU, sys.stderr)
            return 1

    data_commita = git("show", "-s", "--format=%cs", sha).strip()
    bledy = 0
    for sciezka in dokumenty:
        try:
            tytul, markdown = przygotuj(sciezka, git("show", "%s:%s" % (sha, sciezka)), sha, data_commita)
            if args.podglad:
                print("=== %s → strona %s\nTytuł: %s\n\n%s" % (sciezka, DOKUMENTY[sciezka] or "(brak)", tytul, markdown))
                continue
            publikuj(DOKUMENTY[sciezka], tytul, markdown, token)
            log("opublikowano %s (commit %s) → „%s”" % (sciezka, sha[:7], tytul))
        except (BladNotion, RuntimeError, ValueError, StopIteration, subprocess.CalledProcessError, urllib.error.URLError) as blad:
            bledy += 1
            log("błąd publikacji %s (commit %s): %s" % (sciezka, sha[:7], blad), sys.stderr)
    return 1 if bledy else 0


if __name__ == "__main__":
    sys.exit(main())
