# AGENTS.md

Wskazowki dla Codex przy pracy nad tym repozytorium.

## Czym jest AlkanInator

Lokalne narzedzie CLI (Python) do masowego przetwarzania katalogow produktow
elektrycznych/oswietleniowych z plikow XLSX/CSV: ekstrakcja atrybutow z tytulow,
generowanie tytulow SEO, wzbogacanie z eksportow WooCommerce/Baselinker,
rekomendacje kategorii i analiza slow kluczowych. Logika biznesowa siedzi w
konfiguracji YAML (`configs/`, `dictionaries/`), nie w kodzie.

Zasada nadrzedna: **oryginalne pliki wejsciowe nigdy nie sa nadpisywane** -
wyniki ida do `output/` i `reports/`. Kazda zmiana atrybutu niesie `source`,
`confidence` i `status` (`OK` / `WARNING` / `DO_SPRAWDZENIA`).

## Uruchamianie

Interpreter to `py` (Windows launcher) - `python` nie jest w PATH.
Skrypty uruchamia sie z katalogu glownego repo, importy w `src/` sa plaskie
(modul importuje `from utils import ...`), wiec `src/` musi byc na sciezce -
co zapewnia uruchomienie `py src/<skrypt>.py`.

Glowne wejscia:

- `py src/run_pipeline.py` - pipeline z ekstrakcja atrybutow z tytulu.
- `py src/run_pipeline_with_parameters.py` - pipeline z uzupelnianiem atrybutow
  z pliku `Parametry.xlsx`.

Oba dziela wspolny rdzen w `src/pipeline_core.py` (analiza wejscia,
klasyfikacja rol, slowa kluczowe, tytuly, walidacja kategorii, eksport).
Roznia sie tylko sposobem zbudowania kolumny roboczej `__working_title`.

## Architektura - mapa modulow

- `utils.py` - I/O (XLSX/CSV jako string, auto-separator), wykrywanie kolumn,
  walidacja EAN/GTIN, normalizacja tekstu.
- `analyze_products.py` - raport jakosci, wykrywanie kolumn, braki atrybutow.
- `extract_attributes.py` - rdzen ekstrakcji (regexy: moc, barwa, IP, strumien,
  gwint, wymiary...). Najbardziej kruchy modul - patrz testy charakteryzujace.
- `optimize_titles.py` + `title_anatomy.py` - generowanie tytulow SEO wedlug
  regul "anatomii tytulu" per typ produktu (`configs/title_anatomy.yaml`).
- `validate_output.py` - role produktow (glowny vs akcesorium), zgodnosc z
  kategoria.
- `catalog_knowledge.py` - wstrzykiwanie wiedzy z eksportu WooCommerce do configu.
- `dictionaries/supplier_global_knowledge.yaml` - globalna wiedza o dostawcach,
  obecnie przede wszystkim Kanlux: suffixy, zasady kategorii, tytulow,
  akcesoriow, atrybutow i bezpiecznych opisow. Przy zadaniach Kanlux czytaj ten
  plik razem z `docs/kanlux_workflow.md`.
- `analyze_keywords.py` - Google Ads Keyword Planner (zaleznosc opcjonalna,
  importowana leniwie; dziala tez offline z `--offline-keywords`).
- Pozostale `enrich_*`, `recommend_*`, `export_*`, `build_kobi_baselinker.py`,
  `learn_product_taxonomy.py` - narzedzia pomocnicze, kazde z wlasnym argparse.

## Testy

Testy uzywaja `unittest` (biblioteka standardowa, bez dodatkowych zaleznosci):

    py -m unittest discover -s tests -p "test_*.py"

Najwazniejsze sa **testy charakteryzujace** (golden-file) dla ekstrakcji i
generowania tytulow - utrwalaja aktualne zachowanie na realnych danych Kanlux,
zeby kazda niezamierzona regresja w regexach/regulach byla od razu widoczna.

Po **SWIADOMEJ** zmianie regul ekstrakcji, configow tytulow lub slownikow
nalezy zregenerowac snapshoty (wymaga `archived_input_files/input_2026-08-05/Alkan_Kanlux_pelne_rodziny.xlsx`):

    py tests/generate_snapshots.py

Fixtures w `tests/fixtures/` sa samowystarczalne - same testy nie potrzebuja
plikow XLSX.

## Konwencje

- Kod i wynikowe pliki: w wiekszosci bez polskich znakow w nazwach pol roboczych,
  ale dane produktowe i tytuly SEO sa po polsku (utf-8, CSV zapisywany jako
  `utf-8-sig` dla Excela).
- Konsola Windows uzywa cp1250 - nie polegaj na `print()` z polskimi znakami w
  skryptach diagnostycznych; zapisuj do pliku w utf-8.
- Status planu wdrozenia: `PLAN_ROBOCZY.md`.

## Zapamietane decyzje uzytkownika

- Przy pracy nad samodzielnymi plikami XLSX w tym repozytorium uzytkownik
  zezwolil na kontrolowany fallback do `openpyxl`, gdy wymagany runtime arkuszy
  nie jest dostepny. Nadal obowiazuje zakaz nadpisywania plikow wejsciowych.
- Dla produktow Hager/Berker pole `Atrybut Produktu: Seria` ma odzwierciedlac
  dokladna grupe przypisana do konkretnego SKU w katalogu. Kategorie moga byc
  szersze i laczyc serie zgodnie z aktualna lista w
  `input/BerkerHagerKategorie.xlsx`; jeden produkt moze nalezec do wielu
  kategorii. Przyklad: seria `R.1|R.3` pozostaje bez `R.8`, ale kategoria to
  `Berker R.1/R.3/R.8`.
- Dla Hager/Berker `B.1` nie jest seria i nie wolno jej zapisywac w atrybucie,
  tytule, opisie ani tagach. Samo slowo katalogowe `Glas` nie oznacza
  automatycznie serii `Glasserie`; `Glasserie` zachowuje sie tylko przy
  jednoznacznym przypisaniu konkretnego SKU w katalogu.
- Skroty w tytulach Hager/Berker nie zmieniaja wartosci atrybutu `Seria`:
  komplet `Lumina soul|Lumina intense|Lumina passion` zapisuje sie w tytule
  jako `Lumina`, komplet `B. Kwadrat|B.3|B.7` jako `B.X`, komplet
  `Q.1|Q.3|Q.7` jako `Q.X`, a komplet `R.1|R.3|R.8` jako `R.X`. Jezeli obok
  kompletnej grupy wystepuja inne serie, pozostaja jawnie widoczne w tytule.
- Standard tytulu Hager/Berker: nazwa/typ produktu i kolor na poczatku, potem
  marka i seria, a kod produktu na koncu. Kazdy tytul dotyczacy serii Lumina ma
  zawierac marke `Hager`; aliasy katalogowe w nawiasach, np. `(Berker 10107600)`,
  pozostaja zachowane.
- Dla Hager/Berker sposob montazu ustala `dictionaries/hager_berker_mounting_rules.yaml`
  wraz z `src/build_hager_berker_mounting_ip.py`. Pole `Sposob montazu` z eksportu
  BMEcat bywa bledne (lacznikom Lumina przypisuje `Natynkowy`, a hager.com podaje
  `Montaz podtynkowy`), dlatego mocowanie pazurkami rozporowymi ma pierwszenstwo
  przed tym polem. Dla ramek, klawiszy, plytek czolowych i innych elementow
  nakladanych sklep nie wypelnia sposobu montazu.

## Synchronizacja instrukcji Codex <-> Claude

Ten projekt prowadza rownolegle Codex i Claude. Oba maja miec te sama wiedze
robocza. Skille Codexa z `C:\Users\Handlowiec\.codex\skills` sa lustrzanie
kopiowane do `.claude/skills/` w tym repo, a wspolne reguly zyja w `AGENTS.md`
i `CLAUDE.md` naraz.

Na poczatku kazdego nowego czatu uruchom:

    py scripts\check_agent_sync.py

Skrypt porownuje w OBIE strony: skille Codexa z kopia w `.claude/skills/`,
naglowki sekcji `AGENTS.md` i `CLAUDE.md`, oraz
`prompts/alkan_product_description_skill.md` ze skillem
`alkan-product-description`.

Co zrobic z wynikiem:

- `TYLKO U CLAUDE` - Claude dodal albo zmienil skill. Przeczytaj plik i skopiuj
  go do `C:\Users\Handlowiec\.codex\skills\<nazwa>\`.
- `TYLKO U CODEXA` - twoja wersja jest nowsza niz kopia. Zaktualizuj
  `.claude/skills/<nazwa>/`.
- `ROZNI SIE` - porownaj obie wersje i scal swiadomie, nie nadpisuj w ciemno.
- `TYLKO W CLAUDE.md` - dopisz brakujaca regule do `AGENTS.md` (i odwrotnie).

Zawsze powiedz uzytkownikowi, co sie zmienilo, zanim ruszysz z wlasciwym
zadaniem.

Kiedy sam dodajesz lub zmieniasz skill albo regule: po zapisaniu w
`C:\Users\Handlowiec\.codex\skills\` skopiuj go do `.claude/skills/<nazwa>/`, a
reguly ogolne dopisz do OBU plikow - `AGENTS.md` i `CLAUDE.md`. Dzieki temu
Claude zobaczy zmiane juz w swoim nastepnym czacie.
