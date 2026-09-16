# CLAUDE.md

Wskazowki dla Claude Code przy pracy nad tym repozytorium.

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

Lancuch poprawek atrybutow Kanlux (Woo + Baselinker):
`import_attribute_knowledge.py` (macierz "Poprawione atrybuty.xlsx" -> wiedza) ->
`run_pipeline_with_parameters.py` (enriched) -> `build_kanlux_baselinker_update.py`
(CSV importu, raporty brakow/audytu, skoroszyt brakow) ->
`build_kanlux_woocommerce_update.py` (plik Woo). Reczne uzupelnienia atrybutow
czyta z `archived_input_files/input_2026-08-05/kanlux_braki_wymaganych_uzupelnione.xlsx` (wygrywaja z automatem,
"n/d" pomijane). Decyzje sklepu o wartosciach atrybutow sa skodyfikowane w
`dictionaries/supplier_global_knowledge.yaml` -> `value_conventions`.

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

## Skille projektowe (kopia zywych skilli Codexa)

Skille Codexa uzytkownika sa lustrzanie skopiowane do `.claude/skills/`.
Zrodlo: `C:\Users\Handlowiec\.codex\skills\<nazwa>\SKILL.md`.
Przed praca nad odpowiednim tematem przeczytaj caly plik skilla.

- `.claude/skills/alkan-baselinker-pipeline/SKILL.md` - nazwy, kategorie,
  opisy, atrybuty i importy Baselinker/eKonektor z plikow dostawcow
  (Kobi/Kanlux/Ospel/Simon/Hager). Zawiera reguly zgodnosci wariantow
  Hager/Berker oraz `references/ekonektor_csv.md` (8 kolumn, cp1250 bez BOM)
  i `scripts/export_ekonektor_csv.py`.
- `.claude/skills/alkan-product-description/SKILL.md` - pojedyncze opisy HTML
  SEO z `description_codex_brief_*.xlsx`. Kazdy SKU pisany recznie, nigdy
  petla w Pythonie ani wspolny szablon akapitow.
- `.claude/skills/alkanlocal-site-workflow/SKILL.md` - projekt WordPress/
  WooCommerce `C:\laragon\www\AlkanLocal` (osobne repo). Zmiany wizualne
  weryfikowane Playwrightem.
- `.claude/skills/ver-upd/SKILL.md` - podsumowanie waznych zmian Git
  (release notes) po polsku, na haslo "ver-upd".

Prompty w `prompts/` sa czescia kontraktu jakosci opisow:
`prompts/codex_description_agent.md` (styl i jakosc),
`prompts/examples/good_description_arot.html` (wzorzec),
`dictionaries/seo_description_knowledge.yaml` (nadrzedne zrodlo regul SEO -
przy konflikcie z promptem wygrywa YAML).
`prompts/alkan_product_description_skill.md` to kopia skilla
`alkan-product-description` - trzymaj oba w zgodzie.

## Zapamietane decyzje uzytkownika

(Sekcja wspolna z `AGENTS.md` - obowiazuje tak samo Claude'a i Codexa.)

- Dopasowanie produktu po SKU w `find_product_row_by_sku` idzie od
  najdokladniejszego: pelny kod -> kod bez sufiksu dostawcy (`/KON`, `/KAN`) ->
  stare luzne dopasowanie po czlonie przed pierwszym ukosnikiem. Niejednoznacznosc
  na dowolnym poziomie **przerywa z bledem**, zamiast wybierac pierwszy wiersz.
  **Powod:** stare dopasowanie obcinalo kod na pierwszym ukosniku i dla Kontakt
  Simon mylilo warianty - `TW6.01/142/KON` i `TW6.01/X/142/KON` dawaly ten sam
  klucz `TW6.01`, wiec brief wariantu `/X/` dostawal dane wariantu bazowego
  razem z cudzym EAN-em. Walidator tego NIE wykrywa, bo opis jest spojny ze
  swoim (blednym) briefem. Przy nowym dostawcy warto porownac SKU z briefow
  z plikiem wejsciowym.
- Frazy do generatora zbiera `scripts/harvest_description_phrases.py` z
  poprawionych opisow (`--html-dir` albo `--input` z kolumna opisu). Wynik idzie
  do `dictionaries/seo_phrase_bank_by_family.yaml` dopiero przez `--merge-into`;
  bez tej flagi skrypt tylko raportuje do `reports/`.
- Bank rodzinowy jest czytany PRZED bankiem globalnym
  (`seo_phrase_bank_harvested.yaml`). Rozdzial jest celowy: bank globalny nie
  wie, do jakiego produktu trafi fraza, wiec `is_safe_flavor` wycina z niego
  slowa `gniazd`, `laczni`, `puszk`, `czujnik`, `podtynkow`, `zacisk`. Bank
  rodzinowy zna rodzine, wiec wolno mu niesc slownictwo branzowe.
- Do banku trafia tylko PROZA z akapitow `<p>`, bez pierwszego. Naglowki,
  punkty list i FAQ sa pomijane: punkty `<li>` to pary cecha-korzysc prawdziwe
  dla jednego produktu, czyli warstwa argumentacyjna, ktorej bank z zasady nie
  skaluje. Odrzucane sa tez zdania z liczba, kodem produktu, marka, kolorem,
  zwrotem do klienta i zdania zalezne od kontekstu.
- Sekcja `by_domain` w banku powstaje przy zbieraniu, ale **nie jest czytana
  automatycznie**. Zdanie potrafi opisywac funkcje swojego produktu, nie
  nazywajac go ("Pojedynczy przedmiot wsuniety do jednego otworu nie otwiera
  toru pradowego" dotyczy przeslon gniazda), a na innej rodzinie bylby to
  falszywy opis. Frazy przenosi sie z `by_domain` do `by_family` recznie.
- Rodzine produktu w generatorze opisow ustala
  `dictionaries/product_family_taxonomy.yaml` (52 rodziny w domenach
  `oswietlenie`, `osprzet`, `automatyka`, `trasy_kablowe`, `aparatura`,
  `kable`, `narzedzia`). **Kazdy nowy typ produktu dostaje wlasna rodzine** -
  nie wolno rozszerzac `ogolny` ani zostawiac produktow w fallbacku.
  Fallback `nieznana` jest neutralny i raportowany; `ogolny` zostal usuniety z
  `LIGHT_FAMILIES`, bo wczesniej kazdy nierozpoznany produkt byl opisywany jak
  oprawa oswietleniowa.
- Przy wejsciu nowego dostawcy albo typu produktu uruchom
  `py scripts/report_family_classification.py --input <plik> --report <raport>`
  i **zglos uzytkownikowi**, ile sztuk wpadlo w `nieznana` i jakie to typy.
- Dopasowanie rodziny jest dwuprzebiegowe: najpierw nazwa i pola typu
  (`primary_fields`), dopiero potem sciezka kategorii sklepu
  (`context_fields`). Powod: kategoria Woo "Gniazdka i Laczniki > ..." brzmi
  tak samo dla calej serii, wiec w jednym worku wrzucalaby ramki i klawisze do
  rodziny gniazd. Czesci nakladane (`zaslepka`, `pokrywa_gniazda`, `klawisz`,
  `plytka_czolowa`, `ramka`) rozpoznaje `starts_with`, czyli rzeczownik glowny:
  "Zaslepka ramki" to zaslepka, a nie ramka.
- Opisy produktowe (ścieżka `codex_brief` -> opis HTML -> review) konczy blok
  FAQ: po akapicie podsumowujacym idzie `<hr>`, naglowek
  `<h3>Najczęściej zadawane pytania</h3>`, dokladnie 3 pary pytanie/odpowiedz
  (kazda odpowiedz 200-300 znakow) i schema `FAQPage` w JSON-LD powtarzajaca
  widoczna tresc. Pytania numerowane sa tylko w czesci widocznej. Regula stoi w
  `dictionaries/seo_description_knowledge.yaml` (`product_description.faq`),
  jest egzekwowana przez `validate_faq_section` w
  `src/generate_product_descriptions.py` i pokryta testem
  `tests/test_generate_product_descriptions.py`.
- Kolumna `Fraza kluczowa` musi byc w pliku wejsciowym briefu. Bez niej
  `find_seo_keyword` zwraca pusty string, a walidator **pomija wszystkie testy
  frazy kluczowej** - opis przechodzi walidacje mimo zlych naglowkow. Fraza ma
  byc doslownym prefiksem nazwy produktu (np. `Gniazdo wtyczkowe`), a nie
  mnogim `Typ produktu + Seria`, ktore generuje `build_seo_keyword`.
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
robocza. Na poczatku kazdego nowego czatu w tym projekcie uruchom:

    py scripts\check_agent_sync.py

(w sesji Cowork/Linux: `python3 scripts/check_agent_sync.py`; jesli katalog
`C:\Users\Handlowiec\.codex\skills` nie jest podpiety do sesji, popros o dostep
przez `device_request_folder_access` - zgoda wygasa razem z sesja.)

Skrypt porownuje w OBIE strony: skille Codexa z kopia w `.claude/skills/`,
naglowki sekcji `AGENTS.md` i `CLAUDE.md`, oraz
`prompts/alkan_product_description_skill.md` ze skillem
`alkan-product-description`.

Co zrobic z wynikiem:

- `TYLKO U CODEXA` - u Codexa przybyl albo zmienil sie skill. Przeczytaj
  zrodlo, zaktualizuj kopie w `.claude/skills/`, przenies istotne reguly tutaj
  i powiadom uzytkownika.
- `TYLKO U CLAUDE` - kopia ma cos, czego Codex nie ma. Skopiuj plik do
  `C:\Users\Handlowiec\.codex\skills\<nazwa>\`.
- `ROZNI SIE` - porownaj obie wersje i scal swiadomie, nie nadpisuj w ciemno.
- `TYLKO W AGENTS.md` - dopisz brakujaca regule do `CLAUDE.md` (i odwrotnie).

Kiedy sam tworzysz albo zmieniasz skill lub regule: zapisz go w
`.claude/skills/<nazwa>/SKILL.md` (nawet jesli powstal jako skill konta Claude -
inaczej Codex go nie zobaczy) i dopisz reguly ogolne do OBU plikow -
`CLAUDE.md` i `AGENTS.md`. Codex ma bliznaczy zapis tej procedury w `AGENTS.md`.
