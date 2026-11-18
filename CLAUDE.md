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

Wzory tytulow i atrybutow sa wspolne dla wszystkich osob pracujacych z repo i
leza w `configs/title_patterns.yaml` (oswietlenie: `configs/title_anatomy.yaml`).
Przed nadaniem nazw sprawdz, czy rodzina ma tam wzor, i trzymaj sie go. Gdy
ustalisz nowy wzor albo zmienisz istniejacy (nowy producent, nowa rodzina,
poprawka po uwagach uzytkownika), dopisz go do `configs/title_patterns.yaml`
z prawdziwymi przykladami i uruchom `py -m pytest tests/test_title_patterns.py`.
Katalogi `output/` i `outputs/` (takze `outputs/<id-sesji>` tworzone przez agenta)
nie sa w gicie - skrypt czy tabela zostawione tylko tam przepadaja dla innych,
wiec wzor zapisany wylacznie w jednorazowym skrypcie nie jest wzorem projektu.

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
  Slownik `dictionaries/woocommerce_catalog_knowledge.yaml` generuje
  `analyze_woocommerce_export.py` z pliku `woo_api_cli.py pull --all` (od 2026-11-05;
  dawniej CSV `wszystko.csv`): sciezki kategorii rozwija do rodzicow, nazwy atrybutow
  bierze z `cache/woo_api/knowledge.json` (najpierw `knowledge` z tego samego sklepu),
  przecinek w wartosci (`2,5`, dane producenta) nie dzieli juz terminu.
  Nazwy atrybutow LOKALNYCH produktow (nie `pa_*`, np. `Kompatybilny z`) bierze z kolumny
  `Product Attributes` pliku z pull (naglowki maja Wielkie Litery), EAN z kolumny `EAN`.
  **Cala wiedza o sklepie jest z PRODUKCJI alkan.pl** (decyzja uzytkownika 2026-11-18;
  wczesniej zrzuty z AlkanLocal, a slownik katalogu z CSV z maja 2026). Odswieza ja jedno
  polecenie `py scripts/refresh_shop_knowledge.py --pull` (najlepiej poza godzinami ruchu):
  `knowledge --profile prod` -> `dictionaries/woo_*_snapshot.tsv` + `cache/woo_api/knowledge.json`
  (kopia w `cache/woo_api/prod/`), lista `woo_attributes` w `kontakt_simon_shop_structure.yaml`,
  raport smieci, pelny `pull --all` -> `woocommerce_catalog_knowledge.yaml` i
  `learned_product_taxonomy.yaml`; potem testy i `py tests/generate_snapshots.py`.
  Produkcja ma smieciowe terminy po importach (`Antenowe, Komputerowe`, `Kaszmirowy`, opisy
  wpisane jako `Materiał`). `shop_attribute_values.load_vocabulary` ich nie uznaje za terminy
  sklepu: klucze `value_map` (takze `IP20` z 328 produktami -> `IP 20`), terminy z licznikiem 0
  sklejone przecinkiem z innych terminow (nie liczby `10,5`) i `shop_junk_terms` w
  `dictionaries/woo_attribute_value_rules.yaml`. Lista do sprzatniecia w Woo (smieci + terminy
  bez produktow): `py scripts/report_shop_junk_terms.py` -> `reports/woo_api/`.
- `dictionaries/supplier_global_knowledge.yaml` - globalna wiedza o dostawcach,
  obecnie przede wszystkim Kanlux: suffixy, zasady kategorii, tytulow,
  akcesoriow, atrybutow i bezpiecznych opisow. Przy zadaniach Kanlux czytaj ten
  plik razem z `docs/kanlux_workflow.md`.
- `analyze_keywords.py` - Google Ads Keyword Planner (zaleznosc opcjonalna,
  importowana leniwie; dziala tez offline z `--offline-keywords`).
- Pozostale `enrich_*`, `recommend_*`, `export_*`, `build_kobi_baselinker.py`,
  `learn_product_taxonomy.py` - narzedzia pomocnicze, kazde z wlasnym argparse.

## Polaczenie ze sklepem Woo (AlkanInator Bridge)

Sklep ma wtyczke `alkan-inator-bridge` (repo AlkanLocal,
`wp-content/plugins/alkan-inator-bridge`) z REST API `alkan-inator/v1`. Klient:
`src/woo_api/` + `py src/woo_api_cli.py`. Polaczenie: `src/woo_api/config.json`
(ignorowany przez git; `url`, `token`, `verify_ssl` - dla lokalnego Laragona
`false`) albo zmienne `ALKAN_WOO_URL` / `ALKAN_WOO_TOKEN`. Kilka sklepow:
`{"default": "local", "profiles": {"local": {...}, "prod": {..., "production": true}}}`
i `--profile prod` (plaski format = profil `local`). Kazda osoba ma wlasny token
(WooCommerce > AlkanInator API, pokazywany raz), dziennik zapisow pokazuje, kto
zapisal. Klient sprawdza wersje API wtyczki (naglowek
`X-Alkan-Inator-Api-Version`, obslugiwana 2.x; 2.1 dodaje `/variant-groups`,
filtr `tag` i `/shipping-preview` - AlkanLocal od 2026-11-02, produkcja do wgrania). Dziala na AlkanLocal
(`https://alkanlocal.test:8443`, profil `local`) i od 2026-11-02 na produkcji
(profil `prod` z `"production": true`; domyslny zostaje `local`, produkcje
wybiera sie jawnie `--profile prod` przed nazwa polecenia). WP All Import
(`dictionaries/woo_all_export_import_workflow.yaml`) nadal dziala na produkcji.
`pull` z produkcji jest znacznie wolniejszy niz z Laragona (kazda strona to
osobne uruchomienie WordPressa ze wszystkimi wtyczkami sklepu): pobieraj zakres
(`--category`, `--modified-after`, `--sku-file`) zamiast `--all`, a pelny
katalog z `--per-page 500` poza godzinami ruchu. Filtr `--sku` dopasowuje SKU
dokladnie, z sufiksem dostawcy (`37246/KAN`); warianty nie sa zwracane.

- `knowledge` - wiedza o sklepie. Nadpisuje
  `dictionaries/woo_attribute_terms_snapshot.tsv` i
  `dictionaries/woo_product_categories_snapshot.tsv` (ten sam uklad co
  `scripts/dump_woo_*.py`, nazwy z odkodowanymi encjami, np. `F&F`), tworzy
  `dictionaries/woo_shop_terms_snapshot.tsv` (klasy wysylkowe, "Wysylka w ciagu",
  tagi), od wtyczki 2.1 takze `dictionaries/woo_variant_groups_snapshot.tsv`
  (grupy WPC Linked Variation: tagi i osie, uklad jak
  `alkanlocal_wpclv_groups.tsv`) i pelny JSON `cache/woo_api/knowledge.json`
  (kolumny, ktore sa zapisywalne, formaty, jednostki). Odswiezaj przed praca na wartosciach
  atrybutow zamiast zrzutow z MySQL.
- `pull --output plik.xlsx` - produkty w ukladzie WP All Export (arkusz
  `Worksheet`, te same naglowki: `Kategorie produktów`, `Product Tags`,
  `Up-Sells`, `Atrybut Produktu: <Etykieta>` z wielkimi literami jak w WP All
  Export). Filtry: `--category` (slug, ID albo sciezka `A>B`, z podkategoriami),
  `--tag` (slug albo nazwa, np. tag grupy wariantow; tylko wtyczka 2.1 - CLI
  odmawia przy starszej, bo ta pominelaby filtr i oddala caly katalog), `--sku`, `--sku-file`, `--ids`, `--status`, `--search`, `--modified-after`,
  `--all`. Dodatkowe kolumny tylko do odczytu: ceny, stan, EAN,
  `Post Modified Date`, `Permalink`, `Product Attributes`. Od 2026-11-02 plik
  z `pull` ma ukryty arkusz `_alkan` z odciskiem kazdego wiersza z chwili
  pobrania (`src/woo_api/fingerprint.py`) - nie usuwac go i nie przepisywac
  danych do nowego skoroszytu, bo przepada przyspieszenie.
- `push --input plik.xlsx` - ZAWSZE najpierw dry-run calego pliku i raport
  `reports/woo_api/*_dry-run.xlsx` (arkusze Zmiany / Problemy / Nowe terminy).
  `--apply` zapisuje tylko produkty ze zmianami; bledy w dry-run wstrzymuja zapis
  (chyba ze `--skip-errors`). `--columns "A,B"` ogranicza zapis do wskazanych
  kolumn - uzywaj go, gdy zadanie dotyczy kilku kolumn (odpowiednik zasady
  "modyfikuj tylko jawnie wskazane kolumny"). Nowe tagi i wartosci atrybutow
  powstaja tylko z `--allow-new-terms`. Nowe zdjecia z `Image URL` dry-run
  sprawdza (JPG 1000x1000, inaczej blad). Wyczyszczenie `Content`,
  `Short Description` albo `Image URL` wymaga `--allow-destructive`. Sklep
  zapisuje produkt tylko, gdy jego zmiany sa identyczne jak w dry-run
  (`_plan_hash`); inaczej blad "zmiany roznia sie od podgladu" - uruchom push
  jeszcze raz.
- `changesets` i `revert <zapis> [--apply]` - kazdy zapis ma identyfikator
  (np. `ai-20261021-153259-plik`) i dziennik starych wartosci w sklepie.
  Cofniecie pomija pola zmienione w sklepie po tym zapisie. Cofnac mozna tez
  przyciskiem w WooCommerce > AlkanInator API.
- Aplikacja "Wgraj zmiany do sklepu" (od 2026-11-02, dla osob bez terminala):
  dwuklik `Wgraj zmiany do sklepu.bat` (albo `py src/woo_preview_app.py`) otwiera
  strone w przegladarce. Nowy komputer: Python z py launcherem + `Instalacja.bat`
  (requests, openpyxl, Pillow; `.bat` aplikacji sam go wola, gdy bibliotek brak -
  `pyw` padal bez komunikatu) + wlasny `config.json`; `Diagnostyka - Wgraj
  zmiany.bat` pokazuje bledy startu. Wybor sklepu (profile z `config.json`, bez tokenow) i
  pliku -> dry-run -> karty produktu przed/po w CSS motywu `alkan` z AlkanLocal
  (`--theme-dir`, bez motywu proste style zastepcze), zmienione pola w ramce,
  tabela zmian, odznaczanie kolumn (= `--columns`) -> "Zatwierdz" -> zapis ->
  identyfikator zapisu i "Cofnij". Stan "przed" to wiersz pliku z polami
  cofnietymi do "bylo" z dry-runu - sklep nie dostaje dodatkowych zapytan. Do
  zapisu idzie kopia z samymi zmienionymi produktami (`cache/woo_preview/`,
  oryginal nietkniety); gdy dry-run przed zapisem rozni sie od podgladu (takze
  stara wartoscia - ktos zmienil produkt w miedzyczasie), zapis staje. Plik z
  innego sklepu (host w kolumnie `Permalink` inny niz wybrany sklep, np. eksport
  z alkan.pl na AlkanLocal) zatrzymuje podglad przed dry-runem - inne ID i adresy
  zdjec dawaly zmiane `Image URL` w calym katalogu i pobieranie tysiecy zdjec. Zapis i
  cofniecie na PRODUKCJI potwierdza czlowiek w okienku Windows (MessageBox) -
  odpowiednik potwierdzenia w terminalu CLI. Serwer slucha tylko na 127.0.0.1,
  API wymaga jednorazowego klucza z adresu, po 15 min bez otwartej strony konczy
  sie sam. alkan.pl blokuje zdjecia z obcym `Referer` (403), dlatego karta ma
  `no-referrer`. Karty pokazuja tez przelacznik wariantow, Produkty kompatybilne,
  Up-Sells i Cross-Sells (dodane/usuniete zaznaczone). Grupy WPC Linked Variation
  (grupa i os = atrybut laczenia) podglad bierze ze sklepu (`/variant-groups`,
  wtyczka 2.1), a rodzenstwo wariantow dociaga po tagu; przy wtyczce 2.0 zostaje
  zrzut `dictionaries/woo_variant_groups_snapshot.tsv` i rodzenstwo tylko z pliku.
  Nota pod przelacznikiem podaje grupe, os, liczbe rodzenstwa i powtorzenia osi
  (`DO_SPRAWDZENIA`), jak raporty wariantow. Produkty spoza pliku: jeden odczyt na
  200 SKU powiazan i na 50 tagow wariantow, tylko dla zmienionych produktow.
  Wysylka: metode dostawy wybiera KLASA WYSYLKOWA (Flexible Shipping, takze waga
  i wymiary), nie tagi. Po otwarciu karty sklep liczy metody i ceny (1 szt.)
  teraz i po zmianie (`/shipping-preview`: ten sam kalkulator co karta produktu,
  na kopii produktu w pamieci, bez zapisu); karta pokazuje metody znikajace i
  nowe oraz linie dostepnosci z `Stock`, `Backorders` i `Wysylka w ciagu` (bez
  magazynu zewnetrznego). Przy wtyczce 2.0 - sama klasa i waga z nota. Kod: `src/woo_preview/` (korzysta z
  `src/woo_api/`, nie zmienia go), test `tests/test_woo_preview.py`.

Zapisywalne kolumny: `Title`, `Content`, `Short Description`, `Status`,
`Kategorie produktów`, `Product Tags`, `Shipping Class`, `Wysyłka w ciągu`,
`Backorders` (`no`/`notify`/`yes`), `Weight`, `Length`, `Width`, `Height`,
`Up-Sells`, `Cross-Sells` (SKU po `|`), `_alkan_compatible_product_skus`
(Produkty kompatybilne; przyjmuje tez `_alkan_compatible_product_ids`),
`Image URL` (pierwszy URL = zdjecie glowne, reszta = galeria; obce URL-e sa
pobierane do mediow) i `Atrybut Produktu: *`. Cen i stanow NIE zapisuje -
naleza do Baselinkera.

Zasady zapisu:
- Pobrany plik odeslany bez zmian daje 0 roznic (sprawdzone na calym katalogu
  AlkanLocal) - kazda pozycja w raporcie Zmiany to realna zmiana.
- Plik z odciskami (`pull` od 2026-11-02) wysyla do dry-runu i zapisu TYLKO
  wiersze zmienione w pliku od pobrania (inny odcisk, wiersz dopisany albo
  wartosc w kolumnie dodanej po pobraniu; pusta komorka nowej kolumny nie liczy
  sie jako zmiana). Niezmienione wiersze nie cofaja zmian zrobionych w sklepie
  po pobraniu. `--all-rows` (w aplikacji "Sprawdz wszystkie wiersze") wysyla
  caly plik jak dawniej; pliki bez arkusza `_alkan` (WP All Export, starsze
  pobrania) tez ida w calosci. Partie po 200 wierszy (limit wtyczki) - kazde
  zadanie to osobny start WordPressa (~2 s), wiec mniej partii = szybciej.
- Pusta komorka czysci pole (pusty atrybut usuwa go z produktu). Kolumny, ktorej
  nie ma w pliku, zapis nie rusza.
- Kategorie: w pliku sa tylko najglebsze sciezki (jak w WP All Export); `A>B`
  przypisuje B razem z kategoriami nadrzednymi. Nowych kategorii nie tworzy.
- Atrybut nowy dla produktu dostaje deklaracje automatycznie (visible=1,
  variation=0) - kolumna `Product Attributes` jest przy zapisie ignorowana.
  Nowych atrybutow globalnych nie tworzy.
- Wartosci dopasowuje bez wielkosci liter, ale dokladna nazwa i terminy juz
  przypisane produktowi wygrywaja (sklep ma duplikaty typu
  "Tworzywo Sztuczne"/"Tworzywo sztuczne").
- Opisy zapisuja sie jako administrator, wiec `<script>` JSON-LD (FAQ) zostaje -
  inaczej niz przy imporcie przez Baselinker/WC REST.
- Gdy produkt zmienil sie w sklepie po pobraniu pliku (`Post Modified Date`),
  dry-run daje ostrzezenie - plik moze cofac cudze zmiany, pobierz go od nowa.
- Nie tworzy nowych produktow i nie obsluguje wariantow `product_variation`.

### Strefa chroniona (bezpieczniki zapisu)

Z AlkanInatora korzysta kilka osob, a docelowo zapis idzie do sklepu
produkcyjnego. Bezpieczniki sa we wtyczce (wylacznik zapisu, limit produktow na
jeden zapis, zgoda na czyszczenie pol, zgodnosc z dry-run, tokeny osob) i w CLI
(potwierdzenie zapisu na produkcje w terminalu). Zasady dla agentow:
- NIE zmieniaj `src/woo_api/`, `src/woo_api_cli.py` ani testow w
  `tests/test_woo_api.py` bez wyraznej prosby uzytkownika w tej rozmowie.
  Testy pilnuja formatu `config.json`, plikow knowledge i XLSX - nie
  dopasowuj ich do zmienionego kodu, tylko zapytaj.
- NIE wolaj API sklepu z pominieciem CLI (wlasne `requests`/`curl`/skrypty z
  tokenem) i nie czytaj ani nie wypisuj tokenu z `config.json`.
- NIE klikaj "Zatwierdz"/"Cofnij" w aplikacji "Wgraj zmiany do sklepu" i nie
  wolaj jej `/api/apply` ani `/api/revert` - zapis z aplikacji zatwierdza
  uzytkownik. Agent moze przygotowac plik i powiedziec, jak go wgrac.
- NIE uruchamiaj `push --apply` ani `revert --apply` na profilu produkcyjnym.
  CLI i tak odmowi bez terminala - przygotuj plik i dry-run, a polecenie z
  `--apply` podaj uzytkownikowi do uruchomienia samodzielnie. Na AlkanLocal
  `--apply` tylko na prosbe uzytkownika.
- NIE obchodz bezpiecznikow: `--allow-destructive`, `--skip-errors`,
  `--skip-image-check` i podnoszenie limitu tylko na wyrazne zyczenie.
  Blad "Zapis wstrzymany" albo "limit" zglos uzytkownikowi zamiast szukac
  obejscia.

## Dokumentacja dla zespolu (Notion)

`docs/changelog.md` (zmiany od najnowszych, jezykiem osob korzystajacych) i
`docs/instrukcja.md` (prosta instrukcja funkcji dla osob bez terminala) sa
publikowane na stronach Notion przez `py -I src/notion_api/publikuj-w-notion.py`
(mapa plik -> ID strony w `DOKUMENTY`; pusty ID = dokument pomijany). Publikowana
jest tresc z commita, nie robocza; po kazdym commicie robi to hook
`.githooks/post-commit` (wlaczenie: `git config core.hooksPath .githooks`).
`--podglad` pokazuje tytul i tresc bez wysylania, `--wszystkie` publikuje oba
dokumenty niezaleznie od commita. Token integracji Notion: `NOTION_TOKEN` albo
`src/notion_api/config.json` (`{"token": "..."}`, ignorowany przez git) - nie
czytaj go ani nie wypisuj. Strony w Notion sa nadpisywane w calosci, wiec poprawki
robi sie tylko w repo.
- Po zmianie widocznej dla uzytkownika (nowe narzedzie, nowy dostawca, nowa
  runda, zmiana zasady) dopisz wpis na gorze `docs/changelog.md` pod dzisiejsza
  data; gdy zmienia sie sposob korzystania - popraw tez `docs/instrukcja.md`.
  Oba pliki pisze sie po polsku z polskimi znakami, bez sciezek do skryptow.
- Test: `tests/test_notion_publish.py`.

## Testy

Testy uzywaja `unittest` (biblioteka standardowa, bez dodatkowych zaleznosci):

    py -m unittest discover -s tests -p "test_*.py"

Najwazniejsze sa **testy charakteryzujace** (golden-file) dla ekstrakcji i
generowania tytulow - utrwalaja aktualne zachowanie na realnych danych Kanlux,
zeby kazda niezamierzona regresja w regexach/regulach byla od razu widoczna.

Po **SWIADOMEJ** zmianie regul ekstrakcji, configow tytulow lub slownikow
nalezy zregenerowac snapshoty (wymaga `archived_input_files/Alkan_Kanlux_pelne_rodziny.xlsx`):

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
  `archived_input_files/input_2026-11-10/BerkerHager/BerkerHagerKategorie.xlsx`; jeden produkt moze nalezec do wielu
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
- Atrybuty filtrowe osprzetu (Kontakt Simon, a ten sam model maja Ospel i
  nowe partie Hager/Berker) ustala sklep, nie dostawca:
  `dictionaries/kontakt_simon_shop_structure.yaml` (odczytany z AlkanLocal).
  `Typ produktu` to jeden z 7 terminow (Gniazda, Laczniki, Przyciski, Ramki,
  Sterowniki i automatyka, Inne, Puszki podlogowe), `Podtyp produktu` jest
  przypiety do typu (np. Gniazda -> Wtyczkowe/Antenowe/USB, Laczniki ->
  Pojedyncze/Schodowe/Krzyzowe). **Ramki nie maja podtypu** - krotnosc idzie
  w atrybucie `Krotnosc` (`1x`..`5x`), material w `Material`, bez sposobu
  montazu. Klawisze w sklepie nie maja typu ani podtypu. Wartosci dostawcy
  ("Nazwa produktu/Rodzaj", ETIM) i wlasne typy ("Gniazdka", "Akcesoria",
  "Ladowarki USB", "Ramka 1- krotna") sa bledem - tak powstala partia
  Simon 54/55 kaszmir 98 z 2026-09-15. Naprawia to
  `src/fix_simon_attributes_to_shop_model.py` (dopasowanie po symbolu
  bazowym do rodzenstwa w innych kolorach, potem reguly po nazwie), test
  `tests/test_fix_simon_attributes_to_shop_model.py`. Wysylka poprawki do
  Baselinkera wymaga `--replace-features`, bo zwykly sync scala cechy i nie
  usuwa zbednych kluczy.
- **Zadnych nowych atrybutow sklepu bez zatwierdzenia.** Do Baselinkera/Woo
  trafiaja tylko cechy z listy `woo_attributes` w
  `dictionaries/kontakt_simon_shop_structure.yaml` (150 globalnych atrybutow
  Woo z produkcji 2026-11-18, lista z `knowledge.json` przez
  `scripts/refresh_shop_knowledge.py`) plus `approved_extra_attributes`. Cechy z
  katalogu dostawcy spoza tej listy (ETIM, opakowania, RAL, "Wysokosc [mm]"
  itp.) sa usuwane z pliku i laduja w raporcie
  (`atrybuty_do_zatwierdzenia` w JSON, arkusz "Atrybuty do zatwierdzenia" w
  skoroszycie kontrolnym) z podpowiedzia istniejacego atrybutu Woo. Nowy
  atrybut powstaje dopiero po decyzji uzytkownika - dopisuje go do
  `approved_extra_attributes` albo zaklada w Woo. **Powod:** integracja
  Baselinker -> Woo tworzy z kazdej cechy globalny atrybut, wiec kazda nowa
  nazwa smieci w filtrach i liscie atrybutow sklepu.
- **Zadnych nowych WARTOSCI atrybutow bez zatwierdzenia.** Kazda wartosc
  atrybutu globalnego spoza terminow sklepu tworzy w Woo nowy termin (przez
  Baselinkera i przez WP All Import). Slownik terminow:
  `dictionaries/woo_attribute_terms_snapshot.tsv` (zrzut z produkcji przez
  `woo_api_cli.py --profile prod knowledge`, bez smieci sklepu - `shop_junk_terms`; zawiera tez
  powiazanie Podtyp -> Typ z filtra). Reguly: `dictionaries/woo_attribute_value_rules.yaml`
  (`value_map` wartosci dostawcy -> termin sklepu, `null` = usun;
  `approved_new_values` - tylko po decyzji uzytkownika; `identifier_attributes`
  jak Symbol nie sa sprawdzane; atrybuty techniczne z jednostka zostaja z
  WARNING "utworzy nowy termin"). Wspolny modul `src/shop_attribute_values.py`,
  narzedzie `src/audit_shop_attribute_values.py --input <XLSX Woo albo CSV
  Baselinkera> [--shop-export <eksport marki>] [--output <kopia>]` - bez
  `--output` tylko raport (arkusze "Nowe wartosci", "Kontrola wartosci",
  "Terminy do usuniecia w Woo"). Kontrole maja juz wbudowana
  `fix_simon_attributes_to_shop_model.py` (poprawia), `build_simon_subseries_woo_update.py`
  i `fill_cable_1m_from_siblings.py` (raportuja). **Po kazdym innym narzedziu
  produkujacym atrybuty uruchom audit na pliku wynikowym przed importem.**
  Kilka terminow jednego atrybutu rozdziela `|` (`Antenowe|Komputerowe`);
  przecinek daje jeden nowy termin bez powiazania z typem, wiec produkt
  znika spod typu w filtrze. **Powod:** partia Simon kaszmir (2026-09-15)
  utworzyla w sklepie 16 smieciowych terminow: `Antenowe, Komputerowe`,
  `Wtyczkowe, USB`, `Kaszmirowy` obok `Kaszmir`, material ETIM
  (`tworzywo sztuczne, PC, bezhalogenowe`) i caly slownik `Typ zaciskow` z
  ETIM (sklep zna tylko Gwintowe / Bezgwintowe / Bezgwintowe - IDC / Do
  lutowania). Poprawka: `output/Simon-Kaszmir_WARTOSCI_POPRAWIONE_2026-10-12.xlsx`
  (WP All Import), kontrola w `reports/Simon-Kaszmir_kontrola_wartosci_2026-10-12.xlsx`.
  Test: `tests/test_shop_attribute_values.py`.

- Nowy kolor istniejacego produktu Simon dostaje tag wariantowy swojego
  rodzenstwa (ten sam symbol bazowy, inne kolory) - robi to
  `src/fill_variant_tags_from_siblings.py --color-codes 42,142` na pelnym
  eksporcie WP All Export; zmienia tylko `Product Tags` nowych produktow.
  Sklep laczy warianty wtyczka WPC Linked Variation po tagu, wiec tag musi byc
  DOKLADNIE taki jak u rodzenstwa. Tagi eventowe (`variant_tags.ignored_tags`,
  np. "ElektroTargi Alkan 2026") nie sa kopiowane; przy rozbieznosci wygrywa
  wiekszosc, remis rozstrzyga `variant_tags.overrides`. Produkty, ktorych
  rodzenstwo nie ma tagow (klawisze, pokrywy) albo ktore nie maja rodzenstwa,
  zostaja bez tagu i sa raportowane - nie wymyslamy nowych grup wariantow.
  Niespojnosci istniejacych produktow trafiaja do arkusza "Anomalie
  rodzenstwa", nie sa poprawiane bez zgody. Test:
  `tests/test_fill_variant_tags_from_siblings.py`.
- **Kazdy raport wariantow podaje, po jakim atrybucie lacza sie produkty**
  (decyzja uzytkownika 2026-10-29). Sam tag i lista SKU nie wystarcza: raport
  ma kolumny z grupa WPC Linked Variation, do ktorej dopisuje sie tag, z osia
  grupy (atrybut, np. `Kolor`, `Kolor producenta`, `Podswietlenie`) i z wartoscia
  tego atrybutu per SKU, plus kontrole, czy wartosci w grupie sie nie powtarzaja
  (powtorzenie = brak osi, `DO_SPRAWDZENIA`). Osie istniejacych grup sa w
  `dictionaries/woo_variant_groups_snapshot.tsv` (z produkcji; dawny zrzut AlkanLocal
  `alkanlocal_wpclv_groups.tsv` nie jest juz czytany). Robia to
  `src/build_simon_woo_relations.py` (sekcja `variant_group` w
  `dictionaries/simon_woo_relations.yaml`) i `src/fill_variant_tags_from_siblings.py`
  (kolumna "Grupa wariantow w sklepie [osie]"); nowe narzedzie od wariantow musi miec to samo.
- **Opisy ze schema JSON-LD nie przechodza przez Baselinker -> Woo.** Baselinker
  zapisuje produkt przez REST API WooCommerce, a kontroler produktow zawsze
  robi `wp_filter_post_kses` na opisie, wiec `<script type="application/ld+json">`
  znika i schema FAQ wyswietla sie w sklepie jako goly tekst JSON (w
  Baselinkerze opis jest poprawny - blad widac dopiero w Woo). Opisy z blokiem
  FAQ wgrywa sie do Woo plikiem WP All Import (dziala jako administrator i
  zachowuje `<script>`). Naprawa juz zepsutych:
  `src/restore_woo_description_scripts.py` na pelnym eksporcie - podmienia opis
  tylko gdy tresc w Woo jest identyczna z nasza poza znacznikami `<script>`
  (recznych poprawek w sklepie nie nadpisuje). Po kazdym imporcie opisow przez
  Baselinker sprawdz w eksporcie Woo, czy `Content` zawiera `<script`. Trwale
  rozwiazanie to mu-plugin po stronie sklepu (przywracanie zwalidowanego
  FAQPage w `woocommerce_rest_pre_insert_product_object`) - do decyzji
  uzytkownika.
- **"Produkty kompatybilne" to osobne pole, nie cross-sell.** Sklep ma wtyczke
  `alkan-compatible-products`: pole meta `_alkan_compatible_product_ids`
  (tablica ID albo ID po przecinku - oba formaty sa czytane), sekcja na karcie
  produktu tuz przed Up-Sell. Cross-sell jest widoczny tylko w koszyku. Dla
  Hager/Berker pole buduje `src/build_hager_berker_compatible_products.py`
  wylacznie z relacji producenta: bloki "Wspolpracuje z" z katalogu PDF
  (`output/Hager-Berker-Katalog-PDF_BAZA_KOMPATYBILNOSCI.xlsx`) oraz relacje
  `mandatory` i `accessories` z BMEcat (arkusz "Relacje produktowe"). Plik
  Hagera NIE zna pojecia cross-sell - istniejace Cross-Sells/Up-Sells pochodza
  z wlasnych regul sklepu (`build_hager_berker_cross_upsell.py`) i zostaja bez
  zmian. Relacje dzialaja w obie strony (ramka wymienia klawisz -> na karcie
  klawisza jest ramka). Punktacja, limit 8 sztuk, roznorodnosc i lista slow
  obnizajacych ranking wersji specjalnych (NEMA, nadruk, KNX...) sa w
  `dictionaries/hager_berker_compatible_products.yaml`. Element wkladany i
  ramka musza sie zgadzac kolorem; gdy producent nie ma ramki w kolorze
  produktu (pomaranczowe gniazdo), relacja zostaje ze statusem `WARNING`.
  Produkty bez relacji producenta dostaja tylko PROPOZYCJE z reguly "ramka
  1-krotna tej samej serii i koloru" (`DO_SPRAWDZENIA`, arkusz "Propozycje bez
  relacji") - do importu trafiaja dopiero z `--include-series-fallback`.
  Pomijane sa stare duplikaty, warianty tego samego produktu (przelacza je
  wtyczka wariantow) i produkty juz obecne w Up-Sells tej samej karty. ID w
  pliku importu pochodza z eksportu podanego w `--input`, wiec eksport musi byc
  z tego sklepu, do ktorego idzie import. Uzytkownik woli dostac z powrotem
  SUROWY eksport sklepu z wypelniona kolumna, a nie osobny plik: `--fill
  <eksport.xlsx>` zapisuje kopie eksportu (domyslnie
  `output/Kompatybilne-Berker-Hager_UZUPELNIONE_<data>.xlsx`), zmienia tylko
  kolumne `_alkan_compatible_product_ids` (ID po przecinku), bierze ID i
  aktualne Up-Sells z tego eksportu, a reczne wpisy ze sklepu zostawia na
  poczatku listy i tylko je dopelnia. Produkty z eksportu, ktorych nie ma w
  pliku atrybutow, dostaja rekord z samego tytulu (stare kody Berker sa wtedy
  rozpoznawane jako duplikaty i zostaja puste). Test:
  `tests/test_build_hager_berker_compatible_products.py`.

- **API Kontakt-Simon (e-katalog).** Klient: `src/simon_api.py`
  (`https://ekatalog.kontakt-simon.com.pl/api/v2/`, naglowek `X-ApiKey`,
  artykuly po symbolu w paczkach po 25, drzewo kategorii
  `catalog/category/tree.json` z `articleIdList`). Klucz lezy lokalnie w
  `Simon API/Program.php` (`define("API_KEY", ...)`; katalog jest w
  `.gitignore`) albo w zmiennej `SIMON_API_KEY` - nie kopiowac go do innych
  plikow ani raportow. `src/baselinker_api/config.json` ma tylko token
  Baselinkera. Odpowiedzi sa cache'owane w `cache/simon_api/` (`--offline`
  pracuje bez sieci, `--refresh` odswieza).
- **Podserie Kontakt Simon (Line, Duo, Nature, GO, Premium, Flash,
  K45/SF/CIMA).** Model z 2026-09-28: atrybut `Seria` = pelna nazwa
  (`Simon 55 Line`), zatwierdzony nowy atrybut `Podseria` = `Line`
  (`approved_extra_attributes` w `kontakt_simon_shop_structure.yaml`), a w
  tytule `<Seria> <Podseria>` tuz przed kodem producenta (`... Simon 55 Line
  TR1/111`; samodzielne tokeny `LINE`/`DUO`/`Nature` z tytulu sa usuwane).
  Podserie dostaja TYLKO produkty, ktore producent przypisuje do podserii
  (ramki, sterowniki GO, puszki natynkowe "do ramek Premium", akcesoria z
  kategorii ramek Line); mechanizmy (laczniki, gniazda, klawisze) pasuja do
  wszystkich ramek serii i zostaja bez podserii. Zrodlo: parametr `Seria` w
  API (np. `Simon 55 Nature`, `Simon 55 GO`, `Simon 54 Premium`), a gdy API
  podaje gola serie - nazwa kategorii e-katalogu (`Ramki Simon 55 Line`,
  `Ramki Simon 55 Duo`). Reguly: `dictionaries/simon_subseries_rules.yaml`,
  skrypt `src/build_simon_subseries_woo_update.py` (kopia eksportu Woo z
  podmienionymi `Title`/`Seria` i nowa kolumna `Atrybut Produktu: Podseria`
  za `Seria`, skoroszyt kontrolny w `reports/`), test
  `tests/test_build_simon_subseries_woo_update.py`. Seria Woo rozna od serii
  producenta (np. 85 mechanizmow Simon 54 z `Simon 54 Premium`) jest
  poprawiana ze statusem WARNING (arkusz "Korekty serii");
  `--keep-woo-series` tylko raportuje. Produkty pasujace do kilku podserii
  naraz (puszki natynkowe "do ramek Line i Duo", uszczelki IP44) dostaja
  wszystkie podserie po `|` w atrybutach (`Podseria` = `Line|Duo`, `Seria` =
  `Simon 55 Line|Simon 55 Duo`), a tytul zostaje bez zmian (status WARNING).
  Dla Simon Connect podseria to system K45/SF/CIMA/KFR, ale `Seria` i tytul
  zostaja.
- **Nowa partia Kontakt Simon bez rodzenstwa w sklepie (od 2026-10-20, Simon 24).**
  `src/build_simon_baselinker_from_api.py --input <zestawienie z Ekonoma>` bierze
  dane z API e-katalogu, a nazwy ze slownika `dictionaries/simon_product_names.yaml`:
  `<rdzen> <kolor> <seria> <kod>` ("Ramka 1-krotna biała Simon 24 NR1/11"); bialy i
  kaszmirowy odmieniaja sie przez rodzaj, kolory "... mat" nie. Rdzen jest per
  symbol bazowy (`/..A` = klapka transparentna); "Push In" w nazwie tylko gdy API
  podaje montaz Push In - skrypt to sprawdza, tak jak kolor z kodu. Wynik jest
  surowy: potem `fix_simon_attributes_to_shop_model.py` (Typ/Podtyp, kontrola
  wartosci). `description` zostaje puste - `build_description_html` pisal o
  gniazdach jezykiem oprawy ("oprawa jest zabezpieczona"); opisy pisze sie per
  SKU. Uszczelki NU1-NU5/NUC1-NUC5/NU1W to Simon Basic (API i katalog), nie
  Simon 24 - ta seria ma wlasne NUU1/NUV/NUH. Rodziny opisow `uszczelka`,
  `adapter_osprzetu` dodane, `narzedzie` przeniesiona wyzej (NDF wpadalo w
  gniazdo komputerowe). Test: `tests/test_build_simon_baselinker_from_api.py`.
- **Zdjecia produktow ZAWSZE: 1000x1000 JPG na bialym tle** - w kazdym kanale:
  Baselinker, WP All Import i `woo_api_cli.py push` (przypomnienie uzytkownika
  2026-10-21, potwierdzone 2026-10-29). Nigdy surowe PNG z przezroczystoscia ani
  500x500 ze strony producenta - najpierw najwiekszy oryginal (np. API e-katalogu
  Simona), potem obrobka `src/baselinker_api/images.py` (`_encoded_image`).
  `push` NIE obrabia zdjec (obce URL-e wtyczka pobiera do mediow tak, jak sa), ale
  od 2026-10-29 je SPRAWDZA: dry-run pobiera kazde zdjecie nowe dla produktu i daje
  blad, gdy to nie JPG 1000x1000 RGB (ostrzezenie, gdy rogi nie sa biale). Produkt ze
  zlym zdjeciem nie zapisuje sie takze z `--skip-errors`; `--skip-image-check` tylko na
  wyrazne zyczenie (`src/woo_api/push.py`, test `tests/test_woo_api.py`).
  Baselinker: `baselinker_api_cli.py sync` pobiera kazdy adres z `images_urls` i obrabia
  go jak plik lokalny (`src/baselinker_api/images.py`: przezroczystosc -> biale tlo, pad do
  kwadratu 1000x1000, JPEG), a nie wysyla surowego `url:` - wczesniej do Base trafialy
  500x500 PNG Simona z przezroczystym tlem. `--raw-image-urls` tylko na wyrazne zyczenie.
  Strona Kontakt-Simon dla produktow "WKROTCE W SPRZEDAZY" zwraca z HTTP 200 grafike
  "brak zdjecia" (przekreslony aparat 413x413) - `build_simon_baselinker_from_api.py`
  rozpoznaje ja po SHA256 (`PLACEHOLDER_IMAGE_SHA256`) i zostawia produkt bez zdjecia.
  Zdjecie w Base usuwa sie przez `addInventoryProduct` z `images: {"0": ""}`.
  Test: `tests/test_baselinker_api.py`.
- **Uzupelnianie serii Simon juz wgranej do sklepu (od 2026-10-28, Simon 24).** Wejscie:
  eksport WP All Export z produkcji (id, Title, Sku, Image..., Up-Sells, Cross-Sells,
  `_alkan_compatible_product_ids`, Product Tags, Content). Opisy: `scripts/build_simon_woo_brief_input.py`
  (tytul z Woo + atrybuty z CSV Baselinkera + parametry z cache API + "Pasuje do" z
  `dictionaries/simon_24_catalog_facts.yaml`, czyli katalogu serii ver. 26.01) -> briefy
  `--mode codex_brief` -> opisy pisane recznie per SKU (wszystkie kolory symbolu w jednej paczce,
  zeby swiadomie roznicowac teksty) -> `scripts/validate_description_batch.py`. Relacje i zdjecia:
  `src/build_simon_woo_relations.py` z regulami `dictionaries/simon_woo_relations.yaml`:
  tag grupy wariantow = rdzen tytulu bez koloru, pradu, "Push In", klapki i kodu + `24 simon`
  (jak `... 55 simon`); Up-Sells modulu = ramki NR1-NR5 tego samego koloru (49A = kolor 49),
  Up-Sells ramki = popularne moduly w kolorze; Cross-Sells = uzupelnienia w kolorze bez pozycji z
  Up-Sells; kompatybilne TYLKO z katalogu producenta (uszczelki NUV/NUH -> ramki, NSC/NSH ->
  ramki), lustrzane. Zdjecia: oryginal z API e-katalogu (1000x1000 PNG; strona Simona daje tylko
  500x500) -> JPG 1000x1000 na bialym tle, nazwa pliku = slug tytulu, alt = tytul; pliki wgrywa
  sie do `wp-content/uploads/wpallimport/files/` i importuje WP All Import (opisy z FAQ JSON-LD
  tez musza isc WP All Import). Nowe tagi trzeba dopisac do grupy WPC Linked Variation
  "Kolor - Simon" (zrodlo product_tag, os = atrybut `Kolor`; arkusz "Grupy wariantow" w raporcie podaje
  grupe, os i wartosc `Kolor` per SKU, z kontrola powtorzen), inaczej przelacznik kolorow sie nie pojawi. Katalog
  ma pierwszenstwo przed API przy "dostepny tylko w wersji IP20" (NW6/2, NW7/2, NP2, NP31).
  Test: `tests/test_build_simon_woo_relations.py`.
- **Ile paneli LED w jednej paczce zalezy od producenta.** Paczka kurierska
  to 2 kartony zbiorcze producenta: Philips/Signify/PILA 8 szt. (karton 4 szt.),
  Ledvance 12 szt. (karton 6 szt.), pozostali 10 szt. Flexible Shipping liczy
  paczki progami 30/60/90 kg, wiec liczbe sztuk ustawia waga wysylkowa panelu
  w klasie Gabaryt = 28,8 kg / szt. w paczce (3,6 / 2,4 / 2,88 kg). Waga
  gabarytowa tez musi sie zmiescic (GLS dzielnik 5000): przy braku danych
  kartonu `fill_product_weights.py` obniza wysokosc wysylkowa (WARNING), przy
  zweryfikowanym kartonie tylko raportuje. Reguly:
  `dictionaries/panel_shipping_packing.yaml`, test
  `tests/test_fill_product_weights.py`.
- **Odcinki kabli 1 m dziedzicza dane po konfekcjach 25/50/100 m.** Odcinek
  ma SKU bazowe (`5907702810553/EKB`), rodzenstwo ten sam SKU z `/25m`, `/50m`,
  `/100m` i ten sam EAN. `src/fill_cable_1m_from_siblings.py --input <eksport>`
  (nowe wiersze rozpoznaje po tagu `Odcinek 1m`) kopiuje z rodzenstwa
  atrybuty, `Dane producenta`, klase wysylkowa i tagi obecne u wiekszosci
  rodzenstwa (tag grupy wariantow, kampania, tagi paczkomatowe - `Wyklucz
  paczkomat` tez, bo 1 m kupuje sie na metry); `DISABLE` i tagi eventowe nie
  przechodza. **Producent kabli to zawsze `Polski Producent`** (decyzja
  uzytkownika 2026-10-07, nie NKT/Elektrokabel/Zamel/Inny), `Dlugosc` = `1 m`
  ze spacja. Napiecie/izolacja/ekran/zyla ochronna uzupelniane z tytulu tylko
  gdy puste. Tytul odcinka: przekroj z `x` (nie `×`/`X`, jak u rodzenstwa),
  napiecie z jednostka (`450/750V`), kolor z atrybutu, na koncu `1m`. Wynik: kopia eksportu w `output/`, kontrola w `reports/`.
  Eksport bez `Product Tags`/`hwp_product_gtin` (szablon z kolumnami `Attribute Value (pa_...)`)
  wymaga `--tags-source <starszy eksport>`; nowe odcinki to wtedy szkice z SKU bazowym
  i konfekcjami, dostaja tagi `Domyslny dla grupy|Odcinek 1m|<tagi rodzenstwa>`, a plik
  wynikowy ma TYLKO te wiersze (puste tagi reszty WP All Import by wyczyscil). Pusty opis
  = opis konfekcji 100 m bez dlugosci; waga = waga konfekcji / dlugosc, ale nie mniej niz
  sama miedz zyl (n x mm2 x 8,9 g/m) - nizsza waga rodzenstwa jest bledna i idzie do
  sprawdzenia (np. YDYp 2x2,5 100 m = 1,3 kg). Test:
  `tests/test_fill_cable_1m_from_siblings.py`.
- **Puste atrybuty kabli uzupelnia sie z referencji**, nie recznie:
  `src/fill_cable_attributes_from_references.py --input <eksport> --output <kopia>`.
  Kolejnosc: rodzenstwo (ten sam SKU bazowy w innych konfekcjach) > tytul (napiecie
  `450/750V` -> 750, `0,6/1kV` -> 1000, `żo` -> Zyla ochronna Tak) > wartosc typowa
  typu z `dictionaries/cable_attribute_rules.yaml` (YDY/YDYp/OWY/OMY/LGY/YKY: PVC,
  nieekranowany; LGY do 1 mm2 = 500 V, od 1,5 mm2 = 750 V). Istniejace wartosci
  zostaja - rozbieznosci ida do arkusza "Konflikty". Wartosci w regulach to terminy
  sklepu. Po imporcie partii kabli sprawdz eksport tym narzedziem i audytem wartosci.
  Test: `tests/test_fill_cable_attributes_from_references.py`.
- **Tytuly i opisy kabli nie wymieniaja producenta** (decyzja uzytkownika
  2026-10-12). Dotyczy tez linii handlowych (`BiTLAN` Bitnera, `SOLARFLEX-X`
  Helukabla) i kabli sieciowych Bitner. W opisie nie ma tez kodow producenta:
  wiersze `Indeks producenta`, `Numer katalogowy`, `Indeks/kod produktu`,
  `Producent`, `Marka` i kody typu `TI0006` w zdaniach sa usuwane; kod na koncu
  tytulu zostaje. Neutralne "zgodnie z zaleceniami producenta" i
  `Seria producenta: H05VV-F` zostaja. Nazwy nie wycina sie regexem - zostaja
  zdania typu "YDY 5X1 750V marki to"; kazde zdanie redaguje sie recznie, a
  zdania chwalace firme usuwa w calosci. Pierwsza partia:
  `scripts/strip_cable_manufacturer_mentions.py` (lista poprawek per produkt,
  kontrola koncowa `FORBIDDEN`), wynik
  `output/Kable-bez-producenta_2026-10-12.xlsx` (WP All Import - czesc opisow
  ma `<script>` FAQ).
- **Usuwanie produktow = draft + przekierowanie 301** (od 2026-10-12).
  `src/build_product_removal_redirects.py --input <eksport> --redirects
  <eksport JSON Redirection>` z `--categories <kategorie z produkcji>` albo
  `--verify-online`. Eksport WP All Export musi miec `Permalink` (albo
  `Slug`), `Product categories` i `Status`; uzytkownik dopisuje kolumny
  `Do usunięcia` (tak/x/1) i `Zamiennik` (SKU, ID albo URL). Cel: zamiennik,
  a bez niego najglebsza kategoria produktu (remis: nizsze term_id, jak
  wtyczka `alkan-product-category-redirect`); kategoria, ktora zostalaby
  pusta, ustepuje rodzicowi. Kazdy cel przechodzi przez istniejace
  przekierowania (bez lancuchow 301 -> 301). Bez zamiennika raport podaje
  propozycje (ta sama kategoria, zgodne atrybuty i tytul) ze statusem
  `DO_SPRAWDZENIA` - nie trafiaja do importu same. Niejednoznaczny SKU
  zamiennika, nieopublikowany zamiennik, brak URL albo petla blokuja produkt:
  bez redirectu i bez draftu. Wynik: kopia eksportu ze `Status` = `draft`
  (WP All Import), CSV i JSON dla Redirection (grupa `Alkan - usunięte
  produkty`), raport z arkuszami "Istniejace do zmiany" (reguly prowadzace
  na usuwany produkt) i "Lancuchy istniejacych". **Powod:** lokalne drzewo
  kategorii AlkanLocal jest nieaktualne (2026-10-12 brakowalo 29 z 97
  kategorii, na ktore przekierowuje produkcja), a produkcja ma 73 lancuchy
  301 -> 301 (np. 63 reguly na przekierowane `/gniazdka-i-laczniki/podtynkowe/`).
  Wtyczka `alkan-product-category-redirect` jest tylko w AlkanLocal i dla
  draftu nie dziala, dopoki produkt nie trafi do kosza. Test:
  `tests/test_build_product_removal_redirects.py`.

- **Nazwy do Ekonoma to NIE tytuly SEO** (od 2026-10-14). Haslo "robimy nazwy
  do ekonoma" = krotka nazwa dla sprzedawcow z samymi parametrami przydatnymi
  przy ladzie: typ + marka + gwint/dlugosc/rozmiar + W + lm + K (+ IP, DALI,
  czujnik, awaryjna tam, gdzie ma to znaczenie; np. naswietlacz: W, lm, K,
  czujnik). Wejscie: `input/Nazwy do ekonom/<Dostawca>/` - CSV z Ekonoma
  (cp1250, `;`, kolumny LP; Kod towaru; Nazwa towaru; Nazwa dodatkowa; Kod
  producenta = EAN; ...) i PDF-y katalogow jako zrodlo lumenow (po EAN).
  `Nazwa towaru` ma limit 40 znakow - nazwa moze byc dluzsza, reszta idzie do
  `Nazwa dodatkowa` (ciecie na spacji, spacja na poczatku drugiej czesci).
  Marka jest w nazwie (decyzja uzytkownika - "najwyzej sie usunie"; 2026-10-19 na chwile zdjeta
  dla Ledvance i przywrocona tego samego dnia: "jednak chcemy nazwe producenta"). Ledvance: zawsze
  jedna marka - Ledvance albo marka z opakowania Osram/Bellalux (`brand_words`,
  `strip_brand_from_names` usuwa "Ledvance Bellalux"); `brand_in_name: false` zostaje jako
  wylacznik (marka tylko u blizniakow roznych marek). Linia (Value / Performance / Superior / Star...)
  "jesli jest" w kazdym typie Ledvance zaraz po marce (`line_all_types`; typy z seria/modelem maja ja
  w serii) - **tylko zapisana slowem** (VALUE, PERFORMANCE, SUPERIOR, Entry, Pro, Advanced/ADV, Essential, PFM =
  Performance jako samodzielne slowo) albo z rodziny na karcie GTIN (`line_from_card_family`: rodzina "LED CLASSIC B V"
  jest Value, bo karta tej rodziny pisze "VALUE CLAS B 60"; tylko rodziny bez sprzecznych kart, 2026-10-20).
  Sama litera V/P/S w nazwie z Ekonoma ("MR16 DIM S", "SPOT ADJ P 8W") NIE jest linia ani
  wyroznikiem (decyzja 2026-10-19: "moga byc bledne" - karta GTIN nie podaje klasy, "URBAN LITE S" to rozmiar).
  "Ultra Output" (UO) dochodzi do klasy: "Ledvance Performance Ultra Output". Rodzina opraw (`families`, np.
  Urban Area / Lite S / Villa Compact / Wall Vario) stoi za marka przed linia: "... Ledvance Urban Lite S". We wszystkich
  typach Ledvance marka + linia stoja na KONCU (`brand_last`, `move_brand_tail`): parametry i cechy, potem
  wersja / wyroznik, marka + linia, na koncu tylko "op. N szt." ("Downlight LED fi200 18W 4000K bialy
  Ledvance Value"). Srednica jako "220 mm", nie "fi220" (`diameter_mm`); oprawa
  kwadratowa - bok za SQUARE / SQ albo przed SQ i wymiar z mm jako "330x330 mm" (bez liczby nic). Od 2026-10-19 te same zasady
  maja Kobi, Kanlux, GTV i Bemko ("zrob innych producentow wedlug tego"): marka na koncu, a seria oprawy
  z parametrami za nia jak linia (`series_after_brand`; opis z poczatku serii malymi - "z akumulatorem",
  "puszkowy" - zostaje przed parametrami), zarowki w kolejnosci gwint, W, lm, K, ksztalt..., Tuba zamiast
  Swietlowki, "Oprawa pod tuby" zamiast "Oprawa swietlowkowa" (`name_replace`). To zastepuje wczesniejsze
  "marka za rzeczownikiem" (Kobi przed "do ...", Högert za fraza rzeczownika). Ogolnej "Oprawa"/"Oprawa LED"
  ma nie byc: rodzaj bierze sie z kategorii producenta (`noun_overrides` w YAML - wzorzec starej nazwy ->
  rzeczownik, z `types` ograniczajacym do typow ogolnych; GTV: sciezka kategorii na karcie gtv.com.pl,
  Bemko: naglowek karty w katalogu 2026, Kobi: opis w feedzie). Bez zrodla (wycofane, stare kartoteki Kobi
  spoza feedu) zostaje "Oprawa" - nie zgadujemy. Wynik: `reports/Ekonom-ogolne-oprawy_sprawdzone_2026-10-19.xlsx`.
  Zarowki LED
  Ledvance maja kolejnosc ze starego narzedzia ("Zarowka LED GU10 2.7W 230lm 4000K 36ST Philips CorePro"):
  gwint, W, lm, K, ksztalt, kat, klosz, sciemnialna, czujnik, [marka], linia, protokol Smart+ - pseudo-atrybut
  `brand` w `attrs` ustawia pozycje marki. Rzeczowniki Ledvance (2026-10-19): swietlowki to tuby (`Tuba LED`,
  `Tuba fluorescencyjna`, `pod tuby LED` przez `tube_words`), oprawy wpuszczane/spoty to `Downlight LED`, a ogolnej
  "Oprawa LED" ma nie byc - kazda oprawa dostaje konkretny typ (uliczna, plafon, lampa stojaca, do zadaszen,
  scienna, awaryjna, modul awaryjny...). Brak lm
  = nazwa bez lm i status `DO_SPRAWDZENIA`; Ledvance najpierw bierze brakujace lm z karty
  `ledvance.com/en/gtin/<EAN>` (`lumen_from_gtin_card`, `scripts/fetch_ledvance_web.py --missing-lumen <raport>`),
  ale tylko gdy moc i barwa karty zgadzaja sie z Ekonomem (Power Select "55 W / 49 W" + "6650 lm / 5900 lm" - stopien
  o mocy z Ekonoma; Multi Lumen "40W" = najwyzszy stopien 39,3 W; komplet SETn - lm na sztuke; kilka mocy w nazwie
  "11,3/14W" / "4-7-9W" -> lm w tej samej kolejnosci "1900/2300lm" / "500-720-840lm"). Narzedzie
  `src/build_ekonom_names.py --supplier <x> --input <csv> --pdf-dir <katalog>`,
  reguly typow `dictionaries/ekonom_name_rules.yaml`, wynik: CSV do Ekonoma
  w `output/` + skoroszyt kontrolny w `reports/` (arkusze "Nazwy",
  "Powtorzenia", "Roznice w powtorzeniach"). **Plik wynikowy w `output/` ma
  kodowanie UTF-8 z BOM, separator `;` i TYLKO kolumny `Kod towaru`, `Nazwa towaru`,
  `Nazwa dodatkowa`** (decyzja uzytkownika 2026-11-18, `OUTPUT_FIELDS` /
  `write_ekonom_csv`). Dotyczy wylacznie pliku wynikowego - CSV z Ekonoma na wejsciu
  dalej czyta sie w cp1250 z pelnym ukladem, a skoroszyt kontrolny sie nie zmienia. **Powtorzenia nazw zglasza sie
  uzytkownikowi** razem z tym, czym roznia sie produkty - to on decyduje,
  jaki atrybut dodac do nazwy; nie dobiera sie wyroznika samodzielnie.
  Decyzje Ledvance (2026-10-14): zarowki dostaja linie (Value, Performance,
  Superior, Star...), ksztalt (A60, swieczka, kulka, globe, sztyft, PAR16...),
  klosz (przezroczysta/matowa), "sciemnialna", a reflektorowe (PAR16, R50,
  MR16) kat swiecenia z nazwy albo z wiersza PDF, Smart+ protokol (WiFi /
  Zigbee / BT / Matter), z czujnikiem - "z czujnikiem"; oprawy kat swiecenia/optyke i kolor
  obudowy, panele UGR19; gdy nazwa dalej sie powtarza, typy z
  `version_suffix: true` dostaja wyroznik wersji z nazwy katalogowej
  (FIX/ADJUST, DARKLIGHT, HIGH OUTPUT...). Produkty bez parametrow swietlnych
  (runda 6): oprawy dekoracyjne = rodzaj z montazu (Lampa wiszaca, Kinkiet,
  Plafon...) + rozmiar + gwint x ilosc + kolor + marka + seria; Smart+ =
  rodzaj + "Smart+" + protokol + parametry + marka + seria (seria na koncu za marka od 2026-10-20,
  `series_after_brand`: "Lampa stolowa E27 brazowa Ledvance Decor Cardboard Square"); dekor z pilotem
  (WITH REMOTE CONTROL / + RC) = "z pilotem", a "with" / "W/" / "+ W" nie zostaja w serii ("Tasma LED 5m TW z pilotem
  Ledvance Cob Strips"); tasmy FLEX RGB / AUDIO / BENDABLE z pilotem to dekor, nie akcesorium "Pilot"; "Reflektor do
  szynowych opraw TRACKLIGHT SPOT REFLECTOR D75 SP" = "Odblysnik do Tracklight Spot 75 mm SP" (akcesorium); akcesoria = rodzaj po polsku +
  "do <system>" (TruSys Flex, Downlight Comfort...); akcesoria do tasm LED,
  DALI i moduly z kodem modelu na koncu; lampy wyladowcze z typem, W, trzonkiem.
  Slowniki rodzajow i systemow sa w YAML (`mount_nouns`, `device_nouns`,
  `accessory_nouns`, `accessory_systems`). Ledvance runda 7 (2026-10-15):
  te same lampy sa sprzedawane jako OSRAM / LEDVANCE / BELLALUX z osobnymi EAN -
  nazwa dostaje marke z opakowania (kolumny EAN marek w polskim katalogu
  `zrodla_swiatla_led`, koncowka kodu `...2X1OSRAM` / `FS3 OSRAM` w nazwie);
  multipak dostaje na koncu `op. N szt.` (z nazwy: `3pack`, `FS3`, sklejony kod
  `4X3LEDV`; z katalogu tylko `retail_lamps_program`, liczba po trwalosci
  15 000 h); barwa z katalogu retail poprawia nazwe, gdy moc sie zgadza, a gdy
  wiersz katalogu opisuje inny produkt - DO_SPRAWDZENIA bez zmiany nazwy.
  Nazwy, ktore dalej sie powtarzaja, dostaja wyroznik (`distinguish`): **tyle
  informacji, zeby sie roznily, nie wiecej** - kazdy produkt dostaje najmniejszy
  zestaw slow, ktorego nie ma w calosci zaden inny z grupy (slowa przed liczbami;
  "Globe 80", "36°", "czarny", "Superior", "z czujnikiem zmierzchu"). Zrodla po
  kolei: katalog PDF (tylko katalogi lamp z nazwa przy EAN, tylko w podgrupie
  jednego stylu zapisu - retail vs broszura to ten sam produkt), stara nazwa
  w Ekonomie, nazwa ze strony ledvance.com (`scripts/fetch_ledvance_web.py` ->
  `cache/ledvance_web/`); kilka rund, bo np. "-150" odrozni sie dopiero po
  oddzieleniu "-150 OEM". Najlepsze zrodlo to karta `ledvance.com/en/gtin/<EAN>`
  (curl dziala; dokladnie ten EAN: nazwa, kod zamowieniowy, sztuk w opakowaniu) -
  daje multipaki "op. N szt." i, gdy opisuje pod EAN inny produkt niz Ekonom
  (moc/barwa/gwint), nazwe budowana z karty (`fix_from_gtin_page`, WARNING ze
  stara nazwa w uwagach; 2026-10-15: 66 bledow kartotek, np. 9x "T8 EM PLUS 600
  8W 830" to tuby 60-150 cm 8-20W). Podpowiedzi wyszukiwarki Ledvance dopasowuja
  EAN nieostro - tylko gdy moc, barwa i gwint zgadzaja sie z Ekonomem. Kod barwy
  9xx (940, 93040) = "CRI 90", ale gdy powtorzone nazwy roznia sie CRI na karcie GTIN, wyroznikiem jest CRI z
  karty (`add_card_cri`, 2026-10-20: "MR16/PAR16 DIM S" = CRI 97, "DIM P" = CRI 90). Nazwa wlasna oprawy (2026-10-20,
  "brakuje nazw wlasnych") z rodziny na karcie GTIN (`family_from_card`): bez slow technicznych (DALI, SENSOR, GEN 2, IP,
  UGR, wymiary, kolor) i bez slowa, ktore juz stoi na poczatku nazwy; stoi NA KONCU za marka, przed linia (`family_after_noun:
  false`; decyzja 2026-10-20 "nazwa wlasna jednak na koncu wszedzie", uchyla "zaraz za rzeczownikiem" z v9-v15, dotyczy
  tez Urban i opraw pod tuby): "Naswietlacz LED 50W 5000lm 6500K 100° bialy Ledvance Floodlight Compact", "Panel LED 60x60
  ... Ledvance Flex", "Naswietlacz LED 150W ... Ledvance Floodlight R30 Performance", "Oprawa hermetyczna pod tube LED
  1x120cm IP65 Ledvance Damp Proof Housing"; za tym tylko "op. N szt.".
  Czujnik / awaryjna / DALI z rodziny karty, gdy Ekonom ich nie podaje ("SURFACE SLIM RD S 200" = karta "... SENSOR").
  Kod optyki (N / W / VW / VN / WW / NFL / OV / DS) zastepuje kat swiecenia z karty ("Beam angle", pole `beam` w
  cache, `fetch_ledvance_web.py --missing-beam`): Low Bay Flex N = 60°, W = 90°, VW = 120°, OV = 40x90°; bez karty
  zostaje "opt. N". "840 W" to barwa + optyka W, nie moc. Tylko w oprawach hermetycznych (`family_extra`): HE /
  HIGH EFFICIENCY -> "Damp Proof High Efficiency", 5X THROUGHWIRING / 5TH -> "5x Throughwiring", THROUGHWIRING / TH ->
  "Throughwiring", ECO CLASS (z karty; Ekonom pisze "LED ECO") -> "Eco Class Damp Proof" na poczatku, HLO -> "High Lumen
  Output"; "Narrow / Very Narrow / Wide" wypada, bo jest kat (`family_strip`); w downlightach / tubach HE bez zmian.
  Karta "Phase out" z nazwa bez liczb ("DAMP PROOF ECO VERY WIDE") daje nazwe wlasna. Dlugosc "0.6M / 1.2M" = 60 / 120cm.
  Plafony (przeglad 2026-10-20): srednica z samej liczby 200-699 (`diameter_bare`: "SURFACE BULKHEAD 250", "SF DISC 220 V"
  - V to klasa, nie napiecie), "30CM" = 300 mm, "30X30CM" = 300x300 mm, "Square 200mm" / "SQUARE DALI 500" = kwadrat,
  CIRC = okragly; kolory dwutonowe "Black and Blue" -> "czarno-niebieski", "White and Wood Decor" -> "bialy/drewno"
  (`two_tone_colors`, `stem` w `housing_colors`); "LED Color + White" to nazwa, nie kolor obudowy (`rename`,
  `color_ignore`); "Orbis Click Sensor" -> nazwa z karty ("Orbis Square / Sparkledot / Plate Click", `refine_from_name`).
  Czujnik / awaryjna / DALI tez z nazwy produktu na karcie; gdy potwierdza je tylko nazwa, a rodzina nie - WARNING. Skroty Ledvance EL = awaryjna, SN = z czujnikiem (`feature_abbrev`). Plafony: Click-DIM / Click-CCT /
  Remote-CCT (atrybut `click`, Orbis; przy Click-CCT bez osobnego "CCT"). GU5.3 zawsze z napieciem: "GU5.3 12V"
  (230V tylko gdy nazwa / karta tak mowi, `gu53_voltage`); "Damp Proof", "Panel Compact", a takze pojedyncze
  "Bulkhead", "Batten", "Spot", "Floodlight" to nazwy wlasne (2026-10-20: "tutaj nie powinien byc bulkhead jako nazwa
  wlasna?" - `generic_words` puste); rodzina bez niczego poza rzeczownikiem ("HIGH BAY V", "PANEL PERFORMANCE 600")
  daje sam rzeczownik + linie, bo "High Bay" / "Panel" wypada jako powtorzenie.
  Bez karty albo z karta "Phase out" / "Coming soon" (to status, nie rodzina - nie przenosi tez linii): nazwa wlasna
  innych kart, ktora stoi doslownie w nazwie z Ekonoma ("DAMP PROOF HE DA 1500" -> Damp Proof), w tym rodziny z
  podpowiedzi wyszukiwarki ("PLANON™" w Ekonomie = Planon; karta 4058075470712 zwraca HTTP 500). Oprawa okragla
  (ROUND / RD / CIRCULAR w nazwie z Ekonoma) ma "okragly/okragla" zaraz po rzeczowniku: "Panel LED okragly
  300 mm 19W ... Ledvance Planon Frameless", "Plafon LED okragly 250 mm ..." (atrybut `round` w panel / downlight / plafon / wpuszczana); gdy nazwa wlasna mowi juz Round / RD / Circular
  ("Ceiling Round", "Surface Slim RD"), "okragly" sie pomija (2026-10-20). Gdy nazwa z karty nie pasuje do zadnego typu,
  zostaje nazwa z Ekonoma (WARNING), a nie surowa angielska nazwa karty. Wyroznik sprzeczny z nazwa (matowa/przezroczysta,
  ksztalt) daje DO_SPRAWDZENIA zamiast dopisku. "Warm/Cool/Tunable White" to
  barwa swiatla, nie kolor obudowy. Multipaki tez z "x10", "4x", "Triple Pack".
  **Cen nie uzywamy jako dowodu** (decyzja uzytkownika). Reszte powtorzen
  przeglada `scripts/review_ekonom_duplicates.py` (nazwy katalogowe obok
  grup, kolumna "Decyzja uzytkownika"). Kobi (od 2026-10-14): atrybuty z
  feedu Ceneo `--feed input/Nazwy do ekonom/Kobi/ceneo_kobi.xml` (dopasowanie
  kodu z `Kod towaru` 004666/KOB do `Kod producenta` feedu), marka = linia
  z feedu (Kobi, Kobi Premium, Kobi Pro, Kobi Design, LED2B, LLS...), nazwa z
  feedu zastepuje tylko stare nazwy WIELKIMI LITERAMI, a feed jest pomijany,
  gdy moc albo barwa nie zgadzaja sie z nazwa w Ekonomie (uwaga w raporcie).
  Kody ksztaltu Kobi: GS = A60, MB = kulka, SW = swieczka, F... = filament;
  "LX" = czujnik (rodzaj z `Typ czujnika` feedu). Kobi runda 2 (2026-10-15): marka idzie za
  rzeczownikiem, przed "do ..." ("Rozeta sufitowa Kobi Pro do KOLINE K2"; wyjatek "Oprawa do nadbudowania",
  "Lampka na monitor"), oprawa na wymienne zrodlo (1xGU10, G5.3/GU10) nie wymaga lm, swietlowki T8 maja
  dlugosc, wyroznik z samej cyfry/litery dostaje slowo modelu ("SOLARIS 2", "Quazar 15"). Produkty
  spoza feedu (stare kartoteki, ~230) nie maja lm z zadnego zrodla - nastepcy w feedzie to czesto inna
  seria (ZOE != SIGARO), wiec lm nie przenosimy. Kobi runda 3 (2026-11-05): wejscie to katalog
  `archived_input_files/input_2026-11-10/EkonomNazwy/Kobi` ("kata" + 391 nowych kartotek "kart"), feed `ceneo_kobi.xml` jest tylko w stashu git
  `4ad6dd9` (bez niego nazwy traca kolory i nazwy z feedu). Nazwa od kodu modelu ("LED VESPA PRO 40W 4000K Lampa
  uliczna", "HERMETIC 1X60 Oprawa pod tuby led") tez bierze nazwe z feedu (`prefer_feed_name_model_first`), a
  `name_cleanup` czysci i nazwe feedu (`feed_name_cleanup`: "PRAWE", "Kob i", "25-36- 40W", "do nabudowania").
  HERMETIC / HERMIC / HERMES n x dlugosc = oprawa pod tuby ("pod tuby LED 2x120cm"), plafon na zarowki (2xE27, 1X60W)
  to "Plafon", nie "Plafon LED"; marka feedu "Nexus" = Kobi Pro (`brand_map`); VESPA PRO = oprawa uliczna (stare VESPA spoza feedu zostaja naswietlaczem); oczka
  OH-xx z ksztaltem i modelem; "STALE" to nie stal. "LED <MODEL> W K" bez rodzaju (DEFENDER, INDUO) = "Oprawa LED"
  + model, DO_SPRAWDZENIA - rodzaju nie zgadujemy (SUNFLASH zostaje bez reguly; POWER BOX = "Bank energii 600W Kobi Power Box",
  decyzja uzytkownika). Kanlux (od 2026-10-14):
  parametry z plikow `Parametry*.xlsx` z kanlux.pl (`--params-dir`, format dlugi
  Kod Kanlux / Nazwa atrybutu / Wartosc, dopasowanie po `Kod towaru` 37246/KAN) sa
  zrodlem prawdy i nadpisuja nazwe (kod `WW` u zarowek bywa 2700K, nie 3000K).
  Nazwa Kanlux = kod modelu + polski typ - rzeczownik bierze sie z typu
  ("Oprawa liniowa LED"), seria to kod modelu bez parametrow (`AL-SL-MAT-NT`).
  Typ "Oprawa pyloszczelna" w nazwie to "Oprawa hermetyczna" (takze odmiana: "Akcesorium
  oprawy hermetycznej"; `name_replace` Kanlux, decyzja uzytkownika 2026-11-09).
  "Moc maksymalna" przy wymiennym zrodle (trzonek GU10/E27) to limit zarowki, nie
  moc oprawy. Osprzet DOMO/LOGI/TEKNO: typ + seria + kolor + indeks na koncu (kody
  kolorow bi/kr/sr/gr/sz/pe/cm w YAML; `cm` = czarny mat - tak podaja parametry Kanlux LOGI);
  Kanlux runda 2 (2026-10-15): kody spoza plikow Parametry bierze `scripts/fetch_kanlux_web.py`
  (podpowiedzi `kanlux.com/pl/ajax/products/autocomplete/<kod>` + karta `/pl/produkt/<sku>/<seourl>`,
  cache `cache/kanlux_web/`, czytany przez `params_web_cache`); kod nieznany wyszukiwarce = produkt
  wycofany - zostaje z danymi z nazwy, bez dalszej pracy (jak Kobi spoza feedu). Oprawa z trzonkiem w
  parametrach (G13/GU10/E27...) nie wymaga mocy/lm/K i ma w nazwie trzonek z krotnoscia (NOTUS 418 ->
  4xG13); MOD/MODv2 = "Moduł LED"; powtorzenia rozniace sie rozmiarem z kodu dostaja rozmiar przy serii
  (STOBI DLP 350 / 450);
  wylaczniki KMB/KRD i styczniki KCP dekodowane z kodu (B16 1P 10kA).
  Kanlux runda 3 (2026-11-04): wejscie to katalog `archived_input_files/input_2026-11-10/EkonomNazwy/Kanlux` ("kata" = baza z rund 1-2,
  "kart" = 1156 nowych kartotek jako "<40> | <reszta>"), `--input <katalog>` jak Bemko. 14 plikow
  `Parametry*.xlsx` z rund 1-2 nie ma juz w `input/` - sa tylko w stashu git `4ad6dd9` ("input/Nazwy do
  ekonom/Kanlux/"); `--params-dir` z nimi odtwarza v2, a reszte kodow pokrywa `fetch_kanlux_web.py`
  (3386 kart, 94 wycofane). Bez tych plikow kata traci lm/kolor u kodow wycofanych z kanlux.com. Pusty klucz
  karty kanlux.com (ikona klasy energetycznej = "E") nie jest atrybutem - dawal marke "E". Nowe kody
  aparatury: KMBN6 (C10 1P+N 6kA), KRO = "Wylacznik RCBO" (B10 2P 30mA typ A 6kA), KMI / KCS (63A 3P),
  KMC (20A 1NO+1NC), KSD / KSPD (T1+T2 3P+N, M = wkladka, TT = T2 TT), KMCCB-F (125A 4P). "UGR≤" idzie do
  Ekonoma jako "UGR<" (cp1250 dawal "UGR?"); IP powtorzone z typu ("Oprawa punktowa LED IP44") zostaje raz.
  Grupa Ekonoma "Rozgalezniki, gniazda, wtyczki" (ZARK / PIN / GRIP / STAN, 2026-11-05): typy `rozgaleznik`, `wtyczka`,
  `gniazdo_przenosne`, cechy z karty kanlux.com (`param` w `features`: "Liczba gniazd", "Rodzaj uziemienia" typ E = z
  uziemieniem / typ F = Schuko, "Prad znamionowy"), uklad jak rozgalezniki Kobi / Bemko: "Rozgaleznik 3 gn. z uziemieniem
  bialy ZARK KRT/2-1 Kanlux", "Gniazdo przenosne 1 gn. Schuko IP44 czarne GRIP GP/S Kanlux".
  GTV (od
  2026-10-14): `Kod producenta` to kod GTV (LD-SMBD45-10), nie EAN - katalog PDF jest
  czytany po kodzie i tylko z komorek z jednostka (`1000 lm`, `650 lm/m`, `3200 K`);
  lm z nazwy maja pierwszenstwo, katalog uzupelnia braki, a `150lm/W` x moc daje lm.
  Rzeczownik = slowa typu z poczatku nazwy (`type_words` w YAML), model = reszta
  pierwszego czlonu; "max. 20W" przy oprawach na zarowke to limit zrodla, nie moc.
  Kody HT... to narzedzia i odziez Högert (marka `Högert` za fraza rzeczownika, przed przyimkiem i
  liczba: "Palnik gazowy Högert", "Cyfrowy miernik uniwersalny Högert", "Brzeszczot Högert do ...";
  odziez "MODEL rzeczownik" -> "Rzeczownik Högert MODEL rozmiar"), GT-... to linia `GTV G-TECH`.
  GTV runda 2 (2026-10-15): karta `gtv.com.pl/produkt/<kod>/` (tabela "Dane techniczne") uzupelnia braki
  lm/K/IP - `scripts/fetch_gtv_web.py`, cache `cache/gtv_web/` (`web_params_cache`); 404 = wycofany, bez
  dalszej pracy. Slowa pisane na stronie GTV czesciej malymi ("z przewodem", "typu") poprawiaja wielkosc
  liter starych nazw WIELKIMI, czesciej WIELKIMI = model. Powtorzenia Högert (identyczne nazwy w Ekonomie)
  rozstrzyga sie danymi z `hoegert.com` (`scripts/fetch_hoegert_web.py --review`) - zestawienie roznic
  idzie do uzytkownika, atrybut do nazwy wybiera on.
  Bemko (od 2026-10-14): nazwy w Ekonomie to juz parametry WIELKIMI LITERAMI; marka z konca nazwy
  (Bemko / Luxmena / Kafler / Schelinger), a gdy jej brak - z rodzenstwa w rodzinie kodu (`brand_family`,
  np. C66-LFC: jedna nazwa "Luxmena" -> cala rodzina Luxmena; uwaga w raporcie). Lumeny brakujace w nazwie
  biora sie z kart katalogowych (`--pdf-dir ".../Bemko/Karty katalogowe"`, `pdf_match: code_table` - wiersz
  kod | EAN | opak. | kolor | K | W | lm, moc musi sie zgadzac); lm z nazwy wygrywa, rozbieznosc >10% z
  katalogiem idzie do uwag. `code_checks`: kod producenta przeczacy nazwie (-6K przy 4000K, A03-V7-2-63-100
  przy "25A 300mA", tasma ACW przy "BIALA CIEPLA") daje DO_SPRAWDZENIA - nazwy nie poprawiamy. Czujnik z
  kodu (-MS/-MD/-PIR), gdy nazwa go nie podaje. Aparatura: "Wylacznik nadpradowy Schelinger B16 1P 10kA",
  RCBO, roznicowopradowe, ograniczniki T1+T2; przedluzacze: gn. / z uziemieniem / wylacznik / przekroj /
  dlugosc; oprawy "EMPTY" = pod T8 (18W = 60cm, 36W = 120cm, 58W = 150cm). Cechy per typ (`features`,
  `features_pre`, `features_post`, wspolne listy `feature_sets`, `group` = jedna wartosc z grupy).
  Bemko runda 2 (2026-10-15): karty `bemko.pl` po kodzie (`scripts/fetch_bemko_web.py`, cache `cache/bemko_web/`;
  wyszukiwarka `/pl/wyniki-wyszukiwania/1?search=<kod>` -> kafelek "NAZWA (KOD)" -> karta z tabela parametrow).
  **bemko.pl blokuje IP po ~250 zapytaniach w 5 watkach** - pobierac 1 watkiem z przerwa 2 s (domyslne), a
  "brak produktu" z okresu blokady jest falszywy. Pole karty `Marka` zastepuje marke z rodzenstwa; lm/K/IP z
  karty tylko w brakach; gdy nazwa w Ekonomie przeczy karcie (moc, barwa, prad A/mA, `code_checks`) - nazwa z
  karty, WARNING ze stara nazwa w uwagach (`apply_web_card`, np. Plato 450-360 = 36W 3250lm, nie "24W 2160LM").
  Akcesoria `inne`: marka za rzeczownikiem przed "do/z/na" (`brand_before_prep`), caly opis bez wycinania
  powtorzen (`series_all_segments`).
  Bemko runda 3 (2026-11-03): wejscie to katalog `archived_input_files/input_2026-11-10/EkonomNazwy/Bemko` ("kata" = baza z rund 1-2 w ukladzie
  Ekonoma, "kart" = 340 nowych kartotek jako "<40> | <reszta>"), `--input <katalog>` jak Signify. Zestawienie gubi
  spacje na cieciu: pierwsza czesc krotsza niz 40 znakow dostaje spacje przed reszta (`read_csv`), chyba ze to
  liczba i jej ciag dalszy / jednostka ("4000" + "K", GTV '1/4", 6' + '0 mm' - Ekonom liczy cudzyslow podwojnie).
  Literowki Ekonoma ("600 0LM", "4400L M", "I P44", "AT 1H", "z CZUP44") naprawia `name_cleanup`. Nowe typy:
  zarnik R7s, zarowka halogenowa, swietlowka, ramka montazowa do paneli (z wymiarem 60x60 / 30x120), zasilacz do
  paneli (mA), oprawa stadionowa, awaryjna z lm (`awaryjna_lm`: panel / hermetyczna z modulem awaryjnym niesie
  wymiar, lm i K; male awaryjne bez nich), Plato 24V DC; "OPRAWA RAMA LED ... PLAZA" = Panel LED. Plafon LED bez
  lm w nazwie (moc, bez trzonka) to `plafon_led` z lm z karty. "CCT" z nazwy nie jest nadpisywane barwami karty;
  nazwy kart WIELKIMI LITERAMI (takze z jednym "mini") nie ucza pisowni slow (`load_web_params`).
  Przeglad 2026-11-04: ogranicznik "KLASY T1T2" = T1+T2; w akcesoriach `inne` "+" zostaje ("(4NO+0NC)"), uklad
  "2x15" tez, jednostki malymi (`series_unit_lower`: "290mm", "1,5m", "160x160mm", "fi60mm", "10 szt."), kody
  SSO / LFX WIELKIMI, "grafit" / "typ" malymi.
  Przeglad Ledvance partiami (2026-10-21, v19): `name_cleanup` naprawia zapis Ekonoma (PANELANON -> PLANON,
  SUNATHOME -> SUN@HOME, Obris -> Orbis, sklejone RGBTWBK / ERECESSEDMOUNTFRAME, skroty PEND / CLICKD / REM / CAM), a
  `undo_excel_date` cofa daty Excela w `Nazwa dodatkowa` ("05-lip" = 7/5, "27-maj" = 27/5, "maj-40" = 40/5 - koncowki
  kodow tasm). Wlasne typy maja transformatory do halogenow (Ekonom pisze "Statecznik"), zarowki do piekarnika /
  sygnalizacyjne / tradycyjne, latarki, zestawy Google Nest, klosz HQL, akcesoria High Bay i adaptery czujnikow; kamery,
  termostaty, piloty, wentylator, panel solarny i kotwy Endura / Urban to wlasne rzeczowniki, nie "Oprawa". Karta GTIN
  uzupelnia braki tylko w `card_fill_types` i gdy opisuje ten sam produkt; przebudowa z karty zachowuje typ zarowki
  (Smart+ Spot R50 / GU10 to zarowka) i z Ekonoma ksztalt, TW / RGBW, linie, DIM, PCR, rozmiar banki - tylko gdy moc to
  literowka (inaczej to inny produkt). Czujnik z rodzajem (ruchu / zmierzchu), NAV "/I" = z zaplonnikiem, Multicolor =
  RGBW tylko w Smart+, A75 / A100 = A60 (zamiennik mocy, nie ksztalt).
  Decyzje po v19 (2026-10-21, v20; lista `reports/Ekonom-Ledvance_do_decyzji_2026-10-21.md`): karta GTIN z innym
  trzonkiem, ksztaltem banki albo dlugoscia tuby to INNY produkt - zostaje nazwa z Ekonoma, DO_SPRAWDZENIA
  (`gtin_other_product`; sama inna moc = literowka, nadal nazwa z karty). Klosz Gold / Smoke / Amber po angielsku, jak na
  opakowaniu (zastepuje "przezroczysta"); ksztalty Vintage tez po angielsku. "3 STEP DIM" = "3-stopniowa", Duo Click /
  GlowDim nazwa z opakowania. "z pilotem" tez w oprawach parametrycznych, "sciemnialna" nie (to mowi DALI / 1DIM).
  "ON-OFF" tylko przy blizniaku DALI / 1DIM (`drop_lone_onoff`). Optyki uliczne zawsze kodem "opt. RV20ST" (RV i RW to
  rozne rozsyly). "Retrofit" jest wyroznikiem; "BW" w swieczkach wypada ("plomyk"), takze jako wyroznik. Print Sky: numer
  wzoru przy nazwie wlasnej ("Print Sky 1"), SETn = "op. n szt.". Decor Spot "3-punktowy". HQI D / DL: K z karty GTIN
  (< 5000 K = WARNING). Kamery Smart+ ze swiatlem to oprawy z kamera (`kamera_oprawa`: "Kinkiet LED z kamera Smart+ WiFi
  15W 950lm 3000K ..."). Surowe nazwy EN tlumaczy sie tylko po akceptacji arkusza z `scripts/propose_ledvance_translations.py`.
  Zestawienie Ekonoma (2026-10-26, `input/Zestawienie Ledvance.csv`): bez kolumny `Nazwa dodatkowa`, obie czesci w
  `Nazwa towaru` jako `<40 znakow> | <reszta>` - `read_csv` je rozdziela, a wynik ma ZAWSZE osobna `Nazwa dodatkowa`
  (od 2026-11-18 tylko `Kod towaru;Nazwa towaru;Nazwa dodatkowa`, UTF-8 z BOM - patrz wyzej). Uklad zestawienia
  z ` | ` NIE jest formatem do Ekonoma (poprawka uzytkownika 2026-10-27). Nowa partia idzie przez przebieg laczny z baza v20 (`Ledvance.csv` + zestawienie), zeby wyrozniki
  liczyly sie tez wzgledem produktow juz nazwanych; do wyniku ida tylko nowe pozycje, a nazwy w v20 zostaja. Przed
  przebiegiem pobierz karty GTIN nowych EAN (`fetch_ledvance_web.py --ean-file`). Stare kartoteki Osram z kodem na
  poczatku maja wlasne typy i porzadki: zarowki samochodowe (zawsze Osram; NBU / NL = Night Breaker Unlimited / Laser),
  Linestra = "Zarowka liniowa", HALOSTAR / HALOLINE / HALOLUX / "ZAR. HAL." = halogen, HCI / HMI = metalohalogenkowa,
  "L15/76" = L 15W/76 (Natura, Fluora bez K), "DS7/41-827" = DULUX S, "ST8A-0.9M" = tuba T8 90 cm, lm "1K2LM" / "10KLM",
  marka Osram tez przy koncowce EAN ("OSRAM 7426"), DUOBOX = op. 2 szt.
  Swietlowki liniowe Ledvance to "Swietlowka", nie "Tuba fluorescencyjna" (decyzja uzytkownika 2026-10-26); "Tuba LED"
  i oprawy "pod tuby" bez zmian.
  Signify (od 2026-11-03, runda 1): wejscie to KATALOG `archived_input_files/input_2026-11-10/EkonomNazwy/Signify` z 4 zestawieniami (FAR = oprawy,
  PHP = zrodla; "kat" w ukladzie Ekonoma, "kart" jako "<40> | <reszta>") - `--input <katalog>` liczy nazwy i powtorzenia
  razem, a wynik zapisuje osobno dla kazdego pliku (`output/Ekonom-nazwy-Signify_<plik>_<data>.csv`, uklad Ekonoma).
  `Kod producenta` to 12NC Philipsa; karta `lighting.philips.pl/.../product-detail-page.<12NC>_EU.html`
  (`scripts/fetch_signify_web.py`, cache `cache/signify_web/`, ~25 min / 2700 kodow w 10 watkach, potem
  `--retry-missing`) jest zrodlem prawdy jak pliki Parametry Kanlux (`params_web_loader: signify`): moc, lm, K, IP, trzonek,
  kolor, wymiary, DALI, UGR. 404 = brak w ofercie (2026-11-03: 817 z 2674) - dane tylko z nazwy, bez dalszej pracy.
  Zmiane parametru wzgledem Ekonoma notuje uwaga "lumen z karty ...: X (w Ekonomie: Y)" (`params_change_note`).
  Swietlowki, kompakty, halogeny i HID biora moc NOMINALNA z nazwy (`params_skip: [power]`; karta podaje 79,8 W przy 80W).
  Wzor: parametry, potem marka (Philips / PILA), kod rodziny oprawy w pisowni producenta (BY120P, RC132V, DN142B) i linia
  (MASTER, CorePro, CoreLine, Ledinaire... - z nazwy, a gdy jej brak z kategorii karty) na koncu: "Downlight LED 150 mm
  10,7W 1000lm 3000K IP44 UGR19 DALI bialy Philips DN142B CoreLine", "Zarowka LED E27 4W 840lm 2700K A60 przezroczysta
  sciemnialna Philips MASTER". "LED classic 40W" / "PILA LED 60W" to ZAMIENNIK mocy - bez karty zostaje "zamiennik 40W"
  i DO_SPRAWDZENIA. Kod Philips "LED55/740", "36S/840" = 5500 / 3600 lm i 4000K; PSD / PSR = DALI; "IP20/44" = IP44.
  Stare narzedzie nazwalo czesc swietlowek TL-D Colored / TL-E / TL Mini i zarowek do piekarnika "Tuba LED" - kategoria
  karty to poprawia (`card_rename`). Decyzje 2026-11-03 ("dodaj to, co opisales w grupach powtorzen, i nazwy wlasne"):
  wyrozniki z nazwy i z nazwy katalogowej karty (`features_card_name`) - optyka drogowa kodem "opt. DM10", klasa I / II,
  optyka naswietlaczy i High Bay ("opt. sym." / "opt. asym." / "opt. OFA52" / "opt. wąska" / "opt. szer." - za kolorem, tuz przed marka; wytyczne
  2026-11-04: "Naswietlacz LED 31W 4263lm 3000K IP66 szary opt. sym. Philips BVP111 CoreLine"), linie swietlne "opt. WB przewody 5x1,5 HE", panele
  OC / NOC, generacja G4 / G5, "awaryjna 3h" (ELB3), downlighty kat 36° i oslona "opt. C / WR", tuby i zarowki "sciemnialna",
  "CRI 90" (kod 9xx / CRI95), ksztalt plomyk (BA35) i A67; reszte blizniakow rozroznia kod wersji z nazwy katalogowej karty
  (`version_suffix_all`, `version_from_card`, np. PSR / PSD). Nazwa wlasna rodziny (`proper_names`, slownik kategorii karty)
  stoi przed kodem: "... Philips Malaga BRP102 CoreLine", "Philips ClearWay BGP307"; produkty bez karty dostaja ja po kodzie
  rodziny. Slowa kategorii (Downlight, Highbay, Waterproof, Batten, Panel) nie sa nazwa wlasna. Powtorzenia 251 -> 63 grupy
  (28 identycznych juz w Ekonomie). Kontrola 2026-11-03: TForce "60-35W" = 35 W rzeczywiste (odwrotnie niz "LEDBulb
  8-60W"), "LED760/WW" (3-4 cyfry) to juz lumeny, a "LED42/NW" (2 cyfry) = 4200 lm; MDU = czujnik ruchu; MR16 bez
  trzonka = GU5.3, PAR20/30/38 i R63/R80 = E27; kod W3 / W5 paneli przed ELB3, bo "ELB3 W3" czytalo sie jako moc 3W.
  Braki (2026-11-03): kod bez strony czesto ma jeszcze karte PDF `docs.signify.com/assets/Leaflet/Philips_PROF/FP/<pl_PL |
  en_GB | en_AA>/<12NC>_EU.<jezyk>.PROF.FP.pdf` - `scripts/fetch_signify_leaflets.py` (po `fetch_signify_web.py`) zapisuje ja
  do `cache/signify_web` jak karte strony (`source: pdf`, uwagi "karta PDF docs.signify.com"); 495 z 813 kodow, DO_SPRAWDZENIA
  491 -> 147 (reszta: wycofane zarowki konsumenckie LED classic / PILA i stare oprawy bez zadnej karty). Karta, ktora zmienia
  co najmniej dwa z: moc / lm (>15 %) / barwa, opisuje inny produkt niz kartoteka - nazwa z karty, WARNING "karta przeczy
  nazwie w Ekonomie". Dlugosc belek / hermetycznych / linii i wymiar paneli nominalny z nazwy (`params_keep_name`: L1500 =
  150cm, nie 145 z karty); WW / NW przy slupkach i najazdowych = 3000K / 4000K (potwierdzone kartami PDF).
  Przedluzacze (schemat uzytkownika `archived_input_files/input_2026-11-10/schemat nazwy przedłużaczy V3.xlsx`, 2026-11-05, wszyscy dostawcy;
  V3 zastepuje V2 z 2026-11-04 - w domowych dlugosc i przekroj przeszly przed uziemienie): domowe "Przedluzacz |
  ilosc gniazd | dlugosc | przekroj | uziemienie / Schuko / bez uziemienia | wylacznik | nazwa wlasna | producent", bebnowe "Przedluzacz bebnowy | ilosc gniazd | dlugosc |
  przekroj | rodzaj przewodu (material: guma / PVC) | hermetycznosc | uziemienie / Schuko | nazwa wlasna |
  producent" (uziemienie przy koncu - decyzja uzytkownika do V2, potwierdzona dla V3 2026-11-05 - arkusz V3 nie ma tej kolumny, bez niego B15 i B16 Bemko mialy te same nazwy). Uziemienie w domowych to jedna wartosc (Schuko zastepuje "z uziemieniem"); USB idzie za gniazdami
  ("3 gn. + 2xUSB A+C"), kolor (Kobi) przed nazwa wlasna; kody Kobi "3gn/1,5m/Zu+W" rozpisane slowami, nazwa
  wlasna = seria Kobi (Linea, Linea Pro XC, Connecto...). "Przedluzacz 3 gn. 1,5m 1,5mm2 Schuko z wylacznikiem
  Bemko", "Przedluzacz bebnowy 4 gn. 15m 3x1,5mm2 Schuko Bemko", "Przedluzacz 3 gn. 5m z uziemieniem z
  wylacznikiem czarny Linea Kobi" (kolejnosc cech = kolejnosc `features` w YAML; reguly `przedluzacz` Kobi, Bemko i GTV, `przedluzacz_bebnowy` Bemko i GTV).
  GTV od 2026-11-05: domowe (PF / PS), ogrodowe, bebnowe i przedluzacze Hogert ("w izolacji gumowej" = "guma");
  biurkowe, wpuszczane i katowe MINI to inny rodzaj bez schematu. Uziemienia nie zgadujemy z liczby zyl.
  Opis akcesoriow / typow bez schematu GTV (`series_keep_units`, `series_keep_plus`, `series_balance_parens`):
  jednostka za liczba zostaje ("GLAX 12m"), "+" w skladzie zestawu tez, limit slow nie tnie w nawiasie.
  GTV runda 3 (2026-11-05): katalog `archived_input_files/input_2026-11-10/EkonomNazwy/GTV` (kata + 490 nowych kartotek kart). Znaki spoza cp1250
  przychodza z Ekonoma jako "?": "??" = smieci (usuwane), "?150 mm" = srednica, "mm?" = mm2, "cm?" = cm3, "PF?" = PF>;
  "A60 820lm" to 820 lm (liczba sklejona z litera nie zaczyna lumenow). Nowe grupy powtorzen to glownie identyczne
  nazwy w Ekonomie (Hogert -1/-2, profile GLAX z P na koncu) - do decyzji uzytkownika.
  Elgotech (od 2026-11-10, runda 1): katalog `input/Nazwy Ekonom/Elgotech` ("katalog" w ukladzie Ekonoma + "karto" =
  130 nowych kartotek), `--input <katalog>`. Zrodlo danych: katalog PDF ze strony (elgotech.pl ma tylko PDF) -
  `scripts/extract_elgotech_catalog.py` -> `cache/elgotech/katalog_wiersze.json` (`catalog_rows`, 609 z 1254 kodow):
  przekroj i kolor, ktorych krotka nazwa nie podaje; kolor sprzeczny z nazwa = DO_SPRAWDZENIA (`catalog_color_skip`
  dla bledow tabeli, R-20BS-24 jest czarno-czerwony jak na zdjeciu). Z dzialow katalogu: PK-10 / PS-10-5 =
  "Przedluzacz ogrodowy" pomaranczowy, PZB-1111-xxG = beben metalowy, RT-30CA trojkatny, R-x0CA plaski, model
  przeciwprzepieciowych (PSF / PSF2 / PSF3 / PSF4 / MPSFU / MPSF2U - rozna ochrona) jako nazwa wlasna. Prawie same przedluzacze - schemat V3 jak Kobi / Bemko / GTV:
  "Przedluzacz 5 gn. 2m z uziemieniem z wylacznikiem czarny Economic Elgotech", "Przedluzacz bebnowy maly 4 gn. 20m
  3x1mm2 PVC z uziemieniem Elgotech". Kod producenta stoi w nazwie (na poczatku albo w srodku) i jest wycinany
  (`strip_code`). Zapis Elgotech: z/u = z uziemieniem, b/u = bez uziemienia ("BEZ UZIEM" zawiera "Z UZIEM" - stad
  `\bZ UZIEM`), 2P+Z = z uziemieniem, "unischuko" to wtyczka (nie gniazdo Schuko). Nazwa wlasna: EPS = Economic, PSN =
  Classic (z kodu, tak podaja dlugie nazwy). Kolor tylko w koncowce kodu (`code_colors`, uwaga w raporcie): -2 czarny,
  -3 brazowy, -7 zielony, -9 szary, -5 pomaranczowy (jak OSLONA-5) - ale tylko po co najmniej jednym czlonie modelu,
  bo "GN-2" / "GO-3" to liczba gniazd. Bebny: Y = PVC, G = guma, BQ = PUR (H07BQ-F), PZB1 = maly, SPZD = duzy;
  przewod H07RN-F obok H05RR-F tego samego bebna jest jedyna roznica, wiec idzie do nazwy ("guma H07RN-F").
  Silowe SP5-16 / SP5-32 = 16A / 32A z kodu. Kod domowych <gniazda>[0 = z uziemieniem]<dlugosc> ("PSN-407-1,5" = 4 gn.,
  7m) i dlugosc bebna w kodzie sa sprawdzane z nazwa (`code_checks`) - 17 kartotek, w ktorych kod przeczy nazwie,
  ma DO_SPRAWDZENIA (nazwy nie poprawiamy) - to glownie dlugie opisy w Ekonomie skopiowane z sasiedniego kodu
  (PSN-307W opisany jak 306W), stad tez 14 grup powtorzen. Wynik v2: 1221 OK / 14 WARNING / 19 DO_SPRAWDZENIA.
  Plastrol (od 2026-11-10, runda 1): `input/Nazwy Ekonom/Plastrol` - "karto" i "katalog" sa identyczne (78
  kartotek), wiec przebieg na jednym pliku (`--input <plik karto>`; katalog dalby kazdy produkt dwa razy). Bez
  zrodla danych poza nazwa. Przedluzacze wg V3 jak Elgotech; bebny z linia (Professional / Heavy / Hand / Extreme /
  Eco Line) i przewodem (H05RR-F = guma, H07BQ-F = PUR, H05VV-F = PVC). X1 / X3 = "Przedluzacz budowlany" 1 / 3 gn.
  (karty TME), "X3" zostaje jako nazwa wlasna; "2p" bez "+Z" = bez uziemienia. KEL to marka Plastrolu - produkty
  z "KEL" w nazwie maja marke KEL zamiast Plastrol. Kolor przewodu przedluzaczy X ("przewod kolor czarny") to nie
  kolor wtyczki i gniazda. Wynik: 76 OK / 2 WARNING (W-99986 i W-97615 maja w Ekonomie identyczna nazwe "WT-16UP").
  PCE (od 2026-11-10, runda 1): `input/Nazwy Ekonom/PCE` - "karto" to same uslugi (KGO, marketing, samochody), produkty
  sa tylko w "katalog" (577, `--input <plik katalog>`). Bez zrodla danych poza nazwa. Zlacza CEE: prad, bieguny skrotem
  branzowym (2P+Z = 3P, 3P+Z = 4P, 3P+Z+N = 5P, 5P+Z+N = 7P), napiecie (24V AC / DC), IP (ostatnie w nazwie), seria na
  koncu: "Wtyczka przenosna 16A 5P 400V IP44 Shark TT PCE". Rozdzielnice budowlane: wyjscia (32/5, 2x16/5, Nx230V -
  "GS" = gniazdo 230V, "16/5 GS" = jedno, nie piec), zabezpieczenia (z RCD 40A / z RCBO / z wyl. nadpradowymi / bez
  zabezpieczen), rozlacznik, wtyczka zasilajaca przenosnych, nazwa typu: "Rozdzielnica stacjonarna 32/5 2x230V z RCD 40A
  IP44 Adamow PCE"; "Wejscie 32/5" / "wtyczka 16/5" to zasilanie, nie wyjscie. Bebny wg V3, trojfazowe z wyjsciami jak
  rozdzielnice. Jednofazowe 10/16A: "z zab. bocznym VDE" = Schuko, "uni-schuko" = Unischuko, "z bolcem" = z uziemieniem.
  Bledy opisow w Ekonomie poprawione wg zrodel (przeglad 2026-11-10; wartosc z cechy `code_match`, uwaga ze zrodlem i
  WARNING z `code_notes`): wtyczki tablicowe 613 / 623 / 6132 / 6232 = 230V (Ekonom 400V przy 2P+Z 6h), 75132-6 = IP67
  (pcelectric.at), 8615-6 = 7715-6 z dwoma wprowadzeniami przewodu (certyfikat OVE), puszki Light -171 = 155x179x100
  (Ekonom skopiowal -161; Onninen / TME), EX-088 = wymiar EP-088. Star 9M / 12M rozni liczba modulow ("9 mod.").
  Zrodla PCE: pce.pl (WooCommerce, karta z parametrami, "Pliki BMEcat"; Store API wylaczone, `?s=` prowadzi do serii),
  pcelectric.at (`/shop/en/?plugin=dynamicsearch&query=<kod>`, szybko daje 429), Onninen / TME po kodzie. NAJLEPSZE
  ZRODLO: baza XLS PCE (pce.pl > Info & media > Pliki BMEcat > "Baza produktowa w formacie xls" = etim-mapper.com, POST
  formularza -> zip; 10683 produkty, nazwa PL, opis, arkusze klas ETIM z IP, pradem, biegunami, wymiarami, typem gniazd),
  zapis lokalny `cache/pce/`. Kontrola 2026-11-10: 470 z 577 kodow w bazie (bez Argenta, Combi / Oro, ppoz natynkowych);
  poprawione tylko to, co stoi w nazwie / opisie producenta: PV = IP65 (IP54 to naglowek grupy w Ekonomie), lampy - IP z
  bazy (tez naglowek), gniazda z rozlacznikiem = IP44/54, Taurus2 25711 = z klapka / 25712 = z zatyczka (Ekonom odwrotnie)
  i IP54, 0351-6 = IP44, Kolo 9769215W i XREEL 92503H48153 = IP54. IP z samego ETIM (rozdzielnice budowlane IP54,
  ppoz podtynkowe IP44, Light IP65 - pole ETIM, opis mowi IP66/67) i wymiary zewnetrzne ETIM puszek zostaja jak w
  Ekonomie - raport `reports/Ekonom-nazwy-PCE_kontrola-z-baza-PCE_2026-11-10.xlsx`. GS w bebnach = CEE 7/5 typ E
  (z uziemieniem), w Ecoline typ F. Wynik: 537 OK / 40 WARNING (poprawki z uwaga); jedyne powtorzenie 90EKO006 /
  90EKO010 - na pce.pl identyczne parametry.
  Pawbol (od 2026-11-10, runda 1): `input/Nazwy Ekonom/Pawbol` ("katalog" 2120 + "karto" 288 nowych), `--input <katalog>`.
  Zrodlo nazwy: cennik Pawbol 2026 (PDF ze strefy klienta, 2031 kodow) i karty pawbol.pl (REST API WordPressa, typ
  `produkt`, kod w komentarzu tresci `<!-- H.7015B -25 -->`) - `scripts/fetch_pawbol_web.py` -> `cache/pawbol_web/`,
  loader `params_web_loader: pawbol`, `prefer_feed_name: always` (opis producenta zastepuje nazwe z Ekonoma,
  `feed_name_cleanup` czysci go tymi samymi regulami). "Produkt wycofany" w cenniku i kody spoza obu zrodel (ok. 630,
  glownie rozdzielnice R-BOX z zabezpieczeniami i stare oprawy BZ) = dane tylko z nazwy w Ekonomie. Klucz feedu to
  `Kod producenta` (`_feed_key`), bo czlon przed ukosnikiem mylil warianty (E.4150/N, E.4150/SZ). Wzor jak PCE: rzeczownik
  + parametry + kolor + IP, nazwa wlasna serii (R-BOX 240 4S, MAX-BOX 16S, S-BOX, PK60 Lux, ProDuo, Beta, Pantera...) i
  Pawbol na koncu, za marka tylko "op. N szt.". Rozdzielnice: wyjscia (32/5, 2x16/5, Nx230V; 250V z nazw = 230V),
  "z wył. nadprądowymi i RCD 40A" (M.01 / M.02-4/40/0,03), przelacznik L-0-P / rozlacznik 0-1, termik, "z rączką"
  (R w modelu 240R, cennik "+ rączka"; "b/R" = bez raczki), "z 2 podstawami" (-PS) / "z podstawą" (-PS_1). Bebny wg V3,
  trojfazowe = "Przedłużacz bębnowy trójfazowy" z wyjsciami jak rozdzielnica. Oprawy: E27 = "Plafon" bez mocy (moc to
  limit zarowki), LED = "Plafon LED" z moca, OWAL = "Oprawa kanałowa" (Owal 60 / 100 to rozmiar); BN = 4000K (karty EPREL
  pawbol.pl), BZ / BC = "barwa zimna / ciepła" (K nieznane), lm brak w zrodlach - bez DO_SPRAWDZENIA. Puszki "do pustych
  ścian" (K / PG / G/K) to inny rzeczownik niz podtynkowe. Osprzet BETA: uklad z KODU (G = gniazdo, GG / GGG, W =
  jednobiegunowy, W2 = swiecznikowy, WS / WK / WD), bo cennik 2026 opisuje zestawy GGW / W2W2 jak pojedyncze; kolor
  starych kartotek z kodu (D.31 bialy, 32 brazowy, 36 szary, 37 czarny). Bledy cennika poprawione z kodu / strony z
  uwaga (`code_notes`, `code_color_fix`): E.4132 LZ 10/Ż = zolta, E.4316/54SZ = sztyftowa, B.1923 = wtyczka przenosna,
  E.4096/125 = 1x10-35mm2. Cechy tylko z nazwy w Ekonomie: atrybut `features_raw` (maszty R.8312 / R.8313 "podstawa
  15kg / 40kg"). Akcesoria bez schematu (odgromowe, drobne): `noun_from: rest_head` + `rest` (opis producenta bez
  przecinkow i bez 230V / IK / "samogasnąca"). Wynik: 2367 OK / 41 WARNING; 15 grup powtorzen - pary stara / nowa
  kartoteka o identycznym opisie u producenta (przewody D.35xx i D.350x-N, OWAL D.3152 / D.31524, zlacza R.8006 /
  R.8060, S-BOX 036B dwa razy, C.1524-IT / -ITM) - do decyzji uzytkownika.
  F-Elektro (od 2026-11-10, runda 1): `input/Nazwy Ekonom/F Electro` ("katalog" 2925 + "karto" 21 nowych), `--input
  <katalog>`, klucz dostawcy `felektro` (plik wynikowy "F-Elektro" z `label`). Wlasnej strony nie ma (f-elektro.pl nie
  istnieje); nazwy w Ekonomie sa opisowe po polsku i sa zrodlem. Onninen ma karty wszystkich produktow marki, a adres
  karty to tytul producenta z kodem: `scripts/fetch_felektro_onninen.py --sitemap` (mapa strony, robots jej nie blokuje)
  -> `cache/onninen_felektro/adresy.json`, `--codes-from-report <kontrola>` -> `karty.json` (kody z arkusza
  Powtorzenia); loader `params_web_loader: onninen`. Tytul daje to, co Ekonom zgubil: wersje IP65 rozdzielnic (te same
  nazwy bez " - gniazda": LAGO F3.031x / F3.178x / F3.183x, ULISSE 74185 / 74186, FEMO 74187, LEGATO 75101 / 75102,
  TRIBOX F3.058x) i przedluzaczy 92978, przelacznik RB-300 (Ekonom 3x ta sama nazwa). Opis karty Onninen jest
  generowany (do LAGO 11M z gniazdami CEE dopisuje "dwa gniazda 250V") - z niego tylko IP, tylko dla zlaczy CEE
  (`ip_card_desc`, pary IP44 / IP67, WARNING) i tylko, gdy opis podaje jedno IP; w rozdzielnicach opis podaje IP
  obudowy takze przy gniazdach IP44. Wzor jak PCE / Pawbol: "Rozdzielnica 16/5 2x230V z wył. nadprądowymi i RCD 40A
  R-240 8M F-Elektro", "Rozdzielnica 2x32/5 2x230V z przełącznikiem L-0-P 32A z wtyczką 32/5 na przewodzie 5x4mm2 1,5m
  R-240 F-Elektro" (seria z liczba modulow na koncu, "(32A-Z)" = przelacznik zamykany, "wtyk odbiornikowy" = zasilanie
  w obudowie, "2x250V" = 2x230V). Bebny wg V3 (OW = guma, OPD = guma H07RN-F, OWY = PVC; uziemienia nie zgadujemy):
  "Przedłużacz bębnowy trójfazowy 2x32/5 1x230V 40m 5x4mm2 guma H07RN-F MAKARA-460 F-Elektro". Osprzet (puszki, LZ,
  bloki, szafy, dlawice...) = rzeczownik + reszta opisu z Ekonoma, "op. N szt." za marka. "/FE/" i "/VC/" to znaczniki
  marki (VC: F7.0197-F7.0201 - marka do potwierdzenia, uwaga w raporcie). Wynik: 2873 OK / 73 WARNING; 16 grup
  powtorzen - pary zlaczy CEE bez IP na karcie (cena i waga jak IP44 / IP67), F7.0102 / F7.0201 (/VC/), ULISSE
  74184.11 / .12, szyna F2.0112 / F2.0329 - do decyzji uzytkownika.
  Przeglad nazw wlasnych Kobi / GTV / Kanlux (2026-11-16, "czy nie zgubiles nazw wlasnych"): porownanie slow
  "Stara nazwa" z "Nowa nazwa" w raportach kontrolnych pokazalo, ze Kobi gubil serie we wszystkich typach
  parametrycznych (atrybut `series` mial tylko `dekor`) - teraz kazdy typ oprawy / zrodla Kobi ma `series` (za marka
  przez `series_after_brand`: "... Kobi Premium Nivera Square", "... Kobi Premium Delyo MHN"), porzadki serii tylko w
  tych typach (`series_drop: [*kobi_param_series_drop]` - rodzaj z nazwy feedu, kat 36ST, UGR, CRI, DALI; w `dekor`
  seria niesie tez DALI), `series_drop_skip` zostawia SQUARE / ROUND w typach bez cechy ksztaltu, `series_dedupe` usuwa
  slowa serii obecne juz w nazwie. Akcesoria z nazwa z feedu biora system z nazwy w Ekonomie (`features_sys_raw`:
  "Szynoprzewod ... Kobi Nextrack NT", Nelio), zasilacze Lifud kod LF-. GTV: linia zarowek / tub jawna lista
  (`features_line`: Spin, Glow, Grid, Ceramic, Slim, Wave, Vintage, Innovo, SHL), model z 2. czlonu, gdy 1. to sam opis
  (`series_next_segment`: Tiago, Escada), producent zasilacza tuz za marka (`tail_attrs`: "GTV Boke"), zakres napiecia
  "3-54V" i sterowanie "1-10V" nie ucinaja serii, wiszacy myslnik ("Osl SLIM-") wypada. Kanlux trzyma serie w kodzie
  modelu; doszly linie tub GLASSv2-v5 / IQ-LED EX i wylacznikow IDEAL TS (`features_line_raw`). `features_line*`,
  `features_sys*` i `features_maker` nie dopisuja sie, gdy nazwa juz je zawiera. Wejscie: archiwum
  `archived_input_files/input_2026-11-10/EkonomNazwy/<Dostawca>` + feed Kobi i Parametry Kanlux ze stasha `4ad6dd9`
  (przed zmianami przebieg odtwarzal poprzednie wyniki 1:1). Wynik: Kobi 643, GTV 64, Kanlux 23 zmienionych nazw,
  statusy i powtorzenia bez zmian; zestawienie `reports/Ekonom-nazwy-wlasne_przeglad_Kobi-Kanlux-GTV_2026-11-16.xlsx`
  (stara / poprzednia / nowa nazwa, arkusz "Nie zmienione - do decyzji"). Przy kazdym dostawcy warto zrobic ta sama
  kontrole: slowa WIELKIMI ze starej nazwy, ktorych nie ma w nowej, posortowane po liczbie wystapien.
  Kobi: nazwa wlasna PRZED marka (decyzja uzytkownika 2026-11-17, uchyla "seria za marka" powyzej):
  `series_before_brand: true` stawia serie z typow parametrycznych tuz przed marka, a typy z jawna pozycja marki
  (zarowki, zestaw szynowy, akcesoria z systemem) maja w `attrs` `series, brand` / `features_sys_raw, brand`; marka
  zamyka nazwe (za nia tylko "op. N szt."): "Taśma LED COB 320LED/m 50m 6500K 12V IP20 Tramo Kobi Premium",
  "Oprawa uliczna LED 50W 4000K z czujnikiem Cyoto Kobi Premium", "Szynoprzewod ... czarny Nextrack NT Kobi",
  protokol smart tez przed marka ("... RGB A60 WiFi Kobi"). Zmiana 1161 nazw, statusy i powtorzenia bez zmian;
  inni dostawcy (Bemko, GTV, Kanlux, Ledvance) zostaja z seria / linia za marka.
  Przeglad tego samego dnia: zestawy HERMETIC = typ `zestaw_hermetyczny` (atrybut `set_tubes` "z tubami LED 2x120cm",
  `power_per_tube` "2x18W"); cecha tylko z nazwy w Ekonomie dopisywana do nazwy z feedu przez `feed_name_from_raw`
  (lacznik "ZEWNĘTRZN" -> "T zewnętrzny prawy"); Windstar (grzejniki, wentylatory, klimatory) = typ `windstar`
  (rodzaj, moc, kolor, nazwa wlasna przed marka); "barwa NEUTRALNA BIAŁA" w Ekonomie to barwa swiatla, nie seria
  ani kolor obudowy.
  Test: `tests/test_build_ekonom_names.py`.

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
