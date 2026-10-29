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

## Polaczenie ze sklepem Woo (AlkanInator Bridge)

Sklep ma wtyczke `alkan-inator-bridge` (repo AlkanLocal,
`wp-content/plugins/alkan-inator-bridge`) z REST API `alkan-inator/v1`. Klient:
`src/woo_api/` + `py src/woo_api_cli.py`. Polaczenie: `src/woo_api/config.json`
(ignorowany przez git; `url`, `token`, `verify_ssl` - dla lokalnego Laragona
`false`) albo zmienne `ALKAN_WOO_URL` / `ALKAN_WOO_TOKEN`. Token generuje sie w
WooCommerce > AlkanInator API (pokazywany raz). Na razie dziala tylko na
AlkanLocal (`https://alkanlocal.test:8443`); na produkcji wtyczki jeszcze nie ma,
tam nadal obowiazuje WP All Import (`dictionaries/woo_all_export_import_workflow.yaml`).

- `knowledge` - wiedza o sklepie. Nadpisuje
  `dictionaries/woo_attribute_terms_snapshot.tsv` i
  `dictionaries/woo_product_categories_snapshot.tsv` (ten sam uklad co
  `scripts/dump_woo_*.py`, nazwy z odkodowanymi encjami, np. `F&F`), tworzy
  `dictionaries/woo_shop_terms_snapshot.tsv` (klasy wysylkowe, "Wysylka w ciagu",
  tagi) i pelny JSON `cache/woo_api/knowledge.json` (kolumny, ktore sa
  zapisywalne, formaty, jednostki). Odswiezaj przed praca na wartosciach
  atrybutow zamiast zrzutow z MySQL.
- `pull --output plik.xlsx` - produkty w ukladzie WP All Export (arkusz
  `Worksheet`, te same naglowki: `Kategorie produktów`, `Product Tags`,
  `Up-Sells`, `Atrybut Produktu: <Etykieta>` z wielkimi literami jak w WP All
  Export). Filtry: `--category` (slug, ID albo sciezka `A>B`, z podkategoriami),
  `--sku`, `--sku-file`, `--ids`, `--status`, `--search`, `--modified-after`,
  `--all`. Dodatkowe kolumny tylko do odczytu: ceny, stan, EAN,
  `Post Modified Date`, `Permalink`, `Product Attributes`.
- `push --input plik.xlsx` - ZAWSZE najpierw dry-run calego pliku i raport
  `reports/woo_api/*_dry-run.xlsx` (arkusze Zmiany / Problemy / Nowe terminy).
  `--apply` zapisuje tylko produkty ze zmianami; bledy w dry-run wstrzymuja zapis
  (chyba ze `--skip-errors`). `--columns "A,B"` ogranicza zapis do wskazanych
  kolumn - uzywaj go, gdy zadanie dotyczy kilku kolumn (odpowiednik zasady
  "modyfikuj tylko jawnie wskazane kolumny"). Nowe tagi i wartosci atrybutow
  powstaja tylko z `--allow-new-terms`. Nowe zdjecia z `Image URL` dry-run
  sprawdza (JPG 1000x1000, inaczej blad).
- `changesets` i `revert <zapis> [--apply]` - kazdy zapis ma identyfikator
  (np. `ai-20261021-153259-plik`) i dziennik starych wartosci w sklepie.
  Cofniecie pomija pola zmienione w sklepie po tym zapisie. Cofnac mozna tez
  przyciskiem w WooCommerce > AlkanInator API.

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
  `dictionaries/kontakt_simon_shop_structure.yaml` (137 globalnych atrybutow
  Woo z AlkanLocal 2026-09-17) plus `approved_extra_attributes`. Cechy z
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
  `dictionaries/woo_attribute_terms_snapshot.tsv` (zrzut z bazy AlkanLocal,
  odswiezany `py scripts/dump_woo_attribute_terms.py`; zawiera tez
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
  `dictionaries/alkanlocal_wpclv_groups.tsv`. Robia to
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
  reguly typow `dictionaries/ekonom_name_rules.yaml`, wynik: CSV w formacie
  Ekonoma w `output/` + skoroszyt kontrolny w `reports/` (arkusze "Nazwy",
  "Powtorzenia", "Roznice w powtorzeniach"). **Powtorzenia nazw zglasza sie
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
  seria (ZOE != SIGARO), wiec lm nie przenosimy. Kanlux (od 2026-10-14):
  parametry z plikow `Parametry*.xlsx` z kanlux.pl (`--params-dir`, format dlugi
  Kod Kanlux / Nazwa atrybutu / Wartosc, dopasowanie po `Kod towaru` 37246/KAN) sa
  zrodlem prawdy i nadpisuja nazwe (kod `WW` u zarowek bywa 2700K, nie 3000K).
  Nazwa Kanlux = kod modelu + polski typ - rzeczownik bierze sie z typu
  ("Oprawa liniowa LED"), seria to kod modelu bez parametrow (`AL-SL-MAT-NT`).
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
  wylaczniki KMB/KRD i styczniki KCP dekodowane z kodu (B16 1P 10kA). GTV (od
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
  `Nazwa towaru` jako `<40 znakow> | <reszta>` - `read_csv` je rozdziela, a wynik idzie ZAWSZE w ukladzie Ekonoma
  (`LP;Kod towaru;Nazwa towaru;Nazwa dodatkowa;Kod producenta;JM;Cena netto;Waluta`, jak `Ledvance.csv`; `JM` z `J.M.`,
  `Cena netto` z `Cena C3 - Cena Katalogowa`, `Waluta` = PLN). Uklad zestawienia z ` | ` NIE jest formatem do Ekonoma
  (poprawka uzytkownika 2026-10-27). Nowa partia idzie przez przebieg laczny z baza v20 (`Ledvance.csv` + zestawienie), zeby wyrozniki
  liczyly sie tez wzgledem produktow juz nazwanych; do wyniku ida tylko nowe pozycje, a nazwy w v20 zostaja. Przed
  przebiegiem pobierz karty GTIN nowych EAN (`fetch_ledvance_web.py --ean-file`). Stare kartoteki Osram z kodem na
  poczatku maja wlasne typy i porzadki: zarowki samochodowe (zawsze Osram; NBU / NL = Night Breaker Unlimited / Laser),
  Linestra = "Zarowka liniowa", HALOSTAR / HALOLINE / HALOLUX / "ZAR. HAL." = halogen, HCI / HMI = metalohalogenkowa,
  "L15/76" = L 15W/76 (Natura, Fluora bez K), "DS7/41-827" = DULUX S, "ST8A-0.9M" = tuba T8 90 cm, lm "1K2LM" / "10KLM",
  marka Osram tez przy koncowce EAN ("OSRAM 7426"), DUOBOX = op. 2 szt.
  Swietlowki liniowe Ledvance to "Swietlowka", nie "Tuba fluorescencyjna" (decyzja uzytkownika 2026-10-26); "Tuba LED"
  i oprawy "pod tuby" bez zmian.
  Test: `tests/test_build_ekonom_names.py`.

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
