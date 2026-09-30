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
