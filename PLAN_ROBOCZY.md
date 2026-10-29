# PLAN ROBOCZY - AlkanInator

Ten plik jest robocza checklista wdrozenia planu z `Plan.txt`.
Aktualizujemy go po kazdym zakonczonym kroku, zeby bylo jasne:

- co jest zrobione,
- co jest w toku,
- co zostalo do zrobienia,
- jakie decyzje techniczne juz zapadly,
- jakie problemy trzeba jeszcze rozstrzygnac.

Legenda statusow:

- `[ ]` do zrobienia
- `[~]` w toku
- `[x]` zrobione
- `[!]` blokada / wymaga decyzji

---

## 0. Zasady prowadzenia projektu

- [x] Utworzyc osobny plik roboczy do sledzenia planu.
- [x] Nie nadpisywac oryginalnych plikow produktowych.
- [x] Kazda zmiana danych produktowych ma miec slad: stara wartosc, nowa wartosc, zrodlo, confidence, status.
- [x] Codex nie przetwarza recznie calego pliku produktow.
- [x] Codex pracuje na probkach, raportach, configach i bledach.
- [x] Program lokalny masowo przetwarza pliki CSV/XLSX.
- [x] Internet dodac dopiero po stabilnym MVP lokalnym.

---

## 1. Przygotowanie struktury projektu

- [x] Utworzyc katalog `archived_input_files/`.
- [x] Utworzyc katalog `output/`.
- [x] Utworzyc katalog `configs/`.
- [x] Utworzyc katalog `configs/categories/`.
- [x] Utworzyc katalog `configs/producers/`.
- [x] Utworzyc katalog `dictionaries/`.
- [x] Utworzyc katalog `cache/`.
- [x] Utworzyc katalog `reports/`.
- [x] Utworzyc katalog `src/`.
- [x] Utworzyc katalog `prompts/`.
- [x] Utworzyc `README.md`.
- [x] Utworzyc `requirements.txt`.
- [x] Ustalic minimalne zaleznosci Pythona.

---

## 2. MVP bez internetu

Cel MVP:

1. Wczytac plik CSV/XLSX.
2. Wyciagnac atrybuty z tytulow.
3. Wygenerowac nowe tytuly z configu.
4. Zrobic raport roznic.

Zakres:

- [x] Import CSV.
- [x] Import XLSX.
- [x] Wybór arkusza XLSX przez `--sheet`.
- [x] Automatyczne wykrywanie podstawowych kolumn.
- [x] Analiza kolumn.
- [x] Ekstrakcja atrybutow z tytulow.
- [x] Normalizacja podstawowych wartosci.
- [x] Generowanie nowych tytulow SEO.
- [x] Raport zmian.
- [x] Eksport XLSX.
- [x] Eksport CSV.
- [x] Komenda pipeline: `py src/run_pipeline.py --input archived_input_files/products.xlsx --category "Oprawy sufitowe"`.

Pliki docelowe MVP:

- [x] `src/run_pipeline.py`
- [x] `src/analyze_products.py`
- [x] `src/extract_attributes.py`
- [x] `src/optimize_titles.py`
- [x] `src/export_results.py`
- [x] `src/utils.py`
- [x] `configs/global.yaml`
- [x] `configs/categories/oprawy-sufitowe.yaml`
- [x] `reports/quality_report.html`
- [x] `reports/missing_attributes.csv`
- [x] `reports/changes_report.xlsx`
- [x] `output/products_optimized.xlsx`
- [x] `output/products_optimized.csv`

---

## 3. Modul: analiza pliku produktow

Narzędzie docelowe: `analyze-products`

- [x] Wczytywac plik wejściowy CSV/XLSX.
- [x] Wypisywac liste kolumn.
- [x] Liczyc produkty.
- [x] Liczyc kategorie.
- [x] Liczyc producentow.
- [x] Wykrywac puste pola.
- [x] Wykrywac tytuly za krotkie.
- [x] Wykrywac tytuly za dlugie.
- [ ] Wykrywac tytuly bez producenta.
- [ ] Wykrywac tytuly bez waznych parametrow.
- [x] Sprawdzac poprawnosc EAN.
- [x] Wykrywac duplikaty SKU.
- [x] Wykrywac duplikaty EAN.
- [ ] Wykrywac niespojne wartosci atrybutow.
- [x] Generowac `reports/quality_report.html`.
- [x] Generowac `reports/missing_attributes.csv`.
- [x] Generowac `reports/category_summary.csv`.

---

## 4. Modul: ekstrakcja atrybutow

Narzędzie docelowe: `extract-attributes`

- [x] Ekstrahowac moc, np. `24W`.
- [x] Ekstrahowac napiecie, np. `230V`.
- [x] Ekstrahowac barwe swiatla, np. `3000K`, `4000K`, `6500K`.
- [x] Ekstrahowac strumien swietlny, np. `2000lm`.
- [x] Ekstrahowac IP, np. `IP20`, `IP44`, `IP65`.
- [x] Ekstrahowac IK, np. `IK08`.
- [x] Ekstrahowac kat swiecenia, np. `120°`.
- [x] Ekstrahowac kolor.
- [x] Ekstrahowac producenta.
- [x] Ekstrahowac serie.
- [x] Ekstrahowac model.
- [ ] Ekstrahowac material.
- [ ] Ekstrahowac sposob montazu.
- [ ] Ekstrahowac wymiary.
- [ ] Ekstrahowac srednice.
- [ ] Ekstrahowac typ przewodu.
- [ ] Ekstrahowac przekroj przewodu.
- [ ] Ekstrahowac liczbe zyl.
- [x] Zapisac zrodlo atrybutu.
- [x] Zapisac confidence dla atrybutu.

---

## 5. Modul: normalizacja atrybutow

Narzędzie docelowe: `normalize-attributes`

- [x] Utworzyc `dictionaries/colors.yaml`.
- [ ] Utworzyc `dictionaries/power.yaml`.
- [ ] Utworzyc `dictionaries/ip_rating.yaml`.
- [ ] Utworzyc `dictionaries/color_temperature.yaml`.
- [ ] Utworzyc `dictionaries/mounting_types.yaml`.
- [ ] Utworzyc `dictionaries/cable_types.yaml`.
- [x] Utworzyc `dictionaries/producers.yaml`.
- [x] Normalizowac barwe swiatla do formatu np. `4000K`.
- [x] Normalizowac kolory.
- [x] Normalizowac moc.
- [x] Normalizowac IP.
- [x] Normalizowac producentow.
- [ ] Oznaczac wartosci nieznane jako wymagajace sprawdzenia.

---

## 6. Modul: config kategorii

Narzędzie docelowe: `suggest-category-config`

- [x] Obslugiwac osobne configi kategorii.
- [x] Zdefiniowac pola wymagane.
- [x] Zdefiniowac pola opcjonalne.
- [x] Zdefiniowac szablon tytulu SEO.
- [x] Zdefiniowac maksymalna dlugosc tytulu.
- [x] Zdefiniowac rekomendowane filtry.
- [x] Zdefiniowac slowa zabronione w tytulach.
- [x] Zdefiniowac reguly normalizacji.
- [x] Przygotowac config `oprawy-sufitowe.yaml`.
- [ ] Przygotowac config `zrodla-swiatla.yaml`.
- [ ] Przygotowac config `przewody.yaml`.
- [ ] Przygotowac config `osprzet-elektryczny.yaml`.
- [ ] Przygotowac config `rozdzielnice.yaml`.
- [ ] Przygotowac config `aparatura-modulowa.yaml`.

---

## 7. Modul: generowanie tytulow SEO

Narzędzie docelowe: `optimize-titles`

- [x] Generowac tytul wedlug `title_template`.
- [x] Pomijac puste pola w szablonie.
- [x] Pilnowac maksymalnej dlugosci tytulu.
- [x] Usuwac slowa zabronione.
- [x] Zachowac producenta, jesli jest dostepny.
- [x] Nie wymyslac parametrow bez zrodla.
- [x] Zapisywac `old_title`.
- [x] Zapisywac `new_title`.
- [x] Zapisywac `title_status`.
- [x] Zapisywac `title_length`.
- [x] Zapisywac `changed_fields`.
- [x] Zapisywac `warnings`.

---

## 8. Modul: raport roznic

Narzędzie docelowe: `generate-report`

- [x] Pokazywac stary tytul.
- [x] Pokazywac nowy tytul.
- [ ] Pokazywac stare atrybuty.
- [x] Pokazywac nowe atrybuty.
- [ ] Pokazywac dodane atrybuty.
- [ ] Pokazywac zmienione atrybuty.
- [ ] Pokazywac brakujace atrybuty.
- [x] Pokazywac zrodlo danych.
- [x] Pokazywac confidence.
- [x] Pokazywac status.
- [x] Pokazywac ostrzezenia.
- [x] Eksportowac raport do XLSX.
- [x] Eksportowac raport do CSV.

---

## 9. Modul: walidacja wynikow

Narzędzie docelowe: `validate-output`

- [ ] Sprawdzac, czy SKU nie zniknelo.
- [ ] Sprawdzac, czy EAN nie zostal przypadkowo zmieniony.
- [ ] Sprawdzac, czy producent nie zostal blednie nadpisany.
- [ ] Sprawdzac, czy tytul nie jest pusty.
- [ ] Sprawdzac, czy tytul nie jest za dlugi.
- [ ] Sprawdzac dziwne znaki w tytule.
- [ ] Sprawdzac wymagane atrybuty.
- [ ] Sprawdzac zgodnosc atrybutow ze slownikami.
- [ ] Oznaczac niski confidence jako `DO_SPRAWDZENIA`.
- [x] Sprawdzac, czy produkt pasuje do zadeklarowanej kategorii.
- [ ] Sprawdzac, czy kategoria nie zmienila sie bez powodu.
- [ ] Obslugiwac status `OK`.
- [ ] Obslugiwac status `WARNING`.
- [ ] Obslugiwac status `DO_SPRAWDZENIA`.
- [ ] Obslugiwac status `ERROR`.

---

## 10. Modul: rekomendacja filtrow

Narzędzie docelowe: `recommend-filters`

- [ ] Analizowac czestotliwosc wystepowania atrybutu.
- [ ] Analizowac liczbe unikalnych wartosci.
- [ ] Wykrywac balagan w wartosciach.
- [ ] Odrzucac techniczne smieci jako filtry.
- [ ] Odrzucac atrybuty prawie zawsze identyczne.
- [ ] Oznaczac dobre filtry.
- [ ] Oznaczac filtry opcjonalne.
- [ ] Oznaczac filtry niezalecane.
- [ ] Oznaczac atrybuty wymagajace czyszczenia.
- [ ] Generowac `reports/filter_recommendations.csv`.

---

## 11. Modul: internet i zrodla zewnetrzne

Narzędzie docelowe: `research-attributes`

Ten etap robimy dopiero po MVP.

- [ ] Wyszukiwac po EAN.
- [ ] Wyszukiwac po SKU.
- [ ] Wyszukiwac po producencie i modelu.
- [ ] Preferowac oficjalna strone producenta.
- [ ] Preferowac karty katalogowe PDF.
- [ ] Obslugiwac oficjalne pliki produktowe producenta.
- [ ] Obslugiwac zaufanych dystrybutorow.
- [ ] Traktowac inne sklepy jako zrodlo pomocnicze.
- [ ] Traktowac marketplace jako zrodlo pomocnicze.
- [ ] Zapisywac URL zrodla.
- [ ] Zapisywac date pobrania.
- [ ] Zapisywac confidence score.
- [ ] Zapisywac status `AUTO`.
- [ ] Zapisywac status `DO_SPRAWDZENIA`.
- [ ] Cache wynikow w `cache/web_sources.sqlite`.
- [ ] Cache wynikow w `cache/researched_products.json`.

---

## 12. Integracja z Codexem

- [x] Przygotowac instrukcje pracy dla Codexa w `README.md`.
- [x] Przygotowac prompt `prompts/category_config_prompt.md`.
- [x] Przygotowac prompt `prompts/title_rules_prompt.md`.
- [x] Przygotowac prompt `prompts/filter_recommendation_prompt.md`.
- [x] Przygotowac prompt `prompts/attribute_cleanup_prompt.md`.
- [x] Ustalic workflow: analiza -> config -> ekstrakcja -> tytuly -> walidacja -> raport.
- [x] Codex ma analizowac glownie raporty bledow i ostrzezen.

---

## 13. Kolejnosc sprintow

### Sprint 1 - fundament

- [x] Struktura projektu.
- [x] Import XLSX/CSV.
- [x] Analiza kolumn.
- [x] Raport brakow.
- [x] Wykrywanie SKU/EAN/producenta/kategorii.

### Sprint 2 - atrybuty

- [x] Ekstrakcja podstawowych parametrow z tytulow.
- [x] Slowniki wartosci.
- [x] Normalizacja kolorow.
- [x] Normalizacja mocy.
- [x] Normalizacja IP.
- [x] Normalizacja barwy swiatla.

### Sprint 3 - tytuly

- [x] Config kategorii.
- [x] Generator tytulow SEO.
- [x] Raport stary tytul vs nowy tytul.

### Sprint 4 - walidacja i eksport

- [ ] Walidacja wynikow.
- [ ] Statusy `OK`, `WARNING`, `ERROR`, `DO_SPRAWDZENIA`.
- [ ] Eksport pliku wynikowego.

### Sprint 5 - filtry

- [ ] Rekomendacja filtrow.
- [ ] Analiza jakosci atrybutow.
- [ ] Raport filtrow per kategoria.

### Sprint 6 - internet

- [ ] Research z internetu.
- [ ] Cache wynikow.
- [ ] Confidence score.
- [ ] Obsluga zrodel.

---

## 14. Aktualny status

Status ogolny: MVP lokalne dziala na pliku testowym.

Zrobione:

- [x] Odczytano `Plan.txt`.
- [x] Utworzono `PLAN_ROBOCZY.md`.
- [x] Utworzono strukture projektu.
- [x] Wdrozono MVP bez internetu.
- [x] Uruchomiono pipeline na `archived_input_files/sample_products.csv`.
- [x] Wygenerowano pliki w `output/` i `reports/`.
- [x] Dodano raport `suspicious_category_fit.csv` dla produktow podejrzanych w kategorii.
- [x] EAN `5905339372628` jest oznaczany jako `DO_SPRAWDZENIA`.
- [x] Uruchomiono pipeline na pelnym arkuszu `7. Wszystkie SKU (master)`.
- [x] Wynik pelnego testu: 499 produktow, 359 tytulow OK, 140 WARNING, 67 podejrzanych kategorii.
- [x] Walidacja kategorii sprawdza tytul zrodlowy pod akcesoria typu klosz/ramka/zasilacz/uchwyt.
- [x] Dodano `product_role`: rozdzielenie produktow glownych od akcesoriow.
- [x] Dodano osobne szablony tytulow dla akcesoriow: klosz, ramka, zasilacz, uchwyt.
- [x] Wynik klasyfikacji mastera: 452 produkty glowne, 47 akcesoriow.
- [x] Dopracowano SEO ramek: wymiary, ksztalt i kolor w tytule.
- [x] Dodano raport strategii SEO per typ produktu: `title_strategy_report.csv`.
- [x] Dodano pierwsze reguly tytulow per kategoria: panele, naswietlacze, high bay, hermetyczne, plafoniery.
- [x] Dodano ekstrakcje: gwint, srednica, material.
- [x] Dodano akcesoria: pilot, wspornik/uchwyt, soczewka.
- [x] Dodano akcesorium `Czujnik` dla przypadkow typu `MAH PRO SENSOR C`.
- [x] Dopracowano plafoniery drewniane i produkty na gwint E27.
- [x] Dopracowano nietypowe zapisy mocy: `LED-10-B`, `150WD`, `150WS`, `11,8WCCT`.
- [x] Dopracowano zakresy strumienia: `34000 / 25500 / 17000Lm` -> `17000-34000lm`.
- [x] Wynik po kolejnej iteracji: 445 tytulow OK, 54 WARNING.

W toku:

- [x] Test na probce 10 produktow z realnego pliku `archived_input_files/Alkan_Kanlux_pelne_rodziny.xlsx`.

Najblizszy krok:

- [ ] Dopracowac reguly dla 117 produktow glownych bez wykrytej mocy.
- [ ] Dopracowac osobne strategie SEO dla plafonier dekoracyjnych/drewnianych i nietypowych naswietlaczy.

Plan na 2026-10-15 - Nazwy do Ekonoma, Ledvance (przeglad WARNING):

- [x] Przeglad grup powtorzen (2026-10-15): `scripts/review_ekonom_duplicates.py`.
- [x] Runda 7 (decyzje 2026-10-15, cen nie uzywamy): marka z opakowania Osram/Bellalux (430 nazw),
      multipak `op. N szt.` (182), barwa z katalogu (5 poprawionych, 3 DO_SPRAWDZENIA - katalog
      opisuje inny produkt). Powtorzenia 734 -> 580 grup. Wynik
      `output/Ekonom-nazwy-Ledvance_2026-10-15_v2.csv`, przeglad reszty
      `reports/Ekonom-Ledvance_przeglad_powtorzen_2026-10-15_v2.xlsx`.
- [x] Runda 8 (2026-10-15): wyrozniki z nazwy katalogowej (`catalog_suffix`) - 149 nazw,
      powtorzenia 580 -> 529 grup. Wynik `output/Ekonom-nazwy-Ledvance_2026-10-15_v3.csv`,
      przeglad `reports/Ekonom-Ledvance_przeglad_powtorzen_2026-10-15_v3.xlsx`.
- [x] Runda 9 (2026-10-15, "tyle zeby sie wyroznialo, nie wiecej; czego brak - z sieci"): minimalny
      wyroznik (`distinguish`) z katalogu PDF (140), starej nazwy (263) i ledvance.com (42, GTIN
      zweryfikowany na kartach produktu). Poprawka: "Warm White" to nie kolor obudowy (ENDURA DG).
      Powtorzenia 529 -> 327 grup. Wynik `output/Ekonom-nazwy-Ledvance_2026-10-15_v4.csv`.
- [x] Runda 10 (2026-10-15): karty `ledvance.com/en/gtin/<EAN>` dla 739 produktow z powtorzen (735 znalezione):
      285 multipakow, 66 kartotek z nazwa innego produktu przebudowanych z karty, CRI 90, dlugosci, gwint,
      kilka rund wyroznikow. Powtorzenia 327 -> 81 grup (162 szt.). Wynik `..._2026-10-15_v5.csv`.
- [!] Zostaje 81 grup: 75 blizniakow (karta producenta identyczna - 57 calkiem, 18 rozni tylko karton
      zbiorczy) = ten sam produkt pod dwoma EAN; 5 grup sprzecznych z Ekonomem (TruSys DALI/bez DALI,
      Vintage B/P, HQL PRO); 1 bez karty. Decyzja uzytkownika: scalic w Ekonomie czy dopisac koncowke EAN.
- [ ] Przejrzec WARNING z `reports/Ekonom-nazwy-Ledvance_kontrola_2026-10-14.xlsx`: 895 pozycji,
      z tego 753 w grupach powtorzen (734 grupy, 1677 produktow) i 142 inne (typ bez schematu).
- [ ] Dla produktow z tymi samymi atrybutami (ta sama nowa nazwa) ustalic, czy to faktycznie
      zamienniki (np. nowsza wersja / inny EAN tego samego produktu), czy rozne produkty, ktorych
      nazwa nie odroznia. Najwiecej grup: "identyczna nazwa w Ekonomie" (392), "inne" (130),
      "linia produktu" (120), "ksztalt" (66), "opakowanie" (36), "kolor obudowy" (32), "optyka / kat" (30).
- [ ] Zrodla do porownania: EAN z katalogow PDF Ledvance (`input/Nazwy do ekonom/Ledvance/`),
      kod zamowieniowy z nazwy katalogowej, roznice w arkuszu "Powtorzenia".
- [ ] Wynik: lista zamiennikow (do oznaczenia / scalenia w Ekonomie - decyzja uzytkownika) i lista
      roznych produktow z atrybutem, ktory trzeba dodac do nazwy (decyduje uzytkownik, potem regula
      w `dictionaries/ekonom_name_rules.yaml` + test).
- [ ] Bemko (czeka na decyzje z 2026-10-14): 17 grup powtorzen, marka z rodzenstwa po kodzie,
      19 rozbieznosci lm nazwa vs karty katalogowe - `reports/Ekonom-nazwy-Bemko_kontrola_2026-10-14.xlsx`.
- [ ] Bemko runda 2 (2026-10-15): `_v2` z kartami bemko.pl (90/1325 - IP zablokowane przez bemko.pl).
      Do dogrania po odblokowaniu: `py scripts/fetch_bemko_web.py --input "input/Nazwy do ekonom/Bemko/Bemko.csv"`,
      potem przebudowa `_v3` i przeglad par powtorzen (RCBO GKL/GVL, K7/V7, PC/PM, LD/LDP, LDS/LDS2, SL03A/B).
- [x] Ledvance zamkniety na 2026-10-15 decyzja uzytkownika ("tyle moze zostac"): 81 grup blizniakow /
      sprzecznosci zostaje w `..._v5` bez dalszych zmian.
- [x] Ledvance 2026-10-19: nazwy bez producenta ("jest niepotrzebne"). Marka Osram/Ledvance/Bellalux
      zostaje tylko w 357 nazwach blizniakow (ta sama lampa pod inna marka, inny EAN) - decyzja uzytkownika,
      bez sztucznych wyroznikow. Powtorzenia bez zmian (81 grup). Wynik `output/Ekonom-nazwy-Ledvance_2026-10-19.csv`,
      kontrola `reports/Ekonom-nazwy-Ledvance_kontrola_2026-10-19.xlsx`.
- [x] Ledvance 2026-10-19 v2: kolejnosc atrybutow zarowek LED jak w starym narzedziu (gwint, W, lm, K, ...,
      linia na koncu). 3307 nazw zmienionych, inne typy bez zmian. Powtorzenia 81 -> 85: 4 pary PCR swieczka E14 7,5W
      (kartoteki "PCR CLASSIC A 19W E27" przebudowane z karty GTIN) roznily sie wczesniej tylko szykiem slowa PCR.
      Wynik `output/Ekonom-nazwy-Ledvance_2026-10-19_v2.csv`, szablony `reports/Ekonom-Ledvance_szablony_nazw_2026-10-19_v2.xlsx`.
- [x] Ledvance 2026-10-19 v4: swietlowki -> Tuba LED / Tuba fluorescencyjna, wpuszczane -> Downlight LED (ENDURA do
      dekor, Smart+ Spot GU10/R50 i "Oprawa LED CLASSIC/LAMPS/STAR" do zarowek), "Oprawa LED" rozbita na typy (0 zostalo):
      uliczna (SLCOMP/SL FLEX), plafon (SF DISC/CIRC, LED CLICK, SF BLKH), lampa stojaca, oprawa do zadaszen (CANOPY), scienna
      (WALLPACK), awaryjna / modul awaryjny (z moca; akcesoria EM Exit zostaja akcesoriami), Moduł LED (LED RING), oczyszczacz UVC.
      Poprawione srednice z kodu barwy ("fi830" w AR SPOT RD 6W 830). Powtorzenia 85 -> 94: 9 par to zdublowane kartoteki
      ("Oprawa LED LAMPS ..." = "Lampa LED ...", "Oprawa LED TUBE" = "Tuba LED TUBE"). Otwarte: 132 oprawy AR (architektoniczne:
      FLO/ING/WWASH/SPOT/CEIL) jako Downlight - do decyzji. Wynik `..._2026-10-19_v4.csv`, szablony `..._szablony_nazw_2026-10-19_v3.xlsx`.
- [x] Ledvance 2026-10-19 v5: poprawka "Value Value" / "Superior Superior" / "ECO Eco" (litera klasy V/S/E porownywana
      po rozwinieciu). 15 nazw; 5 par stary/nowy EAN tego samego produktu (4058075... / 4099854...) wraca do powtorzen (99 grup).
- [x] Ledvance 2026-10-19 v6: "jednak chcemy nazwe producenta" - marka wraca do wszystkich nazw (zawsze jedna:
      Ledvance 9562 / Osram 388 / Bellalux 131). Powtorzenia 99 -> 105: 6 par stara/nowa kartoteka (PCR CLASSIC A 13W
      E27 = swieczka 4,9W wg karty GTIN, VALUE CLASSIC 4,9W 6500K), wczesniej rozroznianych tylko marka u jednej z pary.
      Wynik `output/Ekonom-nazwy-Ledvance_2026-10-19_v6.csv`, szablony `reports/Ekonom-Ledvance_szablony_nazw_2026-10-19_v4.xlsx`.
- [x] Ledvance 2026-10-19 v7: linia w kazdym typie, gdy jest (zaraz po marce), takze ze skrotu V/P/S przed moca.
      2312 nazw zmienionych. Powtorzenia 105 -> 152: 53 nowe grupy to pary stara/nowa kartoteka tego samego produktu
      ("VALUE CLASSIC A 75 10W" = "CLASSIC A V 75 10W"), wczesniej rozdzielone tylko brakiem "Value" u jednej z pary.
      Wynik `output/Ekonom-nazwy-Ledvance_2026-10-19_v7.csv`, szablony `reports/Ekonom-Ledvance_szablony_nazw_2026-10-19_v5.xlsx`.
- [x] Ledvance 2026-10-19 v8: marka + linia na koncu we wszystkich typach (jak w zarowkach), wersja i wyroznik przed
      marka; "antybakteryjna" jako cecha przed marka; podwojne slowa obok siebie sklejane (`dedupe_adjacent`:
      "globe Globe 95", "SENSOR SENSOR"). Powtorzenia bez zmian (152). Wynik `..._2026-10-19_v8.csv`,
      szablony `reports/Ekonom-Ledvance_szablony_nazw_2026-10-19_v6.xlsx`.
- [x] Ledvance 2026-10-19 v9: srednica "220 mm" zamiast "fi220" (621 nazw). Wynik `..._2026-10-19_v9.csv`,
      szablony `..._szablony_nazw_2026-10-19_v7.xlsx`.

- [x] Ledvance 2026-10-19 v11 (uwagi z wieczora: "niektore linie dzialaja i sa przypisane, a niektore nie"). Wynik
      `output/Ekonom-nazwy-Ledvance_2026-10-19_v11.csv`, kontrola `reports/Ekonom-nazwy-Ledvance_kontrola_2026-10-19_v11.xlsx`,
      szablony `reports/Ekonom-Ledvance_szablony_nazw_2026-10-19_v8.xlsx`. 422 nazwy zmienione (275 tub, 35 liniowych,
      34 plafony, 40 downlightow/wpuszczanych, stateczniki/halogeny/HQI "Pro"); statusy i 152 grupy powtorzen bez zmian.
      Kobi/Kanlux/GTV/Bemko przebudowane kontrolnie - 0 zmian.
      - Linia z litery V/P/S takze przed dlugoscia ("EM S 1500", "UO P 1200", "HF P UO 1200", "EXT P HO54 1149mm", "10.5W830").
        Nowe linie: Entry, Pro, Advanced (ADV), Essential, PFM = Performance (tylko samodzielne slowo oprawy - w kodach
        zasilaczy "DR DS-PFM-20" i akcesoriow tasm "PFM -CSD" zostaje kod), "PARATHOM DIM PRO" = Parathom Pro.
        "Ultra Output" / UO dopisywane do klasy: "Ledvance Performance Ultra Output" (tuby i oprawy LINEAR ULTRA OUTPUT).
        Wszystkie tuby z linia w starej nazwie ja maja (0 brakow).
      - Srednica z 2 cyfr ("TWIST PRO 70mm", "85mm") i "500 R"; wymiar kwadratowy "330x330 mm" z "SQUARE 330", "500 SQ",
        "SLIM SQUARE 155" oraz "292X292MM" (Orbis Alicia, Orbis Click 430x430mm). Ramki akcesoriow "105x105 mm".
- [x] Ledvance 2026-10-19 v12: linia TYLKO ze slowa - sama litera V/P/S wylaczona ("moga byc bledne": 4099854455995
      "MR16 DIM S" dostal Superior bez zrodla, "URBAN LITE S" = rozmiar S). Wylaczone tez `version_letters` (wyrozniki
      V/P/S/E). Rodzina Urban za marka (`families`): "... Ledvance Urban Lite S" (176 szt.). 2265 nazw zmienionych,
      WARNING 222 -> 301, powtorzenia 152 -> 161 grup (60 zniknelo, 69 doszlo). Wynik `..._2026-10-19_v12.csv`,
      szablony `..._szablony_nazw_2026-10-19_v9.xlsx`. Inni dostawcy bez zmian.
- [x] Ledvance 2026-10-20: MR16/PAR16 DIM S/P rozdzielone CRI z karty GTIN ("CRI 97" / "CRI 90", `add_card_cri`,
      CRI zapisywane przez `scripts/fetch_ledvance_web.py`) - 41 nazw, 0 powtorzen MR16/PAR16 DIM. Kody zamowieniowe
      i nazwy rodzin z kart (LSSPCLB40, LEDPIN20, PAR165036, TUBE, SPOT, 1XLAMP) zablokowane jako wyrozniki. Nowe karty
      w cache uruchomily tez stare reguly: 12x "op. 3 szt." (PCR / CLASSIC V 40), 6 kartotek "LAMPS ... Globe/Edison"
      przebudowanych z karty na plafony (WARNING - EAN w Ekonomie wskazuje inny produkt). 125 grup powtorzen.
      Wynik `output/Ekonom-nazwy-Ledvance_2026-10-20.csv`, kontrola `reports/Ekonom-nazwy-Ledvance_kontrola_2026-10-20.xlsx`.
- [x] Ledvance 2026-10-20 v2: brakujace lm z karty ledvance.com/en/gtin/<EAN> (`lumen_from_gtin_card`; pobrane 4384 karty
      `scripts/fetch_ledvance_web.py --missing-lumen ... --workers 4`, ~15 min). Brak lm 4429 -> 833, DO_SPRAWDZENIA 4554 -> 1085,
      OK 5232 -> 8130. lm tylko gdy moc i barwa karty = Ekonom (246 pominietych z uwaga). Wynik `..._2026-10-20_v2.csv`.
- [x] Ledvance 2026-10-20 v3: karty z kilkoma stopniami mocy (Power Select / Multi Lumen, "55 W / 49 W" + "6650 lm / 5900 lm",
      tez "250  W") i komplety SETn. Brak lm 833 -> 398 (238 moc/barwa inna - zestawy "4 x 2,6W", Smart+ bez mocy, rewizje
      Planon 20W -> 19W; 115 kart bez lm; 45 bez karty). Wynik `..._2026-10-20_v3.csv`.
- [x] Ledvance 2026-10-20 v4: reszta opraw wielostopniowych - karty pobrane ponownie (stary parser gubil "15700  lm"),
      kilka mocy w nazwie ("11,3/14W", "900/1200W", "4-7-9W") -> lm w tej samej kolejnosci. Brak lm 398 -> 333, zadna karta
      z kilkoma stopniami nie zostala bez lm (222 moc/barwa inna niz Ekonom, 66 kart bez lm, 45 bez karty). Wynik `..._v4.csv`.
- [x] Ledvance 2026-10-20 v5 ("tak"): linia z rodziny na karcie GTIN (`line_from_card_family`, 47 rodzin potwierdzonych
      slowem, 0 sprzecznych) - +367 linii (Value 303, Performance 50, Superior 14). Tuby z czujnikiem ruchu maja "z czujnikiem"
      (atrybut sensor). CRI z karty tylko dla lamp CRI >= 90. Zablokowane wyrozniki CLA 60 / Valuecla / Clas.
      Powtorzenia 208 -> 258 (glownie pary stara/nowa kartoteka, ktore roznila tylko linia). Wynik `..._v5.csv`.
- [x] Ledvance 2026-10-20 v6 ("brakuje nazw wlasnych" - uwaga szefa): nazwa wlasna oprawy z rodziny na karcie GTIN
      (`family_from_card`, za marka jak Urban), slowo powtorzone z poczatku nazwy obciete (decyzja: "Panel LED ... Ledvance
      Flex"). 2296 nazw zmienionych, ~2100 opraw z nazwa wlasna; ~600 opraw ma w karcie tylko kategorie, reszta bez karty.
      Dociagniete karty dla opraw/akcesoriow bez cache (`fetch_ledvance_web.py --ean-file`). Nazwa z karty bez typu
      ("LED Line 118 ...") nie zastepuje juz polskiej nazwy z Ekonoma. Powtorzenia 258 -> 233. Wynik `..._v6.csv`.
- [x] Ledvance 2026-10-20 v7 ("wyszukaj w internecie wszystkie i znajdz sposob zeby je przypisac"): pobrane karty GTIN dla
      wszystkich opraw/akcesoriow (3319/3349, 2x HTTP 500). Nazwa wlasna w 2953/3298 opraw (v5: tylko Urban). "Damp Proof",
      "Panel Compact" to nazwy wlasne (filtr kategorii tylko dla jednego slowa); bez karty / "Phase out" - nazwa innych kart
      doslownie w nazwie z Ekonoma albo rodzina z podpowiedzi wyszukiwarki; PHOTOCELL zostaje ("Floodlight Photocell").
      Reszta (~345) to rodziny bez nazwy poza kategoria + linia (HIGH BAY V, PANEL PERFORMANCE, FLOODLIGHT PERFORMANCE).
      Blad z v5 naprawiony: "Phase out" byl traktowany jak rodzina w `line_from_card_family` - 84 swietlowki LUMILUX mialy
      niesluszne "Value". Nowe karty daly tez 131 x "op. N szt." akcesoriom i 90 linii tubom/kompaktom. Do sprawdzenia:
      4 moduly PANEL-CUBE (karta podaje inna barwe niz kod -830/-840 w nazwie). Wynik `..._v7.csv`, 233 grupy powtorzen.
- [x] Ledvance 2026-10-20 v8 (uwaga: 4058075470712 bez "PLANON Frameless", "jak jest round to dodac, ze plafon jest okragly"):
      "™" psul dopasowanie nazwy z Ekonoma (PLANONTM); atrybut `round` -> "okragly/okragla" po rzeczowniku dla panel /
      downlight / plafon / wpuszczana (170 nazw: 92 plafony, 74 downlighty, 4 panele). ROUND w dekor / smart / akcesoriach
      zostaje w serii jak dotad. Wynik `..._v8.csv`, statusy i powtorzenia bez zmian.
- [x] Ledvance 2026-10-20 v9: nazwa wlasna zaraz za rzeczownikiem ("Naswietlacz LED Floodlight Compact 50W ... Ledvance",
      decyzja uzytkownika; `family_after_noun`, tez Urban), oprawy pod tuby: "pod tube LED Damp Proof Housing 1x120cm".
      Czujnik / awaryjna z rodziny karty (38 + 18), gdy Ekonom podaje tylko litere ("SURFACE SLIM RD S 200") - zniknely
      smieciowe wyrozniki "RD 200", "Eueledv", "Blkh 250". 2945 nazw zmienionych (2862 tylko kolejnosc). Wynik `..._v9.csv`.
- [x] Ledvance 2026-10-20 v10 ("co znaczy N" -> "tak"): kody optyki zamienione na kat z karty GTIN (`beam`, 297 kart
      pobranych ponownie `--missing-beam`) - 287 nazw (TruSys, Low Bay Flex, Damp Proof Max, szynowe, ECO HP FLOOD).
      Naprawiony blad: optyka "W" za kodem barwy ("140W 840 W") wypadala jak moc. Zostaje 1x "opt. PW45LST" (optyka
      drogowa). Wynik `..._v10.csv`, statusy i powtorzenia bez zmian.
- [x] Ledvance 2026-10-20 v11 (uwagi: HE = High Efficiency tylko w hermetykach, GU5.3 + 12V, zgubiony Click w Orbis):
      "Damp Proof High Efficiency" (+ "5x Throughwiring"), "GU5.3 12V" (134 zarowki, 0 x 230V), Click-DIM / Click-CCT /
      Remote-CCT w plafonach (75), EL / SN -> awaryjna / z czujnikiem. Zniknely resztki "HE THROUGHWIRING", "5TH", "HE SN".
      243 nazwy zmienione, statusy i powtorzenia bez zmian. Wynik `..._v11.csv`.
- [x] Ledvance 2026-10-20 v12 (przeglad nazw wlasnych hermetykow, 281 szt. / 58 rodzin): Eco Class Damp Proof (+ High
      Lumen Output), Throughwiring (Flex / Compact / Submarine / Eco Class), Damp Proof Eco spojnie (bez Narrow, wersje VW
      z "Phase out" tez Eco), Smart+ "Submarine" bez "Damp Proof", dlugosc 0.6M / 1.2M (tez Office Line). 57 nazw.
- [ ] Do decyzji (hermetyki): "DAMP PROOF ... HOUSING TH.-KIT 5x2.5" (4 szt.) - karta: rodzina DAMP PROOF HOUSING
      ACCESSORIES, wiec to pewnie zestaw okablowania przelotowego, a nazwa mowi "Oprawa hermetyczna pod tube"; pary
      DAMP PROOF EMERGENCY gen 1 / gen 2 roznia sie tylko resztka "EL" (karta: GEN 2); "DAMP PROOF ECO 3TH ON-OFF" -
      znaczenie 3TH niepewne, brak dlugosci (karta: 1200 / 1500).
- [x] Ledvance 2026-10-20 v13 (przeglad plafonow, 314 szt. / 104 rodziny): srednica dla ~60 plafonow bez "mm" (Surface
      Bulkhead, Disc, Compact, Circ, Bamboo Paper, Orbis Disc 30CM), kwadraty "Square 200mm", "Color + White" (bez falszywego
      "bialy"), kolory dwutonowe (czarno-niebieski, biale/drewno - takze Smart+ Orbis Moon, Decor Impact), Orbis Square /
      Sparkledot / Plate Click, czujnik Orbis IP44 i DALI w Wallwasherach z nazwy karty (zniknely smieci ALDA/ALON).
      128 nazw. 1 WARNING: Essentials Floodlights EU Plug 20W - karta "Sensor", rodzina bez czujnika.
- [ ] Do decyzji (plafony): "Ceiling Moia" - karta nazywa ja dzis "Orbis Slim Moia"; "+ RC" w Color + White = pilot
      (dopisac "z pilotem"?); Orbis Spiral (5 kartotek z nazwa zarowki w Ekonomie) - rodzina "Main and Side Light".
- [x] Ledvance 2026-10-20 v14 ("jak w nazwie wlasnej jest round to nie musimy dawac okragly"): bez "okragly" przy
      Round / RD / Circular w nazwie wlasnej - 160 nazw (Ceiling Round, Bulkhead Round, Surface Circular, Slim Round, AR
      Spot / Floodlight / Inground Round, Surface Slim RD...). "okragly" zostaje w 12 (Planon Frameless, Click, Color + White).
- [x] Ledvance 2026-10-20 v15 ("tutaj nie powinien byc bulkhead jako nazwa wlasna?"): pojedyncze slowo rodziny tez jest
      nazwa wlasna - Bulkhead (6), Batten (6), Spot (16), Floodlight (52, "FLOODLIGHT PERFORMANCE"), Floodlights (5 Smart+).
      85 nazw, statusy i powtorzenia bez zmian.
- [x] Ledvance 2026-10-20 v16 ("nazwa wlasna jednak na koncu wszedzie"): decyzja z v9 odwrocona - nazwa wlasna (rodzina
      z karty GTIN, Urban, Damp Proof Housing pod tuby) za marka, przed linia: "Naswietlacz LED 150W ... Ledvance Floodlight
      R30 Performance". 3028 nazw (3022 tylko kolejnosc; 6 x Damp Proof Housing MT GEN3 dostalo "GEN3", bo bez tego roznily
      sie od wersji E tylko pozycja "MT"). Statusy i 233 grupy powtorzen bez zmian. Wynik `..._v16.csv`.
- [x] Ledvance 2026-10-20 v17 ("Lampa stolowa Decor Cardboard Square E27 ... - nie powinno byc na koncu?"): seria
      dekoru i Smart+ (Decor, Endura, Orbis, SUN@HOME, Tracklight...) tez za marka (`series_after_brand`, jak inni
      dostawcy): "Lampa stolowa E27 brazowa Ledvance Decor Cardboard Square". 924 nazwy (923 tylko kolejnosc; "DOT-IT 3"
      z "3 pack" stracil resztke "3"). Akcesoria "do <system>" bez zmian. Statusy i 233 grupy powtorzen bez zmian.
- [x] Ledvance 2026-10-20 v18 ("tak napraw" - zdublowany "Reflektor", "Cob Strips With"): 6 odblyskow Tracklight Spot
      -> "Odblysnik do Tracklight Spot 75 mm SP" (akcesorium), "z pilotem" w 15 tasmach / Linear Slim (WITH REMOTE CONTROL,
      + RC), 7 tasm FLEX RGB / AUDIO / BENDABLE z pilotem bylo "Pilot RC 3m" - teraz Tasma LED; bez "With" / "W" / "+ W"
      w seriach (Bundle, Garden Spot). 31 nazw. Nowa grupa powtorzen (234): Flex TV Mood Light 4058075666917 /
      4099854317743 - ta sama tasma, stara i nowa kartoteka (bylo rozroznione resztka "With").
- [x] Ledvance 2026-10-21 v19 (przeglad recznie partiami po typie, 13 partii, "sprawdz 3 razy"): 1600 nazw zmienionych
      wzgledem v18; statusy 8646 OK / 1015 WARNING / 420 DO_SPRAWDZENIA (v18: 8515 / 933 / 633), braki barwy 205 -> 126,
      lm 286 -> 236. Naprawy zapisu Ekonoma (`name_cleanup`: PANELANON/PANELUS/PANELUG, SUNATHOME, Obris, RGBTWBK, PEND,
      CLICKD, REM, CAM, ERECESSEDMOUNTFRAME) i dat Excela w `Nazwa dodatkowa` (`undo_excel_date`: tasmy "927-965/5" byly
      "927m"). Nowe typy: transformator (HTM/HTL/HTN/HTi/ET Parrot), zarowka_piekarnik, zarowka_sygnalizacyjna,
      zarowka_tradycyjna, latarka, zestaw_nest, klosz_hql, akcesorium_hb, adapter_czujnika. Karta GTIN uzupelnia braki
      (moc, ksztalt, TW, protokol, T8) w `card_fill_types`; przebudowa z karty zachowuje typ zarowki (Smart+ Spot R50/R63/GU10
      byly "Downlight"), ksztalt, TW/RGBW, linie, DIM, PCR i rozmiar banki z Ekonoma. Czujnik ruchu / zmierzchu, NAV "/I" =
      z zaplonnikiem, A75/A100 = A60, kamery / termostaty / piloty / wentylator / panel solarny / kotwy jako wlasne
      rzeczowniki, ramki natynkowe / wpuszczane paneli z wysokoscia, akumulatory Ah/V, tuby 72 cm i FOOD / FLUORA.
      Wynik `output/Ekonom-nazwy-Ledvance_2026-10-21_v19.csv`, kontrola `reports/Ekonom-nazwy-Ledvance_kontrola_2026-10-21_v19.xlsx`.
- [x] Ledvance 2026-10-21 v20 (decyzje uzytkownika do listy z v19, 19 z 21 wdrozonych, test `LedvanceDecisions20261021Test`):
      karta GTIN z innym trzonkiem / ksztaltem / dlugoscia tuby -> nazwa z Ekonoma + DO_SPRAWDZENIA; Gold/Smoke/Amber po
      angielsku; 3-stopniowa / Duo Click / GlowDim; "z pilotem" w oprawach; ON-OFF tylko przy blizniaku DALI/1DIM; "opt. RV20ST";
      Retrofit wyroznikiem; bez "BW"; Print Sky 1 + SETn = op. n szt.; Decor Spot N-punktowy; HQI K z karty (5 kart pobranych);
      kamery = "Kinkiet / Naswietlacz / Reflektor LED z kamera" z W/lm/K. 591 zmian wzgledem v19; statusy 8634 OK / 966 WARNING /
      481 DO_SPRAWDZENIA (+61 DO_SPR z kart innego produktu), 290 grup powtorzen. Inni dostawcy bez zmian (sprawdzone).
      Wynik `output/Ekonom-nazwy-Ledvance_2026-10-21_v20.csv`, kontrola `reports/Ekonom-nazwy-Ledvance_kontrola_2026-10-21_v20.xlsx`.
- [x] Ledvance 2026-10-26 - zestawienie nowych kartotek (`input/Zestawienie Ledvance.csv`, 585 poz., tylko 4 EAN w bazie v20):
      format "nazwa | reszta" w jednej kolumnie, karty GTIN pobrane dla 584 EAN, przebieg laczny z v20 (wyrozniki QTz / FIT5 /
      Plus / opal / 24° / PTi MINI wzgledem juz nazwanych). Nowe typy / porzadki dla starych nazw Osram (samochodowe, Linestra,
      HALOSTAR, HCI / HMI, L15/76, DS7/41-827, SUBSTITUBE, 1K2LM), moc z kodu zamowieniowego karty ("LVPR16501204,5W") juz nie
      jest moca. Wynik `output/Ekonom-nazwy-Ledvance-zestawienie_2026-10-26.csv`, kontrola
      `reports/Ekonom-nazwy-Ledvance-zestawienie_kontrola_2026-10-26.xlsx` (503 OK / 70 WARNING / 12 DO_SPRAWDZENIA). Reguly zmieniaja
      21 nazw v20 przy ponownym przebiegu (Natura, BASIC 640 = 4000K, Eco Area, SUBSTITUBE, bez "Zestaw") - v21 nie generowany.
- [ ] Do decyzji (2026-10-26): 4 grupy powtorzen w zestawieniu (identyczne nazwy w Ekonomie: LUMILUX T5 14W 840 x2, T8 36W 840 /
      840-1, T8 30W 840 / "30 W/840", SENSOR CEILING 360DEG x2) i 43 nazwy takie same jak w v20 (glownie blizniaki stara / nowa
      kartoteka: LUMILUX T5, Value Class A60, Entry EM, PIN G9) - arkusz "Takie same nazwy w v20".
- [ ] Do decyzji (v20): powtorzenia nowe wzgledem v18 - `reports/Ekonom-Ledvance_powtorzenia_do_decyzji_v20_2026-10-21.xlsx`
      (69 grup: 18 tub "TUBE T8 EM 1200 mm" / "LEDTUBE T8 EM V 1200", 6 par swieczek BA / BW, blizniaki stara / nowa kartoteka
      Vintage Gold, multipaki Smart+ "3pack" przy pojedynczych); tlumaczenia EN->PL do akceptacji -
      `reports/Ekonom-Ledvance_tlumaczenia_EN-PL_do_akceptacji_2026-10-21.xlsx` (213 propozycji, 98 fraz).
- [x] Do decyzji (v19) - pelna lista z przykladami w `reports/Ekonom-Ledvance_do_decyzji_2026-10-21.md` (zapytac
      uzytkownika na poczatku nastepnej sesji z Ledvance): 55 nowych grup powtorzen (28 zarowek, 26 tub) - wczesniej rozroznione tylko smieciami
      ("T8 36", "CLP 40", "Valcla 100", "Retrofit"); 75 kart GTIN opisujacych inny produkt niz Ekonom (inny trzonek /
      ksztalt); lista 19 pytan z przegladu (Vintage, klosze Smoke/Amber/Gold, CRI 90, slownik EN->PL akcesoriow, kody optyk
      ulicznych RV25ST, Retrofit jako wyroznik...).
- [ ] Do decyzji: zestawy Smart+ Bundle ("Bundle Flex Light Google Nest", "Terminal EU Repeater & 2") - nazwy nadal
      surowe; "z pilotem" tylko w dekorze - plafony Color + White "+ RC", zarowki RGBW "with remote control",
      Endura Flood Split Solar bez dopisku.
- [ ] Do decyzji: T5 "HO 24" / "HE 21" (zamiennik swietlowki HO / HE) wypadly z nazw tub - byly tylko wyroznikiem.
- [ ] Decyzja uzytkownika: nowe karty uruchomily `fix_from_gtin_page` dla 253 kartotek (WARNING): ~200 to drobne rewizje
      mocy (7,5->6,8W, 8,5->8W, 4->3,4W), 19 inna barwa, 16 moc >15%, 6 inny gwint; czesc nazw z karty brzydka
      ("Oprawa LED 9W", "LED Line 118 125 300° DIM 15W ..."). Pytanie: moc z karty czy z Ekonoma przy drobnych rewizjach.
      Tez 132 multipaki "op. N szt." z karty. Powtorzenia 125 -> 208 (87 nowych = blizniaki, ktore roznil tylko brak lm u jednego).
- [ ] Decyzja uzytkownika (reszta z 69 grup): 12 grup CLASSIC V vs P (rozni tylko trwalosc 10 000 / 15 000 h),
      blizniaki z ta sama karta, 6 obudow DAMP PROOF.
- [ ] (stare) Decyzja uzytkownika: 69 nowych grup powtorzen (rozniace sie tylko litera) - karty GTIN zestawione w
      `reports/Ekonom-Ledvance_powtorzenia_bez_liter_2026-10-19.xlsx` (kolumna "Decyzja uzytkownika"):
      12 grup MR16/PAR16 DIM S vs P = CRI 97 vs CRI 90 (propozycja: "CRI 97"/"CRI 90" z karty), 12 grup CLASSIC V vs P
      roznia sie tylko trwaloscia 10 000 / 15 000 h, 38 grup blizniakow z ta sama karta, 6 obudow DAMP PROOF bez danych.
- [ ] Do decyzji (Ledvance, nie ruszane): pozostale WIELKIE wyrozniki tub (PLASTIC, PLUS, HIGH TEMPERATURE, HIGHBAY,
      BIO-LUMILUX); "RD"/"SQ" jako wyroznik opraw AR i URBAN WALL bez wymiaru (okragla / kwadratowa?); "SURFACE BULKHEAD
      250/300" - liczba to pewnie srednica; panele dalej "60x60" w cm; "TWIST PRO ... White Black" -> kolor "czarny".
- [x] 2026-10-19: zasady Ledvance przeniesione na Kobi/Kanlux/GTV/Bemko (marka + seria na koncu, kolejnosc zarowek,
      Tuba, "Oprawa pod tuby", srednica mm, bez podwojnych slow). Statusy i powtorzenia bez zmian (Kobi 10, Kanlux 0,
      GTV 115, Bemko 16). Wyniki `output/Ekonom-nazwy-<Dostawca>_2026-10-19.csv`.
- [x] Ogolne "Oprawa"/"Oprawa LED" sprawdzone u producentow (2026-10-19, "tak sprawdz te produkty"): GTV 160/163
      z kategorii gtv.com.pl (Plafon, Downlight, Panel, Kinkiet, elewacyjna, kanalowa, meblowa...), Bemko 27/27 (katalog
      2026: OXAN kloszowa = liniowa, KANLO plafoniera, ALICE z rodziny kodow C39 = plafon), Kobi 3 (NEXFORCE z feedu,
      Aquarius). Zostaje: GTV 3 wycofane (Lion-UGR, Celta, Faro), Kobi 47 starych kartotek spoza feedu (HPL, IMPA,
      Lahti, Nuuk... - brak danych w feedzie i w sieci). Przy okazji: srednica nie z "AC220-240V" ani z kata "110°".
      bemko.pl dalej nie odpowiada (timeout). Wyniki `..._2026-10-19_v2.csv` (Kobi, GTV, Bemko), raport
      `reports/Ekonom-ogolne-oprawy_sprawdzone_2026-10-19.xlsx`.

Nazwy do Ekonoma, Kobi - likwidacja bledow (2026-10-15):

- [x] Przebudowa na aktualnych regulach i przeglad bledow: wynik `output/Ekonom-nazwy-Kobi_2026-10-15_v3.csv`,
      kontrola `reports/Ekonom-nazwy-Kobi_kontrola_2026-10-15_v3.xlsx` (962 OK / 234 DO_SPRAWDZENIA / 182
      WARNING, 10 grup powtorzen). Poprawione m.in.: ucinanie modelu o dlugosc rzeczownika ("Akumulator
      Kobi 0", "Gniazdko Smart Kobi IFI"), marka przed "do ...", lm sklejone z IP ("54550lm"), dlugosc
      swietlowek T8, oprawy na GU10/R7s bez wymogu lm, "Zestaw szynowy" zamiast reflektora, Bu/przekroj
      przedluzaczy, zasilacz puszkowy, wyrozniki "SOLARIS 2"/"Quazar 15". 171 nazw zmienionych wzgledem `_v2`.
- [x] Zrodla lm: kobi.pl ma puste karty produktow; b2b.kobi.pl (`/pl/query/<fraza>`, karta bez logowania)
      pokazuje te same parametry co feed i nie zna starych kodow. Dopasowanie do nastepcy w feedzie po
      parametrach dalo 6/234, w tym 2 bledne (ZOE -> SIGARO) - nie wdrozone.
- [x] ~230 starych kartotek bez lm (spoza feedu) - decyzja 2026-10-15: "jak Kobi nie interesuja te produkty
      to nas tez nie" - zostaja bez lm, bez dalszej pracy.
- [ ] Decyzja uzytkownika: 10 grup powtorzen to zasilacze - 9 par stara kartoteka (0043xx/002383, brak w
      feedzie) + nowa (005xxx) o identycznej nazwie; para 002391/004602 jest w feedzie z roznymi EAN.
- [ ] Do potwierdzenia: oprawy STREET (S/M/L, ME/P, 20.20) bez znaczenia kodow optyki; HPL/IMPA/NEXFORCE
      jako "Oprawa" (typ dekor), "HPL1 45W 90" - 90 to pewnie kat swiecenia.
- [ ] `distinguish` / `pack_size` z Ledvance dla Kobi - nie bylo potrzeby (powtorzenia to tylko zasilacze).

Nazwy do Ekonoma, Kanlux - runda 2 (2026-10-15):

- [x] Parametry z kanlux.com dla 1979 kodow spoza plikow Parametry: `scripts/fetch_kanlux_web.py` -> 711 w
      ofercie (karty w `cache/kanlux_web/`), 1268 wycofanych (wyszukiwarka ich nie zna) - zostaja z danymi z
      nazwy, bez dalszej pracy. Wynik `output/Ekonom-nazwy-Kanlux_2026-10-15_v2.csv`, kontrola
      `reports/Ekonom-nazwy-Kanlux_kontrola_2026-10-15_v2.xlsx`: 4441 OK / 317 DO_SPRAWDZENIA (315 wycofanych) /
      241 WARNING (typy bez schematu: rozdzielnice, wentylatory, tasmy izolacyjne), 0 grup powtorzen.
- [x] Poprawione: lm z klucza Φuse (FIZU, SIGO), "max 7,3" przy zintegrowanym LED, lm panelu bez mA
      (BLINGO 2280-3800lm), oprawy z trzonkiem bez wymogu mocy/lm/K + "4xG13", MOD = Moduł LED, tasmy
      kolorowe z barwa, zasilacz TRETO 40W (nie "16/40W"), rozmiar przy serii (STONO/STOBI/TUBEO/MERSA),
      `cm` = czarny mat (parametry Kanlux, nie zgadniety czarny metalik).
- [ ] Do potwierdzenia: nazwy typow bez schematu (np. "Rozgałęźniki, gniazda, wtyczki Kanlux GRIP GP/S" -
      rzeczownik z kategorii Kanlux).

Nazwy do Ekonoma, GTV - runda 2 (2026-10-15):

- [x] Karty gtv.com.pl (`scripts/fetch_gtv_web.py`, 730 kodow z problemem: 703 na stronie, 27 wycofanych)
      uzupelniaja lm/K/IP. Wynik `output/Ekonom-nazwy-Gtv_2026-10-15_v2.csv`, kontrola
      `reports/Ekonom-nazwy-Gtv_kontrola_2026-10-15_v2.xlsx`: 5689 OK / 886 WARNING / 25 DO_SPRAWDZENIA (bylo 54).
- [x] Poprawione: wielkosc liter starych nazw WIELKIMI wg pisowni na gtv.com.pl ("z przewodem", "do szyny
      1-fazowej"), marka Högert za fraza rzeczownika ("Cyfrowy miernik uniwersalny Högert"), zapinki jako
      akcesoria, "max. 42W" przy tasmie, "2W1" to nie moc, "840LED" to nie 4000K, powtorzone przymiotniki
      ("kwadratowa ... kwadratowa"), "1,5m" nie daje "5m", "BIAŁY NEUTRAL" = neutralna biała.
- [ ] Decyzja uzytkownika: 115 grup powtorzen (99 Högert - identyczne nazwy w Ekonomie). Zestawienie roznic z
      hoegert.com: `reports/Ekonom-Hogert_powtorzenia_2026-10-15.xlsx` (kolumna "Decyzja uzytkownika").

Blokady:

- [ ] Brak.

---

## 15. Notatki techniczne

- Projekt ma byc lokalnym narzedziem CLI.
- Pierwszy etap nie korzysta z internetu.
- Oryginalny plik produktowy zawsze zostaje bez zmian.
- Wyniki ida do `output/` i `reports/`.
- Dane niepewne dostaja status `DO_SPRAWDZENIA`.
