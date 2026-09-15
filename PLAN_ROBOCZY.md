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

Blokady:

- [ ] Brak.

---

## 15. Notatki techniczne

- Projekt ma byc lokalnym narzedziem CLI.
- Pierwszy etap nie korzysta z internetu.
- Oryginalny plik produktowy zawsze zostaje bez zmian.
- Wyniki ida do `output/` i `reports/`.
- Dane niepewne dostaja status `DO_SPRAWDZENIA`.
