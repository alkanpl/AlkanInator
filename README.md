# AlkanInator

Lokalne narzedzie CLI do masowej analizy produktow, ekstrakcji atrybutow z tytulow i generowania tytulow SEO na podstawie configu kategorii.

## MVP

Pierwsza wersja dziala bez internetu:

1. Wczytuje plik CSV/XLSX.
2. Wykrywa podstawowe kolumny.
3. Analizuje braki i problemy.
4. Wyciaga podstawowe atrybuty z tytulow.
5. Normalizuje wybrane wartosci.
6. Generuje nowe tytuly SEO z configu.
7. Zapisuje plik wynikowy i raporty.

## Instalacja

```bash
pip install -r requirements.txt
```

## Uzycie

```bash
python src/run_pipeline.py --input input/products.xlsx --category "Oprawy sufitowe"
```

Mozesz tez wskazac config i katalogi wyjsciowe:

```bash
python src/run_pipeline.py --input input/products.csv --config configs/categories/oprawy-sufitowe.yaml --output output/products_optimized.xlsx --reports-dir reports
```

Domyslnie pipeline dolacza wiedze z `dictionaries/woocommerce_catalog_knowledge.yaml`, jesli ten plik istnieje. Mozesz wskazac inny slownik albo wylaczyc te warstwe pusta wartoscia:

```bash
python src/run_pipeline.py --input input/products.csv --catalog-knowledge dictionaries/woocommerce_catalog_knowledge.yaml
python src/run_pipeline.py --input input/products.csv --catalog-knowledge ""
```

Dla pliku z wieloma arkuszami XLSX wskaz arkusz:

```bash
py src/run_pipeline.py --input archived_input_files/input_2026-08-05/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --config configs/categories/kanlux-oswietlenie.yaml --output output/kanlux_master_optimized.xlsx --reports-dir reports/kanlux_master
```

Uzupelnianie brakujacych atrybutow z pliku parametrow Kanlux:

```powershell
py src/enrich_from_parameters.py --input archived_input_files/input_2026-08-05/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --parameters archived_input_files/input_2026-08-05/Parametry.xlsx --config configs/categories/kanlux-oswietlenie.yaml --output output/alkan_kanlux_master_with_parameters.xlsx --reports-dir reports/parameter_enrichment
```

Budowanie slownika typow, producentow i kategorii z pelnego eksportu produktow:

```powershell
py src/learn_product_taxonomy.py --input input/Produkty-Export-2026-April-30-0723.xlsx --output dictionaries/learned_product_taxonomy.yaml --reports-dir reports/product_taxonomy
```

Pelna analiza eksportu produktow WooCommerce CSV:

```powershell
py src/analyze_woocommerce_export.py --input input/wszystko.csv --reports-dir reports/woocommerce_catalog --dictionary dictionaries/woocommerce_catalog_knowledge.yaml
```

### Import i eksport WooCommerce w Alkan

W tym projekcie masowe aktualizacje istniejacych produktow WooCommerce sa wykonywane przez wtyczki **WP All Export** i **WP All Import**, a nie przez wbudowany importer CSV WooCommerce.

- Plik `input/Hager-Berker-Gniazdka.xlsx` jest wzorcem struktury eksportu i ponownego importu.
- Plik wynikowy powinien pozostac pelnym XLSX: ten sam arkusz, wszystkie produkty, te same naglowki i ta sama kolejnosc kolumn.
- Nalezy zmieniac tylko pola objete zadaniem. Nie wolno redukowac pliku do `SKU`, tagow i slotow atrybutow w formacie wbudowanego importera WooCommerce.
- Wartosci atrybutow sa w kolumnach `Atrybut Produktu: <nazwa>`.
- Deklaracje atrybutow produktu sa zapisane jako serializowana tablica PHP w kolumnie `Product Attributes`. Dodanie nowego atrybutu wymaga uzupelnienia zarowno kolumny wartosci, jak i tej deklaracji.
- Tagi sa aktualizowane w `Product Tags`; `Up-Sells` i `Cross-Sells` pozostaja bez zmian, jezeli zadanie ich nie dotyczy.
- Nowa kolumne atrybutu mozna dodac do pelnego eksportu bez usuwania lub przebudowywania kolumn zrodlowych.

Cross-sell i up-sell Hager/Berker generowane od nowa w pełnej kopii eksportu WP All Import, z potwierdzeniem kompatybilności w katalogu PDF z `input/`:

```powershell
py -X utf8 src/build_hager_berker_catalog_cross_upsell.py
```

Skrypt indeksuje kody i sekcje `Współpracuje z` w katalogu, modyfikuje wyłącznie kolumny `Cross-Sells` i `Up-Sells`, zapisuje osobny plik do importu oraz raport audytowy z kodami i stronami katalogowymi dla każdego powiązania. Relacje wariantowe, stare duplikaty i nieistniejące SKU są pomijane; sprawdzane jest też kodowanie polskich znaków.

Szczegolowa regula i kanoniczny wzorzec sa zapisane w `dictionaries/woo_all_export_import_workflow.yaml`.

Ten slownik zasila kolejne uruchomienia pipeline'u: rozszerza liste znanych producentow, kolorow, serii, typow produktow, wartosci IP, kategorii i rekomendowanych filtrow bez recznego kopiowania tych danych do configow kategorii.

Przygotowanie eksportu CSV z Baselinkera pod pipeline AlkanInatora:

```powershell
py src/prepare_baselinker_input.py --input input/przykładowycsv.csv --output output/baselinker_prepared.csv --default-category "Gniazdka i Łączniki"
```

Skrypt rozbija kolumne `features` na normalne kolumny, mapuje najwazniejsze parametry na format uzywany przez pipeline i zapisuje raport w `reports/baselinker_prepare`.

Eksport wyniku AlkanInatora do CSV importowego Baselinkera:

```powershell
py src/export_to_baselinker_csv.py --input input/alkan_kanlux_master_names_from_parameters_v10_zolte_export.xlsx --output output/baselinker_import_alkan_kanlux_v10_zolte.csv
```

Skrypt tworzy CSV z kolumnami zgodnymi z przykladem Baselinkera: `product_id;name;sku;ean;manufacturer_name;category;description;features;images_urls`. Nazwa produktu bierze sie domyslnie z `new_title`, kategoria z `proponowana_kategoria_1`, a parametry techniczne trafiaja do JSON w kolumnie `features`.

Bezposredni eksport i synchronizacja przez API BaseLinkera:

```powershell
$env:BASELINKER_TOKEN = "token ustawiony tylko w tej sesji PowerShell"
py src/baselinker_api_cli.py inventories
py src/baselinker_api_cli.py export --inventory-id 12345 --output output/baselinker_backup.csv
py src/baselinker_api_cli.py sync --inventory-id 12345 --input output/baselinker_import.csv
py src/baselinker_api_cli.py sync --inventory-id 12345 --input output/baselinker_import.csv --apply
```

Polecenie `sync` bez `--apply` zawsze wykonuje dry-run. Przed kazdym zapisem powstaja
`backup_before.json`, `catalog_metadata_before.json`, raport `audit.csv` ze zmianami
stara/nowa wartosc oraz `summary.json`
w `reports/baselinker_api_sync/<data_godzina>/`. Domyslnie cechy sa scalane, puste pola
nie czyszcza danych w BaseLinkerze, a blad jednego rekordu blokuje cala paczke. Rozszerzenia
takie jak tworzenie brakujacych kategorii/producentow, czesciowa wysylka i zmiana SKU/EAN
wymagaja jawnych flag widocznych w `py src/baselinker_api_cli.py sync --help`.
Tryb `--match-by auto` dopasowuje najpierw po SKU, potem po EAN, a `product_id` traktuje
jako fallback, poniewaz w historycznych plikach projektu ta kolumna moze zawierac ID WooCommerce.

Propozycje kategorii na podstawie nazwy, typu produktu i wiedzy katalogowej:

```powershell
py src/recommend_categories.py --input output/gazetka_produkty_nazwy_final_20260526_po_suffixach_z_opisami.xlsx --output output/category_recommendations_gazetka_20260526.xlsx --top-n 3
```

Skrypt dopisuje kolumny `proponowana_kategoria_1..N`, `category_score_1..N`, `category_confidence_1..N` i `category_reasons_1..N`. Domyslnie korzysta z `dictionaries/woocommerce_catalog_knowledge.yaml` oraz `dictionaries/learned_product_taxonomy.yaml`.
Zaakceptowane reczne korekty po SKU mozna dopisywac w `dictionaries/category_recommendation_overrides.yaml`.

Analiza slow kluczowych per typ produktu przez Google Keyword Planner:

```powershell
py src/analyze_keywords.py --input output/products_optimized.xlsx --output output/products_with_keywords.xlsx --reports-dir reports/keywords --customer-id 1234567890
```

Ten sam etap mozna wlaczyc w glownym workflow:

```powershell
py src/run_pipeline.py --input input/products.xlsx --category "Oprawy sufitowe" --output output/products_optimized.xlsx --reports-dir reports/products --analyze-keywords --google-ads-config google-ads.yaml
```

Skrypt wykrywa typ produktu z kolumn `attr_typ`, `Typ produktu`, `Typ` albo `Rodzaj produktu`, pobiera propozycje slow kluczowych, sprawdza historyczne `avg_monthly_searches` i dopisuje `primary_keyword`, `secondary_keyword`, metryki oraz `keyword_candidates_json`. Konfiguracja Google Ads jest czytana z `--google-ads-config`, `GOOGLE_ADS_CONFIGURATION_FILE_PATH`, `~/google-ads.yaml` albo lokalnego `google-ads.yaml`; `--customer-id` mozna pominac, jesli YAML ma pole `customer_id`. Do testow bez API mozna uzyc `--offline-keywords` z CSV zawierajacym `keyword`, `avg_monthly_searches` i opcjonalnie `product_type`.

Anatomia tytulow produktowych:

```powershell
py src/run_pipeline_with_parameters.py --input archived_input_files/input_2026-08-05/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --parameters archived_input_files/input_2026-08-05/Parametry.xlsx --config configs/categories/kanlux-oswietlenie.yaml --output output/alkan_kanlux_master_names.xlsx --reports-dir reports/alkan_kanlux_master_names --title-anatomy configs/title_anatomy.yaml
```

Globalny config `configs/title_anatomy.yaml` buduje tytul wedlug schematu `primary_keyword` albo typ produktu jako fallback, potem atrybuty per typ produktu, producent i kod producenta. Jezeli dla typu nie ma zaakceptowanej reguly anatomii, pipeline wraca do starego template'u YAML i dodaje warning `title_anatomy_missing_accepted_rule`. Raporty trafiaja do `reports/<run>/title_anatomy/title_anatomy_review.xlsx` - jeden plik z trzema arkuszami: `Przeglad typow`, `Duplikaty` i `Ostrzezenia`. Kazdy arkusz ma dwie kolumny do recznego przegladu: `proponowana zmiana` (aktualny wynik programu) i `oczekiwany wynik` (pusta kolumna do wpisania docelowej wartosci).

Tagowanie wariantow i uzupelnianie klas wysylkowych bez zmiany nazw ani atrybutow:

```powershell
py src/enrich_variants_and_shipping.py --input "C:\Users\Handlowiec\Downloads\Kanlux (3).xlsx" --output output/kanlux_variants_shipping.xlsx
```

Skrypt dopisuje kolumny kontrolne `variant_group_key`, `variant_group_size`, `variant_tag`, `variant_link_tag`, `variant_differentiating_attributes`, `variant_attribute_summary`, `variant_reason`, `shipping_class_predicted`, `shipping_class_confidence`, `shipping_class_reason` i `shipping_category_guess`. Jesli istnieje kolumna `Product Tags`/`Tagi`, tag wariantu jest do niej dopisywany tylko dla rodzin, ktore realnie roznia sie atrybutami. Raport `reports/<nazwa_wyjscia>/variant_tag_report.xlsx` pokazuje tagi, SKU w grupie i atrybuty laczace wariacje. Jesli istnieje kolumna `Shipping Class`/`Klasa wysylkowa`, puste komorki sa uzupelniane przewidywana klasa, a istniejace wartosci pozostaja bez zmian. Flaga `--overwrite-shipping` pozwala nadpisac istniejace klasy.

## Wyniki

Pipeline tworzy:

- `output/products_optimized.xlsx`
- `output/products_optimized.csv`
- `reports/quality_report.html`
- `reports/missing_attributes.csv`
- `reports/changes_report.xlsx`
- `reports/changes_report.csv`
- `reports/category_summary.csv`

Oryginalny plik wejsciowy nie jest nadpisywany.

## Testy

Testy uzywaja biblioteki standardowej `unittest` (bez dodatkowych zaleznosci):

```powershell
py -m unittest discover -s tests -p "test_*.py"
```

Kluczowe sa testy charakteryzujace (golden-file) dla ekstrakcji atrybutow i
generowania tytulow: utrwalaja aktualne zachowanie na realnych danych, wiec
kazda niezamierzona regresja w regexach lub regulach od razu wywala test.

Po swiadomej zmianie regul ekstrakcji, configow tytulow lub slownikow
zregeneruj snapshoty (wymaga `archived_input_files/input_2026-08-05/Alkan_Kanlux_pelne_rodziny.xlsx`):

```powershell
py tests/generate_snapshots.py
```
