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
py src/run_pipeline.py --input input/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --config configs/categories/kanlux-oswietlenie.yaml --output output/kanlux_master_optimized.xlsx --reports-dir reports/kanlux_master
```

Uzupelnianie brakujacych atrybutow z pliku parametrow Kanlux:

```powershell
py src/enrich_from_parameters.py --input input/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --parameters input/Parametry.xlsx --config configs/categories/kanlux-oswietlenie.yaml --output output/alkan_kanlux_master_with_parameters.xlsx --reports-dir reports/parameter_enrichment
```

Budowanie slownika typow, producentow i kategorii z pelnego eksportu produktow:

```powershell
py src/learn_product_taxonomy.py --input input/Produkty-Export-2026-April-30-0723.xlsx --output dictionaries/learned_product_taxonomy.yaml --reports-dir reports/product_taxonomy
```

Pelna analiza eksportu produktow WooCommerce CSV:

```powershell
py src/analyze_woocommerce_export.py --input input/wszystko.csv --reports-dir reports/woocommerce_catalog --dictionary dictionaries/woocommerce_catalog_knowledge.yaml
```

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

Propozycje kategorii na podstawie nazwy, typu produktu i wiedzy katalogowej:

```powershell
py src/recommend_categories.py --input output/gazetka_produkty_nazwy_final_20260526_po_suffixach_z_opisami.xlsx --output output/category_recommendations_gazetka_20260526.xlsx --top-n 3
```

Skrypt dopisuje kolumny `proponowana_kategoria_1..N`, `category_score_1..N`, `category_confidence_1..N` i `category_reasons_1..N`. Domyslnie korzysta z `dictionaries/woocommerce_catalog_knowledge.yaml` oraz `dictionaries/learned_product_taxonomy.yaml`.
Zaakceptowane reczne korekty po SKU mozna dopisywac w `dictionaries/category_recommendation_overrides.yaml`.

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
