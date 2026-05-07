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
