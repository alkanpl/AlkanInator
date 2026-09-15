# Workflow Kanlux do Baselinkera

To jest opis po ludzku, jak teraz powinien isc Kanlux w AlkanInatorze.

## O co chodzi

Kanlux nie powinien isc jako zwykle kopiowanie danych z pliku producenta. Plik od producenta jest tylko zrodlem: kod, EAN, nazwa Kanlux, parametry, obrazki, opisy techniczne itd.

Docelowo robimy z tego dane sklepowe:

- normalna nazwa SEO/B2C,
- kategoria z naszego Woo/Alkan, nie kategoria producenta,
- opis wygenerowany lokalnie,
- atrybuty zgodne z Woo catalog knowledge,
- CSV w formacie Baselinkera.

Najwazniejsza zasada: do `features` Baselinkera trafiaja tylko takie nazwy i wartosci atrybutow, ktore sa znane w Woo catalog knowledge. Jesli producent zapisuje to inaczej, najpierw trzeba znormalizowac do naszej wersji.

## Glowne pliki

- Input glowny Kanlux:
  `archived_input_files/Alkan_Kanlux_pelne_rodziny.xlsx`

- Arkusz glowny:
  `7. Wszystkie SKU (master)`

- Parametry Kanlux, jesli trzeba uzupelniac braki:
  `archived_input_files/Parametry.xlsx`

- Config Kanlux:
  `configs/categories/kanlux-oswietlenie.yaml`

- Wiedza z Woo:
  `dictionaries/woocommerce_catalog_knowledge.yaml`

- Nauczona taksonomia / suffixy / kategorie:
  `dictionaries/learned_product_taxonomy.yaml`

- Reczne poprawki kategorii:
  `dictionaries/category_recommendation_overrides.yaml`

## 1. Najpierw wiedza z Woo

Pipeline powinien opierac sie na tym, co juz istnieje w Woo.

Jesli mamy swiezy eksport WooCommerce, najpierw warto odswiezyc slownik:

```powershell
py src/analyze_woocommerce_export.py --input archived_input_files/wszystko.csv --reports-dir reports/woocommerce_catalog --dictionary dictionaries/woocommerce_catalog_knowledge.yaml
```

Ten slownik daje:

- znanych producentow,
- znane atrybuty,
- znane wartosci atrybutow,
- dane producenta,
- kategorie,
- serie, kolory, IP itd.

Dane producenta maja byc brane globalnie z tego slownika. Czyli jak produkt ma producenta `Kanlux`, to system ma dopisac:

`KANLUX; Objazdowa 1-3 41-922 Radzionkow; kanlux@kanlux.pl`

Nie z pliku Kanlux i nie recznie w skrypcie.

## 2. Przygotowanie / sprawdzenie pliku Kanlux

Najpierw sprawdzamy, czy plik ma SKU, EAN, producenta, kategorie i nazwy.

Dla obecnego mastera bylo:

- 499 produktow,
- 0 bez SKU,
- 0 bez EAN,
- glowne grupy: plafoniery/oprawy sufitowe, hermetyczne, panele LED, naswietlacze, high bay.

Do samego przygotowania mozna uzyc:

```powershell
py src/run_pipeline.py --input archived_input_files/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --config configs/categories/kanlux-oswietlenie.yaml --output output/kanlux_master_optimized.xlsx --reports-dir reports/kanlux_master
```

Ten krok robi:

- wykrycie kolumn,
- raport jakosci,
- wyciagniecie atrybutow z nazw i kolumn,
- klasyfikacje: produkt glowny czy akcesorium,
- wygenerowanie nazw,
- walidacje kategorii,
- raport brakujacych atrybutow.

## 3. Uzupelnianie parametrow z Parametry.xlsx

Jezeli master Kanlux nie ma wszystkich parametrow albo nazwy wychodza slabo, trzeba wzbogacic dane z `Parametry.xlsx`.

```powershell
py src/run_pipeline_with_parameters.py --input archived_input_files/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --parameters archived_input_files/Parametry.xlsx --config configs/categories/kanlux-oswietlenie.yaml --output output/alkan_kanlux_master_with_parameters.xlsx --reports-dir reports/kanlux_with_parameters
```

Ten wariant:

- buduje indeks parametrow po kodach/SKU,
- dopisuje brakujace parametry,
- robi raport dopasowan,
- robi raport konfliktow,
- potem generuje nazwy tak samo jak normalny pipeline.

To jest zwykle lepsza sciezka dla Kanluxa, bo Kanlux ma duzo technicznych danych, ktore sa potrzebne do dobrych tytulow: moc, lumeny, CCT, IP, kat, kolor, wymiary, gwint, czujnik.

## 4. Jak powstaja nazwy

Nazwy Kanlux robia sie glownie z `configs/categories/kanlux-oswietlenie.yaml`.

Config ma osobne reguly dla typow produktow, np.:

- plafony LED,
- oprawy sufitowe,
- panele LED,
- naswietlacze,
- high bay,
- oprawy hermetyczne,
- akcesoria typu ramka, zasilacz, uchwyt, linka, klips, czujnik itd.

Nazwa nie powinna byc slepym przepisaniem `Nazwa Kanlux`.

Przyklad logiki:

- jesli to panel LED, nazwa ma isc w stylu `Panel LED ...`,
- jesli to hermetyczna, `Oprawa hermetyczna LED ...`,
- jesli to naswietlacz, `Naswietlacz LED ...`,
- jesli to plafon, uzywamy `Plafon`, nie `Plafoniera`,
- kolory i ksztalty maja byc odmienione pod typ produktu, np. `oprawa biala`, ale `plafon bialy`.

Kod producenta/SKU jest dopisywany na koncu nazwy, bo tak jest ustawione w configu:

```yaml
append_sku_after_producer: true
```

Wazne: unikalnosc tytulu trzeba sprawdzac przed dopisaniem kodu producenta. Po dopisaniu kodu prawie wszystko bedzie sztucznie unikalne.

## 5. Kategorie

Kategorie nie ida z Kanluxa 1:1.

Pipeline ma dobrac kategorie sklepowa z Woo/Alkan przez:

- `dictionaries/woocommerce_catalog_knowledge.yaml`,
- `dictionaries/learned_product_taxonomy.yaml`,
- `dictionaries/category_recommendation_overrides.yaml`,
- `src/recommend_categories.py`.

W pliku wynikowym pojawiaja sie kolumny typu:

- `proponowana_kategoria_1`,
- `category_score_1`,
- `category_reasons_1`,
- `category_recommendation_status`.

Po pipeline trzeba sprawdzic raport:

```text
reports/<nazwa>/suspicious_category_fit.csv
```

Tam sa produkty, gdzie kategoria moze byc podejrzana.

Jesli poprawka jest jednorazowa i pewna, dopisac override do:

```text
dictionaries/category_recommendation_overrides.yaml
```

## 6. Opisy

Opis do Baselinkera ma byc lokalnie wygenerowany.

Nie kopiujemy opisow producenta jako glownego opisu sklepowego. Opis producenta moze zostac w pliku kontrolnym jako material pomocniczy, ale do `description` ma isc opis z naszego generatora.

Przy eksporcie do Baselinkera skrypt bierze opis z kolumn:

- `Opis HTML`,
- `description`,
- `Opis`.

Dlatego po pipeline trzeba sprawdzic, czy opis HTML faktycznie istnieje i nie jest pusty.

## 7. Atrybuty / features

Do Baselinkera idzie JSON w kolumnie `features`.

Eksport robi to przez:

```powershell
py src/export_to_baselinker_csv.py --input output/alkan_kanlux_master_with_parameters.xlsx --output output/baselinker_import_alkan_kanlux.csv
```

Ten skrypt:

- bierze nazwe z `new_title` albo podobnej kolumny finalnej,
- bierze SKU z `SKU`, `Kod Producenta`, `Kod` albo `sku`,
- dopisuje suffix producenta, np. `/KAN`,
- bierze kategorie z `proponowana_kategoria_1`,
- buduje `features`,
- dopisuje `Dane producenta` z Woo catalog knowledge po producencie,
- zapisuje CSV ze srednikiem i UTF-8 BOM.

Wazne dla atrybutow:

- nie wrzucac do Baselinkera przypadkowych nazw atrybutow,
- nie wrzucac wartosci, ktorych nie ma w Woo catalog knowledge,
- jesli Kanlux ma np. `IP20`, a u nas jest `IP 20`, to trzeba znormalizowac do `IP 20`,
- jesli producent ma `4000K`, a u nas wartosc to `4000`, to do features idzie `4000`,
- dane producenta ida globalnie z catalog knowledge, nie z inputu.

## 8. Obrazki

Jesli jest osobny plik z linkami do zdjec, mozna go podac przy eksporcie:

```powershell
py src/export_to_baselinker_csv.py --input output/alkan_kanlux_master_with_parameters.xlsx --output output/baselinker_import_alkan_kanlux.csv --links-file archived_input_files/Linki.xlsx --image-key-column "Kod Kanlux" --main-image-column "Zdjecie glowne"
```

Wtedy skrypt probuje dopasowac zdjecie po kodzie produktu.

## 9. Dodatkowe wzbogacenie gotowego CSV

Jesli mamy juz CSV Baselinkera, ale brakuje kategorii albo danych producenta, mozna puscic:

```powershell
py src/enrich_baselinker_export.py --input output/baselinker_import_alkan_kanlux.csv --output output/baselinker_import_alkan_kanlux_enriched.csv
```

Ten skrypt:

- dopisuje brakujace kategorie,
- poprawia czesc kategorii regulami Kanlux,
- dopisuje dane producenta z Woo catalog knowledge,
- robi raport korekt kategorii.

To jest bardziej krok ratunkowy/kontrolny. Lepiej, zeby glowny eksport od razu mial dobre kategorie i dane producenta.

## 10. Warianty i klasy wysylkowe

To jest osobny krok, nie zmienia nazw ani atrybutow:

```powershell
py src/enrich_variants_and_shipping.py --input "C:\Users\Handlowiec\Downloads\Kanlux (3).xlsx" --output output/kanlux_variants_shipping.xlsx
```

Skrypt dopisuje:

- tagi wariantow,
- grupy wariantow,
- atrybuty roznicujace,
- sugerowana klase wysylkowa,
- raport wariantow.

Uzywac po imporcie/eksporcie, kiedy trzeba ogarnac warianty i wysylke.

## 11. Co sprawdzic przed wrzuceniem do Baselinkera

Minimum kontroli:

- liczba wierszy zgadza sie z inputem,
- `name` nie jest puste,
- `sku` nie jest puste,
- `ean` nie jest puste, chyba ze naprawde nie bylo w zrodle,
- `manufacturer_name` to `Kanlux`,
- `category` to kategoria Woo/Alkan, nie kategoria producenta,
- `description` nie jest puste,
- `features` jest poprawnym JSON-em,
- `features` zawiera `Producent`,
- `features` zawiera `Dane producenta`,
- nie ma atrybutow ani wartosci spoza Woo catalog knowledge,
- tytuly sa unikalne przed dopisaniem kodu producenta,
- jednostki sa pisane razem tam gdzie powinny, np. `4000lm`, `600mm`, `20W`,
- `Plafon`, nie `Plafoniera`,
- kolory/ksztalty sa odmienione normalnie po polsku.

## 12. Najczestsza praktyczna kolejnosc

Najczesciej robilbym tak:

```powershell
py src/run_pipeline_with_parameters.py --input archived_input_files/Alkan_Kanlux_pelne_rodziny.xlsx --sheet "7. Wszystkie SKU (master)" --parameters archived_input_files/Parametry.xlsx --config configs/categories/kanlux-oswietlenie.yaml --output output/alkan_kanlux_master_with_parameters.xlsx --reports-dir reports/kanlux_with_parameters
```

Potem sprawdzenie raportow:

```text
reports/kanlux_with_parameters/missing_attributes.csv
reports/kanlux_with_parameters/suspicious_category_fit.csv
reports/kanlux_with_parameters/title_strategy_report.csv
```

Potem eksport:

```powershell
py src/export_to_baselinker_csv.py --input output/alkan_kanlux_master_with_parameters.xlsx --output output/baselinker_import_alkan_kanlux.csv
```

Na koncu walidacja CSV i ewentualnie:

```powershell
py src/enrich_baselinker_export.py --input output/baselinker_import_alkan_kanlux.csv --output output/baselinker_import_alkan_kanlux_enriched.csv
```

## 13. Jak myslec o Kanluxie

Kanlux ma byc traktowany podobnie jak Kobi, tylko tu podstawowy input jest bardziej tabelkowy i duzo logiki siedzi w configu `kanlux-oswietlenie.yaml`.

Najpierw rozpoznajemy produkt i jego parametry, potem tworzymy sklepowa nazwe, potem dobieramy sklepowa kategorie, potem dopiero robimy Baselinkera.

Nie odwrotnie.

