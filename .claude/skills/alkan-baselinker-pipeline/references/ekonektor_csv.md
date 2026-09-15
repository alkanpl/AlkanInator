# eKonektor — masowe dodanie i edycja produktów

Źródło reguł: `eKonektor - masowe dodanie i edycja produktów do Base (1).docx`.

## Format CSV

Zapisać dokładnie osiem wymaganych kolumn w podanej kolejności:

1. `TOWARIDX` — kod towaru z Ekonoma.
2. `ACTIVE` — zawsze `1`.
3. `PRODUCT_CODE` — kod towaru w Ekonoma; w eKonektorze staje się kodem produktu.
4. `NAME` — docelowa nazwa produktu w eKonektorze.
5. `TAX_ID` — zawsze `1`.
6. `UNIT` — wartość ze słownika eKonektora „Jednostki miar”, kolumna „Wartość w sklepie”.
7. `PRODUCER` — wartość ze słownika eKonektora „Producenci”, kolumna „Wartość w sklepie”.
8. `CATEGORY` — wartość ze słownika eKonektora „Kategorie”, kolumna „Kategoria produktu”; domyślnie `Bez kategorii`.

Użyć separatora `;`. Wszystkie pola muszą być niepuste. Dla istniejących indeksów Alkan zachować pełny kod Ekonoma z sufiksem dostawcy zarówno w `TOWARIDX`, jak i `PRODUCT_CODE`, np. `KZN023/HAG`.

## Kodowanie

Zapisać plik jako ANSI zgodne z aktywną polską stroną kodową: Windows-1250 (`cp1250`), bez BOM. Przed zapisem zastąpić znaki niedostępne w cp1250 w sposób zachowujący znaczenie, co najmniej:

- `²` → `2`
- `³` → `3`
- `≤` → `<=`
- `≥` → `>=`
- `×` → `x`
- nietypowe łączniki i minusy → `-`

Po normalizacji wymagać ścisłego kodowania cp1250; nie używać zapisu z `errors=replace`.

## Ustalone wartości katalogowe Alkan

- Hager: `UNIT=szt`, `PRODUCER=HAGER`.
- Gdy nie ma potwierdzonego mapowania kategorii eKonektora: `CATEGORY=Bez kategorii`.

## Walidacja pliku

- Potwierdzić dokładnie osiem nagłówków i ich kolejność.
- Potwierdzić liczbę rekordów względem źródła.
- Potwierdzić brak pustych pól.
- Potwierdzić `ACTIVE=1` i `TAX_ID=1` w każdym wierszu.
- Potwierdzić `TOWARIDX=PRODUCT_CODE` i unikalność kodów.
- Ponownie odczytać wynik jako cp1250 i sprawdzić, że plik nie ma BOM.
- Otworzyć próbkę w Excelu jako Windows-1250 z separatorem średnikowym i sprawdzić polskie znaki.

## Procedura importu w eKonektorze

1. Upewnić się, że wszystkie indeksy istnieją w Kartotece towarów Ekonoma.
2. Otworzyć `Produkty` → `Kartoteki` → `Import z pliku CSV`.
3. Wskazać plik i ustawić `ANSI (aktywna strona kodowa)`.
4. Wybrać `Przetwarzaj plik`; żadna kolumna ani wartość nie może być czerwona. Zielony `PRODUCT_CODE` oznacza aktualizację istniejącego produktu.
5. Wybrać `Importuj`.
6. W zakresie `Produkty przygotowane do sklepu` zaznaczyć wszystkie produkty i użyć `Popraw`.
7. Zaznaczyć: produkt aktywny w sklepie, przenoszenie stanów magazynowych oraz automatyczne przenoszenie cen sprzedaży z Ekonoma.
8. Zapisać, ponownie przefiltrować produkty przygotowane do sklepu i wybrać `Wyślij do Baselinker`.
