# AlkanInator – instrukcja

AlkanInator to zestaw narzędzi i słowników z regułami naszego sklepu. Robi z plików dostawców gotowe dane sklepowe: nazwy, kategorie, atrybuty, opisy, zdjęcia, warianty i pliki importu do Baselinkera, WooCommerce i Ekonoma.

Nie trzeba umieć programować. Narzędziami steruje Claude (albo Codex) w VS Code: Ty dajesz pliki i piszesz po polsku, co ma powstać, a on uruchamia narzędzia i oddaje plik z raportem do kontroli. Decyzje handlowe zostają po Twojej stronie.

Co się ostatnio zmieniło: [Zmiany w AlkanInatorze](changelog.md).

## Pierwsze uruchomienie

1. Otwórz w VS Code folder `AlkanInator` i kliknij ikonę Claude – to jest czat, w którym zlecasz pracę.
2. Jeśli chcesz sam wgrywać zmiany do sklepu (aplikacja „Wgraj zmiany do sklepu”):
   - zainstaluj Pythona z python.org i w instalatorze zostaw zaznaczone „py launcher”,
   - kliknij dwukrotnie `Instalacja.bat` w folderze AlkanInator,
   - wygeneruj własny token w WooCommerce → AlkanInator API (pokazuje się tylko raz) i poproś Claude'a o założenie pliku `src/woo_api/config.json`. Tokenu nie wysyłaj nikomu – po nim sklep wie, kto zapisał zmianę.
3. Gdy aplikacja nie startuje, uruchom `Diagnostyka - Wgraj zmiany.bat` i skopiuj komunikat z okna.

## Jak zlecać pracę

**1. Wrzuć pliki do `input/`.** Załóż folder dostawcy i wgraj tam wszystko, co dotyczy partii: plik dostawcy (parametry, EAN, zdjęcia), listę SKU, świeży eksport sklepu (gdy produkty już są w sklepie), katalog PDF. Działają XLSX, CSV i XML. Oryginały nigdy nie są nadpisywane.

**2. Napisz zlecenie.** Dobre zlecenie mówi: gdzie są pliki, co z nimi zrobić i co ma być wynikiem. Jedna partia i jeden dostawca na jedno zlecenie.

> W input/Dostawca/ są cennik.xlsx (parametry i zdjęcia) i indeksy.xlsx (lista SKU do dodania). Przygotuj nazwy, kategorie, atrybuty z listy sklepu i zdjęcia. Najpierw plik kontrolny XLSX, po akceptacji CSV do Baselinkera.

**3. Sprawdź wynik.** Gotowe pliki są w `output/`, raporty i skoroszyty kontrolne w `reports/`. Otwieraj je w Excelu. Każda zmiana ma źródło, pewność i status:

| Status | Co znaczy |
|---|---|
| OK | wartość potwierdzona, może iść do importu |
| WARNING | działa, ale warto rzucić okiem |
| DO_SPRAWDZENIA | Claude nie miał pewności – decyzja należy do Ciebie |

Zacznij od wierszy `DO_SPRAWDZENIA`. Import robisz dopiero po przejrzeniu pliku kontrolnego.

**Poprawki.** Uwagi pisz normalnie w czacie, a przy długiej liście w `input/Do poprawy.txt` (SKU i to, co ma być zamiast obecnej wartości). Claude poprawia regułę, a nie pojedynczą komórkę, i oddaje nowy plik z końcówką `_v2`, `_v3` (Excel blokuje otwarte pliki).

**Decyzje na stałe.** Gdy coś ustalisz, napisz „zapamiętaj”. Decyzja trafia do słowników i instrukcji projektu, więc nie wróci jako pytanie przy kolejnym dostawcy – także u Codexa.

## Co można zrobić

### Nowa partia produktów do Baselinkera

Nazwy, kategorie Alkan, atrybuty, zdjęcia i plik importu z plików dostawcy (Kanlux, Kobi, Ospel, Kontakt Simon, Hager/Berker i inni). Dla Kontakt Simon dane biorą się wprost z e-katalogu producenta.

- **Zlecenie:** „Zrób nazwy i plik do Baselinkera dla input/Dostawca/cennik.xlsx”.
- **Wynik:** skoroszyt kontrolny, a po akceptacji CSV do Baselinkera albo plik eKonektora.
- **Uwaga:** cechy i wartości spoza listy sklepu nie trafiają do importu, tylko do arkusza „Atrybuty do zatwierdzenia” – nowy atrybut albo nowa wartość powstaje dopiero po Twojej decyzji.

### Opisy produktów

Opisy HTML SEO z blokiem FAQ (3 pytania i odpowiedzi plus schema dla Google). Każdy opis jest pisany osobno dla danego SKU, nie z szablonu.

- **Zlecenie:** „Napisz opis HTML dla SKU 37176/KAN” albo „opisy dla wszystkich produktów z briefu …”.
- **Uwaga:** opisy z FAQ wgrywa się do sklepu przez WP All Import albo „Wgraj zmiany do sklepu”. Przez Baselinkera schema FAQ się psuje i w sklepie widać goły tekst.

### Nazwy do Ekonoma

Krótkie nazwy dla sprzedawców przy ladzie – nie tytuły SEO. Same parametry przydatne przy sprzedaży: typ, gwint, moc, lumeny, barwa, IP, marka na końcu. Obsługiwani dostawcy: Ledvance, Signify (Philips/PILA), Kanlux, Kobi, GTV, Bemko.

- **Zlecenie:** „Robimy nazwy do Ekonoma dla input/EkonomNazwy/Dostawca”.
- **Wynik:** CSV w układzie Ekonoma (`Nazwa towaru` do 40 znaków, reszta w `Nazwa dodatkowa`) i skoroszyt kontrolny.
- **Uwaga:** gdy dwa produkty dostają tę samą nazwę, Claude pokazuje, czym się różnią, a Ty decydujesz, co dopisać do nazwy.

### Zmiany w produktach, które już są w sklepie

Poprawki tytułów, opisów, kategorii, tagów, atrybutów, wag, wymiarów, klas wysyłkowych, zdjęć, Up-Sells i Cross-Sells. Cen i stanów nie zmieniamy – należą do Baselinkera.

- **Zlecenie:** „Pobierz ze sklepu produkty z kategorii … i popraw …” albo „wgraj zmiany z pliku output/…xlsx na produkcję”.
- **Wynik:** plik do wgrania w aplikacji „Wgraj zmiany do sklepu” (opis niżej) albo gotowe polecenie do wklejenia.
- **Uwaga:** każdy zapis ma identyfikator i da się go cofnąć.

### Warianty i powiązania produktów

Tagi wariantów dla nowych kolorów (przełącznik kolorów na karcie produktu), Up-Sells, Cross-Sells i „Produkty kompatybilne” (osobna sekcja na karcie produktu, tylko z relacji producenta).

- **Zlecenie:** „Uzupełnij tagi wariantów w eksporcie input/…xlsx”.
- **Uwaga:** raport zawsze podaje, po jakim atrybucie łączą się warianty (np. `Kolor`). Nowy tag trzeba dopisać do grupy wariantów w sklepie, inaczej przełącznik się nie pojawi.

### Kable

Odcinki 1 m dziedziczą dane po konfekcjach 25/50/100 m, a puste atrybuty uzupełniają się z rodzeństwa, tytułu i typowych wartości (YDY, OWY, LGY…). Producent kabli to zawsze „Polski Producent”, a tytuły i opisy nie wymieniają producenta.

- **Zlecenie:** „Uzupełnij odcinki 1 m w eksporcie kabli input/…xlsx”.

### Wagi i wysyłka

Wagi i wymiary wysyłkowe, np. panele LED pakowane tak, żeby w paczce kurierskiej mieściła się właściwa liczba sztuk (Philips/PILA 8, Ledvance 12, pozostali 10).

### Usuwanie produktów

Usuwany produkt dostaje status szkicu i przekierowanie 301 na zamiennik albo kategorię – bez łańcuchów przekierowań.

- **Co dać:** eksport z kolumnami `Do usunięcia` (tak/x/1) i `Zamiennik` (SKU, ID albo adres).
- **Wynik:** plik ze statusem `draft` do WP All Import i plik przekierowań do wtyczki Redirection.

### Kontrola przed importem

Każdy plik z atrybutami można sprawdzić względem wartości, które już są w sklepie – wychwytuje literówki i wartości, które utworzyłyby w filtrach sklepu nowe, śmieciowe pozycje.

- **Zlecenie:** „Sprawdź wartości atrybutów w pliku output/…xlsx przed importem”.

## Wgraj zmiany do sklepu – krok po kroku

1. Kliknij dwukrotnie `Wgraj zmiany do sklepu.bat`. Strona otworzy się w przeglądarce.
2. Wybierz sklep (lokalny albo produkcja) i plik XLSX z produktami, kliknij „Pokaż podgląd zmian”. Na tym etapie w sklepie nic się nie zmienia.
3. Przejrzyj karty „Przed zmianą” i „Po zmianie” – zmienione pola są w ramce. Karta pokazuje też warianty, powiązane produkty i metody dostawy przed i po zmianie.
4. Jeśli jakiegoś pola nie chcesz zmieniać, odznacz je w „Zapisuj tylko zaznaczone pola” i kliknij „Przelicz podgląd”.
5. Kliknij „Zatwierdź i zapisz”. Na produkcji Windows dodatkowo poprosi o potwierdzenie w okienku.
6. Po zapisie dostajesz identyfikator zapisu. Przycisk „Cofnij ten zapis” przywraca poprzednie wartości (pola zmienione w sklepie później zostają).

Dobrze wiedzieć:
- Plik pobrany przez AlkanInatora pamięta stan z chwili pobrania, więc wysyłane są tylko wiersze, które zmieniłeś. Nie usuwaj z niego ukrytego arkusza i nie przepisuj danych do nowego skoroszytu.
- Pusta komórka czyści pole w sklepie. Kolumny, której nie ma w pliku, zapis nie rusza.
- Plik z innego sklepu (np. eksport z produkcji wgrywany na sklep lokalny) jest zatrzymywany przed podglądem.
- Nowe zdjęcia muszą być JPG 1000×1000 na białym tle, inaczej podgląd pokaże błąd.
- Aplikacja zamyka się sama po 15 minutach bez otwartej strony.

## Zasady, których pilnujemy zawsze

- Pliki wejściowe nigdy nie są nadpisywane – wyniki idą do `output/` i `reports/`.
- Zdjęcia produktów zawsze 1000×1000 JPG na białym tle, w każdym kanale.
- Nowe atrybuty i nowe wartości atrybutów w sklepie tylko po Twoim zatwierdzeniu.
- Opisy z blokiem FAQ do sklepu przez WP All Import albo „Wgraj zmiany do sklepu”, nie przez Baselinkera.
- Kilka wartości jednego atrybutu rozdziela się znakiem `|` (`Antenowe|Komputerowe`), nie przecinkiem.
- Wysyłka do Baselinkera i zapis na produkcji tylko na Twoje polecenie. Na produkcji zapis zatwierdza człowiek.

## Ściąga – gotowe hasła do czatu

| Hasło | Co się stanie |
|---|---|
| Zrób nazwy i plik do Baselinkera dla input/… | nowa partia: nazwy, kategorie, atrybuty, zdjęcia |
| Napisz opis HTML dla SKU … | opis SEO z FAQ dla jednego produktu |
| Robimy nazwy do Ekonoma dla … | krótkie nazwy dla sprzedawców |
| Wgraj zmiany z pliku … na produkcję | gotowy plik i polecenie do wgrania |
| Uzupełnij tagi wariantów w eksporcie … | tagi przełącznika kolorów |
| Zaktualizowałem Do poprawy | Claude wdraża uwagi z input/Do poprawy.txt |
| Zapamiętaj: … | decyzja zapisana na stałe |
| Ile produktów wpadło w rodzinę „nieznana”? | kontrola klasyfikacji nowego dostawcy |
| ver-upd | podsumowanie ostatnich zmian w projekcie |

Nie wiesz, jak coś zlecić? Zapytaj wprost: „jak to zrobić w AlkanInatorze?”.
