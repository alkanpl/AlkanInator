# Zmiany w AlkanInatorze

Najważniejsze zmiany dla osób korzystających z AlkanInatora, od najnowszych. Jak z nich korzystać: [instrukcja](instrukcja.md).

## 2026-11-05

- **Wiedza o katalogu sklepu ze sklepu, nie z ręcznego eksportu.** Słownik kategorii, atrybutów, producentów i danych producenta (z niego korzystają nazwy, kategorie i pliki do Baselinkera) można odświeżyć z pliku pobranego ze sklepu przez API. Przy okazji wartości z przecinkiem nie są już cięte („2,5” zamiast „2” i „5”), a dane producenta są kompletne (doszli m.in. Legrand, Rosa, TME).
- **Nazwy do Ekonoma – GTV, poprawki po przeglądzie ostrzeżeń.** Przedłużacze GTV i Högert według wspólnego schematu („Przedłużacz 4 gn. Schuko 3m GTV”, „Przedłużacz bębnowy 4 gn. 30m 3x1,5mm2 Schuko GTV”, „Przedłużacz bębnowy 30m 3x2,5mm2 guma Högert”). W akcesoriach nawias nie urywa się w połowie, a skład zestawu i jednostki zostają („Zestaw GLAX … (profil 2mb + osłonka mleczna + 2 zaślepki końcowe) GTV”, „GLAX 12m”). „1-fazowy” i „2x gniazdo” małymi, „USB A+C” wielkimi, poprawione literówki z Ekonoma („gniazko”, „uziemieniemieniem”, „włacznik”).
- **Nazwy do Ekonoma – Kobi.** Przedłużacze Kobi przebudowane według wspólnego schematu, a czujniki z rodzajem („z czujnikiem ruchu”).

## 2026-11-04

- **Nazwy do Ekonoma – przedłużacze według wspólnego schematu.** Domowe: liczba gniazd, uziemienie / Schuko / bez uziemienia, wyłącznik, długość, przekrój, nazwa własna, producent – np. „Przedłużacz 3 gn. Schuko z wyłącznikiem 1,5m 1,5mm2 Bemko”, „Przedłużacz 3 gn. z uziemieniem z wyłącznikiem 5m czarny Linea Kobi”. Bębnowe: liczba gniazd, długość, przekrój, materiał przewodu, hermetyczność, uziemienie / Schuko, nazwa własna, producent – „Przedłużacz bębnowy 4 gn. 15m 3x1,5mm2 Schuko Bemko”. Skróty Kobi (Zu, Bu, W) są rozpisane słowami, a seria Kobi (Linea, Connecto…) wróciła do nazwy.
- **Nazwy do Ekonoma – Kanlux, runda 3.** 1156 nowych kartotek razem z bazą z poprzednich rund, bez powtórzeń nazw. Nowe kody aparatury: wyłączniki KMBN6 i RCBO KRO, rozłączniki KMI/KCS, styczniki KMC, ograniczniki przepięć KSD/KSPD, wyłączniki kompaktowe KMCCB-F. „UGR≤” trafia do Ekonoma jako „UGR<” (wcześniej „UGR?”).
- **Changelog i instrukcja w Notion.** Ta strona i instrukcja aktualizują się w Notion same po każdym commicie.

## 2026-11-03

- **Nazwy do Ekonoma – Signify (Philips / PILA), runda 1.** Parametry z kart lighting.philips.pl po 12NC, a dla kodów bez strony z kart PDF Signify. Nazwy własne rodzin opraw (Malaga, ClearWay…) i linie (MASTER, CorePro, CoreLine…) na końcu nazwy. Powtórzenia nazw: z 251 do 63 grup do decyzji.
- **Nazwy do Ekonoma – Bemko, runda 3.** 340 nowych kartotek, nowe typy (żarniki R7s, ramki i zasilacze do paneli, oprawy stadionowe, awaryjne z lumenami), poprawki literówek z Ekonoma. 21 grup powtórzeń do decyzji.
- **„Wgraj zmiany do sklepu” na nowym komputerze.** `Instalacja.bat` instaluje potrzebne biblioteki, a `Diagnostyka - Wgraj zmiany.bat` pokazuje błędy startu.

## 2026-11-02

- **Aplikacja „Wgraj zmiany do sklepu”** – zapis do sklepu bez terminala. Podgląd kart produktu przed i po zmianie, odznaczanie pól, „Zatwierdź i zapisz”, „Cofnij ten zapis”. Na produkcji zapis potwierdza się w okienku Windows.
- **Połączenie z produkcją.** Pobieranie i zapis działają na sklepie produkcyjnym (domyślnie nadal sklep lokalny).
- **Szybsze wgrywanie pobranych plików.** Plik pobrany ze sklepu pamięta stan z chwili pobrania, więc do sklepu idą tylko wiersze zmienione w pliku – niezmienione nie cofają cudzych zmian.
- **Podgląd wysyłki i wariantów.** Karta w podglądzie pokazuje metody dostawy i ceny przed i po zmianie, grupę wariantów i przełącznik wariantów.

## 2026-10-29

- **Bezpośrednie połączenie ze sklepem (AlkanInator Bridge).** Pobieranie produktów do XLSX w układzie WP All Export, porównanie pliku ze sklepem przed zapisem, zapis tylko zmian, dziennik zapisów i cofanie. Każda osoba ma własny token.
- **Kontrola zdjęć przy zapisie.** Nowe zdjęcie, które nie jest JPG 1000×1000, blokuje zapis produktu.
- **Wspólne wzory tytułów** dla wszystkich osób pracujących z AlkanInatorem.
- **Wagi wysyłkowe paneli LED** ustawione tak, żeby paczka kurierska mieściła właściwą liczbę sztuk zależnie od producenta (Philips/PILA 8, Ledvance 12, pozostali 10). Wagi i wymiary rur karbowanych.

## 2026-10-28

- **Simon 24 w sklepie.** Opisy pisane per SKU, zdjęcia 1000×1000 na białym tle, tagi wariantów (grupa „Kolor - Simon”), Up-Sells z ramkami w kolorze, Cross-Sells i produkty kompatybilne z katalogu producenta.

## 2026-10-26

- **Nazwy do Ekonoma – nowa partia Ledvance** z zestawienia Ekonoma, liczona razem z bazą, żeby nowe nazwy nie powtarzały już nadanych. Świetlówki liniowe nazywają się „Świetlówka”.

## 2026-10-21

- **Nazwy do Ekonoma – Ledvance v20.** Wdrożone 21 decyzji z przeglądu partiami: inne produkty na kartach GTIN, klosze Gold / Smoke / Amber, optyki uliczne kodem, kamery Smart+ jako oprawy z kamerą.

## 2026-10-20

- **Nowe partie Kontakt Simon prosto z e-katalogu producenta** (pierwsza: Simon 24). Nazwy ze słownika serii, zdjęcia z największego oryginału, bez grafik „brak zdjęcia”.
- **Nazwy własne opraw Ledvance** (Floodlight Compact, Damp Proof…) na końcu nazwy, za marką.

## 2026-10-14 – 2026-10-19

- **Nowa funkcja: nazwy do Ekonoma** – krótkie nazwy dla sprzedawców z samymi parametrami przydatnymi przy ladzie. Ledvance, Kobi, Kanlux, GTV i Bemko. Brakujące lumeny i parametry z katalogów PDF i stron producentów, powtórzenia nazw zgłaszane do decyzji, marka na końcu nazwy.

## 2026-10-12

- **Kontrola wartości atrybutów przed importem.** Wartości spoza słownika sklepu są wychwytywane, zanim utworzą w filtrach nowe pozycje (poprawka partii Simon kaszmir).
- **Usuwanie produktów:** szkic i przekierowanie 301 na zamiennik albo kategorię, bez łańcuchów przekierowań.
- **Kable bez producenta** w tytułach i opisach.

## 2026-10-07

- **Odcinki kabli 1 m** dziedziczą atrybuty, tagi, opis i klasę wysyłkową po konfekcjach 25/50/100 m. Producent kabli zawsze „Polski Producent”. Puste atrybuty kabli uzupełniają się z rodzeństwa, tytułu i typowych wartości.

## 2026-09-28 – 2026-09-30

- **Podserie Kontakt Simon** (Line, Duo, Nature, GO, Premium…) w atrybucie `Podseria` i w tytule.
- **Słowa kluczowe Incobex** z Keyword Plannera.
- **Instrukcja „Claude + AlkanInator”** dla osób zlecających pracę.

## 2026-09-15 – 2026-09-18

- **Produkty kompatybilne Hager/Berker** – osobna sekcja na karcie produktu, budowana tylko z relacji producenta (katalog PDF i BMEcat).
- **Tagi wariantów dla nowych kolorów Simon** kopiowane od rodzeństwa, żeby działał przełącznik kolorów.
- **Atrybuty tylko z listy sklepu.** Cechy dostawcy spoza listy trafiają do raportu do zatwierdzenia, a nie do importu.
- **Opisy z FAQ** – blok 3 pytań i odpowiedzi ze schemą dla Google, rodziny produktów w generatorze opisów, bank naturalnych fraz z poprawionych opisów. Naprawa opisów, w których Baselinker zepsuł schemę FAQ.
- **Wysyłka do Baselinkera przez API**, narzędzia Hager/Berker (zdjęcia, relacje, sposób montażu, Cross-Sells i Up-Sells), poprawki atrybutów Simon do modelu sklepu.

## Maj – lipiec 2026

- Pierwsze wersje AlkanInatora: pipeline Kanlux do Baselinkera i WooCommerce, ekstrakcja atrybutów z tytułów, tytuły SEO według „anatomii tytułu”, opisy produktów i opisy do gazetki, analiza słów kluczowych Google Ads, rekomendacje kategorii. Testy pilnują, żeby poprawki reguł nie psuły wcześniejszych wyników.
