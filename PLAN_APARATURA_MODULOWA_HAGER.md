# Plan przygotowania atrybutów, nazw i opisów — aparatura modułowa Hager

## 1. Cel i zakres

Celem jest przygotowanie spójnego zestawu atrybutów, anatomii nazw oraz zasad
tworzenia opisów dla produktów z pliku:

- `input/Aparatura-Modulowa-Hager.xlsx`

Głównym polskim źródłem zbiorczym dla tej partii jest oficjalny katalog Hager:

- `input/23PL013_KATALOG ROZDZIAL ENERGII_2023.pdf`

Katalog ma 1265 stron. Automatyczna kontrola symboli potwierdziła obecność
**96/96 kodów** z pliku docelowego. Dla 95 produktów istnieje bezpośredni wiersz
w tabeli danej rodziny. Kod `CDA440D` występuje tylko w tabeli danych
technicznych (straty mocy, Ui, Uimp), natomiast tabela handlowa pokazuje
`CDA440J`; ten jeden rekord pozostaje do rozstrzygnięcia po EAN/ID.

Plik `input/Hager aparatura na www.xlsx` traktujemy pomocniczo jako listę
produktów/cen na WWW, a nie jako źródło parametrów technicznych.

Zakres bieżącej partii to **96 produktów**. Plik pomocniczy ma 134 pozycje,
z czego 77 SKU pokrywa się z partią docelową, 19 występuje tylko w pliku
docelowym, a 57 tylko w pliku pomocniczym. Dodatkowych 57 produktów nie należy
automatycznie dołączać bez osobnej decyzji.

Plan nie zakłada nadpisywania wejściowych skoroszytów. Wynik wdrożenia ma trafić
do `output/`, a raporty do `reports/`.

## 2. Inwentaryzacja partii

| Rodzina | Liczba | Docelowa kategoria Alkan |
|---|---:|---|
| Wyłączniki nadprądowe MCB | 49 | Aparatura modułowa > Wyłączniki nadprądowe |
| Wyłączniki różnicowoprądowe RCCB | 12 | Aparatura modułowa > Wyłączniki różnicowoprądowe |
| Wyłączniki różnicowo-nadprądowe RCBO | 3 | Aparatura modułowa > Wyłączniki różnicowoprądowe |
| Rozłączniki izolacyjne | 24 | Aparatura modułowa > Rozłączniki izolacyjne |
| Rozłącznik bezpiecznikowy | 1 | Aparatura modułowa > Rozłączniki izolacyjne |
| Styczniki modułowe | 2 | Aparatura modułowa > Styczniki |
| Ograniczniki przepięć SPD | 2 | Aparatura modułowa > Ochronniki przepięciowe |
| Szyny grzebieniowe | 3 | Aparatura modułowa > Szyny łączeniowe |

Skoroszyt ma 136 kolumn, lecz wartości produktowe występują tylko w 10 kolumnach
atrybutowych. Największe braki w bieżących danych:

- `Symbol`: brak w 33 z 96 produktów, mimo że kod można bezpiecznie wyprowadzić
  z SKU przez usunięcie suffixu `/HAG`;
- `Rodzaj`: brak w 5 produktach;
- `Prąd znamionowy`: brak w 3 produktach;
- `Liczba biegunów`: brak w 2 produktach;
- nie ma osobnych kolumn na Icn, Icu, prąd różnicowy IΔn, napięcie sterowania,
  konfigurację styków ani kluczowe parametry SPD.

## 3. Hierarchia źródeł

Każdy parametr musi mieć `source`, `confidence` i `status`.

1. **Aktualna polska karta produktu Hager dopasowana po symbolu** — źródło
   rozstrzygające dla bieżącego statusu produktu, EAN oraz wartości, które mogły
   zmienić się po wydaniu katalogu w 2023 roku.
2. **Lokalny polski katalog Hager 2023** — główne źródło wsadowe dla nazw rodzin,
   tabel wariantów i parametrów technicznych całej partii. W przypadku konfliktu
   z aktualną kartą internetową pierwszeństwo ma aktualna karta.
3. **Bieżący eksport WooCommerce** — źródło identyfikatorów, SKU i materiał
   kontrolny; nie jest źródłem prawdy dla parametrów.
4. **Obecny tytuł i opis** — wyłącznie źródło pomocnicze. Parametry z tekstu
   można przyjąć dopiero po zgodności z oficjalną kartą.
5. Brak jednoznacznego dopasowania lub sprzeczność danych daje
   `DO_SPRAWDZENIA`; nie wolno uzupełniać parametru zgadywaniem z podobnego SKU.

Oficjalne strony Hager używają symbolu produktu jako stabilnego klucza i podają
parametry właściwe dla danej rodziny. Przykładowe źródła:

- [MBN106E — wyłącznik nadprądowy](https://hager.com/pl/katalog/produkt/mbn106e-mcb-6ka-1p-b-6a)
- [NDN320 — wyłącznik nadprądowy 10 kA](https://hager.com/de/katalog/produkt/ndn320-ls-schalter-3-polig-10ka-15ka-d-20a-3m)
- [CDA225J — RCCB](https://hager.com/pl/katalog/produkt/cda225j-rccb-2p-25a-30ma-typ-a)
- [ADA916D — RCBO](https://hager.com/pl/katalog/produkt/ada916d-rcbo-1p-n-6ka-b-16a-30ma-typ-a)
- [SBN163 — rozłącznik izolacyjny](https://hager.com/pl/katalog/produkt/sbn163-modulowy-rozlacznik-izolacyjny-1p-63a)
- [ESC225 — stycznik](https://hager.com/pl/katalog/produkt/esc225-stycznik-230vac-2no-25a)
- [SPA931 — ogranicznik przepięć](https://hager.com/pl/katalog/produkt/spa931-spd-ogr-przepiec-t1-t2-mov-4p-tn-s-tt)
- [KDN363A — szyna grzebieniowa](https://hager.com/pl/katalog/produkt/kdn363a-szyna-grzeb-widelkowa-3p-10-mm2-12m)
- [L71M — rozłącznik bezpiecznikowy](https://hager.com/pl/katalog/produkt/l71m-rozlacznik-bezpiecznikowy-d02-1p-63a)

### 3.1. Mapa stron w polskim katalogu Hager

| Rodzina / seria | Strony z produktami | Strony techniczne |
|---|---:|---:|
| MCB MBN, charakterystyka B, Icn 6 kA | 996–997 | 1022 i dalsze |
| MCB MCN, charakterystyka C, Icn 6 kA | 999–1000 | 1022 i dalsze |
| MCB NDN, charakterystyka D, Icn 10 kA / Icu 15 kA | 1009 | 1022 i dalsze |
| Szyny grzebieniowe KDN | 1018–1019 | opis kompatybilności w tabelach |
| RCCB CDC typ AC | 1056 | 1086 i dalsze |
| RCCB CDA typ A | 1058 | 1086 i dalsze |
| RCBO ADC typ AC | 1077 | 1090–1092 |
| RCBO ADA typ A | 1078 | 1090–1092 |
| SPD SPA400 / SPA931 | 1102–1103 | 1108 i dalsze |
| Rozłączniki izolacyjne SBN | 1170 | 1196 i dalsze |
| Styczniki ESC | 1185 | dalsze tabele w sekcji styczników |
| Rozłącznik bezpiecznikowy L71M | 874 | 886 oraz tabela danych na 469 |

Numery stron powyżej odpowiadają numerom wydrukowanym w katalogu i numerom stron
PDF. Katalog służy do automatycznego uzupełnienia partii, natomiast polskie karty
internetowe będą kontrolą aktualności i źródłem EAN.

## 4. Priorytety atrybutów

- **P0 — nazwa i podstawowe filtry:** atrybut rozróżniający wariant i potrzebny
  do decyzji zakupowej. Musi być kompletny przed publikacją.
- **P1 — opis i filtry techniczne:** ważny dla doboru i porównania, ale zwykle
  nie powinien wydłużać nazwy.
- **P2 — pełna specyfikacja:** dane dla instalatora i tabeli technicznej. Nie
  są potrzebne w tytule, ale należy je zachować, gdy są dostępne.

### 4.1. Atrybuty wspólne

| Priorytet | Atrybuty |
|---|---|
| P0 | Rodzaj produktu, Producent, Symbol/kod producenta, Układ biegunów, Prąd znamionowy |
| P1 | Napięcie znamionowe łączeniowe Ue, Rodzaj napięcia AC/DC, Liczba modułów, Sposób montażu, Stopień ochrony IP, Norma produktowa |
| P2 | Ui, Uimp, częstotliwość, zakres temperatury pracy, przekrój przewodu sztywnego i elastycznego, moment dokręcania, wymiary, straty mocy, trwałość mechaniczna i elektryczna |

`Układ biegunów` przechowujemy jako wartość handlową, np. `1P`, `1P+N`, `3P`,
`3P+N`. Osobno można zachować numeryczne `Liczba biegunów` i `Liczba biegunów
chronionych`. Nie używamy zamiennie określeń `1-fazowy` i `1-biegunowy`.

### 4.2. Wyłączniki nadprądowe MCB — 49 produktów

**P0:**

- charakterystyka wyzwalania B/C/D;
- prąd znamionowy In [A];
- układ biegunów;
- znamionowa zwarciowa zdolność wyłączania Icn [kA];
- symbol Hager.

**P1:** napięcie Ue, liczba modułów, klasa ograniczenia energii, IP, sposób i
zakres podłączenia, norma EN/IEC 60898-1.

**P2:** Icu według IEC 60947-2, Ui, Uimp, straty mocy, moment dokręcania,
temperatura pracy, trwałość i wymiary. Nie przenosimy do opisu całej tabeli
korekcji prądu dla każdej temperatury — pozostaje ona w karcie technicznej.

Anatomia nazwy:

`Wyłącznik nadprądowy {charakterystyka}{prąd_A} {układ_biegunów} {Icn_kA} Hager {symbol}`

Przykłady:

- `Wyłącznik nadprądowy B6 1P 6kA Hager MBN106E`
- `Wyłącznik nadprądowy C40 1P 6kA Hager MCN140E`
- `Wyłącznik nadprądowy D20 3P 10kA Hager NDN320`

Zdolność zwarciowa musi pochodzić z karty danego SKU. Serie MBN/MCN w tej partii
mają warianty 6 kA, natomiast NDN jest rodziną 10 kA; nie wolno ustawić jednej
wartości dla wszystkich MCB.

### 4.3. Wyłączniki różnicowoprądowe RCCB — 12 produktów

**P0:**

- typ produktu `RCCB` / wyłącznik różnicowoprądowy bez członu nadprądowego;
- prąd znamionowy In [A];
- znamionowy prąd różnicowy IΔn [mA];
- typ wykrywanego prądu różnicowego A/AC;
- układ biegunów;
- warunkowa zdolność zwarciowa Icn [kA];
- symbol Hager.

**P1:** napięcie Ue, liczba modułów, IP, sposób montażu, typ zacisku,
przekroje przewodów, norma EN/IEC 61008-1.

**P2:** zdolność załączania i wyłączania Idm, Ui, Uimp, temperatura pracy,
straty mocy, trwałość i wymiary.

Anatomia nazwy:

`Wyłącznik różnicowoprądowy {In_A} {IΔn_mA} typ {typ_RCD} {układ_biegunów} {Icn_kA} Hager {symbol}`

Przykład:

`Wyłącznik różnicowoprądowy 25A 30mA typ A 1P+N 6kA Hager CDA225J`

RCCB nie ma charakterystyki nadprądowej B/C/D. Nie wolno w opisie twierdzić,
że samodzielnie zapewnia ochronę przeciążeniową lub zwarciową.

### 4.4. Wyłączniki różnicowo-nadprądowe RCBO — 3 produkty

**P0:**

- typ produktu `RCBO` / wyłącznik różnicowoprądowy z członem nadprądowym;
- charakterystyka członu nadprądowego B/C/D;
- prąd znamionowy In [A];
- prąd różnicowy IΔn [mA];
- typ wykrywanego prądu różnicowego A/AC;
- układ biegunów i liczba biegunów chronionych;
- Icn [kA];
- symbol Hager.

**P1:** napięcie Ue, liczba modułów, IP, typ zacisku, sposób montażu, norma
EN/IEC 61009-1.

**P2:** Ui, Uimp, zakres temperatury, przekroje przewodów, moment dokręcania,
trwałość, straty mocy i wymiary.

Anatomia nazwy:

`Wyłącznik różnicowo-nadprądowy RCBO {charakterystyka}{In_A} {IΔn_mA} typ {typ_RCD} {układ_biegunów} {Icn_kA} Hager {symbol}`

Przykłady:

- `Wyłącznik różnicowo-nadprądowy RCBO B16 30mA typ A 1P+N 6kA Hager ADA916D`
- `Wyłącznik różnicowo-nadprądowy RCBO B10 30mA typ AC 1P+N 6kA Hager ADC910D`

Konieczne są dwie oddzielne kolumny: `Charakterystyka wyzwalania` oraz
`Typ prądu różnicowego`. Wartości `B` i `A/AC` opisują inne właściwości i nie
mogą być przechowywane w jednym polu `Typ wyłącznika`.

### 4.5. Rozłączniki izolacyjne — 24 produkty

**P0:** układ biegunów, prąd znamionowy, napięcie Ue, symbol.

**P1:** liczba modułów, kategoria użytkowania AC-21/AC-22, Icw, Icm, warunkowy
prąd zwarciowy Inc, sposób montażu, typ zacisku, IP.

**P2:** Ui, Uimp, częstotliwość, przekroje przewodów, moment dokręcania,
temperatura, trwałość, straty mocy i wymiary.

Anatomia nazwy:

`Rozłącznik izolacyjny modułowy {układ_biegunów} {prąd_A} {napięcie_VAC} Hager {symbol}`

Przykład:

`Rozłącznik izolacyjny modułowy 1P 63A 230V AC Hager SBN163`

### 4.6. Rozłącznik bezpiecznikowy — 1 produkt

**P0:** typ i wykonanie rozłącznika, rozmiar wkładki D01/D02, układ biegunów,
prąd znamionowy, napięcie AC/DC, symbol.

**P1:** dopuszczalne prądy wkładek, warunkowy prąd zwarciowy Icc, kategoria
użytkowania, sposób montażu i liczba modułów.

**P2:** przekroje przewodów, moment dokręcania, temperatura, wymiary i straty.

Anatomia nazwy:

`Rozłącznik bezpiecznikowy modułowy {rozmiar_wkładki} {układ_biegunów} {prąd_A} {napięcie} Hager {symbol}`

Przykład:

`Rozłącznik bezpiecznikowy modułowy D02 1P 63A 400V AC Hager L71M`

### 4.7. Styczniki modułowe — 2 produkty

**P0:** prąd znamionowy, napięcie sterowania cewki, konfiguracja styków
NO/NC, liczba biegunów, kategoria użytkowania, symbol.

**P1:** napięcie łączeniowe Ue, liczba modułów, prądy AC-1/AC-3, częstotliwość,
sposób montażu i przekroje przewodów.

**P2:** Ui, Uimp, pobór cewki, straty mocy, temperatura, trwałość, moment
dokręcania i wymiary.

Anatomia nazwy:

`Stycznik modułowy {prąd_A} {napięcie_cewki} {konfiguracja_styków} {kategoria_użytkowania} Hager {symbol}`

Przykłady:

- `Stycznik modułowy 25A 230V AC 2NO AC-7a/b Hager ESC225`
- `Stycznik modułowy 25A 230V AC 4NO AC-7a/b Hager ESC425`

`Napięcie sterowania cewki` musi być oddzielone od `Napięcia łączeniowego Ue`.
Konfigurację `2Z 0R` normalizujemy do `2NO` i opcjonalnie zachowujemy polski
synonim w opisie.

### 4.8. Ograniczniki przepięć SPD — 2 produkty

**P0:**

- typ SPD T1/T2/T1+T2;
- technologia, np. MOV, jeżeli potwierdzona;
- układ biegunów;
- rodzaj sieci TN-C/TN-S/TT;
- Iimp [kA], In [kA], Imax [kA] — tylko parametry występujące dla danego SKU;
- poziom ochrony Up [kV];
- maksymalne napięcie pracy ciągłej Uc [V];
- obecność styku zdalnej sygnalizacji;
- symbol.

**P1:** liczba modułów, Ue, maksymalne zabezpieczenie wstępne, sposób montażu,
IP, wymienność wkładów i zintegrowany bezpiecznik.

**P2:** temperatura, przekroje przewodów, moment dokręcania i wymiary.

Anatomia nazwy:

`Ogranicznik przepięć SPD {typ_SPD} {układ_biegunów} {sieć} {Iimp/In/Imax} Up≤{Up} Hager {symbol}`

Przykłady:

- `Ogranicznik przepięć SPD T1+T2 4P TN-S/TT Iimp 50kA In 50kA Up≤1,2kV Hager SPA931`
- `Ogranicznik przepięć SPD T1 3P TN-C Iimp 37,5kA Up≤1,5kV Hager SPA400`

Dla SPD trzeba rozróżnić wartość na tor i wartość całkowitą. Przykładowo
SPA400 jest opisywany jako 37,5 kA łącznie, a karta techniczna pokazuje 12,5 kA
dla toru. Pole musi jednoznacznie nazywać sposób pomiaru, aby nie porównywać
niezgodnych wartości.

### 4.9. Szyny grzebieniowe — 3 produkty

**P0:** wykonanie widełkowe/kołkowe, układ biegunów, przekrój szyny [mm²],
prąd znamionowy, liczba modułów, długość, kompatybilność, symbol.

**P1:** napięcie Ue, Uimp, typ złącza, położenie/układ połączenia.

**P2:** materiał, izolacja i dodatkowe wymiary, jeżeli karta je podaje.

Anatomia nazwy:

`Szyna grzebieniowa {typ_złącza} {układ_biegunów} {przekrój_mm2} {prąd_A} {liczba_modułów} Hager {symbol}`

Przykład:

`Szyna grzebieniowa widełkowa 3P 10mm² 63A 12 modułów Hager KDN363A`

Kompatybilność z RCCB/MCB/SPD jest ważniejsza od napięć Ui/Uimp i powinna być
widoczna w opisie oraz filtrach, ale nie zawsze zmieści się w nazwie.

## 5. Normalizacja danych i słownika

1. Kod producenta to SKU bez `/HAG`; suffix pozostaje w polu SKU.
2. Producent zawsze `Hager`, bez wariantów `HAGER` i `HAG`.
3. Jednostki zapisujemy spójnie: `16A`, `30mA`, `6kA`, `230V AC`, `10mm²`
   w nazwie oraz `16 A`, `30 mA`, `6 kA`, `230 V AC`, `10 mm²` w tabeli
   technicznej/opisie.
4. W wartościach liczbowych przechowujemy liczbę i jednostkę w nazwie pola;
   nie dokładamy drugi raz jednostki przy eksporcie.
5. `1P+N` i `3P+N` zachowujemy jako układy, a nie upraszczamy do `2P` i `4P`.
6. `Typ RCD: A/AC` nie może być mieszany z charakterystyką MCB/RCBO: B/C/D.
7. `Icn`, `Icu`, `Icw`, `Icm`, `Iimp`, `In`, `Imax`, `Up` i `Uc` są
   odrębnymi parametrami. Nie wolno łączyć ich w jedno pole `Zdolność zwarciowa`.
8. W nazwach unikamy zdublowanych synonimów, np. `B+C Typ T1+T2`; preferujemy
   aktualne oznaczenie techniczne `T1+T2`.

## 6. Zasady opisu produktu

Opis ma być generowany lokalnie na podstawie zatwierdzonych atrybutów. Nie
kopiujemy obecnego HTML ani tekstu producenta.

Docelowa struktura:

1. krótki akapit identyfikujący produkt i jego funkcję;
2. sekcja `Najważniejsze cechy` z 5–8 atrybutami P0/P1;
3. akapit o typowym zastosowaniu, bez obietnic wykraczających poza kartę;
4. `Dane techniczne` w uporządkowanej liście/tabeli;
5. kod producenta i zgodność z właściwą normą;
6. opcjonalnie zgodne akcesoria lub kompatybilne aparaty.

Reguły bezpieczeństwa merytorycznego:

- MCB chroni przed przeciążeniami i zwarciami, nie jest ochronnikiem przepięć;
- RCCB bez członu nadprądowego nie zastępuje zabezpieczenia nadprądowego;
- tylko RCBO łączy funkcję różnicową i nadprądową;
- zastosowania B/C/D opisujemy ostrożnie, bez dobierania zabezpieczenia za
  instalatora;
- nie deklarujemy przydatności do konkretnej instalacji wyłącznie na podstawie
  tytułu;
- nie używamy pustych fraz typu `najwyższa jakość`, `idealny wybór` ani
  niepotwierdzonych twierdzeń o bezpieczeństwie.

## 7. Znane błędy do obowiązkowej kontroli

1. `CDA440D/HAG` ma w tytule symbol `CDA440J`. Oba symbole istnieją w katalogach
   Hager. Polski katalog potwierdza `CDA440D` w tabeli technicznej na stronie 464,
   ale w tabeli produktowej na stronie 1058 występuje `CDA440J`. Wiersz wymaga
   potwierdzenia po EAN/ID, a nie automatycznej zmiany.
2. `CDC463J/HAG` ma w tytule `30hA`; powinno zostać zweryfikowane jako `30mA`.
3. `MBN110E/HAG` i `MBN116E/HAG` używają `1-fazowy` zamiast układu `1P`.
4. `ADA916D/HAG` jest obecnie nazwany jak RCCB, lecz oficjalna karta Hager
   klasyfikuje go jako RCBO z członem nadprądowym.
5. Opisy części MCB błędnie mówią o ochronie przeciwprzepięciowej lub
   „ochronie przed napięciem”; takie zdania trzeba usunąć.
6. `SPA931` zawiera jednocześnie `B+C` i `T1+T2`; nazwę trzeba uprościć bez
   utraty parametrów.

## 8. Etapy wdrożenia

### Etap A — mapa źródeł

- odczytać wszystkie 96 wierszy;
- zbudować `symbol = SKU bez /HAG`;
- dopasować 95 produktów bezpośrednio do tabel rodzin w polskim katalogu;
- dla `CDA440D/HAG` wykonać osobną kontrolę po EAN/ID;
- dopasować aktualną polską stronę i kartę Hager po symbolu;
- zapisać stronę PDF, URL, EAN, status dopasowania oraz datę weryfikacji;
- nie rozszerzać partii o 57 produktów z pliku pomocniczego.

### Etap B — konfiguracja rodzin

- dodać konfigurację `configs/categories/hager-aparatura-modulowa.yaml`;
- dodać słownik `dictionaries/hager_modular_apparatus_knowledge.yaml`;
- zapisać mapowanie prefiksów/rodzin: MBN/MCN/NDN, CDC/CDA, ADA/ADC, SBN,
  ESC, SPA, KDN i L7xM;
- prefiks służy do wstępnej klasyfikacji, ale nie zastępuje karty SKU.

### Etap C — wzbogacenie i normalizacja

- uzupełnić P0 dla 100% produktów;
- uzupełnić P1 wszędzie, gdzie oficjalna karta je podaje;
- zachować dane P2 w kontrolnym XLSX;
- każdemu polu nadać `source`, `confidence`, `status`;
- sprzeczności skierować do osobnego arkusza `DO_SPRAWDZENIA`.

### Etap D — generowanie nazw

- generować nazwę według anatomii rodziny;
- kod producenta zawsze na końcu po marce: `Hager {symbol}`;
- sprawdzać zgodność tytułu z atrybutami i brak duplikatów;
- przygotować po 3–5 nazw próbnych dla każdej rodziny przed pełnym przebiegiem.

### Etap E — generowanie opisów

- pisać każdy produkt na podstawie jego zatwierdzonych danych;
- wspólny układ HTML może być rodzinowy, ale wartości i tekst zastosowania muszą
  wynikać z konkretnego SKU;
- obecny `Content` zachować wyłącznie jako kolumnę kontrolną;
- nie kopiować producenta ani nie pozostawiać starego, błędnego HTML.

### Etap F — eksport i kontrola

Wygenerować:

- `output/hager_aparatura_modulowa_control.xlsx`;
- `output/hager_aparatura_modulowa_baselinker.csv`;
- `reports/hager_aparatura_modulowa/validation.md`.

## 9. Kryteria akceptacji

- dokładnie 96 wierszy wynikowych;
- 96/96 ma producenta, SKU, symbol, rodzaj produktu i kategorię;
- 100% atrybutów P0 kompletnych lub jawnie oznaczonych `DO_SPRAWDZENIA`;
- każda nazwa zawiera właściwy symbol, a żaden symbol z innego wariantu;
- brak mieszania RCCB z RCBO;
- brak mieszania `Typ RCD A/AC` z charakterystyką `B/C/D`;
- zgodność wartości w nazwie, atrybutach i opisie;
- wszystkie jednostki znormalizowane;
- brak powtórzeń marki, symbolu i parametrów w nazwie;
- brak błędnych twierdzeń o funkcji ochronnej;
- kategoria pochodzi z lokalnej taksonomii Alkan, nie z katalogu producenta;
- źródło i status zapisane dla każdego wzbogaconego parametru;
- próbka każdej rodziny zatwierdzona przed wygenerowaniem pełnej partii.

## 10. Źródła standardów i modelu atrybutów

- [ETIM EC000042 — Wyłącznik nadprądowy](https://viewer.etim-international.com/class/EC000042?lang=pl-PL)
  potwierdza m.in. charakterystykę, liczbę biegunów, prąd, napięcie, Icn/Icu,
  liczbę modułów, IP i przekroje przewodów jako odrębne cechy.
- [IEC 61008-1:2024](https://webstore.iec.ch/en/publication/67980) rozdziela RCCB
  bez integralnego zabezpieczenia nadprądowego od RCBO.
- [IEC 61009-1:2024](https://webstore.iec.ch/en/publication/67981) określa RCBO
  jako wyłączniki różnicowoprądowe z integralnym zabezpieczeniem nadprądowym.
- [IEC 60898-1:2015](https://webstore.iec.ch/en/publication/21972) obejmuje
  wyłączniki nadprądowe do instalacji domowych i podobnych.
- [IEC 60947-3:2020](https://webstore.iec.ch/en/publication/59785) obejmuje
  rozłączniki, rozłączniki izolacyjne i aparaty bezpiecznikowe.
- [IEC 61095:2023](https://webstore.iec.ch/en/publication/72007) obejmuje
  styczniki elektromechaniczne i kategorie użytkowania AC-7.
- [IEC 61643-11:2025](https://webstore.iec.ch/en/publication/65314) obejmuje SPD
  dla obwodów niskiego napięcia AC.
