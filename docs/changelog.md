# Zmiany w AlkanInatorze

Najważniejsze zmiany dla osób korzystających z AlkanInatora, od najnowszych. Jak z nich korzystać: [instrukcja](instrukcja.md).

## 2026-11-18

- **Nazwy do Ekonoma – nowy format pliku wynikowego.** Plik z gotowymi nazwami ma teraz kodowanie UTF-8 z BOM i tylko trzy kolumny: Kod towaru, Nazwa towaru i Nazwa dodatkowa (bez LP, kodu producenta, jednostki, ceny i waluty). Dzięki UTF-8 znaki takie jak „°” czy „≤” nie zamieniają się już w „?”. Pliki z Ekonoma na wejściu i raport kontrolny bez zmian.
- **Wiedza o sklepie pochodzi teraz z produkcji alkan.pl.** Po podłączeniu sklepu wiedza AlkanInatora nie była przebudowana. Słownik katalogu (kategorie, atrybuty, producenci, serie, kolory) pochodził z eksportu z maja, a słowniki terminów, kategorii, tagów i grup wariantów z lokalnej kopii sklepu. Teraz wszystko jest z produkcji z 18 listopada: 12 152 produkty, 7738 wartości atrybutów, 150 atrybutów globalnych, 490 grup wariantów. Lista atrybutów, które wolno wysyłać do Baselinkera i Woo, ma 17 nowych pozycji (m.in. Gwarancja, Klosz, Typ gwintu, Stopień odporności [IK], Źródło światła). Ubyły z niej 4 atrybuty, których sklep nie ma: Bateria, Biegunowość, Ilość styków zwiernych/rozwiernych i Znamionowy przekrój poprzeczny. Podgląd „Wgraj zmiany do sklepu” i uzupełnianie tagów wariantów korzystają z grup wariantów z produkcji. Całość odświeża się jednym poleceniem, najlepiej poza godzinami ruchu w sklepie.
- **Śmieciowe wartości w sklepie nie są traktowane jak poprawne.** Produkcja ma wartości po dawnych błędach importów: „Antenowe, Komputerowe”, „Kaszmirowy”, opisy i zdania wpisane jako materiał („Meets the product standard's requirements.”, „PKWiU”). Kontrola wartości atrybutów ich nie przyjmuje i zamienia na poprawne wartości sklepu („Kaszmir”, „Antenowe” i „Komputerowe” osobno). „IP20” zmienia się teraz zawsze na „IP 20”, z którego korzysta 5329 produktów, choć w sklepie 328 produktów ma jeszcze „IP20”. Nowy raport „terminy do sprzątnięcia” zbiera 30 takich śmieci i 2849 wartości bez żadnego produktu, do decyzji, co usunąć w Woo. Nic nie jest usuwane automatycznie.

## 2026-11-17

- **Nazwy do Ekonoma – Kobi: nazwa własna przed marką.** Seria (nazwa własna produktu) stoi teraz przed marką, a marka zamyka nazwę: „Taśma LED COB 320LED/m 50m 6500K 12V IP20 Tramo Kobi Premium” (zamiast „… IP20 Kobi Premium Tramo”), „Oprawa uliczna LED 50W 4000K z czujnikiem Cyoto Kobi Premium”, „Plafon LED 12W 1140lm CCT IP54 czarny Nivera Square Kobi Premium”, „Szynoprzewód 3-obwodowy 2m czarny Nextrack NT Kobi”, „Bank energii 600W Power Box Kobi”. Żarówki smart mają protokół przed marką („Żarówka LED E27 9,5W RGB A60 WiFi Kobi”). Zmienia się 1161 nazw Kobi; statusy i powtórzenia bez zmian. Inni producenci bez zmian.
- **Nazwy do Ekonoma – Kobi: poprawki po przeglądzie.** Zestawy HERMETIC mają w nazwie tuby i ich moc („Oprawa hermetyczna z tubami LED 2x120cm 2x18W 3600lm 4000K IP65 szara Hermetic Kobi” zamiast „Oprawa hermetyczna LED 128cm 18W …”). Łączniki szynoprzewodu Nextrack NT mają z powrotem „zewnętrzny” z nazwy w Ekonomie, a stronę małymi literami („… 3-obwodowy T zewnętrzny prawy biały Nextrack NT Kobi”). Grzejniki, termowentylatory, wentylatory i klimatory Kobi Windstar mają nazwę według tego samego wzoru co oprawy: „Grzejnik konwektorowy 2000W biały Sondo Turbo Kobi Windstar”, „Wentylator podłogowy 40W biały Viento-R Kobi Windstar”. Akumulatory piszą, do czego pasują („Akumulator 18650 7,4V 5,4Ah do Solar Street 15W Kobi”). Nexforce N2 52W nie ma już „Neutralna” w serii i ma szary kolor obudowy z feedu (wcześniej „biała” z opisu barwy światła). Poprawione też „(SATYNA)” w oczku Rybie Oko i wiszący myślnik w kostce K006. Zmienia się 48 nazw, a 24 ostrzeżenia znikają.

## 2026-11-16

- **Nazwy do Ekonoma – Kobi, GTV i Kanlux: wróciły nazwy własne.** Przegląd porównał każdą nową nazwę ze starą nazwą z Ekonoma i pokazał, że część nazw serii i linii wypadała. Najwięcej w Kobi: naświetlacze, panele, plafony, oprawy uliczne, liniowe, hermetyczne, high bay, reflektory szynowe, downlighty i taśmy miały same parametry i markę. Teraz seria stoi na końcu, za marką, tak jak u innych producentów: „Plafon LED 12W 1140lm CCT IP54 czarny Kobi Premium Nivera Square”, „Naświetlacz LED 30W 2550lm 4000K 110° IP65 czarny Kobi Premium Delyo MHN”, „Panel LED 60x60 25-36-40W 6400lm CCT UGR19 biały Kobi Pro Capri G3”, „Reflektor szynowy 10W 900lm CCT 36° biały Kobi Pro Nextrack Core”, „Oprawa uliczna LED 100W 14000lm 4000K IP66 szara Kobi Pro Vespa”. Akcesoria szynoprzewodu mają z powrotem system („Szynoprzewód 3-obwodowy 2m czarny Kobi Nextrack NT”), a zasilacze do paneli – kod („Zasilacz LED do paneli 36W LF-GIF040YS900H Kobi”). Zmienia się 643 nazw Kobi. GTV: linie żarówek („Żarówka LED E14 3W 165lm 2200K świeczka filament GTV Spin”, a także Glow, Grid, Ceramic, Slim, Wave, Vintage), tuby SHL, oprawy Tiago i Escada oraz zasilacze Boke – i nie są już ucięte („… DALI-PUSH-1-10V regulowany 250-700mA SELV” zamiast „DALI-PUSH-1- … 3-”, „Osl Slim” zamiast „Osl SLIM-”); zmienia się 64 nazw. Kanlux miał serie w porządku (kod modelu zostaje); doszły tylko linie tub („Tuba LED T8 120cm 18W 2880lm 4000K Kanlux GLASSv4”) i wyłączników IDEAL TS – 23 nazwy. Statusy i powtórzenia bez zmian. Lista wszystkich zmienionych nazw obok poprzednich jest w raporcie przeglądu.

## 2026-11-10

- **Nazwy do Ekonoma – F-Elektro, runda 1.** 2946 nazw (katalog i 21 nowych kartotek). Rozdzielnice budowlane z wyjściami, zabezpieczeniami, przełącznikiem, zasilaniem i serią z liczbą modułów na końcu: „Rozdzielnica 16/5 2x230V z wył. nadprądowymi i RCD 40A R-240 8M F-Elektro”, „Rozdzielnica 2x32/5 2x230V z przełącznikiem L-0-P 32A z wtyczką 32/5 na przewodzie 5x4mm2 1,5m R-240 F-Elektro”. Przedłużacze bębnowe według wspólnego schematu („Przedłużacz bębnowy trójfazowy 2x32/5 1x230V 40m 5x4mm2 guma H07RN-F MAKARA-460 F-Elektro”, „Przedłużacz bębnowy 4 gn. 30m 3x1,5mm2 PVC FILO-230 F-Elektro”), złącza siłowe („Gniazdo siłowe tablicowe skośne 16A 5P z puszką F-Elektro”) i osprzęt z opisem z Ekonoma („Szybkozłączka uniwersalna 0,14-4mm2 3 styki F-Elektro op. 10 szt.”). Ekonom miał wiele par kartotek z identyczną nazwą. Według kart produktów w hurtowni Onninen to wersje o innej szczelności, więc nazwy dostały „IP65” (rozdzielnice LAGO, CAJA, ULISSE, FEMO, LEGATO, TRIBOX i przedłużacze siłowe) albo „IP44” / „IP67” (złącza siłowe), a rozdzielnice RB-300 – przełącznik, którego nazwa w Ekonomie nie podawała. Powtarza się jeszcze 16 par (głównie złącza siłowe, dla których karta nie podaje szczelności) – do decyzji.
- **Nazwy do Ekonoma – Pawbol, runda 1.** 2408 nazw (katalog i 288 nowych kartotek). Źródłem jest cennik Pawbol 2026 i strona producenta, a nie skrócony opis z Ekonoma, więc nazwy są pełne i jednolite. Rozdzielnice budowlane z wyjściami, zabezpieczeniami i serią: „Rozdzielnica 2x32/5 4x230V z wył. nadprądowymi i RCD 40A IP44 R-BOX LUX 320 13S Pawbol”, „Rozdzielnica 32/5 16/5 2x230V z rączką IP44 R-BOX 240 4S Pawbol”. Przedłużacze bębnowe według wspólnego schematu („Przedłużacz bębnowy fi300 4 gn. 25m 3x1,5mm2 guma IP44 z uziemieniem z termikiem Pawbol”), oprawy z mocą i barwą tylko przy LED („Plafon LED 13W 4000K klosz mleczny srebrny z czujnikiem ruchu IP54 Pantera Pawbol”, „Oprawa kanałowa E27 klosz przezroczysty szklany z siatką metalową biała IP44 Owal 60 Pawbol”), osprzęt hermetyczny Beta („Wyłącznik jednobiegunowy z gniazdem podwójnym biały klapka biała IP44 Beta Pawbol”), puszki, rury, kanały, listwy zaciskowe, uchwyty („Uchwyt kablowy z gwoździem YDYp 3x1,5 10/6 płaski biały Pawbol op. 100 szt.”) i osprzęt odgromowy. Błędy cennika poprawione według kodu produktu i strony producenta, z uwagą w raporcie (np. zestawy Beta z podwójnym gniazdem, kolor listwy LZ 10/Ż, wtyczka B.1923). Powtarza się 15 par – to stare i nowe kartoteki tego samego produktu (np. przewody z wtyczką D.3507 i D.3504-3) – do decyzji.
- **Nazwy do Ekonoma – PCE, runda 1.** 577 nazw. Złącza siłowe z prądem, liczbą biegunów, napięciem, szczelnością i serią: „Wtyczka przenośna 16A 5P 400V IP44 Shark TT PCE”, „Gniazdo tablicowe skośne 32A 4P 400V IP44 z puszką podtynkową PCE”, „Gniazdo z przełącznikiem L-0-P 16A 5P 400V + 2x230V IP44 Combo-Pol PCE”. Rozdzielnice budowlane z wyjściami, zabezpieczeniami i nazwą: „Rozdzielnica stacjonarna 32/5 2x230V z RCD 40A IP44 Adamów PCE”, „Rozdzielnica przenośna 3x32/5 z wtyczką 32/5 IP44 Pieszyce PCE”. Przedłużacze bębnowe według wspólnego schematu („Przedłużacz bębnowy 4 gn. 20m 3x1,5mm2 guma IP44 z uziemieniem z termikiem XREEL PCE”), do tego obudowy Argenta, rozdzielnice modułowe, PV i ppoż., puszki Light, rozłączniki i lampy. Błędy opisów w Ekonomie poprawione według danych producenta i hurtowni, z uwagą w raporcie: wtyczki tablicowe 613-6, 623-6, 6132-6 i 6232-6 to 230V (w Ekonomie 400V), gniazdo 75132-6 to IP67 (w Ekonomie IP44), puszki Light -171 mają wymiar 155x179x100 (w Ekonomie skopiowany z -161), a wtyczka 8615-6 ma dwa wprowadzenia przewodu. Wszystkie nazwy porównane z bazą produktów PCE: rozdzielnice PV mają IP65, lampy – szczelność z karty producenta (np. „IPX6”, „IP65”), gniazda z rozłącznikiem „IP44/54”, a gniazda Taurus2 25711 i 25712 – klapkę i zatyczkę tak jak u producenta (w Ekonomie były zamienione). Rozdzielnice Star różnią się liczbą modułów („9 mod.”, „12 mod.”). Jedyna powtórzona nazwa to rozdzielnice Ecoline 90EKO006 i 90EKO010 – producent podaje dla nich identyczne parametry.
- **Nazwy do Ekonoma – Plastrol, runda 1.** 78 nazw. Przedłużacze według wspólnego schematu: „Przedłużacz bębnowy 4 gn. 30m 3x1,5mm2 guma IP44 z uziemieniem Professional Line Plastrol”, „Przedłużacz budowlany 3 gn. 10m 3x2,5mm2 guma IP44 Schuko czarno-czerwony Plastrol”, „Przedłużacz 1 gn. 10m 2x1mm2 bez uziemienia pomarańczowy Eco Plastrol”. Do tego gniazda i wtyczki rozbieralne, płytki odgałęźne („Płytka odgałęźna 5-torowa 16mm2 czarna Plastrol”), lampy warsztatowe i latarki. Produkty KEL mają markę KEL („Naświetlacz akumulatorowy z głośnikiem 40W 6500lm 6400mAh IPX4 KEL”). Dwie kartoteki wtyczki WT-16UP mają w Ekonomie identyczną nazwę – do decyzji.
- **Nazwy do Ekonoma – Elgotech, runda 1.** 1254 nazwy (katalog i 130 nowych kartotek). Przedłużacze według wspólnego schematu: „Przedłużacz 5 gn. 2m z uziemieniem z wyłącznikiem czarny Economic Elgotech”, „Przedłużacz bębnowy mały 4 gn. 20m 3x1mm2 PVC z uziemieniem Elgotech”, „Przedłużacz bębnowy 2 gn. + 1 gn. siłowe 20m 5x2,5mm2 guma H07RN-F IP44 Elgotech”. Do tego gniazda przenośne, rozgałęźniki, wtyczki i przewody z wtyczką („Przewód z wtyczką zalewaną kątową Unischuko 1m 3x1mm2 PVC czarny Elgotech”). Kolor, którego nazwa w Ekonomie nie podaje, pochodzi z końcówki kodu (-2 czarny, -5 pomarańczowy, -7 zielony…). 17 kartotek, w których kod mówi co innego niż nazwa (np. długość), jest do sprawdzenia.
- **Nazwy do Ekonoma – Elgotech, poprawki według katalogu producenta.** Brakujący przekrój i kolor uzupełnione z katalogu Elgotech („Przedłużacz 3 gn. 3m 3x1mm2 z uziemieniem biały Classic Elgotech”). Przedłużacze PK-10 i PS-10-5 to „Przedłużacz ogrodowy … pomarańczowy”, bębny PZB-1111 – „metalowy”, rozgałęźniki „trójkątny” (RT-30CA) i „płaski” (R-30CA), a przedłużacze przeciwprzepięciowe mają na końcu model (PSF3, PSF4, MPSFU…), bo różnią się ochroną. Takich samych nazw zostało 14 grup – wszystkie przy kartotekach, w których opis w Ekonomie nie zgadza się z kodem.

## 2026-11-09

- **Nazwy do Ekonoma – Kanlux: „hermetyczna” zamiast „pyłoszczelna”.** Oprawy, które Kanlux nazywa pyłoszczelnymi, mają teraz w nazwie „Oprawa hermetyczna”, tak jak u innych producentów: „Oprawa hermetyczna LED 58cm 19W 2400lm 4000K IP65 szara Kanlux MAH HI RYF”, „Oprawa hermetyczna pod tuby 127cm IP65 2xG13 szara Kanlux DICHT”, a akcesoria to „Akcesorium oprawy hermetycznej”. Zmienia się 92 nazwy (87 nowych kartotek i 5 z bazy), reszta bez zmian.

## 2026-11-05

- **Nazwy do Ekonoma – GTV: nowe kartoteki.** Zbudowane razem z bazą 490 nowych kartotek. Poprawione znaki, które Ekonom gubi (kodowanie): „Smarownica 400 cm3”, „Szczypce do zaciskania tulejek 6-16 mm2”, bez „??” w nazwach narzędzi Högert, i lumeny żarówki A60 („820lm” zamiast „60820lm”).
- **Nazwy do Ekonoma – Kobi: nowe kartoteki.** Zbudowane razem z bazą 391 nowych kartotek. Oprawy hermetyczne HERMETIC / HERMIC to oprawy pod tuby („Oprawa hermetyczna pod tuby LED 2x120cm IP65 szara Kobi”), VESPA PRO to oprawa uliczna, plafony na żarówki nazywają się „Plafon 2xE27 chrom Kobi Samira S”, oczka mają kształt i model („Downlight kwadratowy stały czarny OH-20 Kobi”), a stare nazwy WIELKIMI LITERAMI dostały rodzaj na początku („Kinkiet 1xG9 gipsowy biały Kobi Bueno Sześcian”). Przy okazji poprawione „do nadbudowania” (zamiast „do nabudowania”), „Lampa solarna” bez drugiego „solarna” i kształt opraw po polsku („okrągła”, „kwadratowa” zamiast Round / Square). Oprawy HPL, NEXLINE i reflektory NEXTRACK NT mają markę Kobi Pro zamiast „Nexus” z feedu. KOBI POWER BOX to „Bank energii 600W Kobi Power Box”.
- **Nazwy do Ekonoma – Kanlux: rozgałęźniki, wtyczki i gniazda przenośne.** Sześć kartotek, które miały w nazwie tylko grupę „Rozgałęźniki, gniazda, wtyczki”, dostało nazwy z danych ze strony Kanlux: „Rozgałęźnik 3 gn. z uziemieniem biały ZARK KRT/2-1 Kanlux”, „Wtyczka z uziemieniem 16A IP44 czarna GRIP W/US Kanlux”, „Gniazdo przenośne 1 gn. Schuko IP44 czarne GRIP GP/S Kanlux”. Reszta nazw Kanlux bez zmian.
- **Nazwy do Ekonoma – przedłużacze według schematu V3.** W przedłużaczach domowych długość i przekrój stoją teraz zaraz za liczbą gniazd, przed uziemieniem i wyłącznikiem: „Przedłużacz 3 gn. 1,5m 1,0mm2 z uziemieniem z wyłącznikiem Bemko”, „Przedłużacz 3 gn. 5m z uziemieniem z wyłącznikiem czarny Linea Kobi”, „Przedłużacz 4 gn. 3m Schuko GTV”. Zmienia się 127 nazw (Kobi 55, Bemko 66, GTV 6). Przedłużacze bębnowe zostają bez zmian – uziemienie albo Schuko nadal na końcu, przed producentem („Przedłużacz bębnowy 4 gn. 15m 3x1,5mm2 Schuko Bemko”).
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
