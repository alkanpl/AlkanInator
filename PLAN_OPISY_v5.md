# PLAN: Opisy SEO v5 — runda krytyki "12 punktów" (Do poprawy.txt)

Stan wejściowy: rundy v1-v4 wdrożone. Ogrodowa + wszystkie rodziny mają już
konkretne benefity (`benefits_luminaire`), naprawiony błąd czujnika (`has_feature`),
fraza 3x, próbka 30/30 OK, 165 testów. Plik: `src/generate_product_descriptions.py`.
Próbka: `reports/opisy_probka_2026-06-25.{html,xlsx}`.

UWAGA o sprzecznościach z poprzednią rundą: user w tej krytyce ODWRACA dwie wcześniejsze
decyzje (poniżej P8 i P9). Działamy wg NAJNOWSZEJ krytyki (pisemnej).

## Fakt do zakomunikowania (NIE błąd)
- Linia `<p><em>rodzina: ... | walidacja: OK</em></p>` jest TYLKO w podglądzie HTML
  (`write_sample_report`, l. ~376), NIE ma jej w kolumnie `opis_html` idącej do sklepu.
  To metadane do przeglądu. Nie usuwać z opisu (bo go tam nie ma); ewentualnie zostawić
  w podglądzie. Wyjaśnić userowi.

## 12 punktów krytyki -> przyczyny -> fix

### Helpery wspólne (DODAĆ po `has_feature`, ~l.1224)
- `LIGHT_FAMILIES = {high_bay, panel, hermetyczna, plafon, naswietlacz, oprawa, ogrodowa, ogolny}`
- `is_light_product(family)` -> family in LIGHT_FAMILIES (False dla ramka/zasilacz/akcesorium)
- `subject_for(family)` -> "Ta ramka"/"Ten zasilacz"/"To akcesorium"/"Ten panel"/"Ten plafon"/
  "Ten naświetlacz"/"Ta oprawa"/"Ta lampa"(ogrodowa)/"Ten produkt"(domyślnie). Do akapitu po H2
  i zamknięć BEZ frazy.
- `ip_meaning(ip)` -> konkretne znaczenie po cyfrach (re.sub \D): 
  20="przeznaczenie do suchych wnętrz, bez narażenia na wodę i podwyższoną wilgotność";
  23/40 podobnie suche; 44="ochronę przed bryzgami wody z każdej strony oraz ciałami obcymi >1 mm";
  54="ochronę przed pyłem i zachlapaniem wodą"; 65="pełną pyłoszczelność i odporność na strugę wody,
  co pozwala na montaż na zewnątrz"; 66/67 mocniej; default="warunki podane w klasie szczelności".
- `dim_compact(attrs)` -> "595 x 595 x 25 mm" (lub "śr. 430 mm"); używać w PROZIE (nie w spec).
  Bierze Wymiary, inaczej Długość/Szerokość/Wysokość (strip " mm", join " x ", +" mm"); fallback średnica.

### P1 + P6 — powtarzalne zakończenia (kopiuj-wklej)  [najważniejsze wg usera]
Przyczyna: `CLOSING_BUILDERS` = tylko 3 generyczne zamknięcia, mocno się powtarzają.
Fix: 5-6 zamknięć, ZALEŻNYCH od rodziny, z wariancją seedem. Każde zawiera frazę kluczową
(to będzie 3. wystąpienie frazy - patrz P9). Zmienić sygnaturę na `closing(keyword, attrs, family)`
i wywołanie w `build_description_html`. Treść per rodzina (oprawa świecąca vs ramka vs zasilacz
vs akcesorium vs ogrodowa).

### P2 + P3 — FAQ szablonowe i BŁĘDNE LOGICZNIE dla nie-opraw  [krytyczne]
Przyczyna: `faq_candidates` (l.~1320) ma generyczne "Gdzie najlepiej sprawdzi się... doświetlić
{after_place}" i "Na co zwrócić uwagę przed zakupem" dla WSZYSTKICH rodzin. Ramka/zasilacz/
czujnik NIE świecą -> "doświetlić przestrzeń" to błąd.
Fix:
- Gałąź `if not is_light_product(family)`: pytania o KOMPATYBILNOŚĆ/montaż/serię, NIE o świecenie.
  - ramka: "Do jakich paneli pasuje ta ramka?" (format, np. 60x60), "Kiedy potrzebna jest ramka
    natynkowa?" (sufit pełny/bez kasetonu), "Z czego wykonana / jaki kolor".
  - zasilacz: "Do jakich opraw pasuje ten zasilacz?" (moc/napięcie/seria), "Czy obsługuje
    ściemnianie?" (1-10V/DALI jeśli atrybut), "Gdzie zamontować zasilacz?" (dostęp serwisowy, IP).
  - czujnik (akcesorium "czujnik"): "Z jaką oprawą współpracuje?" (seria MAH PRO), "Jak
    automatyzuje pracę oprawy?", "Gdzie montować czujnik?".
- Usunąć/ograniczyć generyczne "Na co zwrócić uwagę przed zakupem" (za ogólne) — zastąpić
  bardziej konkretnymi lub dać tylko jako ostatni fallback.
- FAQ IP: użyć `ip_meaning` (patrz P4).
- Czujnik FAQ dla rodziny ogrodowa/naswietlacz już poprawione (furtka/wejście...).

### P3 — produkty nie-świecące traktowane jak oprawy (proza)
Przyczyna: `build_after_h2` (l.~1289) gdy brak faktów świetlnych -> `light_sentence` =
"Moc 40W pozwala dobrać oprawę do wielkości przestrzeni" (BŁĄD dla zasilacza). Intro/closing
też miejscami świetlne.
Fix:
- `build_after_h2`: zacznij od `subject_for(family)` (BEZ frazy — P9). 
  - light product: fakty z `secondary_fact_sentences` (świetlne); brak -> `light_sentence`.
  - non-light: NOWA `nonlight_fact_sentences(attrs, family)`:
    - ramka: zgodność z formatem panelu (dim_compact), kolor, materiał, montaż natynkowy poza
      sufitem kasetonowym; ZERO o świeceniu.
    - zasilacz: moc/napięcie = potwierdzenie KOMPATYBILNOŚCI z panelem (nie dobór do przestrzeni),
      ściemnianie 1-10V/DALI jeśli atrybut, IP/miejsce montażu.
    - akcesorium/czujnik: zgodność z serią/oprawą, automatyzacja (czujnik), montaż.
- `benefits_zasilacz` już OK (kompatybilność). Sprawdzić, że `intro_zasilacz_*` nie mówią
  "dobrać oprawę do przestrzeni" (obecnie OK: "dobrać do panelu/systemu").

### P4 — mylące IP20/IP44 ("odporność na pył i wilgoć")  [globalnie]
Przyczyna: stałe frazy "Stopień ochrony IP X określa odporność na pył i wilgoć" w FAQ
(l.~1356/1358) i benefitach (`benefits_luminaire` ctx, `benefit_common`).
Fix: wszędzie używać `ip_meaning(ip)`:
- FAQ "W jakich warunkach montować": `f"Stopień ochrony {ip} oznacza {ip_meaning(ip)}. ..."`.
- FAQ "Czy nadaje się na zewnątrz" (ogrodowa/naswietlacz) — IP65 OK, ale dla IP20/44 NIE pisać
  że nadaje się na zewnątrz.
- benefits_luminaire `raw["ip"]` -> `f"Stopień ochrony {ip} - oznacza {ip_meaning(ip)}."`
- benefit_common IP (jeśli jeszcze używane) -> ip_meaning.

### P5 — "dopasuj do miejsca montażu" zamiast KONKRETNYCH miejsc
Fix: `after_place(family)` rozbudować na konkretne listy (już częściowo). W FAQ "Gdzie sprawdzi się"
(tylko dla light) dać konkretnie: hermetyczna="korytarz techniczny, garaż, warsztat, zaplecze";
plafon="korytarz, wejście, klatka, pomieszczenie gospodarcze"; itd. Unikać "wybraną przestrzeń".

### P7 — Panel BLINGO uproszczony vs ręczny wzorzec
Ograniczenie: fakty typu "6 elips McAdama / L70B50 / zasilacz stałoprądowy / 115° / stal /
27 mm" NIE są w atrybutach -> nie wolno zmyślać (anti-halucynacja). 
Fix realny: pogłębić panel tym, co MAMY (skuteczność lm/W, Ra, UGR, ściemnianie, profil =
dim_compact wysokość, klasa energ.). H3 i akapit po H2 bardziej "panelowe". Resztę
zakomunikować userowi jako limit danych (ewentualnie: dodać te pola do słownika/Parametry).

### P8 — `<li>` ma zaczynać się od myślnika  [ODWRÓCENIE poprzedniej decyzji]
Fix: `render_benefit_item` -> `f"  <li>- <strong>{lead}</strong> - {rest}</li>"` (myślnik na
początku) oraz `f"  <li>- {item}</li>"` (już jest). 
ZAKTUALIZOWAĆ TEST `test_benefit_item_bolds_feature_before_dash` na nowy format.
Sprawdzić walidator listy (`test_validator_detects_..._bad_list_format`) czy nie wymaga innego.

### P9 — fraza w 2. <p> (akapit po H2)  [ODWRÓCENIE poprzedniej decyzji]
User: po H2 drugi akapit NIE może zaczynać się frazą. 
Fix: 
- `build_after_h2`: zaczynać od `subject_for(family)`, BEZ frazy kluczowej.
- 3. wystąpienie frazy przenieść do ZAMKNIĘCIA (closing zawiera frazę).
- Liczność frazy nadal = 3 (intro + H2 + closing).
- `validate_generated_description`: dodać check, że akapit po H2 NIE zawiera frazy (regresja),
  utrzymać keyword_count == 3.

### P10 — niespójne jednostki (38W vs 38 W)
Fix: w PROZIE generatora wszędzie ze spacją: "38 W", "4560 lm", "4000 K", "120 lm/W".
- `light_sentence`: "Moc {p} W i strumień {f} lm...", "Barwa {t} K...".
- intra/secondary gdzie jest "{x}W/{x}lm/{x}K" bez spacji -> dodać spację.
- H1 (nazwa produktu) jest spoza generatora (pipeline tytułów) — NIE ruszamy tu; ewentualnie
  osobno. Zakomunikować, że h1 = nazwa katalogowa.

### P11 — wymiary mechaniczne ("długość 595 mm, szerokość...")
Fix: w prozie/benefitach używać `dim_compact` ("595 x 595 x 25 mm"). `dim()` zostaje dla
miejsc, gdzie potrzebny opisowy wariant, ale benefit "Wymiary {d}" -> dim_compact.

### P12 — H3 zbyt ogólne i powtarzalne
Fix: `H3_DEV_HEADINGS` -> `H3_DEV_BY_FAMILY` (dict per rodzina), wariancja seedem. Przykłady:
- ogrodowa: "Oświetlenie ogrodu bez prowadzenia przewodów", "Światło przy wejściu, ścieżce i tarasie"
- panel: "Równe światło do biura i ciągów komunikacyjnych"
- ramka: "Montaż panelu poza sufitem kasetonowym"
- zasilacz: "Zasilanie i kompletacja paneli", "Dobór zasilacza do oprawy"
- hermetyczna: "Oświetlenie techniczne do garażu, warsztatu i zaplecza"
- high_bay: "Mocne światło do hal i magazynów"
- plafon: "Światło do korytarzy, wejść i pomieszczeń pomocniczych"
- naswietlacz: "Doświetlenie elewacji, podjazdu i placu"
- akcesorium/czujnik: "Dopasowanie do oprawy i serii"

## Kolejność wdrożenia (bezpieczna)
1. Helpery (LIGHT_FAMILIES, is_light_product, subject_for, ip_meaning, dim_compact).
2. render_benefit_item + update testu (P8).  
3. ip_meaning w benefits_luminaire + FAQ (P4).
4. build_after_h2 + nonlight_fact_sentences + subject (P3, P9).
5. CLOSING_BUILDERS per rodzina + keyword (P1, P6, P9) + zmiana wywołania.
6. faq_candidates per typ produktu (P2, P3, P5) + ip_meaning.
7. H3_DEV_BY_FAMILY (P12).
8. Jednostki ze spacją + dim_compact w prozie (P10, P11).
9. validate_generated_description: keyword==3 i brak frazy w akapicie po H2 (P9).
10. Testy (`py -m unittest discover -s tests -p "test_*.py"`), regeneracja próbki
    `py src/generate_product_descriptions.py --mode generate --input
    output/kanlux_baselinker_attributes_update_2026-06-11.csv --output
    reports/opisy_probka_<data>.xlsx --sample 30 --sample-html reports/opisy_probka_<data>.html`.
11. Weryfikacja na próbce: ramka/zasilacz/czujnik (brak "doświetlić"), IP20/44 poprawne,
    fraza poza 2. <p>, myślniki w <li>, jednostki ze spacją, różne zakończenia/H3/FAQ.

## Po akceptacji
Cały katalog: kolumna Content w pliku Woo + osobny plik kontrolny (`add_descriptions`).

## Pliki
- `src/generate_product_descriptions.py` (rdzeń)
- `tests/test_generate_product_descriptions.py` (update test render_benefit_item + ew. nowe:
  ip_meaning, brak frazy w akapicie po H2, FAQ nie-opraw bez "doświetlić")
- (opcjonalnie) `dictionaries/seo_description_knowledge.yaml` — mapy IP/per-rodzina.
