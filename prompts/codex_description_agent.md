# Codex Agent: indywidualne opisy produktów Alkan

## Cel

Twórz opisy HTML SEO jako autor, po jednym produkcie. Skrypty lokalne służą
wyłącznie do przygotowania briefu, zapisania review i walidacji. Nie wolno
generować prozy skryptem, pętlą ani zestawem gotowych akapitów per typ produktu.

Główne repo:

```text
C:\Users\Handlowiec\Desktop\AlkanInator
```

Finalne pliki zapisuj w `reports\descriptions`, nigdy w `.codex\worktrees`.

## Źródło prawdy

Przeczytaj z briefu XLSX:

- `Produkt`: pełna nazwa, SKU, EAN i kategoria;
- `Product facts`: jedyne fakty dozwolone jako cechy produktu;
- `Compatibility facts`: wyłącznie informacje typu „pasuje do”;
- `Rejected facts`: dane, których nie wolno przypisać produktowi;
- `Prompt dla Codex`: instrukcja dla konkretnego SKU.

Nie pobieraj parametrów z surowego pliku produktowego. Przy akcesorium moc
oprawy może występować wyłącznie jako kontekst kompatybilności, nigdy jako moc
akcesorium.

## Proces jednego SKU

Jeżeli użytkownik nie podał SKU ani ścieżki, najpierw uruchom:

```powershell
py src\codex_description_agent.py --list-briefs
```

Komenda zwraca cały najnowszy zestaw. Jeżeli zwróci 10 briefów, wykonaj kolejno
wszystkie 10. Nie wybieraj starego pliku z głównego katalogu tylko dlatego, że
jest krótszą ścieżką albo pojawił się jako pierwszy w wynikach wyszukiwania.

1. Otwórz jeden brief z wybranej kolejki. Nie czytaj kolejnego przed ukończeniem obecnego.
2. Ustal główną frazę: zwykle typ produktu z pełnej nazwy.
3. Wybierz 2-4 najważniejsze fakty i przełóż je na rzeczywiste korzyści.
4. Ustal indywidualny kierunek tekstu: zastosowanie, montaż, światło,
   odporność, kompatybilność albo wzornictwo, zależnie od produktu.
5. Napisz HTML samodzielnie. Nie twórz Pythona składającego zdania.
6. Zapisz HTML, złóż review helperem i uruchom walidację.
7. Popraw wszystkie `ERROR` oraz `WARNING`.
8. Dopiero wtedy przejdź do następnego SKU.

## Standard SEO

- Zacznij pierwszy `<p>` od `<strong>pełnej nazwy produktu</strong>`.
- Pierwsze 2-3 zdania mają wyjaśnić, czym jest produkt, gdzie się sprawdza i
  jakie potwierdzone cechy są najważniejsze.
- Dodaj konkretne `<h2>` zawierające naturalną frazę produktową i sensowną
  korzyść lub wyróżnik. Akapit bezpośrednio po `<h2>` zaczyna się frazą
  kluczową i rozwija realne zastosowanie produktu.
- Dodaj `<h3>Najważniejsze cechy</h3>` oraz listę korzyści wynikających z faktów.
- Dodaj `<h3>Specyfikacja techniczna</h3>`. Każdy parametr zapisz jako
  `<li><strong>Nazwa:</strong> wartość</li>`.
- Bezpośrednio po liście specyfikacji dodaj osobny końcowy `<p>`. Podsumuj w nim
  najważniejszy praktyczny powód wyboru produktu i naturalnie zachęć do zakupu.
  Używaj wyłącznie potwierdzonych faktów. Opis nie może kończyć się na `</ul>`.
- Opcjonalnie dodaj sekcję zastosowań, jeśli wynika ona z typu produktu i nie
  wymaga wymyślania właściwości.
- Przy bogatych danych celuj w 1500-2200 znaków bez spacji. Przy skromnym
  akcesorium napisz krócej i konkretniej zamiast sztucznie wypełniać tekst.

Przykład AROT od użytkownika jest wzorcem rytmu, konkretności i hierarchii
HTML. Przeczytaj `prompts\examples\good_description_arot.html` przed pisaniem.
Nie przenoś z niego do produktów Kanlux materiału, IP, temperatur, odporności,
zastosowań ani innych parametrów.

## FAQ schema

Na końcu opisu dodaj widoczną sekcję FAQ oraz odpowiadający jej blok danych
strukturalnych `FAQPage` w formacie JSON-LD.

Zasady:

- FAQ musi być widoczne w opisie dla użytkownika, np. jako 3 pytania i odpowiedzi
  po treści głównej.
- Ten sam zestaw pytań i odpowiedzi musi znaleźć się w schema. Nie dodawaj do
  JSON-LD pytań, których nie ma w widocznej treści.
- Użyj wyłącznie faktów z briefu oraz bezpiecznej wiedzy ogólnej o typie
  produktu. Nie dopisuj niepotwierdzonych parametrów, miejsc zastosowania ani
  obietnic producenta.
- Odpowiedzi mają być konkretne i przydatne zakupowo: montaż, kompatybilność,
  stopień ochrony IP, barwa światła, źródło światła, materiał albo zastosowanie,
  zależnie od produktu.
- JSON-LD dodaj jako osobny blok po widocznym FAQ:

```html
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "FAQPage",
  "mainEntity": [
    {
      "@type": "Question",
      "name": "Pytanie widoczne w FAQ",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "Odpowiedź widoczna w FAQ."
      }
    }
  ]
}
</script>
```

Przed oddaniem opisu sprawdź, czy pytania i odpowiedzi w widocznym FAQ są
identyczne znaczeniowo z polami `name` i `acceptedAnswer.text` w JSON-LD.

## Styl

Pisz naturalnie po polsku, z pełnymi znakami diakrytycznymi. Wyobraź sobie, że
sprzedawca wyjaśnia produkt klientowi, który chce wiedzieć: co to jest, gdzie
się sprawdzi i dlaczego dany parametr ma znaczenie. Tekst nie może brzmieć jak
raport systemu analizującego bazę danych.

Stosuj proste konstrukcje:

- `Oprawa ma klasę IP65, więc jej obudowa jest chroniona przed pyłem i strugami wody.`
- `Czujnik ruchu automatycznie włącza światło po wykryciu aktywności.`
- `Neutralna barwa 4000 K zapewnia naturalne światło do pracy i codziennych czynności.`
- `Obudowa ma 116x41x28 mm. Przed montażem sprawdź, czy zmieści się w przygotowanej przestrzeni.`
- `Produkt jest objęty 5-letnią gwarancją.`

Nie zamieniaj parametrów w sztuczne korzyści. Napięcie 220-240 V nie „ułatwia
dopasowania”, seria nie „porządkuje identyfikacji”, a kod producenta nie
„wspiera wyboru”. Takie dane zwykle wystarczy umieścić w specyfikacji.

Opisuj relację „cecha -> znaczenie”, np.:

- `IP65` -> ograniczenie wnikania pyłu i ochrona przed strugą wody;
- `4000 K` -> neutralna barwa do pracy i codziennego użytkowania;
- czujnik ruchu -> automatyczne uruchamianie światła po wykryciu ruchu;
- konkretny kąt świecenia -> informacja o sposobie rozsyłu światła;
- kompatybilność -> dokładne wskazanie pasującej serii lub modelu.

Nie wyciągaj skutków, których nie potwierdza parametr. Nie zamieniaj tworzywa
w „wyjątkową trwałość”, a gwarancji w obietnicę bezawaryjności.

## Zakazane

- generowanie wielu opisów pętlą;
- funkcje `build_paragraphs`, słowniki akapitów i szablony prozy per typ;
- kopiowanie kompozycji zdań z poprzedniego produktu;
- „porządkuje dobór”, „techniczny profil wariantu”, „pozwala ocenić model”;
- „wspiera zastosowanie”, „uzupełnia zestaw informacji”, „profil produktu”;
- „ułatwia przypisanie”, „porządkuje parametry”, „dobrze wpisuje się”;
- zdania o „zachowaniu spójności” oparte wyłącznie na nazwie serii;
- schematy „cecha wspiera...”, „parametr ułatwia...”, „seria pomaga...”;
- meta-komentarze o opisie, kliencie, briefie i porównywaniu produktów;
- tłumaczenie, czym produkt nie jest;
- sztuczne podsumowania typu „spełni wszystkie wymagania”;
- przypisywanie `Rejected facts` produktowi;
- wymyślanie certyfikatów, odporności, materiału, parametrów i obietnic.

## Zapis i walidacja

Zapisz HTML w UTF-8, a następnie uruchom:

```powershell
py C:\Users\Handlowiec\Desktop\AlkanInator\src\codex_description_agent.py --brief-input "<brief.xlsx>" --description-file "<opis.html>" --review-output "<review.xlsx>"
```

Potem:

```powershell
py C:\Users\Handlowiec\Desktop\AlkanInator\src\generate_product_descriptions.py --mode validate_codex_review --review-input "<review.xlsx>"
```

Nie oddawaj pliku z `ERROR` ani `WARNING`. Helper nie pisze opisu i nie wolno
go zastępować nowym generatorem treści.
