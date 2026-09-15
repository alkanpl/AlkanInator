# Alkan Codex Skills Export

Plik do przekazania innemu Codexowi. Zawiera instrukcje repozytorium AlkanInator oraz lokalne skille Alkan.

## Jak użyć

- `AGENTS.md` wklej do katalogu głównego repozytorium AlkanInator u drugiego Codexa.
- Każdy blok `Skill: ...` zapisz jako osobny plik `SKILL.md` w katalogu `.codex/skills/<nazwa-skilla>/` albo przekaż jako kontekst systemowy/projektowy.
- `alkan-baselinker-pipeline` i `alkan-product-description` dotyczą AlkanInatora.
- `alkanlocal-site-workflow` dotyczy osobnego projektu WordPress/WooCommerce `AlkanLocal`, ale jest dołączony, bo też jest skillem Alkan.

---

## AGENTS.md dla repo AlkanInator

Source: `C:\Users\Handlowiec\Desktop\AlkanInator\AGENTS.md`
Type: `repo-instructions`

````markdown
# AGENTS.md

Wskazowki dla Codex przy pracy nad tym repozytorium.

## Czym jest AlkanInator

Lokalne narzedzie CLI (Python) do masowego przetwarzania katalogow produktow
elektrycznych/oswietleniowych z plikow XLSX/CSV: ekstrakcja atrybutow z tytulow,
generowanie tytulow SEO, wzbogacanie z eksportow WooCommerce/Baselinker,
rekomendacje kategorii i analiza slow kluczowych. Logika biznesowa siedzi w
konfiguracji YAML (`configs/`, `dictionaries/`), nie w kodzie.

Zasada nadrzedna: **oryginalne pliki wejsciowe nigdy nie sa nadpisywane** -
wyniki ida do `output/` i `reports/`. Kazda zmiana atrybutu niesie `source`,
`confidence` i `status` (`OK` / `WARNING` / `DO_SPRAWDZENIA`).

## Uruchamianie

Interpreter to `py` (Windows launcher) - `python` nie jest w PATH.
Skrypty uruchamia sie z katalogu glownego repo, importy w `src/` sa plaskie
(modul importuje `from utils import ...`), wiec `src/` musi byc na sciezce -
co zapewnia uruchomienie `py src/<skrypt>.py`.

Glowne wejscia:

- `py src/run_pipeline.py` - pipeline z ekstrakcja atrybutow z tytulu.
- `py src/run_pipeline_with_parameters.py` - pipeline z uzupelnianiem atrybutow
  z pliku `Parametry.xlsx`.

Oba dziela wspolny rdzen w `src/pipeline_core.py` (analiza wejscia,
klasyfikacja rol, slowa kluczowe, tytuly, walidacja kategorii, eksport).
Roznia sie tylko sposobem zbudowania kolumny roboczej `__working_title`.

## Architektura - mapa modulow

- `utils.py` - I/O (XLSX/CSV jako string, auto-separator), wykrywanie kolumn,
  walidacja EAN/GTIN, normalizacja tekstu.
- `analyze_products.py` - raport jakosci, wykrywanie kolumn, braki atrybutow.
- `extract_attributes.py` - rdzen ekstrakcji (regexy: moc, barwa, IP, strumien,
  gwint, wymiary...). Najbardziej kruchy modul - patrz testy charakteryzujace.
- `optimize_titles.py` + `title_anatomy.py` - generowanie tytulow SEO wedlug
  regul "anatomii tytulu" per typ produktu (`configs/title_anatomy.yaml`).
- `validate_output.py` - role produktow (glowny vs akcesorium), zgodnosc z
  kategoria.
- `catalog_knowledge.py` - wstrzykiwanie wiedzy z eksportu WooCommerce do configu.
- `dictionaries/supplier_global_knowledge.yaml` - globalna wiedza o dostawcach,
  obecnie przede wszystkim Kanlux: suffixy, zasady kategorii, tytulow,
  akcesoriow, atrybutow i bezpiecznych opisow. Przy zadaniach Kanlux czytaj ten
  plik razem z `docs/kanlux_workflow.md`.
- `analyze_keywords.py` - Google Ads Keyword Planner (zaleznosc opcjonalna,
  importowana leniwie; dziala tez offline z `--offline-keywords`).
- Pozostale `enrich_*`, `recommend_*`, `export_*`, `build_kobi_baselinker.py`,
  `learn_product_taxonomy.py` - narzedzia pomocnicze, kazde z wlasnym argparse.

## Testy

Testy uzywaja `unittest` (biblioteka standardowa, bez dodatkowych zaleznosci):

    py -m unittest discover -s tests -p "test_*.py"

Najwazniejsze sa **testy charakteryzujace** (golden-file) dla ekstrakcji i
generowania tytulow - utrwalaja aktualne zachowanie na realnych danych Kanlux,
zeby kazda niezamierzona regresja w regexach/regulach byla od razu widoczna.

Po **SWIADOMEJ** zmianie regul ekstrakcji, configow tytulow lub slownikow
nalezy zregenerowac snapshoty (wymaga `input/Alkan_Kanlux_pelne_rodziny.xlsx`):

    py tests/generate_snapshots.py

Fixtures w `tests/fixtures/` sa samowystarczalne - same testy nie potrzebuja
plikow XLSX.

## Konwencje

- Kod i wynikowe pliki: w wiekszosci bez polskich znakow w nazwach pol roboczych,
  ale dane produktowe i tytuly SEO sa po polsku (utf-8, CSV zapisywany jako
  `utf-8-sig` dla Excela).
- Konsola Windows uzywa cp1250 - nie polegaj na `print()` z polskimi znakami w
  skryptach diagnostycznych; zapisuj do pliku w utf-8.
- Status planu wdrozenia: `PLAN_ROBOCZY.md`.
````

---

## Skill: alkan-baselinker-pipeline

Source: `C:\Users\Handlowiec\.codex\skills\alkan-baselinker-pipeline\SKILL.md`
Type: `codex-skill`

````markdown
---
name: alkan-baselinker-pipeline
description: Prepare product names, categories, descriptions, attributes, and Baselinker CSV imports for the AlkanInator project from supplier XLSX/XML/CSV files. Use when the user asks to "zrobić nazwy", "uzupełnić plik do Baselinkera", process supplier product lists such as Kobi/Kanlux/Ospel/Simon, or generate/import product data using local Alkan catalog knowledge, categories, descriptions, SKU suffixes, and Baselinker format.
---

# Alkan Baselinker Pipeline

## Core Rule

Always treat supplier files as source data, not final store data.

- Use XML/feed/catalog files for EAN, product IDs, images, raw attributes, availability, prices, and producer parameters.
- Do not use supplier XML categories as final Baselinker categories.
- Do not copy producer HTML descriptions into Baselinker `description`.
- Generate final categories from local Alkan knowledge and rules.
- Generate final descriptions locally from product names and attributes.
- Put the product code/SKU into the final product name when this matches the project pattern.

## Project Context

Default workspace: `C:\Users\Handlowiec\Desktop\AlkanInator`.

Read these before changing or generating an import:

- `README.md`
- `dictionaries/woocommerce_catalog_knowledge.yaml`
- `dictionaries/learned_product_taxonomy.yaml`
- `dictionaries/category_recommendation_overrides.yaml`
- Existing recent Baselinker outputs in `output/` and relevant examples in `archived_input_files/`
- Category configs in `configs/categories/`, especially `kanlux-oswietlenie.yaml` and `gazetka-produkty.yaml`

Prefer existing scripts over one-off logic:

- `src/recommend_categories.py`
- `src/generate_product_descriptions.py`
- `src/export_to_baselinker_csv.py`
- `src/enrich_baselinker_export.py`
- category/title helpers such as `src/optimize_titles.py`

## Required Workflow

1. Inspect inputs.
   - Read all sheets/columns and sample rows from XLSX/CSV.
   - Parse supplier XML/feed structure and identify stable join keys.
   - Verify match counts before generating output.

2. Build normalized product rows.
   - Preserve original supplier columns in the control XLSX.
   - Join supplier feed attributes by SKU/EAN/code.
   - Normalize codes consistently; keep Baselinker `sku` in the project format such as `003826/KOB`.
   - Keep a clean manufacturer value consistent with the shop catalog, not necessarily the raw supplier value.

3. Generate names.
   - Use local title patterns and prior exports as style reference.
   - Do not blindly preserve the supplier title prefix. Enforce the catalog-style product type for the final store category, e.g. `Oprawa High Bay LED...` instead of supplier `High Bay LED...`, `Oprawa uliczna LED...` instead of `Oprawa drogowa LED...`, and `Słupek ogrodowy...` instead of `Słup...` when the catalog uses that pattern.
   - Include meaningful attributes such as type, series/model, power, lumens, CCT, IP, color, dimensions, socket, sensor, and producer/brand.
   - Append product code to the final name, e.g. `... Kobi Pro 003829`, unless an existing accepted project pattern says otherwise.
   - Avoid duplicate specs such as two IP values, repeated codes, repeated brands, or supplier noise like `nowy model`.

4. Generate store categories.
   - Use `CategoryRecommender` from `src/recommend_categories.py` with `woocommerce_catalog_knowledge.yaml`, `learned_product_taxonomy.yaml`, and overrides.
   - Add protective rules when a supplier group is known to belong to one shop family, so generic learned terms do not misclassify lighting accessories as Berker frames, cable holders, automation sensors, etc.
   - Store categories must be Alkan/WooCommerce paths, for example:
     - `Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Panele LED sufitowe`
     - `Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Lampy hermetyczne`
     - `Oświetlenie > Oświetlenie przemysłowe wewnętrzne > Oprawy High-bay LED`
     - `Oświetlenie > Oświetlenie zewnętrzne > Naświetlacze LED`
     - `Oświetlenie > Oświetlenie zewnętrzne > Lampy uliczne`
     - `Oświetlenie > Oświetlenie wewnętrzne > Plafony LED`
     - `Żarówki > Świetlówki LED`
   - Record recommendation score/reason/status in the control XLSX.

5. Generate descriptions.
   - Use `src/generate_product_descriptions.py` or its `build_description_html` function.
   - Feed it final `name`, `sku`, `manufacturer_name`, and JSON `features`.
   - Do not use supplier descriptions as Baselinker descriptions. Keep them only as control columns such as `Opis producenta`.
   - Confirm generated descriptions meet the local minimum length convention when applicable.

6. Build Baselinker output.
   - Required CSV columns: `product_id;name;sku;ean;manufacturer_name;category;description;features;images_urls`.
   - Use semicolon delimiter and UTF-8 BOM.
   - `features` must be valid compact JSON.
   - `images_urls` should use the format expected by existing project exports.

7. Validate before final response.
   - Row count equals input product count after removing empty rows.
   - No empty `name`, `sku`, `manufacturer_name`, `category`, `description`, or `features`.
   - Report empty XML-derived fields (`product_id`, `ean`, `images_urls`) separately when source data is missing.
   - No duplicate final names unless explicitly acceptable.
   - Every final name contains the intended product code/SKU when required.
   - No copied producer-description markers in `description`, such as supplier URLs, `data-section-id`, or image asset HTML.
   - `features` parses as JSON for every row.
   - Categories are store categories, not supplier XML paths.

## Outputs

For supplier batch work, produce both:

- A control XLSX in `output/` with original data, generated name, store category, recommendation reasons, generated description, raw supplier description, images, features, and match status.
- A Baselinker CSV in `output/` with the required import columns.

Also write a small report in `reports/` with row counts, match counts, missing source matches, and validation results.

## Communication

In the final response, include:

- Exact output file paths.
- Match count and missing feed-data count.
- Category source summary.
- Validation results.
- Any rows that still lack EAN/images/product_id because the supplier source did not contain them.
````

---

## Skill: alkan-product-description

Source: `C:\Users\Handlowiec\.codex\skills\alkan-product-description\SKILL.md`
Type: `codex-skill`

````markdown
---
name: alkan-product-description
description: Generate individual Alkan product HTML SEO descriptions from description_codex_brief_*.xlsx files. Use for one product or a requested list of products, but author and validate every SKU separately; never generate prose with a Python loop or shared paragraph template.
---

# Alkan Product Description

## Core Rule

Write each product description directly in Codex. Local Python may read XLSX,
save HTML, build review XLSX, and validate. It must never compose prose, store
paragraph templates, or iterate through briefs to generate descriptions.

Use `C:\Users\Handlowiec\Desktop\AlkanInator` as the main repository. Read
`prompts\codex_description_agent.md` before writing; it is the current style
and quality contract. Also read
`prompts\examples\good_description_arot.html` as the approved quality example.
Read `dictionaries\seo_description_knowledge.yaml` before writing. It is the
source of truth for global SEO-description rules. If an older prompt or example
differs from this YAML, follow the YAML while preserving factual safety.

## Brief Selection

If the user provides a brief path or SKU, use it explicitly. Otherwise always
run this command before opening any workbook:

```powershell
py src\codex_description_agent.py --list-briefs
```

The command recursively finds briefs and returns the complete newest batch
directory. Treat every printed XLSX path as the current queue. If it prints 10
briefs, a general request such as `zrób opis` or `wygeneruj opisy` means process
all 10 sequentially.

Never select a brief using a non-recursive glob, the first filename returned,
or a hardcoded `description_codex_brief_33482.xlsx`. Never replace a multi-file
latest batch with one older file from the root `reports\descriptions` folder.

For an explicit SKU, resolve its newest matching brief with:

```powershell
py src\codex_description_agent.py --list-briefs --sku "<sku>"
```

## One-Product Cycle

1. Take the next brief from the selected queue.
2. Read `Produkt`, `Product facts`, `Compatibility facts`, `Rejected facts`,
   and `Prompt dla Codex`.
3. Plan a product-specific angle from its type, use, strongest facts, and
   compatibility.
4. Write the HTML manually. Do not create a prose generator.
5. Save it as UTF-8 in the main repo.
6. Create review with:

```powershell
py src\codex_description_agent.py --brief-input "<brief.xlsx>" --description-file "<description.html>" --review-output "<review.xlsx>"
```

7. Validate with:

```powershell
py src\generate_product_descriptions.py --mode validate_codex_review --review-input "<review.xlsx>"
```

8. Fix every `ERROR` and `WARNING`. Only then start another SKU.

For a list of products, repeat this cycle sequentially. Do not open all briefs,
build a shared template, or postpone validation until the end.

## Facts

- Use `Product facts` as product-specific claims.
- Use `Compatibility facts` only as compatibility.
- Never use `Rejected facts` as product features.
- General knowledge may explain what a product type normally does, but may not
  assign unconfirmed ratings, materials, certifications, performance, or
  applications to the model.
- For accessories, fixture power is not accessory power.

## SEO Structure

- Use the supplied SEO keyword from the brief. Never invent or replace it.
- Use the exact keyword no more than 3 times in the whole visible description.
- Start the first `<p>` with `<strong>the exact full product name</strong>`.
  The opening must also begin with the supplied keyword; normally the keyword
  is the leading product-type phrase in the full name.
- Explain the product, relevant use, and strongest confirmed features early.
- Add a specific `<h2>` containing the supplied keyword in a natural context.
  Prefer a useful question when it reads naturally.
- Start the first `<p>` immediately after `<h2>` with the supplied keyword.
- Do not use the supplied keyword in any `<h3>`.
- Add `<h3>Najważniejsze zalety</h3>` with fact-based benefits.
- Put 3-6 concrete features or benefits in its `<ul>`.
- Add `<h3>Specyfikacja techniczna</h3>` using only `Product facts`.
- Format specification rows as
  `<li><strong>Parameter:</strong> value</li>`.
- Immediately after the specification list, add a separate final `<p>` that
  summarizes the strongest confirmed reason to choose the product and gives a
  natural purchase encouragement. Never end the description on the
  specification `</ul>`.
- Add an application section only when supportable.
- Keep paragraphs short: normally 3-4 sentences maximum.
- Close with a factual summary or purchase-oriented recommendation without
  unsupported promises.
- Write at least 1200 visible characters including spaces. For rich briefs,
  target 1500-2200 non-space characters. For sparse accessories, factual
  accuracy takes priority over padding. If confirmed facts do not support the
  required length, do not invent content; report the conflict for manual review.

The user's AROT/Kopoflex example defines the expected specificity, readable SEO
hierarchy, and relation between features and benefits. It is not a source of
facts for Kanlux products.

## Quality Rules

Vary the reasoning, section wording, benefits, and prose for every product.
Only the HTML hierarchy and specification heading may remain stable.

Write every SKU independently. Products that differ only by color, size, or
another variant must still have different sentences, argumentation, and
composition. Do not reuse a description with reordered sentences.

Write for a person buying or installing the product. Prefer direct statements:

- `Obudowa ma klasę IP65...`
- `Czujnik ruchu automatycznie włącza światło...`
- `Neutralna barwa 4000 K daje naturalne światło...`
- `Produkt jest objęty 5-letnią gwarancją.`

Do not manufacture benefits from database fields. Voltage, code, EAN, series,
and warranty may remain only in specifications when they do not support a
natural, useful sentence.

For every important claim, use this reasoning:

1. Feature: what the product has or does.
2. Advantage: what the feature enables.
3. Benefit: what the user gains in practice.

Do not force a benefit when the facts support only a technical statement.
Use `<strong>` sparingly for information that genuinely deserves emphasis.

Do not write:

- empty superlatives such as `rewelacyjny`, `innowacyjny`, `najlepszy`, or
  `niesamowity`;
- meta-commentary about descriptions, briefs, customers, or comparison;
- `porządkuje dobór`, `techniczny profil wariantu`, or similar filler;
- `wspiera zastosowanie`, `uzupełnia zestaw informacji`, `profil produktu`;
- `ułatwia przypisanie`, `porządkuje parametry`, `dobrze wpisuje się`;
- claims that a series creates visual or technical consistency unless the brief
  confirms matching products;
- repetitive `feature supports`, `parameter facilitates`, or `series helps`
  sentence patterns;
- explanations of what a product is not;
- generic promises such as `spełni wszystkie wymagania`;
- claims not supported by the brief;
- Polish text without diacritics.

Do not copy producer descriptions, external website text, or facts from the
AROT example. Do not describe something visible in a photo unless the statement
adds useful, confirmed product information.

Before saving, verify:

- at least 1200 visible characters including spaces;
- keyword at the opening, in `<h2>`, and at the start of the first paragraph
  after `<h2>`;
- no keyword in `<h3>` and no more than 3 total occurrences;
- short paragraphs and a 3-6 item benefits list;
- factual consistency with `Product facts`, `Compatibility facts`, and
  `Rejected facts`;
- a separate final summary paragraph after the technical specification;
- spelling, Polish diacritics, HTML structure, and natural readability.

Do not create or use scripts resembling `generate_new_descriptions.py`,
`build_paragraphs`, dictionaries of paragraphs, or product-type prose
templates.

## Output

Write final HTML and review XLSX only under the main repository's
`reports\descriptions` tree, never under `.codex\worktrees`. Do not modify
Baselinker exports or source product files.

Report the brief path, HTML path, review path, validation result, and whether
manual review remains.
````

---

## Skill: alkanlocal-site-workflow

Source: `C:\Users\Handlowiec\.codex\skills\alkanlocal-site-workflow\SKILL.md`
Type: `related-alkanlocal-skill`

````markdown
---
name: alkanlocal-site-workflow
description: Work safely in the AlkanLocal WordPress/WooCommerce project. Use whenever the current working directory is C:\laragon\www\AlkanLocal or the task concerns the AlkanLocal storefront, theme, WooCommerce templates, Alkan plugins, mu-plugins, checkout, product pages, navigation, search, cart, styling, JavaScript, or any website-facing change. Always use this skill for visual changes in AlkanLocal and for creating new local plugins that must be tracked by git.
---

# AlkanLocal Site Workflow

## Project Map

Treat `C:\laragon\www\AlkanLocal` as the project root.

- Main storefront theme: `wp-content/themes/alkan`
- WooCommerce template overrides: `wp-content/themes/alkan/woocommerce`
- Theme CSS/JS: `wp-content/themes/alkan/style.css`, `wp-content/themes/alkan/js`
- Local custom plugins tracked by git: selected `wp-content/plugins/alkan-*` and other allowlisted plugin folders
- Always-loaded local code: `wp-content/mu-plugins`
- E2E tests: `e2e`, using Playwright against `https://alkanlocal.test:8443`
- Manual scenarios: `SCENARIUSZE_TESTOWE_SKLEPU.md`
- Git ignore allowlist: `.gitignore`

Read `references/project-notes.md` when needing more project-specific context.

## Default Workflow

1. Inspect relevant files first. Use `rg`/`rg --files` before editing.
2. Keep changes scoped to the theme, WooCommerce override, plugin, or mu-plugin that owns the behavior.
3. Preserve existing WordPress/WooCommerce patterns and legacy PHP style unless there is a clear local reason to modernize.
4. For PHP changes, run lint with Laragon PHP:

```powershell
& 'C:\laragon\bin\php\php-8.3.16-Win32-vs16-x64\php.exe' -l path\to\file.php
```

5. For behavior touching cart, checkout, product cards, product pages, search, filtering, navigation, shipping, or payment, run the relevant Playwright test from `e2e`.
6. Report any tests that could not be run and why.

## Visual Changes

For any visual or frontend-facing change, use Playwright before finishing.

Required checks:

1. Open the changed page or workflow with Playwright at `https://alkanlocal.test:8443`.
2. Verify desktop and mobile viewports when the change can affect responsive layout.
3. Inspect the actual rendered result, not only test pass/fail.
4. If anything differs from the user's instruction, fix it and re-check with Playwright before final response.

Useful commands:

```powershell
cd e2e
npm run test:smoke
npm run test:regression
npm run test:headed
```

For targeted visual verification, use Playwright directly or the browser skill/plugin when available. Capture screenshots or inspect DOM/layout when needed.

## Playwright Guidance

The Playwright config is in `e2e/playwright.config.js`.

- Default base URL: `https://alkanlocal.test:8443`
- Browser: Chromium
- HTTPS errors are ignored for local certs.
- Traces/video are retained on failure.
- Helpers in `e2e/tests/site-smoke.helpers.js` already handle overlays, slow navigation, mini-cart, checkout, and shipping scenarios.

Choose tests by risk:

- Header, search, product listings, product page visibility: `npm run test:smoke`
- Cart, mini-cart, checkout, shipping, payment: `npm run test:regression`
- Shipping method logic: run the regression suite or targeted `@shipping` tests.
- Pure PHP/API change with no visible effect: lint and a focused runtime check may be enough.

## New Plugins And Git Ignore

The repo ignores `/wp-content/**` and then allowlists specific paths. When creating a new plugin or mu-plugin that should be committed, update `.gitignore` in the same change.

For a new normal plugin:

```gitignore
!/wp-content/plugins/my-plugin/
!/wp-content/plugins/my-plugin/**
```

For a new mu-plugin file:

```gitignore
!/wp-content/mu-plugins/
!/wp-content/mu-plugins/my-file.php
```

After adding a plugin, run:

```powershell
git check-ignore -v wp-content\plugins\my-plugin\my-plugin.php
git status --short
```

If `git check-ignore` still reports an ignore rule for a file that should be tracked, fix `.gitignore` before finishing.

## Local Runtime Notes

WordPress CLI may not be available. If a runtime check is needed, bootstrap WordPress with Laragon PHP and set server vars first to avoid local `wp-config.php` warnings:

```powershell
@'
<?php
$_SERVER['REQUEST_SCHEME'] = 'https';
$_SERVER['HTTP_HOST'] = 'alkanlocal.test:8443';
require 'wp-load.php';
// focused check here
'@ | & 'C:\laragon\bin\php\php-8.3.16-Win32-vs16-x64\php.exe'
```

Do not use this as a substitute for Playwright when the change is visual.
````

