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
