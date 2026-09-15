---
name: alkan-baselinker-pipeline
description: Prepare product names, categories, descriptions, attributes, Baselinker imports, and eKonektor mass-product CSV files for the AlkanInator project from supplier XLSX/XML/CSV files. Use when the user asks to "zrobić nazwy", "uzupełnić plik do Baselinkera", prepare an eKonektor/eConnector import, process supplier product lists such as Kobi/Kanlux/Ospel/Simon, or generate/import product data using local Alkan catalog knowledge, categories, descriptions, SKU suffixes, and required platform formats.
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

## Hager/Berker Variant Plugin Compatibility

- Apply newly confirmed grouping rules to future builds. Do not reopen or rewrite historical output files unless the user explicitly asks.
- The storefront plugin changes one variant axis at a time. A valid transition must keep every other variant-axis value identical.
- Never link a two-product group when the products differ on more than one axis.
- In a larger multi-axis group, exclude or flag any product that has no peer differing on exactly one axis; it would be visible but impossible to switch from. For example, `10116019/HAG` with the unique `1x` + `Bez orientacji` combination stays standalone, while `10126019/HAG` may remain in a connected multiplicity/orientation grid.
- Validate that each accepted product has at least one single-axis transition and that no accepted two-product group uses multiple axes.
- Preserve Polish characters as UTF-8 and reject common mojibake sequences such as `Å‚`, `Ä…`, `Å›`, `Ã³`, or `â€”` in generated text and reports.

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
   - In Polish product names, inflect the color adjective to agree with the main product noun, not with a later noun in the phrase, and do not copy an uninflected or masculine supplier form blindly. Examples: `pokrywa`, `puszka`, `zaślepka`, `ładowarka` → `kaszmirowa`; `gniazdo`, `wyjście`, `przyłącze` and plural `klawisze` → `kaszmirowe`; `przycisk`, `łącznik`, `klawisz`, `ściemniacz`, `element` → `kaszmirowy`.
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

7. Build eKonektor output when requested.
   - Read `references/ekonektor_csv.md` completely before generating the file.
   - Prefer `scripts/export_ekonektor_csv.py` for deterministic CSV encoding and validation.
   - Do not mix the eight-column eKonektor schema with the Baselinker CSV schema.

8. Validate before final response.
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
