# AlkanInator - Google Ads API Tool Design Documentation

## 1. Tool Overview

**Tool name:** AlkanInator  
**Tool type:** Internal product SEO and catalog enrichment utility  
**Requested access:** Google Ads API Basic Access  
**Permissible use requested:** Researching keywords and recommendations  
**Primary API use:** Keyword research through Google Ads Keyword Planner data

AlkanInator is an internal command-line tool used by Alkan to prepare e-commerce product catalog data before importing products into Baselinker/WooCommerce workflows. The tool processes supplier product spreadsheets, extracts and normalizes product attributes, generates SEO-friendly product names, and enriches output files with keyword research data.

The Google Ads API integration is limited to keyword research. It does not create, edit, pause, remove, or otherwise manage Google Ads campaigns, ad groups, ads, assets, billing, users, or account settings.

## 2. Business Purpose

Alkan sells electrical and lighting products online. The catalog contains many technical products with similar names, such as LED panels, hermetic luminaires, high bay fixtures, floodlights, and ceiling luminaires.

The tool uses Google Ads keyword planning data to select better product-type keywords for product naming and SEO metadata decisions. For each product type, the tool identifies keyword candidates and chooses the two best keywords based on average monthly search volume.

Example product types:

- `panel LED`
- `oprawa hermetyczna LED`
- `High Bay przemyslowy`
- `naswietlacz LED`
- `plafon LED`

Example output fields:

- `primary_keyword`
- `primary_keyword_avg_monthly_searches`
- `secondary_keyword`
- `secondary_keyword_avg_monthly_searches`
- `keyword_candidates_json`
- `keyword_status`
- `keyword_source`

## 3. Users and Access Model

The tool is for internal use only. It is operated by Alkan staff on a local workstation as a command-line script.

There is no public web interface, no third-party access, no customer-facing login, and no external user dashboard. The tool is not provided as SaaS and is not used by external advertisers or agencies.

The OAuth user is an authorized Google account with access to the relevant Google Ads account. Credentials are stored locally in `google-ads.yaml`, which is excluded from version control.

## 4. Google Ads API Scope and Services

The integration uses the official Google Ads Python client library.

The current implementation uses:

- `KeywordPlanIdeaService.GenerateKeywordIdeas`
- `KeywordPlanIdeaService.GenerateKeywordHistoricalMetrics`

The requested data is limited to:

- keyword ideas for product-type seed phrases,
- historical keyword metrics,
- average monthly searches.

The tool does not use:

- `GoogleAdsService.Search` for campaign reporting,
- campaign creation or management services,
- ad group services,
- ad services,
- asset services,
- billing services,
- account creation services,
- user management services,
- recommendation apply/dismiss operations.

## 5. Data Flow

1. The operator provides a supplier/product spreadsheet, for example `Alkan_Kanlux_pelne_rodziny.xlsx`.
2. The pipeline extracts product attributes and determines the product type from columns such as `attr_typ`, `Typ produktu`, `Typ`, or `Rodzaj produktu`.
3. The keyword step groups products by unique product type.
4. For each unique product type, the tool sends the product type as a seed phrase to Google Ads API.
5. The tool retrieves keyword ideas and filters out:
   - brand names,
   - SKU/product codes,
   - EAN/GTIN-like numeric values,
   - duplicate phrases,
   - overly long phrases.
6. The tool keeps up to five keyword candidates per product type.
7. The tool requests historical metrics for these candidates.
8. The tool ranks candidates by `avg_monthly_searches`.
9. The best keyword is written as `primary_keyword`; the second-best keyword is written as `secondary_keyword`.
10. The enriched spreadsheet is saved locally for review and later catalog import work.

## 6. Data Storage and Retention

The tool stores all generated outputs locally in the project workspace.

Stored files may include:

- enriched product spreadsheet outputs in `output/`,
- keyword reports in `reports/<run-name>/keywords/`,
- local metric cache in `cache/keyword_planner_metrics.json`.

The cache contains keyword phrases, targeting settings, source name, and average monthly search values. It does not contain campaign performance data, ad spend data, conversion data, customer lists, or user identity data.

The API credential file `google-ads.yaml` contains OAuth and developer-token configuration. It is excluded from version control through `.gitignore`.

## 7. Security Controls

The tool uses the following security controls:

- Credentials are stored locally and are not committed to the repository.
- `google-ads.yaml`, `.env`, and `.env.*` are ignored by Git.
- The tool is operated manually by an internal user.
- No credentials are displayed in CLI output.
- No Google Ads data is shared with third parties.
- Output files remain on the local workstation unless manually exported by the operator.
- The tool performs read-only keyword planning requests and does not mutate Google Ads account state.

## 8. Quota and Rate Control

The tool is designed for low-volume catalog enrichment.

Quota controls:

- Requests are made per unique product type, not per individual product.
- The tool checks a maximum of five keyword candidates per product type.
- Historical metrics are cached locally to avoid repeated API calls for the same keyword, geography, language, and network.
- Typical test runs contain 5-20 unique product types.
- Full catalog runs are expected to stay well below the Basic Access limit of 15,000 operations per day.

Default targeting:

- Geography: Poland
- Language: Polish
- Network: Google Search

These settings can be configured through CLI flags.

## 9. User Workflow

The keyword step can be run as a standalone step:

```powershell
py src/analyze_keywords.py `
  --input output/products_optimized.xlsx `
  --output output/products_with_keywords.xlsx `
  --reports-dir reports/keywords `
  --google-ads-config google-ads.yaml
```

It can also be run as part of the main catalog workflow:

```powershell
py src/run_pipeline_with_parameters.py `
  --input input/Alkan_Kanlux_pelne_rodziny.xlsx `
  --sheet "7. Wszystkie SKU (master)" `
  --parameters input/Parametry.xlsx `
  --config configs/categories/kanlux-oswietlenie.yaml `
  --output output/alkan_kanlux_master_names_with_keywords.xlsx `
  --reports-dir reports/alkan_kanlux_master_names_with_keywords `
  --analyze-keywords `
  --google-ads-config google-ads.yaml
```

## 10. Error Handling and Reporting

The tool does not fail the entire catalog pipeline if Google Ads API calls fail for a product type. Instead, it writes status and diagnostic fields.

Product-level status values include:

- `OK`
- `NO_TYPE`
- `NO_DATA`
- `ERROR`

The tool writes API/provider errors to:

```text
reports/<run-name>/keywords/keyword_api_errors.csv
```

It also writes:

```text
reports/<run-name>/keywords/keyword_type_summary.csv
reports/<run-name>/keywords/keyword_missing_types.csv
```

## 11. Compliance Position

AlkanInator is not a full-service Google Ads management tool. It is an internal, specialized keyword research and product catalog enrichment utility.

The Google Ads API integration is limited to planning/research services and does not provide campaign creation, campaign management, ad management, reporting dashboards, billing operations, account creation, or user management.

The requested Basic Access is needed because `KeywordPlanIdeaService` is restricted for lower access levels and is required to retrieve keyword ideas and keyword historical metrics for production account keyword research.

## 12. References

- Google Ads API Access Levels and Permissible Use: https://developers.google.com/google-ads/api/docs/api-policy/access-levels
- Google Ads API Required Minimum Functionality: https://developers.google.com/google-ads/api/docs/api-policy/rmf
- Google Ads API Python Client Configuration: https://developers.google.com/google-ads/api/docs/client-libs/python/configuration
- Keyword Planning - Generate Keyword Ideas: https://developers.google.com/google-ads/api/docs/keyword-planning/generate-keyword-ideas
- Keyword Planning - Generate Historical Metrics: https://developers.google.com/google-ads/api/docs/keyword-planning/generate-historical-metrics
