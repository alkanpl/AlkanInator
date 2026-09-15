# AlkanLocal Project Notes

Use these notes only when the task needs local project context.

## Architecture

AlkanLocal is a local WordPress/WooCommerce installation rooted at `C:\laragon\www\AlkanLocal`.

The active custom storefront work is mainly in:

- `wp-content/themes/alkan/functions.php`
- `wp-content/themes/alkan/style.css`
- `wp-content/themes/alkan/js`
- `wp-content/themes/alkan/woocommerce`

Custom local plugins include allowlisted folders under `wp-content/plugins`, especially `alkan-*`, plus selected existing integrations. Always check `.gitignore` before assuming a new plugin is tracked.

`wp-content/mu-plugins` is used for always-loaded local helpers. The repo's ignore rules require explicit allowlisting of mu-plugin files.

## Test Surface

`e2e/package.json` defines:

- `npm run test:e2e` -> all Playwright tests
- `npm run test:smoke` -> `@smoke`
- `npm run test:regression` -> `@regression`
- `npm run test:headed` -> headed run
- `npm run report` -> Playwright HTML report
- `npm run codegen` -> Playwright codegen against `https://alkanlocal.test:8443`

Existing smoke tests cover:

- homepage load
- live search suggestions
- navigation from search to product
- category listing, sorting, filters
- mobile homepage search/cart visibility

Existing regression tests cover:

- add product to cart
- mini-cart content
- checkout progression
- free shipping threshold
- shipping method matrix scenarios

## Verification Expectations

For visual tasks, inspect the rendered page through Playwright at the affected URL. Use desktop and mobile viewport checks when layout can change. If the result does not match the user's instruction, continue editing and re-check before final response.

For PHP-only changes, lint every edited PHP file with Laragon PHP. For runtime-sensitive WordPress behavior, bootstrap `wp-load.php` with `REQUEST_SCHEME=https` and `HTTP_HOST=alkanlocal.test:8443`.
