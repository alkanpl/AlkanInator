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

Python is not guaranteed to be available through `PATH` on this Windows machine. When running Python scripts or quick validators, use the explicit interpreter path:

```powershell
& 'C:\Users\Handlowiec\AppData\Local\Programs\Python\Python314\python.exe' script.py
```

WP-CLI is installed as a PHAR at `C:\wp-cli\wp-cli.phar`. The directory is on `PATH`, but there may be no `wp.bat`/`wp.cmd` wrapper, so prefer invoking it through Laragon PHP:

```powershell
& 'C:\laragon\bin\php\php-8.3.16-Win32-vs16-x64\php.exe' 'C:\wp-cli\wp-cli.phar' --info
```

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
