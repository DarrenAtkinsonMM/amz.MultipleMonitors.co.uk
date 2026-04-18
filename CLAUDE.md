# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**MultipleMonitors.co.uk** is an eCommerce site selling monitors, stands, and trading computers. It runs **ProductCart v5.3.00**, a classic ASP (VBScript) platform on Windows Server 2022 / IIS with SQL Server (database: `stagemm`). Payments use SagePay (independent of ProductCart).

The ProductCart vendor (NetSource Commerce) is **sunsetting January 11, 2027** — all vendor server dependencies must be removed before then. See `maintain.md` for the full remediation plan with decoded source, file-by-file instructions, and verification checklist.

## Architecture

### Three-layer structure

- **Customer storefront** (`shop/pc/`, 513 ASP files) — product catalog, cart, checkout, affiliate system. Has **zero** vendor phone-home calls.
- **Admin panel** (`shop/130707/`, 909 ASP files) — store management, order processing. Contains license check and telemetry calls that need removal.
- **Shared includes** (`shop/includes/`, 92 ASP files) — database access, error handling, security, email, encryption, shipping integrations.

### Key entry points

- `default.asp` — site home page (root level)
- `shop/pc/HomeCode.asp` — storefront home content
- `shop/130707/login.asp` — admin login (includes encrypted `AdminLoginInclude.asp`)
- `shop/130707/menu.asp` — admin dashboard
- `shop/includes/common.asp` — master include file loaded on every page (pulls in opendb, settings, ErrorHandler, etc.)

### Important files for the vendor removal project

| File | Purpose |
|------|---------|
| `maintain.md` | **Complete remediation plan** with decoded encrypted source and replacement code |
| `shop/130707/AdminLoginInclude.asp` | Encrypted admin auth — contains 3 vendor HTTP calls (decoded in maintain.md) |
| `shop/130707/login.asp` | Declares `VBScript.Encode` language on line 1; needs changing to `VBScript` |
| `shop/includes/ErrorHandler.asp` | Included on every page via common.asp; has 6 vendor HTTP calls |
| `shop/includes/coreMethods/security.asp` | `pcs_updateDefinitions()` calls `ws.productcart.com` |
| `shop/includes/coreMethods/webservices.asp` | Marketplace URLs (already disabled via settings) |
| `shop/130707/productcartlive.asp` | Admin "Check for Updates" page |
| `shop/includes/pcSurlLvs.asp` | Store URL fingerprint file written by license check |

## Development & Deployment

- **No build step** — IIS auto-compiles Classic ASP on request
- **No package manager, linter, or test framework** — testing is manual in-browser
- **Deployment** is file copy to the server
- **Frontend libs:** jQuery 1.10.2, Bootstrap 3, Animate.css

## Code Conventions

- **VBScript prefix conventions:** `sc*` = store constants (from settings.asp), `pc*`/`pcs_*`/`pcf_*` = ProductCart functions/vars, `pcv_*` = private/verification vars
- **Includes** use `<!--#include file="..." -->` SSI directives
- **Encrypted files** use `<%@ LANGUAGE = VBScript.Encode %>` — the decoded content of `AdminLoginInclude.asp` is documented in `maintain.md`
- **Database access** uses ADO (`ADODB.Connection`/`ADODB.RecordSet`) with string-concatenated SQL
- **Error handling** pattern: `On Error Resume Next` with conditional checks on `err.number`
- **Custom category pages** follow naming pattern `CUSTOMCAT-*.asp` in `shop/pc/`

## Critical Context

- The storefront works independently and will survive the vendor shutdown with only a ~2s delay per new session (ErrorHandler.asp timeout)
- Admin login will work but with 30-60s timeout delays until vendor calls are removed
- All vendor communication errors default to "pass" — no functionality is permanently lost
- SagePay payment processing is completely independent of ProductCart licensing
- `shop/includes/settings.asp` contains all store configuration constants
- `shop/includes/opendb.asp` manages the database connection

## 2026 Redesign

We are currently implementing a redesign to the main site based off the findings in the audit document `multiplemonitors-website-audit.md` — this site is actually a subdomain called amz.multiplemonitors.co.uk — we are connected to the main / live DB file as the changes for this project are to code only and will not touch the DB.

Current-generation mockups live in the `redesign/` folder. The earlier mockups in `html/` (`index2.html`, `trading-computers2.html`, `bundles.html`) and the stylesheet derived from them (`css/mm-2026.css`, scoped under `.mm26`) have been retired — they represent a superseded design direction.

### CSS architecture for the redesign

The redesign stylesheet is `css/mm-site.css`. It sits alongside (not on top of) the legacy stack and is loaded **last** in `shop/pc/inc_headerCSS.asp`, after Bootstrap 3, `style.css`, `responsive.css`, and `blue.css`.

`mm-site.css` has two halves:

1. **Sitewide chrome — unscoped.** Topbar, nav (`.site-header`, `.navwrap`, `.mainnav`, `.mobnav`, `.cart-btn`), and `footer` rules apply on every page. These replace the legacy header/footer in the first rollout phase. Note: the `footer` tag selector matches any `<footer>` element sitewide, so the footer HTML swap must land in the same deploy as the stylesheet to avoid a half-styled footer.
2. **Content design system — scoped under `.mm-site`.** All content components (hero, trust strip, journey cards, pillars, bundle band, depth teasers, Darren CTA, reviews, benchmark panels, comparison table, FAQ, guide band) only apply when a page wraps its main content in `<div class="mm-site">…</div>`. Legacy pages that don't opt in are untouched.

The scoping is deliberate: the mockups' `.container`, `.row`, `.col-*`, `.btn*`, and `.card` class names would otherwise collide with Bootstrap 3 and `style.css`. Keeping them under `.mm-site` avoids specificity wars with the 53 `!important` declarations in `style.css`.

**Chrome grid rename.** Because chrome must co-exist with Bootstrap 3 on un-migrated pages, the chrome's centring container uses `.mm-container` instead of `.container` (same 1280 px max-width, 24 px gutters). Inside `.mm-site`, the content area uses `.container` / `.row` / `.col-sm-*` / `.col-md-*` as the mockups expect — the descendant selector keeps them from fighting Bootstrap's top-level `.container`.

### Design tokens

`mm-site.css` defines a `:root` set of CSS custom properties for colours (`--brand`, `--brand-deep`, `--brand-soft`, `--accent`, `--accent-deep`, `--ink`, `--slate`, `--muted`, `--line`, `--tp-green`, `--up`, `--down`, etc.), radii, and shadows. Fonts are EB Garamond (display serif), Geist (body sans), and JetBrains Mono (micro/mono), loaded via `@import` in the stylesheet and also linked in the mockup `<head>` for faster first paint.

### Rollout plan

1. **Header + footer sitewide.** Add the `<link>` to `mm-site.css` (already done) and swap header/footer HTML to the new chrome markup (with `.mm-container`) in one deploy.
2. **Page-by-page content migration.** Redesigned landing and product pages wrap their main content in `<div class="mm-site">` and use the new component classes.
3. **All tested on `amz.` staging**, then flipped live in a single cutover.
4. **Legacy CSS stays put.** Don't edit `style.css`, `responsive.css`, or `blue.css` during the migration. They retire naturally once the last legacy page is migrated.

### Mockup prototyping workflow

Mockups live in `redesign/` (`newhome.html`, `trading.html`) and link to `../css/mm-site.css`. Content between the chrome and footer is wrapped in `<div class="mm-site">`. When adding a new mockup, keep the chrome markup identical to the existing mockups, scope new components under `.mm-site` in `mm-site.css`, and put shared bits (tokens, buttons, grid) in the top of the scoped section with page-specific blocks lower down.

### Relevant files

| File | Purpose |
|------|---------|
| `css/mm-site.css` | 2026 design-system stylesheet — sitewide chrome + `.mm-site`-scoped content components |
| `shop/pc/inc_headerCSS.asp` | Storefront `<head>` CSS loader — `mm-site.css` is the last `<link>` |
| `redesign/newhome.html` | Homepage mockup — reference implementation of the redesigned layout |
| `redesign/trading.html` | Trading-computers category mockup — adds firms strip, benchmark panels, compare table, FAQ, guide band, sticky CTA |
| `multiplemonitors-website-audit.md` | Audit document driving redesign priorities |