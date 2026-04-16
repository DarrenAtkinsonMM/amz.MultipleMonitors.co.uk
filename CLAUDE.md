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
