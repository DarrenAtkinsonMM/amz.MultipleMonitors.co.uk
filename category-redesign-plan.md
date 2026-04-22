# Category-page redesign — approach & per-page plans

This document captures the investigation and decisions for porting the 2026 redesign mockups onto the live category pages. It's written to be **reusable across all category pages**, with a stands-specific implementation plan as the first worked example.

Related docs:
- [`CLAUDE.md`](CLAUDE.md) — 2026 redesign overview, CSS architecture, rollout phase
- [`multiplemonitors-website-audit.md`](multiplemonitors-website-audit.md) — audit driving the redesign
- [`maintain.md`](maintain.md) — ProductCart vendor-shutdown (Jan 2027) remediation
- [`chrome-rollout-plan.md`](chrome-rollout-plan.md) — header/footer rollout that preceded this work

## 1. Context

The redesign is being rolled out page-by-page:
1. **Done**: [`css/mm-site.css`](css/mm-site.css) deployed; header + footer chrome swapped sitewide; [`default.asp`](default.asp) (homepage) migrated under `.mm-site`.
2. **This phase**: category pages — starting with **stands** (this doc), then likely computers, monitors, bundles, arrays.
3. **Later**: product detail pages.

Each category page currently has its own file under [`shop/pc/`](shop/pc/) following the pattern `CUSTOMCAT-<slug>.asp` (stands, computers, tradingcomputers, monitors, bundles1/2/3, arrays1/2/3). All of them currently run the full ProductCart product-display stack.

## 2. What the mockups actually need from the page

Across the redesigned category mockups, the dynamic surface is **tiny**:

- Most of each page is **static marketing content** (hero, trust strip, benefit pillars, founder story, specs, cross-links, CTAs).
- The only dynamic content is a **small, fixed set of product tiles** — [`redesign/stands.html`](redesign/stands.html) has 12 in 3 groups; the trading page has a similar structure.
- Each tile is a **link-only card**: image, title, "From £X", "View …" CTA. No inline Buy Now, no list-price / savings block, no wholesale price row, no stock badge, no review stars, no quantity selector.

Compare this with the current live renderer [`shop/pc/pcShowProduct-Standard.asp`](shop/pc/pcShowProduct-Standard.asp), which emits `col-md-4 .product-detail` Bootstrap-3 cards with More-Info + Buy-Now buttons, list-price/savings, and a wholesale-tier branch. The markup and feature set don't match — at all.

## 3. ProductCart machinery on today's category pages

Typical `CUSTOMCAT-*.asp` wiring (stands is representative — 602 lines):

| Area | What it does | Needed under the redesign? |
|---|---|---|
| [`includes/common.asp`](shop/includes/common.asp) | Sessions, DB (`connTemp`), language, affiliate, store constants, `money()` via [`currencyformatinc.asp`](shop/includes/currencyformatinc.asp) | **Keep** |
| `pcStartSession.asp` | Session init, affiliate tracking | **Keep** |
| [`header_wrapper.asp`](shop/pc/header_wrapper.asp) / [`footer_wrapper.asp`](shop/pc/footer_wrapper.asp) | Redesigned chrome (already deployed) | **Keep** |
| Category display-settings lookup (`pcCats_PageStyle`, columns, rows, mobile override) | Varies the grid layout per admin-configured category settings | **Drop** — mockup fixes the layout |
| Sort-order resolution (`prodsort`, `POrder`, BTO-aware ordering, ~70 LOC of SQL) | Lets users reorder products | **Drop** — mockup has a fixed order |
| Solr faceted search branch (`SRCH_CSFON`, `pcv_strCSFilters`) | Optional site-search facets | **Drop** — these categories aren't faceted |
| ADODB pagination (`rs.PageSize`, `rs.AbsolutePage`) | Server-side pagination | **Drop** — small fixed product sets |
| Mobile session override (`session("Mobile")="1"` → forced 1-col layout) | Pre-responsive-CSS adaptation | **Drop** — `mm-site.css` is responsive |
| `pcShowProducts.asp` dispatcher + `pcShowProduct-Standard.asp` tile | Emits tile HTML | **Drop** for redesigned pages — emit the mockup markup inline |
| `pcGetPrdPrices.asp` (wholesale-tier / BTO fallback / parent-price) | Per-product price resolution | **Drop** for tiles (not shown on cards); still runs on detail pages unchanged |
| `prv_incFunctions.asp`, `pcCheckPricingCats.asp`, `pcValidateHeader.asp`, `pcValidateQty.asp`, `atc_viewprd.asp`, `bulkAddToCart.asp`, `common_checkout.asp`, `SearchConstants.asp`, `prv_getSettings.asp` | Reviews, category pricing, header validation, qty validation, add-to-cart dialog, bulk cart ops, checkout helpers, search consts, settings lookup | **Drop** on redesigned category pages |

**Product detail pages continue to use the full stack unchanged** — add-to-cart, wholesale pricing, BTO options, parent-price all still work.

## 4. The reusable approach (any redesigned category page)

For each redesigned category page, rewrite `CUSTOMCAT-<slug>.asp` in place to a mostly-static ASP file. Keep the filename, so URL rewrites and SEO stay intact. Keep ProductCart session/chrome. Skip the product-display chain.

### 4.1 Skeleton

```asp
<%@ LANGUAGE = VBScript %>
<!--#include file="../includes/common.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
' Category ID for this page (check categories_products / categories table)
Dim pIdCategory : pIdCategory = 5    ' stands = 5 (example)
%>
<!--#include file="header_wrapper.asp"-->

<div class="mm-site">

  <!-- All static mockup sections pasted in verbatim -->
  <section class="hero"> … </section>
  <section class="truststrip"> … </section>
  …

  <!-- The single dynamic section -->
  <section class="s depth" id="range">
    <div class="container">
      <div class="section-head reveal"> … </div>
      <%
      Dim query, rs, pcArray
      query = "SELECT p.idProduct, p.sku, p.description, p.price, p.smallImageUrl, p.pcUrl " & _
              "FROM products p INNER JOIN categories_products cp ON p.idProduct=cp.idProduct " & _
              "WHERE cp.idCategory=" & pIdCategory & " " & _
              "  AND p.active=-1 AND p.configOnly=0 AND p.removed=0 " & _
              "ORDER BY cp.POrder ASC, p.description ASC"
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open query, connTemp, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then pcArray = rs.GetRows()
      Set rs = Nothing

      ' Bucket into groups, render each group as a <div class="range-group">
      ' containing a .bundle-cards wrapper with .bundle-card links.
      ' Group/eyebrow/style derivation per-page (see §5 for stands).
      %>
    </div>
  </section>

  <!-- More static sections (bundle cross-link, Darren CTA, etc.) -->

</div>

<!--#include file="footer_wrapper.asp"-->
```

### 4.2 Per-card markup (from the mockups — do not change)

```html
<a href="/shop/pc/<productSlug>.htm" class="bundle-card reveal">
  <div class="bundle-card__media">
    <img src="/shop/pc/catalog/<smallImageUrl>" alt="<title>">
  </div>
  <div class="bundle-card__body">
    <div class="bundle-card__eyebrow"><eyebrow></div>
    <h4 class="bundle-card__title"><title></h4>
    <div class="bundle-card__price">
      <span class="bundle-card__from">From</span>
      <span class="bundle-card__amount">£<price></span>
    </div>
    <span class="btn btn-primary bundle-card__cta">View <kind> <i class="fa fa-arrow-right"></i></span>
  </div>
</a>
```

### 4.3 Price rendering

Single price per product, VAT-exclusive display (stores are VAT-inclusive in the DB). Use the existing `money()` formatter from [`shop/includes/currencyformatinc.asp`](shop/includes/currencyformatinc.asp):

```vbscript
scCursign & money(productPrice / 1.2)
```

No wholesale branch on tiles. No list-price/savings. The "From" prefix is brand voice; there's no price-min calculation.

### 4.4 Product URL

Use `products.pcUrl` as the href segment. Follow the existing convention used throughout the storefront (`/shop/pc/<pcUrl>.htm` or `/products/<slug>/` — match whatever the mockup and current detail pages use).

### 4.5 What still works unchanged

- Add-to-cart from product detail pages.
- Wholesale / trade customer pricing on detail pages.
- Affiliate tracking, session cart, checkout.
- All other (non-redesigned) category pages continue to use `pcShowProducts.asp` / `pcShowProduct-Standard.asp` exactly as today.

## 5. Stands page — specific implementation plan

### 5.1 File to edit
[`shop/pc/CUSTOMCAT-stands.asp`](shop/pc/CUSTOMCAT-stands.asp) — rewrite in place (drops ~602 LOC → ~250-350 LOC, mostly static HTML).

### 5.2 Source-of-truth mockup
[`redesign/stands.html`](redesign/stands.html). Paste hero (lines 84-110), trust strip (115-148), pillars (155-193), design/mfg story (198-222), modular upgrade path (228-278), 28″ screens (284-306), adjustability (311-363), shared specs (368-435), product range (440-649), bundle cross-link (655-692), Darren CTA (697-715), and the inline `<script>` blocks (782-803).

### 5.3 Category ID
`5` (confirmed by the current hardcoding in [`CUSTOMCAT-stands.asp:49`](shop/pc/CUSTOMCAT-stands.asp#L49)).

### 5.4 Decisions

1. **Grouping**: derive from SKU prefix.
   - Suffix `p2*`, `p3*` → "Dual & Triple-screen stands"
   - Suffix `p4*` → "Quad-screen stands"
   - Suffix `p5*`, `p6*`, `p8*` → "Five, Six & Eight-screen stands"
   - New stand products land in the correct group automatically as long as SKUs stay on-pattern.
2. **Pricing**: single price per product, `money(products.price / 1.2)` prefixed with "From" as brand voice. No price-min, no BTO fallback.
3. **Wholesale customers**: retail-only on tiles. Wholesale pricing remains on the detail page unchanged. No `session("customerType")` branch in the card renderer.
4. **Eyebrow labels** ("2-Screen · Vertical" etc.): derive from SKU via a small `Select Case` block. Screen count from the suffix digit, style from trailing letters:
   - `v` → Vertical · `h` → Horizontal · `p` → Pyramid · `s` → Square
   - `sp` / `rp` → Pole · `r` → Side-by-side · `8r` → 2-over-2 quad
   - Unmatched SKU → fall through to a sensible default (e.g. just "N-Screen").
5. **Filename**: keep `CUSTOMCAT-stands.asp`.
6. **What's removed from the live page**:
   - "Discover Why Synergy Stand" question-mark CTA box (lines 554-567).
   - "Save Money, Get Free Cables & Free Delivery with a Bundle" callaction (lines 572-598) — replaced by the fuller Bundle cross-link section in the mockup.
   - Legacy Bootstrap-3 grid wrapper (`.bg-smog .product-grid .container .row`).
   - Sort/pagination/Solr/mobile-override ASP scaffolding.

### 5.5 Verification
1. Load `/stands/` on `amz.` staging — visual regression vs [`redesign/stands.html`](redesign/stands.html) at desktop and mobile widths.
2. Click each of the 12 tiles → confirm each links to the correct product detail page and add-to-cart still works from the detail page (header cart count updates).
3. Load the page as a wholesale customer (`session("customerType")=1`) — confirm tiles show retail "From £X" only and the detail page still applies the trade price.
4. Spot-check other category pages (monitors, computers, bundles, arrays) — confirm unchanged (shared includes untouched).
5. View source: chrome and footer markup should match the homepage and trading mockup.
6. Page weight should be lighter than the current version.

## 6. Applying to other category pages

The same skeleton + decisions apply to other `CUSTOMCAT-*` pages once their mockups are ready:

| Page | File | Category ID | Notes |
|---|---|---|---|
| Stands | [`shop/pc/CUSTOMCAT-stands.asp`](shop/pc/CUSTOMCAT-stands.asp) | 5 | First page — see §5 |
| Computers | [`shop/pc/CUSTOMCAT-computers.asp`](shop/pc/CUSTOMCAT-computers.asp) | 14 | Mockup: [`redesign/trading.html`](redesign/trading.html) (partial fit — has compare table, benchmark panels) |
| Trading computers | [`shop/pc/CUSTOMCAT-tradingcomputers.asp`](shop/pc/CUSTOMCAT-tradingcomputers.asp) | TBC | May converge with / replace the computers page |
| Monitors | [`shop/pc/CUSTOMCAT-monitors.asp`](shop/pc/CUSTOMCAT-monitors.asp) | TBC | Mockup TBD |
| Bundles 1/2/3 | [`shop/pc/CUSTOMCAT-bundles1.asp`](shop/pc/CUSTOMCAT-bundles1.asp) etc. | 5 / TBC | Multi-step builder — different shape, likely keeps its own dispatcher variant |
| Arrays 1/2/3 | [`shop/pc/CUSTOMCAT-arrays1.asp`](shop/pc/CUSTOMCAT-arrays1.asp) etc. | TBC | Multi-step builder — same caveat |

**Open per-page questions** we'd ask again when we get to each:
- Grouping scheme (SKU derivation may or may not translate).
- Whether any page-specific features from the mockup require extra data (e.g. benchmarks for trading computers, compare table rows).
- Whether any tile variant needs Buy Now / stock state (bundles/arrays builder pages probably do).

## 7. What we're not changing

- [`shop/pc/pcShowProducts.asp`](shop/pc/pcShowProducts.asp), [`shop/pc/pcShowProduct-Standard.asp`](shop/pc/pcShowProduct-Standard.asp), [`shop/pc/pcGetPrdPrices.asp`](shop/pc/pcGetPrdPrices.asp) — still used by every non-migrated category page and by product detail flow.
- Product detail pages.
- Admin panel (`shop/130707/`).
- Cart / checkout / payment.
- Database schema.
- [`css/mm-site.css`](css/mm-site.css) — no new classes required; all card/grid/hero styles already landed with the homepage rollout.

## 8. Alignment with vendor-shutdown remediation

Reducing the ProductCart surface area on storefront pages is aligned with the Jan 2027 vendor sunset work in [`maintain.md`](maintain.md): every page that stops depending on `pcShowProducts`, `pcGetPrdPrices`, `pcCheckPricingCats`, and the settings/sort/Solr scaffolding is one less thing to keep working after the vendor shuts down. The admin-panel vendor phone-home calls are a separate cleanup (covered in `maintain.md`), but the storefront simplifications here are a helpful tailwind.
