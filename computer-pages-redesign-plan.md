# Computer product pages — live-priced redesign

## Context

The current [shop/pc/viewPrd-TraderPC.asp](shop/pc/viewPrd-TraderPC.asp) is a heavily customised, hand-built product page for product ID **333**. It produces the UX we want but all 16 option groups' prices are hardcoded into `<option>` `title` attributes and display text. When option pricing changes in the admin panel, the ASP file must be manually edited — an ongoing maintenance pain.

[shop/pc/viewPrd-Computers.asp](shop/pc/viewPrd-Computers.asp) solves the live-pricing problem by using ProductCart's built-in `pcs_OptionsN` renderer (defined at [shop/pc/viewPrdCode.asp:2749](shop/pc/viewPrdCode.asp#L2749)), but at the cost of the UI — it produces generic `<select>` dropdowns we don't want.

We want both: the modern layout of [redesign/traderpc.html](redesign/traderpc.html) *and* live option pricing from the database. The approach proven on [shop/pc/CUSTOMCAT-stands.asp](shop/pc/CUSTOMCAT-stands.asp) — bypass the ProductCart render pipeline, query the DB directly, own the HTML — is the right pattern, extended here to keep ProductCart's **cart submission contract** (so `instPrd.asp` still works unchanged).

Scope: four computer product pages (Trader PC is the first; three other machines follow). Marketing copy and specs differ per machine; the configurator shape and shared marketing sections are reused.

## Approach

**One ASP file per machine, sharing marketing-section includes, driving the options configurator from live DB data, submitting to the existing `instPrd.asp` cart flow.**

The pattern is a mash-up of:
- The stands-page direct-DB + custom-render approach (no `pcShowProducts.asp` chain).
- The `pcs_makeOptionBox` option-pricing SQL (lifted into our own render loop).
- The mockup's radio-card configurator, sticky summary, and live total (ported from inline `<style>`/`<script>` into `mm-site.css` and a per-page JS block).
- ProductCart's cart contract: POST to `/shop/pc/instPrd.asp` with `idproduct`, `quantity`, `OptionGroupCount`, and `idOption1..idOptionN` form fields whose values are valid `idoptoptgrp` IDs. `instPrd.asp` re-queries live prices on its own ([shop/pc/instPrd.asp:717-749](shop/pc/instPrd.asp#L717-L749)), so whatever prices we show on the page are display-only — the server is authoritative.

## File plan

### New / rewritten ASP files

| File | Purpose |
|---|---|
| `shop/pc/viewPrd-TraderPC.asp` | Rewrite — Trader PC (idProduct 333). Built first, acts as the template. |
| `shop/pc/viewPrd-TraderPC-old.asp` | Backup of the current file (follows the `CUSTOMCAT-stands-old.asp` convention already set). |
| `shop/pc/viewPrd-<machine-2>.asp` | Replicate for machine 2 after Trader PC is verified live. |
| `shop/pc/viewPrd-<machine-3>.asp` | Same. |
| `shop/pc/viewPrd-<machine-4>.asp` | Same. |

### New shared includes (extracted from the mockup)

| File | Section it renders |
|---|---|
| `shop/pc/inc_trustStripTrader.asp` | BBC / Trustpilot / Since 2008 / Published benchmarks badges |
| `shop/pc/inc_firmsStrip.asp` | "Trusted by hedge funds, prop desks…" logo row |
| `shop/pc/inc_tradersReviews.asp` | 3-card Trustpilot reviews grid (same set across all four pages) |
| `shop/pc/inc_tradersFaq.asp` | 6–8 FAQ accordion (generic trader questions) |
| `shop/pc/inc_bundleBand.asp` | "Bundle & save" dark band with breakdown card |
| `shop/pc/inc_darrenCTA.asp` | "15-minute call with Darren" CTA block |
| `shop/pc/inc_guideBand.asp` | "Free trader's buying guide" email capture band |
| `shop/pc/inc_stickyCTA.asp` | Sticky bottom-right "Configure" overlay. Takes machine name + starting-price via pre-set VBScript vars. |

### CSS additions

Port the mockup's inline `<style>` block (lines 25–600+ of [redesign/traderpc.html](redesign/traderpc.html)) into [css/mm-site.css](css/mm-site.css) under the existing `.mm-site` scope. New component classes to add:

- `.pd-hero`, `.pd-hero-grid`, `.pd-gallery`, `.pd-gallery__main/__chip/__sku/__thumbs`, `.pd-thumb`
- `.pd-buybox`, `.pd-tp`, `.pd-price`, `.pd-cutoff`, `.pd-incl`, `.pd-cta`, `.pd-foot`
- `.configurator`, `.cfg-head`, `.cfg-grid`, `.cfg-options-wrap`, `.cfg-row`, `.cfg-row__head/__label/__selected/__help`, `.cfg-options`, `.cfg-option` (+ `.is-selected`, `.opt-name`, `.opt-price.std`, `.opt-price.inc`)
- `.cfg-summary`, `.cfg-summary__card/__head/__list/__cta/__trust`, `.cfg-total`, `.cfg-vat`, `.cfg-impact`, `.cfg-impact--cpu/--gpu`, `.cfg-impact__row/__stars/__mon/__ctx`
- `.full-spec`, `.spec-full`, `.spec-full__grid`, `.spec-row`
- `.xlink-card` (upsell to next-tier card), `.section-head-narrow`, `.container-narrow`
- `.talk-link`

Leave the already-present `.bundle*`, `.truststrip`, `.reviews`, `.faq-*`, `.darren-inline`, `.guide-band-section`, `.sticky-cta`, `.btn*`, `.reveal` alone — reused as-is.

## Technical pattern — the Trader PC page, end to end

### Head of file: includes + DB load

```asp
<%@ LANGUAGE=VBScript %>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
  Const PRODUCT_ID = 333       ' Trader PC
  Const VAT_RATE   = 1.2

  ' --- 1. Load product base row ---
  Dim prdSql, prdRs
  Dim pName, pSku, pBasePrice, pSmallImg

  prdSql = "SELECT description, sku, price, smallImageUrl " & _
           "FROM products " & _
           "WHERE idProduct = " & PRODUCT_ID & _
           "  AND active = -1 AND removed = 0"
  Set prdRs = connTemp.Execute(prdSql)
  If Not prdRs.EOF Then
    pName      = prdRs("description") & ""
    pSku       = prdRs("sku") & ""
    pBasePrice = CDbl(prdRs("price"))
    pSmallImg  = prdRs("smallImageUrl") & ""
  End If
  prdRs.Close : Set prdRs = Nothing

  ' --- 2. Load option groups for this product (mirrors pcs_OptionsN) ---
  Dim ogSql, ogRs, ogCount, ogRows
  ogSql = "SELECT DISTINCT og.idOptionGroup, og.OptionGroupDesc, " & _
          "       po.pcProdOpt_Required, po.pcProdOpt_Order " & _
          "FROM pcProductsOptions po " & _
          "INNER JOIN optionsGroups og ON og.idOptionGroup = po.idOptionGroup " & _
          "INNER JOIN options_optionsGroups oog ON oog.idOptionGroup = og.idOptionGroup " & _
          "                                   AND oog.idProduct = po.idProduct " & _
          "WHERE po.idProduct = " & PRODUCT_ID & " " & _
          "ORDER BY po.pcProdOpt_Order, og.OptionGroupDesc"
  Set ogRs = connTemp.Execute(ogSql)
  If Not ogRs.EOF Then ogRows = ogRs.GetRows() : ogCount = UBound(ogRows, 2) + 1 Else ogCount = 0
  ogRs.Close : Set ogRs = Nothing

  ' --- 3. Options-per-group query is re-run inside the render loop below
  '        (same SQL as pcs_makeOptionBox, unchanged) ---
%>
```

### Configurator render loop (per option group)

The mockup's 7 visible groups map onto the DB's many ProductCart option groups via a **per-machine whitelist** (VBScript array of `idOptionGroup` IDs in the order we want them shown, plus a per-group presentation hint: `radio` or `extra`). Everything not in the whitelist is dropped from the visible configurator; we either remove those groups from the product in admin or emit them as hidden inputs defaulted to the "standard" option.

Render each whitelisted group as:

```html
<div class="cfg-row" data-group="cpu">
  <div class="cfg-row__head">
    <div class="cfg-row__label"><span class="n">1</span>Processor</div>
    <div class="cfg-row__selected" data-selected>[default option descrip]</div>
  </div>
  <p class="cfg-row__help">[per-group help copy, hardcoded per page]</p>
  <div class="cfg-options" role="radiogroup">
    <!-- For each option in the group, from options_optionsGroups ORDER BY sortOrder: -->
    <button type="button"
            class="cfg-option [is-selected for first/default]"
            data-name="[optionDescrip]"
            data-delta="[round((price - defaultPrice) / VAT_RATE)]"
            data-idoptoptgrp="[idoptoptgrp]">
      <span class="opt-name">[optionDescrip]</span>
      <span class="opt-price [std|inc]">[Included | + £X]</span>
    </button>
    <!-- ... -->
  </div>
  <input type="hidden" name="idOption<N>" value="[default idoptoptgrp]">
</div>
```

The SQL for the inner loop is identical to `pcs_makeOptionBox` ([shop/pc/viewPrdCode.asp:2855-2865](shop/pc/viewPrdCode.asp#L2855-L2865)):

```sql
SELECT oog.idoptoptgrp, oog.price, oog.Wprice, oog.sortOrder, oog.InActive,
       o.idOption, o.optionDescrip
FROM options_optionsGroups oog
INNER JOIN options o ON oog.idOption = o.idOption
WHERE oog.idOptionGroup = <groupID>
  AND oog.idProduct     = 333
ORDER BY oog.sortOrder, oog.price, o.optionDescrip
```

Price-per-customer-type branch matches the ProductCart one at [shop/pc/viewPrdCode.asp:2918-2924](shop/pc/viewPrdCode.asp#L2918-L2924): wholesale customers (`Session("customerType") = 1`) use `Wprice`, else `price`.

### Multi-select "Extras" group

The mockup's Extras section visually shows four independent toggles (Wireless KB+mouse, Speakers, Wi-Fi+BT, Bootable backup). In the DB these already exist as **four separate option groups** on product 333 (confirmed in the first investigation: groups Mouse & Keyboard, Speakers, WiFi Card, Backup Drive). Render them under a single `.cfg-row[data-group="extras"]` heading but produce **four hidden inputs** (`idOption8`, `idOption9`, `idOption11`, etc.), each toggling between the group's "No / none" `idoptoptgrp` (default) and its "Yes" `idoptoptgrp` when the card is clicked. The pairing {no-idoptoptgrp, yes-idoptoptgrp} per extra is a small per-machine JS/VBScript lookup table built from the DB load.

### Form submission

Wrap the whole configurator (options + sticky summary) in one form:

```html
<form method="post" action="/shop/pc/instPrd.asp" id="cfgForm">
  <input type="hidden" name="idproduct"         value="333">
  <input type="hidden" name="quantity"          value="1">
  <input type="hidden" name="OptionGroupCount"  value="<%= ogCount %>">
  <!-- idOption1..idOptionN rendered inline per group, above -->
  …configurator markup…
  <button type="submit" class="btn btn-primary btn-lg cfg-summary__cta">
    <i class="fa fa-shopping-basket"></i> Add to basket
  </button>
</form>
```

The current mockup's "Add to basket" is a plain `<a>` — it becomes a form submit. The sticky summary sidebar lives inside this form.

### Client-side JS (per-page inline)

- **Selection**: on `.cfg-option` click, toggle `.is-selected` within the row, update the row's `data-selected` text, set the row's hidden `idOption<N>` input to `dataset.idoptoptgrp`.
- **Running total**: sum of all `data-delta` values across selected options, added to `BASE_EX = pBasePrice / 1.2`. Render ex-VAT and inc-VAT, update `data-total-ex`, `data-total-inc`, and `data-sticky-price`.
- **Summary list**: mirror `data-selected` into the right-hand summary list's `data-sum` spans.
- **Impact stars + GPU monitor-list**: lookup table keyed on `idoptoptgrp` (per-machine JS config object) → `{ speed, mt, gfx, ai, mons }`. Lifted from the mockup's substring matching but keyed on IDs for robustness.
- **Auto-upgrade rows** in full-spec (motherboard/PSU/cooler/fans vary with CPU choice): per-machine JS map `cpuIdOptOptGrp → { mobo, cooler, psu, fans }`, applied into `data-spec` spans. (See open decision below — confirm whether these are real upgrades or purely aspirational copy.)
- **Gallery thumbs**, **sticky CTA visibility**, **reveal-on-scroll** — port from mockup unchanged.

### Chrome

Same pattern as `CUSTOMCAT-stands.asp`:

```asp
<!--#include file="header_wrapper.asp"-->
<div class="mm-site">
  …hero / trust strip / specs / configurator / full-spec / benchmarks / upsell / reviews / FAQ / guide / sticky CTA…
</div>
<!--#include file="footer_wrapper.asp"-->
```

[shop/pc/header_wrapper.asp](shop/pc/header_wrapper.asp) already loads [css/mm-site.css](css/mm-site.css) via [shop/pc/inc_headerCSS.asp](shop/pc/inc_headerCSS.asp); no loader changes needed.

## Per-page content (hardcoded, differs per machine)

- Hero image, gallery thumbs, SKU caption
- Buy-box pitch copy, price-from banner, `.pd-incl` three-up (these are machine-specific)
- `<title>`, `<meta description>`
- 6 base-config spec cards (`.spec-card`)
- "In the box" chip list
- Per-group `.cfg-row__help` copy (one sentence per option group)
- Full-spec row set (some static, some driven by JS)
- CPU benchmark panel (per-CPU scores, bars) — the four CPU options on this machine
- Upsell card (link to the next-tier machine, or bundle band if this is the top tier)
- Per-machine JS config objects: impact-star ratings table, auto-upgrade table, extras `{no,yes}` id pairs

## Per-page shared content (via `<!--#include-->`)

Chrome, trust strip, firms strip, bundle band, reviews, FAQ, Darren CTA, guide band, sticky CTA.

## Rollout order

1. **Port CSS first**: copy the mockup's inline `<style>` block into `mm-site.css` under `.mm-site`. No behaviour change yet; the mockup pages keep rendering the same because the inline styles still win over site CSS in the mockup file.
2. **Extract shared marketing-section includes** from the mockup's HTML into the `inc_*.asp` files listed above. Keep copy identical to the mockup.
3. **Build `viewPrd-TraderPC-v2.asp`** as a new file (not yet replacing the live one). Wire the DB query, configurator render, form, and JS config.
4. **Local test on `amz.` staging** — click through every option, verify running total matches, submit form, verify the cart shows the right line items and ProductCart-calculated prices match.
5. **Admin-side price change test**: edit any option price in the admin panel, reload the product page, confirm the new price flows through without touching the ASP file (the whole point of this work).
6. **Cutover**: rename `viewPrd-TraderPC.asp` → `viewPrd-TraderPC-old.asp`, rename `-v2.asp` → live filename.
7. **Replicate for the other three machines** once the pattern is proven. Each follow-up is mostly content + per-machine JS config; the configurator logic is copy-paste.

## Open decisions (to resolve during implementation, not blocking the plan)

1. **Product IDs for the other three machines.** Trader PC is 333; the others need identifying (check admin product list or the `products` table for computer-category products).
2. **Option-group whitelist per machine.** Current Trader PC has 16 groups in the DB; the new design exposes 7 visible slots (CPU, RAM, Screens/GPU, Storage, OS, Warranty, Extras-as-4). For the dropped groups (e.g. DVD drive, MS Office, Second hard drive), pick one: (a) remove the link in admin from `pcProductsOptions`, or (b) keep in DB and emit as hidden `<input>` defaulted to the "standard" `idoptoptgrp` so the cart line is consistent. **Recommendation**: (b) — less risk of admin-side breakage, reversible, cart still shows consistent base config.
3. **Auto-upgrade rows in full-spec** (motherboard/PSU/cooler swap with CPU). Is this real (i.e. the build genuinely changes) or aspirational copy? If real, the mapping belongs in the DB long-term; for the first pass a per-page JS object is fine. Needs a quick check with Darren.
4. **Reviews set.** Are the same 3 reviews shown on all four machine pages (via the shared include), or does each machine filter to reviews that mention it? The mockup copy references "Trader PC" and "Trader Pro" by name. **Recommendation**: one shared set referencing the *Trader range* generically (avoid machine-specific review text), so the include stays simple.
5. **`configOnly` check.** The stands query filters `configOnly = 0`. For computer products, several of them may be `configOnly = -1` (built-to-order). Needs the filter loosened — `removed = 0 AND active = -1` is probably enough.

## Verification

Manual browser testing on `amz.multiplemonitors.co.uk` staging:

- **Golden path**: load `viewPrd-TraderPC.asp` → scroll through page → change each option group → verify `.cfg-summary__list`, `data-total-ex`, `data-total-inc`, and the sticky-CTA price all update consistently → click "Add to basket" → inspect cart → confirm each line item's price matches (`instPrd.asp` does its own DB lookup, so this is the authoritative check on price parity).
- **Admin-side edit**: change one option's retail price in admin (`options_optionsGroups.price`), reload the product page with a cache-buster, verify the new price appears in both the card and the summary total.
- **Wholesale customer (`Session("customerType") = 1`)**: log in as a wholesale user and confirm `Wprice` is used throughout.
- **Mobile viewport**: sticky summary should switch to a bottom-pinned bar (mockup behaviour — port as CSS media query). Verify gallery thumbs, configurator radio cards, FAQ accordion all behave.
- **Accessibility**: keyboard-nav through the configurator — `.cfg-option` buttons are focusable, Enter/Space toggles, aria-selected updates; stars have `aria-label="N of 5"`.
- **Chrome scroll listeners** (sticky CTA, reveal-on-scroll) don't leak across pages or double-fire after hot-reload.
- **Regression check**: the other three unmigrated computer pages (e.g. `viewPrd-Computers.asp`) still render correctly — no legacy CSS accidentally clobbered by new rules.

## Critical files referenced

- [shop/pc/viewPrd-TraderPC.asp](shop/pc/viewPrd-TraderPC.asp) — current hardcoded page to replace
- [shop/pc/viewPrd-Computers.asp](shop/pc/viewPrd-Computers.asp) — reference for ProductCart's live-priced approach
- [shop/pc/viewPrdCode.asp:2749-2924](shop/pc/viewPrdCode.asp#L2749-L2924) — `pcs_OptionsN` / `pcs_makeOptionBox` SQL we're lifting
- [shop/pc/instPrd.asp:475-749](shop/pc/instPrd.asp#L475-L749) — cart submission handler; consumes `idOption1..idOptionN` and re-queries prices
- [shop/pc/CUSTOMCAT-stands.asp](shop/pc/CUSTOMCAT-stands.asp) — direct-DB + bypass-ProductCart template we're extending
- [shop/pc/header_wrapper.asp](shop/pc/header_wrapper.asp), [shop/pc/footer_wrapper.asp](shop/pc/footer_wrapper.asp) — chrome
- [shop/pc/inc_headerCSS.asp](shop/pc/inc_headerCSS.asp) — CSS loader (already includes `mm-site.css` last)
- [css/mm-site.css](css/mm-site.css) — target stylesheet; new configurator classes go here under `.mm-site`
- [redesign/traderpc.html](redesign/traderpc.html) — mockup (visual + inline JS reference)
- [category-redesign-plan.md](category-redesign-plan.md), [chrome-rollout-plan.md](chrome-rollout-plan.md) — prior-rollout context
