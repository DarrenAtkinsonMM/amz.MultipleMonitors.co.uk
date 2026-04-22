# Bundles builder page — migration to single-page 2026 redesign

## Context

The current bundle configurator at [/bundles/](shop/pc/CUSTOMCAT-bundles1.asp) walks customers through three sequential page loads — stand → screens → PC — each hitting the DB for a hard-coded category ID (5 / 6 / 14) and passing picks forward as `sid` / `mid` / `cid` querystrings. Step 3 then redirects to the matching PC product page, carrying those IDs through so the product page can display the bundle context.

The new mockup at [redesign/bundles.html](redesign/bundles.html) collapses the whole flow into one page with a JS-driven stepper (stand → screens → computer), a running sidebar totalling the bundle, and a single "Configure PC & Order Bundle" CTA that deep-links into the already-live [viewPrd-TraderPC-bundle-v2.asp](shop/pc/viewPrd-TraderPC-bundle-v2.asp). The mockup ships with hardcoded `BUNDLE_CONFIG` JS arrays (12 stands, 6 screens, 4 computers) — good for design, but the prices are frozen.

**Goal**: bring the mockup to production on `/bundles/`, keep the mockup's hardcoded curated content (names, short bullets, composite images), but make **price** live against the products table so admin price changes flow through without touching code. The arrays keep their structure; the `id` field switches from mockup strings (`'s2v'`, `'ultra'`) to real `products.idProduct` integers so the CTA URL composes cleanly and the DB lookup is trivial.

We're matching the pattern already proved on [customcat-stands.asp](shop/pc/customcat-stands.asp) ([category-redesign-plan.md](category-redesign-plan.md)): classic ASP opens, queries DB once with `GetRows()`, renders the static mockup content, and writes the product array into a `<script>` block server-side.

**This plan is about the builder page only.** The companion plan at [bundle-pages-redesign-plan.md](bundle-pages-redesign-plan.md) covers the bundle end-pages (the per-PC product pages the builder hands off to).

## Approach

**New file**: [shop/pc/bundlebuilder.asp](shop/pc/bundlebuilder.asp) — single-page ASP that replaces the 3-step flow.

**URL rewrite**: update the `Bundles Category` rule in [web.config:191-194](web.config#L191-L194) to rewrite `^bundles/$` to `/shop/pc/bundlebuilder.asp`. Leave the orphaned `/bundles-2/` and `/bundles-3/` rewrite rules and the three `CUSTOMCAT-bundles*.asp` files in place for now — harmless dead-ends; tidy up in a later pass.

**Hydration pattern** — two-stage:
1. ASP defines the *static* shape of each array (id, name, screens/discount/six/eight, bullets, images) as VBScript arrays at the top of the page.
2. A single SQL batch pulls `idProduct, price` for every ID referenced and stores it in a VBScript dictionary. During JSON emission, each item's price is looked up and injected; any missing ID is skipped with `LogErrorToDatabase()`.

Result: one trip to the DB per page load (roughly 22 products in a single `IN (...)` query); the JS state machine and rendering from the mockup stays 100% intact.

**VAT**: DB `products.price` is VAT-inclusive. The mockup and [inc_bundleContext.asp:130](shop/pc/inc_bundleContext.asp#L130) display ex-VAT — divide by `MM_VAT_RATE = 1.2` before writing into the JS array. Keep the existing `discount` values (25/50/100) ex-VAT — they already match that convention.

**CTA routing** (per decision — fall back to legacy PC pages):

| mockup key | idProduct | CTA target |
|---|---|---|
| ultra   | 306 | `/products/ultra-multi-monitor-pc/?sid=&mid=&cid=306` |
| extreme | 307 | `/products/extreme-multi-screen-computer/?sid=&mid=&cid=307` |
| trader  | 333 | `/shop/pc/viewPrd-TraderPC-bundle-v2.asp?sid=&mid=&cid=333` |
| pro     | 343 | `/products/trader-pro-pc/?sid=&mid=&cid=343` |

(These destinations are lifted verbatim from the current [CUSTOMCAT-bundles3.asp:5-21](shop/pc/CUSTOMCAT-bundles3.asp#L5-L21) dispatcher.) Store the per-computer CTA URL as an extra `cta` field on each item in the `computers` array so the builder JS can read it off the picked computer object rather than hard-coding the switch in two places.

## Files to create / modify

| File | Change |
|---|---|
| [shop/pc/bundlebuilder.asp](shop/pc/bundlebuilder.asp) | **NEW** — single-page builder, ports mockup body into ASP with server-hydrated `BUNDLE_CONFIG` |
| [web.config](web.config) | Update rule at line 191-194: rewrite `^bundles/$` → `/shop/pc/bundlebuilder.asp` |
| [css/mm-site.css](css/mm-site.css) | **Likely add** scoped rules for builder-specific components (`.stepper`, `.stage`, `.stage-head`, `.stand-grid.cols-*`, `.bundle-sidebar`, `#mmb-viz`, etc.) — lift verbatim from the `<style>` block in [redesign/bundles.html](redesign/bundles.html). Scope under `.mm-site` to match the existing pattern |
| CUSTOMCAT-bundles1.asp / 2 / 3 | **Leave untouched** — orphaned but preserved in case any external/printed link still points at `/bundles-2/` or `/bundles-3/` |

## Implementation detail

### 1. ASP skeleton

Mirror [customcat-stands.asp:1-17](shop/pc/customcat-stands.asp#L1-L17):

```asp
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "bundlebuilder.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<% Const MM_VAT_RATE = 1.2 %>
```

Then `header_wrapper.asp` → `<div class="mm-site">` → body → `</div>` → `footer_wrapper.asp` exactly as the stands page does.

### 2. Static config tables (VBScript)

Three 2-D arrays at the top of the page, one per category. Each row carries everything except price:

```asp
' stands: id | name                | screens | discount | img                          | arrayimg
Dim mmStands
mmStands = Array( _
  Array(287, "Dual Vertical",     2,  25, "/images/bundles/bun-s2v-med.png",  "s2v" ), _
  Array(288, "Dual Horizontal",   2,  25, "/images/bundles/bun-s2h-med.png",  "s2h" ), _
  Array(312, "Triple Horizontal", 3,  25, "/images/bundles/bun-s3h-med.png",  "s3h" ), _
  Array(324, "Triple Pyramid",    3,  25, "/images/bundles/bun-s3p-med.png",  "s3p" ), _
  Array(313, "Quad Square",       4,  50, "/images/bundles/bun-s4s-med.png",  "s4s" ), _
  ' ... 12 stands total
)
```

Analogous `mmScreens` (with `desc1/desc2/desc3` bullets) and `mmComputers` (with `six`, `eight`, `cta`, bullets, `img`, `bunimg`).

**IDs to pin down before coding** — known from [category-redesign-plan.md](category-redesign-plan.md) / [CUSTOMCAT-bundles1.asp](shop/pc/CUSTOMCAT-bundles1.asp):

- Stands (cat 5): 287 Dual · 312 Triple · 313 Quad · 314 Six Multi-Pole · 318 Five Pyramid · 319 Eight · 324 Triple Pyramid · 325 Quad Pyramid · 327 Quad Horizontal · 338 Six — **12 stands in mockup, 10 known → Darren to supply the 2 missing** (Dual Horizontal 2v-vs-2h distinction + Quad Multi Pole variant)
- Monitors: 304 (21.5") · 317 (24") known; **6 in mockup → Darren to supply 4 more IDs**
- Computers: 306 Ultra · 307 Extreme · 333 Trader · 343 Trader Pro — all four known

A short SQL against `products` filtered by `sku LIKE 'stand%'` / `'mon%'` etc. can confirm these quickly at implementation time.

### 3. One-shot DB hydration

Collect all IDs into a single `IN (...)` query — one round-trip, then build a dictionary keyed by `idProduct`:

```asp
Dim mmPriceDict : Set mmPriceDict = Server.CreateObject("Scripting.Dictionary")

Dim mmAllIds, mmItem
mmAllIds = ""
For Each mmItem In mmStands    : mmAllIds = mmAllIds & mmItem(0) & "," : Next
For Each mmItem In mmScreens   : mmAllIds = mmAllIds & mmItem(0) & "," : Next
For Each mmItem In mmComputers : mmAllIds = mmAllIds & mmItem(0) & "," : Next
mmAllIds = Left(mmAllIds, Len(mmAllIds) - 1)   ' strip trailing comma

Dim mmSql, mmRs
mmSql = "SELECT idProduct, price FROM products " & _
        "WHERE idProduct IN (" & mmAllIds & ") " & _
        "  AND active = -1 AND removed = 0"
Set mmRs = connTemp.Execute(mmSql)
Do While Not mmRs.EOF
  mmPriceDict.Add CLng(mmRs("idProduct")), CDbl(mmRs("price"))
  mmRs.MoveNext
Loop
mmRs.Close : Set mmRs = Nothing
```

All IDs are numeric (sourced from our own arrays, not user input) — no injection risk.

### 4. JS emission

A helper sub `mmEmitBundleArray` walks each static array, looks up the price in the dictionary, divides by VAT, and emits the JS object literal. Items whose ID isn't in the dictionary are skipped with `LogErrorToDatabase("bundlebuilder missing product id " & id)`.

```asp
Sub mmEmitStandJS(row)
  Dim id, px
  id = CLng(row(0))
  If Not mmPriceDict.Exists(id) Then
    Call LogErrorToDatabase()   ' page-wide err handler sub in common.asp
    Exit Sub
  End If
  px = Int((mmPriceDict(id) / MM_VAT_RATE) + 0.5)   ' round to whole pounds, matches mockup style
  Response.Write "      { id:" & id & ", name:""" & row(1) & """, price:" & px & _
                 ", screens:" & row(2) & ", discount:" & row(3) & _
                 ", img:""" & row(4) & """, arrayimg:""" & row(5) & """ }," & vbCrLf
End Sub
```

Similar subs for screens (adds `desc1/2/3`) and computers (adds `six`, `eight`, `cta`, `desc1/2/3`, `bunimg`).

### 5. JS tweaks in the ported mockup

Everything in [redesign/bundles.html:528-920](redesign/bundles.html) ports as-is with three tiny changes:

1. The `BUNDLE_CONFIG` literal is replaced by the ASP-emitted block.
2. **IDs are now integers** — any `===` comparison against string IDs in the mockup JS (search for `state.stand.id`, `state.screens.id`, `state.computer.id` comparisons) will keep working; JS doesn't care about int-vs-string at runtime. Worth a quick grep to confirm no literal string IDs sneak in.
3. The final CTA builds its URL from `state.computer.cta`:
   ```js
   window.location = state.computer.cta
     + '?sid=' + state.stand.id
     + '&mid=' + state.screens.id
     + '&cid=' + state.computer.id;
   ```
   Replaces whatever placeholder anchor the mockup has on `#mmb-cta`.

### 6. CSS migration

The mockup's `<style>` block (in-file, ~500 lines) needs to land in [css/mm-site.css](css/mm-site.css) under the `.mm-site`-scoped half, matching the way the stands-page builder components were already added. Most rules already exist (`.container`, `.reveal`, `.truststrip`, `.hero`, `.pillars`, `.bundle-card`). The **new** classes specific to the builder — `.stepper`, `.stage`, `.stand-grid.cols-4/3/2`, `.bundle-sidebar`, `#mmb-viz` SVG, `.mmb-pill`, `.bp-sidebar` — need to be extracted and prefixed with `.mm-site` where they aren't uniquely-named.

### 7. URL rewrite update

In [web.config](web.config) change:
```xml
<rule name="Bundles Category" stopProcessing="true">
  <match url="^bundles/$" ignoreCase="true" />
  <action type="Rewrite" url="/shop/pc/CUSTOMCAT-bundles1.asp" appendQueryString="true" />
</rule>
```
to point at `/shop/pc/bundlebuilder.asp`. No other web.config edits needed — the `bundles-2` / `bundles-3` rules keep working as an escape hatch but nothing links there any more.

## Verification

1. **Build runs** — hit `http://amz.multiplemonitors.co.uk/bundles/` and confirm the new page renders without an ASP error. Check `ErrorLog` for any "missing product id" entries (indicates the array has an ID not in the DB).
2. **Prices live** — in admin, change a stand's price by £1, reload `/bundles/`, confirm the new price shows in both the stand card and the sidebar total. Revert.
3. **State machine** — pick stand → screens → PC, confirm sidebar total matches `(stand + screens × qty + computer + upgrade − discount)` and the "Configure" CTA activates only at step 3.
4. **CTA routing** — pick each of the 4 computers in turn, confirm the URL in the browser after clicking Configure:
   - Trader → `/shop/pc/viewPrd-TraderPC-bundle-v2.asp?sid=…&mid=…&cid=333` (loads the new v2 page)
   - Ultra/Extreme/Trader Pro → their `/products/…/` legacy slug with `sid/mid/cid` preserved
5. **Deep-link round-trip** — from the v2 product page, click "Change stand" → returns to `/bundles/?sid=…&mid=…&cid=…&edit=stand` (already supported by the mockup's query-string state restoration at [redesign/bundles.html:870](redesign/bundles.html#L870) — verify it still works).
6. **Missing product soft-fail** — temporarily set `active=0` on one stand in the DB, reload, confirm that stand is silently skipped rather than breaking the page, and an entry appears in `ErrorLog`. Restore.
7. **VAT + wholesale** — the page intentionally shows retail ex-VAT only (matches stands-page convention). Cart + checkout still branch on `Session("customerType")` downstream in `instPrd.asp` / cart pages — no change there.
8. **Mobile** — the mockup's sidebar stacks below the picker at narrow widths; re-check after porting the CSS into `mm-site.css`, since scope wrapping can subtly change media-query specificity.

## Out of scope (flag for follow-up)

- Moving the non-Trader PCs to new v2 bundle end-pages (`viewPrd-UltraPC-bundle-v2.asp` etc.) — parallel workstream tracked in [bundle-pages-redesign-plan.md](bundle-pages-redesign-plan.md); the `cta` field on each computer makes adding them later a one-line change.
- DB-backed bullet descriptions for screens/computers — keep hardcoded for now as agreed.
- Removing the orphaned `CUSTOMCAT-bundles2.asp` / `3.asp` and their rewrite rules — defer until we're confident no link still hits `/bundles-2/` or `/bundles-3/`.
