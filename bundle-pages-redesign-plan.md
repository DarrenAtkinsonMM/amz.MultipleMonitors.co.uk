# Bundle end-page — architecture recommendation

## Context

Today the bundle flow is a 3-step wizard ([CUSTOMCAT-bundles1.asp](shop/pc/CUSTOMCAT-bundles1.asp) → stand, [bundles2](shop/pc/CUSTOMCAT-bundles2.asp) → monitors, [bundles3](shop/pc/CUSTOMCAT-bundles3.asp) → PC). When the user picks a PC in step 3 they land on the standalone PC product page (`viewPrd-TraderPC.asp`, etc.) with `?sid=<stand>&mid=<monitor>&cid=<pc>` in the querystring. That legacy page includes [shop/pc/bundle-breadcrumb.asp](shop/pc/bundle-breadcrumb.asp) above its own content as a "bundle ribbon" — it works but it's cosmetic, and the legacy `bundle-breadcrumb.asp` is also what hands [shop/pc/inc_headerDAJS.asp](shop/pc/inc_headerDAJS.asp) the data it needs to build the hidden `idproduct2`/`idproduct3`/`QtyM*`/`pCnt=3` form inputs that make the "Add to basket" action actually add three products with a `£25`/`£50`/`£100` discount.

In parallel with this work, the 3-step wizard is being replaced by a single-page bundle builder ([redesign/bundles.html](redesign/bundles.html) — stand / screens / PC picks all on one page, with a "Configure the PC" CTA at the end). It produces the **same `sid`/`mid`/`cid` querystring contract** on the way out, so the bundle end-page plan below is agnostic to which builder sent the user.

We've just landed [shop/pc/viewPrd-TraderPC-v2.asp](shop/pc/viewPrd-TraderPC-v2.asp) — the first of the four redesigned standalone PC pages — and now need to decide how the matching bundle end-page looks and is built. The mockup at [redesign/bundleproduct.html](redesign/bundleproduct.html) makes clear the bundle page is **bundle-first, not PC-first**: the hero, price, savings callout, sidebar summary and CTAs all represent the whole bundle, and the PC configurator is a sub-section of the page. The old "insert a breadcrumb header above the PC page" approach can't represent this.

## Recommendation

**Build one dedicated bundle ASP file per machine** (4 files), as siblings of the standalone v2 pages. Keep them lean by sharing the PC-specific marketing content (configurator, full spec, benchmarks, reviews, FAQ) via per-machine includes that **both** the standalone and bundle pages call into. Do **not** try to make the standalone v2 page also render as the bundle page — the hero, sidebar, CTAs, picks section, compatibility band and form-submission contract all differ enough that conditional branching would hurt readability without saving real work.

This matches the user's framing: "manually creating [a bundle-specific page per PC] is acceptable and manageable… I'd rather that than over-complicate the new computer pages."

### File plan

New ASP files (one per machine; Trader PC is the template):

| File | Purpose |
|---|---|
| `shop/pc/viewPrd-TraderPC-bundle-v2.asp` | Bundle end-page for `cid=333`. Takes `sid`, `mid`, `cid` querystrings. Built first as the template. |
| `shop/pc/viewPrd-<machine-2>-bundle-v2.asp` | Replicate for machine 2 after Trader PC is verified live. |
| `shop/pc/viewPrd-<machine-3>-bundle-v2.asp` | Same. |
| `shop/pc/viewPrd-<machine-4>-bundle-v2.asp` | Same. |

New shared includes (used by both standalone and bundle variants):

| File | Purpose |
|---|---|
| `shop/pc/inc_bundleContext.asp` | **The core bundle-data include.** Validates + parses `sid`/`mid`/`cid` querystrings (copy the `IsNumeric`/redirect pattern from [bundle-breadcrumb.asp:17-48](shop/pc/bundle-breadcrumb.asp#L17-L48)); queries `products` for live stand and monitor rows (description, price, imageUrl); derives monitor count from the stand name (Dual=2, Triple=3, Quad=4, Five=5, Six=6, Eight=8); maps the stand ID → bundle discount (`£25`/`£50`/`£100`, port the `Select Case bunBCsid` block from [bundle-breadcrumb.asp:52-141](shop/pc/bundle-breadcrumb.asp#L52-L141) but strip out the hardcoded image filenames — images come from the live DB rows now); emits a VBScript sub `mmEmitBundleHiddenInputs()` that writes the `idproduct2`/`idproduct3`/`QtyM<sid>`/`QtyM<mid>`/`pCnt=3` hidden inputs [inc_headerDAJS.asp:148-150](shop/pc/inc_headerDAJS.asp#L148-L150) expects so `instPrd.asp` processes the cart add as a bundle. |

New per-machine PC marketing includes — extracted from `viewPrd-TraderPC-v2.asp` so both variants share the same content:

| File | Extracted from |
|---|---|
| `shop/pc/inc_traderPC_specGrid.asp` | [viewPrd-TraderPC-v2.asp:322-387](shop/pc/viewPrd-TraderPC-v2.asp#L322-L387) — base-config spec grid + "in the box" chip list |
| `shop/pc/inc_traderPC_configurator.asp` | [viewPrd-TraderPC-v2.asp:100-177, 395-494](shop/pc/viewPrd-TraderPC-v2.asp#L100-L177) — `mmRenderOptionGroup` sub + the configurator render loop + form wrapper. Parameterised so the bundle page can swap in the bundle sidebar (`.bp-sidebar`) in place of the standalone `.cfg-summary`. |
| `shop/pc/inc_traderPC_benchmarks.asp` | [viewPrd-TraderPC-v2.asp:499-543](shop/pc/viewPrd-TraderPC-v2.asp#L499-L543) — per-machine CPU benchmark panels |
| `shop/pc/inc_traderPC_fullspec.asp` | The full-spec block (not yet in v2; will be added as part of this work per the mockup) |
| `shop/pc/inc_traderPC_reviews.asp` | [viewPrd-TraderPC-v2.asp:584-635](shop/pc/viewPrd-TraderPC-v2.asp#L584-L635) |
| `shop/pc/inc_traderPC_faq.asp` | [viewPrd-TraderPC-v2.asp:640-708](shop/pc/viewPrd-TraderPC-v2.asp#L640-L708) |
| `shop/pc/inc_traderPC_configuratorJS.asp` | [viewPrd-TraderPC-v2.asp:731-975](shop/pc/viewPrd-TraderPC-v2.asp#L731-L975) — the configurator / gallery / sticky / impact-stars JS. The bundle page reuses this verbatim, just with extra hooks to update the `bp-card` bundle-total fields. |

New bundle-only include:

| File | Purpose |
|---|---|
| `shop/pc/inc_bundleSidebar.asp` | The `.bp-sidebar` markup from the mockup — CPU/GPU impact panels (same as standalone) + `.bp-card` bundle breakdown with stand / screens / PC line items, subtotal, bundle discount, total, "You're saving" banner. Driven by the VBScript vars set by `inc_bundleContext.asp`. |

### Data flow

On a bundle end-page request (e.g. `viewPrd-TraderPC-bundle-v2.asp?sid=312&mid=304&cid=333`):

1. `inc_bundleContext.asp` validates the querystring (redirects to `/bundles/` on invalid/missing) and loads live stand + monitor DB rows, computes monitor count and bundle discount.
2. The bundle page loads the PC base row (same logic as the v2 standalone) keyed off `cid` (which must equal the page's `MM_PRODUCT_ID` — log and redirect if not, to prevent cross-PC bundle-URL tampering).
3. `inc_traderPC_configurator.asp` renders the configurator identically to the standalone page, emitting a `<form>` that posts to `/shop/pc/instPrd.asp` with `idproduct`/`quantity`/`OptionGroupCount`/`idOption1..N` **plus** the extra hidden bundle inputs (`idproduct2`=sid, `QtyM<sid>`=1, `idproduct3`=mid, `QtyM<mid>`=monitor count, `pCnt`=3). Those three extra inputs are emitted by `mmEmitBundleHiddenInputs` — the bundle page calls it inside the form; the standalone page doesn't.
4. `instPrd.asp` sees `pCnt=3` and runs its existing cross-sell branch, which marks the PC as the parent and attaches the computed bundle discount (`pcCartArray(…,28)`). No changes to `instPrd.asp` are needed.
5. Client-side JS updates the bundle total by adding the live configurator delta to a `BUNDLE_BASE_EX = (standPrice + monitorPrice * monCount + pcBasePrice - discount) / 1.2` baseline, emitted as a VBScript-rendered JS constant.

### Builder → bundle-page wiring

Whichever builder is live when a given PC's bundle page ships has to point its "Configure the PC" CTA at the new bundle end-page for that PC:

- **New single-page builder** ([redesign/bundles.html](redesign/bundles.html) → its live ASP equivalent): its PC-step Configure CTA should build the URL as `viewPrd-<machine>-bundle-v2.asp?sid=X&mid=Y&cid=Z` for the four migrated machines, falling back to the legacy `viewPrd-<machine>.asp` URL for PCs not yet migrated.
- **Legacy 3-step wizard** ([CUSTOMCAT-bundles3.asp](shop/pc/CUSTOMCAT-bundles3.asp)): while this is still in production, update the PC-card link generation for the four migrated machines the same way. Once the single-page builder replaces it, this edit becomes moot.

The `sid`/`mid`/`cid` querystring format is identical in both builders, so the bundle end-pages themselves don't care which builder sent the user.

### `bundle-breadcrumb.asp` — what happens to it

Keep it as-is for now. While the legacy 3-step wizard is still live, its step-1/2/3 branches are the headers for those pages; while any legacy `viewPrd-*.asp` hasn't migrated, its `bunBCpage = "pc"` branch at [bundle-breadcrumb.asp:534-604](shop/pc/bundle-breadcrumb.asp#L534-L604) is still the bundle ribbon shown above the unmigrated PC pages. Once (a) the single-page builder replaces the 3-step wizard **and** (b) all four PC bundle pages are migrated, the whole file is dead and can be deleted in one go. None of this is scope here — just noting the exit.

### Rollout order

1. Extract the PC-specific marketing blocks from [viewPrd-TraderPC-v2.asp](shop/pc/viewPrd-TraderPC-v2.asp) into the new `inc_traderPC_*.asp` includes; verify the standalone page still renders identically.
2. Add bundle-specific CSS (`.bp-hero`, `.bp-gallery__*`, `.bp-picks`, `.bp-pick-card`, `.bp-compat`, `.bp-sidebar`, `.bp-card`, `.bp-items`, `.bp-sub`, `.bp-total`, `.bp-saved`, `.bp-ribbon`, `.bp-savings`, `.bp-price`, `.bp-incl`, `.bp-cta`, `.bp-foot`) to [css/mm-site.css](css/mm-site.css) under the existing `.mm-site` scope — port verbatim from the inline `<style>` in [redesign/bundleproduct.html](redesign/bundleproduct.html).
3. Build `inc_bundleContext.asp` and `inc_bundleSidebar.asp`.
4. Build `viewPrd-TraderPC-bundle-v2.asp` and point whichever builder is live (legacy `CUSTOMCAT-bundles3.asp` today, or the new single-page builder once that ships) at it for `cid=333`.
5. Test end-to-end on `amz.` staging (see Verification below).
6. Replicate for the other three machines.

## Critical files to reference

- [redesign/bundleproduct.html](redesign/bundleproduct.html) — target mockup
- [shop/pc/viewPrd-TraderPC-v2.asp](shop/pc/viewPrd-TraderPC-v2.asp) — the pattern the bundle page extends
- [shop/pc/bundle-breadcrumb.asp](shop/pc/bundle-breadcrumb.asp) — source of the stand→discount mapping and the querystring-validation idiom (to be ported into `inc_bundleContext.asp`, not reused directly)
- [shop/pc/inc_headerDAJS.asp](shop/pc/inc_headerDAJS.asp) (lines ~79–150) — reference for the multi-product hidden-input shape `instPrd.asp` expects
- [shop/pc/CUSTOMCAT-bundles3.asp](shop/pc/CUSTOMCAT-bundles3.asp) — step-3 PC picker; needs the "routing to bundle page" edit per machine as each ships
- [shop/pc/instPrd.asp](shop/pc/instPrd.asp) (lines ~209–275) — cart submission; **no changes needed** — confirms the bundle contract is already in place
- [shop/pc/viewPrdCode.asp:1226-1292](shop/pc/viewPrdCode.asp#L1226-L1292) — `funBundlesCalcs`, the existing stand+monitor lookup function; can be reused inside `inc_bundleContext.asp` or lifted and tidied — either is fine
- [computer-pages-redesign-plan.md](computer-pages-redesign-plan.md) — the standalone-page plan this extends
- [css/mm-site.css](css/mm-site.css) — target stylesheet for the new `.bp-*` classes

## Verification

Manual browser testing on `amz.multiplemonitors.co.uk`:

- **Full builder → bundle page**: start at `/bundles/`, pick a stand → pick a monitor → pick the Trader PC, confirm the URL you land on is `viewPrd-TraderPC-bundle-v2.asp?sid=…&mid=…&cid=333` and the hero, picks section, and sidebar all reflect the stand + monitor you chose with live DB prices.
- **Direct deep-link**: hit `viewPrd-TraderPC-bundle-v2.asp?sid=312&mid=304&cid=333` directly and confirm it renders correctly.
- **Invalid querystrings**: `sid=abc` (non-numeric), `sid=99999` (unknown product), missing `mid`, `cid` mismatch with page — confirm each redirects to `/404.html` or `/bundles/` (whichever `inc_bundleContext.asp` picks).
- **Configurator → cart**: change every option group, confirm the sidebar `data-pc-line`, `data-pc-pri`, `data-sub`, `data-bun-ex`, `data-bun-inc`, `data-bun-saved` all update live. Click "Add bundle to basket", confirm the cart shows **three line items** (stand, monitor ×N, configured PC) and the bundle discount is applied to the correct line (inspect the cart array values per [atc_viewprd.asp:53-77](shop/pc/atc_viewprd.asp#L53-L77)).
- **Discount per stand type**: test one stand from each discount tier (dual = £25, quad = £50, six-way = £100) and confirm the sidebar and cart totals all match.
- **Standalone page unaffected**: load the Trader PC standalone (`viewPrd-TraderPC-v2.asp`) and confirm it still renders identically after the marketing-section extraction to includes. No bundle sidebar, no bundle picks, no bundle hidden inputs.
- **Admin price edit**: change one option price in admin, reload the bundle page, confirm the new price flows through to the configurator card, the sidebar PC line, and the bundle total (same live-priced guarantee as the standalone page).
- **Mobile viewport**: `.bp-hero-grid`, `.bp-picks__grid`, `.cfg-grid` all stack; the `.bp-sidebar` un-stickies below 992 px (mockup behaviour).
- **Legacy flow regression**: pick a PC whose bundle page hasn't migrated yet (e.g. Ultra PC), confirm you still land on its legacy `viewPrd-*.asp` with the `bundle-breadcrumb.asp` ribbon and the cart still behaves.
