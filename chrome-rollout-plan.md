# Rolling out the 2026 chrome (header + footer) sitewide

## Context

We want to begin rolling the `redesign/newhome.html` look onto the live storefront. The fastest, lowest-risk starting point is the **chrome** (topbar, nav, footer) because:

1. It appears on every page, so a single change produces a sitewide visual refresh.
2. It's emitted from two include files (`header_wrapper.asp`, `footer_wrapper.asp`) that are referenced by **204 ASP pages** — one change, one test pass.
3. `css/mm-site.css` is **already loaded last** in `shop/pc/inc_headerCSS.asp:62`, and its chrome rules (`.site-header`, `.topbar`, `.navwrap`, `footer`, etc.) are intentionally **unscoped** so they'll apply to the new markup the moment we ship it.
4. The content-area rules in `mm-site.css` are scoped under `.mm-site` — legacy pages that don't opt in continue to render with Bootstrap 3 + `style.css` as they do today. This matches what the redesign was built for.

The goal of this first rollout is purely: **swap the chrome HTML to the new markup in the two wrapper files, preserving every piece of dynamic behaviour (cart, login, active-menu highlight, analytics, Schema.org/OpenGraph, mobile viewport, favicon) that the current wrappers emit.**

## How the current wrappers fit together

Every storefront page (e.g. [default.asp](default.asp)) includes the wrappers:

```
default.asp
 ├─ shop/pc/header_wrapper.asp             ← opens <html> and <head>
 │   ├─ shop/pc/inc_headerV5.asp           ← meta, title, CSS loader call
 │   │   └─ shop/pc/inc_headerCSS.asp      ← all stylesheets, mm-site.css last
 │   ├─ shop/pc/inc_headerDAJS.asp         ← emits </head> and <body>
 │   └─ [legacy nav markup: lines 91-142]  ← THE CHROME WE'RE REPLACING
 │       └─ shop/pc/smallQuickCart.asp     ← AngularJS cart + login widget
 │
 ├─ <page content>
 │
 └─ shop/pc/footer_wrapper.asp
     ├─ [legacy footer markup: lines 23-120] ← THE CHROME WE'RE REPLACING
     ├─ shop/pc/inc_footer.asp              ← GA, CartStack, Pinterest
     ├─ [Drip + cookie.js + misc JS]
     └─ </body></html>
```

Key emission points to be aware of when editing:
- `<html>` opens in [header_wrapper.asp:2](shop/pc/header_wrapper.asp#L2) (includes `prefix="og:..."` for product pages).
- `<head>` opens in [header_wrapper.asp:3](shop/pc/header_wrapper.asp#L3) (with `itemscope itemtype="http://schema.org/WebSite"`).
- `</head>` and `<body>` are emitted from `inc_headerDAJS.asp` — **one of four `<body>` variants** depending on `pcv_strViewPrdStyle` (product pages get `onLoad="pageLD()" id="page-top" data-spy="scroll" data-target=".navbar-custom"`). We leave that file alone.
- The nav chrome to replace is [header_wrapper.asp:91-142](shop/pc/header_wrapper.asp#L91-L142).
- Footer chrome to replace is [footer_wrapper.asp:23-120](shop/pc/footer_wrapper.asp#L23-L120).
- `<div id="wrapper">` opens at [header_wrapper.asp:91](shop/pc/header_wrapper.asp#L91) and closes at [footer_wrapper.asp:122](shop/pc/footer_wrapper.asp#L122). Keep it — it doesn't collide with the new chrome and some legacy JS may rely on it.

## Rollout approach — single cutover, two files

The chrome is two coupled pieces of HTML. Ship both in the same deploy so we never see half-new chrome.

### Step 1 — Refactor `smallQuickCart.asp` to emit the new `.cart-btn`

[smallQuickCart.asp](shop/pc/smallQuickCart.asp) currently emits login text + basket count using AngularJS inside `#quickCartContainer` (the `QuickCartCtrl` controller binds `shoppingcart.totalQuantity`, `shoppingcart.daQuickCart`, `shoppingcart.total`, `shoppingcart.checkoutStage`). The new chrome splits this into two locations:

- **Topbar** — "Existing Customer Login" link (xs-hidden) and "Basket (n)" link (xs-only fallback).
- **Nav actions** — pill-shaped `.cart-btn` with basket icon and item count.

Rewrite `smallQuickCart.asp` so it emits **both** locations, still wrapped in `#quickCartContainer` with the same AngularJS controller. Preserve:
- Link to `/shop/pc/custPref.asp` (login).
- Link to `/shop/pc/viewCart.asp` (basket).
- AngularJS bindings: `{{shoppingcart.totalQuantity}}` and the checkout-stage totals.
- The existing `ng-show`/`ng-hide` logic so the widget only shows totals once `shoppingcart.totalQuantity > 0`.
- The guard `If Instr(Ucase(...SCRIPT_NAME), "GW") = 0 Then` that hides the widget on payment-gateway callback pages.

The topbar's structure in [newhome.html:27-39](redesign/newhome.html#L27-L39) needs this include split into two output points, so the cleanest refactor is:
1. `smallQuickCart.asp` keeps producing the login + basket links as today but with the new markup for the topbar position.
2. A new small include (`smallCartButton.asp`) emits the pill `.cart-btn` for the nav-actions position.
3. Both must live inside a single `#quickCartContainer` ancestor so the Angular controller covers them, OR each gets its own `data-ng-controller="QuickCartCtrl"` scope (safe — Angular allows multiple instantiations).

### Step 2 — Refactor `header_wrapper.asp` (lines 91-142 only)

Replace the `<nav class="navbar navbar-custom navbar-fixed-top">` block with the chrome markup from [newhome.html:26-77](redesign/newhome.html#L26-L77). Integration points:

- **Brand logo** — the new markup references `/images/mm-logo-trans.png`. Confirm that asset exists on the live site (it's referenced from mockups via the full `https://www.multiplemonitors.co.uk/` URL so it must already exist). If so, use the relative path `/images/mm-logo-trans.png`. If not, either upload it or keep `/images/logo.png` for this first deploy.
- **Active-menu highlight** — the existing `Select Case` block at [header_wrapper.asp:50-87](shop/pc/header_wrapper.asp#L50-L87) populates variables like `topmenuHome`, `topmenuComputers`, `topmenuArrays`, `topmenuBundles`, `topmenuStands`, `topmenuBlog`. Rename those to emit the new CSS class — `.mainnav a.is-trader` is the "active" treatment in the new design. So each `<a>` in `.mainnav` and `.mobnav` renders with `class="<%=topmenuX%>"` where `topmenuX` is set to `"is-trader"` (string without quotes) when that section matches. Blog isn't in the new nav — either add it or drop the Blog case.
- **Phone/email** — currently hardcoded in the topbar-connects div. Keep the same hardcoded values in the new `.topbar` block.
- **Cart widget** — where the mockup has `<a href="/shop/pc/viewcart.asp" class="cart-btn">`, include the new `smallCartButton.asp` instead (from Step 1).
- **Login link in topbar** — the mockup has `<a href="#">Existing Customer Login</a>`. Point this at `/shop/pc/custPref.asp` to match what smallQuickCart does today.
- `<div id="wrapper">` still opens just before the nav — **preserve it** to keep the footer's closing `</div>` matched.

Leave untouched: the prefix/OG stuff on `<html>`, the `<head>` opener, the Schema.org itemscope, all meta / viewport / favicon logic, the GTM block (lines 32-42), and the `<div id="pcMainService">` AngularJS bootstrap.

### Step 3 — Refactor `footer_wrapper.asp` (lines 22-120 only)

Replace the legacy `<footer>…</footer>` block with the chrome markup from [newhome.html:451-509](redesign/newhome.html#L451-L509). Integration points:

- Keep `<% Year(Now()) %>` for the copyright year (the mockup uses a JS year-updater; the server-side version is equivalent and simpler — prefer it, and drop the `<span id="year">` + JS block).
- Drop the legacy "Recently Viewed" column (it pulled from `smallRecentProducts.asp`) — the new footer design doesn't have that column. If you want to keep the recently-viewed feature anywhere, that's a separate decision; for this rollout, dropping it is intentional.
- Drop the "FREE BUYERS GUIDE" column / Drip-form button from the old footer. The new design doesn't have a footer newsletter either (only the inline `.guide-card` on the homepage). Drip's **script** tag stays in the global JS block at the bottom of the file — only the visible CTA goes.
- Keep:
  - `</div>` that closes `#wrapper` — immediately after the new `</footer>`.
  - All JS blocks after `</div>` (wow.js, custom.js, ekko-lightbox.min.js, jquery.scrollTo.js, the Drip snippet, cookie.js, the `</body></html>`).
  - The OPC-variant branching at [footer_wrapper.asp:131-150](shop/pc/footer_wrapper.asp#L131-L150) that swaps `custom.js` for `custom-opc.js` on checkout.
  - The `<a href="#" class="scrollup">` back-to-top anchor (or reassess: the new design doesn't show one — decision for step 1 is "keep it working" so leave it in).
  - The `call closeDB()` at line 125.
  - The fallback branching at lines 1-19 (Facebook/mobile footers) — it's inert (both `server.Execute` calls are commented) but harmless; leave as-is.

### Step 4 — Inline the new chrome's mobile-menu JS

The new nav mobile toggle uses:
```
onclick="document.getElementById('mobnav').classList.toggle('is-open')"
```
That's self-contained inline JS — no new global script needed. Bootstrap 3's `data-toggle="collapse"` from the legacy nav is no longer referenced, but its `bootstrap.min.js` is still loaded for other components (modals, carousels, the review carousel at [default.asp:207](default.asp#L207)). Leave Bootstrap JS loaded.

Same for the `.reveal` scroll-in animation: that script only runs on pages that wrap content in `.mm-site` (because only those pages will have `.reveal` elements). Adding it as a tiny inline `<script>` at the end of `footer_wrapper.asp` is cheap and no-op on legacy pages. **Decision:** defer this to the per-page migration phase, not this chrome rollout. Reveal is content, not chrome.

## Files to modify

| File | Change |
|---|---|
| [shop/pc/header_wrapper.asp](shop/pc/header_wrapper.asp) | Replace lines 91-142 with new topbar + `.site-header`/`.navwrap`/`.mobnav`. Update the `Select Case` block at 50-87 to emit `is-trader` class names. |
| [shop/pc/footer_wrapper.asp](shop/pc/footer_wrapper.asp) | Replace lines 22-120 with new `<footer>` markup. Leave everything after line 120 alone. |
| [shop/pc/smallQuickCart.asp](shop/pc/smallQuickCart.asp) | Refactor to emit topbar markup (login link + xs basket fallback). |
| [shop/pc/smallCartButton.asp](shop/pc/smallCartButton.asp) | **New file** — pill `.cart-btn` with AngularJS count binding, included from the new nav. |

Files **NOT** modified in this rollout: `inc_headerV5.asp`, `inc_headerDAJS.asp`, `inc_headerCSS.asp` (already wired), `inc_footer.asp`, `common.asp`, and the 204 pages that include the wrappers.

## Behaviour that must continue working

Verify the following post-deploy — one check per item:

- [ ] **Active-page nav highlight** — open `/`, `/computers/`, `/display-systems/`, `/bundles/`, `/stands/`, and a product page; the right nav item gets the `is-trader` underline.
- [ ] **Basket count in nav `.cart-btn`** — add a product to basket, confirm the `(0)` pill updates to `(1)` without a hard refresh (proves AngularJS binding survived).
- [ ] **Login link in topbar** — clicking routes to `/shop/pc/custPref.asp`.
- [ ] **Mobile menu toggle** — at <992px width, the hamburger opens the `.mobnav` panel.
- [ ] **Logo href** — `/` link.
- [ ] **Phone/email links** — `tel:` and `mailto:` still work from topbar.
- [ ] **Favicon + page title** — unchanged (handled outside the chrome).
- [ ] **Schema.org / OpenGraph meta** — view-source on a product page still shows the schema itemscope and OG prefix.
- [ ] **Google Tag Manager** — GTM still fires (check Network → `gtm.js` requested).
- [ ] **Analytics / Drip / CartStack / Pinterest** — footer_wrapper's post-`</footer>` includes untouched; tags still fire.
- [ ] **Checkout page** — `shop/pc/checkout.asp` loads `custom-opc.js` (not `custom.js`); the OPC branching still triggers.
- [ ] **Payment gateway callback pages (`gw*.asp`)** — the `smallQuickCart.asp` guard `If Instr(Ucase, "GW") = 0` still hides the cart widget there.
- [ ] **`<div id="wrapper">` open/close still balanced** — view-source confirms one open, one close.
- [ ] **Pages that bypass the wrappers** (~30 popup/admin/utility pages, plus `/guide/*.asp`, `/landing/trading.asp`) — unchanged, rendering as before. Verify one guide page loads visually intact.
- [ ] **Legacy content CSS still applies** — homepage hero (`#intro.intro`), welcome section, services, testimonials — all still styled by `style.css`/`responsive.css`/`blue.css` exactly as before.

## Verification / testing

1. Deploy the four file changes to the **amz.multiplemonitors.co.uk** staging subdomain (the current working copy).
2. Walk the checklist above.
3. Spot-check in Chrome DevTools:
   - Mobile viewport (≤640px), tablet (641-991px), desktop (≥992px), large-desktop (≥1165px where `.brand-est` appears).
   - Network tab — confirm no 404s on new asset paths (logo, mm-logo-trans-w.png in footer).
4. Once staging is green, the cutover to the live site is a file copy of those four files.

## Explicitly out of scope for this first rollout

- Migrating any page's **content area** into `.mm-site` — that's the per-page migration phase that comes next (homepage → `default.asp` using `redesign/newhome.html`, trading → `redesign/trading.html`, etc.).
- Removing legacy CSS (`style.css`, `responsive.css`, `blue.css`) — they retire naturally as pages migrate.
- Admin panel (`shop/130707/*`) — not touched by this rollout. Admin has its own layout.
- `/guide/*.asp`, `/landing/trading.asp`, and other pages that bypass the wrappers — they keep their current look until intentionally migrated.
- Dynamic menu generation from DB categories — the new nav is hardcoded in markup, matching the mockup. If DB-driven nav is desired later, that's a separate project.
- Search UI — neither the current header nor the new chrome has visible search. No change.
- Recently-viewed footer column and the Drip "Get It Now" CTA button — removed as part of the footer redesign. The Drip tracking **script** stays loaded site-wide.
