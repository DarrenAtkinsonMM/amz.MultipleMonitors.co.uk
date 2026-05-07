<%
' ============================================================
' viewprd-monitor-v2.asp
' 2026 redesign — Monitor product page.
' Template renders any single monitor SKU from the products
' table.
'
' Resolution order (single SELECT, indexed lookup):
'   1. Request.QueryString("slug") — preserved across the
'      Server.Transfer from viewPrdRouter.asp on friendly-URL
'      hits. Resolves WHERE pcUrl = ?
'   2. Request.QueryString("idProduct") — direct deep-link
'      fallback for any legacy code linking to this page with
'      ?idProduct=N
'   3. Hardcoded test fallback to the 27" Iiyama ProLite
'      XUB2792QSN, used when the page is loaded with no params.
'
' Once mmIdProduct is hydrated, Session("idProductRedirect") is
' set so legacy ProductCart includes (include-metatags.asp,
' inc_footer.asp, apps/pcBackInStock) see the right value.
' Page-local rendering uses the captured mmIdProduct, so a
' concurrent tab overwriting the session var cannot corrupt
' this page.
'
' Phase 1 scope — see /we-need-to-build-misty-hopper.md:
'   * Name, SKU, price, main image and short description come
'     from the products table.
'   * Per-monitor variant copy (h1, gallery thumbs, spec
'     tables) is hardcoded for the 27" Iiyama for now. Other
'     monitors will load dynamic header data but use this copy
'     until follow-up work pulls per-monitor specs from custom
'     fields or a new pcMonitorSpecs table.
' ============================================================
%>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "viewprd-monitor-v2.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<%
Const MM_VAT_RATE = 1.2

' ------------------------------------------------------------
' 1. Product base row — slug-first, single indexed SELECT
' ------------------------------------------------------------
Dim mmIdProduct, mmName, mmTopDesc, mmSku, mmBasePriceInc, mmImageUrl, mmSmallImageUrl
Dim mmSDesc, mmPcUrl, mmAdditionalImages, mmAltTagText, mmStock
mmIdProduct = 0
mmName = ""              : mmSku = ""              : mmBasePriceInc = 0
mmImageUrl = ""          : mmSmallImageUrl = ""    : mmSDesc = ""
mmPcUrl = ""             : mmAdditionalImages = "" : mmAltTagText = ""
mmStock = 0              : mmTopDesc = ""

Dim mmSlug, mmQsIdProduct, mmWhere
mmSlug        = Trim(Request.QueryString("slug") & "")
mmQsIdProduct = Trim(Request.QueryString("idProduct") & "")

If mmSlug <> "" And mmSlugIsSafe(mmSlug) Then
  mmWhere = "pcUrl = '" & Replace(mmSlug, "'", "''") & "'"
ElseIf mmQsIdProduct <> "" And IsNumeric(mmQsIdProduct) Then
  mmWhere = "idProduct = " & CLng(mmQsIdProduct)
Else
  mmWhere = "pcUrl = '27-iiyama-qhd-monitor'"
End If

Dim mmPrdSql, mmPrdRs
mmPrdSql = "SELECT idProduct, description, detailstop, sku, price, imageUrl, smallImageUrl, " & _
           "       sDesc, pcUrl, pcProd_AdditionalImages, " & _
           "       pcProd_AltTagText, stock " & _
           "FROM products " & _
           "WHERE " & mmWhere & _
           "  AND active = -1 AND removed = 0"
Set mmPrdRs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
mmPrdRs.Open mmPrdSql, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If err.number <> 0 Then
  On Error Goto 0
  call LogErrorToDatabase()
  Set mmPrdRs = Nothing
  call closeDB()
  Response.Redirect "techErr.asp?err=" & pcStrCustRefID
End If
On Error Goto 0

If mmPrdRs.EOF Then
  mmPrdRs.Close : Set mmPrdRs = Nothing
  Response.Redirect "/shop/pc/msg.asp?message=88"
End If

mmIdProduct        = CLng(mmPrdRs("idProduct"))
mmName             = mmPrdRs("description") & ""
mmSku              = mmPrdRs("sku") & ""
mmBasePriceInc     = CDbl(mmPrdRs("price"))
mmImageUrl         = mmPrdRs("imageUrl") & ""
mmSmallImageUrl    = mmPrdRs("smallImageUrl") & ""
mmSDesc            = mmPrdRs("sDesc") & ""
mmTopDesc          = mmPrdRs("detailstop") & ""
mmPcUrl            = mmPrdRs("pcUrl") & ""
mmAdditionalImages = mmPrdRs("pcProd_AdditionalImages") & ""
mmAltTagText       = mmPrdRs("pcProd_AltTagText") & ""
If Not IsNull(mmPrdRs("stock")) Then mmStock = CLng(mmPrdRs("stock"))
mmPrdRs.Close : Set mmPrdRs = Nothing

' Expose to legacy ProductCart includes (metatags helper,
' back-in-stock, footer cleanup). Page-local mmIdProduct is
' the source of truth for this page.
Session("idProductRedirect") = mmIdProduct

Dim mmBasePriceEx
mmBasePriceEx = mmBasePriceInc / MM_VAT_RATE

' ------------------------------------------------------------
' 2. Helpers
' ------------------------------------------------------------
Function mmFormatMoney(ByVal v)
  mmFormatMoney = FormatNumber(v, 2, -1, 0, -1)
End Function
Function mmFormatMoney0(ByVal v)
  mmFormatMoney0 = FormatNumber(v, 0, -1, 0, -1)
End Function

Dim mmBasePriceExDisp, mmBasePriceIncDisp
mmBasePriceExDisp  = mmFormatMoney0(mmBasePriceEx)
mmBasePriceIncDisp = mmFormatMoney(mmBasePriceInc)

' Resolve main image with fallbacks (matches viewprd-stand-v2.asp)
Dim mmMainImgSrc
If mmImageUrl <> "" Then
  mmMainImgSrc = "/shop/pc/catalog/" & mmImageUrl
ElseIf mmSmallImageUrl <> "" Then
  mmMainImgSrc = "/shop/pc/catalog/" & mmSmallImageUrl
Else
  mmMainImgSrc = "/shop/pc/catalog/no_image.gif"
End If

' Canonical URL + absolute image URL for schema.org / share targets
Dim mmCanonicalUrl, mmCanonicalImg
If mmPcUrl <> "" Then
  mmCanonicalUrl = "https://www.multiplemonitors.co.uk/products/" & mmPcUrl & "/"
Else
  mmCanonicalUrl = "https://www.multiplemonitors.co.uk/shop/pc/viewprd-monitor-v2.asp?idProduct=" & mmIdProduct
End If
mmCanonicalImg = "https://www.multiplemonitors.co.uk" & mmMainImgSrc

' ------------------------------------------------------------
' 3. Page-level metadata consumed by inc_headerV5.asp
'    (set BEFORE the header_wrapper include so it wins)
' ------------------------------------------------------------
Dim pcv_PageName
pcv_PageName = mmName & " — UK monitor for multi-screen setups | Multiple Monitors"

' Highlight the Monitor Arrays tab in the main nav
topmenuArrays = " class=""is-trader"""
%>
<!--#include file="header_wrapper.asp"-->

<!-- ===================================================================
     Schema.org Product JSON-LD
     =================================================================== -->
<script type="application/ld+json">
{
  "@context": "https://schema.org/",
  "@type": "Product",
  "name": "<%= Replace(mmName, """", "\""") %>",
  "sku": "<%= Replace(mmSku, """", "\""") %>",
  "mpn": "<%= Replace(mmSku, """", "\""") %>",
  "brand": { "@type": "Brand", "name": "Iiyama" },
  "manufacturer": { "@type": "Organization", "name": "Iiyama" },
  "description": "<%= Replace(Replace(mmSDesc, """", "\"""), vbCrLf, " ") %>",
  "image": [ "<%= mmCanonicalImg %>" ],
  "offers": {
    "@type": "Offer",
    "url": "<%= mmCanonicalUrl %>",
    "priceCurrency": "GBP",
    "price": "<%= mmFormatMoney(mmBasePriceEx) %>",
    "priceValidUntil": "2027-12-31",
    "availability": "<% If mmStock > 0 Then %>https://schema.org/InStock<% Else %>https://schema.org/PreOrder<% End If %>",
    "itemCondition": "https://schema.org/NewCondition",
    "shippingDetails": {
      "@type": "OfferShippingDetails",
      "shippingRate": { "@type": "MonetaryAmount", "value": "10.00", "currency": "GBP" },
      "shippingDestination": { "@type": "DefinedRegion", "addressCountry": "GB" }
    }
  }
}
</script>

<style>
  /* ============================================================
     Page-specific: Monitor product page.
     Scoped under .mm-site so it can't leak onto legacy pages.
     Reuses the .pd-* product-page components first introduced
     on traderpc.html / standproduct.html. No page-specific
     additions beyond the breadcrumb and the .pd-* trunk.
     ============================================================ */

  /* -- Breadcrumb ------------------------------------------------- */
  .mm-site .breadcrumb {
    background:#fff; border-bottom:1px solid var(--line);
    font-family:'JetBrains Mono', monospace; font-size:11px;
    letter-spacing:.14em; text-transform:uppercase; color:var(--muted);
  }
  .mm-site .breadcrumb .inner {
    display:flex; align-items:center; gap:10px;
    padding:14px 0; flex-wrap:wrap;
  }
  .mm-site .breadcrumb a { color:var(--muted); }
  .mm-site .breadcrumb a:hover { color:var(--brand); }
  .mm-site .breadcrumb .sep { color:var(--line-strong); }
  .mm-site .breadcrumb .current { color:var(--ink); }

  /* -- Product hero --------------------------------------------- */
  .mm-site .pd-hero {
    position:relative; padding:48px 0 58px; overflow:hidden;
    background:
      radial-gradient(55% 55% at 20% 25%, rgba(15,110,168,.06), transparent 70%),
      radial-gradient(40% 50% at 95% 80%, rgba(242,167,27,.06), transparent 70%),
      linear-gradient(180deg, var(--sand) 0%, var(--sand-2) 100%);
  }
  .mm-site .pd-hero::before {
    content:""; position:absolute; inset:0; pointer-events:none; opacity:.45;
    background-image:
      linear-gradient(rgba(14,27,44,.045) 1px, transparent 1px),
      linear-gradient(90deg, rgba(14,27,44,.045) 1px, transparent 1px);
    background-size:64px 64px;
    mask-image:radial-gradient(ellipse at 50% 40%, #000 30%, transparent 80%);
    -webkit-mask-image:radial-gradient(ellipse at 50% 40%, #000 30%, transparent 80%);
  }
  .mm-site .pd-hero .container { position:relative; z-index:1; }
  .mm-site .pd-hero-grid {
    display:grid; grid-template-columns:1fr; gap:42px; align-items:start;
  }
  @media (min-width:992px) {
    .mm-site .pd-hero-grid { grid-template-columns:1.1fr 1fr; gap:64px; }
  }

  /* Gallery (left column) */
  .mm-site .pd-gallery { position:relative; }
  .mm-site .pd-gallery__main {
    position:relative;
    background:
      radial-gradient(55% 55% at 30% 30%, rgba(15,110,168,.06), transparent 70%),
      linear-gradient(160deg, #F4F8FB 0%, #E8F1F8 100%);
    border:1px solid var(--line); border-radius:var(--radius-xl);
    aspect-ratio:4/3; display:flex; align-items:center; justify-content:center;
    padding:40px; overflow:hidden;
    box-shadow:var(--shadow-lg), inset 0 0 0 1px rgba(255,255,255,.6);
  }
  .mm-site .pd-gallery__main::before {
    content:""; position:absolute; inset:0; pointer-events:none; opacity:.35;
    background-image:
      linear-gradient(rgba(14,27,44,.06) 1px, transparent 1px),
      linear-gradient(90deg, rgba(14,27,44,.06) 1px, transparent 1px);
    background-size:28px 28px;
    mask-image:radial-gradient(ellipse at 50% 50%, #000 10%, transparent 70%);
    -webkit-mask-image:radial-gradient(ellipse at 50% 50%, #000 10%, transparent 70%);
  }
  .mm-site .pd-gallery__main img {
    position:relative; max-height:100%; width:auto; max-width:82%;
    filter:drop-shadow(0 20px 28px rgba(14,27,44,.18));
    transition:transform .35s ease;
  }
  .mm-site .pd-gallery__main:hover img { transform:scale(1.03) translateY(-4px); }

  /* Floating chip — "In stock · 3-year warranty" */
  .mm-site .pd-gallery__chip {
    position:absolute; top:18px; left:18px;
    background:rgba(14,27,44,.92); color:#fff;
    border:1px solid rgba(255,255,255,.08); border-radius:999px;
    padding:7px 12px 7px 10px;
    font-family:'JetBrains Mono', monospace; font-size:10.5px;
    letter-spacing:.14em; text-transform:uppercase;
    display:inline-flex; align-items:center; gap:8px;
    backdrop-filter:blur(6px); z-index:2;
  }
  .mm-site .pd-gallery__chip .dot {
    width:6px; height:6px; border-radius:50%; background:var(--up);
    box-shadow:0 0 0 0 rgba(33,166,122,.6);
    animation:pulse 2.2s ease-out infinite;
  }
  .mm-site .pd-gallery__chip .acc { color:var(--accent); margin-right:3px; }

  /* SKU mark — bottom-right of gallery */
  .mm-site .pd-gallery__sku {
    position:absolute; right:18px; bottom:18px;
    font-family:'JetBrains Mono', monospace; font-size:10px;
    letter-spacing:.18em; text-transform:uppercase;
    color:var(--muted); background:rgba(255,255,255,.85);
    border:1px solid var(--line); border-radius:4px;
    padding:4px 8px; backdrop-filter:blur(4px); z-index:2;
  }

  /* Thumbnail strip */
  .mm-site .pd-gallery__thumbs {
    display:grid; grid-template-columns:repeat(4, 1fr); gap:10px; margin-top:14px;
  }
  .mm-site .pd-thumb {
    aspect-ratio:1/1;
    background:linear-gradient(160deg, #F4F8FB, #E8F1F8);
    border:1px solid var(--line); border-radius:var(--radius);
    display:flex; align-items:center; justify-content:center; padding:8px;
    cursor:pointer; transition:all .2s ease; overflow:hidden;
    position:relative;
  }
  .mm-site .pd-thumb img { max-height:100%; max-width:100%; object-fit:contain; }
  .mm-site .pd-thumb:hover { border-color:var(--brand); transform:translateY(-2px); }
  .mm-site .pd-thumb.is-active {
    border-color:var(--brand); box-shadow:0 0 0 2px rgba(15,110,168,.18);
  }

  /* Buy-box (right column) */
  .mm-site .pd-buybox { display:flex; flex-direction:column; gap:18px; }
  .mm-site .pd-buybox .eyebrow { margin-bottom:0; }
  .mm-site .pd-buybox h1 {
    font-family:'EB Garamond', 'Georgia', serif;
    font-size:clamp(38px, 5vw, 52px); line-height:1.03;
    letter-spacing:-.028em; font-weight:500; color:var(--ink);
    margin:0;
  }
  .mm-site .pd-buybox h1 em {
    font-style:italic; color:var(--brand); font-weight:400;
  }
  .mm-site .pd-buybox .pitch {
    font-size:16.5px; color:var(--slate); line-height:1.55; margin:0;
    max-width:520px;
  }

  /* Price block */
  .mm-site .pd-price {
    display:flex; align-items:baseline; flex-wrap:wrap; gap:12px 18px;
    padding:20px 0 18px;
    border-top:1px solid var(--line); border-bottom:1px solid var(--line);
  }
  .mm-site .pd-price__from {
    font-family:'JetBrains Mono', monospace; font-size:11px;
    letter-spacing:.18em; text-transform:uppercase; color:var(--muted);
  }
  .mm-site .pd-price__num {
    font-family:'EB Garamond', 'Georgia', serif; font-weight:500;
    font-size:56px; line-height:1; letter-spacing:-.028em; color:var(--ink);
    display:inline-flex; align-items:baseline; gap:4px;
  }
  .mm-site .pd-price__num .sym {
    font-size:28px; color:var(--brand); margin-right:2px;
  }
  .mm-site .pd-price__vat {
    font-family:'JetBrains Mono', monospace; font-size:12px;
    color:var(--muted); letter-spacing:.08em;
  }
  .mm-site .pd-price__vat b { color:var(--ink); font-weight:500; }
  .mm-site .pd-price__vat .ship {
    display:block; margin-top:4px;
    font-family:'Geist', sans-serif; font-size:13px; letter-spacing:0;
    color:var(--slate); text-transform:none;
  }

  /* Delivery cutoff */
  .mm-site .pd-cutoff {
    display:inline-flex; align-items:center; gap:10px;
    background:rgba(33,166,122,.07); border:1px solid rgba(33,166,122,.22);
    border-radius:999px; padding:9px 14px 9px 12px;
    font-size:13.5px; color:var(--ink); font-weight:500;
  }
  .mm-site .pd-cutoff .dot {
    width:7px; height:7px; border-radius:50%; background:var(--up);
    box-shadow:0 0 0 0 rgba(33,166,122,.55);
    animation:pulse 2.2s ease-out infinite;
  }
  .mm-site .pd-cutoff b {
    font-family:'JetBrains Mono', monospace; font-size:12px; font-weight:500;
    color:var(--up); letter-spacing:.05em;
  }

  /* Inclusion strip */
  .mm-site .pd-incl {
    display:grid; grid-template-columns:repeat(3, 1fr); gap:16px;
    margin:4px 0 6px;
  }
  @media (max-width:520px) { .mm-site .pd-incl { grid-template-columns:1fr; gap:10px; } }
  .mm-site .pd-incl .item {
    display:flex; align-items:center; gap:10px; font-size:13px; color:var(--slate);
    line-height:1.3;
  }
  .mm-site .pd-incl .item .fa {
    color:var(--brand); font-size:16px; width:18px; text-align:center; flex-shrink:0;
  }
  .mm-site .pd-incl .item b { color:var(--ink); font-weight:500; display:block; }
  .mm-site .pd-incl .item small {
    font-family:'JetBrains Mono', monospace; font-size:10px;
    letter-spacing:.1em; color:var(--muted); text-transform:uppercase;
  }

  /* Quantity stepper + CTA row */
  .mm-site .pd-action {
    display:grid; grid-template-columns:auto 1fr; gap:12px; align-items:stretch;
  }
  @media (max-width:480px) {
    .mm-site .pd-action { grid-template-columns:1fr; }
  }
  .mm-site .pd-qty {
    display:inline-flex; align-items:stretch;
    border:1px solid var(--line-strong); border-radius:var(--radius);
    background:#fff; overflow:hidden;
  }
  .mm-site .pd-qty button {
    width:40px; border:none; background:#fff; color:var(--ink);
    font-size:16px; cursor:pointer; transition:background .15s ease;
  }
  .mm-site .pd-qty button:hover { background:var(--sand); color:var(--brand); }
  .mm-site .pd-qty input {
    width:44px; text-align:center; border:none;
    border-left:1px solid var(--line); border-right:1px solid var(--line);
    font-family:'EB Garamond', 'Georgia', serif; font-weight:500;
    font-size:18px; color:var(--ink); background:#fff;
    -moz-appearance:textfield; appearance:textfield;
  }
  .mm-site .pd-qty input::-webkit-outer-spin-button,
  .mm-site .pd-qty input::-webkit-inner-spin-button { -webkit-appearance:none; margin:0; }
  .mm-site .pd-action .btn-primary {
    flex:1 1 auto; min-width:0; justify-content:center; font-size:15px;
  }
  .mm-site .pd-action .btn-primary .fa { transition:transform .2s ease; }
  .mm-site .pd-action .btn-primary:hover .fa { transform:translateX(4px); }

  /* Mini footnote beneath CTAs */
  .mm-site .pd-foot {
    font-family:'JetBrains Mono', monospace; font-size:10.5px;
    letter-spacing:.12em; text-transform:uppercase;
    color:var(--muted); display:flex; gap:14px; flex-wrap:wrap;
  }
  .mm-site .pd-foot span { display:inline-flex; align-items:center; gap:6px; }
  .mm-site .pd-foot .fa { color:var(--brand); font-size:11px; }
  .mm-site .pd-foot a { color:var(--brand); }
</style>

<div class="mm-site">

<!-- ===================================================================
     BREADCRUMB
     =================================================================== -->
<nav class="breadcrumb" aria-label="Breadcrumb">
  <div class="container inner">
    <a href="/">Home</a>
    <span class="sep">/</span>
    <a href="/display-systems/">Monitors</a>
    <span class="sep">/</span>
    <span class="current"><%= mmName %></span>
  </div>
</nav>

<!-- ===================================================================
     PRODUCT HERO — gallery + buy-box
     =================================================================== -->
<section class="pd-hero">
  <div class="container">
    <div class="pd-hero-grid">

      <!-- Gallery column -->
      <div class="pd-gallery reveal" id="pdGallery">
        <div class="pd-gallery__main">
          <span class="pd-gallery__chip">
            <span class="dot"></span><span class="acc">IN</span>STOCK&nbsp;&middot;&nbsp;3-YEAR&nbsp;WARRANTY
          </span>
          <img id="pdMainImg"
               src="<%= mmMainImgSrc %>"
               alt="<%= Server.HTMLEncode(mmName) %>" />
          <% If mmSku <> "" Then %>
          <span class="pd-gallery__sku">SKU &middot; <%= Server.HTMLEncode(mmSku) %></span>
          <% End If %>
        </div>

        <%
        ' TODO (phase-2): parse mmAdditionalImages (comma-separated)
        ' to drive these thumbs dynamically. Hardcoded for the
        ' 27" Iiyama mockup for now — all four thumbs reuse the
        ' main image since per-monitor angles aren't wired yet.
        %>
        <div class="pd-gallery__thumbs">
          <div class="pd-thumb is-active"
               data-img="<%= mmMainImgSrc %>"
               title="Front view">
            <img src="<%= mmMainImgSrc %>" alt="Front view of <%= Server.HTMLEncode(mmName) %>" />
          </div>
          <div class="pd-thumb"
               data-img="<%= mmMainImgSrc %>"
               title="Three-quarter view">
            <img src="<%= mmMainImgSrc %>" alt="Three-quarter view showing 2 mm frameless bezel" />
          </div>
          <div class="pd-thumb"
               data-img="<%= mmMainImgSrc %>"
               title="Rear inputs">
            <img src="<%= mmMainImgSrc %>" alt="Rear view showing HDMI, DisplayPort and USB-C inputs" />
          </div>
          <div class="pd-thumb"
               data-img="<%= mmMainImgSrc %>"
               title="Height-adjust stand">
            <img src="<%= mmMainImgSrc %>" alt="Height-adjustable stand with pivot and swivel" />
          </div>
        </div>
      </div>

      <!-- Buy-box column -->
      <aside class="pd-buybox reveal" style="transition-delay:.08s">
        <div class="eyebrow">Monitor &middot; 27&Prime; Quad HD IPS</div>
        <h1><%= mmName %></h1>
        <p class="pitch">
          <% If mmTopDesc <> "" Then %>
            <%= mmTopDesc %>
          <% Else %>
            The 27-inch IPS panel we fit to our Quad HD bundles and multi-monitor arrays &mdash;
            2560&thinsp;&times;&thinsp;1440 at 100&nbsp;Hz, 4&nbsp;ms, USB-C 65&nbsp;W single-cable dock, and a
            2&nbsp;mm three-side frameless bezel built for tight multi-screen layouts.
          <% End If %>
        </p>

        <div class="pd-price">
          <div>
            <div class="pd-price__from">Price</div>
            <div class="pd-price__num"><span class="sym">&pound;</span><%= mmBasePriceExDisp %></div>
          </div>
          <div class="pd-price__vat">
            <b>&pound;<%= mmBasePriceIncDisp %></b> inc VAT
            <span class="ship">UK delivery &pound;10 flat &middot; free when paired with a stand, bundle or array</span>
          </div>
        </div>

        <div>
          <span class="pd-cutoff">
            <span class="dot"></span>
            Order before <b><%= daFunDelCutOff() %></b> &middot; delivered <%= daFunDelDateReturn(0,0) %>
          </span>
        </div>

        <div class="pd-incl">
          <div class="item">
            <i class="fa fa-cube"></i>
            <div><b>UK stock</b><small>Ready to ship</small></div>
          </div>
          <div class="item">
            <i class="fa fa-shield"></i>
            <div><b>3-year warranty</b><small>Iiyama cover</small></div>
          </div>
          <div class="item">
            <i class="fa fa-truck"></i>
            <div><b>2-day dispatch</b><small>UK courier</small></div>
          </div>
        </div>

        <form method="post" action="/shop/pc/instPrd.asp" class="pd-action" id="addToCartForm">
          <input type="hidden" name="idproduct" value="<%= mmIdProduct %>">
          <input type="hidden" name="OptionGroupCount" value="0">
          <div class="pd-qty" role="group" aria-label="Quantity">
            <button type="button" aria-label="Decrease quantity"
                    onclick="(function(i){i.value=Math.max(1,(parseInt(i.value,10)||1)-1);})(document.getElementById('pdQty'));">&minus;</button>
            <input id="pdQty" name="quantity" type="number" min="1" value="1" aria-label="Quantity" />
            <button type="button" aria-label="Increase quantity"
                    onclick="(function(i){i.value=(parseInt(i.value,10)||1)+1;})(document.getElementById('pdQty'));">+</button>
          </div>
          <button type="submit" class="btn btn-primary btn-lg">
            Add Monitor To Your Basket <i class="fa fa-arrow-right"></i>
          </button>
        </form>

        <div class="pd-foot">
          <span><i class="fa fa-phone"></i>Call Darren &mdash; 0330 223 66 55</span>
          <span><i class="fa fa-plug"></i>3&nbsp;m video cable included</span>
        </div>
      </aside>

    </div>
  </div>
</section>

<!-- ===================================================================
     FULL SPECIFICATION &mdash; two-card .ya-specs grid
     =================================================================== -->
<section class="s">
  <div class="container">
    <div class="section-head-narrow reveal">
      <div class="eyebrow" style="justify-content:center; display:inline-flex;">Spec sheet</div>
      <h2 style="margin-top:14px;">Full specification.</h2>
      <p>Every number that matters, in one place &mdash; straight from the Iiyama datasheet and our own workshop test fits.</p>
    </div>

    <div class="ya-specs reveal">

      <div class="ya-specs__col">
        <div class="ya-specs__lbl">Display</div>
        <h4>27&Prime; Iiyama ProLite XUB2792QSN</h4>
        <table class="ya-specs__table">
          <tbody>
            <tr><th>Manufacturer</th>   <td>Iiyama</td></tr>
            <tr><th>Model</th>          <td>ProLite XUB2792QSN</td></tr>
            <tr><th>Size</th>           <td>27&Prime; (27&Prime; diagonal)</td></tr>
            <tr><th>Resolution</th>     <td>2560 &times; 1440 (Quad HD)</td></tr>
            <tr><th>Refresh rate</th>   <td>100 Hz</td></tr>
            <tr><th>Response time</th>  <td>4 ms (GtG)</td></tr>
            <tr><th>Panel type</th>     <td>IPS</td></tr>
          </tbody>
        </table>
      </div>

      <div class="ya-specs__col">
        <div class="ya-specs__lbl">Connectivity &amp; build</div>
        <h4>Ports, mounts &amp; what&rsquo;s in the box</h4>
        <table class="ya-specs__table">
          <tbody>
            <tr><th>Inputs</th>         <td>HDMI 1.4 &middot; DisplayPort 1.4 &middot; USB-C (65&nbsp;W) &middot; USB-B hub</td></tr>
            <tr><th>Bezel width</th>    <td>2 mm (three-side frameless)</td></tr>
            <tr><th>VESA mount</th>     <td>100 &times; 100</td></tr>
            <tr><th>Warranty</th>       <td>3-year manufacturer (Iiyama)</td></tr>
            <tr><th>Cables included</th><td>3&nbsp;m DisplayPort <em>or</em> HDMI &middot; UK power</td></tr>
          </tbody>
        </table>
      </div>

    </div>
  </div>
</section>

</div><!-- /.mm-site -->

<!-- ===================================================================
     PAGE-SPECIFIC JS
     Reveal-on-scroll is already wired in footer_wrapper.asp.
     =================================================================== -->
<script>
(function(){
  // -- Gallery thumbs swap main image --
  var thumbs = document.querySelectorAll('.pd-thumb[data-img]');
  var main   = document.getElementById('pdMainImg');
  if (!main) return;
  thumbs.forEach(function(t){
    t.addEventListener('click', function(){
      document.querySelectorAll('.pd-thumb.is-active').forEach(function(x){ x.classList.remove('is-active'); });
      t.classList.add('is-active');
      main.src = t.dataset.img;
      var imgEl = t.querySelector('img');
      if (imgEl && imgEl.alt) main.alt = imgEl.alt;
    });
  });
})();
</script>

<!--#include file="footer_wrapper.asp"-->
