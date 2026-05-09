<%
' ============================================================
' viewprd-stand-v2.asp
' 2026 redesign - Synergy Stand product page.
' Template renders any single stand SKU from the products table.
'
' Resolution order (single SELECT, indexed lookup):
'   1. Request.QueryString("slug") - preserved across the
'      Server.Transfer from viewPrdRouter.asp on friendly-URL
'      hits. Resolves WHERE pcUrl = ?
'   2. Request.QueryString("idProduct") - direct deep-link
'      fallback for any legacy code linking to this page with
'      ?idProduct=N
'   3. Hardcoded test fallback to the Quad Square Synergy Stand
'      (slug "quad-monitor-stand"), used when the page is loaded
'      with no params for direct-load testing.
'
' Once mmIdProduct is hydrated, Session("idProductRedirect") is
' set so legacy ProductCart includes (include-metatags.asp,
' inc_footer.asp, apps/pcBackInStock) see the right value.
' Page-local rendering uses the captured mmIdProduct, so a
' concurrent tab overwriting the session var cannot corrupt
' this page.
'
' Per-stand copy split:
'   * Name, SKU, price, main image and short description come
'     from the products table.
'   * Eyebrow line, inclusion-strip copy, the six "specs at a
'     glance" cards, in-the-box list, micro-band trio, VESA
'     intro, six dim-stats, and dimension-diagram image paths
'     come from shop/includes/standSpecs.asp, keyed by SKU.
'     Page falls back to hardcoded Quad Square defaults when a
'     SKU isn't registered there.
'   * Gallery thumbs come from pcProductsImages (row 0 = main
'     product image, rows 1+ ordered by pcProdImage_Order).
'     Click swaps hero to pcProdImage_LargeUrl.
' ============================================================
%>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "viewprd-stand-v2.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<!--#include file="../includes/standSpecs.asp"-->
<%
Const MM_VAT_RATE = 1.2

' ------------------------------------------------------------
' 1. Product base row - slug-first, single indexed SELECT
' ------------------------------------------------------------
Dim mmIdProduct, mmName, mmTopDesc, mmSku, mmBasePriceInc, mmImageUrl, mmSmallImageUrl, mmLargeImageUrl
Dim mmSDesc, mmPcUrl, mmAdditionalImages, mmAltTagText, mmStock
mmIdProduct = 0
mmName = ""              : mmSku = ""              : mmBasePriceInc = 0
mmImageUrl = ""          : mmSmallImageUrl = ""    : mmLargeImageUrl = ""
mmSDesc = ""             : mmPcUrl = ""            : mmAdditionalImages = ""
mmAltTagText = ""        : mmStock = 0             : mmTopDesc = ""

Dim mmSlug, mmQsIdProduct, mmWhere
mmSlug        = Trim(Request.QueryString("slug") & "")
mmQsIdProduct = Trim(Request.QueryString("idProduct") & "")

If mmSlug <> "" And mmSlugIsSafe(mmSlug) Then
  mmWhere = "pcUrl = '" & Replace(mmSlug, "'", "''") & "'"
ElseIf mmQsIdProduct <> "" And IsNumeric(mmQsIdProduct) Then
  mmWhere = "idProduct = " & CLng(mmQsIdProduct)
Else
  mmWhere = "pcUrl = 'quad-monitor-stand'"
End If

Dim mmPrdSql, mmPrdRs
mmPrdSql = "SELECT idProduct, description, detailstop, sku, price, imageUrl, smallImageUrl, largeImageURL, " & _
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
mmLargeImageUrl    = mmPrdRs("largeImageURL") & ""
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

' Per-stand reference content - keyed by SKU. Returns Nothing
' if SKU isn't registered in standSpecs.asp; mmMetaStr falls
' back to hardcoded Quad Square defaults in that case.
Dim mmMeta
Set mmMeta = mmGetStandMeta(mmSku)

' Machine name exposed to the Darren CTA include
Dim mmMachineName : mmMachineName = mmName

' ------------------------------------------------------------
' 2. Helpers (same shape as viewprd-monitor-v2.asp)
' ------------------------------------------------------------
Function mmFormatMoney(ByVal v)
  mmFormatMoney = FormatNumber(v, 2, -1, 0, -1)
End Function
Function mmFormatMoney0(ByVal v)
  mmFormatMoney0 = FormatNumber(v, 0, -1, 0, -1)
End Function
Function mmMetaStr(ByVal key, ByVal fallback)
  If Not mmMeta Is Nothing Then
    If mmMeta.Exists(key) Then
      mmMetaStr = mmMeta(key)
      Exit Function
    End If
  End If
  mmMetaStr = fallback
End Function
' Standalone helper because VBScript's `And` does not short-circuit -
' inlining `Not mmMeta Is Nothing And mmMeta.Exists(...)` raises
' "Object variable not set" when mmMeta is Nothing.
Function mmMetaHas(ByVal key)
  mmMetaHas = False
  If Not mmMeta Is Nothing Then
    If mmMeta.Exists(key) Then mmMetaHas = True
  End If
End Function

Dim mmBasePriceExDisp, mmBasePriceIncDisp
mmBasePriceExDisp  = mmFormatMoney0(mmBasePriceEx)
mmBasePriceIncDisp = mmFormatMoney(mmBasePriceInc)

' Resolve main hero image: large > standard > small > placeholder.
Dim mmMainImgSrc
If mmLargeImageUrl <> "" Then
  mmMainImgSrc = "/shop/pc/catalog/" & mmLargeImageUrl
ElseIf mmImageUrl <> "" Then
  mmMainImgSrc = "/shop/pc/catalog/" & mmImageUrl
ElseIf mmSmallImageUrl <> "" Then
  mmMainImgSrc = "/shop/pc/catalog/" & mmSmallImageUrl
Else
  mmMainImgSrc = "/shop/pc/catalog/no_image.gif"
End If

' Row-0 thumb tile uses the lighter image so we don't load the
' heavy hero asset twice on first paint.
Dim mmMainThumbSrc
If mmImageUrl <> "" Then
  mmMainThumbSrc = "/shop/pc/catalog/" & mmImageUrl
ElseIf mmSmallImageUrl <> "" Then
  mmMainThumbSrc = "/shop/pc/catalog/" & mmSmallImageUrl
Else
  mmMainThumbSrc = mmMainImgSrc
End If

' Canonical URL + absolute image URL for schema.org / share targets
Dim mmCanonicalUrl, mmCanonicalImg
If mmPcUrl <> "" Then
  mmCanonicalUrl = "https://www.multiplemonitors.co.uk/products/" & mmPcUrl & "/"
Else
  mmCanonicalUrl = "https://www.multiplemonitors.co.uk/shop/pc/viewprd-stand-v2.asp?idProduct=" & mmIdProduct
End If
mmCanonicalImg = "https://www.multiplemonitors.co.uk" & mmMainImgSrc

' ------------------------------------------------------------
' 2b. Gallery thumbs - row 0 = main product image,
'     rows 1+ = pcProductsImages ordered by pcProdImage_Order.
'     Empty result set is valid: only row 0 renders.
' ------------------------------------------------------------
Dim mmGalSql, mmGalRs, mmThumbs(), mmThumbCount, i
Dim mmThumbU, mmThumbL, mmThumbA

mmGalSql = "SELECT pcProdImage_Url, pcProdImage_LargeUrl, pcProdImage_AltTagText " & _
           "FROM pcProductsImages WHERE idProduct = " & mmIdProduct & _
           " ORDER BY pcProdImage_Order"
Set mmGalRs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
mmGalRs.Open mmGalSql, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If err.number <> 0 Then
  On Error Goto 0
  call LogErrorToDatabase()
  Set mmGalRs = Nothing
  call closeDB()
  Response.Redirect "techErr.asp?err=" & pcStrCustRefID
End If
On Error Goto 0

mmThumbCount = mmGalRs.RecordCount + 1   ' +1 for synthesised row 0
ReDim mmThumbs(mmThumbCount - 1, 2)       ' cols: 0=thumbSrc, 1=largeSrc, 2=alt

mmThumbs(0, 0) = mmMainThumbSrc
mmThumbs(0, 1) = mmMainImgSrc
mmThumbs(0, 2) = mmName

i = 1
Do While Not mmGalRs.EOF
  mmThumbU = mmGalRs("pcProdImage_Url") & ""
  mmThumbL = mmGalRs("pcProdImage_LargeUrl") & ""
  mmThumbA = Trim(mmGalRs("pcProdImage_AltTagText") & "")
  mmThumbs(i, 0) = "/shop/pc/catalog/" & mmThumbU
  If mmThumbL <> "" Then
    mmThumbs(i, 1) = "/shop/pc/catalog/" & mmThumbL
  Else
    mmThumbs(i, 1) = mmThumbs(i, 0)   ' fall back to small if no large
  End If
  If mmThumbA = "" Then mmThumbA = mmName
  mmThumbs(i, 2) = mmThumbA
  i = i + 1
  mmGalRs.MoveNext
Loop
mmGalRs.Close : Set mmGalRs = Nothing

' ------------------------------------------------------------
' 3. Page-level metadata consumed by inc_headerV5.asp
'    (set BEFORE the header_wrapper include so it wins)
' ------------------------------------------------------------
Dim pcv_PageName
pcv_PageName = mmName & " &mdash; UK monitor stand for multi-screen setups | Multiple Monitors"

' Highlight the Stands tab in the main nav
topmenuStands = " class=""is-trader"""

' Resolve max-monitor envelope numbers used in the SVG diagram
Dim mmMonMaxW, mmMonMaxH
mmMonMaxW = mmMetaStr("monMaxWidth",  "630")
mmMonMaxH = mmMetaStr("monMaxHeight", "450")
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
  "brand": { "@type": "Brand", "name": "Synergy Stands" },
  "manufacturer": { "@type": "Organization", "name": "Multiple Monitors Ltd" },
  "description": "<%= Replace(Replace(mmSDesc, """", "\"""), vbCrLf, " ") %>",
  "image": [<% For i = 0 To UBound(mmThumbs, 1) %>"https://www.multiplemonitors.co.uk<%= mmThumbs(i, 1) %>"<% If i < UBound(mmThumbs, 1) Then %>,<% End If %><% Next %>],
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
  }<% If mmMetaHas("aggRatingValue") And mmMetaHas("aggRatingCount") Then %>,
  "aggregateRating": {
    "@type": "AggregateRating",
    "ratingValue": "<%= mmMeta("aggRatingValue") %>",
    "reviewCount": "<%= mmMeta("aggRatingCount") %>",
    "bestRating": "5"
  }<% End If %>
}
</script>

<style>
  /* ============================================================
     Page-specific: Stand product page.
     Scoped under .mm-site so it can't leak onto legacy pages.
     Verbatim port of the inline CSS from redesign/standproduct.html
     - .pd-* trunk shared with monitor-v2 / trader-v2, plus the
     stand-specific .fit-check, .mon-dim, .vesa-badges, .dim-card,
     .reassure, .micro-band rules.
     Followup ticket: lift the shared .pd-*, .spec-*, .bundle-*,
     .darren-*, .sticky-cta blocks into mm-site.css once a third
     product-page template has stabilised the patterns.
     ============================================================ */

  /* -- Breadcrumb (mirrors viewPrd-Monitor-v2.asp) ---------------- */
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

  /* Floating chip */
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

  /* SKU mark */
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
    display:grid; grid-template-columns:repeat(6, 1fr); gap:10px; margin-top:14px;
  }
  @media (max-width:480px) {
    .mm-site .pd-gallery__thumbs { grid-template-columns:repeat(4, 1fr); }
    .mm-site .pd-gallery__thumbs .pd-thumb:nth-child(n+5) { display:none; }
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

  /* Trustpilot mini widget */
  .mm-site .pd-tp {
    display:inline-flex; align-items:center; gap:10px;
    padding:6px 0;
  }
  .mm-site .pd-tp .tp-stars { gap:2px; }
  .mm-site .pd-tp b {
    font-family:'EB Garamond', 'Georgia', serif; font-size:17px; color:var(--ink); font-weight:500;
  }
  .mm-site .pd-tp small {
    font-family:'JetBrains Mono', monospace; font-size:11px; letter-spacing:.08em;
    color:var(--muted);
  }
  .mm-site .pd-tp a {
    font-size:12.5px; color:var(--brand); font-weight:500;
    padding-left:10px; border-left:1px solid var(--line); margin-left:2px;
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

  /* -- Monitor-dimensions diagram (right panel of compatibility) - */
  .mm-site .mon-dim {
    background:
      radial-gradient(45% 55% at 30% 30%, rgba(15,110,168,.06), transparent 70%),
      linear-gradient(160deg, #F4F8FB 0%, #E8F1F8 100%);
    border:1px solid var(--line); border-radius:var(--radius-lg);
    padding:8px; margin-top:14px;
    display:flex; align-items:center; justify-content:center;
    position:relative; overflow:hidden;
  }
  .mm-site .mon-dim::before {
    content:""; position:absolute; inset:0; pointer-events:none; opacity:.3;
    background-image:
      linear-gradient(rgba(14,27,44,.06) 1px, transparent 1px),
      linear-gradient(90deg, rgba(14,27,44,.06) 1px, transparent 1px);
    background-size:28px 28px; border-radius:var(--radius-lg);
    mask-image:radial-gradient(ellipse at 50% 50%, #000 10%, transparent 75%);
    -webkit-mask-image:radial-gradient(ellipse at 50% 50%, #000 10%, transparent 75%);
  }
  .mm-site .mon-dim__svg {
    width:100%; max-width:230px; height:auto;
    display:block; position:relative; z-index:1;
  }
  .mm-site .mon-dim__svg .dim-label {
    font-family:'JetBrains Mono', monospace; font-size:22px;
    letter-spacing:.12em; fill:var(--ink); font-weight:500;
    paint-order:stroke fill; stroke:#EEF3F8; stroke-width:8px; stroke-linejoin:round;
  }
  .mm-site .mon-dim__svg .dim-tag {
    font-family:'JetBrains Mono', monospace; font-size:18px;
    letter-spacing:.16em; fill:var(--brand-deep); font-weight:500;
    text-transform:uppercase;
  }
  .mm-site .mon-dim__svg .dim-line { stroke:var(--slate); stroke-width:1; fill:none; }
  .mm-site .mon-dim__svg .dim-tick { stroke:var(--slate); stroke-width:1; }
  .mm-site .mon-dim__svg .mon-body { fill:url(#monGrad); stroke:var(--ink); stroke-width:3; }
  .mm-site .mon-dim__svg .mon-bezel { fill:none; stroke:rgba(14,27,44,.35); stroke-width:1.5; }
  .mm-site .vesa-mark { stroke:var(--brand); stroke-width:1.5; fill:none; }
  .mm-site .vesa-halo { fill:var(--brand-soft); }
  .mm-site .vesa-dot  { fill:var(--brand); }

  /* -- Reassurance pill ----------------------------------------- */
  .mm-site .reassure {
    display:flex; align-items:center; gap:10px;
    width:500px; max-width:100%;
    margin:-32px 0 18px auto;
    background:linear-gradient(160deg, rgba(33,166,122,.08), rgba(33,166,122,.02));
    border:1px solid rgba(33,166,122,.28); border-radius:999px;
    padding:8px 20px 8px 8px;
    font-size:13.5px; color:var(--ink); line-height:1.35;
  }
  .mm-site .reassure .tick {
    width:28px; height:28px; border-radius:50%;
    background:var(--up); color:#fff;
    display:inline-flex; align-items:center; justify-content:center;
    font-size:13px; flex-shrink:0;
  }
  .mm-site .reassure b {
    font-family:'EB Garamond', 'Georgia', serif; font-weight:500;
    color:var(--ink); font-size:15px; letter-spacing:-.005em; margin-right:2px;
  }

  /* -- VESA badges ---------------------------------------------- */
  .mm-site .vesa-badges {
    display:flex; flex-wrap:wrap; gap:14px; margin:22px 0 4px;
  }
  .mm-site .vesa-badge {
    display:inline-flex; flex-direction:column; align-items:center; gap:6px;
    background:#fff; border:1px solid var(--line-strong);
    border-radius:var(--radius-lg); padding:24px 20px 14px; min-width:110px;
    transition:all .2s ease;
  }
  .mm-site .vesa-badge:hover { border-color:var(--brand); box-shadow:var(--shadow); transform:translateY(-2px); }
  .mm-site .vesa-badge .pat {
    display:grid; grid-template-columns:repeat(2, 12px); grid-template-rows:repeat(2, 12px);
    gap:22px;
  }
  .mm-site .vesa-badge .pat.p100 { gap:30px; }
  .mm-site .vesa-badge .pat .h {
    width:12px; height:12px; border-radius:50%;
    background:var(--brand); box-shadow:0 0 0 3px var(--brand-soft);
  }
  .mm-site .vesa-badge .ttl {
    font-family:'EB Garamond', 'Georgia', serif; font-size:20px; font-weight:500;
    color:var(--ink); letter-spacing:-.01em; margin-top:6px;
  }
  .mm-site .vesa-badge .tag {
    font-family:'JetBrains Mono', monospace; font-size:10px;
    letter-spacing:.14em; text-transform:uppercase; color:var(--muted);
  }

  /* -- Dimensions diagram frame --------------------------------- */
  .mm-site .dim-card {
    background:
      radial-gradient(45% 55% at 30% 30%, rgba(15,110,168,.06), transparent 70%),
      linear-gradient(160deg, #F4F8FB 0%, #E8F1F8 100%);
    border:1px solid var(--line); border-radius:var(--radius-xl);
    padding:36px; display:flex; flex-direction:column; align-items:center; gap:20px;
    box-shadow:inset 0 0 0 1px rgba(255,255,255,.6);
    position:relative;
  }
  .mm-site .dim-card::before {
    content:""; position:absolute; inset:0; pointer-events:none; opacity:.35;
    background-image:
      linear-gradient(rgba(14,27,44,.06) 1px, transparent 1px),
      linear-gradient(90deg, rgba(14,27,44,.06) 1px, transparent 1px);
    background-size:28px 28px; border-radius:var(--radius-xl);
    mask-image:radial-gradient(ellipse at 50% 50%, #000 10%, transparent 75%);
    -webkit-mask-image:radial-gradient(ellipse at 50% 50%, #000 10%, transparent 75%);
  }
  .mm-site .dim-card img { max-width:100%; height:auto; position:relative; z-index:1; }
  .mm-site .dim-tabs {
    display:flex; gap:8px; align-self:flex-start; position:relative; z-index:1;
  }
  .mm-site .dim-tab {
    appearance:none; cursor:pointer;
    background:#fff; border:1px solid var(--line); color:var(--ink);
    padding:8px 14px; border-radius:999px;
    font-family:'JetBrains Mono', monospace; font-size:10.5px;
    letter-spacing:.14em; text-transform:uppercase; font-weight:500;
    transition:transform .15s ease, background .15s ease, color .15s ease, border-color .15s ease, box-shadow .15s ease;
  }
  .mm-site .dim-tab:hover { transform:translateY(-1px); box-shadow:0 2px 6px rgba(14,27,44,.08); }
  .mm-site .dim-tab.is-active { background:var(--brand); color:#fff; border-color:var(--brand); }
  .mm-site .dim-stats {
    display:grid; grid-template-columns:repeat(3, 1fr); gap:18px;
    margin-top:22px;
  }
  @media (max-width:640px) { .mm-site .dim-stats { grid-template-columns:1fr 1fr; } }
  .mm-site .dim-stat b {
    display:block; font-family:'EB Garamond', 'Georgia', serif;
    font-size:28px; font-weight:500; color:var(--ink); letter-spacing:-.018em;
    line-height:1;
  }
  .mm-site .dim-stat b .u {
    font-family:'JetBrains Mono', monospace; font-size:12px;
    color:var(--muted); margin-left:4px; letter-spacing:.08em; font-weight:500;
  }
  .mm-site .dim-stat small {
    display:block; margin-top:6px;
    font-family:'JetBrains Mono', monospace; font-size:10.5px;
    letter-spacing:.14em; text-transform:uppercase; color:var(--muted);
  }

  /* -- Micro-band assembly trio --------------------------------- */
  .mm-site .micro-band {
    padding:28px 0; background:#fff;
    border-top:1px solid var(--line); border-bottom:1px solid var(--line);
  }
  .mm-site .micro-band .hero-mini {
    display:grid; grid-template-columns:repeat(3, 1fr); gap:18px;
    margin:0; padding:0; border:0;
  }
  @media (max-width:680px) { .mm-site .micro-band .hero-mini { grid-template-columns:1fr; } }
  .mm-site .micro-band .hero-mini .item { justify-content:center; text-align:left; }
  .mm-site .micro-band .hero-mini .item .fa {
    color:var(--accent); font-size:22px;
    width:34px; height:34px; border-radius:50%;
    background:rgba(242,167,27,.12); display:inline-flex;
    align-items:center; justify-content:center; flex-shrink:0;
  }
  .mm-site .micro-band .hero-mini .item span b {
    display:block; font-family:'EB Garamond', 'Georgia', serif;
    font-size:18px; color:var(--ink); font-weight:500; letter-spacing:-.005em;
  }
  .mm-site .micro-band .hero-mini .item span small {
    font-family:'JetBrains Mono', monospace; font-size:10.5px;
    letter-spacing:.12em; text-transform:uppercase; color:var(--muted);
  }
</style>

<div class="mm-site">

<!-- ===================================================================
     BREADCRUMB
     =================================================================== -->
<nav class="breadcrumb" aria-label="Breadcrumb">
  <div class="container inner">
    <a href="/">Home</a>
    <span class="sep">/</span>
    <a href="/stands/">Synergy Stands</a>
    <span class="sep">/</span>
    <span class="current"><%= mmName %></span>
  </div>
</nav>

<!-- ===================================================================
     PRODUCT HERO - gallery + buy-box
     =================================================================== -->
<section class="pd-hero">
  <div class="container">
    <div class="pd-hero-grid">

      <!-- Gallery column -->
      <div class="pd-gallery reveal" id="pdGallery">
        <div class="pd-gallery__main">
          <span class="pd-gallery__chip">
            <span class="dot"></span><span class="acc">UK</span>MADE&nbsp;&middot;&nbsp;LIFETIME&nbsp;WARRANTY
          </span>
          <img id="pdMainImg"
               src="<%= mmMainImgSrc %>"
               alt="<%= Server.HTMLEncode(mmName) %>" />
          <% If mmSku <> "" Then %>
          <span class="pd-gallery__sku">SKU &middot; <%= Server.HTMLEncode(mmSku) %></span>
          <% End If %>
        </div>

        <div class="pd-gallery__thumbs">
          <% For i = 0 To UBound(mmThumbs, 1) %>
          <div class="pd-thumb<% If i = 0 Then %> is-active<% End If %>"
               data-img="<%= mmThumbs(i, 1) %>"
               title="<%= Server.HTMLEncode(mmThumbs(i, 2)) %>">
            <img src="<%= mmThumbs(i, 0) %>"
                 alt="<%= Server.HTMLEncode(mmThumbs(i, 2)) %>" />
          </div>
          <% Next %>
        </div>
      </div>

      <!-- Buy-box column -->
      <aside class="pd-buybox reveal" style="transition-delay:.08s">
        <div class="eyebrow"><%= mmMetaStr("eyebrow", "Synergy Stand") %></div>
        <h1><%= mmMetaStr("Title", "Quad Square") %> <em>Synergy Stand</em></h1>
        <p class="pitch">
          <% If mmTopDesc <> "" Then %>
            <%= mmTopDesc %>
          <% ElseIf mmSDesc <> "" Then %>
            <%= mmSDesc %>
          <% Else %>
            <%= mmMetaStr("pitch", "All-steel UK-designed and UK-built monitor stand. Modular, lifetime warranty, fits standard VESA monitors.") %>
          <% End If %>
        </p>

        <div class="pd-tp">
          <span class="tp-stars"><span></span><span></span><span></span><span></span><span></span></span>
          <b>4.9</b>
          <small>&middot; 90+ reviews</small>
          <a href="https://uk.trustpilot.com/review/multiplemonitors.co.uk" target="_blank" rel="noopener">Read reviews <i class="fa fa-arrow-up-right" style="font-size:10px;"></i></a>
        </div>

        <div class="pd-price">
          <div>
            <div class="pd-price__from">Price</div>
            <div class="pd-price__num"><span class="sym">&pound;</span><%= mmBasePriceExDisp %></div>
          </div>
          <div class="pd-price__vat">
            <b>&pound;<%= mmBasePriceIncDisp %></b> inc VAT
            <span class="ship">UK delivery &pound;10 flat &middot; international from &pound;20</span>
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
            <i class="fa fa-flag"></i>
            <div><b>UK-made</b><small>Since 2016</small></div>
          </div>
          <div class="item">
            <i class="fa fa-shield"></i>
            <div><b>Lifetime warranty</b><small>All parts</small></div>
          </div>
          <div class="item">
            <i class="fa fa-truck"></i>
            <div><b>2-day dispatch</b><small>UK stock</small></div>
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
            Add Stand To Your Basket <i class="fa fa-arrow-right"></i>
          </button>
        </form>

        <div class="pd-foot">
          <span><i class="fa fa-phone"></i>Call Darren &mdash; 0330 223 66 55</span>
          <span><i class="fa fa-calendar"></i>Saturday &amp; custom delivery on request</span>
        </div>
      </aside>

    </div>
  </div>
</section>

<!-- ===================================================================
     TRUST STRIP - reused sitewide partial
     =================================================================== -->
<!--#include file="inc_trustStripTrader.asp"-->

<!-- ===================================================================
     KEY SPECS - 6-card grid + "In the box"
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="section-head-narrow reveal">
      <h5 style="margin-bottom:14px;">Specifications at a glance</h5>
      <h2><%= mmMetaStr("ScreenText", "Four") %> Screens. <span class="display-em"><%= mmMetaStr("ColumnText", "One") %> elegant column.</span></h2>
      <p>Strong, stable and adjustable, supplied with everything you need to mount your screens.</p>
    </div>

    <div class="spec-grid">
      <%
      Dim mmCardN, mmCardDelay
      For mmCardN = 1 To 6
        Select Case mmCardN
          Case 1 : mmCardDelay = ""
          Case 2 : mmCardDelay = " style=""transition-delay:.06s"""
          Case 3 : mmCardDelay = " style=""transition-delay:.12s"""
          Case 4 : mmCardDelay = ""
          Case 5 : mmCardDelay = " style=""transition-delay:.06s"""
          Case 6 : mmCardDelay = " style=""transition-delay:.12s"""
        End Select
      %>
      <div class="spec-card reveal"<%= mmCardDelay %>>
        <div class="spec-card__icon"><i class="fa <%= mmMetaStr("specCard" & mmCardN & "_icon", "fa-check") %>"></i></div>
        <div class="spec-card__label"><%= mmMetaStr("specCard" & mmCardN & "_label", "&nbsp;") %></div>
        <div class="spec-card__value"><%= mmMetaStr("specCard" & mmCardN & "_value", "&nbsp;") %></div>
        <div class="spec-card__desc"><%= mmMetaStr("specCard" & mmCardN & "_desc", "&nbsp;") %></div>
      </div>
      <% Next %>
    </div>

    <div class="spec-box reveal" style="transition-delay:.18s">
      <div class="spec-box__lead">
        <div class="spec-box__icon"><i class="fa fa-archive"></i></div>
        <div>
          <div class="spec-box__label">In the box</div>
          <div class="spec-box__title">Everything you need, supplied together.</div>
        </div>
      </div>
      <div class="spec-chips">
        <%
        If mmMetaHas("inTheBox") Then
          Dim mmBoxItems, mmBoxItem
          mmBoxItems = mmMeta("inTheBox")
          For Each mmBoxItem In mmBoxItems
        %>
        <span class="spec-chip"><i class="fa fa-check"></i><%= mmBoxItem %></span>
        <%
          Next
        Else
        %>
        <span class="spec-chip"><i class="fa fa-check"></i>Base &amp; central column</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Arm assemblies</span>
        <span class="spec-chip"><i class="fa fa-check"></i>VESA plates</span>
        <span class="spec-chip"><i class="fa fa-check"></i>All fixings</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Cable management ties</span>
        <span class="spec-chip"><i class="fa fa-check"></i>All assembly tools</span>
        <% End If %>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     ASSEMBLY / WARRANTY MICRO-TRIO
     =================================================================== -->
<section class="micro-band">
  <div class="container">
    <div class="hero-mini reveal">
      <div class="item">
        <i class="fa fa-wrench"></i>
        <span>
          <b><%= mmMetaStr("microAssembly", "30&ndash;60 min assembly") %></b>
          <small><%= mmMetaStr("microAssemblySub", "No drilling, no wall fixings") %></small>
        </span>
      </div>
      <div class="item">
        <i class="fa fa-cog"></i>
        <span>
          <b><%= mmMetaStr("microTools", "All tools included") %></b>
          <small><%= mmMetaStr("microToolsSub", "Nothing to buy separately") %></small>
        </span>
      </div>
      <div class="item">
        <i class="fa fa-certificate"></i>
        <span>
          <b><%= mmMetaStr("microWarranty", "Lifetime warranty") %></b>
          <small><%= mmMetaStr("microWarrantySub", "On every steel part, forever") %></small>
        </span>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     COMPATIBILITY / VESA / MONITOR DIMENSIONS
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Will my monitors fit?</h5>
        <h2>Check if your screens, <span class="display-em">fit on this stand</span>.</h2>
        <p style="max-width:760px; margin-top:12px;">Every Synergy Stand is designed to use the two VESA patterns found on most modern monitors. Check the pattern on the back of your monitor and confirm its dimensions sit within the envelope shown below.</p>
      </div>
    </div>
    <div class="reassure" role="note">
      <span class="tick" aria-hidden="true"><i class="fa fa-check"></i></span>
      <span>Verified compatible with every monitor sold in our bundles and arrays.</span>
    </div>
    <div class="bench-panels">
      <div class="bench-panel reveal">
        <h4>VESA compatibility</h4>
        <span class="sub">Both patterns, every mount</span>

        <div class="vesa-badges">
          <div class="vesa-badge">
            <div class="pat" aria-hidden="true"><span class="h"></span><span class="h"></span><span class="h"></span><span class="h"></span></div>
            <div class="ttl">75 &times; 75</div>
            <div class="tag">VESA MIS-D, 75</div>
          </div>
          <div class="vesa-badge">
            <div class="pat p100" aria-hidden="true"><span class="h"></span><span class="h"></span><span class="h"></span><span class="h"></span></div>
            <div class="ttl">100 &times; 100</div>
            <div class="tag">VESA MIS-D, 100</div>
          </div>
        </div>

        <p style="margin-top:20px; color:var(--slate); line-height:1.55;"><%= mmMetaStr("vesaIntro1", "If you already have screens or are purchasing some new ones then look for a <b style=""color:var(--ink);"">VESA 75</b> or <b style=""color:var(--ink);"">VESA 100</b> rating, sometimes described as a '<b style=""color:var(--ink);"">Wall mount interface</b>'.") %></p>
        <p style="margin-top:20px; color:var(--slate); line-height:1.55;"><%= mmMetaStr("vesaIntro2", "This is the four screw holes in the back of a screen in a square configuration. If your monitor has them then it should be compatible with this Synergy Stand.") %></p>
        <p style="color:var(--muted); font-size:13px; margin-top:10px;">
          Got an unusual monitor? <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">Ring 0330 223 66 55</a> and we&rsquo;ll check it for you.
        </p>
      </div>

      <div class="bench-panel reveal" style="transition-delay:.08s;">
        <h4>Monitor dimensions</h4>
        <span class="sub">Maximum envelope per screen</span>

        <div class="mon-dim">
          <svg class="mon-dim__svg" viewBox="0 0 460 360" role="img"
               aria-label="Diagram of a single monitor showing maximum width <%= mmMonMaxW %> mm and maximum height <%= mmMonMaxH %> mm with the VESA mounting point in the centre">
            <defs>
              <linearGradient id="monGrad" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stop-color="#FFFFFF"/>
                <stop offset="100%" stop-color="#E8F1F8"/>
              </linearGradient>
              <marker id="dimArrow" viewBox="0 0 10 10" refX="9" refY="5"
                      markerWidth="7" markerHeight="7" orient="auto-start-reverse">
                <path d="M 0 0 L 10 5 L 0 10 z" fill="#455065"/>
              </marker>
            </defs>

            <rect class="mon-body"  x="40" y="70" width="340" height="240" rx="4"/>
            <rect class="mon-bezel" x="48" y="78" width="324" height="224" rx="2"/>

            <circle class="vesa-halo" cx="210" cy="190" r="14"/>
            <line   class="vesa-mark" x1="198" y1="190" x2="222" y2="190"/>
            <line   class="vesa-mark" x1="210" y1="178" x2="210" y2="202"/>
            <circle class="vesa-dot"  cx="210" cy="190" r="2.5"/>
            <text class="dim-tag" x="210" y="232" text-anchor="middle">VESA centre</text>

            <line class="dim-tick" x1="40"  y1="64" x2="40"  y2="34" stroke-dasharray="2,2"/>
            <line class="dim-tick" x1="380" y1="64" x2="380" y2="34" stroke-dasharray="2,2"/>
            <line class="dim-line" x1="40"  y1="40" x2="380" y2="40"
                  marker-start="url(#dimArrow)" marker-end="url(#dimArrow)"/>
            <text class="dim-label" x="210" y="44" text-anchor="middle"><%= mmMonMaxW %> mm</text>

            <line class="dim-tick" x1="386" y1="70"  x2="416" y2="70"  stroke-dasharray="2,2"/>
            <line class="dim-tick" x1="386" y1="310" x2="416" y2="310" stroke-dasharray="2,2"/>
            <line class="dim-line" x1="410" y1="70"  x2="410" y2="310"
                  marker-start="url(#dimArrow)" marker-end="url(#dimArrow)"/>
            <g transform="translate(414 190) rotate(-90)">
              <text class="dim-label" x="0" y="4" text-anchor="middle"><%= mmMonMaxH %> mm</text>
            </g>
          </svg>
        </div>

        <p style="margin-top:20px; color:var(--slate); line-height:1.55;"><%= mmMetaStr("monMaxNote1", "These limits still leave room for a <b style=""color:var(--ink);"">gentle curve</b> across a multi-screen setup.") %></p>
        <p style="margin-top:14px; color:var(--slate); line-height:1.55;"><%= mmMetaStr("monMaxNote2", "The height assumes the VESA pattern sits roughly in the middle of the screen. If your VESA holes are towards the top on the back of your screen then the max height may be reduced somewhat.") %></p>
        <p style="color:var(--muted); font-size:13px; margin-top:10px;">
          Not sure? <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">Ring 0330 223 66 55</a> and we&rsquo;ll check the model for you.
        </p>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     FOOTPRINT & DIMENSIONS - 6 stat grid + tabbed dimension card
     =================================================================== -->
<section class="s specs">
  <div class="container">
    <div class="hero-grid">
      <div class="reveal">
        <div class="eyebrow">Footprint &amp; dimensions</div>
        <h2>The space of two. <span class="display-em">The capacity of four.</span></h2>
        <p class="lead"><%= mmMetaStr("dimsLead", "A single central column keeps the desk surface below the screens clear for keyboard, notebook and coffee.") %></p>

        <div class="dim-stats">
          <% Dim mmStatN
          For mmStatN = 1 To 6 %>
          <div class="dim-stat">
            <b><%= mmMetaStr("dimStat" & mmStatN & "_value", "&nbsp;") %></b>
            <small><%= mmMetaStr("dimStat" & mmStatN & "_label", "&nbsp;") %></small>
          </div>
          <% Next %>
        </div>

        <% If mmMetaStr("dimPdf", "") <> "" Then %>
        <p style="margin-top:22px; font-size:13px; color:var(--muted);">
          <a href="<%= mmMetaStr("dimPdf", "") %>" style="color:var(--brand); font-weight:500;"><i class="fa fa-file-pdf-o"></i> Download the full assembly &amp; dimension PDF</a>
        </p>
        <% End If %>
      </div>

      <div class="reveal" style="transition-delay:.08s">
        <div class="dim-card">
          <div class="dim-tabs" role="tablist" aria-label="Stand dimension view">
            <button type="button" class="dim-tab is-active" role="tab" aria-selected="true"
                    data-dim-img="<%= mmMetaStr("dimImgFront", "/images/stands/dim-4s.jpg") %>"
                    data-dim-alt="<%= mmMetaStr("dimImgFrontAlt", "Front-elevation engineering drawing of the stand with dimensions in millimetres") %>">
              Front profile
            </button>
            <button type="button" class="dim-tab" role="tab" aria-selected="false"
                    data-dim-img="<%= mmMetaStr("dimImgSide", "/images/stands/dim-side-tall.jpg") %>"
                    data-dim-alt="<%= mmMetaStr("dimImgSideAlt", "Side-elevation engineering drawing of the stand with dimensions in millimetres") %>">
              Side profile
            </button>
          </div>
          <img id="dimImg"
               src="<%= mmMetaStr("dimImgFront", "/images/stands/dim-4s.jpg") %>"
               alt="<%= mmMetaStr("dimImgFrontAlt", "Front-elevation engineering drawing of the stand with dimensions in millimetres") %>" />
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     BUNDLE UPSELL - dark band
     =================================================================== -->
<section class="bundle">
  <div class="container">
    <div class="bundle-grid">
      <div class="reveal">
        <h5>Complete your setup</h5>
        <h2>Save money and get free upgrades with <em>a bundle or monitor array</em>.</h2>
        <p>Monitor arrays include the stand and screens, you get free cables and free delivery. Bundles include the stand, screens and a multi-screen PC with free upgrades, free cables, free delivery and a bundle discount.</p>
        <div class="bundle-pills">
          <span class="bundle-pill"><i class="fa fa-check"></i>Free 3&nbsp;m video cables</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free UK delivery</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free WiFi card<span>*</span></span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free speakers<span>*</span></span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Auto bundle discount<span>*</span></span>
        </div>
        <div style="display:flex; gap:12px; flex-wrap:wrap;">
          <a href="/display-systems/" class="btn btn-accent btn-lg">See monitor arrays <i class="fa fa-arrow-right"></i></a>
          <a href="/bundles/" class="btn btn-accent btn-lg">See PC bundles <i class="fa fa-arrow-right"></i></a>
        </div>
        <p class="bundle-foot"><span class="bundle-foot__star">*</span>Included on computer bundles only</p>
      </div>
      <div class="reveal" style="transition-delay:.1s">
        <div class="save-card">
          <span class="save-tag">Example &middot; 4-screen computer bundle</span>
          <div class="kicker">Typical saving vs buying separately</div>
          <div class="big"><small>&pound;</small>190</div>
          <div class="sub">Synergy Stand + Four Screens + Multi-Screen PC.</div>
          <div class="breakdown">
            <div class="r"><span>4&thinsp;&times;&thinsp;3&nbsp;m video cables</span><b>&pound;60</b></div>
            <div class="r"><span>WiFi, BT &amp; speakers</span><b>&pound;60</b></div>
            <div class="r"><span>UK mainland delivery</span><b>&pound;20</b></div>
            <div class="r"><span>Bundle discount</span><b>&pound;50</b></div>
            <div class="r total"><span>Total savings</span><b>&minus;&thinsp;&pound;190</b></div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     DARREN CTA - shared partial
     =================================================================== -->
<!--#include file="inc_darrenCTA.asp"-->

<!-- ===================================================================
     STICKY ADD-TO-BASKET CTA
     =================================================================== -->
<form method="post" action="/shop/pc/instPrd.asp" class="sticky-cta" id="stickyCta">
  <input type="hidden" name="idproduct" value="<%= mmIdProduct %>">
  <input type="hidden" name="quantity" value="1">
  <input type="hidden" name="OptionGroupCount" value="0">
  <div class="txt">
    <strong><%= mmName %> &middot; &pound;<%= mmBasePriceExDisp %> + VAT</strong>
    <span>Order before <%= daFunDelCutOff() %> &middot; delivered <%= daFunDelDateReturn(0,0) %></span>
  </div>
  <button type="submit" class="btn btn-primary btn-sm">Add to basket <i class="fa fa-arrow-right"></i></button>
</form>

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
  if (main) {
    thumbs.forEach(function(t){
      t.addEventListener('click', function(){
        document.querySelectorAll('.pd-thumb.is-active').forEach(function(x){ x.classList.remove('is-active'); });
        t.classList.add('is-active');
        main.src = t.dataset.img;
        var imgEl = t.querySelector('img');
        if (imgEl && imgEl.alt) main.alt = imgEl.alt;
      });
    });
  }

  // -- Dimensions tabs - swap front/side profile image --
  var dimImg  = document.getElementById('dimImg');
  var dimTabs = document.querySelectorAll('.dim-tab[data-dim-img]');
  if (dimImg && dimTabs.length) {
    dimTabs.forEach(function(t){
      t.addEventListener('click', function(){
        dimTabs.forEach(function(x){ x.classList.remove('is-active'); x.setAttribute('aria-selected','false'); });
        t.classList.add('is-active');
        t.setAttribute('aria-selected','true');
        dimImg.src = t.dataset.dimImg;
        dimImg.alt = t.dataset.dimAlt;
      });
    });
  }

  // -- Sticky Add-to-Basket - visible after hero, hidden near footer --
  var sticky = document.getElementById('stickyCta');
  var hero   = document.querySelector('.pd-hero');
  var footerEl = document.querySelector('footer');
  if (sticky && hero && footerEl) {
    function onStickyScroll(){
      var y = window.scrollY || window.pageYOffset;
      var heroBottom = hero.getBoundingClientRect().bottom + y;
      var footerTop  = footerEl.getBoundingClientRect().top + y;
      var viewportBottom = y + window.innerHeight;
      if (y > heroBottom + 120 && viewportBottom < footerTop) {
        sticky.classList.add('visible');
      } else {
        sticky.classList.remove('visible');
      }
    }
    window.addEventListener('scroll', onStickyScroll, { passive:true });
    onStickyScroll();
  }
})();
</script>

<!--#include file="footer_wrapper.asp"-->
