<%
' ============================================================
' viewprd-stand-v2.asp
' 2026 redesign — Synergy Stand product page.
' Template renders any single Synergy Stand from the products
' table.
'
' Resolution order (single SELECT, indexed lookup):
'   1. Request.QueryString("slug") — preserved across the
'      Server.Transfer from viewPrdRouter.asp on friendly-URL
'      hits. Resolves WHERE pcUrl = ?
'   2. Request.QueryString("idProduct") — direct deep-link
'      fallback for any legacy code linking to this page with
'      ?idProduct=N
'   3. Hardcoded test fallback to the Quad Square Synergy
'      Stand, used when the page is loaded with no params.
'
' Once mmIdProduct is hydrated, Session("idProductRedirect") is
' set so legacy ProductCart includes (include-metatags.asp,
' inc_footer.asp, apps/pcBackInStock) see the right value.
' Page-local rendering uses the captured mmIdProduct, so a
' concurrent tab overwriting the session var cannot corrupt
' this page.
'
' Phase 1 scope — see /we-need-to-build-playful-catmull.md:
'   * Name, SKU, price, main image, short description and
'     canonical URL come from the products table.
'   * Per-stand variant copy (specs grid, FAQ, sibling cards,
'     mounting/cable/footprint sections, gallery thumbs) is
'     hardcoded for the Quad Square Synergy Stand. Other 11
'     stands will load dynamic header data but use this copy
'     until follow-up work pulls per-stand specs from custom
'     fields or a new pcStandSpecs table.
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
<%
Const MM_VAT_RATE = 1.2

' ------------------------------------------------------------
' 1. Product base row — slug-first, single indexed SELECT
' ------------------------------------------------------------
Dim mmIdProduct, mmName, mmSku, mmBasePriceInc, mmImageUrl, mmSmallImageUrl
Dim mmSDesc, mmPcUrl, mmAdditionalImages, mmAltTagText, mmStock
mmIdProduct = 0
mmName = ""              : mmSku = ""              : mmBasePriceInc = 0
mmImageUrl = ""          : mmSmallImageUrl = ""    : mmSDesc = ""
mmPcUrl = ""             : mmAdditionalImages = "" : mmAltTagText = ""
mmStock = 0

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
mmPrdSql = "SELECT idProduct, description, sku, price, imageUrl, smallImageUrl, " & _
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

' Resolve main image with fallbacks (matches viewPrd-TraderPC.asp)
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
  mmCanonicalUrl = "https://www.multiplemonitors.co.uk/shop/pc/viewprd-stand-v2.asp?idProduct=" & mmIdProduct
End If
mmCanonicalImg = "https://www.multiplemonitors.co.uk" & mmMainImgSrc

' ------------------------------------------------------------
' 3. Builder context — bundle / array hand-off
'    Mirrors the sid/arr/cid pattern in the legacy
'    viewPrd-Stands.asp. When set the CTA links to the array
'    or bundle builder rather than dropping into the cart.
' ------------------------------------------------------------
Dim mmCtxSid, mmCtxMid, mmCtxCid, mmIsArrayBuild, mmIsBundleBuild
mmCtxSid = Trim(Request.QueryString("sid") & "")
mmCtxMid = Trim(Request.QueryString("mid") & "")
mmCtxCid = Trim(Request.QueryString("cid") & "")
mmIsArrayBuild  = (mmCtxSid <> "" And Request.QueryString("arr") = "1")
mmIsBundleBuild = (mmCtxSid <> "" And mmCtxCid <> "" And Not mmIsArrayBuild)

' ------------------------------------------------------------
' 4. Page-level metadata consumed by inc_headerV5.asp
'    (set BEFORE the header_wrapper include so it wins)
' ------------------------------------------------------------
Dim pcv_PageName
pcv_PageName = mmName & " — UK-made monitor stand | Multiple Monitors"

' Highlight the Stands tab in the main nav
topmenuStands = " class=""is-trader"""
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
  },
  "aggregateRating": {
    "@type": "AggregateRating",
    "ratingValue": "4.9",
    "reviewCount": "90",
    "bestRating": "5"
  }
}
</script>

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
    <span class="current"><%= Server.HTMLEncode(mmName) %></span>
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
            <span class="dot"></span><span class="acc">UK</span>MADE&nbsp;&middot;&nbsp;LIFETIME&nbsp;WARRANTY
          </span>
          <img id="pdMainImg"
               src="<%= mmMainImgSrc %>"
               alt="<%= Server.HTMLEncode(mmName) %>" />
          <div class="pd-video" id="pdVideoOverlay">
            <button type="button" aria-label="Play 60-second walkthrough video" onclick="window.open('/pop-pages/stand-video.asp?s=<%= LCase(Server.URLEncode(mmSku)) %>','stand-video','width=780,height=520');">
              <i class="fa fa-play"></i>
            </button>
          </div>
          <% If mmSku <> "" Then %>
          <span class="pd-gallery__sku">SKU &middot; <%= Server.HTMLEncode(mmSku) %></span>
          <% End If %>
        </div>

        <%
        ' TODO (phase-2): parse mmAdditionalImages (comma-separated)
        ' to drive these thumbs dynamically. Hardcoded for the
        ' Quad Square mockup for now.
        %>
        <div class="pd-gallery__thumbs thumbs-6">
          <div class="pd-thumb is-active has-video"
               data-img="<%= mmMainImgSrc %>"
               data-video="1"
               title="60-second walkthrough video">
            <img src="<%= mmMainImgSrc %>" alt="Walkthrough video poster &mdash; <%= Server.HTMLEncode(mmName) %>" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/4s-front-angle-med.jpg">
            <img src="/shop/pc/catalog/4s-front-angle-thm.jpg" alt="Front three-quarter view showing 2&times;2 screen arrangement" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/4s-front-med.jpg">
            <img src="/shop/pc/catalog/4s-front-thm.jpg" alt="Head-on front view of four screens in square layout" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/4s-rear-angle-med.jpg">
            <img src="/shop/pc/catalog/4s-rear-angle-thm.jpg" alt="Rear three-quarter view showing central column and cable channel" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/arm-pivot.jpg">
            <img src="/shop/pc/catalog/arm-pivot.jpg" alt="Close-up of arm pivot joint" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/arm-vesa-height.jpg">
            <img src="/shop/pc/catalog/arm-vesa-height.jpg" alt="Close-up of VESA mount and 30&nbsp;mm fine height adjustment" />
          </div>
        </div>
      </div>

      <!-- Buy-box column -->
      <aside class="pd-buybox reveal" style="transition-delay:.08s">
        <div class="eyebrow">Synergy Stand &middot; 4-screen square</div>
        <h1>Quad Square <em>Synergy Stand</em></h1>
        <p class="pitch">
          <% If mmSDesc <> "" Then %>
            <%= mmSDesc %>
          <% Else %>
            Four screens in the footprint of two. All-steel, UK-designed and UK-built &mdash;
            a square 2&times;2 layout that hides cables behind the column and scales up to six screens later.
          <% End If %>
        </p>

        <div class="pd-tp">
          <span class="tp-stars"><span></span><span></span><span></span><span></span><span></span></span>
          <b>4.9</b>
          <small>&middot; 90+ reviews</small>
          <a href="#reviews">Read reviews <i class="fa fa-arrow-down" style="font-size:10px;"></i></a>
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

        <% If mmIsArrayBuild Then %>
          <div class="pd-action">
            <a href="/display-systems-2/?sid=<%= mmIdProduct %>&amp;mid=<%= Server.URLEncode(mmCtxMid) %>" class="btn btn-primary btn-lg" style="grid-column:1 / -1;">
              Add Stand To Your Array <i class="fa fa-arrow-right"></i>
            </a>
          </div>
        <% ElseIf mmIsBundleBuild Then %>
          <div class="pd-action">
            <a href="/bundles-2/?sid=<%= mmIdProduct %>&amp;mid=<%= Server.URLEncode(mmCtxMid) %>&amp;cid=<%= Server.URLEncode(mmCtxCid) %>" class="btn btn-primary btn-lg" style="grid-column:1 / -1;">
              Add Stand To Your Bundle <i class="fa fa-arrow-right"></i>
            </a>
          </div>
        <% Else %>
          <form method="POST" action="/shop/pc/instPrd.asp" class="pd-action">
            <input type="hidden" name="idproduct" value="<%= mmIdProduct %>" />
            <div class="pd-qty" role="group" aria-label="Quantity">
              <button type="button" aria-label="Decrease quantity" onclick="(function(i){i.value=Math.max(1,(parseInt(i.value,10)||1)-1);})(document.getElementById('pdQty'));">&minus;</button>
              <input id="pdQty" name="Qty" type="number" min="1" value="1" aria-label="Quantity" />
              <button type="button" aria-label="Increase quantity" onclick="(function(i){i.value=(parseInt(i.value,10)||1)+1;})(document.getElementById('pdQty'));">+</button>
            </div>
            <button type="submit" class="btn btn-primary btn-lg">
              Add Stand To Your Basket <i class="fa fa-arrow-right"></i>
            </button>
          </form>
        <% End If %>

        <div class="pd-foot">
          <span><i class="fa fa-phone"></i>Call Darren &mdash; 0330 223 66 55</span>
          <span><i class="fa fa-calendar"></i>Saturday &amp; custom delivery on request</span>
        </div>
      </aside>

    </div>
  </div>
</section>

<!-- ===================================================================
     TRUST STRIP — stand-flavoured copy (BBC, Trustpilot, Est 2008,
     stands sold). Inlined here rather than reusing the trader-pc
     include because the fourth tile differs.
     =================================================================== -->
<section class="truststrip">
  <div class="container">
    <div class="inner">
      <div class="trust-item bbc">
        <div class="icon"><i class="fa fa-television"></i></div>
        <div>
          <div class="label">As seen on the <span class="bbc-mark">BBC</span></div>
          <div class="val">Traders: Millions by the Minute</div>
        </div>
      </div>
      <div class="trust-item tp">
        <div class="icon"><i class="fa fa-star"></i></div>
        <div>
          <div class="label">Trustpilot &middot; 4.9&thinsp;/&thinsp;5</div>
          <div class="val">90+ Unsolicited Reviews</div>
        </div>
      </div>
      <div class="trust-item">
        <div class="icon"><i class="fa fa-clock-o"></i></div>
        <div>
          <div class="label">Established 2008</div>
          <div class="val">17+ years of experience</div>
        </div>
      </div>
      <div class="trust-item accent">
        <div class="icon"><i class="fa fa-th-large"></i></div>
        <div>
          <div class="label">Sold since 2016</div>
          <div class="val">3,000+ stands in use</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     KEY SPECS GRID — Quad Square copy (per-stand later)
     =================================================================== -->
<section class="s specs">
  <div class="container">
    <div class="section-head-narrow reveal">
      <h5 style="margin-bottom:14px;">Specifications at a glance</h5>
      <h2>Four screens. <span class="display-em">One elegant column.</span></h2>
      <p>Everything below the hero has one job: answer the questions real buyers ask before they pay. No marketing fluff, no hidden footnotes.</p>
    </div>

    <div class="spec-grid">
      <div class="spec-card reveal">
        <div class="spec-card__icon"><i class="fa fa-th-large"></i></div>
        <div class="spec-card__label">Layout</div>
        <div class="spec-card__value">4 screens, 2&times;2 square</div>
        <div class="spec-card__desc">Two screens above, two below &mdash; takes the footprint of two side-by-side monitors and stacks the other two vertically. Ideal for traders watching four instruments at once.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.06s">
        <div class="spec-card__icon"><i class="fa fa-expand"></i></div>
        <div class="spec-card__label">Max screen size</div>
        <div class="spec-card__value">Up to 28&Prime; per mount</div>
        <div class="spec-card__desc">Designed with room to angle outer screens inward for a proper curve &mdash; not the &ldquo;up to 24&Prime;&rdquo; limit you see on imported stands.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.12s">
        <div class="spec-card__icon"><i class="fa fa-balance-scale"></i></div>
        <div class="spec-card__label">Max weight</div>
        <div class="spec-card__value">8&nbsp;kg per screen</div>
        <div class="spec-card__desc">Per mount. Comfortably carries any 24&Prime;&ndash;28&Prime; office or gaming monitor we&rsquo;ve tested &mdash; including curved ultrawides up to 29&Prime;.</div>
      </div>
      <div class="spec-card reveal">
        <div class="spec-card__icon"><i class="fa fa-crosshairs"></i></div>
        <div class="spec-card__label">VESA</div>
        <div class="spec-card__value">75&times;75 &amp; 100&times;100</div>
        <div class="spec-card__desc">Fits every monitor we sell, and the vast majority of Dell, LG, Samsung, AOC, BenQ and ASUS displays &mdash; no adapters, no guesswork.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.06s">
        <div class="spec-card__icon"><i class="fa fa-sliders"></i></div>
        <div class="spec-card__label">Adjustability</div>
        <div class="spec-card__value">6 degrees of freedom</div>
        <div class="spec-card__desc">Height, arm hinge, horizontal slide, pivot, tilt and 30&nbsp;mm of fine height adjust at every mount. Set once, locks solid.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.12s">
        <div class="spec-card__icon"><i class="fa fa-cubes"></i></div>
        <div class="spec-card__label">Materials</div>
        <div class="spec-card__value">All-steel, no plastic</div>
        <div class="spec-card__desc">No load-bearing plastic parts. The entire frame, column and arms are welded and powder-coated steel.</div>
      </div>
    </div>

    <div class="spec-box reveal" style="transition-delay:.18s">
      <div class="spec-box__lead">
        <div class="spec-box__icon"><i class="fa fa-archive"></i></div>
        <div>
          <div class="spec-box__label">In the box</div>
          <div class="spec-box__title">Everything you need, one carton.</div>
        </div>
      </div>
      <div class="spec-chips">
        <span class="spec-chip"><i class="fa fa-check"></i>Base &amp; central column</span>
        <span class="spec-chip"><i class="fa fa-check"></i>4 &times; arm assembly</span>
        <span class="spec-chip"><i class="fa fa-check"></i>4 &times; VESA plates</span>
        <span class="spec-chip"><i class="fa fa-check"></i>All mounting hardware</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Cable management clips</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Allen keys &amp; tools</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Printed guide</span>
        <span class="spec-chip"><i class="fa fa-check"></i><a href="/synergy-assembly.pdf" style="color:inherit;">Assembly PDF</a></span>
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
          <b>20&ndash;60 min assembly</b>
          <small>No drilling, no wall fixings</small>
        </span>
      </div>
      <div class="item">
        <i class="fa fa-cog"></i>
        <span>
          <b>All tools included</b>
          <small>Nothing to buy separately</small>
        </span>
      </div>
      <div class="item">
        <i class="fa fa-certificate"></i>
        <span>
          <b>Lifetime warranty</b>
          <small>On every steel part, forever</small>
        </span>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     MOUNTING / DESK INTERFACE
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="hero-grid">
      <div class="reveal">
        <div class="eyebrow">How it mounts</div>
        <h2>Clamps to your desk. <span class="display-em">No drilling, no damage.</span></h2>
        <p class="lead">
          The Quad Square attaches to the back edge of your desk with a heavy-duty steel C-clamp &mdash;
          the same clamp we&rsquo;ve shipped on every Synergy Stand since 2016. Fits desks from
          15&nbsp;mm thin veneer up to 55&nbsp;mm solid oak, and comes off as easily as it goes on.
        </p>
        <p style="color:var(--slate); margin-top:14px;">
          If your desk has a rear cable grommet, you can route the column through it for a cleaner look.
          The stand needs a desk edge to clamp to &mdash; it&rsquo;s not designed to stand freely on the floor
          or on a wall.
        </p>
        <div class="hero-mini" style="margin-top:22px;">
          <div class="item"><i class="fa fa-check" style="color:var(--brand);"></i><span>Clamp fits 15&ndash;55&nbsp;mm desk tops</span></div>
          <div class="item"><i class="fa fa-check" style="color:var(--brand);"></i><span>Grommet option on request</span></div>
          <div class="item"><i class="fa fa-check" style="color:var(--brand);"></i><span>Removable without marks</span></div>
        </div>
      </div>

      <div class="hero-visual reveal photo-todo" style="transition-delay:.1s">
        <img src="/shop/pc/catalog/arm-slot.jpg"
             alt="Close-up of the steel C-clamp securing the Synergy Stand column to the back edge of a desk"
             style="border-radius:var(--radius-xl); background:#F4F8FB; padding:24px;" />
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     ADJUSTABILITY DETAIL
     =================================================================== -->
<section class="s-tight" style="border-top:1px solid var(--line); border-bottom:1px solid var(--line); background:var(--sand);">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Designed for real-world use</h5>
        <h2>Six ways to get every screen <span class="display-em">exactly where you need it</span>.</h2>
        <p style="max-width:760px; margin-top:12px;">Adjustability isn&rsquo;t a single thing &mdash; it&rsquo;s the difference between a stand you fight and a stand you forget about. Every Synergy Stand mount gives you six independent degrees of freedom, then locks solid once positioned.</p>
      </div>
    </div>

    <div class="bench-panels">
      <div class="bench-panel reveal">
        <h4>Per-screen adjustment</h4>
        <span class="sub">Six degrees of freedom at every mount</span>
        <div style="margin-top:18px; display:grid; gap:14px;">
          <div style="display:flex; gap:14px; align-items:flex-start;">
            <i class="fa fa-arrows-v" style="color:var(--brand); font-size:18px; margin-top:3px; width:22px;"></i>
            <div><b style="color:var(--ink);">Height position</b><br><small style="color:var(--slate);">Mount arms at any height up the central column &mdash; great for stacking 2&times;2 at any screen size.</small></div>
          </div>
          <div style="display:flex; gap:14px; align-items:flex-start;">
            <i class="fa fa-refresh" style="color:var(--brand); font-size:18px; margin-top:3px; width:22px;"></i>
            <div><b style="color:var(--ink);">Arm hinge</b><br><small style="color:var(--slate);">Arms hinge from the centre so outer screens pull forward into a curve.</small></div>
          </div>
          <div style="display:flex; gap:14px; align-items:flex-start;">
            <i class="fa fa-arrows-h" style="color:var(--brand); font-size:18px; margin-top:3px; width:22px;"></i>
            <div><b style="color:var(--ink);">Horizontal slide</b><br><small style="color:var(--slate);">Screens slide along the arm to set spacing.</small></div>
          </div>
          <div style="display:flex; gap:14px; align-items:flex-start;">
            <i class="fa fa-compass" style="color:var(--brand); font-size:18px; margin-top:3px; width:22px;"></i>
            <div><b style="color:var(--ink);">Pivot</b><br><small style="color:var(--slate);">Each screen pivots left or right independently.</small></div>
          </div>
          <div style="display:flex; gap:14px; align-items:flex-start;">
            <i class="fa fa-sort" style="color:var(--brand); font-size:18px; margin-top:3px; width:22px;"></i>
            <div><b style="color:var(--ink);">Tilt</b><br><small style="color:var(--slate);">Wide range of up / down tilt on every screen.</small></div>
          </div>
          <div style="display:flex; gap:14px; align-items:flex-start;">
            <i class="fa fa-sliders" style="color:var(--brand); font-size:18px; margin-top:3px; width:22px;"></i>
            <div><b style="color:var(--ink);">30&nbsp;mm fine height adjust</b><br><small style="color:var(--slate);">Per-mount micro-adjust so top edges line up perfectly across the row.</small></div>
          </div>
        </div>
      </div>

      <div class="bench-panel reveal" style="transition-delay:.08s; display:flex; flex-direction:column;">
        <h4>Illustrated</h4>
        <span class="sub">The same mount, six degrees of freedom</span>
        <div style="flex:1; display:flex; align-items:center; justify-content:center; margin-top:20px;">
          <img src="/images/pages/ss-flexible.png" alt="Diagram of Synergy Stand adjustability &mdash; height, tilt, pivot, slide, hinge" style="max-width:100%; height:auto;">
        </div>
        <p class="bench-caption">Everything locks solid once positioned &mdash; this is a stand you set up once and forget, not one you fight with every Monday morning.</p>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     COMPATIBILITY / VESA / FIT CHECKER
     =================================================================== -->
<section class="s specs">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Will my monitors fit?</h5>
        <h2>If it&rsquo;s on our site, <span class="display-em">it fits this stand</span>.</h2>
        <p style="max-width:760px; margin-top:12px;">Every Synergy Stand is designed around the two VESA patterns used by 95&thinsp;% of modern monitors. If you&rsquo;re bringing your own screens, check the pattern on the back of your monitor or use the quick check below.</p>
      </div>
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

        <p style="margin-top:20px; color:var(--slate); line-height:1.55;">
          Verified compatible with every monitor we sell &mdash; and the vast majority of
          <b style="color:var(--ink);">Dell U-series</b>, <b style="color:var(--ink);">LG UltraGear</b>,
          <b style="color:var(--ink);">Samsung ViewFinity</b>, <b style="color:var(--ink);">AOC</b>,
          <b style="color:var(--ink);">BenQ</b> and <b style="color:var(--ink);">ASUS</b> displays up to 28&Prime; / 8&nbsp;kg.
        </p>
        <p style="color:var(--muted); font-size:13px; margin-top:10px;">
          Got an unusual monitor? <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">Ring 0330 223 66 55</a> and we&rsquo;ll check it for you.
        </p>
      </div>

      <div class="bench-panel reveal" style="transition-delay:.08s;">
        <div class="fit-check">
          <div>
            <h4 style="margin-bottom:4px;">Quick check</h4>
            <div class="sub">Your monitor vs this stand</div>
          </div>

          <div class="fit-field">
            <div class="fit-field__lbl">Screen size</div>
            <div class="fit-pills">
              <span class="fit-pill">24&Prime;</span>
              <span class="fit-pill is-on">27&Prime;</span>
              <span class="fit-pill">28&Prime;</span>
              <span class="fit-pill">29&Prime;+</span>
            </div>
          </div>

          <div class="fit-field">
            <div class="fit-field__lbl">Weight</div>
            <div class="fit-pills">
              <span class="fit-pill">&lt;&nbsp;4&nbsp;kg</span>
              <span class="fit-pill is-on">4&ndash;6&nbsp;kg</span>
              <span class="fit-pill">6&ndash;8&nbsp;kg</span>
              <span class="fit-pill">&gt;&nbsp;8&nbsp;kg</span>
            </div>
          </div>

          <div class="fit-field">
            <div class="fit-field__lbl">VESA</div>
            <div class="fit-pills">
              <span class="fit-pill">75 &times; 75</span>
              <span class="fit-pill is-on">100 &times; 100</span>
              <span class="fit-pill">200 &times; 100</span>
              <span class="fit-pill">Other</span>
            </div>
          </div>

          <div class="fit-verdict">
            <div class="tick" aria-hidden="true"><i class="fa fa-check"></i></div>
            <div>
              <b>27&Prime; / 5.2&nbsp;kg / 100&times;100 &mdash; fits perfectly.</b>
              <small>Under 8&nbsp;kg &middot; standard VESA &middot; within 28&Prime; limit</small>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     CABLE MANAGEMENT
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="hero-grid">
      <div class="hero-visual reveal photo-todo">
        <img src="/shop/pc/catalog/4s-rear-angle-med.jpg"
             alt="Rear view of the Quad Square Synergy Stand showing the central column with cable routing channel"
             style="border-radius:var(--radius-xl); background:#F4F8FB; padding:24px;" />
      </div>
      <div class="reveal" style="transition-delay:.08s">
        <div class="eyebrow">Cable management</div>
        <h2>Cables hidden. <span class="display-em">Not &ldquo;tucked&rdquo;.</span></h2>
        <p class="lead">
          The central column is hollow. Power, DisplayPort, HDMI and USB runs go in at the base
          and come out at each mount &mdash; out of sight, out of your cat&rsquo;s way, out of
          the way of vacuuming.
        </p>
        <p style="color:var(--slate); margin-top:14px;">
          Every stand ships with cable management clips for the lengths that run outside the column
          (mainly between the arm and the monitor). On a clean install you see the stand, you see the
          screens, and nothing else.
        </p>
        <div class="hero-mini" style="margin-top:20px;">
          <div class="item"><i class="fa fa-random" style="color:var(--brand);"></i><span>Internal column routing</span></div>
          <div class="item"><i class="fa fa-link" style="color:var(--brand);"></i><span>Clips for external runs</span></div>
          <div class="item"><i class="fa fa-plug" style="color:var(--brand);"></i><span>Access from the base</span></div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     FOOTPRINT & DIMENSIONS
     =================================================================== -->
<section class="s specs">
  <div class="container">
    <div class="hero-grid">
      <div class="reveal">
        <div class="eyebrow">Footprint &amp; dimensions</div>
        <h2>The space of two. <span class="display-em">The capacity of four.</span></h2>
        <p class="lead">
          The Quad Square takes roughly the same desk space as a single 27&Prime; monitor on its stock stand &mdash;
          and gives you four screens in its place. A single central column keeps the desk surface below
          the screens clear for keyboard, notebook and coffee.
        </p>

        <div class="dim-stats">
          <div class="dim-stat"><b>470<span class="u">mm</span></b><small>Base width</small></div>
          <div class="dim-stat"><b>210<span class="u">mm</span></b><small>Base depth</small></div>
          <div class="dim-stat"><b>780<span class="u">mm</span></b><small>Column height</small></div>
          <div class="dim-stat"><b>450<span class="u">mm</span></b><small>Arm reach (each)</small></div>
          <div class="dim-stat"><b>8<span class="u">kg</span></b><small>Max per mount</small></div>
          <div class="dim-stat"><b>32<span class="u">kg</span></b><small>Max total load</small></div>
        </div>

        <p style="margin-top:22px; font-size:13px; color:var(--muted);">
          <a href="/synergy-assembly.pdf" style="color:var(--brand); font-weight:500;"><i class="fa fa-file-pdf-o"></i> Download the full assembly &amp; dimension PDF</a>
        </p>
      </div>

      <div class="reveal" style="transition-delay:.08s">
        <div class="dim-card">
          <div class="desk-silhouette" aria-hidden="true">
            <svg viewBox="0 0 520 320" xmlns="http://www.w3.org/2000/svg">
              <rect x="20" y="240" width="480" height="14" fill="#CED6E0" rx="2"/>
              <rect x="40" y="254" width="10" height="54" fill="#CED6E0"/>
              <rect x="470" y="254" width="10" height="54" fill="#CED6E0"/>
              <rect x="180" y="226" width="160" height="14" fill="rgba(15,110,168,.25)" stroke="#0F6EA8" stroke-width="1.5" rx="2"/>
              <rect x="252" y="40" width="16" height="186" fill="#1B2A3E"/>
              <rect x="130" y="60" width="120" height="80" fill="#1B2A3E" rx="3"/>
              <rect x="270" y="60" width="120" height="80" fill="#1B2A3E" rx="3"/>
              <rect x="130" y="150" width="120" height="80" fill="#1B2A3E" rx="3"/>
              <rect x="270" y="150" width="120" height="80" fill="#1B2A3E" rx="3"/>
              <rect x="134" y="64" width="112" height="72" fill="#0F6EA8" opacity=".28" rx="2"/>
              <rect x="274" y="64" width="112" height="72" fill="#0F6EA8" opacity=".28" rx="2"/>
              <rect x="134" y="154" width="112" height="72" fill="#0F6EA8" opacity=".28" rx="2"/>
              <rect x="274" y="154" width="112" height="72" fill="#0F6EA8" opacity=".28" rx="2"/>
              <line x1="180" y1="285" x2="340" y2="285" stroke="#7A8699" stroke-width="1"/>
              <line x1="180" y1="280" x2="180" y2="290" stroke="#7A8699" stroke-width="1"/>
              <line x1="340" y1="280" x2="340" y2="290" stroke="#7A8699" stroke-width="1"/>
              <text x="260" y="305" text-anchor="middle" font-family="JetBrains Mono, monospace" font-size="11" fill="#455065" letter-spacing="1">470 mm</text>
            </svg>
          </div>
          <img src="/images/pages/dim-4s.jpg"
               alt="Engineering side- and front-elevation drawings of the Quad Square Synergy Stand with dimensions in millimetres"
               style="max-width:100%; margin-top:4px;" />
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     MODULAR SCALE PATH
     =================================================================== -->
<section class="bundle">
  <div class="container">
    <div class="bundle-grid">
      <div class="reveal">
        <h5>Upgrade path</h5>
        <h2>Start at four. Scale to six or eight. <em>Don&rsquo;t buy twice.</em></h2>
        <p>The Quad Square uses the same base assembly and central column as the six-monitor and eight-monitor Synergy Stands. When your next two screens arrive next year, you buy arms and mounts &mdash; not a whole new stand.</p>
        <p style="color:#C7D2DF; margin-top:14px;">We&rsquo;ve heard <em>&ldquo;wouldn&rsquo;t it be easier if I had just two more screens&rdquo;</em> so many times we designed the whole system around it. Parts for every configuration are always in stock.</p>
        <div class="bundle-pills" style="margin-top:20px;">
          <span class="bundle-pill"><i class="fa fa-check"></i>Same base, four configurations</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Arms ship in 2 working days</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>No re-assembly of the base</span>
        </div>
        <div style="display:flex; gap:12px; flex-wrap:wrap; margin-top:20px;">
          <a href="/stands/" class="btn btn-accent btn-lg">See the full range <i class="fa fa-arrow-right"></i></a>
        </div>
      </div>

      <div class="reveal" style="transition-delay:.1s">
        <div class="save-card">
          <span class="save-tag">Scale path</span>
          <div class="kicker">The same base, three configurations</div>
          <div class="breakdown" style="margin-top:6px; gap:14px;">
            <div class="r" style="align-items:center;">
              <span style="display:flex; align-items:center; gap:12px;">
                <img src="/shop/pc/catalog/4s-front-angle-thm.jpg" alt="Quad Square Synergy Stand" style="width:56px; height:56px; object-fit:contain; background:#fff; border-radius:4px;">
                <span><b style="color:var(--ink);">You start here</b><br><small style="color:var(--muted);">4 screens, 2&times;2</small></span>
              </span>
              <b>This stand</b>
            </div>
            <div class="r" style="align-items:center;">
              <span style="display:flex; align-items:center; gap:12px;">
                <img src="/shop/pc/catalog/6r-front-angle-thm.jpg" alt="Six Monitor Synergy Stand" style="width:56px; height:56px; object-fit:contain; background:#fff; border-radius:4px;">
                <span><b style="color:var(--ink);">Scale up</b><br><small style="color:var(--muted);">6 screens, 3 across</small></span>
              </span>
              <b>+ 2 arms</b>
            </div>
            <div class="r" style="align-items:center;">
              <span style="display:flex; align-items:center; gap:12px;">
                <img src="/shop/pc/catalog/8r-front-angle-thm.jpg" alt="Eight Monitor Synergy Stand" style="width:56px; height:56px; object-fit:contain; background:#fff; border-radius:4px;">
                <span><b style="color:var(--ink);">Go big</b><br><small style="color:var(--muted);">8 screens, 2-over-2 quad</small></span>
              </span>
              <b>+ 4 arms</b>
            </div>
            <div class="r total"><span>Same base assembly throughout</span><b>&mdash;</b></div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     SIBLING STAND CROSS-LINKS
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Considering a different layout?</h5>
        <h2>Four screens, <span class="display-em">four ways</span>.</h2>
        <p style="max-width:760px; margin-top:12px;">Quad Square is our most popular four-screen configuration, but it isn&rsquo;t right for every desk or every workflow. Here&rsquo;s how the closest alternatives compare.</p>
      </div>
    </div>

    <div class="bundle-cards" style="margin-top:8px;">
      <a href="/shop/pc/Triple-Monitor-Stand-p3h.htm" class="bundle-card reveal">
        <div class="bundle-card__media">
          <img src="/shop/pc/catalog/3h-front-angle-med.jpg" alt="Triple Monitor Synergy Stand in horizontal layout">
        </div>
        <div class="bundle-card__body">
          <div class="bundle-card__eyebrow">3-Screen &middot; Horizontal</div>
          <h4 class="bundle-card__title">Triple Horizontal</h4>
          <p style="font-size:13px; color:var(--slate); margin:0 0 12px; line-height:1.5;">Three screens side-by-side. Most popular for single-market traders and writers.</p>
          <div class="bundle-card__price">
            <span class="bundle-card__from">From</span>
            <span class="bundle-card__amount">&pound;175</span>
          </div>
          <span class="btn btn-primary bundle-card__cta">View stand <i class="fa fa-arrow-right"></i></span>
        </div>
      </a>
      <a href="/shop/pc/Quad-Monitor-Stand-p4p.htm" class="bundle-card reveal" style="transition-delay:.06s">
        <div class="bundle-card__media">
          <img src="/shop/pc/catalog/4p-front-angle-med.jpg" alt="Quad Pyramid Synergy Stand with one monitor over three">
        </div>
        <div class="bundle-card__body">
          <div class="bundle-card__eyebrow">4-Screen &middot; Pyramid</div>
          <h4 class="bundle-card__title">Quad Pyramid</h4>
          <p style="font-size:13px; color:var(--slate); margin:0 0 12px; line-height:1.5;">One screen sat above a row of three. Good for news/email at the top, charts below.</p>
          <div class="bundle-card__price">
            <span class="bundle-card__from">From</span>
            <span class="bundle-card__amount">&pound;225</span>
          </div>
          <span class="btn btn-primary bundle-card__cta">View stand <i class="fa fa-arrow-right"></i></span>
        </div>
      </a>
      <a href="/shop/pc/Quad-Monitor-Stand-p4h.htm" class="bundle-card reveal" style="transition-delay:.12s">
        <div class="bundle-card__media">
          <img src="/shop/pc/catalog/4h-front-angle-med.jpg" alt="Quad Horizontal Synergy Stand with four screens in a single row">
        </div>
        <div class="bundle-card__body">
          <div class="bundle-card__eyebrow">4-Screen &middot; Horizontal</div>
          <h4 class="bundle-card__title">Quad Horizontal</h4>
          <p style="font-size:13px; color:var(--slate); margin:0 0 12px; line-height:1.5;">Four screens in a single wide row. Needs a deep desk (~1.6&nbsp;m wide).</p>
          <div class="bundle-card__price">
            <span class="bundle-card__from">From</span>
            <span class="bundle-card__amount">&pound;235</span>
          </div>
          <span class="btn btn-primary bundle-card__cta">View stand <i class="fa fa-arrow-right"></i></span>
        </div>
      </a>
      <a href="/shop/pc/Six-Monitor-Stand-p6r.htm" class="bundle-card reveal" style="transition-delay:.18s">
        <div class="bundle-card__media">
          <img src="/shop/pc/catalog/6r-front-angle-med.jpg" alt="Six Monitor Synergy Stand with two rows of three screens">
        </div>
        <div class="bundle-card__body">
          <div class="bundle-card__eyebrow">6-Screen &middot; 2 rows of 3</div>
          <h4 class="bundle-card__title">Six Monitor Stand</h4>
          <p style="font-size:13px; color:var(--slate); margin:0 0 12px; line-height:1.5;">Quad Square&rsquo;s bigger sibling. Same base assembly, adds two extra screens.</p>
          <div class="bundle-card__price">
            <span class="bundle-card__from">From</span>
            <span class="bundle-card__amount">&pound;275</span>
          </div>
          <span class="btn btn-primary bundle-card__cta">View stand <i class="fa fa-arrow-right"></i></span>
        </div>
      </a>
    </div>
  </div>
</section>

<!-- ===================================================================
     REVIEWS (placeholder copy until Trustpilot import lands)
     =================================================================== -->
<section class="s reviews" id="reviews">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Stand reviews</h5>
        <h2>Why <span class="display-em">Quad Square</span> owners pick it.</h2>
        <p>All reviews are voluntary &mdash; we don&rsquo;t ask for them. Placeholder copy below pending the Trustpilot import.</p>
      </div>
      <div class="tp-summary">
        <span class="tp-stars"><span></span><span></span><span></span><span></span><span></span></span>
        <span><b>4.9</b> <small>&middot; based on 90+ reviews</small></span>
        <a href="#" class="link" style="margin-left:10px;">See all on Trustpilot <i class="fa fa-external-link"></i></a>
      </div>
    </div>

    <div class="reviews-grid">
      <div class="review reveal">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <span class="platform">4-screen trader &middot; 27&Prime; Dells</span>
        <h4>The square layout saved my desk</h4>
        <p>I was running triple horizontal but lost too much desk depth. Swapped to the Quad Square, got a fourth screen, and now there&rsquo;s room for a notebook and coffee in front of the keyboard. Arms lock solid &mdash; none of the sag I had with the VIVO stand before it.</p>
        <div class="meta">
          <div class="ava">AT</div>
          <div class="who">Alex T., London</div>
          <div class="when">03&thinsp;/&thinsp;2026</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.08s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <span class="platform">Design studio &middot; 4-up</span>
        <h4>Fits four 27&Prime; LGs with room to <em>breathe</em></h4>
        <p>Our studio runs four 27&Prime; LGs for Figma, reference, video reviews and Slack. The competitor stand we trialled couldn&rsquo;t clear all four panels without them overlapping. Synergy&rsquo;s arms have just enough reach to give each screen its own space. Solid UK-made kit.</p>
        <div class="meta">
          <div class="ava">SK</div>
          <div class="who">Sarah K., Brighton</div>
          <div class="when">02&thinsp;/&thinsp;2026</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.16s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <span class="platform">Ops desk &middot; added 2 arms</span>
        <h4>Scaled from 4 to 6 in fifteen minutes</h4>
        <p>Bought the Quad Square 18 months ago. When our ops team grew we just ordered two more arms and slotted them onto the same central column. No re-building the base, no new stand. Twenty minutes start-to-finish. Exactly what they promised.</p>
        <div class="meta">
          <div class="ava">MR</div>
          <div class="who">Mark R., Manchester</div>
          <div class="when">01&thinsp;/&thinsp;2026</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     FAQ
     =================================================================== -->
<section class="s depth" id="faq">
  <div class="container-narrow">
    <div class="section-head reveal" style="display:block; margin-bottom:38px;">
      <h5>Quad Square questions</h5>
      <h2>The six questions we&rsquo;re asked <span class="display-em">most often</span>.</h2>
      <p style="margin-top:12px;">Specific to this stand &mdash; not generic stand-shop answers. Got a question not listed? <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">Call us on 0330 223 66 55</a>.</p>
    </div>

    <div class="faq-list reveal">
      <details class="faq-item" open>
        <summary>Will my 27&Prime; or 28&Prime; curved monitors fit?</summary>
        <div class="faq-body">
          <p><strong>Yes &mdash; up to 28&Prime; flat or curved is supported on every mount.</strong> The arms have enough reach and the pivot has enough angle that even 28&Prime; curved panels can be angled into a gentle curve across the top row and the bottom row.</p>
          <p>The limit you&rsquo;ll hit first is usually weight, not size: 8&nbsp;kg per mount is generous for any 24&Prime;&ndash;27&Prime; panel but a handful of 28&Prime; curved models approach that. If in doubt, tell us the make &amp; model when you order and we&rsquo;ll confirm before we ship.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Can I upgrade to five, six or eight screens later?</summary>
        <div class="faq-body">
          <p>Yes. The Quad Square uses the exact same central column and base as our five-, six- and eight-monitor stands. When you&rsquo;re ready to grow, you buy arms and VESA plates &mdash; not a whole new stand.</p>
          <p>Typical upgrade prices: +&pound;70 for a pair of arms (takes you to six screens); +&pound;140 for four arms (takes you to eight). We keep these parts in stock year-round. See the <a href="/stands/" style="color:var(--brand); font-weight:500;">full range</a> for exactly what each configuration looks like.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Desk clamp or grommet mount &mdash; which should I pick?</summary>
        <div class="faq-body">
          <p>The Quad Square ships with a desk clamp as standard. It fits desks with a back edge between 15&nbsp;mm and 55&nbsp;mm thick &mdash; which is the vast majority of office and home desks.</p>
          <p>If your desk has a rear cable grommet and you want the column to pass through it, ask for the grommet mount option at no extra cost. If your desk is <em>against</em> a wall and the clamp won&rsquo;t fit, a freestanding base plate is available as a paid option (+&pound;35) &mdash; mention it when you order.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How does the cable management work?</summary>
        <div class="faq-body">
          <p>The central column is hollow. Power, DisplayPort, HDMI and USB cables go in at the base (or at the top, either works), run inside the column, and exit at each arm through a slot near the VESA plate.</p>
          <p>The short runs between the column and each monitor are held in place with included clips. In practice, on a clean install, all you see is the stand and the screens &mdash; no dangling cables, no zip ties, no mess.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How long does assembly take and what tools do I need?</summary>
        <div class="faq-body">
          <p>Most customers build a Quad Square on their own in 30&ndash;45 minutes. Two people can do it in 20. Nothing to drill, nothing to fix to the wall. Every tool you need (two Allen keys and a small spanner) ships in the box.</p>
          <p>If you&rsquo;d rather watch than read the instructions, the <a href="/pages/synergy-stand-assembly-videos/" style="color:var(--brand); font-weight:500;">Synergy Stand assembly videos</a> walk through it step-by-step. You can also download the <a href="/synergy-assembly.pdf" style="color:var(--brand); font-weight:500;">printable assembly PDF</a>.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>What&rsquo;s the warranty if something goes wrong?</summary>
        <div class="faq-body">
          <p>Lifetime on every steel part. If an arm, a clamp, a VESA plate, or the column itself ever fails or bends, we replace it free, forever. Fixings (nuts, bolts, plastic clips) are covered for five years.</p>
          <p>In 9 years of selling Synergy Stands we&rsquo;ve seen a total of 11 warranty claims, mostly minor (replacement clips, re-shipped VESA plates). It&rsquo;s a simple system of steel parts &mdash; not much goes wrong.</p>
        </div>
      </details>
    </div>

    <div class="darren-inline reveal">
      <div class="avatar"><i class="fa fa-user"></i></div>
      <div>
        <h4>Question not on the list?</h4>
        <p>Seventeen years of pre-sale conversations means we&rsquo;ve heard most things. Phone or email Darren &mdash; he&rsquo;ll give you a straight answer, or tell you honestly if the Quad Square isn&rsquo;t the right stand for you.</p>
      </div>
      <div>
        <a href="tel:03302236655" class="btn btn-primary"><i class="fa fa-phone"></i>0330 223 66 55</a>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     BUNDLE UPSELL
     =================================================================== -->
<section class="bundle">
  <div class="container">
    <div class="bundle-grid">
      <div class="reveal">
        <h5>Complete your setup</h5>
        <h2>Add four screens and a PC. <em>Save &pound;200&thinsp;+</em>.</h2>
        <p>The Quad Square is designed to work with our monitor range and our trading computers &mdash; because we built them together. Pair it with four screens and a Trader PC as a bundle and the free cables, free UK delivery and bundle discount more than cover themselves.</p>
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
          <span class="save-tag">Example &middot; 4-screen trader bundle</span>
          <div class="kicker">Typical saving vs buying separately</div>
          <div class="big"><small>&pound;</small>210</div>
          <div class="sub">on Quad Square + four 27&Prime; screens + Trader PC.</div>
          <div class="breakdown">
            <div class="r"><span>4&thinsp;&times;&thinsp;3&nbsp;m video cables</span><b>&pound;60</b></div>
            <div class="r"><span>WiFi, BT &amp; speakers</span><b>&pound;60</b></div>
            <div class="r"><span>UK mainland delivery</span><b>&pound;20</b></div>
            <div class="r"><span>Bundle discount</span><b>&pound;70</b></div>
            <div class="r total"><span>Total savings</span><b>&minus;&thinsp;&pound;210</b></div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     DARREN CTA
     =================================================================== -->
<section class="darren" id="darren">
  <div class="container">
    <div class="darren-grid">
      <div class="darren-photo reveal">
        <img src="/images/pages/darren.jpg" alt="Darren Atkinson, founder of Multiple Monitors Ltd">
      </div>
      <div class="reveal" style="transition-delay:.08s">
        <h5>Still deciding on a stand?</h5>
        <h2>Talk to <em>Darren</em> &mdash; the founder, not a call centre.</h2>
        <p>Seventeen years of speccing these stands means most of our customers&rsquo; questions have pretty direct answers. &ldquo;Will my screens fit?&rdquo; &ldquo;Which configuration for my desk?&rdquo; &ldquo;Can I add screens later?&rdquo; Fifteen minutes on the phone is usually enough to figure out what you need.</p>
        <div class="darren-ctas">
          <a href="tel:03302236655" class="btn btn-primary btn-lg"><i class="fa fa-phone"></i>0330 223 66 55</a>
          <a href="#" class="btn btn-ghost btn-lg"><i class="fa fa-calendar"></i>Book a 15-min call</a>
        </div>
        <div class="darren-sig">&mdash; Darren Atkinson, founder, Multiple Monitors Ltd</div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     STICKY ADD-TO-BASKET CTA
     =================================================================== -->
<div class="sticky-cta" id="stickyCta">
  <div class="txt">
    <strong><%= Server.HTMLEncode(mmName) %> &middot; &pound;<%= mmBasePriceExDisp %> + VAT</strong>
    <span>Order by <%= daFunDelCutOff() %> &middot; delivered <%= daFunDelDateReturn(0,0) %></span>
  </div>
  <% If mmIsArrayBuild Then %>
    <a href="/display-systems-2/?sid=<%= mmIdProduct %>&amp;mid=<%= Server.URLEncode(mmCtxMid) %>" class="btn btn-primary btn-sm">Add to array <i class="fa fa-arrow-right"></i></a>
  <% ElseIf mmIsBundleBuild Then %>
    <a href="/bundles-2/?sid=<%= mmIdProduct %>&amp;mid=<%= Server.URLEncode(mmCtxMid) %>&amp;cid=<%= Server.URLEncode(mmCtxCid) %>" class="btn btn-primary btn-sm">Add to bundle <i class="fa fa-arrow-right"></i></a>
  <% Else %>
    <a href="#pdQty" class="btn btn-primary btn-sm" onclick="document.getElementById('pdQty').focus(); return true;">Add to basket <i class="fa fa-arrow-right"></i></a>
  <% End If %>
</div>

</div><!-- /.mm-site -->

<!-- ===================================================================
     PAGE-SPECIFIC JS
     Reveal-on-scroll is already wired in footer_wrapper.asp.
     =================================================================== -->
<script>
(function(){
  // -- Sticky CTA: visible after the hero, hidden before the footer --
  var sticky    = document.getElementById('stickyCta');
  var hero      = document.querySelector('.pd-hero');
  var footerEl  = document.querySelector('footer');
  if (sticky && hero && footerEl) {
    function onScroll(){
      var y = window.scrollY || window.pageYOffset;
      var heroBottom     = hero.getBoundingClientRect().bottom + y;
      var footerTop      = footerEl.getBoundingClientRect().top + y;
      var viewportBottom = y + window.innerHeight;
      if (y > heroBottom + 120 && viewportBottom < footerTop) sticky.classList.add('visible');
      else                                                    sticky.classList.remove('visible');
    }
    window.addEventListener('scroll', onScroll, { passive:true });
    onScroll();
  }

  // -- Gallery thumbs swap main image; toggle play overlay on video thumb --
  var thumbs = document.querySelectorAll('.pd-thumb[data-img]');
  var main   = document.getElementById('pdMainImg');
  var videoOverlay = document.getElementById('pdVideoOverlay');
  if (main) {
    thumbs.forEach(function(t){
      t.addEventListener('click', function(){
        document.querySelectorAll('.pd-thumb.is-active').forEach(function(x){ x.classList.remove('is-active'); });
        t.classList.add('is-active');
        main.src = t.dataset.img;
        var imgEl = t.querySelector('img');
        if (imgEl && imgEl.alt) main.alt = imgEl.alt;
        if (videoOverlay) {
          videoOverlay.style.display = t.dataset.video === '1' ? '' : 'none';
        }
      });
    });
  }

  // -- Fit-check pills: mutually exclusive within each .fit-field --
  document.querySelectorAll('.fit-field').forEach(function(field){
    var pills = field.querySelectorAll('.fit-pill');
    pills.forEach(function(p){
      p.addEventListener('click', function(){
        pills.forEach(function(x){ x.classList.remove('is-on'); });
        p.classList.add('is-on');
      });
    });
  });
})();
</script>

<!--#include file="footer_wrapper.asp"-->
