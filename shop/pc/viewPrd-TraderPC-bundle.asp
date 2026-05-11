<%
' ============================================================
' viewPrd-TraderPC-bundle-v2.asp
' 2026 redesign - Trader PC bundle end-page (idProduct 333).
'
' Receives sid/mid/cid querystrings from the bundle builder and
' renders the bundle-first product page: stand + screens + PC
' composite hero, three-pick audit section, compatibility band,
' PC configurator (live DB prices, same as viewPrd-TraderPC-v2.asp),
' bundle breakdown sidebar, full PC spec + bundle build summary.
'
' Posts to /shop/pc/instPrd.asp with the PC's idOption1..N plus
' the 3-item bundle hidden inputs (idproduct2/3, QtyM*, pCnt=3)
' emitted by inc_bundleContext.asp's mmEmitBundleHiddenInputs().
'
' See /bundle-pages-redesign-plan.md for the architecture.
' ============================================================
%>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "viewPrd-TraderPC-bundle-v2.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<%
' ------------------------------------------------------------
' Page constants - must be set BEFORE inc_bundleContext.asp.
' ------------------------------------------------------------
Const MM_PRODUCT_ID = 333
Const MM_VAT_RATE   = 1.2

' ------------------------------------------------------------
' Legacy querystring URL (?sid=&mid=&cid=) -> 301 to slug URL.
' Old bookmarks and search-engine hits land here via the
' "Trader PC Bundle" rewrite rule in web.config; we look up
' the stand + monitor slugs by id and bounce to the canonical
' /products/trader-pc/<stand>/<monitor>/ form before
' inc_bundleContext.asp runs (it only knows the slug shape).
' cid is stripped silently - the URL path already implies PC.
' ------------------------------------------------------------
If Request.QueryString("sid") <> "" Then
  Dim mmLegacySid, mmLegacyMid, mmLegacyStandSlug, mmLegacyMonSlug
  mmLegacySid = Request.QueryString("sid") & ""
  mmLegacyMid = Request.QueryString("mid") & ""
  If Not IsNumeric(mmLegacySid) Or Not IsNumeric(mmLegacyMid) Or mmLegacyMid = "" Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", "/bundles/"
    Response.End
  End If
  mmLegacyStandSlug = mmLookupPcUrlBundle(CLng(mmLegacySid))
  mmLegacyMonSlug   = mmLookupPcUrlBundle(CLng(mmLegacyMid))
  If mmLegacyStandSlug = "" Or mmLegacyMonSlug = "" Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", "/bundles/"
    Response.End
  End If
  Response.Status = "301 Moved Permanently"
  Response.AddHeader "Location", "/products/trader-pc/" & mmLegacyStandSlug & "/" & mmLegacyMonSlug & "/"
  Response.End
End If

Function mmLookupPcUrlBundle(ByVal idp)
  Dim sql, rs
  mmLookupPcUrlBundle = ""
  sql = "SELECT pcUrlBundle FROM products " & _
        "WHERE idProduct = " & CLng(idp) & _
        "  AND active = -1 AND removed = 0"
  Set rs = connTemp.Execute(sql)
  If Not rs.EOF Then mmLookupPcUrlBundle = rs("pcUrlBundle") & ""
  rs.Close : Set rs = Nothing
End Function
%>
<!--#include file="inc_bundleContext.asp"-->
<%
' ------------------------------------------------------------
' 1. Trader PC base row
' ------------------------------------------------------------
Dim mmName, mmSku, mmBasePriceInc, mmImageUrl, mmSmallImageUrl
mmName = "Trader PC" : mmSku = "" : mmBasePriceInc = 0 : mmImageUrl = "" : mmSmallImageUrl = ""

Dim mmPrdSql, mmPrdRs
mmPrdSql = "SELECT description, sku, price, imageUrl, smallImageUrl " & _
           "FROM products " & _
           "WHERE idProduct = " & MM_PRODUCT_ID & _
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

If Not mmPrdRs.EOF Then
  mmName          = mmPrdRs("description") & ""
  mmSku           = mmPrdRs("sku") & ""
  mmBasePriceInc  = CDbl(mmPrdRs("price"))
  mmImageUrl      = mmPrdRs("imageUrl") & ""
  mmSmallImageUrl = mmPrdRs("smallImageUrl") & ""
End If
mmPrdRs.Close : Set mmPrdRs = Nothing

Dim mmBasePriceEx
mmBasePriceEx = mmBasePriceInc / MM_VAT_RATE

' ------------------------------------------------------------
' 2. Option groups assigned to this product
' ------------------------------------------------------------
Dim mmOgSql, mmOgRs, mmOgRows, mmOgCount
mmOgCount = 0

mmOgSql = "SELECT DISTINCT og.idOptionGroup, og.OptionGroupDesc, " & _
          "       po.pcProdOpt_Required, po.pcProdOpt_Order " & _
          "FROM pcProductsOptions po " & _
          "INNER JOIN optionsGroups og ON og.idOptionGroup = po.idOptionGroup " & _
          "INNER JOIN options_optionsGroups oog ON oog.idOptionGroup = og.idOptionGroup " & _
          "                                   AND oog.idProduct = po.idProduct " & _
          "WHERE po.idProduct = " & MM_PRODUCT_ID & " " & _
          "ORDER BY po.pcProdOpt_Order, og.OptionGroupDesc"

Set mmOgRs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
mmOgRs.Open mmOgSql, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If err.number <> 0 Then
  On Error Goto 0
  call LogErrorToDatabase()
  Set mmOgRs = Nothing
  call closeDB()
  Response.Redirect "techErr.asp?err=" & pcStrCustRefID
End If
On Error Goto 0

If Not mmOgRs.EOF Then
  mmOgRows  = mmOgRs.GetRows()
  mmOgCount = UBound(mmOgRows, 2) + 1
End If
mmOgRs.Close : Set mmOgRs = Nothing

Dim mmMachineName : mmMachineName = mmName
mmSetBundleSeo mmName
%>
<!--#include file="inc_traderPcConfigurator.asp"-->
<%
' Auto-pick a GPU that supports the bundle's monitor count. The
' Trader PC's default GPU drives 4 screens, so 5/6/8-screen bundles
' need an upgrade picked before mmRenderOptionGroup runs. Returns 0
' for 1-4 monitors, leaving the default option selected.
mmGpuPreselectIdoptoptgrp = mmTraderPcGpuIdForMonCount(mmBunMonCount)

Dim mmBasePriceExDisp, mmBasePriceIncDisp
mmBasePriceExDisp  = mmFormatMoney0(mmBasePriceEx)
mmBasePriceIncDisp = mmFormatMoney(mmBasePriceInc)

Dim mmMainImgSrc
If mmImageUrl <> "" Then
  mmMainImgSrc = "/shop/pc/catalog/" & mmImageUrl
ElseIf mmSmallImageUrl <> "" Then
  mmMainImgSrc = "/shop/pc/catalog/" & mmSmallImageUrl
Else
  mmMainImgSrc = "/shop/pc/catalog/no_image.gif"
End If

' Bundle totals used in opening HTML + as JS constants
Dim mmBunSubtotalEx, mmBunTotalEx, mmBunTotalInc
mmBunSubtotalEx = mmBunStandPriceEx + mmBunMonSubtotalEx + mmBasePriceEx
mmBunTotalEx    = mmBunSubtotalEx - mmBunDiscount
mmBunTotalInc   = mmBunTotalEx * MM_VAT_RATE

' "Change" URLs back to the builder with current picks preserved.
Dim mmBunChangeBase
mmBunChangeBase = "/bundles/?sid=" & mmBunSid & "&mid=" & mmBunMid & "&cid=" & MM_PRODUCT_ID
%>
<!--#include file="header_wrapper.asp"-->

<div class="mm-site">

<!-- ===================================================================
     BREADCRUMB
     =================================================================== -->
<nav class="breadcrumb" aria-label="Breadcrumb">
  <div class="container inner">
    <a href="/">Home</a>
    <span class="sep">/</span>
    <a href="/bundles/">Bundles</a>
    <span class="sep">/</span>
    <span class="current">Your bundle</span>
  </div>
</nav>

<form method="post" action="/shop/pc/instPrd.asp" id="cfgForm">
  <input type="hidden" name="idproduct1"                value="<%= MM_PRODUCT_ID %>">
  <input type="hidden" name="QtyM<%= MM_PRODUCT_ID %>"  value="1">
  <input type="hidden" name="OptionGroupCount"          value="<%= mmOgCount %>">
  <% mmEmitBundleHiddenInputs() %>

<!-- ===================================================================
     BUNDLE HERO - stand+screens composite beside PC case
     =================================================================== -->
<section class="bp-hero">
  <div class="container">
    <div class="bp-hero-grid">

      <!-- Composite gallery -->
      <div class="bp-gallery reveal">
        <div class="bp-gallery__main">
          <span class="bp-gallery__chip">
            <span class="dot"></span><span class="acc">BUNDLE</span>TESTED&nbsp;TOGETHER
          </span>
          <img class="bp-gallery__array"
               src="<%= mmBunStandImgSrc %>"
               alt="<%= Server.HTMLEncode(mmBunStandDispName) %>" />
          <img class="bp-gallery__pc"
               src="<%= mmMainImgSrc %>"
               alt="<%= Server.HTMLEncode(mmName) %>" />
          <span class="bp-gallery__tag">
            <i class="fa fa-cube"></i>Ships in one delivery
          </span>
        </div>
        <div class="bp-gallery__thumbs">
          <div class="bp-thumb is-active">
            <img src="<%= mmBunStandImgSrc %>" alt="Full bundle composite" />
            <span class="bp-thumb__lbl">Bundle</span>
          </div>
          <div class="bp-thumb">
            <img src="<%= mmMainImgSrc %>" alt="<%= Server.HTMLEncode(mmName) %>" />
            <span class="bp-thumb__lbl">PC</span>
          </div>
          <div class="bp-thumb">
            <img src="<%= mmBunStandImgSrc %>" alt="<%= Server.HTMLEncode(mmBunStandDispName) %>" />
            <span class="bp-thumb__lbl">Stand</span>
          </div>
          <div class="bp-thumb">
            <img src="<%= mmBunMonImgSrc %>" alt="<%= Server.HTMLEncode(mmBunMonDispName) %>" />
            <span class="bp-thumb__lbl">Screens</span>
          </div>
          <div class="bp-thumb placeholder">
            <i class="fa fa-play-circle-o"></i>
            <span>60-sec<br>unbox</span>
          </div>
        </div>
      </div>

      <!-- Buybox -->
      <aside class="bp-buybox reveal" style="transition-delay:.08s">
        <span class="bp-ribbon">
          <span class="ico"><i class="fa fa-check"></i></span>
          Configured <b>&middot; ready to order</b>
        </span>
        <div class="eyebrow">Your bundle &middot; <%= mmBunMonCount %>-screen trader setup</div>
        <h1>Your <em>Trader&nbsp;Bundle</em></h1>
        <p class="pitch">
          <%= mmBunMonCount %>&times; <%= Server.HTMLEncode(mmBunMonDispName) %> on the <%= Server.HTMLEncode(mmBunStandDispName) %> stand,
          plus a UK-built <%= Server.HTMLEncode(mmName) %> &mdash; tested together, shipped together,
          with every cable you need. Customise the PC below; the rest of the bundle is already sorted.
        </p>

        <div class="bp-price">
          <div>
            <div class="bp-price__from">Bundle total from</div>
            <div class="bp-price__num"><span class="sym">&pound;</span><span data-hero-ex><%= mmFormatMoney0(mmBunTotalEx) %></span></div>
          </div>
          <div class="bp-price__meta">
            <span><b>&pound;<span data-hero-inc><%= mmFormatMoney(mmBunTotalInc) %></span></b> inc VAT</span>
            <span style="text-transform:none; font-family:'Geist', sans-serif; letter-spacing:0;">
              Stand &pound;<%= mmFormatMoney0(mmBunStandPriceEx) %>
              &middot; <%= mmBunMonCount %> screens &pound;<%= mmFormatMoney0(mmBunMonSubtotalEx) %>
              &middot; <%= Server.HTMLEncode(mmName) %> <span data-hero-pc>&pound;<%= mmBasePriceExDisp %></span>
            </span>
            <span class="bun"><i class="fa fa-gift"></i>Free UK delivery included</span>
          </div>
        </div>

        <% If mmBunDiscount > 0 Then %>
        <div class="bp-savings">
          <p class="bp-savings__line">
            You&rsquo;re saving <b>&pound;<span data-hero-saved><%= mmBunDiscount + 80 %></span></b>
            vs piecing this together separately.
          </p>
          <div class="bp-savings__pills">
            <span>Free Wi-Fi / BT card <b>&pound;40</b></span>
            <span>Free speakers <b>&pound;20</b></span>
            <span>Free premium cables <b>&pound;<%= mmCableCost %></b></span>
            <span>Free UK delivery <b>&pound;20</b></span>
            <span>Bundle discount <b>&pound;<%= mmBunDiscount %></b></span>
          </div>
        </div>
        <% End If %>

        <div class="bp-incl">
          <div class="item">
            <i class="fa fa-flag"></i>
            <div><b>UK-built</b><small>Since 2008</small></div>
          </div>
          <div class="item">
            <i class="fa fa-shield"></i>
            <div><b>5-year PC cover</b><small>Lifetime support</small></div>
          </div>
          <div class="item">
            <i class="fa fa-undo"></i>
            <div><b>30-day returns</b><small>On the whole bundle</small></div>
          </div>
        </div>

        <div class="bp-cta">
          <button type="submit" class="btn btn-primary btn-lg">
            <i class="fa fa-shopping-basket"></i>Add bundle to basket
          </button>
          <a href="#configure" class="ghost-link"><i class="fa fa-sliders"></i>Customise the PC</a>
          <a class="ghost-link" id="copyBundleLink"><i class="fa fa-link"></i>Copy bundle link</a>
        </div>

        <div class="bp-foot">
          <span><i class="fa fa-check"></i>Every cable included</span>
          <span><i class="fa fa-check"></i>32-hour stress-tested</span>
          <span><i class="fa fa-check"></i>One delivery, one invoice</span>
        </div>
      </aside>

    </div>
  </div>
</section>

<!-- ===================================================================
     TRUST STRIP (shared include)
     =================================================================== -->
<!--#include file="inc_trustStripTrader.asp"-->

<!-- ===================================================================
     YOUR-BUNDLE PICKS - three completed picks, swappable
     =================================================================== -->
<section class="bp-picks">
  <div class="container">
    <div class="bp-picks__head reveal">
      <div class="eyebrow">Your bundle &middot; three picks</div>
      <h2>The <span class="display-em">bundle</span> you built.</h2>
      <p>You completed these three steps in the builder. Everything else &mdash; cables, stand hardware, graphics-card spec &mdash; is sorted automatically. Swap any of them before you check out.</p>
    </div>

    <div class="bp-picks__grid reveal">
      <a class="bp-pick-card" href="<%= mmBunChangeBase %>&edit=stand">
        <div class="bp-pick-card__vis">
          <img src="<%= mmBunStandImgSrc %>" alt="<%= Server.HTMLEncode(mmBunStandDispName) %>">
          <span class="bp-pick-card__tick"><i class="fa fa-check"></i></span>
        </div>
        <div class="bp-pick-card__body">
          <div class="bp-pick-card__meta">
            <span class="step">Step 1</span>
            <span class="sep">&middot;</span>
            <span>Stand</span>
          </div>
          <div class="bp-pick-card__name"><%= Server.HTMLEncode(mmBunStandDispName) %></div>
          <div class="bp-pick-card__desc"><%= mmBunMonCount %>-Screen Synergy Stand</div>
        </div>
        <span class="bp-pick-card__change">Change <i class="fa fa-arrow-right"></i></span>
      </a>

      <a class="bp-pick-card" href="<%= mmBunChangeBase %>&edit=screens">
        <div class="bp-pick-card__vis">
          <img src="<%= mmBunMonImgSrc %>" alt="<%= Server.HTMLEncode(mmBunMonDispName) %>">
          <span class="bp-pick-card__tick"><i class="fa fa-check"></i></span>
        </div>
        <div class="bp-pick-card__body">
          <div class="bp-pick-card__meta">
            <span class="step">Step 2</span>
            <span class="sep">&middot;</span>
            <span>Screens</span>
          </div>
          <div class="bp-pick-card__name"><%= mmBunMonDispName %></div>
          <div class="bp-pick-card__desc"></div>
        </div>
        <span class="bp-pick-card__change">Change <i class="fa fa-arrow-right"></i></span>
      </a>

      <a class="bp-pick-card" href="<%= mmBunChangeBase %>&edit=computer">
        <div class="bp-pick-card__vis">
          <img src="<%= mmMainImgSrc %>" alt="<%= Server.HTMLEncode(mmName) %>">
          <span class="bp-pick-card__tick"><i class="fa fa-check"></i></span>
        </div>
        <div class="bp-pick-card__body">
          <div class="bp-pick-card__meta">
            <span class="step">Step 3</span>
            <span class="sep">&middot;</span>
            <span>Computer</span>
          </div>
          <div class="bp-pick-card__name"><%= Server.HTMLEncode(mmName) %></div>
          <div class="bp-pick-card__desc">UK Custom Built</div>
        </div>
        <span class="bp-pick-card__change">Change <i class="fa fa-arrow-right"></i></span>
      </a>
    </div>

    <div class="bp-picks__foot reveal">
      <span class="bar"></span>
      <a href="<%= mmBunChangeBase %>"><i class="fa fa-arrow-left"></i>Back to the bundle builder</a>
      <span class="bar"></span>
    </div>
  </div>
</section>

<!-- ===================================================================
     COMPATIBILITY REASSURANCE
     =================================================================== -->
<section class="bp-compat">
  <div class="container">
    <div class="inner">
      <span class="bp-compat__badge">
        <span class="tick"><i class="fa fa-check"></i></span>
        We&rsquo;ve checked
      </span>
      <p class="bp-compat__copy">
        <strong>VESA plates matched, graphics setup ready for <%= mmBunMonCount %> screens, the right premium cables included.</strong>
        Your bundle is stress-tested <em>before</em> it ships. One delivery, one invoice, one UK phone number if anything isn&rsquo;t right.
      </p>
    </div>
  </div>
</section>

<!-- ===================================================================
     CONFIGURATOR - accordion option rows (shared include) + hybrid
     sidebar: PC selections list at top, bundle breakdown below.
     =================================================================== -->
<section class="configurator" id="configure">
  <div class="container">
    <div class="cfg-head reveal">
      <div>
        <h5>Tune the PC to your workload</h5>
        <h2>Customise the PC <em>for your needs</em>.</h2>
      </div>
      <a href="tel:03302236655" class="talk-link"><i class="fa fa-phone"></i>Or call &mdash; 0330 223 66 55</a>
    </div>

    <div class="cfg-grid">

      <!-- Options column -->
      <div class="cfg-options-wrap reveal">
<%
Dim mmI, mmOgId, mmOgDesc, mmOgShort, mmOgRaw, mmOgKindStr, mmOgHelpHtml
For mmI = 0 To mmOgCount - 1
  mmOgId       = mmOgRows(0, mmI)
  mmOgRaw      = mmOgRows(1, mmI) & ""
  mmOgDesc     = mmFriendlyOgName(mmOgRaw)
  mmOgShort    = mmFriendlyOgShortName(mmOgRaw)
  mmOgKindStr  = mmOgKind(mmOgRaw)
  mmOgHelpHtml = mmOgHelp(mmOgRaw)
  Call mmRenderOptionGroup(mmOgId, mmOgDesc, mmOgShort, mmI + 1, mmOgKindStr, mmOgHelpHtml)
Next
%>
      </div>

      <!-- Hybrid sidebar: PC selections list + bundle breakdown -->
      <aside class="cfg-summary reveal" style="transition-delay:.08s">

        <div class="cfg-impact cfg-impact--system">
          <div class="cfg-impact__head">
            <h5>System Performance</h5>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">CPU Speed</span>
            <span class="cfg-impact__stars" data-rating="speed"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Multi-Tasking</span>
            <span class="cfg-impact__stars" data-rating="mt"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Graphics Power</span>
            <span class="cfg-impact__stars" data-rating="gfx"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Monitors Supported</span>
            <span class="cfg-impact__val" data-rating="screens"></span>
          </div>
        </div>

        <div class="bp-card">
          <div class="bp-card__head">
            <h5>Your bundle</h5>
            <span class="tick"><i class="fa fa-circle" style="color:var(--up); font-size:8px;"></i>LIVE</span>
          </div>

          <!-- PC selections list (driven by data-summary-list) -->
          <ul class="cfg-summary__list" data-summary-list></ul>

          <!-- Bundle 3-item breakdown -->
          <ul class="bp-items">
            <li>
              <span class="bp-items__ico"><i class="fa fa-check"></i></span>
              <span class="bp-items__body">
                <span class="bp-items__role">Multi Screen Computer</span>
                <span class="bp-items__name"><%= Server.HTMLEncode(mmName) %> (Live total)</span>
              </span>
              <span class="bp-items__pri">&pound;<span data-pc-pri><%= mmBasePriceExDisp %></span></span>
            </li>
            <li>
              <span class="bp-items__ico"><i class="fa fa-check"></i></span>
              <span class="bp-items__body">
                <span class="bp-items__role">Multi Screen Stand</span>
                <span class="bp-items__name"><%= Server.HTMLEncode(mmBunStandName) %></span>
              </span>
              <span class="bp-items__pri">&pound;<%= mmFormatMoney0(mmBunStandPriceEx) %></span>
            </li>
            <li>
              <span class="bp-items__ico"><i class="fa fa-check"></i></span>
              <span class="bp-items__body">
                <span class="bp-items__role">Screens &middot; (&pound;<%= mmFormatMoney0(mmBunMonPriceEx) %> each)</span>
                <span class="bp-items__name"><%= mmBunMonDispName %> &times; <%= mmBunMonCount %></span>
              </span>
              <span class="bp-items__pri">&pound;<%= mmFormatMoney0(mmBunMonSubtotalEx) %></span>
            </li>
          </ul>

          <div class="bp-sub">
            <div class="bp-sub__row">
              <span>Subtotal</span>
              <span class="val">&pound;<span data-sub><%= mmFormatMoney0(mmBunSubtotalEx) %></span></span>
            </div>
            <% If mmBunDiscount > 0 Then %>
            <div class="bp-sub__row disc">
              <span>Bundle discount</span>
              <span class="val">&minus;&pound;<span data-bun-disc><%= mmBunDiscount %></span></span>
            </div>
            <% End If %>
          </div>

          <div class="bp-total">
            <span class="lbl">Bundle total</span>
            <span class="amt"><span class="sym">&pound;</span><span data-bun-ex><%= mmFormatMoney0(mmBunTotalEx) %></span></span>
          </div>
          <div class="bp-vat">
            <b>&pound;<span data-bun-inc><%= mmFormatMoney(mmBunTotalInc) %></span></b> inc VAT
          </div>

          <% If mmBunDiscount > 0 Then %>
          <div class="bp-saved">
            <span class="bp-saved__lbl"><i class="fa fa-gift"></i>You&rsquo;re saving</span>
            <span class="bp-saved__amt"><span class="sym">&pound;</span><span data-bun-saved><%= mmBunDiscount + 80 %></span></span>
          </div>
          <% End If %>

          <button type="submit" class="btn btn-primary btn-lg bp-card__cta">
            <i class="fa fa-shopping-basket"></i>Add bundle to basket
          </button>

          <div class="bp-card__trust">
            <i class="fa fa-shield"></i>
            <div>
              <strong>5-year PC cover &middot; Lifetime UK support &middot; 30-day money-back.</strong>
              One delivery. Built to order, typically ships in 3&ndash;5 working days.
            </div>
          </div>

          <a href="#full-spec" class="cfg-summary__speclink">
            <i class="fa fa-list"></i>View full bundle specification
            <i class="fa fa-angle-down" aria-hidden="true"></i>
          </a>
        </div>
      </aside>

    </div><!-- /cfg-grid -->
  </div>
</section>

<!-- ===================================================================
     FULL SPECIFICATION - live-updating PC spec (data-spec rows fed
     by traderpc.js metadata) + bundle-aware rows + bundle build
     summary. Rows with data-spec are populated by renderFullSpec()
     in the IIFE below; rows without are static / always-included.
     =================================================================== -->
<section class="full-spec" id="full-spec">
  <div class="container">

    <div class="section-head-narrow reveal">
      <h5>Full PC &amp; Bundle Specification</h5>
      <h2>Everything included in <span class="display-em">the <%= Server.HTMLEncode(mmName) %> and your bundle</span>.</h2>
      <p>Every component &mdash; the ones you just picked, and the ones we include as standard. When you choose a CPU that needs a bigger board, quieter cooler or more power, the affected parts auto-upgrade with it.</p>
    </div>

    <div class="spec-full reveal">
      <div class="spec-full__grid">
        <div class="spec-row"><span class="spec-row__lbl">Processor</span><span class="spec-row__val" data-spec="cpu">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Motherboard</span><span class="spec-row__val" data-spec="mobo">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Memory</span><span class="spec-row__val" data-spec="ram">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Graphics</span><span class="spec-row__val" data-spec="gpu">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Primary storage</span><span class="spec-row__val" data-spec="storage">&mdash;</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Secondary Storage</span><span class="spec-row__val" data-spec="2nddrive">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">CPU cooler</span><span class="spec-row__val" data-spec="cooler">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Case</span><span class="spec-row__val">Fractal Design Core 1100 &middot; sound-dampened</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Power supply</span><span class="spec-row__val" data-spec="psu">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Audio</span><span class="spec-row__val">8-channel HD audio &middot; on-board</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Monitors</span><span class="spec-row__val"><%= mmBunMonCount %>&times; <%= mmBunMonDispName %></span></div>
        <div class="spec-row"><span class="spec-row__lbl">Stand</span><span class="spec-row__val"><%= Server.HTMLEncode(mmBunStandName) %></span></div>
        <div class="spec-row"><span class="spec-row__lbl">Cables</span><span class="spec-row__val"><%= mmBunMonCount %>&times; 3m Long Premium Digital Cables (Free)</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Network</span><span class="spec-row__val">Gigabit Ethernet LAN &middot; wired</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Wireless internet</span><span class="spec-row__val" data-spec="wifi">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">USB ports</span><span class="spec-row__val">3&times; USB 3.2 &middot; 3&times; USB 2.0 &middot; 1&times; USB-C</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Operating system</span><span class="spec-row__val" data-spec="os">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Included software</span><span class="spec-row__val">DisplayFusion multi-monitor &middot; installed &amp; licensed</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Microsoft Office</span><span class="spec-row__val" data-spec="office">&mdash;</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Input Devices</span><span class="spec-row__val" data-spec="inputs">&mdash;</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Optical Drive</span><span class="spec-row__val" data-spec="optical">&mdash;</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Speakers</span><span class="spec-row__val" data-spec="speakers">&mdash;</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">BlueTooth</span><span class="spec-row__val" data-spec="bluetooth">&mdash;</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Backup system</span><span class="spec-row__val" data-spec="backup">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Warranty</span><span class="spec-row__val" data-spec="warranty">&mdash;</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Support</span><span class="spec-row__val">Lifetime phone &amp; remote support &middot; no clock</span></div>
        <div class="spec-row" data-spec-optional hidden><span class="spec-row__lbl">Extras</span><span class="spec-row__val" data-spec="extras">&mdash;</span></div>
      </div>
    </div>

    <div class="build-summary reveal">
      <div class="build-summary__label">Your bundle</div>
      <div class="build-summary__line" data-build-line><%= Server.HTMLEncode(mmName) %> &mdash; with <%= mmBunMonCount %>&times; <%= Server.HTMLEncode(mmBunMonDispName) %> array</div>
      <div class="build-summary__price">
        <span class="price-main"><span class="sym">&pound;</span><span data-build-ex><%= mmFormatMoney0(mmBunTotalEx) %></span></span>
        <span class="price-vat">+ VAT &middot; inc &pound;<span data-build-inc><%= mmFormatMoney(mmBunTotalInc) %></span></span>
      </div>
    </div>

    <div class="build-cta reveal">
      <button type="button" class="btn btn-primary btn-lg" data-build-submit>
        <i class="fa fa-shopping-basket"></i>Add bundle to basket
      </button>
      <a href="#configure" class="btn btn-ghost">
        <i class="fa fa-arrow-up"></i>Change configuration
      </a>
    </div>

    <div class="build-micro reveal">
      <span><i class="fa fa-truck"></i>One delivery, free UK mainland</span>
      <span><i class="fa fa-shield"></i>5-year cover &middot; lifetime support</span>
      <span><i class="fa fa-undo"></i>30-day bundle money-back</span>
    </div>

  </div>
</section>

<!-- ===================================================================
     FAQ - bundle-flavoured, per-machine
     =================================================================== -->
<section class="s depth" id="faq">
  <div class="container-narrow">
    <div class="section-head reveal" style="display:block; margin-bottom:38px;">
      <h5>Bundle &amp; PC questions</h5>
      <h2>The questions we get at this <span class="display-em">decision point</span>.</h2>
      <p style="margin-top:12px;">Specific to this bundle and this machine &mdash; not generic PC-shop answers. Got one not listed? <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">Call us on 0330 223 66 55</a>.</p>
    </div>

    <div class="faq-list reveal">
      <details class="faq-item" open>
        <summary>Can I swap the stand or screens in this bundle?</summary>
        <div class="faq-body">
          <p><strong>Yes &mdash; nothing is locked in until you complete checkout.</strong> Use the &ldquo;Change&rdquo; link on any of the three chips near the top of this page to go back into the builder with your current picks preserved. Swap the stand size, change screen model, even change PC &mdash; the bundle discount and free extras recalculate automatically.</p>
          <p>If you just want to keep this bundle but tune the PC spec, use the configurator above &mdash; the bundle total updates live with every choice.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Can I upgrade the CPU, RAM or storage later?</summary>
        <div class="faq-body">
          <p><strong>Yes &mdash; every Trader PC is designed for upgrades.</strong> The motherboard supports every 14th-gen Intel CPU option we sell (up to the i9 14900KF), so you can start with the i5 14400F now and upgrade later without changing the board.</p>
          <p>RAM is straightforward &mdash; two DIMM slots, up to 64&nbsp;GB DDR4 3200. Storage adds the same way: a second M.2 slot and 4 SATA ports for extra SSDs or HDDs. We&rsquo;ll walk you through any upgrade on the phone &mdash; usually a 15-minute job.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Does Windows come pre-installed and activated?</summary>
        <div class="faq-body">
          <p>Yes. Every machine ships with Windows 11 Home fully installed, activated, and tuned for trading workloads &mdash; Windows Defender exclusions for your platforms, telemetry minimised, power plan set to high-performance, scheduled updates set for out-of-hours. DisplayFusion is also pre-installed and licensed.</p>
          <p>Windows 11 Pro is available at +&pound;45 if you need BitLocker, Remote Desktop host, or domain join.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>What exactly arrives in the bundle delivery?</summary>
        <div class="faq-body">
          <p>Your <%= Server.HTMLEncode(mmName) %> (configured as above), <%= mmBunMonCount %> <%= Server.HTMLEncode(mmBunMonDispName) %> monitors, the <%= Server.HTMLEncode(mmBunStandDispName) %> Synergy Stand with all mount plates and assembly hardware, <%= mmBunMonCount %> premium 3&thinsp;m DisplayPort cables, a Wi-Fi / Bluetooth card (fitted), a pair of desktop speakers, a UK power lead, a printed setup guide, and a recovery USB drive.</p>
          <p>Everything in one delivery. If you&rsquo;re missing so much as a screw, call us and we&rsquo;ll courier it next day.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>What happens if something fails under warranty?</summary>
        <div class="faq-body">
          <p>In year one on the PC, we come to you (OnSite) or collect and repair &mdash; at our discretion, usually depending on what&rsquo;s failed. Years two to five are collection or return-to-base. You can extend OnSite to year two (+&pound;75) or year three (+&pound;150) at checkout.</p>
          <p>Stand: multi-year warranty (frame is typically covered for life). Screens: manufacturer warranty, typically 3 years. Any failed component across the whole bundle &mdash; one number to call.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How do I move my trading software across from my old PC?</summary>
        <div class="faq-body">
          <p>We do this for free as part of lifetime support. Phone or email us the day your bundle arrives &mdash; we&rsquo;ll remote-connect (TeamViewer or similar), help you install MT4/MT5/NinjaTrader/TradingView, migrate your templates, indicators, EAs and chart layouts, and get your broker connections back up. Usually takes 45&ndash;90 minutes depending on how much you&rsquo;re moving.</p>
        </div>
      </details>
    </div>
  </div>
</section>

</form>

<!-- ===================================================================
     DARREN CTA (shared include - uses mmMachineName)
     =================================================================== -->
<!--#include file="inc_darrenCTA.asp"-->

<!-- ===================================================================
     STICKY CTA - bundle total (mirrors the main form via JS submit)
     =================================================================== -->
<div class="sticky-cta" id="stickyCta">
  <div class="txt">
    <strong>Your Trader Bundle &middot; &pound;<span data-sticky-price><%= mmFormatMoney0(mmBunTotalEx) %></span> + VAT</strong>
    <span>Order today &middot; typically ships in 3&ndash;5 working days</span>
  </div>
  <a href="#configure" class="btn btn-primary btn-sm">Configure <i class="fa fa-arrow-right"></i></a>
</div>

<!-- Link-copied toast -->
<div class="bp-toast" id="bpToast"><i class="fa fa-check"></i>Bundle link copied to clipboard</div>

</div><!-- /.mm-site -->

<!-- Per-option metadata (friendly names, ratings, GPU/CPU specs).
     Keyed by idoptoptgrp; loaded before the IIFE that reads it.
     Same metadata file as viewprd-traderpc.asp - the PC inside the
     bundle is the same idProduct 333 product.
     The bundle flag is set BEFORE the metadata file loads so the
     hide loop in the IIFE below picks the right hideOnBundle /
     hideOnStandalone branch. -->
<script>window.MM_IS_BUNDLE = true;</script>
<script src="/js/products/traderpc.js"></script>

<!-- ===================================================================
     PAGE-SPECIFIC JS - configurator + bundle math, full-spec live
     updates, gallery thumbs, sticky CTA, copy-bundle-link toast.
     =================================================================== -->
<script>
(function(){
  // ------- Constants emitted from VBScript -------
  var PC_BASE_EX       = <%= mmBasePriceEx %>;
  var STAND_EX         = <%= mmBunStandPriceEx %>;
  var SCREENS_EX       = <%= mmBunMonSubtotalEx %>;     // already includes ×count
  var BUNDLE_DISCOUNT  = <%= mmBunDiscount %>;
  var MON_COUNT        = <%= mmBunMonCount %>;
  var VAT_RATE         = <%= MM_VAT_RATE - 1 %>;        // e.g. 0.20 for 20% VAT
  // Savings shown in the marketing callout - bundle discount plus the
  // value of the free wifi card (£40), free speakers (£20) and free
  // UK delivery (£20).
  var SAVINGS_EXTRAS   = (BUNDLE_DISCOUNT > 0) ? 80 : 0;
  var CABLE_COST = MON_COUNT * 15

  // ------- Option metadata lookup -------
  // Per-option overrides + ratings live in /js/products/traderpc.js,
  // keyed by idoptoptgrp. Missing ids return null and the page falls
  // back to the DB description with default ratings.
  function metaFor(id) {
    var t = window.MM_OPTION_META;
    return (t && id != null && t[id]) ? t[id] : null;
  }

  // ------- State -------
  var rows = document.querySelectorAll('.cfg-row');
  var state = {}; // group -> { name, delta, idoptoptgrp, meta }

  rows.forEach(function(row){
    var group = row.dataset.group;
    var sel   = row.querySelector('.cfg-option.is-selected') || row.querySelector('.cfg-option');
    if (sel) {
      state[group] = {
        name:        sel.dataset.name,
        delta:       parseInt(sel.dataset.delta || '0', 10),
        idoptoptgrp: sel.dataset.idoptoptgrp,
        meta:        metaFor(sel.dataset.idoptoptgrp)
      };
    }
  });

  // ------- Hide options flagged in metadata -------
  // Removes the option button from the DOM when its meta entry
  // sets hide:true (everywhere), hideOnBundle:true (this page),
  // or hideOnStandalone:true (the standalone product page only,
  // never matches here). If the hidden option was the group's
  // default-selected one, promote the next remaining option,
  // refresh state, and overwrite the hidden idOption input so
  // the cart posts what the user actually sees.
  var isBundle = window.MM_IS_BUNDLE === true;
  rows.forEach(function(row){
    var group = row.dataset.group;
    var hiddenInput = row.querySelector('input[type="hidden"][name^="idOption"]');
    var lostSelection = false;
    row.querySelectorAll('.cfg-option').forEach(function(btn){
      var m = metaFor(btn.dataset.idoptoptgrp);
      if (!m) return;
      var shouldHide = m.hide === true
        || (isBundle  && m.hideOnBundle     === true)
        || (!isBundle && m.hideOnStandalone === true);
      if (!shouldHide) return;
      if (btn.classList.contains('is-selected')) lostSelection = true;
      btn.parentNode.removeChild(btn);
    });
    if (lostSelection) {
      var next = row.querySelector('.cfg-option');
      if (next) {
        next.classList.add('is-selected');
        state[group] = {
          name:        next.dataset.name,
          delta:       parseInt(next.dataset.delta || '0', 10),
          idoptoptgrp: next.dataset.idoptoptgrp,
          meta:        metaFor(next.dataset.idoptoptgrp)
        };
        if (hiddenInput) hiddenInput.value = next.dataset.idoptoptgrp;
      }
    }
  });

  // ------- Apply friendly names from metadata -------
  // Override the option-button label and the row's selected caption
  // when MM_OPTION_META supplies a `name`. DB description stays as
  // the data-name fallback for any option without a meta entry.
  document.querySelectorAll('.cfg-option').forEach(function(btn){
    var m = metaFor(btn.dataset.idoptoptgrp);
    if (!m || !m.name) return;
    var nameEl = btn.querySelector('.opt-name');
    if (nameEl) nameEl.textContent = m.name;
  });
  rows.forEach(function(row){
    var s = state[row.dataset.group];
    if (!s || !s.meta || !s.meta.name) return;
    var selectedEl = row.querySelector('[data-selected]');
    if (selectedEl) selectedEl.textContent = s.meta.name;
  });

  // ------- GPU group: sort & bucket by meta.screens -------
  // Reorders the GPU group's option buttons ascending by
  // meta.screens, then by data-delta (price) within each bucket.
  // Inserts a "{N} Screen Options" caption before each new bucket.
  // Options without a screens value fall into an "Other Options"
  // bucket at the end. No-op if no GPU option has screens set.
  (function rearrangeGpuByScreens(){
    var gpuGroup = null;
    for (var k in state) {
      if (!state.hasOwnProperty(k)) continue;
      var sm = state[k] && state[k].meta;
      if (sm && sm.gpuPower !== undefined) { gpuGroup = k; break; }
    }
    if (!gpuGroup) return;

    var row = document.querySelector('.cfg-row[data-group="' + gpuGroup + '"]');
    if (!row) return;
    var optsWrap = row.querySelector('.cfg-options');
    if (!optsWrap) return;

    var buttons = Array.prototype.slice.call(optsWrap.querySelectorAll('.cfg-option'));
    var hasScreens = buttons.some(function(b){
      var m = metaFor(b.dataset.idoptoptgrp);
      return m && typeof m.screens === 'number';
    });
    if (!hasScreens) return;

    buttons.sort(function(a, b){
      var ma = metaFor(a.dataset.idoptoptgrp) || {};
      var mb = metaFor(b.dataset.idoptoptgrp) || {};
      var sa = (typeof ma.screens === 'number') ? ma.screens : Infinity;
      var sb = (typeof mb.screens === 'number') ? mb.screens : Infinity;
      if (sa !== sb) return sa - sb;
      return parseInt(a.dataset.delta || '0', 10) - parseInt(b.dataset.delta || '0', 10);
    });

    optsWrap.innerHTML = '';
    var lastBucket = undefined;
    buttons.forEach(function(btn){
      var m = metaFor(btn.dataset.idoptoptgrp) || {};
      var thisBucket = (typeof m.screens === 'number') ? m.screens : null;
      if (thisBucket !== lastBucket) {
        var heading = document.createElement('p');
        heading.textContent = (thisBucket !== null) ? (thisBucket + ' Screen Options') : 'Other Options';
        optsWrap.appendChild(heading);
        lastBucket = thisBucket;
      }
      optsWrap.appendChild(btn);
    });
  })();

  // ------- Accordion (one row open at a time) -------
  document.querySelectorAll('.cfg-row__head').forEach(function(head){
    head.addEventListener('click', function(){
      var row = head.closest('.cfg-row');
      var willOpen = !row.classList.contains('is-open');
      document.querySelectorAll('.cfg-row.is-open').forEach(function(r){
        r.classList.remove('is-open');
        var h = r.querySelector('.cfg-row__head');
        if (h) h.setAttribute('aria-expanded', 'false');
      });
      if (willOpen) {
        row.classList.add('is-open');
        head.setAttribute('aria-expanded', 'true');
      }
    });
  });

  // ------- Formatting -------
  function fmt0(n)  { return n.toLocaleString('en-GB', { minimumFractionDigits: 0, maximumFractionDigits: 0 }); }
  function fmt2(n)  { return n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function decodeHtml(s) {
    return String(s || '')
      .replace(/&middot;/g, '·')
      .replace(/&nbsp;/g,  ' ')
      .replace(/&amp;/g,   '&')
      .replace(/&lt;/g,    '<')
      .replace(/&gt;/g,    '>')
      .replace(/&quot;/g,  '"');
  }

  function renderStars(el, n) {
    if (!el) return;
    // Clamp to 0..10 then snap to nearest half so 6.4 and 6.6 both
    // resolve cleanly to 6.5 / 6 respectively.
    var clamped = Math.max(0, Math.min(10, Number(n) || 0));
    var snapped = Math.round(clamped * 2) / 2;
    var full    = Math.floor(snapped);
    var half    = (snapped - full) === 0.5;
    var html    = '';
    for (var i = 1; i <= 10; i++) {
      if (i <= full)                    html += '<span class="star">★</span>';
      else if (i === full + 1 && half)  html += '<span class="star half">★</span>';
      else                              html += '<span class="star faint">★</span>';
    }
    var changed = el.dataset.prev !== String(snapped);
    el.innerHTML = html;
    if (changed) {
      el.classList.remove('is-changed');
      void el.offsetWidth;                   // force reflow so animation retriggers
      el.classList.add('is-changed');
      el.dataset.prev = String(snapped);
    }
  }

  function renderNumber(el, n) {
    if (!el) return;
    var val = (n == null || isNaN(Number(n))) ? '' : String(Number(n));
    var changed = el.dataset.prev !== val;
    el.textContent = val;
    if (changed) {
      el.classList.remove('is-changed');
      void el.offsetWidth;                   // force reflow so animation retriggers
      el.classList.add('is-changed');
      el.dataset.prev = val;
    }
  }

  // Identify the CPU/RAM/GPU groups by which meta field the
  // currently-selected option carries. If no option in any group has
  // meta yet, the helper returns null and updateImpact uses defaults.
  function findGroupByMetaField(field) {
    for (var k in state) {
      if (!state.hasOwnProperty(k)) continue;
      var m = state[k] && state[k].meta;
      if (m && m[field] !== undefined) return k;
    }
    return null;
  }
  function cpuState() { var k = findGroupByMetaField('cpuSpeed');   return k ? state[k] : null; }
  function ramState() { var k = findGroupByMetaField('ramMtBonus'); return k ? state[k] : null; }
  function gpuState() { var k = findGroupByMetaField('gpuPower');   return k ? state[k] : null; }

  function updateImpact() {
    var cpu = cpuState(), ram = ramState(), gpu = gpuState();
    var cpuMeta = cpu && cpu.meta;
    var ramMeta = ram && ram.meta;
    var gpuMeta = gpu && gpu.meta;

    var speed  = (cpuMeta && cpuMeta.cpuSpeed     != null) ? cpuMeta.cpuSpeed     : 5;
    var mtBase = (cpuMeta && cpuMeta.cpuMultiTask != null) ? cpuMeta.cpuMultiTask : 5;
    var ramBn  = (ramMeta && ramMeta.ramMtBonus   != null) ? ramMeta.ramMtBonus   : 0;
    var mt     = Math.min(10, mtBase + ramBn);
    var gfx    = (gpuMeta && gpuMeta.gpuPower     != null) ? gpuMeta.gpuPower     : 5;
    var ai     = (gpuMeta && gpuMeta.gpuAi        != null) ? gpuMeta.gpuAi        : 4;
    var screens = (gpuMeta && gpuMeta.screens     != null) ? gpuMeta.screens      : '';

    renderStars(document.querySelector('[data-rating="speed"]'), speed);
    renderStars(document.querySelector('[data-rating="mt"]'),    mt);
    renderStars(document.querySelector('[data-rating="gfx"]'),   gfx);
    renderNumber(document.querySelector('[data-rating="screens"]'), screens);

    // Inline CPU info card (inside CPU accordion body).
    var mthread = (cpuMeta && cpuMeta.cpuMultiThread != null) ? cpuMeta.cpuMultiThread : 5;
    renderStars(document.querySelector('[data-cpu-stat="speed"]'),   speed);
    renderStars(document.querySelector('[data-cpu-stat="mt"]'),      mt);
    renderStars(document.querySelector('[data-cpu-stat="mthread"]'), mthread);
    var ctxCpuEl = document.querySelector('[data-cpu-stat="ctx"]');
    var coresEl  = document.querySelector('[data-cpu-stat="cores"]');
    if (ctxCpuEl) ctxCpuEl.textContent = (cpuMeta && (cpuMeta.specText || cpuMeta.name)) || '';
    if (coresEl)  coresEl.textContent  = (cpuMeta && cpuMeta.coresLabel) || '';

    // Inline GPU info card (inside GPU accordion body).
    renderStars(document.querySelector('[data-gpu-stat="gfx"]'), gfx);
    renderNumber(document.querySelector('[data-gpu-stat="ai"]'), ai);
    var ctxGpuEl = document.querySelector('[data-gpu-stat="ctx"]');
    var vramEl   = document.querySelector('[data-gpu-stat="vram"]');
    var portsEl  = document.querySelector('[data-gpu-stat="ports"]');
    var resEl    = document.querySelector('[data-gpu-stat="res"]');
    if (ctxGpuEl) ctxGpuEl.textContent = (gpuMeta && (gpuMeta.gpuLabel || gpuMeta.name)) || '';
    if (vramEl)   vramEl.textContent   = (gpuMeta && gpuMeta.vram) || '';
    if (portsEl)  portsEl.textContent  = (gpuMeta && gpuMeta.outputs) || '';
    if (resEl)    resEl.textContent    = (gpuMeta && gpuMeta.resolutions) || '';
  }

  // ------- Summary list -------
  // Top of the hybrid sidebar - one row per option group with the
  // short label, currently-selected option name, and price delta.
  function renderSummary() {
    var listEl = document.querySelector('[data-summary-list]');
    if (!listEl) return;
    var html = '';
    rows.forEach(function(row){
      var group = row.dataset.group;
      var labelText = row.dataset.shortLabel;
      if (!labelText) {
        var lbl = row.querySelector('.cfg-row__label');
        labelText = lbl ? lbl.textContent.replace(/^\d+/, '').trim() : group;
      }
      var s = state[group] || { name: '', delta: 0 };
      var priCls  = s.delta > 0 ? 'pri inc' : 'pri';
      var priText = s.delta > 0 ? '+ £' + fmt0(s.delta) : 'Inc.';
      var displayedName = (s.meta && s.meta.name) ? s.meta.name : decodeHtml(s.name);
      if (displayedName && displayedName.trim().toLowerCase() === 'none') return;
      html += '<li>' +
                '<span class="lbl">' + labelText + '</span>' +
                '<span class="val">' + displayedName + '</span>' +
                '<span class="' + priCls + '">' + priText + '</span>' +
              '</li>';
    });
    listEl.innerHTML = html;
  }

  // ------- Full specification (live-updating spec table) -------
  // Reads each option's meta.specKey / specText to populate the
  // matching [data-spec="<key>"] cell. CPU and GPU options can also
  // carry derived-component fields (cooler/mobo/fans on CPU, psu on
  // GPU) which fill the auto-upgrade rows. lineParts feeds the
  // bundle-items "Computer" row [data-pc-line]; the build-summary
  // line is left as the server-rendered bundle line.
  function setSpecVal(key, value, isUpgraded, hideRow) {
    var el = document.querySelector('[data-spec="' + key + '"]');
    if (!el) return;
    var row = el.closest('.spec-row');
    if (hideRow) {
      if (row) row.hidden = true;
      return;
    }
    if (row) row.hidden = false;
    el.textContent = value;
    el.classList.toggle('is-upgraded', !!isUpgraded);
  }
  function renderFullSpec() {
    document.querySelectorAll('.spec-row[data-spec-optional]').forEach(function(r){ r.hidden = true; });

    Object.keys(state).forEach(function(g){
      var s = state[g], m = s && s.meta;
      if (!m) return;
      if (m.specRows) {
        Object.keys(m.specRows).forEach(function(k){
          if (m.specSkip) setSpecVal(k, '', false, true);
          else            setSpecVal(k, m.specRows[k], false, false);
        });
        return;
      }
      if (!m.specKey) return;
      if (m.specSkip) { setSpecVal(m.specKey, '', false, true); return; }
      var text = m.specText || m.name || decodeHtml(s.name);
      setSpecVal(m.specKey, text, false, false);
    });

    var cpu = cpuState() && cpuState().meta;
    if (cpu) {
      if (cpu.cooler) setSpecVal('cooler', cpu.cooler, !!cpu.coolerUpgraded);
      if (cpu.mobo)   setSpecVal('mobo',   cpu.mobo,   !!cpu.moboUpgraded);
    }
    var gpu = gpuState() && gpuState().meta;
    if (gpu) {
      if (gpu.psu)  setSpecVal('psu',  gpu.psu,  !!gpu.psuUpgraded);
      if (gpu.mobo) setSpecVal('mobo', gpu.mobo, !!gpu.moboUpgraded);
    }

    // Update the bundle-items "Computer" row label with a short
    // CPU + RAM + storage summary. The build-summary line stays as
    // the server-rendered "PC name - with NxMonitor array".
    var ram = ramState() && ramState().meta;
    var storage = (function(){
      for (var k in state) {
        if (state.hasOwnProperty(k) && state[k].meta && state[k].meta.specKey === 'storage') return state[k].meta;
      }
      return null;
    })();
    var lineParts = [];
    if (cpu) lineParts.push((cpu.name || '').replace(/^Intel\s+/, ''));
    if (ram && (ram.ramShort || ram.name)) lineParts.push(ram.ramShort || ram.name);
    if (storage && (storage.storageShort || storage.name)) lineParts.push(storage.storageShort || storage.name);
    var pcLineEl = document.querySelector('[data-pc-line]');
    if (pcLineEl && lineParts.length > 0) pcLineEl.textContent = lineParts.filter(Boolean).join(' · ');
  }

  // ------- Recalc + totals -------
  // Computes PC ex-VAT then bundle totals (PC + stand + screens -
  // discount), updates every price hook on the page in one pass.
  function recalc() {
    var pcTotal = PC_BASE_EX;
    Object.keys(state).forEach(function(g){ pcTotal += state[g].delta || 0; });

    var bundleSubtotal = pcTotal + STAND_EX + SCREENS_EX;
    var bundleTotalEx  = bundleSubtotal - BUNDLE_DISCOUNT;
    var bundleTotalInc = bundleTotalEx * (1 + VAT_RATE);
    var savedAmt       = BUNDLE_DISCOUNT + SAVINGS_EXTRAS + CABLE_COST;

    // Hero
    var heroEx    = document.querySelector('[data-hero-ex]');
    var heroInc   = document.querySelector('[data-hero-inc]');
    var heroPc    = document.querySelector('[data-hero-pc]');
    var heroSaved = document.querySelector('[data-hero-saved]');
    if (heroEx)    heroEx.textContent    = fmt0(bundleTotalEx);
    if (heroInc)   heroInc.textContent   = fmt2(bundleTotalInc);
    if (heroPc)    heroPc.textContent    = '£' + fmt0(pcTotal);
    if (heroSaved) heroSaved.textContent = fmt0(savedAmt);

    // Sidebar - PC line price + bundle totals
    var pcPriEl  = document.querySelector('[data-pc-pri]');
    var subEl    = document.querySelector('[data-sub]');
    var bunEx    = document.querySelector('[data-bun-ex]');
    var bunInc   = document.querySelector('[data-bun-inc]');
    var bunSaved = document.querySelector('[data-bun-saved]');
    if (pcPriEl)  pcPriEl.textContent  = fmt0(pcTotal);
    if (subEl)    subEl.textContent    = fmt0(bundleSubtotal);
    if (bunEx)    bunEx.textContent    = fmt0(bundleTotalEx);
    if (bunInc)   bunInc.textContent   = fmt2(bundleTotalInc);
    if (bunSaved) bunSaved.textContent = fmt0(savedAmt);

    // Build summary (bottom of full-spec) - bundle totals
    var buildEx  = document.querySelector('[data-build-ex]');
    var buildInc = document.querySelector('[data-build-inc]');
    if (buildEx)  buildEx.textContent  = fmt0(bundleTotalEx);
    if (buildInc) buildInc.textContent = fmt2(bundleTotalInc);

    // Sticky CTA - bundle total
    var stickyEl = document.querySelector('[data-sticky-price]');
    if (stickyEl) stickyEl.textContent = fmt0(bundleTotalEx);

    renderSummary();
    updateImpact();
    renderFullSpec();
  }

  // ------- Wire up clicks on option buttons -------
  document.querySelectorAll('.cfg-row .cfg-option').forEach(function(btn){
    btn.addEventListener('click', function(){
      var row   = btn.closest('.cfg-row');
      if (!row) return;
      var group = row.dataset.group;
      row.querySelectorAll('.cfg-option').forEach(function(b){ b.classList.remove('is-selected'); });
      btn.classList.add('is-selected');
      var meta = metaFor(btn.dataset.idoptoptgrp);
      state[group] = {
        name:        btn.dataset.name,
        delta:       parseInt(btn.dataset.delta || '0', 10),
        idoptoptgrp: btn.dataset.idoptoptgrp,
        meta:        meta
      };
      var selectedEl = row.querySelector('[data-selected]');
      if (selectedEl) selectedEl.textContent = (meta && meta.name) ? meta.name : btn.dataset.name;
      var hidden = row.querySelector('input[type="hidden"][name^="idOption"]');
      if (hidden) hidden.value = btn.dataset.idoptoptgrp;
      recalc();
    });
  });

  // Initial paint
  recalc();

  // ------- Spec-panel "Add to basket" submits the configurator form -------
  var buildSubmit = document.querySelector('[data-build-submit]');
  var cfgForm     = document.getElementById('cfgForm');
  if (buildSubmit && cfgForm) {
    buildSubmit.addEventListener('click', function(){ cfgForm.submit(); });
  }

  // ------- Gallery thumbs (decorative) -------
  document.querySelectorAll('.bp-thumb').forEach(function(t){
    t.addEventListener('click', function(){
      document.querySelectorAll('.bp-thumb.is-active').forEach(function(x){ x.classList.remove('is-active'); });
      t.classList.add('is-active');
    });
  });

  // ------- Copy bundle link -------
  var copyBtn   = document.getElementById('copyBundleLink');
  var copyToast = document.getElementById('bpToast');
  if (copyBtn && copyToast) {
    copyBtn.addEventListener('click', function(e){
      e.preventDefault();
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(window.location.href);
        }
      } catch (err) { /* noop */ }
      copyToast.classList.add('is-on');
      clearTimeout(copyBtn._t);
      copyBtn._t = setTimeout(function(){ copyToast.classList.remove('is-on'); }, 1800);
    });
  }

  // ------- Sticky CTA visibility -------
  var sticky   = document.getElementById('stickyCta');
  var hero     = document.querySelector('.bp-hero');
  var footerEl = document.querySelector('footer');
  if (sticky && hero && footerEl) {
    function onScroll(){
      var y = window.scrollY || window.pageYOffset;
      var heroBottom     = hero.getBoundingClientRect().bottom + y;
      var footerTop      = footerEl.getBoundingClientRect().top + y;
      var viewportBottom = y + window.innerHeight;
      if (y > heroBottom + 120 && viewportBottom < footerTop) sticky.classList.add('visible');
      else                                                    sticky.classList.remove('visible');
    }
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
  }
})();
</script>

<!--#include file="footer_wrapper.asp"-->
