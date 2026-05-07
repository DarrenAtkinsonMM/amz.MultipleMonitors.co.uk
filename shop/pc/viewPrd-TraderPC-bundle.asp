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

' ------------------------------------------------------------
' 3. Sub: render one option-group row + its option buttons.
'    (Identical to viewPrd-TraderPC-v2.asp.)
' ------------------------------------------------------------
Sub mmRenderOptionGroup(ByVal ogId, ByVal ogDesc, ByVal ogIndex)
  Dim sql, rs, rows, count
  count = 0
  sql = "SELECT oog.idoptoptgrp, oog.price, oog.Wprice, oog.sortOrder, " & _
        "       oog.InActive, o.idOption, o.optionDescrip " & _
        "FROM options_optionsGroups oog " & _
        "INNER JOIN options o ON oog.idOption = o.idOption " & _
        "WHERE oog.idOptionGroup = " & ogId & " " & _
        "  AND oog.idProduct = " & MM_PRODUCT_ID & " " & _
        "  AND (oog.InActive = 0 OR oog.InActive IS NULL) " & _
        "ORDER BY oog.sortOrder, oog.price, o.optionDescrip"

  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sql, connTemp, adOpenStatic, adLockReadOnly, adCmdText
  If Not rs.EOF Then
    rows  = rs.GetRows()
    count = UBound(rows, 2) + 1
  End If
  rs.Close : Set rs = Nothing
  If count = 0 Then Exit Sub

  Dim priceCol
  If Session("customerType") = 1 Then priceCol = 2 Else priceCol = 1

  Dim firstPriceInc
  firstPriceInc = CDbl(rows(priceCol, 0))

  Dim firstDescrip, firstId
  firstDescrip = rows(6, 0) & ""
  firstId      = rows(0, 0)

  Dim groupKey
  groupKey = "g" & ogIndex
%>
  <div class="cfg-row" data-group="<%= groupKey %>">
    <div class="cfg-row__head">
      <div class="cfg-row__label"><span class="n"><%= ogIndex %></span><%= Server.HTMLEncode(ogDesc) %></div>
      <div class="cfg-row__selected" data-selected><%= Server.HTMLEncode(firstDescrip) %></div>
    </div>
    <div class="cfg-options" role="radiogroup">
<%
  Dim j, thisIdOptGrp, thisPriceInc, thisDescrip, deltaInc, deltaEx
  For j = 0 To count - 1
    thisIdOptGrp = rows(0, j)
    thisPriceInc = CDbl(rows(priceCol, j))
    thisDescrip  = rows(6, j) & ""
    deltaInc = thisPriceInc - firstPriceInc
    deltaEx  = CLng(Round(deltaInc / MM_VAT_RATE, 0))

    Dim cls, priceTxt, priceCls
    If j = 0 Then
      cls = "cfg-option is-selected"
    Else
      cls = "cfg-option"
    End If
    If deltaEx <= 0 Then
      priceTxt = "Included"
      priceCls = "std"
    Else
      priceTxt = "+ &pound;" & deltaEx
      priceCls = "inc"
    End If
%>
      <button type="button" class="<%= cls %>"
              data-name="<%= Server.HTMLEncode(thisDescrip) %>"
              data-delta="<%= deltaEx %>"
              data-idoptoptgrp="<%= thisIdOptGrp %>">
        <span class="opt-name"><%= Server.HTMLEncode(thisDescrip) %></span>
        <span class="opt-price <%= priceCls %>"><%= priceTxt %></span>
      </button>
<%
  Next
%>
    </div>
    <input type="hidden" name="idOption<%= ogIndex %>" value="<%= firstId %>">
  </div>
<%
End Sub

' Display helpers
Function mmFormatMoney(ByVal v)
  mmFormatMoney = FormatNumber(v, 2, -1, 0, -1)
End Function
Function mmFormatMoney0(ByVal v)
  mmFormatMoney0 = FormatNumber(v, 0, -1, 0, -1)
End Function

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
mmBunChangeBase = "/bundles/?sid=" & mmBunSid & "&mid=" & mmBunMid & "&cid=" & mmBunCid
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
  <input type="hidden" name="idproduct"        value="<%= MM_PRODUCT_ID %>">
  <input type="hidden" name="quantity"         value="1">
  <input type="hidden" name="OptionGroupCount" value="<%= mmOgCount %>">
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
            vs piecing this together from separate suppliers.
          </p>
          <div class="bp-savings__pills">
            <span>Bundle discount <b>&minus;&pound;<%= mmBunDiscount %></b></span>
            <span>Free Wi-Fi card <b>&pound;40</b></span>
            <span>Free speakers <b>&pound;20</b></span>
            <span>Free UK delivery <b>&pound;20</b></span>
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
          <div class="bp-pick-card__desc"><%= mmBunMonCount %>-screen steel array &middot; UK-made Synergy</div>
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
            <span>Screens &middot; <%= mmBunMonCount %> &times;</span>
          </div>
          <div class="bp-pick-card__name"><%= Server.HTMLEncode(mmBunMonDispName) %></div>
          <div class="bp-pick-card__desc">Matched for multi-screen arrays &middot; thin bezel</div>
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
          <div class="bp-pick-card__desc">UK-built &middot; i5 14th gen &middot; tune below</div>
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
        <strong>VESA plates matched, graphics card spec&rsquo;d for <%= mmBunMonCount %> screens, every cable the right length.</strong>
        Your bundle is bench-tested <em>before</em> it ships &mdash; one delivery, one invoice, one UK phone number if anything isn&rsquo;t right.
      </p>
    </div>
  </div>
</section>

<!-- ===================================================================
     CONFIGURATOR - DB-driven option rows. Sidebar is the bundle
     breakdown (.bp-sidebar), not the standalone PC summary.
     =================================================================== -->
<section class="configurator" id="configure">
  <div class="container">
    <div class="cfg-head reveal">
      <div>
        <h5>Tune the PC to your workload</h5>
        <h2>Customise the PC the way you&rsquo;ll <em>actually use it</em>.</h2>
      </div>
      <a href="tel:03302236655" class="talk-link"><i class="fa fa-phone"></i>Or call &mdash; 0330 223 66 55</a>
    </div>

    <div class="cfg-grid">

      <!-- Options column -->
      <div class="cfg-options-wrap reveal">
<%
Dim mmI, mmOgId, mmOgDesc
For mmI = 0 To mmOgCount - 1
  mmOgId   = mmOgRows(0, mmI)
  mmOgDesc = mmOgRows(1, mmI) & ""
  Call mmRenderOptionGroup(mmOgId, mmOgDesc, mmI + 1)
Next
%>
      </div>

      <!-- Bundle sidebar -->
      <aside class="bp-sidebar reveal" style="transition-delay:.08s">

        <div class="cfg-impact cfg-impact--cpu">
          <div class="cfg-impact__head">
            <h5>CPU Impact</h5>
            <span class="cfg-impact__ctx" data-ctx-cpu></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">CPU Speed</span>
            <span class="cfg-impact__stars" data-rating="speed"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Multi-Tasking</span>
            <span class="cfg-impact__stars" data-rating="mt"></span>
          </div>
        </div>

        <div class="cfg-impact cfg-impact--gpu">
          <div class="cfg-impact__head">
            <h5>Graphics Impact</h5>
            <span class="cfg-impact__ctx" data-ctx-gpu></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Graphics Power</span>
            <span class="cfg-impact__stars" data-rating="gfx"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">AI Performance</span>
            <span class="cfg-impact__stars" data-rating="ai"></span>
          </div>
          <div class="cfg-impact__mon">
            <div class="cfg-impact__mon-lbl">Supports simultaneously</div>
            <ul class="cfg-impact__mon-list" data-mons></ul>
          </div>
        </div>

        <div class="bp-card">
          <div class="bp-card__head">
            <h5>Your bundle</h5>
            <span class="tick"><i class="fa fa-circle" style="color:var(--up); font-size:8px;"></i>LIVE</span>
          </div>

          <ul class="bp-items">
            <li>
              <span class="bp-items__ico"><i class="fa fa-cube"></i></span>
              <span class="bp-items__body">
                <span class="bp-items__role">Stand &middot; <%= Server.HTMLEncode(mmBunStandDispName) %></span>
                <span class="bp-items__name"><%= Server.HTMLEncode(mmBunStandName) %></span>
              </span>
              <span class="bp-items__pri">&pound;<%= mmFormatMoney0(mmBunStandPriceEx) %></span>
            </li>
            <li>
              <span class="bp-items__ico"><i class="fa fa-desktop"></i></span>
              <span class="bp-items__body">
                <span class="bp-items__role">Screens &middot; <%= mmBunMonCount %> &times; <%= Server.HTMLEncode(mmBunMonDispName) %></span>
                <span class="bp-items__name"><%= mmBunMonCount %> &times; &pound;<%= mmFormatMoney0(mmBunMonPriceEx) %></span>
              </span>
              <span class="bp-items__pri">&pound;<%= mmFormatMoney0(mmBunMonSubtotalEx) %></span>
            </li>
            <li class="is-live">
              <span class="bp-items__ico"><i class="fa fa-microchip"></i></span>
              <span class="bp-items__body">
                <span class="bp-items__role">Computer &middot; <%= Server.HTMLEncode(mmName) %> (live)</span>
                <span class="bp-items__name" data-pc-line><%= Server.HTMLEncode(mmName) %></span>
              </span>
              <span class="bp-items__pri">&pound;<span data-pc-pri><%= mmBasePriceExDisp %></span></span>
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
        </div>
      </aside>

    </div><!-- /cfg-grid -->
  </div>
</section>

<!-- ===================================================================
     FULL SPECIFICATION - live-updating PC spec + bundle build summary
     =================================================================== -->
<section class="full-spec" id="full-spec">
  <div class="container">

    <div class="section-head-narrow reveal">
      <h5>Full PC specification</h5>
      <h2>Everything in <span class="display-em">the <%= Server.HTMLEncode(mmName) %> inside your bundle</span>.</h2>
      <p>Every component &mdash; the ones you just picked, and the ones we include as standard. When you choose a CPU that needs a bigger board, quieter cooler or more power, the affected parts auto-upgrade with it.</p>
    </div>

    <div class="spec-full reveal">
      <div class="spec-full__grid">
        <div class="spec-row"><span class="spec-row__lbl">Processor</span><span class="spec-row__val" data-spec="cpu">Intel i5 14400F &middot; 10C/16T</span></div>
        <div class="spec-row"><span class="spec-row__lbl">CPU cooler</span><span class="spec-row__val" data-spec="cooler">be quiet! Pure Rock 2 silent tower</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Motherboard</span><span class="spec-row__val" data-spec="mobo">MSI PRO B760M-P DDR4</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Memory</span><span class="spec-row__val" data-spec="ram">16 GB DDR4 3200 &middot; Corsair Vengeance</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Graphics</span><span class="spec-row__val" data-spec="gpu">nVidia RTX A400 &middot; 4&nbsp;GB &middot; 4 screens</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Primary storage</span><span class="spec-row__val" data-spec="storage">500 GB Kingston NVMe &middot; M.2 &middot; 3,500&nbsp;MB/s read</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Case</span><span class="spec-row__val">Fractal Design Core 1100 &middot; sound-dampened</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Power supply</span><span class="spec-row__val" data-spec="psu">be quiet! Pure Power 12 500&thinsp;W &middot; 80+ Gold</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Case cooling</span><span class="spec-row__val" data-spec="fans">2&times; be quiet! Silent Wings 4 (140&thinsp;mm)</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Monitors</span><span class="spec-row__val"><%= mmBunMonCount %> &times; <%= Server.HTMLEncode(mmBunMonDispName) %></span></div>
        <div class="spec-row"><span class="spec-row__lbl">Stand</span><span class="spec-row__val"><%= Server.HTMLEncode(mmBunStandName) %> &middot; steel &middot; UK-manufactured</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Cables</span><span class="spec-row__val"><%= mmBunMonCount %>&times; premium 3&thinsp;m DisplayPort &middot; included free</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Network</span><span class="spec-row__val">Gigabit Ethernet LAN &middot; Wi-Fi AX + Bluetooth (bundle)</span></div>
        <div class="spec-row"><span class="spec-row__lbl">USB ports</span><span class="spec-row__val">3&times; USB 3.2 &middot; 3&times; USB 2.0 &middot; 1&times; USB-C</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Operating system</span><span class="spec-row__val" data-spec="os">Windows 11 Home &middot; pre-activated &middot; trader-tuned</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Included software</span><span class="spec-row__val">DisplayFusion multi-monitor &middot; installed &amp; licensed</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Warranty</span><span class="spec-row__val" data-spec="warranty">5-year hardware cover &middot; 1-year OnSite</span></div>
        <div class="spec-row"><span class="spec-row__lbl">Support</span><span class="spec-row__val">Lifetime phone &amp; remote support &middot; no clock</span></div>
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
      <button type="submit" class="btn btn-primary btn-lg">
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

<!-- ===================================================================
     PAGE-SPECIFIC JS - configurator, gallery, sticky CTA, impact stars
     =================================================================== -->
<script>
(function(){
  // Bundle constants - emitted from VBScript
  var PC_BASE_EX      = <%= mmBasePriceEx %>;
  var STAND_PRICE_EX  = <%= mmBunStandPriceEx %>;
  var SCREENS_EX      = <%= mmBunMonSubtotalEx %>;
  var BUNDLE_DISCOUNT = <%= mmBunDiscount %>;
  var MON_COUNT       = <%= mmBunMonCount %>;
  var VAT_RATE        = <%= MM_VAT_RATE - 1 %>;
  // Savings shown in the marketing callout - discount plus the value
  // of the free speakers + wifi + delivery (20+40+20=80).
  var SAVINGS_FREEBIES = (BUNDLE_DISCOUNT > 0) ? 80 : 0;

  // ------- State -------
  var rows = document.querySelectorAll('.cfg-row');
  var state = {};

  rows.forEach(function(row){
    var group = row.dataset.group;
    var sel   = row.querySelector('.cfg-option.is-selected') || row.querySelector('.cfg-option');
    if (sel) {
      state[group] = {
        name:  sel.dataset.name,
        delta: parseInt(sel.dataset.delta || '0', 10),
        idoptoptgrp: sel.dataset.idoptoptgrp
      };
    }
  });

  // ------- Formatting -------
  function fmt0(n) { return n.toLocaleString('en-GB', { minimumFractionDigits: 0, maximumFractionDigits: 0 }); }
  function fmt2(n) { return n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function decodeHtml(s) {
    return String(s || '')
      .replace(/&middot;/g, '·')
      .replace(/&nbsp;/g,  ' ')
      .replace(/&amp;/g,   '&')
      .replace(/&lt;/g,    '<')
      .replace(/&gt;/g,    '>')
      .replace(/&quot;/g,  '"');
  }

  // ------- Live performance ratings (substring matching) -------
  function getCpuRating(name) {
    if (/14900KF/.test(name)) return { speed: 5, mtBase: 5 };
    if (/14700KF/.test(name)) return { speed: 5, mtBase: 4 };
    if (/14600KF/.test(name)) return { speed: 4, mtBase: 3 };
    return { speed: 3, mtBase: 3 };
  }
  function getRamBonus(name) {
    if (/64\s?GB/i.test(name)) return 1;
    if (/32\s?GB/i.test(name)) return 1;
    return 0;
  }
  function getGpuRating(name) {
    if (/RTX\s?5050/i.test(name)) {
      return {
        gfx: 5, ai: 5, label: 'RTX 5050 · 8 screens',
        mons: [
          { n: 8, res: '4K @ 120 Hz' },
          { n: 8, res: '1440p @ 240 Hz' },
          { n: 8, res: '1080p @ 360 Hz' }
        ]
      };
    }
    if (/Dual/i.test(name) && /A400/i.test(name)) {
      return {
        gfx: 3, ai: 2, label: 'Dual A400 · 8 screens',
        mons: [
          { n: 8, res: '4K @ 60 Hz' },
          { n: 8, res: '1440p @ 144 Hz' },
          { n: 8, res: '1080p @ 240 Hz' }
        ]
      };
    }
    return {
      gfx: 3, ai: 2, label: 'RTX A400 · 4 screens',
      mons: [
        { n: 4, res: '4K @ 60 Hz' },
        { n: 4, res: '1440p @ 144 Hz' },
        { n: 4, res: '1080p @ 240 Hz' }
      ]
    };
  }

  function renderStars(el, n) {
    if (!el) return;
    var html = '';
    for (var i = 1; i <= 5; i++) {
      html += (i <= n) ? '★' : '<span class="faint">★</span>';
    }
    var changed = el.dataset.prev !== String(n);
    el.innerHTML = html;
    if (changed) {
      el.classList.remove('is-changed');
      void el.offsetWidth;
      el.classList.add('is-changed');
      el.dataset.prev = String(n);
    }
  }

  // Identify CPU / RAM / GPU groups by scanning option descriptions.
  function findGroupByDescrip(test) {
    for (var k in state) {
      if (state.hasOwnProperty(k) && test(state[k].name)) return k;
    }
    return null;
  }
  function cpuState() {
    var k = findGroupByDescrip(function(n){ return /\bi[579]\b|Intel\s+(Core|i[3-9])/i.test(n); });
    return k ? state[k] : null;
  }
  function ramState() {
    var k = findGroupByDescrip(function(n){ return /\d+\s?GB\b/i.test(n) && /DDR/i.test(n); });
    if (!k) k = findGroupByDescrip(function(n){ return /\b(16|32|64)\s?GB\b/i.test(n); });
    return k ? state[k] : null;
  }
  function gpuState() {
    var k = findGroupByDescrip(function(n){ return /(RTX|A400|GPU|screen|monitor)/i.test(n); });
    return k ? state[k] : null;
  }

  function shortLabel(name) {
    var s = decodeHtml(name);
    var dot = s.indexOf('·');
    if (dot > -1) s = s.slice(0, dot).trim();
    s = s.replace(/^Intel\s+/, '').replace(/\s+DDR[345]\s*\d*$/, '').trim();
    return s.length > 28 ? s.slice(0, 26) + '…' : s;
  }

  function updateImpact() {
    var cpu = cpuState(), ram = ramState(), gpu = gpuState();
    var cpuName = cpu ? cpu.name : '';
    var ramName = ram ? ram.name : '';
    var gpuName = gpu ? gpu.name : '';

    var cpuR = getCpuRating(decodeHtml(cpuName));
    var ramB = getRamBonus(decodeHtml(ramName));
    var speed = cpuR.speed;
    var mt    = Math.min(5, cpuR.mtBase + ramB);
    var gpuR  = getGpuRating(decodeHtml(gpuName));

    renderStars(document.querySelector('[data-rating="speed"]'), speed);
    renderStars(document.querySelector('[data-rating="mt"]'),    mt);
    renderStars(document.querySelector('[data-rating="gfx"]'),   gpuR.gfx);
    renderStars(document.querySelector('[data-rating="ai"]'),    gpuR.ai);

    var cpuCtx = document.querySelector('[data-ctx-cpu]');
    var gpuCtx = document.querySelector('[data-ctx-gpu]');
    if (cpuCtx) cpuCtx.textContent = shortLabel(cpuName) + (ramName ? ' · ' + shortLabel(ramName) : '');
    if (gpuCtx) gpuCtx.textContent = gpuR.label;

    var monsEl = document.querySelector('[data-mons]');
    if (monsEl) {
      monsEl.innerHTML = gpuR.mons.map(function(m){
        return '<li><b>' + m.n + '×</b><span class="res">' + m.res + '</span></li>';
      }).join('');
    }
  }

  // ------- Live full-spec echo (CPU-driven auto-upgrades) ------
  function setSpecVal(key, value, isUpgraded) {
    var el = document.querySelector('[data-spec="' + key + '"]');
    if (!el) return;
    el.textContent = value;
    if (isUpgraded) el.classList.add('is-upgraded');
    else            el.classList.remove('is-upgraded');
  }

  function updateFullSpec() {
    var cpu = cpuState(), ram = ramState(), gpu = gpuState();
    var cpuName = cpu ? decodeHtml(cpu.name) : '';
    var ramName = ram ? decodeHtml(ram.name) : '';
    var gpuName = gpu ? decodeHtml(gpu.name) : '';

    var isK       = /14600KF|14700KF|14900KF/.test(cpuName);
    var isHighK   = /14700KF|14900KF/.test(cpuName);
    var isRtx5050 = /RTX\s?5050/i.test(gpuName);
    var isDualGpu = /Dual/i.test(gpuName) && /A400/i.test(gpuName);

    var cooler = 'be quiet! Pure Rock 2 silent tower';
    var coolerUp = false;
    if (isHighK)   { cooler = 'be quiet! Dark Rock Pro 5 (135 mm tower)'; coolerUp = true; }
    else if (isK)  { cooler = 'be quiet! Dark Rock 4 (120 mm tower)';     coolerUp = true; }

    var mobo = 'MSI PRO B760M-P DDR4';
    var moboUp = false;
    if (isK) { mobo = 'MSI PRO Z790-P DDR4'; moboUp = true; }

    var psu = 'be quiet! Pure Power 12 500 W · 80+ Gold';
    var psuUp = false;
    if (isRtx5050) { psu = 'be quiet! Pure Power 12 650 W · 80+ Gold'; psuUp = true; }

    var fans = '2× be quiet! Silent Wings 4 (140 mm)';
    var fansUp = false;
    if (isK) { fans = '3× be quiet! Silent Wings 4 (140 mm)'; fansUp = true; }

    var gpuText;
    if (isRtx5050)      gpuText = 'nVidia RTX 5050 · 8 GB GDDR7 · 8 screens';
    else if (isDualGpu) gpuText = 'nVidia RTX A400 (Dual) · 2× 4 GB · 8 screens';
    else                gpuText = 'nVidia RTX A400 · 4 GB · 4 screens';

    setSpecVal('cpu',     cpuName, false);
    setSpecVal('cooler',  cooler,  coolerUp);
    setSpecVal('mobo',    mobo,    moboUp);
    setSpecVal('ram',     ramName ? ramName + ' · Corsair Vengeance' : '', false);
    setSpecVal('gpu',     gpuText, false);
    setSpecVal('psu',     psu,     psuUp);
    setSpecVal('fans',    fans,    fansUp);
  }

  // ------- PC "Computer" sidebar row -------
  function updatePcRow(pcTotal) {
    var lineEl = document.querySelector('[data-pc-line]');
    var priEl  = document.querySelector('[data-pc-pri]');
    if (lineEl) {
      var cpu = cpuState(), ram = ramState();
      var parts = [];
      if (cpu) parts.push(shortLabel(cpu.name));
      if (ram) parts.push(shortLabel(ram.name));
      if (parts.length > 0) lineEl.textContent = parts.join(' · ');
    }
    if (priEl) priEl.textContent = fmt0(pcTotal);
  }

  // ------- Main recalc -------
  function recalc() {
    var pcTotal = PC_BASE_EX;
    Object.keys(state).forEach(function(g){ pcTotal += state[g].delta || 0; });

    var bundleSubtotal = STAND_PRICE_EX + SCREENS_EX + pcTotal;
    var bundleTotalEx  = bundleSubtotal - BUNDLE_DISCOUNT;
    var bundleTotalInc = bundleTotalEx * (1 + VAT_RATE);

    // Hero
    var heroEx    = document.querySelector('[data-hero-ex]');
    var heroInc   = document.querySelector('[data-hero-inc]');
    var heroPc    = document.querySelector('[data-hero-pc]');
    var heroSaved = document.querySelector('[data-hero-saved]');
    if (heroEx)    heroEx.textContent    = fmt0(bundleTotalEx);
    if (heroInc)   heroInc.textContent   = fmt2(bundleTotalInc);
    if (heroPc)    heroPc.textContent    = '£' + fmt0(pcTotal);
    if (heroSaved) heroSaved.textContent = fmt0(BUNDLE_DISCOUNT + SAVINGS_FREEBIES);

    // Sidebar breakdown
    updatePcRow(pcTotal);
    var subEl    = document.querySelector('[data-sub]');
    var bunEx    = document.querySelector('[data-bun-ex]');
    var bunInc   = document.querySelector('[data-bun-inc]');
    var bunSaved = document.querySelector('[data-bun-saved]');
    if (subEl)    subEl.textContent    = fmt0(bundleSubtotal);
    if (bunEx)    bunEx.textContent    = fmt0(bundleTotalEx);
    if (bunInc)   bunInc.textContent   = fmt2(bundleTotalInc);
    if (bunSaved) bunSaved.textContent = fmt0(BUNDLE_DISCOUNT + SAVINGS_FREEBIES);

    // Build summary (bottom of full-spec)
    var buildEx  = document.querySelector('[data-build-ex]');
    var buildInc = document.querySelector('[data-build-inc]');
    if (buildEx)  buildEx.textContent  = fmt0(bundleTotalEx);
    if (buildInc) buildInc.textContent = fmt2(bundleTotalInc);

    // Sticky CTA - bundle total
    var stickyEl = document.querySelector('[data-sticky-price]');
    if (stickyEl) stickyEl.textContent = fmt0(bundleTotalEx);

    updateImpact();
    updateFullSpec();
  }

  // ------- Wire up option clicks -------
  document.querySelectorAll('.cfg-row .cfg-option').forEach(function(btn){
    btn.addEventListener('click', function(){
      var row = btn.closest('.cfg-row');
      if (!row) return;
      var group = row.dataset.group;
      row.querySelectorAll('.cfg-option').forEach(function(b){ b.classList.remove('is-selected'); });
      btn.classList.add('is-selected');
      state[group] = {
        name:  btn.dataset.name,
        delta: parseInt(btn.dataset.delta || '0', 10),
        idoptoptgrp: btn.dataset.idoptoptgrp
      };
      var selectedEl = row.querySelector('[data-selected]');
      if (selectedEl) selectedEl.innerHTML = btn.dataset.name;
      var hidden = row.querySelector('input[type="hidden"][name^="idOption"]');
      if (hidden) hidden.value = btn.dataset.idoptoptgrp;
      recalc();
    });
  });

  recalc();

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
