<%
' ============================================================
' viewPrd-TraderPC-v2.asp
' 2026 redesign — Trader PC product page (idProduct 333).
' Rewritten to pull live option pricing from the DB while
' keeping ProductCart's cart submission contract (POST to
' instPrd.asp with idOption1..idOptionN).
' See /computer-pages-redesign-plan.md for the approach.
' ============================================================
%>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "viewPrd-TraderPC-v2.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<%
' ------------------------------------------------------------
' 1. Product base row
' ------------------------------------------------------------
Const MM_PRODUCT_ID = 333
Const MM_VAT_RATE   = 1.2

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
  mmOgRows = mmOgRs.GetRows()
  mmOgCount = UBound(mmOgRows, 2) + 1
End If
mmOgRs.Close : Set mmOgRs = Nothing

' Machine name exposed to the Darren CTA include
Dim mmMachineName : mmMachineName = mmName

' ------------------------------------------------------------
' 3. Sub: render one option-group row + its option buttons.
'    Queries options_optionsGroups for live prices (mirrors
'    pcs_makeOptionBox logic in viewPrdCode.asp:2855).
'    Also emits the hidden idOption<N> input the cart needs.
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
    rows = rs.GetRows()
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

' Helpers for page display
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
%>
<!--#include file="header_wrapper.asp"-->

<div class="mm-site">

<!-- ===================================================================
     BREADCRUMB
     =================================================================== -->
<nav class="breadcrumb">
  <div class="container inner">
    <a href="/">Home</a>
    <span class="sep">/</span>
    <a href="/trading-computers/">Trading Computers</a>
    <span class="sep">/</span>
    <span class="current"><%= Server.HTMLEncode(mmName) %></span>
  </div>
</nav>

<!-- ===================================================================
     PRODUCT HERO (gallery + buy-box)
     =================================================================== -->
<section class="pd-hero">
  <div class="container">
    <div class="pd-hero-grid">

      <!-- Gallery column -->
      <div class="pd-gallery reveal">
        <div class="pd-gallery__main">
          <span class="pd-gallery__chip">
            <span class="dot"></span><span class="acc">5YR</span>HARDWARE&nbsp;COVER
          </span>
          <img id="pdMainImg" src="<%= mmMainImgSrc %>" alt="<%= Server.HTMLEncode(mmName) %>" />
          <% If mmSku <> "" Then %>
          <span class="pd-gallery__sku">SKU &middot; <%= Server.HTMLEncode(mmSku) %></span>
          <% End If %>
        </div>
        <div class="pd-gallery__thumbs">
          <div class="pd-thumb is-active" data-img="<%= mmMainImgSrc %>">
            <img src="<%= mmMainImgSrc %>" alt="<%= Server.HTMLEncode(mmName) %>" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/antecp7_detail.jpg">
            <img src="/shop/pc/catalog/antecp7_detail.jpg" alt="Case detail" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/intel12cpu_general.jpg">
            <img src="/shop/pc/catalog/intel12cpu_general.jpg" alt="Intel CPU" />
          </div>
          <div class="pd-thumb" data-img="/shop/pc/catalog/nvidiaT600_general.jpg">
            <img src="/shop/pc/catalog/nvidiaT600_general.jpg" alt="nVidia RTX GPU" />
          </div>
          <div class="pd-thumb placeholder">
            <i class="fa fa-play-circle-o"></i>
            <span>60-sec<br>walkthrough</span>
          </div>
        </div>
      </div>

      <!-- Buy-box column -->
      <aside class="pd-buybox reveal" style="transition-delay:.08s">
        <div class="eyebrow">2026 Refresh &middot; Designed for Traders</div>
        <h1>Trader <em>PC</em></h1>
        <p class="pitch">
          The trader&rsquo;s entry point. Built in the UK around the Intel i5 14400F,
          spec&rsquo;d for MT4, TradingView, TradeStation and NinjaTrader up to four screens.
          Silent, tested, shipped with everything you need.
        </p>

        <div class="pd-tp">
          <span class="tp-stars"><span></span><span></span><span></span><span></span><span></span></span>
          <b>4.9</b>
          <small>&middot; 90+ reviews</small>
          <a href="#reviews">Read reviews <i class="fa fa-arrow-down" style="font-size:10px;"></i></a>
        </div>

        <div class="pd-price">
          <div>
            <div class="pd-price__from">From</div>
            <div class="pd-price__num"><span class="sym">&pound;</span><span data-base-ex><%= mmBasePriceExDisp %></span></div>
          </div>
          <div class="pd-price__vat">
            <b>&pound;<span data-base-inc><%= mmBasePriceIncDisp %></span></b> inc VAT<br>
            <span style="text-transform:none; font-family:'Geist', sans-serif; letter-spacing:0; color:var(--slate);">+ UK delivery &pound;10 &middot; international by quote</span>
          </div>
        </div>

        <div class="pd-incl">
          <div class="item">
            <i class="fa fa-flag"></i>
            <div><b>UK&ndash;built</b><small>Since 2008</small></div>
          </div>
          <div class="item">
            <i class="fa fa-shield"></i>
            <div><b>5-year cover</b><small>1yr OnSite</small></div>
          </div>
          <div class="item">
            <i class="fa fa-life-ring"></i>
            <div><b>Lifetime support</b><small>Phone &amp; remote</small></div>
          </div>
        </div>

        <div class="pd-cta">
          <a href="#configure" class="btn btn-primary btn-lg">
            Configure &amp; order <i class="fa fa-arrow-right"></i>
          </a>
        </div>

        <div class="pd-foot">
          <span><i class="fa fa-wrench"></i>Customise to your needs</span>
          <span><i class="fa fa-check"></i>32 hour stress-tested before delivery</span>
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
     KEY SPECS GRID (base configuration — per-machine copy)
     =================================================================== -->
<section class="s specs">
  <div class="container">
    <div class="section-head-narrow reveal">
      <h5 style="margin-bottom:14px;">Base configuration</h5>
      <h2>Perfect for Traders <span class="display-em">here&rsquo;s why</span>.</h2>
      <p>The Trader PC is a fantastic choice for traders running platforms like Trading View, MT4/5, broker platforms like CMC, IG, or Interactive Brokers.</p>
    </div>

    <div class="spec-grid">
      <div class="spec-card reveal">
        <div class="spec-card__icon"><i class="fa fa-microchip"></i></div>
        <div class="spec-card__label">Processors</div>
        <div class="spec-card__value">Intel 14th Generation</div>
        <div class="spec-card__desc">Our benchmark tests show these CPUs are perfectly suited to running trading and charting platforms really well. Pick from i5's, i7's or even the i9.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.06s">
        <div class="spec-card__icon"><i class="fa fa-database"></i></div>
        <div class="spec-card__label">Memory</div>
        <div class="spec-card__value">16&nbsp;GB - 64&nbsp;GB DDR4</div>
        <div class="spec-card__desc">RAM is working memory for your computer. The more RAM you have, the more programs, charts and files you can have open at the same time.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.12s">
        <div class="spec-card__icon"><i class="fa fa-hdd-o"></i></div>
        <div class="spec-card__label">Storage</div>
        <div class="spec-card__value">500&nbsp;GB - 4&nbsp;TB NVMe SSD</div>
        <div class="spec-card__desc">Your SSD drive is where you store your files and folders, and where Windows is installed. For most traders a 500&nbsp;GB drive is more than enough.</div>
      </div>
      <div class="spec-card reveal">
        <div class="spec-card__icon"><i class="fa fa-desktop"></i></div>
        <div class="spec-card__label">Graphics</div>
        <div class="spec-card__value">nVidia Multi-Screen Cards</div>
        <div class="spec-card__desc">The default nVidia RTX card can support up to 4 monitors. You can change to a setup that can run 8 screens, or add more graphics power if you want.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.06s">
        <div class="spec-card__icon"><i class="fa fa-windows"></i></div>
        <div class="spec-card__label">Software</div>
        <div class="spec-card__value">Windows 11</div>
        <div class="spec-card__desc">Pre-installed, activated, tuned for trading workloads. We also supply a fully licensed multi-monitor software suite called DisplayFusion.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.12s">
        <div class="spec-card__icon"><i class="fa fa-shield"></i></div>
        <div class="spec-card__label">Warranty</div>
        <div class="spec-card__value">5-year hardware cover</div>
        <div class="spec-card__desc">1st year OnSite / collection. Extend to 2 or 3 years for extra piece of mind. Plus <strong>lifetime</strong> remote support for the life of the machine.</div>
      </div>
    </div>

    <div class="spec-box reveal" style="transition-delay:.18s">
      <div class="spec-box__lead">
        <div class="spec-box__icon"><i class="fa fa-archive"></i></div>
        <div>
          <div class="spec-box__label">In the box</div>
          <div class="spec-box__title">Everything you need to trade.</div>
        </div>
      </div>
      <div class="spec-chips">
        <span class="spec-chip"><i class="fa fa-check"></i>Fractal Design case</span>
        <span class="spec-chip"><i class="fa fa-check"></i>BeQuiet 500&thinsp;W PSU</span>
        <span class="spec-chip"><i class="fa fa-check"></i>UK power lead</span>
        <span class="spec-chip"><i class="fa fa-check"></i>DisplayFusion licence</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Recovery drive</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Setup guide</span>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     CONFIGURATOR — DB-driven option rows inside one cart form.
     Posts to /shop/pc/instPrd.asp with idOption1..idOptionN whose
     values are valid idoptoptgrp IDs; instPrd.asp re-queries live
     prices on submit, so the on-page prices are display-only.
     =================================================================== -->
<section class="configurator" id="configure">
  <div class="container">
    <div class="cfg-head reveal">
      <div>
        <h5>Build your <%= Server.HTMLEncode(mmName) %></h5>
        <h2>Configure it the way you&rsquo;ll <em>actually use it</em>.</h2>
      </div>
      <a href="tel:03302236655" class="talk-link"><i class="fa fa-phone"></i>Or call &mdash; 0330 223 66 55</a>
    </div>

    <form method="post" action="/shop/pc/instPrd.asp" id="cfgForm">
      <input type="hidden" name="idproduct"        value="<%= MM_PRODUCT_ID %>">
      <input type="hidden" name="quantity"         value="1">
      <input type="hidden" name="OptionGroupCount" value="<%= mmOgCount %>">

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

        <!-- Sticky summary sidebar -->
        <aside class="cfg-summary reveal" style="transition-delay:.08s">

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

          <div class="cfg-summary__card">
            <div class="cfg-summary__head">
              <h5>Your <%= Server.HTMLEncode(mmName) %></h5>
              <span class="tick"><i class="fa fa-check" style="margin-right:4px;"></i>Live</span>
            </div>

            <ul class="cfg-summary__list" data-summary-list></ul>

            <div class="cfg-total">
              <span class="lbl">Your price</span>
              <span class="amt"><span class="sym">&pound;</span><span data-total-ex><%= mmBasePriceExDisp %></span></span>
            </div>
            <div class="cfg-vat">
              <b>&pound;<span data-total-inc><%= mmBasePriceIncDisp %></span></b> inc VAT
            </div>

            <button type="submit" class="btn btn-primary btn-lg cfg-summary__cta">
              <i class="fa fa-shopping-basket"></i>Add to basket
            </button>

            <div class="cfg-summary__trust">
              <i class="fa fa-truck"></i>
              <div>
                <strong>Free UK delivery on trader bundles.</strong>
                Single PCs from &pound;10. Built to order, typically ships in 3&ndash;5 working days.
              </div>
            </div>
          </div>
        </aside>

      </div><!-- /cfg-grid -->
    </form>
  </div>
</section>

<!-- ===================================================================
     BENCHMARKS PANEL — per-machine (indicative data, hardcoded)
     =================================================================== -->
<section class="s depth" id="benchmarks">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>TraderSpec data</h5>
        <h2>What <span class="display-em">each CPU upgrade</span> actually buys you.</h2>
        <p style="max-width:720px; margin-top:12px;">Single-thread chart-rendering performance across the four CPU options available on Trader PC. Higher is better. The i5 14400F is the sweet spot for most traders &mdash; the i7 and i9 earn their price only if you&rsquo;re running heavier workloads.</p>
      </div>
    </div>

    <div class="bench-panels">
      <div class="bench-panel reveal">
        <h4>Single-thread chart performance</h4>
        <span class="sub">Higher is better &middot; indicative mockup data</span>
        <div class="bench-bars">
          <div class="bench-row"><span class="name ours">i9 14900KF (+&pound;265)</span><div class="barwrap"><div class="bar" style="width:98%"></div></div><span class="val">98</span></div>
          <div class="bench-row"><span class="name ours">i7 14700KF (+&pound;145)</span><div class="barwrap"><div class="bar" style="width:93%"></div></div><span class="val">93</span></div>
          <div class="bench-row"><span class="name ours">i5 14600KF (+&pound;65)</span><div class="barwrap"><div class="bar" style="width:88%"></div></div><span class="val">88</span></div>
          <div class="bench-row"><span class="name ours">i5 14400F (std)</span><div class="barwrap"><div class="bar" style="width:84%"></div></div><span class="val">84</span></div>
          <div class="bench-row"><span class="name">Off-the-shelf PC</span><div class="barwrap"><div class="bar alt" style="width:45%"></div></div><span class="val">45</span></div>
        </div>
        <p class="bench-caption">MT4, TradingView, MultiCharts &mdash; chart rendering is a single-thread problem. Going from the i5 14400F to the 14600KF is a 5% uplift for &pound;65; the i7 and i9 gains are real but diminishing.</p>
      </div>

      <div class="bench-panel reveal" style="transition-delay:.08s">
        <h4>Backtest &amp; multi-task throughput</h4>
        <span class="sub">Higher is better &middot; indicative mockup data</span>
        <div class="bench-bars">
          <div class="bench-row"><span class="name ours">i9 14900KF (+&pound;265)</span><div class="barwrap"><div class="bar" style="width:95%"></div></div><span class="val">95</span></div>
          <div class="bench-row"><span class="name ours">i7 14700KF (+&pound;145)</span><div class="barwrap"><div class="bar" style="width:86%"></div></div><span class="val">86</span></div>
          <div class="bench-row"><span class="name ours">i5 14600KF (+&pound;65)</span><div class="barwrap"><div class="bar" style="width:76%"></div></div><span class="val">76</span></div>
          <div class="bench-row"><span class="name ours">i5 14400F (std)</span><div class="barwrap"><div class="bar" style="width:64%"></div></div><span class="val">64</span></div>
          <div class="bench-row"><span class="name">Off-the-shelf PC</span><div class="barwrap"><div class="bar alt" style="width:38%"></div></div><span class="val">38</span></div>
        </div>
        <p class="bench-caption">NinjaTrader strategy analyser, TradeStation backtests, Bloomberg multi-session. If you do these things &mdash; or think you will &mdash; step up a CPU tier, or step up to the <a href="/products/trader-pro/" style="color:var(--brand); font-weight:500;">Trader Pro</a>.</p>
      </div>
    </div>

    <p style="text-align:center; margin:36px 0 0;" class="reveal">
      <a href="https://traderspec.com" target="_blank" rel="noopener" class="btn btn-ghost">
        See the full methodology on TraderSpec.com <i class="fa fa-external-link"></i>
      </a>
    </p>
  </div>
</section>

<!-- ===================================================================
     CROSS-LINK BAND — upsell to next-tier
     =================================================================== -->
<section class="xlink">
  <div class="container">
    <div class="xlink-grid">

      <a href="/shop/pc/viewprod.asp?idproduct=343" class="xlink-card reveal">
        <h5>Considering Trader Pro instead?</h5>
        <h3>Step up to <em>Trader Pro</em></h3>
        <p>If you&rsquo;re running NinjaTrader strategy analysers overnight, using Bloomberg, or going past 6 screens, the extra cores of the Trader Pro earn their price. Core Ultra CPUs &middot; DDR5 &middot; same 5-year cover.</p>
        <div class="xlink-card__foot">
          <span class="from">From<b>&pound;1,345</b></span>
          <span class="arr">See Trader Pro <i class="fa fa-arrow-right"></i></span>
        </div>
      </a>

      <a href="/trading-computers/" class="xlink-card reveal" style="transition-delay:.06s">
        <h5>Not sure which is right?</h5>
        <h3>Compare <em>side-by-side</em></h3>
        <p>Our comparison table rates both machines on the platforms you actually run &mdash; MT4, TradingView, NinjaTrader, TradeStation, Bloomberg, backtesting &mdash; using live benchmark data. Pick the right one in 60 seconds.</p>
        <div class="xlink-card__foot">
          <span class="from">Takes<b>60 sec</b></span>
          <span class="arr">Compare both <i class="fa fa-arrow-right"></i></span>
        </div>
      </a>

    </div>
  </div>
</section>

<!-- ===================================================================
     FIRMS STRIP (shared include)
     =================================================================== -->
<!--#include file="inc_firmsStrip.asp"-->

<!-- ===================================================================
     REVIEWS (per-machine — hardcode real ones here)
     =================================================================== -->
<section class="s reviews" id="reviews">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Trader reviews</h5>
        <h2>Traders who picked the <span class="display-em">Trader PC</span>.</h2>
        <p>All reviews are voluntary &mdash; we don&rsquo;t ask for them.</p>
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
        <span class="platform">MT4 / MT5 &middot; 4-screen</span>
        <h4>Perfect spec for a part-time FX trader</h4>
        <p>Wanted a proper machine without going mad on spec. Trader PC with the i5 14600KF and 32&nbsp;GB RAM runs four MT5 instances, TradingView and a browser with 20+ tabs without breaking a sweat. Arrived silent &mdash; I honestly thought it was off for the first ten minutes.</p>
        <div class="meta">
          <div class="ava">MD</div>
          <div class="who">Michael D., Manchester</div>
          <div class="when">03&thinsp;/&thinsp;2026</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.08s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <span class="platform">TradingView &middot; 4-screen</span>
        <h4>Darren talked me <em>down</em> from the Pro</h4>
        <p>Originally asked for a Trader Pro. Darren asked what I actually do &mdash; TradingView, IG, maybe 15 charts &mdash; and told me the Trader PC would do it with money to spare. Saved me about &pound;400 I would have spent chasing specs I&rsquo;d never use. Four months in, zero regrets.</p>
        <div class="meta">
          <div class="ava">RL</div>
          <div class="who">Rachel L., Bristol</div>
          <div class="when">02&thinsp;/&thinsp;2026</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.16s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <span class="platform">NinjaTrader &middot; 6-screen upgrade</span>
        <h4>Upgraded from 4 to 6 screens &mdash; painless</h4>
        <p>Bought a Trader PC last year running 4 screens. This month I added the 8-screen GPU upgrade &mdash; Multiple Monitors posted the card, I installed it with a ten-minute phone call walking me through it. Try getting that level of support from Scan or Amazon.</p>
        <div class="meta">
          <div class="ava">JP</div>
          <div class="who">Jon P., Edinburgh</div>
          <div class="when">01&thinsp;/&thinsp;2026</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     FAQ (per-machine — hardcoded)
     =================================================================== -->
<section class="s depth" id="faq">
  <div class="container-narrow">
    <div class="section-head reveal" style="display:block; margin-bottom:38px;">
      <h5>Trader PC questions</h5>
      <h2>The six questions we get <span class="display-em">most often</span>.</h2>
      <p style="margin-top:12px;">Specific to this machine &mdash; not generic PC-shop answers. Got a question not listed? <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">Call us on 0330 223 66 55</a>.</p>
    </div>

    <div class="faq-list reveal">
      <details class="faq-item" open>
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
        <summary>What&rsquo;s included in the box?</summary>
        <div class="faq-body">
          <p>The PC itself, a UK power lead, a printed setup guide, and a recovery USB drive for Windows reinstalls. Anything you add in the configurator (keyboard, speakers, Wi-Fi card, etc.) arrives in the same carton, bench-tested with the machine.</p>
          <p>What&rsquo;s <em>not</em> in the box: monitors and cables. If you need those, a <a href="/bundles/" style="color:var(--brand); font-weight:500;">trader bundle</a> is typically &pound;100&ndash;&pound;200 cheaper than buying separately.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How do I connect my existing monitors?</summary>
        <div class="faq-body">
          <p>The standard nVidia RTX A400 has four mini-DisplayPort outputs and ships with four mini-DP to DisplayPort adapters. If your monitors have HDMI only, we can throw in DP-to-HDMI cables at cost &mdash; just call and ask.</p>
          <p>If you&rsquo;re running more than four screens or monitors with unusual inputs, tell us what you have when you order and we&rsquo;ll ship with the right cables in the box.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>What happens if a part fails under warranty?</summary>
        <div class="faq-body">
          <p>In year one, we come to you (OnSite) or collect and repair &mdash; at our discretion, usually depending on what&rsquo;s failed. Years two to five are collection or return-to-base. You can extend OnSite to year two (+&pound;75) or year three (+&pound;150) at checkout.</p>
          <p>In practice: if a trader&rsquo;s machine goes down during market hours, we&rsquo;ll often courier a replacement part the same day and recover yours at leisure. Not a written policy &mdash; just what we do because we&rsquo;d want someone to do it for us.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How do I move my trading software and data across from my old PC?</summary>
        <div class="faq-body">
          <p>We do this for free as part of lifetime support. Phone or email us the day your new machine arrives &mdash; we&rsquo;ll remote-connect (TeamViewer or similar), help you install MT4/MT5/NinjaTrader/TradingView, migrate your templates, indicators, EAs and chart layouts, and get your broker connections back up. Usually takes 45&ndash;90 minutes depending on how much you&rsquo;re moving.</p>
        </div>
      </details>
    </div>

    <div class="darren-inline reveal">
      <div class="avatar"><i class="fa fa-user"></i></div>
      <div>
        <h4>Question not on the list?</h4>
        <p>Seventeen years of pre-sale conversations means we&rsquo;ve heard most things. Phone or email Darren &mdash; he&rsquo;ll give you a straight answer, or tell you honestly if the Trader PC isn&rsquo;t the right machine for you.</p>
      </div>
      <div>
        <a href="tel:03302236655" class="btn btn-primary"><i class="fa fa-phone"></i>0330 223 66 55</a>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     DARREN CTA (shared include — uses mmMachineName)
     =================================================================== -->
<!--#include file="inc_darrenCTA.asp"-->

<!-- ===================================================================
     STICKY CONFIGURE CTA (inline — per-page text)
     =================================================================== -->
<div class="sticky-cta" id="stickyCta">
  <div class="txt">
    <strong><%= Server.HTMLEncode(mmName) %> &middot; &pound;<span data-sticky-price><%= mmBasePriceExDisp %></span> + VAT</strong>
    <span>Order today &middot; typically ships in 3&ndash;5 working days</span>
  </div>
  <a href="#configure" class="btn btn-primary btn-sm">Configure <i class="fa fa-arrow-right"></i></a>
</div>

</div><!-- /.mm-site -->

<!-- ===================================================================
     PAGE-SPECIFIC JS — configurator, gallery, sticky CTA, impact stars
     =================================================================== -->
<script>
(function(){
  var BASE_EX  = <%= mmBasePriceEx %>;
  var VAT_RATE = 0.20;

  // ------- State -------
  var rows = document.querySelectorAll('.cfg-row');
  var state = {}; // group -> { name, delta, idoptoptgrp }

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

  // ------- Live performance ratings (page-specific mapping) -------
  // Match on stable substrings of the option description.
  function getCpuRating(name) {
    if (/14900KF/.test(name)) return { speed: 5, mtBase: 5 };
    if (/14700KF/.test(name)) return { speed: 5, mtBase: 4 };
    if (/14600KF/.test(name)) return { speed: 4, mtBase: 3 };
    return { speed: 3, mtBase: 3 };          // i5 14400F (default)
  }
  function getRamBonus(name) {
    if (/64\s?GB/i.test(name)) return 1;
    if (/32\s?GB/i.test(name)) return 1;
    return 0;                                 // 16 GB (default)
  }
  function getGpuRating(name) {
    if (/RTX\s?5050/i.test(name)) {
      return {
        gfx: 5, ai: 5, label: 'RTX 5050 · 8 screens',
        mons: [
          { n: 8, res: '4K @ 120 Hz' },
          { n: 8, res: '1440p @ 240 Hz' },
          { n: 8, res: '1080p @ 360 Hz' }
        ]
      };
    }
    if (/Dual/i.test(name) && /A400/i.test(name)) {
      return {
        gfx: 3, ai: 2, label: 'Dual A400 · 8 screens',
        mons: [
          { n: 8, res: '4K @ 60 Hz' },
          { n: 8, res: '1440p @ 144 Hz' },
          { n: 8, res: '1080p @ 240 Hz' }
        ]
      };
    }
    return {
      gfx: 3, ai: 2, label: 'RTX A400 · 4 screens',
      mons: [
        { n: 4, res: '4K @ 60 Hz' },
        { n: 4, res: '1440p @ 144 Hz' },
        { n: 4, res: '1080p @ 240 Hz' }
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
      void el.offsetWidth;                   // force reflow so animation retriggers
      el.classList.add('is-changed');
      el.dataset.prev = String(n);
    }
  }

  // Group keys coming from the DB are "g1", "g2"... — we identify the
  // CPU/RAM/GPU groups by looking at the option descriptions in state.
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
  function shortLabel(name) {
    var s = decodeHtml(name);
    var dot = s.indexOf('·');
    if (dot > -1) s = s.slice(0, dot).trim();
    s = s.replace(/^Intel\s+/, '').replace(/\s+DDR[345]\s*\d*$/, '').trim();
    return s.length > 28 ? s.slice(0, 26) + '…' : s;
  }

  // ------- Summary list -------
  function renderSummary() {
    var listEl = document.querySelector('[data-summary-list]');
    if (!listEl) return;
    var html = '';
    rows.forEach(function(row){
      var group = row.dataset.group;
      var lbl   = row.querySelector('.cfg-row__label');
      var labelText = lbl ? lbl.textContent.replace(/^\d+/, '').trim() : group;
      var s = state[group] || { name: '', delta: 0 };
      var priCls  = s.delta > 0 ? 'pri inc' : 'pri';
      var priText = s.delta > 0 ? '+ £' + fmt0(s.delta) : 'Std';
      html += '<li>' +
                '<span class="lbl">' + labelText + '</span>' +
                '<span class="val">' + decodeHtml(s.name) + '</span>' +
                '<span class="' + priCls + '">' + priText + '</span>' +
              '</li>';
    });
    listEl.innerHTML = html;
  }

  // ------- Recalc + totals -------
  function recalc() {
    var total = BASE_EX;
    Object.keys(state).forEach(function(g){ total += state[g].delta || 0; });

    var totalExEl  = document.querySelector('[data-total-ex]');
    var totalIncEl = document.querySelector('[data-total-inc]');
    var stickyEl   = document.querySelector('[data-sticky-price]');
    if (totalExEl)  totalExEl.textContent  = fmt0(total);
    if (totalIncEl) totalIncEl.textContent = fmt2(total * (1 + VAT_RATE));
    if (stickyEl)   stickyEl.textContent   = fmt0(total);

    renderSummary();
    updateImpact();
  }

  // ------- Wire up clicks on option buttons -------
  document.querySelectorAll('.cfg-row .cfg-option').forEach(function(btn){
    btn.addEventListener('click', function(){
      var row   = btn.closest('.cfg-row');
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

  // Initial paint
  recalc();

  // ------- Gallery thumbs -> main image -------
  var thumbs = document.querySelectorAll('.pd-thumb[data-img]');
  var main   = document.getElementById('pdMainImg');
  if (main) {
    thumbs.forEach(function(t){
      t.addEventListener('click', function(){
        document.querySelectorAll('.pd-thumb.is-active').forEach(function(x){ x.classList.remove('is-active'); });
        t.classList.add('is-active');
        main.src = t.dataset.img;
      });
    });
  }

  // ------- Sticky CTA visibility -------
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
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
  }
})();
</script>

<!--#include file="footer_wrapper.asp"-->
