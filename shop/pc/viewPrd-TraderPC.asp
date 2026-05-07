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
Sub mmRenderOptionGroup(ByVal ogId, ByVal ogDesc, ByVal ogShort, ByVal ogIndex, ByVal ogKind, ByVal ogHelpHtml)
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

  Dim groupKey, openCls, ariaExpanded
  groupKey = "g" & ogIndex
  If ogIndex = 1 Then
    openCls = " is-open"
    ariaExpanded = "true"
  Else
    openCls = ""
    ariaExpanded = "false"
  End If
%>
  <div class="cfg-row<%= openCls %>" data-group="<%= groupKey %>" data-short-label="<%= Server.HTMLEncode(ogShort) %>">
    <button type="button" class="cfg-row__head"
            aria-expanded="<%= ariaExpanded %>"
            aria-controls="cfg-body-<%= groupKey %>">
      <span class="cfg-row__head-main">
        <span class="cfg-row__label"><span class="n"><%= ogIndex %></span><%= Server.HTMLEncode(ogDesc) %></span>
        <span class="cfg-row__selected" data-selected><%= Server.HTMLEncode(firstDescrip) %></span>
      </span>
      <i class="fa fa-chevron-down cfg-row__chev" aria-hidden="true"></i>
    </button>
    <div class="cfg-row__body" id="cfg-body-<%= groupKey %>">
      <div class="cfg-row__body-inner">
<% If Len(ogHelpHtml) > 0 Then %>
        <p class="cfg-row__help"><%= ogHelpHtml %></p>
<% End If %>
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
<%
  If ogKind = "cpu" Then
%>
        <div class="cfg-impact cfg-impact--cpu cfg-cpu-info" aria-live="polite">
          <div class="cfg-impact__head">
            <h5>Processor Impact</h5>
            <span class="cfg-impact__ctx" data-cpu-stat="ctx"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">CPU Speed</span>
            <span class="cfg-impact__stars" data-cpu-stat="speed"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Multi-Tasking</span>
            <span class="cfg-impact__stars" data-cpu-stat="mt"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Multi-Threaded</span>
            <span class="cfg-impact__stars" data-cpu-stat="mthread"></span>
          </div>
          <div class="cfg-impact__row cfg-impact__row--text">
            <span class="cfg-impact__lbl">Cores, Threads  &amp; GHz</span>
            <span class="cfg-impact__val" data-cpu-stat="cores"></span>
          </div>
        </div>
<%
  ElseIf ogKind = "gpu" Then
%>
        <div class="cfg-impact cfg-impact--gpu cfg-gpu-info" aria-live="polite">
          <div class="cfg-impact__head">
            <h5>Graphics Impact</h5>
            <span class="cfg-impact__ctx" data-gpu-stat="ctx"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">Graphics Power</span>
            <span class="cfg-impact__stars" data-gpu-stat="gfx"></span>
          </div>
          <div class="cfg-impact__row">
            <span class="cfg-impact__lbl">AI Performance TOPS Score (Higher is better)</span>
            <span class="cfg-impact__val" data-gpu-stat="ai"></span>
          </div>
          <div class="cfg-impact__row cfg-impact__row--text">
            <span class="cfg-impact__lbl">Video memory</span>
            <span class="cfg-impact__val" data-gpu-stat="vram"></span>
          </div>
          <div class="cfg-impact__row cfg-impact__row--text">
            <span class="cfg-impact__lbl">Monitor Ports</span>
            <span class="cfg-impact__val" data-gpu-stat="ports"></span>
          </div>
          <div class="cfg-impact__row cfg-impact__row--text">
            <span class="cfg-impact__lbl">Resolutions Supported</span>
            <span class="cfg-impact__val" data-gpu-stat="res"></span>
          </div>
        </div>
<%
  End If
%>
      </div>
    </div>
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

' Friendly display names for option groups — DB stays unchanged.
' mmFriendlyOgName     -> long form, used in the accordion heading.
' mmFriendlyOgShortName-> short form, used in the right-side summary
'                         list. Falls back to the original DB desc
'                         when no override is set.
' Match is case-insensitive against optionsGroups.OptionGroupDesc.
Function mmFriendlyOgName(ByVal ogDesc)
  Dim k : k = LCase(Trim(ogDesc & ""))
  Select Case k
    Case "os"             : mmFriendlyOgName = "Operating System"
    Case "boot hard drive": mmFriendlyOgName = "Hard Drive"
    Case "2nd hard drive": mmFriendlyOgName = "Second Hard Drive"
    Case "keyb. / mouse" : mmFriendlyOgName = "Keyboard & Mouse"
    Case "graphics cards" : mmFriendlyOgName = "Graphics Setup"
    Case "ms office" : mmFriendlyOgName = "Microsoft Office"
    Case "bluetooth" : mmFriendlyOgName = "Bluetooth Adapter"
    Case Else             : mmFriendlyOgName = ogDesc
  End Select
End Function

' Per-accordion descriptive paragraph rendered above the option
' buttons. Keyed off the raw DB description so it survives any
' display-name changes in mmFriendlyOgName. Returns inline HTML so
' the text can include links and entities (&mdash;, &nbsp;, etc.).
' Returns "" for unmapped groups, in which case the renderer skips
' the <p class="cfg-row__help"> element entirely.
Function mmOgHelp(ByVal ogRaw)
  Dim k : k = LCase(Trim(ogRaw & ""))
  Select Case k
    Case "cpu"
      mmOgHelp = "The CPU is the biggest factor in how fast your computer feels in daily use. " & _
                 "For charting generally a faster speed will help things feel snappy, multi-tasking " & _
                 "is important if you run lots of screens, charts, or multiple platforms. Multi-threaded performance " & _
                 "is important for things like backtesting and data intensive applications."
    Case "ram"
      mmOgHelp = "RAM is important for multi-tasking, running out of RAM makes a PC run very slowly. " & _
                 "16&nbsp;GB is plenty for MT4/TradingView setups across up to 4 screens. Upgrade if you want to keep more charts &amp; " & _ 
                 "tabs open or run more intensive platforms like Ninja Trader, Bloomberg, etc..."
    Case "graphics cards"
      mmOgHelp = "Graphics cards power your screens &amp; provide monitor outputs which each connected screen needs. " & _
                 "More graphics power can help run higher resolution screens smoothly and they also help support more graphical workloads. " & _ 
                 "The AI TOPS score dicates how well locally installed AI models will perform."
    Case "boot hard drive"
      mmOgHelp = "This will be your 'C drive', it is where Windows and your programs are installed. 500Gb is usually  " & _
                 "enough for most traders. Upgrade if you want extra storage space for files and folders."
    Case "2nd hard drive"
      mmOgHelp = "This is a second physical hard drive in your PC. Only add if you specifically want or need a second hard drive." & _ 
                "NVMe SSD's are fast and silent, traditional drives have larger capacities but are slower and can give off " & _ 
                "a faint humming noise and clicks."
    Case "os"
      mmOgHelp = "Both editions come pre-installed, activated and tuned for trading. Home edition is fine for most, " & _
                 "Pro edition is mainly for corporate networks or if you need enhanced remote desktop connectivity."
    Case "ms office"
      mmOgHelp = "Supplied as a lifetime license key that can be used on 1 PC only, no subscription required. Home edition gets you " & _
                 "Word, Excel, PowerPoint &amp; OneNote, the Business edition also includes Outlook."
    Case "wireless card"
      mmOgHelp = "This adds wireless Internet connection capabilities to your PC, select if you use wifi to connect at home or in the " & _
                 "office. It also includes Bluetooth functionality as well."
    Case "keyb. / mouse"
      mmOgHelp = "Add a Logitech wired or wireless mouse and keyboard set to your new PC. " 
    Case "speakers"
      mmOgHelp = "Add a USB powered set of Logitech speakers to your new PC. " 
    Case "optical drive"
      mmOgHelp = "Add a DVD ReWriter drive to your PC. This may require a different case however the price for the drive includes any case swap fee. " 
    Case "bluetooth"
      mmOgHelp = "Add Bluetooth connection functionality to your new PC. Not required if you have selected the Wifi card as this already includes Bluetooth. " 
    Case "backup system"
      mmOgHelp = "This is an extra physical hard drive inside your PC along with software that clones your C drive to it on a regular schedule. In the event . " & _ 
                "of a Windows corruption, virus, or drive failure you can get back instantly up and running with all your programs and files still installed."
    Case "warranty cover"
      mmOgHelp = "Every Trader PC gets 5-year hardware cover by default including 1 year of a on-site enhanced package. Extend the on-site length if you " & _
                 "for 2 or 3 years for extra peace of mind."
    Case Else
      mmOgHelp = ""
  End Select
End Function

' Classifies an option group by its raw DB description so the
' accordion renderer knows when to inject the inline CPU/GPU
' "impact" cards beneath the option buttons.
Function mmOgKind(ByVal ogRaw)
  Dim k : k = LCase(Trim(ogRaw & ""))
  If k = "cpu" Or InStr(k, "processor") > 0 Then
    mmOgKind = "cpu"
  ElseIf k = "graphics cards" Or InStr(k, "graphics") > 0 Then
    mmOgKind = "gpu"
  Else
    mmOgKind = ""
  End If
End Function

Function mmFriendlyOgShortName(ByVal ogDesc)
  Dim k : k = LCase(Trim(ogDesc & ""))
  Select Case k
    Case "os" : mmFriendlyOgShortName = "OS" 
    Case "warranty cover" : mmFriendlyOgShortName = "warranty"
    Case "wireless card" : mmFriendlyOgShortName = "WiFi"  
    Case "boot hard drive" : mmFriendlyOgShortName = "SSD Drive"
     Case "2nd hard drive" : mmFriendlyOgShortName = "2nd Drive"
    Case "optical drive" : mmFriendlyOgShortName = "Optical"
    Case "backup system" : mmFriendlyOgShortName = "Backup"
    Case "keyb. / mouse" : mmFriendlyOgShortName = "Inputs"
    Case "graphics cards" : mmFriendlyOgShortName = "GPU"
    Case "bluetooth" : mmFriendlyOgShortName = "BT"
    Case Else             : mmFriendlyOgShortName = ogDesc
  End Select
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

        <!-- Sticky summary sidebar -->
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

            <a href="#full-spec" class="cfg-summary__speclink">
              <i class="fa fa-list"></i>View full specification
              <i class="fa fa-angle-down" aria-hidden="true"></i>
            </a>
          </div>
        </aside>

      </div><!-- /cfg-grid -->
    </form>
  </div>
</section>

<!-- ===================================================================
     FULL SPECIFICATION — live-updating, echoes configurator state.
     Rows with data-spec are populated by renderFullSpec() in the IIFE
     below; rows without are static / always-included components.
     =================================================================== -->
<section class="full-spec" id="full-spec">
  <div class="container">

    <div class="section-head-narrow reveal">
      <h5>Full specification</h5>
      <h2>Everything in <span class="display-em">your <%= Server.HTMLEncode(mmName) %></span>.</h2>
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
      <div class="build-summary__label">Your build</div>
      <div class="build-summary__line" data-build-line>&mdash;</div>
      <div class="build-summary__price">
        <span class="price-main"><span class="sym">&pound;</span><span data-build-ex><%= mmBasePriceExDisp %></span></span>
        <span class="price-vat">+ VAT &middot; inc &pound;<span data-build-inc><%= mmBasePriceIncDisp %></span></span>
      </div>
    </div>

    <div class="build-cta reveal">
      <button type="button" class="btn btn-primary btn-lg" data-build-submit>
        <i class="fa fa-shopping-basket"></i>Add to basket
      </button>
      <a href="#configure" class="btn btn-ghost">
        <i class="fa fa-arrow-up"></i>Change configuration
      </a>
    </div>

    <div class="build-micro reveal">
      <span><i class="fa fa-shield"></i>5-year hardware cover included</span>
      <span><i class="fa fa-life-ring"></i>Lifetime phone &amp; remote support</span>
    </div>

  </div>
</section>

<!-- ===================================================================
     FIRMS STRIP (shared include)
     =================================================================== -->
<!--#include file="inc_firmsStrip.asp"-->


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

<!-- Per-option metadata (friendly names, ratings, GPU/CPU specs).
     Keyed by idoptoptgrp; loaded before the IIFE that reads it. -->
<script src="/js/products/traderpc.js"></script>

<script>
(function(){
  var BASE_EX  = <%= mmBasePriceEx %>;
  var VAT_RATE = 0.20;

  // ------- Option metadata lookup -------
  // Per-option overrides + ratings live in /js/products/traderpc-v2.js,
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

  // ------- Hide options flagged with meta.hide -------
  // Removes the option button from the DOM. If the hidden option
  // was the group's default-selected one, promote the next
  // remaining option, refresh state, and overwrite the hidden
  // idOption input so the cart posts what the user actually sees.
  rows.forEach(function(row){
    var group = row.dataset.group;
    var hiddenInput = row.querySelector('input[type="hidden"][name^="idOption"]');
    var lostSelection = false;
    row.querySelectorAll('.cfg-option').forEach(function(btn){
      var m = metaFor(btn.dataset.idoptoptgrp);
      if (!m || m.hide !== true) return;
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

    // Inline CPU info card (inside CPU accordion body) — extends the
    // sidebar's CPU Speed / Multi-Tasking with a Multi-Threaded star
    // row and the "Cores & threads" text row.
    var mthread = (cpuMeta && cpuMeta.cpuMultiThread != null) ? cpuMeta.cpuMultiThread : 5;
    renderStars(document.querySelector('[data-cpu-stat="speed"]'),   speed);
    renderStars(document.querySelector('[data-cpu-stat="mt"]'),      mt);
    renderStars(document.querySelector('[data-cpu-stat="mthread"]'), mthread);
    var ctxCpuEl = document.querySelector('[data-cpu-stat="ctx"]');
    var coresEl  = document.querySelector('[data-cpu-stat="cores"]');
    if (ctxCpuEl) ctxCpuEl.textContent = (cpuMeta && (cpuMeta.specText || cpuMeta.name)) || '';
    if (coresEl)  coresEl.textContent  = (cpuMeta && cpuMeta.coresLabel) || '';

    // Inline GPU info card (inside GPU accordion body) — same star
    // ratings as the sidebar panel plus Video memory, monitor ports and resolutions supported.
    renderStars(document.querySelector('[data-gpu-stat="gfx"]'), gfx);
    renderNumber(document.querySelector('[data-gpu-stat="ai"]'), ai);
    var ctxGpuEl = document.querySelector('[data-gpu-stat="ctx"]');
    var vramEl   = document.querySelector('[data-gpu-stat="vram"]');
    var portsEl  = document.querySelector('[data-gpu-stat="ports"]');
    var resEl  = document.querySelector('[data-gpu-stat="res"]');
    if (ctxGpuEl) ctxGpuEl.textContent = (gpuMeta && (gpuMeta.gpuLabel || gpuMeta.name)) || '';
    if (vramEl)   vramEl.textContent   = (gpuMeta && gpuMeta.vram) || '';
    if (portsEl)  portsEl.textContent  = (gpuMeta && gpuMeta.outputs) || '';
    if (resEl)  resEl.textContent  = (gpuMeta && gpuMeta.resolutions) || '';
  }
  // ------- Summary list -------
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
  // GPU) which fill the auto-upgrade rows. The "Your build" line and
  // build-summary prices are updated alongside.
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
  function renderFullSpec(total) {
    // Optional rows (wifi/bluetooth/office/backup/extras) start each
    // render hidden — they're only re-shown when a non-skip option is
    // currently selected for that key.
    document.querySelectorAll('.spec-row[data-spec-optional]').forEach(function(r){ r.hidden = true; });

    Object.keys(state).forEach(function(g){
      var s = state[g], m = s && s.meta;
      if (!m) return;
      // specRows: { key: text, ... } for one option that drives multiple
      // spec rows (e.g. a "Wifi 6 with Bluetooth" combo card).
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
      if (gpu.psu) setSpecVal('psu', gpu.psu, !!gpu.psuUpgraded);
      if (gpu.mobo)   setSpecVal('mobo',   gpu.mobo,   !!gpu.moboUpgraded);
    }

    var ram = ramState() && ramState().meta;
    var storage = (function(){
      for (var k in state) {
        if (state.hasOwnProperty(k) && state[k].meta && state[k].meta.specKey === 'storage') return state[k].meta;
      }
      return null;
    })();
    var os = (function(){
      for (var k in state) {
        if (state.hasOwnProperty(k) && state[k].meta && state[k].meta.specKey === 'os') return state[k].meta;
      }
      return null;
    })();

    var lineParts = [];
    if (cpu) lineParts.push((cpu.name || '').replace(/^Intel\s+/, ''));
    if (ram && (ram.ramShort || ram.name)) lineParts.push(ram.ramShort || ram.name);
    if (storage && (storage.storageShort || storage.name)) lineParts.push(storage.storageShort || storage.name);
    if (gpu && gpu.screens) lineParts.push(gpu.screens + ' screens');
    if (os && os.name) lineParts.push(os.name);
    var lineEl = document.querySelector('[data-build-line]');
    if (lineEl) lineEl.textContent = lineParts.filter(Boolean).join(' · ');

    var exEl  = document.querySelector('[data-build-ex]');
    var incEl = document.querySelector('[data-build-inc]');
    if (exEl)  exEl.textContent  = fmt0(total);
    if (incEl) incEl.textContent = fmt2(total * (1 + VAT_RATE));
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
    renderFullSpec(total);
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
