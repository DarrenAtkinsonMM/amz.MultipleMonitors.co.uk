<%
' ==============================================================
' inc_traderPcConfigurator.asp
' 2026 redesign - shared configurator helpers for the Trader PC
' product (idProduct 333) and its bundle end-page.
'
' Extracted from viewprd-traderpc.asp so the same Sub +
' Functions can drive both viewprd-traderpc.asp and
' viewPrd-TraderPC-bundle.asp without drift.
'
' Provides:
'   Sub  mmRenderOptionGroup(ogId, ogDesc, ogShort, ogIndex,
'                            ogKind, ogHelpHtml)
'     - emits one collapsible accordion option group
'   Function mmFormatMoney(v)        2dp money string
'   Function mmFormatMoney0(v)       0dp money string
'   Function mmFriendlyOgName(d)     long display name for headings
'   Function mmFriendlyOgShortName(d) short label for sidebar list
'   Function mmOgHelp(d)             contextual help paragraph (HTML)
'   Function mmOgKind(d)             "cpu" / "gpu" / "" classifier
'
' Prerequisites - the including page must define BEFORE this
' include:
'   Const MM_PRODUCT_ID = <pc idProduct>
'   Const MM_VAT_RATE   = 1.2
' and have an open ADO connection in connTemp (from common.asp).
' ==============================================================

' Module-level - bundle page sets this before the option-group loop
' to force a specific GPU option to render selected (e.g. when the
' bundle's monitor count exceeds the default GPU's screen capacity).
' Standalone product page leaves it at 0 and gets today's behaviour.
Dim mmGpuPreselectIdoptoptgrp : mmGpuPreselectIdoptoptgrp = 0

' Per-PC GPU lookup. The Trader PC's default GPU drives 4 screens, so
' bundles of 1-4 monitors keep the default. 5/6/8-screen bundles must
' upgrade to a GPU that can physically drive the stand.
Function mmTraderPcGpuIdForMonCount(ByVal monCount)
  Select Case monCount
    Case 5, 6 : mmTraderPcGpuIdForMonCount = 18519   ' nVidia RTX 5050 x2
    Case 8    : mmTraderPcGpuIdForMonCount = 18519   ' nVidia RTX A400 x2
    Case Else : mmTraderPcGpuIdForMonCount = 0       ' no override (1-4 mons)
  End Select
End Function

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

  ' Default selection is the first/cheapest row. When ogKind = "gpu"
  ' and the bundle page has set a preselect idoptoptgrp, find that
  ' row and select it instead - the firstPriceInc baseline above is
  ' unchanged so the override option still renders with its real
  ' upgrade delta.
  Dim selectedIdx, selectedDescrip, selectedIdoptoptgrp
  selectedIdx = 0
  If ogKind = "gpu" And mmGpuPreselectIdoptoptgrp > 0 Then
    Dim k
    For k = 0 To count - 1
      If CLng(rows(0, k)) = CLng(mmGpuPreselectIdoptoptgrp) Then
        selectedIdx = k
        Exit For
      End If
    Next
  End If
  selectedDescrip     = rows(6, selectedIdx) & ""
  selectedIdoptoptgrp = rows(0, selectedIdx)

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
        <span class="cfg-row__selected" data-selected><%= Server.HTMLEncode(selectedDescrip) %></span>
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
    If j = selectedIdx Then
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
        <input type="hidden" name="idOption<%= ogIndex %>" value="<%= selectedIdoptoptgrp %>">
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

' Friendly display names for option groups - DB stays unchanged.
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
%>
