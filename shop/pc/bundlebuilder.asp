<%
' ============================================================
' bundlebuilder.asp
' 2026 redesign - single-page bundle builder.
' Replaces the 3-step flow in CUSTOMCAT-bundles1/2/3.asp.
' See /bundles-builder-redesign-plan.md at repo root.
'
' The stand / screen / computer arrays are defined as static
' VBScript tables below (names, images, screen counts, etc.
' stay hardcoded per the mockup). A single SQL round-trip
' fetches live retail prices (VAT-inclusive) for every
' idProduct referenced, and the JS BUNDLE_CONFIG object at
' the bottom is emitted with ex-VAT prices injected from that
' dictionary. Admin price changes flow through on reload; no
' code change needed.
' ============================================================
%>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "bundlebuilder.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<%
Const MM_VAT_RATE = 1.2

' ------------------------------------------------------------
' Static config tables.
' Columns per table documented in each section header.
' Rows with idProduct = 0 are TBC placeholders - silently
' skipped at render time until a real DB id is filled in.
' ------------------------------------------------------------

' stands columns: 0=idProduct 1=jsKey 2=name 3=screens 4=discount 5=img 6=arrayimg
Dim mmStands
mmStands = Array( _
  Array(326, "s2v",  "Dual Vertical",      2,  25, "/images/bundles/bun-s2v-med.png",  "s2v"  ), _
  Array(287, "s2h",  "Dual Horizontal",    2,  25, "/images/bundles/bun-s2h-med.png",  "s2h"  ), _
  Array(312, "s3h",  "Triple Horizontal",  3,  25, "/images/bundles/bun-s3h-med.png",  "s3h"  ), _
  Array(324, "s3p",  "Triple Pyramid",     3,  25, "/images/bundles/bun-s3p-med.png",  "s3p"  ), _
  Array(313, "s4s",  "Quad Square",        4,  50, "/images/bundles/bun-s4s-med.png",  "s4s"  ), _
  Array(337, "s4sp", "Quad Multi Pole",    4,  50, "/images/bundles/bun-s4sp-med.png", "s4sp" ), _
  Array(325, "s4p",  "Quad Pyramid",       4,  50, "/images/bundles/bun-s4p-med.png",  "s4p"  ), _
  Array(327, "s4h",  "Quad Horizontal",    4,  50, "/images/bundles/bun-s4h-med.png",  "s4h"  ), _
  Array(318, "s5p",  "Five Pyramid",       5,  50, "/images/bundles/bun-s5p-med.png",  "s5p"  ), _
  Array(338, "s6r",  "Six Monitor",        6, 100, "/images/bundles/bun-s6r-med.png",  "s6r"  ), _
  Array(314, "s6rp", "Six Multi Pole",     6, 100, "/images/bundles/bun-s6rp-med.png", "s6rp" ), _
  Array(319, "s8r",  "Eight Monitor",      8, 100, "/images/bundles/bun-s8r-med.png",  "s8r"  ) _
)

' screens columns: 0=idProduct 1=jsKey 2=name 3=desc1 4=desc2 5=desc3 6=img 7=arrayimg
Dim mmScreens
mmScreens = Array( _
  Array(304, "scr24s", "21.5"" AOC - Full HD",     "1920 x 1080 FHD", "Thin bezel",     "VA Panel",  "/shop/pc/catalog/acer22_thumb.jpg",      "a22"), _
  Array(317, "scr24i", "24"" Acer - Full HD",      "1920 x 1080 FHD", "Thin bezel",     "VA Panel",  "/shop/pc/catalog/acer22_thumb.jpg",      "a24"), _
  Array(328, "scr27s", "27"" Acer - Full HD",      "1920 x 1080 FHD", "Thin bezel",     "VA Panel",  "/shop/pc/catalog/acer22_thumb.jpg",      "a27"), _
  Array(320, "scr27i", "24"" Iiyama - Full HD",    "1920 x 1080 FHD", "Thin bezel",     "IPS Panel", "/shop/pc/catalog/iiyama23ips-thumb.jpg", "i23"), _
  Array(344, "scrAw",  "27"" AOC - Quad HD",       "2560 x 1440 QHD", "Thin bezel",     "IPS Panel", "/shop/pc/catalog/aoc27-thumb.jpg",       "i27"), _
  Array(345, "scrIiy", "27"" Iiyama - Quad HD",    "2560 x 1440 QHD", "Thin Bezel",     "IPS Panel", "/shop/pc/catalog/iiyama23ips-thumb.jpg", "i27")  _
)

' computers columns: 0=idProduct 1=jsKey 2=name 3=six 4=eight 5=desc1 6=desc2 7=desc3 8=img 9=bunimg 10=cta
Dim mmComputers
mmComputers = Array( _
  Array(306, "ultra",   "Ultra PC",   165, 165, "Fast everyday computer",         "Perfect for business / office use",           "Multi-screen ready out of the box",         "/images/bundles/bun-ultra-pc.png",   "/images/bundles/case1-bun.png", "/products/ultra-multi-monitor-pc/"),             _
  Array(307, "extreme", "Extreme PC",  65, 175, "High-end workstation",           "Powerful Intel or AMD CPUs",                  "Highly configurable, support up to 12 screens", "/images/bundles/bun-extreme-pc.png", "/images/bundles/case1-bun.png", "/products/extreme-multi-screen-computer/"),      _
  Array(333, "trader",  "Trader PC",  165, 165, "Designed for multi-screen trading", "Great for MT4, TradingView & broker platforms", "Quiet, stable & fast performance",          "/images/bundles/bun-trader-pc.png",  "/images/bundles/case1-bun.png", "/products/trader-pc/"),       _
  Array(343, "pro",     "Trader Pro", 65,  175, "Built for Professional Traders", "Run platforms like NinjaTrader & Bloomberg easily", "Intels fastest CPUs & DDR5 RAM",            "/images/bundles/bun-pro-pc.png",     "/images/bundles/case1-bun.png", "/products/trader-pro-pc/")                       _
)

' ============================================================
' Deep-link preselect pipeline:
'   1. Hydrate prices + slugs (single SQL round-trip) so we
'      can resolve slug-form deeplinks before validating.
'   2. Read slug querystrings (?stand=/?monitor=/?computer=)
'      from the Bundles N Segments rewrite rules in web.config
'      and resolve each to an idProduct via mmSlugDict.
'   3. Numeric ?sid=/?mid=/?cid= querystrings act as fallback
'      so legacy URLs keep working unchanged.
'   4. Validate the resolved ids against the static arrays.
'      Ordering rules (matches CUSTOMCAT-bundles2.asp / bundles3.asp):
'        - cid without sid OR without mid -> 301 /bundles/
'        - mid without sid                -> 301 /bundles/
'        - any id non-numeric             -> 301 /bundles/
'        - id not in static array         -> 301 to longest
'                                            valid prefix
' Valid ids are emitted as MMB_PRESELECT (below). Zeros mean
' "not set" so the JS bootstrap can treat them as falsy.
' ============================================================

' ----- Step 1: single-round-trip price + slug hydration -----
' Collects every non-zero idProduct referenced in the static
' arrays and fetches retail price + pcUrlBundle. All IDs are
' VBScript-numeric (sourced from our own arrays, never user
' input) so direct string concat is safe.
Dim mmPriceDict : Set mmPriceDict = Server.CreateObject("Scripting.Dictionary")
Dim mmSlugDict  : Set mmSlugDict  = Server.CreateObject("Scripting.Dictionary")

Dim mmAllIds, mmRow, mmId
mmAllIds = ""
For Each mmRow In mmStands
    mmId = CLng(mmRow(0))
    If mmId > 0 Then mmAllIds = mmAllIds & mmId & ","
Next
For Each mmRow In mmScreens
    mmId = CLng(mmRow(0))
    If mmId > 0 Then mmAllIds = mmAllIds & mmId & ","
Next
For Each mmRow In mmComputers
    mmId = CLng(mmRow(0))
    If mmId > 0 Then mmAllIds = mmAllIds & mmId & ","
Next

If Len(mmAllIds) > 0 Then
    mmAllIds = Left(mmAllIds, Len(mmAllIds) - 1)

    Dim mmSql, mmRs
    mmSql = "SELECT idProduct, price, pcUrlBundle FROM products " & _
            "WHERE idProduct IN (" & mmAllIds & ") " & _
            "  AND active = -1 AND removed = 0"

    On Error Resume Next
    Set mmRs = connTemp.Execute(mmSql)
    If err.number <> 0 Then
        On Error Goto 0
        Call LogErrorToDatabase()
    Else
        On Error Goto 0
        Do While Not mmRs.EOF
            mmPriceDict.Add CLng(mmRs("idProduct")), CDbl(mmRs("price"))
            mmSlugDict.Add  CLng(mmRs("idProduct")), mmRs("pcUrlBundle") & ""
            mmRs.MoveNext
        Loop
        mmRs.Close
        Set mmRs = Nothing
    End If
End If

' ----- Steps 2 + 3: parse querystrings (slug first, then id) -----
Dim mmPreSid, mmPreMid, mmPreCid
mmPreSid = 0 : mmPreMid = 0 : mmPreCid = 0

Dim mmRawSid, mmRawMid, mmRawCid, mmGuardRow
mmRawSid = Trim(Request.QueryString("sid") & "")
mmRawMid = Trim(Request.QueryString("mid") & "")
mmRawCid = Trim(Request.QueryString("cid") & "")

Dim mmRawStandSlug, mmRawMonSlug, mmRawCompSlug, mmResolvedId
mmRawStandSlug = LCase(Trim(Request.QueryString("stand") & ""))
mmRawMonSlug   = LCase(Trim(Request.QueryString("monitor") & ""))
mmRawCompSlug  = LCase(Trim(Request.QueryString("computer") & ""))

If mmRawStandSlug <> "" Then
    mmResolvedId = mmFindIdByBundleSlug(mmStands, mmRawStandSlug)
    If mmResolvedId > 0 Then mmRawSid = CStr(mmResolvedId)
End If
If mmRawMonSlug <> "" Then
    mmResolvedId = mmFindIdByBundleSlug(mmScreens, mmRawMonSlug)
    If mmResolvedId > 0 Then mmRawMid = CStr(mmResolvedId)
End If
If mmRawCompSlug <> "" Then
    mmResolvedId = mmFindIdByBundleSlug(mmComputers, mmRawCompSlug)
    If mmResolvedId > 0 Then mmRawCid = CStr(mmResolvedId)
End If

' ----- Optional ?edit= querystring: forces the builder to open
' on a specific stage instead of the default "deepest preselected
' slot" rule. Used by the bp-picks "Change" links on the final
' bundle page so clicking Change > Screens lands on the screens
' panel (otherwise all 3 picks would auto-resolve to 'computer').
Dim mmRawEdit, mmPreEdit
mmRawEdit = LCase(Trim(Request.QueryString("edit") & ""))
mmPreEdit = ""
If mmRawEdit = "stand" Or mmRawEdit = "screens" Or mmRawEdit = "computer" Then
    mmPreEdit = mmRawEdit
End If

' ----- Step 4: validate resolved ids against static arrays -----
If (mmRawCid <> "" And (mmRawSid = "" Or mmRawMid = "")) _
Or (mmRawMid <> "" And mmRawSid = "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", "/bundles/"
    Response.End
End If

If mmRawSid <> "" Then
    Dim mmOkSid : mmOkSid = False
    If IsNumeric(mmRawSid) Then
        mmPreSid = CLng(mmRawSid)
        For Each mmGuardRow In mmStands
            If CLng(mmGuardRow(0)) = mmPreSid Then mmOkSid = True : Exit For
        Next
    End If
    If Not mmOkSid Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader "Location", "/bundles/"
        Response.End
    End If
End If

If mmRawMid <> "" Then
    Dim mmOkMid : mmOkMid = False
    If IsNumeric(mmRawMid) Then
        mmPreMid = CLng(mmRawMid)
        For Each mmGuardRow In mmScreens
            If CLng(mmGuardRow(0)) = mmPreMid Then mmOkMid = True : Exit For
        Next
    End If
    If Not mmOkMid Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader "Location", "/bundles/?sid=" & mmPreSid
        Response.End
    End If
End If

If mmRawCid <> "" Then
    Dim mmOkCid : mmOkCid = False
    If IsNumeric(mmRawCid) Then
        mmPreCid = CLng(mmRawCid)
        For Each mmGuardRow In mmComputers
            If CLng(mmGuardRow(0)) = mmPreCid Then mmOkCid = True : Exit For
        Next
    End If
    If Not mmOkCid Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader "Location", "/bundles/?sid=" & mmPreSid & "&mid=" & mmPreMid
        Response.End
    End If
End If

' ------------------------------------------------------------
' JS emission helpers. Each returns "" for missing/TBC rows so
' the caller can skip them without emitting a trailing comma
' that breaks the JS object literal.
' ------------------------------------------------------------

Function mmPriceEx(ByVal idp)
    ' Returns ex-VAT price (whole pounds) or -1 if the id is
    ' missing from the price dictionary. Rounded to match the
    ' mockup's whole-pound price display.
    Dim idL : idL = CLng(idp)
    If idL <= 0 Then
        mmPriceEx = -1
        Exit Function
    End If
    If Not mmPriceDict.Exists(idL) Then
        Call LogErrorToDatabase()
        mmPriceEx = -1
        Exit Function
    End If
    mmPriceEx = Int((mmPriceDict(idL) / MM_VAT_RATE) + 0.5)
End Function

Function mmSlug(ByVal idp)
    ' Returns the Bundle slug for an id, or "" if missing.
    ' Used by mmEmitStand / mmEmitScreen so the JS CTA can
    ' build /products/trader-pc/<stand-slug>/<monitor-slug>/.
    Dim idL : idL = CLng(idp)
    If idL <= 0 Then mmSlug = "" : Exit Function
    If Not mmSlugDict.Exists(idL) Then mmSlug = "" : Exit Function
    mmSlug = mmSlugDict(idL) & ""
End Function

Function mmFindIdByBundleSlug(ByVal arr, ByVal targetSlug)
    ' Reverse-lookup: scans a static array (mmStands / mmScreens
    ' / mmComputers), reads each row's idProduct, fetches its
    ' pcUrlBundle slug from mmSlugDict, returns the idProduct
    ' whose slug equals targetSlug. Zero if no match.
    Dim row, idp, slug
    mmFindIdByBundleSlug = 0
    For Each row In arr
        idp = CLng(row(0))
        If idp > 0 And mmSlugDict.Exists(idp) Then
            slug = LCase(mmSlugDict(idp) & "")
            If slug <> "" And slug = targetSlug Then
                mmFindIdByBundleSlug = idp
                Exit Function
            End If
        End If
    Next
End Function

Function mmJsStr(ByVal s)
    ' JS-string escape: backslash, double quote, then strip
    ' newlines defensively. Mockup content is plain text so
    ' this is enough.
    Dim t : t = s & ""
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    mmJsStr = t
End Function

Sub mmEmitStand(ByVal row)
    Dim idp, px
    idp = CLng(row(0))
    px  = mmPriceEx(idp)
    If px < 0 Then Exit Sub
    Response.Write "      { id:" & idp & _
                   ", slug:""" & mmJsStr(mmSlug(idp)) & """" & _
                   ", name:""" & mmJsStr(row(2)) & """" & _
                   ", price:" & px & _
                   ", screens:" & row(3) & _
                   ", discount:" & row(4) & _
                   ", img:""" & mmJsStr(row(5)) & """" & _
                   ", arrayimg:""" & mmJsStr(row(6)) & """ }," & vbCrLf
End Sub

Sub mmEmitScreen(ByVal row)
    Dim idp, px
    idp = CLng(row(0))
    px  = mmPriceEx(idp)
    If px < 0 Then Exit Sub
    Response.Write "      { id:" & idp & _
                   ", slug:""" & mmJsStr(mmSlug(idp)) & """" & _
                   ", name:""" & mmJsStr(row(2)) & """" & _
                   ", price:" & px & _
                   ", desc1:""" & mmJsStr(row(3)) & """" & _
                   ", desc2:""" & mmJsStr(row(4)) & """" & _
                   ", desc3:""" & mmJsStr(row(5)) & """" & _
                   ", img:""" & mmJsStr(row(6)) & """" & _
                   ", arrayimg:""" & mmJsStr(row(7)) & """ }," & vbCrLf
End Sub

Sub mmEmitComputer(ByVal row)
    Dim idp, px
    idp = CLng(row(0))
    px  = mmPriceEx(idp)
    If px < 0 Then Exit Sub
    Response.Write "      { id:" & idp & _
                   ", name:""" & mmJsStr(row(2)) & """" & _
                   ", price:" & px & _
                   ", six:"   & row(3) & _
                   ", eight:" & row(4) & _
                   ", desc1:""" & mmJsStr(row(5)) & """" & _
                   ", desc2:""" & mmJsStr(row(6)) & """" & _
                   ", desc3:""" & mmJsStr(row(7)) & """" & _
                   ", img:""" & mmJsStr(row(8)) & """" & _
                   ", bunimg:""" & mmJsStr(row(9)) & """" & _
                   ", cta:""" & mmJsStr(row(10)) & """ }," & vbCrLf
End Sub

' Page-level metadata consumed by inc_headerV5.asp + GenerateMetaTags
Dim pcv_PageName, pcv_DefaultDescription
pcv_PageName = "Multiple Monitor Bundles - PC, Stand & Screens | Multiple Monitors"
pcv_DefaultDescription = "Save up to £300 with a multi-screen computer, stand and screen bundle. Includes free PC upgrades, free premium cabling, free delivery and a bundle discount."
%>
<!--#include file="header_wrapper.asp"-->

<div class="mm-site">

<!-- ===================================================================
     HERO - bundle positioning
     =================================================================== -->
<section class="hero">
  <div class="container">
    <div class="hero-grid">
      <div class="reveal">
        <div class="eyebrow">Complete bundles &middot; Since 2008</div>
        <h1>
          <em>Everything you need,</em> delivered together, for less.
        </h1>
        <p class="lead">
          PC, stand, screens and every cable, tested together, shipped together. One box, one invoice, one UK phone number if anything isn't right. Save up to &pound;300 vs ordering the same items individually from us, and it actually works when you plug it in.
        </p>
        <div class="hero-ctas">
          <a href="#builder" class="btn btn-primary btn-lg">Start building <i class="fa fa-arrow-right"></i></a>
          <a href="#starters" class="btn btn-ghost btn-lg">See popular starting points</a>
        </div>
        <div class="hero-mini">
          <div class="item"><i class="fa fa-gift"></i><span>Free cables, wifi, speakers &amp; UK delivery</span></div>
          <div class="item"><i class="fa fa-shield"></i><span>5-year hardware cover &middot; lifetime support</span></div>
        </div>
      </div>

      <div class="hero-visual reveal" style="transition-delay:.1s">
        <img src="/images/pages/trading-image.png" alt="Complete multi-screen bundle &mdash; stand, screens and PC delivered together" style="display:block; max-width:100%; height:auto; margin-left:auto;" />
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     TRUST STRIP
     =================================================================== -->
<section class="truststrip" id="trust">
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
        <div class="icon"><i class="fa fa-truck"></i></div>
        <div>
          <div class="label">Delivered</div>
          <div class="val">2,000+ Bundles</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     BUILDER - two-column: stand picker + running sidebar
     =================================================================== -->
<section class="builder" id="builder">
  <div class="container">

    <div class="section-head reveal" style="margin-bottom:30px;">
      <div>
        <span class="eyebrow">Build your bundle</span>
        <h2>Three steps to your, <span class="display-em">perfect bundle.</span></h2>
      </div>
      <div class="ab-reassure"><i class="fa fa-check-circle"></i>Every stand &amp; screen combination works with every computer.
</div>
    </div>

    <div class="page-body-grid">

      <!-- LEFT: stand picker -->
      <div class="reveal">
        <ol class="stepper" id="mmb-stepper" data-edit-stage="<%= mmPreEdit %>">
          <li class="step is-current" data-step="stand">
            <span class="step-num">1</span>
            <div class="step-body">
              <span class="step-lbl">Choose your stand</span>
              <span class="step-val">&mdash;</span>
            </div>
          </li>
          <li class="step" data-step="screens">
            <span class="step-num">2</span>
            <div class="step-body">
              <span class="step-lbl">Pick your screens</span>
              <span class="step-val">&mdash;</span>
            </div>
          </li>
          <li class="step" data-step="computer">
            <span class="step-num">3</span>
            <div class="step-body">
              <span class="step-lbl">Select your PC</span>
              <span class="step-val">&mdash;</span>
            </div>
          </li>
        </ol>

        <div class="stage">
          <div class="stage-head">
            <h3 id="mmb-stage-title">Start with the stand. <span class="muted">Everything else flexes around it.</span></h3>
            <button type="button" class="btn btn-primary" id="mmb-next-btn" style="display:none"></button>
          </div>

          <div class="stand-grid" id="mmb-grid">
            <!-- Cards rendered by the bundle builder script at the end of this file -->
          </div>

          <div class="mmb-info" id="mmb-info" style="display:none">
            <div class="mmb-info__msg">
              <i class="fa fa-arrow-right" aria-hidden="true"></i>
              <div class="mmb-info__copy">
                <strong>Next: customise the PC (CPU, RAM, storage).</strong>
                <span>Your stand, screens<span id="mmb-info-savings"></span> stay locked in.</span>
              </div>
            </div>
            <button type="button" class="btn btn-primary mmb-info__cta is-disabled" id="mmb-info-cta">Pick a PC to continue <i class="fa fa-arrow-right"></i></button>
          </div>
        </div>
      </div>

      <!-- RIGHT: running bundle sidebar -->
      <aside class="bundle-sidebar reveal" data-variant="A" style="transition-delay:.1s">
        <div class="bsb-card">
          <div class="bsb-head">
            <span class="eyebrow">Your bundle</span>
            <span class="bsb-pct num" id="mmb-pct">0%</span>
          </div>
          <div class="bsb-viz" id="mmb-viz">
            <svg width="220" height="160" viewBox="0 0 110 80" xmlns="http://www.w3.org/2000/svg">
              <!-- Stand base + pole -->
              <rect x="13" y="68" width="42" height="3" fill="var(--ink)"></rect>
              <rect x="32.5" y="12" width="3" height="56" fill="var(--ink)"></rect>
              <!-- 2x2 monitor array -->
              <g><rect x="5.5" y="18" width="28" height="18" fill="transparent" stroke="var(--ink)" stroke-width="1.2"></rect><rect x="8.5" y="21" width="22" height="12" fill="var(--ink)" opacity=".6"></rect></g>
              <g><rect x="34.5" y="18" width="28" height="18" fill="transparent" stroke="var(--ink)" stroke-width="1.2"></rect><rect x="37.5" y="21" width="22" height="12" fill="var(--ink)" opacity=".6"></rect></g>
              <g><rect x="5.5" y="37" width="28" height="18" fill="transparent" stroke="var(--ink)" stroke-width="1.2"></rect><rect x="8.5" y="40" width="22" height="12" fill="var(--ink)" opacity=".6"></rect></g>
              <g><rect x="34.5" y="37" width="28" height="18" fill="transparent" stroke="var(--ink)" stroke-width="1.2"></rect><rect x="37.5" y="40" width="22" height="12" fill="var(--ink)" opacity=".6"></rect></g>
              <!-- Computer case (front view, tall) -->
              <g>
                <rect x="84" y="32" width="22" height="36" fill="transparent" stroke="var(--ink)" stroke-width="1.2"></rect>
                <rect x="85.5" y="33.5" width="19" height="33" fill="var(--ink)" opacity=".6"></rect>
                <rect x="85.5" y="33.5" width="19" height="4" fill="var(--ink)"></rect>
                <circle cx="101" cy="35.5" r="0.9" fill="#fff" opacity=".9"></circle>
                <rect x="87.5" y="35" width="2" height="1" fill="#fff" opacity=".6"></rect>
                <rect x="90.5" y="35" width="2" height="1" fill="#fff" opacity=".6"></rect>
                <line x1="87.5" y1="42" x2="102.5" y2="42" stroke="#fff" stroke-width="0.4" opacity=".55"></line>
                <line x1="87.5" y1="46" x2="102.5" y2="46" stroke="#fff" stroke-width="0.4" opacity=".5"></line>
                <line x1="87.5" y1="50" x2="102.5" y2="50" stroke="#fff" stroke-width="0.4" opacity=".5"></line>
                <line x1="87.5" y1="54" x2="102.5" y2="54" stroke="#fff" stroke-width="0.4" opacity=".5"></line>
                <line x1="87.5" y1="58" x2="102.5" y2="58" stroke="#fff" stroke-width="0.4" opacity=".5"></line>
                <line x1="87.5" y1="62" x2="102.5" y2="62" stroke="#fff" stroke-width="0.4" opacity=".5"></line>
              </g>
            </svg>
          </div>
          <ol class="bsb-list" id="mmb-list">
            <li data-slot="stand">
              <span class="bsb-l">Stand</span>
              <span class="bsb-r"><em class="muted">Not selected</em></span>
              <span class="bsb-p num">&mdash;</span>
            </li>
            <li data-slot="screens">
              <span class="bsb-l">Screens</span>
              <span class="bsb-r"><em class="muted">Not selected</em></span>
              <span class="bsb-p num">&mdash;</span>
            </li>
            <li data-slot="computer">
              <span class="bsb-l">Computer</span>
              <span class="bsb-r"><em class="muted">Not selected</em></span>
              <span class="bsb-p num">&mdash;</span>
            </li>
          </ol>
          <div class="bsb-includes">
            <div class="bsb-includes-hd"><i class="fa fa-gift"></i>Included free</div>
            <div class="bsb-includes-rows">
              <div><span>Wifi / BT card</span><b class="num">&pound;40</b></div>
              <div><span>Speakers</span><b class="num">&pound;20</b></div>
              <div><span>Premium cables</span><b class="num" id="mmb-cables">&mdash;</b></div>
              <div><span>UK delivery</span><b class="num">&pound;20</b></div>
            </div>
          </div>
          <div class="bsb-totals">
            <div class="bsb-row"><span>Subtotal</span><b class="num" id="mmb-subtotal">&mdash;</b></div>
            <div class="bsb-row" style="color:var(--accent-deep);"><span>Bundle discount</span><b class="num" id="mmb-discount" style="color:var(--accent-deep);">&mdash;</b></div>
            <div class="bsb-row total"><span>Total</span><b class="num" id="mmb-total">&mdash;</b></div>
            <div class="bsb-row" style="color:var(--up);"><span>Total savings</span><b class="num" id="mmb-savings" style="color:var(--up);">&mdash;</b></div>
            <p class="bsb-vat-note">All prices exclude VAT</p>
          </div>
          <button class="btn btn-primary btn-lg bsb-cta is-disabled" id="mmb-cta">Keep building <i class="fa fa-arrow-right"></i></button>
          <div class="bsb-trust" id="mmb-trust">
            <i class="fa fa-shield"></i>
            <span>5-year PC cover &middot; Lifetime UK support &middot; 30-day money-back guarantee</span>
          </div>
        </div>
      </aside>

    </div>
  </div>
</section>

<!-- ===================================================================
     WHAT'S IN EVERY BUNDLE - dark band, 4 inline reassurance icons
     =================================================================== -->
<section class="bundle bundle-includes">
  <div class="container">
    <ul class="bib-row reveal">
      <li><i class="fa fa-plug"></i><span>Free Premium 3m Long Cables</span></li>
      <li><i class="fa fa-wifi"></i><span>Free WiFi / Bluetooth &amp; Speakers</span></li>
      <li><i class="fa fa-truck"></i><span>Free UK Delivery</span></li>
      <li><i class="fa fa-tag"></i><span>Bundle Savings Discount</span></li>
    </ul>
  </div>
</section>

<!-- ===================================================================
     POPULAR STARTING POINTS
     =================================================================== -->
<section class="s-tight" id="starters" style="border-top:1px solid var(--line); border-bottom:1px solid var(--line);">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Popular starting points</h5>
        <h2>Not sure where to start? <span class="display-em">Try one of these.</span></h2>
        <p style="margin-top:12px; max-width:700px;">Four configurations our customers pick most often. These are jumping-off points &mdash; you can customise any of them before checkout, or keep building in the configurator above.</p>
      </div>
    </div>

    <div class="bundle-cards">
      <a href="/products/trader-pc/quad-pyramid-stand/27-iiyama-qhd/" class="bundle-card reveal bundle-card-bg">
        <div class="bundle-card__media">
          <img src="/images/bundles/dual-tra-bundle.jpg" alt="Dual-screen trader bundle">
        </div>
        <div class="bundle-card__lines">
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>2&times; 21.5" Screens</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Dual Horizontal Stand</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Ultra PC</div>
        </div>
        <div class="bundle-card__price">
          <span class="bundle-card__from">From</span>
          <span class="bundle-card__amount">&pound;1,195</span>
        </div>
        <span class="btn btn-primary bundle-card__cta">View bundle <i class="fa fa-arrow-right"></i></span>
      </a>

      <a href="/products/trader-pc/triple-pyramid-stand/24-acer/" class="bundle-card reveal bundle-card-bg" style="transition-delay:.06s">
        <div class="bundle-card__media">
          <img src="/images/bundles/3pyr-tra-bundle.jpg" alt="Triple Pyrmaid Trader bundle">
        </div>
        <div class="bundle-card__lines">
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>3&times; 24" Screens</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Triple Pyramid Stand</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Trader PC</div>
        </div>
        <div class="bundle-card__price">
          <span class="bundle-card__from">From</span>
          <span class="bundle-card__amount">&pound;1,465</span>
        </div>
        <span class="btn btn-primary bundle-card__cta">View bundle <i class="fa fa-arrow-right"></i></span>
      </a>

      <a href="#builder" class="bundle-card reveal bundle-card-bg" style="transition-delay:.12s">
        <div class="bundle-card__media">
          <img src="/images/bundles/triple-pro-bundle.jpg" alt="Triple screen extreme bundle">
        </div>
        <div class="bundle-card__lines">
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>3&times; 21.5" Screens</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Triple Screen Stand</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Extreme PC</div>
        </div>
        <div class="bundle-card__price">
          <span class="bundle-card__from">From</span>
          <span class="bundle-card__amount">&pound;1,600</span>
        </div>
        <span class="btn btn-primary bundle-card__cta">View bundle <i class="fa fa-arrow-right"></i></span>
      </a>

      <a href="#builder" class="bundle-card reveal bundle-card-bg" style="transition-delay:.18s">
        <div class="bundle-card__media">
          <img src="/images/bundles/qpyr-ult-bundle.jpg" alt="Four pyramid Trader Pro bundle">
        </div>
        <div class="bundle-card__lines">
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>4&times; 24" Screens</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Quad Pyramid Stand</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Trader Pro PC</div>
        </div>
        <div class="bundle-card__price">
          <span class="bundle-card__from">From</span>
          <span class="bundle-card__amount">&pound;1,880</span>
        </div>
        <span class="btn btn-primary bundle-card__cta">View bundle <i class="fa fa-arrow-right"></i></span>
      </a>

      <a href="/products/trader-pc/six-stand/24-acer/" class="bundle-card reveal bundle-card-bg" style="transition-delay:.18s">
        <div class="bundle-card__media">
          <img src="/images/bundles/six-ult-bundle.jpg" alt="Six screen trader bundle">
        </div>
        <div class="bundle-card__lines">
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>6&times; 24" Screens</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Six Screen Stand</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Trader PC</div>
        </div>
        <div class="bundle-card__price">
          <span class="bundle-card__from">From</span>
          <span class="bundle-card__amount">&pound;1,925</span>
        </div>
        <span class="btn btn-primary bundle-card__cta">View bundle <i class="fa fa-arrow-right"></i></span>
      </a>

      <a href="#builder" class="bundle-card reveal bundle-card-bg" style="transition-delay:.18s">
        <div class="bundle-card__media">
          <img src="/images/bundles/quad-tra-bundle.jpg" alt="Four Square Trader Pro bundle">
        </div>
        <div class="bundle-card__lines">
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>4&times; 27" QHD Screens</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Quad Square Stand</div>
          <div class="bundle-card__desc"><i class="fa fa-check-circle"></i>Trader Pro PC</div>
        </div>
        <div class="bundle-card__price">
          <span class="bundle-card__from">From</span>
          <span class="bundle-card__amount">&pound;2,140</span>
        </div>
        <span class="btn btn-primary bundle-card__cta">View bundle <i class="fa fa-arrow-right"></i></span>
      </a>
    </div>

    <div style="text-align:center; margin-top:30px;" class="reveal">
      <a href="#builder" class="btn btn-ghost btn-lg">Or build your own from scratch <i class="fa fa-arrow-right"></i></a>
    </div>
  </div>
</section>

<!-- ===================================================================
     FAQ
     =================================================================== -->
<section class="s depth" id="faq">
  <div class="container-narrow">
    <div class="section-head reveal" style="display:block; margin-bottom:38px;">
      <h5>Bundle questions</h5>
      <h2>The questions we get <span class="display-em">before ordering</span>, answered.</h2>
      <p style="margin-top:12px;">17 years of bundle conversations has given us a solid list. If yours isn&rsquo;t here, <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">call us</a> on 0330 223 66 55.</p>
    </div>

    <div class="faq-list reveal">
      <details class="faq-item" open>
        <summary>Are all of your stands, screens and computers compatible?</summary>
        <div class="faq-body">
          <p>Yes, every Synergy stand we sell is compatible with every monitor we offer and all of our PC options. Pick any stand layout that you like with any screen size / resolution that fits your needs, and when you select the PC we automatically select a graphics setup that is fully capable of powering your monitor array.</p>
          <p>This is why bundles are so popular, we have taken all the headaches and guesswork out of it for you.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>What&rsquo;s actually included for free in a bundle?</summary>
        <div class="faq-body">
          <p>Every bundle ships with: a premium quality digital 3m long video cable for each screen, a free high speed wifi card which also has Bluetooth functionality, a free set of desktop speakers, and free UK mainland delivery.</p><p>On top of that we also automatically apply a bundle discount ranging from &pound;25 &ndash; &pound;100 off depending on the bundle size.</p>
          <p>On a 6-screen bundle this works out to around &pound;270 of included value vs. sourcing the same parts separately.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Can I change a bundle after I&rsquo;ve picked one of your starting points?</summary>
        <div class="faq-body">
          <p>Absolutely. The popular starting points are just that, starting points. You can swap the stand, change screen size, pick a different PC, bump the RAM, switch the CPU, anything. Nothing is locked in until you complete checkout.</p>
          <p>An easy way to customise your bundle is to use the bundle configurator towards the top pof this page, just pick a stand layout to get started then move through the three steps.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>I already have a couple of screens, can I use them with a new bundle?</summary>
        <div class="faq-body">
          <p>If your screens have a digital monitor port (HDMI, DisplayPort or DVI) then they should be compatible with any of our computers. To mount them on a Synergy Stand they would also need a VESA interface, the four screw holes on the back of the screen.</p>
          <p>If you want to buy a bundle with fewer screens the best thing to do is configure the bundle normally, add it to the basket and then you can reduce the number of screens in the basket before placing the order.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How long does a bundle take to build and deliver?</summary>
        <div class="faq-body">
          <p>We build all computers up for each order, this build and testing process takes around 4 - 5 days to complete. This includes a 32 hour stress-test for the computer. We then dispatch everything together on a pre-12 next working day basis for all UK mainland orders. Basically if you order on a Monday then you are usually accepting the delivery the following Monday.</p>
          <p>We do show a delivery estimate on the bundle pages and in the checkout, this is usually very accurate. If you need a specific delivery date you can just let us know, we can usually hit it with enough notice.</p>
          <p>Also, everything arrives in one delivery, so you&rsquo;re not chasing boxes from three different suppliers.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Do you delivery bundles internationally?</summary>
        <div class="faq-body">
          <p>Yes, we have delivered bundles internationally a number of times. Some locations combined with a higher number of screens can result in very high delivery costs. Whilst we can always provide a quote we may sometimes advise that you purchase screens locally with Multiple Monitors suppling just the computer and the stand.</p>
          <p>Doing this can result in large delivery cost savings. In these cases we are happy to advise on computer and stand compatibility and we can also supply the right cables if required. If you are unsure then talk to us before ordering.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Can I save my bundle and share it with someone else?</summary>
        <div class="faq-body">
          <p>Yes. Once yolu have selected your stand, screens and PC you will arrive on the bundle page. In the top section of this page is a 'Copy Bundle Link' button which you can then paste into an email or save it to your browser favourites.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>What does the warranty cover?</summary>
        <div class="faq-body">
          <p>Every PC in a bundle is covered by an enhanced onsite warranty for 1 - 3 years which runs alongside a 5-year hardware labour warranty. You also get lifetime UK phone, email and remote desktop support. The Synergy Stand carries a separate lifetime warranty on all parts, and the screens are covered by the screen original manufacturer&rsquo;s warranty (which is typically 3 years).</p>
          <p>You can view our full <a href="/pages/warranty/">warranty details here</a>.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Can I return a bundle if it&rsquo;s not right?</summary>
        <div class="faq-body">
          <p>We offer a 30-day money-back guarantee on all of our computer sales. Stands and screens can be returned up to 14 days from point of delivery. You can view our full <a href="/pages/returns/">returns policy here</a>.</p>
        </div>
      </details>
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
        <h5>Still deciding?</h5>
        <h2>Talk to <em>Darren</em> &mdash; the founder, not a call centre.</h2>
        <p>Seventeen years of speccing these bundles means most customers&rsquo; questions have pretty direct answers. Tell him the screens you want, the platforms you run, your budget, and he&rsquo;ll spec it properly &mdash; and tell you if a smaller bundle would do the job.</p>
        <div class="darren-ctas">
          <a href="tel:03302236655" class="btn btn-primary btn-lg"><i class="fa fa-phone"></i>0330 223 66 55</a>
          <a href="#" class="btn btn-ghost btn-lg js-book-call"><i class="fa fa-calendar"></i>Book a 15-min call</a>
        </div>
        <div class="darren-sig">&mdash; Darren Atkinson, founder, Multiple Monitors Ltd</div>
      </div>
    </div>
  </div>
</section>

</div><!-- /.mm-site -->

<!-- ===================================================================
     BUNDLE BUILDER - data + render
     BUNDLE_CONFIG below is emitted from ASP with live DB prices.
     The render and state machine are ported verbatim from
     redesign/bundles.html with two changes:
       1. item.id values are now numeric (the real products.idProduct).
       2. The final CTA builds its URL from state.computer.cta so non-
          Trader PCs fall through to their legacy product pages.
     =================================================================== -->
<script>
(function(){
  const MMB_PRESELECT = { sid: <%= mmPreSid %>, mid: <%= mmPreMid %>, cid: <%= mmPreCid %> };
  const MMB_EDIT_STAGE = (document.getElementById('mmb-stepper').getAttribute('data-edit-stage') || '');

  const BUNDLE_CONFIG = {
    stands: [
<% Dim i
   For i = 0 To UBound(mmStands)     : Call mmEmitStand(mmStands(i))     : Next %>
    ],
    screens: [
<% For i = 0 To UBound(mmScreens)    : Call mmEmitScreen(mmScreens(i))   : Next %>
    ],
    computers: [
<% For i = 0 To UBound(mmComputers)  : Call mmEmitComputer(mmComputers(i)) : Next %>
    ]
  };

  const STAGES = {
    stand: {
      title: 'Start with the stand. <span class="muted">Everything else flexes around it.</span>',
      next:  { goto:'screens',  label:'Next: Pick Screens <i class="fa fa-arrow-right"></i>' },
      cols:  4,
      items: () => BUNDLE_CONFIG.stands,
      meta:  s => '&pound;' + s.price + ' &middot; ' + s.screens + ' screens',
    },
    screens: {
      title: 'Now pick your screens. <span class="muted">Same model across the stand.</span>',
      next:  { goto:'computer', label:'Next: Pick Computer <i class="fa fa-arrow-right"></i>' },
      cols:  3,
      items: () => BUNDLE_CONFIG.screens,
      meta:  (s, state) => {
        const qty = state.stand ? state.stand.screens : 1;
        return '&pound;' + s.price + ' each &middot; ' + qty + ' &times; &pound;' + (s.price * qty);
      },
    },
    computer: {
      title: 'Finally, the PC. <span class="muted">Graphics matched to your stand.</span>',
      next:  null,
      cols:  2,
      items: () => BUNDLE_CONFIG.computers,
      meta:  (c, state) => {
        const up = computerUpgrade(c, state.stand);
        const total = c.price + up;
        if (!up) return '&pound;' + total;
        const kit = state.stand.screens >= 7 ? '8-screen' : '6-screen';
        return '&pound;' + total + ' &middot; ' + kit;
      },
    },
  };

  const state = { stand:null, screens:null, computer:null, view:null };
  const $ = id => document.getElementById(id);
  const fmt = n => '&pound;' + n.toLocaleString('en-GB');

  function currentStage() {
    if (state.view === 'stand')                      return 'stand';
    if (state.view === 'screens'  && state.stand)    return 'screens';
    if (state.view === 'computer' && state.screens)  return 'computer';
    if (!state.stand)    return 'stand';
    if (!state.screens)  return 'screens';
    if (!state.computer) return 'computer';
    return 'done';
  }

  function isComplete() {
    return !!(state.stand && state.screens && state.computer);
  }

  function computerUpgrade(computer, stand) {
    if (!computer || !stand) return 0;
    if (stand.screens >= 7) return computer.eight || 0;
    if (stand.screens >= 5) return computer.six   || 0;
    return 0;
  }
  function computerPrice(computer, stand) {
    if (!computer) return 0;
    return computer.price + computerUpgrade(computer, stand);
  }

  function subtotal() {
    let t = 0;
    if (state.stand)    t += state.stand.price;
    if (state.stand && state.screens) t += state.screens.price * state.stand.screens;
    if (state.computer) t += computerPrice(state.computer, state.stand);
    return t;
  }

  function renderGrid() {
    const stage = currentStage();
    const cfg   = STAGES[stage === 'done' ? 'computer' : stage];
    const grid  = $('mmb-grid');
    const selectedId = state[stage === 'done' ? 'computer' : stage]
      ? state[stage === 'done' ? 'computer' : stage].id : null;

    $('mmb-stage-title').innerHTML = cfg.title;

    const stageKey = stage === 'done' ? 'computer' : stage;
    const picked   = state[stageKey];
    const nextBtn  = $('mmb-next-btn');
    if (cfg.next && picked) {
      nextBtn.innerHTML     = cfg.next.label;
      nextBtn.dataset.goto  = cfg.next.goto;
      nextBtn.style.display = '';
    } else {
      nextBtn.style.display = 'none';
      nextBtn.innerHTML     = '';
      nextBtn.removeAttribute('data-goto');
    }

    grid.className = 'stand-grid cols-' + (cfg.cols || 4);

    grid.innerHTML = cfg.items().map(item => {
      const sel  = item.id === selectedId ? ' is-selected' : '';
      const meta = cfg.meta(item, state);

      if (stageKey === 'computer' || stageKey === 'screens') {
        const bullets = [item.desc1, item.desc2, item.desc3]
          .filter(Boolean)
          .map(d => '<li>' + d + '</li>')
          .join('');
        return (
          '<button type="button" class="stand-card ' +  stageKey + ' is-detail' + sel + '" data-id="' + item.id + '">' +
            '<div class="detail-top">' +
              '<div class="detail-vis"><img src="' + item.img + '" alt="' + item.name + '"></div>' +
              (bullets ? '<ul class="detail-bullets">' + bullets + '</ul>' : '') +
            '</div>' +
            '<div class="detail-bottom">' +
              '<div class="detail-name">' + item.name + '</div>' +
              '<div class="detail-price num">' + meta + '</div>' +
            '</div>' +
          '</button>'
        );
      }

      const desc = item.description ? '<div class="stand-desc">' + item.description + '</div>' : '';
      return (
        '<button type="button" class="stand-card' + sel + '" data-id="' + item.id + '">' +
          '<div class="stand-vis"><img src="' + item.img + '" alt="' + item.name + '"></div>' +
          '<div class="stand-name">' + item.name + '</div>' +
          desc +
          '<div class="stand-meta num">' + meta + '</div>' +
        '</button>'
      );
    }).join('');

    $('mmb-info').style.display = (isComplete()) ? '' : 'none';
  }

  function renderStepper() {
    const stage = currentStage();
    document.querySelectorAll('#mmb-stepper .step').forEach(li => {
      const which  = li.dataset.step;
      const picked = state[which];
      li.classList.toggle('is-current', stage === which || (stage === 'done' && which === 'computer'));
      li.querySelector('.step-val').innerHTML = picked ? picked.name : '&mdash;';
    });
  }

  function renderSidebar() {
    const viz = $('mmb-viz');
    if (state.stand && state.screens) {
      const arrSrc = '/images/bundles/' + state.stand.arrayimg + '-' + state.screens.arrayimg + '-blg.png';
      const arrAlt = state.stand.name + ' with ' + state.screens.name;
      let html = '<img src="' + arrSrc + '" alt="' + arrAlt + '" style="max-width:170px; max-height:130px; object-fit:contain;">';
      if (state.computer && state.computer.bunimg) {
        html += '<img src="' + state.computer.bunimg + '" alt="' + state.computer.name + '" style="max-height:130px; width:auto; margin-left:12px; object-fit:contain;">';
      }
      viz.innerHTML = html;
    } else if (state.stand) {
      viz.innerHTML = '<img src="' + state.stand.img + '" alt="' + state.stand.name + '" style="max-width:180px; max-height:140px; object-fit:contain;">';
    }

    document.querySelectorAll('#mmb-list li').forEach(li => {
      const slot  = li.dataset.slot;
      const pick  = state[slot];
      const rEl   = li.querySelector('.bsb-r');
      const pEl   = li.querySelector('.bsb-p');
      if (!pick) {
        li.classList.remove('is-done');
        rEl.innerHTML = '<em class="muted">Not selected</em>';
        pEl.innerHTML = '&mdash;';
      } else if (slot === 'screens') {
        const qty = state.stand ? state.stand.screens : 1;
        li.classList.add('is-done');
        rEl.textContent = qty + ' × ' + pick.name;
        pEl.innerHTML   = fmt(pick.price * qty);
      } else if (slot === 'computer') {
        li.classList.add('is-done');
        rEl.textContent = pick.name;
        pEl.innerHTML   = fmt(computerPrice(pick, state.stand));
      } else {
        li.classList.add('is-done');
        rEl.textContent = pick.name;
        pEl.innerHTML   = fmt(pick.price);
      }
    });

    const sub      = subtotal();
    const cables   = state.stand ? 15 * state.stand.screens : 0;
    const discount = state.stand ? (state.stand.discount || 0) : 0;
    const savings  = state.stand ? (cables + 80 + discount) : 0;
    const total    = Math.max(0, sub - discount);

    $('mmb-cables').innerHTML   = state.stand ? fmt(cables)          : '&pound;30';
    $('mmb-subtotal').innerHTML = sub         ? fmt(sub)             : '&mdash;';
    $('mmb-discount').innerHTML = discount    ? '&minus; ' + fmt(discount) : '&pound;25 - &pound;100';
    $('mmb-total').innerHTML    = sub         ? fmt(total)           : '&mdash;';
    $('mmb-savings').innerHTML  = state.stand ? fmt(savings)         : '&pound;135 - &pound;300';

    const done = [state.stand, state.screens, state.computer].filter(Boolean).length;
    $('mmb-pct').textContent = Math.round((done / 3) * 100) + '%';

    // Sidebar CTA. Label advertises the deepest unpicked slot (or
    // "Configure..." when complete). Enabled state is driven off the
    // currently viewed stage's pick - the user has to make a selection
    // on the stage they're looking at before the button activates, even
    // if that stage isn't the deepest unpicked one.
    const cta = $('mmb-cta');
    cta.classList.remove('is-disabled');
    delete cta.dataset.goto;

    let ctaHtml, ctaGoto;
    if (isComplete()) {
      ctaHtml = 'Configure PC &amp; Order Bundle <i class="fa fa-arrow-right"></i>';
      ctaGoto = null;
    } else if (!state.stand) {
      ctaHtml = 'Pick a stand <i class="fa fa-arrow-right"></i>';
      ctaGoto = null;
    } else if (!state.screens) {
      ctaHtml = 'Next: Pick Screens <i class="fa fa-arrow-right"></i>';
      ctaGoto = 'screens';
    } else {
      ctaHtml = 'Next: Pick Computer <i class="fa fa-arrow-right"></i>';
      ctaGoto = 'computer';
    }

    const viewStage = currentStage();
    const viewSlot  = viewStage === 'done' ? 'computer' : viewStage;
    const viewReady = !!state[viewSlot];

    cta.innerHTML = ctaHtml;
    if (!viewReady) {
      cta.classList.add('is-disabled');
    } else if (ctaGoto) {
      cta.dataset.goto = ctaGoto;
    }

     // Inline info-row CTA on the computer stage — mirrors the sidebar
    // CTA's state so a user looking at PC cards has the next-step
    // action right under the grid, not just in the right rail.
    const infoCta = $('mmb-info-cta');
    const trust = $('mmb-trust');
    if (infoCta) {
      if (isComplete()) {
        infoCta.classList.remove('is-disabled');
        infoCta.innerHTML = 'Configure PC &amp; Order Bundle <i class="fa fa-arrow-right"></i>';
        trust.classList.add('bsb-trust--next');
        trust.innerHTML =
          '<ul>' +
            '<li><i class="fa fa-check-circle"></i><span>Now pick your CPU, RAM, &amp; other options</span></li>' +
            '<li><i class="fa fa-check-circle"></i><span>Bundle savings of ' + fmt(savings) + ' stay locked in</span></li>' +
          '</ul>';
      } else {
        infoCta.classList.add('is-disabled');
        infoCta.innerHTML = 'Pick a PC to continue <i class="fa fa-arrow-right"></i>';
        trust.classList.remove('bsb-trust--next');
        trust.innerHTML =
          '<i class="fa fa-shield"></i>' +
          '<span>5-year PC cover &middot; Lifetime UK support &middot; 30-day money-back guarantee</span>';
      }
    }

    // Inline savings figure — only meaningful once a stand is picked
    // (which is always true on the computer stage, but guarded for
    // safety in case the banner is ever shown earlier).
    const infoSavings = $('mmb-info-savings');
    if (infoSavings) {
      infoSavings.innerHTML = state.stand ? ' and ' + fmt(savings) + ' saving' : '';
    }
  }

  function render() {
    renderGrid();
    renderStepper();
    renderSidebar();
  }

  document.getElementById('mmb-next-btn').addEventListener('click', function(){
    const goto = this.dataset.goto;
    if (!goto) return;
    state.view = goto;
    render();
  });

  document.getElementById('mmb-stepper').addEventListener('click', function(e){
    const li = e.target.closest('.step');
    if (!li) return;
    const which = li.dataset.step;
    if (which === 'screens'  && !state.stand)   return;
    if (which === 'computer' && !state.screens) return;
    state.view = which;
    render();
  });

  // Pick an item for the current stage. IDs are numeric, so compare
  // against Number(data-id) -- the attribute is always a string.
  document.getElementById('mmb-grid').addEventListener('click', function(e){
    const btn = e.target.closest('[data-id]');
    if (!btn) return;
    const id    = Number(btn.dataset.id);
    const stage = currentStage() === 'done' ? 'computer' : currentStage();

    if (stage === 'stand') {
      state.stand = BUNDLE_CONFIG.stands.find(s => s.id === id) || state.stand;
    } else if (stage === 'screens') {
      state.screens = BUNDLE_CONFIG.screens.find(s => s.id === id) || state.screens;
    } else if (stage === 'computer') {
      state.computer = BUNDLE_CONFIG.computers.find(c => c.id === id) || null;
    }
    state.view = stage;
    render();
  });

  // Final CTA -- compose the per-computer target URL.
  // Trader PC (id 333) uses the canonical slug URL
  //   /products/trader-pc/<stand-slug>/<monitor-slug>/
  // Other PCs still use the legacy ?sid=&mid=&cid= form because
  // their bundle end-pages haven't been rebuilt yet.
  function gotoBundle(e) {
    e.preventDefault();
    if (!isComplete()) return;
    if (state.computer.id === 333 && state.stand.slug && state.screens.slug) {
      window.location = state.computer.cta + state.stand.slug + '/' + state.screens.slug + '/#bp-picks';
      return;
    }
    const sep = state.computer.cta.indexOf('?') > -1 ? '&' : '?';
    window.location = state.computer.cta + sep +
      'sid=' + state.stand.id +
      '&mid=' + state.screens.id +
      '&cid=' + state.computer.id;
  }
  // Sidebar CTA: when bundle isn't complete, the renderSidebar block
  // stashes the next stage on dataset.goto so the same button can
  // advance the view rather than no-op. When complete, dataset.goto is
  // cleared and we fall through to gotoBundle.
  document.getElementById('mmb-cta').addEventListener('click', function(e){
    if (this.classList.contains('is-disabled')) { e.preventDefault(); return; }
    const goto = this.dataset.goto;
    if (goto) {
      e.preventDefault();
      state.view = goto;
      render();
      return;
    }
    gotoBundle(e);
  });
  document.getElementById('mmb-info-cta').addEventListener('click', gotoBundle);

  // Scroll-reveal animation -- same pattern as the redesigned
  // stands/home pages.
  const els = document.querySelectorAll('.reveal');
  if ('IntersectionObserver' in window) {
    const io = new IntersectionObserver(entries => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('is-in');
          io.unobserve(entry.target);
        }
      });
    }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
    els.forEach(el => io.observe(el));
  } else {
    els.forEach(el => el.classList.add('is-in'));
  }

  if (MMB_PRESELECT.sid) {
    const s = BUNDLE_CONFIG.stands.find(x => x.id === MMB_PRESELECT.sid);
    if (s) state.stand = s;
  }
  if (state.stand && MMB_PRESELECT.mid) {
    const m = BUNDLE_CONFIG.screens.find(x => x.id === MMB_PRESELECT.mid);
    if (m) state.screens = m;
  }
  if (state.screens && MMB_PRESELECT.cid) {
    const c = BUNDLE_CONFIG.computers.find(x => x.id === MMB_PRESELECT.cid);
    if (c) state.computer = c;
  }

  if (state.computer)                   state.view = 'computer';
  else if (state.screens)                state.view = 'screens';
  else if (state.stand)                state.view = 'stand';

  // ?edit=stand|screens|computer (used by bp-picks "Change" links on the
  // final bundle page) overrides the default deepest-slot view so the
  // builder opens on the panel the user clicked Change for.
  if (MMB_EDIT_STAGE === 'stand' || MMB_EDIT_STAGE === 'screens' || MMB_EDIT_STAGE === 'computer') {
    state.view = MMB_EDIT_STAGE;
  }

  render();

  if (MMB_PRESELECT.sid) {
    const el = document.getElementById('builder');
    if (el) el.scrollIntoView({ behavior: 'auto', block: 'start' });
  }
})();
</script>

<!--#include file="footer_wrapper.asp"-->
