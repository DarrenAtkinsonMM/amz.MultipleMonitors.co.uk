<%
' ============================================================
' CUSTOMCAT-stands.asp
' 2026 redesign — Synergy Stands category page.
' Rewritten in place (was 602 lines of ProductCart scaffolding);
' now a mostly-static page with a small direct-DB query for the
' 12 stand tiles grouped by screen count. See
' /category-redesign-plan.md at repo root for the approach.
' ============================================================
%>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "customcat-stands.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<%
' ------------------------------------------------------------
' Load stands products (category 5) in one pass.
' Fields: 0=idProduct 1=sku 2=description 3=price
'         4=smallImageUrl 5=pcUrl
' ------------------------------------------------------------
Dim mmStandsSql, mmStandsRs, mmStandsRows, mmStandsCount
mmStandsCount = 0

mmStandsSql = "SELECT p.idProduct, p.sku, p.description, p.price, " & _
              "p.imageUrl, p.pcUrl " & _
              "FROM products p " & _
              "INNER JOIN categories_products cp ON p.idProduct = cp.idProduct " & _
              "WHERE cp.idCategory = 5 " & _
              "  AND p.active = -1 AND p.configOnly = 0 AND p.removed = 0 " & _
              "ORDER BY p.price ASC"

Set mmStandsRs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
mmStandsRs.Open mmStandsSql, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If err.number <> 0 Then
    On Error Goto 0
    call LogErrorToDatabase()
    Set mmStandsRs = Nothing
    call closeDB()
    Response.Redirect "techErr.asp?err=" & pcStrCustRefID
End If
On Error Goto 0

If Not mmStandsRs.EOF Then
    mmStandsRows = mmStandsRs.GetRows()
    mmStandsCount = UBound(mmStandsRows, 2) + 1
End If
Set mmStandsRs = Nothing

' ------------------------------------------------------------
' Screen count from the first word of the product description.
' Dual=2, Triple=3, Quad=4, Five=5, Six=6, Eight=8.
' Returns 0 for anything unrecognised (product will be skipped).
' ------------------------------------------------------------
Function mmStandScreenCount(ByVal descr)
    Dim firstWord, pos, s
    s = LCase(Trim(descr & ""))
    pos = InStr(s, " ")
    If pos > 0 Then
        firstWord = Left(s, pos - 1)
    Else
        firstWord = s
    End If
    Select Case firstWord
        Case "dual"   : mmStandScreenCount = 2
        Case "triple" : mmStandScreenCount = 3
        Case "quad"   : mmStandScreenCount = 4
        Case "five"   : mmStandScreenCount = 5
        Case "six"    : mmStandScreenCount = 6
        Case "eight"  : mmStandScreenCount = 8
        Case Else     : mmStandScreenCount = 0
    End Select
End Function

' ------------------------------------------------------------
' Style label from SKU suffix letters (everything after the
' first digit). v/h/p/s/sp/rp/r map to Vertical/Horizontal/
' Pyramid/Square/Pole/Pole/Side-by-side. For the 8-screen
' stand, a lone 'r' suffix means "2-over-2 quad".
' Returns "" if the SKU doesn't match a known style.
' ------------------------------------------------------------
Function mmStandStyle(ByVal sku, ByVal screenCount)
    Dim s, i, ch, tail
    s = LCase(Trim(sku & ""))
    tail = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            tail = Mid(s, i + 1)
            Exit For
        End If
    Next
    Select Case tail
        Case "v"  : mmStandStyle = "Vertical"
        Case "h"  : mmStandStyle = "Horizontal"
        Case "p"  : mmStandStyle = "Pyramid"
        Case "s"  : mmStandStyle = "Square"
        Case "sp" : mmStandStyle = "Pole"
        Case "rp" : mmStandStyle = "Pole"
        Case "r"
            If screenCount = 8 Then
                mmStandStyle = "4-over-4"
            Else
                mmStandStyle = "Side-by-side"
            End If
        Case Else
            mmStandStyle = ""
    End Select
End Function

' ------------------------------------------------------------
' Render every card whose screen count is in the allowed list.
' allowed is a comma-separated string like "2,3" or "5,6,8".
' ------------------------------------------------------------
Sub mmRenderStandGroup(ByVal allowed)
    Dim i, screens, idProduct, sku, descr, price, img, purl
    Dim eyebrow, style, href, imgSrc, priceDisp, altText, delayIdx

    If mmStandsCount < 1 Then Exit Sub

    delayIdx = 0
    For i = 0 To mmStandsCount - 1
        idProduct = mmStandsRows(0, i)
        sku       = mmStandsRows(1, i) & ""
        descr     = mmStandsRows(2, i) & ""
        price     = mmStandsRows(3, i)
        img       = mmStandsRows(4, i) & ""
        purl      = mmStandsRows(5, i) & ""

        screens = mmStandScreenCount(descr)
        If screens > 0 And InStr("," & allowed & ",", "," & screens & ",") > 0 Then

            style = mmStandStyle(sku, screens)
            If style <> "" Then
                eyebrow = screens & "-Screen &middot; " & style
            Else
                eyebrow = screens & "-Screen"
            End If

            If img <> "" Then
                imgSrc = "/shop/pc/catalog/" & img
            Else
                imgSrc = "/shop/pc/catalog/no_image.gif"
            End If

            If purl <> "" Then
                href = "/products/" & purl & "/"
            Else
                href = "/shop/pc/viewPrd.asp?idproduct=" & idProduct
            End If

            altText   = Server.HTMLEncode(descr)
            priceDisp = scCursign & money(price / 1.2)

            Response.Write "<a href=""" & href & """ class=""bundle-card reveal"""
            If delayIdx > 0 Then
                Response.Write " style=""transition-delay:." & Right("0" & (delayIdx * 6), 2) & "s"""
            End If
            Response.Write ">" & vbCrLf
            Response.Write "  <div class=""bundle-card__media"">" & vbCrLf
            Response.Write "    <img src=""" & imgSrc & """ alt=""" & altText & """>" & vbCrLf
            Response.Write "  </div>" & vbCrLf
            Response.Write "  <div class=""bundle-card__body"">" & vbCrLf
            Response.Write "    <div class=""bundle-card__eyebrow"">" & eyebrow & "</div>" & vbCrLf
            Response.Write "    <h4 class=""bundle-card__title"">" & altText & "</h4>" & vbCrLf
            Response.Write "    <div class=""bundle-card__price"">" & vbCrLf
            Response.Write "      <span class=""bundle-card__from"">From</span>" & vbCrLf
            Response.Write "      <span class=""bundle-card__amount"">" & priceDisp & "</span>" & vbCrLf
            Response.Write "    </div>" & vbCrLf
            Response.Write "    <span class=""btn btn-primary bundle-card__cta"">View stand <i class=""fa fa-arrow-right""></i></span>" & vbCrLf
            Response.Write "  </div>" & vbCrLf
            Response.Write "</a>" & vbCrLf

            delayIdx = delayIdx + 1
        End If
    Next
End Sub
%>
<!--#include file="header_wrapper.asp"-->

<div class="mm-site">

<!-- ===================================================================
     HERO
     =================================================================== -->
<section class="hero">
  <div class="container">
    <div class="hero-grid">
      <div class="reveal">
        <div class="eyebrow">Synergy Stands &middot; UK designed &amp; manufactured</div>
        <h1>
          Synergy Stands, our own UK-designed, UK-manufactured <em>modular monitor mounts</em>.
        </h1>
        <p class="lead">
          Developed by us, manufactured in the UK to our specifications. A modular system that scales from two to six screens on a single assembly. Built to hold up day after day, with thousands in use across trader desks, operations rooms, and home offices.
        </p>
        <div class="hero-ctas">
          <a href="#range" class="btn btn-primary btn-lg">See the range <i class="fa fa-arrow-right"></i></a>
        </div>
        <div class="hero-mini">
          <div class="item"><i class="fa fa-industry"></i><span>UK-designed &amp; UK-made</span></div>
          <div class="item"><i class="fa fa-th-large"></i><span>Modular &middot; 2 to 8 screen packages</span></div>
          <div class="item"><i class="fa fa-clock-o"></i><span>Sold since 2016</span></div>
        </div>
      </div>

      <div class="hero-visual reveal" style="transition-delay:.1s">
        <img src="/images/pages/ss-6r.png" alt="Six-screen Synergy Stand in a curved configuration" />
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
     BENEFIT CARDS
     =================================================================== -->
<section class="s-tight" style="border-top:1px solid var(--line); border-bottom:1px solid var(--line);">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>What makes a Synergy Stand different</h5>
        <h2>Not a generic import. <span class="display-em">Our own product, made properly.</span></h2>
        <p style="max-width:760px; margin-top:12px;">Most multi-monitor stands on the market are rebranded imports, priced on volume. Ours aren&rsquo;t. Four things set a Synergy Stand apart from what you&rsquo;ll find on Amazon.</p>
      </div>
    </div>

    <div class="pillars">
      <div class="pillar reveal">
        <div class="icon"><i class="fa fa-industry"></i></div>
        <h4>UK-designed &amp; UK-made</h4>
        <p>We designed the Synergy Stand. These are not cheap rebranded imports. We work with a specialist UK design and manufacturing partner to build what the market didn&rsquo;t.</p>
        <div class="tag">UK DESIGN &amp; BUILD</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.06s">
        <div class="icon"><i class="fa fa-th-large"></i></div>
        <h4>Modular, 2 to 8 screens</h4>
        <p>Start with two screens. Add arms and mounts as your needs grow, three, four, five and six monitor capable mounts using a single central column. No wasted spend if you scale up.</p>
        <div class="tag">MODULAR SYSTEM</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.12s">
        <div class="icon"><i class="fa fa-shield"></i></div>
        <h4>All-steel, built for daily use</h4>
        <p>Every part is steel, not a cheap metal frame with plastic joints. Plastic bends, flexes, and fails under the weight of real screens. Synergy Stands are built for desks that run ten hours a day, every day.</p>
        <div class="tag">ALL-STEEL</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.18s">
        <div class="icon"><i class="fa fa-sliders"></i></div>
        <h4>Adjustability that works</h4>
        <p>Height position, arm hinge, horizontal slide, pivot, tilt and 30&nbsp;mm of fine height adjustment at every arm mount. Six degrees of freedom per screen and everything locks solid once positioned.</p>
        <div class="tag">FULL ADJUSTMENT</div>
      </div>
    </div>

  </div>
</section>

<!-- ===================================================================
     DESIGN & MANUFACTURING STORY
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="hero-grid">
      <div class="reveal">
        <div class="eyebrow">Where the Synergy Stand came from</div>
        <h2>Designed by us. <span class="display-em">Made in the UK.</span> Refined since 2016.</h2>
        <p class="lead">
          At Multiple Monitors we spent a long time battling with inadequate and expensive stands. After years of frustration we developed the Synergy Stand range. The result of a decade-plus collaboration with a specialist UK design and manufacturing team, producing the stands we knew the market needed, but nobody was making.
        </p>
        <p style="color:var(--slate); margin-top:14px; max-width:640px;">
          Every stand we ship is manufactured in the UK to our specifications and packaged in our workshop. Multiple generations of refinement, driven by real customer feedback, have gone into the system you buy today.
        </p>
        <div class="hero-mini" style="margin-top:22px;">
          <div class="item"><i class="fa fa-check"></i><span>Our own design</span></div>
          <div class="item"><i class="fa fa-check"></i><span>UK manufactured</span></div>
          <div class="item"><i class="fa fa-check"></i><span>10+ years of real-world refinement</span></div>
        </div>
      </div>

      <div class="hero-visual reveal" style="transition-delay:.1s">
        <img src="/images/pages/ss-4p.png" alt="Quad Pyramid Synergy Stand &mdash; UK-designed modular assembly" />
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     MODULAR UPGRADE PATH
     =================================================================== -->
<section class="bundle">
  <div class="container">
    <div class="bundle-grid">
      <div class="reveal">
        <h5>The modular system</h5>
        <h2>Start small. Scale up. <em>Don&rsquo;t buy twice.</em></h2>
        <p>We often hear from our customers things like <em>&lsquo;wouldn&rsquo;t it be easier if I just had one more screen&rsquo;</em>. We hear it a lot. The Synergy Stand is a modular system, not a fixed product, buy one today, buy the extra parts as you need them.</p>
        <p style="color:#C7D2DF; margin-top:14px;">Starting with two screens? The same base stand accepts additional arms. Scale up to four, five or six as your needs grow &mdash; no need to buy a whole new stand.</p>
        <div class="bundle-pills" style="margin-top:20px;">
          <span class="bundle-pill"><i class="fa fa-check"></i>One base, multiple configurations</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Add arms &amp; mounts later</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Same parts, always in stock</span>
        </div>
      </div>

      <div class="reveal" style="transition-delay:.1s">
        <div class="save-card">
          <span class="save-tag">Scale path</span>
          <div class="kicker">The same base, three configurations</div>
          <div class="breakdown" style="margin-top:6px; gap:14px;">
            <div class="r" style="align-items:center;">
              <span style="display:flex; align-items:center; gap:12px;">
                <img src="/shop/pc/catalog/2h-front-angle-thm.jpg" alt="Dual Synergy Stand" style="width:56px; height:56px; object-fit:contain; background:#fff; border-radius:4px;">
                <span><b style="color:var(--ink);">Start</b><br><small style="color:var(--muted);">2 screens</small></span>
              </span>
              <b>Dual Stand</b>
            </div>
            <div class="r" style="align-items:center;">
              <span style="display:flex; align-items:center; gap:12px;">
                <img src="/shop/pc/catalog/4s-front-angle-thm.jpg" alt="Quad Square Synergy Stand" style="width:56px; height:56px; object-fit:contain; background:#fff; border-radius:4px;">
                <span><b style="color:var(--ink);">Scale</b><br><small style="color:var(--muted);">4 screens</small></span>
              </span>
              <b>Add arms</b>
            </div>
            <div class="r" style="align-items:center;">
              <span style="display:flex; align-items:center; gap:12px;">
                <img src="/shop/pc/catalog/6r-front-angle-thm.jpg" alt="Six-screen Synergy Stand" style="width:56px; height:56px; object-fit:contain; background:#fff; border-radius:4px;">
                <span><b style="color:var(--ink);">Grow</b><br><small style="color:var(--muted);">6 screens</small></span>
              </span>
              <b>Add more</b>
            </div>
            <div class="r total"><span>Same base assembly throughout</span><b>&mdash;</b></div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     28" SCREENS + CURVE
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="hero-grid">
      <div class="hero-visual reveal">
        <img src="/images/pages/ss-3p.png" alt="Triple Pyramid Synergy Stand showing curved layout" />
      </div>
      <div class="reveal" style="transition-delay:.08s">
        <div class="eyebrow">Built for today&rsquo;s bigger screens</div>
        <h2>Supports up to 28&Prime; screens <span class="display-em">with room to curve</span>.</h2>
        <p class="lead">
          Screens keep getting larger. 24&Prime; &amp; 27&Prime; widescreens are now our most popular sizes, and many customers now like to go bigger. We designed the Synergy Stand knowing that trend wasn&rsquo;t going away.
        </p>
        <p style="color:var(--slate); margin-top:14px;">
          Often competitor stands specify &lsquo;up to 24&Prime;&rsquo; which leaves no room to angle the outer screens inward for a proper curved layout. Every Synergy Stand is designed to comfortably mount monitors up to and including 28&Prime; widescreens, and still achieve a gentle curve at full screen count.
        </p>
        <div class="hero-mini" style="margin-top:20px;">
          <div class="item"><i class="fa fa-expand"></i><span>Up to 28&Prime; per screen</span></div>
          <div class="item"><i class="fa fa-refresh"></i><span>Comfortable curve at full count</span></div>
        </div>
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
        <p style="max-width:760px; margin-top:12px;">Adjustability isn&rsquo;t a single thing, it&rsquo;s the difference between a stand you fight and a stand you forget about. Every Synergy Stand gives you six independent degrees of freedom, then locks solid once positioned.</p>
      </div>
    </div>

    <div class="bench-panels">
      <div class="bench-panel reveal">
        <h4>Per-screen adjustment</h4>
        <span class="sub">Six degrees of freedom at every mount</span>
        <div style="margin-top:18px; display:grid; gap:14px;">
          <div style="display:flex; gap:14px; align-items:flex-start;">
            <i class="fa fa-arrows-v" style="color:var(--brand); font-size:18px; margin-top:3px; width:22px;"></i>
            <div><b style="color:var(--ink);">Height position</b><br><small style="color:var(--slate);">Mount arms at any height up the central column.</small></div>
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
     SHARED SPECIFICATIONS
     =================================================================== -->
<section class="s specs">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Shared specifications</h5>
        <h2>Built to the <span class="display-em">same standard</span>.</h2>
        <p style="max-width:760px; margin-top:12px;">Whichever stand you pick, these specifications are the same. Every Synergy Stand shares the same core engineering, no variations in quality, no compromises as screen count scales up.</p>
      </div>
    </div>

    <div class="spec-grid">
      <div class="spec-card reveal">
        <div class="spec-card__icon"><i class="fa fa-crosshairs"></i></div>
        <div class="spec-card__label">Mounting standard</div>
        <div class="spec-card__value">VESA 75&times;75 &amp; 100&times;100</div>
        <div class="spec-card__desc">Fits every monitor we sell. No adapters, no surprises.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.06s">
        <div class="spec-card__icon"><i class="fa fa-cubes"></i></div>
        <div class="spec-card__label">Materials</div>
        <div class="spec-card__value">All-steel throughout</div>
        <div class="spec-card__desc">No weak or load-bearing plastic parts anywhere.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.12s">
        <div class="spec-card__icon"><i class="fa fa-expand"></i></div>
        <div class="spec-card__label">Max screen size</div>
        <div class="spec-card__value">Up to 28&Prime; per mount</div>
        <div class="spec-card__desc">With room to curve the outer screens at full count.</div>
      </div>
      <div class="spec-card reveal">
        <div class="spec-card__icon"><i class="fa fa-sliders"></i></div>
        <div class="spec-card__label">Adjustability</div>
        <div class="spec-card__value">Six degrees of freedom</div>
        <div class="spec-card__desc">Height, arm hinge, horizontal slide, pivot, tilt, 30&nbsp;mm fine adjust.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.06s">
        <div class="spec-card__icon"><i class="fa fa-wrench"></i></div>
        <div class="spec-card__label">Assembly</div>
        <div class="spec-card__value">20&ndash;60 min, no drilling</div>
        <div class="spec-card__desc">All tools included. No wall fixings. No drilling through desks.</div>
      </div>
      <div class="spec-card reveal" style="transition-delay:.12s">
        <div class="spec-card__icon"><i class="fa fa-certificate"></i></div>
        <div class="spec-card__label">Warranty</div>
        <div class="spec-card__value">Lifetime on all parts</div>
        <div class="spec-card__desc">Every stand, every configuration.</div>
      </div>
    </div>

    <div class="spec-box reveal" style="transition-delay:.18s">
      <div class="spec-box__lead">
        <div class="spec-box__icon"><i class="fa fa-archive"></i></div>
        <div>
          <div class="spec-box__label">In the box</div>
          <div class="spec-box__title">Everything you need, one delivery.</div>
        </div>
      </div>
      <div class="spec-chips">
        <span class="spec-chip"><i class="fa fa-check"></i>Assembly guide</span>
        <span class="spec-chip"><i class="fa fa-check"></i>All mounting hardware</span>
        <span class="spec-chip"><i class="fa fa-check"></i>VESA plates &amp; screws</span>
        <span class="spec-chip"><i class="fa fa-check"></i>Cable management</span>
      </div>
    </div>

  </div>
</section>

<!-- ===================================================================
     PRODUCT RANGE — GROUPED BY SCREEN COUNT
     (Dynamic: the only DB-backed section on the page.)
     =================================================================== -->
<section class="s depth" id="range">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>The range</h5>
        <h2>12 Synergy Stands, <span class="display-em">pick your perfect layout</span>.</h2>
        <p style="max-width:760px; margin-top:12px;">All stands in the range share the same core components, you can add arms &amp; mounts later to scale up.</p>
      </div>
    </div>

    <!-- 2- & 3-screen -->
    <div class="range-group reveal" style="margin-top:8px;">
      <div style="display:flex; align-items:baseline; gap:14px; margin-bottom:18px; flex-wrap:wrap;">
        <h3 style="margin:0;">Dual &amp; Triple-screen stands</h3>
        <span class="eyebrow" style="margin:0;">2 - 3 screens</span>
      </div>
      <div class="bundle-cards">
<% mmRenderStandGroup "2,3" %>
      </div>
    </div>

    <!-- 4-screen -->
    <div class="range-group reveal" style="margin-top:56px;">
      <div style="display:flex; align-items:baseline; gap:14px; margin-bottom:18px; flex-wrap:wrap;">
        <h3 style="margin:0;">Quad-screen stands</h3>
        <span class="eyebrow" style="margin:0;">4 screens</span>
      </div>
      <div class="bundle-cards">
<% mmRenderStandGroup "4" %>
      </div>
    </div>

    <!-- 5-, 6- & 8-screen -->
    <div class="range-group reveal" style="margin-top:56px;">
      <div style="display:flex; align-items:baseline; gap:14px; margin-bottom:18px; flex-wrap:wrap;">
        <h3 style="margin:0;">Five, Six &amp; Eight-screen stands</h3>
        <span class="eyebrow" style="margin:0;">5 - 8 screens</span>
      </div>
      <div class="bundle-cards">
<% mmRenderStandGroup "5,6,8" %>
      </div>
    </div>

  </div>
</section>


<!-- ===================================================================
     BUNDLE CROSS-LINK
     =================================================================== -->
<section class="bundle">
  <div class="container">
    <div class="bundle-grid">
      <div class="reveal">
        <h5>Complete monitor arrays &amp; bundles</h5>
        <h2>Need some screens or a PC with your stand? <em>Save money</em> with a monitor array or computer bundle.</h2>
        <p>We offer a range of screens and computers that work perfectly with our Synergy Stands. Monitor arrays come with screens, stand, free cabling and free UK delivery. Bundles with a PC included can save you up to &pound;300.</p>
        <div class="bundle-pills">
          <span class="bundle-pill"><i class="fa fa-check"></i>Free premium long-length cables</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free UK delivery</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free WiFi card<span>*</span></span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free Speakers<span>*</span></span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Auto bundle discount<span>*</span></span>
        </div>
        <div style="display:flex; gap:12px; flex-wrap:wrap;">
          <a href="/display-systems/" class="btn btn-accent btn-lg">See monitor arrays <i class="fa fa-arrow-right"></i></a>
          <a href="/bundles/" class="btn btn-accent btn-lg">See bundle deals <i class="fa fa-arrow-right"></i></a>
        </div>
        <p class="bundle-foot"><span class="bundle-foot__star">*</span>Available on computer bundles only</p>
      </div>
      <div class="reveal" style="transition-delay:.1s">
        <div class="save-card">
          <span class="save-tag">Example &middot; 6-screen bundle</span>
          <div class="kicker">Typical saving vs buying separately</div>
          <div class="big"><small>&pound;</small>270</div>
          <div class="sub">on a six-screen Synergy Stand &plus; 6 screens &plus; Trader PC bundle.</div>
          <div class="breakdown">
            <div class="r"><span>6&thinsp;&times;&thinsp;3m video cables</span><b>&pound;90</b></div>
            <div class="r"><span>Wifi, BT &amp; speakers</span><b>&pound;60</b></div>
            <div class="r"><span>UK mainland delivery</span><b>&pound;20</b></div>
            <div class="r"><span>Bundle discount</span><b>&pound;100</b></div>
            <div class="r total"><span>Total savings</span><b>&minus; &pound;270</b></div>
          </div>
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
      <h5>Common questions</h5>
      <h2>Your Synergy Stand questions <span class="display-em">answered</span>.</h2>
      <p style="margin-top:12px;">If your question isn&rsquo;t listed, <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">call us</a> on 0330 223 66 55 and we will help you out.</p>
    </div>

    <div class="faq-list reveal">

      <details class="faq-item" open>
        <summary>Which screens fit on the stands?</summary>
        <div class="faq-body">
          <p>There are two things to consider when deciding on screens for a Synergy Stand. First is the physical dimensions. For any stand which has arms then the maximum width that can be used is a 28&Prime; widescreen monitor. Anything bigger will not physically fit, anything smaller will fit perfectly.</p>
          <p>The second requirement is the VESA mounting interface. This is basically the four screw holes on the back of the screen. This is how they attach to our mounts. Sometimes called a wall mount interface, you need either a VESA 75 or VESA 100 interface on the back.</p>
          <p>If you meet these two requirements then there is no reason it shouldn't fit onto a Synergy Stand.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>I want to use Ultrawide screens, do you have a solution for them?</summary>
        <div class="faq-body">
          <p>The Dual Vertical Synergy Stand can handle ultrawide screens, we have customers using them to hold a stack (2 over 2) of 49&Prime; ultrawides with no problems what so ever.</p>
          <p>The Synergy Stand can easily handle the weight due to the steel construction, the problem is the width, however using the Dual Vertical to hold a stack works perfectly well.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Do you offer a clamp solution instead of a free standing base?</summary>
        <div class="faq-body">
          <p>No, we are not a fan of desk clamp mounts, the only base option we have is a stable freestanding base.</p>
          <p>Desk clamps can restrict where a stand can be placed on a desk, and we've seen a few instances of other clamp based stands failing, leaving monitors falling onto the desk or the floor.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How stable is it? Will it tip over? Do I need to drill into my desk?</summary>
        <div class="faq-body">
          <p>We designed the stand solution so the centre of weight is right over the middle of the base plate for maximum stability. In addition, due to all the components being made from steel, once you mount your screens on this will be a weighty setup.</p>
          <p>If you tried to push it over on purpose you'd struggle to do it.</p>
          <p>There is an option to secure it to a desk if you want to drill a hole through, the base plate has a hole to accomodate this however that is more for security / theft prevention rather than stability. For most customers there is no desk drilling required.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Do I need a special desk? What about sit / stand desks?</summary>
        <div class="faq-body">
          <p>You don't need a special desk, our freestanding system will sit on any sturdy desk. We recommend a desk with four legs, supported at each corner, rather than pedestal desks with one central column.</p>
          <p>Standing desks are usually also fine. For mechanically operated ones just check that the lifting capacity is enough. The weight of each Synergy Stand configuration is detailed on each stand product page. Then just add the weight of your screens.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Do I need any screws to attach my screens?</summary>
        <div class="faq-body">
          <p>No, we provide the correct length VESA compatible screws for any screen which follows the VESA interface guidelines.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Do you sell screens to go with the stand?</summary>
        <div class="faq-body">
          <p>Yes, we have a range of 21.5&Prime; to 27&Prime; screens at either Full HD or Quad HD resolutions which are fully compatible with our stands. We sell stands and screens as Monitor Arrays and they come with free long length premium cabling and free UK mainland delivery. <a href="/display-systems/">You can build your perfect monitor array here</a>.</p>
          <p>We also sell computer, stand and screen bundles, all perfectly matched and designed to work together. Bundles include your choice or PC, your preferred Synergy Stand, your monitor pick, free cabling, a couple of free computer upgrades, free UK mainland delivery, and a bundle discount. <a href="/bundles/">Build your perfect bundle here</a>.</p>
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
        <h5>Still deciding on a stand?</h5>
        <h2>Talk to <em>Darren</em>, the founder, not a call centre.</h2>
        <p>Eighteen years of selling multi-screen stands means we know what does and does not work. &ldquo;Will my screens fit?&rdquo; &ldquo;Which configuration for my desk?&rdquo; &ldquo;Can I add screens later?&rdquo; Five to ten minutes on the phone is usually enough to figure out what you need.</p>
        <div class="darren-ctas">
          <a href="tel:03302236655" class="btn btn-primary btn-lg"><i class="fa fa-phone"></i>0330 223 66 55</a>
          <a href="#" class="btn btn-ghost btn-lg js-book-call"><i class="fa fa-calendar"></i>Book a call</a>
        </div>
        <div class="darren-sig">&mdash; Darren Atkinson, founder, Multiple Monitors Ltd</div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     STICKY CTA — "Talk to Darren"
     (Dynamic: Ultra + Extreme prices pulled from DB.)
     =================================================================== -->
<div class="sticky-cta" id="stickyCta">
  <div class="txt">
    <strong>View The Synergy Stand Range</strong>
    <span>Choose from 2 - 8 screen layouts</span>
  </div>
  <a href="#range" class="btn btn-primary btn-sm">Pick a stand <i class="fa fa-arrow-right"></i></a>
</div>

</div><!-- /.mm-site -->

<script>
  // Sticky CTA — visible after scrolling past the hero, hidden near the footer.
  // Wrapped in DOMContentLoaded because this <script> emits before
  // footer_wrapper.asp renders the <footer>, so the element isn't in the
  // DOM yet at parse time.
  (function(){
    function init(){
      var sticky = document.getElementById('stickyCta');
      if (!sticky) return;
      var hero = document.querySelector('.hero');
      var footer = document.querySelector('footer');
      if (!hero || !footer) return;
      function onScroll(){
        var y = window.scrollY || window.pageYOffset;
        var heroBottom = hero.getBoundingClientRect().bottom + y;
        var footerTop = footer.getBoundingClientRect().top + y;
        var viewportBottom = y + window.innerHeight;
        if (y > heroBottom + 200 && viewportBottom < footerTop) {
          sticky.classList.add('visible');
        } else {
          sticky.classList.remove('visible');
        }
      }
      window.addEventListener('scroll', onScroll, { passive:true });
      onScroll();
    }
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', init);
    } else {
      init();
    }
  })();
</script>

<!--#include file="footer_wrapper.asp"-->
