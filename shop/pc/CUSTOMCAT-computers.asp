<%
' ============================================================
' CUSTOMCAT-computers.asp
' 2026 redesign — Computers landing page.
' Rewritten in place (was 547 lines of ProductCart scaffolding);
' now a mostly-static page with a small direct-DB query for the
' four live prices: Ultra (306), Extreme (307), Trader PC (333),
' Trader Pro (343). See /category-redesign-plan.md at repo root
' for the approach.
' ============================================================
%>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "customcat-computers.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="prv_getSettings.asp"-->
<%
' ------------------------------------------------------------
' Pull the four PC prices (and pcUrl for the two CTAs) in one
' query, keyed by idProduct:
'   306 = Ultra         307 = Extreme
'   333 = Trader PC     343 = Trader Pro
' ------------------------------------------------------------
Dim mmPcSql, mmPcRs
Dim mmPriceUltra, mmPriceExtreme, mmPriceTrader, mmPriceTraderPro
Dim mmUrlUltra, mmUrlExtreme

mmPriceUltra = 0 : mmPriceExtreme = 0 : mmPriceTrader = 0 : mmPriceTraderPro = 0
mmUrlUltra = "" : mmUrlExtreme = ""

mmPcSql = "SELECT idProduct, price, pcUrl FROM products " & _
          "WHERE idProduct IN (306, 307, 333, 343) " & _
          "  AND active = -1 AND removed = 0"

Set mmPcRs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
mmPcRs.Open mmPcSql, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If err.number <> 0 Then
    On Error Goto 0
    call LogErrorToDatabase()
    Set mmPcRs = Nothing
    call closeDB()
    Response.Redirect "techErr.asp?err=" & pcStrCustRefID
End If
On Error Goto 0

Do While Not mmPcRs.EOF
    Select Case CLng(mmPcRs("idProduct"))
        Case 306
            mmPriceUltra = mmPcRs("price")
            mmUrlUltra   = mmPcRs("pcUrl") & ""
        Case 307
            mmPriceExtreme = mmPcRs("price")
            mmUrlExtreme   = mmPcRs("pcUrl") & ""
        Case 333
            mmPriceTrader = mmPcRs("price")
        Case 343
            mmPriceTraderPro = mmPcRs("price")
    End Select
    mmPcRs.MoveNext
Loop
Set mmPcRs = Nothing

' VAT-exclusive display with store currency symbol.
' Returns empty string if the row is missing so the page degrades cleanly.
Function mmPcPrice(ByVal vatIncPrice)
    If IsNumeric(vatIncPrice) And vatIncPrice > 0 Then
        mmPcPrice = scCursign & money(vatIncPrice / 1.2)
    Else
        mmPcPrice = ""
    End If
End Function

' Build detail-page href from pcUrl, falling back to the
' numeric querystring if pcUrl is blank.
Function mmPcHref(ByVal purl, ByVal idProd)
    If purl <> "" Then
        mmPcHref = "/shop/pc/" & purl & ".htm"
    Else
        mmPcHref = "/shop/pc/viewPrd.asp?idproduct=" & idProd
    End If
End Function
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
        <div class="eyebrow">Multi-screen specialists &middot; Since 2008</div>
        <h1>
          Multi-screen PCs for professionals who need <em>more than two displays</em>.
        </h1>
        <p class="lead">
          Built in the UK since 2008. High build quality, stress-testing and lifetime support as standard across our range. Configured for CAD, finance, analytics, security operations, and any workload that&rsquo;s outgrown a standard desktop.
        </p>
        <div class="hero-ctas">
          <a href="#use-cases" class="btn btn-primary btn-lg">See which PC fits your work <i class="fa fa-arrow-right"></i></a>
        </div>
        <div class="hero-mini">
          <div class="item"><i class="fa fa-industry"></i><b>UK-built</b><span>in our workshop</span></div>
          <div class="item"><i class="fa fa-users"></i><b>3,500+</b><span>PCs delivered</span></div>
          <div class="item"><i class="fa fa-shield"></i><b>5-year cover</b><span>&middot; lifetime support</span></div>
        </div>
      </div>

      <div class="hero-visual has-tower reveal" style="transition-delay:.1s">
        <div class="hero-tower">
          <div class="badge">Ultra &middot; UK Built</div>
          <span class="halo" aria-hidden="true"></span>
          <img class="tower-img" src="/images/computers/ultra-open.png" alt="Inside view of an Ultra multi-screen PC, custom built in the UK" />
          <div class="captions">
            <span>32-hr stress test</span>
            <span>5-yr cover</span>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     TRADER REDIRECT STRIP
     =================================================================== -->
<section class="trader-redirect">
  <div class="container">
    <div class="inner reveal">
      <div class="icon" aria-hidden="true"><i class="fa fa-line-chart"></i></div>
      <p>
        <strong>Building a trading setup?</strong> Our Trader PC and Trader Pro are spec&rsquo;d specifically for MT4, NinjaTrader, TradeStation, Bloomberg and other trading platforms &mdash; with published benchmark data.
      </p>
      <a href="/trading-computers/" class="detour-link">
        See the trading computers <i class="fa fa-arrow-right"></i>
      </a>
    </div>
  </div>
</section>

<!-- ===================================================================
     USE-CASE GRID
     =================================================================== -->
<section class="s" id="use-cases">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Which machine do I need?</h5>
        <h2>By the work, <span class="display-em">not by the spec sheet</span>.</h2>
        <p style="max-width:720px; margin-top:12px;">Pick the card that looks closest to the work you actually do. We&rsquo;ll point you at the right starting machine &mdash; and you&rsquo;re two clicks from a human who can tailor from there.</p>
      </div>
    </div>

    <div class="use-cases">

      <!-- Security & Operations -->
      <article class="use-case reveal" data-rec="extreme" style="transition-delay:.18s">
        <div class="uc-illus"><i class="fa fa-shield"></i></div>
        <h3>Security &amp; Operations</h3>
        <p>Monitoring dashboards and multi-feed video walls for SOC, NOC and control-room environments. Built for 24/7 use with proper thermals.</p>
        <div class="chips">
          <span class="mm-chip">SOC</span>
          <span class="mm-chip">NOC</span>
          <span class="mm-chip">IVMS</span>
          <span class="mm-chip">CCTV VMS</span>
        </div>
        <div class="uc-rec">
          <span>Start with the <b>Extreme</b></span>
          <span class="pc-name">EXTREME &rarr;</span>
        </div>
      </article>

      <!-- Finance & Accounting -->
      <article class="use-case reveal" data-rec="ultra" style="transition-delay:.06s">
        <div class="uc-illus"><i class="fa fa-table"></i></div>
        <h3>Finance &amp; Accounting</h3>
        <p>Heavy Excel models, multiple broker platforms, accounting suites and research screens, all running at once without the extreme fan noise.</p>
        <div class="chips">
          <span class="mm-chip">Excel</span>
          <span class="mm-chip">Sage</span>
          <span class="mm-chip">QuickBooks</span>
          <span class="mm-chip">Xero</span>
        </div>
        <div class="uc-rec">
          <span>Start with the <b>Ultra</b></span>
          <span class="pc-name">ULTRA &rarr;</span>
        </div>
      </article>

      <!-- Coding & Development -->
      <article class="use-case reveal" data-rec="ultra" style="transition-delay:.12s">
        <div class="uc-illus"><i class="fa fa-code"></i></div>
        <h3>Coding &amp; Development</h3>
        <p>Programming IDE&rsquo;s tend to be fairly lightweight on computer resources making them a great fit for our Ultra PC.</p>
        <div class="chips">
          <span class="mm-chip">VS Code</span>
          <span class="mm-chip">PyCharm</span>
          <span class="mm-chip">Visual Studio</span>
        </div>
        <div class="uc-rec">
          <span>Start with the <b>Ultra</b></span>
          <span class="pc-name">ULTRA &rarr;</span>
        </div>
      </article>

      <!-- CAD & Engineering -->
      <article class="use-case reveal" data-rec="extreme">
        <div class="uc-illus"><i class="fa fa-cube"></i></div>
        <h3>CAD &amp; Engineering</h3>
        <p>Heavy 3D assemblies, multi-viewport drawings, and workstation-class rendering. Running on stable hardware tested for sustained load.</p>
        <div class="chips">
          <span class="mm-chip">Solidworks</span>
          <span class="mm-chip">AutoCAD</span>
          <span class="mm-chip">Fusion 360</span>
          <span class="mm-chip">Revit</span>
        </div>
        <div class="uc-rec">
          <span>Start with the <b>Extreme</b></span>
          <span class="pc-name">Extreme &rarr;</span>
        </div>
      </article>

      <!-- Analytics & Data -->
      <article class="use-case reveal" data-rec="extreme" style="transition-delay:.24s">
        <div class="uc-illus"><i class="fa fa-bar-chart"></i></div>
        <h3>Analytics &amp; Data</h3>
        <p>Large dashboards, Python/Jupyter notebooks, heavy pivot models and extract-transform jobs. More cores, more RAM, less waiting.</p>
        <div class="chips">
          <span class="mm-chip">Power BI</span>
          <span class="mm-chip">Tableau</span>
          <span class="mm-chip">SQL</span>
          <span class="mm-chip">Python</span>
        </div>
        <div class="uc-rec">
          <span>Start with the <b>Extreme</b></span>
          <span class="pc-name">EXTREME &rarr;</span>
        </div>
      </article>

      <!-- General Multi-Screen Productivity -->
      <article class="use-case reveal" data-rec="ultra" style="transition-delay:.30s">
        <div class="uc-illus"><i class="fa fa-th-large"></i></div>
        <h3>General Multi-Screen Productivity</h3>
        <p>Writers, project managers, entrepreneurs and anyone with too many windows open. When your off-the-shelf desktop runs out of breath.</p>
        <div class="chips">
          <span class="mm-chip">Writing</span>
          <span class="mm-chip">Web Interfaces</span>
          <span class="mm-chip">Research</span>
        </div>
        <div class="uc-rec">
          <span>Start with the <b>Ultra</b></span>
          <span class="pc-name">ULTRA &rarr;</span>
        </div>
      </article>

    </div>

    <p style="text-align:center; color:var(--muted); font-size:17px; margin:30px 0 0; font-family:'EB Garamond', serif; font-style:italic;" class="reveal">
      Not sure which fits? <a href="#darren" style="color:var(--brand); font-weight:500;">Talk to our team</a> most decisions are settled in a ten-minute call.
    </p>
  </div>
</section>

<!-- ===================================================================
     THE TWO PCs — Ultra vs Extreme
     (Dynamic: prices + Configure URLs pulled from DB.)
     =================================================================== -->
<section class="s depth" id="pcs">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Ultra &amp; Extreme</h5>
        <h2>Two machines, <span class="display-em">the range in summary</span>.</h2>
        <p style="max-width:760px; margin-top:12px;">The Ultra can handle most standard workloads fine and can be spec&rsquo;d to a high level. The Extreme exists for the jobs where more processing power is required, complex dashboards, rendering, number crunching and large multi-tasking workloads.</p>
      </div>
    </div>

    <div class="pc-picks">
      <!-- Ultra -->
      <article class="pc-pick reveal" data-tier="ultra">
        <div class="pc-pick__head">
          <div>
            <div class="kicker">The flexible multi-screen computer</div>
            <h3>Ultra</h3>
            <p class="lead-line">Our most popular non-trader machine. Highly configurable across a range of tasks.</p>
          </div>
          <div class="pc-pick__img">
            <img src="/images/bundles/bun-ultra-pc.png" alt="Ultra multi-screen PC" />
          </div>
        </div>

        <div class="price-row">
          <span class="from">From</span>
          <span class="amount"><%= mmPcPrice(mmPriceUltra) %></span>
          <small>+ VAT</small>
        </div>
        <ul class="pc-spec-list">
          <li><span class="k">CPU</span><span class="v">Intel Core i3 &rarr; i9 (14th gen)</span></li>
          <li><span class="k">RAM</span><span class="v">16 &ndash; 64 GB DDR4</span></li>
          <li><span class="k">Screens</span><span class="v">Up to 8 displays</span></li>
        </ul>
        <p class="who">For anyone looking for a capable multi-screen productivity PC.</p>
        <div class="cta-row">
          <a href="<%= mmPcHref(mmUrlUltra, 306) %>" class="btn btn-primary">Configure your Ultra <i class="fa fa-arrow-right"></i></a>
          <a href="tel:03302236655" class="btn btn-ghost"><i class="fa fa-phone"></i>Talk it through</a>
        </div>
      </article>

      <!-- Extreme -->
      <article class="pc-pick reveal" data-tier="extreme" style="transition-delay:.08s">
        <div class="pc-pick__head">
          <div>
            <div class="kicker">The uncompromising workstation class PC</div>
            <h3>Extreme</h3>
            <p class="lead-line">Assembled using the latest and fastest processors, designed for the workloads that break on everything else.</p>
          </div>
          <div class="pc-pick__img">
            <img src="/images/bundles/bun-extreme-pc.png" alt="Extreme high-end multi-screen workstation" />
          </div>
        </div>

        <div class="price-row">
          <span class="from">From</span>
          <span class="amount"><%= mmPcPrice(mmPriceExtreme) %></span>
          <small>+ VAT</small>
        </div>
        <ul class="pc-spec-list">
          <li><span class="k">CPU</span><span class="v">Intel Core Ultra 5 &rarr; 9 // AMD 9000 Series</span></li>
          <li><span class="k">RAM</span><span class="v">16 &ndash; 128 GB DDR5</span></li>
          <li><span class="k">Screens</span><span class="v">Up to 12 displays</span></li>
        </ul>
        <p class="who">For power users who need serious processing power on their desk.</p>
        <div class="cta-row">
          <a href="<%= mmPcHref(mmUrlExtreme, 307) %>" class="btn btn-primary">Configure your Extreme <i class="fa fa-arrow-right"></i></a>
          <a href="tel:03302236655" class="btn btn-ghost"><i class="fa fa-phone"></i>Talk it through</a>
        </div>
      </article>
    </div>

    <p style="text-align:center; margin:26px 0 0; color:var(--muted); font-size:14px;" class="reveal">
      Side-by-side differences above. Basically if the Ultra won&rsquo;t do it, the Extreme will. And if that&rsquo;s not clear, <a href="tel:03302236655" style="color:var(--brand);">phone us</a>.
    </p>
  </div>
</section>

<!-- ===================================================================
     CROSS-LINK TO TRADING
     (Dynamic: Trader PC + Trader Pro prices pulled from DB.)
     =================================================================== -->
<section class="s">
  <div class="container">
    <div class="computers-xlink reveal">
      <div>
        <h5>Trading?</h5>
        <h3>Our Trader PC and Trader Pro are configured <em>specifically</em> for trading platforms.</h3>
        <p>Published benchmark data. MT4/5 Trading View, NinjaTrader, TradeStation, Bloomberg and more tested.</p>
      </div>
      <div class="cta-col">
        <a href="/trading-computers/" class="btn btn-accent btn-lg">Go to trading computers <i class="fa fa-arrow-right"></i></a>
        <small>Trader PC from <%= mmPcPrice(mmPriceTrader) %> &middot; Trader Pro from <%= mmPcPrice(mmPriceTraderPro) %></small>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     WHAT YOU GET (shared across the range)
     =================================================================== -->
<section class="s">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Shared across the range</h5>
        <h2>Same workshop. Same build quality. <span class="display-em">No matter which one you pick.</span></h2>
        <p style="max-width:720px; margin-top:12px;">Four computer options, built to one standard. The Ultra and Extreme share the same workshop, the same stress-tests, and the same 5-year cover as the Trader PC and Trader Pro. The spec changes the build quality and support doesn&rsquo;t.</p>
      </div>
    </div>

    <div class="pillars">
      <div class="pillar reveal">
        <div class="icon"><i class="fa fa-industry"></i></div>
        <h4>UK-built, in our workshop</h4>
        <p>Every PC assembled, stress-tested and packaged up in our UK workshop. Our in-house team builds and tests every single PC before dispatch.</p>
        <div class="tag">UK ASSEMBLY</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.06s">
        <div class="icon"><i class="fa fa-flask"></i></div>
        <h4>32-hour stress tests</h4>
        <p>Before it ships, every PC runs a 32-hour intensive stress test across CPU, RAM, hard drive and graphics card. If a component is going to fail, it fails here, not at your desk.</p>
        <div class="tag">32-HR SOAK</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.12s">
        <div class="icon"><i class="fa fa-shield"></i></div>
        <h4>Extended hardware cover</h4>
        <p>Extendable on-site support packages for every computer we sell, combine with 5 year cover and lifetime telephone, email and remote desktop support.</p>
        <div class="tag">5YR + LIFETIME</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.18s">
        <div class="icon"><i class="fa fa-desktop"></i></div>
        <h4>Multi-screen graphics</h4>
        <p>Tried and tested multi-screen capable graphics setups. Run from 1 - 12 screens at a variety of resolutions without any hassle or compatibility issues.</p>
        <div class="tag">PRO GPU OPTIONS</div>
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
          <div class="label">Since 2008</div>
          <div class="val">3,500+ PCs Delivered</div>
        </div>
      </div>
      <div class="trust-item accent">
        <div class="icon"><i class="fa fa-life-ring"></i></div>
        <div>
          <div class="label">Lifetime UK support</div>
          <div class="val">The team that built your PC</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     REVIEWS
     =================================================================== -->
<section class="s reviews" id="reviews">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>What customers say</h5>
        <h2>From the desks, labs &amp; studios <span class="display-em">we build for</span>.</h2>
        <p>All reviews are voluntary &mdash; we don&rsquo;t ask for them.</p>
      </div>
      <div class="tp-summary">
        <span class="tp-stars"><span></span><span></span><span></span><span></span><span></span></span>
        <span><b>4.9</b> <small>&middot; based on 90+ reviews</small></span>
        <a href="https://uk.trustpilot.com/review/multiplemonitors.co.uk" class="link" style="margin-left:10px;">See all on Trustpilot <i class="fa fa-external-link"></i></a>
      </div>
    </div>

    <div class="reviews-grid">
      <div class="review reveal">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <h4>First-class service</h4>
        <p>First-class service from initial enquiry to delivery. Great advice received during the process from Darren, who is very knowledgable and responsive. Really happy with my customised PC package, the price, and peace of mind that after-sales support is there if needed.</p>
        <div class="meta">
          <div class="ava">GF</div>
          <div class="who">Gavin Foster</div>
          <div class="when">08&thinsp;/&thinsp;2025</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.08s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <h4>These Guys are fantastic</h4>
        <p>These Guys are fantastic! Had a 6 screen system including monitor stand. I&rsquo;m definitely not a techy and this worked straight out of the box. I asked for some advice after approx 10 months about screen configuration, they logged into my PC to assist in getting this altered to suit my needs. Nothing too much trouble. 100% reccomend from me.</p>
        <div class="meta">
          <div class="ava">JC</div>
          <div class="who">James Andrew Clegg</div>
          <div class="when">08&thinsp;/&thinsp;2025</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.16s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <h4>Very happy with the team</h4>
        <p>I was particularly impressed in the way they helped me to choose the right machine specification while taking their time to answer all my questions to ensure I was happy. I wouldn&rsquo;t hesitate to recommend Multiple Monitors to anyone looking for a high-power and quality set-up from a company that delivers fabulous customer support.</p>
        <div class="meta">
          <div class="ava">GD</div>
          <div class="who">Gerry Drew</div>
          <div class="when">04&thinsp;/&thinsp;2025</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     HONEST OVERLAP PANEL
     =================================================================== -->
<section class="s">
  <div class="container">
    <div class="honest-panel reveal">
      <div class="hp-grid">
        <div>
          <h5 class="eyebrow" style="margin:0;">Straight answer</h5>
          <h2>Ultra vs Trader PC &mdash; <em>the honest version</em>.</h2>
        </div>
        <div>
          <p>You might see a Trader PC and wonder how it relates to the Ultra? <strong>The honest answer is they share the same base platform</strong>. The Trader PC is configured, spec&rsquo;d and tested specifically for trading software workloads. The Ultra is the same base machine starting at a slight lower spec but with a broader range of options for general-purpose multi-screen work.</p>
          <p>Same build quality, same workshop, same 5-year cover. Pick the one with the configuration that fits your use case, the underlying hardware quality is identical. <strong>It&rsquo;s the same story for the Extreme and Trader Pro</strong>.</p>
          <div class="attr"><span class="dot"></span>Written by Darren &middot; Multiple Monitors</div>
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
      <h2>The questions we answer <span class="display-em">on the phone</span>, every week.</h2>
      <p style="margin-top:12px;">If yours isn&rsquo;t here, <a href="tel:03302236655" style="color:var(--brand); font-weight:500;">call us</a> on 0330 223 66 55 &mdash; Darren takes most of them himself.</p>
    </div>

    <div class="faq-list reveal">

      <details class="faq-item" open>
        <summary>What&rsquo;s the difference between your Ultra and your Trader PC?</summary>
        <div class="faq-body">
          <p>They share the same base platform. The <strong>Trader PC</strong> is configured and tested specifically for trading software, we run benchmark tests against platforms like MT4, NinjaTrader, TradeStation, Bloomberg. The <strong>Ultra</strong> is the same workshop, the same build quality, the same 5-year cover, configured for broader multi-screen work: CAD, finance, analytics, productivity.</p>
          <p>If you&rsquo;re trading, the <strong>Trader PC</strong> is the right place to start. If you&rsquo;re not, the Ultra is a better option. It&rsquo;s exactly the same story for our Extreme and Trader Pro.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Will a multi-screen PC work with my existing monitors?</summary>
        <div class="faq-body">
          <p>Almost certainly. Every Ultra and Extreme ships with graphics that drive 4 digital screens as standard, and we have options to support 6, 8, 10 or 12 screens. Digital monitor ports can easily be switched from one type to another, so HDMI, DisplayPort and even the older DVI ports can all be supported, tell us what your monitors use and we&rsquo;ll advise on the right cables.</p>
          <p>Key things to speak to us about are if you plan on driving a lot of higher resolution screens such as 4K and some of the larger ultrawide screens. They can be supported but it&rsquo;s best to get the right graphics setup decided upon upfront.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>I use Solidworks / heavy Excel / Power BI / some other software - which PC will handle it?</summary>
        <div class="faq-body">
          <p>The honest answer depends on your specific usage. A 50&thinsp;MB Excel workbook with a thousand formulas is different from a 1&thinsp;GB Power Query model with live SQL connections. One person's workload can be wildly different to another's, even when running similar software.</p>
          <p>The best thing to do is speak to us first, we can discuss your workload, look at your current setup, and then spec something appropriate for your needs.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>How many screens do I actually need for my workload?</summary>
        <div class="faq-body">
          <p>For most finance and productivity workflows, three or four screens is the real productivity sweet-spot. Beyond that you can start having to turn your head to take everything in. For dashboards and monitoring work, six to eight is where most customers settle. Anything over eight is usually for a very specific workflow and worth a call before you buy.</p>
          <p>A good idea is to count the windows you currently keep open while you work. That&rsquo;s roughly how many screens you want.</p>
          <p>Also worth discussing is whether using higher resolution screens like QHD might allow you to fit your workload on to a lower number of screens.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Do I need a workstation-class / high end graphics card?</summary>
        <div class="faq-body">
          <p>It depends on what you are doing but often the answer is no.</p>
          <p>For example, anyone running general business, Internet, email, office and dashboard type software on standard resolution screens (Full HD) then standard graphics cards are more than enough.</p>
          <p>If you&rsquo;re running more graphically intense packages or anything that actually uses the GPU to render or perform calculations (like AI models) then opting for a high-powered graphics card is usually worth the extra cost.</p>
          <p>Running larger numbers of higher resolution screens can also benefit from a bump in graphics power. If you&rsquo;re not sure just give us a call.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>Can you spec a machine for a specific industry or workflow?</summary>
        <div class="faq-body">
          <p>That&rsquo;s most of what the pre-sale conversation is. Tell us the industry and the application, and we&rsquo;ll put together the right CPU/RAM/GPU combination from what we&rsquo;ve already built for similar customers. We&rsquo;ve shipped Ultra and Extreme builds for architecture studios, accountancy firms, security integrators, analytics teams, universities and control rooms. If it involves more than two screens and a real workload, chances are we&rsquo;ve done something close.</p>
          <p>For procurement teams buying multiple matched machines: mention it up front. There are bulk-build efficiencies that only kick in past three or four units.</p>
        </div>
      </details>

      <details class="faq-item">
        <summary>I need a full new setup including computer and screens, can you help?</summary>
        <div class="faq-body">
          <p>We certainly can. Any of our computers are available in a bundle deal. Bundles include your choice of computer, multi-screen Synergy Stand, screens, cabling, some free computer upgrades, free UK mainland delivery, and a bundle discount.</p>
          <p>Save up to &pound;300 - <a href="/bundles/">View our bundle deals</a></p>
        </div>
      </details>

      <details class="faq-item">
        <summary>I need a few computers, are there any discounts available?</summary>
        <div class="faq-body">
          <p>For procurement teams buying multiple machines then let us know. There are some bulk-build efficiencies that can kick in past three or four units.</p>
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
        <h2>Talk to <em>Darren</em>, the founder, not a call centre.</h2>
        <p>Eighteen years of speccing multi-screen builds for traders, finance teams, control rooms and professional users of all walks. Tell him about your work and he&rsquo;ll steer you to the right machine, often a cheaper one than you were about to buy. No hard sell.</p>
        <div class="darren-ctas">
          <a href="tel:03302236655" class="btn btn-primary btn-lg"><i class="fa fa-phone"></i>0330 223 66 55</a>
          <a href="#" class="btn btn-ghost btn-lg js-book-call"><i class="fa fa-calendar"></i>Book a 15-min call</a>
        </div>
        <div class="darren-sig">&mdash; Darren Atkinson, founder, Multiple Monitors Ltd</div>
        <div class="darren-note">&ldquo;If we think you don&rsquo;t need us, we&rsquo;ll tell you. That&rsquo;s why the Trustpilot score looks like it does.&rdquo;</div>
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
    <strong>Not sure which PC?</strong>
    <span>Ultra from <%= mmPcPrice(mmPriceUltra) %> &middot; Extreme from <%= mmPcPrice(mmPriceExtreme) %> + VAT</span>
  </div>
  <a href="tel:03302236655" class="btn btn-primary btn-sm"><i class="fa fa-phone"></i>Talk to our Team</a>
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
