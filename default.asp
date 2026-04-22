<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="shop/includes/common.asp"-->
<!--#include file="shop/includes/common_checkout.asp"-->
<!--#include file="shop/includes/CashbackConstants.asp"-->
<!--#include file="shop/pc/HomeCode.asp"-->
<!--#include file="shop/pc/prv_incFunctions.asp"-->
<%
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "home.asp"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="shop/pc/pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="shop/pc/prv_getSettings.asp"-->


<!--#include file="shop/pc/header_wrapper.asp"-->

<div class="mm-site">

<!-- ===================================================================
     HERO
     =================================================================== -->
<section class="hero">
  <div class="container">
    <div class="hero-grid">
      <div class="reveal">
        <div class="eyebrow">Multi-screen experts &middot; Since 2008</div>
        <h1>
          Multi-screen specialists, &amp; the leading UK <em>trading computer</em> supplier.
        </h1>
        <p class="lead">
          Trusted by hedge funds, prop trading firms, financial institutions and thousands of professionals who need more screens. UK-built, benchmark-tested, supplied with lifetime support.
        </p>
        <div class="hero-ctas">
          <a href="#trust" class="btn btn-primary btn-lg">Find what you need <i class="fa fa-arrow-right"></i></a>
        </div>
        <div class="hero-mini">
          <div class="item"><i class="fa fa-shield"></i><span>5-year hardware cover</span></div>
          <div class="item"><i class="fa fa-headphones"></i><span>Lifetime UK support</span></div>
        </div>
      </div>

      <div class="hero-visual reveal" style="transition-delay:.1s">
        <img src="/images/pages/trading-image.png" alt="Multi-screen trading computer setup" />
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
          <div class="val">4,500+ multi-screen systems</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     TWO-DOOR ROUTING
     =================================================================== -->
<section class="s depth" id="doors">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>What can we help you with?</h5>
        <h2>Specialist solutions for every requirement.</h2>
      </div>
      <a href="#darren" class="talk-link">Not sure? Talk to Darren <i class="fa fa-arrow-right"></i></a>
    </div>

    <div class="doors">

      <a href="/trading-computers/" class="door door-trader reveal">
        <span class="icon-bg"><i class="fa fa-line-chart"></i></span>
        <div class="num">01 / FOR TRADERS</div>
        <h3>I&rsquo;m a trader</h3>
        <p>Specialist trading PCs designed using our own independent benchmark data. Customers use MT4, NinjaTrader, TradingView, Bloomberg and more. Trusted by hedge funds and individual traders since 2008.</p>
        <span class="arrow-link">See the Trader range <i class="fa fa-arrow-right"></i></span>
      </a>

      <a href="/computers/" class="door reveal" style="transition-delay:.08s">
        <span class="icon-bg"><i class="fa fa-desktop"></i></span>
        <div class="num">02 / FOR PROFESSIONALS</div>
        <h3>I need a multi-screen PC</h3>
        <p>Multi-monitor capable computers for any professional workload &mdash; CAD and engineering, financial control, security operations, data analytics, and anyone whose work outgrew a standard desktop / laptop.</p>
        <span class="arrow-link">See multi-screen PCs <i class="fa fa-arrow-right"></i></span>
      </a>

      <div class="door door-split reveal" style="transition-delay:.16s">
        <span class="icon-bg"><i class="fa fa-th-large"></i></span>
        <div class="num">03 / STANDS &amp; SCREENS</div>
        <h3>I just need stands or screens</h3>
        <p>View and buy a range of rock solid, UK designed and manufactured, modular Synergy Stands. Pair the stands with some screens for a professional multi-monitor display system.</p>
        <div class="door-links">
          <a href="/stands/" class="arrow-link">Stands <i class="fa fa-arrow-right"></i></a>
          <a href="/display-systems/" class="arrow-link">Monitor arrays <i class="fa fa-arrow-right"></i></a>
        </div>
      </div>

    </div>
  </div>
</section>

<!-- ===================================================================
     WHY-US PILLARS
     =================================================================== -->
<section class="s-tight" style="border-top:1px solid var(--line); border-bottom:1px solid var(--line);">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>What makes us different</h5>
        <h2>Built for <span class="display-em">people who use</span> their screens all day.</h2>
      </div>
    </div>

    <div class="pillars">
      <div class="pillar reveal">
        <div class="icon"><i class="fa fa-industry"></i></div>
        <h4>UK-built, from our workshop</h4>
        <p>Every PC assembled, stress-tested and packed in our UK workshop. Not drop-shipped, not white-labelled &mdash; our team puts hands on every build.</p>
        <div class="tag">UK ASSEMBLY</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.06s">
        <div class="icon"><i class="fa fa-line-chart"></i></div>
        <h4>In-house benchmark testing</h4>
        <p>TraderSpec.com is our own benchmark data &mdash; the only UK trading-computer specialist with published test results across real trading platforms.</p>
        <div class="tag">TRADERSPEC &middot; 2011&ndash;</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.12s">
        <div class="icon"><i class="fa fa-shield"></i></div>
        <h4>5-year hardware cover</h4>
        <p>Two years longer than US specialists typically offer, and matched by nobody selling generic multi-screen builds. Included on every PC.</p>
        <div class="tag">5YR STANDARD</div>
      </div>
      <div class="pillar reveal" style="transition-delay:.18s">
        <div class="icon"><i class="fa fa-life-ring"></i></div>
        <h4>Lifetime UK phone support</h4>
        <p>Talk to the people who built your machine &mdash; not an overseas first-line script. The same team, years later, knowing your setup.</p>
        <div class="tag">FOR THE LIFE OF THE PC</div>
      </div>
    </div>

  </div>
</section>

<!-- ===================================================================
     CUSTOMER LOGO STRIP
     =================================================================== -->
<section class="logos">
  <div class="container">
    <div class="title">Used by traders, analysts and engineers at</div>
    <div class="logos-row">
      <div class="logo">Capital &amp; Crest</div>
      <div class="logo sans">Northlake Trading</div>
      <div class="logo mono">MERIDIAN&nbsp;FX</div>
      <div class="logo">Blakemore Analytics</div>
      <div class="logo sans">Arden Quant</div>
      <div class="logo mono">VERTEX&nbsp;DESK</div>
    </div>
    <div class="title" style="margin-top:14px; color:var(--muted); font-style:italic; letter-spacing:.02em; font-family:'Fraunces',serif; text-transform:none; font-size:13px;">Representative placeholder &mdash; real customer logos subject to permission.</div>
  </div>
</section>

<!-- ===================================================================
     REVIEWS (Trustpilot style)
     =================================================================== -->
<section class="s reviews">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>What customers say</h5>
        <h2>90+ unsolicited reviews. <span class="display-em">4.9 stars.</span></h2>
        <p>All reviews are voluntary, we don't ask for them.</p>
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
        <span class="platform">NinjaTrader &middot; 6-screen</span>
        <h4>Finally, a PC that keeps up with my charts</h4>
        <p>Moved from a self-built gaming rig that kept stuttering on tick data. This machine has been bulletproof for nine months, and Darren took the time to actually understand my workflow before speccing it. Night and day.</p>
        <div class="meta">
          <div class="ava">JM</div>
          <div class="who">James M.</div>
          <div class="when">03&thinsp;/&thinsp;2026</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.08s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <span class="platform">Solidworks &middot; CAD studio</span>
        <h4>Replaced three workstations across the studio</h4>
        <p>We needed something that would handle Solidworks assemblies across six screens without breaking a sweat. Darren&rsquo;s team built three matched machines &mdash; delivered, configured, no fuss. Support has picked up within a ring every time we&rsquo;ve called.</p>
        <div class="meta">
          <div class="ava">SP</div>
          <div class="who">Sarah P., Director</div>
          <div class="when">02&thinsp;/&thinsp;2026</div>
        </div>
      </div>
      <div class="review reveal" style="transition-delay:.16s">
        <div class="stars">&#9733;&#9733;&#9733;&#9733;&#9733;</div>
        <span class="platform">TradeStation &middot; Prop firm</span>
        <h4>Talked me out of a pointless i9 upgrade</h4>
        <p>Half-expected a hard sell for the most expensive CPU on the page. Instead they pointed me at their own benchmark data showing the i7 would actually be faster for my platform, saved me &pound;300, and it&rsquo;s performed exactly as they said. Rare kind of honest.</p>
        <div class="meta">
          <div class="ava">DR</div>
          <div class="who">Daniel R.</div>
          <div class="when">01&thinsp;/&thinsp;2026</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     BUNDLE SAVINGS BAND
     =================================================================== -->
<section class="bundle">
  <div class="container">
    <div class="bundle-grid">
      <div class="reveal">
        <h5>Complete bundles</h5>
        <h2>Buy the <em>whole setup</em>, save &pound;200+ vs piecing it together.</h2>
        <p>Pick a stand, screens, and a PC - We build it, test it, and deliver it in one simple package, at a discount.</p>
        <div class="bundle-pills">
          <span class="bundle-pill"><i class="fa fa-check"></i>Free long-length cables</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free wifi card</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free speakers</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Free UK delivery</span>
          <span class="bundle-pill"><i class="fa fa-check"></i>Auto bundle discount</span>
        </div>
        <div style="display:flex; gap:12px; flex-wrap:wrap;">
          <a href="/bundles/" class="btn btn-accent btn-lg">Build your bundle <i class="fa fa-arrow-right"></i></a>
        </div>
      </div>
      <div class="reveal" style="transition-delay:.1s">
        <div class="save-card">
          <span class="save-tag">Example &middot; 6-screen</span>
          <div class="kicker">Typical saving on a 6-screen bundle</div>
          <div class="big"><small>&pound;</small>270</div>
          <div class="sub">vs ordering the same stand, screens, PC and cables separately.</div>
          <div class="breakdown">
            <div class="r"><span>6&thinsp;&times;&thinsp;3m long high quality video cables</span><b>&pound;90</b></div>
            <div class="r"><span>Free PC Upgrades (Wifi, BT & Speakers)</span><b>&pound;60</b></div>
            <div class="r"><span>Free UK mainland delivery</span><b>&pound;20</b></div>
            <div class="r"><span>Bundle discount</span><b>&pound;100</b></div>
            <div class="r total"><span>Total savings</span><b>&minus; &pound;270</b></div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===================================================================
     DEPTH TEASERS (TraderSpec + Guide + Blog)
     =================================================================== -->
<section class="s depth">
  <div class="container">
    <div class="section-head reveal">
      <div>
        <h5>Go deeper</h5>
        <h2>Published benchmarks. Honest buying advice. <span class="display-em">No pressure.</span></h2>
      </div>
    </div>

    <div class="depth-grid">

      <!-- TraderSpec teaser -->
      <div class="card ts-card reveal">
        <h5>TraderSpec&thinsp;.com &middot; our own data</h5>
        <h3>How the industry&rsquo;s CPUs actually perform on live trading platforms.</h3>
        <p>We benchmark every CPU we fit against the real software traders run &mdash; not manufacturer benchmarks. Here&rsquo;s a snapshot. Competitors don&rsquo;t publish this; we do.</p>
        <div class="ts-bars">
          <div class="ts-bar">
            <div class="name">i7-14700K</div>
            <div class="track"><div class="fill" style="width:94%"></div></div>
            <div class="val">94</div>
          </div>
          <div class="ts-bar">
            <div class="name">Ryzen 9 7900X</div>
            <div class="track"><div class="fill alt" style="width:81%"></div></div>
            <div class="val">81</div>
          </div>
          <div class="ts-bar">
            <div class="name">i5-14600K</div>
            <div class="track"><div class="fill alt" style="width:78%"></div></div>
            <div class="val">78</div>
          </div>
          <div class="ts-bar">
            <div class="name">Ryzen 5 7600</div>
            <div class="track"><div class="fill alt" style="width:62%"></div></div>
            <div class="val">62</div>
          </div>
        </div>
        <div class="ts-footnote">Single-thread index &middot; MT4 tick replay &middot; lower is slower. <a href="/trading-computers/" style="font-weight:500;">Full chart &rsquo; TraderSpec data&nbsp;<i class="fa fa-arrow-right" style="font-size:11px"></i></a></div>
      </div>

      <!-- Buyers guide -->
      <div class="card guide-card reveal" style="transition-delay:.08s">
        <h5><i class="fa fa-book" style="margin-right:6px;"></i>Free buyer&rsquo;s guide</h5>
        <h3>The 23-page guide we wish every customer read first.</h3>
        <p>Plain-English answers on CPU choice, screen counts, GPU requirements, and how to avoid over-spending on specs that won&rsquo;t help you. Sent as a PDF.</p>
        <form class="guide-form" onsubmit="event.preventDefault(); this.querySelector('input').value=''; this.querySelector('input').placeholder='Thanks &mdash; check your inbox'">
          <input type="email" placeholder="you@example.com" aria-label="Email" required>
          <button type="submit" class="btn btn-accent">Send me the PDF</button>
        </form>
        <div style="font-size:12px; color:#7A8699; margin-top:8px;">No spam &middot; unsubscribe any time &middot; UK GDPR</div>
      </div>

      <!-- Blog -->
      <div class="card blog-card reveal" style="transition-delay:.16s">
        <div class="thumb" style="background-image:linear-gradient(135deg, rgba(14,27,44,.3), rgba(14,27,44,.1)), url('/images/ts-blog.jpg');">
          <span class="meta-tag">Benchmarks</span>
        </div>
        <div class="body">
          <div class="date">March 2026 &middot; 6 min read</div>
          <h3>TraderSpec.com: what we&rsquo;ve learned from 10,000 benchmark runs.</h3>
          <a class="read" href="/blog/">Read the article <i class="fa fa-arrow-right"></i></a>
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
        <h5>Still deciding?</h5>
        <h2>Talk to <em>Darren</em> &mdash; the founder, not a call centre.</h2>
        <p>Seventeen years of speccing these builds means most of our customers&rsquo; questions have pretty direct answers. No scripts, no hard sell &mdash; a 15-minute call is usually enough to figure out whether we&rsquo;re right for you.</p>
        <div class="darren-ctas">
          <a href="tel:03302236655" class="btn btn-primary btn-lg"><i class="fa fa-phone"></i>0330 223 66 55</a>
          <a href="#" class="btn btn-ghost btn-lg"><i class="fa fa-calendar"></i>Book a 15-min call</a>
        </div>
        <div class="darren-sig">&mdash; Darren Atkinson, founder, Multiple Monitors Ltd</div>
      </div>
    </div>
  </div>
</section>

</div><!-- /.mm-site -->

<!--#include file="shop/pc/orderCompleteTracking.asp"-->
<!--#include file="shop/pc/inc-Cashback.asp"-->
<!--#include file="shop/pc/footer_wrapper.asp"-->