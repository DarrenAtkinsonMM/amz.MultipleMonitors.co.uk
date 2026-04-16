<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<!--#include file="../shop/includes/settings.asp"-->
<!--#include file="../shop/includes/storeconstants.asp"-->
<!--#include file="../shop/includes/ErrorHandler.asp"-->
<!--#include file="../shop/includes/stringfunctions.asp"-->
<!--#include file="../shop/includes/opendb.asp"-->
<!--#include file="../shop/includes/adovbs.inc"-->
<!--#include file="../shop/includes/languages.asp"--> 
<!--#include file="../shop/includes/currencyformatinc.asp"-->
<!--#include file="../shop/includes/pcProductOptionsCode.asp"--> 
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
<!--#include file="../shop/pc/pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'-------------------------------
' declare local variables
'-------------------------------
dim pcStrHPStyle, pcStrHPDesc, pcIntHPFirst, pcIntHPShowSKU, pcIntHPShowImg, pcIntHPFeaturedCount, pcIntHPFeaturedOrder
dim pcIntHPSpcCount, pcIntHPSpcOrder, pcIntHPNewCount, pcIntHPSNewOrder, pcIntHPBestCount, pcIntHPBestOrder
Dim query, conntemp, rsProducts, rsDisc, pDiscountPerQuantity, pTotalCount
Dim pcv_intBackOrder, pStock, pNoStock, pFormQuantity, pserviceSpec

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Multiple Monitors Trading Computers | Multiple Monitors</title>
<meta name="Description" content="Supplying a wide range of multiple monitor display systems and computers ideal for stock trading."/>
<meta name="Keywords" content="multiple, monitor, dual, triple, quad, display, screen"/>
<meta name="Robots" content="index,follow"/>
<link href="/lightbox/lightbox.css" rel="stylesheet" media="screen, projection" type="text/css" />
<script type="text/javascript" src="/thickbox/jquery-latest.js"></script>
<script type="text/javascript" src="/thickbox/thickbox.js"></script>
<link rel="stylesheet" href="/thickbox/thickbox.css" type="text/css" media="screen" />
<script type="text/javascript" src="/lightbox/lightbox.js"></script>
<link type="text/css" rel="stylesheet" href="/shop/pc/pcStorefront.css" />
<link href="/da.css" rel="stylesheet" media="screen, projection" type="text/css" />
<!--[if lt IE 7]>
<link href="/ie.css" rel="stylesheet"  media="screen,projection" type="text/css" />
<![endif]-->
</head>
<body>
<div id="wrapper">
<div id="header-logo">
<p class="site-id"><a href="/">Multiple Monitors</a></p>
<p class="site-tag">Multi-display systems, computers and accessories...</p>
</div><!-- header-logo -->
<!--#include file="../shop/pc/SmallShoppingCart.asp"-->
<div id="cat-nav">
<ul>
<li><a href="/">HOME</a></li>
<li><a href="/display-systems/">DISPLAY SYSTEMS</a></li>
<li><a href="/computers/">COMPUTERS</a></li>
<li><a href="/bundles/">MONITOR &amp; PC BUNDLES</a></li>
<li><a href="/accessories/">ACCESSORIES</a></li>
<li><a href="/guide/">RESOURCES</a></li>
</ul>
</div><!-- cat-nav -->

<div id="product-page">
<h1>Trading Computer Systems</h1>
<img src="/images/lp/trading-computer-main.jpg" alt="Offer ends 28th February!" class="lp-img-main" />
<div id="l-page-top">
<p class="pd-header">Professional Multi Monitor &amp; Computer Solutions</p>
<p>We are specialists in multi display systems and computers and provide a wide range of multiple monitor products.</p>
<p>Our solutions are a perfect match for anyone looking to build a fast, reliable trading platform for office or home use.</p>
<p>Reasons to select a Multiple Monitors solution:</p>
<ul>
<li>Ultra fast PC's built in-house with the best components available</li>
<li>Top quality TFT screens and ergonomic mounting solutions</li>
<li>Clear &amp; competitive pricing</li>
<li>Excellent customer service levels</li>
<li>Fast &amp; free delivery on all our products</li>
</ul>
</div>
<div id="l-page-main">
<p class="pd-header">Trading Computers</p>
<p>If you're looking for a new trading computer system you need it to be fast and reliable to ensure you have no down time through the trading day. </p>
<p>All our systems are constructed using only top quality components rather than the cheaper less reliable parts that many larger retailers use.</p>
<p>We have different solutions depending on your exact requirements:</p>
<p><strong>Dual Monitor Capable Computers</strong></p>
<ul>
<li><a href="/products/intel-dual-monitor-computer/">Intel Dual Core PC</a> - &pound;595, supports 2 screens, solid performance</li>
<li><a href="/products/ultra-dual-monitor-computer/">Intel Quad Core PC</a> - &pound;985, supports 2 screens, ultra fast Intel quad core</li>
<li><a href="/products/extreme-dual-monitor-computer/">Intel Core i7 PC</a> - &pound;1,275, supports 2 screens, latest Intel i7 processor</li>
</ul>
<p><strong>Triple / Quad Monitor Capable Computers</strong></p>
<ul>
<li><a href="/products/intel-quad-monitor-computer/">Intel Dual Core PC</a> - &pound;745, supports 4 screens, solid performance</li>
<li><a href="/products/ultra-quad-monitor-computer/">Intel Quad Core PC</a> - &pound;1,265, supports 4 screens, ultra fast Intel quad core</li>
<li><a href="/products/extreme-quad-monitor-computer/">Intel Core i7 PC</a> - &pound;1,425, supports 4 screens, latest Intel i7 processor</li>
</ul>
<p><strong>Five + Monitor Capable Computers</strong></p>
<ul>
<li><a href="/products/ultra-six-monitor-computer/">Intel Quad Core PC</a> - &pound;1,415, supports 6 screens, ultra fast Intel quad core</li>
<li><a href="/products/extreme-six-monitor-computer/">Intel Core i7 PC</a> - &pound;1,625, supports 6 screens, latest Intel i7 processor</li>
<li><a href="/products/extreme-eight-monitor-computer/">Intel Core i7 PC</a> - &pound;1,725, supports 8 screens, latest Intel i7 processor</li>
</ul>
<p class="pd-header">Trading Display Systems</p>
<p>The perfect compliments to a new trading computer are our multi monitor display systems.</p>
<p>Featuring top quality, fast response TFT screens and dedicated multiple monitor stands to perfectly mount and align them, these are the premier solution for trading platforms.</p>
<p>See our individual monitor pages for further details:</p>
<ul>
<li><a href="/dual-monitors/">Dual Monitor Displays</a> - From &pound;435 featuring ergonomic cable managed stands</li>
<li><a href="/triple-monitors/">Triple Monitor Displays</a> - From &pound;645 featuring ergonomic tilt & lift stands</li>
<li><a href="/quad-monitors/">Quad Monitor Displays</a> - From &pound;765 featuring horizontal, square and pyramid arrays</li>
<li><a href="/control-stations/">Control Station Displays</a> - From &pound;1,235 featuring 5,6 and 8 screen displays</li>
</ul>
<p class="pd-header">Bundle Discounts</p>
<p>If you're looking to buy a complete monitor and PC system together we have great bundle deals available saving you even more money off our already competitive pricing.</p>
<p>Explore all our <a href="/bundles/">Bundle Deals here</a>.</p>
<p>&nbsp;</p>
</div>
<div class="product-offers">
            <p class="po-contact">Contact us on:</p>
            <p class="po-tel">0845 508 53 77</p>
            <p class="po-email">or <a href="/shop/pc/contact.asp" class="po-a">send us an email</a> for all enquiries about trading systems.</p>

            </div>
            <div class="product-offers">
            <p class="po-title">Complete Peace of Mind!</p>
            <ul class="po-ul">
            <li class="po-li"><a href="/pages/warranty/" class="po-a">Full Warranties Provided</a></li>
            <li class="po-li"><a href="/pages/returns/" class="po-a">7 Day FREE Returns Policy</a></li>
            <li class="po-li"><a href="/pages/delivery/" class="po-a">Aftersales Support Line</a></li>
            </ul>
            </div>
            <div class="product-offers">
            <p class="po-title">Bundle Discount Prices</p>
            <p class="po-body">Save money when buying a new computer and multi screen display solution together!</p>
            </div>
</div><!-- product-page -->


<div id="footer">
<div class="footer-line"></div>
<div id="footer-areas">
<p class="footer-title">Site Areas</p>
<ul>
<li><a href="/">Home</a></li>
<li><a href="/pages/about-us/">About Us</a></li>
<li><a href="/shop/pc/contact.asp">Contact Us</a></li>
<li><a href="/pages/testimonials/">Testimonials</a></li>
<li><a href="/pages/site-map/">Sitemap</a></li>

</ul>
</div><!-- footer-areas -->
<div id="footer-legal">
<p class="footer-title">Policies &amp; Legal</p>
<ul>
<li><a href="/pages/delivery/">Delivery Information</a></li>
<li><a href="/pages/warranty/">Warranty Information</a></li>
<li><a href="/pages/returns/">Returns Policy</a></li>
<li><a href="/pages/privacy-policy/">Privacy Policy</a></li>
<li><a href="/pages/terms/">Terms &amp; Conditions</a></li>

</ul>
</div><!-- footer-legal -->
<!--#include file="../shop/pc/smallRecentProducts.asp"-->
<p class="footer-copy">&copy; 2009 MultipleMonitors.co.uk</p>
</div><!-- footer -->
</div><!-- wrapper -->
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script src="/ga_keyword2.js" type="text/javascript"></script>
<script type="text/javascript">
var pageTracker = _gat._getTracker("UA-5648327-3");
pageTracker._trackPageview();
</script>
</body>
</html>