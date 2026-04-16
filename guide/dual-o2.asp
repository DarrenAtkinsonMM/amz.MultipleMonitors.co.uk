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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Dual Monitor Displays - Multiple Monitors</title>
<meta name="Description" content="The Dual Monitors Displays outcome page brought to you by the Multiple Monitors Guide."/>
<meta name="Keywords" content="multiple, monitor, dual, triple, quad, display, screen"/>
<%Response.Buffer=True%>
<%Response.charset="iso-8859-1"%>
<link rel="stylesheet" href="/styles/general.css" type="text/css" />
<link rel="stylesheet" href="/styles/tabcontent.css" type="text/css" media="screen" />
<link rel="stylesheet" href="/styles/fancybox.css" type="text/css" media="screen">
<script type="text/javascript" src="/js/tabcontent.js"></script>

<script type="text/javascript" src="/js/jquery-1.3.2.min.js"></script> 
<script type="text/javascript" src="/js/jquery.fancybox-1.2.1.js"></script> 
<script type="text/javascript" src="/js/jquery.easing.1.3.js"></script>
<!--[if lte IE 6]>
		<link type="text/css" rel="stylesheet" href="/styles/pngfix.css" />
<![endif]-->
</head>     
<body>
<div id="branding"><a href="/" target="_self">
    <img class="logo" title="MultipleMonitors" alt="MultipleMonitors" src="/images/generic/mm-logo.gif" /></a>
    <div class="strapline"><p>Multi-display systems, computers and accessories...</p></div>
    <div class="basket">
		<!--#include file="../shop/pc/SmallShoppingCart.asp"-->
    </div>
</div>

<div id="navigation">
  <div id="mainNav">
	<ul>
		<li class="first"><a title="Home" href="/" >Home</a></li> 
		<li><a title="Monitor Arrays" href="/display-systems/" >Monitor Arrays</a></li> 
		<li><a title="Computers" href="/computers/" >Computers</a></li> 
		<li><a title="Bundles" href="/bundles/" >Bundles</a></li>
		<li><a title="Accessories" href="/accessories/" >Accessories</a></li>
        <li class="last active"><a title="Resources" href="/guide/" >Resources</a></li> 
	</ul>  
</div><!-- END: mainNav -->

</div><!-- END: navigation -->
<div class="clear"></div>

<div id="pageContent">
<div id="tool-page">
<h1>Dual Monitor Displays</h1>
<p>Your current computer is capable of supporting dual monitors which means it is fully compatible with any of our <a href="/dual-monitors/">dual monitor displays</a>.</p>
<p>You can purchase any of our dual monitor display systems confident in the knowledge that installation will simply be a matter of connecting up the leads to your computer.</p>
<p><strong>Dual Monitor Displays:</strong></p>
<div id="tool-featured">
<div class="tf-prod">
		<p class="tf-title"><a href="/products/17-dual-monitors/" >17&quot; Dual Monitors</a></p>
		<a href="/products/17-dual-monitors/" ><img src="/shop/pc/catalog/dual-monitors-stock_1467_thumb.jpg" alt="17&quot; Dual Monitors" class="tf-img" /></a>
		<p class="tf-price">Price &pound;425.00</p>
			<p class="tf-price"><a href="/shop/pc/instPrd.asp?idproduct=21"><img src="/shop/pc/images/pc/mm-buynow.gif" alt="Add to the cart: 17&quot; Dual Monitors" class="tf-img2" /></a><a href="/products/17-dual-monitors/" ><img src="/shop/pc/images/pc/mm-moreinfo.gif" alt="Show product details for 17&quot; Dual Monitors" class="tf-img2" /></a>
		</p>  
</div>
<div class="tf-prod">
		<p class="tf-title"><a href="/products/19-dual-monitors/" >19&quot; Dual Monitors</a></p>
		<a href="/products/19-dual-monitors/" ><img src="/shop/pc/catalog/dual-monitors-stock_1467_thumb.jpg" alt="19&quot; Dual Monitors" class="tf-img" /></a>
		<p class="tf-price">Price &pound;460.00</p>
			<p class="tf-price"><a href="/shop/pc/instPrd.asp?idproduct=22"><img src="/shop/pc/images/pc/mm-buynow.gif" alt="Add to the cart: 19&quot; Dual Monitors" class="tf-img2" /></a><a href="/products/19-dual-monitors/" ><img src="/shop/pc/images/pc/mm-moreinfo.gif" alt="Show product details for 19&quot; Dual Monitors" class="tf-img2" /></a>
		</p>  
</div>
<div class="tf-prod">
		<p class="tf-title"><a href="/products/22-dual-widescreens/" >22&quot; Dual Widescreens</a></p>
 			<a href="/products/22-dual-widescreens/" ><img src="/shop/pc/catalog/dual-monitors-stock_1467_thumb.jpg" alt="22&quot; Dual Widescreens" class="tf-img" /></a>
		<p class="tf-price">Price &pound;495.00</p>
			<p class="tf-price"><a href="/shop/pc/instPrd.asp?idproduct=3"><img src="/shop/pc/images/pc/mm-buynow.gif" alt="Add to the cart: 22&quot; Dual Widescreens" class="tf-img2" /></a><a href="/products/22-dual-widescreens/" ><img src="/shop/pc/images/pc/mm-moreinfo.gif" alt="Show product details for 22&quot; Dual Widescreens" class="tf-img2" /></a>
		</p>  
</div>
</div><!-- tool-featured -->
<p>View all <a href="/dual-monitors/">dual monitor displays</a>.</p>
<div id="tool-ans-top">
<p>Previous Answers</p>
</div>
<div id="tool-ans-mid">
<ul>
<li>You would like to run a dual display (2 screen) system</li>
<li>You would like to use your current computer to run dual monitors</li>
<li>You current computer can already support dual monitors</li>
</ul>
<p>If this is incorrect simply <a href="/guide/dual-q2/">go back a step</a> or <a href="/guide/">restart the guide</a>.</p>
</div>
<div id="tool-ans-bot">
</div>
</div><!-- content-page -->
</div>
<div id="footer">
    <div id="footer-content">   
      <div id="footer-links">
        <img src="/images/generic/footer-logo.gif" alt="MultipleMonitors" style="float:right" /> 
            
  <div class="footer-column">
<p>Site Areas:</p>
                <ul>
					<li><a title="Home" href="/">Home</a></li>
					<li><a title="About Us" href="/pages/about-us/">About Us</a></li>
					<li><a title="Contact Us" href="/shop/pc/contact.asp">Contact Us</a></li>
					<li><a title="Testimonials" href="/pages/testimonials/">Testimonials</a></li>
					<li><a title="Site Map" href="/pages/site-map/">Site Map</a></li>
				</ul>
		</div>
                
				<div class="footer-column">
                    <p>Policies &amp; Legal:</p>
                    <ul>
                        <li><a title="Delivery Information" href="/pages/delivery/">Delivery Information</a></li>
                        <li><a title="Warranty Information" href="/pages/warranty/">Warranty Information</a></li>
                        <li><a title="Returns Policy" href="/pages/returns/">Returns Policy</a></li>
                        <li><a title="Privacy Policy" href="/pages/privacy-policy/">Privacy Policy</a></li>
                        <li><a title="Terms &amp; Conditions" href="/pages/terms/">Terms &amp; Conditions</a></li>
                    </ul>
				</div>
                
				<div class="footer-column">
                    <p>Recently Viewed Products:</p>
					<!--#include file="../shop/pc/smallRecentProducts.asp"-->
                </div>
		<div class="clear"> </div>

			<div id="contact">
                <img src="/images/generic/get-in-touch.gif" alt="Get in Touch" />
                <p class="call">0845 508 5377</p>
                <p class="mail"><a href="mailto:sales@multiplemonitors.co.uk">sales@multiplemonitors.co.uk</a></p>
            </div>

	  </div><!-- END: footer-links -->

			<div id="copyright">
			  <p>&copy; Copyright 2009 MultipleMonitors Ltd.</p>
	  		</div>           
    </div><!-- END: footer-content -->
</div><!-- END: footer -->

<%
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing
%>
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