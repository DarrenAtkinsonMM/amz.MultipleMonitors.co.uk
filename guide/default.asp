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
<title>Multiple Monitors Guide | Multiple Monitors</title>
<meta name="Description" content="Supplying a wide range of multiple monitor display systems and accessories for business and home use across the UK."/>
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
<h1>Multiple Monitors Guide (Beta)</h1>
<p>Welcome to the Multiple Monitor's guide to successfully setting up a multiple display system.</p>
<p>One of the biggest headaches when moving to a multi screen system is trying to work out which computer would be the best to buy, or which parts on your existing computer need upgrading, etc...</p>
<p>It's confusing even for the most experienced of us, so we have put together this tool to try and help guide you through the process.</p>
<h2>How many displays do you want to run?</h2>
<div class="tool-but">
  <a href="/guide/dual-q1/"><img src="/images/tools/dual-but.jpg" alt="Dual Screens" width="210" height="184" class="tool-but-img" /></a></div>
<div class="tool-but">
  <a href="triple-q1/"><img src="/images/tools/triple-but.jpg" alt="Triple Screens" width="210" height="184" class="tool-but-img" /></a></div>
<div class="tool-but">
  <a href="quad-q1/"><img src="/images/tools/quad-but.jpg" alt="Quad Screens" width="210" height="184" class="tool-but-img" /></a></div>
<p class="tool-clear"><strong>More about the tool...</strong></p>
<p>Simply answer the few simple questions as they appear to discover if your current computer system is capable of running the desired amount of extra screens, and if not, exactly what extra hardware will you need to get the perfect multi monitor setup!</p>
<p><strong>Please note</strong>, at this time this tool is only for PC users, we will be adding Mac options and guides very soon though...</p>
<p>&nbsp;</p>
  </div>
<!-- content-page -->
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