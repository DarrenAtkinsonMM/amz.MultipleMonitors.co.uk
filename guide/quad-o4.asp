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
<title>Quad Screen Upgrade Options - Multiple Monitors</title>
<meta name="Description" content="The options available to you when attempting to upgrade your current computer to support quad monitors."/>
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
<h1>Quad Monitor Upgrade Options</h1>
<p>Using the answers you have supplied we have determined the following options for you to consider in your attempt to run quad monitors:</p>
<p><strong>Install Dedicated Quad Graphics Card</strong></p>
<p>You can purchase a 'quad head' graphics card which can be installed into your computer to give you the ability to output to up to 4 different screens from the one card.</p>
<p>The benefits of this solution are that upgrading and installation is no more difficult than for any standard graphics card, however 'quad head' cards usually prove much more expensive than other graphics cards.</p>
<p>If you decide to go down this route we recommend the nVidia NVS 440 graphics cards which are available from a few different vendors. You should be looking to pay around £350 for the card alone.</p>
<p><strong>Install 2nd Graphics Card</strong></p>
<p>Depending on your specific computer you may be able to install a secondary graphics card, if you select one which supports 2 monitor outputs then you will have a system that can output to a four screen display.</p>
<p>The downside of this option is that it is often unclear if you have the right expansion ports inside your computer for an additional graphics card, and if you do you must ensure the cards are compatible with each other. The general rule is that you should stick to the same 'type' of card, i.e. two nVidia GeForce cards or two ATI cards.</p>
<div id="tool-ans-top">
<p>Previous Answers</p>
</div>
<div id="tool-ans-mid">
<ul>
<li>You would like to run a quad display (4 screen) system</li>
<li>You would like to use your current computer to run quad screens</li>
<li>You current computer can support dual monitors</li>
<li>You require dedicated graphics power to each screen</li>
</ul>
<p>If this is incorrect simply <a href="/guide/quad-q3/">go back a step</a> or <a href="/guide/">restart the guide</a>.</p>
</div>
<div id="tool-ans-bot">
</div>
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