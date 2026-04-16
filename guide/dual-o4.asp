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
<title>New Multiple Monitors Graphics Card Required | Multiple Monitors</title>
<meta name="Description" content="This is the outcome page taken from the Multiple Monitors Guide displayed when a new graphics card is required for your computer."/>
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
<h1>New Graphics Card Required</h1>
<p>Based on the information you supplied the only option to upgrade your current computer to support dual monitors is to replace or add a new graphics card into your system.</p>
<p>The good news is that the majority of new graphics cards on sale have the ability to output to two distinct monitors built directly into them and prices can start as low as just &pound;35.</p>
<p><strong>Specific Requirements</strong></p>
<p>Without knowing your exact model of computer and the hardware currently inside it, it would be impossible for us to suggest a specific solution for you.</p>
<p>We recommend you check with your computer documentation or with the place you purchased your PC from to see what options you have in terms of upgrading the graphics card.</p>
<p><strong>Typical Example</strong></p>
<p>Most modern computers have what is called a PCI-E internal graphics port inside them. If you have one of these ports then there are many graphics cards available which will support dual monitors.</p>
<p>Simply buy a new PCI-E graphics card with dual outputs on it and install into your PC as directed by the instructions supplied with the card.</p>
<p>We aim to sell a range of recommended graphics cards soon however in the mean time, if you are stuck and need further help feel free to <a href="/pop-contact.aspx?KeepThis=true&TB_iframe=true&height=334&width=576" class="thickbox">contact us</a> and we will do our best to guide you through this process.</p>
<div id="tool-ans-top">
<p>Previous Answers</p>
</div>
<div id="tool-ans-mid">
<ul>
<li>You would like to run a dual display (2 screen) system</li>
<li>You would like to use your current computer to run dual monitors</li>
<li>You current computer can not support dual monitors</li>
<li>You require dedicated graphics power to each screen</li>
</ul>
<p>If this is incorrect simply <a href="/guide/dual-q3/">go back a step</a> or <a href="/guide/">restart the guide</a>.</p>
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