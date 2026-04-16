<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<!--#include file="shop/includes/settings.asp"-->
<!--#include file="shop/includes/storeconstants.asp"-->
<!--#include file="shop/includes/ErrorHandler.asp"-->
<!--#include file="shop/includes/stringfunctions.asp"-->
<!--#include file="shop/includes/opendb.asp"-->
<!--#include file="shop/includes/adovbs.inc"-->
<!--#include file="shop/includes/languages.asp"--> 
<!--#include file="shop/includes/currencyformatinc.asp"-->
<!--#include file="shop/includes/pcProductOptionsCode.asp"--> 
<!--#INCLUDE file="HomeCode.asp"-->
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
'-------------------------------
' declare local variables
'-------------------------------
dim pcStrHPStyle, pcStrHPDesc, pcIntHPFirst, pcIntHPShowSKU, pcIntHPShowImg, pcIntHPFeaturedCount, pcIntHPFeaturedOrder
dim pcIntHPSpcCount, pcIntHPSpcOrder, pcIntHPNewCount, pcIntHPSNewOrder, pcIntHPBestCount, pcIntHPBestOrder
Dim query, conntemp, rsProducts, rsDisc, pDiscountPerQuantity, pTotalCount
Dim pcv_intBackOrder, pStock, pNoStock, pFormQuantity, pserviceSpec

'*******************************
' LOAD HOMEPAGE SETTINGS
'*******************************
' Refer to "pcadmin/manageHomePage.asp" to see features added to this page

call opendb()

query=  "SELECT pcHPS_FeaturedCount,pcHPS_Style,pcHPS_PageDesc,pcHPS_First,pcHPS_ShowSKU,pcHPS_ShowImg," &_
        "pcHPS_SpcCount,pcHPS_SpcOrder,pcHPS_NewCount,pcHPS_NewOrder,pcHPS_BestCount,pcHPS_BestOrder FROM pcHomePageSettings;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntHPFeaturedCount=rs("pcHPS_FeaturedCount")
	pcStrHPStyle=rs("pcHPS_Style")	
	pcStrHPDesc=replace(rs("pcHPS_PageDesc"),"''","'")
	pcIntHPFirst=rs("pcHPS_First")
	pcIntHPShowSKU=rs("pcHPS_ShowSKU")
	pcIntHPShowImg=rs("pcHPS_ShowImg")
	pcIntHPSpcCount=rs("pcHPS_SpcCount")
	pcIntHPSpcOrder=rs("pcHPS_SpcOrder")
	pcIntHPNewCount=rs("pcHPS_NewCount")
	pcIntHPNewOrder=rs("pcHPS_NewOrder")
	pcIntHPBestCount=rs("pcHPS_BestCount")
	pcIntHPBestOrder=rs("pcHPS_BestOrder")
end if

if pcIntHPFeaturedCount = "" or not validNum(pcIntHPFeaturedCount) or pcIntHPFeaturedCount < 0 then
	pcIntHPFeaturedCount = 3
end if

' // Note: 0 is an acceptable value and it indicates that Specials should not be shown
if pcIntHPSpcCount = "" or not validNum(pcIntHPSpcCount) or pcIntHPSpcCount < 0 then
	pcIntHPSpcCount = 4
end if

' // Note: 0 is an acceptable value and it indicates that New Arrivals should not be shown
if pcIntHPNewCount = "" or not validNum(pcIntHPNewCount) or pcIntHPNewCount < 0 then
	pcIntHPNewCount = 4
end if

' // Note: 0 is an acceptable value and it indicates that Best Sellers should not be shown
if pcIntHPBestCount = "" or not validNum(pcIntHPBestCount) or pcIntHPBestCount < 0 then
	pcIntHPBestCount = 4
end if

set rs=nothing
call closedb()

pShowSKU = pcIntHPShowSKU
if pShowSKU = "" or isNull(pShowSKU) then
	pShowSKU = -1 ' If 0, then the SKU is hidden
end if

pShowSmallImg = pcIntHPShowImg
if pShowSmallImg = "" or isNull(pShowSmallImg) then
	pShowSmallImg = -1 ' If 0, then the Image is hidden
end if

'*******************************
' END LOAD HOMEPAGE SETTINGS
'*******************************


'*******************************
' GET page style
'*******************************
' Load the page style: check to see if a querystring
' or a form is sending the page style.
Dim pcPageStyle

pcPageStyle = LCase(Request.QueryString("pageStyle"))
if pcPageStyle = "" then
	pcPageStyle = LCase(Request.Form("pageStyle"))
end if

if pcPageStyle = "" then
    pcPageStyle = pcStrHPStyle
end if

if pcPageStyle = "" then
	pcPageStyle = LCase(bType)
end if

if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

'*******************************
' GET Featured Products from DB
'*******************************

call openDb()
if session("CustomerType")<>"1" then
	query1= " AND ((categories.pccats_RetailHide)=0)"
else
	query1=""
end if

query="SELECT distinct products.idProduct,products.sku,products.description,products.price,products.listHidden,products.listPrice,products.serviceSpec,products.bToBPrice,products.smallImageUrl,products.noprices,products.stock,products.noStock, products.pcprod_HideBTOPrice,products.formQuantity,pcprod_OrdInHome,products.pcProd_BackOrder,products.pcUrl FROM products,categories_products,categories WHERE products.active=-1 AND products.showInHome=-1 AND products.configOnly=0 AND products.removed=0 AND formQuantity = 0 AND categories_products.idProduct=products.idProduct AND categories.idCategory=categories_products.idCategory AND categories.iBTOhide=0 " & query1 & " order by pcprod_OrdInHome asc"

set rsProducts=server.CreateObject("ADODB.Recordset")
set rsProducts=conntemp.execute(query)
if Err.number <> 0 then
	set rsProducts=nothing
	call closeDb()  
	response.redirect "techErr.asp?error="&Server.UrlEncode("Error db in " & pcStrPageName & " - Error: "&Err.Description)
end If
if NOT rsProducts.eof then
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
end if
set rsProducts = nothing


'*******************************
' Set Total Count
'*******************************
pTotalCount=pcv_intProductCount

'*******************************
' Start: Set variables for "M" display
'*******************************
if pcPageStyle = "m" then
	'Check if customers are allowed to order products
	dim iShow
	iShow=0
	If scOrderlevel=0 then
		iShow=1
	end if
	If scOrderlevel=1 AND session("customerType")="1" then
		iShow=1
	End if
	
	Dim pCnt, pAddtoCart, pAllCnt
	'reset count variables
	pCnt=Cint(0)
	pAllCnt=Cint(0)
	
	'Loop until the total number of products to show
	if pcIntHPFirst<>0 then
		pCnt=pCnt+1
		pAllCnt=pAllCnt+1
		pcTempHPFeaturedCount=pcIntHPFeaturedCount+1
	else
		pcTempHPFeaturedCount=pcIntHPFeaturedCount
	end if

	'// Run through the products to count all products, products with options, and BTO products
	do while (pCnt < pcv_intProductCount) and (pCnt < pcTempHPFeaturedCount)		
		
		pidrelation=pcArray_Products(0,pCnt) '// rsCount("idProduct")
		pserviceSpec=pcArray_Products(6,pCnt) '// rsCount("serviceSpec")	
		pStock=pcArray_Products(10,pCnt) '// rsCount("stock")
		pNoStock=pcArray_Products(11,pCnt) '// rsCount("noStock")
		pcv_intBackOrder=pcArray_Products(15,pCnt) '// rs("pcProd_BackOrder")
		
		pCnt=pCnt+1
		
		' Check which items will have multi qty enabled,
		pcv_SkipCheckMinQty=-1 
		If pcf_AddToCart(pidrelation)=False Then
			pAllCnt=pAllCnt+1
		End If	
		
	loop
	
	pcv_SkipCheckMinQty=0
		
	' If all items on the page are either BTO or have options,
	' do not show the quantity column or the Add to Cart button.						
	if cint(pAllCnt) <> cint(pCnt) then 
		pAddtoCart = 1
	end if
end if	
'*******************************
' End: Set variables for "M" display
'*******************************


'*******************************
' Build the page
'*******************************
%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="Description" content="Multiple Monitors are the UK's only dedicated multi-screen computer and stand specialists. We only build multi-monitor computers and have our own range of multi monitor stands allowing us to offer an unmatched selection of computers, monitor arrays and multi-screen bundles.">
	<META name="Keywords" content="multiple, monitor, dual, triple, quad, display, screen">
    <meta name="author" content="">
    <title>Multi-Screen Computer & Stand Specialists | Multiple Monitors</title>
<%Response.Buffer=True%>
<%Response.charset="iso-8859-1"%>
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
    <!-- css -->
    <link href="css/bootstrap.min.css" rel="stylesheet" type="text/css">
    <link href="css/font-awesome.min.css" rel="stylesheet" type="text/css" />
    <link href="/css/ekko-lightbox.min.css" rel="stylesheet" />
	<link href="css/animate.css" rel="stylesheet" />
    <link href="css/style.css" rel="stylesheet">
    <link href="css/responsive.css" rel="stylesheet">
	<!-- template skin -->
	<link href="css/blue.css" rel="stylesheet">
<%
private const scIncHeader="1"
private const scJQuery="1"
%>
<link rel="canonical" href="https://www.multiplemonitors.co.uk/" />
<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','https://www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-5648327-3', 'auto');
  ga('send', 'pageview');

</script>
</head>
<body id="page-top" data-spy="scroll" data-target=".navbar-custom">
<div id="wrapper">
    <nav class="navbar navbar-custom navbar-fixed-top" role="navigation">
		<div class="top-area">
			<div class="container">
				<div class="row">
					<div class="col-sm-6 col-md-6 topbar-connects">
					<p class="text-left">
						<span class="tb-contact-bx tb-phone"><a class="text-white" href="tel:0845 508 5377"><i class="fa fa-phone"></i>0845 508 5377</a></span>
						<a class="tb-contact-bx tb-mail" href="mailto:sales@multiplemonitors.co.uk"><i class="fa fa-envelope"></i>sales@multiplemonitors.co.uk</a>
					</p>
					</div>
					<div class="col-sm-6 col-md-6 text-right top-user-box">
						<!--#include file="shop/pc/SmallShoppingCart.asp"-->
					</div>
				</div>
			</div>
		</div>
        <div class="container navigation">
			<div class="row">
				<div class="navbar-header page-scroll col-lg-4 col-md-3 col-xs-12">
					<button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-main-collapse">
						<i class="fa fa-bars"></i>
					</button>
					<a class="navbar-brand" href="/">
						<img src="/images/logo.png" alt="" width="353" height="44" />
					</a>
				</div>

				<!-- Collect the nav links, forms, and other content for toggling -->
				<div class="collapse navbar-collapse navbar-right navbar-main-collapse">
				  <ul class="nav navbar-nav">
					<li class="active"><a href="/">Home</a></li>
					<li><a href="/computers/">Computers</a></li>
					<li><a href="/display-systems/">Monitor Arrays</a></li>
					<li><a href="/bundles/">Bundles</a></li>
					<li><a href="/stands/">Stands</a></li>
					<li><a href="/blog/">Blog</a></li>
				  </ul>
				</div>
				<!-- /.navbar-collapse -->
            </div>
        </div>
        <!-- /.container -->
		<div class="top-tagline-bar">
			<div class="container top-tagline">
				<div class="row">
					<p class="tagline-txt col-lg-8 col-sm-12">Multi-screen displays & computers. As seen on BBC's 'Traders: Millions by the Minute'</p>
					<p class="subscribe-txt col-lg-4 hidden-ss color">Get Exclusive Special Offers ! <a href="/pages/email-signup/">SIGN UP NOW</a></p>
				</div>
				<!-- /.container -->
			</div>
		</div>
    </nav>
	

	<!-- Header: intro -->
    <header id="intro" class="intro">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-6">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h1 class="h-light font-light home-heading">Multi-Screen Computers, Stands &amp; Monitors</h1>
							<p class="home-tagline text-uppercase text-white medium">WORKING WITH MULTI-SCREEN EQUIPMENT IS A PROVEN WAY TO INCREASE PRODUCTIVITY AND REDUCE ERRORS</p>
						</div>
						<div class="wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
							<p class="home-head-text text-white text-justify">Make the switch to a multiple monitor setup quickly and easily with our range of multi-monitor computers, multiple screen stands and monitors, all delivered to you in a stress free, ready to go straight out of the box package.</p>
						</div>
						<div class="wow fadeInRight home-contactbx" data-wow-delay="0.1s">
							<h4 class="h-semi text-uppercase h-contact-heading">Contact us today</h4>
							<h3 class="h-light font-light">
								<a href="tel:0845 508 5377" class="hlink-contact hlink-phone text-white"><i class="fa fa-mobile-phone"></i> 0845 508 5377</a>
							</h3>
							<h3 class="h-light font-light">
								<a href="mailto:sales@multiplemonitors.co.uk" class="hlink-contact hlink-mail text-white"><i class="fa  fa-envelope"></i> sales@multiplemonitors.co.uk</a>
							</h3>
						</div>
					</div>
					<div class="col-md-6 home-header-image">
						<div class="wow fadeInRight text-center" data-wow-offset="0" data-wow-delay="0.1s">
							<img class="home-himage" src="/images/pages/trading-image.png" alt="" />
						</div>
					</div>					
				</div>		
			</div>
		</div>	
		<a href="#welcome" id="wg-toplink" class="scroll">&#xf107;</a>		
    </header>
	
	<!-- /Header: intro -->

	<!-- Section: Welcome -->
    <section id="welcome" class="home-section paddingtop-40 paddingbot-40 bg-smog">
		<div class="container">
			<div class="row">
				<div class="col-sm-12 col-md-6">
					<div class="wow fadeInUp" data-wow-delay="0.2s">
						<div class="welcome-box">
							<h2 class="h-light color-med marginbot-0 lineh1">Welcome to</h2>
							<h2 class="h-semi color-med margintop-0">Multiple<span class="color-focus">Monitors</span></h2>
							<p class="home-tagline text-uppercase">The UK's number 1 source for Multi-screen computers & arrays.</p>
							<p class="welcome-para lineh2">We are experts at building multiple monitor capable computers and offer a range of PC and monitor arrays dedicated to this task.<br/>We specialise in just one area to ensure that you get the best possible performance from your equipment, whatever your budget.
							</p>
							<a href="/pages/about-us/" class="btn btn-skin btn-wc">Learn More About Us <i class="fa fa-angle-right"></i></a>
						</div>
					</div>
				</div>
				<div class="col-sm-12 col-md-6">
					<div class="wow fadeInUp" data-wow-delay="0.2s">
						<div class="home-blog">
							<h2 class="h-light color-med marginbot-30">From the <strong class="medium">Blog</strong></h2>
							<div class="home-blog-listing row">
								<div class="col-sm-6">
									<div class="blog-list-col">
										<div class="blog-image">
											<a href="/blog/intel-9th-generation-cpu/"><img src="/images/i7-blog.jpg" /></a>
										</div>
										<div class="blog-detail">
											<h6 class="h-semi"><a href="/blog/intel-9th-generation-cpu/">New 9th Gen. Intel CPU's</a></h6>
											<p>A new Intel chip refresh sees 3 new processor releases including an all new i9 CPU</p>
										</div>
									</div>
								</div>
								<div class="col-sm-6">
									<div class="blog-list-col">
										<div class="blog-image">
											<a href="/blog/introducing-traderspec/"><img src="/images/ts-blog.jpg" /></a>
										</div>
										<div class="blog-detail">
											<h6 class="h-semi"><a href="/blog/introducing-traderspec/">Introducing TraderSpec.com</a></h6>
											<p>Learn all about our new website dedicated to Trading and Technology, launched Feb 2018.</p>
										</div>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
			</div>
		</div>

	</section>
	<!-- /Section: Welcome -->
	<section id="callaction" class="callact-row">	
           <div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="callaction">
							<div class="row">
								<div class="col-md-10">
									<div class="wow fadeInUp" data-wow-delay="0.1s">
									<div class="cta-text">
									<h2 class="h-bold font-light disp-inline">Save Money,</h2>
									<h3 class="h-light font-light disp-inline">Get Free Cables & Free Delivery with a Bundle</h3>
									</div>
									</div>
								</div>
								<div class="col-md-2">
									<div class="wow fadeInRight" data-wow-delay="0.1s">
										<div class="cta-btn">
										<a data-toggle="lightbox" data-title="Multi-Screen Bundles" href="/pop-pages/bundles.htm" class="btn btn-outline">Learn More</a>	
										</div>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
            </div>
	</section>	
	

	<!-- Section: services -->
    <section id="service" class="home-services">
		<div class="container">
			<div class="row">
				<div class="col-sm-6 col-md-4 hs-box">
					<div class="wow fadeInUp" data-wow-delay="0.2s">
						<div class="row">
							<div class="col-xs-12">
								<h6 class="h-semi"><a href="/pages/build-process/">Build Process &amp; Guarantee</a></h6>
							</div>	
							<div class="col-xs-8 hs-detail">
								<p>Once we have your order we will custom build your new array or PC using only the best quality parts available.</p>
								<a class="btn btn-skin btn-sm text-uppercase" href="/pages/build-process/">Learn More</a>
							</div>
							<div class="col-xs-4">
								<img src="img\pc-img.png" alt="Build Process &amp; Guarantee" />
							</div>
						</div>
					</div>
				</div>		
				<div class="col-sm-6 col-md-4 hs-box">
					<div class="wow fadeInUp" data-wow-delay="0.2s">
						<div class="row">
							<div class="col-xs-12">
								<h6 class="h-semi"><a href="/pages/trading-computers/">Trading Computers</a></h6>
							</div>	
							<div class="col-xs-8 hs-detail">
								<p>Our computers and monitor arrays are perfect for traders, see the top benefits of buying a new trading computer</p>
								<a class="btn btn-skin btn-sm text-uppercase" href="/pages/trading-computers/">Learn More</a>
							</div>
							<div class="col-xs-4">
								<img src="img\screen-img.png" alt="Trading Computers" />
							</div>
						</div>
					</div>
				</div>	
				<div class="col-sm-6 col-md-4 hs-box">
					<div class="wow fadeInUp" data-wow-delay="0.2s">
						<div class="row">
							<div class="col-xs-12 hs-detail">
								<h6 class="h-semi"><a href="/shop/pc/support.asp?cmode=1">Multiple Monitors Support Area</a></h6>
								<p>Looking for support on your Multiple Monitors PC or Monitor Array? <strong>Experience our support area now</strong> to view support articles or to request help with a specific issue.</p>
								<a class="btn btn-skin btn-sm text-uppercase" href="/shop/pc/checkout.asp?cmode=1">Learn More</a>
							</div>
						</div>
					</div>
				</div>			
			</div>		
		</div>
	</section>
	<!-- /Section: services -->
		
	<!-- Section: testimonial -->
    <section id="testimonial" class="testimonial-section">
		<div class="carousel-reviews broun-block">
			<div class="container testimonial-wrap">
				<div class="row">
					<h2 class="text-center h-light tm-heading">What our <strong class="medium color-focus">Customers say</strong></h2>
					<div id="carousel-reviews" class="carousel slide" data-ride="carousel">
						<div class="carousel-inner">
							<div class="item active">
								<div class="col-sm-12">
									<div class="tm-details text-center">
										<h3 class="color-med">Computer arrived on sunday and is fantastic!<br /> Very well put together - many thanks.</h3>
										<p class="tm-by color-med">Tom Boszko <span>(Six 19" Array and Intel i7 PC)</span></p>
									</div>
								</div>
							</div>
							<div class="item">
								<div class="col-sm-12">
									<div class="tm-details text-center">
										<h3 class="color-med">The new 6 screen trading station is brilliant, well worth the money and a great investment. It has made my job much easier.</h3>
										<p class="tm-by color-med">Geoff Wheeler <span>(6 x 22" Monitors Array & Extreme i7 PC)</span></p>
									</div>
								</div>
							</div>
							<div class="item">
								<div class="col-sm-12">
									<div class="tm-details text-center">
										<h3 class="color-med">Just a quick email to say my triple screen system arrived yesterday and it looks great. Thanks for all your help.</h3>
										<p class="tm-by color-med">David Coomber <span>(Triple 20" Widescreens)</span></p>
									</div>
								</div>
							</div>
						</div>
						<a class="left carousel-control" href="#carousel-reviews" role="button" data-slide="prev"><span class="glyphicon glyphicon-chevron-left"></span></a>
						<a class="right carousel-control" href="#carousel-reviews" role="button" data-slide="next"><span class="glyphicon glyphicon-chevron-right"></span></a>
					</div>
				</div>
			</div>
		</div>
	</section>
	<!-- /Section: testimonial -->
	<footer>	
		<div class="container">
			<div class="row">
				<div class="col-sm-12 col-md-5 mobi-first-row">
					<div class="row">
						<div class="col-sm-6 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white">Policies &amp; Legal</h5>
									<div class="footer-content">
										<ul class="footer-list">
											<li><a href="/pages/delivery/">Delivery Information</a></li>
											<li><a href="/pages/international/">International Orders</a></li>
											<li><a href="/pages/warranty/">Warranty Information</a></li>
											<li><a href="/pages/returns/">Returns Policy</a></li>
											<li><a href="/pages/privacy-policy/">Privacy Policy</a></li>
											<li><a href="/pages/terms/">Terms &amp; Conditions</a></li>
                                            <li><a href="/remote">Remote Access</a></li>
										</ul>
									</div>
								</div>
							</div>
						</div>
						<div class="col-sm-6 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white">Recently Viewed</h5>
									<div class="footer-content">
										<ul class="footer-list">
											<!--#include file="shop/pc/smallRecentProducts.asp"-->
										</ul>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
				<div class="col-sm-12 col-md-7">
					<div class="row">
						<div class="col-sm-7 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white h-semi">Get In touch</h5>
									<div class="footer-content footer-git">
										<h4 class="h-bold font-light">
											<a href="tel:0845 508 5377" class="hlink-contact hlink-phone text-white"><i class="fa fa-mobile-phone"></i> 0845 508 5377</a>
										</h4>
										<h6 class="h-semi font-light footer-mail">
											<a href="mailto:sales@multiplemonitors.co.uk" class="hlink-contact hlink-mail text-white"><i class="fa  fa-envelope"></i> sales@multiplemonitors.co.uk</a>
										</h6>
									</div>
								</div>
							</div>
						</div>
						<div class="col-sm-5 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white h-semi">Free Buyers Guide</h5>
									<div class="footer-content footer-subscribe">
										<p>Learn all about  multi-screen PC components with our FREE buyers guide.</p>
										<a href="https://www.getdrip.com/forms/206455195/submissions/new" data-drip-show-form="206455195" class="btn footer-submit bg-skin transition-nm medium manual-optin-trigger">Get It Now</a>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
			</div>	
		</div>
		<div class="sub-footer">
		<div class="container">
			<div class="row">
				<div class="col-sm-5">
					<div class="wow fadeInLeft" data-wow-delay="0.1s">
					<div class="text-left">
					<p>&copy;Copyright 2019 - MultipleMonitors Ltd. All rights reserved.</p>
					</div>
					</div>
				</div>
				<div class="col-sm-7">
					<div class="wow fadeInRight" data-wow-delay="0.1s">
					<div class="text-right">
						<ul class="footer-base-menu">
							<li><a href="/">Home</a></li>
							<li><a href="/blog/">Blog</a></li>
							<li><a href="/support/">Support</a></li>
							<li><a href="/pages/about-us/">About Us</a></li>
							<li><a href="/shop/pc/contact.asp">Contact Us</a></li>
							<li><a href="/pages/testimonials/">Testimonials</a></li>
							<li><a href="/pages/site-map/">Site Map</a></li>
						</ul>
					</div>
					</div>
				</div>
			</div>	
		</div>
		</div>
	</footer>

</div>

<!--#include file="shop/pc/inc_footer.asp" -->
<a href="#" class="scrollup"><i class="fa fa-angle-up active"></i></a>
	<!-- Core JavaScript Files -->
    <script src="js/jquery.min.js"></script>	 
    <script src="js/bootstrap.min.js"></script>
    <script src="js/jquery.easing.min.js"></script>
	<script src="js/wow.min.js"></script>
	<script src="js/jquery.scrollTo.js"></script>
    <script src="/js/ekko-lightbox.min.js"></script>
    <script src="js/custom.js"></script>
<!-- Drip -->
<script type="text/javascript">
  var _dcq = _dcq || [];
  var _dcs = _dcs || {};
  _dcs.account = '1043541';

  (function() {
    var dc = document.createElement('script');
    dc.type = 'text/javascript'; dc.async = true;
    dc.src = '//tag.getdrip.com/1043541.js';
    var s = document.getElementsByTagName('script')[0];
    s.parentNode.insertBefore(dc, s);
  })();
</script>

<script src="https://cc.cdn.civiccomputing.com/8.0/cookieControl-8.0.min.js"></script>
<script src="js/cookie.js"></script>
</body>
</html>