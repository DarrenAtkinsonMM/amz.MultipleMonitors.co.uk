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
								<a href="tel:03302236655" class="hlink-contact hlink-phone text-white"><i class="fa fa-mobile-phone"></i> 0330 223 66 55</a>
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
											<a href="/blog/new-intel-amd-cpu-2025/"><img src="/images/intel-amd-cup-blog.jpg" /></a>
										</div>
										<div class="blog-detail">
											<h6 class="h-semi"><a href="/blog/new-intel-amd-cpu-2025/">New Intel &amp; AMD CPUs.</a></h6>
											<p>Intel's new Core Ultra CPU's and the AMD 9000 series are here, but are they any good?</p>
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
											<p>Learn all about our new website dedicated to Trading and Technology...</p>
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
								<h6 class="h-semi"><a href="/pages/dual-monitor-pc/">Dual Monitor PC</a></h6>
								<p>Discover our range of dual monitor capable computers, ready to use straight out of the box. We also offer a range of dual screen bundles which include PC, stand, screens and cables.</p>
								<a class="btn btn-skin btn-sm text-uppercase" href="/pages/dual-monitor-pc/">Learn More</a>
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

<!--#include file="shop/pc/orderCompleteTracking.asp"-->
<!--#include file="shop/pc/inc-Cashback.asp"-->
<!--#include file="shop/pc/footer_wrapper.asp"-->