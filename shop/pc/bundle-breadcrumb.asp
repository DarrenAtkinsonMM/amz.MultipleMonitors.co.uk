<%
'DA Edit for Bundle Breadcrumb

'Check which page we are on
Select Case Request.ServerVariables("PATH_INFO")
	Case "/shop/pc/CUSTOMCAT-bundles1.asp"
		bunBCpage = "bundle1"
	Case "/shop/pc/CUSTOMCAT-bundles2.asp"
		bunBCpage = "bundle2"
	Case "/shop/pc/CUSTOMCAT-bundles3.asp"
		bunBCpage = "bundle3"
	Case Else
		bunBCpage = "pc"
End Select

'fix missing querystring's - 2nd edit to check for SQL injections
if Request.querystring("mid") = "" Then
	bunBCmid = 0
else
	'Querystring has something check if a number, if not dump user to 404 page
	If IsNumeric(Request.querystring("mid")) Then
		bunBCmid = CInt(Request.querystring("mid"))
	Else
		response.redirect("/404.html")
	End If
end if

if Request.querystring("sid") = "" Then
	bunBCsid = 0
else
	'Querystring has something check if a number, if not dump user to 404 page
	If IsNumeric(Request.querystring("sid")) Then
		bunBCsid = CInt(Request.querystring("sid"))
	Else
		response.redirect("/404.html")
	End If
end if

if Request.querystring("cid") = "" Then
	bunBCcid = 0
else
	'Querystring has something check if a number, if not dump user to 404 page
	If IsNumeric(Request.querystring("cid")) Then
		bunBCcid = CInt(Request.querystring("cid"))
	Else
		response.redirect("/404.html")
	End If
end if


'Work out correct stand
Select Case bunBCsid
	Case 326
		bunStandimg = "/images/bundles/bun-s2v-med.png"
		bunStandTxt = "Dual Vertical"
		bunStandOk = 1
		bunMonNum = "2"
		bunBCdiscount = 25
		bunArrStdImg = "s2v"
	Case 287
		bunStandimg = "/images/bundles/bun-s2h-med.png"
		bunStandTxt = "Dual Horizontal"
		bunStandOk = 1
		bunMonNum = "2"
		bunBCdiscount = 25
		bunArrStdImg = "s2h"
	Case 324
		bunStandimg = "/images/bundles/bun-s3p-med.png"
		bunStandTxt = "Triple Pyramid"
		bunStandOk = 1
		bunMonNum = "3"
		bunBCdiscount = 25
		bunArrStdImg = "s3p"
	Case 312
		bunStandimg = "/images/bundles/bun-s3h-med.png"
		bunStandTxt = "Triple Horizontal"
		bunStandOk = 1
		bunMonNum = "3"
		bunBCdiscount = 25
		bunArrStdImg = "s3h"
	Case 313
		bunStandimg = "/images/bundles/bun-s4s-med.png"
		bunStandTxt = "Quad Square"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4s"
	Case 337
		bunStandimg = "/images/bundles/bun-s4sp-med.png"
		bunStandTxt = "Quad Square"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4sp"
	Case 327
		bunStandimg = "/images/bundles/bun-s4h-med.png"
		bunStandTxt = "Quad Horizontal"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4h"
	Case 325
		bunStandimg = "/images/bundles/bun-s4p-med.png"
		bunStandTxt = "Quad Pyramid"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4p"
	Case 318
		bunStandimg = "/images/bundles/bun-s5p-med.png"
		bunStandTxt = "Five Pyramid"
		bunStandOk = 1
		bunMonNum = "5"
		bunBCdiscount = 50
		bunArrStdImg = "s5p"
	Case 338
		bunStandimg = "/images/bundles/bun-s6r-med.png"
		bunStandTxt = "Six Way"
		bunStandOk = 1
		bunMonNum = "6"
		bunBCdiscount = 100
		bunArrStdImg = "s6r"
	Case 314
		bunStandimg = "/images/bundles/bun-s6rp-med.png"
		bunStandTxt = "Six Way"
		bunStandOk = 1
		bunMonNum = "6"
		bunBCdiscount = 100
		bunArrStdImg = "s6rp"
	Case 319
		bunStandimg = "/images/bundles/bun-s8r-med.png"
		bunStandTxt = "Eight Way"
		bunStandOk = 1
		bunMonNum = "8"
		bunBCdiscount = 100
		bunArrStdImg = "s8r"
	Case else
		bunStandimg = "/images/bundles/bun-question.jpg"
		bunStandOk = 0
		bunBCdiscount = 0
End Select

'Work out correct monitor
Select Case bunBCmid
	Case 315
		bunMonitorimg = "/images/bundles/bun-acersq-med.png"
		bunMonitorTxt = "Acer 17&quot;"
		bunMonitorTxt2 = "Monitors"
		bunMonitorOk = 1
		bunArrMonImg = "a17"
	Case 304
		bunMonitorimg = "/images/bundles/bun-acerwide-med.png"
		bunMonitorTxt = "AOC 21.5&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "a22"
	Case 316
		bunMonitorimg = "/images/bundles/bun-acersq-med.png"
		bunMonitorTxt = "Acer 19&quot;"
		bunMonitorTxt2 = "Monitors"
		bunMonitorOk = 1
		bunArrMonImg = "a19"
	Case 317
		bunMonitorimg = "/images/bundles/bun-acerwide-med.png"
		bunMonitorTxt = "Acer 24&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "a24"
	Case 321
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorTxt = "Iiyama 21.5&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i22"
	Case 328
		bunMonitorimg = "/images/bundles/bun-acerwide-med.png"
		bunMonitorTxt = "Acer 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "a27"
	Case 320
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorTxt = "Iiyama 24&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i23"
	Case 329
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorTxt = "Iiyama 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
 	Case 342
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorTxt = "ViewSonic 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
	Case 344
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorTxt = "AOC 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
 	Case 345
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorTxt = "IIyama 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
	Case else
		bunMonitorimg = "/images/bundles/bun-question.jpg"
		bunMonitorOk = 0
End Select

'Work out correct computer
Select Case bunBCcid
	Case 301
		bunCompimg = "/images/bundles/bun-pro-pc.png"
		bunCompTxt = "Pro PC"
		bunCompOk = 1
	Case 306
		bunCompimg = "/images/bundles/bun-ultra-pc.png"
		bunCompTxt = "Ultra PC"
		bunCompOk = 1
	Case 307
		bunCompimg = "/images/bundles/bun-extreme-pc.png"
		bunCompTxt = "Extreme PC"
		bunCompOk = 1
	Case 333
		bunCompimg = "/images/bundles/bun-trader-pc.png"
		bunCompTxt = "Trader PC"
		bunCompOk = 1
	Case 339
		bunCompimg = "/images/bundles/bun-pro-pc.png"
		bunCompTxt = "Charter PC"
		bunCompOk = 1
   Case 343
		bunCompimg = "/images/bundles/bun-trader-pc.png"
		bunCompTxt = "Trader Pro PC"
		bunCompOk = 1
	Case else
		bunCompimg = "/images/bundles/bun-question.jpg"
		bunCompOk = 0
End Select

'Sort out bundle savings messages
Select Case bunBCdiscount
	Case 25
		bunBCSavTxt1 = "&pound;25 Bundle Discount"
		bunBCSavTxt2 = "Total Savings Of Over &pound;105!"
	Case 50
		bunBCSavTxt1 = "&pound;50 Bundle Discount"
		bunBCSavTxt2 = "Total Savings Of Over &pound;130!"
	Case 100
		bunBCSavTxt1 = "&pound;100 Bundle Discount"
		bunBCSavTxt2 = "Total Savings Of Over &pound;180!"
End Select

if bunBCpage = "bundle1" then
	if bunBCmid = 0 then
	'We have no stand or monitor so show bundle start toolbar
%>
	<!-- Header: intro -->
    <header id="bundle-stands" class="bundle-wrap bg-lyt">
		<div class="intro-content paddingtop-20">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Bundle Deal</h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1">Order everything You need for your new multiple monitor system with one easy and hassle free purchase</h5>
							<p class="grnt-para">We guarantee that you will receive everything you need to instantly get up and running</p>
						</div>
						<div class="wow fadeInUp bd-wrap" data-wow-offset="0" data-wow-delay="0">
							<div class="row">
								<div class="col-md-3 bd-heading">
									<h2 class="color-med h-bold text-uppercase">Bundle <span class="color-focus">Deals</span></h2>
								</div>
								<div class="col-md-9 bd-content">
									<ul class="bd-listing clearfix">
										<li>Up to &pound;100 In Discounts</li>
										<li>Free Speakers (Save &pound;20)</li>
										<li>Free Wifi Card (Save &pound;40)</li>
										<li>Free UK Delivery (Save &pound;20)</li>
										<li>Free Long Length Cables</li>
										<li>Total Savings Of Up To &pound;180!</li>
									</ul>
								</div>
							</div>
						</div>
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message hm1 color-focus">To Get Started Simply Select A Stand For Your Bundle</h3><a name="bundlestart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#bundlestart" id="wg-scrollDown">&#xf107;</a>	
    </header>
	
	<!-- /Header: Bundle -->
<%
	else
	'Means we have a monitor but no stand
		if bunBCcid = 0 then
		'no stand, no pc
%>
	<!-- Header: intro -->
    <header id="bundle-stands" class="bundle-wrap bg-lyt">
		<div class="intro-content paddingtop-20">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Bundle Deal</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-single">
								<div class="row">
									<div class="col-sm-3 bd-product-img">
										<img src="<%=bunMonitorimg%>">
									 </div>
									 <div class="col-sm-9 bd-product-text text-left">
										<h3 class="color-med h-bold">Monitors <span>Selected</span></h3>
										<h5 class="green-link h-bold"><%=bunMonitorTxt%> <span><%=bunMonitorTxt2%></span></h5>
										<a href="/bundles/" class="btn btn-skin semi margintop-20">Restart Bundle Configuration <i class="fa fa-angle-right"></i></a>
									 </div>
								</div>
							</div>
						</div>
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Pick Your Preferred Stand, Select One Below</h3><a name="bundlestart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#bundlestart" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
		else
		'Got a monitor and a PC but no stand
%>
	<!-- Header: intro -->
    <header id="bundle-stands" class="bundle-wrap bg-lyt">
		<div class="intro-content paddingtop-20">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Bundle Deal</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-2col no-bg">
								<div class="row">
									<div class="col-md-7">											
										<div class="row">
											<div class="col-sm-3 bd-product-img">
												<img src="<%=bunMonitorImg%>">
											 </div>
											 <div class="col-sm-9 bd-product-text text-left">
												<h5 class="green-link h-bold margintop-20"><%=bunMonitorTxt%> <span><%=bunMonitorTxt2%></span></h5>
												<a href="/bundles/" class="btn btn-skin semi margintop-10">Restart Bundle Creation <i class="fa fa-angle-right"></i></a>
											 </div>
										</div>
									 </div>
									 <div class="col-md-5 step-bndl">
										<div class="row">
											<div class="col-md-4 col-sm-3 bd-product-img">
												<img src="<%=bunCompImg%>">
											 </div>
											 <div class="col-md-8 col-sm-9 bd-product-text text-left">
												<h5 class="green-link h-bold margintop-20"><%=bunCompTxt%></h5>
												<a href="/bundles/" class="btn btn-skin semi margintop-10">Restart Bundle Creation <i class="fa fa-angle-right"></i></a>
											 </div>
										</div>
									 </div>
								</div>
							</div>
						</div>
						
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Pick Your Preferred Stand, Select One Below</h3><a name="bundlestart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#bundlestart" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
		end if
	end if
end if
if bunBCpage = "bundle2" then
	if bunBCcid = 0 then
	'No monitor, no PC
%>
	<!-- Header: intro -->
    <header id="bundle-screens" class="bundle-wrap bg-lyt">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Bundle Deal</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-single">
								<div class="row">
									<div class="col-sm-3 bd-product-img">
										<img src="<%=bunStandimg%>">
									 </div>
									 <div class="col-sm-9 bd-product-text text-left">
										<h3 class="color-med h-bold">Stand <span>Selected</span></h3>
										<h5 class="green-link h-bold"><%=bunStandTxt%> <span>Synergy Stand</span></h5>
										<a href="/bundles/#arraystart" class="btn btn-skin semi margintop-20">Change Selection <i class="fa fa-angle-right"></i></a>
									 </div>
								</div>
							</div>
						</div>
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Now Add Some Screens, Select Them Below</h3><a name="bundlestart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#bundlestart" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
	else
	'Got stand and PC but no monitor
%>
	<!-- Header: intro -->
    <header id="bundle-screens" class="bundle-wrap bg-lyt">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Bundle Deal</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-2col no-bg">
								<div class="row">
									<div class="col-md-7">											
										<div class="row">
											<div class="col-sm-3 bd-product-img">
												<img src="<%=bunStandImg%>">
											 </div>
											 <div class="col-sm-9 bd-product-text text-left">
												<h5 class="green-link h-bold margintop-20"><%=bunStandTxt%> <span>Synergy Stand</span></h5>
												<a href="/bundles/" class="btn btn-skin semi margintop-10">Restart Bundle Creation <i class="fa fa-angle-right"></i></a>
											 </div>
										</div>
									 </div>
									 <div class="col-md-5 step-bndl">
										<div class="row">
											<div class="col-md-4 col-sm-3 bd-product-img">
												<img src="<%=bunCompImg%>">
											 </div>
											 <div class="col-md-8 col-sm-9 bd-product-text text-left">
												<h5 class="green-link h-bold margintop-20"><%=bunCompTxt%></h5>
												<a href="/bundles/" class="btn btn-skin semi margintop-10">Restart Bundle Creation <i class="fa fa-angle-right"></i></a>
											 </div>
										</div>
									 </div>
								</div>
							</div>
						</div>
						
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Pick Your Preferred Monitors, Select Them Below</h3><a name="bundlestart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#bundlestart" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
	end if
end if
if bunBCpage = "bundle3" then
%>
	<!-- Header: intro -->
    <header id="bundle-screens" class="bundle-wrap bg-lyt">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Bundle Deal</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-multi">
								<h3 class="color-med h-bold text-left marginbot-20">Stand &amp; Screens <span>Selected</span></h3>
								<div class="row">
									<div class="col-sm-3 bd-product-img">
										<img src="/images/bundles/<%=bunArrStdImg%>-<%=bunArrMonImg%>-blg.png">
									 </div>
									 <div class="col-sm-9 bd-product-text text-left">
										<h5 class="green-link h-bold"><%=bunStandTxt%> <span>Synergy Stand</span></h5>
										<a href="/bundles/?sid=&mid=<%=bunBCmid%>&cid=#arraystart" class="btn btn-skin semi margintop-10">Change Selection <i class="fa fa-angle-right"></i></a>
										<h5 class="green-link h-bold margintop-20"><%=bunMonNum%> X <%=bunMonitorTxt%> <span><%=bunMonitorTxt2%></span></h5>
										<a href="/bundles-2/?sid=<%=bunBCsid%>&mid=&cid=" class="btn btn-skin semi margintop-10">Change Selection <i class="fa fa-angle-right"></i></a>
									 </div>
								</div>
							</div>
						</div>
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Finally Select Your Preferred Computer</h3><a name="bundlestart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#bundlestart" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
end if
if bunBCpage = "pc" then
%>
	<!-- Header: intro -->
    <header id="bundle-screens" class="bundle-wrap bg-lyt">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-check-square-o green-link"></i><span> Your Custom</span> Bundle Deal</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-2col no-bg">
								<div class="row">
									<div class="col-md-7">											
										<div class="row">
											<div class="col-sm-3 bd-product-img">
												<img src="/images/bundles/<%=bunArrStdImg%>-<%=bunArrMonImg%>-blg.png">
											 </div>
											 <div class="col-sm-9 bd-product-text text-left">
												<h5 class="green-link h-bold"><%=bunStandTxt%> <span>Synergy Stand</span></h5>
												<a href="/bundles/?sid=&mid=<%=bunBCmid%>&cid=<%=bunBCcid%>#arraystart" class="btn btn-skin semi margintop-10">Change Selection <i class="fa fa-angle-right"></i></a>
												<h5 class="green-link h-bold margintop-20"><%=bunMonNum%> X <%=bunMonitorTxt%> <span><%=bunMonitorTxt2%></span></h5>
												<a href="/bundles-2/?sid=<%=bunBCsid%>&mid=&cid=<%=bunBCcid%>" class="btn btn-skin semi margintop-10">Change Selection <i class="fa fa-angle-right"></i></a>
											 </div>
										</div>
									 </div>
									 <div class="col-md-5 step-bndl">
										<div class="row">
											<div class="col-md-4 col-sm-3 bd-product-img">
												<img src="<%=bunCompImg%>">
											 </div>
											 <div class="col-md-8 col-sm-9 bd-product-text text-left">
												<h5 class="green-link h-bold"><%=bunCompTxt%></h5>
												<a href="/bundles-3/?sid=<%=bunBCsid%>&mid=<%=bunBCmid%>&cid=" class="btn btn-skin semi margintop-10">Change Selection <i class="fa fa-angle-right"></i></a>
											 </div>
										</div>
									 </div>
								</div>
							</div>
						</div>
						
						<div class="wow fadeInUp bd-wrap" data-wow-offset="0" data-wow-delay="0">
							<div class="row">
								<div class="col-md-3 bd-heading">
									<h2 class="color-med h-bold text-uppercase">Bundle <span class="color-focus">Deal</span></h2>
								</div>
								<div class="col-md-9 bd-content">
									<ul class="bd-listing clearfix">
										<li><%=bunBCSavTxt1%></li>
										<li>Free Speakers (Save &pound;20)</li>
                                        <li>Free Wifi Card (Save &pound;40)</li>
										<li>Free UK Delivery (Save &pound;20)</li>
										<li>Free Long Length Cables</li>
										<li><%=bunBCSavTxt2%></li>
									</ul>
								</div>
							</div>
						</div>
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Customise Your PC and order Your Bundle</h3>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#custom-order" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
end if
%>