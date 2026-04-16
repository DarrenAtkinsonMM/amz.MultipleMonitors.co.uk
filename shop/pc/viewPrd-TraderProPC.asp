<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact, LLC.
'ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC.
'Copyright 2001-2007. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<%
if not request.querystring("sid") = "" then
%>
<!--#include file="bundle-breadcrumb.asp"-->
<%
strPTPadding = " paddingtop-0"
else
strPTPadding = ""
end if
%>
	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content<%=strPTPadding%>">
			<div class="container">
				<div class="row">
					<div class="col-sm-8 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
            <%
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' START:  Show product name 
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            pcs_ProductTitle
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' END:  Show product name 
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            %>
						</div>
					</div>
					<div class="col-sm-4 pt-extras text-right">
						<div class="wow fadeInRight" data-wow-offset="0" data-wow-delay="0.1s">
							<div class="title-declaration text-left">
								<p class="t-declaration-head">UK Assembly &amp; Supply</p>
								<p class="t-declaration-text">If you order from a US or EU supplier <br />you will be liable for VAT &amp; shipping.</p>
							</div>
						</div>
					</div>					
				</div>		
			</div>		
		</div>	
    </header>
	<!--#include file="banner.asp"-->
	<!-- /Header: pagetitle -->
	<!-- Section: product-detail -->
    <section id="product-detail" class="paddingtop-60 paddingbot-40 ">
		<div class="container">
			<div class="row">
				<div class="col-sm-12 col-md-4">
					<div class="wow fadeInUp product-view" data-wow-delay="0">
						<div class="productimage-box">
							<div id="product-zoom">
								<div class="pi-box">
									<div id="productbig-image" class="pi-boxfix">
            <% 
            '*****************************************************************************************************
            ' 4) PRODUCT IMAGES
            '*****************************************************************************************************
            
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' START:  Show Product Image (If there is one)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
            pcs_ProductImage
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' END:  Show Product Image (If there is one)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            %>
									</div>
								</div>
								<div id="product-thumbs"> 
            <%
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' START:  Show Additional Product Images (If there are any)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
            pcs_AdditionalImages
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' END:  Show Additional Product Images (If there are any)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			'EDIT TO FULLY CLOSE PRODUCT IMAGE WRAPPER DIV
			
            '*****************************************************************************************************
            ' END PRODUCT IMAGES
            '*****************************************************************************************************	
            
            
            '*****************************************************************************************************
            ' 15) QUANTITY DISCOUNTS ZONE
            '*****************************************************************************************************
            
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' START:  Show quantity discounts
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            'pcs_QtyDiscounts
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' END:  Show quantity discounts
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            
            '*****************************************************************************************************
            ' END QUANTITY DISCOUNTS ZONE
            '*****************************************************************************************************
            %>
								</div>
								<p class="text-center pz-info"><dfn>(Click to see larger image and other views)<dfn></p>
							</div>
						</div>
						<div class="further-info">
							<p class="medium rubric marginbot-10">Further Information:</p>
							<a href="#tech" class="text-underline">Full Computer Spec &amp; Customisation Options</a>
							<a href="#learn" class="text-underline">Learn More About This Computer</a>
						</div>
					</div>
				</div>
        <% ' START RIGHT COLUMN %>
                        <%
                        '*****************************************************************************************************
                        ' 2) GENERAL INFORMATION
                        '*****************************************************************************************************
                        %>
				<div class="col-sm-12 col-md-8">
					<div class="wow fadeInUp" data-wow-delay="0">
						<div class="product-details">
							<p class="rubric bold marginbot-10">Description:</p>
							<div class="product-details-txt">
                        <%
                        pcs_ShowDetailsTop
						%>
							<p><a href="#learn" class="text-underline">Learn More About The Trader Pro PC</a></p>
                            </div>
							<div class="product-price">
		                        <label class="media-middle">Price:</label> <h3 class="price-info disp-inline h-semi color media-middle"><% pcs_ProductPricesNoVat %> + VAT</h3><span class="media-middle vat-info">(<% pcs_ProductPrices %> inc. VAT)</span>
							</div>
							<div class="delivery-details bg-smog">
								<p class="rubric bold marginbot-10">Delivery Details</p>
                                <%
								'Work out Delivery string
								if daFunDelDateBlockTest(1,0) then
									daDelEstimate = "Due to a short workshop closure, orders will now be delivered on <strong class=""color"">" & daFunDelDateReturn(1,0) & "</strong>."
								 'daDelEstimate = "Due to the Christmas and New Year holidays, deliveries will now be made after <strong class=""color"">" & daFunDelDateReturn(1,0) & "</strong>."
								Else
									daDelEstimate = "Order before <strong class=""color"">" & daFunDelCutOff() & "</strong> for delivery on <strong class=""color"">" & daFunDelDateReturn(1,0) & "</strong>"
								end if
								%>
                        		<p class="dcategory"><label class="rubric semi space-right10">UK :</label> <strong class="color">&pound;10</strong> | <%=daDelEstimate%></p>
                        		<p class="dcategory"><label class="rubric semi space-right10">International :</label> International shipping from just &pound;20 - <a data-toggle="lightbox" data-title="International Delivery" class="text-underline" href="/pop-pages/int-del-pop.asp?ProdID=<%=pidProduct%>">View Costs / Timescales</a></p>
								<p class="dcategory"><em>(UK buyers: Saturday delivery is available in the checkout. You can also email us after placing an order to request a specific delivery date, any date after the above estimate is possible.)</em></p>
							</div>
							<a class="btn btn-skin btn-wc semi order-btn margintop-30" href="#tech">Customise &amp; Order Your New PC <i class="fa fa-angle-right"></i></a>
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
								<div class="col-lg-6 cta-gqrow">
									<div class="wow fadeInUp" data-wow-delay="0.1s">
										<div class="row">
											<div class="col-lg-8 col-sm-7 cta-gqcolumn">
												<div class="cta-text double-linetext gq-headings">
													<h2 class="h-bold font-light disp-inline cta-txtline">Got a Question</h2>
													<h5 class="font-light disp-inline cta-txtline">About This Trading PC?</h5>
												</div>
											</div>
										</div>
									</div>
								</div>
								<div class="col-lg-6">
									<div class="wow fadeInRight" data-wow-delay="0.1s">
										<div class="row">
											<div class="col-lg-5 col-sm-6 cta-email cta-2line cta-gqcolumn">
												<i class="fa fa-envelope-o cta-icon"></i><a href="javascript:;" class="twoline-link linkpre-mail">Send us an <strong>Email enquiry</strong></a>
											</div><a name="learn" id="learn"></a>
											<div class="col-lg-7 col-sm-6 cta-2phone">
												<div class="cta-text double-linetext cta-2line cta-gqcolumn">
													<i class="fa fa-phone cta-icon"></i><label class="text-white">Call us on</label>
													<h2 class="h-bold font-light"><a class="hlink-contact hlink-phone text-white" href="tel:03302236655"> 0330 223 66 55</a></h2>
												</div>
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

	<!-- Section: custom-order -->
    <section id="fortrader" class="viewfor-traders bg-smog paddingbot-70">
		<div class="container">
			<div class="row">
				<div class="col-sm-12 lr-detail-wrap">
					<div class="row lr-titleRow paddingbot-60 wow fadeInUp" data-wow-delay="0">
						<div class="lr-col-title col-xs-12">
							<h2 class="h-semi color-med">Perfect for Traders, <span class="color disp-inline">Here's why</span></h2>
							<p class="bigtxt-para text-justify"><i class="fa fa-line-chart lr-titleicon color"></i>The Trader Pro PC offers an unmatched level of performance for professional traders and power users of the various trading platforms. With a default spec faster than all previous generation i9’s this really is a top trading computer. <br /><br />Read on to see how this computer will bring to your trading to a whole new level:</p>
						</div>
					</div>
					<div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline fstrow-img">
								<img src="/images/trader-pc/intel-12th.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6 paddingtop-0">
							<h3 class="h-bold color-med">Fast Processors & RAM</h3>
							<h3 class="color-med lr-subtytl">Latest Intel Core Ultra CPUs</h3>
							<p>The biggest factor in how fast your trading software will run is your CPU.</p>
							<p>In this Trader Pro PC we use the fastest Intel processors, the standard Core Ultra 5 245KF CPU matches the raw speed of a 14th generation i9 processor. This chip would comfortably handle virtually any trading workload.</p>
							<p>If that’s not enough performance for you then go for the new Core Ultra 7 or 9 CPUs for even faster performance. The Core Ultra 9 is only recommended if you have particularly demanding multi-threaded workloads.</p>
                            <p>32GB of ultra-fast DDR5 RAM is perfect for most charting and trading applications however you have the option of increasing this even further if you anticipate running more platforms, charts, or lots of web browser tabs simultaneously.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-even wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/nvidia.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Responsive &amp; Easy To Setup</h3>
							<h3 class="color-med lr-subtytl">Multi-Screen Graphics Capability</h3>
							<p>Many traders want to run multi-screen setups so that they can see their different charts all at once without having to constantly flick between screens.</p>
							<p>With a Trader Pro PC this is made super simple, simply select how many monitors you want to connect in the options below and we will build the PC for you with the right number of monitor ports on the back of it.</p>
							<p>By default you get a four screen PC capable of running up to four 4K or QHD screens. If you need more screens or graphics power then select the 4, 6, 8, 10 or 12 screen option below.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/displayfusion.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Display Fusion</h3>
							<h3 class="color-med lr-subtytl">Multi-Monitor Software</h3>
							<p>When you have a lot of monitors it can become time consuming and even frustrating trying to get the right charts into the right screens.</p>
							<p>Every new Trader Pro PC comes with our exclusive version of DisplayFusion, a suite of tools that integrate with Windows and allow full control over all of your programs, charts and screens.</p>
							<p>Want a specific chart to open in a specific screen? Divide your screens into regions? Or quickly jump your charts into the right size and position, all with just a couple of clicks?</p>
                            <p>DisplayFusion makes it all easy, and is pre-installed ready to go. </p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-even wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/silent-pc.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Silence As Standard</h3>
							<h3 class="color-med lr-subtytl">On All Trader Pro PC's</h3>
							<p>Noise levels are often overlooked and can be hard to estimate before using a computer, however controlling noise is important. Who wants to sit next to a loud computer through long trading sessions?</p>
							<p>As a company policy, all of our computers are really quiet and the Trader Pro PC is no different.</p>
							<p>Ultra-low noise power supplies, silent graphics cards, silent solid state hard drives, along with manual fan tuning ensure that all you'll hear from your Trader Pro is a faint hum at the worst.</p>
                            <p class="detail-callout">Other companies can charge up to &pound;50 extra for a quiet PC build, we include it free of charge.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/free-kit.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">All The Extra Kit</h3>
							<h3 class="color-med lr-subtytl">You Need Included</h3>
							<p>Getting setup with a new computer often means buying extra kit on top of the standard PC unit itself, this can quickly add to the cost of your new system.</p>
							<p>With a Trader Pro we can include pretty much anything you're going to need to get up and running.</p>
							<p>Options include a mouse and keyboard set, Wifi cards and speakers. We can also include adapters to allow you to connect HDMI or DVI screens if you request them after ordering your PC.</p>
                            <p class="detail-callout">Other companies charge between &pound;60 - &pound;100 for this extra kit, we supply it at cost pricing.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-even wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/windows-11.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Windows 11</h3>
							<h3 class="color-med lr-subtytl">Fully Optimised OS</h3>
							<p>Great computer performance and responsiveness is down to more than just the speed of the components, the software setup and configuration can massively impact how fast a computer runs.</p>
							<p>Windows 11 is the main software on your computer, it launches all your trading packages and interfaces, and constantly runs in the background. </p>
							<p>There are many options which can impact how fast Windows performs, we manually adjust settings based on your computers hardware to ensure that it runs at optimal levels right out of the box.</p>
                            <p class="detail-callout">Other companies can charge up to &pound;40 extra for Windows optimisation, we do it free of charge.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/build-delivery.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Quick Build Service</h3>
							<h3 class="color-med lr-subtytl">Get Your Machine Quicker</h3>
							<p>All Trader Pro PC's are custom built for each order, this allows you to select the exact configuration that meets your trading needs rather than just buying what's available at the time.</p>
							<p>One of our experienced, expert technicians will build up the computer for you, pre-install and optimise your Windows and Display Fusion installations, and then put it on a thorough 32-hour stress test to make sure it reaches you in perfect working order.</p>
							<p>Our efficient build and test process takes a maximum of 4 - 5 working days, with delivery made on the next day after this.</p>
                            <p class="detail-callout">Other companies charge between &pound;70 - &pound;300 extra for a 4 - 5 day build time, we do this free of charge.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-even wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/onsite-support.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Un-Matched Support</h3>
							<h3 class="color-med lr-subtytl reducefont-sub">Onsite &amp; Unlimited Remote Support As Standard</h3>
							<p>Traders, perhaps more than most, rely on their computers to make money, a loss of access to the markets could be very costly, that's why ensuring you get great support for your new computer should be a priority.</p>
							<p>All Trader Pros come with our standard 5 year hardware cover, the first year is our unique OnSite / Replacement / Collect service, which means most hardware faults can be resolved without you sending the PC back for repair.</p>
							<p>We also offer lifetime email, phone and remote access support for those times when something goes wrong with your Windows installation (which can happen to anyone, at any time).</p>
						</div>
					</div> <!-- lr-detailRow -->
                    <a name="tech"></a><a name="custom-order"></a>	
				</div>		
			</div>		
		</div>
	</section>
	<!-- /Section:  -->

<input type="hidden" id="productid" name="idproduct1" value="343">
<input type="hidden" id="productqty" name="QtyM343" value="1">
<%=formBundleOptions%>

<input type="hidden" name="OptionGroupCount" value="16">
<!-- Section: custom-order -->
    <section id="traderOptions" class="traderOptions paddingtop-30 paddingbot-40">
		<div class="container">
			<div class="row">
				<div class="col-md-8 traderOptions-wrap">
					<div class="row wow fadeInUp" data-wow-delay="0">
						<div class="lr-col-title col-xs-12">
							<h3 class="h-semi color-med">Choose Your Options</span></h3>
							<p class="bigtxt-para text-justify">We have a selection of upgrades available to add functionality and further improve the performance of your Trader Pro PC:</p>
						</div>
					</div>
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">CPU / Processor</span></h5>
							<p class="text-justify">The heart of your computer, your processor has a direct impact on the speed of your computer and its ability to run multiple programs simultaneously.</p>
							<div class="get-traderOptions">
								<label class="color">CPU / Processor:</label>
								<select id="idOption1" name="idOption1" class="spec-dd" onchange="reCalc();flashCPU();">
                                    <option value="title" class="spec-dd-dis" disabled="">Intel Core Ultra CPUs:</option>
                                    <option value="18497" id="4" title="0">Intel Core Ultra 5 245KF // 4.2 - 5.2GHz // 14C- 14T</option>
                                    <option value="18498" id="5" title="95">Intel Core Ultra 7 265KF // 3.9 - 5.5GHz // 20C - 20T + &pound;95</option>
									<option value="18499" id="6" title="325">Intel Core Ultra 9 285KF // 3.7 - 5.7GHz // 24C - 24T + &pound;325</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
                    <div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">RAM / Memory</span></h5>
							<p class="text-justify">RAM determines how many programs and charts you can open simultaneously without slowing down your computer.</p>
							<div class="get-traderOptions">
								<label class="color">RAM / Memory:</label>
								<select id="idOption2" name="idOption2" class="spec-dd" onchange="reCalc();flashRAM();">
									<option value="18416" id="0" selected title="0">32GB DDR5 5,200MHz</option>
									<option value="18417" id="1" title="295">64GB DDR5 5,200MHz + &pound;295</option>
                                    <option value="18415" id="2" title="875" >128GB DDR5 4,400MHz - &pound;875</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Graphics Setup / Number Of Screens Supported</span></h5>
							<p class="text-justify">A standard Trader Pro PC can power up to four 4K, QHD or FHD monitors. Change this option to support more screens, the Monitor &amp; Resolution panel shows supported resolutions and ports.</p>
						<p class="text-justify">Graphics cards can also help run more demanding graphics workloads and even power AI models, change the options and see the impact in the star ratings opposite.</p>
							<div class="get-traderOptions">
								<label class="color">Monitor Connections:</label>
								<select id="idOption4" name="idOption4" class="spec-dd" onchange="reCalc();flashGPU();">
									<option value="title" class="spec-dd-dis" disabled>Up To 4 Monitor Capable:</option>
									<option value="18459" id="1" title="0" selected>Up to 4 screens - Intel Arc A380 (6GB)</option>
									<option value="18512" id="2" title="145">Up to 4 screens - nVidia RTX 5060 (8GB) + &pound;145</option>
									<option value="18514" id="9" title="395">Up to 4 screens - nVidia RTX 5070 (12GB) + &pound;395</option>
									<option value="title" class="spec-dd-dis" disabled>Up To 6 Monitor Capable:</option>
                                    <option value="18460" id="4" title="65">Up to 6 screens - Intel Arc A380 (6GB) &amp; Intel UHD + &pound;65 </option>
									<option value="title" class="spec-dd-dis" disabled>Up To 8 Monitor Capable:</option>
									<option value="18461" id="5" title="175">Up to 8 screens - Intel Arc A380 (6GB) x2 + &pound;175</option>
									<option value="18513" id="6" title="425">Up to 8 screens - nVidia RTX 5060 (8GB) x2 + &pound;425</option>
									<option value="title" class="spec-dd-dis" disabled>Up To 10 Monitor Capable:</option>
									<option value="18462" id="7" title="225">Up to 10 screens -  Intel Arc A380 (6GB) x2 & Intel UHD + &pound;225</option>
									<option value="title" class="spec-dd-dis" disabled>Up To 12 Monitor Capable:</option>
									<option value="18503" id="8" title="465">Up to 12 screens - nVidia A400 (4GB) x3 + &pound;465</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Hard Drive Capacity</span></h5>
							<p class="text-justify">Your hard drive is where your software and data is stored. 1TB is more enough for Windows and your trading platform installations. Increase for extra file storage capacity.</p>
							<div class="get-traderOptions">
								<label class="color">Hard Drive:</label>
								<select id="idOption3" name="idOption3" class="spec-dd" onchange="reCalc();flashSSD();">
									<option value="18520" id="0" selected title="0">500GB WD NVMe M.2 SSD – (5000MBs/4000MBs)</option>
                                    <option value="18343" id="1" title="85">1TB Kingston NVMe M.2 SSD – (6000MBs/4000MBs) + &pound;85</option>
                                   <option value="18345" id="2" title="175">2TB WD NVMe M.2 SSD – (7250MBs/6900MBs) + &pound;175</option>
									<option value="18458" id="3" title="375">4TB Kingston NVMe M.2 SSD – (3500MBs/2800MBs) + &pound;375</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->

					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Second Hard Drive</span></h5>
							<p class="text-justify">Add a second hard drive if you have larger file storage requirements. Traditional style drives are slower and can increase the noise level of your computer. SSD drives are ultra fast and silent.</p>
							<div class="get-traderOptions">
								<label class="color">Second Hard Drive:</label>
								<select id="idOption14" name="idOption14" class="spec-dd" onchange="reCalc();flashHDD2();">
									<option value="18356" id="0" title="0">Not Required</option>
									<option value="title" class="spec-dd-dis" disabled="">Fast &amp; Silent SSDs:</option>
									<option value="18348" id="7" title="125">1TB Adata NVMe M.2 SSD – (3500MBs/3000MBs) + &pound;125</option>
									<option value="18350" id="8" title="185">2TB Adata NVMe M.2 SSD – (3500MBs/3000MBs) + &pound;185</option>
									<option value="18353" id="9" title="445">4TB Kingston NVMe M.2 SSD – (3500MBs/2800MBs) + &pound;445</option>
									<option value="title" class="spec-dd-dis" disabled="">Traditional Hard Drives:</option>
									<option value="18352" id="4" title="105">4TB Traditional Hard Drive + &pound;105</option>
									<option value="18355" id="5" title="140">6TB Traditional Hard Drive + &pound;140</option>
									<option value="18483" id="6" title="160">8TB Traditional Hard Drive + &pound;160</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Bootable Backup Drive</span></h5>
							<p class="text-justify">This is a backup solution which clones your 'C' drive on a schedule to a dedicated extra internal hard drive and allows for easy and quick recovery from many system issues. Extra drive included in the price. </p>
							<div class="get-traderOptions">
								<label class="color">Backup Drive:</label>
								<select id="idOption11" name="idOption11" class="spec-dd" onchange="reCalc();flashBBD();">
									<option value="18384" id="0" title="0">Not Required</option>
									<option value="18382" id="1" title="145">Bootable Backup Drive Solution + &pound;145</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->

					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Mouse &amp; Keyboard Set </span></h5>
							<p class="text-justify">Any standard mouse and keyboard will work with your Trader Pro PC, you can add either a wired or wireless set to your order here.</p>
							<div class="get-traderOptions">
								<label class="color">Inputs:</label>
								<select id="idOption9" name="idOption9" class="spec-dd" onchange="reCalc();flashKYB();">
									<option value="18371" id="0" title="0">Not Required</option>
                                    <option value="18369" id="1" title="20">Logitech Wired Mouse / Keyboard Set + &pound;20</option>
                                    <option value="18370" id="2" title="25">Logitech Wireless Mouse / Keyboard Set + &pound;25</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Speakers</span></h5>
							<p class="text-justify">Desktop computers do not generally have built in speakers, select whether you would like some suppling with your Trader Pro.</p>
							<div class="get-traderOptions">
								<label class="color">Speakers:</label>
								<select id="idOption6" name="idOption6" class="spec-dd" onchange="reCalc();flashSPK();">
                                	<% 
									if request.querystring("sid") <> "" Then
									%>
									<option value="18362" id="1" title="0">Free Desktop Speakers (Worth &pound;20)</option>
                                    <%
									else
									%>
                                    <option value="18363" id="0" title="0">Not Required</option>
                                    <option value="18361" id="1" title="20">Desktop Speaker Set + &pound;20</option>
                                    <%
									end if
									%>                                  
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
                    <div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Wireless Network Card</span></h5>
							<p class="text-justify">All computers come with a wired network port, if you need a WiFi connection you can add one here. WiFi AX is the fastest connection type and all cards now include a free Bluetooth adapter.</p>
							<div class="get-traderOptions">
								<label class="color">WiFi Card:</label>
								<select id="idOption8" name="idOption8" class="spec-dd" onchange="reCalc();flashWIFI();">
                                	<% 
									if request.querystring("sid") <> "" Then
									%>
                                    <option value="18366" id="1" title="0">Free Wireless AX - 3,000Mbps inc. Bluetooth (Worth &pound;40)</option>
                                    <%
									else
									%>
									<option value="18367" id="0" title="0">Not Required</option>
                                    <option value="18368" id="1" title="40">Wireless AX - 3,000Mbps inc. Bluetooth + &pound;40</option>
                                    <%
									end if
									%>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->




					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Operating System</span></h5>
							<p class="text-justify">Select between Windows 11 Home and Professional editions.</p>
							<div class="get-traderOptions">
								<label class="color">Operating System:</label>
								<select id="idOption10" name="idOption10" class="spec-dd" onchange="reCalc();flashWIN();">
									<option value="18374" id="3" title="0">Windows 11 Home</option>
									<option value="18375" id="4" title="45">Windows 11 Professional + &pound;45</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Microsoft Office</span></h5>
							<p class="text-justify">Microsoft Office Home Edition gives you Word, Excel, PowerPoint & OneNote, if you require Outlook go for the Business Edition. This is a 1 PC, lifetime license.</p>
							<div class="get-traderOptions">
								<label class="color">Microsoft Office:</label>
								<select id="idOption12" name="idOption12" class="spec-dd" onchange="reCalc();flashMSO();">
									<option value="18378" id="4" title="0">Not Required</option>
									<option value="18377" id="6" title="105">Home Edition 2024 + &pound;105</option>
                                    <option value="18376" id="8" title="195">Home & Business Edition 2024 + &pound;195</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row otLast-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Hardware Support Warranty</span></h5>
							<p class="text-justify">All Trader Pro PC's come with 5 year hardware cover as standard. The first year is an Onsite / Replacement / Collect service, for extra peace of mind this can be extended for 2 or 3 years.</p>
							<div class="get-traderOptions">
								<label class="color">Hardware Support:</label>
								<select id="idOption13" name="idOption13" class="spec-dd" onchange="reCalc();flashWAR();">
									<option value="18379" id="7" title="0">5 Year (1 Year OnSite / Replacement / Collect)</option>
									<option value="18380" id="8" title="75">5 Year (2 Year OnSite / Replacement / Collect) + &pound;75</option>
                                    <option value="18381" id="9" title="150">5 Year (3 Year OnSite / Replacement / Collect) + &pound;150</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
				</div>	
				<div class="col-lg-4 col-sm-12 spec-box-wrap">

               						<div id="cust-monitors" class="spec-box spec-box2 wow fadeInRight" data-wow-delay="0.1s"><!-- spec-box2 -->
										<div class="spec-custom">
											<h5 class="specbox-heading">Trading Performance Levels</h5>
											<div class="spec-content">
												<table width="100%" cellpadding="0" cellspacing="0" class="star-ratings marginbot-10">
													<tbody>
														<tr>
									 						<td><strong>CPU Speed: </strong><span class="star-desc">The raw CPU speed, a big factor in trading software performance levels.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0"><span id="stars-speed"></span></td>
                                                        </tr>
														<tr>
															<td><strong>Multi-Tasking: </strong><span class="star-desc">Ability to run multiple trading platforms and software simultaneously without slowing down system performance.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0"><span id="stars-multi"></span></td>
                                                        </tr>
                                                        <tr>
															<td><strong>Multi-Threading: </strong><span class="star-desc">Important for back-testing and some of the more intensive platforms.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0"><span id="stars-mulThr"></span></td>
                                                        </tr>
														<tr>
														<td><strong>Graphics Power: </strong><span class="star-desc">How well it copes with more graphically demanding programs and apps.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0"><span id="stars-GPU"></span></td>
                                                        </tr>
														<tr>
															<td class="paddingbot-10"><strong>AI Performance: </strong><span class="star-desc">The TOPS score for this graphics card (higher is better) is:&nbsp;&nbsp;<span class="spec-tops"><span id="stars-gputops">66</span></span></td>
														</tr>
														<tr>
															<td><strong>Quietness: </strong><span class="star-desc">Noise levels in standard use. <br />10 stars = faint hum, 1 star = jet engine.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0"><span id="stars-quiet"></span></td>
                                                        </tr>
													</tbody>
												</table>
												<a data-toggle="lightbox" data-title="Computer Ratings Explained" class="specbox-link" href="/pop-pages/custpc-tradingstars.htm">Learn More About These Ratings</a>
											</div>
											
										</div>
									</div><!-- spec-box2 end -->
					<div id="cust-monitors" class="spec-box spec-box2 wow fadeInRight" data-wow-delay="0.1s"><!-- spec-box1 -->
										<div class="spec-custom">
											<h5 class="specbox-heading"><span id="optScreensTitle">Monitors &amp; Resolutions Supported</span></h5>
											<div class="spec-content">
												<p><strong>Supported Screen Resolutions:</strong></p>
												<ul class="specbox-list">
													<span id="optRes"></span>
												</ul>
												<p><strong>Available Monitor Ports:</strong></p>
												<span style="font-weight:bold;" id="optScreens"></span>
												<ul class="specbox-list">
													<span id="optPorts"></span>
												</ul>
												<span id="optSpecificPorts"></span>
												<small class="spec-infotxt">*Use the 'Graphics Card Setup' to change this.</small>
											</div>
											
										</div>
									</div><!-- spec-box2 end -->
					<div id="traderspec" class="spec-box spec-box2 wow fadeInRight" data-wow-delay="0.1s">
						<div class="trader-spec-box spec-box-ts">
							<h5 class="color">Full Trader Pro Specifications</h5>
							<div class="spec-content">
								<p><span id="txtCPU"></span></p>
								<p><span id="txtRAM"></span></p>
								<p><span id="txtMB"></span></p>
								<p><span id="txtGPU"></span></p>
								<p><span id="txtSSD"></span></p>
								<span id="txtHDD2"></span>
								<span id="txtBBD"></span>
								<p>Corsair 3000D Case</p>
								<p><span id="txtPSU"></span></p>
								<span id="txtDVD"></span>
								<span id="txtWIFI"></span>
								<p>Gigabit Ethernet LAN Adapter</p>
								<p>8 Channel High Definition Audio Sound Card</p>
								<p>3 x USB 3, 3 x USB 2 &amp; 1 x USB Type-C Ports</p>
								<p><span id="txtCPUCool"></span></p>
								<p id="pKYB"><span id="txtKYB"></span></p>
								<span id="txtSPK"></span>
								<span id="txtBT"></span>
								<p><span id="txtWIN"></span></p>
								<span id="txtMSO"></span>
								<p><span id="txtWAR"></span></p>
								<p class="uppricefont"><strong class="">Trader Pro Price:</strong> <strong class="color">&pound;<span id="pcPrice"></span></strong> <strong class="pri1">+ VAT</strong></p>
                                <span id="txtBunPrice"></span>
								  <p>(&pound;<span id="vatPrice"></span> inc. VAT)</p><br>
								  <div align="center">
                                  <input type="submit" value="ORDER YOUR TRADER PRO NOW" class="btn btn-skin btn-sm text-uppercase" />
							</div>

						</div>
					</div>
				</div>	
			</div>		
		</div>
        </div>
	</section>
	<!-- /Section:  -->
    <div id="footer-appear">&nbsp;</div>