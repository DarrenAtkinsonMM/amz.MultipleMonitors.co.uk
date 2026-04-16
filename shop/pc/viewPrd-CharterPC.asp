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
							<p><a href="#learn" class="text-underline">Learn More About The Charter PC</a></p>
                            </div>
							<div class="product-price">
		                        <label class="media-middle">Price:</label> <h3 class="price-info disp-inline h-semi color media-middle"><% pcs_ProductPricesNoVat %> + VAT</h3><span class="media-middle vat-info">(<% pcs_ProductPrices %> inc. VAT)</span>
							</div>
							<div class="delivery-details bg-smog">
								<p class="rubric bold marginbot-10">Delivery Details</p>
                                <%
								'Work out Delivery string
								if daFunDelDateBlockTest(1,0) then
									'daDelEstimate = "Due to the Easter break, orders will be delivered on <strong class=""color"">" & daFunDelDateReturn(1,0) & "</strong>."
								 daDelEstimate = "Due to  the Christmas break, delivery will now be around <strong class=""color"">" & daFunDelDateReturn(1,0) & "</strong>."
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
							<p class="bigtxt-para text-justify"><i class="fa fa-line-chart lr-titleicon color"></i>The Charter PC is a fantastic choice for traders looking to purchase their first trading computer, or for more experienced traders running common web trading platforms like IG or CMC. It is also a great option for users of MetaTrader 4 / 5. <br /><br />Read on to see how this computer will bring to your trading to a whole new level:</p>
						</div>
					</div>
					<div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline fstrow-img">
								<img src="/images/trader-pc/intel-10th.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6 paddingtop-0">
							<h3 class="h-bold color-med">Fast Processors & RAM</h3>
							<h3 class="color-med lr-subtytl">Intel 10th Generation CPUs</h3>
							<p>The biggest factor in how fast your trading software will run is your CPU.</p>
							<p>In this Charter PC we use the fast 10th generation Intel Comet Lake processors, our benchmark testing results show that they are a great choice when looking to run one or two trading platforms whilst keeping costs down.</p>
							<p>The i3 10105F CPU is fast at running most charting platforms, the i5, i7, and i9 options below increase the speed and the multi-tasking ability further allowing extra charts or running a second broker or charting system at the same time. </p>
                            <p>8GB of fast DDR4 RAM is ample for most charting applications however you have the option of increasing this if you anticipate running more platforms, charts, or lots of web browser tabs simultaneously.</p>
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
							<p>With a Charter PC this is made super simple, simply select how many monitors you want to connect in the options below and we will build the PC for you with the right number of monitor ports on the back of it.</p>
							<p>By default you get a triple screen capable PC, 4, 6, or even 8 screen versions can be selected below. You can run standard FHD screens, or even go to higher resolution QHD or 4K monitors.</p>
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
							<p>Every new Charter PC comes with our exclusive version of DisplayFusion, a suite of tools that integrate with Windows and allow full control over all of your programs, charts and screens.</p>
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
							<h3 class="color-med lr-subtytl">On All Charter PC's</h3>
							<p>Noise levels are often overlooked and can be hard to estimate before using a computer, however controlling noise is important. Who wants to sit next to a loud computer through long trading sessions?</p>
							<p>As a company policy, all of our computers are really quiet and the Charter PC is no different.</p>
							<p>Ultra-low noise power supplies, silent graphics cards, silent solid state hard drives, along with manual fan tuning ensure that all you'll hear from your Charter PC is a faint hum at the worst.</p>
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
							<p>With a Charter PC we include pretty much anything you're going to need in with it as standard.</p>
							<p>This includes a wired mouse and keyboard set, which can be removed or switched to wireless. We can also include adapters to allow you to connect HDMI or DVI screens if you request them after ordereding your PC.</p>
                            <p class="detail-callout">Other companies charge between &pound;60 - &pound;100 for this extra kit, we supply it free of charge.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-even wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/windows-11.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Select Windows 10 or 11</h3>
							<h3 class="color-med lr-subtytl">Fully Optimised OS</h3>
							<p>Great computer performance and responsiveness is down to more than just the speed of the components, the software setup and configuration can massively impact how fast a computer runs.</p>
							<p>Windows 10/11 is the main software on your computer, it launches all your trading packages and interfaces, and constantly runs in the background. </p>
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
							<p>All Charter PC's are custom built for each order, this allows you to select the exact configuration that meets your trading needs rather than just buying what's available at the time.</p>
							<p>One of our experienced, expert technicians will then build up the computer for you, pre-install and optimise your Windows and Display Fusion installations, and then put it on a thorough 32-hour stress test to make sure it reaches you in perfect working order.</p>
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
							<p>All Charter PC's come with our standard 5 year hardware cover, the first year is our unique OnSite / Replacement / Collect service, which means most hardware faults can be resolved without you sending the PC back for repair.</p>
							<p>We also offer lifetime email, phone and remote access support for those times when something goes wrong with your Windows installation (which can happen to anyone, at any time).</p>
						</div>
					</div> <!-- lr-detailRow -->
                    <a name="tech"></a><a name="custom-order"></a>	
				</div>		
			</div>		
		</div>
	</section>
	<!-- /Section:  -->

<input type="hidden" id="productid" name="idproduct1" value="339">
<input type="hidden" id="productqty" name="QtyM339" value="1">
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
							<p class="bigtxt-para text-justify">We have a selection of upgrades available to add functionality and further improve the performance of your Charter PC:</p>
						</div>
					</div>
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">CPU / Processor</span></h5>
							<p class="text-justify">The heart of your computer, your processor has a direct impact on the speed of your computer and its ability to run multiple programs simultaneously.</p>
							<div class="get-traderOptions">
								<label class="color">CPU / Processor:</label>
								<select id="idOption1" name="idOption1" class="spec-dd" onchange="reCalc();flashCPU();">
									<option value="title" class="spec-dd-dis" disabled="">Intel 10th Generation CPUs:</option>
									<option value="18190" id="0" title="0">Intel i3 10105F // 3.7 - 4.4GHz // 4C - 8T</option>
									<option value="18191" id="1" title="105">Intel i5 10600KF // 4.1 - 4.8GHz // 6C - 12T + &pound;105.00</option>
                                    <option value="18193" id="2" title="215">Intel i7 10700KF // 3.8 - 5.1GHz // 8C - 16T + &pound;215.00</option>
                                    <option value="18195" id="3" title="345">Intel i9 10900KF // 3.7 - 5.3GHz // 10C - 20T + &pound;345.00</option>
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
									<option value="18199" id="0" title="0">8GB DDR4 2,666MHz</option>
									<option value="18196" id="1" title="45">16GB DDR4 2,666MHz + &pound;45.00</option>
                                    <option value="18197" id="2" title="95">32GB DDR4 2,666MHz + &pound;95.00</option>
									<option value="18198" id="2" title="195">64GB DDR4 2,666MHz + &pound;195.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Number Of Screens Supported</span></h5>
							<p class="text-justify">A standard Charter PC can power up to three 5K monitors. Change the option to support more screens, the Monitor &amp; Resolution panel shows supported resolutions and ports.</p>
							<div class="get-traderOptions">
								<label class="color">Monitor Connections:</label>
								<select id="idOption4" name="idOption4" class="spec-dd" onchange="reCalc();flashGPU();">
									<option value="18316" id="0" title="0">Up to 3 screens - nVidia T400 (2GB)</option>
									<option value="18318" id="1" title="85">Up to 4 screens - nVidia T600 (4GB) + &pound;85.00</option>
                                    <option value="18317" id="2" title="125">Up to 6 screens - nVidia T400 (2GB) x2 + &pound;125.00</option>
									<option value="18319" id="3" title="275">Up to 8 screens - nVidia T600 (4GB) x2 + &pound;275.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Hard Drive Capacity</span></h5>
							<p class="text-justify">Your hard drive is where your software and data is stored. 250GB is enough for Windows and your trading platform installations. Increase for extra file storage capacity.</p>
							<div class="get-traderOptions">
								<label class="color">Hard Drive:</label>
								<select id="idOption3" name="idOption3" class="spec-dd" onchange="reCalc();flashSSD();">
									<option value="18322" id="0" title="0">250GB Adata NVMe M.2 SSD – (3500MBs/1200MBs)</option>
									<option value="18324" id="1" title="45">500GB Adata NVMe M.2 SSD – (3500MBs/2300MBs) + &pound;45.00</option>
                                    <option value="18321" id="2" title="95">1TB Adata NVMe M.2 SSD – (3500MBs/3000MBs) + &pound;95.00</option>
									<option value="18323" id="3" title="195">2TB Adata NVMe M.2 SSD – (3500MBs/3000MBs) + &pound;195.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->

					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Second Hard Drive</span></h5>
							<p class="text-justify">Add a second hard drive if you have larger file storage requirements. </p>
							<div class="get-traderOptions">
								<label class="color">Second Hard Drive:</label>
								<select id="idOption14" name="idOption14" class="spec-dd" onchange="reCalc();flashHDD2();">
									<option value="18209" id="0" title="0">Not Required</option>
									<option value="title" class="spec-dd-dis" disabled="">Traditional Hard Drives:</option>
									<option value="18210" id="1" title="55">1TB Traditional Hard Drive + &pound;55.00</option>
									<option value="18211" id="2" title="75">2TB Traditional Hard Drive + &pound;75.00</option>
									<option value="18212" id="3" title="85">3TB Traditional Hard Drive + &pound;85.00</option>
									<option value="18217" id="4" title="115">4TB Traditional Hard Drive + &pound;115.00</option>
									<option value="18213" id="5" title="170">6TB Traditional Hard Drive + &pound;170.00</option>
									<option value="title" class="spec-dd-dis" disabled="">Fast &amp; Silent SSDs:</option>
									<option value="18214" id="6" title="65">500GB WD Blue SSD (500MBs/500MBs) + &pound;65.00</option>
									<option value="18215" id="7" title="95">1TB WD Blue SSD (500MBs/500MBs) + &pound;95.00</option>
									<option value="18216" id="8" title="195">2TB WD Blue SSD (500MBs/500MBs) + &pound;195.00</option>
									<option value="18320" id="9" title="345">4TB WD Blue SSD (500MBs/500MBs) + &pound;345.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Bootable Backup Drive</span></h5>
							<p class="text-justify">This is a backup solution which clones your 'C' drive on a schedule to a spare internal hard drive and allows for easy and quick recovery from many system issues. </p>
							<div class="get-traderOptions">
								<label class="color">Backup Drive:</label>
								<select id="idOption11" name="idOption11" class="spec-dd" onchange="reCalc();flashBBD();">
									<option value="18035" id="0" title="0">Not Required</option>
									<option value="18034" id="1" title="115">Bootable Backup Drive Solution + &pound;115.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->

					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Mouse &amp; Keyboard Set </span></h5>
							<p class="text-justify">We supply a wired mouse and keyboard set with your Charter PC, you can switch to a wireless set or remove them if you'd prefer.</p>
							<div class="get-traderOptions">
								<label class="color">Inputs:</label>
								<select id="idOption9" name="idOption9" class="spec-dd" onchange="reCalc();flashKYB();">
									<option value="18122" id="0" title="-10">Not Required - &pound;10.00</option>
                                    <option value="18033" id="1" title="0" selected>Wired Mouse / Keyboard Set</option>
                                    <option value="18123" id="2" title="15">Wireless Mouse / Keyboard Set + &pound;15.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Speakers</span></h5>
							<p class="text-justify">Desktop computers do not generally have built in speakers, select whether you would like some suppling with your Charter PC.</p>
							<div class="get-traderOptions">
								<label class="color">Speakers:</label>
								<select id="idOption6" name="idOption6" class="spec-dd" onchange="reCalc();flashSPK();">
                                	<% 
									if request.querystring("sid") <> "" Then
									%>
									<option value="18119" id="1" title="0">Free Desktop Speakers (Worth &pound;20)</option>
                                    <%
									else
									%>
                                    <option value="18120" id="0" title="0">Not Required</option>
                                    <option value="18006" id="1" title="20">Desktop Speaker Set + &pound;20.00</option>
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
							<p class="text-justify">All computers come with a wired network port, if you need a WiFi connection selection your option here. Select the AC card for faster fibre optic connections.</p>
							<div class="get-traderOptions">
								<label class="color">WiFi Card:</label>
								<select id="idOption8" name="idOption8" class="spec-dd" onchange="reCalc();flashWIFI();">
                                	<% 
									if request.querystring("sid") <> "" Then
									%>
                                    <option value="18128" id="1" title="0">Free Wireless N - 300Mbps (Worth &pound;20.00)</option>
                                    <option value="18127" id="2" title="20">Wireless AC - 867Mbps (Worth &pound;40.00) + &pound;20.00</option>
                                    <%
									else
									%>
									<option value="18124" id="0" title="0">Not Required</option>
                                    <option value="18032" id="1" title="20">Wireless N - 300Mbps + &pound;20.00</option>
                                    <option value="18125" id="2" title="40">Wireless AC - 867Mbps + &pound;40.00</option>
                                    <%
									end if
									%>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->



  					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Bluetooth Functionality</span></h5>
							<p class="text-justify">If you need Bluetooth capability for a wireless headset or keyboard / mouse then add it here. </p>
							<div class="get-traderOptions">
								<label class="color">Bluetooth:</label>
								<select id="idOption16" name="idOption16" class="spec-dd" onchange="reCalc();flashBT();">
									<option value="18218" id="0" title="0">Not Required</option>
									<option value="18219" id="1" title="10">USB Bluetooth Adapter + &pound;10.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Operating System</span></h5>
							<p class="text-justify">Select between Windows 10 or 11, Home and Professional editions.</p>
							<div class="get-traderOptions">
								<label class="color">Operating System:</label>
								<select id="idOption10" name="idOption10" class="spec-dd" onchange="reCalc();flashWIN();">
									<option value="17998" id="1" title="0">Windows 10 Home</option>
									<option value="18277" id="3" title="0">Windows 11 Home</option>
									<option value="18074" id="2" title="45">Windows 10 Professional + &pound;45.00</option>
									<option value="18278" id="4" title="45">Windows 11 Professional + &pound;45.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Microsoft Office</span></h5>
							<p class="text-justify">Microsoft Office gives you Word, Excel, PowerPoint & OneNote, if you require Outlook go for the Business Edition. This is a 1 PC, lifetime license.</p>
							<div class="get-traderOptions">
								<label class="color">Microsoft Office:</label>
								<select id="idOption12" name="idOption12" class="spec-dd" onchange="reCalc();flashMSO();">
									<option value="18011" id="4" title="0">Not Required</option>
									<option value="18061" id="6" title="105">Home & Student 2021 Edition + &pound;105.00</option>
                                    <option value="18060" id="8" title="195">Home & Business 2021 Edition + &pound;195.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row otLast-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Hardware Support Warranty</span></h5>
							<p class="text-justify">All Charter PC's come with 5 year hardware cover as standard. The first year is an Onsite / Replacement / Collect service, for extra peace of mind this can be extended for 2 or 3 years.</p>
							<div class="get-traderOptions">
								<label class="color">Hardware Support:</label>
								<select id="idOption13" name="idOption13" class="spec-dd" onchange="reCalc();flashWAR();">
									<option value="18009" id="7" title="0">5 Year (1 Year OnSite / Replacement / Collect)</option>
									<option value="18008" id="8" title="75">5 Year (2 Year OnSite / Replacement / Collect) + &pound;75.00</option>
                                    <option value="18010" id="9" title="150">5 Year (3 Year OnSite / Replacement / Collect) + &pound;150.00</option>
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
												<table width="100%" cellpadding="0" cellspacing="0" class="star-ratings marginbot-20">
													<tbody>
														<tr>
									 						<td><strong>Speed / Responsiveness: </strong><span class="star-desc">PC raw speed ignoring multi-tasking workloads, a big factor in trading software performance levels.</span></td>
															
														</tr>
                                                        <tr>
                                                        	<td><span id="stars-speed"></span></td>
                                                        </tr>
														<tr>
															<td><strong>Multi-Tasking: </strong><span class="star-desc">Ability to run multiple trading platforms and software simultaneously without impacting system performance.</span></td>
														</tr>
                                                        <tr>
                                                        	<td><span id="stars-multi"></span></td>
                                                        </tr>
                                                        <tr>
															<td><strong>Multi-Threading: </strong><span class="star-desc">Multi-threaded performance. Important for back-testing and some of the more intensive platforms.</span></td>
														</tr>
                                                        <tr>
                                                        	<td><span id="stars-mulThr"></span></td>
                                                        </tr>
														<tr>
															<td><strong>Quietness: </strong><span class="star-desc">Noise levels in standard use. <br />10 stars = faint hum, 1 star = jet engine.</span></td>
														</tr>
                                                        <tr>
                                                        	<td><span id="stars-quiet"></span></td>
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
							<h5 class="color">Full Charter PC Specifications</h5>
							<div class="spec-content">
								<p><span id="txtCPU"></span></p>
								<p><span id="txtRAM"></span></p>
								<p>Fast B560 Chipset Motherboard</p>
								<p><span id="txtGPU"></span></p>
								<p><span id="txtSSD"></span></p>
								<span id="txtHDD2"></span>
								<span id="txtBBD"></span>
								<p>Antec VSK ELite Case</p>
								<p>Low Noise 500W Quiet Power Supply</p>
								<span id="txtDVD"></span>
								<span id="txtWIFI"></span>
								<p>Gigabit Ethernet LAN Adapter</p>
								<p>8 Channel High Definition Audio Sound Card</p>
								<p>3 x USB 3, 3 x USB 2 &amp; 1 x USB Type-C Ports</p>
								<p>Low Noise CPU Cooler</p>
								<p id="pKYB"><span id="txtKYB"></span></p>
								<span id="txtSPK"></span>
								<span id="txtBT"></span>
								<p><span id="txtWIN"></span></p>
								<span id="txtMSO"></span>
								<p><span id="txtWAR"></span></p>
								<p class="uppricefont"><strong class="">Charter PC Price:</strong> <strong class="color">&pound;<span id="pcPrice"></span></strong> <strong class="pri1">+ VAT</strong></p>
                                <span id="txtBunPrice"></span>
								  <p>(&pound;<span id="vatPrice"></span> inc. VAT)</p><br>
								  <div align="center">
                                  <input type="submit" value="ORDER YOUR CHARTER PC NOW" class="btn btn-skin btn-sm text-uppercase" />
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