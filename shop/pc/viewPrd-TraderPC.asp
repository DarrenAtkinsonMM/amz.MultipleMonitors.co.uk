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
							<p><a href="#learn" class="text-underline">Learn More About The Trader PC</a></p>
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
								' daDelEstimate = "Due to the Christmas and New Year holidays, deliveries will now be made after <strong class=""color"">" & daFunDelDateReturn(1,0) & "</strong>."
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
							<p class="bigtxt-para text-justify"><i class="fa fa-line-chart lr-titleicon color"></i>The Trader PC is a fantastic choice for traders who are looking for the absolute best in class performance levels for their trading platforms. Read on to see how the Trader PC can handle anything you throw at it:</p>
						</div>
					</div>
					<div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline fstrow-img">
								<img src="/images/trader-pc/intel-core.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6 paddingtop-0">
							<h3 class="h-bold color-med">Fast Processors & RAM</h3>
							<h3 class="color-med lr-subtytl">Choose Responsive Intel Chips</h3>
							<p>When it comes to running trading platforms our benchmark tests show that the biggest impact in how responsive they are is the processor.</p>
                            <p>Having an ultra-fast CPU in your trading computer ensures the absolute best performance possible.</p>
							<p>With the Trader PC you get the Intel 14th generation chips, these run trading software faster than all older Intel and AMD chips.</p>
                            <p>For power users going for processors with more CPU cores will offer you fantastic multi-tasking and multi-threaded performance. </p>
                            <p>Combining a fast processor with the 16GB or 32GB of DDR4 RAM makes for an great combination for pretty much any trading workload.</p>
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
							<h3 class="color-med lr-subtytl">Multi-Screen Support</h3>
							<p>For traders seeing the right information at the right time can be the difference between success and failure. With a multi-screen trading computer you can position your charts and programs at any point on any screen.</p>
							<p>We take all the hassle out of achieving a multi-screen system, simply select how many screens you want to run using the options below and we do the rest. </p>
							<p>Your Trader PC will be delivered pre-configured to connect to your selected number of screens right out of the box, connect standard resolution FHD screens, or go for higher resolution QHD, 4K, or even 5K monitors.</p>
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
							<p>Something that can be a problem for anyone running a multi-screen computer system, especially traders, is getting the right info in the right screens quickly.</p>
							<p>Our exclusive version of DisplayFusion, solves this and many other multi-screen problems with ease.</p>
                            <p>You can control exactly where programs and charts open and automatically jump windows to pretty much anywhere you want. This ability, combined with extended taskbars, and screen partitions will make a massive difference to your experience of using a multi-screen trading PC.</p>
                            <p>It comes pre-installed and ready to use with your new Trader PC.</p>
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
							<h3 class="color-med lr-subtytl">On All Trader PC's</h3>
							<p>Traders can often find themselves sat in front of a PC for long periods of time and a noisy PC setup can quickly become both annoying and frustrating, it's one of the things many traders ask us about.</p>
                            <p>Some try to reduce noise by using noise insulation materials however this often increases internal system temperatures which is not a good idea.</p>
							<p>We eliminate noise by building all our computers using ultra-low noise components, quiet cooling fans, passively cooled graphics cards and silent SSD hard drives, this guarantees you will hardly hear a thing from your new Trader PC.</p>
                            <p class="detail-callout">Many companies charge up to &pound;50 extra for a quiet PC build, we include this at no cost to you.</p>
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
							<h3 class="color-med lr-subtytl">You Need Available</h3>
							<p>Every 'Trader PC' comes with options for all the kit you need to get up and running instantly, straight out of the box.</p>
							<p>You can select a Logitech Wired or Wireless mouse and keyboard set, some traders prefer a wired set to avoid any potential battery or connection issues.</p>
							<p>We can also provide graphics adapters to let you easily connect your machine up to your choice of monitors so you don't have to worry about screen compatibility.</p>
                            <p>Everything can be supplied with your Trader PC in one simple purchase.</p>
                            <p class="detail-callout">Most companies charge between &pound;80 - &pound;120 extra for this equipment, we supply it all at cost pricing.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-even wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/windows-11.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Choose Windows 11</h3>
							<h3 class="color-med lr-subtytl">Tuned for Performance</h3>
							<p>You could have the fastest processor available, lots of RAM, and an ultra-fast hard drive, but if the software which runs on your PC is mis-configured or runs slowly then it's all completely pointless.</p>
							<p>The main piece of software on any computer is the operating system, Windows 11 for the Trader PC, ensuring that it runs quickly can make a real impact on how your computer feels in day to day use.</p>
							<p>We take steps to tune Windows performance for each Trader PC to so that it runs as fast as it can, taking full advantage of your computer's hardware. </p>
                            <p class="detail-callout">Some companies charge up to &pound;40 extra for Windows optimisation, we do this for you free of charge.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/build-delivery.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">Fast Build &amp; Test</h3>
							<h3 class="color-med lr-subtytl">Begin Trading Faster</h3>
							<p>Once you have configured your new Trader PC to meet your specific needs, the order is passed across to our workshop where your new machine will be custom built just for you.</p>
							<p>The build process includes physically assembling your computer and then installing and configuring Windows and Display Fusion so that it works right out of the box. The final step is a 32 hour stress test to make sure everything is in full working order.</p>
							<p>This custom computer build and test routine takes 4 - 5 working days, with delivery made on the next working day.</p>
                            <p class="detail-callout">Many companies will charge from &pound;70 right up to &pound;300 for a 4 - 5 day build, we do this as standard at no extra cost.</p>
						</div>
					</div> <!-- lr-detailRow -->
					<div class="row lr-detailRow lr-even wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/onsite-support.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">5 Year Hardware Cover</h3>
							<h3 class="color-med lr-subtytl reducefont-sub">&amp; Unlimited Remote Support As Standard</h3>
							<p>Computers are great, until they stop working properly, and with the best will in the world nobody can guarantee that any PC will keep working indefinitely. If you rely on your trading computer to make your money then it's wise to have great support in place.</p>
							<p>Our team of technicians will support you and your machine remotely via email, telephone and remote access sessions for the lifetime of your PC.</p>
							<p>If the worst happens and a part fails in your computer then our unique OnSite / Replacement / Collect service means that we can get you back up and running without the need for you to send the PC back to us in most cases.</p>
						</div>
					</div>
                    <div class="row lr-detailRow lr-odd wow fadeInUp" data-wow-delay="0">
						<div class="lr-mage-col text-center col-md-6 pright-md">
							<div class="lr-mage displ-inline">
								<img src="/images/trader-pc/trading-customers.jpg" alt="" />
							</div>
						</div>
						<div class="lr-details-col col-md-6">
							<h3 class="h-bold color-med">You're In Great Company</h3>
							<h3 class="color-med lr-subtytl">Traders Trust Us</h3>
							<p>We have worked with a lot of traders over the past 13+ years and have supplied equipment to customers of all sizes.</p>
                            <p>Here are a small sample of our trading customers who actively trade on Multiple Monitors trading computers every single day.</p>
                            <p>We are capable of supporting the individual trader working from a home office, right through to hedge funds with multi-millions under management with an in-house team of traders.</p>
                            <p class="detail-callout">We have customers running everything from MT4, IG Index, Pro-Realtime, CMC Markets, NinjaTrader, TradeStation, Bloomberg and anything in-between, if you want to run it we have probably supported it at some point in time.</p>
						</div>
					</div> <!-- lr-detailRow -->
 <!-- lr-detailRow --><a name="tech"></a><a name="custom-order"></a>	
				</div>		
			</div>		
		</div>
	</section>
	<!-- /Section:  -->

<input type="hidden" id="productid" name="idproduct1" value="333">
<input type="hidden" id="productqty" name="QtyM333" value="1">
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
							<p class="bigtxt-para text-justify">We have a selection of upgrades available to add functionality and further improve the performance of your Trader PC:</p>
						</div>
					</div>
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">CPU / Processor</span></h5>
							<p class="text-justify">The number one difference to how fast your computer will perform, CPU's impact speed and multi-tasking performance levels.</p>
							<div class="get-traderOptions">
								<label class="color">CPU / Processor:</label>
								<select id="idOption1" name="idOption1" class="spec-dd" onchange="reCalc();flashCPU();">
									<option value="title" class="spec-dd-dis" disabled="">Intel 14th Generation CPUs:</option>
									<option value="18464" id="3" title="0">Intel i5 14400F // 2.5 - 4.7GHz // 10C - 16T </option>
									<option value="18479" id="3" title="65">Intel i5 14600KF // 3.5 - 5.3GHz // 14C - 20T + &pound;65.00</option>
                                    <option value="18480" id="4" title="145">Intel i7 14700KF // 3.4 - 5.6GHz // 20C - 28T + &pound;145.00</option>
									<option value="18496" id="5" title="265">Intel i9 14900KF // 3.2 - 6.0GHz // 24C - 32T + &pound;265.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
                    <div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">RAM / Memory</span></h5>
							<p class="text-justify">Your memory or RAM dictates how many programs and charts your trading computer can hold open without slowing down your PC.</p>
							<div class="get-traderOptions">
								<label class="color">RAM / Memory:</label>
								<select id="idOption2" name="idOption2" class="spec-dd" onchange="reCalc();flashRAM();">
									<option value="17950" id="0" selected title="0">16GB DDR4 3,200MHz</option>							
									<option value="17951" id="1" title="125">32GB DDR4 3,200MHz + &pound;125.00</option>
                                    <option value="18064" id="2" title="325">64GB DDR4 3,200MHz + &pound;325.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Number Of Screens Supported</span></h5>
							<p class="text-justify">Your new Trader PC can power up to four 4K screens, to run more screens simply change the option here. The Monitor &amp; Resolution panel shows supported resolutions and ports.</p>
							<div class="get-traderOptions">
								<label class="color">Monitor Connections:</label>
								<select id="idOption4" name="idOption4" class="spec-dd" onchange="reCalc();flashGPU();">
									<option value="title" class="spec-dd-dis" disabled>Up To 4 Monitor Capable:</option>
									<option value="18518" id="1" title="0">Up to 4 screens - nVidia RTX A400 (4GB)</option>
									<option value="18510" id="5" title="125">Up to 4 screens - nVidia RTX 5050 (8GB) + &pound;125</option>
                                    <option value="title" class="spec-dd-dis" disabled>Up To 8 Monitor Capable:</option>
									<option value="18519" id="3" title="165">Up to 8 screens - nVidia RTX A400 (4GB) x2 + &pound;165</option>
									<option value="18511" id="6" title="395">Up to 8 screens - nVidia RTX 5050 (8GB) x2 + &pound;395</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Hard Drive Capacity</span></h5>
							<p class="text-justify">Used to store your installed programs and files, 500GB is a decent amount for most. Increase if you want more room to store data files and folders.</p>
							<div class="get-traderOptions">
								<label class="color">Hard Drive:</label>
								<select id="idOption3" name="idOption3" class="spec-dd" onchange="reCalc();flashSSD();">
									<option value="18393" id="1" selected title="0">500GB Kingston NVMe M.2 SSD – (3500MBs/2300MBs)</option>
                                    <option value="18390" id="2" title="65">1TB Kingston NVMe M.2 SSD – (6000MBs/4000MBs) + &pound;65.00</option>
                                   <option value="18392" id="3" title="125">2TB WD NVMe M.2 SSD – (6000MBs/5000MBs) + &pound;125.00</option>
									<option value="18465" id="4" title="375">4TB Kingston NVMe M.2 SSD –  (3500MBs/2800MBs) + &pound;375.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Second Hard Drive</span></h5>
							<p class="text-justify">Add in a second hard drive if you have larger file storage needs. </p>
							<div class="get-traderOptions">
								<label class="color">Second Hard Drive:</label>
								<select id="idOption14" name="idOption14" class="spec-dd" onchange="reCalc();flashHDD2();">
									<option value="18229" id="0" title="0">Not Required</option>
									<option value="title" class="spec-dd-dis" disabled="">Fast &amp; Silent SSDs:</option>
									<option value="18235" id="7" title="125">1TB Kingston M.2 NVMe SSD (3500MBs/3000MBs) + &pound;125.00</option>
									<option value="18236" id="8" title="185">2TB WD M.2 NVMe SSD (3500MBs/3000MBs) + &pound;185.00</option>
									<option value="18315" id="9" title="445">4TB Kingston M.2 NVMe SSD (3500MBs/2800MBs) + &pound;445.00</option>
									<option value="title" class="spec-dd-dis" disabled="">Traditional Hard Drives:</option>
									<option value="18237" id="4" title="105">4TB Traditional Hard Drive + &pound;105.00</option>
									<option value="18233" id="5" title="140">6TB Traditional Hard Drive + &pound;140.00</option>
									<option value="18484" id="6" title="160">8TB Traditional Hard Drive + &pound;160.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Bootable Backup Drive</span></h5>
							<p class="text-justify">For quick and easy recovery from a number of Windows issues you can choose to add a separate internal backup drive.</p>
							<div class="get-traderOptions">
								<label class="color">Backup Drive:</label>
								<select id="idOption11" name="idOption11" class="spec-dd" onchange="reCalc();flashBBD();">
									<option value="18040" id="0" title="0">Not Required</option>
									<option value="18039" id="1" title="145">Bootable Backup Drive Solution + &pound;145.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Optical / DVD Drive</span></h5>
							<p class="text-justify">Add a DVD ReWriter drive if you need one. </p>
							<div class="get-traderOptions">
								<label class="color">DVD Drive:</label>
								<select id="idOption15" name="idOption15" class="spec-dd" onchange="reCalc();flashDVD();">
									<option value="18245" id="0" title="0">Not Required</option>
									<option value="17915" id="1" title="60">DVD ReWriter Drive + &pound;60.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
                    <div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Mouse &amp; Keyboard Set </span></h5>
							<p class="text-justify">Choose between a high quality wireless or wired mouse and keyboard set with your Trader PC, or you can supply your own, any PC compatible set should work fine.</p>
							<div class="get-traderOptions">
								<label class="color">Inputs:</label>
								<select id="idOption9" name="idOption9" class="spec-dd" onchange="reCalc();flashKYB();">
									<option value="18113" id="0" title="0">Not Required</option>
                                    <option value="18114" id="1" title="20">Wired Mouse / Keyboard Set + &pound;20.00</option>
                                    <option value="17894" id="2" title="25">Wireless Mouse / Keyboard Set + &pound;25.00</option>
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
									<option value="18132" id="1" title="0">Free Desktop Speakers (Worth &pound;20)</option>
                                    <%
									else
									%>
                                    <option value="18111" id="0" title="0">Not Required</option>
                                    <option value="17897" id="1" title="20">Desktop Speaker Set + &pound;20.00</option>
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
                                    <option value="18133" id="1" title="0">Free Wireless AX - 3,000Mbps inc. Bluetooth (Worth &pound;40)</option>
                                    <%
									else
									%>
									<option value="18112" id="0" title="0">Not Required</option>
                                    <option value="17966" id="1" title="40">Wireless AX - 3,000Mbps inc. Bluetooth + &pound;40</option>
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
									<option value="18246" id="0" title="0">Not Required</option>
									<option value="18247" id="1" title="10">USB Bluetooth Adapter + &pound;10.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Operating System</span></h5>
							<p class="text-justify">Choose between Windows 11 Home or Professional Edition.</p>
							<div class="get-traderOptions">
								<label class="color">Operating System:</label>
								<select id="idOption10" name="idOption10" class="spec-dd" onchange="reCalc();flashWIN();">
                                    <option value="18443" id="1" title="0">Windows 11 Home</option>
									<option value="18280" id="4" title="45">Windows 11 Professional + &pound;45.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Microsoft Office</span></h5>
							<p class="text-justify">Microsoft Office Home gives you Word, Excel, PowerPoint & OneNote, if you require Outlook go for the Business Edition. This is a 1 PC, lifetime license.</p>
							<div class="get-traderOptions">
								<label class="color">Microsoft Office:</label>
								<select id="idOption12" name="idOption12" class="spec-dd" onchange="reCalc();flashMSO();">
									<option value="17906" id="4" title="0">Not Required</option>
                                    <option value="18063" id="6" title="105">Home Edition 2024 + &pound;105.00</option>
                                    <option value="18062" id="8" title="195">Home & Business Edition 2024 + &pound;195.00</option>
								</select>
							</div>
						</div>
					</div><!-- optiontrade-row -->
					<div class="row wow fadeInUp optiontrade-row otLast-row" data-wow-delay="0">
						<div class="row-traderOptions col-xs-12">
							<h5 class="h-semi color-med">Hardware Support Warranty</span></h5>
							<p class="text-justify">All Trader PC's come with 5 year hardware cover as standard. The first year is an Onsite / Replacement / Collect service, for extra peace of mind this can be extended for 2 or 3 years.</p>
							<div class="get-traderOptions">
								<label class="color">Hardware Support:</label>
								<select id="idOption13" name="idOption13" class="spec-dd" onchange="reCalc();flashWAR();">
									<option value="17992" id="7" title="0">5 Year (1 Year OnSite / Replacement / Collect)</option>
									<option value="17903" id="8" title="75">5 Year (2 Year OnSite / Replacement / Collect) + &pound;75.00</option>
                                    <option value="17905" id="9" title="150">5 Year (3 Year OnSite / Replacement / Collect) + &pound;150.00</option>
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
															<td class="paddingbot-10"><strong>AI Performance: </strong><span class="star-desc">The TOPS score for this graphics card (higher is better) is:&nbsp;&nbsp;<span class="spec-tops"><span id="stars-gputops">43</span></span></td>
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
							<h5 class="color">Full Trader PC Specifications</h5>
							<div class="spec-content">
								<p><span id="txtCPU"></span></p>
								<p><span id="txtRAM"></span></p>
								<p><span id="txtMB"></span></p>
								<p><span id="txtGPU"></span></p>
								<p><span id="txtSSD"></span></p>
								<span id="txtHDD2"></span>
								<span id="txtBBD"></span>
								<p>Fractal Design Case</p>
								<p>BeQuiet Premium 500W Quiet Power Supply</p>
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
								<p class="uppricefont"><strong class="">Trader PC Price:</strong> <strong class="color">&pound;<span id="pcPrice"></span></strong> <strong class="pri1">+ VAT</strong></p>
                                <span id="txtBunPrice"></span>
								  <p>(&pound;<span id="vatPrice"></span> inc. VAT)</p><br>
								  <div align="center">
                                  <input type="submit" value="ORDER YOUR TRADER PC NOW" class="btn btn-skin btn-sm text-uppercase" />
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