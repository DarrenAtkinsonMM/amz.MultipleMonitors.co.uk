<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact, LLC.
'ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC.
'Copyright 2001-2007. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>


	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
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
							<a href="javascript:specjump();void(0);" class="text-underline">Full Computer Spec &amp; Customisation Options</a>
							<a href="javascript:faqjump();void(0);" class="text-underline">Frequently Asked Questions</a>
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
                        <%
                        
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        ' START:  Show product prices
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
                        %>
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
                            <a class="btn btn-skin btn-wc semi order-btn margintop-30" href="javascript:specjump();void(0);">Customise &amp; Order Your New PC <i class="fa fa-angle-right"></i></a>
						</div>
					</div>
				</div>
			</div>
		</div>

	</section>
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
													<h5 class="font-light disp-inline cta-txtline">About This Multi-Screen PC?</h5>
												</div>
											</div>
											<div class="col-lg-4 col-sm-5 cta-question cta-2line cta-gqcolumn">
												<i class="fa fa-question cta-icon"></i><a href="javascript:faqjump();void(0);" class="twoline-link linkpre-question">View Frequently <br/>Asked Questions</a>
											</div>
										</div>
									</div>
								</div>
								<div class="col-lg-6"><a name="customise"></a><a name="faqs"></a>
									<div class="wow fadeInRight" data-wow-delay="0.1s">
										<div class="row">
											<div class="col-lg-5 col-sm-6 cta-email cta-2line cta-gqcolumn">
												<i class="fa fa-envelope-o cta-icon"></i><a href="javascript:;" class="twoline-link linkpre-mail">Send us an <strong>Email enquiry</strong></a>
											</div>
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
    <section id="custom-order" class="product-specs bg-smog">
		<div class="container">
			<div class="row">
				<div class="col-sm-12 specs-wrap wow fadeInUp" data-wow-delay="0.1s">
					<ul id="productSpecs" class="nav nav-tabs">
					   <li class="active pscustom-wide"><a href="#fullSpecs" data-toggle="tab">Full Specification &amp; Customisation Options</a></li>
					   <li class="pscustom-wide2"><a href="#faq" data-toggle="tab">Frequently Asked Questions</a></li>	
					</ul>
					<div id="specsTabContent" class="tab-content">
					    <div class="tab-pane fade in active" id="fullSpecs">
							<div class="row">
								<div id="specForm-wrap" class="col-lg-8 col-sm-12">
                                									
									<div class="specs-title wow fadeInUp" data-wow-delay="0.1s">
										<h5 class="color h-semi margintop-0"><%=Replace(pDescription, "Multi Screen ", "")%> | Specification &amp; Customisation Options</h5>
										<p class="medium subhead-para">
                        <%
                        if not InStr(pSku, "PRO1") = 0 Then
							Response.Write "The Pro PC has a good base spec for solid Windows performance, you can use the options below to enhance performance and change the number of monitors supported."
						end if 
						if not InStr(pSku, "ULT1") = 0 Then
							Response.Write "The Ultra PC has a strong base spec for fast Windows performance, you can use the options below to enhance performance further and change the number of monitors supported."
						end if 
						if not InStr(pSku, "EXT1") = 0 Then
							Response.Write "The Extreme PC has a fantastic spec for superb Windows performance, use the options below to enhance performance further and change the number of monitors supported."
						end if 
						%></p>
									</div>
  <form method="post" action="/shop/pc/instPrd.asp" name="additem" onSubmit="return checkproqty(document.additem.quantity);" class="pcFormsProdSmall" id="pcform">
		<input name="index" type="hidden" value="1">
        <input type="hidden" name="idproduct" value="<%=pidProduct%>">
        <input type="hidden" name="quantity" value="1">
<% pcs_OptionsN %> 
								<div class="col-lg-4 col-sm-12 spec-box-wrap">
									
									<div id="cust-monitors" class="spec-box spec-box1 wow fadeInRight" data-wow-delay="0.1s"><!-- spec-box2 -->
										<div class="spec-custom">
											<h5 class="specbox-heading">System Performance Levels</h5>
											<div class="spec-content">
												<table width="100%" cellpadding="0" cellspacing="0" class="star-ratings marginbot-10">
													<tbody>
														<tr>
									 						<td><strong>CPU Speed: </strong><span class="star-desc">The raw speed of the CPU, important for fast &amp; responsive performance.</span></td>
															
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0 paddingbot-10"><span id="stars-speed"></span></td>
                                                        </tr>
														<tr>
															<td><strong>Multi-Tasking: </strong><span class="star-desc">CPU's ability to handle lots of programs open simultaneously.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0 paddingbot-10"><span id="stars-multi"></span></td>
                                                        </tr>
                                                        <tr>
															<td><strong>Multi-Threading: </strong><span class="star-desc">CPU performance in more intensive multi-threaded workloads.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0 paddingbot-10"><span id="stars-mulThr"></span></td>
                                                        </tr>
									 					<tr>
															<td><strong>Graphics Power: </strong><span class="star-desc">Capability of handling more graphical programs and apps.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0 paddingbot-10"><span id="stars-gpu"><img src="/images/generic/stars5.jpg"></span></td>
                                                        </tr>
									 					<tr>
															<td class="paddingbot-10"><strong>AI Performance: </strong><span class="star-desc">The TOPS score for this graphics card (higher is better) is:&nbsp;&nbsp;<span class="spec-tops"><span id="stars-gputops">43</span></span></td>
														</tr>
														<tr>
															<td class="paddingtop-0"><strong>System Noise Levels: </strong><span class="star-desc">(In standard use) <br />10 stars = faint hum, 1 star = jet engine.</span></td>
														</tr>
                                                        <tr>
                                                        	<td class="paddingtop-0"><span id="stars-quiet"><img src="/images/generic/stars10.jpg"></span></td>
                                                        </tr>
													</tbody>
												</table>
												<a data-toggle="lightbox" data-title="Computer Ratings Explained" class="specbox-link" href="/pop-pages/custpc-stars.htm">Learn More About These Ratings</a>
											</div>
										</div>
									</div><!-- spec-box1 end -->
									<div id="cust-monitors" class="spec-box spec-box2 wow fadeInRight" data-wow-delay="0.1s"><!-- spec-box1 -->
										<div class="spec-custom">
											<h5 class="specbox-heading"><span id="optScreensTitle">Monitors & Resolutions</span></h5>
											<div class="spec-content">
									 			<p><strong>Supported Screen Resolutions:</strong></p>
												<ul class="specbox-list">
													<span id="optRes"></span>
												</ul>
									 			<p><strong>Available Monitor Ports:</strong></p>
												<ul class="specbox-list">
													<span id="optPorts"></span>
												</ul>
												<small class="spec-infotxt">*Use the 'Graphics Card Setup' to change this.</small>
											</div>
											
										</div>
									</div><!-- spec-box2 end -->
									<div id="cust-monitors" class="spec-box spec-box3 wow fadeInRight" data-wow-delay="0.1s"><!-- spec-box3 -->
										<div class="spec-custom">
											<h5 class="specbox-heading">Your Final Price</h5>
											<div class="spec-content">
												<table width="100%" cellpadding="0" cellspacing="0" class="upgrades">
												   <tbody>
														<tr>
															<td>PC Base Price:</td>
															<td align="right">&pound;<span id="basePrice"></span></td>
														</tr>
														<tr>
															<td>Customisation Options:</td>
															<td align="right">&pound;<span id="extrasPrice"></span></td>
														</tr>
														<tr>
															<td><span class="upgpri">Sub Total:</span></td>
															<td align="right"><span class="upgpri">&pound;<span id="subtotalPrice"></span></span></td>
														</tr>
														<tr>
															<td>VAT:</td>
															<td align="right">&pound;<span id="vatPrice"></span></td>
														</tr>
														<tr>
															<td><span class="upgpri">Final Price:</span></td>
															<td align="right"><span class="upgpri">&pound;<span id="finalPrice"></span></span></td>
														</tr>
														<tr>
															<td colspan="2" align="center">
																<input type="submit" value="ORDER YOUR NEW PC NOW" class="btn btn-skin btn-sm text-uppercase" />
															</td>
														</tr>
												    </tbody>
												</table>
											</div>
											
										</div>
									</div><!-- spec-box3 end -->
								</div>
							</div>
							<div class="row spec-additions">
								<div class="col-lg-3 col-sm-6 prd-additions prd-add1 wow fadeInLeft" data-wow-delay="0.1s">
									<h2><strong>Low Noise</strong> <span>Computing</span></h2>
									<p>All our PCs are built to run really, really quiet.</p>
									<p>We choose components to ensure that <span class="rubric">your new PC runs almost silent.</span></p>
									<p>Built in and provided at no extra cost to you.</p>
								</div>
								<div class="col-lg-3 col-sm-6 prd-additions prd-add2 wow fadeInLeft" data-wow-delay="0.2s">
									<h2><strong>free</strong> <span>software</span></h2>
									<p>Arrange your programs quickly and easily across your screens with Display Fusion, <span class="rubric">our exclusive version of the best multi-screen productivity tool</span>.</p>
									<p>Pre-installed and configured for you, no trials, lifetime license.</p>
								</div>
								<div class="col-lg-3 col-sm-6 prd-additions prd-add3 wow fadeInLeft" data-wow-delay="0.4s">
									<h2><strong>warranty</strong> <span>cover</span></h2>
									<p>Lifetime phone, email & remote access PC support.</p>
									<p>5 Year Hardware Cover.</p>
									<p>Hardware problems normally fixed within 1 business day <span class="rubric">- No long wait times for repairs</span>.</p>
								</div>
								<div class="col-lg-3 col-sm-6 prd-additions prd-add4 wow fadeInLeft" data-wow-delay="0.6s">
									<h2><strong>quick build</strong> <span>&amp; delivery</span></h2>
									<p>Order now and <span class="rubric">receive your computer on <%=daFunDelDateReturn(1,0) %>*</span>.</p>
									<p>We do NOT charge extra for this quick build time like others do.</p>
									<p>*UK mainland only.</p>
								</div>
							</div>
					    </div>
					    <div class="tab-pane fade" id="faq">
<% pcs_LongProductDescription %>
					</div>
				</div>			
			</div>		
		</div>
	</section>
	<!-- /Section: custom-order -->
</form>