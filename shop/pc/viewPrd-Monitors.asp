<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact, LLC.
'ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC.
'Copyright 2001-2007. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'Early Impact. To contact Early Impact, please visit www.earlyimpact.com.

'DA Edit - Vimeo Stand URLS
		daVimeoUrl = "/pop-pages/stand-video.asp?s=" & LCase(pSku)
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
								<p class="t-declaration-text">If you order from a US supplier <br />You will be liable  for VAT &amp; Shipping</p>
							</div>
						</div>
					</div>					
				</div>		
			</div>		
		</div>	
    </header>
	
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
							<a href="#tech" class="text-underline">View Monitor Technical Specs</a>
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
								if daFunDelDateBlockTest(0,0) then
									daDelEstimate = "Due to a short workshop closure, orders will be delivered on <strong class=""color"">" & daFunDelDateReturn(0,0) & "</strong>."
								Else
									daDelEstimate = "Order before <strong class=""color"">" & daFunDelCutOff() & "</strong> for delivery on <strong class=""color"">" & daFunDelDateReturn(0,0) & "</strong>"
								end if
								%>
                        		<p class="dcategory"><label class="rubric semi space-right10">UK :</label> <strong class="color">&pound;10</strong> | <%=daDelEstimate%></p>
                        		<p class="dcategory"><label class="rubric semi space-right10">International :</label> International shipping from just &pound;20 - <a data-toggle="lightbox" data-title="International Delivery" class="text-underline" href="/pop-pages/int-del-pop.asp?ProdID=<%=pidProduct%>">View Costs / Timescales</a></p>
								<p class="dcategory"><em>(UK buyers: Saturday delivery is available in the checkout. You can also email us after placing an order to request a specific delivery date, any date after the above estimate is possible.)</em></p>
							</div>
                            <%
							'Swap button target depending on whether we are on a single product, building an array or a bundle
							if request.querystring("sid") = "" Then
							'We are just on a plain product
							%>
							<a class="btn btn-skin btn-wc semi order-btn margintop-30" href="/shop/pc/instPrd.asp?idproduct=<%=pIdProduct%>">Add Monitor To Your Basket <i class="fa fa-angle-right"></i></a>
                            <%
							else
								if request.querystring("arr") = 1 Then
								'We are building an array
								%>
							<a class="btn btn-skin btn-wc semi order-btn margintop-30" href="/display-systems-3/?sid=<%=request.querystring("sid")%>&mid=<%=pIdProduct%>">Add Monitor To Your Array <i class="fa fa-angle-right"></i></a>
                                <%
								else
								'We must be on a bundle
								%>
							<a class="btn btn-skin btn-wc semi order-btn margintop-30" href="/bundles-3/?sid=<%=request.querystring("sid")%>&mid=<%=pIdProduct%>&cid=<%=request.querystring("cid")%>">Add Monitor To Your Bundle <i class="fa fa-angle-right"></i></a>
                                <%
								end if
							end if
							%>
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
									<h2 class="h-bold font-light disp-inline">Free Cables,</h2>
									<h3 class="h-light font-light disp-inline">&nbsp;Buy with a Stand &amp get free long length cables</h3>
									</div>
									</div>
								</div>
								<div class="col-md-2">
									<div class="wow fadeInRight" data-wow-delay="0.1s">
										<div class="cta-btn">
										<a data-toggle="lightbox" data-title="Multi-Screen Arrays" href="/pop-pages/arrays.htm" class="btn btn-outline">Learn More</a><a name="tech"></a>
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
				<div class="col-sm-12 specs-wrap wow fadeInUp" data-wow-delay="0">
					<div id="mt-specifications" class="bg-white">
                    	<h5 class="color h-semi margintop-0 marginbot-30">Monitor Technical Specifications</h5>
						<div class="row">
						<% pcs_LongProductDescription %>
						</div>
					</div>
				</div>			
			</div>		
		</div>
	</section>
	<!-- /Section: custom-order -->
