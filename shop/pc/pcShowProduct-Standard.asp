<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2006. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="<%=pDescription%>"><a class="" href="/products/<%=pUrl%>/"><%=pDescription%></a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 <%if pSmallImageUrl<>"" then%>
								<img src="/shop/pc/catalog/<%response.write pSmallImageUrl%>" alt="<%=pDescription%>" />
							<% else %>
								<img src="/shop/pc/catalog/no_image.gif" alt="<%=pDescription%>" />
							<%end if %>
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p><%=psDesc%></p>
								<% if (pPrice>0) and (pcv_intHideBTOPrice<>"1") then %>
								<h4>Price: <span><%response.write scCursign & money(pPrice/1.2)%></span></h4>
			<%if (pListPrice-pPrice)>0 AND plistHidden<0 AND session("customerType")<>1 then %>
				<p class="pcShowProductListPrice">
					<%=dictLanguage.Item(Session("language")&"_viewPrd_20")%><%=scCursign & money(pListPrice)%>
				</p>
				<p class="pcShowProductSavings">
					<%=dictLanguage.Item(Session("language")&"_prdD1_2") & scCursign & money(pListPrice-pPrice)%>
				</p>
			<% end if
		end if %>
		<% 'if customer category type logged in - show pricing
		if session("customerCategory")<>0 and (dblpcCC_Price>"0") and (pcv_intHideBTOPrice<>"1") then %>
			<p class="pcShowProductPriceW">
				<% response.write session("customerCategoryDesc")& " " & scCursign & money(dblpcCC_Price)%>
			</p>
		<% else %>
			<% if (dblpcCC_Price>"0") and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") then %>
				<p class="pcShowProductPriceW">
					<% response.write dictLanguage.Item(Session("language")&"_prdD1_4")& " " & scCursign & money(dblpcCC_Price)%>
				</p>
			<% end if 
		end if %>
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/products/<%=pUrl%>/">More info</a>
                                    <% If pcf_AddToCart(pIdProduct)=True Then %>
                                    <a title="Buy Now" class="btn product-action pg-green-btn" href="/shop/pc/instPrd.asp?idproduct=<%=pIdProduct%>">Buy Now</a><% End If %>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->