<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<div class="pcClear"></div>
<div id="opcOrderPreview" data-ng-controller="orderDetailCtrl" data-ng-cloak class="ng-cloak">
	<div id="opcOrderPreviewWrapper">
		<ul class="pcListLayout">
			<li id="pcOrderPreview" class="pcCartLayout container-fluid">
				<div class="pcTableHeader row">
					<div class="col-xs-4 col-sm-5"><%=dictLanguage.Item(Session("language")&"_orderverify_27")%></div>
					<div class="col-xs-1"><%=dictLanguage.Item(Session("language")&"_orderverify_25")%></div>
					<div class="col-xs-1 col-sm-3 hidden-xs right"><%=dictLanguage.Item(Session("language")&"_orderverify_32")%></div>
					<div class="col-xs-5 col-sm-3 right"><%=dictLanguage.Item(Session("language")&"_orderverify_28")%></div>
				</div>
				<div id="pcShoppingCartRows">
					<div class="pcShoppingCartRow" data-ng-repeat="shoppingcartitem in shoppingcart.shoppingcartrow">
						<div class="{{shoppingcartitem.rowClass}} row">
							<% '// START 2nd Row - Main Product Data %>
							<div class="row pcCartRowMain">
								<div class="pcViewCartDesc col-xs-4 col-sm-5">
									<span data-ng-bind-html="shoppingcartitem.description | unsafe">{{shoppingcartitem.description}}</span> <span class="opcSku">({{shoppingcartitem.sku}})</span>
									<div data-ng-show="Evaluate(shoppingcart.writeReview)"><a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct={{shoppingcartitem.id}}');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a></div>
								</div>
								<div class="pcViewCartQty col-xs-1">{{shoppingcartitem.quantity}}</div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right">
									<span class="pcUnitPrice subTitle semibold currency">{{shoppingcartitem.DAUnitPrice}}</span>
								</div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right">
									<span class="pcUnitPrice subTitle semibold currency" data-ng-show="(!IsEmpty(shoppingcartitem.productSubTotal) || Evaluate(shoppingcartitem.IsParentofBundle)) || !IsEmpty(shoppingcartitem.xSellBundleSubTotal)">{{shoppingcartitem.DARowPrice}}</span>
									<span class="pcRowPrice subTitle semibold currency" data-ng-show="IsEmpty(shoppingcartitem.productSubTotal) && !Evaluate(shoppingcartitem.IsParentofBundle) && IsEmpty(shoppingcartitem.xSellBundleSubTotal)">{{shoppingcartitem.DARowPrice}}</span>
								</div>
							</div>
							<% '// END 2nd Row - Main Product Data %>
							<% '// START 3rd Row - BTO Product Details %>
							<div class="row pcViewBTOProductHeading hidden-xs" data-ng-show="!IsEmpty(shoppingcartitem.btoConfiguration)">
								<div class="pcViewCartDesc col-xs-4 col-sm-5">
									<span class="small indent">
									<strong data-ng-show="Evaluate(shoppingcartitem.BToConfigTitle)"><%= bTo_dictLanguage.Item(Session("language")&"_viewcart_2")%></strong>
									<strong data-ng-show="!Evaluate(shoppingcartitem.BToConfigTitle)"><%=bTo_dictLanguage.Item(Session("language")&"_viewcart_1")%></strong>
									</span>
								</div>
								<div class="pcViewCartQty col-xs-1"></div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right"></div>
							</div>
							<div class="row pcViewBTOProductDetails hidden-xs" data-ng-repeat="btoLineItem in shoppingcartitem.btoConfiguration">
								<div class="pcViewCartDesc col-xs-4 col-sm-5">
									<span class="small indent">{{btoLineItem.BToConfigCatDescription}}: {{btoLineItem.BToConfigDescription}}</span>
								</div>
								<div class="pcViewCartQty col-xs-1"></div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right"><span class="currency">{{btoLineItem.BToConfigPrice}}</span></div>
							</div>
							<% '// END 3nd Row - BTO Product Details %>
							<% '// START 4th Row - Product Options %>
							<div class="row pcViewCartOptions hidden-xs" data-ng-repeat="productoption in shoppingcartitem.productoptions">
								<div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent">{{productoption.name}}</span></div>
								<div class="pcViewCartQty col-xs-1"></div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"><span class="currency">{{productoption.DADELunitprice}}</span></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right"><span class="currency">{{productoption.DADELprice}}</span></div>
							</div>
							<% '// END 4th Row - Product Options %>
							<% '// START 5th Row - Custom Input Fields %>
							<div class="row pcViewCartCustomInputHeading hidden-xs" data-ng-repeat="customField in shoppingcartitem.customFields">
								<div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent" data-ng-bind-html="customField.xField | unsafe">{{customField.xField}}</span></div>
							</div>
							<% '// END 5th Row - Custom Input Fields %>
							<% '// START 6th Row - BTO Item Discounts %>
							<div class="row pcViewCartItemDiscounts hidden-xs" data-ng-show="!IsEmpty(shoppingcartitem.itemDiscountRowTotal)">
								<div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent"><%=dictLanguage.Item(Session("language")&"_showcart_23")%></span></div>
								<div class="pcViewCartQty col-xs-1"></div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right"><span class="currency discount">{{shoppingcartitem.itemDiscountRowTotal}}</span></div>
							</div>
							<% '// END 6th Row - BTO Item Discounts %>
							<% '// START 7th Row - BTO Additional Charges %>
							<div class="row pcViewCartBTOChargesHeading hidden-xs" data-ng-show="!IsEmpty(shoppingcartitem.additionalCharges)">
								<div class="pcViewCartDesc col-xs-4 col-sm-5">
									<span class="small indent">
									<strong><%=bto_dictLanguage.Item(Session("language")&"_viewcart_3") %></strong>
									</span>
								</div>
								<div class="pcViewCartQty col-xs-1"></div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right"></div>
							</div>
							<div class="row pcViewCartBTOCharges hidden-xs" data-ng-repeat="btoCharge in shoppingcartitem.additionalCharges">
								<div class="pcViewCartDesc col-xs-4 col-sm-5">
									<span class="small indent">{{btoCharge.categoryDesc}}: {{btoCharge.description}}</span>
								</div>
								<div class="pcViewCartQty col-xs-1"></div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right"><span class="currency">{{btoCharge.total}}</span></div>
							</div>
							<% '// END 7th Row - BTO Additional Charges %>
							<% '// START 8th Row - Quantity Discounts %>
							<div class="row pcViewCartQuantityDiscounts" data-ng-show="!IsEmpty(shoppingcartitem.itemQuantityDiscountRowTotal)">
								<div class="col-xs-5 col-sm-9 right">
									<span class="subTitle light">
									<%= dictLanguage.Item(Session("language")&"_showcart_20")%>
									</span>
								</div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right">
									<span class="pcUnitPrice subTitle currency discount">{{shoppingcartitem.itemQuantityDiscountRowTotal}}</span>
								</div>
							</div>
							<% '// END 8th Row - Quantity Discounts %>
							<% '// START 10th Row - Cross Sell Bundle Discount %>
							<div class="row pcViewCartCrossSellDiscounts" data-ng-show="!IsEmpty(shoppingcartitem.xSellBundleDiscount)">
								<div class="col-xs-5 col-sm-9 right"><span class="subTitle light"><%= dictLanguage.Item(Session("language")&"_showcart_26")%></span></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right">
									<span class="pcUnitPrice subTitle semibold currency">{{shoppingcartitem.xSellBundleDiscount}}</span>
								</div>
							</div>
							<% '// END 10th Row - Cross Sell Bundle Discount %>
							<% '// START 11th Row - Cross Sell Bundle Subtotal %>
							<div class="row pcViewCartCrossSellSubtotal" data-ng-show="!IsEmpty(shoppingcartitem.xSellBundleSubTotal)">
								<div class="col-xs-5 col-sm-9 right"><span class="subTitle light">Bundle Subtotal: <%'= dictLanguage.Item(Session("language")&"_showcart_22")%></span></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right">                                    
									<span class="pcRowPrice subTitle semibold currency">{{shoppingcartitem.xSellBundleSubTotal}}</span>                                        
								</div>
							</div>
							<% '// START 9th Row - Product Subtotal %>
							<div class="row pcViewCartSubtotal" data-ng-show="(!IsEmpty(shoppingcartitem.productSubTotal)) && (IsEmpty(shoppingcartitem.xSellBundleSubTotal))">
								<div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_showcart_22")%></span></div>
								<div class="pcViewCartPrice col-xs-5 col-sm-3 right">
									<span data-ng-show="Evaluate(shoppingcartitem.IsParentofBundle)" class="pcUnitPrice subTitle semibold currency">{{shoppingcartitem.DAproductSubTotal}}</span>
									<span data-ng-show="!Evaluate(shoppingcartitem.IsParentofBundle)" class="pcRowPrice subTitle semibold currency">{{shoppingcartitem.DAproductSubTotal}}</span>
								</div>
							</div>
							<% '// END 9th Row - Product Subtotal %>
							<% '// END 11th Row - Cross Sell Bundle Subtotal %>
							<% '// START 12th Row - Gift Wrapping Message %>
							<!--<div class="row pcCartRowGiftWrapMsg" data-ng-show="!IsEmpty(shoppingcartitem.giftWrapMessage)">
								<div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent">{{shoppingcartitem.giftWrapMessage}}</span></div>
								</div>-->
							<% '// END 12th Row - Gift Wrapping Message %>
							<!--<div class="col-xs-12 rowSpacer" data-ng-show="!Evaluate(shoppingcartitem.IsParentofBundle)"></div>-->
						</div>
						<div class="pcClear"></div>
					</div>
				</div>
				<div class="pcClear"></div>
				<!-- Start: Cart Summary -->            
				<div id="pcCartSummary" class="row">
					<div class="col-xs-12">
						<!--
							<div class="pcViewCartSummaryHeading row">
							    <div class="col-xs-12">
							        <h2>Order Summary</h2>
							    </div>              
							</div>
							-->
						<div class="pcViewCartSummaryBody row">
							<% '// SubTotal %>
							
								<div class="row pcCartOrderSubTotal">   
								                           
								        <div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_15")%></span></div>                      
								        <div class="col-xs-5 col-sm-3 right">
								            <span class="pcUnitPrice subTitle currency">{{shoppingcart.DAsubTotalBeforeDiscounts}}</span>
								        </div>
								    
								</div>
								
							<% '// Display Promotions %>
							<div class="row pcCartRowPromotion" data-ng-repeat="promotion in shoppingcart.promotions">
								<div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{promotion.name}}:</span></div>
								<div class="col-xs-5 col-sm-3 right">
									<span class="pcUnitPrice subTitle currency">{{promotion.price}}</span>
								</div>
							</div>
							<% '// Display category-based quantity discounts %>
							<div class="row pcCartRowCategoryDiscounts" data-ng-show="!IsEmpty(shoppingcart.categoryDiscountTotal)">
								<div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_catdisc_2")%></span></div>
								<div class="col-xs-5 col-sm-3 right">
									<span class="pcUnitPrice subTitle currency discount">{{shoppingcart.categoryDiscountTotal}}</span>
								</div>
							</div>
							<% '// Display order subtotal %>
							<!--
								<div class="row pcCartRowSubTotal" data-ng-show="!IsEmpty(shoppingcart.subtotal)">                         
								    <div class="col-xs-5 col-sm-9 right">
								        <span class="SubTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_15")%></span>
								    </div>                       
								    <div class="col-xs-5 col-sm-3 right">
								        <span class="pcUnitPrice subTitle currency">{{shoppingcart.subtotal}}</span>
								    </div>
								</div>
								-->  
							<div id="pcOrderPreviewTotals">
								<% '// Display Payment Type %>
								<div class="row pcCartRowPaymentTypes" data-ng-show="!IsEmpty(shoppingcart.paymentTotal)">
									<div class="col-xs-5 col-sm-9 right">
										<span class="subTitle" data-ng-show="!IsEmpty(shoppingcart.paymentDescription)">{{shoppingcart.paymentDescription}} <%=dictLanguage.Item(Session("language")&"_orderverify_12")%></span>
										<span class="subTitle" data-ng-show="IsEmpty(shoppingcart.paymentDescription)"><%=dictLanguage.Item(Session("language")&"_orderverify_20")%></span>
									</div>
									<div class="col-xs-5 col-sm-3 right">
										<span class="pcUnitPrice subTitle currency">{{shoppingcart.paymentTotal}}</span>
									</div>
								</div>
								<% '// Discount Table Row %>
								<div class="row pcCartRowDiscounts" data-ng-repeat="discount in shoppingcart.discounts">
									<div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{discount.name}}:</span></div>
									<div class="col-xs-5 col-sm-3 right">
										<span class="pcUnitPrice subTitle currency discount">{{discount.DAprice}}</span>
									</div>
								</div>
								<% '// Reward Points Used %>
								<div class="row pcCartRowRewardsUsed" data-ng-show="!IsEmpty(shoppingcart.rewardPointsUsedTotal)">
									<div class="col-xs-5 col-sm-9 right">
										<span class="subTitle">{{shoppingcart.rewardPointsUsedLabel}} <%=dictLanguage.Item(Session("language")&"_orderverify_31")%></span>
									</div>
									<div class="pcViewCartPrice col-xs-5 col-sm-3 right">
										<span class="pcUnitPrice subTitle currency discount">{{shoppingcart.rewardPointsUsedTotal}}</span>
									</div>
								</div>
								<% '// Gift Wrap Total %>
								<div class="row pcCartRowGWTotal" data-ng-show="!IsEmpty(shoppingcart.giftWrapTotal)">
									<div class="col-xs-5 col-sm-9 right">
										<span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_36a")%></span>
									</div>
									<div class="pcViewCartPrice col-xs-5 col-sm-3 right">
										<span class="pcUnitPrice subTitle currency">{{shoppingcart.giftWrapTotal}}</span>
									</div>
								</div>
								<% '// Shipping Total %>
								<div class="row pcCartRowShippingTotal" data-ng-show="!IsEmpty(shoppingcart.shipmentTotal)">
									<div class="col-xs-5 col-sm-9 right">
										<span class="subTitle">{{shoppingcart.shippingMethod}} <%=dictLanguage.Item(Session("language")&"_orderverify_13")%></span>
									</div>
									<div class="pcViewCartPrice col-xs-5 col-sm-3 right">
										<span class="pcUnitPrice subTitle currency">{{shoppingcart.DAshipmentTotal}}</span>
									</div>
								</div>
								<% '// Shipping Handling Fee %>
								<div class="row pcCartRowHandlingTotal" data-ng-show="!IsEmpty(shoppingcart.serviceHandlingFee)">
									<div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_18")%></span></div>
									<div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{shoppingcart.serviceHandlingFee}}</span></div>
								</div>
								<% '// Taxes and VAT %>
								<div class="row pcCartRowHandlingTotal" data-ng-repeat="tax in shoppingcart.taxes">
									<div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{tax.name}}</span></div>
									<div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{tax.amount}}</span></div>
								</div>
								<% '// Gift Certificates %>
								<div class="row pcCartRowHandlingTotal" data-ng-repeat="giftCert in shoppingcart.giftCerts">
									<div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{giftCert.name}}:</span></div>
									<div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency discount">{{giftCert.amount}}</span></div>
								</div>
								<% '// VAT %>
								<div class="row pcCartRowVATTotal" data-ng-show="!IsEmpty(shoppingcart.vatTotal)">
									<div class="col-xs-5 col-sm-9 right"><span class="subTitle">VAT:</span></div>
									<div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{shoppingcart.DAvatTotal}}</span></div>
								</div>
   								<% '// Cart Total %>
								<div class="row pcCartRowTotal">
									<div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_17")%></span></div>
									<div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{shoppingcart.total}}</span></div>
								</div>

								<% '// Reward Points Earned %>
								<div class="row pcCartRowRewardPointsEarned" data-ng-show="!IsEmpty(shoppingcart.rewardPointsAccrued)">
									<div class="col-xs-115 col-sm-12 right">
                                    <span class="help-block">*{{shoppingcart.rewardPointsAccrued}} <%=dictRewardsLanguage.Item(Session("rewards_language")&"_orderverify")%> <%=dictLanguage.Item(Session("language")&"_orderverify_30")%>
                                      <% if (Session("CustomerGuest")="1") then %>
                                      <div> 
                                        <%= dictLanguage.Item(Session("language")&"_CustviewOrd_49")%><%=RewardsLabel%>
                                      </div>
                                      <% end if %>
                                    </span>
                                    </div>
								</div>
							</div>
						</div>
					</div>
				</div>
	</li>
	</ul>
</div>
</div>
<div class="pcClear"></div>
