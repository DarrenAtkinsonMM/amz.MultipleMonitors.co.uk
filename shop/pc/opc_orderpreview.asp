<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<div class="pcClear"></div>
<div id="opcOrderPreview" data-ng-controller="orderSummaryCtrl" data-ng-cloak class="ng-cloak">
	<h1><%=dictLanguage.Item(Session("language")&"_opc_41")%></h1>
	<div id="opcOrderPreviewWrapper">
		<ul class="pcListLayout">
			<% '// START SDBA - Notify Drop-Shipping %>
			<li class="pcSpacer" data-ng-if="Evaluate(shoppingcart.IsDropShipping)"></li>
			<li data-ng-if="Evaluate(shoppingcart.IsDropShipping)">
				<div class="pcAttention">
					<img src="<%=pcf_getImagePath("images","sds_boxes.gif")%>" alt="<%= ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%>" align="left" vspace="5" hspace="10">
					<%= ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%>
				</div>
			</li>
			<li class="pcSpacer" data-ng-if="Evaluate(shoppingcart.IsDropShipping)"></li>
			<% '// END SDBA - Notify Drop-Shipping %>
			<li data-ng-if="Evaluate(shoppingcart.displayDivider)">
				<!--<hr>-->
			</li>
			<% '// START Order Name %>
			<li data-ng-if="!IsEmpty(shoppingcart.orderNickName)">
				<strong><%=dictLanguage.Item(Session("language")&"_CustviewOrd_40")%></strong>{{shoppingcart.orderNickName}}
			</li>
			<% '// END Order Name %>
			<% '// START Delivery Date/ Time %>
			<li data-ng-if="!IsEmpty(shoppingcart.showDateFrmt)">
				<strong><%=dictLanguage.Item(Session("language")&"_orderverify_34")%></strong>
				{{shoppingcart.showDateFrmt}} <span data-ng-if="!IsEmpty(shoppingcart.savTF1)">-{{shoppingcart.savTF1}}</span></p>
			</li>
			<% '// END Delivery Date/ Time %>
			<% '// START Order Comments %>
			<li data-ng-if="!IsEmpty(shoppingcart.orderComments)">
				<strong><%=dictLanguage.Item(Session("language")&"_orderverify_11")%></strong>{{shoppingcart.orderComments}}
			</li>
			<% '// END Order Comments %>
			<li id="pcOrderPreview" class="pcCartLayout container-fluid">
				<div class="pcTableHeader row">
					<div class="col-xs-4 col-sm-5"><%=dictLanguage.Item(Session("language")&"_showcart_6")%></div>
					<div class="col-xs-1"><%=dictLanguage.Item(Session("language")&"_showcart_4")%></div>
					<div class="col-xs-1 col-sm-3 hidden-xs right"><%=dictLanguage.Item(Session("language")&"_showcart_8b")%></div>
					<div class="col-xs-5 col-sm-3 right"><%=dictLanguage.Item(Session("language")&"_showcart_8")%></div>
				</div>
			
                <div class="pcShoppingCartRow {{shoppingcartitem.rowClass}} row" data-ng-repeat="shoppingcartitem in shoppingcart.shoppingcartrow | filter:{productID: '!!'}">
                    <div class="col-xs-12">
                
                        <% '// START 2nd Row - Main Product Data %>
                        <div class="row pcCartRowMain">
                            <div class="pcViewCartDesc col-xs-4 col-sm-5">
                                <span data-ng-bind-html="shoppingcartitem.description|unsafe">{{shoppingcartitem.description}}</span> <span class="opcSku">({{shoppingcartitem.sku}})</span>
                            </div>
                            <div class="pcViewCartQty col-xs-1">{{shoppingcartitem.quantity}}</div>
                            <div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right">
                                <span class="pcUnitPrice subTitle semibold currency">{{shoppingcartitem.UnitPrice}}</span>
                            </div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle semibold currency" data-ng-if="(!IsEmpty(shoppingcartitem.productSubTotal) || Evaluate(shoppingcartitem.IsParentofBundle)) || !IsEmpty(shoppingcartitem.xSellBundleSubTotal)">{{shoppingcartitem.RowPrice}}</span>
                                <span class="pcRowPrice subTitle semibold currency" data-ng-if="IsEmpty(shoppingcartitem.productSubTotal) && !Evaluate(shoppingcartitem.IsParentofBundle) && IsEmpty(shoppingcartitem.xSellBundleSubTotal)">{{shoppingcartitem.RowPrice}}</span>
                            </div>
                        </div>
                        <% '// END 2nd Row - Main Product Data %>
                        <% '// START 3rd Row - BTO Product Details %>
                        <div class="row pcViewBTOProductHeading hidden-xs" data-ng-if="!IsEmpty(shoppingcartitem.btoConfiguration)">
                            <div class="pcViewCartDesc col-xs-4 col-sm-5">
                                <span class="small indent">
                                <strong data-ng-if="Evaluate(shoppingcartitem.BToConfigTitle)"><%= bTo_dictLanguage.Item(Session("language")&"_viewcart_2")%></strong>
                                <strong data-ng-if="!Evaluate(shoppingcartitem.BToConfigTitle)"><%=bTo_dictLanguage.Item(Session("language")&"_viewcart_1")%></strong>
                                </span>
                            </div>
                            <div class="pcViewCartQty col-xs-1"></div>
                            <div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"></div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right"></div>
                        </div>
                        <div class="row pcViewBTOProductDetails hidden-xs" data-ng-repeat="btoLineItem in shoppingcartitem.btoConfiguration">
                            <div class="pcViewCartDesc col-xs-4 col-sm-5">
                                <span class="small indent">
                                	<span data-ng-bind-html="btoLineItem.BToConfigCatDescription|unsafe">{{btoLineItem.BToConfigCatDescription}}</span>: <span data-ng-bind-html="btoLineItem.BToConfigDescription|unsafe">{{btoLineItem.BToConfigDescription}}</span>
                                </span>
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
                            <div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"><span class="currency">{{productoption.unitprice}}</span></div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right"><span class="currency">{{productoption.price}}</span></div>
                        </div>
                        <% '// END 4th Row - Product Options %>
                        <% '// START 5th Row - Custom Input Fields %>
                        <div class="row pcViewCartCustomInputHeading hidden-xs" data-ng-repeat="customField in shoppingcartitem.customFields">
                            <div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent" data-ng-bind-html="customField.xField|unsafe">{{customField.xField}}</span></div>
                        </div>
                        <% '// END 5th Row - Custom Input Fields %>
                        <% '// START 6th Row - BTO Item Discounts %>
                        <div class="row pcViewCartItemDiscounts hidden-xs" data-ng-if="!IsEmpty(shoppingcartitem.itemDiscountRowTotal)">
                            <div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent"><%=dictLanguage.Item(Session("language")&"_showcart_23")%></span></div>
                            <div class="pcViewCartQty col-xs-1"></div>
                            <div class="pcViewCartUnitPrice col-xs-1 col-sm-3 hidden-xs right"></div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right"><span class="currency discount">{{shoppingcartitem.itemDiscountRowTotal}}</span></div>
                        </div>
                        <% '// END 6th Row - BTO Item Discounts %>
                        <% '// START 7th Row - BTO Additional Charges %>
                        <div class="row pcViewCartBTOChargesHeading hidden-xs" data-ng-if="!IsEmpty(shoppingcartitem.additionalCharges)">
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
                        <div class="row pcViewCartQuantityDiscounts" data-ng-if="!IsEmpty(shoppingcartitem.itemQuantityDiscountRowTotal)">
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
                        <% '// START 9th Row - Product Subtotal %>
                        <div class="row pcViewCartSubtotal" data-ng-if="(!IsEmpty(shoppingcartitem.productSubTotal)) && (IsEmpty(shoppingcartitem.xSellBundleSubTotal))">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_showcart_22")%></span></div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right">
                                <span data-ng-if="Evaluate(shoppingcartitem.IsParentofBundle)" class="pcUnitPrice subTitle semibold currency">{{shoppingcartitem.productSubTotal}}</span>
                                <span data-ng-if="!Evaluate(shoppingcartitem.IsParentofBundle)" class="pcRowPrice subTitle semibold currency">{{shoppingcartitem.productSubTotal}}</span>
                            </div>
                        </div>
                        <% '// END 9th Row - Product Subtotal %>
                        <% '// START 10th Row - Cross Sell Bundle Discount %>
                        <div class="row pcViewCartCrossSellDiscounts" data-ng-if="!IsEmpty(shoppingcartitem.xSellBundleDiscount)">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle light"><%= dictLanguage.Item(Session("language")&"_showcart_26")%></span></div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle semibold currency">{{shoppingcartitem.xSellBundleDiscount}}</span>
                            </div>
                        </div>
                        <% '// END 10th Row - Cross Sell Bundle Discount %>
                        <% '// START 11th Row - Cross Sell Bundle Subtotal %>
                        <div class="row pcViewCartCrossSellSubtotal" data-ng-if="!IsEmpty(shoppingcartitem.xSellBundleSubTotal)">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle light">Bundle Subtotal: <%'= dictLanguage.Item(Session("language")&"_showcart_22")%></span></div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right">                                    
                                <span class="pcRowPrice subTitle semibold currency">{{shoppingcartitem.xSellBundleSubTotal}}</span>                                        
                            </div>
                        </div>
                        <% '// END 11th Row - Cross Sell Bundle Subtotal %>
                        <% '// START 12th Row - Gift Wrapping Message %>
                        <!--<div class="row pcCartRowGiftWrapMsg" data-ng-if="!IsEmpty(shoppingcartitem.giftWrapMessage)">
                            <div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent">{{shoppingcartitem.giftWrapMessage}}</span></div>
                            </div>-->
                        <% '// END 12th Row - Gift Wrapping Message %>
                        <!--<div class="col-xs-12 rowSpacer" data-ng-if="!Evaluate(shoppingcartitem.IsParentofBundle)"></div>-->
                    </div>
                    <div class="pcClear"></div>
                </div>
		
				<div class="pcClear"></div>
                
				<!-- Start: Cart Summary -->
				<div id="pcCartSummary" class="row">
					<div class="col-xs-12">

                        <% '// SubTotal %>
                        <div class="row pcCartOrderSubTotal">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_15")%></span></div>
                            <div class="col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle currency">{{shoppingcart.subTotalBeforeDiscounts}}</span>
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
                        <div class="row pcCartRowCategoryDiscounts" data-ng-if="!IsEmpty(shoppingcart.categoryDiscountTotal)">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_catdisc_2")%></span></div>
                            <div class="col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle currency discount">{{shoppingcart.categoryDiscountTotal}}</span>
                            </div>
                        </div>

                        <% '// Display Payment Type %>
                        <div class="row pcCartRowPaymentTypes" data-ng-if="!IsEmpty(shoppingcart.paymentTotal) && Evaluate(shoppingcart.checkoutStage) && Evaluate(shoppingcart.checkoutStage)">
                            <div class="col-xs-5 col-sm-9 right">
                                <span class="subTitle" data-ng-if="!IsEmpty(shoppingcart.paymentDescription)">{{shoppingcart.paymentDescription}} <%=dictLanguage.Item(Session("language")&"_orderverify_12")%></span>
                                <span class="subTitle" data-ng-if="IsEmpty(shoppingcart.paymentDescription)"><%=dictLanguage.Item(Session("language")&"_orderverify_20")%></span>
                            </div>
                            <div class="col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle currency">{{shoppingcart.paymentTotal}}</span>
                            </div>
                        </div>

                        <% '// Discount Table Row %>
                        <div class="row pcCartRowDiscounts" data-ng-repeat="discount in shoppingcart.discounts | filter:{name: '!!'}">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{discount.name}}:</span></div>
                            <div class="col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle currency discount">{{discount.price}}</span>
                            </div>
                        </div>
                            
                    
                        <% '// Reward Points Used %>
                        <div class="row pcCartRowRewardsUsed" data-ng-if="!IsEmpty(shoppingcart.rewardPointsUsedTotal) && Evaluate(shoppingcart.checkoutStage)">
                            <div class="col-xs-5 col-sm-9 right">
                                <span class="subTitle">{{shoppingcart.rewardPointsUsedLabel}} <%=dictLanguage.Item(Session("language")&"_orderverify_31")%></span>
                            </div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle currency discount">{{shoppingcart.rewardPointsUsedTotal}}</span>
                            </div>
                        </div>
                            
                        <% '// Gift Wrap Total %>
                        <div class="row pcCartRowGWTotal" data-ng-if="!IsEmpty(shoppingcart.giftWrapTotal) && Evaluate(shoppingcart.checkoutStage)">
                            <div class="col-xs-5 col-sm-9 right">
                                <span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_36a")%></span>
                            </div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle currency">{{shoppingcart.giftWrapTotal}}</span>
                            </div>
                        </div>
                            
                        <% '// Shipping Total %>
                        <div class="row pcCartRowShippingTotal" data-ng-if="!IsEmpty(shoppingcart.shipmentTotal)">
                            <div class="col-xs-5 col-sm-9 right">
                                <span class="subTitle" data-ng-bind-html="shoppingcart.shippingMethod|unsafe"></span> <span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_13")%></span>
                            </div>
                            <div class="pcViewCartPrice col-xs-5 col-sm-3 right">
                                <span class="pcUnitPrice subTitle currency">{{shoppingcart.shipmentTotal}}</span>
                            </div>
                        </div>
                            
                        <% '// Shipping Handling Fee %>
                        <div class="row pcCartRowHandlingTotal" data-ng-if="!IsEmpty(shoppingcart.serviceHandlingFee)">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_18")%></span></div>
                            <div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{shoppingcart.serviceHandlingFee}}</span></div>
                        </div>
                            
                        <% '// Taxes and VAT %>
                        <div class="row pcCartRowHandlingTotal" data-ng-if="Evaluate(shoppingcart.checkoutStage)" data-ng-repeat="tax in shoppingcart.taxes">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{tax.name}}</span></div>
                            <div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{tax.amount}}</span></div>
                        </div>
                        
                        <% '// Gift Certificates %>
                        <div class="row pcCartRowHandlingTotal" data-ng-if="Evaluate(shoppingcart.checkoutStage)" data-ng-repeat="giftCert in shoppingcart.giftCerts">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{giftCert.name}}:</span></div>
                            <div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency discount">{{giftCert.amount}}</span></div>
                        </div>
                    
                        <% '// Cart Total - NOT Logged In %>
                        <div class="row pcCartRowTotal" data-ng-if="!Evaluate(shoppingcart.checkoutStage) && Evaluate(shoppingcart.haveDiscounts)">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_17")%></span></div>
                            <div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{shoppingcart.subtotal}}</span></div>
                        </div>                           

                        <% '// Cart Total - Logged In %>
                        <div class="row pcCartRowTotal" data-ng-if="Evaluate(shoppingcart.checkoutStage)">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_17")%></span></div>
                            <div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{shoppingcart.total}}</span></div>
                        </div>
                        
                        <% '// VAT %>
                        <div class="row pcCartRowVATTotal" data-ng-if="Evaluate(shoppingcart.checkoutStage)" data-ng-if="!IsEmpty(shoppingcart.vatTotal)">
                            <div class="col-xs-5 col-sm-9 right"><span class="subTitle">{{shoppingcart.vatName}}</span></div>
                            <div class="col-xs-5 col-sm-3 right"><span class="pcUnitPrice subTitle currency">{{shoppingcart.vatTotal}}</span></div>
                        </div>
                        
                        <% '// Reward Points Earned %>
                        <div class="row pcCartRowRewardPointsEarned" data-ng-if="!IsEmpty(shoppingcart.rewardPointsAccrued) && Evaluate(shoppingcart.checkoutStage)">
                            <div class="col-xs-115 col-sm-12 right"><span class="help-block">*{{shoppingcart.rewardPointsAccrued}} <%=dictRewardsLanguage.Item(Session("rewards_language")&"_orderverify")%> <%=dictLanguage.Item(Session("language")&"_orderverify_30")%></span></div>
                        </div>

					</div>
				</div>
	        </li>
	    </ul>
    </div>
</div>
<div class="pcClear"></div>