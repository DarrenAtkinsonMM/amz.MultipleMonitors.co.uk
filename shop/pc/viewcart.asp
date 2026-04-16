<%@  language="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "viewcart.asp"
' This page displays the items in the cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="viewcart_init.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="viewcart_pageLoad.asp"-->
<!--#include file="viewcart_modules.asp"-->
	<!-- Header: pagetitle -->
    <header id="cartcontent" class="cartcontent">
		<div class="ct-content">
			<div class="container">
				<div class="row">
					<div class="col-sm-12 cart-title">
                         <div class="wow fadeInUp" data-wow-offset="0" data-wow-delay="0">
							<h3 class="color marginbot-0">Your Shopping Cart <span>Contains</span></h3>
						</div>
                    </div>					
				</div>		
			</div>		
		</div>	
    </header>
    <!-- Section: product-detail -->
    <section id="cart-detail" class="paddingbot-30 ">
		<div class="container">
			<div class="row">
			<div class="col-sm-12">
			<div class="wow fadeInUp" data-wow-delay="0">
<div id="pcMain" data-ng-cloak class="ng-cloak pcViewCart">
    <div class="pcMainContent" data-ng-controller="orderSummaryCtrl">
        <form method="post" action="cRec.asp" name="recalculate" id="recalculate" class="pcForms" onSubmit="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert(alert_8b); return(false);}">
            
            <% If Len(Trim(strCCSLCheck))>0 Then %>
                <div class="pcErrorMessage">
                    <%= dictLanguage.Item(Session("language")&"_alert_19") & strCCSLcheck %>
                </div>
            <% End If %>
    
            <div data-ng-if="Evaluate(shoppingcart.IsEditOrder)">
                <div class="pcInfoMessage">
                    <%= dictLanguage.Item(Session("language")&"_SB_34A")%> <%= Session("SBEditOrderID")%> <%= dictLanguage.Item(Session("language")&"_SB_34B")%>
                </div>
            </div>
    
            <!-- Start: Messages -->
            <!--#include file="viewcart_messages.asp"-->
            <!-- Start: Messages -->
    
            <!-- Start: Cart -->
            <div id="pcCart" class="pcCartLayout container-fluid">
    
                    <!-- Start: Cart Header -->
                    <div class="pcTableHeader row">                        
                        <div class="col-xs-1"></div>
                        <div class="col-xs-4 col-sm-5"><%=dictLanguage.Item(Session("language")&"_showcart_6")%></div>
                        <div class="col-xs-1 col-sm-2 hidden-xs right"><%=dictLanguage.Item(Session("language")&"_showcart_8b")%></div>
                        <div class="col-xs-4 col-sm-2 right"><%=dictLanguage.Item(Session("language")&"_showcart_8")%></div>
                        <div class="col-xs-2"><%=dictLanguage.Item(Session("language")&"_showcart_4")%></div>
                    </div>
                    <!-- End: Cart Header -->
        
                    <!-- Start: Cart Rows -->
                    <div class="pcShoppingCartRow {{shoppingcartitem.rowClass}} row" data-ng-repeat="shoppingcartitem in shoppingcart.shoppingcartrow | filter:{productID: '!!'}">
                        <div class="col-xs-12">
        
                            <% '// START 1st Row - Bundle Header %>
                            
                            <!--                            
                            <div class="row bundle-header" data-ng-if="Evaluate(shoppingcartitem.IsParentofBundle)">    
                                <div class="col-xs-1"></div>
                                <div class="col-xs-9 col-sm-11"><span class="pcItemDescription title bold">Bundle {{shoppingcartitem.BundleTitle}}</span></div>
                            </div>
                
                            <div class="row bundle-subheader" data-ng-if="Evaluate(shoppingcartitem.IsParentofBundle)">    
                                <div class="col-xs-1"></div>
                                <div class="col-xs-9 col-sm-11"><span class="small light">The following items are included in the bundle:</span></div> 
                            </div>
                            -->
                            
                            <div class="pcSpacer" data-ng-if="!Evaluate(shoppingcartitem.IsParentofBundle)"></div>
                            
                            <% '// END 1st Row - Bundle Header %>
                
                
                            <% '// START 2nd Row - Main Product Data %>
                            
                            <div class="row pcViewCartMain">
                
                                <div class="pcViewCartItem col-xs-1">
                                
                                    <div data-ng-if="!Evaluate(shoppingcart.IsBuyGift)">
                                        <a data-ng-if="Evaluate(shoppingcartitem.ShowImage);" data-ng-href="{{shoppingcartitem.daproductURL}}"><img data-ng-src="catalog/{{shoppingcartitem.ImageURL}}" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")%>{{shoppingcartitem.description}}"></a>                                    
                                    </div>
                
                                    <div data-ng-if="Evaluate(shoppingcart.IsBuyGift)">
                                        <a data-ng-if="Evaluate(shoppingcartitem.ShowImage);" data-ng-href="ggg_viewEP.asp?grCode={{shoppingcart.grCode}}&amp;geID={{shoppingcartitem.geID}}"><img data-ng-src="catalog/{{shoppingcartitem.ImageURL}}" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")%>{{shoppingcartitem.description}}"></a>                                    
                                    </div>
                                
                                </div>
                
                
                                <div class="pcViewCartDesc col-xs-4 col-sm-5">
                
                                    <div data-ng-if="!Evaluate(shoppingcart.IsBuyGift)">
                                        <a data-ng-href="{{shoppingcartitem.daproductURL}}"><span class="pcItemDescription title bold" data-ng-bind-html="shoppingcartitem.description|unsafe">{{shoppingcartitem.description}}</span></a> <span class="pcItemSKU small light">({{shoppingcartitem.sku}})</span> 
                                    </div>
                
                                    <div data-ng-if="Evaluate(shoppingcart.IsBuyGift)">
                                        <a rel="nofollow" data-ng-href="ggg_viewEP.asp?grCode={{shoppingcart.grCode}}&amp;geID={{shoppingcartitem.geID}}"><span class="pcItemDescription title bold" data-ng-bind-html="shoppingcartitem.description|unsafe">{{shoppingcartitem.description}}</span></a> <span class="pcItemSKU small light">({{shoppingcartitem.sku}})</span> 
                                    </div>
                                    
                                    
                                    <div data-ng-if="Evaluate(shoppingcartitem.giftWrapStatus) && Evaluate(shoppingcart.ShowGiftWrapOptions)" class="hidden-xs">
                                        <span  class="pcItemGiftWrap small light"><%= dictLanguage.Item(Session("language")&"_showcart_24")%>&nbsp;<%=dictLanguage.Item(Session("language")&"_showcart_30")%></span>
                                    </div>
                                    <div data-ng-if="!Evaluate(shoppingcartitem.giftWrapStatus) && Evaluate(shoppingcart.ShowGiftWrapOptions)" class="hidden-xs">
                                        <span  class="pcItemGiftWrap small light"><%= dictLanguage.Item(Session("language")&"_showcart_24")%>&nbsp;<%=dictLanguage.Item(Session("language")&"_showcart_25")%></span>
                                    </div>
                                    
                                    <div data-ng-if="Evaluate(shoppingcartitem.IsReconfigurable);">
                                        <a href="Reconfigure.asp?pcCartIndex={{shoppingcartitem.row}}"><span class="pcItemActions small light"><%=dictLanguage.Item(Session("language")&"_css_reconfigure")%></span></a>
                                    </div>
                                    
                                    <div data-ng-if="Evaluate(shoppingcartitem.IsApparel);">
                                        <span  class="small light"><a href="viewPrd.asp?idproduct={{shoppingcartitem.productID}}&amp;index={{shoppingcartitem.row}}&amp;imode=updOrd"><%=dictLanguage.Item(Session("language")&"_showcart_21")%></a></span>
                                    </div>
                
                                </div>
                                
                
                                <div class="pcViewCartUnitPrice col-xs-1 col-sm-2 hidden-xs right">
                                    <span class="pcUnitPrice subTitle semibold currency">{{shoppingcartitem.DAUnitPrice}}</span>
                                </div>
                
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right">
                                    <span class="pcUnitPrice subTitle semibold currency" data-ng-if="(!IsEmpty(shoppingcartitem.productSubTotal) || Evaluate(shoppingcartitem.IsParentofBundle)) || !IsEmpty(shoppingcartitem.xSellBundleSubTotal)">{{shoppingcartitem.DARowPrice}}</span>
                                    <span class="pcRowPrice subTitle bold currency" data-ng-if="IsEmpty(shoppingcartitem.productSubTotal) && !Evaluate(shoppingcartitem.IsParentofBundle) && IsEmpty(shoppingcartitem.xSellBundleSubTotal)">{{shoppingcartitem.DARowPrice}}</span>
                                </div>
                
                                <div class="pcViewCartQty col-xs-2">
 
                
                                    <div data-ng-if="Evaluate(shoppingcartitem.IsRemoveable);">
                                        <a href="cRemv.asp?pcCartIndex={{shoppingcartitem.row}}" class="btn btnItem-delete"><i class="fa fa-close"></i></a>
                                    </div>
                                    
                                    <div id="pcUpdate{{shoppingcartitem.row}}" class="pcUpdateButton">
                                     <input type="text" min="1" id="Cant{{shoppingcartitem.row}}" name="Cant{{shoppingcartitem.row}}" size="3" data-ng-value="shoppingcartitem.quantity" 
                                        class="pcQuantity {{shoppingcartitem.quantityClass}} sc-qInp"
                                        data-ng-readonly="Evaluate(shoppingcartitem.quantityReadOnly);"
                                        data-ng-model="shoppingcartitem.quantity"
                                        data-ng-change="CheckQuantityMins('Cant' + {{shoppingcartitem.row}},{{shoppingcartitem.MinimumQty}},{{shoppingcartitem.QtyValidate}},{{shoppingcartitem.QtyIDEvent}},{{shoppingcartitem.MultiQty}},{{shoppingcartitem.quantityValidate}})"
                                        data-ng-mousedown="ShowUpdateButton(shoppingcartitem.row)"
                                        data-ng-blur="HideUpdateButton(shoppingcartitem.row)"
                                        onkeypress="return handleEnter(this, event)" />
                                        
                                    
                                        <button id="submit" name="Submit" class="btn btn-xs btn-skin btn-updcart updateBtn" data-ng-onclick="javascript: if ((RemainIssue != '') || (RemainIssue1 != '')) { alert(alert_8b); return (false); } else { return (true); }">
                                        <%= dictLanguage.Item(Session("language")&"_css_recalculate") %>
                                        </button>
                                    </div>
                                    
                                    <div class="pcClear"></div>
        
                                    <input type="hidden" name="SavQty{{shoppingcartitem.row}}" data-ng-value="shoppingcartitem.savequantity" />
                                    
                                </div>
                
                            </div>
                
                            <% '// END 2nd Row - Main Product Data %>
                
                
                            <% '// START 3rd Row - BTO Product Details %>
                
                            <div class="row pcViewBTOProductHeading hidden-xs" data-ng-if="!IsEmpty(shoppingcartitem.btoConfiguration)">                            
                                <div class="pcViewCartItem col-xs-1"></div>
                                <div class="pcViewCartDesc col-xs-4 col-sm-5">
                                    <span class="small indent">
                                        <strong data-ng-if="Evaluate(shoppingcartitem.BToConfigTitle)"><%= bTo_dictLanguage.Item(Session("language")&"_viewcart_2")%></strong>
                                        <strong data-ng-if="!Evaluate(shoppingcartitem.BToConfigTitle)"><%=bTo_dictLanguage.Item(Session("language")&"_viewcart_1")%></strong>
                                        <a href="Reconfigure.asp?pcCartIndex={{shoppingcartitem.row}}"><span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_reconfigure")%></span></a>
                                    </span>
                                </div>
                                <div class="pcViewCartQty col-xs-2"></div>
                                <div class="pcViewCartUnitPrice col-xs-1 col-sm-2 hidden-xs right"></div>
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right"></div>
                            </div>
                
                            <div class="row pcViewBTOProductDetails hidden-xs" data-ng-repeat="btoLineItem in shoppingcartitem.btoConfiguration">
                                <div class="pcViewCartItem col-xs-1"></div>
                                <div class="pcViewCartDesc col-xs-4 col-sm-5">
                                    <span class="small indent">
                                    	<span data-ng-bind-html="btoLineItem.BToConfigCatDescription|unsafe">{{btoLineItem.BToConfigCatDescription}}</span>: <span data-ng-bind-html="btoLineItem.BToConfigDescription|unsafe">{{btoLineItem.BToConfigDescription}}</span>
                                    </span>
                                </div>
                                <div class="pcViewCartQty col-xs-2">{{btoLineItem.BToConfigQuantity}}</div>
                                <div class="pcViewCartUnitPrice col-xs-1 col-sm-2 hidden-xs right"></div>
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right"><span class="currency">{{btoLineItem.BToConfigPrice}}</span></div>
                            </div>
                
                            <% '// END 3nd Row - BTO Product Details %>
                
                
                            <% '// START 4th Row - Product Options %>
                
                            <div class="row pcViewCartOptions hidden-xs" data-ng-repeat="productoption in shoppingcartitem.productoptions">                            
                               <div class="pcViewCartItem col-xs-1"></div>
                                <div class="pcViewCartDesc col-xs-4 col-sm-5"><span class="small indent">{{productoption.name}}</span></div>
                                <div class="pcViewCartUnitPrice col-xs-1 col-sm-2 hidden-xs right"><span class="currency small">{{productoption.daunitprice}}</span></div>
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right"><span class="currency small">{{productoption.daprice}}</span></div>
                                <div class="pcViewCartQty col-xs-2"></div>
                            </div>
                
                            <% '// END 4th Row - Product Options %>
                
                
                            <% '// START 5th Row - Custom Input Fields %>
                               
                            <% '// END 5th Row - Custom Input Fields %>
                
                            
                            <% '// START 6th Row - BTO Additional Charges %>
                
                            <div class="row pcViewCartBTOChargesHeading hidden-xs" data-ng-if="!IsEmpty(shoppingcartitem.additionalCharges)">
                                <div class="pcViewCartItem col-xs-1"></div>
                                <div class="pcViewCartDesc col-xs-4 col-sm-5">
                                    <span class="small indent">
                                        <strong><%=bto_dictLanguage.Item(Session("language")&"_viewcart_3") %></strong>
                                        <% if pcv_FinalizedQuote=0 then %>
                                            <a href="RePrdAddCharges.asp?pcCartIndex={{shoppingcartitem.row}}"><span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_reconfigure")%></span></a>
                                        <% end if %>
                                    </span>
                                </div>
                                <div class="pcViewCartQty col-xs-2"></div>
                                <div class="pcViewCartUnitPrice col-xs-1 col-sm-2 hidden-xs right"></div>
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right"></div>
                            </div>
                
                            <div class="row pcViewCartBTOCharges hidden-xs" data-ng-repeat="btoCharge in shoppingcartitem.additionalCharges">                            
                                <div class="pcViewCartItem col-xs-1"></div>
                                <div class="pcViewCartDesc col-xs-4 col-sm-5">
                                    <span class="small indent">{{btoCharge.categoryDesc}}: {{btoCharge.description}}</span>
                                </div>
                                <div class="pcViewCartQty col-xs-2"></div>
                                <div class="pcViewCartUnitPrice col-xs-1 col-sm-2 hidden-xs right"></div>
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right"><span class="currency">{{btoCharge.total}}</span></div>
                            </div>
                
                            <% '// END 6th Row - BTO Additional Charges %>
                            
                            <% '// START 7th Row - BTO Item Discounts %>
                
                            <div class="row pcViewCartItemDiscounts" data-ng-if="!IsEmpty(shoppingcartitem.itemDiscountRowTotal)">
                                <div class="pcViewCartItem col-xs-1"></div>
                                <div class="col-xs-6 col-sm-9 right">
                                    <span class="small indent">
                                        <%=dictLanguage.Item(Session("language")&"_showcart_23")%>
                                    </span>
                                </div>
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right">
                                    <span class="currency discount">{{shoppingcartitem.itemDiscountRowTotal}}</span>
                                </div>
                            </div>
                
                            <% '// END 7th Row - BTO Item Discounts %>
                            
                            <% '// START 8th Row - Quantity Discounts %>
                
                            <div class="row pcViewCartQuantityDiscounts" data-ng-if="!IsEmpty(shoppingcartitem.itemQuantityDiscountRowTotal)">
                                <div class="pcViewCartItem col-xs-1"></div>
                                <div class="col-xs-6 col-sm-9 right">
                                    <span class="subTitle light">
                                        <a data-ng-if="!Evaluate(shoppingcartitem.IsApparel)" href="javascript:openbrowser('priceBreaks.asp?idproduct={{shoppingcartitem.itemQuantityDiscountRowID}}')"><%= dictLanguage.Item(Session("language")&"_showcart_20")%></a>
                                        <a data-ng-if="Evaluate(shoppingcartitem.IsApparel)" href="javascript:openbrowser('app-subPrdDiscount.asp?idproduct={{shoppingcartitem.itemQuantityDiscountRowID}}')"><%= dictLanguage.Item(Session("language")&"_showcart_20")%></a>
                                    </span>
                                </div>
                                <div class="pcViewCartPrice col-xs-4 col-sm-2 right">
                                    <span class="pcUnitPrice subTitle semibold currency discount">{{shoppingcartitem.itemQuantityDiscountRowTotal}}</span>
                                </div>
                            </div>
                
                            <% '// END 8th Row - Quantity Discounts %>
                            
                
                            <% '// START 9th Row - Product Subtotal %>
                            <% '// END 9th Row - Product Subtotal %>
                
                
                            <% '// START 10th Row - Cross Sell Bundle Discount %>
                
                                <div class="row pcViewCartCrossSellDiscounts" data-ng-if="!IsEmpty(shoppingcartitem.xSellBundleDiscount)">
                                    <div class="pcViewCartItem col-xs-1"></div>
                                    <div class="col-xs-6 col-sm-9 right"><span class="subTitle light"><%= dictLanguage.Item(Session("language")&"_showcart_26")%></span></div>
                                    <div class="pcViewCartPrice col-xs-4 col-sm-2 right">
                                        <span class="pcUnitPrice subTitle semibold currency discount">{{shoppingcartitem.xSellBundleDiscount}}</span>
                                    </div>
                                </div>
                
                            <% '// END 10th Row - Cross Sell Bundle Discount %>
                
                
                            <% '// START 11th Row - Cross Sell Bundle Subtotal %>
                
                                <div class="row pcViewCartCrossSellSubtotal" data-ng-if="!IsEmpty(shoppingcartitem.xSellBundleSubTotal)">
                                    <div class="pcViewCartItem col-xs-1"></div>
                                    <div class="col-xs-6 col-sm-9 right"><span class="subTitle light">Bundle Subtotal: <%'= dictLanguage.Item(Session("language")&"_showcart_22")%></span></div>
                                    <div class="pcViewCartPrice col-xs-4 col-sm-2 right">                                    
                                        <span class="pcRowPrice subTitle bold currency">{{shoppingcartitem.xSellBundleSubTotal}}</span>                                        
                                    </div>
                                </div>
                
                            <% '// END 11th Row - Cross Sell Bundle Subtotal %>
                
                           
                            <!--<div class="col-xs-12 rowSpacer" data-ng-if="!Evaluate(shoppingcartitem.IsParentofBundle)"></div>-->
        					<div class="daCartDash"></div>
                        </div>
                    </div>
                    <!-- End: Cart Rows -->
                    
                    <div data-ng-if="Evaluate(shoppingcart.daCblBool);" class="pcShoppingCartRow row">
						<div class="col-xs-12">
							<div class="row pcViewCartMain">
								<div class="pcViewCartItem col-xs-1">
									<img src="/images/cart/cart-cable-thumb.jpg" alt="Free Cables">
								</div>
								<div class="pcViewCartDesc col-xs-4 col-sm-5">
									<span class="pcItemDescription title bold">FREE 3m Digital Video & Power Leads Pack - (Worth {{shoppingcart.daCblValue}}!)</span>
								</div>
								<div class="pcViewCartUnitPrice col-xs-1 col-sm-2 hidden-xs right">
									<span class="pcUnitPrice subTitle semibold currency">£15.00</span>
								</div>
								<div class="pcViewCartPrice col-xs-4 col-sm-2 right">
									<span class="pcUnitPrice subTitle semibold currency"><del>{{shoppingcart.daCblValue}}</del></span>
								</div>
								<div class="col-xs-2">
								</div>
							</div>
                            <div class="daCartDash"></div>
						</div>
					</div>
                    
                    <div data-ng-if="Evaluate(shoppingcart.daMsgBool);">
						<section id="cart-option" class="paddingbot-40 ">
							<div class="container">
								<div class="row">
									<div class="col-sm-12">
										
											<div class="inside-cart-option">
												<h3>{{shoppingcart.daMsgHeader}}</h3>
												<p>{{shoppingcart.daMsgBody}}</p>
												<a href="{{shoppingcart.daMsgUrl}}">{{shoppingcart.daMsgButText}}</a>
											</div>
										
									</div>
								</div>
							</div>
						</section>
					</div>
                    
                    <div data-ng-if="Evaluate(shoppingcart.daBundleDiscountApplied);" class="row daPricingSummary">
                        <div class="col-xs-1 hidden-xs"></div>
                        <div class="col-xs-4 col-sm-5 hidden-xs"></div>
                        <div class="col-xs-1 col-sm-2 right daPricingSummaryResp">Bundle Discount:</div>
                        <div class="col-xs-4 col-sm-2 right">- {{shoppingcart.daBundleDiscount}}</div>
                        <div class="col-xs-2 hidden-xs"></div>
                    </div>	
                    <div class="row daPricingSummary">                        
                        <div class="col-xs-1 hidden-xs"></div>
                        <div class="col-xs-4 hidden-xs col-sm-5"></div>
                        <div class="col-xs-1 col-sm-2 right daPricingSummaryResp">Sub total:</div>
                        <div class="col-xs-4 col-sm-2 right">{{shoppingcart.dasubTotalBeforeDiscounts}}</div>
                        <div class="col-xs-2 hidden-xs"></div>
                    </div> 
                    <div class="row daPricingSummary">                        
                        <div class="col-xs-1 hidden-xs"></div>
                        <div class="col-xs-4 col-sm-5 hidden-xs"></div>
                        <div class="col-xs-1 col-sm-2 right daPricingSummaryResp">UK Delivery:</div>
                        <div class="col-xs-4 col-sm-2 right">{{shoppingcart.daDelCharge}}</div>
                        <div class="col-xs-2 hidden-xs"></div>
                    </div>     
                    <div class="row daPricingSummary">                        
                        <div class="col-xs-1 hidden-xs"></div>
                        <div class="col-xs-4 col-sm-5 hidden-xs"></div>
                        <div class="col-xs-1 col-sm-2 right daPricingSummaryResp">VAT:</div>
                        <div class="col-xs-4 col-sm-2 right">{{shoppingcart.daVAT}}</div>
                        <div class="col-xs-2 hidden-xs"></div>
                    </div>     
                    <div class="row daPricingSummary">                        
                        <div class="col-xs-1 hidden-xs"></div>
                        <div class="col-xs-4 col-sm-5 hidden-xs"></div>
                        <div class="col-xs-1 col-sm-2 right daPricingSummaryResp">Total:</div>
                        <div class="col-xs-4 col-sm-2 right">{{shoppingcart.daFinalTotal}}</div>
                        <div class="col-xs-2 hidden-xs"></div>
                    </div>  
                     <div class="row daPricingDelEst">                        
                        <div class="col-xs-1"></div>
                        <div data-ng-if="!Evaluate(shoppingcart.daFunDelDateBlockTest);">
                        <div class="col-xs-4 col-sm-9 right daDelEstResp"><strong class="color">Delivery Estimate:</strong> Order by <strong class="color">{{shoppingcart.daDelCutOff}}</strong> for delivery on <strong class="color">{{shoppingcart.daDelDate}}</strong>.</div>
                        </div>
                        <div data-ng-if="Evaluate(shoppingcart.daFunDelDateBlockTest);">
                        <div class="col-xs-4 col-sm-9 right daDelEstResp"><strong class="color">Delivery Estimate:</strong> Due to a short workshop closure, orders will now be delivered on <strong class="color">{{shoppingcart.daDelDate}}</strong>.</div>
                        </div>
                        <div class="col-xs-2"></div>
                    </div>                     
 
                      
                     
                    <!-- Start: Cart Recalc Button
                    <div id="pcViewCartRecalculate" class="row">               
                        <div class="col-xs-5 col-sm-6"></div>
                        <div class="col-xs-7 col-sm-6">
                            <button class="pcButton pcButtonRecalculate secondary tiny" id="submit" name="Submit" data-ng-onclick="javascript: if ((RemainIssue != '') || (RemainIssue1 != '')) { alert(alert_8b); return (false); } else { return (true); }">
                                <img src="<%=RSlayout("recalculate")%>" alt="<%= dictLanguage.Item(Session("language")&"_alert_8b") %>" />
                                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_recalculate") %></span>
                            </button>
                        </div>                       
                    </div> 
                     End: Cart Recalc Button -->
                     
                    <div class="row daOrderBut">                        
                        <div class="col-xs-1"></div>
                        <div class="col-xs-4 col-sm-5"></div>
                        <div class="col-xs-1 col-sm-2 hidden-xs right"></div>
                        <div class="col-xs-4 col-sm-2 right daDelCheckBtnResp">
                           <div data-ng-if="!Evaluate(shoppingcart.HaveGcs) && Evaluate(shoppingcart.ShowCheckoutBtn)">                                        
                           	<a class="btn btn-skin btn-wc ckotBtn" href="/shop/pc/onepagecheckout.asp" onClick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert(alert_8b); return(false);} return(checkQtyChange());">
                            	<span class="pcButtonText3">Place Order & Pay</span>
                            </a>
                            <span data-ng-if="Evaluate(shoppingcart.haveGR)">
                            	<a class="pcButton pcButtonAddToRegistry" href="ggg_addtoGR.asp" onClick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert(alert_8b); return(false);}">
                            		<img src="<%=RSlayout("AddToRegistry")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_14")%>">
                                    <span class="pcButtonText4"><%=dictLanguage.Item(Session("language")&"_css_addtoregistry")%></span>
                                </a>
                            </span>
                         </div>
                        </div>
                        <div class="col-xs-2 hidden-xs"></div>
                    </div>     
        
                    <input type="hidden" name="actGCs" value="">        

            </div>
            <!-- End: Cart -->
    
    
            <div class="pcSpacer"></div>
            <%
            '// Show CrossSelling
            'pcs_ShowCrossSelling()            
            %>
    
            <%
            '// START: Saved Cart Modal Content
            %>
            <div class="modal fade" id="pcSavedCartModal" tabindex="-1" role="dialog" aria-labelledby="<%=dictLanguage.Item(Session("language")&"_SaveCart_1")%>" aria-hidden="true">
               <div class="modal-dialog modal-dialog-center">
                  <div class="modal-content">
                     <div class="modal-header">
                        <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                        <h4 class="modal-title"><%=dictLanguage.Item(Session("language")&"_SaveCart_1")%></h4>
                     </div>
                     <div class="modal-body">
                        <div id="pcSaveCartContents">
                            <p><%=dictLanguage.Item(Session("language")&"_SaveCart_2")%></p>
                            <div class="validateTips"></div>
                            <div>
                                <%= dictLanguage.Item(Session("language")&"_SaveCart_3") %>:  
                                <input type="text" name="SavedCartName" id="SavedCartName" />
                            </div>
                        </div>
                     </div>
                     <div class="modal-footer">
                        <div id="pcSaveCartButtons">
                            <a href="javascript:saveCart();" role="button" class="btn btn-default">Save Cart</a>
                            <a role="button" class="btn btn-default" data-dismiss="modal">Cancel</a>
                        </div>
                     </div>
                  </div>
               </div>
            </div>
            <%
            '// END: Saved Cart Modal Content
            %>
        
        </form> 
    </div>   
</div>
<script>
<%if scDispDiscCart="1" then%>
	var scDispDiscCart="1";
<%else%>
	var scDispDiscCart="0";
<%end if%>
</script>
			</div>
            </div>
            </div>
		</div>
    </section>
    <section id="updatecart-option" class="paddingbot-30 ">
		<div class="container">
			<div class="row">
			<div class="col-sm-12 col-md-10">
			<div class="wow fadeInUp" data-wow-delay="0.1s">
			    <div class="cartVisa">
				   <ul>
				   <li><img src="/images/cart/v-1.png"></li>
				   <li><img src="/images/cart/v-2.png"></li>
				   <li><img src="/images/cart/v-3.png"></li>
				   <li><img src="/images/cart/v-4.png"></li>
				   <li><img src="/images/cart/v-5.png"></li>
				   <li><img src="/images/cart/v-6.png"></li>
				   <li><img src="/images/cart/v-7.png"></li>
				   <li><img src="/images/cart/v-8.png"></li>
				   </ul>
				</div>
			</div>
            </div>
            </div>
		</div>
    </section>
<!--#include file="footer_wrapper.asp"-->
