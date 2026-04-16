<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>   
<% If Instr(Ucase(Request.ServerVariables("SCRIPT_NAME")), "GW") = 0 Then %> 
        
<div id="quickCartContainer" data-ng-controller="QuickCartCtrl">
    <div data-ng-cloak class="ng-cloak">


        <div data-ng-hide="shoppingcart.totalQuantity>0" class="pcIconBarViewCart">
            <%'=dictLanguage.Item(Session("language")&"_smallcart_11")%>
            <%'DA - EDIT EMPTY CART%>
            <p class="user-box disp-inline"><a href="/shop/pc/custPref.asp" class="user-action">Existing Customer Login</a></p>
			<p class="user-cart disp-inline"><span class=""><i class="fa fa-shopping-basket"></i> 0 item(s)</span></p>
        </div>  


        <div data-ng-show="shoppingcart.totalQuantity>0" class="pcIconBarViewCart">
        
            <div id="quickcart">
            				<p class="user-box disp-inline"><a href="/shop/pc/custPref.asp" class="user-action">Existing Customer Login</a></p>
							<p class="user-cart disp-inline"><span class=""><i class="fa fa-shopping-basket"></i> <%'= dictLanguage.Item(Session("language")&"_addedtocart_5") %>{{shoppingcart.totalQuantity}} Item(s) -  <%'= dictLanguage.Item(Session("language")&"_smallcart_2") %>
              <%'=dictLanguage.Item(Session("language")&"_showcart_12")%>
              <span data-ng-show="!Evaluate(shoppingcart.checkoutStage)">{{shoppingcart.daQuickCart}}</span>
              <span data-ng-show="Evaluate(shoppingcart.checkoutStage)">{{shoppingcart.total}}</span>
               | <a href="/shop/pc/viewCart.asp" class="user-action">View Basket</a></p>
            </div>
            
            
        </div>
 
 
    </div>  
</div>
<% End If %>
