<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>   
<% If Instr(Ucase(Request.ServerVariables("SCRIPT_NAME")), "GW") = 0 Then %> 
        
<div id="quickCartContainer" data-ng-controller="QuickCartCtrl">
    <div data-ng-cloak class="ng-cloak">

        <div data-ng-hide="shoppingcart.totalQuantity>0" class="pcIconBarViewCart">
            <%=dictLanguage.Item(Session("language")&"_smallcart_11")%>
        </div>  

        <div data-ng-show="shoppingcart.totalQuantity>0" class="pcIconBarViewCart">
        
            <div id="quickcart">
							<%= dictLanguage.Item(Session("language")&"_addedtocart_5") %>{{shoppingcart.totalQuantity}}<%= dictLanguage.Item(Session("language")&"_smallcart_2") %>
							<%=dictLanguage.Item(Session("language")&"_showcart_12")%>
              <a href="viewcart.asp" data-ng-show="!Evaluate(shoppingcart.checkoutStage)">{{shoppingcart.subtotal}}</a>
              <a href="viewcart.asp" data-ng-show="Evaluate(shoppingcart.checkoutStage)">{{shoppingcart.total}}</a>
            </div>

        </div>

    </div>  
</div>
<% End If %>
