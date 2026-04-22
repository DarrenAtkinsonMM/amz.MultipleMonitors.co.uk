<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% If Instr(Ucase(Request.ServerVariables("SCRIPT_NAME")), "GW") = 0 Then %>
<div id="quickCartContainer" data-ng-controller="QuickCartCtrl" data-ng-cloak class="ng-cloak">
	<span class="hide-xs"><a href="/shop/pc/custPref.asp"><i class="fa fa-user" style="color:var(--accent);"></i>Existing Customer Login</a></span>
	<span class="show-xs"><a href="/shop/pc/viewCart.asp"><i class="fa fa-shopping-basket"></i>Basket <strong>({{shoppingcart.totalQuantity || 0}})</strong></a></span>
</div>
<% End If %>
