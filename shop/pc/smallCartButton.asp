<% If Instr(Ucase(Request.ServerVariables("SCRIPT_NAME")), "GW") = 0 Then %>
<div data-ng-controller="QuickCartCtrl" data-ng-cloak class="ng-cloak">
	<a href="/shop/pc/viewCart.asp" class="cart-btn" aria-label="View basket">
		<i class="fa fa-shopping-basket"></i>
		<span>Basket ({{shoppingcart.totalQuantity || 0}})</span>
	</a>
</div>
<% End If %>
