<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<% 
'SB S
pcIsSubscription = findSubscription(Session("pcCartSession"), Session("pcCartIndex"))
If pcIsSubscription then		
	strAndSub = "AND (pcPayTypes_Subscription = 1)"
Else		
	strAndSub = ""		
End if 
'SB E

'SB S
if session("customerType")=1 then
	query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 AND (gwCode=999999 OR gwCode=46 OR gwCode=53 OR gwCode=80 OR gwCode=99) " & strAndSub & " ORDER by paymentPriority;"
else
	query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 and Cbtob=0 AND (gwCode=999999 OR gwCode=46 OR gwCode=53 OR gwCode=80 OR gwCode=99) " & strAndSub & " ORDER by paymentPriority;"
end if
'SB E

hasPayPalButtons = false
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
If NOT rs.eof Then
	hasPayPalButtons = true

	Select Case rs("gwCode")
	'// PayPal Payments Pro (Direct Payments) and PayPal Express Checkout
	Case 46, 999999
		pcStrCheckoutPage = "pcPay_ExpressPay_Start.asp"
	'// PayPal Payments Pro
	Case 53
		pcStrCheckoutPage = "pcPay_ExpressPayPPP_Start.asp"
	'// PayPal Payments Advanced
	Case 80
		pcStrCheckoutPage = "pcPay_ExpressPayPPA_Start.asp"
	'// PayFlow Link
	Case 99
		pcStrCheckoutPage = "pcPay_ExpressPayPFL_Start.asp"
	Case Else
		hasPayPalButtons = false
	End Select

	if hasPayPalButtons then
%>
    
    <div id="pcPayPalButtons" class="pcAltCheckoutButtons">
    	
        <%
			query = "SELECT pcPay_PayPal_Sandbox, pcPay_PayPal_Layout, pcPay_PayPal_Shape, pcPay_PayPal_Size, pcPay_PayPal_Color FROM pcPay_PayPal;"
			set rsP=server.CreateObject("ADODB.RecordSet")
			set rsP=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pcPayPal_Sandbox=rsP("pcPay_PayPal_Sandbox")
			if pcPayPal_Sandbox=1 then
				pcPayPal_Method = "sandbox"
			else
				pcPayPal_Method = "production"
			end if
			pcPayPal_Layout=rsP("pcPay_PayPal_Layout")
			pcPayPal_Shape=rsP("pcPay_PayPal_Shape")
			pcPayPal_Size=rsP("pcPay_PayPal_Size")
			pcPayPal_Color=rsP("pcPay_PayPal_Color")
			set rsP=nothing
		%>
        
        <!-- HTML element inside of which PP button will be populated --> 
        <div id="paypal-button" class="pcAltCheckoutButton"></div>
        
        <!-- Load PayPal JavaScript -->
        <script src="https://www.paypalobjects.com/api/checkout.js"></script>
        
        <script>
			var CREATE_PAYMENT_URL = '<%=pcStrCheckoutPage%>?refer=viewcart.asp';
			paypal.Button.render({
				env: '<%=pcPayPal_Method%>',
				style: {
					layout: '<%=pcPayPal_Layout%>',
					size: '<%=pcPayPal_Size%>',
					shape: '<%=pcPayPal_Shape%>',
					color: '<%=pcPayPal_Color%>'
				},
				funding: {
					<% if scCompanyCountry="US" then %>
					allowed: [ paypal.FUNDING.CARD, paypal.FUNDING.CREDIT ],
					disallowed: [ ]
					<% else %>
					allowed: [ paypal.FUNDING.CARD ],
					disallowed: [ paypal.FUNDING.CREDIT ]
					<% end if %>
				},
				payment: function(resolve) {
					paypal.request.post(CREATE_PAYMENT_URL)
						.then((token) => {
							resolve(token)
						})
				},
				onAuthorize: function(data, actions) {
					return actions.redirect();
				},
				onCancel: function(data, actions) {
					return actions.redirect();
				},
				onError: function(err) {
					// Show an error page here, when an error occurs
				},
				commit: false, // Show “Pay Now” button in PayPal Window
			}, '#paypal-button')
		</script>
    </div>
		<%
	end if
	
End If
set rs=nothing
%>
