<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'Gateway File: Worldpay
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->

<% 
'// See if this is a response back from WorldPay
if request("status")="Y" then
	session("GWAuthCode")=getUserInput(request("rawAuthCode"),0)
	session("GWTransId")=getUserInput(request("transId"),0)
	if session("GWOrderId")="" then
		session("GWOrderId")=getUserInput(request("idOrder"),0)
	end if
	session("GWSessionID")=Session.SessionID 
	'// Payment is received - redirect to gwReturn.asp
    call closeDb()
	response.Redirect("gwReturn.asp?s=true&gw=WorldPay")
end if

'//Set redirect page to the current file name
session("redirectPage")="gwwp.asp"
session("redirectPage2")="https://select.worldpay.com/wcc/purchase"
'//secure-test.wp3.rbsworldpay.com

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
		
'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=getUserInput(request("idOrder"),0)
end if

'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT WP_instID, WP_Currency, WP_testmode FROM WorldPay WHERE wp_id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
WP_instID=rs("WP_instID")
WP_Currency=rs("WP_Currency")
WP_testmode=rs("WP_testmode")

set rs=nothing
%>
<div id="pcMain">
	<div class="pcMainContent">
    
				<form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="instId" value="<%=WP_instID%>"> 
					<input type="hidden" name="cartId" value="<%=session("GWOrderId")%>"> 
					<input type="hidden" name="amount" value="<%=pcBillingTotal%>"> 
					<input type="hidden" name="currency" value="<%=WP_Currency%>"> 
					<input type="hidden" name="desc" value="Online Order, ProductCart Store">
					<% if WP_testmode="YES" then %>
						<input type="hidden" name="testMode" value="100"> 
						<input type=hidden name="name" value="AUTHORISED">
					<% else %>
                        <input type="hidden" name="name" value="<%=pcBillingFirstName&" "&pcBillingLastName%>"> 
					<% end if %>
					<input type="hidden" name="address" value="<%=pcBillingAddress%>"> 
					<input type="hidden" name="postcode" value="<%=pcBillingPostalCode%>"> 
					<input type="hidden" name="country" value="<%=pcBillingCountryCode%>"> 
					<input type="hidden" name="tel" value="<%=pcBillingPhone%>"> 
					<input type="hidden" name="email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="city" value="<%=pcBillingCity%>">
					<input type="hidden" name="state" value="<%=pcBillingState%>">    
					<input type="hidden" name="MC_OrderID" value="<%=session("GWOrderId")%>">

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                    
                    
                    <% call pcs_showBillingAddress %>


            <div class="pcFormItem"> 
			    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
                <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
            </div>

            <div class="pcAttention">
                NOTE: When you click on the 'Place Order' button, you will temporarily leave our Web site and will be taken to a secure payment page on the WorldPay Web site. You will be redirected back to our store once the transaction has been processed. We have partnered with WorldPay, a leader in secure Internet payment processing, to ensure that your transactions are processed securely and reliably.</p>
            </div>

            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
