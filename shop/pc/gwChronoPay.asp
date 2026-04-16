<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'Gateway File: gwChronopay.asp
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'// See if this is a response back from ChronoPay
Transaction_Id = request("transaction_id")
if Transaction_Id <> "" then
	session("GWAuthCode")=""
	session("GWTransId")=request("transaction_id")
	if session("GWOrderId")="" then
		session("GWOrderId")=request("idOrder")
	end if
	session("GWSessionID")=Session.SessionID 
	'// Payment is received - redirect to gwReturn.asp
    call closeDb()
	response.Redirect("gwReturn.asp?s=true&gw=ChronoPay")
end if

'======================================================================================
'// Set redirect page
'======================================================================================
session("redirectPage")="gwChronoPay.asp"
session("redirectPage2")="https://secure.chronopay.com/index_shop.cgi"

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE	
'======================================================================================
': Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
': End Declare and Retrieve Customer's IP Address	

': Declare URL path to gwSubmit.asp	
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
': End Declare URL path to gwSubmit.asp

': Get Order ID and Set to session
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if
': End Get Order ID

': Get customer and order data from the database for this order	
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<%
': End Get customer and order data


': Reset customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
': End Reset customer session

'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'======================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database

query="SELECT CP_ProdID, CP_Currency, CP_testmode FROM pcPay_Chronopay WHERE CP_id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
CP_ProdID=rs("CP_ProdID")
CP_Currency=rs("CP_Currency")
CP_testmode=rs("CP_testmode")

set rs=nothing
%>
<div id="pcMain">
	<div class="pcMainContent">
				<form method="POST" action="<%=session("redirectPage2")%>" name="payment_form" class="pcForms">
					<input type="hidden" name="product_id" value="<%=CP_ProdID%>">
					<% 'select all products from the ProductsOrdered table to post them into the Chronopay.					
					query="SELECT products.idproduct, products.description FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& session("GWOrderId")
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					tempStrDescription=""
					do until rs.eof
						if(tempStrDescription="") then
							tempStrDescription=tempStrDescription & rs("description")
						else
							tempStrDescription=tempStrDescription & " , " & rs("description")
						end if
						rs.moveNext
					loop 
					set rs=nothing 
					%>

					<input type="hidden" name="product_name" value="<%=tempStrDescription%>">
					<input type="hidden" name="product_price" value="<%=pcBillingTotal%>">
					<input type="hidden" name="language" value="En">
					<input type="hidden" name="f_name" value="<%=pcBillingFirstName%>">
					<input type="hidden" name="s_name" value="<%=pcBillingLastName%>">
					<input type="hidden" name="street" value="<%=pcBillingAddress%>">
					<input type="hidden" name="city" value="<%=pcBillingCity%>">
					<input type="hidden" name="state" value="<%=pcBillingState%>">
					<input type="hidden" name="zip" value="<%=pcBillingPostalCode%>">
					<input type="hidden" name="country" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="phone" value="<%=pcBillingPhone%>">
					<input type="hidden" name="email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="cb_url" value="<%=replace((scStoreURL&"/"&scPcFolder&"/pc/gwreturn.asp"),"//","/")%>">

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                    
                    <% call pcs_showBillingAddress %>

            <div class="pcFormItem"> 
			    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
                <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
            </div>
					
            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
