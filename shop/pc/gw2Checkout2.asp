<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% '//Check if this is a post-back
CartOrderID = request("cart_order_id")

if CartOrderID<>"" then
	gwTransID=request("order_number")

	session("GWAuthCode")=""
	session("GWTransId")=gwTransID
	if session("GWOrderId")="" then
		session("GWOrderId")=CartOrderID
	end if
	session("GWSessionID")=Session.SessionID 
	Response.redirect "gwReturn.asp?s=true&gw=twoCheckout"
end if

'//Set redirect page to the current file name
session("redirectPage")="gw2Checkout2.asp"

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
	session("GWOrderId")=request("idOrder")
end if

'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer



'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT store_id, v2co, v2co_TestMode FROM twoCheckout Where id_twoCheckout=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
store_id=rs("store_id")
v2co=rs("v2co")
if (IsNull(v2co)) OR (v2co="") OR (v2co="0") then
	v2co="2"
end if
v2co_TestMode=rs("v2co_TestMode")
set rs=nothing

if v2co="1" then
	session("redirectPage2")="https://www.2checkout.com/checkout/purchase"
else
	session("redirectPage2")="https://www.2checkout.com/2co/buyer/purchase"
end if

%>
<div id="pcMain">
	<div class="pcMainContent">
    
        <form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
            <input type="hidden" name="sid" value="<%=store_id%>">
            <input type="hidden" name="cart_order_id" value="<%=session("GWOrderId")%>">
            <input type="hidden" name="email" value="<%=pcCustomerEmail%>">
            <input type="hidden" name="total" value="<%=pcBillingTotal%>">
            <% if v2co_TestMode=1 then %>
                <input type="hidden" name="demo" value="Y">
            <% end if %>
            <%IF v2co="1" THEN%>
                <input type="hidden" name="merchant_order_id" value='<%=session("GWOrderId")%>' />
                <input type='hidden' name='mode' value='2CO' />
                <input type='hidden' name='li_0_name' value='OrderID<%=session("GWOrderId")%>' />
                <input type='hidden' name='li_0_price' value='<%=pcBillingTotal%>' />
            <%ELSE%>
                <input type="hidden" name="twoCheckout" value="twoCheckout">
                <% 'select all products from the ProductsOrdered table to insert them into the 2Checkout db.					
                query="SELECT products.idproduct, products.description, quantity, unitPrice FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& session("GWOrderId")
                set rs=server.CreateObject("ADODB.Recordset")
                set rs=connTemp.execute(query)
                    if err.number<>0 then
                        call LogErrorToDatabase()
                        set rstemp=nothing
                        call closedb()
                        response.redirect "techErr.asp?err="&pcStrCustRefID
                    end if
                IntProdCnt=0
                do until rs.eof
                    tempIntIdProduct=rs("idproduct")
                    tempStrDescription=rs("description")
                    tempIntQuantity=rs("quantity")
                    tempDblUnitPrice=rs("unitPrice")
                    IntProdCnt=IntProdCnt+1
                    %>
                    <input type="hidden" name="c_prod_<%=IntProdCnt%>" value="Product_<%=tempIntIdProduct%>,<%=tempIntQuantity%>">
                    <input type="hidden" name="id_type" value="1">
                    <input type="hidden" name="c_name_<%=IntProdCnt%>" value="<%=tempStrDescription%>">
                    <input type="hidden" name="c_description_<%=IntProdCnt%>" value="<%=tempStrDescription%>">
                    <input type="hidden" name="c_price_<%=IntProdCnt%>" value="<%=tempDblUnitPrice%>">
                    <% rs.moveNext
                loop 
                set rs=nothing
                %>
            <%END IF%>
            <input type="hidden" name="card_holder_name" size=35 value="<%=pcBillingFirstName&" "&pcBillingLastName%>">
            <input type="hidden" name="street_address" value="<%=pcBillingAddress%>"> 
            <input type="hidden" name="city" value="<%= pcBillingCity%>"> 
            <input type="hidden" name="state" value="<%=pcBillingState%>"> 
            <input type="hidden" name="zip" value="<%= pcBillingPostalCode %>">
            <input type="hidden" name="country" value="<%= pcBillingCountryCode %>"> 
            <input type="hidden" name="phone" value="<%= pcBillingPhone %>"> 


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
        
        <% IF v2co="1" THEN %>
            <script src="<%=pcf_getJSPath("https://www.2checkout.com/static/checkout/javascript","direct.min.js")%>"></script>
        <% END IF %>
            
    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
