<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% If request("Approved")= "APPROVED" AND request("OrderID")<>"" then
	session("GWAuthCode")=request("TransRefNumber")
	session("GWTransId")=request("TransRefNumber")
	session("GWTransType")=request("TransactionType")
	session("GWSessionID")=Session.SessionID 
	if session("GWOrderId")="" then
		session("GWOrderId")=request("OrderID")
	end if
		
	response.redirect "gwReturn.asp?s=true&gw=PSI"
End If

'//Set redirect page to the current file name
session("redirectPage")="gwPSI_H.asp"
		
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
query="SELECT Userid,[Mode],psi_post,psi_testmode FROM PSIGate WHERE id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_Userid=rs("Userid")
pcv_TransType=rs("Mode")
pcv_PSI_Post=rs("psi_post")
pcv_PSI_TestMode=rs("psi_testmode")

if pcv_PSI_TestMode="YES" then 
  session("redirectPage2")= "https://devcheckout.psigate.com/HTMLPost/HTMLMessenger"
Else
 session("redirectPage2")= "https://checkout.psigate.com/HTMLPost/HTMLMessenger" '"https://order.psigate.com/psigate.asp"
end if 

set rs=nothing
%>
<div id="pcMain">
	<div class="pcMainContent">
				<form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="StoreKey" value="<%=pcv_Userid%>">
					<INPUT TYPE="hidden" NAME="PaymentType" VALUE="CC">
					<INPUT TYPE="hidden" NAME="CustomerRefNo" VALUE="<%=session("idCustomer")%>">
					<input type="hidden" name="Bcompany" value="<%=pcBillingCompany%>">
					<input type="hidden" name="Bname" value="<%=pcBillingFirstName&" "&pcBillingLastName%>" size="45">
					<input type="hidden" name="Baddress1" value="<%=pcBillingAddress%>">
					<INPUT TYPE="hidden" NAME="Baddress2" VALUE="<%=pcBillingAddress2%>">
					<input type="hidden" name="Bcity" value="<%=pcBillingCity%>">
					<% if pcBillingStateCode = "" then pcBillingStateCode= pcBillingProvince End if %>
					<input type="hidden" name="Bprovince" value="<%=pcBillingStateCode%>">
					<input type="hidden" name="Bcountry" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="Email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="Bpostalcode" value="<%=pcBillingPostalCode%>" size="15">
					<input type="hidden" name="Phone" value="<%=pcBillingPhone%>" size="20">
					<input type="hidden" name="OrderID" value="<%=session("GWOrderID")%>"> 
					<input type="hidden" name="Userid" value="HTML Posting">
					<input type="hidden" name="CardAction" value="<%=pcv_TransType%>">
					<% if pcv_PSI_TestMode="YES" then %>
						<input type="hidden" name="TestResult" value="A"> <% 'test only %>
					<% end if %>
						<input type="hidden" name="items" value="1">
						<input type="hidden" name="ItemID1" value="Online Order">
					<% If scCompanyName="" then %>
						<input type="hidden" name="Description1" value="Shopping Cart"> 
					<%else %>
						<input type="hidden" name="Description1" value="<%=scCompanyName%>"> 
					<% end if %>
					<input type="hidden" name="Price1" value="<%=pcBillingTotal%>">
					<input type="hidden" name="Quantity1" value="1">
					<%
					pcv_ThanksURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwPSI_H.asp"),"//","/")
					pcv_ThanksURL=replace(pcv_ThanksURL,"https:/","https://")
					pcv_ThanksURL=replace(pcv_ThanksURL,"http:/","http://")
					%>
					<input type="hidden" name="Subtotal" value="<%=money(pcBillingTotal)%>">
					<input type="hidden" name="ThanksURL" value="<%=pcv_ThanksURL%>">
					<%
					pcv_SorryURL=replace((scStoreURL&"/"&scPcFolder&"/pc/sorry_psi.asp"),"//","/")
					pcv_SorryURL=replace(pcv_SorryURL,"https:/","https://")
					pcv_SorryURL=replace(pcv_SorryURL,"http:/","http://")
					%>
					<input type="hidden" name="NoThanksURL" value="<%=pcv_SorryURL%>">
					<input type="hidden" name="Sname" value="<%=pcShippingFirstName&" "&pcShippingLastName%>">
					<input type="hidden" name="Saddress1" value="<%=pcShippingAddress%>">
					<input type="hidden" name="Saddress2" value="<%=pcShippingAddress2%>">
					<input type="hidden" name="Scity" value="<%=pcShippingCity%>">
					<% if pcshippingStateCode = "" then pcshippingStateCode= pcShippingProvince End if %>
					<input type="hidden" name="Sprovince" value="<%=pcshippingStateCode%>">
					<input type="hidden" name="Spostalcode" value="<%=pcShippingPostalCode%>">
					<input type="hidden" name="Scountry" value="<%=pcShippingCountryCode%>">
					<input type="hidden" name="Comments" value="none"> 
					<INPUT TYPE="hidden" NAME="ResponseFormat" VALUE="HTML">
					<%'Response.end%>

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                    
                    <% call pcs_showBillingAddress %>

            <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></div>
                <div class="pcFormField"><input type="text" name="CardNumber" value="" autocomplete="off"></div>
            </div>
            
            <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></div>
                <div class="pcFormField"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="CardExpMonth">
								<option value="01">1</option>
								<option value="02">2</option>
								<option value="03">3</option>
								<option value="04">4</option>
								<option value="05">5</option>
								<option value="06">6</option>
								<option value="07">7</option>
								<option value="08">8</option>
								<option value="09">9</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="12">12</option>
							</select>
							<% dtCurYear=Year(date()) %>
							&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%> 
							<select name="CardExpYear">
								<option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,2)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,2)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,2)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,2)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,2)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,2)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,2)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,2)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,2)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,2)%>"><%=dtCurYear+10%></option>
							</select></div>
            </div>
            

					<% 'If x_CVV="1" Then %>
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155"></div>
                </div>
					<% 'End If %>

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
