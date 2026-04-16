<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->

<% 
'//Set redirect page to the current file name
session("redirectPage")="gwFastCharge.asp"

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
query="SELECT pcPay_FAC_ATSID, pcPay_FAC_TransType, pcPay_FAC_CVV FROM pcPay_FastCharge WHERE pcPay_FAC_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_FAC_ATSID=rs("pcPay_FAC_ATSID")
pcPay_FAC_ATSID=enDeCrypt(pcPay_FAC_ATSID, scCrypPass)
pcPay_FAC_TransType=rs("pcPay_FAC_TransType")
pcv_CVV=rs("pcPay_FAC_CVV")

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************

	Set gwObject = Server.CreateObject("ATS.SecurePost")
	if err.number<>0 then
		FAC_success=0
		strErrorInfo="Unable to send payment information, the required COM Object is not installed on this server."
	else
		gwObject.ATSID = pcPay_FAC_ATSID
		gwObject.Amount = Round(pcBillingTotal,2)
		gwObject.CCName = pcBillingFirstName & " " & pcBillingLastName
		gwObject.CCNumber = Request.Form( "CardNumber" )
		gwObject.ExpMonth = Request.Form( "expMonth" )
		gwObject.ExpYear = Request.Form( "expYear" )
		if pcv_CVV="1" then
			gwObject.CVV2 = Request.Form( "CVV" )
		end if
		gwObject.CI_IPAddress = pcCustIpAddress
		gwObject.MerchantOrderNumber = "ORD-" & session("GWOrderID")
		gwObject.CI_CompanyName = pcBillingCompany
		gwObject.CI_BillAddr1 = pcBillingAddress
		gwObject.CI_BillAddr2 = pcBillingAddress2
		gwObject.CI_BillCity = pcBillingCity
		gwObject.CI_BillState = pcBillingState
		gwObject.CI_BillZip = pcBillingPostalCode
		gwObject.CI_BillCountry = pcBillingCountryCode
		gwObject.CI_Phone = pcBillingPhone
		gwObject.CI_Email = pcCustomerEmail
		gwObject.CI_ShipAddr1 = pcShippingAddress
		gwObject.CI_ShipAddr2 = pcShippingAddress2
		gwObject.CI_ShipCity = pcShippingCity
		gwObject.CI_ShipState = pcShippingState
		gwObject.CI_ShipZip = pcShippingPostalCode
		gwObject.CI_ShipCountry = pcShippingCountryCode
		if pcPay_FAC_TransType="1" then
			gwObject.ProcessSale
		else
			gwObject.ProcessAuth
		end if

		response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
	
		FAC_success=0
		
		If gwObject.ResultAccepted Then
			 FAC_success=1
		End if
	
		Dim FAC_rd_successurl, FAC_rd_resultfailurl
	
		If FAC_success=1 Then
			session("GWAuthCode")=gwObject.ResultAuthCode
			session("GWTransId")=gwObject.ResultRefCode
	
			FAC_rd_successurl="gwReturn.asp?s=true&gw=FAC"
			if pcPay_FAC_TransType<>"1" then
				session("GWTransType")="yes"
			end if
		end if

		If (FAC_success <> 1) then
			strErrorInfo=""
			If gwObject.ResultErrorFlag Then
				strErrorInfo="Error: " & gwObject.LastError
			Else
				strErrorInfo="Declined: " & gwObject.ResultAuthCode
			End If
	
			If (strErrorInfo="") Then
				strErrorInfo="There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."
			End if
		End if
	
	End if

	If FAC_success <> 1 Then
		call closeDb()
        Session("message") = strErrorInfo
        Session("backbuttonURL") = tempURL & "?psslurl=gwFastCharge.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        Response.Redirect "msgb.asp?back=1"
	ElseIf FAC_success=1 Then
		call closeDb()
		Response.Redirect FAC_rd_successurl
	End If
	
	'*************************************************************************************
	' END
	'*************************************************************************************
end if 
%>
<div id="pcMain">
	<div class="pcMainContent">
				<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
				<input type="hidden" name="PaymentSubmitted" value="Go">

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
							<select name="expMonth">
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
							<select name="expYear">
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
							</select>
						</div>
					</div>
                    
					<% If pcv_CVV="1" Then %>
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155"></div>
                </div>
					<% End If %>

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
