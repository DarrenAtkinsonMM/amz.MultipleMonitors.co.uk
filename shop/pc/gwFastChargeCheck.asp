<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwFastChargeCheck.asp"

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
query="SELECT pcPay_FAC_ATSID, pcPay_FAC_TransType, pcPay_FAC_CheckPending FROM pcPay_FastCharge WHERE pcPay_FAC_Id=1;"
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
pcPay_FAC_CheckPending=rs("pcPay_FAC_CheckPending")

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	on error resume next
	Set gwObject = Server.CreateObject("ATS.SecurePost")
	if err.number<>0 then
		FAC_success=0
		strErrorInfo="Unable to send payment information, the required COM Object is not installed on this server."
	else
		gwObject.ATSID = pcPay_FAC_ATSID
		gwObject.Amount = Round(pcBillingTotal,2)
		gwObject.CKName = FIRSTNAME & " " & LASTNAME
		gwObject.CKRoutingNumber = Request.Form( "BANKROUTING" )
		gwObject.CKAccountNumber = Request.Form( "CHECKACCT" )
		gwObject.CI_SSNum = Request.Form( "SSN" )
		gwObject.CI_DLNum=Request.Form( "DLNUM" )
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
	
		response.buffer=true
		response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
		
		FAC_success=0
		
		If gwObject.ResultAccepted Then
			FAC_success=1
		End if
	
		Dim FAC_rd_successurl, FAC_rd_resultfailurl
	
		If FAC_success=1 Then
			session("GWAuthCode")=gwObject.ResultAuthCode
			session("GWTransId")=gwObject.ResultRefCode
			if (pcPay_FAC_TransType<>"1") or (pcPay_FAC_CheckPending="1") then
				session("GWTransType")="yes"
			end if
	
			FAC_rd_successurl="gwReturn.asp?s=true&gw=FAC"
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
        Session("backbuttonURL") = tempURL & "?psslurl=gwFastChargeCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"
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
                <img src="<%=pcf_getImagePath("images","sampleck.gif")%>" width="390" height="230">
            </div>

            <div class="pcFormItem"> 
                <div class="pcFormLabel">Bank Routing Number:</div>
                <div class="pcFormField">
                    <input name="BANKROUTING" type="text" size="35" maxlength="50">
                </div> 
            </div>  
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Checking Account Number:</div>
                <div class="pcFormField">
                    <input name="CHECKACCT" type="text" size="35" autocomplete="off">
                </div> 
            </div>         
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Social Security number:</div>
                <div class="pcFormField">
                    <input name="SSN" type="text" size="20" maxlength="35">
                </div> 
            </div>          
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Drivers License Number:</div>
                <div class="pcFormField">
                    <input name="DLNUM" type="text" size="20" maxlength="35">
                </div> 
            </div>

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
