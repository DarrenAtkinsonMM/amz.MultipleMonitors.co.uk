<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="gwUSAePay_xcenums.inc"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwUSAePayCheck.asp"

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
query="SELECT pcPay_Uep_SourceKey,pcPay_Uep_CheckPending,pcPay_Uep_TestMode FROM pcPay_USAePay WHERE pcPay_Uep_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_Uep_SourceKey=rs("pcPay_Uep_SourceKey")
pcPay_Uep_SourceKey=enDeCrypt(pcPay_Uep_SourceKey, scCrypPass)
pcPay_Uep_CheckPending=rs("pcPay_Uep_CheckPending")
pcPay_Uep_TestMode=rs("pcPay_Uep_TestMode")

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Set XCharge1 = Server.CreateObject("USAePayXChargeCom2.XChargeCom2")
	
	if request.Form("CHECKTYPE")="1" then
		XCharge1.Command = checkcredit
	else
		XCharge1.Command = check
	end if

	XCharge1.Sourcekey = pcPay_Uep_SourceKey
	XCharge1.IP = pcCustIpAddress

	if pcPay_Uep_TestMode="1" then
		XCharge1.Testmode = True
	else
		XCharge1.Testmode = False
	end if

	XCharge1.Routing = Request.Form( "BANKROUTING" )
	XCharge1.Account = Request.Form( "CHECKACCT" )
	XCharge1.SSN = Request.Form( "SSN" )
	XCharge1.DLNum=Request.Form( "DLNUM" )
	XCharge1.DLState=Request.Form( "DLSTATE" )

	XCharge1.Amount = pcBillingTotal
	
	XCharge1.Invoice = "ORD-" & session("GWOrderId")
	XCharge1.Description = "ORDER ID: #" & session("GWOrderId")
	
	XCharge1.TransHolderName = pcBillingFirstName & " " & pcBillingLastName
	XCharge1.Street = pcBillingAddress
	XCharge1.Zip = pcBillingPostalCode
	
	XCharge1.BillFName = pcBillingFirstName
	XCharge1.BillLName = pcBillingLastName
	XCharge1.BillCompany = pcBillingCompany
	XCharge1.BillStreet = pcBillingAddress
	XCharge1.BillStreet2 = pcBillingAddress2
	XCharge1.BillCity = pcBillingCity
	XCharge1.BillState = pcBillingState
	XCharge1.BillZip = pcBillingPostalCode
	XCharge1.BillCountry = pcBillingCountryCode
	XCharge1.BillPhone = pcBillingPhone
	XCharge1.Email = pcCustomerEmail
				
	XCharge1.ShipFName = pcShippingFirstName
	XCharge1.ShipLName = pcShippingLastName
	XCharge1.ShipCompany = pcShippingCompany
	XCharge1.ShipStreet = pcShippingAddress
	XCharge1.ShipStreet2 = pcShippingAddress2
	XCharge1.ShipCity = pcShippingCity
	XCharge1.ShipState = pcShippingState
	XCharge1.ShipZip = pcShippingPostalCode
	XCharge1.ShipCountry = pcShippingCountryCode
	XCharge1.ShipPhone = pcShippingPhone
	XCharge1.Process
	
	response.buffer=true
	response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
				
	uep_success=0
				
	Select Case XCharge1.ResponseStatus
		Case Approved
		uep_success=1
	End Select

	Dim uep_rd_successurl, uep_rd_resultfailurl

	If uep_success=1 Then
		session("GWAuthCode")=XCharge1.ResponseAuthCode
		session("GWTransId")=XCharge1.ResponseReferenceNum

		uep_rd_successurl="gwReturn.asp?s=true&gw=UEP&c=1"
		if pcPay_Uep_CheckPending="1" then
			uep_rd_successurl=uep_rd_successurl
		end if
	end if
				
	If (uep_success <> 1) then
		
		strErrorInfo=""
		
		If XCharge1.ErrorExists = True Then            
			Dim XError
			
			For Each XError In XCharge1.Errors
				strErrorInfo=strErrorInfo&"<br>"
				strErrorInfo=strErrorInfo & "Error code: " & XError.ErrorCode  & " - Error Message: " & XError.ErrorText
			Next
		End If
					
		If (strErrorInfo="") Then
			strErrorInfo="There was a problem completing your order. We apologize for the inconvenience. Please contact customer support to review your order."
		End if
	End if
		
	If uep_success <> 1 Then
        call closeDb()
        Session("message") = strErrorInfo
        Session("backbuttonURL") = tempURL & "?psslurl=gwUSAePayCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"
	ElseIf uep_success=1 Then
		call closeDb()
		Response.Redirect uep_rd_successurl
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
                <div class="pcFormField"><input name="CHECKACCT" type="text" size="35"></div>
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Account Type:</div>
                <div class="pcFormField">
                    <select name="CHECKTYPE">
                        <option value="0">Check</option>
                        <option value="1">Checkcredit</option>
                    </select>
                </div>
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Social Security number:</div>
                <div class="pcFormField"><input name="SSN" type="text" size="20" maxlength="35"></div>
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Drivers License Number:</div>
                <div class="pcFormField"><input name="DLNUM" type="text" size="20" maxlength="35"></div>
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Drivers License Issuing State:</div>
                <div class="pcFormField"><input name="DLSTATE" type="text" size="20" maxlength="35"></div>
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
