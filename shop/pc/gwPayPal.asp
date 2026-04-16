<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="header_wrapper.asp"-->
<% response.Buffer=true %>
<%
'SB-S
If session("SB_SkipPayment")="1" then
	Session("PayPalExpressToken") = "SubscriptionBridge"
	Response.redirect "gwReturn.asp?s=true&gw=PayPalDP"
End if
'SB-E

'SB S
Dev_Testmode = 1
msg=getUserInput(request.querystring("message"),0)
msg=replace(msg, "&lt;BR&gt;", "<BR>")
msg=replace(msg, "&lt;br&gt;", "<br>")
msg=replace(msg, "&lt;b&gt;", "<b>")
msg=replace(msg, "&lt;/b&gt;", "</b>")
msg=replace(msg, "&lt;/font&gt;", "</font>")
msg=replace(msg, "&lt;a href", "<a href")
msg=replace(msg, "&gt;Back&lt;/a&gt;", ">Back</a>")
msg=replace(msg, "&lt;font", "<font")
msg=replace(msg, "&gt;<b>Error&nbsp;</b>:", "><b>Error&nbsp;</b>:")
msg=replace(msg, "&gt;&lt;img src=", "><img src=")
msg=replace(msg, "&gt;&lt;/a&gt;", "></a>")
msg=replace(msg, "&gt;<b>", "><b>")
msg=replace(msg, "&lt;/a&gt;", "</a>")
msg=replace(msg, "&gt;View Cart", ">View Cart")
msg=replace(msg, "&gt;Continue", ">Continue")
msg=replace(msg, "&lt;u>", "<u>")
msg=replace(msg, "&lt;/u>", "</u>")
msg=replace(msg, "&lt;ul&gt;", "<ul>")
msg=replace(msg, "&lt;/ul&gt;", "</ul>")
msg=replace(msg, "&lt;li&gt;", "<li>")
msg=replace(msg, "&lt;/li&gt;", "</li>")
msg=replace(msg, "&gt;", ">") 
msg=replace(msg, "&lt;", "<") 

Session("PayPalExpressToken") = ""
'SB E
%>
<% 
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalDP"

'******************************************************************
'// PayPal Itemized Order
'// To change this value from the default "non-Itemized Order"
'// you will need to change the variable below to the value of 1.
'//
'// For Example: 
'// pcv_strItemizeOrder = 1

'******************************************************************
'// Set to "non-Itemized Order" by Default
pcv_strItemizeOrder = 0	
'******************************************************************


'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass



'//Set redirect page to the current file name
session("redirectPage")="gwPayPal.asp"

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

'///////////////////////////////////////////////////////////////////////////////
'// START: GET DATA FROM DB
'///////////////////////////////////////////////////////////////////////////////

'// Declare Local Variables at once
'>>> pcPay_PayPal_TransType
'>>> PaymentAction
'>>> pcPay_PayPal_Username
'>>> pcPay_PayPal_Password
'>>> pcPay_PayPal_Sandbox
'>>> pcPay_PayPal_Method
'>>> pcPay_PayPal_Signature
objPayPalClass.pcs_SetAllVariables()
objPayPalClass.pcs_SetShipAddress(pcGatewayDataIdOrder)

'///////////////////////////////////////////////////////////////////////////////
'// END: GET DATA FROM DB
'///////////////////////////////////////////////////////////////////////////////


query="SELECT details, comments, taxAmount, shipmentDetails FROM orders WHERE idOrder="&pcGatewayDataIdOrder
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pOrderDetails = rs("details")
pOrderDetails = replace(pOrderDetails,"Amount: ||"," $")
pOrderComments = rs("comments")
pOrderTax = round(rs("taxAmount"),2)
pshipmentDetails = rs("shipmentDetails")
set rs=nothing

'get shipping details...
shipping=split(pshipmentDetails,",")
if ubound(shipping)>1 then
	if NOT isNumeric(trim(shipping(2))) then
		pshipmentCharge="0"
	else
		pshipmentCharge=trim(shipping(2))
		if ubound(shipping)=>3 then
			serviceHandlingFee=trim(shipping(3))
			if NOT isNumeric(serviceHandlingFee) then
				serviceHandlingFee=0
			end if
		else
			serviceHandlingFee=0
		end if
		pshipmentCharge = round(pshipmentCharge,2) + round(serviceHandlingFee,2)
	end if
else
	pshipmentCharge="0"
end if

IF Request.ServerVariables("Content_Length") > 0 AND request("PaymentSubmitted")="Go" then

	'SB S
	'// By pass PayPal if the immediate order value is 0 
	If pcBillingTotal<0 Then
		pcBillingTotal=0
	End If
	If (pcIsSubscription) AND (pcBillingTotal=0) Then	

		session("reqCardNumber")=getUserInput(request.Form("CardNumber"),16)
		session("reqExpMonth")=getUserInput(request.Form("expMonth"),0)
		session("reqExpYear")=getUserInput(request.Form("expYear"),0)
		session("reqCardType")=getUserInput(request.Form("creditCardType"),0)
		session("reqCVV")=getUserInput(request.Form("CVV"),4)		

		pExpiration=getUserInput(request("expMonth"),0) & "/01/" & getUserInput(request("expYear"),0)				
		
		'// Validates expiration
		if DateDiff("d", Month(Now)&"/01/"&Year(now),pExpiration)<=-1 then
			
            call closeDb()
            Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_6")
            Session("backbuttonURL") = tempURL & "?psslurl="&session("redirectPage")&"&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1"

		end if
		
		'// Validate card
		if not IsCreditCard(session("reqCardNumber"), request.form("creditCardType")) then
		
            call closeDb()
            Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_5")
            Session("backbuttonURL") = tempURL & "?psslurl="&session("redirectPage")&"&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1" 
            
		end if 

        call closeDb()
		Response.Redirect("gwReturn.asp?s=true&gw=PayPalDP&GWError=1")
		Response.End 
		
	Else

		'// Normal Payment, Let Pass
		session("reqCardNumber")=getUserInput(request.Form("CardNumber"),16)
		session("reqExpMonth")=getUserInput(request.Form("expMonth"),0)
		session("reqExpYear")=getUserInput(request.Form("expYear"),0)
		session("reqCardType")=getUserInput(request.Form("creditCardType"),0)
		session("reqCVV")=getUserInput(request.Form("CVV"),4)

	End if 
	'SB E


	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' (2) HANDLE POST BACK FORM DATA  
	'		> Important billing info  
	'	
	CardNumber=request.Form("CardNumber")
	CardNumber=replace(CardNumber, " ", "") 'fixes when people enter spaces in CC number
	expYear=request.Form("expYear")
	expMonth=request.Form("expMonth")
	expYear2=request.Form("expYear2")
	expMonth2=request.Form("expMonth2")
	startYear=request.Form("startYear")
	startMonth=request.Form("startMonth")
	CVV=request.Form("CVV")
	CC_TYPE=request.Form("creditCardType")
	ISSUENUMBER=request.Form("ISSUENUMBER")
	
	' (2a) Check the integrity of the data
	'		> Do we have everything that we need?
	'
	reqFieldsOK = true
	
	' ####  card number  
	If reqFieldsOK Then
		retVal = CardNumber
		if (retVal = "") then
			DeclinedString="Invalid credit card number"
			reqFieldsOK = false
		end if
	End If
	
	' ####  valid card number
	if not IsCreditCard(CardNumber,CC_TYPE) AND (CC_TYPE<>"Solo" AND CC_TYPE<>"Maestro") then
			DeclinedString="You have not entered a valid credit card number"
			reqFieldsOK = false      
	end if
	
	' ####  expiration year 
	If reqFieldsOK Then
		retVal = expYear
		if (retVal = "") then
			DeclinedString="Invalid expiration year"
			reqFieldsOK = false
		end if
	End If
	
	' ####  CVV
	if pcPay_PayPal_CVC=1 then
		If reqFieldsOK Then
			retVal = CVV
			if (retVal = "") then
				DeclinedString="Missing CVV Security Code"
				reqFieldsOK = false
			end if
		End IF
	End If
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	If reqFieldsOK Then  ' start data integrity check conditional submission
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
		'///////////////////////////////////////////////////////////////////////////////
		'// START: Direct Payment Method
		'///////////////////////////////////////////////////////////////////////////////

		
		'***********************************************************************
		'// Start: Posting Details to PayPal
		'***********************************************************************

		'---------------------------------------------------------------------------
		' Construct the parameter string that describes the PayPal payment the varialbes 
		' were set in the web form, and the resulting string is stored in nvpstr
		'
		' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
		'---------------------------------------------------------------------------
		nvpstr="" '// clear 
		objPayPalClass.AddNVP "PAYMENTACTION", PaymentAction
		objPayPalClass.AddNVP "IPADDRESS", pcCustIpAddress
		objPayPalClass.AddNVP "AMT", pcf_CurrencyField(money(pcBillingTotal))		
		objPayPalClass.AddNVP "ACCT", CardNumber
		if CVV<>"" then
			objPayPalClass.AddNVP "CVV2", CVV
		end if			
    objPayPalClass.AddNVP "EMAIL", pcCustomerEmail
		objPayPalClass.AddNVP "FIRSTNAME", pcBillingFirstName
		objPayPalClass.AddNVP "LASTNAME", pcBillingLastName
		objPayPalClass.AddNVP "STREET", pcBillingAddress
		objPayPalClass.AddNVP "CITY", pcBillingCity
		objPayPalClass.AddNVP "STATE", pcBillingState
		objPayPalClass.AddNVP "ZIP", pcBillingPostalCode
		objPayPalClass.AddNVP "COUNTRYCODE", pcBillingCountryCode
		objPayPalClass.AddNVP "CURRENCYCODE", pcPay_PayPal_Currency
		objPayPalClass.AddNVP "BUTTONSOURCE", "ProductCart_Cart_DP_US"
		objPayPalClass.AddNVP "INVNUM", session("GWOrderId")		
		if CC_TYPE="Solo" OR CC_TYPE="Maestro" then
			if startMonth<>"" AND startYear<>"" then
				objPayPalClass.AddNVP "STARTDATE", startMonth & startYear
			end if				
			if ISSUENUMBER<>"" then 
				objPayPalClass.AddNVP "STARTDATE", "022007" '// patch paypal bug
				objPayPalClass.AddNVP "ISSUENUMBER", ISSUENUMBER
			end if		
			objPayPalClass.AddNVP "EXPDATE", expMonth2 & expYear2		
			objPayPalClass.AddNVP "CREDITCARDTYPE", "MasterCard" '// patch paypal bug
		else
			objPayPalClass.AddNVP "CREDITCARDTYPE", CC_TYPE
			objPayPalClass.AddNVP "EXPDATE", expMonth & expYear	
		end if

		'// Check for Discounts that are not compatible with "Itemization"
		query="SELECT orders.discountDetails, orders.pcOrd_CatDiscounts FROM orders WHERE orders.idOrder="&pcGatewayDataIdOrder&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.eof then
			pcv_strDiscountDetails=rs("discountDetails")
			pcv_CatDiscounts=rs("pcOrd_CatDiscounts")						
		end if
		
		set rs=nothing

		if pcv_CatDiscounts>0 or trim(pcv_strDiscountDetails)<>"No discounts applied." then
			pcv_strItemizeOrder = 0
		end if
		
		IF pcv_strItemizeOrder = 1 THEN	

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start: Itemized Order
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
			%>
			<!--#include file="pcPay_PayPal_Itemize.asp"-->
			<%	
			'// PayPal requires two decimal places with a "." decimal separator.
			pcv_strFinalTotal= pcf_CurrencyField(money(pcv_strFinalTotal))
			pcv_strFinalShipCharge= pcf_CurrencyField(money(pcv_strFinalShipCharge))
			pcv_strFinalServiceCharge= pcf_CurrencyField(money(pcv_strFinalServiceCharge))
			pcv_strFinalTax= pcf_CurrencyField(money(pcv_strFinalTax))
			ItemTotal= pcf_CurrencyField(money(ItemTotal))

				
			objPayPalClass.AddNVP "ITEMAMT", ItemTotal
			objPayPalClass.AddNVP "SHIPPINGAMT", pcv_strFinalShipCharge
			objPayPalClass.AddNVP "HANDLINGAMT", pcv_strFinalServiceCharge
			objPayPalClass.AddNVP "TAXAMT", pcv_strFinalTax
			
		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' End: Itemized Order
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
			
		End If	
		
		
		'***********************************************************************
		'// Start: Address Override
		'***********************************************************************	
		if pcv_strShippingStateCode="" OR isNULL(pcv_strShippingStateCode)=True then
			pcv_strShippingStateCode=pcv_strShippingProvince
		end if
		if pcv_strShippingStateCode<>"" AND isNULL(pcv_strShippingStateCode)=False then
			objPayPalClass.AddNVP "SHIPTONAME", pcv_strShippingFullName
			objPayPalClass.AddNVP "SHIPTOSTREET", pcv_strShippingAddress
			objPayPalClass.AddNVP "SHIPTOCITY", pcv_strShippingCity
			objPayPalClass.AddNVP "SHIPTOSTATE", pcv_strShippingStateCode
			objPayPalClass.AddNVP "SHIPTOZIP", pcv_strShippingPostalCode
			objPayPalClass.AddNVP "SHIPTOCOUNTRYCODE", pcv_strShippingCountryCode
			objPayPalClass.AddNVP "SHIPTOSTREET2", pcv_strShippingAddress2
			objPayPalClass.AddNVP "SHIPTOPHONENUM", pcv_strShippingPhone
		end if
		'***********************************************************************
		'// End: Address Override
		'***********************************************************************	

		'--------------------------------------------------------------------------- 
		' Make the call to PayPal to set the Express Checkout token
		' If the API call succeded, then redirect the buyer to PayPal
		' to begin to authorize payment.  If an error occurred, show the
		' resulting errors
		'---------------------------------------------------------------------------
		Set resArray = objPayPalClass.hash_call("DoDirectPayment",nvpstr)
		Set Session("nvpResArray")=resArray
		ack = UCase(resArray("ACK"))
		
		if err.number <> 0 then	
			'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
			Dim pcv_PayPalErrMessage
			%><!--#include file="../includes/pcPayPalErrors.asp"--><%										
		end if		
		
		If instr(ack,"SUCCESS")>0 Then

			session("GWTransId")=resArray("TRANSACTIONID")
			session("AVSCode")=resArray("AVSCODE")
			session("CVV2Code")=resArray("CVV2MATCH")
			session("GWAuthCode")=""
			session("GWTransType")=pcPay_PayPal_TransType
			
			if session("GWTransId") <> "" then			
			
				'// Save info in pcPay_PayPal_Authorize if "Authorization"			
				If PaymentAction="Authorization" Then
					
					Dim pTodaysDate
					pTodaysDate=Date()
					if SQL_Format="1" then
						pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
					else
						pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
					end if
					tmpStr="'"& pTodaysDate &"'"
				
					query="INSERT INTO pcPay_PayPal_Authorize (idOrder, amount, paymentmethod, transtype, authcode, idCustomer, captured, AuthorizedDate, CurrencyCode) VALUES ("&pcGatewayDataIdOrder&", "&pcBillingTotal&", 'PayPalDP', '"&paymentAction&"', '"&session("GWTransId")&"', "&pcIdCustomer&", 0," & tmpStr & ", '"&pcPay_PayPal_Currency&"');"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)				
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
				End If		
				
				'Log successful transaction
				call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 1)
			
        call closedb()
				response.redirect "gwReturn.asp?s=true&gw=PayPalWP"
				
			else			
						
				'Log failed transaction
				call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Start: Error Reporting
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				'// Combine User Friendly Errors from "pcPay_PayPal_Errors.asp"
				'// with Code errors from string "Declined String".
				'// Return a formatted error report as the string "pcv_PayPalErrMessage".
				objPayPalClass.GenerateErrorReport()
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' End: Error Reporting
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
				
			end if
		
		'// Unsuccessful Express Checkout / Transaction Not Complete
		Else	
		
			'Log failed transaction
			call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
        
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start: Error Reporting
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			'// Combine User Friendly Errors from "pcPay_PayPal_Errors.asp"
			'// with Code errors from string "DeclinedString".
			'// Return a formatted error report as the string "pcv_PayPalErrMessage".
			objPayPalClass.GenerateErrorReport()
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' End: Error Reporting
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
		
		End If

	Else '// If reqFieldsOK Then
	
		pcv_PayPalErrMessage = DeclinedString
	
	End If ' end data integrity check

	'*************************************************************************************
	' END
	'*************************************************************************************
end if 
%>
<script type=text/javascript>
	function generateCC(){
		var cc_number = new Array(16);
		var cc_len = 16;
		var start = 0;
		var rand_number = Math.random();
		
		switch(document.PaymentForm.creditCardType.value)
				{
			case "Visa":
				cc_number[start++] = 4;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Discover":
				cc_number[start++] = 6;
				cc_number[start++] = 0;
				cc_number[start++] = 1;
				cc_number[start++] = 1;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "MasterCard":
				cc_number[start++] = 5;
				cc_number[start++] = Math.floor(Math.random() * 5) + 1;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Amex":
				cc_number[start++] = 3;
				cc_number[start++] = Math.round(Math.random()) ? 7 : 4 ;
				cc_len = 15;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Maestro":
				cc_number[start++] = 5;
				cc_len = 16;
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
			case "Solo":
				cc_number[start++] = 6;
				cc_len = 16;
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
				}
				
				for (var i = start; i < (cc_len - 1); i++) {
			cc_number[i] = Math.floor(Math.random() * 10);
				}
		
		var sum = 0;
		for (var j = 0; j < (cc_len - 1); j++) {
			var digit = cc_number[j];
			if ((j & 1) == (cc_len & 1)) digit *= 2;
			if (digit > 9) digit -= 9;
			sum += digit;
		}
		
		var check_digit = new Array(0, 9, 8, 7, 6, 5, 4, 3, 2, 1);
		cc_number[cc_len - 1] = check_digit[sum % 10];
		
		document.PaymentForm.CardNumber.value = "";
		for (var k = 0; k < cc_len; k++) {
			document.PaymentForm.CardNumber.value += cc_number[k];
		}
	}
	function generateCC2(){
		switch(document.PaymentForm.creditCardType.value)
				{
			case "Visa":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Discover":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "MasterCard":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Amex":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Maestro":
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
			case "Solo":
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
				}
	}
</script>
<div id="pcMain">
	<div class="pcMainContent">
		<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
		<input type="hidden" name="PaymentSubmitted" value="Go">
			<div class="pcShowContent">

			<% if pcv_PayPalErrMessage <> "" then %>
				<div class="pcErrorMessage">
					The transaction was not performed for the following reasons: 
					<%=pcv_PayPalErrMessage%>
				</div>
			<% end if %>
			<% if Msg<>"" then %>
			  <div class="pcErrorMessage"><%=Msg%></div>
			<% end if %>
			<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
				<div class="pcSpacer"></div>
				<p><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></p>
			<% end if %>
			<div class="pcSpacer"></div>
			<div class="pcSectionTitle">
				<%=dictLanguage.Item(Session("language")&"_GateWay_1")%>
			</div>
			<div class="pcSpacer"></div>
			<p><%=pcBillingFirstName&" "&pcBillingLastName%></p>
			<p><%=pcBillingAddress%></p>
			<% if pcBillingAddress2<>"" then %>
			  <p><%=pcBillingAddress2%></p>
			<% end if %>
			<p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p>
      <p><a href="onepagecheckout.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p>
			<div class="pcSpacer"></div>
			<div class="pcSectionTitle">
				<%=dictLanguage.Item(Session("language")&"_GateWay_5")%>
			</div>
			<div class="pcSpacer"></div>

      <% '// Card Type %>
      <div class="pcFormItem">
        <div class="pcFormLabel">
          <%=dictLanguage.Item(Session("language")&"_GateWay_12")%>
        </div>
        <div class="pcFormField">
          <% if pcPay_PayPal_Method = "sandbox" then %>
            <select name="creditCardType" onChange="javascript:generateCC(); return false;">
          <% else %>
            <select name="creditCardType" onChange="javascript:generateCC2(); return false;">
          <% end if %>	
				  <% 	
					  cardTypeArray=split(pcPay_PayPal_CardTypes,", ")
						i=ubound(cardTypeArray)
						cardCnt=0
						do until cardCnt=i+1
						  cardVar=cardTypeArray(cardCnt)
							select case cardVar
							  case "V"
								  response.write "<option value=""Visa"" selected>Visa</option>"
									cardCnt=cardCnt+1
								case "M" 
									response.write "<option value=""MasterCard"">MasterCard</option>"
									cardCnt=cardCnt+1
								case "A"
									response.write "<option value=""Amex"">American Express</option>"
						  		cardCnt=cardCnt+1
								case "D"
									response.write "<option value=""Discover"">Discover</option>"
									cardCnt=cardCnt+1
							end select
						loop
          %>
          <% If PaymentAction="Authorization" AND pcPay_PayPal_Currency="GBP" Then %>
            <option value="Maestro" <%if CC_TYPE="Maestro" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_20")%></option>
            <option value="Solo" <%if CC_TYPE="Solo" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_21")%></option>
          <% End If %>
          </select>
        </div>
      </div>

      <% '// Card Number %>
      <div class="pcFormItem">
        <div class="pcFormLabel">
          <%=dictLanguage.Item(Session("language")&"_GateWay_7")%>
        </div>
        <div class="pcFormField">
          <input type="text" name="CardNumber" value="" autocomplete="off">
        </div>
      </div>
        
      <%
      '// Maestro/ Solo Cards
      %>
      <div id="UKOptions" style="display: none">
        
        <% '// Issue Number %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=dictLanguage.Item(Session("language")&"_GateWay_13")%>
          </div>
          <div class="pcFormField">
            <input name="ISSUENUMBER" type="text" id="ISSUENUMBER" value="" size="2" maxlength="2">
          </div>
        </div>

        <% '// Issue Date %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=dictLanguage.Item(Session("language")&"_GateWay_14")%>
          </div>
          <div class="pcFormField">
            <%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
            <select name="startMonth">
              <option value="" selected></option>
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
            <select name="startYear">
              <option value="" selected></option>
              <option value="<%=(dtCurYear-10)%>"><%=dtCurYear-10%></option>
              <option value="<%=(dtCurYear-9)%>"><%=dtCurYear-9%></option>
              <option value="<%=(dtCurYear-8)%>"><%=dtCurYear-8%></option>
              <option value="<%=(dtCurYear-7)%>"><%=dtCurYear-7%></option>
              <option value="<%=(dtCurYear-6)%>"><%=dtCurYear-6%></option>
              <option value="<%=(dtCurYear-5)%>"><%=dtCurYear-5%></option>
              <option value="<%=(dtCurYear-4)%>"><%=dtCurYear-4%></option>
              <option value="<%=(dtCurYear-3)%>"><%=dtCurYear-3%></option>
              <option value="<%=(dtCurYear-2)%>"><%=dtCurYear-2%></option>
              <option value="<%=(dtCurYear-1)%>"><%=dtCurYear-1%></option>											
              <option value="<%=(dtCurYear)%>"><%=dtCurYear%></option>
            </select>
          </div>
        </div>

        <% '// Expiration Date %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=dictLanguage.Item(Session("language")&"_GateWay_8")%>
          </div>
          <div class="pcFormField">
            <%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
            <select name="expMonth2">
              <option value="" selected></option>
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
            <select name="expYear2">
              <option value="" selected></option>
              <option value="<%=(dtCurYear-10)%>"><%=dtCurYear-10%></option>
              <option value="<%=(dtCurYear-9)%>"><%=dtCurYear-9%></option>
              <option value="<%=(dtCurYear-8)%>"><%=dtCurYear-8%></option>
              <option value="<%=(dtCurYear-7)%>"><%=dtCurYear-7%></option>
              <option value="<%=(dtCurYear-6)%>"><%=dtCurYear-6%></option>
              <option value="<%=(dtCurYear-5)%>"><%=dtCurYear-5%></option>
              <option value="<%=(dtCurYear-4)%>"><%=dtCurYear-4%></option>
              <option value="<%=(dtCurYear-3)%>"><%=dtCurYear-3%></option>
              <option value="<%=(dtCurYear-2)%>"><%=dtCurYear-2%></option>
              <option value="<%=(dtCurYear-1)%>"><%=dtCurYear-1%></option>											
              <option value="<%=(dtCurYear)%>"><%=dtCurYear%></option>
            </select>
            <div class="pcSmallText"><%=dictLanguage.Item(Session("language")&"_GateWay_15")%></div>
          </div>
        </div>
      </div>
                                     
      <%
        '// Visa/ MasterCard/ Discover/ AMEX
      %>
      <div id="USOptions">
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=dictLanguage.Item(Session("language")&"_GateWay_8")%>
          </div>
          <div class="pcFormField">
            <%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
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
              <option value="<%=(dtCurYear)%>" selected><%=dtCurYear%></option>
              <option value="<%=(dtCurYear+1)%>"><%=dtCurYear+1%></option>
              <option value="<%=(dtCurYear+2)%>"><%=dtCurYear+2%></option>
              <option value="<%=(dtCurYear+3)%>"><%=dtCurYear+3%></option>
              <option value="<%=(dtCurYear+4)%>"><%=dtCurYear+4%></option>
              <option value="<%=(dtCurYear+5)%>"><%=dtCurYear+5%></option>
              <option value="<%=(dtCurYear+6)%>"><%=dtCurYear+6%></option>
              <option value="<%=(dtCurYear+7)%>"><%=dtCurYear+7%></option>
              <option value="<%=(dtCurYear+8)%>"><%=dtCurYear+8%></option>
              <option value="<%=(dtCurYear+9)%>"><%=dtCurYear+9%></option>
              <option value="<%=(dtCurYear+10)%>"><%=dtCurYear+10%></option>
            </select>
          </div>
        </div>
      </div>
        
      <% if pcPay_PayPal_CVC=1 then %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=dictLanguage.Item(Session("language")&"_GateWay_11")%>
          </div>
          <div class="pcFormField">
            <input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4">
          </div>
        </div>
        
        <div class="pcFormItem">
          <div class="pcFormLabel">
            &nbsp;
          </div>
          <div class="pcFormField">
            <img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155">
          </div>
        </div>
      <% End If %>
        
		  <% 
        'SB S 
        if pcIsSubscription Then
      %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=scSBLang7%>
          </div>
          <div class="pcFormField">
            <%= money((pcBillingTotal + pcBillingSubScriptionTotal))%>
          </div>
        </div>
      <% Else %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=dictLanguage.Item(Session("language")&"_GateWay_4")%>
          </div>
          <div class="pcFormField">
            <%= money(pcBillingTotal)%>
          </div>
        </div>
      <% 
        End if
        'SB E 
      %>
        
      <%
        'SB S
			  If pcIsSubscription Then 
      %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=scSBLang8%>
          </div>
          <div class="pcFormField">
            <!--#include file="inc_sb_widget.asp"-->
          </div>
        </div>
      <% End If %>
			<%'SB E %> 
              
			<%
        'SB S
        If pcIsSubscription AND scSBaymentPageText <>"" Then
      %>
        <div class="pcSpacer"></div>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=scSBLang9%>
          </div>
          <div class="pcFormField">
            <%=scSBaymentPageText%>
          </div>
        </div>
      <% End If %>                                
      
      <% If pcIsSubscription AND pcv_intIsTrial AND scSBPaymentPageTrialText <> "" Then %>
        <div class="pcSpacer"></div>
        
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%=scSBLang10%>
          </div>
          <div class="pcFormField">
            <%=scSBPaymentPageTrialText%>
          </div>
        </div>
      <% 
        End if 
        'SB E
      %>

      <div class="pcFormButtons">
        <!--#include file="inc_gatewayButtons.asp"-->
      </div>
                    
		  <div class="pcSpacer"></div>
		</div>
	</form>
	<script type=text/javascript>
		<% if pcPay_PayPal_Method = "sandbox" then %>
			generateCC();
		<% else %>
			generateCC2();
		<% end if %>	
	</script>
</div>
</div>
<!--#include file="footer_wrapper.asp"-->
<%
'SB S
if not pcIsSubscription then
	session("reqCardNumber")=""
	session("reqExpMonth")=""
	session("reqExpYear")=""
	session("reqCVV")=""
End if
'SB E
					
'*************************************************************************************
' FUNCTIONS
' START
'
'*************************************************************************************
 function IsCreditCard(ByRef anCardNumber, ByRef asCardType)
	Dim lsNumber		' Credit card number stripped of all spaces, dashes, etc.
	Dim lsChar			' an individual character
	Dim lnTotal			' Sum of all calculations
	Dim lnDigit			' A digit found within a credit card number
	Dim lnPosition		' identifies a character position In a String
	Dim lnSum			' Sum of calculations For a specific Set
		
	' Default result is False
	IsCreditCard = False
    			
	' ====
	' Strip all characters that are Not numbers.
	' ====
		
	' Loop through Each character inthe card number submited
	For lnPosition = 1 To Len(anCardNumber)
		' Grab the current character
		lsChar = Mid(anCardNumber, lnPosition, 1)
		' if the character is a number, append it To our new number
		if validNum(lsChar) Then lsNumber = lsNumber & lsChar
		
	Next ' lnPosition
		
	' ====
	' The credit card number must be between 13 and 16 digits.
	' ====
	' if the length of the number is less Then 13 digits, then Exit the routine
	if Len(lsNumber) < 13 Then Exit function
		
	' if the length of the number is more Then 16 digits, then Exit the routine
	if Len(lsNumber) > 16 Then Exit function
    			    			
	' Choose action based on Type of card
	Select Case LCase(asCardType)
		' VISA
		Case "visa", "v", "V"
			' if first digit Not 4, Exit function
			if Not Left(lsNumber, 1) = "4" Then Exit function
		' American Express
		Case "american express", "americanexpress", "american", "ax", "A"
			' if first 2 digits Not 37, Exit function
			if Not Left(lsNumber, 2) = "37" AND Not Left(lsNumber, 2) = "34" Then Exit function
		' Mastercard
		Case "mastercard", "master card", "master", "M"
			' if first digit Not 5, Exit function
			if Not Left(lsNumber, 1) = "5" Then Exit function
		' Discover
		Case "discover", "discovercard", "discover card", "D"
			' if first digit Not 6, Exit function
			if Not Left(lsNumber, 1) = "6" Then Exit function
			
		Case Else
	End Select ' LCase(asCardType)
    			
	' ====
	' if the credit card number is less Then 16 digits add zeros
	' To the beginning to make it 16 digits.
	' ====
	' Continue Loop While the length of the number is less Then 16 digits
	While Not Len(lsNumber) = 16
			
		' Insert 0 To the beginning of the number
		lsNumber = "0" & lsNumber
		
	Wend ' Not Len(lsNumber) = 16
		
	' ====
	' Multiply Each digit of the credit card number by the corresponding digit of
	' the mask, and sum the results together.
	' ====
		
	' Loop through Each digit
	For lnPosition = 1 To 16
    				
		' Parse a digit from a specified position In the number
		lnDigit = Mid(lsNumber, lnPosition, 1)
			
		' Determine if we multiply by:
		'	1 (Even)
		'	2 (Odd)
		' based On the position that we are reading the digit from
		lnMultiplier = 1 + (lnPosition Mod 2)
			
		' Calculate the sum by multiplying the digit and the Multiplier
		lnSum = lnDigit * lnMultiplier
			
		' (Single digits roll over To remain single. We manually have to Do this.)
		' if the Sum is 10 or more, subtract 9
		if lnSum > 9 Then lnSum = lnSum - 9
			
		' Add the sum To the total of all sums
		lnTotal = lnTotal + lnSum
    			
	Next ' lnPosition
		
	' ====
	' Once all the results are summed divide
	' by 10, if there is no remainder Then the credit card number is valid.
	' ====
	IsCreditCard = ((lnTotal Mod 10) = 0)
		
End function ' IsCreditCard

'*************************************************************************************
' FUNCTIONS
' END
'*************************************************************************************
%>
