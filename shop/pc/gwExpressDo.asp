<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->

<!--#include file="header_wrapper.asp"-->
<% response.Buffer = true %>
<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalExp"


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


'******************************************************************
'// PayPal Address Override
'// To change this value from the default "no Address Override"
'// you will need to change the variable below to the value of 1.
'//
'// For Example: 
'// pcv_strAddressOverride = 1

'******************************************************************
'// Set to "no Address Override" by Default
pcv_strAddressOverride = 0	
'******************************************************************


'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

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
objPayPalClass.pcs_SetAllVariables()
objPayPalClass.pcs_SetShipAddress(pcGatewayDataIdOrder)

'///////////////////////////////////////////////////////////////////////////////
'// END: GET DATA FROM DB
'///////////////////////////////////////////////////////////////////////////////


'///////////////////////////////////////////////////////////////////////////////
'// START: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////

'// Set our Token
Dim Token
PayerID			= Session("PayerId")
Token			= Session("PayPalExpressToken")
currCodeType	= Session("currencyCodeType")
paymentAmount	= pcBillingTotal '// Session("paymentAmount")
paymentType		= Session("PaymentType")

Session("GWTransType")=pcPay_PayPal_TransType

'***********************************************************************
'// Start: Posting Details to PayPal
'***********************************************************************


'SB S
'// By pass PayPal if the immediate order value is 0 
If pcBillingTotal = 0 AND pcIsSubscription Then 	
		
	session("GWAuthCode")	= "AUTH-ARB" 
	session("GWTransId")	= "0" 
    call closeDb()
	Response.Redirect("gwReturn.asp?s=true&gw=PayPalExp&GWError=1")
	Response.End 
	
Else

	'// Normal Payment, Let Pass

End if 
'SB E

'---------------------------------------------------------------------------
' Construct the parameter string that describes the PayPal payment the varialbes 
' were set in the web form, and the resulting string is stored in nvpstr
'
' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
'---------------------------------------------------------------------------
nvpstr="" '// clear 
objPayPalClass.AddNVP "TOKEN", Token
objPayPalClass.AddNVP "PAYERID", PayerID
if session("PayPalExpressBML") = true then
	objPayPalClass.AddNVP "BUTTONSOURCE", "ProductCart_Cart_BML"
else
	objPayPalClass.AddNVP "BUTTONSOURCE", "ProductCart_Cart_EC_US"
end if

pcPay_PayPal_PaymentPrefix 	= API_PAYMENT_PREFIX
pcPay_PayPal_PaymentIndex 	= 0

objPayPalClass.AddNVP "PAYMENTACTION", PaymentAction
objPayPalClass.AddNVP "AMT", pcf_CurrencyField(money(paymentAmount))
objPayPalClass.AddNVP "CURRENCYCODE", currCodeType
objPayPalClass.AddNVP "INVNUM", session("GWOrderId")

IF pcv_strItemizeOrder = 1 THEN

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Itemized Order
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	%>
	<!--#include file="pcPay_PayPal_Itemize.asp"-->
	<%	
	
	'// Disable line items completely since PayPal Express Checkout doesn't support an Item Total of $0.00.
	If ItemTotal = 0 Then
		If pcv_strFinalShipCharge > 0 Then
			ItemTotal = pcv_strFinalShipCharge
			
			'Get shipment details
			query = "SELECT shipmentDetails FROM orders WHERE idOrder = " & pOID & ";"
			Set rs = connTemp.Execute(query)
			If Not rs.eof Then
				shipmentDetails = rs("shipmentDetails")
				Dim shipmentArr : shipmentArr = Split(shipmentDetails, ",")
				
				If UBound(shipmentArr) > 0  Then
					If Len(shipmentArr(1)) > 0 Then pcv_strSelectedShipping = " (" & shipmentArr(1) & ")"
				End If
			End If
			Set rs = Nothing
			
			' Add a custom line item for shipping
			objPayPalClass.AddNVPLineItem "NAME"&count, dictLanguage.Item(Session("language")&"_PayPal_11") & pcv_strSelectedShipping
			objPayPalClass.AddNVPLineItem "NUMBER"&count, ""
			objPayPalClass.AddNVPLineItem "QTY"&count, 1
			objPayPalClass.AddNVPLineItem "AMT"&count, pcv_strFinalShipCharge
			count = count + 1
			
			pcv_strFinalShipCharge = 0
			pcv_strFinalShipDiscount = 0
		End If
	End If
	
	'// PayPal requires two decimal places with a "." decimal separator.
	pcv_strFinalTotal= pcf_CurrencyField(money(pcv_strFinalTotal))
	pcv_strFinalShipCharge= pcf_CurrencyField(money(pcv_strFinalShipCharge))
	pcv_strFinalShipDiscount= pcf_CurrencyField(money(pcv_strFinalShipDiscount))
	pcv_strFinalServiceCharge= pcf_CurrencyField(money(pcv_strFinalServiceCharge))
	pcv_strFinalTax= pcf_CurrencyField(money(pcv_strFinalTax))
	ItemTotal= pcf_CurrencyField(money(ItemTotal))
		
	objPayPalClass.AddNVP "ITEMAMT", ItemTotal
	objPayPalClass.AddNVP "SHIPPINGAMT", pcv_strFinalShipCharge
	if pcv_strFinalShipDiscount <> 0 Then
		objPayPalClass.AddNVP "SHIPDISCAMT", pcv_strFinalShipDiscount
	End If
	objPayPalClass.AddNVP "HANDLINGAMT", pcv_strFinalServiceCharge
	objPayPalClass.AddNVP "TAXAMT", pcv_strFinalTax
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Itemized Order
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
	
End If	



'***********************************************************************
'// Start: Address Override
'***********************************************************************
If pcv_strAddressOverride = 1 Then

	if pcv_strShippingStateCode="" OR isNULL(pcv_strShippingStateCode)=True then
		pcv_strShippingStateCode=pcv_strShippingProvince
	end if
	if pcv_strShippingStateCode<>"" AND isNULL(pcv_strShippingStateCode)=False then
		objPayPalClass.AddNVP "SHIPTONAME", pcv_strShippingFullName
		objPayPalClass.AddNVP "SHIPTOSTREET", pcv_strShippingAddress
		objPayPalClass.AddNVP "SHIPTOSTREET2", pcv_strShippingAddress2
		objPayPalClass.AddNVP "SHIPTOCITY", pcv_strShippingCity
		objPayPalClass.AddNVP "SHIPTOSTATE", pcv_strShippingStateCode
		objPayPalClass.AddNVP "SHIPTOZIP", pcv_strShippingPostalCode
		objPayPalClass.AddNVP "SHIPTOCOUNTRYCODE", pcv_strShippingCountryCode
		objPayPalClass.AddNVP "SHIPTOPHONENUM", pcv_strShippingPhone
	end if
	
End If
'***********************************************************************
'// End: Address Override
'***********************************************************************	

'--------------------------------------------------------------------------- 
' Make the call to PayPal to set the Express Checkout token
' If the API call succeded, then redirect the buyer to PayPal
' to begin to authorize payment.  If an error occurred, show the
' resulting errors
'---------------------------------------------------------------------------

Set resArray = objPayPalClass.hash_call("DoExpressCheckoutPayment",nvpstr)
Set Session("nvpResArray")=resArray

pcPay_PayPal_PaymentIndex = 0

ack = UCase(resArray("ACK"))

if err.number <> 0 then	
	'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
	Dim pcv_PayPalErrMessage
	%><!--#include file="../includes/pcPayPalErrors.asp"--><%
	session("ExpressCheckoutPayment")=""							
end if

If ack="SUCCESS" Then

	TransactionID=resArray("PAYMENTINFO_"&pcPay_PayPal_PaymentIndex&"_TRANSACTIONID")	
	session("GWTransId")=TransactionID
	
	if session("GWTransId") <> "" then
		
		Dim pTodaysDate
		pTodaysDate=Date()
		if SQL_Format="1" then
			pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
		else
			pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
		end if
		tmpStr="'"& pTodaysDate &"'"
		query="INSERT INTO pcPay_PayPal_Authorize (idOrder, amount, paymentmethod, transtype, authcode, idCustomer, captured, AuthorizedDate, CurrencyCode) VALUES ("&pcGatewayDataIdOrder&", "&pcBillingTotal&", 'PayPalExp', '"&paymentAction&"', '"&session("GWTransId")&"', "&pcIdCustomer&", 0," & tmpStr & ", '"&pcPay_PayPal_Currency&"');"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)				
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		call closedb()
					
		response.redirect "gwReturn.asp?s=true&gw=PayPalExp"			
				
	else			
		
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
		'session("ExpressCheckoutPayment")=""
		
	end if


'// Unsuccessful Express Checkout / Transaction Not Complete
Else	

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
	'session("ExpressCheckoutPayment")=""

End If
'///////////////////////////////////////////////////////////////////////////////
'// END: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////

%>
<div id="pcMain">
	<div class="pcMainContent">
		<div class="pcErrorMessage">
            <%=pcv_PayPalErrMessage%>
        </div>
        <div class="pcFormButtons">
            <a class="pcButton pcButtonBack" href="OnePageCheckout.asp">
                <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
            </a>
        </div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
