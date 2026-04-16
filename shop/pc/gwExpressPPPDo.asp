<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcPayPalPFApiClass.asp"-->

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
objPayPalClass.pcs_SetShipAddress((int(session("GWOrderId"))-scpre))

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
paymentAmount	= pcBillingTotal
paymentType		= Session("PaymentType")

Session("GWTransType")=pcPay_PayPal_TransType

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

objPayPalClass.AddNVP "CURRENCY", pcPay_PayPal_Currency
objPayPalClass.AddNVP "BUTTONSOURCE", "ProductCart_Cart_EC_US"
objPayPalClass.AddNVP "INVNUM", session("GWOrderId")

'// Check for Discounts that are not compatible with "Itemization"
query="SELECT orders.discountDetails, orders.pcOrd_CatDiscounts FROM orders WHERE orders.idOrder="&(int(session("GWOrderId"))-scpre)&";"
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
			objPayPalClass.AddNVP "SHIPTOCOUNTRY", pcv_strShippingCountryCode
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

Set resArray = objPayPalClass.hash_call("DoExpressCheckoutPayment",nvpstr)
Set Session("nvpResArray")=resArray

ack = UCase(resArray("RESULT"))

if err.number <> 0 then	
	'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
	Dim pcv_PayPalErrMessage
	%><!--#include file="../includes/pcPayPalErrors.asp"--><%
	session("ExpressCheckoutPayment")=""							
end if


If ack=0 Then
	TransactionID=resArray("PNREF")	
	session("GWTransId")=TransactionID
	session("AVSCode")=resArray("AVSADDR")
	session("CVV2Code")=resArray("CVV2MATCH")
			
	if session("GWTransId") <> "" then
		
		'// P = PayPal
		paymentMethod = "P"

		'// Save info to PFL table since it's done through the PayFlow SDK
		query="INSERT INTO pcPay_PFL_Authorize (idOrder, orderDate, paySource, amount, paymentmethod, transtype, authcode, captured, gwCode) VALUES ("&pcGatewayDataIdOrder&", '" & Date() & "', 'PayPalExp', "&pcBillingTotal&", '" & paymentMethod & "', '" & PaymentAction & "', '"&session("GWTransId")&"', 0, 53);"
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
		session("ExpressCheckoutPayment")=""
		
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
	session("ExpressCheckoutPayment")=""

End If
'///////////////////////////////////////////////////////////////////////////////
'// END: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////
%>
<div id="pcMain">
	<div class="pcMainContent">
		<div class="pcErrorMessage"><%=pcv_PayPalErrMessage%></div>

    <div class="pcFormButtons">
      <a class="pcButton pcButtonBack" href="OnePageCheckout.asp">
        <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
      </a>
    </div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
