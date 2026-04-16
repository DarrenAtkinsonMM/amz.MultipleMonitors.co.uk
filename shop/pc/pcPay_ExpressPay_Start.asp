<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="DBsv.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<% response.Buffer = true %>
<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalExp"
	
session("ExpressPayMethod") = ""

'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass

'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************


'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=Session("pcCartIndex")

If session("customerType")=1 Then
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
	'Wholesale minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=205"
	end if
Else
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then
	'Retail minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=206"
	end if
End If
		
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

'///////////////////////////////////////////////////////////////////////////////
'// END: GET DATA FROM DB
'///////////////////////////////////////////////////////////////////////////////



'///////////////////////////////////////////////////////////////////////////////
'// START: GET ORDER DETAILS
'///////////////////////////////////////////////////////////////////////////////
'// Order Total
if session("pcPay_PayPalExp_OrderTotal")="" OR session("pcPay_PayPalExp_OrderTotal")="0" then
	session("pcPay_PayPalExp_OrderTotal")=calculateCartTotal(pcCartArray, ppcCartIndex)
end if
OrderTotal=session("pcPay_PayPalExp_OrderTotal")
if OrderTotal="" then
	OrderTotal=0
end if
OrderTotal=money(OrderTotal)
OrderTotal=pcf_CurrencyField(OrderTotal)

'// Category Discounts Total
CatDiscountTotal=calculateCategoryDiscountTotal(pcCartIndex, pcCartArray)
if CatDiscountTotal="" then
	CatDiscountTotal=0
end if
CatDiscountTotal=money(CatDiscountTotal)
CatDiscountTotal=pcf_CurrencyField(CatDiscountTotal)


'// Currency Code Type
currencyCodeType = pcPay_PayPal_Currency

'// Express URLs
url = objPayPalClass.GetURL()
returnURL	= url & "pcPay_ExpressPay_Start.asp?currencyCodeType=" &  currencyCodeType & "&paymentAmount=" & OrderTotal & "&paymentType=" &PaymentAction 
cancelURL	= url & "viewcart.asp?cmd=_express-checkout"

If (scSSL<>"" AND scSSL<>"0" AND scCompanyLogo<>"") Then
	tempURL=scSslURL &"/"& scPcFolder & "/pc/" & "catalog/" & scCompanyLogo
	tempURL=replace(tempURL,"///","/")
	tempURL=replace(tempURL,"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
	logoURL		= tempURL 
End If

'// Sandbox or Live URL
pcv_PayPal_URL	= objPayPalClass.GetECURL(pcPay_PayPal_Method)
'pcv_PayPal_URL = pcv_PayPal_URL & "&cmd=_express-checkout&token="
pcv_PayPal_URL = pcv_PayPal_URL & "?cmd=_express-checkout&token="  '// Use ? instead of & (both work)


'///////////////////////////////////////////////////////////////////////////////
'// END: GET ORDER DETAILS
'///////////////////////////////////////////////////////////////////////////////


'///////////////////////////////////////////////////////////////////////////////
'// START: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////

'// Set our token
Dim Token
Token=Request.Querystring("TOKEN")
session("PayPalExpressToken")=Token

'// Begin Post If No Token
If  Request.QueryString("token")="" Then
	'***********************************************************************
	'// Start: Posting Details to PayPal
	'***********************************************************************

		
	'// Set whether or not to use BML (Bill-Me-Later)
	BML = false
	if LCase(Request.Querystring("bml")) = "true" then
		BML = true
	end if
	session("PayPalExpressBML") = BML
	
		'---------------------------------------------------------------------------
		' Construct the parameter string that describes the PayPal payment the varialbes 
		' were set in the web form, and the resulting string is stored in nvpstr
		'
		' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
		'---------------------------------------------------------------------------
		nvpstr="" '// clear 
		if lcase(request.QueryString("refer")) = "viewcart.asp" OR lcase(request.QueryString("refer")) = "onepagecheckout.asp" then
		else
			objPayPalClass.AddNVP "ADDROVERRIDE", "1"
		end if
		objPayPalClass.AddNVP "RETURNURL", returnURL
		objPayPalClass.AddNVP "CANCELURL", cancelURL
		objPayPalClass.AddNVP "CURRENCYCODE", currencyCodeType
		objPayPalClass.AddNVP "PAYMENTACTION", PaymentAction
		objPayPalClass.AddNVP "ALLOWNOTE", "0"
		if logoURL<>"" then
			objPayPalClass.AddNVP "HDRIMG", logoURL
		end if

		'// Add BML fields if enabled
		if BML = true then
			objPayPalClass.AddNVP "USERSELECTEDFUNDINGSOURCE", "BML"
			objPayPalClass.AddNVP "SOLUTIONTYPE", "SOLE"
			objPayPalClass.AddNVP "LANDINGPAGE", "Billing"
		end if
		
		'SB S		
		If (pcCartArray(1,38)>0) then
			pSubscriptionID = (pcCartArray(1,38))
			%>
			<!--#include file="../includes/pcSBDataInc.asp" --> 
			<%
			objPayPalClass.AddNVP "L_BILLINGTYPE0", "RecurringPayments"
			objPayPalClass.AddNVP "L_BILLINGAGREEMENTDESCRIPTION0", pcv_strLinkID
		End if 
		'SB E

		if lcase(request.QueryString("refer")) = "viewcart.asp" OR lcase(request.QueryString("refer")) = "onepagecheckout.asp" then
		else
			
			query="SELECT shipmentDetails FROM orders WHERE randomNumber=" &session("pcSFIdDbSession")& " AND idCustomer=" &session("idCustomer")
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			pshipmentDetails=rs("shipmentDetails")
			shipping=split(pshipmentDetails,",")
			if ubound(shipping)>1 then
				if NOT isNumeric(trim(shipping(2))) then
				else
					Postage=trim(shipping(2))
					if ubound(shipping)=>3 then
						serviceHandlingFee=trim(shipping(3))
						if NOT isNumeric(serviceHandlingFee) then
							serviceHandlingFee=0
						end if
					else
						serviceHandlingFee=0
					end if
				end if
			end if
			
			objPayPalClass.AddNVP "SHIPPINGAMT", Postage
			objPayPalClass.AddNVP "HANDLINGAMT", serviceHandlingFee
			
			query="SELECT pcCustSession_ShippingFirstName, pcCustSession_ShippingLastName, pcCustSession_ShippingCompany, pcCustSession_ShippingAddress, pcCustSession_ShippingPostalCode, pcCustSession_ShippingStateCode, pcCustSession_ShippingProvince, pcCustSession_ShippingPhone,  pcCustSession_ShippingCity, pcCustSession_ShippingCountryCode, pcCustSession_ShippingAddress2 FROM pcCustomerSessions WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
	
			pcCustSession_ShippingFirstName = rs("pcCustSession_ShippingFirstName")
			pcCustSession_ShippingLastName = rs("pcCustSession_ShippingLastName")
			pcCustSession_ShippingCompany = rs("pcCustSession_ShippingCompany")
			pcv_strShippingAddress = rs("pcCustSession_ShippingAddress")
			pcv_strShippingPostalCode = rs("pcCustSession_ShippingPostalCode")
			pcv_strShippingStateCode = rs("pcCustSession_ShippingStateCode")
			pcv_strShippingProvince = rs("pcCustSession_ShippingProvince")
			pcv_strShippingPhone = rs("pcCustSession_ShippingPhone")
			pcv_strShippingCity = rs("pcCustSession_ShippingCity")
			pcv_strShippingCountryCode = rs("pcCustSession_ShippingCountryCode")
			pcv_strShippingAddress2  = rs("pcCustSession_ShippingAddress2")
			pcv_strShippingFullName = pcCustSession_ShippingFirstName & " "&pcCustSession_ShippingLastName
			set rs=nothing
			
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
		end if
		objPayPalClass.AddNVP "AMT", ccur(OrderTotal) + ccur(Postage) + Ccur(serviceHandlingFee) - Ccur(CatDiscountTotal)
	
		'--------------------------------------------------------------------------- 
		' Make the call to PayPal to set the Express Checkout token
		' If the API call succeded, then redirect the buyer to PayPal
		' to begin to authorize payment.  If an error occurred, show the
		' resulting errors
		'---------------------------------------------------------------------------
		Set resArray = objPayPalClass.hash_call("SetExpressCheckout",nvpstr)
		Set Session("nvpResArray")=resArray


		ack = UCase(resArray("ACK"))
		
		if err.number <> 0 then			
			'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
			Dim pcv_PayPalErrMessage
			%><!--#include file="../includes/pcPayPalErrors.asp"--><%	
			session("ExpressCheckoutPayment")=""							
		end if

		If instr(ack,"SUCCESS")>0 Then
			'// Redirect to paypal.com here
			token = resArray("TOKEN")
			response.write token
			'payPalURL = pcv_PayPal_URL & token
			'objPayPalClass.ReDirectURL(payPalURL)
		Else 
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
				
		End If
	
	'***********************************************************************
	'// End: Posting Details to PayPal
	'***********************************************************************
Else
	'***********************************************************************
	'// Start: Get Details from PayPal
	'***********************************************************************

	'// Create a Session Flag
	session("ExpressCheckoutPayment")="YES"	
	
	'---------------------------------------------------------------------------
	' At this point, the buyer has completed in authorizing payment
	' at PayPal.  The script will now call PayPal with the details
	' of the authorization, incuding any shipping information of the
	' buyer.  Remember, the authorization is not a completed transaction
	' at this state - the buyer still needs an additional step to finalize
	' the transaction
	'---------------------------------------------------------------------------	
	Session("currencyCodeType") = Request.Querystring("currencyCodeType")
	Session("paymentAmount") = Request.Querystring("paymentAmount")
	Session("PaymentType")= Request.Querystring("PaymentType")
	Session("PayerID")= Request.Querystring("PayerID")

	'---------------------------------------------------------------------------
	' Build a second API request to PayPal, using the token as the
	' ID to get the details on the payment authorization
	'
	' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
	'---------------------------------------------------------------------------
	nvpstr="" '// clear 
	objPayPalClass.AddNVP "TOKEN", session("PayPalExpressToken")

		
	'---------------------------------------------------------------------------
	' Make the API call and store the results in an array.  If the
	' call was a success, show the authorization details, and provide
	' an action to complete the payment.  If failed, show the error
	'---------------------------------------------------------------------------
	Set resArray = objPayPalClass.hash_call("GetExpressCheckoutDetails",nvpstr)
	

	ack = UCase(resArray("ACK"))
	Set Session("nvpResArray")=resArray
	
	'// Successful Get Express Details
	If UCase(ack)="SUCCESS" Then


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Set Express Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
		pcStrShippingPhone=resArray("PHONENUM")
		pcv_Payer=resArray("EMAIL")
		session("Payer")=pcv_Payer
		pcv_PayerID=resArray("PAYERID")
		session("PayerId")=pcv_PayerID
		pcv_PayerStatus=resArray("PAYERSTATUS")
		pcv_PayerBusiness=resArray("BUSINESS")	
		pcv_FirstName=resArray("FIRSTNAME")
		pcv_LastName=resArray("LASTNAME")		
		pcv_FullName= pcv_FirstName & " " & pcv_LastName	
		pcv_ShipToName = resArray("SHIPTONAME")	
		pcv_ShipToBusiness =  resArray("SHIPTOBUSINESS")
		pcv_Street1=resArray("SHIPTOSTREET")
		pcv_Street2=resArray("SHIPTOSTREET2")
		pcv_CityName=resArray("SHIPTOCITY")
		pcv_StateOrProvince=resArray("SHIPTOSTATE")
		pcv_StateCode=resArray("SHIPTOSTATE")
		pcv_Country=resArray("SHIPTOCOUNTRYCODE")        
        If pcv_Country="US" Then
            pcv_StateOrProvince = ""
        End If
		If pcv_Country&""="" Then
			pcv_Country=resArray("COUNTRYCODE")
		End If
		pcv_CountryName= resArray("SHIPTOCOUNTRYNAME")
		pcv_PostalCode= resArray("SHIPTOZIP")

		session("ppec_shipto_Name") = pcv_ShipToName
		session("ppec_shipto_Business") = pcv_ShipToBusiness
		session("ppec_shipto_Street1") = pcv_Street1
		session("ppec_shipto_Street2") = pcv_Street2
		session("ppec_shipto_City") = pcv_CityName
		session("ppec_shipto_StateCode") = pcv_StateCode
		session("ppec_shipto_Province") = pcv_StateOrProvince
		session("ppec_shipto_Country") = pcv_Country
		session("ppec_shipto_PostalCode") = pcv_PostalCode
		session("ppec_shipto_Phone") = pcStrShippingPhone
		session("ppec_shipto_Email") = pcv_Payer
		
		if pcv_Country = "AU" or pcv_Country = "CA" then 
		  
			query="SELECT stateCode,stateName FROM states WHERE pcCountryCode = '"&pcv_Country&"' and (stateName='"&pcv_StateCode&"' Or stateCode='"&pcv_StateCode&"')"		
			set rsStates=server.CreateObject("ADODB.RecordSet")
			set rsStates=conntemp.execute(query)
			if not rsstates.eof then 		
			  pcv_StateCode = rsStates("stateCode")
			  pcv_StateOrProvince = ""
			End if
			set rsStates = nothing
			
		End if 

		strEmail=session("Payer")
        strPassword=""
		pCustomerType = 0
		pIdRefer = 0
		pRecvNews = 0
		pcv_strPhoneQuery = ""
		
		if len(pcv_StateCode)>4 then
			pcv_StateCode="" '// Show Province Field, This is not a valid ISO Code
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Set Express Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		'// Set idPayment session
		query="SELECT payTypes.idPayment FROM payTypes WHERE payTypes.active = - 1 AND payTypes.gwCode = 999999;"
		set rs=conntemp.execute(query)
		if not rs.eof then
			session("pcSFIdPayment")=rs("idPayment")
		else
			query="SELECT payTypes.idPayment FROM payTypes WHERE payTypes.active = - 1 AND payTypes.gwCode = 46;"
			set rs=conntemp.execute(query)
			if not rs.eof then
				session("pcSFIdPayment")=rs("idPayment")
			end if
		end if
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Customer Logged into ProductCart
		if session("idCustomer")<>"" and session("idCustomer")<>0 then
			
			query="UPDATE pcCustomerSessions SET idCustomer="&session("idCustomer")&", pcCustSession_ShippingNickName='"&pcf_SanitizeApostrophe(pcv_ShipToName)&"', pcCustSession_ShippingFirstName='"&pcf_SanitizeApostrophe(pcv_FirstName)&"', pcCustSession_ShippingLastName='"&pcf_SanitizeApostrophe(pcv_LastName)&"', pcCustSession_ShippingCompany='"&pcf_SanitizeApostrophe(pcv_PayerBusiness)&"', pcCustSession_ShippingAddress='"&pcf_SanitizeApostrophe(pcv_Street1)&"', pcCustSession_ShippingPostalCode='"&pcf_SanitizeApostrophe(pcv_PostalCode)&"', pcCustSession_ShippingStateCode='"&pcf_SanitizeApostrophe(pcv_StateCode)&"', pcCustSession_ShippingProvince='"&pcf_SanitizeApostrophe(pcv_StateOrProvince)&"', pcCustSession_ShippingPhone='"&pcf_SanitizeApostrophe(pcStrShippingPhone)&"',  pcCustSession_ShippingCity='"&pcf_SanitizeApostrophe(pcv_CityName)&"', pcCustSession_ShippingCountryCode='"&pcf_SanitizeApostrophe(pcv_Country)&"', pcCustSession_ShippingAddress2='"&pcf_SanitizeApostrophe(pcv_Street2)&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
			set rs=server.CreateObject("ADODB.RecordSet")			
			set rs=conntemp.execute(query)			
			set rs=nothing
			call closedb()								
			response.redirect "OnePageCheckout.asp"
		
		'// Customer NOT Logged into ProductCart
		else

			'// Check if Customer Exists
			query="SELECT idCustomer, pcCust_Guest FROM customers WHERE email='"&strEmail&"';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)				
			
			'// Email Does Not Exist - Create New Customer
			if rs.eof then
			
				pcv_dateCustomerRegistration=Date()
				if SQL_Format="1" then
					pcv_dateCustomerRegistration=Day(pcv_dateCustomerRegistration)&"/"&Month(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
				else
					pcv_dateCustomerRegistration=Month(pcv_dateCustomerRegistration)&"/"&Day(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
				end if
						
				query="INSERT INTO customers (name, lastName, email, [password], city, zip, CountryCode, state, stateCode,shippingcity,shippingZip,shippingCountryCode, shippingState, shippingStateCode, phone, address, shippingAddress, customercompany, customerType, IDRefer, CI1, CI2, address2, shippingCompany, shippingAddress2,RecvNews,pcCust_DateCreated,pcCust_Guest) VALUES ('" &pcf_SanitizeApostrophe(pcv_FirstName)& "', '" &pcf_SanitizeApostrophe(pcv_LastName)& "', '" &pcf_SanitizeApostrophe(strEmail)& "', '" &pcf_SanitizeApostrophe(strPassword)&"','" &pcf_SanitizeApostrophe(pcv_CityName)& "','" &pcf_SanitizeApostrophe(pcv_PostalCode)& "','" &pcf_SanitizeApostrophe(pcv_Country)& "', '"&pcf_SanitizeApostrophe(pcv_StateOrProvince)&"', '" &pcf_SanitizeApostrophe(pcv_StateCode)& "','" &pcf_SanitizeApostrophe(pcv_CityName)& "','" &pcf_SanitizeApostrophe(pcv_PostalCode)& "','" &pcf_SanitizeApostrophe(pcv_Country)& "', '"&pcf_SanitizeApostrophe(pcv_StateOrProvince)&"', '" &pcf_SanitizeApostrophe(pcv_StateCode)& "', '" &pcf_SanitizeApostrophe(pcStrShippingPhone)& "', '" &pcf_SanitizeApostrophe(pcv_Street1)& "', '" &pcf_SanitizeApostrophe(pcv_Street1)& "', '"&pcf_SanitizeApostrophe(pcv_PayerBusiness)&"', " &pCustomerType& ","&pIdRefer&",'" &pcf_SanitizeApostrophe(pCI1)& "','" &pcf_SanitizeApostrophe(pCI2)& "', '" &pcf_SanitizeApostrophe(pcv_Street2)& "','','" &pcf_SanitizeApostrophe(pcv_Street2)& "',"&pRecvNews&",'" & pcf_SanitizeApostrophe(pcv_dateCustomerRegistration) & "',1)"
				set rstemp=server.CreateObject("ADODB.RecordSet")				
				set rstemp=conntemp.execute(query)	
				set rstemp=nothing

				query="SELECT idCustomer, pcCust_Guest FROM customers WHERE email='"&strEmail&"' ORDER BY idCustomer DESC;"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)					
				session("idCustomer")=rstemp("idCustomer")	
				session("CustomerGuest")=rstemp("pcCust_Guest")	
				session("isCustomerNew")="YES"				
				set rstemp=nothing		
			
			'// Email Does Exist - Login Customer
			else
				intIdCustomer=rs("idCustomer")
				intCustomerGuest=rs("pcCust_Guest")	
				session("idCustomer")=intIdCustomer	
				session("CustomerGuest")=intCustomerGuest			
				set rs=nothing
			end if

		end if	
		
		session("PPSA")="0"
		session("PPSAID") = ""
		If session("ppec_shipto_Name")&""<>"" then
			shipToNameArry = split(session("ppec_shipto_Name"), " ")
			shipToFirstNameTmp = shipToNameArry(0)
			shipToLastNameTmp = shipToNameArry(1)
			query="SELECT idRecipient, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Email, recipient_Phone, recipient_Fax, recipient_Company, recipient_Address, recipient_Address2, recipient_City, recipient_State, recipient_StateCode, recipient_Zip, recipient_CountryCode, Recipient_Residential FROM recipients WHERE recipient_NickName='PayPal Shipping Address' AND idCustomer="&session("idCustomer")&";"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)				
			If rs.eof then
				query = "INSERT INTO recipients (idCustomer, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Email, recipient_Phone, recipient_Fax, recipient_Company, recipient_Address, recipient_Address2, recipient_City, recipient_State, recipient_StateCode, recipient_Zip, recipient_CountryCode) VALUES ("&session("idCustomer")&",'PayPal Shipping Address', '"&pcf_SanitizeApostrophe(shipToFirstNameTmp)&"', '"&pcf_SanitizeApostrophe(shipToLastNameTmp)&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_Email"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_Phone"))&"', '', '"&pcf_SanitizeApostrophe(session("ppec_shipto_Business"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_Street1"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_Street2"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_City"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_Province"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_StateCode"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_PostalCode"))&"', '"&pcf_SanitizeApostrophe(session("ppec_shipto_Country"))&"');"
			Else
				query="UPDATE recipients SET recipient_FirstName='"&pcf_SanitizeApostrophe(shipToFirstNameTmp)&"', recipient_LastName='"&pcf_SanitizeApostrophe(shipToLastNameTmp)&"', recipient_Email='"&pcf_SanitizeApostrophe(session("ppec_shipto_Email"))&"', recipient_Phone='"&pcf_SanitizeApostrophe(session("ppec_shipto_Phone"))&"', recipient_Company='"&pcf_SanitizeApostrophe(session("ppec_shipto_Business"))&"', recipient_Address='"&pcf_SanitizeApostrophe(session("ppec_shipto_Street1"))&"', recipient_Address2='"&pcf_SanitizeApostrophe(session("ppec_shipto_Street2"))&"', recipient_City='"&pcf_SanitizeApostrophe(session("ppec_shipto_City"))&"', recipient_State='"&pcf_SanitizeApostrophe(session("ppec_shipto_Province"))&"', recipient_StateCode='"&pcf_SanitizeApostrophe(session("ppec_shipto_StateCode"))&"', recipient_Zip='"&pcf_SanitizeApostrophe(session("ppec_shipto_PostalCode"))&"', recipient_CountryCode='"&pcf_SanitizeApostrophe(session("ppec_shipto_Country"))&"' WHERE recipient_NickName='PayPal Shipping Address' AND idCustomer="&session("idCustomer")&";"
			End If
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			query="SELECT idRecipient FROM recipients WHERE recipient_NickName='PayPal Shipping Address' AND idCustomer="&session("idCustomer")&";"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)				
			if NOT rs.eof then
				session("PPSA")="1"
				session("PPSAID") = rs("idRecipient")
			end if
		End If
		
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Update Customer Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	
		
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Sessions
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		query="UPDATE pcCustomerSessions SET idCustomer="&session("idCustomer")&", pcCustSession_ShippingNickName='"&pcf_SanitizeApostrophe(pcv_ShipToName)&"', pcCustSession_ShippingFirstName='"&pcf_SanitizeApostrophe(pcv_FirstName)&"', pcCustSession_ShippingLastName='"&pcf_SanitizeApostrophe(pcv_LastName)&"', pcCustSession_ShippingCompany='"&pcf_SanitizeApostrophe(pcv_PayerBusiness)&"', pcCustSession_ShippingPhone='"&pcf_SanitizeApostrophe(pcStrShippingPhone)&"',  pcCustSession_ShippingAddress='"&pcf_SanitizeApostrophe(pcv_Street1)&"', pcCustSession_ShippingPostalCode='"&pcf_SanitizeApostrophe(pcv_PostalCode)&"', pcCustSession_ShippingStateCode='"&pcf_SanitizeApostrophe(pcv_StateCode)&"', pcCustSession_ShippingProvince='"&pcf_SanitizeApostrophe(pcv_StateOrProvince)&"', pcCustSession_ShippingCity='"&pcf_SanitizeApostrophe(pcv_CityName)&"', pcCustSession_ShippingCountryCode='"&pcf_SanitizeApostrophe(pcv_Country)&"', pcCustSession_ShippingAddress2='"&pcf_SanitizeApostrophe(pcv_Street2)&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Sessions
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

		set rs=nothing
		call closedb()	


		If session("customerType")=1 Then
			if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
			'Wholesale minimum not met, so customer cannot checkout -> show message
				response.redirect "msg.asp?message=205"
			end if
		Else
			if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then
			'Retail minimum not met, so customer cannot checkout -> show message
				response.redirect "msg.asp?message=206"
			end if
		End If
		
		
		response.redirect "OnePageCheckout.asp"
		
	'// Failed Get Express Details
	Else		
	
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
		
	End If	
	'***********************************************************************
	'// End: Get Details from PayPal
	'***********************************************************************
End If
'///////////////////////////////////////////////////////////////////////////////
'// END: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////

function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function

%>
