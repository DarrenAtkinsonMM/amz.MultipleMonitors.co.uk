<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/validation.asp"-->

<!--#include file="pcStartSession.asp" -->
<!--#include file="DBsv.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->

<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc, pcv_strSelectedOptions
Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal%>
<!--#include file="pcCheckPricingCats.asp"-->
<% 	
dim pCartTotalWeight, howMuch

pcStrPageName = "login.asp"

pcCartArray=Session("pcCartSession")
pcCartIndex=Session("pcCartIndex")
ppcCartIndex=Session("pcCartIndex")

'// Vat Settings
pcv_ShowVatId = false
pcv_isVatIdRequired = false
pcv_ShowSSN = false
pcv_isSSNRequired = false
if pshowVatID="1" then pcv_ShowVatId = true
if pVatIdReq="1" then pcv_isVatIdRequired = true
if pshowSSN="1" then pcv_ShowSSN = true
if pSSNReq="1" then pcv_isSSNRequired = true

'Get required fields from database ----
pcv_isBillingFirstNameRequired = true
pcv_isBillingLastNameRequired = true
pcv_isBillingCompanyRequired = false
pcv_isBillingPhoneRequired = true
pcv_isBillingAddressRequired = true
pcv_isBillingPostalCodeRequired = true
pcv_isBillingStateCodeRequired = true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isBillingStateCodeRequired=pcv_strStateCodeRequired
end if
pcv_isBillingProvinceRequired = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isBillingProvinceRequired=pcv_strProvinceCodeRequired
end if
pcv_isBillingCityRequired = true
pcv_isBillingCountryCodeRequired = true
pcv_isBillingAddress2Required = false
pcv_isBillingFaxRequired = false

'Check for Conditional required fields
if RefNewReg="1" AND ViewRefer="1" then
	pcv_isIDreferRequired = true
else
	pcv_isIDreferRequired = false
end if

if NewsReg="1" AND AllowNews="1" then
	pcv_isCRecvNewsRequired = true
else
	pcv_isCRecvNewsRequired = false
end if

pcLoginMode=request.QueryString("lmode")
if pcLoginMode=2 then
	session("ExpressCheckoutPayment")="" '// registration on site, cancel express login
end if
if pcLoginMode="" then
	pcLoginMode=0
end if
if NOT validNum(pcLoginMode) then
	pcLoginMode=0
end if

'pcLoginMode
'If PayPal Express look for session("ExpressCheckoutPayment")="YES"
'0=redirect to ONE PAGE CHECKOUT
'1=edit profile
'2=log in only then redirect to redirect URL or custpref.asp

'Start Special Customer Fields
session("sf_nc_custfields")=""
session("pcSFCustFieldsExist")=""

query="SELECT pcCField_ID, pcCField_Name, pcCField_FieldType, pcCField_Value, pcCField_Length, pcCField_Maximum, pcCField_Required, pcCField_PricingCategories, pcCField_ShowOnReg, pcCField_ShowOnCheckout,'',pcCField_Description FROM pcCustomerFields ORDER BY pcCField_Order ASC, pcCField_Name ASC;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if not rs.eof then
	session("pcSFCustFieldsExist")="YES"
	session("sf_nc_custfields")=rs.GetRows()
end if

set rs=nothing
	
if session("pcSFCustFieldsExist")="YES" AND Session("idCustomer")<>0 then
	pcArr=session("sf_nc_custfields")
	For k=0 to ubound(pcArr,2)
		pcArr(10,k)=""
		query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & Session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.eof then
			pcArr(10,k)=rs("pcCFV_Value")
		end if
		set rs=nothing
	Next
	session("sf_nc_custfields")=pcArr
end if

'End of Special Customer Fields

'START B. If form is submitted
if len(request.Form("pcBillingFirstName"))>0 then
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))

	'// validate all fields unless this is a PayPal Express Order
	If session("ExpressCheckoutPayment")<>"YES" then
		
		if len(getUserInput(request("pcCustomerPassword"),0))>0 then
			pcs_ValidateTextField	"pcCustomerPassword", pcv_isCustomerPasswordRequired, 100	
			pcs_ValidateTextField	"pcCustomerConfirmPassword", pcv_isCustomerPasswordRequired, 100
		end if
		
		If pcLoginMode=1 then
			pcs_ValidateEmailField	"pcCustomerEmail2", true, 50
		End If

		Session("pcSFpcCustomerEmail") = session("pcSFLoginEmail")
		
		If (pcLoginMode=2) AND (Session("pcSFpcCustomerEmail")="") then
			response.Clear()
			response.redirect "checkout.asp?msgmode=8&cmode=1"
		End if
		
		pcs_ValidateTextField	"pcBillingFirstName", pcv_isBillingFirstNameRequired, 70
		pcs_ValidateTextField	"pcBillingLastName", pcv_isBillingLastNameRequired, 50
		pcs_ValidateTextField	"pcBillingCompany", pcv_isBillingCompanyRequired, 50
		pcs_ValidatePhoneNumber	"pcBillingPhone", pcv_isBillingPhoneRequired, 30
		pcs_ValidateTextField	"pcBillingAddress", pcv_isBillingAddressRequired, 70
		pcs_ValidateTextField	"pcBillingPostalCode", pcv_isBillingPostalCodeRequired, 12
		pcs_ValidateTextField	"pcBillingCity", pcv_isBillingCityRequired, 50
		pcs_ValidateTextField	"pcBillingAddress2", pcv_isBillingAddress2Required, 150	
		pcs_ValidatePhoneNumber	"pcBillingFax", pcv_isBillingFaxRequired, 0	
		pcs_ValidateTextField	"pcBillingCountryCode", pcv_isBillingCountryCodeRequired, 4
		pcs_ValidateTextField	"pcBillingProvince", pcv_isBillingProvinceRequired, 50
		pcs_ValidateTextField	"pcBillingStateCode", pcv_isBillingStateCodeRequired, 4		
		
		'// VATID
		If pcv_ShowVatId = True Then
			pcs_ValidateVATIDField "pcBillingVATID", pcv_isVATIDRequired, getUserInput(request("pcBillingCountryCode"),0)
		End If		
		
		'// SSN
		If pcv_ShowSSN = True Then
			pcs_ValidateSSNField "pcBillingSSN", pcv_isSSNRequired, getUserInput(request("pcBillingCountryCode"),0)
		End If
		
		'// Check the PostalCode Length for United States
		If Session("pcSFpcBillingCountryCode")="US" Then
			if len(Session("pcSFpcBillingPostalCode"))<5 then
				pcv_intErr = pcv_intErr + 1
				pcv_strGenericPageError = dictLanguage.Item(Session("language")&"_opc_js_74")
			end if
		End If
		
		'// Sanitize Postal Code
		Session("pcSFpcBillingPostalCode")=pcf_PostCodes(Session("pcSFpcBillingPostalCode"))
		
		'Start Special Customer Fields
		pcv_ShowFieldRequired = Request("pcv_ShowFieldRequired")
		if session("pcSFCustFieldsExist")="YES" then
			pcArr=session("sf_nc_custfields")
			For k=0 to ubound(pcArr,2)
				if pcArr(6,k)="1" AND pcArr(2,k)=0 AND pcv_ShowFieldRequired="" then '// Required?
					pcs_ValidateTextField "custfield" & pcArr(0,k), true, 0
				else
					if pcArr(6,k)="1" AND pcArr(2,k)=1 AND pcv_ShowFieldRequired="" then '// Required?
						pcs_ValidateTextField "custfield" & pcArr(0,k), true, 0
					else
						pcs_ValidateTextField "custfield" & pcArr(0,k), false, 0
					end if
				end if
			Next
		end if
		'End of Special Customer Fields
		
	End If

	pcs_ValidateTextField	"CRecvNews", false, 0


	if RefNewCheckout="1" AND Session("idCustomer")=0 then 
		pcs_ValidateTextField	"IDrefer", false, 0
	end if
		
	if NOT validNum(Session("pcSFIDrefer")) then
		Session("pcSFIDrefer")=0
	end if	
	if NOT validNum(Session("pcSFCRecvNews")) OR Session("pcSFCRecvNews")<>"1" then
		Session("pcSFCRecvNews")="0"
	end if
	
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////

	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
	End If
	session("isCustomerNew")="" '// Clear New Customer Flag - Set on pcPay_ExpressPay_Start.asp to identify new customers.
	
	'/////////////////////////////////////////////////////
	'// Set Local Variables for Billing
	'/////////////////////////////////////////////////////
	pcIntCRecvNews = Session("pcSFCRecvNews")
	pcIntIDrefer = Session("pcSFIDrefer")		
	
	' email
	pcStrCustomerEmail = Session("pcSFpcCustomerEmail")
	pcStrCustomerEmail2 = Session("pcSFpcCustomerEmail2")
	if pcStrCustomerEmail2="" then
		pcStrCustomerEmail2=pcStrCustomerEmail
	end if	
	pcStrCustomerPassword = Session("pcSFpcCustomerPassword")
	pcStrBillingFirstName = Session("pcSFpcBillingFirstName")
	pcStrBillingLastName = Session("pcSFpcBillingLastName")
	pcStrBillingCompany = Session("pcSFpcBillingCompany")
	pcStrBillingPhone = Session("pcSFpcBillingPhone")
	pcStrBillingFax = Session("pcSFpcBillingFax")
	pcStrBillingAddress = Session("pcSFpcBillingAddress")
	pcStrBillingPostalCode = Session("pcSFpcBillingPostalCode")
	pcStrBillingStateCode = Session("pcSFpcBillingStateCode")
	pcStrBillingProvince = Session("pcSFpcBillingProvince")
	pcStrBillingVATID = Session("pcSFpcBillingVATID")
	pcStrBillingSSN = Session("pcSFpcBillingSSN")	
	if pcStrBillingProvince<>"" then
		pcStrBillingStateCode=""
	end if
	pcStrBillingCity = Session("pcSFpcBillingCity")
	pcStrBillingCountryCode = Session("pcSFpcBillingCountryCode")
	pcStrBillingAddress2 = Session("pcSFpcBillingAddress2")	
	'encrypt password
	pcStrCustomerPassword=pcf_PasswordHash(pcStrCustomerPassword)	
	' PRV41 begin
	If request.form("pcAllowReviewEmails")="1" Then
	   pcAllowReviewEmails = 1
	Else	
	   pcAllowReviewEmails = 0
	End if
	' PRV41 end
	
	if Session("idCustomer")<>0 then
		'existing customer, update customer data with form data
		if session("ExpressCheckoutPayment")<>"YES" then			

			query="UPDATE customers SET customers.name=N'"&pcStrBillingFirstName&"', customers.lastName=N'"&pcStrBillingLastName&"', customers.customerCompany=N'"&pcStrBillingCompany&"', customers.phone='"&pcStrBillingPhone&"', customers.email='"&pcStrCustomerEmail2&"', customers.address=N'"&pcStrBillingAddress&"', customers.zip='"&pcStrBillingPostalCode&"', customers.stateCode='"&pcStrBillingStateCode&"', customers.state=N'"&pcStrBillingProvince&"', customers.city=N'"&pcStrBillingCity&"', customers.countryCode='"&pcStrBillingCountryCode&"', customers.address2=N'"&pcStrBillingAddress2&"'" & tmpStrQuery & ", customers.fax='"&pcStrBillingFax&"', customers.pcCust_VATID='" & pcStrBillingVATID & "', customers.pcCust_SSN='" & pcStrBillingSSN & "'"
			if pcStrCustomerPassword <>"" then
				query=query&", customers.password='"&pcStrCustomerPassword&"'"
			end If
			' PRV41 begin
			query=query&", pcCust_AllowReviewEmails=" & pcAllowReviewEmails & " WHERE idCustomer="&Session("idCustomer")&";"
			' PRV41 end
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			if pcStrCustomerPassword<>"" then
				call pcs_SaveUsedPass(Session("idCustomer"),pcStrCustomerPassword)
				
				call pcs_SendResetPassMail(Session("idCustomer"),"")
			end if
			
		end if
	else
		tmpRandomPass=0
		if session("pcSFPassWordExists")="NOREG" then
			session("pcIntCheckoutWR") = 1
			' generate random passwords:
			function randomNumber(limit)
				randomize
				randomNumber=int(rnd*limit)+2
			end function
			pcStrCustomerPassword=randomNumber(99999999)
			pcStrCustomerPassword=pcf_PasswordHash(pcStrCustomerPassword)
			tmpRandomPass=1
		end if
		If session("ExpressCheckoutPayment")<>"YES" then
			' Create the Customer Registration Date
			pcv_dateCustomerRegistration=Date()
			if SQL_Format="1" then
				pcv_dateCustomerRegistration=Day(pcv_dateCustomerRegistration)&"/"&Month(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
			else
				pcv_dateCustomerRegistration=Month(pcv_dateCustomerRegistration)&"/"&Day(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
			end if
			'if this is a new customer, save to database
							
			query="INSERT INTO customers ([name], lastname, pcCust_VATID, pcCust_SSN, email, [password], customerCompany, phone, address, zip, stateCode, state, city, countryCode, IDRefer, address2, RecvNews, fax, pcCust_DateCreated) VALUES (N'"&pcStrBillingFirstName&"', N'"&pcStrBillingLastName&"', '"&pcStrBillingVATID&"', '"&pcStrBillingSSN&"', '"&pcStrCustomerEmail&"', '"&pcStrCustomerPassword&"',N'"&pcStrBillingCompany&"', '"&pcStrBillingPhone&"',N'"&pcStrBillingAddress&"', '"&pcStrBillingPostalCode&"', '"&pcStrBillingStateCode&"', N'"&pcStrBillingProvince&"', N'"&pcStrBillingCity&"', '"&pcStrBillingCountryCode&"', "&pcIntIDrefer&", N'"&pcStrBillingAddress2&"',"&pcIntCRecvNews&", '"&pcStrBillingFax&"', '"& pcv_dateCustomerRegistration &"');"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			'Get new customer ID from database and set the idCustomer session
			query="SELECT customers.idCustomer, customers.email FROM customers WHERE customers.email='"&pcStrCustomerEmail&"';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			Session("idCustomer")=rs("idCustomer")
			set rs=nothing
			
			If tmpRandomPass=0 then
				call pcs_SaveUsedPass(Session("idCustomer"),pcStrCustomerPassword)
			End if

			'Start Special Customer Fields for New Customers
			if session("pcSFCustFieldsExist")="YES" then
				pcArr=session("sf_nc_custfields")
				For k=0 to ubound(pcArr,2)
					tmp_cf=""
					tmp_cf=request.form("custfield" & pcArr(0,k))
					if not IsNull(tmp_cf) then
						tmp_cf=replace(tmp_cf,"'","''")
					end if
					pcArr(3,k)=tmp_cf
				Next
	
				pcv_IDCustomer=Session("idCustomer")				
			
				For k=0 to ubound(pcArr,2)
					query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
				
					if NOT rs.eof then
						query="UPDATE pcCustomerFieldsValues SET pcCFV_Value=N'" & pcArr(3,k) & "' WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
					else
						query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & pcv_IDCustomer & "," & pcArr(0,k) & ",N'" & pcArr(3,k) & "');"
					end if
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					set rs=nothing
				Next

				session("sf_nc_custfields")=""
			end if
			'End of Special Customer Fields for New Customers
			
			'// Send New Customer Emails
			pcv_strNoticeNewCust="1" '// Send to Admin
			pcv_strNewCustEmail="1" '// Send to Customer
			%> <!--#include file="adminNewCustEmail.asp"--> <%
		end if
	end if 
	
	'/////////////////////////////////////////////////////
	'//TAX Zone Check
	'/////////////////////////////////////////////////////
	%>
	<!--#include file="pcTaxZone.asp"-->
	<%
	'/////////////////////////////////////////////////////
	'//END TAX Zone Check
	'/////////////////////////////////////////////////////
	
	'Start Special Customer Fields for existing Customers
	'>>> If this is a new customer "session("sf_nc_custfields")" will be empty.
	'>>> This only runs if this is an existing customer
	if session("pcSFCustFieldsExist")="YES" AND isArray(session("sf_nc_custfields"))=True then
		pcArr=session("sf_nc_custfields")
		For k=0 to ubound(pcArr,2)
			tmp_cf=""
			tmp_cf=session("pcSFcustfield" & pcArr(0,k))
			if not IsNull(tmp_cf) then
				tmp_cf=replace(tmp_cf,"'","''")
			end if
			pcArr(3,k)=tmp_cf
		Next

		pcv_IDCustomer=Session("idCustomer")

		For k=0 to ubound(pcArr,2)
			query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			if NOT rs.eof then
				query="UPDATE pcCustomerFieldsValues SET pcCFV_Value=N'" & pcArr(3,k) & "' WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
			else
				query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & pcv_IDCustomer & "," & pcArr(0,k) & ",N'" & pcArr(3,k) & "');"
			end if
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			set rs=nothing
		Next

		session("sf_nc_custfields")=""
	end if
	'End of Special Customer Fields for existing Customers
	
	select case pcLoginMode
		Case 1
			response.redirect "CustPref.asp"
		Case 2
			
			SaveCustLogin=1%>
			<!--#include file="inc_SaveShoppingCart.asp"-->
			<%call closedb()
			pcTempRedirect=Session("SFStrRedirectUrl")
			Session("SFStrRedirectUrl")=""
			if pcTempRedirect<>"" then
				response.redirect pcTempRedirect
			else
				response.redirect "CustPref.asp?mode=new"
			end if
		Case Else
			response.redirect "CustPref.asp"
	end select
end if

'B. Capture customer's email address.
pcStrEryPassword=session("pcSFEryPassword")
session("pcSFEryPassword")=""

pcStrPassedEmail=session("pcSFLoginEmail")
	
'B-1. If idCustomer Session exist OR If customer enter in an email address and password, check credentials and log them in.
if session("pcSFPassWordExists")="YES" OR session("idCustomer")<>0 then

    '// If this is a guest, then redirect them elsewhere, unless Auto Login.
    If (Not len(session("pcSFPassNotEnter"))>0) Then
        query="SELECT TOP 1 customers.pcCust_Guest FROM customers WHERE ((customers.email)='" & pcStrPassedEmail & "') ORDER BY pcCust_Guest ASC"
        set rs = connTemp.execute(query)
        If Not rs.Eof Then
            pcv_intGuestStatus = rs("pcCust_Guest")
            If pcv_intGuestStatus = 1 Then
                Set rs = Nothing
                call closedb()
                session("pcSFPassNotEnter") = ""
                response.redirect("checkout.asp?msgmode=7&cmode="&session("pcSFCMode"))
            End If
        End If
    End If
    Set rs = Nothing

	'// DeCrypt Temp enCrypt
	pcStrLoginPassword=Decrypt(pcStrEryPassword, 9286803311968)
	
	if session("idCustomer")=0 then
		tmpResult=pcf_CheckNewPassH("",pcStrPassedEmail)
		if tmpResult="0" then
			set rs = nothing
			call closeDb()
			response.redirect "checkout.asp?cmode=2&fmode=1&new=1"
		end if
		tmpResult=pcf_CheckUnlockUser("",pcStrPassedEmail)
		if tmpResult="1" then
			set rs = nothing
			call closeDb()
			response.redirect "msg.asp?message=313"
		end if
	end if

	' PRV41 begin
	if session("idCustomer")=0 then
		query="SELECT [password],customers.idcustomer, customers.pcCust_Guest, customers.pcCust_VATID, customers.pcCust_SSN, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email, customers.address, customers.zip, customers.stateCode, customers.state, customers.city, customers.countryCode, customers.address2, customers.suspend, customers.idCustomerCategory, customers.customerType, customers.RecvNews, fax, customers.pcCust_Locked, pcCust_AllowReviewEmails FROM customers WHERE (customers.email='"&pcStrPassedEmail&"') AND (pcCust_Guest=0 OR pcCust_Guest=2);"
		tmpNewLogin=1
		tmpResult="false"
	else
		query="SELECT idcustomer, customers.pcCust_Guest, customers.pcCust_VATID, customers.pcCust_SSN, [name], lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, address2, suspend, idCustomerCategory, customerType, RecvNews, fax, customers.pcCust_Locked, pcCust_AllowReviewEmails FROM customers WHERE ((customers.idcustomer)="&session("idCustomer")&") AND (pcCust_Guest=0 OR pcCust_Guest=2);"
		tmpNewLogin=0
		tmpResult="true"
	end If
	' PRV41 end
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if NOT rs.eof then
		if session("idCustomer")=0 then
			if session("pcSFPassNotEnter")="1" then
				if pcStrLoginPassword=rs("password") then
					tmpResult="true"
				else
					tmpResult="false"
				end if
				session("pcSFPassNotEnter")=""
			else
				tmpResult=pcf_CheckPassH(pcStrLoginPassword,rs("password"))
			end if
			if Ucase(""&tmpResult)<>"TRUE" then
				pcIntIdcustomer=rs("idcustomer")
				tmpResult1=pcf_SaveLoginLockFailed(pcIntIdcustomer,"1")
				if tmpResult1>"0" then
					set rs=nothing
					call closedb()
					response.redirect "msg.asp?message=314"
				end if
			end if
		end if
	end if
	
	if (NOT rs.eof) AND (Ucase(""&tmpResult)="TRUE") then
    
    
        '///////////////////////////////////////////////////////////////////////////
        '// START: LOGIN SUCCESS
        '///////////////////////////////////////////////////////////////////////////
    
    
		pcIntIdcustomer=rs("idcustomer")
		session("CustomerGuest")=rs("pcCust_Guest")
		if session("CustomerGuest") = "" OR isNULL(session("CustomerGuest")) then
			session("CustomerGuest") = 0
		end if
		pcStrBillingFirstName=pcf_ResetFormField(Session("pcSFpcBillingFirstName"), rs("name"))
		pcStrBillingLastName=pcf_ResetFormField(Session("pcSFpcBillingLastName"), rs("lastName"))
		pcStrBillingCompany=pcf_ResetFormField(Session("pcSFpcBillingCompany"), rs("customerCompany"))
		pcStrBillingPhone=pcf_ResetFormField(Session("pcSFpcBillingPhone"), rs("phone"))
		pcStrCustomerEmail2=pcf_ResetFormField(Session("pcSFpcCustomerEmail2"), rs("email"))
		pcStrBillingAddress=pcf_ResetFormField(Session("pcSFpcBillingAddress"), rs("address"))
		pcStrBillingPostalCode=pcf_ResetFormField(Session("pcSFpcBillingPostalCode"), rs("zip"))
		pcStrBillingStateCode=pcf_ResetFormField(Session("pcSFpcBillingStateCode"), rs("stateCode"))
		pcStrBillingProvince=pcf_ResetFormField(Session("pcSFpcBillingProvince"), rs("state"))
		pcStrBillingCity=pcf_ResetFormField(Session("pcSFpcBillingCity"), rs("city"))
		pcStrBillingCountryCode=pcf_ResetFormField(Session("pcSFpcBillingCountryCode"), rs("countryCode"))
		pcStrBillingAddress2=pcf_ResetFormField(Session("pcSFpcBillingAddress2"), rs("address2"))
		pcIntSuspend=rs("suspend")
		pcIntIdCustomerCategory=rs("idCustomerCategory")
		pcIntCustomerType=rs("customerType")
		pcIntRecvNews=rs("RecvNews")
		pcStrBillingFax=rs("fax")
		pcStrBillingVATID=pcf_ResetFormField(Session("pcSFpcBillingVATID"), Trim(rs("pcCust_VATID")))
		pcStrBillingSSN=pcf_ResetFormField(Session("pcSFpcBillingSSN"), Trim(rs("pcCust_SSN")))
		pcIntCustomerLocked=rs("pcCust_Locked")
		if IsNull(pcIntCustomerLocked) or pcIntCustomerLocked="" then
			pcIntCustomerLocked=0
		end If
		' PRV41 begin
		If IsNull(rs("pcCust_AllowReviewEmails")) Then
		   pcAllowReviewEmails = 0
		else
		   pcAllowReviewEmails = rs("pcCust_AllowReviewEmails")
		End If
		' PRV41 end
		
		if pcStrPassedEmail = "" then
			pcStrPassedEmail = pcStrCustomerEmail2
		end if
		
		'// Locked Customer: Previous: customerType=3, Current: pcCust_Locked=1
		if pcIntCustomerType="3" OR pcIntCustomerLocked="1" then
			set rs = nothing
			call closeDb()
			response.redirect "msg.asp?message=56"
		end if
		
		' save logged customer in session
		Session("idCustomer")=pcIntIdcustomer
		If pcIntCustomerType="1" then
			session("customerType")=1
		Else
			session("customerType")=0
		End If
		
		if tmpNewLogin=1 then
			tmpNewLogin=0
			tmpResult=pcf_SaveLoginLockFailed(Session("idCustomer"),"0")
		end if
			
		session("customerCategory")=pcIntIdCustomerCategory
		
		set rs=nothing
		session("idCustomer")=pcIntIdcustomer
		
		For t=1 to ppcCartIndex
			pcCartArray(t,18)=0
		Next
		%>
		<!--#include file="pcReCalPricesLogin.asp"-->
		<%

		'//Restore Saved Cart of customer
		tmpGUID=getUserInput(Request.Cookies("SavedCartGUID"),0)
		if (tmpGUID<>"") AND ((session("pcCartIndex")="0") OR (IsNull(session("pcCartIndex")))) then%>
			<!--#include file="inc_RestoreShoppingCart.asp"-->
		<%else
		SaveCustLogin=1%>
		<!--#include file="inc_SaveShoppingCart.asp"-->
		<%end if
		
		'// If Google Checkout or Express Checkout Only
		pcv_strShowCheckoutBtn=pcf_PaymentTypes("")	
		'//  
		if pcv_strShowCheckoutBtn=0 AND session("ExpressCheckoutPayment")<>"YES" AND pcLoginMode=2 then
			if Session("SFStrRedirectUrl")<>"" then
				call closedb()
				response.redirect Session("SFStrRedirectUrl")
			else
				call closedb()
				response.redirect "custPref.asp"
			end if
		elseif pcv_strShowCheckoutBtn=0 AND session("ExpressCheckoutPayment")<>"YES" AND pcLoginMode=0 then
			call closedb()
			response.redirect "viewCart.asp"
		end if
		
		'Start Special Customer Fields
		err.clear
		if session("pcSFCustFieldsExist")="YES" then
			pcArr=session("sf_nc_custfields")
			For k=0 to ubound(pcArr,2)
				pcArr(10,k)=""
				query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & pcIntIdcustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if not rs.eof then
					pcArr(10,k)=rs("pcCFV_Value")
				end if
				set rs=nothing
			Next
			session("sf_nc_custfields")=pcArr
		end if
		'End of Special Customer Fields

        '// Add customer session to pcCustomerSessions
        If session("idCustomer")>"0" And len(session("pcSFIdDbSession"))>0 And len(session("pcSFRandomKey"))>0 Then
        
			query="UPDATE pcCustomerSessions SET pcCustomerSessions.idCustomer="&session("idCustomer")&" WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&"));"

			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing

        End If


		if pcLoginMode=2 then
			
			
			Dim pcTempRedirect
			pcTempRedirect=Session("SFStrRedirectUrl")
			Session("SFStrRedirectUrl")=""
			
			if pcTempRedirect="" then	
				call closedb()
				response.redirect "CustPref.asp"
			else
				call closedb()
				response.redirect pcTempRedirect
			end if
			response.end
		end if
        
        '///////////////////////////////////////////////////////////////////////////
        '// END: LOGIN SUCCESS
        '///////////////////////////////////////////////////////////////////////////
		
	else 
    
        '///////////////////////////////////////////////////////////////////////////
        '// START: LOGIN FAILED 
        '///////////////////////////////////////////////////////////////////////////
        
        '// Clear the sessions
        pcs_ClearAllSessions
        'redirect customer back to login page and let them know that there credentials were invalid.
        if scSecurity=1 AND (scUserLogin=1 AND session("pcSFPassWordExists")="YES") then
            session("AttackCount")=session("AttackCount")+1
            if session("AttackCount")>=scAttackCount then
                session("AttackCount")=0%>
                <!--#include file="../includes/sendAlarmEmail.asp" -->
                <%
                call closedb()
                response.redirect("checkout.asp?msgmode=4&cmode="&session("pcSFCMode"))
            end if
        end if
        call closedb()
        response.redirect("checkout.asp?msgmode=2&cmode="&session("pcSFCMode"))
        
        '///////////////////////////////////////////////////////////////////////////
        '// END: LOGIN FAILED 
        '///////////////////////////////////////////////////////////////////////////
        
	end if

	if pcLoginMode=0 or request("opc")="1" then
		response.Clear
		if request("opc")="1" then
			response.write "OK"
		else
			response.redirect "onepagecheckout.asp"
		end if
		response.end
	end if
else
	'// Reset some local variables for a form error on login mode = 2
	pcStrBillingFirstName=pcf_ResetFormField(Session("pcSFpcBillingFirstName"), pcStrBillingFirstName)
	pcStrBillingLastName=pcf_ResetFormField(Session("pcSFpcBillingLastName"), pcStrBillingLastName)
	pcStrBillingVATID=pcf_ResetFormField(Session("pcSFpcBillingVATID"), pcStrBillingVATID)
	pcStrBillingSSN=pcf_ResetFormField(Session("pcSFpcBillingSSN"), pcStrBillingSSN)
	pcStrBillingCompany=pcf_ResetFormField(Session("pcSFpcBillingCompany"), pcStrBillingCompany)
	pcStrBillingPhone=pcf_ResetFormField(Session("pcSFpcBillingPhone"), pcStrBillingPhone)
	pcStrCustomerEmail2=pcf_ResetFormField(Session("pcSFpcCustomerEmail2"), pcStrCustomerEmail2)
	pcStrBillingAddress=pcf_ResetFormField(Session("pcSFpcBillingAddress"), pcStrBillingAddress)
	pcStrBillingPostalCode=pcf_ResetFormField(Session("pcSFpcBillingPostalCode"), pcStrBillingPostalCode)
	pcStrBillingStateCode=pcf_ResetFormField(Session("pcSFpcBillingStateCode"), pcStrBillingStateCode)
	pcStrBillingProvince=pcf_ResetFormField(Session("pcSFpcBillingProvince"), pcStrBillingProvince)
	pcStrBillingCity=pcf_ResetFormField(Session("pcSFpcBillingCity"), pcStrBillingCity)
	pcStrBillingCountryCode=pcf_ResetFormField(Session("pcSFpcBillingCountryCode"), pcStrBillingCountryCode)
	pcStrBillingAddress2=pcf_ResetFormField(Session("pcSFpcBillingAddress2"), pcStrBillingAddress2)
	pcStrBillingFax=pcf_ResetFormField(Session("pcSFpcBillingFax"), pcStrBillingFax)

	
	'B-1. If customer did not enter in password address, verify that the email address they entered does not already exist in the database.		

	query="SELECT customers.idcustomer,customers.password FROM customers WHERE customers.email='"&pcStrPassedEmail&"';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if NOT rs.eof then
		query="SELECT customers.idcustomer,customers.password FROM customers WHERE customers.email like '"&pcStrPassedEmail&"' AND pcCust_Guest=1;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closedb()
			response.Redirect("checkout.asp?msgmode=7&cmode="&session("pcSFCMode"))
		else
			set rs=nothing
			call closedb()
			response.Redirect("checkout.asp?msgmode=3&cmode="&session("pcSFCMode"))
		end if
	end if
	
	set rs=nothing
	'B-2. End
'B-1. End
end if
'B. End

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Section C - Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script type=text/javascript>"&vbcrlf
	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

If session("ExpressCheckoutPayment")<>"YES" Then	
	
	if session("idCustomer")=0 AND session("pcSFPassWordExists")="NO" then
		pcs_JavaCompare		"pcCustomerPassword", "pcCustomerConfirmPassword", true, dictLanguage.Item(Session("language")&"_NewCust_2")
		pcs_JavaTextField	"pcCustomerPassword", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
		pcs_JavaTextField	"pcCustomerConfirmPassword", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	end if
	
	pcs_JavaTextField	"pcBillingFirstName", pcv_isBillingFirstNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingLastName", pcv_isBillingLastNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingCompany", pcv_isBillingCompanyRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingCountryCode", pcv_isBillingCountryCodeRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingAddress", pcv_isBillingAddressRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingAddress2", pcv_isBillingAddress2Required, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingCity", pcv_isBillingCityRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingPostalCode", pcv_isBillingPostalCodeRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcBillingPhone", pcv_isBillingPhoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	if session("idCustomer")<>0 AND pcv_ShowVatId = True then
		pcs_JavaTextField "pcBillingVATID", pcv_isVatIdRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	end if
	if session("idCustomer")<>0 AND pcv_ShowSSN = True then
		pcs_JavaTextField	"pcBillingSSN", pcv_isSSNRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	end if
	
End If

If session("ExpressCheckoutPayment")="YES" Then	
	
	'// Only offer the password update the first time.
	if session("isCustomerNew")="YES" AND scAllowCheckoutWR=1 then
		pcs_JavaCompare		"pcCustomerPasswordPayPal", "pcCustomerConfirmPasswordPayPal", true, dictLanguage.Item(Session("language")&"_NewCust_2")
	end if

End If

if pcLoginMode=1 then
	pcs_JavaCompare		"pcCustomerPassword", "pcCustomerConfirmPassword", true, dictLanguage.Item(Session("language")&"_NewCust_2")
end if

if RefNewCheckout="1" AND Session("idCustomer")=0 then
	pcs_JavaDropDownList	"IDRefer", pcv_isIDreferRequired, dictLanguage.Item(Session("language")&"_NewCust_4")
end if

'// Start Special Customer Fields
if session("pcSFCustFieldsExist")="YES" then
	pcArr=session("sf_nc_custfields")
	For k=0 to ubound(pcArr,2)						
		pcv_ShowField=0
		if (pcArr(8,k)="1") and (pcLoginMode=2 or pcLoginMode=1) then
			pcv_ShowField=1
		end if
		if (pcv_ShowField=0) AND (pcArr(9,k)="1") AND (pcLoginMode=0 or pcLoginMode=1) then
			pcv_ShowField=1
		end if
		if (pcv_ShowField=1) AND (pcArr(7,k)="1") then
			if session("idCustomer")<>0 then
								
				query="SELECT pcCustFieldsPricingCats.idcustomerCategory FROM pcCustFieldsPricingCats INNER JOIN Customers ON (pcCustFieldsPricingCats.pcCField_ID=" & pcArr(0,k) & " AND pcCustFieldsPricingCats.idCustomerCategory=customers.idCustomerCategory) WHERE customers.idcustomer=" & session("idCustomer")
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)												
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if								
				if NOT rs.eof then
					pcv_ShowField=1
				else
					pcv_ShowField=0
				end if
				set rs=nothing
			else
				pcv_ShowField=0
			end if
			
		end if	
		if pcv_ShowField=1 then
			if pcArr(6,k)="1" AND pcArr(2,k)=0 then 'Required?
				pcs_JavaTextField	"custfield" & pcArr(0,k), true, dictLanguage.Item(Session("language")&"_NewCust_4"), ""
			else
				if pcArr(6,k)="1" AND pcArr(2,k)=1 then
					pcs_JavaCheckedBox "custfield" & pcArr(0,k), true, dictLanguage.Item(Session("language")&"_NewCust_4")
				else
					pcs_JavaTextField	"custfield" & pcArr(0,k), false, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
				end if
			end if
		end if '// pcv_ShowField=1
		Session("pcSF_ShowField"&k)=pcv_ShowField
	Next
end if
'End of Special Customer Fields
	
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf

response.write "</script>"&vbcrlf



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: FORM VALIDATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: FIX STRINGS FOR DISPLAY
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pcStrBillingFirstName = pcf_ReverseGetUserInput(pcStrBillingFirstName)
pcStrBillingLastName = pcf_ReverseGetUserInput(pcStrBillingLastName)
pcStrBillingCompany = pcf_ReverseGetUserInput(pcStrBillingCompany)
pcStrBillingPhone = pcf_ReverseGetUserInput(pcStrBillingPhone)
pcStrBillingFax = pcf_ReverseGetUserInput(pcStrBillingFax)
pcStrBillingAddress = pcf_ReverseGetUserInput(pcStrBillingAddress)
pcStrCustomerEmail = pcf_ReverseGetUserInput(pcStrCustomerEmail)
pcStrCustomerEmail2 = pcf_ReverseGetUserInput(pcStrCustomerEmail2)
pcStrBillingPostalCode = pcf_ReverseGetUserInput(pcStrBillingPostalCode)
pcStrBillingStateCode = pcf_ReverseGetUserInput(pcStrBillingStateCode)
pcStrBillingProvince = pcf_ReverseGetUserInput(pcStrBillingProvince)
pcStrBillingCity = pcf_ReverseGetUserInput(pcStrBillingCity)
pcStrBillingCountryCode = pcf_ReverseGetUserInput(pcStrBillingCountryCode)
pcStrBillingAddress2 = pcf_ReverseGetUserInput(pcStrBillingAddress2)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: FIX STRINGS FOR DISPLAY
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<div id="pcMain" class="container-fluid pcLogin">
	<div class="row">
    
        <% if pcLoginMode=1 then %>
    	    <h1><%= dictLanguage.Item(Session("language")&"_Custmoda_1")%></h1>
        <% end if %>
        
        <% if pcLoginMode=2 then %>
    	    <h1><%= dictLanguage.Item(Session("language")&"_RegCustAcc")%></h1>
        <% end if %>

        <form id="loginform" name="loginform" action="login.asp?lmode=<%=pcLoginMode%>" method="post" onSubmit="return Form1_Validator(this);" class="form">
			
        <%
					msg = Session("message")
					Session("message") = ""

					If msg<>"" Then
						%><div class="pcErrorMessage"><%= msg %></div><% 
					End If 
				%>
      
        <% 'A. If new customer create a password %>
        <%
        if session("idCustomer")=0 then
          if session("pcSFPassWordExists")="NO" then
          %>
                <!--
              <div class="pcSectionTitle">
                <%= dictLanguage.Item(Session("language")&"_order_BB")%>
              </div>
              <div class="pcSpacer"></div>
              -->


            <% 'Customer Password %>
            <div class="form-group">
                <label for="password"><%= dictLanguage.Item(Session("language")&"_order_H")%><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div></label>
                <input type="password" class="form-control" name="pcCustomerPassword" id="pcCustomerPassword" autocomplete="off">
                <% pcs_RequiredImageTagHorizontal "pcCustomerPassword", true %>
            </div>

              
            <% 'Customer Password Confirm %>
            <div class="form-group">
                <label for="pcCustomerConfirmPassword"><%= dictLanguage.Item(Session("language")&"_order_I")%><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div></label>
                <input type="password" class="form-control" name="pcCustomerConfirmPassword" id="pcCustomerConfirmPassword" autocomplete="off">
                <% pcs_RequiredImageTagHorizontal "pcCustomerConfirmPassword", true %>
            </div>
			
			<script>
			var validator0
			$pc(document).ready(function () {
			var validator0 = $pc("#loginform").validate({
			rules: {
				pcCustomerPassword: {
					required: true,
					remote: {
						type: 'POST',
						url: "checkPass.asp",
						data: {
							passtype: "R",
							pass: function () {
								return $pc("#pcCustomerPassword").val();
							}
							},
						dataFilter: function(data) {
							var myjson = JSON.parse(data);
							if(myjson.isError == "true") {
								return "\"" + myjson.errorMessage + "\"";
							} else {
								return true;
							}
						}
					}
				},
				pcCustomerConfirmPassword: {
					required: true,
					equalTo: "#pcCustomerPassword"
				}
			},
			messages: {
				pcCustomerPassword: {
					required: "<%=dictLanguage.Item(Session("language")&"_opc_js_4")%>"
				},
				pcCustomerConfirmPassword: {
					required: "<%=dictLanguage.Item(Session("language")&"_opc_js_47")%>",
					equalTo: "<%=dictLanguage.Item(Session("language")&"_opc_js_48")%>"
				}
			}
			});
			});
			</script>						  

          <% 
          end if
        end if 
      %>
      <% 'A. End create a new password %>
      
      <input type="hidden" name="pcBillingReferenceId" value="0">
      <input type="hidden" name="pcBillingNickName" value="">
      
      <!--
      <div class="pcSectionTitle">
        <%= dictLanguage.Item(Session("language")&"_order_J")%>
      </div>
      -->

        
        <% 'Billing First Name %>
        <div class="form-group">
            <label for="pcBillingFirstName"><%= dictLanguage.Item(Session("language")&"_order_C")%><% If pcv_isBillingFirstNameRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingFirstName" value="<%=pcStrBillingFirstName%>" size="20">
            <% pcs_RequiredImageTagHorizontal "pcBillingFirstName", pcv_isBillingFirstNameRequired %>
        </div>

      
        <% 'Billing Last Name %>
        <div class="form-group">
            <label for="pcBillingLastName"><%= dictLanguage.Item(Session("language")&"_order_D")%><% If pcv_isBillingLastNameRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingLastName" value="<%=pcStrBillingLastName%>" size="20">
            <% pcs_RequiredImageTagHorizontal "pcBillingLastName", pcv_isBillingLastNameRequired %>
        </div>

        
        <% 'Billing Company %>
        <div class="form-group">
            <label for="pcBillingCompany"><%= dictLanguage.Item(Session("language")&"_order_E")%><% If pcv_isBillingCompanyRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingCompany" value="<%=pcStrBillingCompany%>" size="30">
            <% pcs_RequiredImageTagHorizontal "pcBillingCompany", pcv_isBillingCompanyRequired %>
        </div>

        
        <% 'VAT %>
        <% if pcv_ShowVatID = True Then %>            
            <% if session("ErrpcBillingVATID")<>"" then %>           
                <div class="form-group">
                    <img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_27") %>
                </div>
                <% session("ErrpcBillingVATID") = "" %>
            <% end if %>
            <div class="form-group">
                <label for="pcBillingVATID"><%= dictLanguage.Item(Session("language")&"_Custmoda_26")%><% If pcv_isVatIdRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
                <input type="text" class="form-control" name="pcBillingVATID" value="<%=pcStrBillingVATID %>" ID="Text1">
                <% pcs_RequiredImageTagHorizontal "pcBillingVATID", pcv_isVatIdRequired %>
            </div>            
        <% end if %> 
        
        <% if pcv_ShowSSN = True then %>	
            <% if session("ErrpcBillingSSN")<>"" then %>
                <div class="form-group">
                    <img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_25") %>
                </div>
                <% session("ErrpcBillingSSN") = "" %>
            <% end if %>
            <div class="form-group">
                <label for="pcBillingSSN"><%= dictLanguage.Item(Session("language")&"_Custmoda_24")%><% If pcv_isSSNRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
                <input type="text" class="form-control" name="pcBillingSSN" value="<%=pcStrBillingSSN %>" ID="Text2">
                <% pcs_RequiredImageTagHorizontal "pcBillingSSN", pcv_isSSNRequired %>
            </div>           
        <% end if %>
                        
        <%        
        '///////////////////////////////////////////////////////////
        '// START: COUNTRY AND STATE/ PROVINCE CONFIG
        '///////////////////////////////////////////////////////////
        ' 
        ' 1) Place this section ABOVE the Country field
        ' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
        ' 3) Additional Required Info
        
        '// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
        pcv_isStateCodeRequired =  pcv_isBillingStateCodeRequired '// determines if validation is performed (true or false)
        pcv_isProvinceCodeRequired =  pcv_IsBillingProvinceRequired '// determines if validation is performed (true or false)
        pcv_isCountryCodeRequired =  pcv_IsBillingCountryCodeRequired '// determines if validation is performed (true or false)
        
        '// #3 Additional Required Info
        pcv_strTargetForm = "loginform" '// Name of Form
        pcv_strCountryBox = "pcBillingCountryCode" '// Name of Country Dropdown
        pcv_strTargetBox = "pcBillingStateCode" '// Name of State Dropdown
        pcv_strProvinceBox =  "pcBillingProvince" '// Name of Province Field
        
        '// Set local Country to Session
        if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
          Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrBillingCountryCode
        end if
        
        '// Set local State to Session
        if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
          Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrBillingStateCode
        end if
        
        '// Set local Province to Session
        if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
          Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  pcStrBillingProvince
        end if
        %>					
        <!--#include file="../includes/javascripts/opc_pcStateAndProvince.asp"-->
        <%
        '///////////////////////////////////////////////////////////
        '// END: COUNTRY AND STATE/ PROVINCE CONFIG
        '///////////////////////////////////////////////////////////
        %>
  
        <%
        '// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince5.asp)
        pcs_CountryDropdown
        %>		
        
        <% 'Billing Address %>
        <div class="form-group">
            <label for="pcBillingAddress"><%= dictLanguage.Item(Session("language")&"_order_K")%><% If pcv_isBillingAddressRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingAddress" value="<%=pcStrBillingAddress%>" size="30">
            <% pcs_RequiredImageTagHorizontal "pcBillingAddress", pcv_isBillingAddressRequired %>
        </div>  

        
        <% 'Billing Address 2 %>
        <div class="form-group">
            <label for="pcBillingAddress2"><%= dictLanguage.Item(Session("language")&"_opc_14")%><% If pcv_isBillingAddress2Required Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingAddress2" value="<%=pcStrBillingAddress2%>" size="30">
            <% pcs_RequiredImageTagHorizontal "pcBillingAddress2", pcv_isBillingAddress2Required %>
        </div>
        
        
        <% 'Billing City %>
        <div class="form-group">
            <label for="pcBillingCity"><%= dictLanguage.Item(Session("language")&"_order_L")%><% If pcv_isBillingCityRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingCity" value="<%=pcStrBillingCity%>" size="30">
            <% pcs_RequiredImageTagHorizontal "pcBillingCity", pcv_isBillingCityRequired %>
        </div>

                    
        <%
          '// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince5.asp)
          pcs_StateProvince
        %>
      
        <% 'Billing Postal Code %>
        <div class="form-group">
            <label for="pcBillingPostalCode"><%= dictLanguage.Item(Session("language")&"_order_O")%><% If pcv_isBillingPostalCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingPostalCode" value="<%=pcStrBillingPostalCode%>" size="10">
            <% pcs_RequiredImageTagHorizontal "pcBillingPostalCode", pcv_isBillingPostalCodeRequired %>
        </div>

            
        <%
        '// Phone Custom Error
        if session("ErrpcBillingPhone")<>"" then 
            %>
            <div class="form-group">
                <img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
            </div>
            <% 
            session("ErrpcBillingPhone") = ""
        end if
        %>
              
        <% 'Billing Phone %>
        <div class="form-group">
            <label for="pcBillingPhone"><%= dictLanguage.Item(Session("language")&"_order_F")%><% If pcv_isBillingPhoneRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingPhone" value="<%=pcStrBillingPhone%>" size="15">
            <% pcs_RequiredImageTagHorizontal "pcBillingPhone", pcv_isBillingPhoneRequired %>
        </div>

        
        <% 'Billing Fax %>
        <div class="form-group">
            <label for="pcBillingFax"><%= dictLanguage.Item(Session("language")&"_order_AA")%><% If pcv_isBillingFaxRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcBillingFax" value="<%=pcStrBillingFax%>" size="15">
            <% pcs_RequiredImageTagHorizontal "pcBillingFax", pcv_isBillingFaxRequired %>
        </div>

        
        <% if pcLoginMode=1 then %>
        
          <% '// Email Custom Error
            if session("ErrpcCustomerEmail2")<>"" then %>
            <div class="form-group">
                <img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
            </div>
              <% session("ErrpcCustomerEmail2") = ""
            end if
          %>
          
        <% 'Billing Email %>
        <div class="form-group">
            <label for="pcCustomerEmail2"><%= dictLanguage.Item(Session("language")&"_order_G")%><% If true Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            <input type="text" class="form-control" name="pcCustomerEmail2" value="<%=pcStrCustomerEmail2%>" size="40">
            <% pcs_RequiredImageTagHorizontal "pcCustomerEmail2", true %>
        </div>
          
        <% 'Billing Password %>
        <div class="form-group">
            <label for="pcCustomerEmail2"><%= dictLanguage.Item(Session("language")&"_order_H")%></label>
            <input type="password" class="form-control" id="pcCustomerPassword" name="pcCustomerPassword" autocomplete="off">
        </div>

          
        <% 'Billing Password Confirmation %>
        <div class="form-group">
            <label for="pcCustomerEmail2"><%= dictLanguage.Item(Session("language")&"_order_I")%></label>
            <input type="password" class="form-control" id="pcCustomerConfirmPassword" name="pcCustomerConfirmPassword" autocomplete="off">
        </div>
		
		<script>
			var validator0
			$pc(document).ready(function () {
			var validator0 = $pc("#loginform").validate({
			rules: {
				pcCustomerPassword: {
					required: false,
					remote: {
						type: 'POST',
						url: "checkPass.asp",
						data: {
							passtype: "Rs",
							pass: function () {
								return $pc("#pcCustomerPassword").val();
							}
							},
						dataFilter: function(data) {
							var myjson = JSON.parse(data);
							if(myjson.isError == "true") {
								return "\"" + myjson.errorMessage + "\"";
							} else {
								return true;
							}
						}
					}
				},
				pcCustomerConfirmPassword: {
					required: false,
					equalTo: "#pcCustomerPassword"
				}
			},
			messages: {
				pcCustomerPassword: {
					required: "<%=dictLanguage.Item(Session("language")&"_opc_js_4")%>"
				},
				pcCustomerConfirmPassword: {
					required: "<%=dictLanguage.Item(Session("language")&"_opc_js_47")%>",
					equalTo: "<%=dictLanguage.Item(Session("language")&"_opc_js_48")%>"
				}
			}
			});
			});
		</script>
		
        <% end if%>
        
        <% 
          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          '// START: Product Reviews Reminder
          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          Dim pRSActive
          
          pRSActive = False
          query = "SELECT TOP 1 pcRS_Active FROM pcRevSettings"
          set rs=server.CreateObject("ADODB.RecordSet")
          set rs=conntemp.execute(query)
          If rs.eof = False Then
             pRSActive = CBool(CLng(rs("pcRS_Active")) <> 0)
          End If
          rs.close
        
          If pRSActive Then ' If Product Reviews are active, ask customer if they want to be notified
            if pcLoginMode=2 then ' Start - New customer that is registering, opt in by default
            %>
              <input type="hidden" value="1" name="pcAllowReviewEmails">
            <% else %>

              <div class="form-group">
                <%= dictLanguage.Item(Session("language")&"_order_FF")%> <input type="checkbox" value="1" name="pcAllowReviewEmails"<% If pcAllowReviewEmails<>0 Then response.write " CHECKED" %> class="clearBorder">
              </div>
            <% 
            end if ' End - New customer registering
          end if
          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          '// END: Product Reviews Reminder
          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        %>
        
        <%
          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          '// START: Special Customer Fields
          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
           
          if session("pcSFCustFieldsExist")="YES" then
            pcArr=session("sf_nc_custfields")
  
            For k=0 to ubound(pcArr,2)
            
            pcv_ShowField=Session("pcSF_ShowField"&k)
            Session("pcSF_ShowField"&k)=""
            
            if pcv_ShowField = 1 then
				
				if pcArr(6,k) = "1" then
                	pcv_isCFRequired = true
                else
                	pcv_isCFRequired = false
                end if
            %>

              <div class="form-group">
                <label for=""><%=pcArr(1,k)%>:<% If pcv_isCFRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
                <%if pcArr(2,k)="1" then%>
                  <input type="checkbox" name="custfield<%=pcArr(0,k)%>" <%if pcArr(10,k)<>"" then%>value="<%=pcArr(10,k)%>"<%else%><%if pcArr(3,k)<>"" then%>value="<%=pcArr(3,k)%>"<%else%>value="1"<%end if%><%end if%> <%if pcArr(10,k)<>"" OR Session("pcSFcustfield"&pcArr(0,k))<>"" then%>checked<%end if%> class="clearBorder">
                <%else%>
                  <input type="text" class="form-control" name="custfield<%=pcArr(0,k)%>" value="<%if session("idcustomer")=0 then%><%if pcArr(3,k)<>"" then%><%=pcArr(3,k)%><%else%><%=Session("pcSFcustfield"&pcArr(0,k))%><%end if%><%else%><%if pcArr(10,k)<>"" then%><%=pcArr(10,k)%><%else%><%=Session("pcSFcustfield"&pcArr(0,k))%><%end if%><%end if%>" size="<%=pcArr(4,k)%>" <%if pcArr(5,k)>"0" then%>maxlength="<%=pcArr(5,k)%>"<%end if%>>
                <%end if%>
                
                <% pcs_RequiredImageTagHorizontal "custfield"&pcArr(0,k), pcv_isCFRequired %>	
              <%if trim(pcArr(11,k))<>"" then%>
                  <span class="help-block"><%=pcArr(11,k)%></span>
              <%end if%>
              </div>
              
            <%else%>
            
              <div class="form-group">
                <%if pcArr(2,k)="1" then%>
                <input type="hidden" name="custfield<%=pcArr(0,k)%>" <%if pcArr(10,k)<>"" then%>value="<%=pcArr(10,k)%>"<%else%><%if pcArr(3,k)<>"" then%>value="<%=pcArr(3,k)%>"<%else%>value="1"<%end if%><%end if%>>
                <%else%>
                <input type="hidden" name="custfield<%=pcArr(0,k)%>" value="<%if session("idcustomer")=0 then%><%if pcArr(3,k)<>"" then%><%=pcArr(3,k)%><%else%><%=Session("pcSFcustfield"&pcArr(0,k))%><%end if%><%else%><%if pcArr(10,k)<>"" then%><%=pcArr(10,k)%><%else%><%=Session("pcSFcustfield"&pcArr(0,k))%><%end if%><%end if%>">
                <%end if%>
                <input type="hidden" name="pcv_ShowFieldRequired" value="NO" />
              </div>
              
            <%end if%>
          <% Next
        end if
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '// END: Special Customer Fields
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        %>
          
        <% ' If the referrer drop down field is enabled, show it for a new customer
        if Session("idCustomer")=0 AND ((RefNewCheckout="1" AND pcLoginMode=0) OR (RefNewReg="1" AND pcLoginMode=2)) then %>

          <div class="form-group">
            <label for=""><%=ReferLabel%><% If pcv_IsIDReferRequired Then %><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div><% End If %></label>
            
              <select class="form-control" name="IDRefer" id="IDRefer">
                <option value="0" <%if Session("pcSFIDrefer")="" then%>selected<%end if%>></option>
                <% 
                query="select idrefer, [name] from Referrer where removed=0 order by SortOrder;"
                set rs=Server.CreateObject("ADODB.Recordset")
                set rs=connTemp.execute(query)
                do while not rs.eof
                  intIdrefer=rs("idrefer")
                  strName=rs("name") %>
                  <option value="<%=intIdrefer%>" <%if Session("pcSFIDrefer")=trim(intIdrefer) then%>selected<%end if%>><%=strName%></option>
                  <% rs.movenext
                loop
                set rs = nothing 
                %>
              </select>
              <%pcs_RequiredImageTagHorizontal "IDRefer", pcv_IsIDReferRequired %>

          </div>
          
        <% end if
        'End If the referrer drop down field is enabled, show it for a new customer %>

        <% 'If newsletter is enabled, show it for new customer and when existing customers edit their account
		if (session("SF_MU_Setup")<>"1") AND (AllowNews="1" and Session("idCustomer")<>0 and pcLoginMode<>0) OR (AllowNews="1" AND Session("idCustomer")=0 AND ((NewsCheckout="1" AND pcLoginMode=0) OR (AllowNews="1" AND pcLoginMode=1) OR (NewsReg="1" AND pcLoginMode=2))) then %>
		<div class="pcSpacer"></div>
        <div class="form-group">
            <input type="checkbox" value="1" name="CRecvNews" <%if pcIntRecvNews="1" then%>checked<%end if%> class="clearBorder" />&nbsp;<%=NewsLabel%>
        </div>
		<% end if	
		'End If newsletter is enabled, show it for new customer
        %>
        <div class="form-group">
          <button class="pcButton pcButtonContinue" id="SubmitCustomerData" name="SubmitCustomerData">
           	<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="Save and Continue" />
            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update") %></span>
          </button>

          <%
            ' Take customer back to account menu if in editing mode
              if pcLoginMode=1 then
          %>
          
            <a class="pcButton pcButtonBack" href="custpref.asp">
              <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
              <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
            </a>
          <%
              end if
          %>
        </div>

  	</form>
    <div class="pcSpacer"></div>
    
  </div>
</div>
<%
'// Clear Un-Managed Sessions
Session("pcSFCRecvNews")=""
Session("pcSFIDrefer")=""
Session("pcSFpcCustomerEmail")=""
Session("pcSFpcCustomerEmail2")=""
Session("pcSFpcShippingReferenceId")=""
Session("pcSFpcShippingResidential")=""
Session("pcSFpcShippingReferenceId")=""
Session("pcSFidPayment")=""
Session("pcSFcomments")=""
Session("pcSFord_OrderName")=""
session("sf_nc_custfields")=""
session("pcSFCustFieldsExist")=""
%>

<!--#include file="footer_wrapper.asp"-->
