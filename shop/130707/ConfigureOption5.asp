<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx Web Services Shipping Configuration"
pageIcon="pcv4_icon_settings.png"
%>
<% Section="shipOpt" %>
<% pcPageName = "ConfigureOption5.asp" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%
Dim pcv_strMethodName
Dim FEDEX_URL, pcv_strErrorMsg, objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML

'// Validate phone
function fnStripPhone(PhoneField)
	PhoneField=replace(PhoneField," ","")
	PhoneField=replace(PhoneField,"-","")
	PhoneField=replace(PhoneField,".","")
	PhoneField=replace(PhoneField,"(","")
	PhoneField=replace(PhoneField,")","")
	fnStripPhone = PhoneField
end function

'**************************************************************************
' START: If registration request was submitted, process request
'**************************************************************************
Dim pcv_strAccountName, pcv_strMeterNumber, pcv_strCarrierCode

pcv_strMethodName = "SubscriptionRequest"
pcv_strVersion = FedExWS_RegistrationVersion
CustomerTransactionIdentifier = "Subscription_Request"

if request.form("submit")<>"" then

	'// Set error count
	pcv_intErr=0

	'// generic error for page
	pcv_strGenericPageError = "At least one required field was empty."

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ValidateTextField	"FedExWSMode", true, 4
	pcs_ValidateTextField	"FedExWS_AccountNumber", true, 20

	pcs_ValidateTextField	"FedExWS_ShippingAddress", true, 200
	pcs_ValidateTextField	"FedExWS_ShippingCity", true, 20
	pcs_ValidateTextField	"FedExWS_ShippingCountryCode", true, 2
	pcs_ValidateTextField	"FedExWS_ShippingStateCode", FedExRequiresStateProvince(Session("pcAdminFedExWS_ShippingCountryCode")), 20
	pcs_ValidateTextField	"FedExWS_ShippingProvinceCode", false, 20
	pcs_ValidateTextField	"FedExWS_ShippingPostalCode", true, 20

	pcs_ValidateTextField	"FedExWS_FirstName", true, 20
	pcs_ValidateTextField	"FedExWS_LastName", true, 20
	pcs_ValidateTextField	"FedExWS_CompanyName", false, 20
	pcs_ValidatePhoneNumber	"FedExWS_PhoneNumber", true, 16
	pcs_ValidateEmailField	"FedExWS_eMailAddress", false, 250
	pcs_ValidateTextField	"FedExWS_Line1", true, 200
	pcs_ValidateTextField	"FedExWS_Line2", false, 200
	pcs_ValidateTextField	"FedExWS_City", true, 20
	pcs_ValidateTextField	"FedExWS_CountryCode", true, 2
	pcs_ValidateTextField	"FedExWS_StateCode", FedExRequiresStateProvince(Session("pcAdminFedExWS_CountryCode")), 20
	pcs_ValidateTextField	"FedExWS_ProvinceCode", false, 20
	pcs_ValidateTextField	"FedExWS_PostalCode", true, 20

	if len(Session("pcAdminFedExWS_Line1"))<1 then
		pcv_intErr=pcv_intErr+1
	end if

	pcs_ValidateTextField	"FedExWS_PostalCode", true, 20

	if len(Session("pcAdminFedExWS_PostalCode"))<1 then
		pcv_intErr=pcv_intErr+1
	end if

	Session("pcAdminFedExWS_AccountNumber") = replace(Session("pcAdminFedExWS_AccountNumber"),"-","")
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If pcv_intErr>0 Then
		response.redirect pcPageName & "?msg=" & pcv_strGenericPageError
	Else

		'// Save collected data in database
		FedExWSAPI_ID=getUserInput(request("FedExWSAPI_ID"),4)

		'// Generate the Query (Save form data)
		if FedExWSAPI_ID=0 then
			query="INSERT INTO FedExWSAPI (FedExAPI_PersonName, FedExAPI_CompanyName, FedExAPI_Department, FedExAPI_PhoneNumber, FedExAPI_FaxNumber, FedExAPI_EmailAddress, FedExAPI_Line1, FedExAPI_Line2, FedExAPI_city, FedExAPI_State, FedExAPI_PostalCode, FedExAPI_Country) VALUES ('"& Session("pcAdminFedExWS_FirstName") & " " & Session("pcAdminFedExWS_LastName") &"', '"&Session("pcAdminFedExWS_CompanyName")&"', '"&Session("pcAdminFedExWS_Department")&"', '"&Session("pcAdminFedExWS_PhoneNumber")&"', '"&Session("pcAdminFedExWS_FaxNumber")&"', '"&Session("pcAdminFedExWS_EmailAddress")&"', '"&Session("pcAdminFedExWS_Line1")&"', '"&Session("pcAdminFedExWS_Line2")&"', '"&Session("pcAdminFedExWS_City")&"', '"&Session("pcAdminFedExWS_StateOrProvinceCode")&"', '"&Session("pcAdminFedExWS_PostalCode")&"', '"&Session("pcAdminFedExWS_CountryCode")&"');"
		else
			query="UPDATE FedExWSAPI SET FedExAPI_PersonName='"& Session("pcAdminFedExWS_FirstName") & " " & Session("pcAdminFedExWS_LastName") &"', FedExAPI_CompanyName='"&Session("pcAdminFedExWS_CompanyName")&"', FedExAPI_Department='"&Session("pcAdminFedExWS_Department")&"', FedExAPI_PhoneNumber='"&Session("pcAdminFedExWS_PhoneNumber")&"', FedExAPI_FaxNumber='"&Session("pcAdminFedExWS_FaxNumber")&"', FedExAPI_EmailAddress='"&Session("pcAdminFedExWS_EmailAddress")&"', FedExAPI_Line1='"&Session("pcAdminFedExWS_Line1")&"', FedExAPI_Line2='"&Session("pcAdminFedExWS_Line2")&"', FedExAPI_city='"&Session("pcAdminFedExWS_City")&"', FedExAPI_State='"&Session("pcAdminFedExWS_StateOrProvinceCode")&"', FedExAPI_PostalCode='"&Session("pcAdminFedExWS_PostalCode")&"', FedExAPI_Country='"&Session("pcAdminFedExWS_CountryCode")&"' WHERE FedExAPI_ID=1;"
		end if

		'// Execute the Query
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Set our Object.
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        set objFedExClass = New pcFedExWSClass

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Build Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		NameOfMethod = "RegisterWebCspUserRequest"
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns=""http://fedex.com/ws/registration/v" & pcv_strVersion & """>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
		objFedExClass.WriteParent NameOfMethod, ""

		objFedExClass.WriteParent "WebAuthenticationDetail", ""
			objFedExClass.WriteParent "CspCredential", ""
				objFedExClass.AddNewNode "Key", pcv_strCSPKey
				objFedExClass.AddNewNode "Password", pcv_strCSPPassword
			objFedExClass.WriteParent "CspCredential", "/"
		objFedExClass.WriteParent "WebAuthenticationDetail", "/"
		
		objFedExClass.WriteParent "ClientDetail", ""
			objFedExClass.AddNewNode "AccountNumber", Session("pcAdminFedExWS_AccountNumber")
			objFedExClass.AddNewNode "ClientProductId", pcv_strClientProductID
			objFedExClass.AddNewNode "ClientProductVersion", pcv_strClientProductVersion
		objFedExClass.WriteParent "ClientDetail", "/"

		'--------------------
		'// TransactionDetail
		'--------------------
		objFedExClass.WriteParent "TransactionDetail", ""
			objFedExClass.AddNewNode "CustomerTransactionId", "Registration Request"
		objFedExClass.WriteParent "TransactionDetail", "/"

		'--------------------
		'// Version
		'--------------------
		objFedExClass.WriteParent "Version", ""
			objFedExClass.AddNewNode "ServiceId", "fcas"
			objFedExClass.AddNewNode "Major", pcv_strVersion
			objFedExClass.AddNewNode "Intermediate", "0"
			objFedExClass.AddNewNode "Minor", "0"
		objFedExClass.WriteParent "Version", "/"

		objFedExClass.WriteSingleParent "Categories", "SHIPPING"
			objFedExClass.WriteParent "ShippingAddress", ""
				objFedExClass.AddNewNode "StreetLines", Session("pcAdminFedExWS_ShippingAddress")
				objFedExClass.AddNewNode "City", Session("pcAdminFedExWS_ShippingCity")
				stateOrProvinceCode = Session("pcAdminFedExWS_ShippingStateCode")
				If Len(stateOrProvinceCode) < 1 Then stateOrProvinceCode = Session("pcAdminFedExWS_ShippingProvinceCode")
        If FedExRequiresStateProvince(Session("pcAdminFedExWS_ShippingCountryCode")) Then
				    objFedExClass.AddNewNode "StateOrProvinceCode", stateOrProvinceCode
        End If
				objFedExClass.AddNewNode "PostalCode", Session("pcAdminFedExWS_ShippingPostalCode")
				objFedExClass.AddNewNode "CountryCode", Session("pcAdminFedExWS_ShippingCountryCode")
			objFedExClass.WriteParent "ShippingAddress", "/"

			objFedExClass.WriteParent "UserContactAndAddress", ""
				objFedExClass.WriteParent "Contact", ""
					objFedExClass.WriteParent "PersonName", ""
						objFedExClass.AddNewNode "FirstName", Session("pcAdminFedExWS_FirstName")
						objFedExClass.AddNewNode "LastName", Session("pcAdminFedExWS_LastName")
					objFedExClass.WriteParent "PersonName", "/"
					objFedExClass.AddNewNode "CompanyName", Session("pcAdminFedExWS_CompanyName")
					objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminFedExWS_PhoneNumber"))
					objFedExClass.AddNewNode "EMailAddress", Session("pcAdminFedExWS_eMailAddress")
				objFedExClass.WriteParent "Contact", "/"
				objFedExClass.WriteParent "Address", ""
					objFedExClass.AddNewNode "StreetLines", Session("pcAdminFedExWS_Line1") & " " & Session("pcAdminFedExWS_Line2")
					objFedExClass.AddNewNode "City", Session("pcAdminFedExWS_City")
          If FedExRequiresStateProvince(Session("pcAdminFedExWS_CountryCode")) Then
						stateOrProvinceCode = Session("pcAdminFedExWS_StateCode")
						If Len(stateOrProvinceCode) < 1 Then stateOrProvinceCode = Session("pcAdminFedExWS_ProvinceCode")
					 	objFedExClass.AddNewNode "StateOrProvinceCode", stateOrProvinceCode
          End If
					objFedExClass.AddNewNode "PostalCode", Session("pcAdminFedExWS_PostalCode")
					objFedExClass.AddNewNode "CountryCode", Session("pcAdminFedExWS_CountryCode")
				objFedExClass.WriteParent "Address", "/"
			objFedExClass.WriteParent "UserContactAndAddress", "/"

		objFedExClass.EndXMLTransaction NameOfMethod

		'// Print out our newly formed request xml
		'response.Clear()
		'response.ContentType="text/xml"
		'response.Write(fedex_postdataWS)
		'response.End()

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Send Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		call objFedExClass.SendXMLRequest(fedex_postdataWS)
		'// Print out our response
		'response.Clear()
		'response.ContentType="text/xml"
		'response.write FEDEXWS_result
		'response.end

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Load Our Response.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.LoadXMLResults(FEDEXWS_result)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Baseline Logging
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Log our Transaction
		call objFedExClass.pcs_LogTransaction(fedex_postdataWS, NameOfMethod&"_in"&q&".in", true)
		'// Log our Response
		call objFedExClass.pcs_LogTransaction(FEDEXWS_result, NameOfMethod&"_out"&q&".out", true)

        fedex_xmlPrefix = objFedExClass.GetXMLPrefix(pcv_strVersion)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Redirect with a Message OR complete some task.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// ERROR
		pcv_strNotificationCode = objFedExClass.ReadResponseNode("//<VER>RegisterWebCspUserReply", "<VER>Notifications/<VER>Severity")
		If pcv_strNotificationCode <> "SUCCESS" Then
			pcv_strErrorMessage = objFedExClass.ReadResponseNode("//<VER>RegisterWebCspUserReply", "<VER>Notifications/<VER>Message")
			response.redirect "ConfigureOption5.asp?msg="&pcv_strErrorMessage
			response.end
		End If

		'// Web User Credentials
		pcv_strWUKey = objFedExClass.ReadResponseNode("//<VER>RegisterWebCspUserReply", "<VER>Credential/<VER>Key")

		pcv_strWUPassword = objFedExClass.ReadResponseNode("//<VER>RegisterWebCspUserReply", "<VER>Credential/<VER>Password")

			'// Ensure that the MeterNumber exists
		if pcv_strWUKey&""="" OR pcv_strWUPassword&""="" then
			response.redirect pcPageName & "?msg=There was an error activating your FedEx account. The FedEx servers may be down temporarily. Please try again later."
		else
			'// Process Subscribe Request!
			NameOfMethod = "SubscriptionRequest"
			fedex_postdataWS=""
			fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns=""http://fedex.com/ws/registration/v" & pcv_strVersion & """>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
			
			objFedExClass.WriteParent NameOfMethod, ""
	
			objFedExClass.WriteParent "WebAuthenticationDetail", ""
				objFedExClass.WriteParent "CspCredential", ""
					objFedExClass.AddNewNode "Key", pcv_strCSPKey
					objFedExClass.AddNewNode "Password", pcv_strCSPPassword
				objFedExClass.WriteParent "CspCredential", "/"
				objFedExClass.WriteParent "UserCredential", ""
					objFedExClass.AddNewNode "Key", pcv_strWUKey
					objFedExClass.AddNewNode "Password", pcv_strWUPassword
				objFedExClass.WriteParent "UserCredential", "/"
			objFedExClass.WriteParent "WebAuthenticationDetail", "/"
			
			objFedExClass.WriteParent "ClientDetail", ""
				objFedExClass.AddNewNode "AccountNumber", Session("pcAdminFedExWS_AccountNumber")
				objFedExClass.AddNewNode "ClientProductId", pcv_strClientProductID
				objFedExClass.AddNewNode "ClientProductVersion", pcv_strClientProductVersion
			objFedExClass.WriteParent "ClientDetail", "/"

			objFedExClass.WriteParent "Version", ""
				objFedExClass.AddNewNode "ServiceId", "fcas"
				objFedExClass.AddNewNode "Major", pcv_strVersion
				objFedExClass.AddNewNode "Intermediate", "0"
				objFedExClass.AddNewNode "Minor", "0"
			objFedExClass.WriteParent "Version", "/"

			objFedExClass.AddNewNode "CspSolutionId", pcv_strCSPSolutionID
			objFedExClass.AddNewNode "CspType", "CERTIFIED_SOLUTION_PROVIDER"

			objFedExClass.WriteParent "Subscriber", ""
				objFedExClass.AddNewNode "AccountNumber", Session("pcAdminFedExWS_AccountNumber")
				objFedExClass.WriteParent "Contact", ""
					objFedExClass.AddNewNode "PersonName", Session("pcAdminFedExWS_LastName")
					objFedExClass.AddNewNode "CompanyName", Session("pcAdminFedExWS_CompanyName")
					objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminFedExWS_PhoneNumber"))
					objFedExClass.AddNewNode "EMailAddress", Session("pcAdminFedExWS_eMailAddress")
				objFedExClass.WriteParent "Contact", "/"
				objFedExClass.WriteParent "Address", ""
					objFedExClass.AddNewNode "StreetLines", Session("pcAdminFedExWS_Line1") & " " & Session("pcAdminFedExWS_Line2")
					objFedExClass.AddNewNode "City", Session("pcAdminFedExWS_City")
          If Session("pcAdminFedExWS_CountryCode") <> "GB" Then
						stateOrProvinceCode = Session("pcAdminFedExWS_StateCode")
						If Len(stateOrProvinceCode) < 1 Then stateOrProvinceCode = Session("pcAdminFedExWS_ProvinceCode")
						objFedExClass.AddNewNode "StateOrProvinceCode", stateOrProvinceCode
          End If
					objFedExClass.AddNewNode "PostalCode", Session("pcAdminFedExWS_PostalCode")
					objFedExClass.AddNewNode "CountryCode", Session("pcAdminFedExWS_CountryCode")
				objFedExClass.WriteParent "Address", "/"
			objFedExClass.WriteParent "Subscriber", "/"

			objFedExClass.WriteParent "AccountShippingAddress", ""
				objFedExClass.AddNewNode "StreetLines", Session("pcAdminFedExWS_ShippingAddress")
				objFedExClass.AddNewNode "City", Session("pcAdminFedExWS_ShippingCity")
        If Session("pcAdminFedExWS_ShippingCountryCode") <> "GB" Then
					stateOrProvinceCode = Session("pcAdminFedExWS_ShippingStateCode")
					If Len(stateOrProvinceCode) < 1 Then stateOrProvinceCode = Session("pcAdminFedExWS_ShippingProvinceCode")
				 	objFedExClass.AddNewNode "StateOrProvinceCode", stateOrProvinceCode
        End If
				objFedExClass.AddNewNode "PostalCode", Session("pcAdminFedExWS_ShippingPostalCode")
				objFedExClass.AddNewNode "CountryCode", Session("pcAdminFedExWS_ShippingCountryCode")
			objFedExClass.WriteParent "AccountShippingAddress", "/"

			objFedExClass.EndXMLTransaction NameOfMethod

			'// Print out our newly formed request xml
			'response.Clear()
			'response.ContentType="text/xml"
			'response.Write(fedex_postdataWS)
			'response.End()

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Send Our Transaction.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			call objFedExClass.SendXMLRequest(fedex_postdataWS)
			
			'// Print out our response
			'response.Clear()
			'response.ContentType="text/xml"
			'response.write FEDEXWS_result
			'response.end

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Load Our Response.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			call objFedExClass.LoadXMLResults(FEDEXWS_result)

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Baseline Logging
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Log our Transaction
			call objFedExClass.pcs_LogTransaction(fedex_postdataWS, NameOfMethod&"_in"&q&".in", true)
			'// Log our Response
			call objFedExClass.pcs_LogTransaction(FEDEXWS_result, NameOfMethod&"_out"&q&".out", true)

			fedex_xmlPrefix = objFedExClass.GetXMLPrefix(pcv_strVersion)

			'// ERROR
			pcv_strNotificationCode = objFedExClass.ReadResponseNode("//<VER>SubscriptionReply", "<VER>Notifications/<VER>Severity")
			If pcv_strNotificationCode="SUCCESS" Then
			Else
				pcv_strErrorMessage = objFedExClass.ReadResponseNode("//<VER>SubscriptionReply", "<VER>Notifications/<VER>Message")
				response.redirect pcPageName & "?msg="&pcv_strErrorMessage
			End If

			'// Web User Credentials
			pcv_strWUMeterNumber = objFedExClass.ReadResponseNode("//<VER>SubscriptionReply", "<VER>MeterDetail/<VER>MeterNumber")
			
			query="UPDATE ShipmentTypes SET [password]='"&pcv_strWUMeterNumber&"', userID='"&Session("pcAdminFedExWS_AccountNumber")&"', AccessLicense='LIVE', FedExKey='"&pcv_strWUKey&"', FedExPwd='"&pcv_strWUPassword&"' WHERE (((ShipmentTypes.idShipment)=9));"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing

				if err.number<>0 then
					call closedb()
					response.redirect pcPageName & "?msg=There was an error activating your FedEx account. Please submit your registration request again."
				else
					'// Generate the Query (Update shipment types)
					query="UPDATE ShipmentTypes SET AccessLicense='LIVE' WHERE (((ShipmentTypes.idShipment)=9));"
					set rs=server.CreateObject("ADODB.RecordSet")
					'// Execute the Query
					set rs=conntemp.execute(query)
					set rs=nothing
					call closedb()
				'// No errors, redirect to next step
					session("FedExWSSetUP")="YES"
					pcs_ClearAllSessions()
					response.redirect "FEDEXWS_EditShipOptions.asp"
					response.end

				end if
			end if


	end if
end if
'**************************************************************************
' END: If registration request was submitted, process request
'**************************************************************************




'**************************************************************************
' START: Was FedExWS was previously registered by querying the database ?
'**************************************************************************
if request("changeMode")="" then
	query="SELECT userID FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=9));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	If rs.eof then
		'FedExWS was previously activated - redirect
		session("FedExWSSetUP")="YES"
		response.redirect "FEDEXWS_EditShipOptions.asp"
		response.end
	end if
	set rs=nothing
end if
'**************************************************************************
' END: Was FedExWS was previously registered by querying the database ?
'**************************************************************************
%>

<%
'**************************************************************************
' START: Get Fed Ex credentials
'**************************************************************************
'// Get Access License
query="SELECT ShipmentTypes.AccessLicense, ShipmentTypes.userID FROM ShipmentTypes WHERE idShipment=9;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

Session("pcAdminFedExWS_AccountNumber") = rs("userID")
strAccessLicense=rs("AccessLicense")

if len(strAccessLicense)<1 then
	strAccessLicense="TEST"
end if

'// Get Form Data
query="SELECT FedExAPI_ID, FedExAPI_PersonName, FedExAPI_CompanyName, FedExAPI_Department, FedExAPI_PhoneNumber, FedExAPI_FaxNumber, FedExAPI_EmailAddress, FedExAPI_Line1, FedExAPI_Line2, FedExAPI_city, FedExAPI_State, FedExAPI_PostalCode, FedExAPI_Country FROM FedExWSAPI;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	FedExWSAPI_ID=rs("FedExAPI_ID")
	if request("changeMode")="Y" then
		Session("pcAdminFedExWS_PersonName")=rs("FedExAPI_PersonName")
		Session("pcAdminFedExWS_CompanyName")=rs("FedExAPI_CompanyName")
		Session("pcAdminFedExWS_Department")=rs("FedExAPI_Department")
		Session("pcAdminFedExWS_PhoneNumber")=rs("FedExAPI_PhoneNumber")
		Session("pcAdminFedExWS_PagerNumber")=rs("FedExAPI_PagerNumber")
		Session("pcAdminFedExWS_FaxNumber")=rs("FedExAPI_FaxNumber")
		Session("pcAdminFedExWS_EmailAddress")=rs("FedExAPI_EmailAddress")
		Session("pcAdminFedExWS_Line1")=rs("FedExAPI_Line1")
		Session("pcAdminFedExWS_Line2")=rs("FedExAPI_Line2")
		Session("pcAdminFedExWS_city")=rs("FedExAPI_city")
		Session("pcAdminFedExWS_StateOrProvinceCode")=rs("FedExAPI_State")
		Session("pcAdminFedExWS_PostalCode")=rs("FedExAPI_PostalCode")
		Session("pcAdminFedExWS_Country")=rs("FedExAPI_Country")
	end if
else
	FedExWSAPI_ID=0
end if
set rs=nothing
'**************************************************************************
' END: Get Fed Ex credentials
'**************************************************************************
%>
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
	<input type="hidden" name="FedExWSAPI_ID" value="<%=FedExWSAPI_ID%>">
	<input type="hidden" name="changeMode" value="<%=request("changeMode")%>">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2">
			If you have any problems with the registration/subscription process, contact FedEx Technical Support at 1.800.820.1336 or via email at <a href="mailto:websupport@fedex.com">websupport@fedex.com</a>.
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<% if intErrCnt>0 then %>
			<tr>
			<td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="4">
				<tr>
					<td width="4%" valign="top"><img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"></td>
					<td width="96%" valign="top" class="message"><font color="#FF9900"><b>
					  <% response.write intErrCnt&" error(s) were located. <ul>"&strErrMsg&"</ul>"%></b></font></td>
				</tr>
		</table>
			</td>
			</tr>
		<% end if %>

		<tr>
			<td colspan="2">
			<span class="pcCPnotes">Click &quot;Continue&quot; below to submit your FedEx subscription request.</span>
			<input name="FedExWSMode" type="hidden" value="LIVE" />
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Account Details</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%"><p>FedEx Account Number:</p></td>
			<td width="77%"><p><input name="FedExWS_AccountNumber" type="text" value="<%=pcf_FillFormField("FedExWS_AccountNumber", true)%>" size="15" maxlength="25">
			<%pcs_RequiredImageTag "FedExWS_AccountNumber", true%></p></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Shipping Address</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%"><p>Address: </p></td>
			<td width="77%">
			  <p><input name="FedExWS_ShippingAddress" type="text" value="<%=pcf_FillFormField("FedExWS_ShippingAddress", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_ShippingAddress", true%></p></td>
		</tr>
		<tr>
			<td width="23%"><p>City: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_ShippingCity" type="text" value="<%=pcf_FillFormField("FedExWS_ShippingCity", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_ShippingCity", true%></p></td>
		</tr>
        
            <%
			'///////////////////////////////////////////////////////////
			'// START: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			' 
			' 1) Place this section ABOVE the Country field
			' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
			' 3) Additional Required Info
					
			'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
			pcv_isStateCodeRequired = FedExRequiresStateProvince(Session("pcAdminFedExWS_ShippingCountryCode")) '// determines if validation is performed (true or false)
			pcv_isProvinceCodeRequired = FedExRequiresStateProvince(Session("pcAdminFedExWS_ShippingCountryCode")) '// determines if validation is performed (true or false)
			pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)
					
			'// #3 Additional Required Info
			pcv_strTargetForm = "form1" '// Name of Form
			pcv_strCountryBox = "FedExWS_ShippingCountryCode" '// Name of Country Dropdown
			pcv_strTargetBox = "FedExWS_ShippingStateCode" '// Name of State Dropdown
			pcv_strProvinceBox =  "FedExWS_ShippingProvinceCode" '// Name of Province Field
					
            pcv_strSessionPrefix = ""

			'// Set local Country to Session
			if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
				Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session("pcAdminFedExWS_ShippingCountryCode")
			end if
					
			'// Set local State to Session
			if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
				Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminFedExWS_ShippingStateCode")
			end if
					
			'// Set local Province to Session
			if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
				Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session("pcAdminFedExWS_ShippingProvinceCode")
			end if

			%>					
			<!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
			<%
			'///////////////////////////////////////////////////////////
			'// END: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			%>		
			<%
			'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
			pcs_CountryDropdown

			'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
			pcs_StateProvince
			%>	

		<tr>
			<td width="23%"><p>Postal Code:  </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_ShippingPostalCode" type="text" value="<%=pcf_FillFormField("FedExWS_ShippingPostalCode", true)%>" size="15" maxlength="20">
			  <%pcs_RequiredImageTag "FedExWS_ShippingPostalCode", true%></p>
            </td>
	    </tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Contact Address</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%"><p>First Name: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_FirstName" type="text" value="<%=pcf_FillFormField("FedExWS_FirstName", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_FirstName", true%></p></td>
		</tr>
		<tr>
			<td width="23%"><p>Last Name: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_LastName" type="text" value="<%=pcf_FillFormField("FedExWS_LastName", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_LastName", true%></p></td>
		</tr>
		<tr>
			<td width="23%"><p>Company Name: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_CompanyName" type="text" value="<%=pcf_FillFormField("FedExWS_CompanyName", false)%>" size="30" maxlength="100"></p></td>
		</tr>
		<% if len(Session("ErrFedExWS_PhoneNumber"))>0 then %>
		<tr>
			<td width="23%"></td>
			<td width="77%"><p>
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Phone Number.</p></td>
		</tr>
		<% end if %>
		<tr>
			<td width="23%"><p>Phone Number: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_PhoneNumber" type="text" value="<%=pcf_FillFormField("FedExWS_PhoneNumber", true)%>" size="16" maxlength="16">
			  <%pcs_RequiredImageTag "FedExWS_PhoneNumber", true%></p></td>
		</tr>
		<% if len(Session("ErrFedExWS_eMailAddress"))>0 then %>
		<tr>
			<td width="23%"></td>
			<td width="77%"><p>
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Email Address.</p></td>
		</tr>
		<% end if %>
		<tr>
			<td width="23%"><p>Email Address:  </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_eMailAddress" type="text" value="<%=pcf_FillFormField("FedExWS_eMailAddress", true)%>" size="40" maxlength="250">
			  <%pcs_RequiredImageTag "FedExWS_eMailAddress", true%></p></td>
			</tr>
		<tr>
			<td width="23%"><p>Address: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_Line1" type="text" value="<%=pcf_FillFormField("FedExWS_Line1", true)%>" size="40" maxlength="250">
			  <%pcs_RequiredImageTag "FedExWS_Line1", true%></p></td>
			</tr>
		<tr>
			<td width="23%"><p>&nbsp;</p></td>
			<td width="77%"><p>
			<input name="FedExWS_Line2" type="text" value="<%=pcf_FillFormField("FedExWS_Line2", false)%>" size="40" maxlength="250"></p></td>
			</tr>
		<tr>
			<td width="23%"><p>City: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_City" type="text" value="<%=pcf_FillFormField("FedExWS_City", true)%>" size="30" maxlength="250">
			  <%pcs_RequiredImageTag "FedExWS_City", true%></p>
			</td>
		</tr>

        <%
		'///////////////////////////////////////////////////////////
		'// START: COUNTRY AND STATE/ PROVINCE CONFIG
		'///////////////////////////////////////////////////////////
		' 
		' 1) Place this section ABOVE the Country field
		' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
		' 3) Additional Required Info
					
		'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
		pcv_isStateCodeRequired = FedExRequiresStateProvince(Session("pcAdminFedExWS_CountryCode")) '// determines if validation is performed (true or false)
		pcv_isProvinceCodeRequired = FedExRequiresStateProvince(Session("pcAdminFedExWS_CountryCode")) '// determines if validation is performed (true or false)
		pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)
					
		'// #3 Additional Required Info
		pcv_strTargetForm = "form1" '// Name of Form
		pcv_strCountryBox = "FedExWS_CountryCode" '// Name of Country Dropdown
		pcv_strTargetBox = "FedExWS_StateCode" '// Name of State Dropdown
		pcv_strProvinceBox =  "FedExWS_ProvinceCode" '// Name of Province Field
					
        pcv_strSessionPrefix = ""

		'// Set local Country to Session
		if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session("pcAdminFedExWS_CountryCode")
		end if
					
		'// Set local State to Session
		if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminFedExWS_StateOrProvinceCode")
		end if
					
		'// Set local Province to Session
		if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session("pcAdminFedExWS_StateOrProvinceCode")
		end if

		'///////////////////////////////////////////////////////////
		'// END: COUNTRY AND STATE/ PROVINCE CONFIG
		'///////////////////////////////////////////////////////////
		%>		
		<%
		'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
		pcs_CountryDropdown

		'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
		pcs_StateProvince
		%>
		<tr>
			<td width="23%"><p>Postal Code: </p></td>
			<td width="77%"><p>
			  <input name="FedExWS_PostalCode" type="text" value="<%=pcf_FillFormField("FedExWS_PostalCode", true)%>" size="15" maxlength="20">
			  <%pcs_RequiredImageTag "FedExWS_PostalCode", true%></p></td>
			</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2">
			<input type="submit" name="Submit" value="Continue" class="btn btn-primary">
			&nbsp;
			<input type="button" class="btn btn-default"  name="back" value="Back" onClick="JavaScript:history.back();">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->