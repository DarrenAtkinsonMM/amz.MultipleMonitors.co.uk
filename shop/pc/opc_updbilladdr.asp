<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/pcUSPSClass.asp"-->
<!--#include file="../includes/validation.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="opc_contentType.asp" -->

<%
Call SetContentType()

Dim pcv_strCatcher
pcv_strCatcher = Session("pcCartIndex")
If pcv_strCatcher=0 Then
	pcv_strCatcher=""		
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
End If


%><!--#include file="../includes/pcServerSideValidation.asp" --><%

Function generatePassword(passwordLength)
	Dim sDefaultChars
	Dim iCounter
	Dim sMyPassword
	Dim iPickedChar
	Dim iDefaultCharactersLength
	Dim iPasswordLength
	
	sDefaultChars="abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
	iPasswordLength=passwordLength
	iDefaultCharactersLength = Len(sDefaultChars) 
	Randomize
	For iCounter = 1 To iPasswordLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1) 
		sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
	Next 
	generatePassword = sMyPassword
End Function

pcErrMsg=""

'//////////////////////////////////////////////////////////////////////////
'// START: VALIDATE BILLING
'//////////////////////////////////////////////////////////////////////////

pcStrBillingFirstName=URLDecode(getUserInput(request("billfname"),50))
pcStrBillingLastName=URLDecode(getUserInput(request("billlname"),50))
pcStrBillingCompany=URLDecode(getUserInput(request("billcompany"),150))
pcStrBillingVATID=URLDecode(getUserInput(request("billVATID"),150))
pcStrBillingSSN=URLDecode(getUserInput(request("billSSN"),150))
pcStrBillingPhone=URLDecode(getUserInput(request("billphone"),20))
pcStrCustomerEmail=URLDecode(getUserInput(request("billemail"),150))
pcStrBillingAddress=URLDecode(getUserInput(request("billaddr"),255))
pcStrBillingPostalCode=URLDecode(getUserInput(request("billzip"),10))
pcStrBillingStateCode=URLDecode(getUserInput(request("billstate"),4))
pcStrBillingProvince=URLDecode(getUserInput(request("billprovince"),150))
pcStrBillingCity=URLDecode(getUserInput(request("billcity"),50))
pcStrBillingCountryCode=URLDecode(getUserInput(request("billcountry"),4))
pcStrBillingAddress2=URLDecode(getUserInput(request("billaddr2"),255))
pcStrBillingFax=URLDecode(getUserInput(request("billfax"),20))
If scComResShipAddress = "0" Then
    pcIntShippingResidential=URLDecode(getUserInput(request("pcAddressType"),0))
    If pcIntShippingResidential<>"" Then
        If Not IsNumeric(pcIntShippingResidential) Then
            pcIntShippingResidential=scComResShipAddress
        End If
    End If
End If
If len(pResidentialShipping)=0 Then
    Select Case scComResShipAddress
        Case "1"
            pcIntShippingResidential="-1"
        Case "2"
            pcIntShippingResidential="0"
        Case "3"
            if session("customerType")="1" then
                pcIntShippingResidential="0"
            else
                pcIntShippingResidential="-1"
            end if
    End Select
End If

pcStrNewPass1=""
if scGuestCheckoutOpt=2 then
	pcStrNewPass1=URLDecode(getUserInput(request("billpass"),250))
	pcStrNewPass2=URLDecode(getUserInput(request("billrepass"),250))
end if

'Check the PostalCode Length for United States
If pcStrBillingCountryCode="US" Then
	if len(pcStrBillingPostalCode)<5 then
		response.clear
		Call SetContentType()
		response.Write("ZIPLENGTH")
		response.End()
	end if
End If

if pcStrBillingFirstName="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_58")&"</li>"
end if
if pcStrBillingLastName="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_59")&"</li>"
end if
if pcStrBillingAddress="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_60")&"</li>"
end if
if pcStrBillingCity="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_61")&"</li>"
end if
if pcStrBillingCountryCode="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_62")&"</li>"
end if
if (pcStrBillingCountryCode="US") OR (pcStrBillingCountryCode="CA") then
	if pcStrBillingStateCode="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_63")&"</li>"
	end if
	if pcStrBillingPostalCode="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_64")&"</li>"
	end if
end if

'// Validate billing address, if enabled.
Dim USPS_postdata, USPS_result, srvUSPSXmlHttp, objOutputXMLDoc  
pcv_boolIsBillingAddressValidated = getUserInput(request("IsBillingAddressValidated"), 5)
If (pcv_boolIsBillingAddressValidated = "") Or (pcv_boolIsBillingAddressValidated = "false") Then
    pcs_ValidateBillingAddress()
End If

if session("idCustomer")="" OR session("idCustomer")=0 then
	if pcStrCustomerEmail="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_65")&"</li>"
	else
		pcStrCustomerEmail=replace(pcStrCustomerEmail," ","")
		if instr(pcStrCustomerEmail,"@")=0 or instr(pcStrCustomerEmail,".")=0 then 
			pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_66")&"</li>"
		end if
	end if
	
	if scGuestCheckoutOpt=2 then
		if pcErrMsg="" then
			query="SELECT idCustomer FROM Customers WHERE [email] like '" & pcStrCustomerEmail & "';"
			set rs=connTemp.execute(query)
			if not rs.eof then
				pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_5a")&"</li>"
			end if
			set rs=nothing
		end if
	
		if pcStrNewPass1="" then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_5") & "</li>"
		end if

		if pcStrNewPass2="" then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_4") & "</li>"
		end if

		if pcStrNewPass1<>pcStrNewPass2 then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_3") & "</li>"
		end if
	end if
	
end if
if session("idCustomer")>"0" then
	query="SELECT idCustomer FROM Customers WHERE idcustomer=" & session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_67")&"</li>"
	end if
end if
tmpRecvNews=URLDecode(getUserInput(request("CRecvNews"),0))
if tmpRecvNews="" then
	tmpRecvNews=0
end if
if not IsNumeric(tmpRecvNews) then
	tmpRecvNews=0
end if
Session("pcSFCRecvNews")=tmpRecvNews


'//////////////////////////////////////////////////////////////////////////
'// END: VALIDATE BILLING
'//////////////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////////////
'// START: UPDATE BILLING
'//////////////////////////////////////////////////////////////////////////
if pcErrMsg="" then

	tmpIDRefer=URLDecode(getUserInput(request("IDRefer"),0))
	if tmpIDRefer<>"" then
		if not IsNumeric(tmpIDRefer) then
			tmpIDRefer=0
		end if
	else
		if session("idCustomer")="0" OR session("idCustomer")="" then
			tmpIDRefer=0
		end if
	end if
	Session("pcSFIDrefer")=tmpIDRefer
		
	if session("idCustomer")>"0" then

		if session("CustomerGuest") = "" OR isNULL(session("CustomerGuest")) then
			session("CustomerGuest") = 0
		end if
		query="UPDATE Customers SET [name]=N'" & pcStrBillingFirstName & "',lastName=N'" & pcStrBillingLastName & "',customerCompany=N'" & pcStrBillingCompany & "', pcCust_VATID='" & pcStrBillingVATID & "', pcCust_SSN='" & pcStrBillingSSN & "', phone='" & pcStrBillingPhone & "',address=N'" & pcStrBillingAddress & "',zip='" & pcStrBillingPostalCode & "',stateCode='" & pcStrBillingStateCode & "',state=N'" & pcStrBillingProvince & "',city=N'" & pcStrBillingCity & "',countryCode='" & pcStrBillingCountryCode & "',address2=N'" & pcStrBillingAddress2 & "',fax='" & pcStrBillingFax & "'"
		if session("CustomerGuest")>"0" then
			query=query & ",email='" & pcStrCustomerEmail & "'"
		end if
		
		tmpCustomerGuest = session("CustomerGuest")
		if pcStrNewPass1<>"" then
			pcPassword=pcf_PasswordHash(pcStrNewPass1)
			tmpCustomerGuest=0
			query=query & ",password='" & pcPassword & "'"
			query=query & ",pcCust_Guest=" & tmpCustomerGuest & ""
		end if
			
		query=query & " WHERE idcustomer=" & session("idCustomer") & " AND pcCust_Guest=" & session("CustomerGuest") & ";"

    call opendb()
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		set rs=nothing
		OKmsg="OK"
		
		session("CustomerGuest") = tmpCustomerGuest
	else
		if pcStrNewPass1<>"" then
			pcPassword=pcStrNewPass1
			tmpCustomerGuest=0
		else
			pcPassword=generatePassword(10)
			tmpCustomerGuest=1
		end if
		if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" OR scGuestCheckoutOpt=1 then
			tmpCustomerGuest=1
		end if
		pcPassword=pcf_PasswordHash(pcPassword)
		if tmpCustomerGuest=0 then
		query="SELECT idCustomer FROM Customers WHERE [email] like '" & pcStrCustomerEmail & "' AND pcCust_Guest<>1;"
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		if not rs.eof then
			tmpCustomerGuest=2
		end if
		set rs=nothing
		end if

		pcv_dateCustomerRegistration=Date()
		if SQL_Format="1" then
			pcv_dateCustomerRegistration=Day(pcv_dateCustomerRegistration)&"/"&Month(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
		else
			pcv_dateCustomerRegistration=Month(pcv_dateCustomerRegistration)&"/"&Day(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
		end if
		
		query="INSERT INTO Customers (Password,customerType,email,[name],lastName,customerCompany,pcCust_VATID,pcCust_SSN,phone,address,zip,stateCode,state,city,countryCode,address2,fax,pcCust_DateCreated,pcCust_Guest) VALUES ('" & pcPassword & "',0,'" & pcStrCustomerEmail & "',N'" & pcStrBillingFirstName & "',N'" & pcStrBillingLastName & "',N'" & pcStrBillingCompany & "','" & pcStrBillingVATID & "','" & pcStrBillingSSN & "','" & pcStrBillingPhone & "',N'" & pcStrBillingAddress & "','" & pcStrBillingPostalCode & "','" & pcStrBillingStateCode & "',N'" & pcStrBillingProvince & "',N'" & pcStrBillingCity & "','" & pcStrBillingCountryCode & "',N'" & pcStrBillingAddress2 & "','" & pcStrBillingFax & "','" & pcv_dateCustomerRegistration & "'," & tmpCustomerGuest & ");"
		set rsBA=connTemp.execute(query)		
		query="SELECT TOP 1 idCustomer,pcCust_Guest,[name], lastName, email FROM Customers WHERE email like '" & pcStrCustomerEmail & "' ORDER BY idCustomer DESC;"
		set rsBA=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsBA=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		if not rsBA.eof then
			session("idCustomer")=rsBA("idCustomer")
			session("CustomerGuest")=rsBA("pcCust_Guest")
			session("CustomerType")=0
			pcStrBillingFirstName = rsBA("name")
			pcStrBillingLastName = rsBA("lastName")
			pcStrCustomerEmail = rsBA("email")
			'// Send New Customer Emails
			pcv_strNoticeNewCust="1" '// Send to Admin
			If session("CustomerGuest")="0" Then
				pcv_strNewCustEmail="1" '// Send to Customer
			End If
%>
<!--#include file="adminNewCustEmail.asp"-->
<%
		else
			pcErrMsg=dictLanguage.Item(Session("language")&"_opc_68")
		end if
		set rsBA=nothing
		OKmsg="NEW"
	end if
end if
if pcErrMsg<>"" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_69")&"<br><ul>" & pcErrMsg & "</ul>"
	response.write pcErrMsg
else
	
	'Start Special Customer Fields
	tmpCustCFList=""
	pcSFCustFieldsExist=""
	
	query="SELECT pcCField_ID, pcCField_Name, pcCField_FieldType, pcCField_Value, pcCField_Length, pcCField_Maximum, pcCField_Required, pcCField_PricingCategories, pcCField_ShowOnReg, pcCField_ShowOnCheckout,'',pcCField_Description,0 FROM pcCustomerFields ORDER BY pcCField_Name ASC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closeDb()
		response.clear
		Call SetContentType()
		response.write "ERROR"
		response.End
	end if
	if not rs.eof then
		pcSFCustFieldsExist="YES"
		tmpCustCFList=rs.GetRows()
	end if
	set rs=nothing

	if pcSFCustFieldsExist="YES" AND Session("idCustomer")<>0 then
	pcArr=tmpCustCFList
	For k=0 to ubound(pcArr,2)
		pcArr(10,k)=""
		query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & Session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		if not rs.eof then
			pcArr(10,k)=rs("pcCFV_Value")
		end if
		set rs=nothing
	Next
	tmpCustCFList=pcArr
	end if

	if pcSFCustFieldsExist="YES" then
		pcArr=tmpCustCFList
		For k=0 to ubound(pcArr,2)						
			pcv_ShowField=0
			if pcArr(9,k)="1" then
				pcv_ShowField=1
			end if
			if (pcv_ShowField=1) AND (pcArr(7,k)="1") then
			if session("idCustomer")>"0" then
				query="SELECT pcCustFieldsPricingCats.idcustomerCategory FROM pcCustFieldsPricingCats INNER JOIN Customers ON (pcCustFieldsPricingCats.pcCField_ID=" & pcArr(0,k) & " AND pcCustFieldsPricingCats.idCustomerCategory=customers.idCustomerCategory) WHERE customers.idcustomer=" & session("idCustomer")
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)	
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
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
			pcArr(12,k)=pcv_ShowField
		Next
		tmpCustCFList=pcArr
	end if

	if pcSFCustFieldsExist="YES" then
	pcArr=tmpCustCFList
						
	For k=0 to ubound(pcArr,2)
		pcv_ShowField=pcArr(12,k)
		if pcv_ShowField=1 then
			if pcArr(5,k)>"0" then
				pcArr(10,k)=URLDecode(getUserInput(request("custfield" & pcArr(0,k)),pcArr(5,k)))
			else
				pcArr(10,k)=URLDecode(getUserInput(request("custfield" & pcArr(0,k)),0))
			end if
			if pcArr(10,k)<>"" then
				query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE idcustomer=" & session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if
				if NOT rs.eof then
					query="UPDATE pcCustomerFieldsValues SET pcCFV_Value=N'" & pcArr(10,k) & "' WHERE idcustomer=" & session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
				else
					query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & session("idCustomer") & "," & pcArr(0,k) & ",N'" & pcArr(10,k) & "');"
				end if
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if
				set rs=nothing
			else
				query="DELETE FROM pcCustomerFieldsValues WHERE idcustomer=" & session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if
				set rs=nothing
			end if
		end if
	Next
	
	end if
	
	tmpStr=""
	
	tmpIDRefer=URLDecode(getUserInput(request("IDRefer"),0))
	if tmpIDRefer<>"" then
		if not IsNumeric(tmpIDRefer) then
			tmpIDRefer=0
		end if
		tmpStr="IDRefer=" & tmpIDRefer
	end if


	if (AllowNews="1") AND (NewsCheckout="1") then
		if tmpStr<>"" then
			tmpStr=tmpStr & ","
		end if
		tmpRecvNews=URLDecode(getUserInput(request("CRecvNews"),0))
		if tmpRecvNews="" then
			tmpRecvNews=0
		end if
		if not IsNumeric(tmpRecvNews) then
			tmpRecvNews=0
		end if
		tmpStr=tmpStr & "RecvNews=" & tmpRecvNews
		Session("pcSFCRecvNews")=tmpRecvNews
	end if
	
	if Session("pcCustomerTermsAgreed")="1" then
		if tmpStr<>"" then
			tmpStr=tmpStr & ","
		end if
		tmpStr=tmpStr & "pcCust_AgreeTerms=1"
	end if
	
	if tmpStr<>"" then
		query="UPDATE Customers SET " & tmpStr & " WHERE idCustomer=" & session("idCustomer") & ";"
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		set rs=nothing
	end if
%>
<!--#include file="DBsv.asp"-->
<%
	call opendb
	query="UPDATE pcCustomerSessions SET pcCustSession_BillingStateCode='"&pcStrBillingStateCode&"', pcCustSession_BillingCity=N'"&pcStrBillingCity&"', pcCustSession_BillingProvince=N'"&pcStrBillingProvince&"', pcCustSession_BillingPostalCode='"&pcStrBillingPostalCode&"', pcCustSession_BillingCountryCode='"&pcStrBillingCountryCode&"', pcCustSession_ShippingResidential='"&pcIntShippingResidential&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closeDb()
		response.clear
		Call SetContentType()
		response.write "ERROR"
		response.End
	end if
	set rs=nothing
%>
<%
	if pcErrMsg<>"" then
		pcErrMsg=dictLanguage.Item(Session("language")&"_opc_69")&"<br><ul>" & pcErrMsg & "</ul>"
		response.write pcErrMsg
	end if

	response.write OKmsg

end if
'//////////////////////////////////////////////////////////////////////////
'// END: UPDATE BILLING
'//////////////////////////////////////////////////////////////////////////

call closeDb()
response.End()
%>
