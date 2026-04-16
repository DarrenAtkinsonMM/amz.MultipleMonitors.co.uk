<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%Dim pcStrPageName
pcStrPageName = "opc_AmzUpdShipAddr.asp"%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_AmazonHeader.asp" -->
<%
	tmpscXML=".3.0"

	AmzOrderId=getUserInput(request("id"),0)
	if AmzOrderId="" then
		response.write "Error: Cannot get Amazon Order ID#"
		response.End()
	end if
	session("AmzOrderID")=AmzOrderId
    session("AmzBillAgreementID")=""
	
	'Get Shipping Address
	QueryStr=""
	tmpTimeStamp=GenAmazonTimeStamp(UtcNow())
	QueryStr=QueryStr & "AWSAccessKeyId=" & pcAMZAccessKeyID
	QueryStr=QueryStr & "&Action=GetOrderReferenceDetails"
	QueryStr=QueryStr & "&AddressConsentToken=" & AmazonURLEnCode(session("Amz_access_token"))
	QueryStr=QueryStr & "&AmazonOrderReferenceId=" & session("AmzOrderID")
	QueryStr=QueryStr & "&SellerId=" & pcAMZSellerID
	QueryStr=QueryStr & "&SignatureMethod=HmacSHA256"
	QueryStr=QueryStr & "&SignatureVersion=2"
	QueryStr=QueryStr & "&Timestamp=" & tmpTimeStamp
	QueryStr=QueryStr & "&Version=2013-01-01"
	
	StringtoSign="POST" & vbLf
	StringtoSign=StringtoSign & pcAMZHost & vbLf
	StringtoSign=StringtoSign & pcAMZUI & vbLf
	StringtoSign=StringtoSign & QueryStr
	
	Set sha256 = GetObject( "script:" & Server.MapPath("sha256md5.txt") )
	StringtoSign=Server.URLEncode(sha256.b64_hmac_sha256(pcAMZSecretKey, StringtoSign))
	
	QueryStr=QueryStr & "&Signature=" & StringtoSign
	
	Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
	xml.open "POST", pcAMZEndPoint, False
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send QueryStr
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText
	'response.write strRetVal
	
	Set ReXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
	Set ReXML = xml.responseXML
	Set iRoot = ReXML.documentElement
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/AmazonOrderReferenceId")
	
	If (tmpNode is Nothing) OR (tmpNode.Text="") OR (tmpNode.Text<>session("AmzOrderID")) Then
		response.write "Error: Cannot Get Amazon Shipping Address"
		response.End()
	End if
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/DestinationType")
	If (tmpNode is Nothing) OR (tmpNode.Text="") Then
		response.write "Error: Cannot Get Amazon Shipping Address"
		response.End()
	End if
	desType=tmpNode.Text
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/AddressLine1")
	If (tmpNode is Nothing) Then
		response.write "Error: Cannot Get Amazon Shipping Address"
		response.End()
	End if
	If  (tmpNode.Text="") then
		response.write "Error: Cannot Get Amazon Shipping Address"
		response.End()
	End if
	AmzShipAddr1=tmpNode.Text
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/AddressLine2")
	AmzShipAddr2=""
	If Not (tmpNode is Nothing) then
		AmzShipAddr2=tmpNode.Text
	end if
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/City")
	AmzShipCity=""
	If Not (tmpNode is Nothing) then
		AmzShipCity=tmpNode.Text
	end if
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/StateOrRegion")
	AmzShipState=""
	If Not (tmpNode is Nothing) then
		AmzShipState=tmpNode.Text
	end if
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/CountryCode")
	AmzShipCountry=""
	If Not (tmpNode is Nothing) then
		AmzShipCountry=tmpNode.Text
	end if
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/PostalCode")
	AmzShipZip=""
	If Not (tmpNode is Nothing) then
		AmzShipZip=tmpNode.Text
	end if
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/Phone")
	AmzShipPhone=""
	If Not (tmpNode is Nothing) then
		AmzShipPhone=tmpNode.Text
	end if
	
	Set tmpNode=iRoot.selectSingleNode("GetOrderReferenceDetailsResult/OrderReferenceDetails/Destination/" & desType & "Destination/Name")
	AmzShipName=""
	If Not (tmpNode is Nothing) then
		AmzShipName=tmpNode.Text
	end if
	
	pcStrShippingNickName=AmzShipName
	if Instr(pcStrShippingNickName," ")>0 then
		pcStrShippingFirstName=Left(pcStrShippingNickName,Instr(pcStrShippingNickName," ")-1)
		pcStrShippingLastName=Mid(pcStrShippingNickName,Instr(pcStrShippingNickName," ")+1,len(pcStrShippingNickName))
	else
		pcStrShippingFirstName=pcStrShippingNickName
		pcStrShippingLastName=""
	end if
	pcStrShippingCompany=""
	pcStrShippingPhone=AmzShipPhone
	pcStrShippingEmail=""
	pcStrShippingAddress=replace(AmzShipAddr1, "'", "''")
	if pcStrShippingNickName="" then
		pcStrShippingNickName=replace(pcStrShippingAddress, "'", "''")
	end if
	pcStrShippingPostalCode=AmzShipZip
	pcStrShippingCountryCode=AmzShipCountry
	if (pcStrShippingCountryCode="US") OR (pcStrShippingCountryCode="CA") then
		pcStrShippingStateCode=AmzShipState
		pcStrShippingProvince=""
	else
		pcStrShippingStateCode=""
		pcStrShippingProvince=AmzShipState
	end if
	pcStrShippingCity=AmzShipCity
	pcStrShippingAddress2=replace(AmzShipAddr2, "'", "''")
	pcStrShippingFax=""
	pcIntShippingResidential="1"
	
	query="SELECT address FROM Customers WHERE idCustomer=" & session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		tmpBillAddr=rs("address")
		if IsNull(tmpBillAddr) OR tmpBillAddr="" then
			session("AmazonFirstTime")="1"
			query="UPDATE Customers SET address=N'" & pcStrShippingAddress & "',address2=N'" & pcStrShippingAddress2 & "', city=N'" & pcStrShippingCity & "', state=N'" & pcStrShippingProvince & "', stateCode='" & pcStrShippingStateCode & "', zip='" & pcStrShippingPostalCode & "', countryCode='" & pcStrShippingCountryCode & "', phone='" & pcStrShippingPhone & "',pcCust_AgreeTerms=1 WHERE idCustomer=" & session("idCustomer") & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	end if
	set rs=nothing
	
	query="SELECT idRecipient,idcustomer,recipient_FullName,recipient_NickName FROM recipients where idcustomer=" & session("idCustomer") & " AND recipient_FullName LIKE '" & pcStrShippingNickName & "';"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		IDReci=rs("idRecipient")
		session("pcShipOpt")=IDReci
		set rs=nothing
		query="UPDATE Recipients SET recipient_FullName=N'" & pcStrShippingNickName & "',recipient_NickName=N'" & pcStrShippingNickName & "',recipient_FirstName=N'" & pcStrShippingFirstName & "',recipient_LastName=N'" & pcStrShippingLastName & "',recipient_Email='" & pcStrShippingEmail & "',recipient_Phone='" & pcStrShippingPhone & "',recipient_Fax='" & pcStrShippingFax & "',recipient_Company=N'" & pcStrShippingCompany & "', recipient_Address=N'" & pcStrShippingAddress & "',recipient_Address2=N'" & pcStrShippingAddress2 & "',recipient_City=N'" & pcStrShippingCity & "',recipient_State=N'" & pcStrShippingProvince & "',recipient_StateCode='" & pcStrShippingStateCode & "',recipient_Zip='" & pcStrShippingPostalCode & "',recipient_CountryCode='" & pcStrShippingCountryCode & "'"
		if pcIntShippingResidential<>"" then
			query=query & ",Recipient_Residential=" & pcIntShippingResidential
		end if
		query=query & " WHERE idcustomer=" & session("idCustomer") & " AND idRecipient=" & IDReci & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
	else
		'// Add Recipient
		tmp1=""
		tmp2=""
		if pcIntShippingResidential<>"" then
			tmp1=",Recipient_Residential"
			tmp2="," & pcIntShippingResidential
		end if
		query="INSERT INTO Recipients (idcustomer,recipient_FullName,recipient_NickName,recipient_FirstName,recipient_LastName,recipient_Email,recipient_Phone,recipient_Fax,recipient_Company,recipient_Address,recipient_Address2,recipient_City,recipient_State,recipient_StateCode,recipient_Zip,recipient_CountryCode" & tmp1 & ") VALUES (" & session("idCustomer") & ",N'" & pcStrShippingNickName & "',N'" & pcStrShippingNickName & "',N'" & pcStrShippingFirstName & "',N'" & pcStrShippingLastName & "','" & pcStrShippingEmail & "','" & pcStrShippingPhone & "','" & pcStrShippingFax & "',N'" & pcStrShippingCompany & "',N'" & pcStrShippingAddress & "',N'" & pcStrShippingAddress2 & "',N'" & pcStrShippingCity & "',N'" & pcStrShippingProvince & "','" & pcStrShippingStateCode & "','" & pcStrShippingPostalCode & "','" & pcStrShippingCountryCode & "'" & tmp2 & ");"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="SELECT TOP 1 idRecipient FROM Recipients WHERE idcustomer=" & session("idCustomer") & " ORDER BY idRecipient DESC;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			IDReci=rs("idRecipient")
			session("pcShipOpt")=IDReci
		end if
		set rs=nothing
	end if
	
	%>
	<!--#include file="DBsv.asp"-->
	<%
		
	pcShipOptA="1"
	
	query="SELECT payTypes.idPayment FROM payTypes WHERE (payTypes.active = - 1) AND (payTypes.gwCode = 88);"
	set rs=connTemp.execute(query)
	if not rs.eof then
		AmzidPayment=rs("idPayment")
	end if
	set rs=nothing

	query="UPDATE pcCustomerSessions SET pcCustSession_IdPayment=" & AmzidPayment& ", pcCustSession_ShowShipAddr=" & pcShipOptA & ", idCustomer="&session("idCustomer")&", pcCustSession_ShippingNickName=N'"&pcStrShippingNickName&"', pcCustSession_ShippingFirstName=N'"&pcStrShippingFirstName&"', pcCustSession_ShippingLastName=N'"&pcStrShippingLastName&"', pcCustSession_ShippingCompany=N'"&pcStrShippingCompany&"', pcCustSession_ShippingPhone='"&pcStrShippingPhone&"', pcCustSession_ShippingAddress=N'"&pcStrShippingAddress&"', pcCustSession_ShippingPostalCode='"&pcStrShippingPostalCode&"', pcCustSession_ShippingStateCode='"&pcStrShippingStateCode&"', pcCustSession_ShippingProvince=N'"&pcStrShippingProvince&"', pcCustSession_ShippingCity=N'"&pcStrShippingCity&"', pcCustSession_ShippingCountryCode='"&pcStrShippingCountryCode&"', pcCustSession_ShippingAddress2=N'"&pcStrShippingAddress2&"', pcCustSession_ShippingResidential='"&pcIntShippingResidential&"', pcCustSession_ShippingFax='"&pcStrShippingFax&"', pcCustSession_ShippingEmail='"&pcStrShippingEmail&"', pcCustSession_TF1='"&pcSFTF1&"', pcCustSession_DF1='"&pcSFDF1&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	if session("AmazonFirstTime")="1" then
		query="UPDATE pcCustomerSessions SET pcCustSession_BillingStateCode='"&pcStrShippingStateCode&"', pcCustSession_BillingCity=N'"&pcStrShippingCity&"', pcCustSession_BillingProvince=N'"&pcStrShippingProvince&"', pcCustSession_BillingPostalCode='"&pcStrShippingPostalCode&"', pcCustSession_BillingCountryCode='"&pcStrShippingCountryCode&"', pcCustSession_ShippingResidential='"&pcIntShippingResidential&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	
	If DeliveryZip = "1" Then
		query="SELECT * from zipcodevalidation WHERE zipcode='" &pcStrShippingPostalCode& "'"
		set rsZipCodeObj=server.CreateObject("ADODB.RecordSet")
		set rsZipCodeObj=conntemp.execute(query)
		if rsZipCodeObj.eof then
			set rsZipCodeObj=nothing		
			response.write "Error: "&dictLanguage.Item(Session("language")&"_Custmoda_23")
			call closedb()
			response.End()
		end if
	End If
	
	response.write "OK|*|" & session("pcShipOpt")
	session("AmazonFirstTime")=""
	call closeDb()
	response.End()
%>	

<%
Function AmazonURLEnCode(tmpStr)
Dim tmp1

tmp1=tmpStr
tmp1=replace(Server.URLEncode(tmp1),"%2E",".")
tmp1=replace(tmp1,"%5F","_")
tmp1=replace(tmp1,"%2D","-")
tmp1=replace(tmp1,"%7E","~")
tmp1=replace(tmp1,"+","%20")
AmazonURLEnCode=tmp1

End Function

Function AddZero(tmpStr)
	if Clng(tmpStr)<10 then
		AddZero="0" & tmpStr
	else
		AddZero=tmpStr
	end if
End Function

Function GenAmazonTimeStamp(tmpDate)
Dim tmp1
	tmp1=Year(tmpDate) & "-" & AddZero(Month(tmpDate)) & "-" & AddZero(Day(tmpDate)) & "T" & AddZero(Hour(tmpDate)) & ":" & AddZero(Minute(tmpDate)) & ":" & AddZero(Second(tmpDate)) & "Z"
	GenAmazonTimeStamp=AmazonURLEnCode(tmp1)
End Function


Function UtcNow()
UtcNow = serverdate.toUTCString()
UtcNow = CDate(Replace(Right(UtcNow, Len(UtcNow) - Instr(UtcNow, ",")), "UTC", ""))
End Function
%>
<script language="JScript" runat="server">
var serverdate=new Date();
</script>
