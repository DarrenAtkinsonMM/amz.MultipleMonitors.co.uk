
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/pcUSPSClass.asp"-->

<%
If USPS_AddressValidation = 1 AND request("CountryCode") = "US" Then
	Dim objUSPSClass, srvUSPSXmlHttp, USPS_postdata, USPS_result, xmlDoc
	
	pcv_PostalCode = request("zip")
	
	set objUSPSClass = New pcUSPSClass
	
	objUSPSClass.NewXMLTransaction "CityStateLookup", "CityStateLookupRequest", USPS_Id
	strXMLClosingTag="CityStateLookupRequest"
	
	objUSPSClass.WriteParent "ZipCode", ""
		objUSPSClass.AddNewNode "Zip5", pcv_PostalCode, 1
	objUSPSClass.WriteParent "ZipCode", "/"
	
	ObjUSPSClass.WriteParent strXMLClosingTag, "/"
	
	USPS_postdata=replace(USPS_postdata, "&XML", "andXML")
	USPS_postdata=replace(USPS_postdata, "&", "and")
	USPS_postdata=replace(USPS_postdata, "andamp;", "and")
	USPS_postdata=replace(USPS_postdata, "andXML", "&XML")
	
	call objUSPSClass.SendXMLRequest(USPS_postdata, USPS_AccessLicense)
	
	Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
	if xmlDoc.loadXML(USPS_result) then
		if xmlDoc.selectSingleNode("CityStateLookupResponse/ZipCode/Error/Number") Is Nothing then
			Response.Write "VALID"
		else
			Response.Write getUserInput(xmlDoc.selectSingleNode("CityStateLookupResponse/ZipCode/Error/Description").text, 0)
		end if
	end if
else
	Response.Write "VALID"
End if
%>