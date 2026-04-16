<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard" %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<style>
	#pcCPmain ul {
		margin: 0px;
		padding: 0;
	}

	#pcCPmain ul li {
		margin: 0px;
	}

	div.menu ul {
	text-align:left;
	margin:0 0 0 60px;
	padding:0;
	cursor:pointer;
	}

	div.menu ul li {
	display:inline;
	list-style:none;
	margin:0 0.3em;
	cursor:pointer;
	font-size:12px;
	}

	div.menu ul li a {
	position:relative;
	z-index:0;
	font-weight:bold;
	border:solid 2px #e1e1e1;
	border-bottom-width:0;
	padding:0.3em;
	background-color:#ffffcc;
	color:black;
	text-decoration:none;
	cursor:pointer;
	font-size:12px;
	}

	div.menu ul li a.current {
	background-color:#F5F5F5;
	border:solid 2px #CCCCCC;
	border-bottom-width:0;
	position:relative;z-index:2;
	cursor:pointer;
	font-size:12px;
	}

	div.menu ul li a.current:hover {
	background-color:#F5F5F5;
	cursor:pointer;
	font-size:12px;
	}

	div.menu ul li a:hover {
	z-index:2;
	background-color:#F5F5F5;
	border-bottom:0;
	cursor:pointer;
	font-size:12px;
	}

	div.menu a span {display:none;}

	div.menu a:hover span {
		display:block;
		position:absolute;
		top:2.3em;
		background-color:#F5F5F5;
		border-bottom:thin dotted gray;
		border-top:thin dotted gray;
		font-weight:normal;
		left:0;
		padding:1px 2px;
		cursor:pointer;
		font-size:12px;
	}

	div.panes {
		padding: 1em;
		border: dashed 2px #CCCCCC;
		background-color: #F5F5F5;
		display: none;
		text-align:left;
		position:relative;z-index:1;
		margin-top:0.15em;
	}

	div.navbox {
		display: table-cell;
		padding: .3em;
		font-size: 12px;
		font-weight:bold;
		border: solid 2px #CCCCCC;
		background-color: #F5F5F5;
		text-align:left;
	}

	div.NavOrderClass1 {
		padding: 0.2em;
		font-size:12px;
		font-weight:bold;
		background-color: #B0E0E6;
		display: none;
		text-align:left;
		margin-top:0em;
		border-bottom: 1px solid #CCCCCC;
		border-left: 1px solid #CCCCCC;
		border-right: 1px solid #CCCCCC;
	}

	div.NavOrderClass2 {
		padding: 0.2em;
		font-size:12px;
		font-weight:bold;
		background-color: #FFFFFF;
		display: none;
		text-align:left;
		margin-top:0em;
		border-bottom: 1px solid #CCCCCC;
		border-left: 1px solid #CCCCCC;
		border-right: 1px solid #CCCCCC;
	}
</style>

<%
  
Public Function CamelCase(str)
  Dim arr, i
  arr = Split(str, " ")
  For i = LBound(arr) To UBound(arr)
    word = arr(i)
    arr(i) = UCase(Left(word, 1)) & Mid(word, 2)
  Next
  CamelCase = Join(arr, " ")
End Function
	
Sub FedEx_CurrencyControl(currencyAmountName, currencyTypeName, isRequired)
  CurrencyTypes = Array( _
    Array("US Dollars (USD)", "USD"), _
    Array("Canadian Dollars (CAD)", "CAD") _
  )

  %>
    <% '// Currency Value %>
    <input id="<%= currencyAmountName %>" name="<%= currencyAmountName %>" type="text" value="<%=pcf_FillFormField(currencyAmountName, isRequired)%>" title="Amount" placeholder="Amount" size="10">
	  <%pcs_RequiredImageTag currencyAmountName, isRequired %> &nbsp;
                    
    <% '// Currency Type %>
    <select id="<%= currencyTypeName %>" name="<%= currencyTypeName %>" title="Currency">
      <% If Not isRequired Then %>
        <option value="">Select Currency</option>
      <% End If %>
      <%
        For Each CurrencyType In CurrencyTypes
          %>
            <option value="<%= CurrencyType(1) %>" <%= pcf_SelectOption(currencyTypeName, CurrencyType(1)) %>><%= CurrencyType(0) %></option>
          <%
        Next
      %>
    </select>
	  <%pcs_RequiredImageTag currencyTypeName, isRequired %>
  <%
End Sub

Sub FedEx_WeightControl(weightValueName, weightUnitsName, isRequired)
  WeightUnits = Array( _
    Array("Pounds", "LB"), _
    Array("Kilograms", "KG") _
  )

  %>
    <% '// Weight Value %>
    <input id="<%= weightValueName %>" name="<%= weightValueName %>" type="text" value="<%=pcf_FillFormField(weightValueName, isRequired)%>" title="Weight" placeholder="Weight" size="5">
	  <%pcs_RequiredImageTag weightValueName, isRequired %> &nbsp;
                    
    <% '// Weight Units %>
    <select id="<%= weightUnitsName %>" name="<%= weightUnitsName %>" title="Units">
      <% If Not isRequired Then %>
        <option value="">Select Units</option>
      <% End If %>
      <%
        For Each Unit In WeightUnits
          %>
            <option value="<%= Unit(1) %>" <%= pcf_SelectOption(weightUnitsName, Unit(1)) %>><%= Unit(0) %></option>
          <%
        Next
      %>
    </select>
	  <%pcs_RequiredImageTag weightUnitsName, isRequired %>
  <%
End Sub



Sub FedEx_DimensionsControl(lengthName, widthName, heightName, unitsName, isRequired)
  DimensionsUnits = Array( _
    Array("Inches", "IN"), _
    Array("Centimeters", "CM") _
  )

  %>
    <input type="text" id="<%= lengthName %>" name="<%= lengthName %>" value="<%=pcf_FillFormField(lengthName, false)%>" title="Length" placeholder="Length" size="5">
    <%pcs_RequiredImageTag lengthName, isRequired %> &nbsp;
    <input type="text" id="<%= widthName %>"  name="<%= widthName %>" value="<%=pcf_FillFormField(widthName, false)%>" title="Width" placeholder="Width" size="5">
    <%pcs_RequiredImageTag widthName, isRequired %> &nbsp;
    <input type="text" id="<%= heightName %>" name="<%= heightName %>" value="<%=pcf_FillFormField(heightName, false)%>" title="Height" placeholder="Height" size="5">
    <%pcs_RequiredImageTag heightName, isRequired %> &nbsp;
                    
    <select id="<%= unitsName %>" name="<%= unitsName %>" title="Units">
      <option value="">Select Units</option>
      <%
        For Each Unit In DimensionsUnits
          %>
            <option value="<%= Unit(1) %>" <%= pcf_SelectOption(unitsName, Unit(1)) %>><%= Unit(0) %></option>
          <%
        Next
      %>
    </select>
    <%pcs_RequiredImageTag unitsName, isRequired %><br />
  <%
End Sub

Sub FedEx_RequestedShippingDocumentType(controlName)
  ShippingDocumentTypes = Array( _
    "CERTIFICATE_OF_ORIGIN", _
    "COMMERCIAL_INVOICE", _
    "CUSTOMER_SPECIFIED_LABELS", _
    "CUSTOM_PACKAGE_DOCUMENT", _
    "CUSTOM_SHIPMENT_DOCUMENT", _
    "DANGEROUS_GOODS_SHIPPERS_DECLARATION", _
    "EXPORT_DECLARATION", _
    "FREIGHT_ADDRESS_LABEL", _
    "GENERAL_AGENCY_AGREEMENT", _
    "LABEL", _
    "NAFTA_CERTIFICATE_OF_ORIGIN", _
    "OP_900", _
    "PRO_FORMA_INVOICE", _
    "RETURN_INSTRUCTIONS" _
    )
  %>
    <select name="<%= controlName %>" id="<%= controlName %>">
      <option value="" <%=pcf_SelectOption(controlName,"")%>>Select Option</option>
      <% For Each DocumentType In ShippingDocumentTypes %>
        <option value="<%= DocumentType %>" <%=pcf_SelectOption(controlName, DocumentType)%>><%= CamelCase(LCase(Replace(DocumentType, "_", " "))) %></option>
      <% Next %>
    </select>
    <%pcs_RequiredImageTag controlName, false%>
  <%
End Sub

%>

<%
Dim objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strCarrierCode
Dim pcv_strTrackingNumber, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim fedex_xmlPrefix, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_URL, pcv_strErrorMsg, pcv_strAction

'// Define objects used to create and send Google Checkout Order Processing API requests
Dim xmlRequest
Dim xmlResponse
Dim attrGoogleOrderNumber
Dim elemAmount
Dim elemReason
Dim elemComment
Dim elemCarrier
Dim elemTrackingNumber
Dim elemMessage
Dim elemSendEmail
Dim elemMerchantOrderNumber
Dim transmitResponse

function fnStripPhone(PhoneField)
	PhoneField=replace(PhoneField," ","")
	PhoneField=replace(PhoneField,"-","")
	PhoneField=replace(PhoneField,".","")
	PhoneField=replace(PhoneField,"(","")
	PhoneField=replace(PhoneField,")","")
	fnStripPhone = PhoneField
end function

function sanitizeField(UserInput)
	if UserInput<>"" AND isNULL(UserInput)=False then
		UserInput=replace(UserInput,"&"," ")
	end if
	sanitizeField=UserInput
end function

'// If there is no Tracking Number, provide a random number for the log file.
'// This is due to the fact there could be multiple errors for the same ID.
function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// OPEN DATABASE


'// PACKAGE COUNT
pcv_strPackageCount=request("PackageCount")
pcv_strSessionPackageCount=Session("pcAdminPackageCount")
if pcv_strSessionPackageCount="" OR len(pcv_strPackageCount)>0 then
	pcPackageCount=pcv_strPackageCount
	Session("pcAdminPackageCount")=pcPackageCount
else
	pcPackageCount=pcv_strSessionPackageCount
end if
if pcPackageCount="" then
	pcPackageCount=1
end if
pcArraySize = (pcPackageCount -1)

'// GET ORDER ID
pcv_strOrderID=request("idorder")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_strOrderID)>0 then
	pcv_intOrderID=pcv_strOrderID
	Session("pcAdminOrderID")=pcv_intOrderID
else
	pcv_intOrderID=pcv_strSessionOrderID
end if

'// REDIRECT
if pcv_intOrderID="" then
	call closeDb()
response.redirect "menu.asp"
end if

query="SELECT orders.pcOrd_GoogleIDOrder FROM orders WHERE idOrder="& pcv_intOrderID
set rs=server.CreateObject("ADODB.RecordSet")
Set rs=conntemp.execute(query)
if Not rs.eof then
	pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder") '// determine if this is a google order
end if
set rs=nothing

'// ITEM COUNT
pcv_count=Request("count")
if pcv_count="" then
	pcv_count=0
end if

'// CREATE THE ARRAY
Dim pcLocalArray()

'// SIZE THE ARRAY
ReDim pcLocalArray(pcArraySize)

'// POPULATE THE ARRAY
if request.form("submit")<>"" OR request.form("submit1")<>"" then
	For xPackageCount=0 to pcArraySize
		pcLocalArray(xPackageCount) = Request("pcAdminPrdList" & (xPackageCount+1))
	Next
else
	if Session("pcGlobalArray")<>"" then
		pcArray_TmpGlobalReturn = split(Session("pcGlobalArray"), chr(124))
		For xPackageCount = LBound(pcArray_TmpGlobalReturn) TO UBound(pcArray_TmpGlobalReturn)
			pcLocalArray(xPackageCount) = pcArray_TmpGlobalReturn(xPackageCount)
		Next
	end if
end if

'// UPDATE ARRAY
If pcv_count <> 0 Then
	For i=1 to pcv_count
		if request("C" & i)="1" then
			pcv_strTmpList=pcv_strTmpList & request("IDPrd" & i) & ","
		end if
	Next
	pcLocalArray((pcPackageCount-1)) = pcv_strTmpList
End If

'// CONVERT ARRAY TO SESSIONS
For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
	Session("pcAdminPrdList"&(xArrayCount+1)) = pcLocalArray(xArrayCount)
Next

'// ARRAY TO PASS TO OTHER PAGES
pcv_strItemsList = join(pcLocalArray, chr(124))

'// SESSION FOR REDIRECTS
Session("pcGlobalArray") = pcv_strItemsList

'////////////////////////////////////////////
'// END: PRODUCT ID LIST FOUR
'////////////////////////////////////////////

'// PAGE NAME
pcPageName="FedExWS_ManageShipmentsRequest.asp"
ErrPageName="FedExWS_ManageShipmentsRequest.asp"

'// ACTION
pcv_strAction = request("Action")

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExWSClass

'// FEDEX CREDENTIALS
query = "SELECT ShipmentTypes.userID, ShipmentTypes.password, ShipmentTypes.AccessLicense, ShipmentTypes.FedExKey, ShipmentTypes.FedExPwd "
query = query & "FROM ShipmentTypes "
query = query & "WHERE (((ShipmentTypes.idShipment)=9));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	FedExAccountNumber=rs("userID")
	FedExMeterNumber=rs("password")
	pcv_strEnvironment=rs("AccessLicense")
	FedExkey=rs("FedExKey")
	FedExPassword=rs("FedExPwd")
end if
set rs=nothing

if session("pcAdminPayorAccountNumber")&"" = "" then
	session("pcAdminPayorAccountNumber") = FedExAccountNumber
end if

if session("pcAdminDutiesAccountNumber")&"" = "" then
	session("pcAdminDutiesAccountNumber") = FedExAccountNumber
end if

If Session("pcAdminCODAccountNumber") = "Please use shipper account number" Then
	Session("pcAdminCODAccountNumber") = Session("pcAdminPayorAccountNumber")
End If

'// DATE FUNCTION
function ShowDateFrmt(x)
	ShowDateFrmt = x
end function
			
DGMaxContainers = 5

'// SELECT DATA SET
' >>> Tables: pcPackageInfo
query="SELECT orders.idCustomer, orders.ShippingFullname, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, "
query = query & "orders.shippingCountryCode, orders.shippingZip, orders.shippingCompany, orders.shippingAddress2, orders.pcOrd_shippingPhone, orders.pcOrd_ShippingEmail, "
query = query & "orders.SRF, orders.shipmentDetails, orders.OrdShipType, orders.OrdPackageNum, orders.pcOrd_ShipWeight, orders.pcOrd_ShippingFax "
query = query & "FROM orders "
query = query & "WHERE orders.idOrder=" & pcv_intOrderID &" "

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then
	Dim pidorder, pidcustomer, pshippingAddress, pshippingCity, pshippingStateCode, pshippingState, pshippingZip, pshippingPhone, pshippingCountryCode, pshippingCompany, pshippingAddress2, pShippingEmail, SRF

	'// ORDER INFO
	pidorder=scpre+int(pcv_intOrderID)
	pidcustomer=rs("idcustomer")

	'// DESTINATION ADDRESS
	pShippingFullname=rs("ShippingFullname")
	pshippingAddress=rs("shippingAddress")
	pshippingCity=rs("shippingCity")
	pshippingStateCode=rs("shippingStateCode")
	pshippingState=rs("shippingState")
	pshippingZip=rs("shippingZip")
	pshippingPhone=rs("pcOrd_shippingPhone")
	pshippingCountryCode=rs("shippingCountryCode")
	pshippingCompany=rs("shippingCompany")
	pshippingAddress2=rs("shippingAddress2")
	pShippingEmail=rs("pcOrd_ShippingEmail")
	pSRF=rs("SRF")
	pshipmentDetails=rs("shipmentDetails")
	pOrdShipType=rs("ordShipType")
	pOrdPackageNum=rs("ordPackageNum")
	pcOrd_ShipWeight=rs("pcOrd_ShipWeight")
	pshippingFax=rs("pcOrd_ShippingFax")
end if
set rs=nothing


' Shipment
If pSRF="1" then
	pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_b")
else
	shipping=split(pshipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
			varShip="0"
			pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
		else
			Shipper=shipping(0)
			Service=shipping(1)
			Postage=trim(shipping(2))
			if ubound(shipping)=3 then
				serviceHandlingFee=trim(shipping(3))
				if NOT isNumeric(serviceHandlingFee) then
					serviceHandlingFee=0
				end if
			else
				serviceHandlingFee=0
			end if
		end if
	else
		varShip="0"
		pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	end if
end if

'// SHIPPER EMAIL
query="SELECT * FROM emailsettings WHERE id=1;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number <> 0 then
	set rs = nothing
	
end if
ownerEmail=rs("ownerEmail")
set rs = nothing

'// SHIP TIME
if Session("pcAdminShipTime") = "" then
	todaysDate=now()
	pcv_strShipTime = FormatDateTime(todaysDate,4) & ":00"
	Session("pcAdminShipTime") = pcv_strShipTime
end if


'// Residential Delivery
if Session("pcAdminResidentialDelivery")&"" = "" then
	if pOrdShipType=0 then
		pcv_strResidentialDelivery = "true"
	else
		pcv_strResidentialDelivery = "false"
	end if
	Session("pcAdminResidentialDelivery") = pcv_strResidentialDelivery
end if

'// SHIP CONSTANTS

if Session("pcAdminSMHubID") = "" then
	pcv_strHubId = FEDEXWS_SMHUBID
	Session("pcAdminSMHubID") = pcv_strHubId
end if

'// DropType
if Session("pcAdminDropoffType") = "" then
	pcv_strDropoffType = FEDEXWS_DROPOFF_TYPE
	Session("pcAdminDropoffType") = pcv_strDropoffType
end if

'//Rate Request Type
if Session("pcAdminRateRequestType") = "" then
	pcv_strRateRequestType = FEDEXWS_LISTRATE
	if pcv_strRateRequestType="0" then
		pcv_strRateRequestType = "ACCOUNT"
	end if
	if pcv_strRateRequestType="-1" then
		pcv_strRateRequestType = "LIST"
	end if
	if pcv_strRateRequestType="-2" then
		pcv_strRateRequestType = "PREFERRED"
	end if
	Session("pcAdminRateRequestType") = pcv_strRateRequestType
end if
'// PackageType
if Session("pcAdminPackaging1") = "" then
	pcv_strPackaging = FEDEXWS_FEDEX_PACKAGE
	Session("pcAdminPackaging1") = pcv_strPackaging
	Session("pcAdminPackaging2") = pcv_strPackaging
	Session("pcAdminPackaging3") = pcv_strPackaging
	Session("pcAdminPackaging4") = pcv_strPackaging
end if

'// L
if Session("pcAdminLength1") = "" then
	pcv_strLength = FEDEXWS_LENGTH
	Session("pcAdminLength1") = pcv_strLength
	Session("pcAdminLength2") = pcv_strLength
	Session("pcAdminLength3") = pcv_strLength
	Session("pcAdminLength4") = pcv_strLength
end if

'// W
if Session("pcAdminWidth1") = "" then
	pcv_strWidth = FEDEXWS_WIDTH
	Session("pcAdminWidth1") = pcv_strWidth
	Session("pcAdminWidth2") = pcv_strWidth
	Session("pcAdminWidth3") = pcv_strWidth
	Session("pcAdminWidth4") = pcv_strWidth
end if

'// H
if Session("pcAdminHeight1") = "" then
	pcv_strHeight = FEDEXWS_HEIGHT
	Session("pcAdminHeight1") = pcv_strHeight
	Session("pcAdminHeight2") = pcv_strHeight
	Session("pcAdminHeight3") = pcv_strHeight
	Session("pcAdminHeight4") = pcv_strHeight
end if

'// U
if Session("pcAdminUnits1") = "" then
	pcv_strUnits = FEDEXWS_DIM_UNIT
	Session("pcAdminUnits1") = pcv_strUnits
	Session("pcAdminUnits2") = pcv_strUnits
	Session("pcAdminUnits3") = pcv_strUnits
	Session("pcAdminUnits4") = pcv_strUnits
end if

if Session("pcAdminWeightUnits1") = "" then
	pcv_strWeightUnits = scShipFromWeightUnit
  If pcv_strWeightUnits = "LBS" Then
    pcv_strWeightUnits = "LB"
  End If
	Session("pcAdminWeightUnits1") = pcv_strWeightUnits
end if

'// SHIPPER INFO
if Session("pcAdminOriginPersonName") = "" then
	pcv_strOriginPersonName = scOriginPersonName
	Session("pcAdminOriginPersonName") = pcv_strOriginPersonName
end if

if Session("pcAdminOriginCompanyName") = "" then
	pcv_strOriginCompanyName = scShipFromName
	Session("pcAdminOriginCompanyName") = pcv_strOriginCompanyName
end if

if Session("pcAdminOriginDepartment") = "" then
	pcv_strOriginDepartment = scOriginDepartment
	Session("pcAdminOriginDepartment") = pcv_strOriginDepartment
end if

if Session("pcAdminOriginPhoneNumber") = "" then
	pcv_strOriginPhoneNumber = scOriginPhoneNumber
	Session("pcAdminOriginPhoneNumber") = pcv_strOriginPhoneNumber
end if
if Session("pcAdminOriginPagerNumber") = "" then
	pcv_strOriginPagerNumber = scOriginPagerNumber
	Session("pcAdminOriginPagerNumber") = pcv_strOriginPagerNumber
end if
if Session("pcAdminOriginFaxNumber") = "" then
	pcv_strOriginFaxNumber = scOriginFaxNumber
	Session("pcAdminOriginFaxNumber") = pcv_strOriginFaxNumber
end if
if Session("pcAdminOriginEmailAddress") = "" then
	pcv_strOriginEmailAddress = ownerEmail
	Session("pcAdminOriginEmailAddress") = pcv_strOriginEmailAddress
	If Session("pcAdminNotificationShipperEmail")&""="" Then Session("pcAdminNotificationShipperEmail") = pcv_strOriginEmailAddress
end if

'// ORIGIN ADDRESS
if Session("pcAdminOriginLine1") = "" then
	pcv_strOriginLine1 = scShipFromAddress1
	Session("pcAdminOriginLine1") = pcv_strOriginLine1
end if
if Session("pcAdminOriginLine2") = "" then
	pcv_strOriginLine2 = scShipFromAddress2
	Session("pcAdminOriginLine2") = pcv_strOriginLine2
end if
if Session("pcAdminOriginCity") = "" then
	pcv_strOriginCity = scShipFromCity
	Session("pcAdminOriginCity") = pcv_strOriginCity
end if
if Session("pcAdminOriginStateOrProvinceCode") = "" then
	pcv_strOriginStateOrProvinceCode = scShipFromState
	Session("pcAdminOriginStateOrProvinceCode") = pcv_strOriginStateOrProvinceCode
end if
if Session("pcAdminOriginPostalCode") = "" then
	pcv_strOriginPostalCode = scShipFromPostalCode
	Session("pcAdminOriginPostalCode") = pcv_strOriginPostalCode
end if
if Session("pcAdminOriginCountryCode") = "" then
	pcv_strOriginCountryCode = scShipFromPostalCountry
	Session("pcAdminOriginCountryCode") = pcv_strOriginCountryCode
end if

'// RECIPIENT
if Session("pcAdminRecipPersonName") = "" then
	pcv_strRecipPersonName = pShippingFullname
	Session("pcAdminRecipPersonName") = pcv_strRecipPersonName
end if
if Session("pcAdminRecipCompanyName") = "" then
	pcv_strRecipCompanyName = pshippingCompany
	Session("pcAdminRecipCompanyName") = pcv_strRecipCompanyName
end if

if Session("pcAdminRecipPhoneNumber") = "" then
	pcv_strRecipPhoneNumber = pshippingPhone
	Session("pcAdminRecipPhoneNumber") = pcv_strRecipPhoneNumber
end if

if Session("pcAdminRecipPhoneNumber") = "" then
	Session("pcAdminDeliveryPhone") = Session("pcAdminRecipPhoneNumber")
end if

if Session("pcAdminRecipFaxNumber") = "" then
	pcv_strRecipFaxNumber = pshippingFax
	Session("pcAdminRecipFaxNumber") = pcv_strRecipFaxNumber
end if

if Session("pcAdminRecipEmailAddress") = "" then
	pcv_strRecipEmailAddress = pShippingEmail
	Session("pcAdminRecipEmailAddress") = pcv_strRecipEmailAddress
	If Session("pcAdminNotificationRecipientEmail")&""="" Then Session("pcAdminNotificationRecipientEmail") = pcv_strRecipEmailAddress
end if

'   >>> Origin Address Conditionals
'// Use the Request object to toggle State (based of Country selection)
isRequiredState =  true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	isRequiredState=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
isRequiredProvince = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	isRequiredProvince=pcv_strProvinceCodeRequired
end if

'// DESTINATION ADDRESS
if Session("pcAdminRecipLine1") = "" then
	pcv_strRecipLine1 = pshippingAddress
	Session("pcAdminRecipLine1") = pcv_strRecipLine1
end if
if Session("pcAdminRecipLine2") = "" then
	pcv_strRecipLine2 = pshippingAddress2
	Session("pcAdminRecipLine2") = pcv_strRecipLine2
end if
if Session("pcAdminRecipCity") = "" then
	pcv_strRecipCity = pshippingCity
	Session("pcAdminRecipCity") = pcv_strRecipCity
end if
if Session("pcAdminRecipStateOrProvinceCode") = "" then
	pcv_strRecipStateOrProvinceCode = pshippingStateCode
	Session("pcAdminRecipStateOrProvinceCode") = pcv_strRecipStateOrProvinceCode
end if
if Session("pcAdminRecipPostalCode") = "" then
	pcv_strRecipPostalCode = pshippingZip
	Session("pcAdminRecipPostalCode") = pcv_strRecipPostalCode
end if



'//try revalidate

'if Session("pcAdminRecipCountryCode") = "" Or Request.Form.Count > 0 then
'	pcs_ValidateTextField	"RecipCountryCode", true, 2
	
'	if Session("pcAdminRecipCountryCode") = "" then
'		pcv_strRecipCountryCode = pshippingCountryCode
'		Session("pcAdminRecipCountryCode") = pcv_strRecipCountryCode
'	end if
'end if

If Session("pcAdminOriginCountryCode")&""="" Then pcs_ValidateTextField	"OriginCountryCode", false, 2
If Session("pcAdminRecipCountryCode")&""="" Then 	pcs_ValidateTextField	"RecipCountryCode", false, 2

if Session("pcAdminRecipCountryCode")&"" = "" then
	pcv_strRecipCountryCode = pshippingCountryCode
	Session("pcAdminRecipCountryCode") = pcv_strRecipCountryCode
end if

pcv_FedExShipmentID = 9

'Automatically load International area
pcv_FedExInternational = false
Session("pcAdminbInternational") = "0"
If Session("pcAdminOriginCountryCode")&""<>"" And Session("pcAdminRecipCountryCode")&""<>"" And _
   Session("pcAdminOriginCountryCode")<>Session("pcAdminRecipCountryCode") Then
	pcv_FedExInternational = true
	Session("pcAdminbInternational") = "1"
End If

If pcv_FedExInternational Then
	isRequiredCVAmount = true
	isRequiredCVCurrency = true
	isRequiredNumberOfPieces = true
	isRequiredDescription = true
	isRequiredCountryOfManufacture = true
	isRequiredCommodityWeight = true
	isRequiredCommodityQuantity = true
	isRequiredCommodityQuantityUnits = true
	isRequiredCommodityUnitPrice = true
	isRequiredDutiesAccountNumber = true
	isRequiredDutiesCountryCode = true
Else
	isRequiredCVAmount = false
	isRequiredCVCurrency = false
	isRequiredNumberOfPieces = false
	isRequiredDescription = false
	isRequiredCountryOfManufacture = false
	isRequiredCommodityWeight = false
	isRequiredCommodityQuantity = false
	isRequiredCommodityQuantityUnits = false
	isRequiredCommodityUnitPrice = false
	isRequiredDutiesAccountNumber = false
	isRequiredDutiesCountryCode = false
End If

isRequiredDutiesPersonName = false

'Automatically load Freight area
pcv_FedExFreight = false
If Session("pcAdminService1")="FEDEX_FREIGHT_PRIORITY" Or Session("pcAdminService1")="FEDEX_FREIGHT_ECONOMY" Then
	pcv_FedExFreight = true
	Session("pcAdminbFreight") = "1"
End If

If pcv_FedExFreight Then
	isRequiredFreightAccountNumber = true
	isRequiredShipmentRoleType = true
	isRequiredTotalHandlingUnits = true
	
	isRequiredFreightContactPersonName = true
	isRequiredFreightContactCompanyName = true
	isRequiredFreightContactStreetLines = true
	isRequiredFreightContactCity = true
	isRequiredFreightContactStateCode = true
	isRequiredFreightContactPostalCode = true
	isRequiredFreightContactCountryCode = true
	
	isRequiredFreightLIPackaging = true
	isRequiredFreightLIPieces = true
	isRequiredFreightLIDescription = true
	isRequiredFreightLIWeightValue = true
	isRequiredFreightLIWeightUnits = true
Else
	isRequiredFreightAccountNumber = false
	isRequiredShipmentRoleType = false
	isRequiredTotalHandlingUnits = false
	
	isRequiredFreightContactPersonName = false
	isRequiredFreightContactCompanyName = false
	isRequiredFreightContactStreetLines = false
	isRequiredFreightContactCity = false
	isRequiredFreightContactStateCode = false
	isRequiredFreightContactPostalCode = false
	isRequiredFreightContactCountryCode = false
	
	isRequiredFreightLIPieces = false
	isRequiredFreightLIPackaging = false
	isRequiredFreightLIDescription = false
	isRequiredFreightLIWeightValue = false
	isRequiredFreightLIWeightUnits = false
End If

isHomeDeliveryTypeRequired = false
isHomeDeliveryDateRequired = false
isHomeDeliveryPhoneRequired = false


'   >>> Recipient Address Conditionals
'// Use the Request object to toggle State (based of Country selection)
isRequiredState2 =  true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired2")
if  len(pcv_strStateCodeRequired)>0 then
	isRequiredState2=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
isRequiredProvince2 = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired2")
if  len(pcv_strProvinceCodeRequired)>0 then
	isRequiredProvince2=pcv_strProvinceCodeRequired
end if

if Session("pcAdminShipToCountryCode") = "US" OR Session("pcAdminShipToCountryCode") = "CA" then
	isRequiredShipToPostal = true
end if


if Session("pcAdminCustomerReference") = "" then
	pcv_strCustomerReference = pidcustomer
	Session("pcAdminCustomerReference") = pcv_strCustomerReference
end if
if Session("pcAdminCustomerInvoiceNumber") = "" then
	CustomerInvoiceNumber = pidorder
	Session("pcAdminCustomerInvoiceNumber") = CustomerInvoiceNumber
end if

if Session("pcAdminLabelFormatType") = "" then
	pcv_strType = "COMMON2D"
	Session("pcAdminLabelFormatType") = pcv_strType
end if

if session("pcAdminLabelStockType") = "" then
	pcv_strLabelStockType = "PAPER_LETTER"
	Session("pcAdminLabelStockType") = pcv_strLabelStockType
end if



'// SET REQUIRED VARIABLES
pcv_strMethodName = "FDXShipRequest"
pcv_strMethodReply = "FDXShipReply"
CustomerTransactionIdentifier = "ProductCart_Test"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_intOrderID))%></b></td>
	</tr>
	<tr>
		<th colspan="2">FedEx<sup>&reg;</sup> Shipment Request</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			This flexible service allows a customer to request shipments and print return labels.  Simply fill out all required fields from each of the console's tabs.
			Then click the "Process Shipment" button to send your request to FedEx.  If any error or warning occurs it will be displayed on your screen.
			Once your order is confirmed you will be redirected back to the Shipping Wizard for FedEx.
			</p>
		</td>
	</tr>
</table>
<table class="pcCPcontent">
	<tr>
		<td>
		  <%
			'*******************************************************************************
			' START: ON POSTBACK
			'*******************************************************************************
			if request.form("submit")<>"" then

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' ServerSide Validate the Required Fields and Formatting.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

				'// Generic error for page
				pcv_strGenericPageError = "At least one required field was empty."

				'// Clear error string
				pcv_strSecondaryErrors = ""
				pcv_strErrorMsg = ""

				'// Get all the dynamic package details and validate
				pcv_xCounter = 1
				pcv_strTotalDeclaredValue = 0

				For pcv_xCounter = 1 to pcPackageCount

					' If its shipped the field is no longer required
					if pcLocalArray(pcv_xCounter-1) = "shipped" then
						pcv_strToggle = false
					else
						pcv_strToggle = true
					end if

					pcs_ValidateTextField	"FaxLetter"&pcv_xCounter, false, 0

					pcs_ValidateTextField	"Service"&pcv_xCounter, pcv_strToggle, 0

					pcs_ValidateTextField	"Length"&pcv_xCounter, false, 5
					pcs_ValidateTextField	"Width"&pcv_xCounter, false, 5
					pcs_ValidateTextField	"Height"&pcv_xCounter, false, 5
					pcs_ValidateTextField	"Units"&pcv_xCounter, false, 3
					pcs_ValidateTextField	"Packaging"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField "ContainerType"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"WeightUnits"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField	"Weight"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField	"declaredvalue"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField	"currency"&pcv_xCounter, false, 0
					Session("pcAdminLength"&pcv_xCounter)=int(Session("pcAdminLength"&pcv_xCounter))
					Session("pcAdminWidth"&pcv_xCounter)=int(Session("pcAdminWidth"&pcv_xCounter))
					Session("pcAdminHeight"&pcv_xCounter)=int(Session("pcAdminHeight"&pcv_xCounter))

					if Session("pcAdminWeight"&pcv_xCounter)="" then Session("pcAdminWeight"&pcv_xCounter) = 0
					if Session("pcAdmindeclaredvalue"&pcv_xCounter)="" then Session("pcAdmindeclaredvalue"&pcv_xCounter) = 0
					if Session("pcAdmincurrency"&pcv_xCounter)="" then Session("pcAdmincurrency"&pcv_xCounter) = "USD"
					pcv_strTotalDeclaredValue = pcv_strTotalDeclaredValue + Session("pcAdmindeclaredvalue"&pcv_xCounter)
				Next

				pcs_ValidateTextField	"TotalShipmentWeight", true, 10

				pcs_ValidateTextField	"ShipmentWeightUnits", true, 2
				If Session("pcAdminShipmentWeightUnits")&""="" Then
					Session("pcAdminShipmentWeightUnits") = "LB"
				End If
				pcs_ValidateTextField	"RateRequestType", true, 10
				If Session("pcAdminRateRequestType")&""="" Then
					Session("pcAdminRateRequestType") = "LIST"
				End If
				Session("pcAdminTotalDeclaredValue")=FormatNumber(pcv_strTotalDeclaredValue,2)
				pcs_ValidateTextField	"CarrierCode", true, 10
				select case session("pcAdminCarrierCode")
					case 1
						session("pcAdminCarrierCode") = "FDXE"
					case 2
						session("pcAdminCarrierCode") = "FDXG"
					case 3
						session("pcAdminCarrierCode") = "FXFR"
				end select
				
				pcs_ValidateTextField "LabelFormatType", false, 0
				pcs_ValidateTextField "LabelImageType", false, 0
				pcs_ValidateTextField "LabelStockType", false, 0
				pcs_ValidateTextField "LabelPrintingOrientation", false, 0

				pcs_ValidateTextField	"ShipDate", true, 0
				pcs_ValidateTextField	"ShipTime", true, 0
				pcs_ValidateTextField	"ReturnShipmentIndicator", false, 0
				pcs_ValidateTextField	"ReturnShipmentReason", false, 0
				if Session("pcAdminCarrierCode") = "FDXE" then
					isRequiredDropoffType = true
				else
					isRequiredDropoffType = false
				end if
				pcs_ValidateTextField "DropoffType", isRequiredDropoffType, 0
				pcs_ValidateTextField	"CurrencyCode", false, 3
				pcs_ValidateTextField	"ListRate", false, 1
				'// Origin
				pcs_ValidateTextField	"OriginPersonName", true, 0
				pcs_ValidateTextField	"OriginCompanyName", true, 0
				pcs_ValidateTextField	"OriginDepartment", false, 10
				pcs_ValidatePhoneNumber	"OriginPhoneNumber", true, 16
				pcs_ValidatePhoneNumber	"OriginPagerNumber", false, 16
				pcs_ValidatePhoneNumber	"OriginFaxNumber", false, 16
				pcs_ValidateEmailField	"OriginEmailAddress", true, 0
				pcs_ValidateTextField	"OriginLine1", true, 0
				pcs_ValidateTextField	"OriginLine2", false, 0
				pcs_ValidateTextField	"OriginCity", true, 0
				pcs_ValidateTextField	"OriginStateOrProvinceCode", isRequiredState, 2
				pcs_ValidateTextField	"OriginProvinceCode", isRequiredProvince, 2
				pcs_ValidateTextField	"OriginPostalCode", true, 16
				pcs_ValidateTextField	"OriginCountryCode", true, 2
				Session("pcAdminOriginPostalCode")=replace(Session("pcAdminOriginPostalCode"),"-","")

				'   >>> Merge Province or State into one variable before we send to FedEx
				if Session("pcAdminOriginProvinceCode") <> "" then
					Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminOriginProvinceCode")
				else
					Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminOriginStateOrProvinceCode")
				end if

				' Recipient
				pcs_ValidateTextField	"RecipPersonName", true, 0
				pcs_ValidateTextField	"RecipCompanyName", false, 0
				pcs_ValidateTextField	"RecipDepartment", false, 10
				pcs_ValidatePhoneNumber	"RecipPhoneNumber", true, 16
				pcs_ValidatePhoneNumber	"RecipPagerNumber", false, 16
				pcs_ValidatePhoneNumber	"RecipFaxNumber", false, 16
				pcs_ValidateEmailField	"RecipEmailAddress", false, 0

				'// Recipient Address
				pcs_ValidateTextField	"RecipCountryCode", true, 2

				'   >>> Recipient Address Conditionals
				if Session("pcAdminRecipCountryCode") = "US" OR Session("pcAdminRecipCountryCode") = "CA" then
					isRequiredRecipPostal = true
				else
					isRequiredRecipPostal = false
				end if
				pcs_ValidateTextField	"RecipLine1", true, 0
				pcs_ValidateTextField	"RecipLine2", false, 0
				pcs_ValidateTextField	"RecipCity", true, 0 '// FDXE-35, FDXG-20
				pcs_ValidateTextField	"RecipStateOrProvinceCode", isRequiredState2, 2
				pcs_ValidateTextField	"RecipProvinceCode", isRequiredProvince2, 2
				'   >>> Merge Province or State into one variable before we send to FedEx
				if Session("pcAdminRecipProvinceCode") <> "" then
					Session("pcAdminRecipStateOrProvinceCode")=Session("pcAdminRecipProvinceCode")
				else
					Session("pcAdminRecipStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
				end if

				pcs_ValidateTextField	"RecipPostalCode", isRequiredRecipPostal, 16

				Session("pcAdminRecipPostalCode")=replace(Session("pcAdminRecipPostalCode"),"-","")

				If pcv_FedExInternational Then
					isRequiredCVAmount = true
					isRequiredCVCurrency = true
					isRequiredNumberOfPieces = true
					isRequiredDescription = true
					isRequiredCountryOfManufacture = true
					isRequiredCommodityWeight = true
					isRequiredCommodityQuantity = true
					isRequiredCommodityQuantityUnits = true
					isRequiredCommodityUnitPrice = true
					isRequiredDutiesAccountNumber = true
					isRequiredDutiesCountryCode = true
				Else
					isRequiredCVAmount = false
					isRequiredCVCurrency = false
					isRequiredNumberOfPieces = false
					isRequiredDescription = false
					isRequiredCountryOfManufacture = false
					isRequiredCommodityWeight = false
					isRequiredCommodityQuantity = false
					isRequiredCommodityQuantityUnits = false
					isRequiredCommodityUnitPrice = false
					isRequiredDutiesAccountNumber = false
					isRequiredDutiesCountryCode = false
				End If

				'// International
				pcs_ValidateTextField	"DutiesAccountNumber", isRequiredDutiesAccountNumber, 12
				pcs_ValidateTextField	"DutiesPersonName", isRequiredDutiesPersonName, 0
				pcs_ValidateTextField	"DutiesCountryCode", isRequiredDutiesCountryCode, 2
				pcs_ValidateTextField	"DutiesPayorType", false, 12

				pcs_ValidateTextField	"PayorType", false, 0 '// Required if PayorType is RECIPIENT or THIRD_PARTY.
				pcs_ValidateTextField	"PayorAccountNumber", false, 0
				pcs_ValidateTextField	"PayorPersonName", false, 0
				pcs_ValidateTextField	"PayorCountryCode", false, 2

				'// Customer Reference
				pcs_ValidateTextField	"CustomerReference", true, 0 '// FDXE-40, FDXG-30
				pcs_ValidateTextField	"CustomerPONumber", false, 30
				pcs_ValidateTextField	"CustomerInvoiceNumber", false, 30

				'// COD
				pcs_ValidateTextField	"AddTransportationCharges", false, 45
				pcs_ValidateTextField	"CollectionAmount", false, 45 '// Required if COD element is provided.  Format: Two explicit decimals (e.g.5.00).
				pcs_ValidateTextField	"CollectionType", false, 45 '// ANY, GUARANTEEDFUNDS, CASH

				'// CODReturn
				' >>> will default to the shipper if not specified
				pcs_ValidateTextField	"TrackingNumber", false, 20 '// Required for a COD multiple-piece shipments only. / Tracking number assigned to the COD remittance
				pcs_ValidateTextField	"ReferenceIndicator", false, 0 '// TRACKING, REFERENCE, PO, INVOICE

				'// Expandable Regions
				pcs_ValidateTextField	"bOrder", false, 0
				pcs_ValidateTextField	"bShip", false, 0
				pcs_ValidateTextField	"bAdditional", false, 0
				pcs_ValidateTextField	"bInternational", false, 0
				pcs_ValidateTextField	"bGround", false, 0
				'// International
				pcs_ValidateTextField	"TotalCustomsValue", false, 0
				pcs_ValidateTextField	"AdmissibilityPackageType", false, 0
				if Session("pcAdminRecipCountryCode")<>"US" then
					pcs_ValidateTextField	"RecipientTIN", false, 0
				else
					pcs_ValidateTextField	"RecipientTIN", false, 0
				end if
				pcs_ValidateTextField	"SenderTINNumber", false, 0
				pcs_ValidateTextField	"SenderTINType", false, 0
				pcs_ValidateTextField	"AESOrFTSRExemptionNumber", false, 0
				pcs_ValidateTextField	"NumberOfPieces", isRequiredNumberOfPieces, 0
				pcs_ValidateTextField	"Description", isRequiredDescription, 0
				pcs_ValidateTextField	"CountryOfManufacture", isRequiredCountryOfManufacture, 0
				pcs_ValidateTextField	"HarmonizedCode", false, 0
				pcs_ValidateTextField	"CommodityWeightValue", isRequiredCommodityWeight, 0
				pcs_ValidateTextField	"CommodityWeightUnits", isRequiredCommodityWeight, 0
				pcs_ValidateTextField	"CommodityQuantity", isRequiredCommodityQuantity, 0
				pcs_ValidateTextField	"CommodityQuantityUnits", isRequiredCommodityQuantityUnits, 0
				pcs_ValidateTextField	"CommodityUnitCurrency", isRequiredCommodityUnitPrice, 0
				pcs_ValidateTextField	"CommodityUnitPrice", isRequiredCommodityUnitPrice, 0
				pcs_ValidateTextField	"CommodityCustomsValue", false, 0
				pcs_ValidateTextField	"CommodityCustomsCurrency", false, 0
				pcs_ValidateTextField	"ExportLicenseNumber", false, 0
				pcs_ValidateTextField	"ExportLicenseExpirationDate", false, 0
				pcs_ValidateTextField	"CIMarksAndNumbers", false, 0
				pcs_ValidateTextField	"B13AFilingOption", false, 0
				pcs_ValidateTextField	"ExportComplianceStatement", false, 0

				'// NAFTA Certificate of Origin
				pcs_ValidateTextField	"bNAFTAOption", false, 0
				pcs_ValidateTextField	"NAFTAPreferenceCriterion", false, 0
				pcs_ValidateTextField	"NAFTAProducerDetermination", false, 0
				pcs_ValidateTextField	"NAFTAProducerID", false, 0
				pcs_ValidateTextField	"NAFTANetCostMethod", false, 0
				
				'// FICE FedEx International Controlled Export
				pcs_ValidateTextField	"bFICEOption", false, 0
				If Session("pcAdminbFICEOption") = "1" Then
					isRequiredFICENumber = true
					isRequiredFICEExpirationDate = true
				Else
					isRequiredFICENumber = false
					isRequiredFICEExpirationDate = false
				End If
				pcs_ValidateTextField	"FICENumber", isRequiredFICENumber, 0
				pcs_ValidateTextField	"FICEExpirationDate", isRequiredFICEExpirationDate, 0
				pcs_ValidateTextField	"FICEEntryNumber", false, 0
				pcs_ValidateTextField	"FICETradeZoneCode", false, 0
				
				'// ITAR International Traffic in Arms Regulations
				pcs_ValidateTextField	"bITAROption", false, 0
				pcs_ValidateTextField	"ITARNumber", false, 0

        '// Additional Shipping Documents (SDS)
        pcs_ValidateTextField "bSOShippingDocumentOption", false, 0
        pcs_ValidateTextField "SDSType", false, 0
        pcs_ValidateTextField "SDSLabelImageType", false, 0
        pcs_ValidateTextField "SDSLabelStockType", false, 0
        pcs_ValidateTextField "SDSProvideInstructions", false, 0
				
        '// Electronic Trade Documents (ETD)
        pcs_ValidateTextField "bSOETDOption", false, 0
        pcs_ValidateTextField "ETDRequestedDocumentCopies", false, 0
        pcs_ValidateTextField "ETDLineNumber", false, 0
        pcs_ValidateTextField "ETDDocumentProducer", false, 0
        pcs_ValidateTextField "ETDDocumentId", false, 0
        pcs_ValidateTextField "ETDDocumentIdProducer", false, 0

				'// Hold At Location
				pcs_ValidateTextField	"HALContactID", false, 0
				pcs_ValidateTextField	"HALPersonName", false, 0
				pcs_ValidateTextField	"HALCompanyName", false, 0
				pcs_ValidateTextField	"HALPhone", false, 0
				pcs_ValidateTextField	"HALPhoneExtension", false, 0
				pcs_ValidateTextField	"HALPager", false, 0
				pcs_ValidateTextField	"HALFax", false, 0
				pcs_ValidateTextField	"HALEmail", false, 0
				pcs_ValidateTextField	"HALLine1", false, 0
				pcs_ValidateTextField	"HALCity", false, 0
				pcs_ValidateTextField	"HALStateOrProvinceCode", false, 0
				pcs_ValidateTextField	"HALPostalCode", false, 0
				pcs_ValidateTextField	"HALUrbanizationCode", false, 0
				pcs_ValidateTextField	"HALCountryCode", false, 0
				pcs_ValidateTextField	"HALResidential", false, 0
				pcs_ValidateTextField	"HALLocationType", false, 0

				'// Dry Ice Shipment
				pcs_ValidateTextField	"SDIPackageCount", false, 0
				pcs_ValidateTextField	"SDIValue", false, 0
				'pcs_ValidateTextField	"SDIUnit", false, 0

				'// International Freight
				pcs_ValidateTextField "BookingConfirmationNumber", false, 12

				'// Freight
				pcs_ValidateTextField "FreightAccountNumber", isRequiredFreightAccountNumber, 12
				pcs_ValidateTextField "FreightShipmentRoleType", isRequiredShipmentRoleType, 0
				pcs_ValidateTextField "FreightTotalHandlingUnits", isRequiredTotalHandlingUnits, 0
				pcs_ValidateTextField "FreightCollectTermsType", false, 0
				pcs_ValidateTextField "FreightClientDiscount", false, 0
				pcs_ValidateTextField "FreightPalletWeightValue", false, 0
				pcs_ValidateTextField "FreightPalletWeightUnits", false, 0
				pcs_ValidateTextField "FreightShipmentDimensionsLength", false, 0
				pcs_ValidateTextField "FreightShipmentDimensionsWidth", false, 0
				pcs_ValidateTextField "FreightShipmentDimensionsHeight", false, 0
				pcs_ValidateTextField "FreightShipmentDimensionsUnits", false, 0
				pcs_ValidateTextField "FreightShipmentComment", false, 0
				
				'// Billing Contact & Address
				pcs_ValidateTextField "FreightContactPersonName", isRequiredFreightContactPersonName, 0
				pcs_ValidateTextField "FreightContactCompanyName", isRequiredFreightContactCompanyName, 0
				pcs_ValidateTextField "FreightContactPagerNumber", false, 0
				pcs_ValidateTextField "FreightContactStreetLines", isRequiredFreightContactStreetLines, 0
				pcs_ValidateTextField "FreightContactCity", isRequiredFreightContactCity, 0
				pcs_ValidateTextField "FreightContactStateCode", isRequiredFreightContactStateCode, 0
				pcs_ValidateTextField "FreightContactPostalCode", isRequiredFreightContactPostalCode, 0
				pcs_ValidateTextField "FreightContactCountryCode", isRequiredFreightContactCountryCode, 0
				
				'// Freight Declared Value
				pcs_ValidateTextField "FreightDVCurrency", false, 0
				pcs_ValidateTextField "FreightDVAmount", false, 0
				pcs_ValidateTextField "FreightDVUnits", false, 0
				
				'// Freight Liability Coverage
				pcs_ValidateTextField "FreightLCType", false, 0
				pcs_ValidateTextField "FreightLCCurrency", false, 0
				pcs_ValidateTextField "FreightLCAmount", false, 0
				
				'// Freight Line Items/Commodities
				pcs_ValidateTextField "FreightLIPackaging", isRequiredFreightLIPackaging, 0
				pcs_ValidateTextField "FreightLIPieces", isRequiredFreightLIPieces, 0
				pcs_ValidateTextField "FreightLIDescription", isRequiredFreightLIDescription, 0
				pcs_ValidateTextField "FreightLIWeightValue", isRequiredFreightLIWeightValue, 0
				pcs_ValidateTextField "FreightLIWeightUnits", isRequiredFreightLIWeightUnits, 0

				pcs_ValidateTextField "FreightLIClass", false, 0
				pcs_ValidateTextField "FreightLIClassProvided", false, 0
				pcs_ValidateTextField "FreightLIHandlingUnits", false, 0
				pcs_ValidateTextField "FreightLIPONumber", false, 0
				pcs_ValidateTextField "FreightLIDimensionsLength", false, 0
				pcs_ValidateTextField "FreightLIDimensionsWidth", false, 0
				pcs_ValidateTextField "FreightLIDimensionsHeight", false, 0
				pcs_ValidateTextField "FreightLIDimensionsUnits", false, 0

				pcs_ValidateTextField "DeliveryInstructions", false, 0
				
				'//Special Services
				pcs_ValidateTextField "ResidentialDelivery", false, 0
				pcs_ValidateTextField "InsideDelivery", false, 0
				pcs_ValidateTextField "InsidePickup", false, 0
				pcs_ValidateTextField "bSOAlcoholOption", false, 0
				pcs_ValidateTextField "AlcoholRecipientType", false, 0
				pcs_ValidateTextField "PharmacyDelivery", false, 0
				pcs_ValidateTextField "bSOSaturdayServices", false, 0
				pcs_ValidateTextField "SaturdayPickup", false, 0
				pcs_ValidateTextField "SaturdayDelivery", false, 0
				pcs_ValidateTextField "bSOSignatureOption", false, 0
				pcs_ValidateTextField "SignatureOption", false, 0
				pcs_ValidateTextField "SignatureRelease", false, 0
				pcs_ValidateTextField "ExtremeLength", false, 0
				pcs_ValidateTextField "OneRate", false, 0
				pcs_ValidateTextField "PriorityAlert", false, 0
				pcs_ValidateTextField "PAContent", false, 0
				pcs_ValidateTextField "PriorityAlertPlus", false, 0
				pcs_ValidateTextField "PAPContent", false, 0
				pcs_ValidateTextField "bSOHAL", false, 0
				pcs_ValidateTextField "bSODryIce", false, 0
				pcs_ValidateTextField "bSODGShip", false, 0
				pcs_ValidateTextField "bISOBrokerSelect", false, 0
				pcs_ValidateTextField "BSOType", false, 0
				pcs_ValidateTextField "BSOAccountNumber", false, 0
				pcs_ValidateTextField "BSOTinType", false, 0
				pcs_ValidateTextField "BSOTinNumber", false, 0
				pcs_ValidateTextField "BSOContactID", false, 0
				pcs_ValidateTextField "BSOPersonName", false, 0
				pcs_ValidateTextField "BSOTitle", false, 0
				pcs_ValidateTextField "BSOCompanyName", false, 0
				pcs_ValidateTextField "BSOPhoneNumber", false, 0
				pcs_ValidateTextField "BSOPhoneExtension", false, 0
				pcs_ValidateTextField "BSOEmailAddress", false, 0
				pcs_ValidateTextField "BSOStreetLines", false, 0
				pcs_ValidateTextField "BSOCity", false, 0
				pcs_ValidateTextField "BSOStateOrProvinceCode", false, 0
				pcs_ValidateTextField "BSOPostalCode", false, 0
				pcs_ValidateTextField "BSOCountryCode", false, 0
				pcs_ValidateTextField "bSOCODCollection", false, 0
				pcs_ValidateTextField "CODAmount", false, 0
				pcs_ValidateTextField "CODCurrency", false, 0
				pcs_ValidateTextField "CODRateType", false, 0
				pcs_ValidateTextField "CODChargeBasis", false, 0
				pcs_ValidateTextField "CODChargeBasisLevel", false, 0
				pcs_ValidateTextField "CODType", false, 0
				pcs_ValidateTextField "CODTinType", false, 0
				pcs_ValidateTextField "CODTinNumber", false, 0
				pcs_ValidateTextField "CODAccountNumber", false, 0
				pcs_ValidateTextField "CODPersonName", false, 0
				pcs_ValidateTextField "CODCompanyName", false, 0
				pcs_ValidateTextField "CODPhoneNumber", false, 0
				pcs_ValidateTextField "CODTitle", false, 0
				pcs_ValidateTextField "CODStreetLines", false, 0
				pcs_ValidateTextField "CODCity", false, 0
				pcs_ValidateTextField "CODState", false, 0
				pcs_ValidateTextField "CODPostalCode", false, 0
				pcs_ValidateTextField "CODCountryCode", false, 0
				
				'//Customs
				pcs_ValidateTextField "CVAmount", isRequiredCVAmount, 0
				pcs_ValidateTextField "CVCurrency", isRequiredCVAmount, 0
				pcs_ValidateTextField "CICAmount", false, 0
				pcs_ValidateTextField "CICCurrency", false, 0
				pcs_ValidateTextField "CMCAmount", false, 0
				pcs_ValidateTextField "CMCCurrency", false, 0
				pcs_ValidateTextField "CFCAmount", false, 0
				pcs_ValidateTextField "CFCCurrency", false, 0
				pcs_ValidateTextField "CCIPurpose", false, 0
				pcs_ValidateTextField "CCIInvoiceNumber", false, 0
				pcs_ValidateTextField "CCICustomerReference", false, 0
				pcs_ValidateTextField	"CCITermsOfSale", false, 0
				pcs_ValidateTextField "CCIComments", false, 0
				pcs_ValidateTextField "CCDOptionType", false, 0
				If Session("pcAdminCCDOptionType") = "OTHER" Then
					isCCDOptionDescriptionRequired = true
				Else
					isCCDOptionDescriptionRequired = false
				End If
				pcs_ValidateTextField "CCDOptionDescription", isCCDOptionDescriptionRequired, 0

				'//Importer of Record
				pcs_ValidateTextField "IORPersonName", false, 0
				pcs_ValidateTextField "IORCompanyName", false, 0
				pcs_ValidateTextField "IORPhoneNumber", false, 0
				pcs_ValidateTextField "IORAddress", false, 0
				pcs_ValidateTextField "IORCity", false, 0
				pcs_ValidateTextField "IORStateOrProvince", false, 0
				pcs_ValidateTextField "IORCountryCode", false, 0
				pcs_ValidateTextField "IORPostalCode", false, 0
				
				'//Shipper Notification
				pcs_ValidateTextField	"ShipperNotificationFormat", false, 0
				
				pcs_ValidateTextField	"NotificationShipperEnabled", false, 0
				pcs_ValidateTextField	"NotificationShipperEmail", false, 0
				
				pcs_ValidateTextField	"NotificationRecipientEnabled", false, 0
				pcs_ValidateTextField	"NotificationRecipientEmail", false, 0
				
				pcs_ValidateTextField	"NotificationThirdPartyEnabled", false, 0
				pcs_ValidateTextField	"NotificationThirdPartyEmail", false, 0
				
				pcs_ValidateTextField	"NotificationBrokerEnabled", false, 0
				pcs_ValidateTextField	"NotificationBrokerEmail", false, 0
				
				pcs_ValidateTextField	"NotificationOtherEnabled", false, 0
				pcs_ValidateTextField	"NotificationOtherEmail", false, 0

        If Session("pcAdminService1") = "GROUND_HOME_DELIVERY" Then
          isHomeDeliveryTypeRequired = false
          isHomeDeliveryDateRequired = false
          isHomeDeliveryPhoneRequired = false
        End If

				pcs_ValidateTextField		"HomeDeliveryType", isHomeDeliveryTypeRequired, 14
				pcs_ValidateTextField		"HomeDeliveryDate", isHomeDeliveryDateRequired, 0
				pcs_ValidatePhoneNumber	"HomeDeliveryPhone", isHomeDeliveryPhoneRequired, 16
				pcs_ValidateTextField		"HomeDeliveryInstructions", false, 74

				pcs_ValidateTextField "RCIdValue", false, 0
				pcs_ValidateTextField "RCIdType", false, 0

				'Express Freight
				pcs_ValidateTextField "EFPackingListEnclosed", false, 0
				pcs_ValidateTextField "EFShippersLoadAndCount", false, 0
				pcs_ValidateTextField "EFBookingConfirmationNumber", false, 0

				'Dangerous Goods
				pcs_ValidateTextField "DGAccessibility", false, 0
				pcs_ValidateTextField "DGAircraftOnly", false, 0
				pcs_ValidateTextField "DGHazardousMaterials", false, 0
				pcs_ValidateTextField "DGORMD", false, 0
				pcs_ValidateTextField "DGContainerType", false, 0
				pcs_ValidateTextField "DGContainerCount", false, 0
				pcs_ValidateTextField "DGPackagingCount", false, 0
				pcs_ValidateTextField "DGPackagingUnits", false, 0
				
				for i = 0 to DGMaxContainers
					pcs_ValidateTextField "DGCommodityID" & i, false, 0
					pcs_ValidateTextField "DGPackingGroup" & i, false, 0
					pcs_ValidateTextField "DGContainerAircraftOnly" & i, false, 0
					pcs_ValidateTextField "DGPackingInstructions" & i, false, 0
					pcs_ValidateTextField "DGShippingName" & i, false, 0
					pcs_ValidateTextField "DGHazardClass" & i, false, 0
					pcs_ValidateTextField "DGQuantityAmount" & i, false, 0
					pcs_ValidateTextField "DGQuantityUnits" & i, false, 0
				next
				
				pcs_ValidateTextField "DGContactName", false, 0
				pcs_ValidateTextField "DGContactTitle", false, 0
				pcs_ValidateTextField "DGContactPlace", false, 0
				pcs_ValidateTextField "DGEmergencyContactNumber", false, 0
				pcs_ValidateTextField "DGOfferor", false, 0
				pcs_ValidateTextField "DocumentsOnly", false, 0

				'SMARTPOST
				pcs_ValidateTextField "SMIndicia", false, 0
				pcs_ValidateTextField "SMAncillaryEndorsement", false, 0
				pcs_ValidateTextField "SMHubID", false, 0

				'// Additional Validation for Numerics
				if isNULL(Session("pcAdminDocumentsOnly")) then
					Session("pcAdminDocumentsOnly")="0"
				end if

				if isNULL(Session("pcAdminDGHazardousMaterials")) OR Session("pcAdminDGHazardousMaterials")<>"1" then
					Session("pcAdminDGHazardousMaterials")="0"
				end if
				
				if isNULL(Session("pcAdminDGORMD")) OR Session("pcAdminDGORMD")<>"1" then
					Session("pcAdminDGORMD")="0"
				end if

				if isNULL(Session("pcAdminEFPackingListEnclosed")) OR Session("pcAdminEFPackingListEnclosed")<>"1" then
					Session("pcAdminEFPackingListEnclosed")="0"
				end if

				if isNULL(Session("pcAdminResidentialDelivery")) OR Session("pcAdminResidentialDelivery")<>"true" then
					Session("pcAdminResidentialDelivery")="false"
				end if

				if NOT validNum(Session("pcAdminInsideDelivery")) OR Session("pcAdminInsideDelivery")<>"1" then
					Session("pcAdminInsideDelivery")="0"
				end if
				if NOT validNum(Session("pcAdminbSOCODCollection")) OR Session("pcAdminbSOCODCollection")<>"1" then
					Session("pcAdminbSOCODCollection")="0"
				end if
				if NOT validNum(Session("pcAdminbSOAlcoholOption")) OR Session("pcAdminbSOAlcoholOption")<>"1" then
					Session("pcAdminbSOAlcoholOption")="0"
				end if
				if NOT validNum(Session("pcAdminSaturdayPickup")) OR Session("pcAdminSaturdayPickup")<>"1" then
					Session("pcAdminSaturdayPickup")="0"
				end if
				if NOT validNum(Session("pcAdminSaturdayDelivery")) OR Session("pcAdminSaturdayDelivery")<>"1" then
					Session("pcAdminSaturdayDelivery")="0"
				end if
				if NOT validNum(Session("pcAdminbSOHAL")) OR Session("pcAdminbSOHAL")<>"1" then
					Session("pcAdminbSOHAL")="0"
				end if
				if NOT validNum(Session("pcAdminbSODryIce")) OR Session("pcAdminbSODryIce")<>"1" then
					Session("pcAdminbSODryIce")="0"
				end if
				if NOT validNum(Session("pcAdminbSODGShip")) OR Session("pcAdminbSODGShip")<>"1" then
					Session("pcAdminbSODGShip")="0"
				end if
				if NOT validNum(Session("pcAdminDGAircraftOnly")) OR Session("pcAdminDGAircraftOnly")<>"1" then
					Session("pcAdminDGAircraftOnly")="0"
				end if
				if NOT validNum(Session("pcAdminbISOBrokerSelect")) OR Session("pcAdminbISOBrokerSelect")<>"1" then
					Session("pcAdminbISOBrokerSelect")="0"
				end if
				if NOT validNum(Session("pcAdminbNAFTAOption")) OR Session("pcAdminbNAFTAOption")<>"1" then
					Session("pcAdminbNAFTAOption")="0"
				end if
				if NOT validNum(Session("pcAdminbFICEOption")) OR Session("pcAdminbFICEOption")<>"1" then
					Session("pcAdminbFICEOption")="0"
				end if
				if NOT validNum(Session("pcAdminbITAROption")) OR Session("pcAdminbITAROption")<>"1" then
					Session("pcAdminbITAROption")="0"
				end if
				If Session("pcAdminOtherShipmentNotification")="1" OR Session("pcAdminOtherDeliveryNotification")="1" OR Session("pcAdminOtherExceptionNotification")="1" then
					pcs_ValidateEmailField "OtherNotification", true, 0
				else
					pcs_ValidateEmailField "OtherNotification", false, 0
				end if
				if Session("pcAdminCommodityQuantity")="" then
					Session("pcAdminCommodityQuantity")=1
				end if
				if Session("pcAdminNumberOfPieces")="" then
					Session("pcAdminNumberOfPieces")=0
				end if

				'//Smart Post Flag
				mySP = 0

				if Session("pcAdminService1") = "SMART_POST" then
					mySP = 1
				end if
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Check for Validation Errors. Do not proceed if there are errors.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				If pcv_intErr>0 Then
					call closeDb()
response.redirect pcPageName & "?sub=1&msg=" & pcv_strGenericPageError
				Else

					'///////////////////////////////////////////////////////////////////////
					'// START LOOP
					'///////////////////////////////////////////////////////////////////////
					pcv_xCounter = 1
					pcv_strTotalDeclaredValue = 0
					pcv_strTotalWeight = 0
					errnum = 0
					For pcv_xCounter = 1 to pcPackageCount


					'// Reverse Address if Return Shipment
'						if Session("pcAdminReturnShipmentIndicator")="PRINT_RETURN_LABEL" then
'							pcv_a=Session("pcAdminOriginPersonName")
'							pcv_b=Session("pcAdminOriginCompanyName")
'							pcv_c=Session("pcAdminOriginDepartment")
'							pcv_d=Session("pcAdminOriginPhoneNumber")
'							pcv_e=Session("pcAdminOriginPagerNumber")
'							pcv_f=Session("pcAdminOriginFaxNumber")
'							pcv_g=Session("pcAdminOriginEmailAddress")
'							pcv_h=Session("pcAdminOriginLine1")
'							pcv_i=Session("pcAdminOriginLine2")
'							pcv_j=Session("pcAdminOriginCity")
'							pcv_k=Session("pcAdminOriginStateOrProvinceCode")
'							pcv_l=Session("pcAdminOriginPostalCode")
'							pcv_m=Session("pcAdminOriginCountryCode")
'
'							Session("pcAdminOriginPersonName")=Session("pcAdminRecipPersonName")
'							Session("pcAdminOriginCompanyName")=Session("pcAdminRecipCompanyName")
'							Session("pcAdminOriginDepartment")=Session("pcAdminRecipDepartment")
'							Session("pcAdminOriginPhoneNumber")=Session("pcAdminRecipPhoneNumber")
'							Session("pcAdminOriginPagerNumber")=Session("pcAdminRecipPagerNumber")
'							Session("pcAdminOriginFaxNumber")=Session("pcAdminRecipFaxNumber")
'							Session("pcAdminOriginEmailAddress")=Session("pcAdminRecipEmailAddress")
'							Session("pcAdminOriginLine1")=Session("pcAdminRecipLine1")
'							Session("pcAdminOriginLine2")=Session("pcAdminRecipLine2")
'							Session("pcAdminOriginCity")=Session("pcAdminRecipCity")
'							Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
'							Session("pcAdminOriginPostalCode")=Session("pcAdminRecipPostalCode")
'							Session("pcAdminOriginCountryCode")=Session("pcAdminRecipCountryCode")
'
'							Session("pcAdminRecipPersonName")=pcv_a
'							Session("pcAdminRecipCompanyName")=pcv_b
'							Session("pcAdminRecipDepartment")=pcv_c
'							Session("pcAdminRecipPhoneNumber")=pcv_d
'							Session("pcAdminRecipPagerNumber")=pcv_e
'							Session("pcAdminRecipFaxNumber")=pcv_f
'							Session("pcAdminRecipEmailAddress")=pcv_g
'							Session("pcAdminRecipLine1")=pcv_h
'							Session("pcAdminRecipLine2")=pcv_i
'							Session("pcAdminRecipCity")=pcv_j
'							Session("pcAdminRecipStateOrProvinceCode")=pcv_k
'							Session("pcAdminRecipPostalCode")=pcv_l
'							Session("pcAdminRecipCountryCode")=pcv_m
'						end if

						'// If the package was processed, skip it.
						if pcLocalArray(pcv_xCounter-1) <> "shipped" then
						
              pcv_strVersion = FedExWS_ShipVersion

              ' No prefix or namespace in the request
              fedex_xmlPrefix = ""
              fedex_xmlNamespace = ""

						  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						  ' Build Our Transaction.
						  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						  objFedExClass.NewXMLLabelWS "ProcessShipmentRequest", FedExKey, FedExPassword, FedExAccountNumber, FedExMeterNumber, pcv_strVersion, "ship"

							'--------------------
							'// TransactionDetail
							'--------------------
							objFedExClass.WriteParent "TransactionDetail", ""
								objFedExClass.AddNewNode "CustomerTransactionId", Session("pcAdminCustomerInvoiceNumber")
							objFedExClass.WriteParent "TransactionDetail", "/"

							'--------------------
							'// Version
							'--------------------
							objFedExClass.WriteParent "Version", ""
								objFedExClass.AddNewNode "ServiceId", "ship"
								objFedExClass.AddNewNode "Major", pcv_strVersion
								objFedExClass.AddNewNode "Intermediate", "0"
								objFedExClass.AddNewNode "Minor", "0"
							objFedExClass.WriteParent "Version", "/"

							'--------------------
							'// RequestedShipment
							'--------------------
							objFedExClass.WriteParent "RequestedShipment", ""
								objFedExClass.AddNewNode "ShipTimestamp", Session("pcAdminShipDate") & "T" & Session("pcAdminShipTime")
								objFedExClass.AddNewNode "DropoffType", Session("pcAdminDropoffType")
								objFedExClass.AddNewNode "ServiceType", Session("pcAdminService1")
								objFedExClass.AddNewNode "PackagingType", Session("pcAdminPackaging1")
								'// Off for smartpost
								If mySP=0 Then
									objFedExClass.WriteParent "TotalWeight", ""
										objFedExClass.AddNewNode "Units", Session("pcAdminShipmentWeightUnitS")
										objFedExClass.AddNewNode "Value", Session("pcAdminTotalShipmentWeight")
									objFedExClass.WriteParent "TotalWeight", "/"
								End If

								'--------------------------------
								'// RequestedShipment/Shipper
								'--------------------------------
								objFedExClass.WriteParent "Shipper", ""
									objFedExClass.AddNewNode "AccountNumber", ""
									If Session("pcAdminSenderTINNumber")&""<>"" Then
										objFedExClass.WriteParent "Tins", ""
											objFedExClass.AddNewNode "TinType", Session("pcAdminSenderTINType")
											objFedExClass.AddNewNode "Number", fnStripPhone(Session("pcAdminSenderTINNumber"))
										objFedExClass.WriteParent "Tins", "/"
									End If
									objFedExClass.WriteParent "Contact", ""
										objFedExClass.AddNewNode "PersonName", sanitizeField(Session("pcAdminOriginPersonName"))
										objFedExClass.AddNewNode "CompanyName", sanitizeField(Session("pcAdminOriginCompanyName"))
										objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminOriginPhoneNumber"))
										objFedExClass.AddNewNode "PagerNumber", fnStripPhone(Session("pcAdminOriginPagerNumber"))
										objFedExClass.AddNewNode "FaxNumber", fnStripPhone(Session("pcAdminOriginFaxNumber"))
										objFedExClass.AddNewNode "EMailAddress", Session("pcAdminOriginEmailAddress")
									objFedExClass.WriteParent "Contact", "/"
									objFedExClass.WriteParent "Address", ""
										objFedExClass.AddNewNode "StreetLines", Session("pcAdminOriginLine1")
										objFedExClass.AddNewNode "StreetLines", Session("pcAdminOriginLine2")
										objFedExClass.AddNewNode "City", Session("pcAdminOriginCity")
										
										'Workaround for FedEx old PQ code
										stateOrProvinceCode = Session("pcAdminOriginStateOrProvinceCode")
										if stateOrProvinceCode = "QC" then
											stateOrProvinceCode = "PQ"
										end if
										objFedExClass.AddNewNode "StateOrProvinceCode", stateOrProvinceCode
										
										objFedExClass.AddNewNode "PostalCode", Session("pcAdminOriginPostalCode")

										objFedExClass.AddNewNode "CountryCode", Session("pcAdminOriginCountryCode")
										objFedExClass.AddNewNode "Residential", "false"
									objFedExClass.WriteParent "Address", "/"
								objFedExClass.WriteParent "Shipper", "/"

								'--------------------------------
								'// RequestedShipment/Recipient
								'--------------------------------
								objFedExClass.WriteParent "Recipient", ""
									objFedExClass.AddNewNode "AccountNumber", ""
									objFedExClass.WriteParent "Contact", ""
										objFedExClass.AddNewNode "PersonName", sanitizeField(Session("pcAdminRecipPersonName"))
										objFedExClass.AddNewNode "CompanyName", sanitizeField(Session("pcAdminRecipCompanyName"))
										objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminRecipPhoneNumber"))
										objFedExClass.AddNewNode "PagerNumber", fnStripPhone(Session("pcAdminRecipPagerNumber"))
										objFedExClass.AddNewNode "FaxNumber", fnStripPhone(Session("pcAdminRecipFaxNumber"))
										objFedExClass.AddNewNode "EMailAddress", Session("pcAdminRecipEmailAddress")
									objFedExClass.WriteParent "Contact", "/"
									objFedExClass.WriteParent "Address", ""
										objFedExClass.AddNewNode "StreetLines", Session("pcAdminRecipLine1")
										objFedExClass.AddNewNode "StreetLines", Session("pcAdminRecipLine2")
										objFedExClass.AddNewNode "City", Session("pcAdminRecipCity")
										if Session("pcAdminRecipCountryCode")="US" OR Session("pcAdminRecipCountryCode")="CA" then
											'Workaround for FedEx old PQ code
											stateOrProvinceCode = Session("pcAdminRecipStateOrProvinceCode")
											if stateOrProvinceCode = "QC" then
												stateOrProvinceCode = "PQ"
											end if
											objFedExClass.AddNewNode "StateOrProvinceCode", stateOrProvinceCode
										end if

										objFedExClass.AddNewNode "PostalCode", Session("pcAdminRecipPostalCode")
										objFedExClass.AddNewNode "CountryCode", Session("pcAdminRecipCountryCode") '"US"
										objFedExClass.AddNewNode "Residential", Session("pcAdminResidentialDelivery")
									objFedExClass.WriteParent "Address", "/"
								objFedExClass.WriteParent "Recipient", "/"

								'--------------------------------------------
								'// RequestedShipment/ShippingChargesPayment
								'--------------------------------------------
								objFedExClass.WriteParent "ShippingChargesPayment", ""
									objFedExClass.AddNewNode "PaymentType", Session("pcAdminPayorType")
									objFedExClass.WriteParent "Payor", ""
										objFedExClass.WriteParent "ResponsibleParty", ""
											objFedExClass.AddNewNode "AccountNumber", Session("pcAdminPayorAccountNumber")
											objFedExClass.WriteParent "Contact", ""
												objFedExClass.AddNewNode "PersonName", Session("pcAdminPayorPersonName")
											objFedExClass.WriteParent "Contact", "/"
											objFedExClass.WriteParent "Address", ""
												objFedExClass.AddNewNode "CountryCode", Session("pcAdminPayorCountryCode")
											objFedExClass.WriteParent "Address", "/"
										objFedExClass.WriteParent "ResponsibleParty", "/"
									objFedExClass.WriteParent "Payor", "/"
								objFedExClass.WriteParent "ShippingChargesPayment", "/"

								'---------------------------------------------
								'// RequestedShipment/SpecialServicesRequested
								'---------------------------------------------
								BSO = 0
								FUT = 0
								INS = 0
								INP = 0
								PHAR = 0
								EL = 0
								ONERATE = 0
								SAT_DELIVERY = 0
								SAT_PICKUP = 0

								COD = 0
								HAL = 0
								ENS = 0
								RET = 0
								FICE = 0
								ITAR = 0
								HDP = 0
								ETD = 0

								If Session("pcAdminbISOBrokerSelect")="1" Then BSO = 1
								If DateDiff("d", FedExDateFormat(Date()), Session("pcAdminShipDate")) >= 1 Then FUT = 1
								If Session("pcAdminInsideDelivery")="1" Then INS = 1
								If Session("pcAdminInsidePickup")="1" Then INP = 1
								If Session("pcAdminPharmacyDelivery")="1" Then PHAR = 1
								If Session("pcAdminExtremeLength")="1" Then EL = 1
								If Session("pcAdminOneRate")="1" Then ONERATE = 1
								If Session("pcAdminbSOSaturdayServices")="1" Then
										If Session("pcAdminSaturdayDelivery")="1" Then SAT_DELIVERY = 1
										If Session("pcAdminSaturdayPickup")="1" Then SAT_PICKUP = 1
								End If
								If Session("pcAdminbSOCODCollection")="1" AND Session("pcAdminService1")<>"FEDEX_GROUND" Then COD = 1
								If Session("pcAdminbSOHAL")="1" Then HAL = 1

                '// START Email Notifications
								pcTempNotificationShipperEmail = 0
								pcTempNotificationRecipientEmail = 0
								pcTempNotificationThirdPartyEmail = 0
								pcTempNotificationBrokerEmail = 0
								pcTempNotificationOtherEmail = 0
										
								If Session("pcAdminNotificationShipperEnabled")="1" AND Session("pcAdminNotificationShipperEmail")&""<>"" Then
									pcTempNotificationShipperEmail = 1
									ENS = 1
								End If
										
								If Session("pcAdminNotificationRecipientEnabled")="1" AND Session("pcAdminNotificationRecipientEmail")&""<>"" Then
									pcTempNotificationRecipientEmail = 1
									ENS = 1
								End If
										
								If Session("pcAdminNotificationThirdPartyEnabled")="1" AND Session("pcAdminNotificationThirdPartyEmail")&""<>"" Then
									pcTempNotificationThirdPartyEmail = 1
									ENS = 1
								End If
										
								If Session("pcAdminNotificationBrokerEnabled")="1" AND Session("pcAdminNotificationBrokerEmail")&""<>"" Then
									pcTempNotificationBrokerEmail = 1
									ENS = 1
								End If
										
								If Session("pcAdminNotificationOtherEnabled")="1" AND Session("pcAdminNotificationOtherEmail")&""<>"" Then
									pcTempNotificationOtherEmail = 1
									ENS = 1
								End If
                '// END Email Notifications

                If Session("pcAdminReturnShipmentIndicator")="PRINT_RETURN_LABEL" Then RET = 1
								If Session("pcAdminbFICEOption") = "1" Then FICE = 1
								If Session("pcAdminbITAROption") = "1" Then ITAR = 1
								If Session("pcAdminHomeDeliveryType")&""<>"" AND Session("pcAdminService1") = "GROUND_HOME_DELIVERY" Then HDP = 1
                If Session("pcAdminbSOETDOption")&""<>"" Then ETD = 1

								If BSO = 1 OR FUT = 1 OR INS = 1 OR INP = 1 OR PHAR = 1 OR EL = 1 OR ONERATE = 1 OR SAT_DELIVERY = 1 OR SAT_PICKUP = 1 OR COD = 1 OR HAL = 1 OR ENS = 1 OR RET = 1 OR FICE = 1 OR ITAR = 1 OR HDP = 1 OR ETD = 1 Then
								  objFedExClass.WriteParent "SpecialServicesRequested", ""
										
                  '// Options with NO related detail (go first)
									If BSO = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "BROKER_SELECT_OPTION"
									If FUT = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "FUTURE_DAY_SHIPMENT"
									If INS = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "INSIDE_DELIVERY"
									If INP = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "INSIDE_PICKUP"
		
									'// Pharmacy Delivery
									If PHAR = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "PHARMACY_DELIVERY"
				
									'// Extreme Length
									if EL=1 then
										objFedExClass.AddNewNode "SpecialServiceTypes", "EXTREME_LENGTH"
									end if

									'// One Rate
									if ONERATE=1 then
										objFedExClass.AddNewNode "SpecialServiceTypes", "FEDEX_ONE_RATE"
									end if

                  If RET = 0 Then		
										If SAT_DELIVERY = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "SATURDAY_DELIVERY"											
										If SAT_PICKUP = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "SATURDAY_PICKUP"
                  End If

                 '// Options with related detail
									If COD = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "COD"
									If HAL = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "HOLD_AT_LOCATION"
									If ENS = 1 Then	objFedExClass.AddNewNode "SpecialServiceTypes", "EMAIL_NOTIFICATION"
									If RET = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "RETURN_SHIPMENT"
									If FICE = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "INTERNATIONAL_CONTROLLED_EXPORT_SERVICE"
									If ITAR = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "INTERNATIONAL_TRAFFIC_IN_ARMS_REGULATIONS"
                  If HDP = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "HOME_DELIVERY_PREMIUM"
									If ETD = 1 Then objFedExClass.AddNewNode "SpecialServiceTypes", "ELECTRONIC_TRADE_DOCUMENTS"

										'// Collect On Delivery
										if COD = 1 then
											objFedExClass.WriteParent "CodDetail", ""
												objFedExClass.WriteParent "CodCollectionAmount", ""
													objFedExClass.AddNewNode "Currency", Session("pcAdminCODCurrency")
													objFedExClass.AddNewNode "Amount", Session("pcAdminCODAmount")
												objFedExClass.WriteParent "CodCollectionAmount", "/"
												If Session("pcAdminCODRateType")&""<>"" Then
													objFedExClass.WriteParent "AddTransportationChargesDetail", ""
														objFedExClass.AddNewNode "RateTypeBasis", Session("pcAdminCODRateType")
														objFedExClass.AddNewNode "ChargeBasis", Session("pcAdminCODChargeBasis")
														objFedExClass.AddNewNode "ChargeBasisLevel", Session("pcAdminCODChargeBasisLevel")
													objFedExClass.WriteParent "AddTransportationChargesDetail", "/"
												End If
												objFedExClass.AddNewNode "CollectionType", Session("pcAdminCODType")
												objFedExClass.WriteParent "CodRecipient", ""
													objFedExClass.AddNewNode "AccountNumber", Session("pcAdminCODAccountNumber")
													
													If Session("pcAdminCODTinType")&""<>"" Then
														objFedExClass.WriteParent "Tins", ""
															objFedExClass.AddNewNode "TinType", Session("pcAdminCODTinType")
															objFedExClass.AddNewNode "Number", Session("pcAdminCODTinNumber")
														objFedExClass.WriteParent "Tins", "/"
													End If
													objFedExClass.WriteParent "Contact", ""
														objFedExClass.AddNewNode "PersonName", Session("pcAdminCODPersonName")
														objFedExClass.AddNewNode "Title", Session("pcAdminCODTitle")
														objFedExClass.AddNewNode "CompanyName", Session("pcAdminCODCompanyName")
														objFedExClass.AddNewNode "PhoneNumber", Session("pcAdminCODPhoneNumber")
													objFedExClass.WriteParent "Contact", "/"
													objFedExClass.WriteParent "Address", ""
														objFedExClass.AddNewNode "StreetLines", Session("pcAdminCODStreetLines")
														objFedExClass.AddNewNode "City", Session("pcAdminCODCity")
														objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminCODState")
														objFedExClass.AddNewNode "PostalCode", Session("pcAdminCODPostalCode")
														objFedExClass.AddNewNode "CountryCode", Session("pcAdminCODCountryCode")
													objFedExClass.WriteParent "Address", "/"
												objFedExClass.WriteParent "CodRecipient", "/"

											objFedExClass.WriteParent "CodDetail", "/"

										end if
                                        
										'// Hold At Location
										if HAL = 1 then
											objFedExClass.WriteParent "HoldAtLocationDetail", ""
												objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminHALPhone"))
												objFedExClass.WriteParent "LocationContactAndAddress", ""
													objFedExClass.WriteParent "Contact", ""
														objFedExClass.AddNewNode "ContactId", Session("pcAdminHALContactID")
														objFedExClass.AddNewNode "PersonName", Session("pcAdminHALPersonName")
														objFedExClass.AddNewNode "CompanyName", Session("pcAdminHALCompanyName")
														objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminHALPhone"))
														objFedExClass.AddNewNode "PhoneExtension", fnStripPhone(Session("pcAdminHALPhoneExtension"))
														objFedExClass.AddNewNode "PagerNumber", fnStripPhone(Session("pcAdminHALPager"))
														objFedExClass.AddNewNode "FaxNumber", fnStripPhone(Session("pcAdminHALFax"))
														objFedExClass.AddNewNode "EMailAddress", fnStripPhone(Session("pcAdminHALEmail"))
													objFedExClass.WriteParent "Contact", "/"
													objFedExClass.WriteParent "Address", ""
														objFedExClass.AddNewNode "StreetLines", Session("pcAdminHALLine1")
														objFedExClass.AddNewNode "City", Session("pcAdminHALCity")
														objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminHALStateOrProvinceCode")
														objFedExClass.AddNewNode "PostalCode", Session("pcAdminHALPostalCode")
														objFedExClass.AddNewNode "UrbanizationCode", Session("pcAdminHALUrbanizationCode")
														objFedExClass.AddNewNode "CountryCode", Session("pcAdminHALCountryCode")
														objFedExClass.AddNewNode "Residential", Session("pcAdminHALResidential")
													objFedExClass.WriteParent "Address", "/"
												objFedExClass.WriteParent "LocationContactAndAddress", "/"
												objFedExClass.AddNewNode "LocationType", Session("pcAdminHALLocationType")
											objFedExClass.WriteParent "HoldAtLocationDetail", "/"
										end if

                                        '// Email Notifications
										if ENS = 1 then	
											objFedExClass.WriteParent "EMailNotificationDetail", ""
												'objFedExClass.AddNewNode "PersonalMessage", "Personal Message Details"
												If pcTempNotificationShipperEmail = 1 Then
													objFedExClass.WriteParent "Recipients", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "SHIPPER"
														objFedExClass.AddNewNode "EMailAddress", Session("pcAdminNotificationShipperEmail")
														objFedExClass.AddNewNode "Format", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", ""
															objFedExClass.AddNewNode "LanguageCode", "EN"
														objFedExClass.WriteParent "Localization", "/"
													objFedExClass.WriteParent "Recipients", "/"
												End If
												If pcTempNotificationRecipientEmail = 1 Then
													objFedExClass.WriteParent "Recipients", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "RECIPIENT"
														objFedExClass.AddNewNode "EMailAddress", Session("pcAdminNotificationRecipientEmail")
														objFedExClass.AddNewNode "Format", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", ""
															objFedExClass.AddNewNode "LanguageCode", "EN"
														objFedExClass.WriteParent "Localization", "/"
													objFedExClass.WriteParent "Recipients", "/"
												End If
												If pcTempNotificationThirdPartyEmail = 1 Then
													objFedExClass.WriteParent "Recipients", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "THIRD_PARTY"
														objFedExClass.AddNewNode "EMailAddress", Session("pcAdminNotificationThirdPartyEmail")
														objFedExClass.AddNewNode "Format", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", ""
															objFedExClass.AddNewNode "LanguageCode", "EN"
														objFedExClass.WriteParent "Localization", "/"
													objFedExClass.WriteParent "Recipients", "/"
												End If
												If pcTempNotificationBrokerEmail = 1 Then
													objFedExClass.WriteParent "Recipients", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "BROKER"
														objFedExClass.AddNewNode "EMailAddress", Session("pcAdminNotificationBrokerEmail")
														objFedExClass.AddNewNode "Format", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", ""
															objFedExClass.AddNewNode "LanguageCode", "EN"
														objFedExClass.WriteParent "Localization", "/"
													objFedExClass.WriteParent "Recipients", "/"
												End If
												If pcTempNotificationOtherEmail = 1 Then
													objFedExClass.WriteParent "Recipients", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "OTHER"
														objFedExClass.AddNewNode "EMailAddress", Session("pcAdminNotificationOtherEmail")
														objFedExClass.AddNewNode "Format", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", ""
															objFedExClass.AddNewNode "LanguageCode", "EN"
														objFedExClass.WriteParent "Localization", "/"
													objFedExClass.WriteParent "Recipients", "/"
												End If
											objFedExClass.WriteParent "EMailNotificationDetail", "/"
										End If

                                        '// Return Shipment
										if RET = 1 then
											objFedExClass.WriteParent "ReturnShipmentDetail", ""
												objFedExClass.AddNewNode "ReturnType", "PRINT_RETURN_LABEL"

                        If Session("pcAdminReturnShipmentReason")&""<>"" Then
													objFedExClass.WriteParent "Rma", ""
														objFedExClass.AddNewNode "Reason", Session("pcAdminReturnShipmentReason")
													objFedExClass.WriteParent "Rma", "/"
                        End If
												
												if Session("pcAdminSaturdayDelivery")<>"0" or Session("pcAdminSaturdayPickup")<>"0" then
													objFedExClass.WriteParent "ReturnEMailDetail", ""
													if Session("pcAdminSaturdayDelivery")<>"0" then
														objFedExClass.AddNewNode "AllowedSpecialServices", "SATURDAY_DELIVERY"
													end if
													if Session("pcAdminSaturdayPickup")<>"0" then
														objFedExClass.AddNewNode "AllowedSpecialServices", "SATURDAY_PICKUP"
													end if
													objFedExClass.WriteParent "ReturnEMailDetail", "/"
												end if
											
											objFedExClass.WriteParent "ReturnShipmentDetail", "/"
										end if

										'// FICE
										If FICE = 1 Then
											objFedExClass.WriteParent "InternationalControlledExportDetail", ""
												objFedExClass.AddNewNode "Type", Session("pcAdminFICEType")
												objFedExClass.AddNewNode "ForeignTradeZoneCode", Session("pcAdminFICETradeZoneCode")
												objFedExClass.AddNewNode "EntryNumber", Session("pcAdminFICEEntryNumber")
												objFedExClass.AddNewNode "LicenseOrPermitNumber", Session("pcAdminFICENumber")
												objFedExClass.AddNewNode "LicenseOrPermitExpirationDate", Session("pcAdminFICEExpirationDate")
											objFedExClass.WriteParent "InternationalControlledExportDetail", "/"
										End If
										
										'// ITAR
										If ITAR = 1 Then
											objFedExClass.WriteParent "InternationalTrafficInArmsRegulationsDetail", ""
												objFedExClass.AddNewNode "LicenseOrExemptionNumber", Session("pcAdminITARNumber")
											objFedExClass.WriteParent "InternationalTrafficInArmsRegulationsDetail", "/"
										End If

										'// Home Delivery Premium
										if HDP = 1 Then
											objFedExClass.WriteParent "HomeDeliveryPremiumDetail", ""
												objFedExClass.AddNewNode "HomeDeliveryPremiumType", Session("pcAdminHomeDeliveryType")
												objFedExClass.AddNewNode "Date", objFedExClass.pcf_FedExDateFormat(CDate(Session("pcAdminHomeDeliveryDate")))
												objFedExClass.AddNewNode "PhoneNumber", Session("pcAdminHomeDeliveryPhone")
											objFedExClass.WriteParent "HomeDeliveryPremiumDetail", "/"
										end if

										'// Electronic Trade Documents
										if ETD = 1 Then
											objFedExClass.WriteParent "EtdDetail", ""
												objFedExClass.AddNewNode "RequestedDocumentCopies", Session("pcAdminETDRequestedDocumentCopies")
											  'objFedExClass.WriteParent "DocumentReferences", ""
												'  objFedExClass.AddNewNode "LineNumber", Session("pcAdminETDLineNumber")
												'  objFedExClass.AddNewNode "DocumentProducer", Session("pcAdminETDDocumentProducer")
												'  objFedExClass.AddNewNode "DocumentType", "ETD_LABEL"
												'  objFedExClass.AddNewNode "DocumentId", Session("pcAdminETDDocumentId")
												'  objFedExClass.AddNewNode "DocumentIdProducer", Session("pcAdminETDDocumentIdProducer")
											  'objFedExClass.WriteParent "DocumentReferences", "/"
											objFedExClass.WriteParent "EtdDetail", "/"
										end if

									objFedExClass.WriteParent "SpecialServicesRequested", "/"
								End If

								'---------------------------------------------
								'// SMARTPOST
								'---------------------------------------------
								if mySP = 1 Then
									objFedExClass.WriteParent "SmartPostDetail", ""
										objFedExClass.AddNewNode "Indicia", Session("pcAdminSMIndicia")
										objFedExClass.AddNewNode "AncillaryEndorsement", Session("pcAdminSMAncillaryEndorsement")
										objFedExClass.AddNewNode "HubId", Session("pcAdminSMHubID")
									objFedExClass.WriteParent "SmartPostDetail", "/"
								end if
								'---------------------------------------------
								'// RequestedShipment/CustomsClearanceDetail
								'---------------------------------------------
								
								If Session("pcAdminEFShippersLoadAndCount")&"" <>"" Then
									objFedExClass.WriteParent "ExpressFreightDetail", ""
										objFedExClass.AddNewNode "PackingListEnclosed", Session("pcAdminEFPackingListEnclosed")
										objFedExClass.AddNewNode "ShippersLoadAndCount", Session("pcAdminEFShippersLoadAndCount")
										objFedExClass.AddNewNode "BookingConfirmationNumber", Session("pcAdminEFBookingConfirmationNumber")
									objFedExClass.WriteParent "ExpressFreightDetail", "/"
								End If
								

								'// Freight Details
								If Session("pcAdminService1")="FEDEX_FREIGHT_PRIORITY" Or Session("pcAdminService1")="FEDEX_FREIGHT_ECONOMY" Then
									objFedExClass.WriteParent "FreightShipmentDetail", ""
										objFedExClass.AddNewNode "FedExFreightAccountNumber", Session("pcAdminFreightAccountNumber")
										
										' Billing Contact and Address
										objFedExClass.WriteParent "FedExFreightBillingContactAndAddress", ""
											objFedExClass.WriteParent "Contact", ""
												objFedExClass.AddNewNode "PersonName", Session("pcAdminFreightContactPersonName")
												objFedExClass.AddNewNode "CompanyName", Session("pcAdminFreightContactCompanyName")
												objFedExClass.AddNewNode "PagerNumber", Session("pcAdminFreightPagerNumber")
											objFedExClass.WriteParent "Contact", "/"
											objFedExClass.WriteParent "Address", ""
												objFedExClass.AddNewNode "StreetLines", Session("pcAdminFreightContactStreetLines")
												objFedExClass.AddNewNode "City", Session("pcAdminFreightContactCity")
												objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminFreightContactStateCode")
												objFedExClass.AddNewNode "PostalCode", Session("pcAdminFreightContactPostalCode")
												objFedExClass.AddNewNode "CountryCode", Session("pcAdminFreightContactCountryCode")
											objFedExClass.WriteParent "Address", "/"
										objFedExClass.WriteParent "FedExFreightBillingContactAndAddress", "/"
										
										' Role and Collect Terms Type
										objFedExClass.AddNewNode "Role", Session("pcAdminFreightShipmentRoleType")
										objFedExClass.AddNewNode "CollectTermsType", Session("pcAdminFreightCollectTermsType")
										
										' Declared Value
										If Len(Session("pcAdminFreightDVAmount")) > 0 Then
											objFedExClass.WriteParent "DeclaredValuePerUnit", ""
												objFedExClass.AddNewNode "Currency", Session("pcAdminFreightDVCurrency")
												objFedExClass.AddNewNode "Amount", Session("pcAdminFreightDVAmount")
											objFedExClass.WriteParent "DeclaredValuePerUnit", "/"
											objFedExClass.AddNewNode "DeclaredValueUnits", Session("pcAdminFreightDVUnits")
										End If

										' Liability Coverage
										If Len(Session("pcAdminFreightLCAmount")) > 0 Then
											objFedExClass.WriteParent "LiabilityCoverageDetail", ""
												objFedExClass.AddNewNode "CoverageType", Session("pcAdminFreightLCType")
												objFedExClass.WriteParent "CoverageAmount", ""
													objFedExClass.AddNewNode "Currency", Session("pcAdminFreightLCCurrency")
													objFedExClass.AddNewNode "Amount", Session("pcAdminFreightLCAmount")
												objFedExClass.WriteParent "CoverageAmount", "/"
											objFedExClass.WriteParent "LiabilityCoverageDetail", "/"
										End If

										' Optional Items
										objFedExClass.AddNewNode "TotalHandlingUnits", Session("pcAdminFreightTotalHandlingUnits")
										objFedExClass.AddNewNode "ClientDiscountPercent", Session("pcAdminFreightClientDiscountPercent")
										If Len(Session("pcAdminFreightPalletWeightValue")) > 0 Then
											objFedExClass.WriteParent "PalletWeight", ""
												objFedExClass.AddNewNode "Units", Session("pcAdminFreightPalletWeightUnits")
												objFedExClass.AddNewNode "Value", Session("pcAdminFreightPalletWeightValue")
											objFedExClass.WriteParent "PalletWeight", "/"
										End If
										If Len(Session("pcAdminFreightShipmentDimensionsLength")) > 0 Then
											objFedExClass.WriteParent "ShipmentDimensions", ""
												objFedExClass.AddNewNode "Length", Session("pcAdminFreightShipmentDimensionsLength")
												objFedExClass.AddNewNode "Width", Session("pcAdminFreightShipmentDimensionsWidth")
												objFedExClass.AddNewNode "Height", Session("pcAdminFreightShipmentDimensionsHeight")
												objFedExClass.AddNewNode "Units", Session("pcAdminFreightShipmentDimensionsUnits")
											objFedExClass.WriteParent "ShipmentDimensions", "/"
										End If

										objFedExClass.AddNewNode "Comment", Session("pcAdminFreightShipmentComment")
										
										' Line Items
										objFedExClass.WriteParent "LineItems", ""
											objFedExClass.AddNewNode "FreightClass", Session("pcAdminFreightLIClass")
											objFedExClass.AddNewNode "ClassProvidedByCustomer", Session("pcAdminFreightLIClassProvided")
											objFedExClass.AddNewNode "HandlingUnits", Session("pcAdminFreightLIHandlingUnits")
											objFedExClass.AddNewNode "Packaging", Session("pcAdminFreightLIPackaging")
											objFedExClass.AddNewNode "Pieces", Session("pcAdminFreightLIPieces")
											objFedExClass.AddNewNode "PurchaseOrderNumber", Session("pcAdminFreightLIPONumber")
											objFedExClass.AddNewNode "Description", Session("pcAdminFreightLIDescription")
											If Len(Session("pcAdminFreightLIWeightValue")) > 0 Then
												objFedExClass.WriteParent "Weight", ""
													objFedExClass.AddNewNode "Units", Session("pcAdminFreightLIWeightUnits")
													objFedExClass.AddNewNode "Value", Session("pcAdminFreightLIWeightValue")
												objFedExClass.WriteParent "Weight", "/"
											End If

											If Len(Session("pcAdminFreightLIDimensionsLength")) > 0 Then
												objFedExClass.WriteParent "Dimensions", ""
													objFedExClass.AddNewNode "Length", Session("pcAdminFreightLIDimensionsLength")
													objFedExClass.AddNewNode "Width", Session("pcAdminFreightLIDimensionsWidth")
													objFedExClass.AddNewNode "Height", Session("pcAdminFreightLIDimensionsHeight")
													objFedExClass.AddNewNode "Units", Session("pcAdminFreightLIDimensionsUnits")
												objFedExClass.WriteParent "Dimensions", "/"
											End If
										objFedExClass.WriteParent "LineItems", "/"

									objFedExClass.WriteParent "FreightShipmentDetail", "/"

									objFedExClass.AddNewNode "DeliveryInstructions", Session("pcAdminDeliveryInstructions")
								End If
				
								'// Delivery Instructions (for Home Delivery)
								If Session("pcAdminService1")="GROUND_HOME_DELIVERY" Then
									objFedExClass.AddNewNode "DeliveryInstructions", Session("pcAdminHomeDeliveryInstructions")
								End If

								CCD = 0

								If BSO = 1 OR Session("pcAdminRCIdValue")&""<>"" OR Session("pcAdminDutiesAccountNumber")&"" <> "" OR Session("pcAdminCVAmount")&"" <> "" OR Session("pcAdminCICAmount")&"" <> "" OR Session("pcAdminCFCAmount")&"" <> "" OR Session("pcAdminNumberOfPieces")>0 OR Session("pcAdminB13AFilingOption")&"" <> "" OR Session("pcAdminbNAFTAOption")="1" Then
									CCD = 1
								End if

								if CCD = 1 And Session("pcAdminbInternational")="true" then
									objFedExClass.WriteParent "CustomsClearanceDetail", ""
										If BSO = 1 Then
											objFedExClass.WriteParent "Brokers", ""
												objFedExClass.AddNewNode "Type", Session("pcAdminBSOType")
													
												objFedExClass.WriteParent "Broker", ""
													objFedExClass.AddNewNode "AccountNumber", Session("pcAdminBSOAccountNumber")
												
													objFedExClass.WriteParent "Tins", ""
														objFedExClass.AddNewNode "TinType", Session("pcAdminBSOTinType")
														objFedExClass.AddNewNode "Number", Session("pcAdminBSOTinNumber")
													objFedExClass.WriteParent "Tins", "/"
													objFedExClass.WriteParent "Contact", ""
														objFedExClass.AddNewNode "ContactId", Session("pcAdminBSOContactID")
														objFedExClass.AddNewNode "PersonName", Session("pcAdminBSOPersonName")
														objFedExClass.AddNewNode "Title", Session("pcAdminBSOTitle")
														objFedExClass.AddNewNode "CompanyName", Session("pcAdminBSOCompanyName")
														objFedExClass.AddNewNode "PhoneNumber", Session("pcAdminBSOPhoneNumber")
														objFedExClass.AddNewNode "PhoneExtension", Session("pcAdminBSOPhoneExtension")
														objFedExClass.AddNewNode "EMailAddress", Session("pcAdminBSOEmailAddress")
													objFedExClass.WriteParent "Contact", "/"
													objFedExClass.WriteParent "Address", ""
														objFedExClass.AddNewNode "StreetLines", Session("pcAdminBSOStreetLines")
														objFedExClass.AddNewNode "City", Session("pcAdminBSOCity")
														objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminBSOStateOrProvinceCode")
														objFedExClass.AddNewNode "PostalCode", Session("pcAdminBSOPostalCode")
														objFedExClass.AddNewNode "CountryCode", Session("pcAdminBSOCountryCode")
													objFedExClass.WriteParent "Address", "/"
												objFedExClass.WriteParent "Broker", "/"
											objFedExClass.WriteParent "Brokers", "/"
										End If
										
										If Session("pcAdminCCDOptionType")&"" <> "" Then
											objFedExClass.WriteParent "CustomsOptions", ""
												objFedExClass.AddNewNode "Type", Session("pcAdminCCDOptionType")
												objFedExClass.AddNewNode "Description", Session("pcAdminCCDOptionDescription")
											objFedExClass.WriteParent "CustomsOptions", "/"
										End If

										'// Importer of Record
										If Session("pcAdminIORPersonName")&"" <> "" Then
											objFedExClass.WriteParent "ImporterOfRecord", ""
												objFedExClass.WriteParent "Contact", ""
													objFedExClass.AddNewNode "PersonName", Session("pcAdminIORPersonName")
													objFedExClass.AddNewNode "CompanyName", Session("pcAdminIORCompanyName")
													objFedExClass.AddNewNode "PhoneNumber", Session("pcAdminIORPhoneNumber")
												objFedExClass.WriteParent "Contact", "/"
												objFedExClass.WriteParent "Address", ""
													objFedExClass.AddNewNode "StreetLines", Session("pcAdminIORAddress")
													objFedExClass.AddNewNode "City", Session("pcAdminIORCity")
													objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminIORStateOrProvince")
													objFedExClass.AddNewNode "PostalCode", Session("pcAdminIORPostalCode")
													objFedExClass.AddNewNode "CountryCode", Session("pcAdminIORCountryCode")
												objFedExClass.WriteParent "Address", "/"
											objFedExClass.WriteParent "ImporterOfRecord", "/"
										End If

										If Session("pcAdminRCIdValue")&""<>"" Then
											objFedExClass.WriteParent "RecipientCustomsId", ""
												objFedExClass.AddNewNode "Type", Session("pcAdminRCIdType")
												objFedExClass.AddNewNode "Value", Session("pcAdminRCIdValue")
											objFedExClass.WriteParent "RecipientCustomsId", "/"
										End If

										If Session("pcAdminDutiesAccountNumber")&"" <> "" Then
											objFedExClass.WriteParent "DutiesPayment", ""
												objFedExClass.AddNewNode "PaymentType", Session("pcAdminDutiesPayorType")
												objFedExClass.WriteParent "Payor", ""
													objFedExClass.WriteParent "ResponsibleParty", ""
														objFedExClass.AddNewNode "AccountNumber", Session("pcAdminDutiesAccountNumber")
														objFedExClass.WriteParent "Contact", ""
															objFedExClass.AddNewNode "PersonName", Session("pcAdminDutiesPersonName")
														objFedExClass.WriteParent "Contact", "/"
														objFedExClass.WriteParent "Address", ""
															objFedExClass.AddNewNode "CountryCode", Session("pcAdminDutiesCountryCode")
														objFedExClass.WriteParent "Address", "/"
													objFedExClass.WriteParent "ResponsibleParty", "/"
												objFedExClass.WriteParent "Payor", "/"
											objFedExClass.WriteParent "DutiesPayment", "/"
										End If

										If Session("pcAdminDocumentsOnly") <> "0" Then
											if Session("pcAdminDocumentsOnly")="1" Then
												objFedExClass.AddNewNode "DocumentContent", "NON_DOCUMENTS"
											else
												objFedExClass.AddNewNode "DocumentContent", "DOCUMENTS_ONLY"
											end if
										End If

										If Session("pcAdminCVAmount")&"" <> "" Then
											objFedExClass.WriteParent "CustomsValue", ""
												objFedExClass.AddNewNode "Currency", Session("pcAdminCVCurrency")
												objFedExClass.AddNewNode "Amount", Session("pcAdminCVAmount")
											objFedExClass.WriteParent "CustomsValue", "/"
										End If

										If Session("pcAdminCICAmount")&"" <> "" Then
											objFedExClass.WriteParent "InsuranceCharges", ""
												objFedExClass.AddNewNode "Currency", Session("pcAdminCICCurrency")
												objFedExClass.AddNewNode "Amount", Session("pcAdminCICAmount")
											objFedExClass.WriteParent "InsuranceCharges", "/"
										End If

										If Session("pcAdminCFCAmount")&"" <> "" OR session("pcAdminCCITermsOfSale")&""<>"" Then
											objFedExClass.WriteParent "CommercialInvoice", ""
												objFedExClass.AddNewNode "Comments", Session("pcAdminCCIComments")
												If Session("pcAdminCFCAmount")&"" <> "" Then
													objFedExClass.WriteParent "FreightCharge", ""
														objFedExClass.AddNewNode "Currency", Session("pcAdminCFCCurrency")
														objFedExClass.AddNewNode "Amount", Session("pcAdminCFCAmount")
													objFedExClass.WriteParent "FreightCharge", "/"
												End If
				
												If Session("pcAdminCMCAmount")&"" <> "" Then
													objFedExClass.WriteParent "TaxesOrMiscellaneousCharge", ""
														objFedExClass.AddNewNode "Currency", Session("pcAdminCMCCurrency")
														objFedExClass.AddNewNode "Amount", Session("pcAdminCMCAmount")
													objFedExClass.WriteParent "TaxesOrMiscellaneousCharge", "/"
												End If

												objFedExClass.AddNewNode "Purpose", Session("pcAdminCCIPurpose")
												objFedExClass.AddNewNode "CustomerInvoiceNumber", Session("pcAdminCCIInvoiceNumber")

												If session("pcAdminCCICustomerReference")&""<>"" Then
													objFedExClass.WriteParent "CustomerReferences", ""
														objFedExClass.AddNewNode "CustomerReferenceType", "CUSTOMER_REFERENCE"
														objFedExClass.AddNewNode "Value", Session("pcAdminCCICustomerReference")
													objFedExClass.WriteParent "CustomerReferences", "/"
												End If

												If session("pcAdminCCITermsOfSale")&""<>"" Then
													objFedExClass.AddNewNode "TermsOfSale", session("pcAdminCCITermsOfSale")
												End If
											objFedExClass.WriteParent "CommercialInvoice", "/"
										End If

										If Session("pcAdminNumberOfPieces")>0 Then
											objFedExClass.WriteParent "Commodities", ""
												objFedExClass.AddNewNode "NumberOfPieces", Session("pcAdminNumberOfPieces")
												objFedExClass.AddNewNode "Description", Session("pcAdminDescription")
												objFedExClass.AddNewNode "CountryOfManufacture", Session("pcAdminCountryOfManufacture")
												objFedExClass.WriteParent "Weight", ""
													objFedExClass.AddNewNode "Units", Session("pcAdminCommodityWeightUnits")
													objFedExClass.AddNewNode "Value", Session("pcAdminCommodityWeightValue")
												objFedExClass.WriteParent "Weight", "/"
												objFedExClass.AddNewNode "Quantity", Session("pcAdminCommodityQuantity")
												objFedExClass.AddNewNode "QuantityUnits", Session("pcAdminCommodityQuantityUnits")

												objFedExClass.WriteParent "UnitPrice", ""
													objFedExClass.AddNewNode "Currency", Session("pcAdminCommodityUnitCurrency")
													objFedExClass.AddNewNode "Amount", Session("pcAdminCommodityUnitPrice")
												objFedExClass.WriteParent "UnitPrice", "/"
												objFedExClass.WriteParent "CustomsValue", ""
													objFedExClass.AddNewNode "Currency", Session("pcAdminCommodityCustomsCurrency")
													objFedExClass.AddNewNode "Amount", Session("pcAdminCommodityCustomsValue")
												objFedExClass.WriteParent "CustomsValue", "/"
										
												'// NAFTA Stuff
												If Session("pcAdminbNAFTAOption")="1" Then
													objFedExClass.WriteParent "NaftaDetail", ""
														objFedExClass.AddNewNode "PreferenceCriterion", Session("pcAdminNAFTAPreferenceCriterion")
														objFedExClass.AddNewNode "ProducerDetermination", Session("pcAdminNAFTAProducerDetermination")
														objFedExClass.AddNewNode "ProducerId", Session("pcAdminNAFTAProducerID")
														objFedExClass.AddNewNode "NetCostMethod", Session("pcAdminNAFTANetCostMethod")
													objFedExClass.WriteParent "NaftaDetail", "/"
												End If
										
											objFedExClass.WriteParent "Commodities", "/"
										End If


										If Session("pcAdminB13AFilingOption")&"" <> "" Then
											objFedExClass.WriteParent "ExportDetail", ""
												objFedExClass.AddNewNode "B13AFilingOption", Session("pcAdminB13AFilingOption")
												objFedExClass.AddNewNode "ExportComplianceStatement", Session("pcAdminExportComplianceStatement")
											objFedExClass.WriteParent "ExportDetail", "/"
										End If
										
										'// Regulatory Types
										If Session("pcAdminbNAFTAOption")="1" Then
											objFedExClass.AddNewNode "RegulatoryControls", "NAFTA"
										End If

									objFedExClass.WriteParent "CustomsClearanceDetail", "/"
								end if
								
								'---------------------------------------------
								'// RequestedShipment/LabelSpecification
								'---------------------------------------------
								objFedExClass.WriteParent "LabelSpecification", ""
									objFedExClass.AddNewNode "LabelFormatType", Session("pcAdminLabelFormatType")
									objFedExClass.AddNewNode "ImageType", Session("pcAdminLabelImageType")
									objFedExClass.AddNewNode "LabelStockType", Session("pcAdminLabelStockType")
									objFedExClass.AddNewNode "LabelPrintingOrientation", Session("pcAdminLabelPrintingOrientation")
								objFedExClass.WriteParent "LabelSpecification", "/"

								'---------------------------------------------
								'// RequestedShipment/ShippingDocumentSpecification
								'---------------------------------------------
								If Session("pcAdminbSOShippingDocumentOption") = "1" Then
									objFedExClass.WriteParent "ShippingDocumentSpecification", ""
										objFedExClass.AddNewNode "ShippingDocumentTypes", Session("pcAdminSDSType")
										
                    Select Case Session("pcAdminSDSType")
                    Case "CERTIFICATE_OF_ORIGIN":
										  objFedExClass.WriteParent "CertificateOfOrigin", ""
											  objFedExClass.WriteParent "DocumentFormat", ""
												  objFedExClass.AddNewNode "ImageType", Session("pcAdminSDSLabelImageType")
												  objFedExClass.AddNewNode "StockType", Session("pcAdminSDSLabelStockType")
											  objFedExClass.WriteParent "DocumentFormat", "/"
										  objFedExClass.WriteParent "CertificateOfOrigin", "/"
                    Case "COMMERCIAL_INVOICE":
										  objFedExClass.WriteParent "CommercialInvoiceDetail", ""
											  objFedExClass.WriteParent "Format", ""
												  objFedExClass.AddNewNode "ImageType", Session("pcAdminSDSLabelImageType")
												  objFedExClass.AddNewNode "StockType", Session("pcAdminSDSLabelStockType")
											  objFedExClass.WriteParent "Format", "/"
										  objFedExClass.WriteParent "CommercialInvoiceDetail", "/"
                    Case "FREIGHT_ADDRESS_LABEL":
										  objFedExClass.WriteParent "FreightAddressLabelDetail", ""
											  objFedExClass.WriteParent "Format", ""
												  objFedExClass.AddNewNode "ImageType", Session("pcAdminSDSLabelImageType")
												  objFedExClass.AddNewNode "StockType", Session("pcAdminSDSLabelStockType")
												  objFedExClass.AddNewNode "ProvideInstructions", Session("pcAdminSDSProvideInstructions")
											  objFedExClass.WriteParent "Format", "/"
										  objFedExClass.WriteParent "FreightAddressLabelDetail", "/"
                    End Select
								
									objFedExClass.WriteParent "ShippingDocumentSpecification", "/"
								End If
								
								objFedExClass.WriteSingleParent "RateRequestTypes",  Session("pcAdminRateRequestType")

								'---------------------------------------------
								'// RequestedShipment/MasterTrackingId
								'---------------------------------------------
								if cint(pcPackageCount) > 1 then
									objFedExClass.WriteParent "MasterTrackingId", ""
										if pcv_xCounter>1 then
											'// Required for multiple-piece shipping if PackageSequenceNumber value is greater than one.
											objFedExClass.AddNewNode "TrackingNumber", Session("MasterTrackingNumber")

										end if
									objFedExClass.WriteParent "MasterTrackingId", "/"
								end if

								'-------------------------------------------------
								'// RequestedShipment/PackageCount
								'-------------------------------------------------
								objFedExClass.WriteSingleParent "PackageCount", pcPackageCount

								'-------------------------------------------------
								'// RequestedShipment/PackageDetail
								'-------------------------------------------------
								'objFedExClass.AddNewNode "PackageDetail", "INDIVIDUAL_PACKAGES"

								'-------------------------------------------------
								'// RequestedShipment/RequestedPackageLineItems
								'-------------------------------------------------
								objFedExClass.WriteParent "RequestedPackageLineItems", ""
									objFedExClass.AddNewNode "SequenceNumber", pcv_xCounter
									IF mySP = 0 THEN
										objFedExClass.WriteParent "InsuredValue", ""
											objFedExClass.AddNewNode "Currency", Session("pcAdmincurrency"&pcv_xCounter)
											objFedExClass.AddNewNode "Amount", Session("pcAdmindeclaredvalue"&pcv_xCounter)
										objFedExClass.WriteParent "InsuredValue", "/"
									END IF
									objFedExClass.WriteParent "Weight", ""
										pcvTempWeightUnit = ""
										if Session("pcAdminWeightUnits"&pcv_xCounter) = "LB" then
											pcvTempWeightUnit = "LB"
										else
											pcvTempWeightUnit = "KG"
										end if
										objFedExClass.AddNewNode "Units", pcvTempWeightUnit
										IF mySP = 0 THEN
											objFedExClass.AddNewNode "Value", Session("pcAdminWeight"&pcv_xCounter)
										Else
											objFedExClass.AddNewNode "Value", Session("pcAdminWeight"&pcv_xCounter)
										End If
									objFedExClass.WriteParent "Weight", "/"

									If TRIM(Session("pcAdminPackaging1"))="YOUR_PACKAGING" then
										objFedExClass.WriteParent "Dimensions", ""
											objFedExClass.AddNewNode "Length", Session("pcAdminLength"&pcv_xCounter) '"12"
											objFedExClass.AddNewNode "Width", Session("pcAdminWidth"&pcv_xCounter) '"13"
											objFedExClass.AddNewNode "Height", Session("pcAdminHeight"&pcv_xCounter) '"14"
											objFedExClass.AddNewNode "Units", Session("pcAdminUnits"&pcv_xCounter) '"IN"
										objFedExClass.WriteParent "Dimensions", "/"
									End If

									' START: CUSTOMER REFERENCES
									objFedExClass.WriteParent "CustomerReferences", ""
										objFedExClass.AddNewNode "CustomerReferenceType", "CUSTOMER_REFERENCE"
										objFedExClass.AddNewNode "Value", Session("pcAdminCustomerReference")
									objFedExClass.WriteParent "CustomerReferences", "/"
									If Session("pcAdminCustomerInvoiceNumber")&""<>"" Then
									objFedExClass.WriteParent "CustomerReferences", ""
										objFedExClass.AddNewNode "CustomerReferenceType", "INVOICE_NUMBER"
										objFedExClass.AddNewNode "Value", Session("pcAdminCustomerInvoiceNumber")
									objFedExClass.WriteParent "CustomerReferences", "/"
									End If
									If Session("pcAdminCustomerPONumber")&""<>"" Then
									objFedExClass.WriteParent "CustomerReferences", ""
										objFedExClass.AddNewNode "CustomerReferenceType", "P_O_NUMBER"
										objFedExClass.AddNewNode "Value", Session("pcAdminCustomerPONumber")
									objFedExClass.WriteParent "CustomerReferences", "/"
									End If
									' START: SPECIAL SERVICES
									
									DG=0
									if Session("pcAdminbSODGShip")<>"0" then
										DG=1
									end if

									if DG=1 OR (Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService"&pcv_xCounter)="FEDEX_GROUND") OR Session("pcAdminbSODGShip")<>"0" OR Session("pcAdminbSODryIce")="1" OR Session("pcAdminContainerType"&pcv_xCounter)="1" OR Session("pcAdminPriorityAlert")&""<>"" OR Session("pcAdminPriorityAlertPlus")&""<>"" OR Session("pcAdminSignatureOption")&""<>"" OR Session("pcAdminbSOAlcoholOption")&""<>"" then
										objFedExClass.WriteParent "SpecialServicesRequested", ""
											
											'// Alcohol
											if Session("pcAdminbSOAlcoholOption")<>"0" then
													objFedExClass.AddNewNode "SpecialServiceTypes", "ALCOHOL"
											end if

											if Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService"&pcv_xCounter)="FEDEX_GROUND" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "COD"
											end if

											IF DG=1 THEN
												objFedExClass.AddNewNode "SpecialServiceTypes", "DANGEROUS_GOODS"
											END IF

											if Session("pcAdminbSODryIce")="1" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "DRY_ICE"
											end if

											If Session("pcAdminContainerType"&pcv_xCounter)="1" Then
												objFedExClass.AddNewNode "SpecialServiceTypes", "NON_STANDARD_CONTAINER"
											End If
                                            
											if Session("pcAdminSignatureOption")&""<>"" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "SIGNATURE_OPTION"
											end if

											'//priority alert
											if Session("pcAdminPriorityAlert")&""<>"" OR Session("pcAdminPriorityAlertPlus")&""<>"" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "PRIORITY_ALERT"
											end if

											'//COD FOR FEDEX GROUND
											if Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService"&pcv_xCounter)="FEDEX_GROUND" then
												objFedExClass.WriteParent "CodDetail", ""
													objFedExClass.WriteParent "CodCollectionAmount", ""
														objFedExClass.AddNewNode "Currency", Session("pcAdminCODCurrency") 'Session("asdf")
														objFedExClass.AddNewNode "Amount", Session("pcAdminCODAmount")
													objFedExClass.WriteParent "CodCollectionAmount", "/"

													If Session("pcAdminCODRateType")&""<>"" Then
														objFedExClass.WriteParent "AddTransportationChargesDetail", ""
															objFedExClass.AddNewNode "RateTypeBasis", Session("pcAdminCODRateType")
															objFedExClass.AddNewNode "ChargeBasis", Session("pcAdminCODChargeBasis")
															objFedExClass.AddNewNode "ChargeBasisLevel", Session("pcAdminCODChargeBasisLevel")
														objFedExClass.WriteParent "AddTransportationChargesDetail", "/"
													End If

													objFedExClass.AddNewNode "CollectionType", Session("pcAdminCODType")
													objFedExClass.WriteParent "CodRecipient", ""													
														objFedExClass.AddNewNode "AccountNumber", Session("pcAdminCODAccountNumber")
													
														If Session("pcAdminCODTinType")&""<>"" Then
															objFedExClass.WriteParent "Tins", ""
																objFedExClass.AddNewNode "TinType", Session("pcAdminCODTinType")
																objFedExClass.AddNewNode "Number", Session("pcAdminCODTinNumber")
															objFedExClass.WriteParent "Tins", "/"
														End If
														objFedExClass.WriteParent "Contact", ""
															objFedExClass.AddNewNode "PersonName", Session("pcAdminCODPersonName")
															objFedExClass.AddNewNode "Title", Session("pcAdminCODTitle")
															objFedExClass.AddNewNode "CompanyName", Session("pcAdminCODCompanyName")
															objFedExClass.AddNewNode "PhoneNumber", Session("pcAdminCODPhoneNumber")
														objFedExClass.WriteParent "Contact", "/"
														objFedExClass.WriteParent "Address", ""
															objFedExClass.AddNewNode "StreetLines", Session("pcAdminCODStreetLines")
															objFedExClass.AddNewNode "City", Session("pcAdminCODCity")
															objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminCODState")
															objFedExClass.AddNewNode "PostalCode", Session("pcAdminCODPostalCode")
															objFedExClass.AddNewNode "CountryCode", Session("pcAdminCODCountryCode")
															objFedExClass.AddNewNode "Residential", "false"
														objFedExClass.WriteParent "Address", "/"
													objFedExClass.WriteParent "CodRecipient", "/"
													objFedExClass.AddNewNode "ReferenceIndicator", "INVOICE"
												objFedExClass.WriteParent "CodDetail", "/"
											end if

											'// Dangerous Goods
											IF dg=1 THEN
												objFedExClass.WriteParent "DangerousGoodsDetail", ""
													If session("pcAdminDGAccessibility")&""<>"" Then
														objFedExClass.AddNewNode "Accessibility", session("pcAdminDGAccessibility")
														objFedExClass.AddNewNode "CargoAircraftOnly", session("pcAdminDGAircraftOnly")
													End If
													If session("pcAdminDGHazardousMaterials")="1" Then
														objFedExClass.AddNewNode "Options", "HAZARDOUS_MATERIALS"
													End If
													If session("pcAdminDGBattery")="1" Then
														objFedExClass.AddNewNode "Options", "BATTERY"
													End If
													If session("pcAdminDGORMD")="1" Then
														objFedExClass.AddNewNode "Options", "ORM_D"
													End If
													'// Add containers
													if session("pcAdminDGContainerCount") > 0 then
														objFedExClass.WriteParent "Containers", ""
															objFedExClass.AddNewNode "ContainerType", session("pcAdminDGContainerType")
															objFedExClass.AddNewNode "NumberOfContainers", session("pcAdminDGContainerCount")
															for i = 1 to session("pcAdminDGContainerCount")
																objFedExClass.WriteParent "HazardousCommodities", ""
																	objFedExClass.WriteParent "Description", ""
																		objFedExClass.AddNewNode "Id", session("pcAdminDGCommodityID" & i)
																		objFedExClass.AddNewNode "SequenceNumber", i
																		objFedExClass.AddNewNode "PackingGroup", session("pcAdminDGPackingGroup" & i)
																		objFedExClass.WriteParent "PackingDetails", ""
																			objFedExClass.AddNewNode "CargoAircraftOnly", session("pcAdminDGContainerAircraftOnly" & i)
																			objFedExClass.AddNewNode "PackingInstructions", session("pcAdminDGPackingInstructions" & i)
																		objFedExClass.WriteParent "PackingDetails", "/"
																		objFedExClass.AddNewNode "ProperShippingName", session("pcAdminDGShippingName" & i)
																		objFedExClass.AddNewNode "HazardClass", session("pcAdminDGHazardClass" & i)
																	objFedExClass.WriteParent "Description", "/"
																	objFedExClass.WriteParent "Quantity", ""
																		objFedExClass.AddNewNode "Amount", session("pcAdminDGQuantityAmount" & i)
																		objFedExClass.AddNewNode "Units", session("pcAdminDGQuantityUnits" & i)
																	objFedExClass.WriteParent "Quantity", "/"
																objFedExClass.WriteParent "HazardousCommodities", "/"
															next
														objFedExClass.WriteParent "Containers", "/"
													end if
													If session("pcAdminDGPackagingUnits")&""<>"" Then
													objFedExClass.WriteParent "Packaging", ""
														objFedExClass.AddNewNode "Count", session("pcAdminDGPackagingCount")
														objFedExClass.AddNewNode "Units", session("pcAdminDGPackagingUnits")
													objFedExClass.WriteParent "Packaging", "/"
													End If
													
													objFedExClass.WriteParent "Signatory", ""
														objFedExClass.AddNewNode "ContactName", session("pcAdminDGContactName")
														objFedExClass.AddNewNode "Title", session("pcAdminDGContactTitle")
														objFedExClass.AddNewNode "Place", session("pcAdminDGContactPlace")
													objFedExClass.WriteParent "Signatory", "/"
													
													If session("pcAdminDGEmergencyContactNumber")&""<>"" Then
														objFedExClass.AddNewNode "EmergencyContactNumber", session("pcAdminDGEmergencyContactNumber")
													End If
													
													objFedExClass.AddNewNode "Offeror", session("pcAdminDGOfferor")
												objFedExClass.WriteParent "DangerousGoodsDetail", "/"
											END IF

											'// Dry Ice Shipment
											if Session("pcAdminbSODryIce")="1" then
												objFedExClass.WriteParent "DryIceWeight", ""
													objFedExClass.AddNewNode "Units", "KG"
													objFedExClass.AddNewNode "Value", Session("pcAdminSDIValue")
												objFedExClass.WriteParent "DryIceWeight", "/"
											end if

											'//SignatureOption
											if Session("pcAdminSignatureOption")&""<>"" then
												objFedExClass.WriteParent "SignatureOptionDetail", ""
													objFedExClass.AddNewNode "OptionType", Session("pcAdminSignatureOption")
													objFedExClass.AddNewNode "SignatureReleaseNumber", Session("pcAdminSignatureRelease")
												objFedExClass.WriteParent "SignatureOptionDetail", "/"
											end if
													
											if Session("pcAdminPriorityAlert")&""<>"" OR Session("pcAdminPriorityAlertPlus")&""<>"" then
												objFedExClass.WriteParent "PriorityAlertDetail", ""
													if Session("pcAdminPriorityAlertPlus")&""<>"" then
														objFedExClass.AddNewNode "EnhancementTypes", "PRIORITY_ALERT_PLUS"
														objFedExClass.AddNewNode "Content", Session("pcAdminPAPContent")
													else
														objFedExClass.AddNewNode "Content", Session("pcAdminPAContent")
													end if
												objFedExClass.WriteParent "PriorityAlertDetail", "/"
											end if
				
											'// Alcohol
											if Session("pcAdminbSOAlcoholOption")<>"0" AND Session("pcAdminAlcoholRecipientType")&""<>"" then
												objFedExClass.WriteParent "AlcoholDetail", ""
													objFedExClass.AddNewNode "RecipientType", Session("pcAdminAlcoholRecipientType")
												objFedExClass.WriteParent "AlcoholDetail", "/"
											end if

										objFedExClass.WriteParent "SpecialServicesRequested", "/"
									End If
								
								objFedExClass.WriteParent "RequestedPackageLineItems", "/"
							objFedExClass.WriteParent "RequestedShipment", "/"
						objFedExClass.EndXMLTransaction "ProcessShipmentRequest"

						strLogID= Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
						
						'response.Clear()
						'response.contenttype = "text/xml"
						'response.write fedex_postdataWS
						'response.end

						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Send Our Transaction.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						Set srvFEDEXWSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
						Set objOutputXMLDocWS = Server.CreateObject("Microsoft.XMLDOM")
						Set objFedExStream = Server.CreateObject("ADODB.Stream")
						Set objFEDEXXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
						objFEDEXXmlDoc.async = False
						objFEDEXXmlDoc.validateOnParse = False
						if err.number>0 then
							err.clear
						end if

						srvFEDEXWSXmlHttp.open "POST", FedExWSURL&"/ship", false
						srvFEDEXWSXmlHttp.send(fedex_postdataWS)
						FEDEXWS_result = srvFEDEXWSXmlHttp.responseText

						'response.Clear()
						'response.contenttype = "text/xml"
						'response.write FEDEXWS_result
						'response.end

						'// Print out our response

						if trim(FEDEXWS_result)="" then
							call closeDb()
response.redirect ErrPageName & "?msg=FedEx was unable to send a response. There may have been a connection error. Please try again."
						end if

						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Load Our Response.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						call objFedExClass.LoadXMLResults(FEDEXWS_result)
						objOutputXMLDocWS.loadXML FEDEXWS_result

                        'fedex_xmlPrefix = objFedExClass.GetXMLPrefix("")

						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Check for errors from FedEx.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						if pcv_xCounter = 1 then
							'// master package error, no processing done
							pcv_strErrorMsg = objFedExClass.ReadResponseNode("//<VER>ProcessShipmentReply", "<VER>Notifications/<VER>Severity")
							strTmpLabelImage = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>Parts/<VER>Image")

							if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="WARNING" OR (pcv_strErrorMsg="NOTE" AND strTmpLabelImage&""<>"") then
								pcv_strErrorMsg = Cstr("")
							else
								pcv_strErrorMsg = objFedExClass.ReadResponseNode("//<VER>ProcessShipmentReply", "<VER>Notifications/<VER>Message")
							end if

              '// Also try the default namespace in case of a FedEx error
							if pcv_strErrorMsg&""="" then
							  pcv_strErrorMsg = objFedExClass.ReadResponseNode("//ns:ProcessShipmentReply", "ns:Notifications/ns:Severity")
							  strTmpLabelImage = objFedExClass.ReadResponseNode("//ns:CompletedPackageDetails", "ns:Label/ns:Parts/ns:Image")
        
							  if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="WARNING" OR (pcv_strErrorMsg="NOTE" AND strTmpLabelImage&""<>"") then
								  pcv_strErrorMsg = Cstr("")
							  else
								  pcv_strErrorMsg = objFedExClass.ReadResponseNode("//ns:ProcessShipmentReply", "ns:Notifications/ns:Message")
							  end if
							end if

							if pcv_strErrorMsg&""="" then
								pcv_strErrorMsg = objFedExClass.ReadResponseNode("//soapenv:Fault", "faultstring")
								pcv_isFault = "&fault="&sanitizeField(Session("pcAdminRecipPersonName"))&"_Res_"& strLogID &".txt"
							end if
							
							if len(pcv_strErrorMsg)>0 then
								'response.Clear()
								'response.contenttype = "text/xml"
								'response.write FEDEXWS_result
								'response.end
								
								'/////////////////////////////////////////////////////////////
								'// POSTBACK ERROR LOGGING
								'/////////////////////////////////////////////////////////////
								pcv_intRandomNumber = randomNumber(999999999)
								'// Log our Transaction
								call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "ErrLog__"& sanitizeField(Session("pcAdminRecipPersonName")) &"_Req.xml", true)
								'// Log our Error Response
								call objFedExClass.pcs_LogTransaction(FEDEXWS_result, "ErrLog__"& sanitizeField(Session("pcAdminRecipPersonName")) &"_Res.xml", true)
								'/////////////////////////////////////////////////////////////
								'// Display the Error
								call closeDb()
								response.redirect ErrPageName & "?msg=Your shipment was not processed for the following reason. " & pcv_strErrorMsg & pcv_isFault
							else
								pcLocalArray(pcv_xCounter-1) = "shipped"
								pcv_strItemsList = join(pcLocalArray, chr(124))
								Session("pcGlobalArray") = pcv_strItemsList
								'/////////////////////////////////////////////////////////////
								'// POSTBACK LOGGING
								'/////////////////////////////////////////////////////////////
								'// Tracking Number for Logs
								pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>TrackingIds/<VER>TrackingNumber")

								if pcv_strTrackingNumber<>"" then
									'// Log our Transaction
									call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "Req_" & pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".xml", true)
									'// Log our Response
									call objFedExClass.pcs_LogTransaction(FEDEXWS_result, "Res_" & pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".xml", true)
								else
									'// Log our Transaction
									call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "Req_" & pcv_strMethodName&"_noTracking_"&pcv_xCounter&".xml", true)
									'// Log our Response
									call objFedExClass.pcs_LogTransaction(FEDEXWS_result, "Res_" & pcv_strMethodName&"_noTracking_"&pcv_xCounter&".xml", true)
								end if
								'/////////////////////////////////////////////////////////////
							end if
						else
							'// tack package errors, same checks with no redirect
							pcv_strErrorMsg = ""
							call objFedExClass.XMLResponseVerifyCustom(ErrPageName)
							if len(pcv_strErrorMsg)>0 then
								'/////////////////////////////////////////////////////////////
								'// POSTBACK ERROR LOGGING
								'/////////////////////////////////////////////////////////////
								pcv_intRandomNumber = randomNumber(999999999)
								'// Log our Transaction
								call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "ErrLog__"& pcv_intRandomNumber &".xml", true)
								'// Log our Error Response
								call objFedExClass.pcs_LogTransaction(FEDEX_result, "ErrLog__"& pcv_intRandomNumber &".xml", true)
								'/////////////////////////////////////////////////////////////
								'// Pend an error string
								errnum = errnum + 1
								pcv_strSecondaryErrors = pcv_strSecondaryErrors & "<br />" & errnum & ".) " & pcv_strErrorMsg & "<br /> "
							else
								pcLocalArray(pcv_xCounter-1) = "shipped"
								pcv_strItemsList = join(pcLocalArray, chr(124))
								Session("pcGlobalArray") = pcv_strItemsList
								'/////////////////////////////////////////////////////////////
								'// POSTBACK LOGGING
								'/////////////////////////////////////////////////////////////
								'// Tracking Number for Logs
								pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>TrackingIds/<VER>TrackingNumber")
								if pcv_strTrackingNumber<>"" then
									'// Log our Transaction
									call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "Req_" & pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".xml", true)
									'// Log our Response
									call objFedExClass.pcs_LogTransaction(FEDEX_result, "Req_" & pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".xml", true)
								else
									'// Log our Transaction
									call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "Req_" & pcv_strMethodName&"_noTracking_"&pcv_xCounter&".xml", true)
									'// Log our Response
									call objFedExClass.pcs_LogTransaction(FEDEX_result, "Req_" & pcv_strMethodName&"_noTracking_"&pcv_xCounter&".xml", true)
								end if
								'/////////////////////////////////////////////////////////////
							end if
						end if

						'if len(pcv_strErrorMsg) < 1 then
						'	call objFedExClass.pcs_LogTransaction(fedex_postdataWS, sanitizeField(Session("pcAdminRecipPersonName")) & "_" & pcv_xCounter & "_Req" & ".xml", true)
						'	call objFedExClass.pcs_LogTransaction(FEDEXWS_result, sanitizeField(Session("pcAdminRecipPersonName")) & "_" & pcv_xCounter & "_Res" & ".xml", true)
						'end if

						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Redirect with a Message OR complete some task.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						if NOT len(pcv_strErrorMsg)>0 then
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' Set Our Response Data to Local.
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'//Notifications
							pcv_NotificationSeverity = objFedExClass.ReadResponseNode("//<VER>Notifications", "<VER>Severity")
							pcv_NotificationSource = objFedExClass.ReadResponseNode("//<VER>Notifications", "<VER>Source")
							pcv_NotificationCode = objFedExClass.ReadResponseNode("//<VER>Notifications", "<VER>Code")
							pcv_NotificationMessage = objFedExClass.ReadResponseNode("//<VER>Notifications", "<VER>Message")
							pcv_NotificationLocalizedMessage = objFedExClass.ReadResponseNode("//<VER>Notifications", "<VER>LocalizedMessage")

							pcv_CustomerTransactionId = objFedExClass.ReadResponseNode("//<VER>TransactionDetail", "<VER>CustomerTransactionId")

							pcv_VersionServiceId = objFedExClass.ReadResponseNode("//<VER>Version", "<VER>ServiceId")
							pcv_VersionMajor = objFedExClass.ReadResponseNode("//<VER>Version", "<VER>Major")
							pcv_VersionIntermediate = objFedExClass.ReadResponseNode("//<VER>Version", "<VER>Intermediate")
							pcv_VersionMinor = objFedExClass.ReadResponseNode("//<VER>Version", "<VER>Minor")

							pcv_UsDomestic = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>UsDomestic")
							pcv_CarrierCode = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>CarrierCode")
							
							'//if multi-piece shipment get master tracking id
							session("MasterTrackingIdType") = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>MasterTrackingId/<VER>TrackingIdType")
							session("MasterFormId") = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>MasterTrackingId/<VER>FormId")
							session("MasterTrackingNumber") = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>MasterTrackingId/<VER>TrackingNumber")

							pcv_ServiceTypeDescription = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>ServiceTypeDescription")
							pcv_PackagingDescription = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>PackagingDescription")

							pcv_ShipmentOriginLocationNumber = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>OperationalDetail/<VER>OriginLocationNumber")
							pcv_ShipmentDestinationLocationNumber = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>OperationalDetail/<VER>DestinationLocationNumber")
							pcv_ShipmentTransitTime = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>OperationalDetail/<VER>TransitTime")
							pcv_ShipmentCustomTransitTime = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>OperationalDetail/<VER>CustomTransitTime")
							pcv_ShipmentIneligibleForMoneyBackGuarantee = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>OperationalDetail/<VER>IneligibleForMoneyBackGuarantee")
							pcv_ShipmentDeliveryEligibilities = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>OperationalDetail/<VER>DeliveryEligibilities")
							pcv_ShipmentServiceCode = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>OperationalDetail/<VER>ServiceCode")

							pcv_ShipmentRateType = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>RateType")
							pcv_ShipmentRateZone = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>RateZone")
							pcv_ShipmentRatedWeightMethod = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>RatedWeightMethod")
							pcv_ShipmentDimDivisor = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>DimDivisor")
							pcv_ShipmentFuelSurchargePercent = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>FuelSurchargePercent")
							pcv_ShipmentTotalBillingWeightUnits = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalBillingWeight/<VER>Units")
							pcv_ShipmentTotalBillingWeightValue = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalBillingWeight/<VER>Value")

							pcv_ShipmentTotalBaseChargeCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalBaseCharge/<VER>Currency")
							pcv_ShipmentTotalBaseChargeAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalBaseCharge/<VER>Amount")
							pcv_ShipmentTotalFreightDiscountsCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalFreightDiscounts/<VER>Currency")
							pcv_ShipmentTotalFreightDiscountsAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalFreightDiscounts/<VER>Amount")
							pcv_ShipmentTotalNetFreightCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalNetFreight/<VER>Currency")
							pcv_ShipmentTotalNetFreightAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalNetFreight/<VER>Amount")
							pcv_ShipmentTotalSurchargesCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalSurcharges/<VER>Currency")
							pcv_ShipmentTotalSurchargesAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalSurcharges/<VER>Amount")
							pcv_ShipmentTotalNetFedExChargeCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalNetFedExCharge/<VER>Currency")
							pcv_ShipmentTotalNetFedExChargeAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalNetFedExCharge/<VER>Amount")
							pcv_ShipmentTotalTaxesCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalTaxes/<VER>Currency")
							pcv_ShipmentTotalTaxesAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalTaxes/<VER>Amount")
							pcv_ShipmentTotalNetChargeCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalNetCharge/<VER>Currency")
							pcv_ShipmentTotalNetChargeAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalNetCharge/<VER>Amount")
							pcv_ShipmentTotalRebatesCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalRebates/<VER>Currency")
							pcv_ShipmentTotalRebatesAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>TotalRebates/<VER>Amount")

							pcv_ShipmentSurchargesType = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>Surcharges/<VER>SurchargeType")
							pcv_ShipmentSurchargesLevel = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>Surcharges/<VER>Level")
							pcv_ShipmentSurchargesDesc = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>Surcharges/<VER>Description")
							pcv_ShipmentSurchargesCurrency = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>Surcharges/<VER>Amount/<VER>Currency")
							pcv_ShipmentSurchargesAmount = objFedExClass.ReadResponseNode("//<VER>ShipmentRateDetails", "<VER>Surcharges/<VER>Amount/<VER>Amount")

							pcv_SequenceNumber = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>SequenceNumber")

							pcv_TrackingIdType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>TrackingIds/<VER>TrackingIdType")
							pcv_TrackingNumber = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>TrackingIds/<VER>TrackingNumber")

							pcv_GroupNumber = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>GroupNumber")

							pcv_PackageRateType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>RateType")
							pcv_PackageRatedWeightMethod = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>RatedWeightMethod")
							pcv_PackageBillingWeightUnit = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>BillingWeight/<VER>Units")
							pcv_PackageBillingWeightValue = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>BillingWeight/<VER>Value")
							pcv_PackageBaseChargeCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>BaseCharge/<VER>Currency")
							pcv_PackageBaseChargeAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>BaseCharge/<VER>Amount")
							pcv_PackageTotalFreightDiscountsCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalFreightDiscounts/<VER>Currency")
							pcv_PackageTotalFreightDiscountsAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalFreightDiscounts/<VER>Amount")
							pcv_PackageNetFreightCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>NetFreight/<VER>Currency")
							pcv_PackageNetFreightAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>NetFreight/<VER>Amount")
							pcv_PackageTotalSurchargesCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalSurcharges/<VER>Currency")
							pcv_PackageTotalSurcharges = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalSurcharges/<VER>Amount")
							pcv_PackageNetFedExChargeCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>NetFedExCharge/<VER>Currency")
							pcv_PackageNetFedExChargeAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>NetFedExCharge/<VER>Amount")
							pcv_PackageTotalTaxesCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalTaxes/<VER>Currency")
							pcv_PackageTotalTaxesAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalTaxes/<VER>Amount")
							pcv_PackageNetChargeCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>NetCharge/<VER>Currency")
							pcv_PackageNetChargeAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>NetCharge/<VER>Amount")
							pcv_PackageTotalRebatesCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalRebates/<VER>Currency")
							pcv_PackageTotalRebatesAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>TotalRebates/<VER>Amount")
							pcv_PackageSurchargesType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>Surcharges/<VER>SurchargeType")
							pcv_PackageSurchargesLevel = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>Surcharges/<VER>Level")
							pcv_PackageSurchargesDesc = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>Surcharges/<VER>Description")
							pcv_PackageSurchargesCurrency = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>Surcharges/<VER>Amount/<VER>Currency")
							pcv_PackageSurchargesAmount = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>PackageRating/<VER>PackageRateDetails/<VER>Surcharges/<VER>Amount/<VER>Amount")

							pcv_BarcodesType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>OperationalDetail/<VER>Barcodes/<VER>BinaryBarcodes/<VER>Type")
							pcv_BarcodesValue = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>OperationalDetail/<VER>Barcodes/<VER>BinaryBarcodes/<VER>Value")
							pcv_StringBarcodesType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>OperationalDetail/<VER>Barcodes/<VER>StringBarcodes/<VER>Type")
							pcv_StringBarcodesValue = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>OperationalDetail/<VER>Barcodes/<VER>StringBarcodes/<VER>Value")
							pcv_GroundServiceCode = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>OperationalDetail/<VER>GroundServiceCode")
        
							If Trim(pcv_CarrierCode)="FXFR" Then 
							
								pcv_TrackingIdType = session("MasterTrackingIdType")
								pcv_TrackingNumber = session("MasterTrackingNumber")
								pcv_strTrackingNumber = pcv_TrackingNumber
	
							End If

							pcv_SignatureOption = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>SignatureOption")
        
              '// Shipping Label Array Format: (Name, Label Type, Image Type, Resolution, Image Data, Label File)
              pcv_ShippingLabelCount = 0

              '// Add Outbound label
							pcv_LabelType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>Type")
							pcv_LabelImageType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>ImageType")
							pcv_LabelShippingDocumentDisposition = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>ShippingDocumentDisposition")
							pcv_LabelResolution = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>Resolution")
							pcv_LabelCopiesToPrint = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>CopiesToPrint")
							pcv_LabelDocumentPartSequenceNumber = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>Parts/<VER>DocumentPartSequenceNumber")
							pcv_LabelImage = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>Label/<VER>Parts/<VER>Image")
              pcv_LabelFileName = "Label" & pcv_TrackingNumber
              If pcv_LabelImage <> "" Then
                ReDim Preserve pcv_ShippingLabels(pcv_ShippingLabelCount)
                pcv_ShippingLabels(pcv_ShippingLabelCount) = Array("Outbound Label", pcv_LabelType, pcv_LabelImageType, pcv_LabelResolution, pcv_LabelImage, pcv_LabelFileName)
                pcv_ShippingLabelCount = pcv_ShippingLabelCount + 1
              End If

              '// Add Shipment Documents
							pcv_strShipmentDocuments = objFedExClass.ReadResponsesArray("//<VER>CompletedShipmentDetail/<VER>ShipmentDocuments", "")
							Dim pcv_ShippingDocuments
							pcv_ShippingDocuments = Split(pcv_strShipmentDocuments, ",")
							pcv_intAdditonalDocuments = 0
							For i = 0 To UBound(pcv_ShippingDocuments)
								pcv_DocumentType = objFedExClass.ReadResponseNodeIdx("//<VER>CompletedShipmentDetail/<VER>ShipmentDocuments", "<VER>Type", i)
								pcv_DocumentImageType = objFedExClass.ReadResponseNodeIdx("//<VER>CompletedShipmentDetail/<VER>ShipmentDocuments", "<VER>ImageType", i)
								pcv_DocumentImage = objFedExClass.ReadResponseNodeIdx("//<VER>CompletedShipmentDetail/<VER>ShipmentDocuments", "<VER>Parts/<VER>Image", i)
                pcv_DocumentFileName = "Doc" & pcv_TrackingNumber & "_" & pcv_intAdditonalDocuments
								
                If pcv_DocumentImage <> "" Then
                  ReDim Preserve pcv_ShippingLabels(pcv_ShippingLabelCount)
                  pcv_ShippingLabels(pcv_ShippingLabelCount) = Array("Shipping Document", pcv_LabelType, pcv_DocumentImageType, NULL, pcv_DocumentImage, pcv_DocumentFileName)
                  pcv_ShippingLabelCount = pcv_ShippingLabelCount + 1

                  pcv_intAdditonalDocuments = pcv_intAdditonalDocuments + 1
                End If
							Next

              ' Add COD Return Label
							pcv_CodLabelType = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>CodReturnDetail/<VER>Label/<VER>Type")
							pcv_CodLabelShippingDocumentDisposition = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>CodReturnDetail/<VER>Label/<VER>ShippingDocumentDisposition")
							pcv_CodLabelResolution = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>CodReturnDetail/<VER>Label/<VER>Resolution")
							pcv_CodLabelCopiesToPrint = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>CodReturnDetail/<VER>Label/<VER>CopiesToPrint")
							pcv_CodLabelDocumentPartSequenceNumber = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>CodReturnDetail/<VER>Label/<VER>Parts/<VER>DocumentPartSequenceNumber")
							pcv_CodLabelImage = objFedExClass.ReadResponseNode("//<VER>CompletedPackageDetails", "<VER>CodReturnDetail/<VER>Label/<VER>Parts/<VER>Image")
              pcv_CodLabelFileName = "COD" & pcv_TrackingNumber
              If pcv_CodLabelImage <> "" Then
                ReDim Preserve pcv_ShippingLabels(pcv_ShippingLabelCount)
                pcv_ShippingLabels(pcv_ShippingLabelCount) = Array("COD Return Label", pcv_CodLabelType, "PNG", pcv_CodLabelResolution, pcv_CodLabelImage, pcv_CodLabelFileName)
                pcv_ShippingLabelCount = pcv_ShippingLabelCount + 1
              End If

              ' Add associated shipment labels
              pcv_AssociatedShipmentType = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Type")
              Select Case pcv_AssociatedShipmentType
              Case "COD_RETURN":
							  pcv_CodLabelType = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Label/<VER>Type")
							  pcv_CodLabelImageType = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Label/<VER>ImageType")
							  pcv_CodLabelShippingDocumentDisposition = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Label/<VER>ShippingDocumentDisposition")
							  pcv_CodLabelResolution = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Label/<VER>Resolution")
							  pcv_CodLabelCopiesToPrint = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Label/<VER>CopiesToPrint")
							  pcv_CodLabelDocumentPartSequenceNumber = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Label/<VER>Parts/<VER>DocumentPartSequenceNumber")
							  pcv_CodLabelImage = objFedExClass.ReadResponseNode("//<VER>CompletedShipmentDetail", "<VER>AssociatedShipments/<VER>Label/<VER>Parts/<VER>Image")
                pcv_CodLabelFileName = "COD" & pcv_TrackingNumber
                If pcv_CodLabelImage <> "" Then
                  ReDim Preserve pcv_ShippingLabels(pcv_ShippingLabelCount)
                  pcv_ShippingLabels(pcv_ShippingLabelCount) = Array("COD Return Label", pcv_CodLabelType, pcv_CodLabelImageType, pcv_CodLabelResolution, pcv_CodLabelImage, pcv_CodLabelFileName)
                  pcv_ShippingLabelCount = pcv_ShippingLabelCount + 1
                End If
              End Select

							'pcv_ShippingLabelCount = 0

							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' START: SAVE SHIPPING LABELS
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							For i = 0 To pcv_ShippingLabelCount - 1
                labelName = pcv_ShippingLabels(i)(0)
                labelType = pcv_ShippingLabels(i)(1)
                labelImageType = pcv_ShippingLabels(i)(2)
                labelResolution = pcv_ShippingLabels(i)(3)
                labelImage = pcv_ShippingLabels(i)(4)
                labelFileName = pcv_ShippingLabels(i)(5)

                fileName = labelFileName & "." & labelImageType

                Response.Write "Adding Label: " & fileName & "<br/>"
									
								'// Create XML for Label
								GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""" & fileName & """>"&labelImage&"</Base64Data>"
									
								'// Load label from the request stream
								objFEDEXXmlDoc.loadXML GraphicXML
	
								'// Use ADO stream to save the binary data
								objFedExStream.Type = 1
								objFedExStream.Open
	
								objFedExStream.Write objFEDEXXmlDoc.selectSingleNode("/Base64Data").nodeTypedValue
									err.clear
								strFileName = objFEDEXXmlDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue
								'Save the binary stream to the file and overwrite if it already exists in folder
								objFedExStream.SaveToFile server.MapPath("FedExLabels\"&strFileName),2
								objFedExStream.Close()
							Next
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' END: SAVE SHIPPING LABELS
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' START: SAVE PACKAGES
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

							'// Get Our Required Data
							pcv_method=Session("pcAdminService"&pcv_xCounter)
							pcv_tracking=pcv_strTrackingNumber
							pcv_shippedDate=Session("pcAdminShipDate")
							pcv_AdmComments=""

							'// Fix quotes on comments
							if pcv_AdmComments<>"" then
								pcv_AdmComments=replace(pcv_AdmComments,"'","''")
							end if

							dim dtShippedDate
							dtShippedDate=pcv_shippedDate
							if pcv_shippedDate<>"" then
								'dtShippedDate=objFedExClass.pcf_FedExDateFormat(dtShippedDate)
								if SQL_Format="1" then
									dtShippedDate=(day(dtShippedDate)&"/"&month(dtShippedDate)&"/"&year(dtShippedDate))
								else
									dtShippedDate=(month(dtShippedDate)&"/"&day(dtShippedDate)&"/"&year(dtShippedDate))
								end if
							end if

							'// Insert Details into Package Info
							if pcv_shippedDate<>"" then
								query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments,pcPackageInfo_MethodFlag) "
								query=query&"VALUES (" & pcv_intOrderID & ",'" & pcv_method & "','" & dtShippedDate & "','" & pcv_tracking & "','" & pcv_AdmComments & "', 3);"
							else
								query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments,pcPackageInfo_MethodFlag) "
								query=query&"VALUES (" & pcv_intOrderID & ",'" & pcv_method & "','" & pcv_tracking & "','" & pcv_AdmComments & "', 3);"
							end if
							set rs=connTemp.execute(query)
							set rs=nothing

							'// Re-Query for the ID
							query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_intOrderID & " ORDER by pcPackageInfo_ID DESC;"
							set rs=connTemp.execute(query)
							pcv_PackageID=rs("pcPackageInfo_ID")
							set rs=nothing
							qry_ID=pcv_intOrderID

							'// Do a Full Update
							if pcv_PackageNetChargeAmount = "" then
								pcv_PackageNetChargeAmount = 0
							end if

							If Session("pcAdminResidentialDelivery")="true" Then
								pcv_ResidentialDelivery = 1
							End If
							If Session("pcAdminResidentialDelivery")="false" Then
								pcv_ResidentialDelivery  =0
							End if

							query=		"UPDATE pcPackageInfo "
							query=query&"SET pcPackageInfo_FDXSPODFlag=0, "
							query=query&"pcPackageInfo_PackageNumber=1, "
							query=query&"pcPackageInfo_PackageWeight=" & Session("pcAdminWeight"&pcv_xCounter) & ", "
							query=query&"pcPackageInfo_ShipToName='" & Session("pcAdminRecipPersonName") & "', "
							query=query&"pcPackageInfo_ShipToAddress1='" & Session("pcAdminRecipLine1") & "', "
							query=query&"pcPackageInfo_ShipToAddress2='" & Session("pcAdminRecipLine2") & "', "
							query=query&"pcPackageInfo_ShipToCity='" & Session("pcAdminRecipCity") & "', "
							query=query&"pcPackageInfo_ShipToStateCode='" & Session("pcAdminRecipStateOrProvinceCode") & "', "
							query=query&"pcPackageInfo_ShipToZip='" & Session("pcAdminRecipPostalCode") & "', "
							query=query&"pcPackageInfo_ShipToCountry='" & Session("pcAdminRecipCountryCode") & "', "
							query=query&"pcPackageInfo_ShipToPhone='" & Session("pcAdminRecipPhoneNumber") & "', "
							query=query&"pcPackageInfo_ShipToEmail='" & Session("pcAdminRecipEmailAddress") & "', "
							query=query&"pcPackageInfo_ShipToResidential=" & pcv_ResidentialDelivery & ", "
							query=query&"pcPackageInfo_PackageDescription='" & pcv_strPackagingDescription & "', "
							query=query&"pcPackageInfo_ShipFromCompanyName='" & Session("pcAdminOriginCompanyName") & "', "
							query=query&"pcPackageInfo_ShipFromAttentionName='" & Session("pcAdminOriginPersonName") & "', "
							query=query&"pcPackageInfo_ShipFromPhoneNumber='" & Session("pcAdminOriginPhoneNumber") & "', "
							query=query&"pcPackageInfo_ShipFromAddress1='" & Session("pcAdminOriginLine1") & "', "
							query=query&"pcPackageInfo_ShipFromAddress2='" & Session("pcAdminOriginLine2") & "', "
							query=query&"pcPackageInfo_ShipFromCity='" & Session("pcAdminOriginCity") & "', "
							query=query&"pcPackageInfo_ShipFromStateProvinceCode='" & Session("pcAdminOriginStateOrProvinceCode") & "', "
							query=query&"pcPackageInfo_ShipFromPostalCode='" & Session("pcAdminOriginPostalCode") & "', "
							query=query&"pcPackageInfo_ShipFromCountryCode='" & Session("pcAdminOriginCountryCode") & "', "
							query=query&"pcPackageInfo_UPSServiceCode='" & pcv_strServiceTypeDescription & "', "
							query=query&"pcPackageInfo_UPSPackageType='" & pcv_strPackagingDescription & "', "
							query=query&"pcPackageInfo_PackageInsuredValue='" & pcv_strDeclaredValue & "', "
							query=query&"pcPackageInfo_PackageLength='" & Session("pcAdminLength"&pcv_xCounter) & "', "
							query=query&"pcPackageInfo_PackageWidth='" & Session("pcAdminWidth"&pcv_xCounter) & "', "
							query=query&"pcPackageInfo_PackageHeight='" & Session("pcAdminHeight"&pcv_xCounter) & "', "
							If pcv_strServiceTypeDescription="" Then
								pcv_strServiceTypeDescription = Session("pcAdminService"&pcv_xCounter)

							End If
							query=query&"pcPackageInfo_ShipMethod='" & "FedEx: " & pcv_strServiceTypeDescription & "', "
							query=query&"pcPackageInfo_FDXCarrierCode='" & Session("pcAdminCarrierCode") & "', "
							query=query&"pcPackageInfo_FDXRate=" & pcv_PackageNetChargeAmount
							query=query&"WHERE pcPackageInfo_ID=" & pcv_PackageID & " ;"
							set rstemp=connTemp.execute(query)
							set rs=nothing

              ' Create shipping label records
							For i = 0 To pcv_ShippingLabelCount - 1
                labelName = pcv_ShippingLabels(i)(0)
                labelType = pcv_ShippingLabels(i)(1)
                labelImageType = pcv_ShippingLabels(i)(2)
                labelResolution = pcv_ShippingLabels(i)(3)
                labelImage = pcv_ShippingLabels(i)(4)
                labelFileName = pcv_ShippingLabels(i)(5)

								labelName = Replace(labelName, "'", "''")

								If labelResolution & "" = "" Then
									labelResolution = "NULL"
								End If

                query = "SELECT * FROM pcPackageLabel WHERE pcPackageInfo_ID = " & pcv_PackageID & " AND pcPackageLabel_Name = '" & labelName & "' AND pcPackageLabel_File = '" & labelFileName & "';"
                set rstemp = connTemp.execute(query)
                If rstemp.eof Then
                  query = "INSERT INTO pcPackageLabel (pcPackageInfo_ID, pcPackageLabel_Name, pcPackageLabel_File, pcPackageLabel_FileType, pcPackageLabel_Resolution, pcPackageLabel_Type, pcPackageLabel_Date)"
                  query = query & " VALUES (" & pcv_PackageID & ", '" & labelName & "', '" & labelFileName & "', '" & labelImageType & "', " & labelResolution & ", '" & labelType & "','" & Now() & "');"
							    connTemp.execute(query)
                End If
                Set rstemp = Nothing
              Next


							'// Delete the old comments
							query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID & ";"

							set rstemp=connTemp.execute(query)

							'// Add the new comments
							query=		"INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier,pcPackageInfo_ID) "
							query=query&"VALUES (" & qry_ID & ",2,'" & pcv_AdmComments & "',0,0," & pcv_PackageID & ");"
							set rstemp=connTemp.execute(query)

							if trim(Session("pcAdminPrdList"&pcv_xCounter))<>"" then
								pcA=split(Session("pcAdminPrdList"&pcv_xCounter),",")
								For i=lbound(pcA) to ubound(pcA)
									if trim(pcA(i)<>"") then
										query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1, pcPackageInfo_ID=" & pcv_PackageID & " WHERE (idorder=" & qry_ID & " AND idProductOrdered=" & pcA(i) & ");"
										set rs=connTemp.execute(query)
										set rs=nothing
									end if
								Next
							else
								query="UPDATE ProductsOrdered SET pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND pcPrdOrd_Shipped=0 AND pcDropShipper_ID=0;"
								set rsQ=connTemp.execute(query)
								set rsQ=nothing
							end if

							pcv_SendCust="1"

							pcv_SendAdmin="0"
							pcv_LastShip="0"
							query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder AND ProductsOrdered.pcPrdOrd_Shipped=0) WHERE Orders.idorder=" & qry_ID & " AND Orders.orderstatus<>4;"
							set rs=connTemp.execute(query)
							if not rs.eof then
								pcv_LastShip="0"
							else
								pcv_LastShip="1"
							end if
							set rs=nothing

							if trim(Session("pcAdminPrdList"&pcv_xCounter))<>"" then
								if pcv_LastShip="1" then
									query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
								else
									query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & qry_ID & ";"
								end if
								set rs=connTemp.execute(query)
								set rs=nothing
							end if
							
							%>
							<!--#include file="../pc/inc_PartShipEmail.asp"-->
							<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' END: SAVE PACKAGES
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						end if	' if NOT len(pcv_strErrorMsg)>0 then
							
						' Determine if there were any errors. If not, redirect with a message.
						if (NOT len(pcv_strErrorMsg)>0) AND ((pcv_xCounter-1)=UBound(pcLocalArray)) then
							'// Destroy the Sessions
							pcs_ClearAllSessions
							Session("pcAdminPackageCount")=""
							Session("pcAdminOrderID")=""
							Session("pcGlobalArray")=""
							For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
								Session("pcAdminPrdList"&(xArrayCount+1))
							Next
							
							'// REDIRECT
							call closeDb()
response.redirect "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg=Your transaction has been completed successfully."
							response.end
						elseif (pcv_xCounter-1)=UBound(pcLocalArray) then

							'// Generate an error report and redisplay the page.
							pcv_strSecondaryErrMsg="Your shipment was processed. <br />However, we have found the following errors:  <br />"
							pcv_strSecondarySolution="<br />You must resolve all errors to avoid delivery problems. "
							pcv_strSecondarySolution=pcv_strSecondarySolution&"First find and correct all form fields with errors. <br />Then click the 'Finish Processing' button. "
							pcv_strSecondarySolution = pcv_strSecondarySolution & "Repeat these steps until all packages are successfully shipped."

							'// ON ERROR - Reverse Address if Return Shipment
							if Session("pcAdminReturnShipmentIndicator")="PRINT_RETURN_LABEL" then
								pcv_a=Session("pcAdminOriginPersonName")
								pcv_b=Session("pcAdminOriginCompanyName")
								pcv_c=Session("pcAdminOriginDepartment")
								pcv_d=Session("pcAdminOriginPhoneNumber")
								pcv_e=Session("pcAdminOriginPagerNumber")
								pcv_f=Session("pcAdminOriginFaxNumber")
								pcv_g=Session("pcAdminOriginEmailAddress")
								pcv_h=Session("pcAdminOriginLine1")
								pcv_i=Session("pcAdminOriginLine2")
								pcv_j=Session("pcAdminOriginCity")
								pcv_k=Session("pcAdminOriginStateOrProvinceCode")
								pcv_l=Session("pcAdminOriginPostalCode")
								pcv_m=Session("pcAdminOriginCountryCode")

								Session("pcAdminOriginPersonName")=Session("pcAdminRecipPersonName")
								Session("pcAdminOriginCompanyName")=Session("pcAdminRecipCompanyName")
								Session("pcAdminOriginDepartment")=Session("pcAdminRecipDepartment")
								Session("pcAdminOriginPhoneNumber")=Session("pcAdminRecipPhoneNumber")
								Session("pcAdminOriginPagerNumber")=Session("pcAdminRecipPagerNumber")
								Session("pcAdminOriginFaxNumber")=Session("pcAdminRecipFaxNumber")
								Session("pcAdminOriginEmailAddress")=Session("pcAdminRecipEmailAddress")
								Session("pcAdminOriginLine1")=Session("pcAdminRecipLine1")
								Session("pcAdminOriginLine2")=Session("pcAdminRecipLine2")
								Session("pcAdminOriginCity")=Session("pcAdminRecipCity")
								Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
								Session("pcAdminOriginPostalCode")=Session("pcAdminRecipPostalCode")
								Session("pcAdminOriginCountryCode")=Session("pcAdminRecipCountryCode")

								Session("pcAdminRecipPersonName")=pcv_a
								Session("pcAdminRecipCompanyName")=pcv_b
								Session("pcAdminRecipDepartment")=pcv_c
								Session("pcAdminRecipPhoneNumber")=pcv_d
								Session("pcAdminRecipPagerNumber")=pcv_e
								Session("pcAdminRecipFaxNumber")=pcv_f
								Session("pcAdminRecipEmailAddress")=pcv_g
								Session("pcAdminRecipLine1")=pcv_h
								Session("pcAdminRecipLine2")=pcv_i
								Session("pcAdminRecipCity")=pcv_j
								Session("pcAdminRecipStateOrProvinceCode")=pcv_k
								Session("pcAdminRecipPostalCode")=pcv_l
								Session("pcAdminRecipCountryCode")=pcv_m
							end if

							call closeDb()
response.redirect ErrPageName & "?msg=" & Server.URLEncode(pcv_strSecondaryErrMsg & pcv_strSecondaryErrors & pcv_strSecondarySolution)

						end if



					End if '// end skip shipped packages


					'// Reverse Address if Return Shipment
					if Session("pcAdminReturnShipmentIndicator")="PRINT_RETURN_LABEL" then
						pcv_a=Session("pcAdminOriginPersonName")
						pcv_b=Session("pcAdminOriginCompanyName")
						pcv_c=Session("pcAdminOriginDepartment")
						pcv_d=Session("pcAdminOriginPhoneNumber")
						pcv_e=Session("pcAdminOriginPagerNumber")
						pcv_f=Session("pcAdminOriginFaxNumber")
						pcv_g=Session("pcAdminOriginEmailAddress")
						pcv_h=Session("pcAdminOriginLine1")
						pcv_i=Session("pcAdminOriginLine2")
						pcv_j=Session("pcAdminOriginCity")
						pcv_k=Session("pcAdminOriginStateOrProvinceCode")
						pcv_l=Session("pcAdminOriginPostalCode")
						pcv_m=Session("pcAdminOriginCountryCode")

						Session("pcAdminOriginPersonName")=Session("pcAdminRecipPersonName")
						Session("pcAdminOriginCompanyName")=Session("pcAdminRecipCompanyName")
						Session("pcAdminOriginDepartment")=Session("pcAdminRecipDepartment")
						Session("pcAdminOriginPhoneNumber")=Session("pcAdminRecipPhoneNumber")
						Session("pcAdminOriginPagerNumber")=Session("pcAdminRecipPagerNumber")
						Session("pcAdminOriginFaxNumber")=Session("pcAdminRecipFaxNumber")
						Session("pcAdminOriginEmailAddress")=Session("pcAdminRecipEmailAddress")
						Session("pcAdminOriginLine1")=Session("pcAdminRecipLine1")
						Session("pcAdminOriginLine2")=Session("pcAdminRecipLine2")
						Session("pcAdminOriginCity")=Session("pcAdminRecipCity")
						Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
						Session("pcAdminOriginPostalCode")=Session("pcAdminRecipPostalCode")
						Session("pcAdminOriginCountryCode")=Session("pcAdminRecipCountryCode")

						Session("pcAdminRecipPersonName")=pcv_a
						Session("pcAdminRecipCompanyName")=pcv_b
						Session("pcAdminRecipDepartment")=pcv_c
						Session("pcAdminRecipPhoneNumber")=pcv_d
						Session("pcAdminRecipPagerNumber")=pcv_e
						Session("pcAdminRecipFaxNumber")=pcv_f
						Session("pcAdminRecipEmailAddress")=pcv_g
						Session("pcAdminRecipLine1")=pcv_h
						Session("pcAdminRecipLine2")=pcv_i
						Session("pcAdminRecipCity")=pcv_j
						Session("pcAdminRecipStateOrProvinceCode")=pcv_k
						Session("pcAdminRecipPostalCode")=pcv_l
						Session("pcAdminRecipCountryCode")=pcv_m
					end if

				Next
				'///////////////////////////////////////////////////////////////////////
				'// END LOOP
				'///////////////////////////////////////////////////////////////////////

			End If ' If pcv_intErr>0 Then

		else
		'*******************************************************************************
		' END: ON POSTBACK
		'*******************************************************************************

		'*******************************************************************************
		' START: LOAD HTML FORM
		'*******************************************************************************

msg=request.querystring("msg")
if msg<>"" then %>
<div class="pcCPmessage">
	<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
</div>
<% end if %>

<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  FORM VALIDATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf

response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

pcs_JavaTextField	"CarrierCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_1"), ""
pcs_JavaTextField	"ShipDate", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_2"), ""
pcs_JavaTextField	"ShipTime", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_3"), ""
pcs_JavaTextField	"OriginPersonName", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_4"), ""
pcs_JavaTextField	"OriginCompanyName", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_5"), ""
pcs_JavaTextField	"OriginPhoneNumber", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_6"), ""
pcs_JavaTextField	"OriginEmailAddress", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_7"), ""
pcs_JavaTextField	"OriginLine1", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_8"), ""
pcs_JavaTextField	"OriginCity", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_9"), ""
pcs_JavaTextField	"OriginPostalCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_10"), ""
pcs_JavaTextField	"OriginCountryCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_11"), ""
pcs_JavaTextField	"RecipPersonName", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_12"), ""
pcs_JavaTextField	"RecipPhoneNumber", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_13"), ""
pcs_JavaTextField	"RecipCountryCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_14"), ""
pcs_JavaTextField	"RecipLine1", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_15"), ""
pcs_JavaTextField	"RecipCity", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_16"), ""
pcs_JavaTextField	"CustomerReference", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_17"), ""

response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf

response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  FORM VALIDATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
			<form name="form1" method="post" action="<%=pcPageName%>" onSubmit="return Form1_Validator(this)" class="pcForms">
				<table class="pcCPcontent">
					<tr>
				<%
				dim strJSOnChangeTabCnt, k, pcPackageCount, intTempJSChangeCnt
				'pcPackageCount = 1
				strTabCnt=""
				for k=1 to pcPackageCount
					if k=1 then
						strTabCnt="""tab5"""
					else
						iCnt=4+int(k)
						strTabCnt=strTabCnt&",""tab"&iCnt&""""
					end if
				next

				strJSOnChangeTabCnt=""
				for k=1 to pcPackageCount
					intTempJSChangeCnt=4+int(k)
					strJSOnChangeTabCnt=strJSOnChangeTabCnt&";change('tabs"&intTempJSChangeCnt&"', '')"
				next %>
				<!--#include file="../includes/javascripts/pcFedExLabelTabs.asp"-->
				<td valign="top">
					<div class="menu">
						<ul>
							<li><a id="tabs1" class="current" onclick="change('tabs1', 'current');change('tabs2', '');change('tabs3', '');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab1')">Ship Settings</a></li>
							<li><a id="tabs2" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab2')">Ship From</a></li>
							<li><a id="tabs3" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', 'current');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab3')">Recipient</a></li>
							<li><a id="tabs4" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current')<%=strJSOnChangeTabCnt%>;showTab('tab4')">Ship Notification</a></li>
							<% strOnclickTabCnt=""
							if pcPackageCount=1 then %>
							<li><a id="tabs5" onclick="setpackagedivs();change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', 'current'); showTab('tab5')">Package Information</a></li>
							<% else %>
								<% for k=1 to pcPackageCount
									intTempPackageCnt=4+int(k)
									strOnclickTabCnt=""
									for l=1 to pcPackageCount
										intCPC=4+int(l)
										if intCPC=intTempPackageCnt then
											strOnclickTabCnt=strOnclickTabCnt&";change('tabs"&intCPC&"', 'current')"
										else
											strOnclickTabCnt=strOnclickTabCnt&";change('tabs"&intCPC&"', '')"
										end if
									next
									%>
									<li><a id="tabs<%=4+int(k)%>" onclick="setpackagedivs();change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '')<%=strOnclickTabCnt%>;showTab('tab<%=intTempPackageCnt%>')">Package <%=k%></a></li>
									<%
								next
							end if %>
						</ul>
					</div>
					<!--
					//////////////////////////////////////////////////////////////////////////////////////////////
					// SHIP SETTINGS
					//////////////////////////////////////////////////////////////////////////////////////////////
					-->
				  <div id="tab1" class="panes" style="display:block">
		<%
				For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
				%>
				<input type="hidden" name="<%="pcAdminPrdList"&(xArrayCount+1)%>" value="<%=pcLocalArray(xArrayCount)%>">
				<% Next %>
				<input type="hidden" name="idorder" value="<%=pcv_intOrderID%>">
				<input type="hidden" name="PackageCount" value="<%=pcPackageCount%>">
				<input type="hidden" name="ItemsList" value="<%=pcv_strItemsList%>">

				<input name="CurrencyCode" type="hidden" id="CurrencyCode" value="USD" size="3" maxlength="3">
				<table class="pcCPcontent">
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td class="pcCPshipping"><span class="titleShip">Ship Settings</span></td>
						<td class="pcCPshipping" align="right">
						<i>(Check box to view)</i>&nbsp;
						<script type=text/javascript>
						function jfShip(){

						var selectValDom = document.forms['form1'];
						if (selectValDom.bShip.checked == true) {
						document.getElementById('Ship').style.display='';
						}else{
						document.getElementById('Ship').style.display='none';
						}
						}
						</script>
						<%
						if Session("pcAdminbShip")="true" then
							pcv_strDisplayStyle="style=""display:block"""
						else
							pcv_strDisplayStyle="style=""display:none"""
						end if
						%>
			<input onClick="jfShip();" name="bShip" id="bShip" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("bShip", "true")%>>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<div id="Ship" <%=pcv_strDisplayStyle%>>
							<script type=text/javascript>
							document.getElementById('bShip').checked=true
							jfShip();
							</script>
							<table width="100%">
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Billing Detail</th>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Payor:</b></td>
									<td align="left">
										<select name="PayorType" id="PayorType">
										<option value="SENDER" <%=pcf_SelectOption("PayorType","SENDER")%>>Sender</option>
										<option value="RECIPIENT" <%=pcf_SelectOption("PayorType","RECIPIENT")%>>Recipient</option>
										<option value="THIRD_PARTY" <%=pcf_SelectOption("PayorType","THIRD_PARTY")%>>3rd Party</option>
										<option value="COLLECT" <%=pcf_SelectOption("PayorType","COLLECT")%>>Collect</option>
										</select>
										<%pcs_RequiredImageTag "PayorType", true%></td>
								</tr>
								<tr>
								<td align="right" valign="top"><b>Payor Account Number:</b></td>
								<td align="left">
								<input name="PayorAccountNumber" type="text" id="PayorAccountNumber" value="<%=pcf_FillFormField("PayorAccountNumber", false)%>"><%pcs_RequiredImageTag "PayorAccountNumber", false%>
								  </td>
								</tr>
								<tr>
								<td align="right" valign="top"><b>Payor Person Name:</b></td>
								<td align="left">
								<input name="PayorPersonName" type="text" id="PayorPersonName" value="<%=pcf_FillFormField("PayorPersonName", false)%>"><%pcs_RequiredImageTag "PayorPersonName", false%>
								  </td>
								</tr>
								<tr>
								<td align="right" valign="top"><b>Payor Country Code:</b></td>
								<td align="left">
								<input name="PayorCountryCode" type="text" id="PayorCountryCode" value="<%=pcf_FillFormField("PayorCountryCode", false)%>">
								<%pcs_RequiredImageTag "PayorCountryCode", false%>
								e.g.	US	</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Rate Settings</th>
								</tr>
								<tr>
									<td colspan="2">
									<div class="pcCPnotes"> By default account rates are retrieved. Select LIST to retrieve FedEx list rates and
									PREFERRED to retrieve rates with your preferred selected currency.       

									</div>
									</td>
								</tr>
								<tr>
									<td width="24%" align="right" valign="top"><b>Rate Request Type:</b></td>
									<td width="76%" align="left">
									<select name="RateRequestType" id="RateRequestType">
										<option value="ACCOUNT" <%=pcf_SelectOption("RateRequestType","ACCOUNT")%>>Use Default</option>
										<option value="LIST" <%=pcf_SelectOption("RateRequestType","LIST")%>>LIST</option>
										<option value="PREFERRED" <%=pcf_SelectOption("RateRequestType","PREFERRED")%>>PREFERRED</option>
									</select>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>

						<tr>
							<th colspan="2">Service Settings  </th>
						</tr>
						<tr>
									<td width="24%" align="right" valign="top"><b>Type of service:</b></td>
									<td width="76%" align="left">
									<script type=text/javascript>
									function setcarriercodedivs() {
									   var div_num = $pc("#CarrierCode").val();
									   if (div_num == 1) {
									   $pc("#groundsettings").hide();
										 $pc("#CODTitle").html("FedEx&reg; Collect on Delivery (C.O.D.)");
										 $pc("#bSOECOD").hide();
										};
									   if (div_num ==2) {
									   $pc("#groundsettings").show();
										 $pc("#CODTitle").html("FedEx Ground&reg; C.O.D.");
										 $pc("#bSOECOD").show();
										};
									}
									</script>
							<%
							'// Set Carrier Code to local
							Select Case Service
								Case "FedEx First Overnight": ShippingSelector="FDXE"
								Case "FedEx First Overnight Freight": ShippingSelector="FDXE"
								Case "FedEx Priority Overnight": ShippingSelector="FDXE"
								Case "FedEx Standard Overnight": ShippingSelector="FDXE"
								Case "FedEx 2Day": ShippingSelector="FDXE"
								Case "FedEx 2Day AM": ShippingSelector="FDXE"
								Case "FedEx Express Saver": ShippingSelector="FDXE"
								Case "FedEx Freight Priority": ShippingSelector="FXFR"
								Case "FedEx Freight Economy": ShippingSelector="FXFR"
								Case "FedEx Ground": ShippingSelector="FDXG"
								Case "FedEx Home Delivery": ShippingSelector="FDXG"
								Case "FedEx International First": ShippingSelector="FDXE"
								Case "FedEx International Priority": ShippingSelector="FDXE"
								Case "FedEx International Economy": ShippingSelector="FDXE"
								Case "FedEx 1Day Freight": ShippingSelector="FDXE"
								Case "FedEx 2Day Freight": ShippingSelector="FDXE"
								Case "FedEx 3Day Freight": ShippingSelector="FDXE"
								Case "FedEx International Priority Freight": ShippingSelector="FDXE"
								Case "FedEx International Economy Freight": ShippingSelector="FDXE"
							End Select

							if Session("pcAdminCarrierCode")="" then
								Session("pcAdminCarrierCode")=ShippingSelector
							end if
							%>
								<select name="CarrierCode" id="CarrierCode" size="1" onchange="setcarriercodedivs();">
								<option value="1" <%=pcf_SelectOption("CarrierCode","FDXE")%>>FedEx Express</option>
								<option value="2" <%=pcf_SelectOption("CarrierCode","FDXG")%>>FedEx Ground</option>
								<option value="3" <%=pcf_SelectOption("CarrierCode","FXFR")%>>FedEx Freight</option>
							</select>
							<%pcs_RequiredImageTag "CarrierCode", true %>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Drop-off Type:</b></td>
								<td align="left">
								<select name="DropoffType" id="DropoffType">
									<option value="REGULAR_PICKUP" <%=pcf_SelectOption("DropoffType","REGULAR_PICKUP")%>>Regular Pickup</option>
									<option value="REQUEST_COURIER" <%=pcf_SelectOption("DropoffType","REQUEST_COURIER")%>>Courier Pickup</option>
									<option value="DROP_BOX" <%=pcf_SelectOption("DropoffType","DROP_BOX")%>>FedEx Express Drop Box</option>
									<option value="BUSINESS_SERVICE_CENTER" <%=pcf_SelectOption("DropoffType","BUSINESS_SERVICE_CENTER")%>>Business Service Center</option>
									<option value="STATION" <%=pcf_SelectOption("DropoffType","STATION")%>>FedEx Station</option>
								</select><%pcs_RequiredImageTag "DropoffType", isRequiredDropoffType %>
							</td>
						</tr>
						<tr>
              <%
                pcv_strDisplayStyle = "display: none"
                If Session("pcAdminReturnShipmentIndicator") = "PRINT_RETURN_LABEL" Then
                  pcv_strDisplayStyle = ""
                End If
              %>
              <script type=text/javascript>
                function toggleReturnsFields() 
                {
                  if ($pc("ReturnShipmentIndicator").val() == "PRINT_RETURN_LABEL") {
                    $pc("#returnReason").show();
                  } else {
                    $pc("#returnReason").hide();
                  }
                }
              </script>
							<td align="right" valign="top"><b>Shipment Type:</b></td>
							<td align="left">
								<select name="ReturnShipmentIndicator" id="ReturnShipmentIndicator" onchange="toggleReturnsFields();">
								<option value="NON_RETURN" <%=pcf_SelectOption("ReturnShipmentIndicator","NON_RETURN")%>>Outgoing Shipment</option>
								<option value="PRINT_RETURN_LABEL" <%=pcf_SelectOption("ReturnShipmentIndicator","PRINT_RETURN_LABEL")%>>Return Shipment</option>
								</select>
								<%pcs_RequiredImageTag "ReturnShipmentIndicator", false%>
						  </td>
						</tr>
            <tr style="<%= pcv_strDisplayStyle %>" id="returnReason">
              <td align="right"><b>RMA Reason:</b></td>
              <td align="left">
                <INPUT type="text" name="ReturnShipmentReason" id="ReturnShipmentReason" value="<%=pcf_FillFormField("ReturnShipmentReason", false)%>">
                <%pcs_RequiredImageTag "ReturnShipmentReason", false%>
              </td>
            </tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Label Settings</th>
						</tr>
            <tr>
							<td align="right" valign="top"><b>Label Type:</b></td>
							<td align="left">
								<select name="LabelFormatType" id="LabelFormatType">
									<option value="COMMON2D" <%=pcf_SelectOption("LabelFormatType","COMMON2D")%>>Common 2D</option>
									<option value="FEDEX_FREIGHT_STRAIGHT_BILL_OF_LADING" <%=pcf_SelectOption("LabelFormatType","FEDEX_FREIGHT_STRAIGHT_BILL_OF_LADING")%>>FedEx Freight Bill of Lading</option>
								</select>
              </td>
            </tr>
            <tr>
							<td align="right" valign="top"><b>Image Type:</b></td>
							<td align="left">
								<select name="LabelImageType" id="LabelImageType">
									<option value="PNG" <%=pcf_SelectOption("LabelImageType","PNG")%>>PNG</option>
									<option value="PDF" <%=pcf_SelectOption("LabelImageType","PDF")%>>PDF</option>
								</select>
              </td>
            </tr>
            <tr>
							<td align="right" valign="top"><b>Stock Type:</b></td>
							<td align="left">
								<select name="LabelStockType" id="LabelStockType">
									<option value="" <%=pcf_SelectOption("LabelStockType","")%>>Use Default</option>
									<option value="PAPER_4X6" <%=pcf_SelectOption("LabelStockType","PAPER_4X6")%>>Paper 4x6</option>
									<option value="PAPER_4X8" <%=pcf_SelectOption("LabelStockType","PAPER_4X8")%>>Paper 4x8</option>
									<option value="PAPER_4X9" <%=pcf_SelectOption("LabelStockType","PAPER_4X9")%>>Paper 4x9</option>
									<option value="PAPER_7X4.75" <%=pcf_SelectOption("LabelStockType","PAPER_7X4.75")%>>Paper 7x4.75</option>
									<option value="PAPER_8.5X11_BOTTOM_HALF_LABEL" <%=pcf_SelectOption("LabelStockType","PAPER_8.5X11_BOTTOM_HALF_LABEL")%>>Paper 8.5x11 (Bottom Half)</option>
									<option value="PAPER_8.5X11_TOP_HALF_LABEL" <%=pcf_SelectOption("LabelStockType","PAPER_8.5X11_TOP_HALF_LABEL")%>>Paper 8.5x11 (Top Half)</option>
									<option value="PAPER_LETTER" <%=pcf_SelectOption("LabelStockType","PAPER_LETTER")%>>Paper Letter</option>
								</select>
              </td>
            </tr>
            <tr>
							<td align="right" valign="top"><b>Printing Orientation:</b></td>
							<td align="left">
								<select name="LabelPrintingOrientation" id="LabelPrintingOrientation">
									<option value="" <%=pcf_SelectOption("LabelPrintingOrientation","")%>>Use Default</option>
									<option value="BOTTOM_EDGE_OF_TEXT_FIRST" <%=pcf_SelectOption("LabelPrintingOrientation","BOTTOM_EDGE_OF_TEXT_FIRST")%>>Bottom edge of text first</option>
									<option value="TOP_EDGE_OF_TEXT_FIRST" <%=pcf_SelectOption("LabelPrintingOrientation","TOP_EDGE_OF_TEXT_FIRST")%>>Top edge of text first</option>
								</select>
              </td>
            </tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Total Shipment Weight  </th>
						</tr>
						<%
						if scShipFromWeightUnit="KGS" then
							intShipWeightPounds=int(pcOrd_ShipWeight/1000)
							intShipWeightOunces=pcv_ShipWeight-(intShipWeightPounds*1000)
						else
							intShipWeightPounds=int(pcOrd_ShipWeight/16) 'intPounds used for USPS
							intShipWeightOunces=pcOrd_ShipWeight-(intShipWeightPounds*16) 'intUniversalOunces used for USPS
						end if

						intMPackageWeight=intShipWeightPounds
						if intMPackageWeight<1 AND intShipWeightOunces<1 then
							intMPackageWeight=0
							intShipWeightOunces=0
						end if

						if intMPackageWeight<1 AND intShipWeightOunces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
							intMPackageWeight=1
						else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
							If intMPackageWeight>0 AND intShipWeightOunces>0 then
								intMPackageWeight=cLng(intMPackageWeight)+1
							End if
							End if
						if request("test")<>"" then
							response.write intMPackageWeight
							response.End()
						end if

						If session("pcAdminTotalShipmentWeight")&""="" Then
							session("pcAdminTotalShipmentWeight") = intMPackageWeight
						End If
						%>
						<tr>
							<td width="24%" align="right" valign="top"><strong>Total Weight:</strong></td>
							<td width="76%" align="left">
                <% FedEx_WeightControl "TotalShipmentWeight", "ShipmentWeightUnits", true %>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>

						<tr>
							<th colspan="2">Date/ Time  </th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Ship Date:</b></td>
							<td align="left">
								<select name="ShipDate">
								<% dtTodayDate=Date()
								Function FedExDateFormat (FedExDate)
								FedExDay=Day(FedExDate)
								FedExMonth=Month(FedExDate)
								FedExYear= Year(FedExDate)
								FedExDateFormat=FedExYear&"-"&Right(Cstr(FedExMonth + 100),2)&"-"&Right(Cstr(FedExDay + 100),2)
								End Function %>
								<option value="<%=FedExDateFormat(dtTodayDate)%>" <%=pcf_SelectOption("ShipDate",FedExDateFormat(dtTodayDate))%>>Today</option>
								<% for d=1 to 10
								if DatePart("W", dtTodayDate+d, VBSUNDAY)=1 then
								else %>
								<option value="<%=FedExDateFormat((dtTodayDate+d))%>" <%=pcf_SelectOption("ShipDate",FedExDateFormat(dtTodayDate+d))%>><%=FormatDateTime((dtTodayDate+d), 1)%></option>
								<% end if
								next %>
								</select>
								<%pcs_RequiredImageTag "ShipDate", true%>
									&nbsp;&nbsp;&nbsp;<b>Ship Time:&nbsp;
									<input name="ShipTime" type="text" id="ShipTime" value="<%=pcf_FillFormField("ShipTime", true)%>"><%pcs_RequiredImageTag "ShipTime", true%>
									* hh:mm:ss </b></td>
							  </tr>
							</table>
							</div>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
			<tr>
						<td class="pcCPshipping" colspan="2"><span class="titleShip">Additional Settings</span></td>
					</tr>
					<tr>
						<td colspan="2">
							<table width="100%">
								<!--<tr>
									<th align="left"><input type="checkbox" name="Alcohol" value="1" class="clearBorder" <%=pcf_CheckOption("Alcohol", "1")%><%pcs_RequiredImageTag "Alcohol", false%>>&nbsp;Alcohol Services</th>
									<td align="left">

									</td>
								</tr>
								<tr>
									<td colspan="2"></td>
							  </tr>-->
								<tr>
									<th colspan="2">
									<script type=text/javascript>
										function jfSOAlcoholOption(){

											var selectValDom = document.forms['form1'];
											if (selectValDom.bSOAlcoholOption.checked == true) {
												document.getElementById('SOAlcoholOption').style.display='';
											}else{
												document.getElementById('SOAlcoholOption').style.display='none';
											}
										}
									</script>
									<%
									if Session("pcAdminbSOAlcoholOption")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOAlcoholOption();" name="bSOAlcoholOption" id="bSOAlcoholOption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOAlcoholOption", "1")%>>
									Alcohol Shipment</th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOAlcoholOption" <%=pcv_strDisplayStyle%>>
											<table>
                        <tr>
                        <td align="right" valign="top"><b>Recipient Type:</b></td>
                          <td align="left">
                            <select name="AlcoholRecipientType" id="AlcoholRecipientType">
                              <option value="CONSUMER" <%=pcf_SelectOption("AlcoholRecipientType","CONSUMER")%>>Consumer</option>
                              <option value="LICENSEE" <%=pcf_SelectOption("AlcoholRecipientType","LICENSEE")%>>Licensee</option>
                            </select>
                            <%pcs_RequiredImageTag "AlcoholRecipientType", false%>
                          </td>
                        </tr>
											</table>
										</div>
									</td>
								</tr>
								<tr>
									<th align="left"><input type="checkbox" name="PharmacyDelivery" value="1" class="clearBorder" <%=pcf_CheckOption("PharmacyDelivery", "1")%><%pcs_RequiredImageTag "PharmacyDelivery", false%>>&nbsp;Pharmacy Delivery</th>
									<td align="left">

									</td>
								</tr>
								<tr>
									<td colspan="2"></td>
							  </tr>
                <tr>
									<th colspan="2">
									<script type=text/javascript>
										function jfSOShippingDocumentOption(){

											var selectValDom = document.forms['form1'];
											if (selectValDom.bSOShippingDocumentOption.checked == true) {
												document.getElementById('SOShippingDocumentOption').style.display='';
											}else{
												document.getElementById('SOShippingDocumentOption').style.display='none';
											}
										}
									</script>
									<%
									if Session("pcAdminbSOShippingDocumentOption")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOShippingDocumentOption();" name="bSOShippingDocumentOption" id="bSOShippingDocumentOption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOShippingDocumentOption", "1")%>>
									Additional Shipping Documents</th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOShippingDocumentOption" <%=pcv_strDisplayStyle%>>
											<Table>
                        <tr>
                        <td align="right" valign="top"><b>Shipping Document Type:</b></td>
                          <td align="left">
                            <%
                              FedEx_RequestedShippingDocumentType "SDSType"
                            %>
                          </td>
                        </tr>
                        <tbody id="SDSDetail">
													<tr>
														<td align="right" valign="top"><b>Image Type:</b></td>
														<td align="left">
															<select name="SDSLabelImageType" id="SDSLabelImageType">
																<option value="PNG" <%=pcf_SelectOption("SDSLabelImageType","PNG")%>>PNG</option>
																<option value="PDF" <%=pcf_SelectOption("SDSLabelImageType","PDF")%>>PDF</option>
															</select>
														</td>
													</tr>
													<tr>
														<td align="right" valign="top"><b>Stock Type:</b></td>
														<td align="left">
															<select name="SDSLabelStockType" id="SDSLabelStockType">
																<option value="" <%=pcf_SelectOption("SDSLabelStockType","")%>>Use Default</option>
																<option value="PAPER_4X6" <%=pcf_SelectOption("SDSLabelStockType","PAPER_4X6")%>>Paper 4x6</option>
																<option value="PAPER_4X8" <%=pcf_SelectOption("SDSLabelStockType","PAPER_4X8")%>>Paper 4x8</option>
																<option value="PAPER_4X9" <%=pcf_SelectOption("SDSLabelStockType","PAPER_4X9")%>>Paper 4x9</option>
																<option value="PAPER_7X4.75" <%=pcf_SelectOption("SDSLabelStockType","PAPER_7X4.75")%>>Paper 7x4.75</option>
																<option value="PAPER_8.5X11_BOTTOM_HALF_LABEL" <%=pcf_SelectOption("SDSLabelStockType","PAPER_8.5X11_BOTTOM_HALF_LABEL")%>>Paper 8.5x11 (Bottom Half)</option>
																<option value="PAPER_8.5X11_TOP_HALF_LABEL" <%=pcf_SelectOption("SDSLabelStockType","PAPER_8.5X11_TOP_HALF_LABEL")%>>Paper 8.5x11 (Top Half)</option>
																<option value="PAPER_LETTER" <%=pcf_SelectOption("SDSLabelStockType","PAPER_LETTER")%>>Paper Letter</option>
															</select>
														</td>
													</tr>
													<tr>
														<td align="right"><b>Provide Instructions:</b></td>
														<td align="left">
															<INPUT type="radio" name="SDSProvideInstructions" value="true" class="clearBorder" <%= pcf_CheckOption("SDSProvideInstructions", "true") %>>Yes
															<INPUT type="radio" name="SDSProvideInstructions" value="false" class="clearBorder" <%= pcf_CheckOption("SDSProvideInstructions", "false") %>>No
														</td>
													</tr>
                        </tbody>
											</Table>
										</div>
									</td>
								</tr>

                <tr>
									<th colspan="2">
									<script type=text/javascript>
                  function jfSOETDOption(){

                    var selectValDom = document.forms['form1'];
                    if (selectValDom.bSOETDOption.checked == true) {
                      document.getElementById('SOETDOption').style.display='';
                    }else{
                      document.getElementById('SOETDOption').style.display='none';
                    }
                  }
									</script>
									<%
									if Session("pcAdminbSOETDOption")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOETDOption();" name="bSOETDOption" id="bSOETDOption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOETDOption", "1")%>>
									Electronic Trade Documents (ETD)</th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOETDOption" <%=pcv_strDisplayStyle%>>
											<Table>
                        <tr>
                        <td align="right" valign="top"><b>Requested Document Copies:</b></td>
                          <td align="left">
                            <%
                              FedEx_RequestedShippingDocumentType "ETDRequestedDocumentCopies"
                            %>
                          </td>
                        </tr>
                        <!--<tbody id="ETDDetail">
                        	<tr>
                            <td align="right" valign="top"><b>Line Number:</b></td>
                            <td align="left">
                              <input type="number" min="0" name="ETDLineNumber" id="ETDLineNumber" value="<%=pcf_FillFormField("ETDLineNumber", false)%>"/>
                            </td>
                          </tr>
                        	<tr>
                            <td align="right" valign="top"><b>Document Producer:</b></td>
                            <td align="left">
                              <select name="ETDDocumentProducer" id="ETDDocumentProducer">
                                <option value="CUSTOMER" <%=pcf_SelectOption("ETDDocumentProducer","CUSTOMER")%>>Customer</option>
                                <option value="FEDEX_CLS" <%=pcf_SelectOption("ETDDocumentProducer","FEDEX_CLS")%>>FedEx Common Label Server</option>
                                <option value="FEDEX_GTM" <%=pcf_SelectOption("ETDDocumentProducer","FEDEX_GTM")%>>FedEx Global Trade Manager</option>
                                <option value="OTHER" <%=pcf_SelectOption("ETDDocumentProducer","OTHER")%>>Other</option>
                              </select>
                            </td>
                          </tr>
                        	<tr>
                            <td align="right" valign="top"><b>Document ID:</b></td>
                            <td align="left">
                              <input type="text" name="ETDDocumentId" id="ETDDocumentId" value="<%=pcf_FillFormField("ETDDocumentId", false)%>"/>
                            </td>
                          </tr>
                        	<tr>
                            <td align="right" valign="top"><b>Document ID Producer:</b></td>
                            <td align="left">
                              <select name="ETDDocumentIdProducer" id="ETDDocumentIdProducer">
                                <option value="CUSTOMER" <%=pcf_SelectOption("ETDDocumentIdProducer","CUSTOMER")%>>Customer</option>
                                <option value="FEDEX_CSHP" <%=pcf_SelectOption("ETDDocumentIdProducer","FEDEX_CSHP")%>>FedEx CSHP</option>
                                <option value="FEDEX_GTM" <%=pcf_SelectOption("ETDDocumentIdProducer","FEDEX_GTM")%>>FedEx Global Trade Manager</option>
                              </select>
                            </td>
                          </tr>
                        </tbody>-->
											</Table>
										</div>
									</td>
								</tr>
                
                
								<tr>
									<th colspan="2">
									<script type=text/javascript>
									function jfSOSaturdayServices(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOSaturdayServices.checked == true) {
									document.getElementById('SOSaturdayServices').style.display='';
									}else{
									document.getElementById('SOSaturdayServices').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminbSOSaturdayServices")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOSaturdayServices();" name="bSOSaturdayServices" id="bSOSaturdayServices" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOSaturdayServices", "1")%>>
									Saturday Services </th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOSaturdayServices" <%=pcv_strDisplayStyle%>>
											<Table>
											  <TR>
												<TD width="51" height="20">&nbsp;</TD>
												<TD width="303"><INPUT tabIndex="25" type="checkbox" value="1" name="SaturdayDelivery" class="clearBorder" <%=pcf_CheckOption("SaturdayDelivery", "1")%>>
												&nbsp;Saturday Delivery</TD>
											  </TR>
											  <TR>
												<TD height="20">&nbsp;</TD>
												<TD height="20"><input tabindex="25" type="checkbox" value="1" name="SaturdayPickup" class="clearBorder" <%=pcf_CheckOption("SaturdayPickup", "1")%>>
&nbsp;Saturday Pickup</TD>
											  </TR>
											</Table>
										</div>
									</td>
								</tr>
								<tr>
									<th colspan="2">
									<script type=text/javascript>
									function jfSOSignatureOption(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOSignatureOption.checked == true) {
									document.getElementById('SOSignatureOption').style.display='';
									}else{
									document.getElementById('SOSignatureOption').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminbSOSignatureOption")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOSignatureOption();" name="bSOSignatureOption" id="bSOSignatureOption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOSignatureOption", "1")%>>
									Signature Options </th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOSignatureOption" <%=pcv_strDisplayStyle%>>
											<Table>
                        <tr>
                        <td align="right" valign="top"><b>Signature Type:</b></td>
                          <td align="left">
                            <select name="SignatureOption" id="SignatureOption">
                              <option value="" <%=pcf_SelectOption("SignatureOption","")%>>No Signature Options</option>
                              <option value="SERVICE_DEFAULT" <%=pcf_SelectOption("SignatureOption","SERVICE_DEFAULT")%>>Service default</option>
            
                              <option value="NO_SIGNATURE_REQUIRED" <%=pcf_SelectOption("SignatureOption","NO_SIGNATURE_REQUIRED")%>>No signature required</option>
                              <option value="INDIRECT" <%=pcf_SelectOption("SignatureOption","INDIRECT")%>>Indirect Signature Required</option>
                              <option value="DIRECT" <%=pcf_SelectOption("SignatureOption","DIRECT")%>>Direct Signature Required</option>
                              <option value="ADULT" <%=pcf_SelectOption("SignatureOption","ADULT")%>>Adult Signature Required</option>
                            </select>
                            <%pcs_RequiredImageTag "SignatureOption", false%>
                          </td>
                        </tr>
                        <tr>
                          <td align="right"><b>Signature Release:</b></td>
                          <td align="left">
                            <INPUT type="text" name="SignatureRelease" id="SignatureRelease" value="<%=pcf_FillFormField("SignatureRelease", false)%>">
                            <%pcs_RequiredImageTag "SignatureRelease", false%>
                            (Deliver Without Signature Only)
                          </td>
                        </tr>
											</Table>
										</div>
									</td>
								</tr>
								<tr>
									<th align="left"><input type="checkbox" name="ExtremeLength" value="1" class="clearBorder" <%=pcf_CheckOption("ExtremeLength", "1")%><%pcs_RequiredImageTag "ExtremeLength", false%>>&nbsp;Extreme Length Package</th>
									<td align="left">
									</td>
								</tr>
								<tr>
									<td colspan="2"></td>
							  </tr>
								<tr>
									<th align="left"><input type="checkbox" name="OneRate" value="1" class="clearBorder" <%=pcf_CheckOption("OneRate", "1")%><%pcs_RequiredImageTag "OneRate", false%>>&nbsp;FedEx One Rate<sup>&reg</sup></th>
									<td align="left">
									</td>
								</tr>
								<tr>
									<td colspan="2"></td>
							  </tr>
								<tr>
									<th align="left">
                  
									<script type=text/javascript>
									function jfSOPriorityAlertOption(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.PriorityAlert.checked == true) {
									document.getElementById('PriorityAlertOption').style.display='';
									
									selectValDom.PriorityAlertPlus.checked = false;
									document.getElementById('PriorityAlertPlusOption').style.display='none';
									}else{
									document.getElementById('PriorityAlertOption').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminPriorityAlert")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
                  
                  <input type="checkbox" name="PriorityAlert" value="1" class="clearBorder" <%=pcf_CheckOption("PriorityAlert", "1")%> onClick="return jfSOPriorityAlertOption();">
									<%pcs_RequiredImageTag "PriorityAlert", false%>FedEx Priority Alert&trade;</th>
									<td align="left">

									</td>
								</tr>
								<tr>
									<td colspan="2">
										<div id="PriorityAlertOption" <%=pcv_strDisplayStyle%>>
											<table>
                        <tr>
                          <td align="right" valign="top"><b>Content: </b></td>
                          <td align="left">
                            <INPUT type="text" name="PAContent" id="PAContent" value="<%=pcf_FillFormField("PAContent", false)%>">
                            <%pcs_RequiredImageTag "PAContent", false%>
                          </td>
                        </tr>
                      </Table>
                    </div>
                  </td>
								</tr>
								<tr>
									<th align="left">
									<script type=text/javascript>
									function jfSOPriorityAlertPlusOption(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.PriorityAlertPlus.checked == true) {
									document.getElementById('PriorityAlertPlusOption').style.display='';
									
									selectValDom.PriorityAlert.checked = false;
									document.getElementById('PriorityAlertOption').style.display='none';
									}else{
									document.getElementById('PriorityAlertPlusOption').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminPriorityAlertPlus")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
                  
                  <input type="checkbox" name="PriorityAlertPlus" value="1" class="clearBorder" <%=pcf_CheckOption("PriorityAlertPlus", "1")%> onClick="return jfSOPriorityAlertPlusOption();">
									<%pcs_RequiredImageTag "PriorityAlertPlus", false%>FedEx Priority Alert Plus&trade;</th>
									<td align="left">

									</td>
								</tr>
								<tr>
									<td colspan="2">
										<div id="PriorityAlertPlusOption" <%=pcv_strDisplayStyle%>>
											<table>
                        <tr>
                          <td align="right" valign="top"><b>Content: </b></td>
                          <td align="left">
                            <INPUT type="text" name="PAPContent" id="PAPContent" value="<%=pcf_FillFormField("PAPContent", false)%>">
                            <%pcs_RequiredImageTag "PAPContent", false%>
                          </td>
                        </tr>
                      </Table>
                    </div>
                  </td>
                </tr>
								<tr>
									<th colspan="2">
									<script type=text/javascript>
									function jfSOCODCollection(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOCODCollection.checked == true) {
									document.getElementById('SOCODCollection').style.display='';
									}else{
									document.getElementById('SOCODCollection').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminbSOCODCollection")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOCODCollection();" name="bSOCODCollection" id="bSOCODCollection" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOCODCollection", "1")%>>
									<span id="CODTitle">FedEx&reg; Collect on Delivery (C.O.D.)</span></th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOCODCollection" <%=pcv_strDisplayStyle%>>
											<Table>
								<tr>
							<td align="right"><b>Collection Amount:</b></td>
							<td align="left">
								<%
									FedEx_CurrencyControl "CODAmount", "CODCurrency", false
								%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Rate Type:</b></td>
							<td align="left">
								<select name="CODRateType" id="CODRateType">
									<option value="" <%=pcf_SelectOption("CODRateType","")%>></option>
									<option value="ACCOUNT" <%=pcf_SelectOption("CODRateType","ACCOUNT")%>>ACCOUNT</option>
									<option value="LIST" <%=pcf_SelectOption("CODRateType","LIST")%>>LIST</option>
								</select>
								<%pcs_RequiredImageTag "CODRateType", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Charge Basis:</b></td>
							<td align="left">
								<select name="CODChargeBasis" id="CODChargeBasis">
									<option value="" <%=pcf_SelectOption("CODChargeBasis","")%>></option>
									<option value="COD_SURCHARGE" <%=pcf_SelectOption("CODChargeBasis","COD_SURCHARGE")%>>COD Surcharge</option>
									<option value="NET_CHARGE" <%=pcf_SelectOption("CODChargeBasis","NET_CHARGE")%>>Net Charge</option>
									<option value="NET_FREIGHT" <%=pcf_SelectOption("CODChargeBasis","NET_FREIGHT")%>>Net Freight</option>
									<option value="TOTAL_CUSTOMER_CHARGE" <%=pcf_SelectOption("CODChargeBasis","TOTAL_CUSTOMER_CHARGE")%>>Total Customer Charge</option>
								</select>
								<%pcs_RequiredImageTag "CODChargeBasis", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Charge Basis Level:</b></td>
							<td align="left">
								<select name="CODChargeBasisLevel" id="CODChargeBasisLevel">
									<option value="" <%=pcf_SelectOption("CODChargeBasisLevel","")%>></option>
									<option value="CURRENT_PACKAGE" <%=pcf_SelectOption("CODChargeBasisLevel","CURRENT_PACKAGE")%>>Current Package</option>
									<option value="SUM_OF_PACKAGES" <%=pcf_SelectOption("CODChargeBasisLevel","SUM_OF_PACKAGES")%>>Sum of Packages</option>
								</select>
								<%pcs_RequiredImageTag "CODChargeBasisLevel", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Collection Type:</b></td>
							<td align="left">
								<select name="CODType" id="CODType">
									<option value="CASH" <%=pcf_SelectOption("CODType","CASH")%>>Cash</option>
									<option value="COMPANY_CHECK" <%=pcf_SelectOption("CODType","COMPANY_CHECK")%>>Company Check</option>

									<option value="GUARANTEED_FUNDS" <%=pcf_SelectOption("CODType","GUARANTEED_FUNDS")%>>Guaranteed Funds</option>
									<option value="PERSONAL_CHECK" <%=pcf_SelectOption("CODType","PERSONAL_CHECK")%>>Personal Check</option>
									<option value="ANY" <%=pcf_SelectOption("CODType","ANY")%>>Any</option>
								</select>
								<%pcs_RequiredImageTag "CODType", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Tin Type:</b></td>
							<td align="left">
								<select name="CODTinType" id="CODTinType">
                	<option value="">Select Option</option>
									<option value="BUSINESS_NATIONAL" <%=pcf_SelectOption("CODTinType","BUSINESS_NATIONAL")%>>Business National</option>
									<option value="BUSINESS_STATE" <%=pcf_SelectOption("CODTinType","BUSINESS_STATE")%>>Business State</option>
									<option value="BUSINESS_UNION" <%=pcf_SelectOption("CODTinType","BUSINESS_UNION")%>>Business Union</option>
									<option value="PERSONAL_NATIONAL" <%=pcf_SelectOption("CODTinType","PERSONAL_NATIONAL")%>>Personal National</option>
									<option value="PERSONAL_STATE" <%=pcf_SelectOption("CODTinType","PERSONAL_STATE")%>>Personal State</option>
								</select>
								<%pcs_RequiredImageTag "CODTinType", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Tin Number:</b></td>
							<td align="left">
								<INPUT type="text" name="CODTinNumber" id="CODTinNumber" value="<%=pcf_FillFormField("CODTinNumber", false)%>">
								<%pcs_RequiredImageTag "CODTinNumber", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Account Number:</b></td>
							<td align="left">
								<INPUT type="text" name="CODAccountNumber" id="CODAccountNumber" value="<%=pcf_FillFormField("CODAccountNumber", false)%>">
								<%pcs_RequiredImageTag "CODAccountNumber", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b> Contact Name:</b></td>
							<td align="left">
								<INPUT type="text" name="CODPersonName" id="CODPersonName" value="<%=pcf_FillFormField("CODPersonName", false)%>">
								<%pcs_RequiredImageTag "CODPersonName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Company Name:</b></td>
							<td align="left">
								<INPUT type="text" name="CODCompanyName" id="CODCompanyName" value="<%=pcf_FillFormField("CODCompanyName", false)%>">
								<%pcs_RequiredImageTag "CODCompanyName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Phone Number:</b></td>
							<td align="left">
								<INPUT type="text" name="CODPhoneNumber" id="CODPhoneNumber" value="<%=pcf_FillFormField("CODPhoneNumber", false)%>">
								<%pcs_RequiredImageTag "CODPhoneNumber", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Title:</b></td>
							<td align="left">
								<INPUT type="text" name="CODTitle" id="CODTitle" value="<%=pcf_FillFormField("CODTitle", false)%>">
								<%pcs_RequiredImageTag "CODTitle", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Street Lines:</b></td>
							<td align="left">
								<INPUT type="text" name="CODStreetLines" id="CODStreetLines" value="<%=pcf_FillFormField("CODStreetLines", false)%>">
								<%pcs_RequiredImageTag "CODStreetLines", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b> City:</b></td>
							<td align="left">
								<INPUT type="text" name="CODCity" id="CODCity" value="<%=pcf_FillFormField("CODCity", false)%>">
								<%pcs_RequiredImageTag "CODCity", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>State:</b></td>
							<td align="left">
								<INPUT type="text" name="CODState" id="CODState" value="<%=pcf_FillFormField("CODState", false)%>">
								<%pcs_RequiredImageTag "CODState", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Postal Code:</b></td>
							<td align="left">
								<INPUT type="text" name="CODPostalCode" id="CODPostalCode" value="<%=pcf_FillFormField("CODPostalCode", false)%>">
								<%pcs_RequiredImageTag "CODPostalCode", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Country Code:</b></td>
							<td align="left">
								<INPUT type="text" name="CODCountryCode" id="CODCountryCode" value="<%=pcf_FillFormField("CODCountryCode", false)%>">
								<%pcs_RequiredImageTag "CODCountryCode", false%>
							</td>
						</tr>
											</Table>
										</div>
									</td>
								</tr>

								<tr>
								  <th colspan="2">
									<script type=text/javascript>
									function jfSOHAL(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOHAL.checked == true) {
									document.getElementById('SOHAL').style.display='';
									}else{
									document.getElementById('SOHAL').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminbSOHAL")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOHAL();" name="bSOHAL" id="bSOHAL" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOHAL", "1")%>>
									Hold At FedEx Location </th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOHAL" <%=pcv_strDisplayStyle%>>
											<Table>
						<tr>
							<td align="right"><b>Contact ID:</b></td>
							<td align="left">
								<INPUT type="text" name="HALContactID" id="HALContactID" value="<%=pcf_FillFormField("HALContactID", false)%>">
								<%pcs_RequiredImageTag "HALContactID", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Contact Name:</b></td>
							<td align="left">
								<INPUT type="text" name="HALPersonName" id="HALPersonName" value="<%=pcf_FillFormField("HALPersonName", false)%>">
								<%pcs_RequiredImageTag "HALPersonName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Company Name:</b></td>
							<td align="left">
								<INPUT type="text" name="HALCompanyName" id="HALCompanyName" value="<%=pcf_FillFormField("HALCompanyName", false)%>">
								<%pcs_RequiredImageTag "HALCompanyName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Phone:</b></td>
							<td align="left">
								<INPUT type="text" name="HALPhone" id="HALPhone" value="<%=pcf_FillFormField("HALPhone", false)%>">
								<%pcs_RequiredImageTag "HALPhone", false%>
								ext.
								<INPUT type="text" size="6" name="HALPhoneExtension" id="HALPhoneExtension" value="<%=pcf_FillFormField("HALPhoneExtension", false)%>">
								<%pcs_RequiredImageTag "HALPhoneExtension", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Pager:</b></td>
							<td align="left">
								<INPUT type="text" name="HALPager" id="HALPager" value="<%=pcf_FillFormField("HALPager", false)%>">
								<%pcs_RequiredImageTag "HALPager", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Fax:</b></td>
							<td align="left">
								<INPUT type="text" name="HALFax" id="HALFax" value="<%=pcf_FillFormField("HALFax", false)%>">
								<%pcs_RequiredImageTag "HALFax", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Email:</b></td>
							<td align="left">
								<INPUT type="text" name="HALEmail" id="HALEmail" value="<%=pcf_FillFormField("HALEmail", false)%>">
								<%pcs_RequiredImageTag "HALEmail", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Address:</b></td>
							<td align="left">
								<INPUT type="text" name="HALLine1" id="HALLine1" value="<%=pcf_FillFormField("HALLine1", false)%>">
								<%pcs_RequiredImageTag "HALLine1", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>City:</b></td>
							<td align="left">
								<INPUT type="text" name="HALCity" id="HALCity" value="<%=pcf_FillFormField("HALCity", false)%>">
								<%pcs_RequiredImageTag "HALCity", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>State or Province Code:</b></td>
							<td align="left">
								<INPUT type="text" name="HALStateOrProvinceCode" id="HALStateOrProvinceCode" value="<%=pcf_FillFormField("HALStateOrProvinceCode", false)%>">
								<%pcs_RequiredImageTag "HALStateOrProvinceCode", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Postal Code:</b></td>
							<td align="left">
								<INPUT name="HALPostalCode" type="text" id="HALPostalCode" value="<%=pcf_FillFormField("HALPostalCode", false)%>">
								<%pcs_RequiredImageTag "HALPostalCode", false%>
							</td>
											  </tr>
						<tr>
							<td align="right"><b>Urbanization Code:</b></td>
							<td align="left">
								<INPUT name="HALUrbanizationCode" type="text" id="HALUrbanizationCode" value="<%=pcf_FillFormField("HALUrbanizationCode", false)%>">
								<%pcs_RequiredImageTag "HALUrbanizationCode", false%>
							</td>
											  </tr>
						<tr>
							<td align="right"><b>Country Code:</b></td>
							<td align="left">
								<INPUT name="HALCountryCode" type="text" id="HALCountryCode" value="<%=pcf_FillFormField("HALCountryCode", false)%>">
								<%pcs_RequiredImageTag "HALCountryCode", false%>
							</td>

						</tr>
						<tr>
						<td align="right"><b>Residential:</b></td>
						<td>
							<input type="checkbox" name="HALResidential" value="true" class="clearBorder" <%=pcf_CheckOption("HALResidential", "true")%>>
						</td>
						</tr>
                    <tr>
                      <td align="right"><b>Location Type:</b></td>
                      <td align="left">
                        <select name="HALLocationType" id="HALLocationType">
                          <option value="" <%=pcf_SelectOption("HALLocationType","")%>>Select One</option>
                          <option value="FEDEX_EXPRESS_STATION" <%=pcf_SelectOption("HALLocationType","FEDEX_EXPRESS_STATION")%>>FedEx Express Station</option>
                          <option value="FEDEX_FACILITY" <%=pcf_SelectOption("HALLocationType","FEDEX_FACILITY")%>>FedEx Facility</option>
                          <option value="FEDEX_FREIGHT_SERVICE_CENTER" <%=pcf_SelectOption("HALLocationType","FEDEX_FREIGHT_SERVICE_CENTER")%>>FedEx Freight Service Center</option>
                          <option value="FEDEX_GROUND_TERMINAL" <%=pcf_SelectOption("HALLocationType","FEDEX_GROUND_TERMINAL")%>>FedEx Ground Terminal</option>
                          <option value="FEDEX_HOME_DELIVERY_STATION" <%=pcf_SelectOption("HALLocationType","FEDEX_HOME_DELIVERY_STATION")%>>FedEx Home Delivery Station</option>
                          <option value="FEDEX_OFFICE" <%=pcf_SelectOption("HALLocationType","FEDEX_OFFICE")%>>FedEx Office</option>
                          <option value="FEDEX_SHIPSITE" <%=pcf_SelectOption("HALLocationType","FEDEX_SHIPSITE")%>>FedEx Ship Site</option>
                          <option value="FEDEX_SMART_POST_HUB" <%=pcf_SelectOption("HALLocationType","FEDEX_SMART_POST_HUB")%>>FedEx SmartPost Hub</option>
                        </select>
                        <%pcs_RequiredImageTag "HALLocationType", false%>
                      </td>
                    </tr>
											</Table>
										</div>
									</td>
								</tr>
                
								<tr>
									<th colspan="2">
									<script type=text/javascript>
									function jfSODGShip(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSODGShip.checked == true) {
									document.getElementById('SODGShip').style.display='';
									}else{
									document.getElementById('SODGShip').style.display='none';
									}
									}
									
									function jfSODGContainers(numContainers) {										
										for (var i = 0; i < <%= DGMaxContainers %>; i++) {
											if (i < numContainers) {
												document.getElementById("DGContainer" + (i + 1)).style.display = '';
											} else {
												document.getElementById("DGContainer" + (i + 1)).style.display = 'none';
											}
										}
									}
									
									</script>
									<%
									if Session("pcAdminbSODGShip")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSODGShip();" name="bSODGShip" id="bSODGShip" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSODGShip", "1")%>>
									Dangerous Goods Shipment </th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="SODGShip" <%=pcv_strDisplayStyle%>>
										<Table>
						<tr>
							<td align="right"><b>Accessibility:</b></td>
							<td align="left">
              	<select name="DGAccessibility" id="DGAccessibility">
                	<option value="">Select Option</option>
                	<option value="ACCESSIBLE" <%=pcf_SelectOption("DGAccessibility","ACCESSIBLE")%>>Accessible</option>
                  <option value="INACCESSIBLE" <%=pcf_SelectOption("DGAccessibility","INACCESSIBLE")%>>Inaccessible</option>
                </select>
								<%pcs_RequiredImageTag "DGAccessibility", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Cargo Aircraft Only:</b></td>
							<td align="left">
								<INPUT type="radio" name="DGAircraftOnly" id="DGAircraftOnly" value="1" class="clearBorder" <%= pcf_CheckOption("DGAircraftOnly", "1") %>>Yes
								<INPUT type="radio" name="DGAircraftOnly" id="DGAircraftOnly" value="0" class="clearBorder" <%= pcf_CheckOption("DGAircraftOnly", "0") %>>No
							</td>
						</tr>
						<tr>
							<td align="right"><b>Hazardous Materials?</b></td>
							<td align="left">
								<INPUT type="radio" name="DGHazardousMaterials" id="DGHazardousMaterials" value="1" class="clearBorder" <%= pcf_CheckOption("DGHazardousMaterials", "1") %>>Yes
								<INPUT type="radio" name="DGHazardousMaterials" id="DGHazardousMaterials" value="0" class="clearBorder" <%= pcf_CheckOption("DGHazardousMaterials", "0") %>>No
							</td>
						</tr>
						<tr>
							<td align="right"><b>Battery?</b></td>
							<td align="left">
								<INPUT type="radio" name="DGBattery" id="DGBattery" value="1" class="clearBorder" <%= pcf_CheckOption("DGBattery", "1") %>>Yes
								<INPUT type="radio" name="DGBattery" id="DGBattery" value="0" class="clearBorder" <%= pcf_CheckOption("DGBattery", "0") %>>No
							</td>
						</tr>
						<tr>
							<td align="right"><b>ORM-D?</b></td>
							<td align="left">
								<INPUT type="radio" name="DGORMD" id="DGORMD" value="1" class="clearBorder" <%= pcf_CheckOption("DGORMD", "1") %>>Yes
								<INPUT type="radio" name="DGORMD" id="DGORMD" value="0" class="clearBorder" <%= pcf_CheckOption("DGORMD", "0") %>>No
							</td>
						</tr>
						<tr>
							<td align="right"><b>Container Type:</b></td>
							<td align="left">
								<INPUT type="text" name="DGContainerType" id="DGContainerType" value="<%=pcf_FillFormField("DGContainerType", false)%>">
								<%pcs_RequiredImageTag "DGContainerType", false%>
							</td>
						</tr>
            <%
							numContainers = 0
							if IsNumeric(session("pcAdminDGContainerCount")) then
								numContainers = cint(session("pcAdminDGContainerCount"))
							end if
						%>
						<tr>
							<td align="right"><b>Number of Containers:</b></td>
							<td align="left">
              	<select name="DGContainerCount" id="DGContainerCount" onChange="return jfSODGContainers(this.options[this.selectedIndex].value);">
              	<% for i = 0 to DGMaxContainers %>
                	<option value="<%= i %>" <%=pcf_SelectOption("DGContainerCount", CStr(i))%>><%= i %></option>
                <% next %>
                </select>
								<%pcs_RequiredImageTag "DGContainerCount", false%>
							</td>
						</tr>
            <% for i = 1 to DGMaxContainers %>
              <tr id="DGContainer<%= i%>" style="<% if i > numContainers then response.write "display: none" %>">
              	<td colspan="2">
                  <table style="border: 1px solid #DFDFDF; padding: 8px; margin: 4px; border-radius: 4px">
                    <tr>
                      <td colspan="2" style="border-bottom: 1px dashed #999;"><b>Container <%= i %></b></td>
                    </tr>
                    <tr>
                      <td align="right"><b>Commodity ID:</b></td>
                      <td align="left">
                        <INPUT type="text" name="DGCommodityID<%= i %>" id="DGCommodityID<%= i %>" value="<%=pcf_FillFormField("DGCommodityID" & i, false)%>">
                        <%pcs_RequiredImageTag "DGCommodityID" & i, false%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"><b>Packing Group:</b></td>
                      <td align="left">
                        <select name="DGPackingGroup<%= i %>" id="DGPackingGroup<%= i %>">
                          <option value="DEFAULT" <%=pcf_SelectOption("DGPackingGroup" & i,"DEFAULT")%>>Default</option>
                          <option value="I" <%=pcf_SelectOption("DGPackingGroup" & i,"I")%>>I</option>
                          <option value="II" <%=pcf_SelectOption("DGPackingGroup" & i,"II")%>>II</option>
                          <option value="III" <%=pcf_SelectOption("DGPackingGroup" & i,"III")%>>III</option>
                        </select>
                        <%pcs_RequiredImageTag "DGPackingGroup" & i, false%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"><b>Cargo Aircraft Only:</b></td>
                      <td align="left">
                        <INPUT type="radio" name="DGContainerAircraftOnly<%= i %>" id="DGContainerAircraftOnly<%= i %>" value="1" class="clearBorder" <%= pcf_CheckOption("DGContainerAircraftOnly" & i, "1") %>>Yes
                        <INPUT type="radio" name="DGContainerAircraftOnly<%= i %>" id="DGContainerAircraftOnly<%= i %>" value="0" class="clearBorder" <%= pcf_CheckOption("DGContainerAircraftOnly" & i, "0") %>>No
                        <%pcs_RequiredImageTag "DGContainerAircraftOnly" & i, false%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"><b>Packing Instructions:</b></td>
                      <td align="left">
                        <INPUT type="text" name="DGPackingInstructions<%= i %>" id="DGPackingInstructions<%= i %>" value="<%=pcf_FillFormField("DGPackingInstructions" & i, false)%>">
                        <%pcs_RequiredImageTag "DGPackingInstructions" & i, false%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"><b>Proper Shipping Name:</b></td>
                      <td align="left">
                        <INPUT type="text" name="DGShippingName<%= i %>" id="DGShippingName<%= i %>" value="<%=pcf_FillFormField("DGShippingName" & i, false)%>">
                        <%pcs_RequiredImageTag "DGShippingName" & i, false%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"><b>Hazard Class:</b></td>
                      <td align="left">
                        <INPUT type="text" name="DGHazardClass<%= i %>" id="DGHazardClass<%= i %>" value="<%=pcf_FillFormField("DGHazardClass" & i, false)%>">
                        <%pcs_RequiredImageTag "DGHazardClass" & i, false%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"><b>Quantity Amount:</b></td>
                      <td align="left">
                        <INPUT type="text" name="DGQuantityAmount<%= i %>" id="DGQuantityAmount<%= i %>" value="<%=pcf_FillFormField("DGQuantityAmount" & i, false)%>">
                        <%pcs_RequiredImageTag "DGQuantityAmount" & i, false%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"><b>Quantity Units:</b></td>
                      <td align="left">
                        <INPUT type="text" name="DGQuantityUnits<%= i %>" id="DGQuantityUnits<%= i %>" value="<%=pcf_FillFormField("DGQuantityUnits" & i, false)%>">
                        <%pcs_RequiredImageTag "DGQuantityUnits" & i, false%>
                      </td>
                    </tr>
                  </table>
               	</td>
              </tr>
            <% next %>
						<tr>
							<td align="right"><b>Package Count:</b></td>
							<td align="left">
								<INPUT type="text" name="DGPackagingCount" id="DGPackagingCount" value="<%=pcf_FillFormField("DGPackagingCount", false)%>">
								<%pcs_RequiredImageTag "DGPackagingCount", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Package Units:</b></td>
							<td align="left">
								<INPUT type="text" name="DGPackagingUnits" id="DGPackagingUnits" value="<%=pcf_FillFormField("DGPackagingUnits", false)%>">
								<%pcs_RequiredImageTag "DGPackagingUnits", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Contact Name:</b></td>
							<td align="left">
								<INPUT type="text" name="DGContactName" id="DGContactName" value="<%=pcf_FillFormField("DGContactName", false)%>">
								<%pcs_RequiredImageTag "DGContactName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Contact Title:</b></td>
							<td align="left">
								<INPUT type="text" name="DGContactTitle" id="DGContactTitle" value="<%=pcf_FillFormField("DGContactTitle", false)%>">
								<%pcs_RequiredImageTag "DGContactTitle", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Contact Place:</b></td>
							<td align="left">
								<INPUT type="text" name="DGContactPlace" id="DGContactPlace" value="<%=pcf_FillFormField("DGContactPlace", false)%>">
								<%pcs_RequiredImageTag "DGContactPlace", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Emergency Contact #:</b></td>
							<td align="left">
								<INPUT type="text" name="DGEmergencyContactNumber" id="DGEmergencyContactNumber" value="<%=pcf_FillFormField("DGEmergencyContactNumber", false)%>">
								<%pcs_RequiredImageTag "DGEmergencyContactNumber", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Offeror:</b></td>
							<td align="left">
								<INPUT type="text" name="DGOfferor" id="DGOfferor" value="<%=pcf_FillFormField("DGOfferor", false)%>">
								<%pcs_RequiredImageTag "DGOfferor", false%>
							</td>
						</tr>
										</Table>
									</div>
									</td>
								</tr>

								<tr>
									<th colspan="2">
									<script type=text/javascript>
									function jfSODryIce(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSODryIce.checked == true) {
									document.getElementById('SODryIce').style.display='';
									}else{
									document.getElementById('SODryIce').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminbSODryIce")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSODryIce();" name="bSODryIce" id="bSODryIce" type="checkbox" class="clearBorder" value=1 <%=pcf_CheckOption("bSODryIce", "1")%>>
									Dry Ice Shipment </th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="SODryIce" <%=pcv_strDisplayStyle%>>
										<table>
											<tr>
												<td width="273" align="right"><b>Dry Ice Package Count:</b></td>
												<td width="345" align="left">
													<input type="text" name="SDIPackageCount" id="SDIPackageCount" value="<%=pcf_FillFormField("SDIPackageCount", false)%>">
													<%pcs_RequiredImageTag "SDIPackageCount", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Dry Ice Weight:</b></td>
												<td align="left">
													<INPUT type="text" name="SDIValue" id="SDIValue" value="<%=pcf_FillFormField("SDIValue", false)%>">&nbsp;&nbsp;KG
													<%pcs_RequiredImageTag "SDIValue", false%>
												</td>
										  </tr>
											<!--<tr>
												<td align="right"><b>Weight Units:</b></td>
												<td align="left">
                          <select name="SDIUnit" id="SDIUnit">
                          	<option value="KG" <%=pcf_SelectOption("SDIUnit","KG")%>>Kilograms (KG)</option>
                          	<option value="LB" <%=pcf_SelectOption("SDIUnit","LB")%>>Pounds (LB)</option>
                          </select>
													<%pcs_RequiredImageTag "SDIUnit", false%></td>
											</tr>-->
										</table>
									</div>
									</td>
								</tr>
							</table>
				</td>
			</tr>
			<tr>
				<td width="50%" class="pcCPshipping"><span class="titleShip">International Settings</span></td>
				<td class="pcCPshipping" align="right">
				<i>(Check box to view)</i>&nbsp;
				<script type=text/javascript>
						function jfInternational(){

						var selectValDom = document.forms['form1'];
						if (selectValDom.bInternational.checked == true) {
						document.getElementById('International').style.display='';
						}else{
						document.getElementById('International').style.display='none';
						}
						}
						</script>
						<%
						if Session("pcAdminbInternational")="1" then
							pcv_strDisplayStyle="style=""display:block"""
						else
							pcv_strDisplayStyle="style=""display:none"""
						end if
						%>
			<input onClick="jfInternational();" name="bInternational" id="bInternational" type="checkbox" class="clearBorder" value="true" <%=pcf_CheckOption("bInternational", "1")%>>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<div id="International" <%=pcv_strDisplayStyle%>>
							<table width="100%">
								<tr>
									<th colspan="2">International Settings</th>
						</tr>

						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td width="24%" align="right" valign="top"><b>Documents  Shipment:</b></td>
							<td width="76%" align="left">
							<input type="radio" name="DocumentsOnly" value="0" class="clearBorder" <%= pcf_CheckOption("DocumentsOnly", "0") %>>Not Applicable
							<input type="radio" name="DocumentsOnly" value="1" class="clearBorder" <%= pcf_CheckOption("DocumentsOnly", "1") %>>Non-Documents
							<input type="radio" name="DocumentsOnly" value="2" class="clearBorder" <%= pcf_CheckOption("DocumentsOnly", "2") %>>Documents Only
							</td>
						</tr>


						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Customs Clearance Detail</th>
						</tr>
						<tr>
							<td width="25%" align="right" valign="top"><b>Customs Amount:</b></td>
							<td width="75%" align="left">
								<% 
									FedEx_CurrencyControl "CVAmount", "CVCurrency", isRequiredCVAmount 
								%>
							</td>
						</tr>
						<tr>
						  <td align="right"><b>Insurance Charges Amount:</b></td>
						  <td align="left">
								<% 
									FedEx_CurrencyControl "CICAmount", "CICCurrency", false 
								%>
							</td>
						  </tr>
						<tr>
						  <td align="right"><b>Taxes or Miscellaneous Charges Amount:</b></td>
							<td align="left">
								<% 
									FedEx_CurrencyControl "CMCAmount", "CMCCurrency", false 
								%>
							</td>
						  </tr>
						<tr>
							<td align="right"><b>Freight Charges Amount:</b></td>
							<td align="left">
								<% 
									FedEx_CurrencyControl "CFCAmount", "CFCCurrency", false 
								%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Purpose:</b></td>
							<td align="left">
							 <select name="CCIPurpose" id="CCIPurpose">
								<option value="" <%=pcf_SelectOption("CCIPurpose","")%>>Select Option</option>
								<option value="NOT_SOLD" <%=pcf_SelectOption("CCIPurpose","NOT_SOLD")%>>Not Sold</option>
								<option value="PERSONAL_EFFECTS" <%=pcf_SelectOption("CCIPurpose","PERSONAL_EFFECTS")%>>Personal Effects</option>
								<option value="REPAIR_AND_RETURN" <%=pcf_SelectOption("CCIPurpose","REPAIR_AND_RETURN")%>>Repair and Return</option>
								<option value="SAMPLE" <%=pcf_SelectOption("CCIPurpose","SAMPLE")%>>Sample</option>
								<option value="SOLD" <%=pcf_SelectOption("CCIPurpose","SOLD")%>>Sold</option>
								</select>
								<%pcs_RequiredImageTag "CCIPurpose", false%>
						</tr>
						<tr>
							<td align="right"><b>Customer Reference:</b></td>
							<td align="left"><INPUT type="text" name="CCICustomerReference" id="CCICustomerReference" value="<%=pcf_FillFormField("CCICustomerReference", false)%>">
							<%pcs_RequiredImageTag "CCICustomerReference", false%></td>
						</tr>
						<tr>
							<td align="right"><b>Terms of Sale:</b></td>
							<td align="left">
							 <select name="CCITermsOfSale" id="CCITermsOfSale">
               	<option value="">Select Option</option>
								<option value="CFR_OR_CPT" <%=pcf_SelectOption("CCITermsOfSale","CFR_OR_CPT")%>>Cost and Freight/Carriage Paid To</option>
								<option value="CIF_OR_CIP" <%=pcf_SelectOption("CCITermsOfSale","CIF_OR_CIP")%>>Cost Insurance and Freight/Carraige Insurance Paid</option>
								<option value="DAT" <%=pcf_SelectOption("CCITermsOfSale","DAT")%>>DAT</option>
								<option value="DDP" <%=pcf_SelectOption("CCITermsOfSale","DDP")%>>Delivered Duty Paid</option>
								<option value="DDU" <%=pcf_SelectOption("CCITermsOfSale","DDU")%>>Delivered Duty Unpaid</option>
								<option value="EXW" <%=pcf_SelectOption("CCITermsOfSale","EXW")%>>Ex Works</option>
								<option value="FOB" <%=pcf_SelectOption("CCITermsOfSale","FOB")%>>Free On Board</option>
								<option value="FCA" <%=pcf_SelectOption("CCITermsOfSale","FCA")%>>Free Carrier</option>
								</select>
								<%pcs_RequiredImageTag "CCITermsOfSale", false%>
						</tr>
						<tr>
							<td align="right"><b>Invoice Number:</b></td>
							<td align="left"><INPUT type="text" name="CCIInvoiceNumber" id="CCIInvoiceNumber" value="<%=pcf_FillFormField("CCIInvoiceNumber", false)%>">
							<%pcs_RequiredImageTag "CCIInvoiceNumber", false%></td>
						</tr>
						<tr>
							<td align="right"><b>Comments:</b></td>
							<td align="left"><INPUT type="text" name="CCIComments" id="CCIComments" value="<%=pcf_FillFormField("CCIComments", false)%>">
							<%pcs_RequiredImageTag "CCIComments", false%></td>
						</tr>
						<tr>
							<td align="right"><b>Customs Option Type:</b></td>
							<td align="left">
							 <select name="CCDOptionType" id="CCDOptionType">
               	<option value="">Select Option</option>
								<option value="COURTESY_RETURN_LABEL" <%=pcf_SelectOption("CCDOptionType","COURTESY_RETURN_LABEL")%>>Courtesy Return Label</option>
								<option value="EXHIBITION_TRADE_SHOW" <%=pcf_SelectOption("CCDOptionType","EXHIBITION_TRADE_SHOW")%>>Exhibition Trade Show</option>
								<option value="FOR_REPAIR" <%=pcf_SelectOption("CCDOptionType","FOR_REPAIR")%>>For Repair</option>
                <option value="FOLLOWING_REPAIR" <%=pcf_SelectOption("CCDOptionType","FOLLOWING_REPAIR")%>>Following Repair</option>
								<option value="FAULTY_ITEM" <%=pcf_SelectOption("CCDOptionType","FAULTY_ITEM")%>>Faulty Item</option>
								<option value="ITEM_FOR_LOAN" <%=pcf_SelectOption("CCDOptionType","ITEM_FOR_LOAN")%>>Item for Loan</option>
								<option value="REJECTED" <%=pcf_SelectOption("CCDOptionType","REJECTED")%>>Rejected</option>
								<option value="REPLACEMENT" <%=pcf_SelectOption("CCDOptionType","REPLACEMENT")%>>Replacement</option>
								<option value="TRIAL" <%=pcf_SelectOption("CCDOptionType","TRIAL")%>>Trial</option>
								<option value="OTHER" <%=pcf_SelectOption("CCDOptionType","OTHER")%>>Other</option>

								</select>
								<%pcs_RequiredImageTag "CCDOptionType", false%>
						</tr>
            <tr id="CCDOptionOther">
							<td align="right"><b>Customs Option Description:</b></td>
							<td align="left">
              	<input type="text" name="CCDOptionDescription" id="CCDOptionDescription" value="<%= pcf_FillFormField("CCDOptionDescription", false) %>" />
                <%pcs_RequiredImageTag "CCDOptionDescription", isCCDOptionDescriptionRequired%>
              </td>
            </tr
						><tr>
						  <td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Commodity</th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Number of Pieces:</b></td>
							<td align="left">
								<input name="NumberOfPieces" type="text" id="NumberOfPieces" value="<%=pcf_FillFormField("NumberOfPieces", false)%>">
								<%pcs_RequiredImageTag "NumberOfPieces", isRequiredNumberOfPieces %>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Description:</b></td>
							<td align="left">
								<input name="Description" type="text" id="Description" value="<%=pcf_FillFormField("Description", false)%>">
								<%pcs_RequiredImageTag "Description", isRequiredDescription %>
							</td>
						</tr>



						<tr>
							<td align="right" valign="top"><b>Country Code of Manufacture:</b></td>
							<td align="left">
								<input name="CountryOfManufacture" type="text" id="CountryOfManufacture" value="<%=pcf_FillFormField("CountryOfManufacture", false)%>">
								<%pcs_RequiredImageTag "CountryOfManufacture", isRequiredCountryOfManufacture %>
							(e.g. US) </td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Harmonized Code:</b></td>
							<td align="left">
								<input name="HarmonizedCode" type="text" id="HarmonizedCode" value="<%=pcf_FillFormField("HarmonizedCode", false)%>">
								<%pcs_RequiredImageTag "HarmonizedCode", false%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Weight:</b></td>	
							<td align="left">						
								<% 
									FedEx_WeightControl "CommodityWeightValue", "CommodityWeightUnits", isRequiredCommodityWeight 
								%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Quantity:</b></td>
							<td align="left">
								<input name="CommodityQuantity" type="text" id="CommodityQuantity" value="<%=pcf_FillFormField("CommodityQuantity", false)%>">
								<%pcs_RequiredImageTag "CommodityQuantity", isRequiredCommodityQuantity %>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Quantity Units:</b></td>
							<td align="left">
								<input name="CommodityQuantityUnits" type="text" id="CommodityQuantityUnits" value="<%=pcf_FillFormField("CommodityQuantityUnits", false)%>">
								<%pcs_RequiredImageTag "CommodityQuantityUnits", isRequiredCommodityQuantityUnits %>
							(e.g. EA) </td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Unit Price:</b></td>
							<td align="left">
								<% 
									FedEx_CurrencyControl "CommodityUnitPrice", "CommodityUnitCurrency", isRequiredCommodityUnitPrice 
								%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Total Customs Value:</b></td>
							<td align="left">
								<% 
									FedEx_CurrencyControl "CommodityCustomsValue", "CommodityCustomsCurrency", isRequiredCommodityUnitPrice 
								%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>B13A Filing Option:</b></td>
							<td align="left">
							 <select name="B13AFilingOption" id="B13AFilingOption">
								<option value="" <%=pcf_SelectOption("B13AFilingOption","")%>>Select Option</option>
								<option value="NOT_REQUIRED" <%=pcf_SelectOption("B13AFilingOption","NOT_REQUIRED")%>>Not Required</option>
								<option value="FILED_ELECTRONICALLY" <%=pcf_SelectOption("B13AFilingOption","FILED_ELECTRONICALLY")%>>Filed Electronically</option>
								<option value="MANUALLY_ATTACHED" <%=pcf_SelectOption("B13AFilingOption","MANUALLY_ATTACHED")%>>Manually Attached</option>
								<option value="SUMMARY_REPORTING" <%=pcf_SelectOption("B13AFilingOption","SUMMARY_REPORTING")%>>Summary Reporting</option>
								</select>
								<%pcs_RequiredImageTag "B13AFilingOption", false%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Export Compliance Statement:</b></td>
							<td align="left">
              	<input type="text" name="ExportComplianceStatement" id="ExportComplianceStatement" value="<%=pcf_FillFormField("ExportComplianceStatement", false)%>"/>
                
								<!--<select name="ExportComplianceStatement" id="ExportComplianceStatement">
								<option value="NO EEI 30.36" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.36")%>>NO EEI 30.36</option>
								<option value="NO EEI 30.37(a)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(a)")%>>NO EEI 30.37(a)</option>
								<option value="NO EEI 30.37(b)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(b)")%>>NO EEI 30.37(b)</option>
								<option value="NO EEI 30.37(f)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(f)")%>>NO EEI 30.37(f)</option>
								<option value="NO EEI 30.37(g)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(g)")%>>NO EEI 30.37(g)</option>
								<option value="NO EEI 30.37(h)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(h)")%>>NO EEI 30.37(h)</option>
								<option value="NO EEI 30.37(i)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(i)")%>>NO EEI 30.37(i)</option>
								<option value="NO EEI 30.37(j)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(j)")%>>NO EEI 30.37(j)</option>
								<option value="NO EEI 30.37(k)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(k)")%>>NO EEI 30.37(k)</option>
								<option value="NO EEI 30.37(l)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(l)")%>>NO EEI 30.37(l)</option>
								<option value="NO EEI 30.37(p)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(p)")%>>NO EEI 30.37(p)</option>
								<option value="NO EEI 30.39" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.39")%>>NO EEI 30.39</option>
								<option value="NO EEI 30.40(a)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(a)")%>>NO EEI 30.40(a)</option>
								<option value="NO EEI 30.40(b)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(b)")%>>NO EEI 30.40(b)</option>
								<option value="NO EEI 30.40(c)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(c)")%>>NO EEI 30.40(c)</option>
								<option value="NO EEI 30.40(d)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(d)")%>>NO EEI 30.40(d)</option>
								<option value="NO EEI 30.02 (d)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.02(d)")%>>NO EEI 30.02(d)</option>
								</select>-->
								<%pcs_RequiredImageTag "ExportComplianceStatement", false%>
							</td>
						</tr>
								<tr>
								  <td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2">International Duties and Taxes</th>
							  </tr>
								<tr>
									<td width="25%" align="right" valign="top"><b>Duties Payor:</b></td>
									<td align="left">
										<select name="DutiesPayorType" id="DutiesPayorType">
										<option value="SENDER" <%=pcf_SelectOption("DutiesPayorType","SENDER")%>>Sender</option>
										<option value="RECIPIENT" <%=pcf_SelectOption("DutiesPayorType","RECIPIENT")%>>Recipient</option>
										<option value="THIRD_PARTY" <%=pcf_SelectOption("DutiesPayorType","THIRD_PARTY")%>>3rd Party</option>
										</select>
										<%pcs_RequiredImageTag "DutiesPayorType", false%>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Duties Payor Account#:</b></td>
									<td align="left">
										<input name="DutiesAccountNumber" type="text" id="DutiesAccountNumber" value="<%=pcf_FillFormField("DutiesAccountNumber", false)%>">
										<%pcs_RequiredImageTag "DutiesAccountNumber", isRequiredDutiesAccountNumber %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Duties Payor Name:</b></td>
									<td align="left">
										<input name="DutiesPersonName" type="text" id="DutiesPersonName" value="<%=pcf_FillFormField("DutiesPersonName", false)%>">
										<%pcs_RequiredImageTag "DutiesPersonName", isRequiredDutiesPersonName %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Duties Payor Country Code:</b></td>
									<td align="left">
									<input name="DutiesCountryCode" type="text" id="DutiesCountryCode" value="<%=pcf_FillFormField("DutiesCountryCode", false)%>">
									<%pcs_RequiredImageTag "DutiesCountryCode", isRequiredDutiesCountryCode %>
									(e.g. US) </td>
								</tr>
								<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2">Importer of Record</th>
							</tr>
							<tr>
								<td align="right"><b>Person Name:</b></td>
								<td align="left"><INPUT type="text" name="IORPersonName" id="IORPersonName" value="<%=pcf_FillFormField("IORPersonName", false)%>">
								<%pcs_RequiredImageTag "IORPersonName", false%></td>
							</tr>
							<tr>
								<td align="right"><b>Company Name:</b></td>
								<td align="left"><INPUT type="text" name="IORCompanyName" id="IORCompanyName" value="<%=pcf_FillFormField("IORCompanyName", false)%>">
								<%pcs_RequiredImageTag "IORCompanyName", false%></td>
							</tr>
							<tr>
								<td align="right"><b>Phone Number:</b></td>
								<td align="left"><INPUT type="text" name="IORPhoneNumber" id="IORPhoneNumber" value="<%=pcf_FillFormField("IORPhoneNumber", false)%>">
								<%pcs_RequiredImageTag "IORPhoneNumber", false%></td>
							</tr>
							<tr>
								<td align="right"><b>Street Address:</b></td>
								<td align="left"><INPUT type="text" name="IORAddress" id="IORAddress" value="<%=pcf_FillFormField("IORAddress", false)%>">
								<%pcs_RequiredImageTag "IORAddress", false%></td>
							</tr>
							<tr>
								<td align="right"><b>City:</b></td>
								<td align="left"><INPUT type="text" name="IORCity" id="IORCity" value="<%=pcf_FillFormField("IORCity", false)%>">
								<%pcs_RequiredImageTag "IORCity", false%></td>
							</tr>
							<tr>
								<td align="right"><b>State/Province Code:</b></td>
								<td align="left"><INPUT type="text" name="IORStateOrProvince" id="IORStateOrProvince" value="<%=pcf_FillFormField("IORStateOrProvince", false)%>">
								<%pcs_RequiredImageTag "IORStateOrProvince", false%></td>
							</tr>
							<tr>
								<td align="right"><b>Country Code:</b></td>
								<td align="left"><INPUT type="text" name="IORCountryCode" id="IORCountryCode" value="<%=pcf_FillFormField("IORCountryCode", false)%>">
								<%pcs_RequiredImageTag "IORCountryCode", false%></td>
							</tr>
							<tr>
								<td align="right"><b>Postal Code:</b></td>
								<td align="left"><INPUT type="text" name="IORPostalCode" id="IORPostalCode" value="<%=pcf_FillFormField("IORPostalCode", false)%>">
								<%pcs_RequiredImageTag "IORPostalCode", false%></td>
							</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Sender TIN Details</th>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Sender TIN Number:</b></td>
									<td align="left">
										<input name="SenderTINNumber" type="text" id="SenderTINNumber" value="<%=pcf_FillFormField("SenderTINNumber", false)%>">
										<%pcs_RequiredImageTag "SenderTINNumber", false%>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Sender TIN or DUNS Type:</b></td>
									<td align="left">
										<select name="SenderTINType" id="SenderTINType">
											<option value="BUSINESS_NATIONAL" <%=pcf_SelectOption("SenderTINType","BUSINESS_NATIONAL")%>>Business National</option>
											<option value="BUSINESS_STATE" <%=pcf_SelectOption("SenderTINType","BUSINESS_STATE")%>>Business State</option>
											<option value="BUSINESS_UNION" <%=pcf_SelectOption("SenderTINType","BUSINESS_UNION")%>>Business Union</option>
											<option value="PERSONAL_NATIONAL" <%=pcf_SelectOption("SenderTINType","PERSONAL_NATIONAL")%>>Personal National</option>
											<option value="PERSONAL_STATE" <%=pcf_SelectOption("SenderTINType","PERSONAL_STATE")%>>Personal State</option>
										</select>
										<%pcs_RequiredImageTag "SenderTINType", false%>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">
									<script type=text/javascript>
									function jfISOBrokerSelect(){
									var selectValDom = document.forms['form1'];
									if (selectValDom.bISOBrokerSelect.checked == true) {
									document.getElementById('ISOBrokerSelect').style.display='';
									}else{
									document.getElementById('ISOBrokerSelect').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminbISOBrokerSelect")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfISOBrokerSelect();" name="bISOBrokerSelect" id="bISOBrokerSelect" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bISOBrokerSelect", "1")%>>&nbsp;
									Broker Select Special Services Option<br>
									<br>
									<div class="pcCPnotes">Broker Select Option should be used for FedEx Express&reg; shipments only.</div></th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="ISOBrokerSelect" <%=pcv_strDisplayStyle%>>
										<Table>
											<tr>
												<td align="right"><b>Type:</b></td>
												<td align="left">
                        	<select name="BSOType" id="BSOType">
                          	<option value="IMPORT" <%=pcf_SelectOption("BSOType","IMPORT")%>>Import</option>
                            <option value="EXPORT" <%=pcf_SelectOption("BSOType","EXPORT")%>>Export</option>
                          </select>
													<%pcs_RequiredImageTag "BSOType", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Account Number:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOAccountNumber" id="BSOAccountNumber" value="<%=pcf_FillFormField("BSOAccountNumber", false)%>">
													<%pcs_RequiredImageTag "BSOAccountNumber", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Tin Type:</b></td>
												<td align="left">
                          <select name="BSOTinType" id="BSOTinType">
                            <option value="BUSINESS_NATIONAL" <%=pcf_SelectOption("BSOTinType","BUSINESS_NATIONAL")%>>Business National</option>
                            <option value="BUSINESS_STATE" <%=pcf_SelectOption("BSOTinType","BUSINESS_STATE")%>>Business State</option>
                            <option value="BUSINESS_UNION" <%=pcf_SelectOption("BSOTinType","BUSINESS_UNION")%>>Business Union</option>
                            <option value="PERSONAL_NATIONAL" <%=pcf_SelectOption("BSOTinType","PERSONAL_NATIONAL")%>>Personal National</option>
                            <option value="PERSONAL_STATE" <%=pcf_SelectOption("BSOTinType","PERSONAL_STATE")%>>Personal State</option>
                          </select>
                          <%pcs_RequiredImageTag "BSOTinType", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Tin Number:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOTinNumber" id="BSOTinNumber" value="<%=pcf_FillFormField("BSOTinNumber", false)%>">
													<%pcs_RequiredImageTag "BSOTinNumber", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Contact ID:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOContactID" id="BSOContactID" value="<%=pcf_FillFormField("BSOContactID", false)%>">
													<%pcs_RequiredImageTag "BSOContactID", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Title:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOTitle" id="BSOTitle" value="<%=pcf_FillFormField("BSOTitle", false)%>">
													<%pcs_RequiredImageTag "BSOTitle", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Contact Name:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOPersonName" id="BSOPersonName" value="<%=pcf_FillFormField("BSOPersonName", false)%>">
													<%pcs_RequiredImageTag "BSOPersonName", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Company Name:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOCompanyName" id="BSOCompanyName" value="<%=pcf_FillFormField("BSOCompanyName", false)%>">
													<%pcs_RequiredImageTag "BSOCompanyName", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Phone Number:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOPhoneNumber" id="BSOPhoneNumber" value="<%=pcf_FillFormField("BSOPhoneNumber", false)%>">
													<%pcs_RequiredImageTag "BSOPhoneNumber", false%>
                          
                          ext.
													<INPUT type="text" name="BSOPhoneExtension" id="BSOPhoneExtension" size="4" value="<%=pcf_FillFormField("BSOPhoneExtension", false)%>">
													<%pcs_RequiredImageTag "BSOPhoneExtension", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Email:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOEmailAddress" id="BSOEmailAddress" value="<%=pcf_FillFormField("BSOEmailAddress", false)%>">
													<%pcs_RequiredImageTag "BSOEmailAddress", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Address:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOStreetLines" id="BSOStreetLines" value="<%=pcf_FillFormField("BSOStreetLines", false)%>">
													<%pcs_RequiredImageTag "BSOStreetLines", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>City:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOCity" id="BSOCity" value="<%=pcf_FillFormField("BSOCity", false)%>">
													<%pcs_RequiredImageTag "BSOCity", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>State/Province Code:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOStateOrProvinceCode" id="BSOStateOrProvinceCode" value="<%=pcf_FillFormField("BSOStateOrProvinceCode", false)%>">
													<%pcs_RequiredImageTag "BSOStateOrProvinceCode", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Postal Code:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOPostalCode" id="BSOPostalCode" value="<%=pcf_FillFormField("BSOPostalCode", false)%>">
													<%pcs_RequiredImageTag "BSOPostalCode", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Country Code:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOCountryCode" id="BSOCountryCode" value="<%=pcf_FillFormField("BSOCountryCode", false)%>">
													<%pcs_RequiredImageTag "BSOCountryCode", false%>
												</td>
											</tr>
										</Table>
									</div>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
             <tr>
              <th colspan="2">
								<script type=text/javascript>
									function jfNAFTAOption() {
										var selectValDom = document.forms['form1'];
										if (selectValDom.bNAFTAOption.checked == true) {
											document.getElementById('bNAFTAContainer').style.display='';
										}else{
											document.getElementById('bNAFTAContainer').style.display='none';
										}
									}
								</script>
									<%
									if Session("pcAdminbNAFTAOption")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfNAFTAOption();" name="bNAFTAOption" id="bNAFTAOption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bNAFTAOption", "1")%>>&nbsp;
                  
                  NAFTA Certificate of Origin
                </th>
            </tr>
						<tr>
            	<td colspan="2">
                <div id="bNAFTAContainer" <%=pcv_strDisplayStyle%>>
                  <Table>
                  	<tr>
                      <td align="right" valign="top"><b>Preference Criterion:</b></td>
                      <td align="left">
                        <select name="NAFTAPreferenceCriterion" id="NAFTAPreferenceCriterion">
                        <option value="">Select Option</option>
                        <option value="A" <%=pcf_SelectOption("NAFTAPreferenceCriterion","A")%>>A</option>
                        <option value="B" <%=pcf_SelectOption("NAFTAPreferenceCriterion","B")%>>B</option>
                        <option value="C" <%=pcf_SelectOption("NAFTAPreferenceCriterion","C")%>>C</option>
                        <option value="D" <%=pcf_SelectOption("NAFTAPreferenceCriterion","D")%>>D</option>
                        <option value="E" <%=pcf_SelectOption("NAFTAPreferenceCriterion","E")%>>E</option>
                        <option value="F" <%=pcf_SelectOption("NAFTAPreferenceCriterion","F")%>>F</option>
                        <%pcs_RequiredImageTag "NAFTAPreferenceCriterion", false %>
                        </select>
                      </td>
                    </tr>            
                    <tr>
                      <td align="right" valign="top"><b>Producer Determination:</b></td>
                      <td align="left">
                        <select name="NAFTAProducerDetermination" id="NAFTAProducerDetermination">
                        <option value="">Select Option</option>
                        <option value="NO_1" <%=pcf_SelectOption("NAFTAProducerDetermination","NO_1")%>>NO_1</option>
                        <option value="NO_2" <%=pcf_SelectOption("NAFTAProducerDetermination","NO_2")%>>NO_2</option>
                        <option value="NO_3" <%=pcf_SelectOption("NAFTAProducerDetermination","NO_3")%>>NO_3</option>
                        <option value="YES" <%=pcf_SelectOption("NAFTAProducerDetermination","YES")%>>YES</option>
                        </select>
                        <%pcs_RequiredImageTag "NAFTAProducerDetermination", false %>
                      </td>
                    </tr>
                    <tr>
                      <td align="right" valign="top"><b>Producer ID:</b></td>
                      <td align="left">
                        <input name="NAFTAProducerID" type="text" id="NAFTAProducerID" value="<%=pcf_FillFormField("NAFTAProducerID", false)%>">
                        <%pcs_RequiredImageTag "NAFTAProducerID", false %>
                      </td>
                    </tr>
                    <tr>
                      <td align="right" valign="top"><b>Net Cost Method:</b></td>
                      <td align="left">
                        <select name="NAFTANetCostMethod" id="NAFTANetCostMethod">
                        <option value="">Select Option</option>
                        <option value="NC" <%=pcf_SelectOption("NAFTANetCostMethod","NC")%>>NC</option>
                        <option value="NO" <%=pcf_SelectOption("NAFTANetCostMethod","NO")%>>NO</option>
                        </select>
                        <%pcs_RequiredImageTag "NAFTANetCostMethod", false %>
                      </td>
                    </tr>
                	</Table>
                </div>
              </td>
              <tr>
                <td colspan="2" class="pcCPspacer"></td>
              </tr>
              <tr>
                <th colspan="2">
								<script type=text/javascript>
									function jfFICEOption() {
										var selectValDom = document.forms['form1'];
										if (selectValDom.bFICEOption.checked == true) {
											document.getElementById('bFICEContainer').style.display='';
										}else{
											document.getElementById('bFICEContainer').style.display='none';
										}
									}
								</script>
									<%
									if Session("pcAdminbFICEOption")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfFICEOption();" name="bFICEOption" id="bFICEOption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bFICEOption", "1")%>>&nbsp;
                  
                  FedEx International Controlled Export (FICE)
                </th>
              </tr>
              <tr>
                <td colspan="2">
                  <div id="bFICEContainer" <%=pcv_strDisplayStyle%>>
                    <Table>
                      <tr>
                        <td align="right" valign="top"><b>Type:</b></td>
                        <td align="left">
                          <select name="FICEType" id="FICEType">
                          <option value=""></option>
                          <option value="DSP_05" <%=pcf_SelectOption("FICEType","DSP_05")%>>DSP-5</option>
                          <option value="DSP_61" <%=pcf_SelectOption("FICEType","DSP_61")%>>DSP-61</option>
                          <option value="DSP_73" <%=pcf_SelectOption("FICEType","DSP_73")%>>DSP-73</option>
                          <option value="DSP_85" <%=pcf_SelectOption("FICEType","DSP_85")%>>DSP-85</option>
                          <option value="DSP_94" <%=pcf_SelectOption("FICEType","DSP_94")%>>DSP-94</option>
                          <option value="DSP_LICENSE_AGREEMENT" <%=pcf_SelectOption("FICEType","DSP_LICENSE_AGREEMENT")%>>DSP License Argreements</option>
                          <option value="DEA_036" <%=pcf_SelectOption("FICEType","DEA_036")%>>DEA-36</option>
                          <option value="DEA_236" <%=pcf_SelectOption("FICEType","DEA_236")%>>DEA-236</option>
                          <option value="DEA_486" <%=pcf_SelectOption("FICEType","DEA_486")%>>DEA_486</option>
                          <option value="WAREHOUSE_WITHDRAWAL" <%=pcf_SelectOption("FICEType","WAREHOUSE_WITHDRAWAL")%>>Warehouse Withdrawal</option>
                          <option value="FROM_FOREIGN_TRADE_ZONE" <%=pcf_SelectOption("FICEType","FROM_FOREIGN_TRADE_ZONE")%>>T&amp;E from a Foreign Trade</option>
                          </select>
                          <%pcs_RequiredImageTag "NAFTANetCostMethod", false %>
                        </td>
                      </tr>  
                      <tr>
                        <td align="right" valign="top"><b>License/Permit Number:</b></td>
                        <td align="left">
                          <input name="FICENumber" type="text" id="FICENumber" value="<%=pcf_FillFormField("FICENumber", false)%>">
                          <%pcs_RequiredImageTag "FICENumber", false%>
                        </td>
                      </tr>  
                      <tr>
                        <td align="right" valign="top"><b>Expiration Date:</b></td>
                        <td align="left">
                          <input name="FICEExpirationDate" type="text" id="FICEExpirationDate" value="<%=pcf_FillFormField("FICEExpirationDate", false)%>">
                          <%pcs_RequiredImageTag "FICEExpirationDate", false%>
                          <span>Format: 2005-03-12</span>
                        </td>
                      </tr>   
                      <tr>
                        <td align="right" valign="top"><b>Entry Number:</b></td>
                        <td align="left">
                          <input name="FICEEntryNumber" type="text" id="FICEEntryNumber" value="<%=pcf_FillFormField("FICEEntryNumber", false)%>">
                          <%pcs_RequiredImageTag "FICEEntryNumber", false%>
                        </td>
                      </tr> 
                      <tr>
                        <td align="right" valign="top"><b>Foreign Trade Zone Code:</b></td>
                        <td align="left">
                          <input name="FICETradeZoneCode" type="text" id="FICETradeZoneCode" value="<%=pcf_FillFormField("FICETradeZoneCode", false)%>">
                          <%pcs_RequiredImageTag "FICETradeZoneCode", false%>
                        </td>
                      </tr>       
                    </Table>
                  </div>
                </td>
                
								<tr>
								  <td colspan="2" class="pcCPspacer"></td>
								</tr>
              <th colspan="2">
								<script type=text/javascript>
									function jfITAROption() {
										var selectValDom = document.forms['form1'];
										if (selectValDom.bITAROption.checked == true) {
											document.getElementById('bITARContainer').style.display='';
										}else{
											document.getElementById('bITARContainer').style.display='none';
										}
									}
								</script>
									<%
									if Session("pcAdminbITAROption")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfITAROption();" name="bITAROption" id="bITAROption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bITAROption", "1")%>>&nbsp;
                  
                  International Traffic in Arms Regulations (ITAR)
                </th>
            </tr>
						<tr>
            	<td colspan="2">
                <div id="bITARContainer" <%=pcv_strDisplayStyle%>>
                  <Table>
                  	<tr>
                      <td align="right" valign="top"><b>License of Exemption Number:</b></td>
                      <td align="left">
												<input name="ITARNumber" type="text" id="ITARNumber" value="<%=pcf_FillFormField("ITARNumber", false)%>">
												<%pcs_RequiredImageTag "ITARNumber", false%>
                      </td>
                    </tr>            
                	</Table>
                </div>
              </td>
                
								<tr>
								  <td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">
									<script type=text/javascript>
									function jfISOCustomsID(){
									var selectValDom = document.forms['form1'];
									if (selectValDom.bISOCustomsID.checked == true) {
									document.getElementById('ISOCustomsID').style.display='';
									}else{
									document.getElementById('ISOCustomsID').style.display='none';
									}
									}
									</script>
									<%
									if Session("pcAdminbISOCustomsID")="1" then
										pcv_strDisplayStyle="style=""display:block"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfISOCustomsID();" name="bISOCustomsID" id="bISOCustomsID" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bISOCustomsID", "1")%>>
									Recipient Customs ID</th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="ISOCustomsID" <%=pcv_strDisplayStyle%>>
										<Table>
											<tr>
											<td align="right" valign="top"><b>Id Type:</b></td>
											<td align="left">
										<select name="RCIdType" id="RCIdType">
											<option value="COMPANY" <%=pcf_SelectOption("RCIdType","COMPANY")%>>Company</option>
											<option value="INDIVIDUAL" <%=pcf_SelectOption("RCIdType","INDIVIDUAL")%>>Individual</option>
											<option value="PASSPORT" <%=pcf_SelectOption("RCIdType","PASSPORT")%>>Passport</option>
										</select>
										<%pcs_RequiredImageTag "RCIdType", false%>
											</td>
										</tr>
										<tr>
											<td align="right" valign="top"><b>Id Number:</b></td>
											<td align="left">
												<input name="RCIdValue" type="text" id="RCIdValue" value="<%=pcf_FillFormField("RCIdValue", false)%>">
												<%pcs_RequiredImageTag "RCIdValue", false%>


							</td>
						</tr>
					</table>
					</div>
				</td>
			</tr>
		</table>
		</div>
        </td>
      </tr>
      <tr>
      	<td colspan="2" class="pcCPspacer"></td>
      </tr> 
			<tr>
				<td width="50%" class="pcCPshipping"><span class="titleShip">Freight Settings</span></td>
				<td class="pcCPshipping" align="right">
					<i>(Check box to view)</i>&nbsp;
					<script type=text/javascript>
						function jfFreight(){

							var selectValDom = document.forms['form1'];
							if (selectValDom.bFreight.checked == true) {
								document.getElementById('Freight').style.display='';
							} else {
								document.getElementById('Freight').style.display='none';
							}
						}
						</script>
						<%
						if Session("pcAdminbFreight")="1" then
							pcv_strDisplayStyle="style=""display:block"""
						else
							pcv_strDisplayStyle="style=""display:none"""
						end if
						%>
						<input onClick="jfFreight();" name="bFreight" id="bFreight" type="checkbox" class="clearBorder" value="true" <%=pcf_CheckOption("bFreight", "1")%>>
					</td>
				</tr>
        <tr>
        	<td colspan="2">
          	<div id="Freight" <%= pcv_strDisplayStyle %>>
            	<table width="100%">
              	<tr>
                	<th colspan="2">Freight Settings</th>
                </tr>
                <tr>
                  <td colspan="2"></td>
                </tr>
								<%
									If Session("pcAdminFreightAccountNumber") & "" = "" Then
										Session("pcAdminFreightAccountNumber") = pcv_strShipmentAccountNumber
									End If
								%>
								<tr>
									<td align="right" valign="top" style="width: 25%"><b>Account Number:</b></td>
									<td align="left">
										<input name="FreightAccountNumber" type="text" id="FreightAccountNumber" value="<%=pcf_FillFormField("FreightAccountNumber", false)%>">
										<%pcs_RequiredImageTag "FreightAccountNumber", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Shipment Role:</b></td>
									<td align="left">
                    <select name="FreightShipmentRoleType" id="FreightShipmentRoleType">
                      <option value=""></option>
                      <option value="SHIPPER" <%=pcf_SelectOption("FreightShipmentRoleType","SHIPPER")%>>Shipper</option>
                      <option value="CONSIGNEE" <%=pcf_SelectOption("FreightShipmentRoleType","CONSIGNEE")%>>Consignee</option>
                    </select>
										<%pcs_RequiredImageTag "FreightShipmentRoleType", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Total Handling Units:</b></td>
									<td align="left">
										<input name="FreightTotalHandlingUnits" type="text" id="FreightTotalHandlingUnits" value="<%=pcf_FillFormField("FreightTotalHandlingUnits", false)%>">
										<%pcs_RequiredImageTag "FreightTotalHandlingUnits", true %>
									</td>
								</tr>  
								<tr>
									<td align="right" valign="top"><b>Collect Terms Type:</b></td>
									<td align="left">
                    <select name="FreightCollectTermsType" id="FreightCollectTermsType">
                      <option value="">Select Option</option>                      
                      <option value="STANDARD" <%=pcf_SelectOption("FreightCollectTermsType","STANDARD")%>>Standard</option>
                      <option value="NON_RECOURSE_SHIPPER_SIGNED" <%=pcf_SelectOption("FreightCollectTermsType","NON_RECOURSE_SHIPPER_SIGNED")%>>Non Recourse Shipper Signed</option>
                    </select>
										<%pcs_RequiredImageTag "FreightCollectTermsType", false %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Client Discount:</b></td>
									<td align="left">
										<input name="FreightClientDiscount" type="text" id="FreightClientDiscount" value="<%=pcf_FillFormField("FreightClientDiscount", false)%>"> %
										<%pcs_RequiredImageTag "FreightClientDiscount", false %>
									</td>
								</tr>   
								<tr>
									<td align="right" valign="top"><b>Pallet Weight:</b></td>
									<td align="left">
                    <% FedEx_WeightControl "FreightPalletWeightValue", "FreightPalletWeightUnits", false %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top">
										<b>Shipment Dimensions:</b><br />
										<span class="pcSmallText">Length x Width x Height</span>
									</td>
									<td align="left">
                    <% FedEx_DimensionsControl "FreightShipmentDimensionsLength", "FreightShipmentDimensionsWidth", "FreightShipmentDimensionsHeight", "FreightShipmentDimensionsUnits", false %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Comment:</b></td>
									<td align="left">
										<input name="FreightShipmentComment" type="text" id="FreightShipmentComment" value="<%=pcf_FillFormField("FreightShipmentComment", false)%>">
										<%pcs_RequiredImageTag "FreightShipmentComment", false %>
									</td>
								</tr>
                <tr>
                  <td colspan="2"></td>
                </tr> 
              	<tr>
                	<th colspan="2">Billing Contact &amp; Address</th>
                </tr>
								<tr>
									<td align="right" valign="top"><b>Person Name:</b></td>
									<td align="left">
										<input name="FreightContactPersonName" type="text" id="FreightContactPersonName" value="<%=pcf_FillFormField("FreightContactPersonName", false)%>">
										<%pcs_RequiredImageTag "FreightContactPersonName", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Company Name:</b></td>
									<td align="left">
										<input name="FreightContactCompanyName" type="text" id="FreightContactCompanyName" value="<%=pcf_FillFormField("FreightContactCompanyName", false)%>">
										<%pcs_RequiredImageTag "FreightContactCompanyName", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Pager Number:</b></td>
									<td align="left">
										<input name="FreightContactPagerNumber" type="text" id="FreightContactPagerNumber" value="<%=pcf_FillFormField("FreightContactPagerNumber", false)%>">
										<%pcs_RequiredImageTag "FreightContactPagerNumber", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Street Address:</b></td>
									<td align="left">
										<input name="FreightContactStreetLines" type="text" id="FreightContactStreetLines" value="<%=pcf_FillFormField("FreightContactStreetLines", false)%>">
										<%pcs_RequiredImageTag "FreightContactStreetLines", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>City:</b></td>
									<td align="left">
										<input name="FreightContactCity" type="text" id="FreightContactCity" value="<%=pcf_FillFormField("FreightContactCity", false)%>">
										<%pcs_RequiredImageTag "FreightContactCity", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>State Code:</b></td>
									<td align="left">
										<input name="FreightContactStateCode" type="text" id="FreightContactStateCode" value="<%=pcf_FillFormField("FreightContactStateCode", false)%>">
										<%pcs_RequiredImageTag "FreightContactStateCode", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Postal Code:</b></td>
									<td align="left">
										<input name="FreightContactPostalCode" type="text" id="FreightContactPostalCode" value="<%=pcf_FillFormField("FreightContactPostalCode", false)%>">
										<%pcs_RequiredImageTag "FreightContactPostalCode", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Country Code:</b></td>
									<td align="left">
										<input name="FreightContactCountryCode" type="text" id="FreightContactCountryCode" value="<%=pcf_FillFormField("FreightContactCountryCode", false)%>">
										<%pcs_RequiredImageTag "FreightContactCountryCode", true %>
									</td>
								</tr>
                <tr>
                  <td colspan="2"></td>
                </tr>
              	<tr>
                	<th colspan="2">Declared Value</th>
                </tr>
								<tr>
									<td align="right" valign="top"><b>Per-Unit Amount:</b></td>
									<td align="left">
										<% 
											FedEx_CurrencyControl "FreightDVAmount", "FreightDVCurrency", false
										%>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b># Units:</b></td>
									<td align="left">
										<input name="FreightDVUnits" type="text" id="FreightDVUnits" value="<%=pcf_FillFormField("FreightDVUnits", false)%>">
										<%pcs_RequiredImageTag "FreightDVUnits", false %>
									</td>
								</tr>
                <tr>
                  <td colspan="2"></td>
                </tr>
              	<tr>
                	<th colspan="2">Liability Coverage</th>
                </tr>
								<tr>
									<td align="right" valign="top"><b>Coverage Type:</b></td>
									<td align="left">
                    <select name="FreightLCType" id="FreightLCType">
                      <option value="">Select Option</option>                      
                      <option value="NEW" <%=pcf_SelectOption("FreightLCType","NEW")%>>New</option>
                      <option value="USED_OR_RECONDITIONED" <%=pcf_SelectOption("FreightLCType","USED_OR_RECONDITIONED")%>>Used or Reconditioned</option>
                    </select>
										<%pcs_RequiredImageTag "FreightLCType", false %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Coverage Amount:</b></td>
									<td align="left">
										<%
											FedEx_CurrencyControl "FreightLCAmount", "FreightLCCurrency", false
										%>
									</td>
								</tr>     
                <tr>
                  <td colspan="2"></td>
                </tr> 
              	<tr>
                	<th colspan="2">Commodities</th>
                </tr>
								<tr>
									<td align="right" valign="top"><b>Packaging:</b></td>
									<td align="left">
          
                    <select name="FreightLIPackaging" id="FreightLIPackaging">
                      <option value="">Select Option</option>                      
                      <option value="BAG" <%=pcf_SelectOption("FreightLIPackaging","BAG")%>>Bag</option>
                      <option value="BARREL" <%=pcf_SelectOption("FreightLIPackaging","BARREL")%>>Barrel</option>
                      <option value="BASKET" <%=pcf_SelectOption("FreightLIPackaging","BASKET")%>>Basket</option>
                      <option value="BOX" <%=pcf_SelectOption("FreightLIPackaging","BOX")%>>Box</option>
                      <option value="BUCKET" <%=pcf_SelectOption("FreightLIPackaging","BUCKET")%>>Bucket</option>
                      <option value="BUNDLE" <%=pcf_SelectOption("FreightLIPackaging","BUNDLE")%>>Bundle</option>
                      <option value="CARTON" <%=pcf_SelectOption("FreightLIPackaging","CARTON")%>>Carton</option>
                      <option value="CASE" <%=pcf_SelectOption("FreightLIPackaging","CASE")%>>Case</option>
                      <option value="CONTAINER" <%=pcf_SelectOption("FreightLIPackaging","CONTAINER")%>>Container</option>
                      <option value="CRATE" <%=pcf_SelectOption("FreightLIPackaging","CRATE")%>>Create</option>
                      <option value="CYLINDER" <%=pcf_SelectOption("FreightLIPackaging","CYLINDER")%>>Cylinder</option>
                      <option value="DRUM" <%=pcf_SelectOption("FreightLIPackaging","DRUM")%>>Drum</option>
                      <option value="ENVELOPE" <%=pcf_SelectOption("FreightLIPackaging","ENVELOPE")%>>Envelope</option>
                      <option value="HAMPER" <%=pcf_SelectOption("FreightLIPackaging","HAMPER")%>>Hamper</option>
                      <option value="OTHER" <%=pcf_SelectOption("FreightLIPackaging","OTHER")%>>Other</option>
                      <option value="PALLET" <%=pcf_SelectOption("FreightLIPackaging","PALLET")%>>Pallet</option>
                      <option value="PIECE" <%=pcf_SelectOption("FreightLIPackaging","PIECE")%>>Piece</option>
                      <option value="REEL" <%=pcf_SelectOption("FreightLIPackaging","REEL")%>>Reel</option>
                      <option value="ROLL" <%=pcf_SelectOption("FreightLIPackaging","ROLL")%>>Roll</option>
                      <option value="SKID" <%=pcf_SelectOption("FreightLIPackaging","SKID")%>>Skid</option>
                      <option value="TANK" <%=pcf_SelectOption("FreightLIPackaging","TANK")%>>Tank</option>
                      <option value="TUBE" <%=pcf_SelectOption("FreightLIPackaging","TUBE")%>>Tube</option>
                    </select>
										<%pcs_RequiredImageTag "FreightLIPackaging", isRequiredFreightLIPackaging %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Pieces:</b></td>
									<td align="left">
										<input name="FreightLIPieces" type="text" id="FreightLIPieces" value="<%=pcf_FillFormField("FreightLIPieces", false)%>">
										<%pcs_RequiredImageTag "FreightLIPieces", isRequiredFreightLIPieces %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Description:</b></td>
									<td align="left">
										<input name="FreightLIDescription" type="text" id="FreightLIDescription" value="<%=pcf_FillFormField("FreightLIDescription", false)%>">
										<%pcs_RequiredImageTag "FreightLIDescription", true %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Weight:</b></td>
									<td align="left">
										<%
											FedEx_WeightControl "FreightLIWeightValue", "FreightLIWeightUnits", true
										%>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Dimensions:</b></td>
									<td align="left">
										<%
											FedEx_DimensionsControl "FreightLIDimensionsLength", "FreightLIDimensionsWidth", "FreightLIDimensionsHeight", "FreightLIDimensionsUnits", false
										%>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Freight Class:</b></td>
									<td align="left">
                    <select name="FreightLIClass" id="FreightLIClass">
                      <option value="">Select Option</option>                      
                      <option value="CLASS_050" <%=pcf_SelectOption("FreightLIClass","CLASS_050")%>>Class 50</option>
                      <option value="CLASS_055" <%=pcf_SelectOption("FreightLIClass","CLASS_055")%>>Class 55</option>
                      <option value="CLASS_060" <%=pcf_SelectOption("FreightLIClass","CLASS_060")%>>Class 60</option>
                      <option value="CLASS_065" <%=pcf_SelectOption("FreightLIClass","CLASS_065")%>>Class 65</option>
                      <option value="CLASS_070" <%=pcf_SelectOption("FreightLIClass","CLASS_070")%>>Class 70</option>
                      <option value="CLASS_077_5" <%=pcf_SelectOption("FreightLIClass","CLASS_077_5")%>>Class 77.5</option>
                      <option value="CLASS_085" <%=pcf_SelectOption("FreightLIClass","CLASS_085")%>>Class 85</option>
                      <option value="CLASS_092_5" <%=pcf_SelectOption("FreightLIClass","CLASS_092_5")%>>Class 92.5</option>
                      <option value="CLASS_100" <%=pcf_SelectOption("FreightLIClass","CLASS_100")%>>Class 100</option>
                      <option value="CLASS_110" <%=pcf_SelectOption("FreightLIClass","CLASS_110")%>>Class 110</option>
                      <option value="CLASS_125" <%=pcf_SelectOption("FreightLIClass","CLASS_125")%>>Class 125</option>
                      <option value="CLASS_150" <%=pcf_SelectOption("FreightLIClass","CLASS_150")%>>Class 150</option>
                      <option value="CLASS_175" <%=pcf_SelectOption("FreightLIClass","CLASS_175")%>>Class 175</option>
                      <option value="CLASS_200" <%=pcf_SelectOption("FreightLIClass","CLASS_200")%>>Class 200</option>
                      <option value="CLASS_250" <%=pcf_SelectOption("FreightLIClass","CLASS_250")%>>Class 250</option>
                      <option value="CLASS_300" <%=pcf_SelectOption("FreightLIClass","CLASS_300")%>>Class 300</option>
                      <option value="CLASS_400" <%=pcf_SelectOption("FreightLIClass","CLASS_400")%>>Class 400</option>
                      <option value="CLASS_500" <%=pcf_SelectOption("FreightLIClass","CLASS_500")%>>Class 500</option>
                    </select>
										<%pcs_RequiredImageTag "FreightLIClass", isRequiredFreightLIClass %>
									</td>
								</tr>
                <tr>
                  <td align="right" valign="top"><b>Class Provided by Customer:</b></td>
                  <td align="left">
                    <input name="FreightLIClassProvided" type="radio" id="FreightLIClassProvided" value="true" <%=pcf_CheckOption("FreightLIClassProvided", "true")%>>Yes
                    <input name="FreightLIClassProvided" type="radio" id="FreightLIClassProvided" value="false" <%=pcf_CheckOption("FreightLIClassProvided", "false")%>>No
                  </td>
                </tr>
								<tr>
									<td align="right" valign="top"><b>Handling Units:</b></td>
									<td align="left">
										<input name="FreightLIHandlingUnits" type="text" id="FreightLIHandlingUnits" value="<%=pcf_FillFormField("FreightLIHandlingUnits", false)%>">
										<%pcs_RequiredImageTag "FreightLIHandlingUnits", isRequiredFreightLIHandlingUnits %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Purchase Order Number:</b></td>
									<td align="left">
										<input name="FreightLIPONumber" type="text" id="FreightLIPONumber" value="<%=pcf_FillFormField("FreightLIPONumber", false)%>">
										<%pcs_RequiredImageTag "FreightLIPONumber", isRequiredFreightLIPONumber %>
									</td>
								</tr>
              </table>
            </div>
          </td>
        </tr>
	    </table>
    </div>


		<!--
		//////////////////////////////////////////////////////////////////////////////////////////////
		// SHIPPER
		//////////////////////////////////////////////////////////////////////////////////////////////
		-->
		<div id="tab2" class="panes">
		<table class="pcCPcontent">
			<tr>
			<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			<th colspan="2">Contact Details</th>
			</tr>
			<tr>
			<td width="25%" align="right"><p>Contact Name:</p></td>
			<td width="75%" align="left"><p>
			<input name="OriginPersonName" type="text" id="OriginPersonName" value="<%=pcf_FillFormField("OriginPersonName", true)%>">
			<%pcs_RequiredImageTag "OriginPersonName", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Company Name:</p></td>
			<td align="left"><p>
			<input name="OriginCompanyName" type="text" id="OriginCompanyName" value="<%=pcf_FillFormField("OriginCompanyName", false)%>">
			<%pcs_RequiredImageTag "OriginCompanyName", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Department:</p></td>
			<td align="left"><p>
			<input name="OriginDepartment" type="text" id="OriginDepartment" value="<%=pcf_FillFormField("OriginDepartment", false)%>">
						<%pcs_RequiredImageTag "OriginDepartment", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrOriginPhoneNumber"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Phone Number.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Phone Number:</p></td>
			<td align="left"><p>
			<input name="OriginPhoneNumber" type="text" id="OriginPhoneNumber" value="<%=pcf_FillFormField("OriginPhoneNumber", true)%>">
			<%pcs_RequiredImageTag "OriginPhoneNumber", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Pager Number:</p></td>
			<td align="left"><p>
			<input name="OriginPagerNumber" type="text" id="OriginPagerNumber" value="<%=pcf_FillFormField("OriginPagerNumber", false)%>">
			<%pcs_RequiredImageTag "OriginPagerNumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Fax Number:</p></td>
			<td align="left"><p>
			<input name="OriginFaxNumber" type="text" id="OriginFaxNumber" value="<%=pcf_FillFormField("OriginFaxNumber", false)%>">
			<%pcs_RequiredImageTag "OriginFaxNumber", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrOriginEmailAddress"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Email Address.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Email Address:</p></td>
			<td align="left"><p>
			<input name="OriginEmailAddress" type="text" id="OriginEmailAddress" value="<%=pcf_FillFormField("OriginEmailAddress", true)%>">
						<%pcs_RequiredImageTag "OriginEmailAddress", true%></p>
			</td>
			</tr>
			<tr>
			<th colspan="2">Location Details</th>
			</tr>

			<%
			'///////////////////////////////////////////////////////////
			'// START: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			'
			pcv_isStateOrProvinceCodeRequired = isRequiredState '// determines if validation is performed (true or false)
			pcv_isProvinceCodeRequired = isRequiredProvince '// determines if validation is performed (true or false)
			pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)

			'// #3 Additional Required Info
			pcv_strTargetForm = "form1" '// Name of Form
			pcv_strCountryBox = "OriginCountryCode" '// Name of Country Dropdown
			pcv_strTargetBox = "OriginStateOrProvinceCode" '// Name of State Dropdown
			pcv_strProvinceBox =  "OriginProvinceCode" '// Name of Province Field

			'// Set local Country to Session
			if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strCountryBox))=True then
				Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
			end if

			'// Set local State to Session
			if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strTargetBox))=True then
				Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminOriginStateOrProvinceCode")
			end if

			'// Set local Province to Session
			if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strProvinceBox))=True then
				Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  Session("pcAdminOriginStateOrProvinceCode")
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
			%>

			<tr>
			<td align="right"><p>Address Line 1:</p></td>
			<td align="left"><p>
			<input name="OriginLine1" type="text" id="OriginLine1" value="<%=pcf_FillFormField("OriginLine1", true)%>">
			<%pcs_RequiredImageTag "OriginLine1", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Address Line 2:</p></td>
			<td align="left"><p>
			<input name="OriginLine2" type="text" id="OriginLine2" value="<%=pcf_FillFormField("OriginLine2", false)%>">
			<%pcs_RequiredImageTag "OriginLine2", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>City:</p></td>
			<td align="left"><p>
			<input name="OriginCity" type="text" id="OriginCity" value="<%=pcf_FillFormField("OriginCity", true)%>">
						<%pcs_RequiredImageTag "OriginCity", true%></p>
			</td>
			</tr>

			<%
			'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
			pcs_StateProvince
			%>

			<tr>
			<td align="right"><p>Postal Code:</p></td>
			<td align="left"><p>
			<input name="OriginPostalCode" type="text" id="OriginPostalCode" value="<%=pcf_FillFormField("OriginPostalCode", true)%>">
			<%pcs_RequiredImageTag "OriginPostalCode", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"></td>
			<td align="left">
			</td>
			</tr>
		</table>

		</div>

		<!--
				//////////////////////////////////////////////////////////////////////////////////////////////
				// RECIPIENT
				//////////////////////////////////////////////////////////////////////////////////////////////
				-->
		<div id="tab3" class="panes">
		<table class="pcCPcontent">
			<tr>
			<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			<th colspan="2">Contact Details</th>
			</tr>
			<tr>
			<td width="25%" align="right"><p>Contact Name:</p></td>
			<td width="75%" align="left"><p>
			<input name="RecipPersonName" type="text" id="RecipPersonName" value="<%=pcf_FillFormField("RecipPersonName", true)%>">
			<%pcs_RequiredImageTag "RecipPersonName", true%></p>
			</td>
			</tr>
			<tr>

			<td align="right"><p>Company Name:</p></td>
			<td align="left"><p>
			<input name="RecipCompanyName" type="text" id="RecipCompanyName" value="<%=pcf_FillFormField("RecipCompanyName", false)%>">
			<%pcs_RequiredImageTag "RecipCompanyName", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Department:</p></td>
			<td align="left"><p>
			<input name="RecipDepartment" type="text" id="RecipDepartment" value="<%=pcf_FillFormField("RecipDepartment", false)%>">
						<%pcs_RequiredImageTag "RecipDepartment", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrRecipPhoneNumber"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Phone Number.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Phone Number:</p></td>
			<td align="left"><p>
			<input name="RecipPhoneNumber" type="text" id="RecipPhoneNumber" value="<%=pcf_FillFormField("RecipPhoneNumber", true)%>">
			<%pcs_RequiredImageTag "RecipPhoneNumber", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Pager Number:</p></td>
			<td align="left"><p>
			<input name="RecipPagerNumber" type="text" id="RecipPagerNumber" value="<%=pcf_FillFormField("RecipPagerNumber", false)%>">
			<%pcs_RequiredImageTag "RecipPagerNumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Fax Number:</p></td>
			<td align="left"><p>
			<input name="RecipFaxNumber" type="text" id="RecipFaxNumber" value="<%=pcf_FillFormField("RecipFaxNumber", false)%>">
			<%pcs_RequiredImageTag "RecipFaxNumber", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrRecipEmailAddress"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Email Address.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Email Address:</p></td>
			<td align="left"><p>
			<input name="RecipEmailAddress" type="text" id="RecipEmailAddress" value="<%=pcf_FillFormField("RecipEmailAddress", false)%>">
						<%pcs_RequiredImageTag "RecipEmailAddress", false%></p>
			</td>
			</tr>
			<tr>
			<th colspan="2">Location Details</th>
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
			pcv_isStateOrProvinceCodeRequired = isRequiredState2 '// determines if validation is performed (true or false)
			pcv_isProvinceCodeRequired = isRequiredProvince2 '// determines if validation is performed (true or false)
			pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)

			'// #3 Additional Required Info
			pcv_strTargetForm = "form1" '// Name of Form
			pcv_strCountryBox = "RecipCountryCode" '// Name of Country Dropdown
			pcv_strTargetBox = "RecipStateOrProvinceCode" '// Name of State Dropdown
			pcv_strProvinceBox =  "RecipProvinceCode" '// Name of Province Field

			'// Set local Country to Session
			if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strCountryBox))=True then
				Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
			end if

			'// Set local State to Session
			if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strTargetBox))=True then
				Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminRecipStateOrProvinceCode")
			end if

			'// Set local Province to Session
			if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strProvinceBox))=True then
				Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  Session("pcAdminRecipStateOrProvinceCode")
			end if

			'// Declare the instance number if greater than 1
			pcv_strFormInstance = "2"
			'///////////////////////////////////////////////////////////
			'// END: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			%>

			<%
			'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
			pcs_CountryDropdown
			%>

			<tr>
			<td align="right"><p>Address Line 1:</p></td>
			<td align="left"><p>
			<input name="RecipLine1" type="text" id="RecipLine1" value="<%=pcf_FillFormField("RecipLine1", true)%>">
			<%pcs_RequiredImageTag "RecipLine1", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Address Line 2:</p></td>
			<td align="left"><p>
			<input name="RecipLine2" type="text" id="RecipLine2" value="<%=pcf_FillFormField("RecipLine2", false)%>">
			<%pcs_RequiredImageTag "RecipLine2", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>City:</p></td>
			<td align="left"><p>
			<input name="RecipCity" type="text" id="RecipCity" value="<%=pcf_FillFormField("RecipCity", true)%>">
						<%pcs_RequiredImageTag "RecipCity", true%></p>
			</td>
			</tr>

			<%
			'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
			pcs_StateProvince
			%>

			<tr>
			<td align="right"><p>Postal Code:</p></td>
			<td align="left"><p>
			<input name="RecipPostalCode" type="text" id="RecipPostalCode" value="<%=pcf_FillFormField("RecipPostalCode", isRequiredRecipPostal)%>">
			<%pcs_RequiredImageTag "RecipPostalCode", isRequiredRecipPostal %></p>
			</td>
			</tr>

			<tr>
			<td align="right"><p>Customer Reference:</p></td>
			<td align="left"><p>
			<input name="CustomerReference" type="text" id="CustomerReference" value="<%=pcf_FillFormField("CustomerReference", true)%>">
						<%pcs_RequiredImageTag "CustomerReference", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Customer PO Number:</p></td>
			<td align="left"><p>
			<input name="CustomerPONumber" type="text" id="CustomerPONumber" value="<%=pcf_FillFormField("CustomerPONumber", false)%>">
						<%pcs_RequiredImageTag "CustomerPONumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Customer Invoice Number:</p></td>
			<td align="left"><p>
			<input name="CustomerInvoiceNumber" type="text" id="CustomerInvoiceNumber" value="<%=pcf_FillFormField("CustomerInvoiceNumber", false)%>">
						<%pcs_RequiredImageTag "CustomerInvoiceNumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right">    </td>
			<td>
			<input type="checkbox" name="ResidentialDelivery" value="true" class="clearBorder" <%=pcf_CheckOption("ResidentialDelivery", "true")%>>
			<strong>This is a Residential Delivery</strong>
			</td>
			</tr>

			<tr>
			<td align="right">    </td>
			<td>
			<input type="checkbox" name="InsideDelivery" value="1" class="clearBorder" <%=pcf_CheckOption("InsideDelivery", "1")%>>
			<strong>Inside Delivery</strong>
			</td>
			</tr>
			<tr>
			<td align="right">    </td>
			<td>
			<input type="checkbox" name="InsidePickup" value="1" class="clearBorder" <%=pcf_CheckOption("InsidePickup", "1")%>>
			<strong>Inside Pickup</strong>
			</td>
			</tr>

			<tr>
			<td align="right"></td>
			<td align="left">
			</td>
			</tr>
		</table>

		</div>


		<!--
				//////////////////////////////////////////////////////////////////////////////////////////////
				// SHIPPING ALERTS
				//////////////////////////////////////////////////////////////////////////////////////////////
				-->
		<div id="tab4" class="panes">
		<table class="pcCPcontent">
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="2"><span class="title">FedEx ShipAlert<sup>&reg;</sup> Notifications</span></td>
			</tr>
			<tr>
				<td colspan="2">
					<p><strong>Shipment notification</strong> &ndash; Automatically send an email message indicating the shipment is on the way.<br>
					  <strong>Delivery notification</strong> &ndash; receive a delivery notification for an express package. <br>
					  <strong>Exception notification</strong> - receive an email notifcation for delivery exceptions.<br>
					  <strong>Email address</strong> &ndash; Enter the email addresses to receive the notifications.                    </p></td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td align="left" colspan="2">Notification Format:&nbsp;&nbsp;<select name="ShipperNotificationFormat" id="select">
				<option value="HTML" <%=pcf_SelectOption("ShipperNotificationFormat","HTML")%>>HTML Format</option>
				<option value="TEXT" <%=pcf_SelectOption("ShipperNotificationFormat","TEXT")%>>Plain Text Format</option>
				<option value="WIRELESS" <%=pcf_SelectOption("ShipperNotificationFormat","WIRELESS")%>>Formatted for Wireless Device</option>
			</select></td>
			</tr>
			<tr>
				<td align="left" style="width: 30%"><input type="checkbox" name="NotificationShipperEnabled" value="1" <%=pcf_CheckOption("NotificationShipperEnabled", "1")%>> <strong>Shipper Notification:</strong></td>
				<td align="left">
        	<input name="NotificationShipperEmail" type="text" id="NotificationShipperEmail" value="<%=pcf_FillFormField("NotificationShipperEmail", false)%>" size="40" />
          <%pcs_RequiredImageTag "NotificationShipperEnabled", false%>
        </td>
			</tr>
			<tr>
				<td align="left"><input type="checkbox" name="NotificationRecipientEnabled" value="1" <%=pcf_CheckOption("NotificationRecipientEnabled", "1")%>> <strong>Recipient Notification:</strong></td>
				<td align="left">
        	<input name="NotificationRecipientEmail" type="text" id="NotificationRecipientEmail" value="<%=pcf_FillFormField("NotificationRecipientEmail", false)%>" size="40" />
          <%pcs_RequiredImageTag "NotificationRecipientEmail", false%>
        </td>
			</tr>
			<tr>
				<td align="left"><input type="checkbox" name="NotificationThirdPartyEnabled" value="1" <%=pcf_CheckOption("NotificationThirdPartyEnabled", "1")%>> <strong>3rd Party Notification:</strong></td>
				<td align="left">
        	<input name="NotificationThirdPartyEmail" type="text" id="NotificationThirdPartyEmail" value="<%=pcf_FillFormField("NotificationThirdPartyEmail", false)%>" size="40" />
          <%pcs_RequiredImageTag "NotificationThirdPartyEmail", false%>
        </td>
			</tr>
			<tr>
				<td align="left"><input type="checkbox" name="NotificationBrokerEnabled" value="1" <%=pcf_CheckOption("NotificationBrokerEnabled", "1")%>> <strong>Broker Notification:</strong></td>
				<td align="left">
        	<input name="NotificationBrokerEmail" type="text" id="NotificationBrokerEmail" value="<%=pcf_FillFormField("NotificationBrokerEmail", false)%>" size="40" />
          <%pcs_RequiredImageTag "NotificationBrokerEmail", false%>
        </td>
			</tr>
			<tr>
				<td align="left"><input type="checkbox" name="NotificationOtherEnabled" value="1" <%=pcf_CheckOption("NotificationOtherEnabled", "1")%>>  <strong>Additional Notification:</strong></td>
				<td align="left">
        	<input name="NotificationOtherEmail" type="text" id="NotificationOtherEmail" value="<%=pcf_FillFormField("NotificationOtherEmail", false)%>" size="40" />
          <%pcs_RequiredImageTag "NotificationOtherEmail", false%>
        </td>
			</tr>
		</table>

		</div>

		<!--
		//////////////////////////////////////////////////////////////////////////////////////////////
		// PACKAGES
		//////////////////////////////////////////////////////////////////////////////////////////////
		-->
		<%
		for k=1 to pcPackageCount
			%>
			<div id="tab<%=4+int(k)%>" class="panes">
				<table class="pcCPcontent">
					<tr>
					<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<%
					'// If the tab was processed, skip it.
					if pcLocalArray(k-1) <> "shipped" then
					%>
					<tr>
					<td colspan="2"><span class="title">Package <%=k%> Information</span></td>
					</tr>
					<tr>
					<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="2">
						  <script type=text/javascript>
								function FaxSelected<%=k%>() {
								  var selectValDom = document.forms['form1'];
								  if (selectValDom.FaxLetter<%=k%>.checked == true) {
								    document.getElementById('FaxTable<%=k%>').style.display='';
								  } else {
								    document.getElementById('FaxTable<%=k%>').style.display='none';
								  }
								}
						  </script>
						<%
						if Session("pcAdminFaxLetter"&k)="true" then
							pcv_strDisplayStyle="style=""display:block"""
						else
							pcv_strDisplayStyle="style=""display:none"""
						end if
						%>
						<input onClick="FaxSelected<%=k%>();" name="FaxLetter<%=k%>" id="FaxLetter<%=k%>" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("FaxLetter"&k, "true")%>>
						Click Here to view <b>package contents</b>.

							<table class="pcCPcontent" ID="FaxTable<%=k%>" <%=pcv_strDisplayStyle%>>
								<tr>
									<td colspan="2" valign="top">
										<%
									  xProductDisplayArray = split(Session("pcAdminPrdList"&k),",")
										For pcv_xCounter=0 to (ubound(xProductDisplayArray)-1)
											pcv_intPackageInfo_ID = xProductDisplayArray(pcv_xCounter)
											' GET THE PACKAGE CONTENTS
											' >>> Tables: products, ProductsOrdered
											query = 		"SELECT ProductsOrdered.pcPackageInfo_ID , products.description, products.idProduct, products.OverSizeSpec "
											query = query & "FROM ProductsOrdered "
											query = query & "INNER JOIN products "
											query = query & "ON ProductsOrdered.idProduct = products.idProduct "
											query = query & "WHERE ProductsOrdered.idProductOrdered=" & pcv_intPackageInfo_ID &" "
                      
                      on error resume next

											set rs2=server.CreateObject("ADODB.RecordSet")
											set rs2=conntemp.execute(query)

											if err.number<>0 then
												'// handle admin error
											end if

											if NOT rs2.eof then
                        %><ul style="padding-left: 20px"><%
												Do until rs2.eof
													pcv_strProductDescription = rs2("description")
													pOverSizeSpec=rs2("OverSizeSpec")
													if pOverSizeSpec="" or isNull(pOverSizeSpec) then
														pOverSizeSpec="NO"
													end if
													if pOverSizeSpec<>"NO" then
														pOSArray=split(pOverSizeSpec,"||")
														if ubound(pOSArray)>2 then
															tOS_width=pOSArray(0)
  														tOS_height=pOSArray(1)
															tOS_length=pOSArray(2)
														else
															tOS_width=FEDEXWS_WIDTH
															tOS_height=FEDEXWS_HEIGHT
															tOS_length=FEDEXWS_LENGTH
														end if
													else
														tOS_width=FEDEXWS_WIDTH
														tOS_height=FEDEXWS_HEIGHT
														tOS_length=FEDEXWS_LENGTH
													end if
													'// You only ship one oversized item per package, override dimensions for this tab
													Session("pcAdminLength"&k) = tOS_length
													Session("pcAdminWidth"&k) = tOS_width
													Session("pcAdminHeight"&k) = tOS_height
													%>
													<li><%=pcv_strProductDescription%></li>
													<%
												rs2.movenext
												Loop
                        %></ul><%
											end if
										Next
										%>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
					<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
					<th colspan="2">Settings <%'=k%></th>
					</tr>
					  <script type=text/javascript>
							function setpackagedivs() {
							   var serviceCode = $pc("#Service1").val();
							   $pc(".smartpost").hide();
							   $pc(".expressfreight").hide();
							   $pc(".freight_options").hide();
							   $pc(".homedelivery").hide();
								 
								 switch (serviceCode) {
									case "GROUND_HOME_DELIVERY":
							   		$pc(".homedelivery").show();
										break;
									case "INTERNATIONAL_PRIORITY_FREIGHT":
									case "INTERNATIONAL_ECONOMY_FREIGHT":
									case "FEDEX_1_DAY_FREIGHT":
									case "FEDEX_2_DAY_FREIGHT":
									case "FEDEX_3_DAY_FREIGHT":
		 						   	$pc(".expressfreight").show();
									 	break;
									case "FEDEX_FREIGHT_PRIORITY":
									case "FEDEX_FREIGHT_ECONOMY":
							   		$pc(".freight_options").show();
										break;
									case "SMART_POST":
							   		$pc(".smartpost").show();
										break;
								 }
								 
							}
							</script>
					<tr>
            <td align="right">
					    <strong>Service Type: </strong>
            </td>
            <td>
					    <select name="Service<%=k%>" id="Service<%=k%>" onchange="setpackagedivs(this);">
						    <%
                  query = "SELECT idShipService,serviceCode,serviceDescription FROM shipService WHERE idShipment=" & pcv_FedExShipmentID & " AND serviceActive = -1 ORDER BY idShipService"
                  set rsService = conntemp.execute(query)
                  if not rsService.eof then
                    do while not rsService.eof
                      idShipService = rsService("idShipService")
                      serviceCode = rsService("serviceCode")
                      serviceDescription = rsService("serviceDescription")

                      '// Auto-select the default shipping option
                      If Service = Replace(serviceDescription, "<sup>&reg;</sup>", "") Then
					              If Session("pcAdminService"&k)="" Then
						              Session("pcAdminService"&k)=serviceCode
					              End If
                      End If

                  
                      %>
                        <option value="<%= serviceCode %>" <%=pcf_SelectOption("Service"&k,serviceCode)%>><%= serviceDescription %></option>
                      <%
									    rsService.MoveNext
                    loop
								    rsService.Close
                  end if
                %>
					    </select>

					    <%pcs_RequiredImageTag "Service"&k, true%>
            </td>
          </tr>
          <tr>
            <td colspan="2">
					    <div class="pcCPnotes">
					      When using FedEx packaging, select the
					      packaging type from the drop-down list.
                <br>
					      When using non-FedEx packaging, select &quot;Your
					      Packaging&quot;, and then enter
					      the dimensions manually.
					    </div>
            </td>
          </tr>
          <tr>
            <td colspan="2" class="pcCPspacer"></td>
          </tr>

          <tr>
					  <td align="right" width="20%">
              <strong>Package Type:</strong>
					  </td>
            <td>
					    <%
					    %>
					    <select name="Packaging<%=k%>" id="Packaging<%=k%>">
					    <option value="FEDEX_10KG_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_10KG_BOX")%>>FedEx&reg; 10kg Box</option>
					    <option value="FEDEX_25KG_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_25KG_BOX")%>>FedEx&reg; 25kg Box</option>
					    <option value="FEDEX_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_BOX")%>>FedEx&reg; Box</option>
					    <option value="FEDEX_ENVELOPE" <%=pcf_SelectOption("Packaging"&k,"FEDEX_ENVELOPE")%>>FedEx&reg; Envelope</option>
					    <option value="FEDEX_SMALL_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_SMALL_BOX")%>>FedEx&reg; Small Box</option>
					    <option value="FEDEX_MEDIUM_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_MEDIUM_BOX")%>>FedEx&reg; Medium Box</option>
					    <option value="FEDEX_LARGE_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_LARGE_BOX")%>>FedEx&reg; Large Box</option>
					    <option value="FEDEX_EXTRA_LARGE_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_EXTRA_LARGE_BOX")%>>FedEx&reg; Extra Large Box</option>
					    <option value="FEDEX_PAK" <%=pcf_SelectOption("Packaging"&k,"FEDEX_PAK")%>>FedEx&reg; Pak</option>
					    <option value="FEDEX_TUBE" <%=pcf_SelectOption("Packaging"&k,"FEDEX_TUBE")%>>FedEx&reg; Tube</option>
					    <option value="YOUR_PACKAGING" <%=pcf_SelectOption("Packaging"&k,"YOUR_PACKAGING")%>>Customer Package</option>
					    </select>
					    <%pcs_RequiredImageTag "Packaging"&k, true%>
					
				    </td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>
					    <input type="checkbox" name="ContainerType<%= k %>" value="1" class="clearBorder" <%= pcf_CheckOption("ContainerType"&k, "1") %>>&nbsp;Non-Standard Container
            </td>
					</tr>
          <tr>
            <td colspan="2" class="pcCPspacer"></td>
          </tr>
					<tr>
					  <th colspan="2">Dimensions and Weight</th>
					</tr>
					<tr>
					  <td colspan="2">
					    <strong>Package Dimensions:</strong>
              <div class="pcCPnotes">
                Maximum 274 cm (~108 inches) in length (always the longest side)
                <br />
					      Maximum 330 cm (~130 inches) in length and girth combined. Girth = (2 x height) + (2 x  width)
              </div>
            </td>
          </tr>
          <tr>
            <td align="right">
              <strong>Dimensions:</strong><br />
							<span class="pcSmallText">Length x Width x Height</span>
            </td>
            <td>
              <% FedEx_DimensionsControl "Length" & k, "Width" & k, "Height" & k, "Units" & k, false %>
            </td>
          </tr>
          <tr>
            <td class="pcCPspacer"></td>"
          </tr>
					<tr>
					  <td colspan="2">
					    <strong>Package Weight:</strong>
					    <div class="pcCPnotes">
					      Enter the weight of the package. If there is more than one package in the shipment, enter the weight of the first package or the total shipment weight.
					    </div>
					  </td>
          </tr>
          <tr>
            <td align="right">
               <strong>Weight:</strong>
            </td>
            <td>
					    <%
					      if pcPackageCount=1 AND Session("pcAdminWeight"&k)="" then
						      '// Get weight
						      Session("pcAdminWeight"&k) = intMPackageWeight
					      end if
              %>
              
              <% FedEx_WeightControl "Weight" & k, "WeightUnits" & k, true %>
            </td>
          </tr>
          <tr>
            <td colspan="2" class="pcCPspacer"></td>
          </tr>
					<tr>
					<th colspan="2">Package Value</th>
					</tr>
					<tr>
					  <td align="right">
					    <strong>Declared Value: </strong>
            </td>
            <td>
							<%
								if Session("pcAdmindeclaredvalue"&k)="" then
									Session("pcAdmindeclaredvalue"&k) = 100
								end if

								if Session("pcAdmincurrency"&k)="" then
									Session("pcAdmincurrency"&k) = "USD"
								end if

								FedEx_CurrencyControl "declaredvalue"&k, "currency"&k, true
							%>
					  </td>
					</tr>
					<% if k = 1 then %>
						<tr>
							<td colspan="2">
							<div class="smartpost">
								<Table>
									<tr>
										<th colspan="2">Smart Post Details</th>
									</tr>
									<tr>
										<td align="right" valign="top"><b>Indicia:</b></td>
										<td align="left">
										<select name="SMIndicia" id="SMIndicia">
											<option value="MEDIA_MAIL" <%=pcf_SelectOption("SMIndicia","PARCEL_SELECT")%>>Media Mail</option>
											<option value="PARCEL_SELECT" <%=pcf_SelectOption("SMIndicia","PARCEL_SELECT")%>>Parcel Select</option>
											<option value="PARCEL_RETURN" <%=pcf_SelectOption("SMIndicia","PARCEL_RETURN")%>>Parcel Return</option>
											<option value="PRESORTED_STANDARD" <%=pcf_SelectOption("SMIndicia","PRESORTED_STANDARD")%>>Presorted Standard</option>
											<option value="PRESORTED_BOUND_PRINTED_MATTER" <%=pcf_SelectOption("SMIndicia","PRESORTED_BOUND_PRINTED_MATTER")%>>Presorted Bound Printed Matter</option>
										  </select>
											<%pcs_RequiredImageTag "SMIndicia", true%>
										</td>
									</tr>
									<tr>
										<td align="right" valign="top" nowrap><b>Ancillary Endorsement:</b></td>
										<td align="left">
										<select name="SMAncillaryEndorsement" id="SMAncillaryEndorsement">
											<option value="CARRIER_LEAVE_IF_NO_RESPONSE" <%=pcf_SelectOption("SMAncillaryEndorsement","CARRIER_LEAVE_IF_NO_RESPONSE")%>>Carrier leave if no response</option>
											<option value="ADDRESS_CORRECTION" <%=pcf_SelectOption("SMAncillaryEndorsement","ADDRESS_CORRECTION")%>>Address Correction</option>
											<option value="RETURN_SERVICE" <%=pcf_SelectOption("SMAncillaryEndorsement","RETURN_SERVICE")%>>Return Service</option>
										  </select>
											<%pcs_RequiredImageTag "SMAncillaryEndorsement", true%>
										</td>
									</tr>
									<tr>
										<td align="right" valign="top"><b>HUB ID:</b></td>
										<td align="left">
										<select name="SMHubID" id="SMHubID">
											<option value="5015" <%=pcf_SelectOption("SMHubID","5015")%>>5015</option>Northborough, MA</option>
											<option value="5087" <%=pcf_SelectOption("SMHubID","5087")%>>5087</option>Edison, NJ</option>
											<option value="5150" <%=pcf_SelectOption("SMHubID","5150")%>>5150</option>Pittsburgh, PA</option>
											<option value="5185" <%=pcf_SelectOption("SMHubID","5185")%>>5185</option>Allentown, PA</option>
											<option value="5254" <%=pcf_SelectOption("SMHubID","5254")%>>5254</option>Martinsburg, WV</option>
											<option value="5281" <%=pcf_SelectOption("SMHubID","5281")%>>5281</option>Charlotte, NC</option>
											<option value="5303" <%=pcf_SelectOption("SMHubID","5303")%>>5303</option>Atlanta, GA</option>
											<option value="5327" <%=pcf_SelectOption("SMHubID","5327")%>>5327</option>Orlando, FL</option>
											<option value="5379" <%=pcf_SelectOption("SMHubID","5379")%>>5379</option>Memphis, TN</option>
											<option value="5431" <%=pcf_SelectOption("SMHubID","5431")%>>5431</option>Grove City, OH</option>
											<option value="5465" <%=pcf_SelectOption("SMHubID","5465")%>>5465</option>Indianapolis, IN</option>
											<option value="5481" <%=pcf_SelectOption("SMHubID","5481")%>>5481</option>Detroit, MI</option>
											<option value="5531" <%=pcf_SelectOption("SMHubID","5531")%>>5531</option>New Berlin, WI</option>
											<option value="5552" <%=pcf_SelectOption("SMHubID","5552")%>>5552</option>Minneapolis, MN</option>
											<option value="5631" <%=pcf_SelectOption("SMHubID","5631")%>>5631</option>St. Louis, MO</option>
											<option value="5648" <%=pcf_SelectOption("SMHubID","5648")%>>5648</option>Kansas, KS</option>
											<option value="5751" <%=pcf_SelectOption("SMHubID","5751")%>>5751</option>Dallas, TX</option>
											<option value="5771" <%=pcf_SelectOption("SMHubID","5771")%>>5771</option>Houston, TX</option>
											<option value="5802" <%=pcf_SelectOption("SMHubID","5802")%>>5802</option>Denver, CO</option>
											<option value="5843" <%=pcf_SelectOption("SMHubID","5843")%>>5843</option>Salt Lake City, UT</option>
											<option value="5854" <%=pcf_SelectOption("SMHubID","5854")%>>5854</option>Phoenix, AZ</option>
											<option value="5902" <%=pcf_SelectOption("SMHubID","5902")%>>5902</option>Los Angeles, CA</option>
											<option value="5929" <%=pcf_SelectOption("SMHubID","5929")%>>5929</option>Chino, CA</option>
											<option value="5958" <%=pcf_SelectOption("SMHubID","5958")%>>5958</option>Sacramento, CA</option>
											<option value="5983" <%=pcf_SelectOption("SMHubID","5983")%>>5983</option>Seattle, WA</option>
										  </select>
											<%pcs_RequiredImageTag "SMHubID", true%>
										</td>
									</tr>
								</Table>
							</div>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<div class="expressfreight">
								<Table>
									<tr>
										<th colspan="2">Express Freight Options</th>
									</tr>
									<tr>
										<td width="54%" align="right" valign="top"><b>Packing List Enclosed?</b></td>
										<td width="46%" align="left">
											<input name="EFPackingListEnclosed" type="radio" value="1">Yes
											<input name="EFPackingListEnclosed" type="radio" value="1">No
										</td>
									</tr>
									<tr>
										<td align="right" valign="top"><b>ShippersLoadAndCount:</b></td>
										<td align="left">
											<input name="EFShippersLoadAndCount" type="text" id="EFShippersLoadAndCount" value="<%=pcf_FillFormField("EFShippersLoadAndCount", false)%>">
											<%pcs_RequiredImageTag "EFShippersLoadAndCount", false%>
										</td>
									</tr>
									<tr>
										<td align="right" valign="top" nowrap><b>Booking Confirmation Number</b></td>
										<td align="left">
											<input name="EFBookingConfirmationNumber" type="text" id="EFBookingConfirmationNumber" value="<%=pcf_FillFormField("EFBookingConfirmationNumber", false)%>">
											<%pcs_RequiredImageTag "EFBookingConfirmationNumber", false%>
										</td>
									</tr>
								</Table>
								</div>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<div class="freight_options">
								<Table width="100%">
									<tr>
									  <th colspan="2">Freight Options</th>
								  </tr>
									<tr>
									  <td align="right" valign="top" style="width: 20%"><b>Delivery Instructions:</b></td>
									  <td align="left">
										<input name="DeliveryInstructions" size="30" type="text" id="DeliveryInstructions" value="<%=pcf_FillFormField("DeliveryInstructions", false)%>">
										<%pcs_RequiredImageTag "DeliveryInstructions", false%>
									  </td>
								</tr>
								</Table>
								</div>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<div class="homedelivery">
								<Table width="100%">
									<tr>
									  <th colspan="2">Home Delivery Options</th>
								  </tr>
									<tr>
									  <td width="25%" align="right" valign="top"><b>Delivery Type:</b></td>
									  <td align="left">
										<select name="HomeDeliveryType" id="HomeDeliveryType">
										  <option value="">Please make a selection.</option>
										  <option value="DATE_CERTAIN" <%=pcf_SelectOption("HomeDeliveryType","DATE_CERTAIN")%>>FedEx Date Certain Home Delivery&reg;</option>
										  <option value="EVENING" <%=pcf_SelectOption("HomeDeliveryType","EVENING")%>>FedEx Evening Home Delivery&reg;</option>
										  <option value="APPOINTMENT" <%=pcf_SelectOption("HomeDeliveryType","APPOINTMENT")%>>FedEx Appointment Home Delivery&reg;</option>
										</select> 
										<% pcs_RequiredImageTag "HomeDeliveryType", isHomeDeliveryTypeRequired %>
										(Required for FedEx Home Delivery&reg; shipments.)
									  </td>
								  </tr>

									<tr>
									  <td align="right" valign="top"><b>Delivery Date:</b></td>
									  <td align="left">
										<input name="HomeDeliveryDate" type="text" id="HomeDeliveryDate" value="<%=pcf_FillFormField("HomeDeliveryDate", false)%>">
										<% pcs_RequiredImageTag "HomeDeliveryDate", isHomeDeliveryDateRequired %>
										e.g.: 2012-02-29</td>
								  </tr>
									<tr>
									  <td align="right" valign="top"><b>Delivery Phone:</b></td>
									  <td align="left">
										<input name="HomeDeliveryPhone" type="text" id="HomeDeliveryPhone" value="<%=pcf_FillFormField("HomeDeliveryPhone", false)%>">
										<% pcs_RequiredImageTag "HomeDeliveryPhone", isHomeDeliveryPhoneRequired %>
									  </td>
								  </tr>
									<tr>
									  <td align="right" valign="top"><b>Delivery Instructions:</b></td>
									  <td align="left">
										<input name="HomeDeliveryInstructions" type="text" id="HomeDeliveryInstructions" value="<%=pcf_FillFormField("HomeDeliveryInstructions", false)%>">
										<% pcs_RequiredImageTag "HomeDeliveryInstructions", false %>
									  </td>
								</tr>
								</Table>
								</div>
							</td>
						</tr>
					<% end if %>
				<% else %>
					<tr>
						<th colspan="2">This package has been shipped.</th>
					</tr>
				<% end if %>
				</table>
			</div>
			<%
		next %>

			<br />
			<br />

			<%
			pcv_strPreviousPage = "Orddetails.asp?id=" & pcv_intOrderID
			pcv_strAddPackagePage = "sds_ShipOrderWizard1.asp?idorder="&pcv_intOrderID&"&PageAction=FedExWs&PackageCount="&pcPackageCount&"&ItemsList="&pcv_strItemsList
			%>

			<p>
				<div align="center">
					<input type="button" class="btn btn-default"  name="Button" value="Start Over" onclick="document.location.href='<%=pcv_strPreviousPage%>'">
					<% if pcPackageCount<4 then %>
					<input type="button" class="btn btn-default"  name="Button" value="Add Another Package" onclick="document.location.href='<%=pcv_strAddPackagePage%>'">
					<% end if %>
					<input type="submit" name="submit" value="Process Shipment">
					<br />
					<br />
					<input type="button" class="btn btn-default"  name="Button" value="Go Back To Order Details" onclick="document.location.href='<%=pcv_strPreviousPage%>'">
				</div>
			</p>
		</td>
		</tr>
	<!--End -->
	</table>
</form>
<%
end if
'*******************************************************************************
' END: LOAD HTML FORM
'*******************************************************************************
%>
</td>
</tr>
</table>
<%

'// DESTROY THE SESSIONS
'pcs_ClearAllSessions
'Session("pcAdminPackageCount")=""
'Session("pcAdminOrderID")=""
'Session("pcGlobalArray")=""
'Session("pcAdminTotalWeight")=""
'Session("pcAdminDeclaredValue")=""
For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
	Session("pcAdminPrdList"&(xArrayCount+1))
Next
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->