<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard - Track Packages" %>
<% Section="mngAcc" %>
<%PmAdmin=9%>
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
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strCarrierCode
Dim pcv_strValue, pcv_strType, pcv_strTrackingNumberUniqueIdentifier, pcv_strShipDateRangeBegin, pcv_strShipDateRangeEnd, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim fedex_postdata, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// GET ORDER ID
pcv_strOrderID=Request("id")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_strOrderID)>0 then
	pcv_intOrderID=pcv_strOrderID
	Session("pcAdminOrderID")=pcv_intOrderID
else
	pcv_intOrderID=pcv_strSessionOrderID
end if

'// PAGE NAME
pcPageName="FedExWS_ManageShipmentsTrack.asp"
ErrPageName="FedExWS_ManageShipmentsResults.asp"

'// ACTION
pcv_strAction = request("Action")

'// OPEN DATABASE


'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExWSClass

'// REQUEST ARRAY OF PACKAGES TO TRACK "PackageInfo_ID"
if pcv_strAction="batch" then
	pcv_strTrackingNumbers=""
	Count=request("count")
	Dim k
	For k=1 to Count
		if (request("check" & k)<>"") then
			pcv_strTrackingNumbers=pcv_strTrackingNumbers & request("check" & k) & ","
		end if
	Next
	xStringLength = len(pcv_strTrackingNumbers)
	if xStringLength>0 then
		pcv_strTrackingNumbers = left(pcv_strTrackingNumbers,(xStringLength-1))
	end if
else
	pcv_strTrackingNumbers = Request("PackageInfo_ID")
end if

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


'// CREATE ARRAY OF PACKAGES
Dim xIdOptCounter, pcArrayTrackingNumbers
if NOT instr(pcv_strTrackingNumbers,",") then
	pcv_strTrackingNumbers = pcv_strTrackingNumbers&","
end if
pcArrayTrackingNumbers = split(pcv_strTrackingNumbers,",")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Tracking FedEx&reg; shipments for Order Number <%=(scpre+int(Session("pcAdminOrderID")))%></th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<span class="pcCPnotes">
			<strong>ATTENTION SHIPPERS:</strong> If your package has not yet been scanned by FedEx then the information on this page may not be accurate.
			FedEx sometimes reuses Tracking Numbers, so a Tracking Number may show data from a previous shipment until its scanned again.
			</span>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
</table>

<table class="pcCPcontent">

	<form name="form1" action="<%=pcPageName%>" method="post" class="pcForms">
	<input name="PackageInfo_ID" type="hidden" value="<%=pcv_strTrackingNumbers%>">
	<input name="id" type="hidden" value="<%=pcv_intOrderID%>">
		<tr>
			<td>
			<%
			'***************************************************************************
			' START LOOP THROUGH TRACKING
			'***************************************************************************


			for xIdOptCounter = 0 to Ubound(pcArrayTrackingNumbers)-1


				pcv_strTmpNumber = pcArrayTrackingNumbers(xIdOptCounter)

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Set Required Data
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' SELECT DATA SET
				' >>> Tables: pcPackageInfo
				query = 		"SELECT pcPackageInfo.pcPackageInfo_ID, pcPackageInfo.pcPackageInfo_TrackingNumber, pcPackageInfo.pcPackageInfo_ShipMethod, pcPackageInfo.pcPackageInfo_FDXCarrierCode "
				query = query & "FROM pcPackageInfo "
				query = query & "WHERE pcPackageInfo.pcPackageInfo_ID=" & pcv_strTmpNumber &" "

				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)

				if NOT rs.eof then
					pcv_strValue=rs("pcPackageInfo_TrackingNumber")
					pcv_strType= "" 'rs("")
					pcv_strShipDateRangeBegin= "" 'rs("")
					pcv_strShipDateRangeEnd= "" 'rs("")
					pcv_strDestinationCountryCode= "" 'rs("")
					pcv_strDestinationPostalCode= "" 'rs("")
					pcv_strLanguageCode= "" 'rs("")
					pcv_strLocaleCode= "" 'rs("")
					pcv_strDetailScans= "" 'rs("")
					pcv_strPagingToken= "" 'rs("")
					pcv_strTrackingNumberUniqueIdentifier= "" 'rs("")
					pcv_strShipMethod=rs("pcPackageInfo_ShipMethod")
					pcv_strCarrierCode=rs("pcPackageInfo_FDXCarrierCode")
				end if
				set rs=nothing
					
				pcv_strShipmentAccountNumber=pcv_strAccountNumber '// Owner's Account Number

				fedex_postdataWS=""
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: Build Transaction
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				if instr(pcv_strShipMethod, "FIRST_OVERNIGHT") then
					pcv_strCarrierCode = "FDXE"
				end if
				if instr(pcv_strShipMethod, "PRIORITY_OVERNIGHT") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "STANDARD_OVERNIGHT") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "FEDEX_2_DAY") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "FEDEX_EXPRESS_SAVER") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "FEDEX_GROUND") then
					pcv_strCarrierCode ="FDXG"
				end if
				if instr(pcv_strShipMethod, "GROUND_HOME_DELIVERY") then
					pcv_strCarrierCode ="FDXG"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_FIRST") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_PRIORITY") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_ECONOMY") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_PRIORITY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_ECONOMY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "FEDEX_1_DAY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "FEDEX_2_DAY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "FEDEX_3_DAY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "SMART_POST") then
					pcv_strCarrierCode = "FXSP"
				end if
							
				pcv_strVersion = FedExWS_TrackVersion
				
				if FedExWS_UseNamespace = true then
					fedex_xmlNamespace = ":v" & pcv_strVersion
				end if
				
				NameOfMethod = "TrackRequest"
				fedex_postdataWS=""
				fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns" & fedex_xmlNamespace & "=""http://fedex.com/ws/track/v" & pcv_strVersion & """>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
				
				objFedExClass.WriteParent NameOfMethod, ""
				
				objFedExClass.WriteParent "WebAuthenticationDetail", ""
					If CSPTurnOn = 1 Then
						objFedExClass.WriteParent "CspCredential", ""
							objFedExClass.AddNewNode "Key", pcv_strCSPKey
							objFedExClass.AddNewNode "Password", pcv_strCSPPassword
						objFedExClass.WriteParent "CspCredential", "/"
					End If
					objFedExClass.WriteParent "UserCredential", ""
						objFedExClass.AddNewNode "Key", FedExkey
						objFedExClass.AddNewNode "Password", FedExPassword
					objFedExClass.WriteParent "UserCredential", "/"
				objFedExClass.WriteParent "WebAuthenticationDetail", "/"
			
				objFedExClass.WriteParent "ClientDetail", ""
					objFedExClass.AddNewNode "AccountNumber", FedExAccountNumber
					objFedExClass.AddNewNode "MeterNumber", FedExMeterNumber
					objFedExClass.AddNewNode "ClientProductId", pcv_strClientProductID
					objFedExClass.AddNewNode "ClientProductVersion", pcv_strClientProductVersion
				objFedExClass.WriteParent "ClientDetail", "/"

				'--------------------
				'// TransactionDetail
				'--------------------
				objFedExClass.WriteParent "TransactionDetail", ""
					objFedExClass.AddNewNode "CustomerTransactionId", "Track Shipment"
					objFedExClass.WriteParent "Localization", ""
						objFedExClass.AddNewNode "LanguageCode", "EN"
					objFedExClass.WriteParent "Localization", "/"
				objFedExClass.WriteParent "TransactionDetail", "/"

				'--------------------
				'// Version
				'--------------------
				objFedExClass.WriteParent "Version", ""
					objFedExClass.AddNewNode "ServiceId", "trck"
					objFedExClass.AddNewNode "Major", pcv_strVersion
					objFedExClass.AddNewNode "Intermediate", "1"
					objFedExClass.AddNewNode "Minor", "0"
				objFedExClass.WriteParent "Version", "/"

				objFedExClass.WriteParent "SelectionDetails", ""
					objFedExClass.AddNewNode "CarrierCode", pcv_strCarrierCode
					objFedExClass.WriteParent "PackageIdentifier", ""
						objFedExClass.AddNewNode "Type", "TRACKING_NUMBER_OR_DOORTAG"
						objFedExClass.AddNewNode "Value", pcv_strValue
					objFedExClass.WriteParent "PackageIdentifier", "/"
				objFedExClass.WriteParent "SelectionDetails", "/"

				objFedExClass.EndXMLTransaction NameOfMethod

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: Build Transaction
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        strLogID = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)

				'// Print out our newly formed request xml
				'response.Clear()
				'response.contenttype = "text/xml"
				'response.write fedex_postdataWS
				'response.end

				call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "Track_" & strLogID & "_Req" & ".xml", true)

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Send Our Transaction.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'call objFedExClass.SendXMLRequest(fedex_postdata, pcv_strEnvironment)
				Set srvFEDEXWSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
				Set objOutputXMLDocWS = Server.CreateObject("Microsoft.XMLDOM")
				Set objFedExStream = Server.CreateObject("ADODB.Stream")
				Set objFEDEXXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
				objFEDEXXmlDoc.async = False
				objFEDEXXmlDoc.validateOnParse = False
				if err.number>0 then
					err.clear
				end if

				srvFEDEXWSXmlHttp.open "POST", FedExWSURL&"/track", false


				srvFEDEXWSXmlHttp.send(fedex_postdataWS)
				FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
				'// Print out our response

				call objFedExClass.pcs_LogTransaction(FEDEXWS_result, "Track_" & strLogID & "_Res" & ".xml", true)
				
				'response.Clear()
				'response.contenttype = "text/xml"
				'response.write FEDEXWS_result
				'response.end

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Load Our Response.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				call objFedExClass.LoadXMLResults(FEDEXWS_result)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Check for errors from FedEx.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// master package error, no processing done
		pcv_strErrorMsg = Cstr("")

		'// Try TrackReply first
		pcv_strErrorMsg = objFedExClass.ReadResponseNode("//TrackReply", "Notifications/Severity")

		if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="NOTE" then
			pcv_strErrorMsg = Cstr("")
		else
			pcv_strErrorMsg = objFedExClass.ReadResponseNode("//TrackReply", "Notifications/Message")

			'// If there's still no message, try TrackDetails
			if pcv_strErrorMsg&""="" then
				pcv_strErrorMsg = objFedExClass.ReadResponseNode("//TrackDetails", "Notifications/Severity")
			
				if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="NOTE" then
					pcv_strErrorMsg = Cstr("")
				else
					pcv_strErrorMsg = objFedExClass.ReadResponseNode("//TrackDetails", "Notification/Message")
				end if
			end if
		end if

		'// Still no error message, try soap error string
		if pcv_strErrorMsg&""="" then
  		pcv_strErrorMsg = objFedExClass.ReadResponseNode("//soapenv:Fault", "faultstring")
		end if

		If pcv_strErrorMsg&"" <> "" Then
			call closeDb()
			response.redirect ErrPageName&"?msg="&pcv_strErrorMsg
		End IF

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Redirect with a Message OR complete some task.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if NOT len(pcv_strErrorMsg)>0 then
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Set Our Response Data to Local.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Available Methods will search unlimited levels of Nodes by separating nodes with a "/".
			' 1) ReadResponseParent
			' 2) ReadResponseNode

			'///////////////////////////////////////////////////////////////////////////////////////////////////
			' Note: these are the primary values, but there are many more possible return values
			'///////////////////////////////////////////////////////////////////////////////////////////////////
					pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//TrackDetails", "TrackingNumber")
					pcv_strSignedForBy = objFedExClass.ReadResponseNode("//TrackDetails", "DeliverySignatureName")
					pcv_strService = objFedExClass.ReadResponseNode("//TrackDetails", "Service/Description")
					pcv_strShipTimeStamp = objFedExClass.ReadResponseNode("//TrackDetails", "ShipTimestamp")
					pcv_strActualDeliveryTimeStamp = objFedExClass.ReadResponseNode("//TrackDetails", "ActualDeliveryTimestamp")
					pcv_strStatusDescription = objFedExClass.ReadResponseNode("//TrackDetails", "StatusDetail/Description")
					pcv_strWeight = objFedExClass.ReadResponseNode("//TrackDetails", "PackageWeight")
					pcv_strPackagingDescription = objFedExClass.ReadResponseNode("//TrackDetails", "Packaging")
					pcv_strEstDeliveryTimestamp = objFedExClass.ReadResponseNode("//TrackDetails", "EstimatedDeliveryTimestamp")

					pcv_strEventDate = objFedExClass.ReadResponseNode("//TrackDetails", "Events/EventType")
					pcv_strEventTime = objFedExClass.ReadResponseNode("//TrackDetails", "Events/Timestamp")
					pcv_strEventType = objFedExClass.ReadResponseNode("//TrackDetails", "Events/EventType")
					pcv_strEventDescription = objFedExClass.ReadResponseNode("//TrackDetails", "Events/EventDescription")
					pcv_strEventStatusExceptionCode = objFedExClass.ReadResponseNode("//TrackDetails", "Events/StatusExceptionCode")
					pcv_strEventStatusExceptionDescription = objFedExClass.ReadResponseNode("//TrackDetails", "Events/StatusExceptionDescription")
					pcv_strEventAddressCity = objFedExClass.ReadResponseNode("//TrackDetails", "Events/Address/City")
					pcv_strEventAddressStateOrProvinceCode = objFedExClass.ReadResponseNode("//TrackDetails", "Events/Address/StateOrProvinceCode")
					pcv_strEventAddressPostalCode = objFedExClass.ReadResponseNode("//TrackDetails", "Events/Address/PostalCode")
					pcv_strEventAddressCountryCode = objFedExClass.ReadResponseNode("//TrackDetails", "Events/Address/CountryCode")
					'//
					tmpShipTimestamp = pcv_strShipTimeStamp
					if instr(tmpShipTimestamp,"T") then
						arrShipTimestamp = split(tmpShipTimestamp, "T")
						tmpShipTime = arrShipTimestamp(1)
						arrShipTimeFormat = split(tmpShipTime,":")
						tmpShipTimeHour = Cint(arrShipTimeFormat(0))
						tmpShipTimeMinutes = arrShipTimeFormat(1)
						tmpShipTimeSeconds = arrShipTimeFormat(2)
						'//Format hour and check for AM/PM
						if tmpShipTimeHour < 12 then
							tmpShipAMPM = "AM"
							tmpShipHour = Cint(tmpShipTimeHour)
						else
							tmpShipAMPM = "PM"
							tmpShipHour = Cint(tmpShipTimeHour) - Cint(12)
						end if
						tmpShipDate = arrShipTimestamp(0)
						arrShipDate = split(tmpShipDate,"-")
						tmpShipDay = arrShipDate(2)
						tmpShipMonth = arrShipDate(1)
						select case tmpShipMonth
							case "01"
								tmpShipMonth = "January"
							case "02"
								tmpShipMonth = "February"
							case "03"
								tmpShipMonth = "March"
							case "04"
								tmpShipMonth = "April"
							case "05"
								tmpShipMonth = "May"
							case "06"
								tmpShipMonth = "June"
							case "07"
								tmpShipMonth = "July"
							case "08"
								tmpShipMonth = "August"
							case "09"
								tmpShipMonth = "September"
							case "10"
								tmpShipMonth = "October"
							case "11"
								tmpShipMonth = "November"
							case "12"
								tmpShipMonth = "December"
						end select
						tmpShipYear = arrShipDate(0)
						FedExShipTimeDateStampF = tmpShipMonth&", "&tmpShipDay&" "&tmpShipYear&" "&tmpShipTimeHour&":"&tmpShipTimeMinutes&" "&tmpShipAMPM
					else
						FedExShipTimeDateStampF = "N/A"
					end if
					
					'//
					tmpActualDeliveryTimestamp = pcv_strActualDeliveryTimestamp
					if instr(tmpActualDeliveryTimestamp,"T") then
						arrActualDeliveryTimestamp = split(tmpActualDeliveryTimestamp, "T")
						tmpActualDeliveryTime = arrActualDeliveryTimestamp(1)
						arrActualTimeFormat = split(tmpActualDeliveryTime,":")
						tmpActualTimeHour = Cint(arrActualTimeFormat(0))
						tmpActualTimeMinutes = arrActualTimeFormat(1)
						tmpActualTimeSeconds = arrActualTimeFormat(2)
						'//Format hour and check for AM/PM
						if tmpActualTimeHour < 12 then
							tmpActualAMPM = "AM"
							tmpActualHour = Cint(tmpActualTimeHour)
						else
							tmpActualAMPM = "PM"
							tmpActualHour = Cint(tmpActualTimeHour) - Cint(12)
						end if
						tmpActualDeliveryDate = arrActualDeliveryTimestamp(0)
						arrActualDeliveryDate = split(tmpActualDeliveryDate,"-")
						tmpActualDeliveryDay = arrActualDeliveryDate(2)
						tmpActualDeliveryMonth = arrActualDeliveryDate(1)
						select case tmpActualDeliveryMonth
							case "01"
								tmpActualDeliveryMonth = "January"
							case "02"
								tmpActualDeliveryMonth = "February"
							case "03"
								tmpActualDeliveryMonth = "March"
							case "04"
								tmpActualDeliveryMonth = "April"
							case "05"
								tmpActualDeliveryMonth = "May"
							case "06"
								tmpActualDeliveryMonth = "June"
							case "07"
								tmpActualDeliveryMonth = "July"
							case "08"
								tmpActualDeliveryMonth = "August"
							case "09"
								tmpActualDeliveryMonth = "September"
							case "10"
								tmpActualDeliveryMonth = "October"
							case "11"
								tmpActualDeliveryMonth = "November"
							case "12"
								tmpActualDeliveryMonth = "December"
						end select
						tmpActualDeliveryYear = arrActualDeliveryDate(0)
						
						FedExActualTimeDateStampF = tmpActualDeliveryMonth&", "&tmpActualDeliveryDay&" "&tmpActualDeliveryYear&" "&tmpActualTimeHour&":"&tmpActualTimeMinutes&" "&tmpActualAMPM
					else
						FedExActualTimeDateStampF = "N/A"
					end if
					
					'//
					tmpEstDeliveryTimestamp = pcv_strEstDeliveryTimestamp
					if instr(tmpEstDeliveryTimestamp,"T") then
						arrEstDeliveryTimestamp = split(tmpEstDeliveryTimestamp, "T")
						tmpEstDeliveryTime = arrEstDeliveryTimestamp(1)
						arrEstTimeFormat = split(tmpEstDeliveryTime,":")
						tmpEstTimeHour = Cint(arrEstTimeFormat(0))
						tmpEstTimeMinutes = arrEstTimeFormat(1)
						tmpEstTimeSeconds = arrEstTimeFormat(2)
						'//Format hour and check for AM/PM
						if tmpEstTimeHour < 12 then
							tmpEstAMPM = "AM"
							tmpEstHour = Cint(tmpEstTimeHour)
						else
							tmpEstAMPM = "PM"
							tmpEstHour = Cint(tmpEstTimeHour) - Cint(12)
						end if
						tmpEstDeliveryDate = arrEstDeliveryTimestamp(0)
						arrEstDeliveryDate = split(tmpEstDeliveryDate,"-")
								tmpEstDeliveryDay = arrEstDeliveryDate(2)
								tmpEstDeliveryMonth = arrEstDeliveryDate(1)
								select case tmpEstDeliveryMonth
									case "01"
										tmpEstDeliveryMonth = "January"
									case "02"
										tmpEstDeliveryMonth = "February"
									case "03"
										tmpEstDeliveryMonth = "March"
									case "04"
										tmpEstDeliveryMonth = "April"
									case "05"
										tmpEstDeliveryMonth = "May"
									case "06"
										tmpEstDeliveryMonth = "June"
									case "07"
										tmpEstDeliveryMonth = "July"
									case "08"
										tmpEstDeliveryMonth = "August"
									case "09"
										tmpEstDeliveryMonth = "September"
									case "10"
										tmpEstDeliveryMonth = "October"
									case "11"
										tmpEstDeliveryMonth = "November"
							case "12"
								tmpEstDeliveryMonth = "December"
						end select
						tmpEstDeliveryYear = arrEstDeliveryDate(0)
					
						FedExEstTimeDateStampF = tmpEstDeliveryMonth&", "&tmpEstDeliveryDay&" "&tmpEstDeliveryYear&" "&tmpEstTimeHour&":"&tmpEstTimeMinutes&" "&tmpEstAMPM
					else
						FedExEstTimeDateStampF = "N/A"
					end if
					
					'select case pcv_strService
					'	case "PRIORITY_OVERNIGHT"
					'		pcv_strService="FedEx Priority Overnight<sup>&reg;</sup>"
					'	case "STANDARD_OVERNIGHT"
					'		pcv_strService="FedEx Standard Overnight<sup>&reg;</sup>"
					'	case "FIRST_OVERNIGHT"
					'		pcv_strService="FedEx First Overnight<sup>&reg;</sup>"
					'	case "FEDEX_2_DAY"
					'		pcv_strService="FedEx 2Day<sup>&reg;</sup>"
					'	case "FEDEX_EXPRESS_SAVER"
					'		pcv_strService="FedEx Express Saver<sup>&reg;</sup>"
					'	case "INTERNATIONAL_PRIORITY"
					'		pcv_strService="FedEx International Priority<sup>&reg;</sup>"
					'	case "INTERNATIONAL_ECONOMY"
						'	pcv_strService="FedEx International Economy<sup>&reg;</sup>"
					'	case "INTERNATIONAL_FIRST"
					'		pcv_strService="FedEx International First<sup>&reg;</sup>"
					'	case "FEDEX_1_DAY_FREIGHT"
					'		pcv_strService="FedEx 1Day<sup>&reg;</sup> Freight"
					'	case "FEDEX_2_DAY_FREIGHT"
					'		pcv_strService="FedEx 2Day<sup>&reg;</sup> Freight"
					'	case "FEDEX_3_DAY_FREIGHT"
					'		pcv_strService="FedEx 3Day<sup>&reg;</sup> Freight"
					'	case "FEDEX_GROUND"
					'		pcv_strService="FedEx Ground<sup>&reg;</sup>"
					'	case "GROUND_HOME_DELIVERY"
					'		pcv_strService="FedEx Home Delivery<sup>&reg;</sup>"
					'	case "INTERNATIONAL_PRIORITY_FREIGHT"
					'		pcv_strService="FedEx International Priority<sup>&reg;</sup> Freight"
					'	case "INTERNATIONAL_ECONOMY_FREIGHT"
					'		pcv_strService="FedEx International Economy<sup>&reg;</sup> Freight"
					'	case "SMART_POST"
					'		pcv_strService="FedEx SmartPost<sup>&reg;</sup>"
					'end select
					%>
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Tracking Number <%=pcv_strTrackingNumber%> Summary</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								  <tr>
									<td colspan="3">

										<table width="100%" border="0" cellpadding="0" cellspacing="0">
											<tr>
												<td width="19%" align="right">Tracking Number:</td>
												<td width="32%" align="left"><%=pcv_strTrackingNumber%></td>
												<td align="right">Service Type:</td>
											<td align="left"><%=pcv_strService%></td>
											</tr>
											<tr>
												<td align="right">Signed For By:</td>
												<td align="left">
												<% if pcv_strSignedForBy<>"" then %>
													<%=pcv_strSignedForBy%>
												<% else %>
													N/A
												<% end if %>
												</td>
												<td align="right">Destination:</td>
												<td align="left">
												<% if pcv_strEventAddressCity<>"" then %>
													<%=pcv_strEventAddressCity%>, <%=pcv_strEventAddressStateOrProvinceCode%>
												<% else %>
													N/A
												<% end if %>												</td>
											</tr>
											<tr>
												<td align="right">Ship Date:</td>
												<td align="left"><%=FedExShipTimeDateStampF%></td>
												<td align="right">Packaging:</td>
												<td align="left">
												<% if pcv_strPackagingDescription<>"" then %>
													<%=pcv_strPackagingDescription%>
												<% else %>
													N/A
												<% end if %>
												</td>
											</tr>
											<tr>
												<td align="right" nowrap="nowrap">Delivery Date/Time:</td>
												<td align="left">
												<% = FedExActualTimeDateStampF %>
												</td>
												<td align="right" nowrap="nowrap">Estimated Delivery Date:</td>
												<td align="left">
													<%=FedExEstTimeDateStampF%>
												</td>
											</tr>
										  <tr>
											<td align="right">Status:</td>
											<td align="left"><%=pcv_strStatusDescription%></td>

											<td width="18%" align="right">&nbsp;</td>
												<td width="31%" align="left">&nbsp;</td>
											</tr>
										</table>

									</td>
								  </tr>
								  <tr>
									<th width="44%"><strong>Date/ Time : Location</strong></th>
									<th width="16%"><strong>Scan Activity </strong></th>
									<th width="40%"><strong>Comments </strong></th>
								  </tr>
								<%

								'// Generate/ Trim Event Type
								arrayFedExEventType = objFedExClass.ReadResponsesArray("//Events", "EventType")
								'// Generate/ Trim Event Description
								arrayFedExEventDescription = objFedExClass.ReadResponsesArray("//Events", "StatusExceptionDescription")
								arrayFedExEventStatusExcDes = objFedExClass.ReadResponsesArray("//Events", "StatusExceptionDescription")
								arrayFedExTimestamp = objFedExClass.ReadResponsesArray("//Events", "Timestamp")
								arrayFedExEventDescription2 = objFedExClass.ReadResponsesArray("//Events", "EventDescription")
								arrayFedExCity = objFedExClass.ReadResponsesArray("//Events", "Address/City")
								arrayFedExStateOrProvinceCode = objFedExClass.ReadResponsesArray("//Events", "Address/StateOrProvinceCode")
								arrayFedExPostalCode = objFedExClass.ReadResponsesArray("//Events", "Address/PostalCode")
								arrayFedExCountryCode = objFedExClass.ReadResponsesArray("//Events", "Address/CountryCode")
								arrayFedExArrivalLocation = objFedExClass.ReadResponsesArray("//Events", "ArrivalLocation")

								arrayFedExEventType = split(arrayFedExEventType, ",")

								if arrayFedExEventDescription&""<>"" then
									arrayFedExEventDescription = split(arrayFedExEventDescription, ",")
								end if

								if arrayFedExEventDescription2&""<>"" then
									arrayFedExEventDescription = split(arrayFedExEventDescription2, ",")
								end if

								if arrayFedExEventStatusExcDes&""<>"" then
									arrayFedExEventStatusExcDes = split(arrayFedExEventStatusExcDes, ",")
								end if

								arrayFedExTimestamp = split(arrayFedExTimestamp, ",")
								arrayFedExCity = split(arrayFedExCity, ",")
								arrayFedExStateOrProvinceCode = split(arrayFedExStateOrProvinceCode, ",")
								arrayFedExPostalCode = split(arrayFedExPostalCode, ",")
								arrayFedExCountryCode = split(arrayFedExCountryCode, ",")
								arrayFedExArrivalLocation = split(arrayFedExArrivalLocation, ",")


								for bIdOptCounter = 0 to Ubound(arrayFedExEventType)-1

								tmpDeliveryTimestamp = arrayFedExTimestamp(bIdOptCounter)
								arrDeliveryTimestamp = split(tmpDeliveryTimestamp, "T")
								tmpDeliveryTime = arrDeliveryTimestamp(1)
								arrTimeFormat = split(tmpDeliveryTime,":")
								tmpTimeHour = Cint(arrTimeFormat(0))
								tmpTimeMinutes = arrTimeFormat(1)
								tmpTimeSeconds = arrTimeFormat(2)
								'//Format hour and check for AM/PM
								if tmpTimeHour < 12 then
									tmpAMPM = "AM"
									tmpHour = Cint(tmpTimeHour)
								else
									tmpAMPM = "PM"
									tmpHour = Cint(tmpTimeHour) - Cint(12)
								end if
								tmpDeliveryDate = arrDeliveryTimestamp(0)
								arrDeliveryDate = split(tmpDeliveryDate,"-")
								tmpDeliveryDay = arrDeliveryDate(2)
								tmpDeliveryMonth = arrDeliveryDate(1)
								select case tmpDeliveryMonth
									case "01"
										tmpDeliveryMonth = "January"
									case "02"
										tmpDeliveryMonth = "February"
									case "03"
										tmpDeliveryMonth = "March"
									case "04"
										tmpDeliveryMonth = "April"
									case "05"
										tmpDeliveryMonth = "May"
									case "06"
										tmpDeliveryMonth = "June"
									case "07"
										tmpDeliveryMonth = "July"
									case "08"
										tmpDeliveryMonth = "August"
									case "09"
										tmpDeliveryMonth = "September"
									case "10"
										tmpDeliveryMonth = "October"
									case "11"
										tmpDeliveryMonth = "November"
									case "12"
										tmpDeliveryMonth = "December"
								end select
								tmpDeliveryYear = arrDeliveryDate(0)

								FedExEventTimeDateStampF = tmpDeliveryMonth&", "&tmpDeliveryDay&" "&tmpDeliveryYear&" "&tmpTimeHour&":"&tmpTimeMinutes&" "&tmpAMPM
									'2012-07-11T00:00:00 %>
								  <tr>
									<td>
									<%
									select case arrayFedExArrivalLocation(bIdOptCounter)
										case "AIRPORT"
											FedExLocation = "Airport"
										case "CUSTOMER"
											FedExLocation = "Customer"
										case "CUSTOMS_BROKER"
											FedExLocation = "Customs Broker"
										case "DELIVERY_LOCATION"
											FedExLocation = "Delivery Location"
										case "DESTINATION_AIRPORT"
											FedExLocation = "Destination Airport"
										case "DESTINATION_FEDEX_FACILITY"
											FedExLocation = "Destination FedEx Facility"
										case "DROP_BOX"
											FedExLocation = "Drop Box"
										case "ENROUTE"
											FedExLocation = "Enroute"
										case "FEDEX_FACILITY"
											FedExLocation = "FedEx Facility"
										case "FEDEX_OFFICE_LOCATION"
											FedExLocation = "FedEx Office Location"
										case "INTERLINE_CARRIER"
											FedExLocation = "Interline Carrier"
										case "NON_FEDEX_FACILITY"
											FedExLocation = "Non-FedEx Facility"
										case "ORIGIN_AIRPORT"
											FedExLocation = "Origin Airport"
										case "ORIGIN_FEDEX_FACILITY"
											FedExLocation = "Origin FedEx Facility"
										case "PICKUP_LOCATION"
											FedExLocation = "Pickup Location"
										case "PLANE"
											FedExLocation = "Plane"
										case "PORT_OF_ENTRY"
											FedExLocation = "Port of Entry"
										case "SORT_FACILITY"
											FedExLocation = "Sort Facility"
										case "TURNPOINT"
											FedExLocation = "Turnpoint"
										case "VEHICLE"
											FedExLocation = "Vehicle"
										case else
											FedExLocation = "Unknown"
									end Select


									if FedExEventTimeDateStampF<>"" then
										%>
										<%=FedExEventTimeDateStampF%>
							    <% else %>
										N/A
									<% end if %>
									: <%=FedExLocation%></td>
									<td nowrap><%=arrayFedExEventDescription(bIdOptCounter)%>
									</td>
									<td><%=arrayFedExEventStatusExcDes(bIdOptCounter)%></td>
								  </tr>
								</table>
								<%
								next
								%>
							</td>
						</tr>
					</table>
				<%
				end if
				'set objFedExClass = nothing

			next
			'***************************************************************************
			' END LOOP THROUGH TRACKING
			'***************************************************************************
			%>
			</td>
		</tr>
	</form>
</table>
<%
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->