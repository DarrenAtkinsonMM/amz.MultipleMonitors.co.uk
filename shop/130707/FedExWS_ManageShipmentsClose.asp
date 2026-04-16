<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard - Close &amp; Print Manifest" %>
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
         
Function FedExDateFormat (FedExDate)
	If Len(FedExDate) < 1 Then Exit Function
	
	FedExDay=Day(FedExDate)
	FedExMonth=Month(FedExDate)
	FedExYear= Year(FedExDate)
	FedExDateFormat=FedExYear&"-"&Right(Cstr(FedExMonth + 100),2)&"-"&Right(Cstr(FedExDay + 100),2)
End Function

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
pcPageName="FedExWS_ManageShipmentsClose.asp"
ErrPageName="FedExWS_ManageShipmentsResults.asp"

'// OPEN DATABASE
call openDb()

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

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Set Required Data
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' SELECT DATA SET
' >>> Tables: pcPackageInfo
query = 		"SELECT pcPackageInfo_ShipMethod, pcPackageInfo_FDXCarrierCode "
query = query & "FROM pcPackageInfo "
query = query & "WHERE idOrder=" & pcv_intOrderID &" "

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then
	pcv_strShipMethod=Replace(rs("pcPackageInfo_ShipMethod"), "FedEx: ", "")
	pcv_strFDXCarrierCode=rs("pcPackageInfo_FDXCarrierCode")
end if
set rs=nothing

select case pcv_strFDXCarrierCode
	case "FDXG"
		pcv_strCarrierCode = "FedEx Ground"
	case "FXSP"
		pcv_strCarrierCode = "FedEx SmartPost"
end select

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>

<form name="form1" action="" method="post" class="pcForms">
	<input name="PackageInfo_ID" type="hidden" value="<%=pcv_strTrackingNumbers%>">
	<input name="id" type="hidden" value="<%=pcv_intOrderID%>">


  <% If Request.Form.Count > 0 Then %>
    <table class="pcCPcontent">

		<tr>
			<td>
			<%			
				pcv_strShipmentAccountNumber=pcv_strAccountNumber '// Owner's Account Number
        
        closeDateTime = Request.Form("closeDateTime")
        closeMethod = Request.Form("closeOption")
        If closeMethod&""="" Then
          closeMethod = "Ground"
        End If

				If IsDate(closeDateTime) Then
					closeDateTime = CDate(closeDateTime)
				End If

				fedex_postdataWS=""
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: Build Transaction
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
				pcv_strVersion = FedExWS_CloseVersion
				
				NameOfMethod = closeMethod & "CloseRequest"

				fedex_postdataWS=""
				fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns=""http://fedex.com/ws/close/v" & pcv_strVersion & """>"&vbcrlf
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
					objFedExClass.AddNewNode "CustomerTransactionId", "Close Shipment"
					objFedExClass.WriteParent "Localization", ""
						objFedExClass.AddNewNode "LanguageCode", "EN"
					objFedExClass.WriteParent "Localization", "/"
				objFedExClass.WriteParent "TransactionDetail", "/"

				'--------------------
				'// Version
				'--------------------
				objFedExClass.WriteParent "Version", ""
					objFedExClass.AddNewNode "ServiceId", "clos"
					objFedExClass.AddNewNode "Major", FedExWS_CloseVersion
					If closeMethod = "Ground" Then
						objFedExClass.AddNewNode "Intermediate", "1"
					Else
						objFedExClass.AddNewNode "Intermediate", "0"
					End If
					objFedExClass.AddNewNode "Minor", "0"
				objFedExClass.WriteParent "Version", "/"

        Function FedExDateTimeFormat(dateTime)
          FedExDateTimeFormat = FedExDateFormat(dateTime) & "T" & FormatDateTime(dateTime, 4) & ":00-04:00"
        End Function

				If closeMethod = "Ground" Then
					objFedExClass.AddNewNode "TimeUpToWhichShipmentsAreToBeClosed", FedExDateTimeFormat(closeDateTime)
				End If

				If closeMethod = "SmartPost" Then
					objFedExClass.AddNewNode "HubId", FEDEXWS_SMHUBID
					objFedExClass.AddNewNode "DestinationCountryCode", "US"
					objFedExClass.AddNewNode "PickUpCarrier", pcv_strFDXCarrierCode
				End If

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

				'call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "Close_" & strLogID & "_Req" & ".xml", true)

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Send Our Transaction.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				Set srvFEDEXWSXmlHttp = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
				Set objOutputXMLDocWS = Server.CreateObject("Microsoft.XMLDOM")
				Set objFedExStream = Server.CreateObject("ADODB.Stream")
				Set objFEDEXXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
				objFEDEXXmlDoc.async = False
				objFEDEXXmlDoc.validateOnParse = False
				if err.number>0 then
					err.clear
				end if

				call objFedExClass.SendXMLCloseRequest(fedex_postdata)

				'// Print out our response
				'response.Clear()
				'response.contenttype = "text/xml"
				'response.write FEDEXWS_result
				'response.end

				'call objFedExClass.pcs_LogTransaction(FEDEXWS_result, "Close_" & strLogID & "_Res" & ".xml", true)

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Load Our Response.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				call objFedExClass.LoadXMLResults(FEDEXWS_result)

		    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		    ' Check for errors from FedEx.
		    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		    '// master package error, no processing done
		    pcv_strErrorMsg = Cstr("")

        pcv_strReplyName = closeMethod & "CloseReply"

		    pcv_strErrorMsg = objFedExClass.ReadResponseNode("//" & pcv_strReplyName, "Notifications/Severity")

		    if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="NOTE" then
			    pcv_strErrorMsg = Cstr("")
		    else
			    pcv_strErrorMsg = objFedExClass.ReadResponseNode("//" & pcv_strReplyName, "Notifications/Message")
		    end if
		
		    if pcv_strErrorMsg&""="" then
  		    pcv_strErrorMsg = objFedExClass.ReadResponseNode("//soapenv:Fault", "faultstring")
		    end if

		    If pcv_strErrorMsg&"" <> "" Then
			    response.redirect ErrPageName&"?msg="&pcv_strErrorMsg
		    End If

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
					pcv_manifestFileName = objFedExClass.ReadResponseNode("//" & pcv_strReplyName, "Manifest/FileName")
					pcv_manifestFile = objFedExClass.ReadResponseNode("//" & pcv_strReplyName, "Manifest/File")
        
          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: SAVE MANIFEST FILE
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          If Len(pcv_manifestFile) > 0 Then
            fileName = "Manifest_" & pcv_manifestFileName & ".txt"
									
					  '// Create XML for Label
					  GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""" & fileName & """>"&pcv_manifestFile&"</Base64Data>"
									
					  '// Load label from the request stream
					  objFEDEXXmlDoc.loadXML GraphicXML
	
					  '// Use ADO stream to save the binary data
					  objFedExStream.Type = 1
					  objFedExStream.Open
	
					  objFedExStream.Write objFEDEXXmlDoc.selectSingleNode("/Base64Data").nodeTypedValue
					  err.clear
					  strFileName = objFEDEXXmlDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue
					  'Save the binary stream to the file and overwrite if it already exists in folder
					  objFedExStream.SaveToFile Server.MapPath("FedExLabels/"&strFileName),2
					  objFedExStream.Close()
          End If
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: SAVE MANIFEST FILE
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					%>
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
            <tr>
              <td colspan="2">
								<% If closeMethod = "Ground" Then %>
									All FedEx Ground orders up to <strong><%= FormatDateTime(closeDateTime) %></strong> have been closed. A manifest file has been generated that you can print out and give to your FedEx driver. Please click the link below to download and/or print the manifest.
								<% End If %>

								<% If closeMethod = "SmartPost" Then %>
									The FedEx SmartPost close request has been executed successfully. 
								<% End If %>
              </td>
            </tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<% If closeMethod = "Ground" Then %>
							<tr>
								<td colspan="2">
									<a href="FedExLabels/<%= fileName %>" target="_blank">Print Manifest</a>
								</td>

							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
						<% End If %>
						<tr>
							<td colspan="2">
								<%
								pcv_strPreviousPage = "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID
								%>
								<a href="<%=pcv_strPreviousPage%>" class="btn btn-primary">Done</a>
							</td>
						</tr>
          </table>
				<%
				end if

				set objFedExClass = nothing
			%>
			</td>
		</tr>
  </table>
  <% Else %>
    <table class="pcCPcontent">
      <tr>
        <td colspan="2">
					This page will allow you to submit a Ground or SmartPost Close request to FedEx. Choose an option below:
				</td>
			</tr>
			<%
				groundChecked = ""
				smartPostChecked = ""
					
				groundDisplay = "style=""display: none"""
				smartPostDisplay = "style=""display: none"""
					
				If pcv_strShipMethod = "FEDEX_GROUND" Or pcv_strShipMethod = "GROUND_HOME_DELIVERY" Then
					groundChecked = "checked"
					groundDisplay = ""
				End If

				If pcv_strShipMethod = "SMART_POST" Then
					smartPostChecked = "checked"
					smartPostDisplay = ""
				End If
			%>
			<script type=text/javascript>
				$pc(document).ready(function() {
					$pc("input[name='closeOption']").click(function() {
						$pc(".closeDiv").hide();
						$pc("#" + $pc(this).val()).show();
					});
				});
			</script>
      <tr>
        <td width="20%" align="right">
          <strong>Close Request Type: </strong>
        </td>
        <td>
          <input type="radio" id="closeOptionGround" name="closeOption" value="Ground" <%= groundChecked %>>
          <label for="closeOptionGround">Ground</label>

          &nbsp;&nbsp;
          
          <input type="radio" id="closeOptionSmartPost" name="closeOption" value="SmartPost" <%= smartPostChecked %>>
          <label for="closeOptionSmartPost">SmartPost</label>
        </td>
      </tr>
      <tbody class="closeDiv" id="Ground" <%= groundDisplay %>>
				<tr>
					<td width="20%" align="right">
						<%  
							closingDateTime = FormatDateTime(Now())
						%>
						<strong>Closing Date: </strong>
					</td>
					<td>
						<input type="hidden" name="closeDateTime" value="<%= closingDateTime %>">
						<%= closingDateTime %>
					</td>
				</tr>
				<tr>
					<td></td>
					<td>
						<p class="pcSmallText">
							*A manifest file will be generated that you can print out and give to your FedEx driver.
						</p>
					</td>
				</tr>
      </tbody>
			<tbody class="closeDiv" id="SmartPost" <%= smartPostDisplay %>>
				<tr>
					<td width="20%" align="right">
						<strong>SmartPost Hub ID:</strong>
					</td>
					<td>
						<%= FEDEXWS_SMHUBID %>
					</td>
				</tr>
				<tr>
					<td width="20%" align="right">
						<strong>Pickup Carrier:</strong>
					</td>
					<td>
						<%= pcv_strCarrierCode %>
					</td>
				</tr>
				
			</tbody>
      <tr>
        <td colspan="2" class="pcCPspacer"></td>
      </tr>
      <tr>
        <td colspan="2">
			    <input type="submit" class="btn btn-primary" name="submit" value="Request Close Shipment">&nbsp;&nbsp;
			    <%
			    pcv_strPreviousPage = "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID
			    %>
	        <a href="<%=pcv_strPreviousPage%>" class="btn btn-default">Back</a>
        </td>
      </tr>
    </table>
  <% End If %>
</form>
<%
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->