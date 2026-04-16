<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard - Cancel and Delete Shipment" %>
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
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, FedExAccountNumber, FedExMeterNumber, pcv_strCarrierCode
Dim pcv_strTrackingNumber, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction

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
pcPageName="FedExWS_ManageShipmentsCancel.asp"
ErrPageName="FedExWS_ManageShipmentsCancel.asp"

'// ACTION
pcv_strAction = request("Action")

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExWSClass

'// OPEN DATABASE


'// GET PACKAGE ID NUMBERS
PackageInfo_ID = Request("PackageInfo_ID")
SessionPackageInfo_ID = Session("pcAdminPackageInfo_ID")
if SessionPackageInfo_ID="" OR len(PackageInfo_ID)>0 then
	pcv_intPackageInfo = PackageInfo_ID
	Session("pcAdminPackageInfo_ID")=pcv_intPackageInfo
else
	pcv_intPackageInfo = SessionPackageInfo_ID
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
	FedExKey=rs("FedExKey")
	FedExPassword=rs("FedExPwd")
end if
set rs=nothing

'// SELECT DATA SET
' >>> Tables: pcPackageInfo
query = 		"SELECT pcPackageInfo.pcPackageInfo_ID, pcPackageInfo.pcPackageInfo_TrackingNumber, pcPackageInfo.pcPackageInfo_ShippedDate, "
query = query & "pcPackageInfo.pcPackageInfo_FDXCarrierCode "
query = query & "FROM pcPackageInfo "
query = query & "WHERE pcPackageInfo.pcPackageInfo_ID=" & pcv_intPackageInfo &" "
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then

'// LOOKUP THE PACKAGE INFO
	pcv_strTrackingNumber=rs("pcPackageInfo_TrackingNumber")
	pcv_strShipDate=rs("pcPackageInfo_ShippedDate")
	pcv_strCarrierCode=rs("pcPackageInfo_FDXCarrierCode")

end if
set rs=nothing

'// SET REQUIRED VARIABLES
pcv_strMethodName = "FDXShipDeleteRequest"
pcv_strMethodReply = "FDXShipDeleteReply"
CustomerTransactionIdentifier = pcv_strTrackingNumber
if pcv_strCarrierCode = "" OR pcv_strCarrierCode = NULL then
	pcv_strCarrierCode = "FDXG"
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Page Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'***************************************************************************
' START: POST BACK
'***************************************************************************
if request.form("submit")<>"" then
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Get all of the required information.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'// Generic error for page
	pcv_strGenericPageError = "At least one required field was empty. "

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If pcv_intErr>0 Then
		call closeDb()
response.redirect "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg=" & pcv_strGenericPageError
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Build Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		NameOfMethod = "DeleteShipmentRequest"
		objFedExClass.NewXMLLabelWS NameOfMethod, FedExKey, FedExPassword, FedExAccountNumber, FedExMeterNumber, FedExWS_ShipVersion, "ship"

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
			objFedExClass.AddNewNode "Major", FedExWS_ShipVersion
			objFedExClass.AddNewNode "Intermediate", "0"
			objFedExClass.AddNewNode "Minor", "0"
		objFedExClass.WriteParent "Version", "/"

		objFedExClass.WriteParent "TrackingId", ""
			objFedExClass.AddNewNode "TrackingIdType", "EXPRESS"
			'objFedExClass.AddNewNode "FormId", "0430"
			objFedExClass.AddNewNode "TrackingNumber", pcv_strTrackingNumber
		objFedExClass.WriteParent "TrackingId", "/"
		objFedExClass.AddNewNode "DeletionControl", "DELETE_ALL_PACKAGES"


		objFedExClass.EndXMLTransaction NameOfMethod
		'// Print out our newly formed request xml
		'response.write fedex_postdataWS&"<HR>"
		'response.end
		
		'Response.Clear()
		'Response.ContentType = "text/xml"
		'Response.Write fedex_postdataWS
		'Response.End()

		'// Log our Transaction
		'call objFedExClass.pcs_LogTransaction(fedex_postdataWS, "Delete_Req_"& pcv_strTrackingNumber &".xml", true)

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

		'Response.Clear()
		'Response.ContentType = "text/xml"
		'Response.Write FEDEXWS_result
		'Response.End()
		
		'// Log our Response
		'call objFedExClass.pcs_LogTransaction(FEDEXWS_result, "Delete_Res_"& pcv_strTrackingNumber &".xml", true)

		'// Print out our response
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Load Our Response.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.LoadXMLResults(FEDEXWS_result)
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Check for errors from FedEx.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// master package error, no processing done
		'pcv_strErrorMsg = objFedExClass.ReadResponseNode("//v13:ShipmentReply", "v13:HighestSeverity/v13:Severity")
		pcv_strErrorMsg = objFedExClass.ReadResponseNode("//v13:ShipmentReply", "v13:Notifications/v13:Severity")
			if err.number<>0 then
				response.write "FedEx Cancel 241: "&err.description
				response.end
			end if
		if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="WARNING" then
			isDeleted = 0
			pcv_strErrorMsg = Cstr("")
			if err.number<>0 then
				response.write "FedEx Cancel 248: "&err.description
				response.end
			end if
		else
			pcv_fault = ""
			if pcv_strErrorMsg="ERROR" OR pcv_strErrorMsg="NOTE" then
				pcv_strErrorMsg = objFedExClass.ReadResponseNode("//v13:ShipmentReply", "v13:Notifications/v13:Message")
			end if
			if pcv_strErrorMsg&""="" then
				pcv_strErrorMsg = objFedExClass.ReadResponseNode("//soapenv:Fault", "faultstring")
				pcv_isFault = "&fault=Delete_Res_"& objFedExClass.RandomNumber(999999999) &".txt"
			end if

			If len(pcv_strErrorMsg)>0 then
				isDeleted = 0
				if pcv_strErrorMsg="Shipment Delete was requested for a tracking number already in a deleted state." then
					isDeleted = 1
				else
					call closeDb()
					response.redirect "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg="&pcv_strErrorMsg & pcv_isFault
				end if
			End If
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Redirect with a Message OR complete some task.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if NOT len(pcv_strErrorMsg)>0 OR isDeleted = 1 then
			if isDeleted = 0 then
				pcv_strHideForm="true"
			end if
	
			'// Insert Code that will delete the package label info
			query="DELETE FROM pcPackageLabel WHERE pcPackageInfo_ID IN "
			query = query & "(SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE pcPackageInfo.pcPackageInfo_ID=" & pcv_intPackageInfo &");"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
	
			'// Insert Code that will delete the package shipment info
			query="DELETE FROM pcPackageInfo "
			query = query & "WHERE pcPackageInfo.pcPackageInfo_ID=" & pcv_intPackageInfo &";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing

			if err.number<>0 then
				response.write "277: "&err.description
				response.end
				
				call closeDb()
response.redirect pcPageName & "?msg=There was an error processing your request. Please try again or contact Customer Service at 1.800.Go.FedEx(R) 800.463.3339."
			else

				'// Restore the Ability to Re-Ship this package.
				query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=0,pcPackageInfo_ID=0 WHERE pcPackageInfo_ID=" & pcv_intPackageInfo & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing

				'// Check the Order's Shipping Status
				pcv_strOrderStatus=3
				query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder) WHERE Orders.idorder=" & pcv_intOrderID & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if not rs.eof then
					do while NOT rs.eof
						pcv_strTempOrdShipped=rs("pcPrdOrd_Shipped")
						if pcv_strTempOrdShipped=1 then
							pcv_strOrderStatus=7
						end if
						rs.movenext
					loop
				end if
				set rs=nothing
				'// Update the Order Status to "Pending" or "Partially Shipped"
				if pcv_strOrderStatus=7 then
					query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & pcv_intOrderID & ";"
				else
					query="UPDATE Orders SET orderStatus=3 WHERE idorder=" & pcv_intOrderID & ";"
				end if
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
				
				if pcv_strOrderStatus=3 then
					call pcs_hookOrderProcessed(pcv_intOrderID)
				end if

				'// Clear the Sessions
				pcs_ClearAllSessions()

				'// Close the Connection
				


				If isDeleted = 1 Then
					call closeDb()
response.redirect "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg="&pcv_strErrorMsg
				Else
					'// Redirect to the Shipment Manager
					call closeDb()
response.redirect "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg=Your Shipment has been deleted."
					response.end
				End If
			end if

			%>


			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">FedEx<sup>&reg;</sup> Shipment Canceled</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
			</table>

		<% end if
	end if

end if
'***************************************************************************
' END: POST BACK
'***************************************************************************
%>
<%
if msg<>"" then
	%>
	<div class="pcCPmessage">
		<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
	</div>
	<%
end if
%>
<% if pcv_strHideForm <> "true" then %>

<form name="form1" action="<%=pcPageName%>" method="post" class="pcForms">
	<input name="PackageInfo_ID" type="hidden" value="<%=pcv_intPackageInfo%>">
	<input name="id" type="hidden" value="<%=pcv_intOrderID%>">

	<table class="pcCPcontent">
		<tr>
			<td colspan="2" valign="top">
				<div class="pcCPmessageInfo">
					<ul style="margin: 0px">
						<li>When shipping with FedEx Express, you must delete a shipment prior to an end-of-day close performed at FedEx.</li>
						<li>When shipping with FedEx Ground, you cannot delete a shipment once a close operation has been performed.</li>
					</ul>
				</div>
			</td>
		</tr>
		<tr>
			<td colspan="2">
				<strong>Are you sure you want to delete this shipment?</strong>
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
			<input type="submit" class="btn btn-primary" name="submit" value="Request Delete Shipment">&nbsp;&nbsp;
			<%
			pcv_strPreviousPage = "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID
			%>
			<a href="<%=pcv_strPreviousPage%>" class="btn btn-default">Back</a>
			</td>
		</tr>
	</table>
</form>
<% end if
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->