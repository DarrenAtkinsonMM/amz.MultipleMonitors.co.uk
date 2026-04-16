<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx Web Services Shipping Configuration - Select Shipping Services" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp" -->
<!--#include file="AdminHeader.asp"-->

		<%
		Sub ProcessServiceVars(Service)
			If request.form("FREE-" & Service)="YES" then
				pcFreeAmount=request.form("AMT-" & Service)
				pcStrFreeShip=pcStrFreeShip&Service&"|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-" & Service)<>"0" AND request.form("HAND-" & Service)<>"" then
				If isNumeric(request.form("HAND-" & Service))=true then
					pcStrHandling=pcStrHandling&Service&"|"&replacecomma(request.form("HAND-" & Service))&"|"&request.form("SHFEE-" & Service)&","
				End If
			End if
			servicePriority=request.form("SP-" & Service)
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&Service&"|"&servicePriority&","
		End Sub

		if request.querystring("mode")="InAct" then
			' inactivate
			set rs=Server.CreateObject("ADODB.Recordset")

			query="UPDATE ShipmentTypes SET active=0, international=0 WHERE idShipment=" & FedExWS_ShipmentID & ";"
			set rs=connTemp.execute(query)

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEXWS"
		end if


		if request.querystring("mode")="Act" then
			' activate
			set rs=Server.CreateObject("ADODB.Recordset")
			query = "SELECT * FROM ShipmentTypes WHERE idShipment=" & FedExWS_ShipmentID & ";"
			set rs=connTemp.execute(query)
			IF RS.EOF THEN
				CALL CLOSEDB()
				RESPONSE.REDIRECT "upddb_v50.asp"
				response.end
			END IF
			query="UPDATE ShipmentTypes SET active=-1, international=0 WHERE idShipment=9;"
			set rs=connTemp.execute(query)

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEXWS"
		end if


		if request.querystring("mode")="del" then
			'remove
			set rs=Server.CreateObject("ADODB.Recordset")
			'clear all informatin out of shipService for service

			'// Deactivate FedEx
			query="UPDATE ShipmentTypes SET shipServer='', active=0, international=0 WHERE idShipment=" & FedExWS_ShipmentID & ";"
			set rs=connTemp.execute(query)

			'// Deactivate all FedEx shipping services
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE idShipment=" & FedExWS_ShipmentID & ";"
			connTemp.execute(query)

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEXWS"
		end if

		'check for real integers
		Function validNum2(strInput)
			DIM iposition		' Current position of the character or cursor
			validNum2 =  true
			if isNULL(strInput) OR trim(strInput)="" then
				validNum2 = false
			else
				'loop through each character in the string and validate that it is a number or integer
				For iposition=1 To Len(trim(strInput))
					if InStr(1, "12345676890", mid(strInput,iposition,1), 1) = 0 then
						validNum2 =  false
						Exit For
					end if
				Next
			end if
		end Function

		if request.form("submit")<>"" then

			pcStrService=request.form("FEDEXWS_SERVICE")
			if pcStrService="" then
				response.redirect "FedEXWS_EditShipOptions.asp?msg="&Server.URLEncode("Select at least one service.")
				response.end
			end if
			pcStrFreeShip=""
			pcStrHandling=""
			servicePriorityStr=""
			
			'// Load form fields for each item
			For Each Service In FedExWS_ShipmentTypes
				ProcessServiceVars(Service)
			Next

			'// Activate FedEx
			query="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=" & FedExWS_ShipmentID & ";"
			connTemp.execute(query)

			'Reset all FedEx services
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE idShipment=" & FedExWS_ShipmentID & ";"
			connTemp.execute(query)


			Dim i
			shipServiceArray=split(pcStrService,", ")

			for i=0 to ubound(shipServiceArray)
				query="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
			next

			pcStrFreeShipArray=split(pcStrFreeShip,",")
			for i=0 to (ubound(pcStrFreeShipArray)-1)
				pcFreeOverAmt=split(pcStrFreeShipArray(i),"|")
				if pcFreeOverAmt(1)>0 then
					pcServiceFree=-1
				else
					pcServiceFree=0
				end if
				query="UPDATE shipService SET serviceFree="&pcServiceFree&",serviceFreeOverAmt="&pcFreeOverAmt(1)&" WHERE serviceCode='"&pcFreeOverAmt(0)&"';"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
			next

			pcStrHandlingArray=split(pcStrHandling,",")
			for i=0 to (ubound(pcStrHandlingArray)-1)
				pcShipHandAmt=split(pcStrHandlingArray(i),"|")
				query="UPDATE shipService SET serviceHandlingFee="&pcShipHandAmt(1)&", serviceShowHandlingFee="&pcShipHandAmt(2)&" WHERE serviceCode='"&pcShipHandAmt(0)&"';"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
			next

				servicePriorityStrArray=split(servicePriorityStr,",")
				for i=0 to (ubound(servicePriorityStrArray)-1)
					SetServicePriority=split(servicePriorityStrArray(i),"|")
					query="UPDATE shipService SET servicePriority="&SetServicePriority(1)&" WHERE serviceCode='"&SetServicePriority(0)&"';"
					set rs=connTemp.execute(query)
				next

			set rs=nothing
			call closedb()
			if session("FedExWSSetUP")="YES" then
				response.redirect "FEDEXWS_EditSettings.asp"
			else
				response.redirect "viewshippingoptions.asp#FedEXWS"
			end if
		else %>


		<% ' START show message, if any %>
			<!--#include file="pcv4_showMessage.asp"-->
		<% 	' END show message %>

			<form name="form1" method="post" action="FedEXWS_EditShipOptions.asp" class="pcForms">
				<table class="pcCPcontent">
					<% 
					
					Function CreateShipServiceStr(serviceIdx)
						formString = ""
						
						pServiceCode=pcv_shipService(0, serviceIdx)
						pServiceActive=pcv_shipService(1, serviceIdx)
						pServicePriority=pcv_shipService(2, serviceIdx)
						pServiceDescription=pcv_shipService(3, serviceIdx)
						pServiceFree=pcv_shipService(4, serviceIdx)
						pServiceFreeOverAmt=pcv_shipService(5, serviceIdx)
						pServiceHandlingFee=pcv_shipService(6, serviceIdx)
						pServiceShowHandlingFee =pcv_shipService(7, serviceIdx)
						if pServiceActive="-1" then
							pServiceCheck="checked"
						else
							pServiceCheck=""
						end if
						if pServiceShowHandlingFee="0" then
							pServiceHandlingFeeChecked="checked"
						else
							pServiceHandlingFeeChecked=""
						end if
						if pServiceFree="-1" then
							pServiceFreeChecked="checked"
						else
							pServiceFreeChecked=""
						end if
						if pServicePriority = 0 then
							pServicePriority = serviceCount
						end if
						pTempString="<tr bgcolor='#DDEEFF'><td width='4%'><input type='checkbox' name='FEDEXWS_SERVICE' value='XXXX' "&pServiceCheck&"></td><td width='77%'><font color='#000000'><b>"&pServiceDescription&"</b></font></td><td width='19%' align='right'><strong>Order:&nbsp;</strong><input name='SP-XXXX' type='text' id='SP-XXXX' size='2' maxlength='10' value='"&pServicePriority&"'></td></tr>||||||||||<tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='FREE-XXXX' type='checkbox' id='FREE-XXXX' value='YES' "&pServiceFreeChecked&">Offer free shipping for orders over "&scCurSign&" <input name='AMT-XXXX' type='text' id='AMT-XXXX' size='6' maxlength='10' value='"&money(pServiceFreeOverAmt)&"'></td></tr><tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee "&scCurSign&" <input name='HAND-XXXX' type='text' id='HAND-XXXX' size='6' maxlength='10' value='"&money(pServiceHandlingFee)&"'></td></tr><tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='SHFEE-XXXX' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='SHFEE-XXXX' value='0' "&pServiceHandlingFeeChecked&">Integrate into shipping rate.<br><br></td></tr>"

						'// Load proper strings
						If pServiceCode = "SMART_POST" Then
							pTempString=replace(pTempString,"XXXX",pServiceCode)
							pTempString=replace(pTempString,"||||||||||","<tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><font color='#ff0000'>To use FedEx SmartPost you must have it enabled for your FedEx account. Please contact your FedEx account representative for more information.</font></td></tr>")
							formString=formString&pTempString
						Else
							pTempString=replace(pTempString,"XXXX",pServiceCode)
							pTempString=replace(pTempString,"||||||||||","")
							formString=formString&pTempString
						End If
						
						CreateShipServiceStr = formString
					End Function
					
					query="SELECT serviceCode, serviceActive, servicePriority, serviceDescription,serviceFree,serviceFreeOverAmt,serviceHandlingFee,serviceShowHandlingFee FROM shipService WHERE idShipment = " & FedExWS_ShipmentID & " ORDER BY servicePriority, idShipService ASC;"
					set rs=connTemp.execute(query)
					If Not rs.eof Then
						pcv_shipService = rs.getRows()
						intShipServiceCount = UBound(pcv_shipService, 2)
					End If
					set rs=nothing
					
					pcv_FormString=""
					
					serviceCount = 0
					
					'// See if there's custom ordering
					query="SELECT idShipment FROM shipService WHERE idShipment = " & FedExWS_ShipmentID & " AND servicePriority > 0;"
					set rs=connTemp.execute(query)
					If Not rs.eof Then
						'// Load with priority ordering
						For i = 0 To intShipServiceCount
							For Each Service In FedExWS_ShipmentTypes
								If pcv_shipService(0, i) = Service Then
									pcv_FormString = pcv_FormString & CreateShipServiceStr(i)
									serviceCount = serviceCount + 1
								End If
							Next
						Next
					Else
						'// Load with default ordering
						For Each Service In FedExWS_ShipmentTypes
							For i = 0 To intShipServiceCount
								If pcv_shipService(0, i) = Service Then
									pcv_FormString = pcv_FormString & CreateShipServiceStr(i)
									serviceCount = serviceCount + 1
								End If
							Next
						Next
					End If
					set rs=nothing
					
					response.write pcv_FormString
					%>

					<tr>
						<td colspan="3"><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">FedEx service marks are owned by Federal Express Corporation and used with permission.</div>
	</td>
					</tr>
					<tr>
						<td colspan="3" align="center"><input type="submit" name="Submit" value="Submit" class="btn btn-primary"></td>
					</tr>
				</table>
			</form>
			<% end if %>
<!--#include file="AdminFooter.asp"-->
