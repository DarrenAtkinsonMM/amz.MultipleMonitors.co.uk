<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Options Summary" %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/CPconstants.asp"-->
<!--#include file="../includes/pcShipTestModes.asp" -->
<!--#include file="AdminHeader.asp"-->

<style type="text/css">
.panel-title a {
 	display: block; 
}
</style>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<%
Sub ShowOptionStatus(enabled, disabled)
	If enabled Then
		%><span title="Shipping option is currently enabled." style="float: right" class="glyphicon glyphicon-ok"></span><%
	ElseIf disabled Then
		%><span title="Shipping option is currently disabled." style="float: right" class="glyphicon glyphicon-minus"></span><%
	Else
		%><span title="Shipping option has not been activated." style="float: right" class="glyphicon glyphicon-remove"></span><%
	End If
End Sub

Dim customServiceActive : customServiceActive = "0"

If request("switch")<>"" then
	pcv_Switch=request("switch")
	pcv_Service=request("service")
	pcv_USPSTM=USPS_TESTMODE
	pcv_UPSTM=UPS_TESTMODE
	if pcv_Service="UPS" then
		if pcv_Switch="TEST" then
			pcv_UPSTM="1"
		else
			pcv_UPSTM="0"
		end if
	end if
	if pcv_Service="USPS" then
		if pcv_Switch="TEST" then
			pcv_USPSTM="1"
		else
			pcv_USPSTM="0"
		end if
	end if
	Dim objFS
	Dim objFile

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")

	'//Get File
	if PPD="1" then
		pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/pcShipTestModes.asp")
	else
		pcStrFileName=Server.Mappath ("../includes/pcShipTestModes.asp")
	end if

	'//Write File
	Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
	objFile.WriteLine CHR(60)&CHR(37)& vbCrLf
	objFile.WriteLine "UPS_TESTMODE = """&pcv_UPSTM&"""" & vbCrLf
	objFile.WriteLine "USPS_TESTMODE = """&pcv_USPSTM&"""" & vbCrLf
	objFile.WriteLine CHR(37)&CHR(62)& vbCrLf
	objFile.Close
	set objFS=nothing
	set objFile=nothing
	'//Redirect
	call closeDb()
response.redirect "viewShippingOptions.asp"
end if

set rs=server.CreateObject("ADODB.RecordSet")

query="SELECT active FROM ShipmentTypes WHERE idShipment=3"
set rs=connTemp.execute(query)
UPSActive=rs("active")
if UPSActive=True or UPSActive<>0 then
	UPSActive="YES"
end if

' check if UPS is disabled
Dim UPSDisabled : UPSDisabled = "NO"
if UPSActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE serviceCode IN ('01','02','03','07','08','11','12','13','14','54','59','65') AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		UPSDisabled = "YES"
	end if
end if


query="SELECT active FROM ShipmentTypes WHERE idShipment=4"
set rs=connTemp.execute(query)
USPSActive=rs("active")
if USPSActive=True or USPSActive<>0 then
	USPSActive="YES"
end if

' check if USPS is disabled
Dim USPSDisabled : USPSDisabled = "NO"
if USPSActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE serviceCode IN ('9901','9902','9903','9904','9905','9906','9907','9908','9909','9910','9911','9912','9914','9915','9916','9917') AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		USPSDisabled = "YES"
	end if
end if


query="SELECT active FROM ShipmentTypes WHERE idShipment=9"
set rs=connTemp.execute(query)
If NOT rs.EOF Then
	FEDEXWSActive=rs("active")
End If
if FEDEXWSActive=True or FEDEXWSActive<>0 then
	FEDEXWSActive="YES"
end if
set rs = nothing

' check if FEDEX WebServices is disabled
Dim FEDEXWSDisabled : FEDEXWSDisabled = "NO"
if FEDEXWSActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE idShipment=" & FedExWS_ShipmentID & " AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		FEDEXWSDisabled = "YES"
	end if
end if


query="SELECT active FROM ShipmentTypes WHERE idShipment=7"
set rs=connTemp.execute(query)
CPActive=rs("active")
if CPActive=True or CPActive<>0 then
	CPActive="YES"
end if


' check if CP is disabled
Dim CPDisabled : CPDisabled = "NO"
if CPActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE idShipment=7 AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		CPDisabled = "YES"
	end if
end if
%>

<div id="accordion" class="panel-group">
    
    <div class="panel panel-default">
        <div class="panel-heading">
            <h4 class="panel-title">
                <a data-toggle="collapse" data-parent="#accordion" href="#collapseOne">
                	<% ShowOptionStatus UPSActive="YES", UPSDisabled="YES" %>
                    UPS&reg; Developer Kit
                </a>
            </h4>
        </div>
        <div id="collapseOne" class="panel-collapse collapse">
            <div class="panel-body">
                <table class="pcCPcontent">
    
                <%
                IF UPSActive="YES" THEN
    
                    if UPS_TESTMODE="1" then %>
                    <tr class="pcShowProductsMheader">
                        <td><p>You are currently runing UPS in <strong>&quot;TEST&quot; mode</strong>. <a href="viewShippingOptions.asp?switch=LIVE&service=UPS">Switch to &quot;LIVE&quot; mode</a>.</p></td>
                    </tr>
                    <tr class="pcShowProductsMheader">
                        <td><p><u>"TEST" mode only affect the printing of UPS shipping labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
                        </td>
                    </tr>
                    <% else %>
                    <tr class="pcShowProductsMheader">
                        <td><p>You are currently runing UPS in <strong>&quot;LIVE&quot; mode</strong>. <a href="viewShippingOptions.asp?switch=TEST&service=UPS">Switch to &quot;TEST&quot; mode</a>.</p></td>
                    </tr>
                    <tr class="pcShowProductsMheader">
                        <td><p><u>This setting only affect the priting of UPS shippings labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
                        </td>
                    </tr>
                    <% end if %>
                    <tr>
                        <td class="pcCPspacer"></td>
                    </tr>
                    <tr>
                        <td>
                        <ul>
                            <li><img src="../pc/images/ups_pri_scr_lbg_sm.jpg" alt="UPS&reg; Developer Kit" align="right">Services: - <a href="UPS_EditShipOptions.asp">Edit</a>
                                <ul>
                                <% query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
                                set rs=connTemp.execute(query)
                                do until rs.eof
                                    VarServiceCode=rs("serviceCode")
                                    select case VarServiceCode
                                    case "01","02","03","07","08","11","12","13","14","54","59","65"
                                        response.write "<li>"&rs("serviceDescription")&"</li>"
                                    end select
                                    rs.movenext
                                loop %>
                                </ul>
                             </li>
                        <% 'if ups_license contains info, do not show this link
                        query="SELECT ups_UserId FROM ups_license WHERE idUPS=1;"
                        set rs=connTemp.execute(query)
                        if len(rs("ups_UserId")&"A")=1 then %>
                            <li>UPS&reg; Developer Kit License Information - <a href="UPS_EditLicense.asp">Edit</a></li>
                        <% end if %>
                        <li>UPS&reg; Developer Kit: User Preferences - <a href="UPS_Preferences.asp">Edit</a></li>
                        <li>UPS&reg; Developer Kit: Shipping Settings - <a href="UPS_EditSettings.asp">Edit</a></li>
                        <li>Current Settings:
                            <ul>
                                <li>Default packaging:
                                <% select case UPS_PACKAGE_TYPE
                                    case "00"
                                        response.write "Unknown"
                                    case "01"
                                        response.write "UPS Letter"
                                    case "02"
                                        response.write "Package"
                                    case "03"
                                        response.write "UPS Tube"
                                    case "04"
                                        response.write "UPS Pak"
                                    case "21"
                                        response.write "UPS Express Box"
                                    case "24"
                                        response.write "UPS 25KG Box&reg;"
                                end select %>
                                </li>
                                <li>Default Account Type:
                                  <% select case UPS_PICKUP_TYPE
                                    case "01"
                                        response.write "Daily Pickup"
                                    case "03"
                                        response.write "Occasional Pickup"
                                    case "11"
                                        response.write "Suggested Retail Rates (UPS Store)"
                                end select %>
                                </li>
                                <li>Default Package Dimensions:
                                    <ul>
                                        <li>Height: <%=UPS_HEIGHT%>&nbsp;<%=UPS_DIM_UNIT%></li>
                                        <li>Width: <%=UPS_WIDTH%>&nbsp;<%=UPS_DIM_UNIT%></li>
                                        <li>Length: <%=UPS_LENGTH%>&nbsp;<%=UPS_DIM_UNIT%></li>
                                    </ul>
                                </li>
                            </ul>
                        </li>
                        <li><a href="OrderShippingOptions.asp?Provider=ups">Set Display Order</a></li>
                        <li><a href="javascript:if (confirm('You are about to permanantly delete all current UPS settings. You will have to register again with the UPS&reg; Developer Kit from your ProductCart Control Panel in order to reactivate UPS. This feature should only be used when UPS rates cannot be retrieved and no explanation other than a misconfigured account can be found. Are you sure you want to complete this action?')) location='pcResetUPS.asp'">Reset UPS&reg; Developer Kit registration.</a></li>
                        <li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='UPS_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
                        <li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='UPS_EditShipOptions.asp?mode=del'">Remove</a></li>
                        </ul>
                      </td>
                    </tr>
    
                <% ELSEIF UPSDisabled = "YES" then %>
                    <tr>
                        <td>
                            <img src="../pc/images/ups_pri_scr_lbg_sm2.jpg" alt="Enable (Reactivate) UPS" hspace="10"><strong>UPS</strong>&nbsp; is disabled - <a href="UPS_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
                        </td>
                    </tr>
                <% ELSE %>
                    <tr>
                        <td>
                            <img src="../pc/images/ups_pri_scr_lbg_sm2.jpg" alt="Activate UPS" hspace="10"><strong>UPS&reg; Developer Kit</strong> is not active - <a href="ConfigureOption1.asp">Activate</a>.
                        </td>
                    </tr>
                <% END IF %>
    
                <tr align="center">
                    <td><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved.</div></td>
                </tr>
             </table>
        </div>
    </div> 
</div>

<div class="panel panel-default">
    <div class="panel-heading">
        <h4 class="panel-title">
            <a data-toggle="collapse" data-parent="#accordion" href="#collapse2">
                <% ShowOptionStatus USPSActive="YES", USPSDisabled="YES" %>
                United States Postal Service (USPS)
            </a>
        </h4>
    </div>
    <div id="collapse2" class="panel-collapse collapse">
        <div class="panel-body">
			<table class="pcCPcontent">

			<% if USPSActive="YES" then %>
				<%
				query="SELECT pcES_UserID,pcES_PassP,pcES_AutoRefill,pcES_TriggerAmount,pcES_LogTrans,pcES_Reg,pcES_TestMode FROM pcEDCSettings WHERE pcES_Reg=1;"
				set rsQ=connTemp.execute(query)

				tmpEDCUserID=0
				if not rsQ.eof then
					EndiciaReg=1
					tmpEDCUserID=rsQ("pcES_UserID")
				else
					EndiciaReg=0
				end if
				set rsQ=nothing%>
				<%if EndiciaReg=0 then
				if USPS_TESTMODE="1" then %>
				<tr class="pcShowProductsMheader">
					<td><p>You are currently runing USPS in <strong>&quot;TEST&quot; mode</strong>. <a href="viewShippingOptions.asp?switch=LIVE&service=USPS">Switch to &quot;LIVE&quot; mode</a></p></td>
				</tr>
				<tr class="pcShowProductsMheader">
					<td><p><u>"TEST" mode only affects the printing of USPS shipping labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
					</td>
				</tr>
				<% else %>
				<tr class="pcShowProductsMheader">
					<td><p>You are currently runing USPS in <strong>&quot;LIVE&quot; mode.</strong>. <a href="viewShippingOptions.asp?switch=TEST&service=USPS">Switch to &quot;TEST&quot; mode</a>.</p></td>
				</tr>
				<tr class="pcShowProductsMheader">
					<td><p><u>This setting only affect the printing of USPS shipping labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
					</td>
				</tr>
				<% end if
				end if %>
				<tr>
					<td class="pcCPspacer"></td>
				</tr>
				<tr>
					<td>
						<ul>
							<li>Active Shipping Services - <a href="USPS_EditShipOptions.asp">Edit</a>
								<ul>
								<% dim USPSServiceFound
								USPSServiceFound=0
								query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
								set rs=connTemp.execute(query)
								do until rs.eof
									VarServiceCode=rs("serviceCode")
									select case VarServiceCode
									case "9901","9902","9903","9904","9905","9906","9907","9908","9909","9910","9911","9912","9914","9915","9916","9917"
										response.write "<li>"&rs("serviceDescription")&"</li>"
										USPSServiceFound=1
									end select
									rs.movenext
								loop
								if USPSServiceFound=0 then
									response.write "<font color=#FF0000>NO services are active, please choose at least one service for the USPS Provider.</font>&nbsp; - <a href='USPS_EditShipOptions.asp'>Add</a>"
								end if
								set rs=nothing %>
								</ul>
							</li>
							<li>USPS License Information - <a href="USPS_EditLicense.asp">Edit</a></li>
							<li>USPS Shipping Settings - <a href="USPS_EditSettings.asp">Edit</a></li>
							<li>Current Shipping Settings:
								<ul>
									<li>Default Express Mail packaging: <%=USPS_EM_PACKAGE%></li>
									<% pcv_PMPackage=""
									select case USPS_PM_PACKAGE
									case "Flat Rate Envelope"
										pcv_PMPackage="Priority Mail Flat Rate Envelope, 12.5&quot; x 9.5&quot;"
									case "Flat Rate Box"
										pcv_PMPackage="Priority Mail Box, 12.25&quot; x 15.5&quot; x 3&quot;"
									case "Flat Rate Box1"
										pcv_PMPackage="Priority Mail Flat Rate Box, 14&quot; x 12&quot; x 3.5&quot;"
									case "Flat Rate Box2"
										pcv_PMPackage="Priority Mail Flat Rate Box, 11.25&quot; x 8.75&quot; x 6&quot;"
									end select %>
									<li>Default Priority Mail packaging: <%=pcv_PMPackage%></li>
								</ul>
							</li>
							<li><a href="OrderShippingOptions.asp?Provider=usps">Set Display Order</a></li>
							<li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='USPS_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
							<li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='USPS_EditShipOptions.asp?mode=del'">Remove</a></li>
						</ul>
					</td>
				</tr>
			</table>
        </div>
    </div> 
</div>

<div class="panel panel-default">
    <div class="panel-heading">
        <h4 class="panel-title">
            <a data-toggle="collapse" data-parent="#accordion" href="#collapse3">
                <% ShowOptionStatus EndiciaReg<>0, false %>
                Endicia's Postage Label Services for USPS
            </a>
        </h4>
    </div>
    <div id="collapse3" class="panel-collapse collapse">
        <div class="panel-body">
			<table class="pcCPcontent">
                <tr valign="top">
                    <td colspan="2">
                        <img src="images/PoweredByEndicia_small.jpg" border="0" align="right" hspace="20">
                        <%if EndiciaReg=0 then%>
                            You can choose Endicia's Postage Label Services to print USPS postage.<br>
                            <a href="EDC_signup.asp?reg=1">Click here</a> to sign-up for an Endicia account.
                        <%else%>
                            <%if tmpEDCUserID="0" OR tmpEDCUserID="" then%>
                                You signed up to use Endicia's service to print USPS postage.<br>
                                Please <a href="EDC_manage.asp">click here</a> to complete the sign up process and activate your account
                            <%else%>
                                You are using Endicia's Postage Label Services.<br>
                                <a href="EDC_manage.asp">Click here</a> to manage your Endicia account.
                                <br><br>
                                <a href="javascript:if (confirm('You are about to remove Endicia and all of its settings. Would you like to continue?')) location='ECD_remove.asp'">Remove</a> Endicia
                            <%end if%>
                      <%end if%>
                    </td>
                </tr>
                <tr>
                    <td cospan="2" class="pcCPspacer"></td>
                </tr>
            <% elseif USPSDisabled = "YES" then %>
                <tr>
                    <td>
                        <strong>USPS</strong>&nbsp; is disabled - <a href="USPS_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
                    </td>
                </tr>
    
            <% else %>
                <tr>
                    <td><strong>USPS</strong>&nbsp; is not active - <a href="ConfigureOption2.asp">Activate</a> <br><span class="pcSmallText">Note: to take advantage of Endicia's Postage Printing Service, you must first activate USPS.</span></td>
                </tr>
            <% end if %>
            </table>
        </div>
    </div> 
</div>

<div class="panel panel-default">
    <div class="panel-heading">
        <h4 class="panel-title">
            <a data-toggle="collapse" data-parent="#accordion" href="#collapse4">
                <% ShowOptionStatus FedExWSActive="YES", FedExWSDisabled="YES" %>
                FedEx Web Services
            </a>
        </h4>
    </div>
    <div id="collapse4" class="panel-collapse collapse">
        <div class="panel-body">
			<table class="pcCPcontent">

			<% if FedExWSActive="YES" then %>

				<tr>
					<td>
						<ul>
						<li>FedEx&reg; Shipping Services - <a href="FedExWS_EditShipOptions.asp">Edit</a>
							<ul>
								<% 
								
								query="SELECT serviceCode, serviceDescription FROM shipService WHERE serviceActive=-1 AND idShipment = " & FedExWS_ShipmentID & " ORDER BY servicePriority, idShipService ASC;"
								set rs=connTemp.execute(query)
								If Not rs.eof Then
									pcv_FedExShipService = rs.getRows()
									intShipServiceCount = UBound(pcv_FedExShipService, 2)
								End If
								set rs=nothing
								
								query="SELECT idShipment FROM shipService WHERE idShipment = " & FedExWS_ShipmentID & " AND serviceActive=-1 AND servicePriority > 0;"
								set rs=connTemp.execute(query)
								If Not rs.eof Then
									'// Load with priority ordering
									For i = 0 To intShipServiceCount
										For Each Service In FedExWS_ShipmentTypes
											If pcv_FedExShipService(0, i) = Service Then
												response.write "<li>"&pcv_FedExShipService(1, i)&"</li>"
												FEDEXWSServiceFound=1
											End If
										Next
									Next
								Else
									'// Load with default ordering
									For Each Service In FedExWS_ShipmentTypes
										For i = 0 To intShipServiceCount
											If pcv_FedExShipService(0, i) = Service Then
												response.write "<li>"&pcv_FedExShipService(1, i)&"</li>"
												FEDEXWSServiceFound=1
											End If
										Next
									Next
								End If
								set rs=nothing 
								
								if FEDEXWSServiceFound=0 then
									response.write "<li>NO services are active, please choose at least one service for the FedEx Provider. - <a href='FEDEXWS_EditShipOptions.asp'>Add</a></li>"
								end if
								%>
							</ul>
						</li>
						<li>FedEx Shipping Settings - <a href="FEDEXWS_EditSettings.asp">Edit</a></li>
						<li>Current Settings:
							<ul>
								<li>Default Package Type:
									<% select case FEDEXWS_FEDEX_PACKAGE
										case "YOUR_PACKAGING"
											response.write "Your Packaging"
										case "FEDEX_TUBE"
											response.write "FedEx&reg; Tube"
										case "FEDEX_PAK"
											response.write "FedEx&reg; Pak"
										case "FEDEX_ENVELOPE"
											response.write "FedEx&reg; Envelope"
										case "FEDEX_SMALL_BOX"
											response.write "FedEx&reg; Small Box"
										case "FEDEX_MEDIUM_BOX"
											response.write "FedEx&reg; Medium Box"
										case "FEDEX_LARGE_BOX"
											response.write "FedEx&reg; Large Box"
										case "FEDEX_EXTRA_LARGE_BOX"
											response.write "FedEx&reg; Extra Large Box"
										case "FEDEX_10KG_BOX"
											response.write "FedEx&reg; 10KG Box"
										case "FEDEX_25KG_BOX"
											response.write "FedEx&reg; 25KG Box"
									end select %>
								</li>
								<li>Default Drop-off Type:
								<% select case FEDEXWS_DROPOFF_TYPE
									case "REGULAR_PICKUP"
										response.write "Regular Pickup"
									case "REQUEST_COURIER"
										response.write "Request Courier"
									case "DROP_BOX"
										response.write "Dropbox"
									case "BUSINESS_SERVICE_CENTER"
										response.write "Business Service Center"
									case "STATION"
										response.write "FedEx&reg; Station"
								end select %>
								</li>
								<li>Default Package Dimensions:
									<ul>
									<li>Height:
									<%=FEDEXWS_HEIGHT%>&nbsp;<%=FEDEXWS_DIM_UNIT%></li>
									<li>Width:
									<%=FEDEXWS_WIDTH%>&nbsp;<%=FEDEXWS_DIM_UNIT%></li>
									<li>Length:
									<%=FEDEXWS_LENGTH%>&nbsp;<%=FEDEXWS_DIM_UNIT%></li>
									</ul>
								</li>
							</ul>
						</li>
						<li><a href="OrderShippingOptions.asp?Provider=fedexWS">Set Display Order</a></li>
						<li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='FedEXWS_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
						<li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='FedEXWS_EditShipOptions.asp?mode=del'">Remove</a></li>
						</ul>
					</td>
				</tr>
			<%elseif FEDEXWSDisabled = "YES" then %>
			<tr>
				<td>
					<img src="../pc/images/fedex_corp_logo.gif" alt="Enable (Reactivate)" style="padding: 5px"><strong>FedEx</strong> is disabled - <a href="FedEXWS_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
				</td>
			</tr>
			<tr align="center">
				<td><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">FedEx service marks are owned by Federal Express Corporation and used with permission.</div></td>
			</tr>

			<% ELSE %>
				<tr>
					<td><img src="../pc/images/fedex_corp_logo.gif" alt="Activate FedEx" style="padding: 5px"><strong>FedEx</strong> is not active - <a href="ConfigureOption5.asp">Activate</a> | <a href="http://wiki.productcart.com/productcart/shipping-federal_express_ws" target="_blank">Help</a></td>
				</tr>
				<tr align="center">
					<td><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">FedEx service marks are owned by Federal Express Corporation and used with permission.</div></td>
				</tr>

			<% end if %>

			</table>
        </div>
    </div> 
</div>

<div class="panel panel-default">
    <div class="panel-heading">
        <h4 class="panel-title">
            <a data-toggle="collapse" data-parent="#accordion" href="#collapse5">
                <% ShowOptionStatus CPActive="YES", CPDisabled="YES" %>
                Canada Post
            </a>
        </h4>
    </div>
    <div id="collapse5" class="panel-collapse collapse">
        <div class="panel-body">
			<table class="pcCPcontent">

			<% if CPActive="YES" then %>

				<tr>
					<td>
						<ul>
						<li>Canada Post Shipping Services - <a href="CP_EditShipOptions.asp">Edit</a>
							<ul>
							<% Dim CPServiceFound
							CPServiceFound=0
							query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 AND idShipment=7 Order by servicePriority;"
							set rs=connTemp.execute(query)
							do until rs.eof
								response.write "<li>"&rs("serviceDescription")&"</li>"
								CPServiceFound=1
								rs.movenext
							loop
							if CPServiceFound=0 then
								response.write "<li>NO services are active, please choose at least one service for Canada Post. - <a href='CP_EditShipOptions.asp'>Add</a></li>"
							end if
							set rs=nothing %>
							</ul>
						</li>
						<li>Canada Post User License - <a href="CP_EditLicense.asp">Edit</a></li>
						<li>Canada Post Shipping Settings - <a href="CP_EditSettings.asp">Edit</a></li>
						<li>Current Shipping Settings:
							<ul>
								<li>Default Package Dimensions:
									<ul>
										<li>Height:
										<%=CP_Height%> cm</li>
										<li>Width:
										<%=CP_Width%> cm</li>
										<li>Length:
										<%=CP_Length%> cm</li>
									</ul>
								</li>
							</ul>
						<li><a href="OrderShippingOptions.asp?Provider=cp">Set Display Order</a></li>
						<li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='CP_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
						<li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='CP_EditShipOptions.asp?mode=del'">Remove</a></li>
						</ul>
					</td>
				</tr>
			<%elseif CPDisabled = "YES" then %>
                <tr>
                    <td>
                        <strong>Canada Post</strong>&nbsp; is disabled - <a href="CP_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
                    </td>
                </tr>
			<% else %>
				<tr>
					<td><strong>Canada Post</strong>&nbsp;is not active - <a href="ConfigureOption4.asp">Add</a></td>
				</tr>
			<% end if %>

			</table>
        </div>
    </div> 
</div>
                    
<% dim iCustomCnt
iCustCnt=0
query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceCode LIKE 'C%';"
set rs=connTemp.execute(query)
%>
<div class="panel panel-default">
	<div class="panel-heading">
		<h4 class="panel-title">
			<a data-toggle="collapse" data-parent="#accordion" href="#collapse6">
				<% ShowOptionStatus rs.eof<>true, false %>
				Custom Shipping Options
			</a>
		</h4>
	</div>
	<div id="collapse6" class="panel-collapse collapse">
		<div class="panel-body">
            <table class="pcCPcontent">
                <%
                if rs.eof then
                %>
                <tr>
                    <td colspan="2">No Custom Shipping Options have been added - <a href="AddCustomShipping.asp">Add</a></td>
                </tr>
                <%
                else
                %>
                <tr>
                    <td colspan="2" style="border-bottom: 1px dashed #CCC;" class="cpLinksList"><a href="AddCustomShipping.asp">Add New</a> : <a href="OrderShippingOptions.asp?Provider=x">Set Display Order</a></td>
                </tr>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <%
                    do until rs.eof
                        VarServiceCode=rs("serviceCode")
                        customServiceActive = rs("serviceActive")
                        if left(VarServiceCode,1)="C" then
                            iCustCnt=iCustCnt+1
                            VaridFlatShipType=replace(VarServiceCode,"C","")
                            query="SELECT FlatShipTypeDesc FROM FlatShipTypes WHERE idFlatShipType="&VaridFlatShipType&";"
                            set rsCustObj=Server.CreateObject("ADODB.RecordSet")
                            set rsCustObj=connTemp.execute(query)
                            %>
                            <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist">
                            <td width="55%">
                                <a href="modFlatShippingRates.asp?refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>"><strong><% =rsCustObj("FlatShipTypeDesc") %></strong></a>
                            </td>
                            <td width="45%" align="right">
                                <span class="cpLinksList">
                                    <%if customServiceActive = "-1" then%>
                                        <a href="modFlatShippingRates.asp?refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>">Edit</a> : <a href="javascript:if (confirm('You are about to permanantly disable (inactivate) this shipping type from the database. Are you sure you want to complete this action?')) location='modFlatShippingRates.asp?mode=InAct&refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>'">Disable (Inactivate)</a> : <a href="javascript:if (confirm('You are about to permanantly delete this shipping type from the database. Are you sure you want to complete this action?')) location='modFlatShippingRates.asp?mode=DEL&refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>'">Remove</a>
                                    <%else%>
                                        <a href="modFlatShippingRates.asp?mode=Act&refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>">Enable (Reactivate)</a>
                                    <%end if%>
                                </span>
                             </td>
                            </tr>
                            <% set rsCustObj=nothing
                        end if
                        rs.movenext
                    loop
                    set rs=nothing
                    
                end if
                %>
            </table>
            </div>
        </div> 
    </div>
    
</div>

<table class="pcCPcontent">
	<tr>
		<td align="center">
			<form class="pcForms" style="margin: 20px;">
				<input type="button" class="btn btn-default"  value="Edit Shipping Settings" onClick="location.href='modFromShipper.asp'">
				&nbsp;<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
			</form>
		 </td>
	</tr>
</table>

<script type=text/javascript>
	$pc( "#acc1" ).accordion({collapsible: true, header: "h5", active:false, heightStyle: "content"});
</script>

<!--#include file="AdminFooter.asp"-->