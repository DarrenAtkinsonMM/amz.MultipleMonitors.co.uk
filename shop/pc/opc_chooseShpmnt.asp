<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 'SB S %>
<!--#include file="inc_sb.asp"-->
<% 'SB E %>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->

<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/FedEXWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/CPconstants.asp"-->

<!--#include file="opc_contentType.asp" -->
<%
dim pcHideEstimateDeliveryTimes
if ( scHideEstimateDeliveryTimes <> "" ) then
	pcHideEstimateDeliveryTimes = scHideEstimateDeliveryTimes
else
	pcHideEstimateDeliveryTimes = "0"
end if


Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

Call SetContentType()

if (request("f")<>"cart") OR (request("ShippingChargeSubmit")="") then
	Call pcs_CheckLoggedIn()
end if

Dim PgType
PgType="2"

Dim pcCartArray
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex" and check to see dbSession was not defined
'*****************************************************************************************************
%>
<!--#include file="pcVerifySession.asp"-->
<%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex" and check to see dbSession was not defined
'*****************************************************************************************************
Sub UpdateNullShipper(tmpvalue)
	query="UPDATE pcCustomerSessions SET pcCustSession_NullShipper='"& tmpvalue &"' WHERE idDbSession="&session("pcSFIdDbSession")&" AND randomKey="&session("pcSFRandomKey")&" AND idCustomer="&session("idCustomer")&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
End Sub

Sub UpdateNullShipRates(tmpvalue)
	query="UPDATE pcCustomerSessions SET pcCustSession_NullShipRates='"& tmpvalue &"' WHERE idDbSession="&session("pcSFIdDbSession")&" AND randomKey="&session("pcSFRandomKey")&" AND idCustomer="&session("idCustomer")&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
End Sub

ppcCartIndex=Session("pcCartIndex")

if request("ShippingChargeSubmit")<>"" then

	pcStrShippingArray=URLDecode(getUserInput(request("Shipping"),0))
	session("pcEstShipping")=pcStrShippingArray
	if (request("f")="cart") AND ((session("idCustomer")="") OR (session("idCustomer")="0")) then
		call closedb()
		response.clear
		Call SetContentType()
		response.write "OK"
		response.end
	end if
	pcIntOrdPackageNumber=URLDecode(getUserInput(request("ordPackageNum"),0))
    If len(pcIntOrdPackageNumber)=0 Then
        pcIntOrdPackageNumber="0"
    End If
	'// If there is any shipping at all, then we should have at least one package.
	if (pcIntOrdPackageNumber="" OR pcIntOrdPackageNumber=0) AND len(pcStrShippingArray)>0 then
		pcIntOrdPackageNumber=1
	end if

	query="UPDATE pcCustomerSessions SET pcCustSession_OrdPackageNumber="&pcIntOrdPackageNumber&", pcCustSession_ShippingArray='"&pcStrShippingArray&"' WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&"));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rs=nothing

	call closedb()
	response.clear
	Call SetContentType()
	response.write "OK"
	response.end
	'response.Redirect("tax.asp")
end if

call pcs_PreCalShipRates()

iShipmentTypeCnt=0
%>
<!--#include file="ShipRates.asp"-->
<%err.number=0
err.description=""

query="SELECT shipService.serviceCode, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation FROM shipService WHERE (((shipService.serviceActive)=-1)) ORDER BY shipService.servicePriority;"

set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if rs.eof then
	call UpdateNullShipper("Yes")
	call closedb()
	response.Clear()
	Call SetContentType()
	response.write "OK|*|<div class='pcSuccessMessage'>" & dictLanguage.Item(Session("language")&"_opc_ship_1") & "</div>"
	response.end
	'response.redirect "tax.asp?idDbSession="& pIdDbSession &"&randomKey="& pRandomKey
else %>
<script type=text/javascript>
function newWindow(file,window) {
		PackageWindow=open(file,window,'resizable=no,width=500,height=600,scrollbars=1');
		if (PackageWindow.opener == null) PackageWindow.opener = self;
}
</script>
<script type="text/javascript">
$(document).ready(function(){
$("#TabbedPanelsShipping").removeClass("hidden-xs");
});
</script>
	<form name="ShipChargeForm" id="ShipChargeForm">
		<div class="pcShowContent">
			<%
			'=============================================================================
			' START optional shipping-related message
			'  - if the feature is on
			'  - if it's setup to show the message at the top
			'=============================================================================
			if PC_SECTIONSHOW="TOP" AND PC_RATESONLY="NO" then %>
				<div class="pcSectionTitle">
					<%=PC_SHIP_DETAIL_TITLE%>
				</div>
				
				<p><%=PC_SHIP_DETAILS%></p>
				
				<div class="pcSpacer"></div>
			<% end if
			'=============================================================================
			' END optional shipping-related message
			'=============================================================================

			'=============================================================================
			' START show package information
			'  - if the feature is on
			'  - if there is more than 1 package
			'=============================================================================
			if scHideProductPackage <> "-1" then
				if pcv_intTotPackageNum>1 then
				%>
					<div class="pcSectionTitle">
						<%=dictLanguage.Item(Session("language")&"_CustviewOrd_38")%>
					</div>
					
					<p><%=ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_h")&pcv_intTotPackageNum&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_i") %></p>
					
					<div class="pcSpacer"></div>
				<%
				end if
			end if
			'=============================================================================
			' END show package information
			'=============================================================================
			%>
      
			<%' load previous entered fields in hidden HTML tags %>
			<input type="hidden" name="ordPackageNum" value="<%=pcv_intTotPackageNum%>">
			<div id="TabbedPanelsShipping">
			<%
			'=============================================================================
			' START shipping provide selection
			' If more than 1 provider is active, ask customer to choose which one to display
			' This feature was introduced to remain compliant with UPS requirements
			'=============================================================================
			dim strTabOrder
			strTabOrder=""

			if iShipmentTypeCnt=>1 then%>
			  <ul id="pcShippingTabs" class="nav nav-tabs">
				<% if instr(strTabShipmentType,"[/TAB][TAB]") then
					strTabShipmentTypeArry=split(strTabShipmentType,"[/TAB]")
					strFirstTab=""
					strTabs=""
					for itab=0 to ubound(strTabShipmentTypeArry)-1
						strTabProvider=replace(strTabShipmentTypeArry(itab),"[TAB]","")
						strTabProviderArry=split(strTabProvider,",")
						if strTabProviderArry(0)=strDefaultProvider then
							strFirstTab="<li class=""active""><a data-toggle=""tab"" href=""#Tab" & strTabProviderArry(0) & """>"&strTabProviderArry(1)&"</a></li>"
							strTabFirst=strTabProviderArry(0)&","
						else
							strTabs=strTabs&"<li><a data-toggle=""tab"" href=""#Tab" & strTabProviderArry(0) & """>"&strTabProviderArry(1)&"</a></li>"
							strTabOrder=strTabOrder&strTabProviderArry(0)&","
						end if
					Next
					strTabOrder=strTabFirst&strTabOrder
				else
					if instr(strTabShipmentType,"[/TAB]") then
						strTabProvider=replace(strTabShipmentType,"[TAB]","")
						strTabProvider=replace(strTabProvider,"[/TAB]","")
						strTabProviderArry=split(strTabProvider,",")
						'strFirstTab="<li class=""active""><a data-toggle=""tab"" href=""#Tab" & strTabProviderArry(0) & """>"&strTabProviderArry(1)&"</a></li>"
						strTabFirst=strTabProviderArry(0)&","
						strTabOrder=strTabFirst
					end if
				end if
				if pcv_boolShowFilteredRates<>"1" then
					response.write strFirstTab
					response.write strTabs
				else
					response.Write "<li class=""active""><a data-toggle=""tab"" href=""#TabShipMap"">" & ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_ja") & "</a></li>"
				end if
				%>
			  </ul>
			<% end if
			'=============================================================================
			' END shipping provider selection
			'=============================================================================

			'=============================================================================
			' START display shipping options
			'=============================================================================
			call pcs_ProcessShipMethods()

			'//Replace last instance of each provider string with an alternate class
			If NOT IsEmpty(strUPS) Then
				strUPSLastInStr = split(strUPS, "<div class='pcTableRow'>")
				strReplace = "<div class='pcTableRow'>"&strUPSLastInStr(Ubound(strUPSLastInStr))
				strNew = "<div class='pcTableRow'>"&strUPSLastInStr(Ubound(strUPSLastInStr))
				strUPS = replace(strUPS, strReplace, strNew)
			End If

			If NOT IsEmpty(strUSPS) Then
				strUSPSLastInStr = split(strUSPS, "<div class='pcTableRow'>")
				strReplace = "<div class='pcTableRow'>"&strUSPSLastInStr(Ubound(strUSPSLastInStr))
				strNew = "<div class='pcTableRow'>"&strUSPSLastInStr(Ubound(strUSPSLastInStr))
				strUSPS = replace(strUSPS, strReplace, strNew)
			End If

			If NOT IsEmpty(strFedEx) Then
				strFedExLastInStr = split(strFedEx, "<div class='pcTableRow'>")
				strReplace = "<div class='pcTableRow'>"&strFedExLastInStr(Ubound(strFedExLastInStr))
				strNew = "<div class='pcTableRow'>"&strFedExLastInStr(Ubound(strFedExLastInStr))
				strFedEx = replace(strFedEx, strReplace, strNew)
			End If

			If NOT IsEmpty(strCP) Then
				strCPLastInStr = split(strCP, "<div class='pcTableRow'>")
				strReplace = "<div class='pcTableRow'>"&strCPLastInStr(Ubound(strCPLastInStr))
				strNew = "<div class='pcTableRow'>"&strCPLastInStr(Ubound(strCPLastInStr))
				strCP = replace(strCP, strReplace, strNew)
			End If

			If NOT IsEmpty(strCUSTOM) Then
				strCUSTOMLastInStr = split(strCUSTOM, "<div class='pcTableRow'>")
				strReplace = "<div class='pcTableRow'>"&strCUSTOMLastInStr(Ubound(strCUSTOMLastInStr))
				strNew = "<div class='pcTableRow'>"&strCUSTOMLastInStr(Ubound(strCUSTOMLastInStr))
				strCUSTOM = replace(strCUSTOM, strReplace, strNew)
			End If

			'//ENSURE THERE IS AT LEAST ONE OPTION CHECKED - Can happen if no rates are returned by the default provider that is set in the CP
			if pcv_Default=0 then
				strFEDEX=replace(strFEDEX,"XCHECK","")
				strUSPS=replace(strUSPS,"XCHECK","")
				strUPS=replace(strUPS,"XCHECK","")
				strCP=replace(strCP,"XCHECK","")
				strCUSTOM=replace(strCUSTOM,"XCHECK","")
			else
				strFEDEX=replace(strFEDEX,"XCHECK","")
				strUSPS=replace(strUSPS,"XCHECK","")
				strUPS=replace(strUPS,"XCHECK","")
				strCP=replace(strCP,"XCHECK","")
				strCUSTOM=replace(strCUSTOM,"XCHECK","")
				strFEDEX=replace(strFEDEX,"FCHECK","")
				strUSPS=replace(strUSPS,"FCHECK","")
				strUPS=replace(strUPS,"FCHECK","")
				strCP=replace(strCP,"FCHECK","")
				strCUSTOM=replace(strCUSTOM,"FCHECK","")
			end if
			inttotalUPSWeight=0
			for uCnt=1 to pcv_intPackageNum
				intTotalUPSWeight=intTotalUPSWeight+session("UPSPackWeight"&uCnt)
			next
			%>

			<%
			strContent=""
			
			strProviderHeader=""
			strProviderHeader=strProviderHeader&"<div class='pcTable pcShipRates'>"
			strProviderHeader=strProviderHeader&"<div class='pcTableHeader'>"
			
			if (pcHideEstimateDeliveryTimes <> "-1") then
				strProviderHeader=strProviderHeader&"<div class='pcShip_ServiceType'>Delivery Option</div>"
				strProviderHeader=strProviderHeader&"<div class='pcShip_DeliveryTime'>Delivery Estimate</div>"
			else
				strProviderHeader=strProviderHeader&"<div class='pcShip_ServiceTypeL'>&nbsp;"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_a")&"</div>"
			end if
			
			strProviderHeader=strProviderHeader&"<div class='pcShip_Rate'>Cost</div>"
			strProviderHeader=strProviderHeader&"<div class='pcSpacer'>&nbsp;</div>"
			strProviderHeader=strProviderHeader&"</div>"

			strUPSFooter="<div class='pcSpacer'></div>"
			strUPSFooter=strUPSFooter&"<div id='pcUPSFooter'>"
			  strUPSFooter=strUPSFooter&"<a href=""javascript:;"" onclick=""newWindow('pcUPSTimeInTransit.asp?sResidential="&pResidentialShipping&"&sPackageCnt="&pcv_intTotPackageNum&"&sWeight="&pShipWeight&"&sState="&universal_destination_provOrstate&"&sCity="&universal_destination_city&"&sPC="&universal_destination_postal&"&sCountry="&universal_destination_country&"','ProductWindow')"">Time In Transit</a>:"
			  strUPSFooter=strUPSFooter&"&nbsp;Calculate estimated transit time for the various UPS services.<hr>"
			  strUPSFooter=strUPSFooter&"<div class='pcUPSLogo'><img src='" & pcf_getImagePath("../UPSLicense","LOGO_S2.png") & "' style='width: 42px; height: 50px'></div>"
			  strUPSFooter=strUPSFooter&"<div class='pcUPSTerms'><p><b>UPS&reg; Developer Kit Rates & Service Selection</b></p><p>Notice: UPS fees do not necessarily represent UPS published rates and may include charges levied by the store owner.</p><br/><p class='pcSmallText'>UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved.</p></div>"
			strUPSFooter=strUPSFooter&"</div>"

			%>
				<% 
                strContent=strContent&"<div class=""tab-content"">"
                pcv_strActiveTab = "in active"
                strTabOrderArry=split(strTabOrder,",")
				for iContent=0 to ubound(strTabOrderArry)-1
					select case ucase(strTabOrderArry(iContent))
					case "USPS"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK"," checked")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div id=""Tab" & strTabOrderArry(iContent) & """ class=""tab-pane fade " & pcv_strActiveTab & " "">"&strProviderHeader&strUSPS&"</div>"
						strContent=strContent&"<div class='pcClear'></div>"
						strContent=strContent&"</div>"
                        pcv_strActiveTab=""
					case "CP"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK"," checked")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div id=""Tab" & strTabOrderArry(iContent) & """ class=""tab-pane fade " & pcv_strActiveTab & " "">"&strProviderHeader&strCP&"</div>"
						strContent=strContent&"<div class='pcClear'></div>"
						strContent=strContent&"</div>"
                        pcv_strActiveTab=""
					case "FEDEXWS"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK"," checked")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div id=""Tab" & strTabOrderArry(iContent) & """ class=""tab-pane fade " & pcv_strActiveTab & " "">"
						strContent=strContent&strProviderHeader&strFEDEX&"</div>"
		        strContent=strContent&"<div class='pcSpacer'></div>"
						strContent=strContent&"<div class='pcFormItem'><p class='pcSmallText'>FedEx service marks are owned by Federal Express Corporation and used with permission.</p></div>"				
		        strContent=strContent&"<div class='pcClear'></div>"
						strContent=strContent&"</div>"
                        pcv_strActiveTab=""
					case "UPS"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK"," checked")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div id=""Tab" & strTabOrderArry(iContent) & """ class=""tab-pane fade " & pcv_strActiveTab & " "">"&strProviderHeader&strUPS&strUPSFooter&"</div>"
						strContent=strContent&"<div class='pcClear'></div>"
						strContent=strContent&"</div>"
                        pcv_strActiveTab=""
					case "CUSTOM"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK"," checked")
						end if
						strContent=strContent&"<div id=""Tab" & strTabOrderArry(iContent) & """ class=""tab-pane fade " & pcv_strActiveTab & " "">"&strProviderHeader&strCUSTOM&"</div>"				
		        strContent=strContent&"<div class='pcClear'></div>"
						strContent=strContent&"</div>"
                        pcv_strActiveTab=""
					end select
				next 
                
                strContent=strContent&"</div>"
                %>
								
				<%if pcv_boolShowFilteredRates="1" then
					call pcs_MapShip()
					
					strContent=""
					
					For iM=0 to MCount
						if tMArr(0,iM)=1 then
							strContent=strContent&"<div class='pcTableRow'>"
							if (tMArr(6,iM)=1) OR ((HasDefaultM=0) AND (iM=0)) then
								x_checked="checked"
							else
								x_checked=""
							end if
								
							strContent=strContent&"<div class='" & serviceTypeClass & "'><input type='radio' name='Shipping' value='"&tMArr(3,iM)&"' class='clearBorder'"&x_checked&">&nbsp;"&tMArr(2,iM)&"</div>"
							if showDeliveryCol then
								strContent=strContent&"<div class='" & deliveryTimeClass & "'>"&tMArr(7,iM)&"</div>"
							end if
							strContent=strContent&"<div class='" & rateClass & "'>"&scCurSign&money(tMArr(4,iM))&"</div>"
							strContent=strContent&"</div>"
						end if
					Next
					
					strContent="<div class=""tab-content"">"&"<div id=""TabShipMap"" class=""tab-pane fade in active "">"&strProviderHeader&strContent&"</div>"				
		        	strContent=strContent&"<div class='pcClear'></div>"
					strContent=strContent&"</div>"
				end if%>
				
				<% response.write strContent %>

			</div>
      
			<% call UpdateNullShipRates("No")
			dim intCRates
			intCRates=0
			If DCnt=0 then
				if scAlwNoShipRates="-1" then
					call UpdateNullShipRates("Yes")
					ShowJSSubmitValidation="1"
					'=============================================================================
					' START show messages about no shipping options available
					' No shipping rates and checkout allowed
					'=============================================================================
					%>
          <p>
          	<% 
							intCRates=1
							response.write ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_g")
						%>
          </p>
          
          
          
          <div class="pcSpacer"></div>
          
          <button class="pcButton pcButtonContinue btn btn-skin btn-wc updateBtn opcShipCntBut" name="ShippingChargeSubmit" id="ShippingChargeSubmit" data-ng-click="updateShippingMethod()">
            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update")%></span>
          </button>
              
				<% else
					'=============================================================================
					' No shipping rates and checkout NOT allowed
					'=============================================================================
					response.Clear()
					Call SetContentType()
					response.write "STOP|*|" & dictLanguage.Item(Session("language")&"_opc_ship_2")
					response.end
				end if
				'=============================================================================
				' END show messages about no shipping options available
				'=============================================================================
			else
			ShowJSSubmitValidation="1" %>
        
        
        <div class="pcSpacer">&nbsp;</div>
          <div class="daOPCShipMsg"><span class="daOPCShipMsgHead">Delivery Notes:</span><ul><li>Delivery dates are estimates only, however are usually very accurate</li><li>If you require a specific delivery date simply email us after placing your order to confirm this</li><li>Full delivery tracking information is emailed through once we dispatch orders</li><li>UK deliveries should receive a 1 hour delivery window via email / txt on the morning of delivery</li><li>International shipping times can sometimes vary by a day or two due to courier or custom delays</li></ul></div>
        
        <button class="pcButton pcButtonContinue btn btn-skin btn-wc updateBtn opcShipCntBut" name="ShippingChargeSubmit" id="ShippingChargeSubmit" data-ng-click="updateShippingMethod()">
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update")%></span>
        </button>
			<% end if %>

			<%if ShowJSSubmitValidation="1" then%>

			<%end if%>


			<%'=============================================================================
			' END display shipping options
			'=============================================================================

			'=============================================================================
			' START optional shipping-related message
			'  - if the feature is on
			'  - if it's setup to show the message at the bottom
			'=============================================================================
			dim intDisplay
			intDisplay=0
			if PC_SECTIONSHOW="BTM" then
				intDisplay=1
				if PC_RATESONLY="YES" then
					if intCRates=1 then
						intDisplay=0
					end if
				end if
			end if
			if intDisplay=1 then
				%>
        <div class="pcShowContent">
        	<div class="pcSectionTitle"><%=PC_SHIP_DETAIL_TITLE%></div>
          <p><%=PC_SHIP_DETAILS%></p>
        </div>
				<%
			end if
			'=============================================================================
			' END optional shipping-related message
			'=============================================================================
			%>
	</div>
	</form>
<% end if %>
<% call closeDb()
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing %>
