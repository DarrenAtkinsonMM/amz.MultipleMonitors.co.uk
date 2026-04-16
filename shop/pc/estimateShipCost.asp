
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->

<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/FedEXWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/CPconstants.asp"-->

<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->

<%
pcStrPageName = "estimateShipCost.asp"

Dim pcCartArray, ppcCartIndex

Dim PgType
PgType="1"

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

'// Load validation resources
pcv_strRequiredIcon = rsIconObj("requiredicon")
pcv_strErrorIcon = rsIconObj("errorfieldicon")
%>

    <div class="col-xs-12">

        <%
        dim pcHideEstimateDeliveryTimes
        
        If scHideEstimateDeliveryTimes <> "" Then
            pcHideEstimateDeliveryTimes = scHideEstimateDeliveryTimes
        Else
            pcHideEstimateDeliveryTimes = "0"
        End If
		
		pcv_isZipRequired = False
        pcv_isCityRequired = False
        pcv_isShipCountryCodeRequired=False

        '// Use the Request object to toggle State (based of Country selection)
        If UPS_USENEGOTIATEDRATES = 1 Then
            pcv_isShipStateCodeRequired=True
        Else
            pcv_isShipStateCodeRequired=False
        End If
        
        '// Use the Request object to toggle Province (based of Country selection)
        pcv_isShipProvinceCodeRequired=False
		
		queryQ="SELECT [idshipservice],[serviceActive],[serviceCode],[serviceDescription],[idShipment] FROM [shipService] WHERE serviceActive=-1;"
		set rsQ=connTemp.execute(queryQ)
		if not rsQ.eof then
			tmpArrQ=rsQ.getRows()
			set rsQ=nothing
			intQ=ubound(tmpArrQ,2)
			tmpShipServ=""
			For iQ=0 to intQ
				tmpVal=""
				if InStr(tmpArrQ(3,iQ),"UPS ")>0 then
					tmpVal="UPS"
				else
					if InStr(tmpArrQ(3,iQ),"FedEx ")>0 then
						tmpVal="FEDEX"
					else
						if InStr(tmpArrQ(3,iQ),"Canada Post ")>0 then
							tmpVal="CP"
						else
							if Left(tmpArrQ(2,iQ),1)="C" then
								tmpVal="CUSTOM"
							else
                                If IsNumeric(tmpArrQ(2,iQ)) Then
                                    If (CLng(tmpArrQ(2,iQ))>=9900) AND (CLng(tmpArrQ(2,iQ))<=9920) then
                                        tmpVal="USPS"
                                    End if
                                End If
							end if
						end if
					end if
				end if
				if tmpVal<>"" then
					if tmpShipServ="" then
						tmpShipServ="***"
					end if
					tmpShipServ=tmpShipServ & tmpVal & "***"
				end if
			Next		
			if Instr(tmpShipServ,"***UPS***")>0 then
				pcv_isZipRequired = true
				pcv_isShipCountryCodeRequired=true
				pcv_isShipStateCodeRequired=true
			end if
			if Instr(tmpShipServ,"***FEDEX***")>0 then
				pcv_isZipRequired = true
				pcv_isShipCountryCodeRequired=true
				pcv_isShipStateCodeRequired=true
			end if
			if Instr(tmpShipServ,"***USPS***")>0 then
				pcv_isZipRequired = true
				pcv_isShipCountryCodeRequired=true
			end if
			if Instr(tmpShipServ,"***CP***")>0 then
				pcv_isZipRequired = true
				pcv_isShipCountryCodeRequired=true
			end if
			if Instr(tmpShipServ,"***CUSTOM***")>0 then
				pcv_isShipCountryCodeRequired=true
			end if
		end if
		set rsQ=nothing

        '// Use the Request object to toggle State (based of Country selection)
        pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
       
        If  len(pcv_strStateCodeRequired)>0 Then
            pcv_isShipStateCodeRequired = pcv_strStateCodeRequired
        End If

        '// Use the Request object to toggle Province (based of Country selection)
        pcv_isShipProvinceCodeRequired=False
        pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
        
        If len(pcv_strProvinceCodeRequired) > 0 Then
            pcv_isShipProvinceCodeRequired=pcv_strProvinceCodeRequired
        End If

        IF request("zip")="" AND request("ddjumpflag")="" then

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' START: Config Client-Side Validation
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            response.write "<script type=text/javascript>"&vbcrlf
                
            response.write "function FormEst_Validator()"&vbcrlf
            response.write "{theForm=document.getElementById('shipCost');"&vbcrlf
            if pcv_iszipRequired then
                response.write "if (theForm.zip.value=='') return(1);"
            end if
            if pcv_isShipStateCodeRequired then
                response.write "if ((theForm.StateCode.value=='') && (theForm.Province.value=='')) return(2);"
            end if
            response.write " return (3);"&vbcrlf
            response.write "}"&vbcrlf
            
            response.write "</script>"&vbcrlf
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' END: Config Client-Side Validation
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            %>
            
            <form action="<%=pcStrPageName%>" method="post" name="shipCost" id="shipCost" class="form">
                    <div class="row">
                        <div class="col-xs-10">  
                            <% 
                            err.clear
                            msg = ""
                            code=getUserInput(Request.QueryString("msg"),0)
                            
                            If code = "1" Then
                                msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
                            End If
                            
                            If msg<>"" Then 
                                %><div class="pcErrorMessage"><%= msg %></div><%
                            End If
                            %>
        
                            <%
                            '///////////////////////////////////////////////////////////
                            '// START: COUNTRY AND STATE/ PROVINCE CONFIG
                            '///////////////////////////////////////////////////////////
                            
                            '// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
                            pcv_isStateCodeRequired = pcv_isShipStateCodeRequired '// determines if validation is performed (true or false)
                            pcv_isProvinceCodeRequired = pcv_isShipProvinceCodeRequired '// determines if validation is performed (true or false)
                            pcv_isCountryCodeRequired = pcv_isShipCountryCodeRequired '// determines if validation is performed (true or false)
                            
                            '// Required Info
                            pcv_strTargetForm = "shipCost" '// Name of Form
                            pcv_strCountryBox = "CountryCode" '// Name of Country Dropdown
                            pcv_strTargetBox = "StateCode" '// Name of State Dropdown
                            pcv_strProvinceBox =  "Province" '// Name of Province Field
                            
                            '// Set local Country to Session
                            if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                                Session(pcv_strSessionPrefix&pcv_strCountryBox) = CountryCode
                            end if
                            
                            '// Set local State to Session
                            if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                                Session(pcv_strSessionPrefix&pcv_strTargetBox) = StateCode
                            end if
                            
                            '// Set local Province to Session
                            if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                                Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Province
                            end if
                            %>
                            <!--#include file="../includes/javascripts/opc_pcStateAndProvince.asp"-->
                            <%
                            '///////////////////////////////////////////////////////////
                            '// END: COUNTRY AND STATE/ PROVINCE CONFIG
                            '///////////////////////////////////////////////////////////
                            %>
                            <%
                            '// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince5.asp)
                            pcs_CountryDropdown
                            %>
                        </div>
                    </div>
                    
                    <%
                    If (UPS_USENEGOTIATEDRATES = 1) OR (pcv_isShipStateCodeRequired) Then
                    %>
                    <div class="row">
                        <div class="col-xs-10">
                    		<%
                    		'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince5.asp)
                    		pcs_StateProvince
                    		%>
                        </div>
                    </div>
                    <%
                    End If
                    %>
                    
                    <style>
                        #LabelCountryCode {
                            display: none;
                        }
                    </style>
					
                    
                    <div class="row">
                        <div class="col-xs-10">
                            <input name="zip" id="zip" class="form-control" placeholder="<%= dictLanguage.Item(Session("language")&"_vShipAdd_5")%>" type="text" size="10" value="<%=pcf_FillFormField ("zip", pcv_isZipRequired) %>"><% pcs_RequiredImageTagHorizontal "zip", pcv_isZipRequired %>
                            <div id="ErrorZipBox" class="pcErrorMessage" style="display: none"></div>
                        </div>
                    </div>
                
                    <%
                    '// ProductCart v4.5 - Commercial vs. Residential
                    Dim pcComResShipAddress
                    if scComResShipAddress = "0" then
                    %>

                    <div class="row">
                        <div class="col-xs-10">
                            <div class="radio-wrapper">
                                <label class="radio-inline">
                                    <input type="radio" name="residentialShipping" value="-1" checked class="clearBorder"> <%= ship_dictLanguage.Item(Session("language")&"_login_c")%>
                                </label>
                                <label class="radio-inline">
                                    <input type="radio" name="residentialShipping" value="0" class="clearBorder"> <%= ship_dictLanguage.Item(Session("language")&"_login_d")%>
                                </label>
                            </div>
                        </div>
                    </div>
                
                    <%
                    else
                        Select Case scComResShipAddress
                            Case "1"
                                pcComResShipAddress="-1"
                            Case "2"
                                pcComResShipAddress="0"
                            Case "3"
                                if session("customerType")="1" then
                                    pcComResShipAddress="0"
                                else
                                    pcComResShipAddress="-1"
                                end if
                        End Select
                        %>
                        <div class="pcFormItem">
                            <div class="pcSpacer"><input type="hidden" name="residentialShipping" value="<%=pcComResShipAddress%>"></div>
                        </div>
                        <%
                    end if
                    %>
                    <div class="pcFormItem">
                        <div class="pcFormButtons">
                            <button class="btn btn-default" type="button" data-ng-click="getEstShipServices();" name="SubmitShip" id="SubmitShip">Estimate</button>
                        </div>
                    </div>

             
                <input type="hidden" name="ddjumpflag" value="YES">
            </form>

        <% ELSE

				call pcs_PreCalShipRates()
				
				if ((session("availableShipStr")="" or session("provider")="") OR request("ddjumpflag")="") AND pCartShipQuantity>0 then %>
					<!--#include file="ShipRates.asp"-->
                    <% session("strDefaultProvider")=strDefaultProvider
                    session("iShipmentTypeCnt")=iShipmentTypeCnt
                    session("strOptionShipmentType")=strOptionShipmentType
                    session("availableShipStr")=availableShipStr
                    session("iUPSFlag")=iUPSFlag
                end if
    
                availableShipStr=session("availableShipStr")
				
                query="SELECT shipService.serviceCode, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation FROM shipService WHERE (((shipService.serviceActive)=-1)) ORDER BY shipService.servicePriority;"
                set rs=Server.CreateObject("ADODB.RecordSet")
                set rs=connTemp.execute(query)
                if rs.eof then
                	Session("nullShipper")="Yes"
                else %>
                
 
                    <%
                    'if count shows that more then 1 shipmentType is active, show customer choice
					if (pcv_boolShowFilteredRates<>"1") then
						if session("iShipmentTypeCnt")>1 then %>
						<div class="pcFormItem">
						  <div class="pcFormItemFull">
	
							<form action="<%=pcStrPageName%>" method="post" id="shipCost" data-target="#QuickViewDialog" class="form">
							  
							  <%
							  pcv_strDefaultProvider = session("strDefaultProvider")
							  if session("provider")<>"" then
								pcv_strDefaultProvider = session("provider")
							  end if
							  %>
							  <select data-ng-init="provider='<%=pcv_strDefaultProvider%>'" data-ng-change="getEstShipServicesS();" data-ng-model="provider" id="provider" name="provider" class="form-control">
							  <% strTempOptionShipmentType=session("strOptionShipmentType")
							  if session("provider")<>"" then
								strTempOptionShipmentType=replace(strTempOptionShipmentType,"value="&session("provider")&"","value="&session("provider")&" selected")
							  else
								strTempOptionShipmentType=replace(strTempOptionShipmentType,"value="&session("strDefaultProvider")&"","value="&session("strDefaultProvider")&" selected")
								session("provider")=session("strDefaultProvider")
							  end if %>
							  <%=strTempOptionShipmentType%>
							  </select>
							  
							  <input type="hidden" name="ddjumpflag" value="YES">
							  <input type="hidden" name="CountryCode" value="<%=Session("pcSFCountryCode")%>">
							  <input type="hidden" name="StateCode" value="<%=Session("pcSFStateCode")%>">
							  <input type="hidden" name="Province" value="<%=Session("pcSFProvince")%>">
							  <input type="hidden" name="city" value="<%=Session("pcSFcity")%>">
							  <input type="hidden" name="zip" value="<%=Session("pcSFzip")%>">
							  <input type="hidden" name="residentialShipping" value="<%=request("residentialShipping")%>">
							</form>
						  </div>
						</div>
											<div class="pcSpacer"></div>
					  <% else
						  session("provider")=session("strDefaultProvider")%>
						<% end if
					else
						session("provider")=""
					end if %>
					<%tSCount=0%>
                    
                    <form action="<%=pcStrPageName%>" id="ShipChargeForm" name="ShipChargeForm" data-target="#QuickViewDialog" method="get"> 
                    
                    <div class="pcRatesSummaryWrapper">
    
                        <% 
                        col_ServiceTypeClass = "col-xs-5"
                        col_DeliveryTimeClass = "col-xs-5"
                        col_RateClass = "col-xs-2"
                        %>
                        <div class="pcRatesSummary">
						
							<%call pcs_ProcessShipMethods()
                            %>
                            <%
                            if pcv_boolShowFilteredRates="1" then
								call pcs_MapShip()
								
								tmpHaveSMp=0
								For iM=0 to MCount
									if tMArr(0,iM)=1 then
									tmpHaveSMp=tmpHaveSMp+1%>
										<%if (tmpHaveSMp="1") then%>
										  <div class="row">
											<div class="<%= col_ServiceTypeClass %>"><%= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_a")%></div>
											<div class="col-xs-5"></div> 
											<div class="<%= col_RateClass %>"><%= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_c")%></div>
										  </div>
										<%end if%>
										<div class="row">
											<div class="<%= col_ServiceTypeClass %>">
												<label class="radio-inline" for="pcShipSelection">
													<input type="radio" name="Shipping" value="<%=tMArr(3,iM)%>" class="clearBorder" <%if tMArr(3,iM)=session("pcEstShipping") then%>checked<%end if%> data-ng-click="updateShippingMethod()"><%=tMArr(2,iM)%>
												</label>
											</div>
											<div class="col-xs-5"></div>
											<div class="<%= col_RateClass %>"><%=scCurSign&money(tMArr(4,iM))%></div>
										</div>
									<%end if
								Next
								
								if tmpHaveSMp=0 then%>
									<div class="row">
									  <div class="col-xs-12">
										<%= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_d")%>
									  </div>
									</div>
								<%end if
								
							end if%>
							<% if CntFree>0 then %>
                                <div class="row">
                                  <div class="col-xs-12">
                                    <%=ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_e")%>
                                  </div>
                                </div>
                            <% end if %>
                            <% if session("iUPSFlag")=1 AND ucase(session("provider"))="UPS"then %>
                                <div class="row">
                                    <hr />
                                </div>
                                <div class="row">
                                    <div class="col-xs-12">
                                        <div style="float: left; width: 15%"><img src="<%=pcf_getImagePath("../UPSLicense","LOGO_S2.jpg")%>" style="width: 45px; height: 50px"></div>
                                        <div style="float: left; width: 85%" class="pcSmallText">
                                            <p><strong>UPS&reg; Developer Kit Rates &amp; Service Selection</strong></p>
                                            <p>Notice: UPS fees do not necessarily represent UPS published rates  and may include charges levied by the store owner.</p>
                                            <p>UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved.</p>
                                        </div>
                                    </div>
                                </div>
							<% end if %>
                            <% If pCartShipQuantity=0 then %>
								<div class="row">
                                  <div class="col-xs-12">
                                    <p><%= dictLanguage.Item(Session("language")&"_opc_ship_1")%></p>
                                  </div>
                                </div>
							<% ElseIf (DCnt=0) AND (pcv_boolShowFilteredRates<>"1") then %>
                                <div class="row">
                                  <div class="col-xs-12">
                                    <p><%= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_d")%></p>
                                  </div>
                                </div>
                        	<% end if %>
                        </div>
    
                    
                    </div>

                    <button class="btn btn-default" type="button" data-ng-click="showEstShip()" name="SubmitShip" id="SubmitShip">Change Postal Code</button>
									
                </form>
                <% end if %>
							<%
              'End If '// If pcv_intErr>0 Then

        END IF '// IF request("SubmitShip")="" AND request("ddjumpflag")="" then
        %>
    </div>
   

<%
call closeDb()
%>
