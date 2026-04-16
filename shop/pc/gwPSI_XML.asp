<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'====================================
'// Turn CVV on - "1"=on, "0"=off
pcv_CVV="1"
'====================================


'//Set redirect page to the current file name
session("redirectPage")="gwPSI_XML.asp"

if session("GWOrderDone")="YES" then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
	session("GWOrderDone")=""
	response.redirect tempURL
end if
		
'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if

'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT Config_File_Name, Userid, [Mode], psi_testmode FROM PSIGate WHERE (((id)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
psi_XMLPassPhrase=rs("Config_File_Name")
psi_XMLStoreID=rs("Userid")
psi_XMLTransType=rs("Mode")
psi_XMLTestmode=rs("psi_testmode")

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	if psi_XMLTestmode="YES" then
		PSiGateURL = "https://dev.psigate.com:27989/Messenger/XMLMessenger"
		psi_XMLPassPhrase="psigate1234"
		psi_XMLStoreID="teststore"
	else
		PSiGateURL = "https://secure.psigate.com:27934/Messenger/XMLMessenger"
	end if
	
	Dim SrvPSiGateXmlHttp, pcPSiGateXMLPostData
	pcPSiGateXMLPostData=""
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<?xml version=""1.0"" encoding=""UTF-8""?>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Order>"
	
					
	'// Check for Discounts that are not compatible with "Itemization"
	query="SELECT orders.discountDetails, orders.pcOrd_CatDiscounts FROM orders WHERE orders.idOrder="&(int(session("GWOrderId"))-scpre)&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rs.eof then
		pcv_strDiscountDetails=rs("discountDetails")
		pcv_CatDiscounts=rs("pcOrd_CatDiscounts")						
	end if
				
	set rs=nothing
	
	pcv_strItemizeOrder = 1

	if pcv_CatDiscounts>0 or trim(pcv_strDiscountDetails)<>"No discounts applied." then
		pcv_strItemizeOrder = 0
	end if
	
	IF pcv_strItemizeOrder = 1 THEN
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Itemized Order
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>
		<!--#include file="pcPay_PSiGate_Itemize.asp"-->
	<%	end if
	
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<StoreID>"&Server.HTMLEncode(psi_XMLStoreID)&"</StoreID>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Passphrase>"&Server.HTMLEncode(psi_XMLPassPhrase)&"</Passphrase>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Tax1>"&Server.HTMLEncode(pcv_strFinalTax)&"</Tax1>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<ShippingTotal>"&Server.HTMLEncode(pcv_strFinalShipCharge)&"</ShippingTotal>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Subtotal>"&Server.HTMLEncode(pcBillingTotal)&"</Subtotal>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<PaymentType>"&Server.HTMLEncode("CC")&"</PaymentType>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardAction>"&Server.HTMLEncode(psi_XMLTransType)&"</CardAction>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardNumber>"&Server.HTMLEncode(Request.Form("Cardnumber"))&"</CardNumber>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardExpMonth>"&Server.HTMLEncode(Request.Form("expMonth"))&"</CardExpMonth>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardExpYear>"&Server.HTMLEncode(Request.Form("expYear"))&"</CardExpYear>"
	If pcv_CVV="1" Then
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardIDCode>1</CardIDCode>"
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardIDNumber>"&Server.HTMLEncode(Request.Form("CVV"))&"</CardIDNumber>"
	end if
	if psi_XMLTestmode="YES" then
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<TestResult>A</TestResult>"
	end if
	if psi_XMLTestmode="YES" then
		pcTestOrderID = Hour(Now) & Minute(Now) & Second(Now)
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<OrderID>"&pcTestOrderID&"PCTEST"&Server.HTMLEncode(session("GWOrderId"))&"</OrderID>"
	else
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<OrderID>"&Server.HTMLEncode(session("GWOrderId"))&"</OrderID>"
	end if
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<UserID>"&Server.HTMLEncode(session("idCustomer"))&"</UserID>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bname>"&Server.HTMLEncode(pcBillingFirstName&" "&pcBillingLastName)&"</Bname>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bcompany>"&Server.HTMLEncode(pcBillingCompany)&"</Bcompany>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Baddress1>"&Server.HTMLEncode(pcBillingAddress)&"</Baddress1>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Baddress2>"&Server.HTMLEncode(pcBillingAddress2)&"</Baddress2>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bcity>"&Server.HTMLEncode(pcBillingCity)&"</Bcity>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bprovince>"&Server.HTMLEncode(pcBillingState)&"</Bprovince>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bpostalcode>"&Server.HTMLEncode(pcBillingPostalCode)&"</Bpostalcode>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bcountry>"&Server.HTMLEncode(pcBillingCountry)&"</Bcountry>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Sname>"&Server.HTMLEncode(pcShippingFirstName&" "&pcShippingLastName)&"</Sname>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Scompany></Scompany>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Saddress1>"&Server.HTMLEncode(pcShippingAddress)&"</Saddress1>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Saddress2>"&Server.HTMLEncode(pcShippingAddress2)&"</Saddress2>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Scity>"&Server.HTMLEncode(pcShippingCity)&"</Scity>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Sprovince>"&Server.HTMLEncode(pcShippingState)&"</Sprovince>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Spostalcode>"&Server.HTMLEncode(pcShippingPostalCode)&"</Spostalcode>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Scountry>"&Server.HTMLEncode(pcShippingCountryCode)&"</Scountry>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Phone>"&Server.HTMLEncode(pcBillingPhone)&"</Phone>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Email>"&Server.HTMLEncode(pcCustomerEmail)&"</Email>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Comments></Comments>"
	if psi_XMLTestmode="YES" then
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CustomerIP>66.249.66.203</CustomerIP>"
	else
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CustomerIP>"&Server.HTMLEncode(pcCustIpAddress)&"</CustomerIP>"
	end if
	
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"</Order>"

	Set SrvPSiGateXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	err.clear
	SrvPSiGateXmlHttp.open "POST", PSiGateURL, false
	SrvPSiGateXmlHttp.send(pcPSiGateXMLPostData)
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	PSiGateResult = SrvPSiGateXmlHttp.responseText

	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
	xmlDoc.async = False
	If xmlDoc.loadXML(SrvPSiGateXmlHttp.responseText) Then
		' Get the results
		pcResultApproved = xmlDoc.documentElement.selectSingleNode("/Result/Approved").Text
		pcResultErrorMsg = xmlDoc.documentElement.selectSingleNode("/Result/ErrMsg").Text
		pcResultTransRefNumber = xmlDoc.documentElement.selectSingleNode("/Result/TransRefNumber").Text
		pcResultCardAuthNumber = xmlDoc.documentElement.selectSingleNode("/Result/CardAuthNumber").Text
		pcResultCardRefNumber = xmlDoc.documentElement.selectSingleNode("/Result/CardRefNumber").Text
	Else
		'//ERROR
		Response.Write "Transaction error or declined.  Error Message: " & pcResultErrorMsg
		response.end
	End If
	If pcResultApproved="APPROVED" then
		session("GWAuthCode")=pcResultCardAuthNumber
		session("GWTransId")=pcResultTransRefNumber
		response.redirect "gwReturn.asp?s=true&gw=PSIGate"
	Else
		if pcResultErrorMsg="" then
			pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"&PSiGateResult
		end if
		Msg=pcResultErrorMsg
	End if

	'*************************************************************************************
	' END
	'*************************************************************************************
end if 
%>
<div id="pcMain">
	<div class="pcMainContent">
				<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="PaymentSubmitted" value="Go">
					<input type="hidden" name="ItemID1" value="Online Order">
					<% If scCompanyName="" then %>
						<input type="hidden" name="Description1" value="Shopping Cart"> 
					<%else %>
						<input type="hidden" name="Description1" value="<%=scCompanyName%>"> 
					<% end if %>
					<input type="hidden" name="Price1" value="<%=pcBillingTotal%>"> 
					<input type="hidden" name="Quantity1" value="1">

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                    
                    <% call pcs_showBillingAddress %>

            <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></div>
                <div class="pcFormField"><input type="text" name="CardNumber" value="" autocomplete="off"></div>
            </div>

					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></div>
						<div class="pcFormField"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="expMonth">
								<option value="01">1</option>
								<option value="02">2</option>
								<option value="03">3</option>
								<option value="04">4</option>
								<option value="05">5</option>
								<option value="06">6</option>
								<option value="07">7</option>
								<option value="08">8</option>
								<option value="09">9</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="12">12</option>
							</select>
							<% dtCurYear=Year(date()) %>
							&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%> 
							<select name="expYear">
								<option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,2)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,2)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,2)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,2)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,2)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,2)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,2)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,2)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,2)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,2)%>"><%=dtCurYear+10%></option>
							</select>
						</div>
					</div>
                    
					<% If pcv_CVV="1" Then %>
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155"></div>
                </div>
					<% End If %>

            <div class="pcFormItem"> 
			    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
                <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
            </div>
					
            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
