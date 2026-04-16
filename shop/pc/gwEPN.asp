<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
'//Set redirect page to the current file name
session("redirectPage")="gwEPN.asp"

'//Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

'//Declare URL path to gwSubmit.asp
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If

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
query="SELECT pcPay_EPN_Account, pcPay_EPN_RestrictKey, pcPay_EPN_CVV, pcPay_EPN_testmode, pcPay_EPN_TranType FROM pcPay_EPN Where pcPay_EPN_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables

pcPay_EPN_Account=rs("pcPay_EPN_Account")
pcPay_EPN_RestrictKey=rs("pcPay_EPN_RestrictKey")
pcv_CVV=rs("pcPay_EPN_CVV")
pcPay_EPN_testmode=rs("pcPay_EPN_testmode")
pcPay_EPN_TranType=rs("pcPay_EPN_TranType")
If IsNull(pcPay_EPN_TranType) Or Len(pcPay_EPN_TranType) < 1 Then
	pcPay_EPN_TranType = "Sale"
End If

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim objXMLHTTP, xml

	sRemoteURL = "https://www.eProcessingNetwork.Com/cgi-bin/tdbe/transact.pl"

	if pcPay_EPN_testmode=1 then
		pcPay_EPN_Account="080880"
		pcPay_EPN_RestrictKey="yFqqXJh9Pqnugfr"
		pcCardNumber="4111111111111111"
	else
		pcCardNumber = request.Form("CardNumber")
		pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)
		pcPay_EPN_RestrictKey=enDeCrypt(pcPay_EPN_RestrictKey, scCrypPass)
	end if

	'Get form variables
	stext="ePNAccount="&pcPay_EPN_Account
	stext=stext & "&RestrictKey="&pcPay_EPN_RestrictKey
	stext=stext & "&TranType="&pcPay_EPN_TranType
	stext=stext & "&CardNo="&pcCardNumber
	stext=stext & "&ExpMonth="&request.Form("expMonth")
	stext=stext & "&ExpYear="&request.Form("expYear")
	stext=stext & "&Total="&pcBillingTotal
	stext=stext & "&Address="&pcBillingAddress
	stext=stext & "&Zip="&pcBillingPostalCode
	stext=stext & "&EMail="&pcCustomerEmail
	stext=stext & "&FirstName="&pcBillingFirstName
	stext=stext & "&LastName="&pcBillingLastName
	if pcv_CVV=1 then
		stext=stext & "&CVV2Type=1"
		stext=stext & "&CVV2="&request.Form("CVV")
	else
		stext=stext & "&CVV2Type=0"
	end if
	stext=stext & "&HTML=No"
	stext=stext & "&Inv="&session("GWOrderID")

	'Create & initialize the XMLHTTP object
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	'Open the connection to the remote server
	xml.Open "POST", sRemoteURL, False
	'Send the request to the eProcessingNetwork Transparent Database Engine
	xml.Send stext

	'store the response
	sResponse = xml.responseText

	'parse the response string and handle appropriately
	sApproval = mid(sResponse, 2, 1)
	
	Dim sFields 
	sFields = Split(sResponse, """")
	sTransID = sFields(9)

	if sApproval = "Y" then
		session("GWTransId")=sTransID
		
		captured = 0
		if pcPay_EPN_TranType = "Sale" then
			captured = 1
		end if		
		
		query="INSERT INTO pcPay_EPN_Authorize (idOrder, amount, transtype, authcode, idCustomer, captured, AuthorizedDate) VALUES ("&pcGatewayDataIdOrder&", "&pcBillingTotal&", '"&pcPay_EPN_TranType&"', '"&session("GWTransId")&"', "&pcIdCustomer&", " & captured & ",'" & now() & "');"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)				
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		
	elseif sApproval = "N" then
		sDeclineReason = "Your transaction has been declined with the " & _
				 "following response: <b>" & mid(sResponse, 3, 16) & "</b><br>"
	else
		sDeclineReason = "The processor was unable to handle your " & _
				 "transaction, having returned the following response: <b>" & _
				 sResponse & "</b><br>"
	end if
	strStatus = xml.Status
	Set xml = Nothing
	

	'save and update order
    call closedb()
	if sApproval = "Y" then
		Response.redirect "gwReturn.asp?s=true&gw=EPN"
	else
        Session("message") = sDeclineReason
        Session("backbuttonURL") = tempURL & "?psslurl=gwEPN.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"        
	End If

	'*************************************************************************************
	' END
	'*************************************************************************************
end if
%>
<div id="pcMain">
	<div class="pcMainContent">
				<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
				<input type="hidden" name="PaymentSubmitted" value="Go">

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
