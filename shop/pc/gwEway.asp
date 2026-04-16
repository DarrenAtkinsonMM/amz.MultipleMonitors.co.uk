<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwEway.asp"

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

pcv_BeagleNotAvailable=0

'//Check if Beagle Field Exists
on error resume next
err.clear
query="SELECT * FROM eWay;"
set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query)
eWay_BeagleActive=rstemp("eWayBeagleActive")
if err.number<>0 then
	pcv_BeagleNotAvailable=0
else
	pcv_BeagleNotAvailable=1
end if
set rstemp=nothing

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT eWayCustomerid, eWayPostMethod, eWayTestmode, eWayCVV"
if pcv_BeagleNotAvailable=1 then
	query=query&", eWayBeagleActive"
end if
query=query&" FROM eWay WHERE eWayID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcEwayCustomerid=rs("eWayCustomerid")
pcEwayPostMethod=rs("eWayPostMethod")
pcEwayPostMethod="XML"
pcEwayTestmode=rs("eWayTestmode")
pcEwayBillingTotal=pcBillingTotal
if pcEwayTestmode=1 then
	pcEwayCustomerid="87654321"
	pcEwayBillingTotal="10.00"
end if
pcEwayCVV = rs("eWayCVV")
if pcv_BeagleNotAvailable=1 then
	pcEwayBeagleActive = rs("eWayBeagleActive")
else
	pcEwayBeagleActive="0"
end if
pcEwayBillingTotal = replacecomma(pcEwayBillingTotal)
pcEwayBillingTotal = (pcEwayBillingTotal*100)

set rs=nothing

if request("PaymentSubmitted")="Go" then
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	if pcEwayTestmode=1 then
	   if pcEwayCVV ="1" Then
	   		pcEwayURL = "https://www.eway.com.au/gateway_cvn/xmltest/testpage.asp"
	   Else
			pcEwayURL = "https://www.eway.com.au/gateway/xmltest/TestPage.asp"
	   End if 	
	else
	   if pcEwayCVV ="1" Then
	   		pcEwayURL = "https://www.eway.com.au/gateway_cvn/xmlpayment.asp"
			if pcEwayBeagleActive = "1" then
				pcEwayURL = "http://www.eway.com.au/gateway_cvn/xmlbeagle.asp"
			end if
	   Else
			pcEwayURL = "https://www.eway.com.au/gateway/xmlpayment.asp"
			if pcEwayBeagleActive = "1" then
				pcEwayURL = "http://www.eway.com.au/gateway_cvn/xmlbeagle.asp"
			end if
	   end if		
	end if

	Dim SrvEWayXmlHttp, pcEwayXMLPostData
	pcEwayXMLPostData=""
	pcEwayXMLPostData=pcEwayXMLPostData&"<?xml version=""1.0"" encoding=""UTF-8""?>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewaygateway>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerID>"&pcEwayCustomerid&"</ewayCustomerID>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayTotalAmount>"&pcEwayBillingTotal&"</ewayTotalAmount>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerFirstName>"&pcBillingFirstName&"</ewayCustomerFirstName>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerLastName>"&pcBillingLastName&"</ewayCustomerLastName>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerEmail>"&pcCustomerEmail&"</ewayCustomerEmail>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerAddress>"&pcBillingAddress&"</ewayCustomerAddress>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerPostcode>"&pcBillingPostalCode&"</ewayCustomerPostcode>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerInvoiceDescription>Online Order</ewayCustomerInvoiceDescription>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerInvoiceRef>"&session("GWOrderId")&"</ewayCustomerInvoiceRef>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardHoldersName>"&pcBillingFirstName&" "&pcBillingLastName&"</ewayCardHoldersName>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardNumber>"&request("CardNumber")&"</ewayCardNumber>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardExpiryMonth>"&request("expMonth")&"</ewayCardExpiryMonth>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardExpiryYear>"&request("expYear")&"</ewayCardExpiryYear>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayTrxnNumber>"&session("GWOrderId")&"</ewayTrxnNumber>"
	
	if pcEwayCVV ="1" Then
		pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCVN>"&request.form("CVV")&"</ewayCVN>"
	End if 
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayOption1></ewayOption1>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayOption2></ewayOption2>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayOption3></ewayOption3>"
	'//eWay's Beagle Fraud Prevention
	if pcEwayBeagleActive = "1" then
		pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerIPAddress>"&xxx&"</ewayCustomerIPAddress>" 
		pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerBillingCountry>"&xxx&"</ewayCustomerBillingCountry>"
	end if 

	pcEwayXMLPostData=pcEwayXMLPostData&"</ewaygateway>"
	'response.write pcEwayXMLPostData&"<HR>"
	
	Set SrvEWayXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	err.clear
	SrvEWayXmlHttp.open "POST", pcEwayURL, false
	SrvEWayXmlHttp.send(pcEwayXMLPostData)
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	EWayResult = SrvEWayXmlHttp.responseText
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
	xmlDoc.async = False
	If xmlDoc.loadXML(SrvEWayXmlHttp.responseText) Then
		' Get the results
		pcResultEwayTrxnStatus = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnStatus").Text
		pcResultEwayTrxnNumber = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnNumber").Text
		pcResultEwayTrxnOption1 = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnOption1").Text
		pcResultEwayTrxnOption2 = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnOption2").Text
		pcResultEwayTrxnOption3 = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnOption3").Text
		pcResultEwayTrxnReference = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnReference").Text
		pcResultEwayAuthCode = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayAuthCode ").Text
		pcResultEwayReturnAmount = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayReturnAmount ").Text
		pcResultEwayTrxnError = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnError").Text
	Else
		'//ERROR
		response.write "Failed to process response"
		response.end
	End If

	if ucase(pcResultEwayTrxnStatus)="TRUE" then
		session("GWAuthCode")=pcResultEwayTrxnReference
		session("GWTransId")=pcResultEwayTrxnNumber
		Set eWay = Nothing
        call closeDb()
		Response.redirect "gwReturn.asp?s=true&gw=eWay"
	else
		Set eWay = Nothing
        call closeDb()
        Session("message") = pcResultEwayTrxnError
        Session("backbuttonURL") = tempURL & "?psslurl=gwEway.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"        
	end if

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
                    
					<% If pcEwayCVV="1" Then %>
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
