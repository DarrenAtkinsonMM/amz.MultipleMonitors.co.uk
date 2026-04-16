<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwTCLink.asp"

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
query="SELECT TCLinkid, TCLinkPassword, TCTestmode, TCCurcode, CVV, avs, TranType FROM tclink Where idTCLink=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_TCLinkid=rs("TCLinkid")
pcv_TCLinkPassword=rs("TCLinkPassword")
'decrypt
pcv_TCLinkPassword=enDeCrypt(pcv_TCLinkPassword, scCrypPass)
pcv_TestMode=rs("TCTestmode")
pcv_CurCode=rs("TCCurcode")
pcv_CVV=rs("CVV")
pcv_AVS=rs("avs")
pcv_TransType=rs("TranType")
				
If pcv_TestMode =1 Then
	pcv_Demo="y"
Else ' Live Mode
	pcv_Demo=""   
End If   
			
If pcv_AVS ="1" Then
	pcv_AVS="y"
Else 
	pcv_AVS="n"   
End If
			
If pcv_CVV ="1" Then
	pcv_CheckCVV ="y"
Else 
	pcv_CheckCVV ="n"   
End If

set rs=nothing

if request("PaymentSubmitted")="Go" then

'*************************************************************************************
' This is where you would post info to the gateway
' START
'*************************************************************************************
		strPostUrl="https://vault.trustcommerce.com/trans"
		
		strRequest = "?custid="&pcv_TCLinkid&_
		"&password="&pcv_TCLinkPassword&_
		"&action="&pcv_TransType&_
		"&media="&"cc"&_
		"&demo="&pcv_Demo&_
		"&amount="&pcBillingTotal*100&_   
		"&cc="&request.Form("CardNumber")&_ 
		"&avs="&pcv_AVS&_  
		"&checkcvv="&pcv_CheckCVV&_  
		"&cvv="& request.Form("CVV")&_	
		"&exp="&request.Form("expMonth")&request.Form("expYear")&_
		"&currency="&pcv_CurCode&_
		"&name="&pcBillingFirstName& " "&pcBillingLastName&_
		"&address1="&pcBillingAddress&_
		"&city="&pcBillingCity&_
		"&state="&pcBillingState&_
		"&zip="&pcBillingPostalCode&_
		"&country="&pcBillingCountryCode&_
		"&phone="&pcBillingPhone&_
		"&email="&pcCustomerEmail&_
		"&shipto_name="&pcShippingFirstName&" "&pcShippingLastName&_
		"&shipto_address1="&pcShippingAddress&_
		"&shipto_city="&pcShippingCity&_
		"&shipto_state="&pcShippingState&_
		"&shipto_zip"&pcShippingPostalCode&_
		"&shipto_country="&pcShippingCountryCode'&_
		
		' Create the WinHTTPRequest Object
		Set objHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objHTTP.Open "POST", strPostUrl & strRequest, false
		objHTTP.Send()    ' Send the HTTP request.
			
			
		If objHTTP.Status = 200 Then  ' HTTP_STATUS_OK=200 
		
		Dim objDictResponse, intDelimiterPos, ResponseArray
		Dim strResponse, strNameValuePair, strName, strValue
		strResponse = replace(objHTTP.ResponseText,chr(10)," ")
		strResponse = rtrim(strResponse)
			
		' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
		ResponseArray = Split(strResponse, " ")
		Set objDictResponse = server.createobject("Scripting.Dictionary")
		For each ResponseItem in ResponseArray
			NameValue = Split(ResponseItem, "=")
			objDictResponse.Add NameValue(0), NameValue(1)
		Next
			
		' Parse the response into local vars
		strAuthorizationNumber = objDictResponse.Item("authcode")
		strstatus        = objDictResponse.Item("status")
		strerror          = objDictResponse.Item("error")
		stroffenders       = objDictResponse.Item("offenders")
		strTransactionID       = objDictResponse.Item("transid")
		strAVSResponseCode     = objDictResponse.Item("avs")
						
		If lcase(strstatus) = "approved" Then
			session("GWAuthCode")=strAuthorizationNumber
			session("GWTransId")=strTransactionID
			session("GWTransType")=pcv_TransType
			Response.redirect "gwReturn.asp?s=true&gw=TCLink"
		Else
	
            call closeDb()
            Session("message") = strstatus&"&nbsp;in&nbsp;"&lcase(stroffenders)&" --Error Type:&nbsp;"&lcase(strerror)
            Session("backbuttonURL") = tempURL & "?psslurl=gwtclink.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1"

		End If  
			
	Else   
		Response.write "Connection Failed..."
		Response.Write "<BR> Https Response <BR>" 
		Response.Write "Status     = " & objHttp.status       & "<BR>"
		Response.Write "StatusText = " & objHttp.statusText   & "<BR>"
		Response.Write "Header     = " & objHttp.getAllResponseHeaders & _
				"<BR>"
		Response.Write "RespText   = " & objHttp.responseText & "<BR>"
		
	End If   
	Set objHttp   = Nothing

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
                    
					<% 'If x_CVV="1" Then %>
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155"></div>
                </div>
					<% 'End If %>

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
