<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwConcord.asp"

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
query="SELECT StoreID, StoreKey, testmode, Curcode, CVV, MethodName FROM concord Where idConcord=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_StoreID=rs("StoreID")
pcv_StoreKey=rs("StoreKey")
'decrypt
pcv_StoreKey=enDeCrypt(pcv_StoreKey, scCrypPass)
pcv_CVV=rs("CVV")
pcv_Curcode=rs("Curcode")
pcv_TestMode=rs("testmode")
pcv_MethodName=rs("MethodName")

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Quote = String(1,34) 
	If pcv_TestMode =1 Then
		strDLLUrl   = "https://stg.dw.us.fdcnet.biz/efsnet.dll"      
	Else ' Live Mode
		strDLLUrl = "https://prod.dw.us.fdcnet.biz/efsnet.dll"
	End If   

	'--------------------------------------------------------------------------
	' For simplicity in this sample, we'll only populate variables for a 
	' SystemCheck.  For other methods, required fields such as 
	' "strTransactionAmount" would need to be populated before calling 
	' ProcessCGIRequest. Notice how easy it is to use any EFSnet method 
	' by simply changing the method name passed to ProcessCGIRequest.
	'--------------------------------------------------------------------------
	' Populate Transaction Request Variables
	If pcv_TestMode =1 Then
		strApplicationID = "ProductCart Test"
	else
		strApplicationID = "ProductCart"
	end if
	
	Select Case pcv_MethodName
		' Build a properly formatted EFSnet Cgi Request Message 
		Case "SysCheck"  ProcessCGIRequest("SystemCheck")  
		Case "Authorize" ProcessCGIRequest("CreditCardAuthorize")
		Case "Settle"    ProcessCGIRequest("CreditCardSettle")  
		Case "Charge"    ProcessCGIRequest("CreditCardCharge") 'Auth/Settle
		Case "Refund"    ProcessCGIRequest("CreditCardRefund")  
	End Select

	'--------------------------------------------------------------------------
	' HELPER FUNCTIONS
	'--------------------------------------------------------------------------
	
	'--------------------------------------------------------------------------
	Public Sub ProcessCGIRequest(strMethod)
		 Dim objHTTP, strRequest 
	
		 strRequest = "Method="       & strMethod										& _
			"&StoreID="               & pcv_StoreID              						& _
			"&StoreKey="              & pcv_StoreKey              						& _   
			"&ApplicationID="         & strApplicationID       	  						& _   
			"&AccountNumber="         & request.Form("CardNumber")						& _   
			"&ExpirationMonth="       & request.Form("expMonth")  						& _
			"&ExpirationYear="        & request.Form("expYear")							& _
			"&CardVerificationValue=" & strCardVerificationValue						& _
			"&Track1="                & strTrack1                						& _
			"&Track2="                & strTrack2                						& _
			"&TerminalID="            & strTerminalID            						& _
			"&CashierNumber="         & strCashierNumber         						& _ 
			"&ReferenceNumber="       & session("GWOrderId")	 						& _  
			"&TransactionAmount="     & pcBillingTotal				    				& _
			"&SalesTaxAmount="        & strSalesTaxAmount       						& _
			"&Currency="              & pcv_Curcode               						& _
			"&BillingName="           & pcBillingFirstName& " "&pcBillingLastName		& _
			"&BillingAddress="        & pcBillingAddress        		 				& _
			"&BillingCity="           & pcBillingCity           						& _
			"&BillingState="          & pcBillingStateCode          		 			& _
			"&BillingPostalCode="     & pcBillingPostalCode     				 		& _
			"&BillingCountry="        & pcBillingCountryCode        		 			& _
			"&BillingPhone="          & pcBillingPhone         		 					& _
			"&BillingEmail="          & pcCustomerEmail          		 				& _
			"&ShippingName="          & pcShippingFirstName&" "&pcShippingLastName		& _
			"&ShippingAddress="       & pcShippingAddress      							& _
			"&ShippingCity="          & pcShippingCity         							& _
			"&ShippingState="         & pcShippingStateCode          					& _
			"&ShippingPostalCode"     & pcShippingPostalCode			 				& _
			"&ShippingCountry="       & pcShippingCountryCode       					& _
			"&ShippingPhone="         & pcShippingPhone         						& _
			"&ShippingEmail="         & strShippingEmail         						& _
			"&ClientIPAddress="       & pcCustIpAddress

		 If VIEW_CGI_REQUEST Then
				Response.Write strRequest & "<BR>"   ' Debug only
		 End If

		 ' Create the WinHTTPRequest Object
		 Set objHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		 objHTTP.Open "POST", strDLLUrl, false
		 objHTTP.Send(strRequest)    ' Send the HTTP request.

		 If objHTTP.Status = 200 Then  ' HTTP_STATUS_OK=200 
			
		 Dim objDictResponse, intDelimiterPos, ResponseArray
		 Dim strResponse, strNameValuePair, strName, strValue
		 strResponse = objHTTP.ResponseText

		If VIEW_CGI_REQUEST Then 
			Response.Write strResponse & "<BR>"   ' Debug only
		End If 

		' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
		ResponseArray = Split(strResponse, "&") 
		Set objDictResponse = server.createobject("Scripting.Dictionary")
		For each ResponseItem in ResponseArray
			NameValue = Split(ResponseItem, "=")
			objDictResponse.Add NameValue(0), NameValue(1)
		Next
       
		' Parse the response into local vars
		strResponseCode        = objDictResponse.Item("ResponseCode")
		strResultCode          = objDictResponse.Item("ResultCode")
		strResultMessage       = objDictResponse.Item("ResultMessage")
		strTransactionID       = objDictResponse.Item("TransactionID")
		strAVSResponseCode     = objDictResponse.Item("AVSResponseCode")
		strCVVResponseCode     = objDictResponse.Item("CVVResponseCode")
		strApprovalNumber      = objDictResponse.Item("ApprovalNumber")
		strAuthorizationNumber = objDictResponse.Item("AuthorizationNumber")
		strTransactionDate     = objDictResponse.Item("TransactionDate")
		strTransactionTime     = objDictResponse.Item("TransactionTime")

		If strResponseCode = 0 Then 
			session("GWAuthCode")=strAuthorizationNumber
			session("GWTransId")=strTransactionID
			session("GWTransType")=pcv_MethodName
			call closeDb()
            Response.redirect "gwReturn.asp?s=true&gw=Concord"
		Else
            call closeDb()
            Session("message") = strResponseCode&"&nbsp;&nbsp;"&lcase(strResultMessage)
            Session("backbuttonURL") = tempURL & "?psslurl=gwconcord.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
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
End Sub
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
