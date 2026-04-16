<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwtclinkCheck.asp"

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
query="SELECT TCLinkid, TCLinkPassword, TCTestmode, TCCurcode, TranType, TCLinkCheckPending FROM tclink Where idTCLink=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
TCLinkid=rs("TCLinkid")
TCLinkPassword=rs("TCLinkPassword")
'decrypt
TCLinkPassword=enDeCrypt(TCLinkPassword, scCrypPass)
TCLinkCheckPending=rs("TCLinkCheckPending")
TCCurcode=rs("TCCurcode")
TCTestmode=rs("TCTestmode")
action="sale"
DLLUrl="https://vault.trustcommerce.com/trans"
If TCTestmode =1 Then
		demo="y"
Else ' Live Mode
			demo=""   
End If   

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim objHTTP, strRequest 
	
	strRequest = "?custid="&TCLinkid&_
	"&password="&TCLinkPassword&_
	"&action="&action&_
	"&media="&"ach"&_
	"&demo="&demo&_
	"&amount="&pcBillingTotal*100&_   
	"&routing="&request.Form("bank_aba_code")&_
	"&account="&request.Form("bank_acct_num")&_
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
	"&shipto_country="&pcShippingCountryCode&_
	"&currency="&TCCurcode&_
	"&ClientIPAddress="&pcCustIpAddress
			
	' Create the WinHTTPRequest Object
	Set objHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
	objHTTP.Open "POST", DLLUrl & strRequest, false
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
			strstatus        = objDictResponse.Item("status")
			strerror          = objDictResponse.Item("error")
			stroffenders       = objDictResponse.Item("offenders")
			strTransactionID       = objDictResponse.Item("transid")
			strAVSResponseCode     = objDictResponse.Item("avs")
				
			If strstatus = "accepted" Then 
				'tordnum=(int(strTransactionID)-scpre)
				'session("AuthorizationNumber")=strAuthorizationNumber
				session("GWTransId")=strTransactionID
				session("TranType")=action
				session("GWTransType")=TCLinkCheckPending
				session("GWAuthCode")=""
				
				Response.redirect "gwReturn.asp?s=true&gw=TCLinkCheck"
			Else

                call closeDb()
                Session("message") = strstatus&"&nbsp;in&nbsp;"&lcase(stroffenders)&" &rsaquo;&nbsp;&lsaquo;Error Type:&nbsp;"&lcase(strerror)
                Session("backbuttonURL") = tempURL & "?psslurl=gwtclinkCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
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
                <img src="<%=pcf_getImagePath("images","sampleck.gif")%>" width="390" height="230"> 
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Name on the Account:</div>
                <div class="pcFormField">
                    <input name="bank_acct_name" type="text" size="35" maxlength="50">
                </div> 
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Bank Routing Number:</div>
                <div class="pcFormField">
                    <input name="bank_aba_code" type="text" size="35">
                </div> 
            </div>
 
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Bank Account Number:</div>
                <div class="pcFormField">
                    <input name="bank_acct_num" type="text" size="35" autocomplete="off">
                </div> 
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Check Number:</div>
                <div class="pcFormField">
                    <input name="check_num" type="text" size="35">
                </div> 
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Bank Account Type:</div>
                <div class="pcFormField">
                    <select name="bank_acct_type">
								<option value="CHECKING">Checking Account</option>
								<option value="SAVINGS">Savings Account</option>
							</select>
                </div> 
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Bank Name:</div>
                <div class="pcFormField">
                    <input name="bank_name" type="text" size="35" maxlength="20">
                </div> 
            </div>       
  
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
