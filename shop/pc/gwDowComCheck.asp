<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->

<!--#include file="header_wrapper.asp"-->
<% 'Gateway specific files
		
		
		session("redirectPage")="gwDowComCheck.asp" 
		
		 Dim pcCustIpAddress
		pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
		
		dim tempURL
		If scSSL="" OR scSSL="0" Then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://") 
		Else
			tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://")
		End If
		
		' Get Order ID
		if session("GWOrderId")="" then
			session("GWOrderId")=request("idOrder")
		end if
		
		pcGatewayDataIdOrder=session("GWOrderID")
		%>
		<!--#include file="pcGateWayData.asp"-->
		<% session("idCustomer")=pcIdCustomer 

 'DELETE FOR HARD CODED VARS
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'//Functions parse response string
'======================================================================================
Function DecodeQueryValue(qValue)
		'Purpose: To URL decode a string
		'Pre: qValue is set to a url encoded value of a query string parameter. ex: "one+two"
		'Post: none
		'Returns: Returns the url decoded value of qValue. ex: "one two"
		Dim i
		Dim qChar
		dim newString
		if IsNull(qValue) = false then
			For i = 1 To Len(qValue)
			qChar = Mid(qValue, i, 1)
				If qChar = "%" Then
				on error resume next
				newString = newString & Chr("&H" & Mid(qValue, i + 1, 2))
				on error goto 0
				i = i + 2
				ElseIf qChar = "+" Then
				newString = newString & " "
				Else
				newString = newString & qChar
				End If
			Next
			
			DecodeQueryValue = Replace(newString, "&lt;", "<")
		else
			DecodeQueryValue = ""
		end if

End Function

Function GetQueryValue(queryString, paramName)
		'Purpose: To return the value of a parameter in an HTTP query string.
		'Pre: queryString is set to the full query string of url encoded name value pairs. ex:
		'"value1=one&value2=two&value3=3"
		' paramName is set to the name of one of the parameters in the queryString. ex: "value2"
		'Post: None
		'Returns: The function returns the query string value assigned to the paramName parameter. ex: "two"
		Dim pos1
		dim pos2
		Dim qString
		qString = "&" & queryString & "&"
		pos1 = InStr(1, qString, paramName & "=")
		If pos1 > 0 Then
		    pos1 = pos1 + Len(paramName) + 1
		    pos2 = InStr(pos1, qString, "&")
			If pos2 > 0 Then
			GetQueryValue = DecodeQueryValue(Mid(qString, pos1, pos2 - pos1))
			End If
		End If
End Function

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your 
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
		query= "SELECT pcPay_Dow_TransType,pcPay_Dow_MerchantID,pcPay_Dow_MerchantPassword,pcPay_Dow_CardTypes,pcPay_Dow_CVC,pcPay_Dow_TestMode FROM pcPay_DowCom where pcPay_Dow_ID=1"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		pcPay_Dow_TransType=rs("pcPay_Dow_TransType")
		pcPay_Dow_MerchantID=rs("pcPay_Dow_MerchantID")
		'decrypt
		pcPay_Dow_MerchantID=enDeCrypt(pcPay_Dow_MerchantID, scCrypPass)
		pcPay_Dow_MerchantPassword=rs("pcPay_Dow_MerchantPassword")
		'decrypt
		pcPay_Dow_MerchantPassword=enDeCrypt(pcPay_Dow_MerchantPassword, scCrypPass)
		pcPay_Dow_CardTypes=rs("pcPay_Dow_CardTypes")
		pcPay_Dow_CVC=rs("pcPay_Dow_CVC")
		pcPay_Dow_TestMode=rs("pcPay_Dow_TestMode")
		
		
		
		
		M="0"
		V="0"
		A="0"
		D="0"
		set rs=nothing
		
	if request("PaymentSubmitted")="Go" then 

			 Dim objXMLHTTP, xml
           
				 
				DataToSend = "type=" & Server.URLEncode("sale") &_
				 "&username=" & Server.URLEncode(pcPay_Dow_MerchantID) &_
			     "&password=" & Server.URLEncode(pcPay_Dow_MerchantPassword) &_
			     "&checkname=" & Server.URLEncode(request.form("x_bank_acct_name")) &_
			     "&checkaba=" & Server.URLEncode( request.form("x_bank_aba_code")) &_
				 "&checkaccount=" & Server.URLEncode(request.form("x_bank_acct_num")) &_
			     "&account_holder_type=" & Server.URLEncode(request.form("x_customer_organization_type")) &_
				 "&account_type=" & Server.URLEncode(request.form("x_bank_acct_type")) &_
			     "&payment=" & Server.URLEncode("check") &_
				 "&phone=" & Server.URLEncode(pcBillingPhone)& _
			     "&amount=" & Server.URLEncode(pcBillingTotal) &_
			     "&firstname=" & Server.URLEncode(pcBillingFirstName) &_
				 "&lastname=" & Server.URLEncode(pcBillingLastName) &_
			     "&address1=" & Server.URLEncode(pcBillingAddress) &_
				 "&city=" & Server.URLEncode(pcBillingCity) &_
				 "&state=" & Server.URLEncode(pcBillingState) &_ 
			     "&zip=" & Server.URLEncode(pcBillingPostalCode)& _				
				 "&country=" & Server.URLEncode(pcBillingCountryCode)& _ 
				 "&ipaddress=" & Server.URLEncode(pcCustIpAddress)& _
				 "&orderid=" & Server.URLEncode(session("GWOrderId"))& _ 
				 "&email=" & Server.URLEncode(pcCustomerEmail)
				 
				
             'Response.write DataToSend &"<BR><BR><BR>"
			'Send the transaction info as part of the querystring
			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			xml.open "POST", "https://secure.DowCommerce.net/api/transact.php", false
			xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
			xml.send(DataToSend)			
			if err.number<>0 then
				pcResultErrorMsg = err.description
			end if
			strStatus = xml.Status			
			
			if strStatus = 200 then 	
				'store the response
				strRetVal = xml.responseText	
				'Response.write strRetVal
				'Response.end 			
				
				if strRetVal <> "" Then		
				
					        pcResultResponseCode = GetQueryValue(strRetVal,"response")
							pcResultTransRefNumber = GetQueryValue(strRetVal,"transactionid") 
							pcResultErrorMsg = GetQueryValue(strRetVal,"responsetext")  
							pcResultAuthCode =  GetQueryValue(strRetVal,"authcode")
					
				Else
		        	'//ERROR
					pcResultErrorMsg = "Transaction error or declined.  Error Message: " & pcResultErrorMsg
                    call closeDb()
                    Session("message") = pcResultErrorMsg
                    Session("backbuttonURL") = tempURL & "?psslurl=gwDowComCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"
				 End If
				If pcResultResponseCode = "1"  then
				 	session("GWAuthCode")=pcResultAuthCode
					session("GWTransId")=pcResultTransRefNumber
					response.redirect "gwReturn.asp?s=true&gw=DowComCheck"
				Else				
					if pcResultErrorMsg="" then
					  pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"              	
					end if
                    call closeDb()
                    Session("message") = pcResultErrorMsg
                    Session("backbuttonURL") = tempURL & "?psslurl=gwDowComCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"	
				End if
			Else		  
				if pcResultErrorMsg="" then
					pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"					
				end if
                call closeDb()
                Session("message") = pcResultErrorMsg
                Session("backbuttonURL") = tempURL & "?psslurl=gwDowComCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                response.redirect "msgb.asp?back=1"			 
		   End if 
	End if 
			
				%>	
<div id="pcMain">
	<div class="pcMainContent">
					<form action="gwDowComCheck.asp" method="POST" name="form1" class="pcForms">
				

				<input type="hidden" name="PaymentSubmitted" value="Go">

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                    
                    <% call pcs_showBillingAddress %>

            <div class="pcFormItem"> 
                <img src="<%=pcf_getImagePath("images","sampleck.gif")%>" width="390" height="230">
            </div>      
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_12")%></div>
                <div class="pcFormField">
                    <input name="x_bank_acct_name" type="text" size="35" maxlength="50">
                </div> 
            </div> 

            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_13")%></div>
                <div class="pcFormField">
                    <input name="x_bank_aba_code" type="text" size="35">
                </div> 
            </div>  
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_14")%></div>
                <div class="pcFormField">
                    <input name="x_bank_acct_num" type="text" size="35">
                </div> 
            </div>

            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_15")%></div>
                <div class="pcFormField">
                    <select name="x_bank_acct_type">
								<option value="checking">Checking</option>
								<option value="savings">Savings</option>
							</select>
                </div> 
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_17")%></div>
                <div class="pcFormField">
                    <input type="radio" name="x_customer_organization_type" value="personal" class="clearBorder">Personal 
								<input type="radio" name="x_customer_organization_type" value="business" class="clearBorder">Business
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
