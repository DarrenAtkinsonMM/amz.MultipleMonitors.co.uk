<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwEcho.asp"

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
query="SELECT transaction_type, merchant_echo_id, merchant_pin, merchant_email, cnp_security  FROM echo Where id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
transaction_type=rs("transaction_type")
merchant_echo_id=rs("merchant_echo_id")
merchant_pin=rs("merchant_pin")
merchant_email=rs("merchant_email")
pcv_CVV=rs("cnp_security")
if cnp_security="" then
	cnp_security=0
end if

set rs=nothing

if request("PaymentSubmitted")="Go" then
	intOrderId = session("GWOrderId")
	if intOrderId&""="" then
		'session has been lost or idle for too long - redirect customer to session failed message.
		response.redirect "msg.asp?message=38"
	end if

	'decrypt
	merchant_echo_id=enDeCrypt(merchant_echo_id, scCrypPass)
	'decrypt
	merchant_pin=enDeCrypt(merchant_pin, scCrypPass)

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	'Send the request to ECHO for processing.
	stext="transaction_type="&transaction_type
	stext=stext & "&order_type=S"
	stext=stext & "&counter=1&debug=F"
	stext=stext & "&merchant_echo_id="&merchant_echo_id
	stext=stext & "&merchant_pin="&merchant_pin
	stext=stext & "&isp_echo_id="
	stext=stext & "&isp_pin="
	stext=stext & "&billing_ip_address="&pcCustIpAddress
	stext=stext & "&merchant_email="&merchant_email
	stext=stext & "&cc_number="&request.Form("CardNumber")
	stext=stext & "&ccexp_month="&request.Form("expMonth")
	stext=stext & "&ccexp_year="&request.Form("expYear")
	if cnp_security=1 then
		stext=stext & "&cnp_security="&request.Form("CVV")
	end if
	stext=stext & "&grand_total="&money(pcBillingTotal)
	stext=stext & "&billing_prefix="
	stext=stext & "&billing_first_name="&pcBillingFirstName
	stext=stext & "&billing_last_name="&pcBillingLastName
	stext=stext & "&billing_address1="&pcBillingAddress
	stext=stext & "&billing_address2="&pcBillingAddress2
	stext=stext & "&billing_city="&pcBillingCity
	stext=stext & "&billing_state="&pcBillingState
	stext=stext & "&billing_zip="&pcBillingPostalCode
	stext=stext & "&billing_country="&pcBillingCountryCode
	stext=stext & "&billing_phone="&pcBillingPhone
	stext=stext & "&billing_fax="
	stext=stext & "&billing_email="&pcCustomerEmail
	stext=stext & "&merchant_trace_nbr="&session("GWOrderId")
	stext=stext & "&auth_code="
	stext=stext & "&order_number="
	stext=stext & "&original_amount="
	stext=stext & "&original_reference="
	stext=stext & "&original_trandate_mm="
	stext=stext & "&original_trandate_dd="
	stext=stext & "&original_trandate_yyyy="
	stext=stext & "&order_info="

	'Send the transaction info as part of the querystring
	err.clear
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)

	xml.open "POST", "https://wwws.echo-inc.com/scripts/INR200.EXE?"&stext, False
	xml.send
	strStatus = xml.Status
	result = xml.responseText
	Set xml = Nothing
	
	'Check the ErrorCode to make sure that the component was able to talk to the authorization network
	If (strStatus <> 200) Then
		Response.Write "An error occurred during processing. Please try again later."
	else
		'ECHO STATUS
		intStartChar = InStr(1,result, "<status>", 0)+8
		intEndChar = InstrRev(result,"</status>")
		ECHO_status=trim(Mid(result, intStartChar, (intEndChar-intStartChar)))
		if ucase(ECHO_status)="D" then
			intStartChar = InStr(1,result, "<decline_code>", 0)+ 14
			intEndChar = InstrRev(result,"</decline_code>")
			ECHO_decline_code=Mid(result, intStartChar, (intEndChar-intStartChar))
			%>
			<!--#include file="gwEchoCodes.asp"-->
			<%
		end if
		if ECHO_status="G" then
			intStartChar = InStr(1,result, "<auth_code>", 0) +11
			intEndChar = InstrRev(result,"</auth_code>")
			ECHO_auth_code=Mid(result, intStartChar, (intEndChar-intStartChar))
			if instr(result,"<echo_reference>") then
				intStartChar = InStr(1,result, "<echo_reference>", 0) +16
				intEndChar = InstrRev(result,"</echo_reference>")
				ECHO_echo_reference=Mid(result, intStartChar, (intEndChar-intStartChar))
			end if
			intStartChar = InStr(1,result, "<merchant_trace_nbr>", 0) +20
			intEndChar = InstrRev(result,"</merchant_trace_nbr>")
			ECHO_merchant_trace_nbr=Mid(result, intStartChar, (intEndChar-intStartChar))
			intStartChar = InStr(1,result, "<order_number>", 0) +14
			intEndChar = InstrRev(result,"</order_number>")
			ECHO_order_number=Mid(result, intStartChar, (intEndChar-intStartChar))
		end if
		if ECHO_status="T" then
			ECHO_message="Timeout waiting for host response. Please try again later or contact the merchant directly"
		end if
		'save and update order 
		If ECHO_status = "G" Then
        
			tordnum=(int(ECHO_merchant_trance_nbr)-scpre)
			session("GWTransType")=transaction_type
			call closeDb()
            Response.redirect "gwReturn.asp?s=true&gw=ECHO"
            
		elseif ECHO_status="D" then
			
            call closeDb()
            Session("message") = "Error code " & ECHO_decline_code&": "& declined_msg
            Session("backbuttonURL") = tempURL & "?psslurl=gwecho.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1"
            
		End If
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
