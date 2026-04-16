<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.CacheControl = "No-Store"
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
%>
<!--#include file="../includes/common.asp"-->
<%
Dim PPAURL
If scSSL="" OR scSSL="0" Then
	PPAURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	PPAURL=replace(PPAURL,"https:/","https://")
	PPAURL=replace(PPAURL,"http:/","http://")
Else
	PPAURL=replace((scSslURL&"/"&scPcFolder&"/pc/"),"//","/")
	PPAURL=replace(PPAURL,"https:/","https://")
	PPAURL=replace(PPAURL,"http:/","http://")
End If

'//Check for a response
ppa_avszip = request("AVSZIP")
ppa_ppref = request("PPREF")
ppa_transactiontime = request("TRANSTIME")
ppa_ziptoship = request("ZIPTOSHIP")
ppa_lastname = request("LASTNAME")
ppa_pnref = request("PNREF")
ppa_avsdata = request("AVSDATA")
ppa_type = request("TYPE")
ppa_citytoship = request("CITYTOSHIP")
ppa_payerid = request("PAYERID")
ppa_tender = request("TENDER")
ppa_pendingreason = request("PENDINGREASON")
ppa_token = request("TOKEN")
ppa_method = request("METHOD")
ppa_avsaddr = request("AVSADDR")
ppa_addresstoship = replace(request("ADDRESSTOSHIP"),"null","")
ppa_securetoken = request("SECURETOKEN")
ppa_securetokenid = request("SECURETOKENID")
ppa_responsemessage = request("RESPMSG")
ppa_firstname = request("FIRSTNAME")
ppa_correlationid = request("CORRELATIONID")
ppa_countrytoship = request("COUNTRYTOSHIP")
ppa_statetoship = request("STATETOSHIP")

ppa_result = request("RESULT")
ppa_result = cstr(ppa_result)
if len(ppa_result)=0 then
	ppa_result="NONE"
end if

ppa_cancelflag = request("cancel_ec_trans")
ppa_prefpsmsg = request("PREFPSMSG")
ppa_hostcode = request("HOSTCODE")
ppa_invoice = request("INVOICE")
'read true order ID from session
ppa_trueOrderID=cLng(session("GWOrderId"))-cLng(scPre)
if clng(ppa_trueOrderID)<0 then
	'load it from 
	ppa_trueOrderID = ppa_invoice
	session("GWOrderId") = Clng(ppa_trueOrderID)+cLng(scPre)
end if
ppa_postfpsmsg = request("POSTFPSMSG") 'Review
ppa_acct = request("ACCT") '7930
ppa_proccvv2 = request("PROCCVV2") 'M
ppa_cvv2match = request("CVV2MATCH") 'Y
ppa_email = request("EMAIL") 
ppa_phone = request("PHONE") '1231231231
ppa_amt = request("AMT") '70.04
ppa_zip = request("ZIP") '92506
ppa_authcode = request("AUTHCODE") '111111
ppa_expdate = request("EXPDATE") '1017
ppa_iavs = request("IAVS") 'N
ppa_tax = request("TAX") '0.00
ppa_cardtype = request("CARDTYPE") '0
ppa_procavs = request("PROCAVS") 'X
ppa_prefpsmsg = request("PREFPSMSG") 'Review%3A+More+than+one+rule+was+triggered+for+Review
ppa_invnum = request("INVNUM") '29

if ppa_securetokenid&""<>"" then 

	query = "SELECT pcPay_PayPal_Signature FROM orders WHERE idOrder="& ppa_trueOrderID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	pcPay_PayPal_SecureTokenId = rs("pcPay_PayPal_Signature")
	
	set rs=nothing
	%>
    <html><body><p><center>Your payment is currently being processed.<br />It can take up to 2 minutes to complete.</center></p>
	<%
	if pcPay_PayPal_SecureTokenId <> ppa_securetokenid then
		ppa_message = "Invalid Secure Token!"
		%>
		<script type=text/javascript>window.parent.location.href='gwReturn.asp?s=true&gw=PayPalAdvanced&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PPAURL&"gwPPA.asp?Message="&ppa_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if
	
	if lcase(ppa_cancelflag) = "true" then
		'Customer cancelled payment."
		ppa_message = "Result: Customer Canceled Payment"
		err.number=0
		err.clear
		%>
		<script type=text/javascript>window.parent.location.href='<%=PPAURL&"gwPPA.asp?Message="&ppa_message%>';</script>
        <noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PPAURL&"gwPPA.asp?Message="&ppa_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if


	if ppa_result = "0" then
		
		'//Update the customer's shipping address for this order
		query = "Update orders set shippingAddress='"&ppa_addresstoship&"', shippingStateCode='"&ppa_statetoship&"', shippingCity='"&ppa_citytoship&"', shippingCountryCode='"&ppa_countrytoship&"', shippingZip='"&ppa_ziptoship&"', ShippingFullName='"&ppa_firstname &" "&ppa_lastname&"' WHERE idOrder="& ppa_trueOrderID
        set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		query = "INSERT INTO pcPay_PFL_Authorize (idOrder, orderDate,paySource, amount, paymentmethod, transtype, authcode,  captured, fraudcode, gwCode) VALUES ("& ppa_trueOrderID &", '"&date()&"', 'PayPalAdvanced', "&ppa_amt&", '"&ppa_method&"', '"&ppa_type&"', '"&ppa_pnref&"',0, '"&ppa_result&"', 80);"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		set rs=nothing
		
		'//gwReturn.asp
		session("GWAuthCode")=ppa_pnref
		session("GWTransId")=ppa_ppref
		session("GWTransType")=ppa_type
		session("AVSCode")=ppa_avszip 
        session("CVV2Code")=ppa_cvv2match
		
		fraudmode = ""
		if cstr(ppa_result) = "126" then
			fraudmode = "review"
			Session("FraudCode") = fraudmode
		end if
		
		session("GWTransType")=ppa_type

		%>
		<script type=text/javascript>window.parent.location.href='gwReturn.asp?s=true&gw=PayPalAdvanced&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PPAURL&"gwPPA.asp?Message="&ppa_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	else
		if lcase(ppa_cancelflag) = "true" then
			'Customer cancelled payment."
			ppa_message = "Result: Customer Canceled Payment"
			err.number=0
			err.clear
			%>
			<script type=text/javascript>window.parent.location.href='<%=PPAURL&"gwPPA.asp?Message="&ppa_message%>';</script>
        	<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PPAURL&"gwPPA.asp?Message="&ppa_message%>">here </a>to continue.</noscript>
			<%
			response.End()
		else
			if ppa_responsemessage&""="" and ppa_prefpsmsg&""<>"" Then
				ppa_responsemessage = ppa_prefpsmsg
			end if
			ppa_message = "The payment could not be completed for the following reasons<br><ul><li>" & ppa_responsemessage &"</li></ul>"
			session("ppa_message") = ppa_message
			RedirectURLA = PPAURL&"gwPPA.asp?Message=session"
			err.number=0
			err.clear
			%>
			<script type=text/javascript>window.parent.location.href='<%=RedirectURLA%>';</script>
      <noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=RedirectURLA%>">here </a>to continue.</noscript>
			<%
			response.End()
		end if
	end if
else
	if ppa_prefpsmsg&""<>"" Then
		ppa_responsemessage = ppa_prefpsmsg
	end if
	ppa_message = "The payment could not be completed for the following reasons<br><ul><li>" & ppa_responsemessage &"</li></ul>"
	session("ppa_message") = ppa_message
	RedirectURLA = PPAURL&"gwPPA.asp?Message=session"
	err.number=0
	err.clear
	%>
	<script type=text/javascript>window.parent.location.href='<%=RedirectURLA%>';</script>
	<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=RedirectURLA%>">here </a>to continue.</noscript>
	<%
	response.End()
end if %>
</body></html>
