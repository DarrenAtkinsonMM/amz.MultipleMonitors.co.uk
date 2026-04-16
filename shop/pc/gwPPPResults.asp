<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.CacheControl = "No-Store"
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
%>
<!--#include file="../includes/common.asp"-->
<%
Dim PPPURL
If scSSL="" OR scSSL="0" Then
	PPPURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	PPPURL=replace(PPPURL,"https:/","https://")
	PPPURL=replace(PPPURL,"http:/","http://")
Else
	PPPURL=replace((scSslURL&"/"&scPcFolder&"/pc/"),"//","/")
	PPPURL=replace(PPPURL,"https:/","https://")
	PPPURL=replace(PPPURL,"http:/","http://")
End If

'//Check for a response
ppp_avszip = request("AVSZIP")
ppp_ppref = request("PPREF")
ppp_transactiontime = request("TRANSTIME")
ppp_ziptoship = request("ZIPTOSHIP")
ppp_lastname = request("LASTNAME")
ppp_pnref = request("PNREF")
ppp_avsdata = request("AVSDATA")
ppp_type = request("TYPE")
ppp_citytoship = request("CITYTOSHIP")
ppp_payerid = request("PAYERID")
ppp_tender = request("TENDER")
ppp_pendingreason = request("PENDINGREASON")
ppp_token = request("TOKEN")
ppp_method = request("METHOD")
ppp_avsaddr = request("AVSADDR")
ppp_addresstoship = request("ADDRESSTOSHIP")
ppp_securetoken = request("SECURETOKEN")
ppp_securetokenid = request("SECURETOKENID")
ppp_responsemessage = request("RESPMSG")
ppp_firstname = request("FIRSTNAME")
ppp_correlationid = request("CORRELATIONID")
ppp_countrytoship = request("COUNTRYTOSHIP")
ppp_statetoship = request("STATETOSHIP")
ppp_result = request("RESULT")
ppp_result = cstr(ppp_result)
if len(ppp_result)=0 then
	ppp_result="NONE"
end if
ppp_cancelflag = request("cancel_ec_trans")
ppp_prefpsmsg = request("PREFPSMSG")
ppp_hostcode = request("HOSTCODE")
ppp_invoice = request("INVOICE")
pOrderID=cLng(session("GWOrderId"))-cLng(scPre)
if clng(pOrderID)<0 then
	pOrderID = ppp_invoice
	session("GWOrderId") = Clng(pOrderID)+cLng(scPre)
end if
ppp_postfpsmsg = request("POSTFPSMSG") 'Review
ppp_acct = request("ACCT") '7930
ppp_proccvv2 = request("PROCCVV2") 'M
ppp_cvv2match = request("CVV2MATCH") 'Y
ppp_email = request("EMAIL") 
ppp_phone = request("PHONE") '1231231231
ppp_amt = request("AMT") '70.04
ppp_zip = request("ZIP") '92506
ppp_authcode = request("AUTHCODE") '111111
ppp_expdate = request("EXPDATE") '1017
ppp_iavs = request("IAVS") 'N
ppp_tax = request("TAX") '0.00
ppp_cardtype = request("CARDTYPE") '0
ppp_procavs = request("PROCAVS") 'X
ppp_prefpsmsg = request("PREFPSMSG") 'Review%3A+More+than+one+rule+was+triggered+for+Review
ppp_invnum = request("INVNUM") '29

if ppp_securetokenid&""<>"" then 
	
	query = "SELECT pcPay_PayPal_Signature FROM orders WHERE idOrder="& pOrderID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	pcPay_PayPal_SecureTokenId = rs("pcPay_PayPal_Signature")
	
	set rs=nothing
	%>
    <html><body><p><center>Your payment is currently being processed.<br />It can take up to 2 minutes to complete.</center></p>
	<%
	
	if pcPay_PayPal_SecureTokenId <> ppp_securetokenid then
		ppp_message = "Invalid Secure Token!"
		%>
		<script type=text/javascript>window.parent.location.href='gwReturn.asp?s=true&gw=PayPalAdvanced&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PPPURL&"gwPPP.asp?Message="&ppp_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if
	
	if lcase(ppa_cancelflag) = "true" then
		'Customer cancelled payment."
		ppp_message = "Result: Customer Canceled Payment"
		
		
		err.number=0
		err.clear
		%>
		<script type=text/javascript>window.parent.location.href='<%=PPPURL&"gwPPP.asp?Message="&ppp_message%>';</script>
        <noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PPPURL&"gwPPP.asp?Message="&ppp_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if

	if ppp_result = "0" then
		
		'//Update the customer's shipping address for this order
		query = "Update orders set shippingAddress='"&ppp_addresstoship&"', shippingStateCode='"&ppp_statetoship&"', shippingCity='"&ppp_citytoship&"', shippingCountryCode='"&ppp_countrytoship&"', shippingZip='"&ppp_ziptoship&"', ShippingFullName='"&ppp_firstname &" "&ppp_lastname&"' WHERE idOrder="& pOrderID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		'//Auth-only orders we will save in a new table
		query = "INSERT INTO pcPay_PFL_Authorize (idOrder, orderDate, paySource, amount, paymentmethod, transtype, authcode, captured, gwCode) VALUES ("& pOrderID &", '"&date()&"', 'PayPalPro', "&ppp_amt&", '"&ppp_method&"', '"&ppp_type&"', '"&ppp_pnref&"', 0, 53);"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		set rs=nothing
		
		'//gwReturn.asp
		session("GWAuthCode")=ppp_pnref
		session("GWTransId")=ppp_ppref
		session("GWTransType")=ppp_type
		session("AVSCode")=ppp_avsdata
        session("CVV2Code")=ppp_cvv2match
		
		fraudmode = ""
		if cstr(ppp_result) = "126" then
			fraudmode = "review"
			Session("FraudCode") = fraudmode
		end if
		
		session("GWTransType")=ppp_type

		%>
		<script type=text/javascript>window.parent.location.href='gwReturn.asp?s=true&gw=PayPalPro&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwPPP.asp?Message="&ppp_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	else
		if lcase(ppp_cancelflag) = "true" then
			'Customer cancelled payment."
			err.number=0
			err.clear
			%>
			<script type=text/javascript>window.parent.location.href='<%=PPPRUL&"gwPPP.asp?Message="&ppp_message%>';</script>
      <noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PPPRUL&"gwPPP.asp?Message="&ppp_message%>">here </a>to continue.</noscript>
			<%
			response.End()
		else
			ppp_message = "The payment could not be completed for the following reasons<br><ul><li>" & ppp_responsemessage &"</li></ul>"
			session("ppp_message") = ppp_message
			
			RedirectURLA = PFLURL&"gwPPP.asp?Message=session"
			%>
				<script type=text/javascript>window.parent.location.href='<%=RedirectURLA%>';</script>
				<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=RedirectURLA%>">here </a>to continue.</noscript>
			<%
			response.End()
		end if
	end if
    
    
else

	if ppp_prefpsmsg&""<>"" Then
		ppp_responsemessage = ppp_prefpsmsg
	end if
	ppp_message = "The payment could not be completed for the following reasons<br><ul><li>" & ppp_responsemessage &"</li></ul>"
	session("ppp_message") = ppp_message
	RedirectURLA = PPPURL&"gwPPP.asp?Message=session"
	
	err.number=0
	err.clear
	%>
	<script type=text/javascript>window.parent.location.href='<%=RedirectURLA%>';</script>
	<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=RedirectURLA%>">here </a>to continue.</noscript>
	<%
	response.End()
end if 
%>
</body></html>
