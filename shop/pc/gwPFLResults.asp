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
<!--#include file="../includes/pcPayPalPFApiClass.asp"-->
<%
Dim PFLURL
If scSSL="" OR scSSL="0" Then
	PFLURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	PFLURL=replace(PFLURL,"https:/","https://")
	PFLURL=replace(PFLURL,"http:/","http://")
Else
	PFLURL=replace((scSslURL&"/"&scPcFolder&"/pc/"),"//","/")
	PFLURL=replace(PFLURL,"https:/","https://")
	PFLURL=replace(PFLURL,"http:/","http://")
End If

'//Check for a response
pfl_avszip = request("AVSZIP")
pfl_ppref = request("PPREF")
pfl_transactiontime = request("TRANSTIME")
pfl_ziptoship = request("ZIPTOSHIP")
pfl_lastname = request("LASTNAME")
pfl_pnref = request("PNREF")
pfl_avsdata = request("AVSDATA")
pfl_type = request("TYPE")
pfl_citytoship = request("CITYTOSHIP")
pfl_payerid = request("PAYERID")
pfl_tender = request("TENDER")
pfl_pendingreason = request("PENDINGREASON")
pfl_token = request("TOKEN")
pfl_method = request("METHOD")
pfl_avsaddr = request("AVSADDR")
pfl_addresstoship = request("ADDRESSTOSHIP")
pfl_securetoken = request("SECURETOKEN")
pfl_securetokenid = request("SECURETOKENID")
pfl_responsemessage = request("RESPMSG")
pfl_firstname = request("FIRSTNAME")
pfl_correlationid = request("CORRELATIONID")
pfl_countrytoship = request("COUNTRYTOSHIP")
pfl_statetoship = request("STATETOSHIP")

pfl_result = request("RESULT")
pfl_result = cstr(pfl_result)
if len(pfl_result)=0 then
	pfl_result="NONE"
end if
pfl_cancelflag = request("cancel_ec_trans")
pfl_prefpsmsg = request("PREFPSMSG")
pfl_hostcode = request("HOSTCODE")
pfl_invoice = request("INVOICE")
PFLpcOrderId=cLng(session("GWOrderId"))-cLng(scPre)
if clng(PFLpcOrderId)<0 then
	PFLpcOrderId = pfl_invoice
	session("GWOrderId") = Clng(PFLpcOrderId)+cLng(scPre)
end if
pfl_postfpsmsg = request("POSTFPSMSG") 'Review
pfl_acct = request("ACCT") '7930
pfl_proccvv2 = request("PROCCVV2") 'M
pfl_cvv2match = request("CVV2MATCH") 'Y
pfl_email = request("EMAIL") 
pfl_phone = request("PHONE") '1231231231
pfl_amt = request("AMT") '70.04
pfl_zip = request("ZIP") '92506
pfl_authcode = request("AUTHCODE") '111111
pfl_expdate = request("EXPDATE") '1017
pfl_iavs = request("IAVS") 'N
pfl_tax = request("TAX") '0.00
pfl_cardtype = request("CARDTYPE") '0
pfl_procavs = request("PROCAVS") 'X
pfl_prefpsmsg = request("PREFPSMSG") 'Review%3A+More+than+one+rule+was+triggered+for+Review
pfl_invnum = request("INVNUM") '29

if pfl_securetokenid&""<>"" then 
	
	query = "SELECT pcPay_PayPal_Signature FROM orders WHERE idOrder="& PFLpcOrderId
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	pcPay_PayPal_SecureTokenId = rs("pcPay_PayPal_Signature")
	
	set rs=nothing
	%>
    <html><body><p><center>Your payment is currently being processed.<br />It can take up to 2 minutes to complete.</center></p>
	<%
	
	if pcPay_PayPal_SecureTokenId <> pfl_securetokenid then
		pfl_message = "Invalid Secure Token!"
		%>
		<script type=text/javascript>window.parent.location.href='gwReturn.asp?s=true&gw=PFLink&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwPFL.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if
	
	if lcase(ppa_cancelflag) = "true" then
		'Customer cancelled payment."
		pfl_message = "Result: Customer Canceled Payment"
		
		
		err.number=0
		err.clear
		%>
		<script type=text/javascript>window.parent.location.href='<%=PFLURL&"gwPFL.asp?Message="&pfl_message%>';</script>
        <noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwPFL.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if

	if cstr(pfl_result) = "0" then
		
		'//Update the customer's shipping address for this order
		query = "Update orders set shippingAddress='"&pfl_addresstoship&"', shippingStateCode='"&pfl_statetoship&"', shippingCity='"&pfl_citytoship&"', shippingCountryCode='"&pfl_countrytoship&"', shippingZip='"&pfl_ziptoship&"', ShippingFullName='"&pfl_firstname &" "&pfl_lastname&"' WHERE idOrder="& PFLpcOrderId
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		'//All orders we will save in a new table
		query = "INSERT INTO pcPay_PFL_Authorize (idOrder, orderDate, paySource, amount, paymentmethod, transtype, authcode,  captured, fraudcode, gwCode) VALUES ("& PFLpcOrderId &", '"&date()&"', 'PFLink', "&pfl_amt&", '"&pfl_method&"', '"&pfl_type&"', '"&pfl_pnref&"',0, '"&ppa_result&"', 99);"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		set rs=nothing
		
		'//gwReturn.asp
		session("GWAuthCode")=pfl_pnref
		session("GWTransId")=pfl_ppref
		session("GWTransType")=pfl_type
		session("AVSCode")=pfl_avsdata
        session("CVV2Code")=pfl_cvv2match
		
		fraudmode = ""
		if cstr(pfl_result) = "126" then
			fraudmode = "review"
			Session("FraudCode") = fraudmode
		end if
		
		err.number=0
		err.clear
		
		session("GWTransType")=pfl_type

		%>
		<script type=text/javascript>window.parent.location.href='gwReturn.asp?s=true&gw=PFLink&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwPFL.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	else
		if lcase(pfl_cancelflag) = "true" then
			'Customer cancelled payment."
			err.number=0
			err.clear
			%>
			<script type=text/javascript>window.parent.location.href='<%=PFLURL&"gwPFL.asp?Message="&pfl_message%>';</script>
      <noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwPFL.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
			<%
			response.End()
		else
			pfl_message = "The payment could not be completed for the following reasons<br><ul><li>" & pfl_responsemessage &"</li></ul>"
			session("pfl_message") = pfl_message
			
			RedirectURLA = PFLURL&"gwPFL.asp?Message=session"
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
	'RESPMSG=Parameter+list+format+error%3A&RESULT=-6
	if pfl_prefpsmsg&""<>"" Then
		pfl_responsemessage = pfl_prefpsmsg
	end if
	pfl_message = "The payment could not be completed for the following reasons<br><ul><li>" & pfl_responsemessage &"</li></ul>"
	session("pfl_message") = pfl_message
	RedirectURLA = PPAURL&"gwPFL.asp?Message=session"
	
	err.number=0
	err.clear
	%>
	<script type=text/javascript>window.parent.location.href='<%=RedirectURLA%>';</script>
	<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=RedirectURLA%>">here </a>to continue.</noscript>
	<%
	response.End()
end if %>
</body></html>
