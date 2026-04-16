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

<!--#include file="header_wrapper.asp"-->
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
	
'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass

'//Set redirect page to the current file name
session("redirectPage")="gwPFL.asp"

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
	session("GWOrderId")=getUserInput(request("idOrder"),0)
end if

'//Retrieve customer data from the database using the current session id
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
	
objPayPalClass.pcs_SetAllVariables()
objPayPalClass.pcs_SetShipAddress((int(session("GWOrderId"))-scpre))

newSecureTokenID = objPayPalClass.genrandomvalue(36)

'//SAVE TOKEN TO ORDER
query = "UPDATE orders SET pcPay_PayPal_Signature = '"&newSecureTokenID&"' WHERE idOrder="& pcGatewayDataIdOrder
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

set rs=nothing

if pcShippingFullName&""="" then
	pcShippingFullName = pcBillingFirstName &" "&pcBillingLastName
end if
if pcShippingAddress&""="" then
	pcShippingAddress = pcBillingAddress
	pcShippingCity = pcBillingCity
	pcShippingStateCode = pcBillingStateCode
	pcShippingPostalCode = pcBillingPostalCode
	pcShippingCountryCode = pcBillingCountryCode
	pcShippingAddress2 = pcBillingAddress2
	pcShippingPhone = pcBillingPhone
end if

stext = "USER="&trim(pcPay_PayPal_Username)
stext = stext & "&VENDOR="&trim(pcPay_PayPal_Vendor)
stext = stext & "&PARTNER="&pcPay_PayPal_Partner
stext = stext & "&PWD="&pcPay_PayPal_Password
stext = stext & "&TRXTYPE="&PaymentAction
stext = stext & "&CREATESECURETOKEN=Y"
stext = stext & "&RETURNURL="&PFLURL&"gwPFLResults.asp"
stext = stext & "&CANCELURL="&PFLURL&"gwPFLResults.asp"
stext = stext & "&ERRORURL="&PFLURL&"gwPFLResults.asp"
stext = stext & "&URLMETHOD=POST"
stext = stext & "&SILENTPOST=False"
stext = stext & "&TEMPLATE=MINLAYOUT"
stext = stext & "&SECURETOKENID="&newSecureTokenID
stext = stext & "&INVNUM="&session("GWOrderId")
stext = stext & "&AMT="&pcBillingTotal
stext = stext & "&BILLTOFIRSTNAME="&pcBillingFirstName
stext = stext & "&BILLTOLASTNAME="&pcBillingLastName
stext = stext & "&BILLTOSTREET="&pcBillingAddress
stext = stext & "&BILLTOSTREET2="&pcBillingAddress2
stext = stext & "&BILLTOCITY="&pcBillingCity
stext = stext & "&BILLTOSTATE="&pcBillingStateCode
stext = stext & "&BILLTOZIP="&pcBillingPostalCode
stext = stext & "&BILLTOPHONENUM="&pcBillingPhone
stext = stext & "&EMAIL="&pcCustomerEmail
stext = stext & "&DISABLERECEIPT=TRUE"
stext = stext & "&ADDROVERRIDE=1"
stext = stext & "&SHIPTONAME="&pcShippingFullName
stext = stext & "&SHIPTOSTREET="&pcShippingAddress
stext = stext & "&SHIPTOCITY="&pcShippingCity
stext = stext & "&SHIPTOSTATE="&pcShippingStateCode
stext = stext & "&SHIPTOZIP="&pcShippingPostalCode
stext = stext & "&SHIPTOCOUNTRYCODE="&pcShippingCountryCode
stext = stext & "&SHIPTOSTREET2="&pcShippingAddress2
stext = stext & "&SHIPTOPHONENUM="&pcShippingPhone
stext = stext & "&BUTTONSOURCE=ProductCart_Cart_PPA"
%>
<!--#include file="pcPay_PPA_Itemize.asp"-->
<%	
'// PayPal requires two decimal places with a "." decimal separator.
pcv_strFinalTotal= pcf_CurrencyField(money(pcv_strFinalTotal))
pcv_strFinalShipCharge= pcf_CurrencyField(money(pcv_strFinalShipCharge))
pcv_strFinalServiceCharge= pcf_CurrencyField(money(pcv_strFinalServiceCharge))
pcv_strFinalTax= pcf_CurrencyField(money(pcv_strFinalTax))
ItemTotal= pcf_CurrencyField(money(ItemTotal))
stext = stext & iString
stext = stext & "&ITEMAMT="&ItemTotal
stext = stext & "&FREIGHTAMT="&pcv_strFinalShipCharge
stext = stext & "&HANDLINGAMT="&pcv_strFinalServiceCharge
stext = stext & "&TAXAMT="&pcv_strFinalTax

'Send the transaction info as part of the querystring
set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
'SB S
xml.open "POST", objPayPalClass.GetPayPalURL(pcPay_PayPal_Method), false
'SB E

xml.Send stext
strStatus = xml.Status

'store the response
strRetVal = xml.responseText
Set xml = Nothing

split_resultXML = split(strRetVal,"&")
j=0
for each item in split_resultXML
  split_param = split(split_resultXML(j),"=")
  formname = split_param(0)
  formvalue = split_param(1)
  if ucase(formname)  = "RESULT" then resultcode_pymt = formvalue
  if ucase(formname)  = "RESPMSG" then resultvalue_pymt = formvalue
  if ucase(formname)  = "SECURETOKEN" or ucase(formname) = "SECURETOKENID" then
	 if trim(iframeString) = "" then
		iframeString = formname &"="& formvalue
	 else
		iframeString = iframeString & "&"& formname &"="& formvalue
	 end if
  end if
  j = j + 1	  
next
dim ppmode
ppmode = ""
if pcPay_PayPal_Method = "sandbox" then ppmode = "MODE=TEST&"

%>
<div id="pcMain">
	<div class="pcMainContent">
		<% if request("Message")&""<>"" then
			myMsg = getUserInput(request("Message"),0)
			
			if lcase(myMsg)="session" then
				myMsg = session("pfl_message")
			end if %>
      <div class="pcErrorMessage"><%=myMsg%></div>
		<% end if %>
		<div class="pcPayPalFrameContainer">
      <iframe class="pcPayPalFrame" id="PFLFrame" name="PFLFrame" src="https://payflowlink.paypal.com/?<%=ppmode%><%=iframeString%>" width="490" height="565" scrolling="no"></iframe>
		</div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
