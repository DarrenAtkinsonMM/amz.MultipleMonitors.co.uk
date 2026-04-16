<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="protx_functions.asp"-->

<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwProtx.asp"

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
'query="SELECT Protxid,ProtxPassword,ProtxTestmode,ProtxCurcode,TxType,avs FROM protx Where idProtx=1;"
'set rs=server.CreateObject("ADODB.RecordSet")
'set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err=DA"&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_VendorName="multiplemonitor"
pcv_rVendorName=pcv_VendorName
pcv_password="Ý©pxP‡Pqõ±”"
	'decrypt
	pcv_password=enDeCrypt(pcv_password, scCrypPass)
ProtxTestmode=0
pcv_CurCode="GBP"
pcv_TxType="DEFERRED"
pcv_AVVS = 1

If ProtxTestmode=1 Then
	'Simulator Account
	vspsite="https://test.sagepay.com/Simulator/VSPFormGateway.asp"
Else ' Live Mode
	If ProtxTestmode=2 then
		'vspsite="https://test.sagepay.com/gateway/service/vspform-register.vsp"
	    vspsite="https://sandbox.opayo.eu.elavon.com/gateway/service/vspform-register.vsp"
	Else
		'vspsite="https://live.sagepay.com/gateway/service/vspform-register.vsp"
	  	vspsite="https://live.opayo.eu.elavon.com/gateway/service/vspform-register.vsp"
	End if
End If	

If request.QueryString("crypt")<>"" then
	pcv_ReceiptCrypt=request.QueryString("crypt")
	
	' ** Decrypt the plaintext string for inclusion in the hidden field **
	'pcv_ReceiptString =SimpleXor(base64decode(pcv_ReceiptCrypt),pcv_password)
	
	' ** Decrypt the plaintext string for inclusion in the hidden field, using AES method **
	pcv_ReceiptString = DecodeAndDecrypt(pcv_ReceiptCrypt, pcv_password)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err=DA2"&pcStrCustRefID
	end if

	' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
	pcv_ResponseArray = Split(pcv_ReceiptString, "&") 
	Set objDictResponse = server.createobject("Scripting.Dictionary")
	For each ResponseItem in pcv_ResponseArray
		pcv_NameValue = Split(ResponseItem, "=")
		objDictResponse.Add pcv_NameValue(0), pcv_NameValue(1)
	Next
       
	' Parse the response into local vars
	pcv_StrStatus = objDictResponse.Item("Status")
	pcv_StrVendorTxCode  = objDictResponse.Item("VendorTxCode")
	If ProtxTestmode =1 Then
		pcv_StrVendorTxCode=replace(pcv_StrVendorTxCode,""&pcv_rVendorName&"","")
	End if
	pcv_StrTxAuthNo = objDictResponse.Item("TxAuthNo")
	pcv_StrAVSCV2 = objDictResponse.Item("AVSCV2")
	pcv_StrVPSTxID = objDictResponse.Item("VPSTxID")
       
	If pcv_StrStatus="OK" then
		session("GWAuthCode")=pcv_StrTxAuthNo
		session("GWTransId")=pcv_StrVPSTxID
		session("GWTransType")=pcv_TxType
		if session("GWOrderId")="" then
			session("GWOrderId")=session("ProtxOrdno")
		end if
		session("GWSessionID")=Session.SessionID 
		Response.redirect "gwReturn.asp?s=true&gw=SagePay"
	Else
		if pcv_strStatus = "ABORT" then 
			   StatusOutput = "You elected to cancel your online payment<BR>Any credit/debit card details you entered have not been sent to the bank. You will not be charged for this transaction. Press the BACK button to try again."
		elseif pcv_strStatus = "NOTAUTHED" then 
			   StatusOutput = "The VSP was unable to authorise your payment<BR>The acquiring bank would not authorise your selected method of payment. You will not be charged for this transaction. Press the BACK button to try again."
		else
			   StatusOutput = "An error has occurred at SagePay<br>Because an error occurred in the payment process, you will not be charged for this transaction, even if an authorisation was given by the bank. Press the BACK button to try again."
		end if

        call closeDb()
        Session("message") = StatusOutput
        Session("backbuttonURL") = tempURL & "?psslurl="&session("redirectPage")&"&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1&message="&StatusOutput

	end if
end if

pcCustomerName = pcBillingFirstName&" "&pcBillingLastName
                    
if scSSL="1" then
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwprotx.asp"),"//","/")
else
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwprotx.asp"),"//","/")
end if
tempURL=replace(tempURL,"http:/","http://")
tempURL=replace(tempURL,"https:/","https://")

ThisVendorTxCode = session("GWOrderId") & "_" & Hour(Now) & Minute(Now) & Second(Now)
If ProtxTestmode =1 Then
	ThisVendorTxCode = ThisVendorTxCode & timer() & rnd()
End if

pcvPostString = "VendorTxCode=" & ThisVendorTxCode & "&"
pcvPostString = pcvPostString + "Amount=" & money(pcBillingTotal) & "&"
pcvPostString = pcvPostString + "Currency=" & pcv_CurCode & "&"
DescString = "" 
for f = 1 to ppcCartIndex
 DescString = DescString  & replace(pcCartArray(f,1),"'","") & " || " 
next
if len(DescString)>0 then
	DescString = left(left(DescString,(len(DescString) - 3)), 100)
else
	DescString=scCompanyName& " Order"
end if
pcvPostString = pcvPostString + "Description="& DescString &"&"
pcvPostString = pcvPostString + "SuccessURL=" & tempURL & "&"
pcvPostString = pcvPostString + "FailureURL=" & tempURL & "&"
pcvPostString = pcvPostString + "CustomerName=" & left(pcCustomerName,100) & "&"
if pcCustomerEmail<>"" then
	pcvPostString = pcvPostString + "CustomerEMail=" & pcCustomerEmail & "&"
end if
pcvPostString = pcvPostString + "VendorEmail=" & scFrmEmail & "&"
pcvPostString = pcvPostString + "SendEmail=0" & "&"		' 1 : send to vendor and customer email		2 : send to vendor email only

pcvPostString = pcvPostString + "BillingFirstnames=" & left(pcBillingFirstName,20) & "&"
pcvPostString = pcvPostString + "BillingSurname=" & left(pcBillingLastName,20) & "&"
pcvPostString = pcvPostString + "BillingAddress1=" & left(pcBillingAddress,100) & "&"
pcvPostString = pcvPostString + "BillingAddress2=" & left(pcBillingAddress2,100) & "&"
pcvPostString = pcvPostString + "BillingCity=" & left(pcBillingCity,40) & "&"
pcvPostString = pcvPostString + "BillingPostCode=" & pcBillingPostalCode &"&"
pcvPostString = pcvPostString + "BillingCountry=" & pcBillingCountryCode & "&"
If Ucase(pcBillingCountryCode)="US" then
	pcvPostString = pcvPostString + "BillingState=" & pcBillingStateCode &"&"
End If
pcvPostString = pcvPostString + "BillingPhone=" & pcBillingPhone & "&"

pcvPostString = pcvPostString + "DeliveryFirstnames=" & left(pcShippingFirstName,20) & "&"
pcvPostString = pcvPostString + "DeliverySurname=" & left(pcShippingLastName,20) & "&"
pcvPostString = pcvPostString + "DeliveryAddress1=" & left(pcShippingAddress,100) & "&"
pcvPostString = pcvPostString + "DeliveryAddress2=" & left(pcShippingAddress2,100) & "&"
pcvPostString = pcvPostString + "DeliveryCity=" & left(pcShippingCity,40) & "&"
pcvPostString = pcvPostString + "DeliveryPostCode=" & pcShippingPostalCode & "&"
pcvPostString = pcvPostString + "DeliveryCountry=" & pcShippingCountryCode & "&"
If Ucase(pcBillingCountryCode)="US" then
	pcvPostString = pcvPostString + "DeliveryState=" & pcShippingStateCode &"&"
End If
pcvPostString = pcvPostString + "DeliveryPhone=" & pcShippingPhone & "&"

pcvPostString = pcvPostString + "ApplyAVSCV2=" & pcv_AVVS

' ** Encrypt the plaintext string for inclusion in the hidden field **
'pcv_Crypt = base64Encode(SimpleXor(pcvPostString,pcv_password))

' ** Encrypt the plaintext string for inclusion in the hidden field, using AES method **
pcv_Crypt = EncryptAndEncode(pcvPostString, pcv_password)

if err.number<>0 then
	call LogErrorToDatabase()
	'set rs=nothing
	'call closedb()
	'response.redirect "techErr.asp?err=DA3"&pcStrCustRefID
end if

session("redirectPage2")=vspsite

set rs=nothing
%>
	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Card Payment Details">Card Payment Details</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->
	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
<div id="pcMain">
	<div class="pcMainContent">
				<form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="VPSProtocol" value="3.00">
					<input type="hidden" name="TxType" value="<%=ucase(pcv_TxType)%>">
					<input type="hidden" name="Vendor" value="<%=pcv_VendorName%>">
					<input type="hidden" name="Crypt" value="<%=pcv_Crypt %>">

					<% If msg<>"" Then %>
              <div class="pcErrorMessage"><%=msg%></div>
          <% End If %>
                  
                  <% call pcs_showBillingAddress %>

          <div class="pcFormItem"> 
        <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
              <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
          </div>

          <div class="pcAttention"><br />
              NOTE: When you click on the 'Place Order' button, you will temporarily leave our Web site and will be taken to a secure payment page on the SagePay Web site. You will be redirected back to our store once the transaction has been processed. We have partnered with SagePay, a leader in secure Internet payment processing, to ensure that your transactions are processed securely and reliably.<br /><br />
          </div>

          <div class="pcFormButtons">
              <!--#include file="inc_gatewayButtons.asp"-->
          </div>
        </form>
    </div>
</div>
					</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->

<!--#include file="footer_wrapper.asp"-->
