<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<script type="text/javascript">
function sameDel() {
	if (document.getElementById('mpDelChk').checked == true) {
		document.getElementById('mpDelFirstName').value = document.getElementById('mpBillFirstName').value;
		document.getElementById('mpDelLastName').value = document.getElementById('mpBillLastName').value;
		document.getElementById('mpDelTel').value = document.getElementById('mpBillTel').value;
		document.getElementById('mpDelAddress1').value = document.getElementById('mpBillAddress1').value;
		document.getElementById('mpDelAddress2').value = document.getElementById('mpBillAddress2').value;
		document.getElementById('mpDelCity').value = document.getElementById('mpBillCity').value;
		document.getElementById('mpDelPCode').value = document.getElementById('mpBillPCode').value;
		document.getElementById('mpDelCountry').value = document.getElementById('mpBillCountry').value;
	} else {
		document.getElementById('mpDelFirstName').value = '';
		document.getElementById('mpDelLastName').value = '';
		document.getElementById('mpDelTel').value = '';
		document.getElementById('mpDelAddress1').value = '';
		document.getElementById('mpDelAddress2').value = '';
		document.getElementById('mpDelCity').value = '';
		document.getElementById('mpDelPCode').value = '';
		document.getElementById('mpDelCountry').value = 'GB';
	}
}
</script>
<% 
'Redirect to secure page if not already secure
If Request.ServerVariables("HTTPS") = "off" Then
	Response.redirect("https://www.multiplemonitors.co.uk/shop/pc/manpay.asp")
end if

'//Set redirect page to the current file name
session("redirectPage")="manpay.asp?amount="&request.QueryString("amount")

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress

pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/manpay.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/manpay.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
		
'//Get Order ID
session("GWOrderId")=1
if session("GWOrderId")="" then
	session("GWOrderId")=1
end if

'//Retrieve customer data from the database using the current session id		
'pcGatewayDataIdOrder=session("GWOrderID")
pcGatewayDataIdOrder=1
%>
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=1

dim connTemp22, rs22
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT Protxid, ProtxTestmode, ProtxCurcode, CVV, avs, TxType, ProtxCardTypes, ProtxApply3DSecure FROM protx Where idProtx=1;"

set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_StrVendorName=rs("Protxid")
pcv_StrTestVendorName=pcv_StrVendorName

pcv_CVV=rs("CVV")
pcv_StrProtxTestmode=rs("ProtxTestmode")
pcv_ProtxCurcode=rs("ProtxCurcode")
avs=rs("avs")
pcv_StrTxType=rs("TxType")
pcv_StrProtxCardTypes=rs("ProtxCardTypes")
pcv_IntProtxApply3DSecure=rs("ProtxApply3DSecure")
if pcv_IntProtxApply3DSecure<>"0" AND pcv_IntProtxApply3DSecure<>"3" then
	pcv_IntProtxApply3DSecure="3"
end if

'MP Variables
mpAmount = request.querystring("amount")
mpTransCode = "mp" & CStr(Date) & CStr(Hour(Time)) & CStr(Minute(Time))
mpTransCode = replace(mpTransCode,"/","")

' ******************************************************************
' SagePay system to connect to
' ******************************************************************
pcv_StrProtocolVersion="3.00"

set rs22=nothing
call closedb()

pcv_StrProtxAmex=0
pcv_StrProtxMaestro=0
strFormCardTypes=""

cardTypeArray=split(pcv_StrProtxCardTypes,", ")

for i=lbound(cardTypeArray) to ubound(cardTypeArray)
	cardVar=cardTypeArray(i)
	select case ucase(cardVar)
		case "VISA"
			strFormCardTypes=strFormCardTypes&"<option value=""VISA"" selected>Visa</option>"
		case "MC"
			strFormCardTypes=strFormCardTypes&"<option value=""MC"" selected>MasterCard</option>"
		case "UKE" 
			strFormCardTypes=strFormCardTypes&"<option value=""UKE"">Visa Debit/Visa Electron</option>"
		case "AMEX"
			strFormCardTypes=strFormCardTypes&"<option value=""AMEX"">American Express</option>"
			pcv_StrProtxAmex=1
	end select
next

strFormCardTypes="<option value=""VISA"" selected>Visa</option><option value=""MC"">MasterCard</option><option value=""UKE"">Visa Debit/Visa Electron</option><option value=""AMEX"">American Express</option>"
pcv_StrProtxAmex=1

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	if pcv_StrProtxTestmode=1 then
		purchaseURL	= "https://test.sagepay.com/Simulator/VSPDirectGateway.asp"
	elseif pcv_StrProtxTestmode=2 then
		purchaseURL = "https://test.sagepay.com/gateway/service/vspdirect-register.vsp"
	elseif pcv_StrProtxTestmode=0 then
		purchaseURL = "https://live.sagepay.com/gateway/service/vspdirect-register.vsp"
	else
		'// problem
	end if
	
	'ThisVendorTxCode=session("GWOrderId")
	ThisVendorTxCode=mpTransCode
		
	Randomize
	
	If pcv_StrProtxTestmode =1 Then
		ThisVendorTxCode=pcv_StrTestVendorName & timer() & rnd()
	End if
	
	'// Validate all input
	pcv_CardNumber=request("CardNumber")
	pcv_CardNumber=replace(pcv_CardNumber,"-","")
	pcv_CardNumber=replace(pcv_CardNumber,".","")
	pcv_CardNumber=replace(pcv_CardNumber," ","")
	if NOT isNumeric(pcv_CardNumber) then
		'Card Number is not a valid format
	end if
		
	'set all the required outgoing properties
	postData = _
	"VPSProtocol=" & pcv_StrProtocolVersion & _
	"&TxType=" & pcv_StrTxType & _
	"&Vendor=" & pcv_StrVendorName

	postData = postData & "&VendorTxCode=" & ThisVendorTxCode
	postData = postData & "&Amount=" & replace(money(mpAmount),",","")
	postData = postData & "&Currency=" & pcv_ProtxCurcode
	postData = postData & "&Description=" & Server.URLEncode( session("GWOrderId") )
	postData = postData & "&CardHolder=" & request.form("mpBillFirstName") & " " & request.form("mpBillLastName")
	postData = postData & "&CardNumber=" & pcv_CardNumber
	if pcv_StrProtxAmex=1 then
		postData = postData & "&StartDate=" & request.form( "StartDate" )
	end if
	postData = postData & "&ExpiryDate=" & request.form( "expMonth" ) & request.form( "expYear" )
	if pcv_StrProtxMaestro=1 then
	 postData = postData & "&IssueNumber=" & request.form( "IssueNumber" )
	end if
	postData = postData & "&CardType=" & request.form( "ProtxCardTypes" )
	If pcv_CVV ="1" Then
		postData = postData & "&CV2=" & request.form( "CVV" )
	End If

	'//New for 2.23
	postData = postData & "&BillingSurname=" & left(request.form("mpBillLastName"),20)
	postData = postData & "&BillingFirstnames=" & left(request.form("mpBillFirstName"),20)
	postData = postData & "&BillingAddress1=" & left(request.form("mpBillAddress1"),100)
	postData = postData & "&BillingAddress2=" & left(request.form("mpBillAddress2"),100)
	postData = postData & "&BillingCity=" & left(request.form("mpBillCity"),40)
	postData = postData & "&BillingPostCode=" & left(request.form("mpBillPCode"),10)
	postData = postData & "&BillingCountry=" & left(request.form("mpBillCountry"),2)
	If ucase(pcBillingCountryCode) = "US" then
		postData = postData & "&BillingState=" & left(pcBillingStateCode,2)
	End If
	postData = postData & "&BillingPhone=" & left(request.form("mpBillTel"),20)
	if pcShippingFullName<>"" then
		pcShippingNameArry=split(pcShippingFullName, " ")
		if ubound(pcShippingNameArry)>0 then
			pcShippingFirstName=pcShippingNameArry(0)
			if ubound(pcShippingNameArry)>1 then
				 tmpShipFirstName = pcShippingFirstName&" "
				 pcShippingLastName = replace(pcShippingFullName,tmpShipFirstName,"")
			else
				pcShippingLastName=pcShippingNameArry(1)
			end if
		else
			pcShippingFirstName=pcShippingFullName
			pcShippingLastName=pcShippingFullName
		end if
	else
		pcShippingFirstName=pcBillingFirstName
		pcShippingLastName=pcBillingLastName
	end if
	
	if len(pcShippingLastName)> 0 then
		postData = postData & "&DeliverySurname=" & left(request.form("mpDelLastName"),20)
        else
		postData = postData & "&DeliverySurname=" & left(request.form("mpDelLastName"),20)
	end if
	
	if len(pcShippingFirstName)> 0 then
		postData = postData & "&DeliveryFirstnames=" & left(request.form("mpDelFirstName"),20)
        else
		postData = postData & "&DeliveryFirstnames=" & left(request.form("mpDelFirstName"),20)
	end if
	
	if len(pcShippingAddress)> 0 then
		postData = postData & "&DeliveryAddress1=" & left(request.form("mpDelAddress1"),100)
        else
		postData = postData & "&DeliveryAddress1=" & left(request.form("mpDelAddress1"),100)
	end if
	
	if len(pcShippingAddress2)> 0 then
		postData = postData & "&DeliveryAddress2=" & left(request.form("mpDelAddress2"),100)
        else
		postData = postData & "&DeliveryAddress2=" & left(request.form("mpDelAddress2"),100)
	end if
	
	if len(pcShippingCity)> 0 then
		postData = postData & "&DeliveryCity=" & left(request.form("mpDelCity"),40)
        else
		postData = postData & "&DeliveryCity=" & left(request.form("mpDelCity"),40)
	end if
	
	if len(pcShippingPostalCode)> 0 then
		postData = postData & "&DeliveryPostCode=" & left(request.form("mpDelPCode"),10)
        else
		postData = postData & "&DeliveryPostCode=" & left(request.form("mpDelPCode"),10)
	end if
	
	if len(pcShippingCountryCode)> 0 then
		postData = postData & "&DeliveryCountry=" & left(request.form("mpDelCountry"),2)
        else
		postData = postData & "&DeliveryCountry=" & left(request.form("mpDelCountry"),2)
	end if
	
	if len(pcShippingStateCode)> 0 then
		If ucase(pcBillingCountryCode) = "US" then
			postData = postData & "&DeliveryState=" & left(pcShippingStateCode,2)
        	else
			postData = postData & "&DeliveryState=" & left(pcBillingStateCode,2)
		End If	
	end if
	
	if len(pcShippingPhone)> 0 then
		postData = postData & "&DeliveryPhone=" & left(request.form("mpDelTel"),20)
        else
		postData = postData & "&DeliveryPhone=" & left(request.form("mpDelTel"),20)
	end if	
	
	postData = postData & "&CustomerEMail=" &request.form("mpBillEmail")
	
	If pcv_CVV ="1" Then
		ApplyAVSCV2=1
		postData = postData & "&ApplyAVSCV2=" & ApplyAVSCV2
	End If 
	postData = postData & "&ClientIPAddress=" & pcCustIpAddress
	
	'** Send the account type to be used for this transaction.  Web sites should us E for e-commerce **
	'** If you are developing back-office applications for Mail Order/Telephone order, use M **
	'** If your back office application is a subscription system with recurring transactions, use C **
	'** Your SagePay account MUST be set up for the account type you choose.  If in doubt, use E **
	postData = postData & "&AccountType=E"

	'// Use this variable to turn your BASKET feature ON/OFF - Default is "OFF"
	ThisShoppingBasket="OFF"

	if ThisShoppingBasket="ON" then
		'select all products from the ProductsOrdered table to insert them into the 2Checkout db.
		call opendb()
		query="SELECT products.idproduct, products.description, quantity, unitPrice FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& session("GWOrderId")
		set rsBasketObj=server.CreateObject("ADODB.Recordset")
		set rsBasketObj=connTemp22.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsBasketObj=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		IntProdCnt=0
		do until rsBasketObj.eof
			tempIntIdProduct=rsBasketObj("idproduct")
			tempStrDescription=rsBasketObj("description")
			tempIntQuantity=rsBasketObj("quantity")
			tempDblUnitPrice=rsBasketObj("unitPrice")
			IntProdCnt=IntProdCnt+1

			strPCBasket = strPCBasket & ":" & tempStrDescription & ":" & tempIntQuantity & ":" & tempDblUnitPrice & ":::" &tempDblUnitPrice

			rsBasketObj.moveNext
		loop 
		set rsBasketObj=nothing
		call closedb() 
		
		postData = postData & "&Basket="&IntProdCnt & strPCBasket
	end if
	
	'** Allow fine control over 3D-Secure checks and rules by changing this value. 0 is Default **
	if pcv_IntProtxApply3DSecure="0" then
		pcv_IntProtxApply3DSecure="0"
	end if
	'** It can be changed dynamically, per transaction, if you wish.  See the VSP Server Protocol document **
	postData = postData & "&Apply3DSecure="&pcv_IntProtxApply3DSecure&""
	'0 = If 3D-Secure checks are possible and rules allow, perform the checks and apply the authorisation rules (default).
	'1 = Force 3D-Secure checks for this transaction only (if your account is 3D-enabled) and apply rules for authorisation. 
	'2 = Do not perform 3D-Secure checks for this transaction only and always authorise. 
	'3 = Force 3D-Secure checks for this transaction (if your account is 3D-enabled) but ALWAYS obtain an auth code, irrespective of rule base.

	'send to SagePay
	set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
	httpRequest.option (9) = 2720
	
	' *** open connection to SagePay
	httpRequest.Open "POST", purchaseURL, False
	
	httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	httpRequest.send postData
	
	responseData = httpRequest.responseText
	
	'*** Following line shows the whole reply for debugging purposes ***
	'response.write "ERROR: "&Err.number & " - "&Err.description &"<br/>"
	'response.write "<b>Response was:</b> " & responseData & "<br/>"
	'response.end
	
	'** An non zero Err.number indicates an error of some kind **
	'** Check for the most common error... unable to reach the purchase URL **  
	strPageError="" 
	if err.number<>0 then
		if Err.number = -2147012889 then
			strPageError="Your server was unable to register this transaction with SagePay." &_
						"  Check that you do not have a firewall restricting the POST and " &_
						"that your server can correctly resolve the address " & strPurchaseURL
		else
			strPageError="An Error has occurred whilst trying to register this transaction.<BR>" &_
						"The Error Number is: " & Err.number & "<BR>" &_
						"The Description given is: " & Err.Description
		end If 
		if strPageError<>"" then
			'response.redirect "msgb.asp?did="&Err.number&" - "&Err.Description&"&message="&server.URLEncode("<b>Error</b>:"&strProtxErrorMsg&"<br><br><a href="""&tempURL&"?psslurl=gwProtxVSP.aspDAANDidCustomer="&session("idCustomer")&"DAANDidOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
			response.write(strPageError)
		end if
	end if
	' ******************************************************************
	' Determine next action
	'** No transport level errors, so the message got the SagePay **
	'** Analyse the response from VSP Direct to check that everything is okay **
	'** Registration results come back in the Status and StatusDetail fields **
	strStatus=findField("Status",responseData)
	strStatusDetail=findField("StatusDetail",responseData)

	if strStatus="3DAUTH" then
		'** This is a 3D-Secure transaction, so we need to redirect the customer to their bank **
		'** for authentication.  First get the pertinent information from the response **
		strMD=findField("MD",responseData)
		strACSURL=findField("ACSURL",responseData)
		strPAReq=findField("PAReq",responseData)
		strPageState="3DRedirect"
	else
		'** If this isn't 3D-Auth, then this is an authorisation result (either successful or otherwise) **
		'** Get the results form the POST if they are there **
		strVPSTxId=findField("VPSTxId",responseData)
		strSecurityKey=findField("SecurityKey",responseData)
		strTxAuthNo=findField("TxAuthNo",responseData)
		strAVSCV2=findField("AVSCV2",responseData)
		strAddressResult=findField("AddressResult",responseData)
		strPostCodeResult=findField("PostCodeResult",responseData)
		strCV2Result=findField("CV2Result",responseData)
		str3DSecureStatus=findField("3DSecureStatus",responseData)
		strCAVV=findField("CAVV",responseData)
	
	
		if strStatus="OK" then
			session("GWAuthCode")=strTxAuthNo
			session("GWTransId")=strVPSTxId
			session("GWTransType")=pcv_StrTxType
			Response.redirect "manpay-ok.asp?s=true&gw=SagePay"
		else
			if strStatus="AUTHENTICATED" then
				session("GWAuthCode")=strTxAuthNo
				session("GWTransId")=strVPSTxId
				session("GWTransType")=pcv_StrTxType
				Response.redirect "manpay-ok.asp?s=true&gw=SagePay"
			end if
			' ** Something has gone wrong, record the error and redirect etc.
			 strProtxErrorType=strStatus
			 strProtxErrorMsg=strStatusDetail
			'REJECTED, NOTAUTHED, ERROR redirect back to payment form
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "&strProtxErrorType&" - "&strProtxErrorMsg&"<br><br><a href="""&tempURL&"?psslurl=gwProtxVSP.aspDAANDidCustomer="&session("idCustomer")&"DAANDidOrder="&session("GWOrderId")&"DAANDamount="&request.querystring("amount")&"""><img src="""&rslayout("back")&"""></a>")&"&e1="&strProtxErrorType&"&e2="&strProtxErrorMsg&"&r="&tempURL&"?psslurl=gwProtxVSP.aspDAANDidCustomer="&session("idCustomer")&"DAANDidOrder="&session("GWOrderId")&"DAANDamount="&request.querystring("amount")
			' ** Write VPSTxID, SecurityKey, Status and StatusDetail to the screen, log file or database
			response.write "<b>Failed</b><br/>"
			response.end
		end if
	
		' ******************************************************************
		' remove the reference to the object
		set httpRequest = nothing
	end if
'*************************************************************************************
' END
'*************************************************************************************
end if 
%>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="">Quick Payment Page</h3>
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
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s"><div id="pcMain">
	<table class="pcMainTable" width="100%">
        <% '** A 3D-Auth response has been returned, so show the bank page inline if possible, or redirect to it otherwise
		if strPageState="3DRedirect" then %>
		<tr>
			<td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
				  <td class="subheader" align="center">3D-Secure Authentication with your Bank</td>
              </tr>
              <tr>
                <td valign="top">
					<table border="0" width="100%">
						<tr>
							<td width="80%">To increase the security of Internet transactions Visa and Mastercard have introduced 3D-Secure (like an online version of Chip and PIN). <br>
							  <br>
						    You have chosen to use a card that is part of the 3D-Secure scheme, so you wi
						    <iframe src="manpay-gwProtx_3DRedirect.asp?amount=<%=request.QueryString("amount")%>" name="3DIFrame" width="100%" height="500" frameborder="0">
                            <% 'Non-IFRAME browser support
					response.write "<SCRIPT LANGUAGE=""Javascript""> function OnLoadEvent() { document.form.submit(); }</" & "SCRIPT>" 		
					response.write "<html><head><title>3D Secure Verification</title></head>"

					response.write "<body OnLoad=""OnLoadEvent();"">"
					response.write "<FORM name=""form"" action=""" & strACSURL &""" method=""POST"">"
					response.write "<input type=""hidden"" name=""PaReq"" value=""" & strPAReq &"""/>"
					response.write "<input type=""hidden"" name=""TermUrl"" value=""" & strYourSiteFQDN & strVirtualDir & "/3DCallback.asp?VendorTxCode=" & strVendorTxCode & """/>"
					response.write "<input type=""hidden"" name=""MD"" value=""" & strMD &"""/>"

					response.write "<NOSCRIPT>" 
					response.write "<center><p>Please click button below to Authenticate your card</p><input type=""submit"" value=""Go""/></p></center>"
					response.write "</NOSCRIPT>"
					response.write "</form></body></html>"%>
                            </iframe>
						    ll need to authenticate yourself with your bank in the section below.</td>
							<td width="20%" align="center"><img src="images/vbv_logo_small.gif" alt="Verified by Visa"><BR><BR><img src="images/mcsc_logo.gif" alt="MasterCard SecureCode"></td>
						</tr>
					</table>
				</td>
              </tr>
			  
			  <tr>
                <td valign="top"><%
					'** Attempt to set up an inline frame here.  If we can't, set up a standard full page redirection **
					Session("MD")=strMD
					Session("PAReq")=strPAReq
					Session("ACSURL")=strACSURL
					Session("VendorTxCode")=strVendorTxCode
				%></td>
			  </tr>

			</table>
           </td>
           </tr>
        <% else %>
		<tr>
			<td>
				<form method="POST" autocomplete="off" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
				<input type="hidden" name="PaymentSubmitted" value="Go">
					<table class="pcShowContent" width="100%">
			
					<% if Msg<>"" then %>
						<tr valign="top"> 
							<td colspan="4">
								<div class="pcErrorMessage"><%=Msg%></div>
							</td>
						</tr>
					<% end if %>
					<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
						<tr>
							<td colspan="4" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="4"><p><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></p></td>
						</tr>
					<% end if %>
                    
                    <% if request.QueryString("pay") <> "" then %>
					<tr>
						<td colspan="4" class="pcSpacer"></td>
					</tr>
					<tr class="pcSectionTitle">
						<td colspan="4"><p><strong>Payment For: <%=request.QueryString("pay")%></strong></p></td>
					</tr>
                    <% end if %>
                    
                    
					<tr>
						<td colspan="4" class="pcSpacer"></td>
					</tr>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><strong><%=dictLanguage.Item(Session("language")&"_GateWay_1")%></strong></p></td>
                        <td colspan="2"><p><strong>Delivery Address</strong></p></td>
					</tr>
					<tr>
						<td colspan="4" class="pcSpacer"></td>
					</tr>
					<tr>
						<td><p>First Name</p></td><td><input name="mpBillFirstName" id="mpBillFirstName" type="text" tabindex="1" /></td><td><p>Same As Billing Address:</p></td><td><input name="mpDelChk" id="mpDelChk" type="checkbox" value="1" onchange="sameDel();" tabindex="10" />
					</tr>									
                    <tr>
						<td><p>Last Name</p></td><td><input name="mpBillLastName" id="mpBillLastName" type="text" tabindex="2" /></td>
                        <td><p>First Name</p></td><td><input name="mpDelFirstName" id="mpDelFirstName" type="text" tabindex="11" /></td>
					</tr>
					<tr>
						<td><p>Phone</p></td><td><input name="mpBillTel" id="mpBillTel" type="text" tabindex="3" /></td>
                        <td><p>Last Name</p></td><td><input name="mpDelLastName" id="mpDelLastName" type="text" tabindex="12" /></td>
					</tr>
					<tr>
						<td><p>Email</p></td><td><input name="mpBillEmail" id="mpBillEmail" type="text" tabindex="4" /></td>
                        <td><p>Phone</p></td><td><input name="mpDelTel" id="mpDelTel" type="text" tabindex="13" /></td>
					</tr>
					<tr>
						<td><p>Address Line 1</p></td><td><input name="mpBillAddress1" id="mpBillAddress1" type="text" tabindex="5" /></td>
                        <td><p>Address Line 1</p></td><td><input name="mpDelAddress1" id="mpDelAddress1" type="text" tabindex="14" /></td>
					</tr>
					<tr>
						<td><p>Address Line 2</p></td><td><input name="mpBillAddress2" id="mpBillAddress2" type="text" tabindex="6" /></td>
                        <td><p>Address Line 2</p></td><td><input name="mpDelAddress2" id="mpDelAddress2" type="text" tabindex="15" /></td>
					</tr>
					<tr>
						<td><p>Town / City</p></td><td><input name="mpBillCity" id="mpBillCity" type="text" tabindex="7" /></td>
                        <td><p>Town / City</p></td><td><input name="mpDelCity" id="mpDelCity" type="text" tabindex="16" /></td>
					</tr>					
                    <tr>
						<td><p>Post Code</p></td><td><input name="mpBillPCode" id="mpBillPCode" type="text" tabindex="8" /></td>
                        <td><p>Post Code</p></td><td><input name="mpDelPCode" id="mpDelPCode" type="text" tabindex="17" /></td>
					</tr>					
                    <tr>
						<td><p>Country</p></td><td><select name="mpBillCountry" id="mpBillCountry" tabindex="9">
						  <option value="AF" >Afghanistan</option>
						  <option value="AX" >Åland Islands</option>
						  <option value="AL" >Albania</option>
						  <option value="DZ" >Algeria</option>
						  <option value="AS" >American Samoa</option>
						  <option value="AD" >Andorra</option>
						  <option value="AO" >Angola</option>
						  <option value="AI" >Anguilla</option>
						  <option value="AQ" >Antarctica</option>
						  <option value="AG" >Antigua and Barbuda</option>
						  <option value="AR" >Argentina</option>
						  <option value="AM" >Armenia</option>
						  <option value="AW" >Aruba</option>
						  <option value="AU" >Australia</option>
						  <option value="AT" >Austria</option>
						  <option value="AZ" >Azerbaijan</option>
						  <option value="BS" >Bahamas</option>
						  <option value="BH" >Bahrain</option>
						  <option value="BD" >Bangladesh</option>
						  <option value="BB" >Barbados</option>
						  <option value="BY" >Belarus</option>
						  <option value="BE" >Belgium</option>
						  <option value="BZ" >Belize</option>
						  <option value="BJ" >Benin</option>
						  <option value="BM" >Bermuda</option>
						  <option value="BT" >Bhutan</option>
						  <option value="BO" >Bolivia</option>
						  <option value="BA" >Bosnia and Herzegovina</option>
						  <option value="BW" >Botswana</option>
						  <option value="BV" >Bouvet Island</option>
						  <option value="BR" >Brazil</option>
						  <option value="IO" >British Indian Ocean Territory</option>
						  <option value="BN" >Brunei Darussalam</option>
						  <option value="BG" >Bulgaria</option>
						  <option value="BF" >Burkina Faso</option>
						  <option value="BI" >Burundi</option>
						  <option value="CI" >C?te D'Ivoire</option>
						  <option value="KH" >Cambodia</option>
						  <option value="CM" >Cameroon</option>
						  <option value="CA" >Canada</option>
						  <option value="CV" >Cape Verde</option>
						  <option value="KY" >Cayman Islands</option>
						  <option value="CF" >Central African Republic</option>
						  <option value="TD" >Chad</option>
						  <option value="CL" >Chile</option>
						  <option value="CN" >China - Peoples Republic of</option>
						  <option value="CX" >Christmas Island</option>
						  <option value="CC" >Cocos (Keeling) Islands</option>
						  <option value="CO" >Colombia</option>
						  <option value="KM" >Comoros</option>
						  <option value="CG" >Congo</option>
						  <option value="CK" >Cook Islands</option>
						  <option value="CR" >Costa Rica</option>
						  <option value="HR" >Croatia</option>
						  <option value="CU" >Cuba</option>
						  <option value="CY" >Cyprus</option>
						  <option value="CZ" >Czech Republic</option>
						  <option value="DK" >Denmark</option>
						  <option value="DJ" >Djibouti</option>
						  <option value="DM" >Dominica</option>
						  <option value="DO" >Dominican Republic</option>
						  <option value="EC" >Ecuador</option>
						  <option value="EG" >Egypt</option>
						  <option value="SV" >El Salvador</option>
						  <option value="GQ" >Equatorial Guinea</option>
						  <option value="ER" >Eritrea</option>
						  <option value="EE" >Estonia</option>
						  <option value="ET" >Ethiopia</option>
						  <option value="FK" >Falkland Islands (Malvinas)</option>
						  <option value="FO" >Faroe Islands</option>
						  <option value="FJ" >Fiji</option>
						  <option value="FI" >Finland</option>
						  <option value="FR" >France</option>
						  <option value="GF" >French Guiana</option>
						  <option value="PF" >French Polynesia</option>
						  <option value="TF" >French Southern Territories</option>
						  <option value="GA" >Gabon</option>
						  <option value="GM" >Gambia</option>
						  <option value="GE" >Georgia</option>
						  <option value="DE" >Germany</option>
						  <option value="GH" >Ghana</option>
						  <option value="GI" >Gibraltar</option>
						  <option value="GR" >Greece</option>
						  <option value="GL" >Greenland</option>
						  <option value="GD" >Grenada</option>
						  <option value="GP" >Guadeloupe</option>
						  <option value="GU" >Guam</option>
						  <option value="GT" >Guatemala</option>
						  <option value="GG" >Guernsey</option>
						  <option value="GN" >Guinea</option>
						  <option value="GW" >Guinea-Bissau</option>
						  <option value="GY" >Guyana</option>
						  <option value="HT" >Haiti</option>
						  <option value="HN" >Honduras</option>
						  <option value="HK" >Hong Kong</option>
						  <option value="HU" >Hungary</option>
						  <option value="IS" >Iceland</option>
						  <option value="IN" >India</option>
						  <option value="ID" >Indonesia</option>
						  <option value="IR" >Iran - Islamic Republic Of</option>
						  <option value="IQ" >Iraq</option>
						  <option value="IE" >Ireland</option>
						  <option value="IM" >Isle of Man</option>
						  <option value="IL" >Israel</option>
						  <option value="IT" >Italy</option>
						  <option value="JM" >Jamaica</option>
						  <option value="JP" >Japan</option>
						  <option value="JE" >Jersey</option>
						  <option value="JO" >Jordan</option>
						  <option value="KZ" >Kazakhstan</option>
						  <option value="KE" >Kenya</option>
						  <option value="KI" >Kiribati</option>
						  <option value="KR" >Korea - Republic of</option>
						  <option value="KW" >Kuwait</option>
						  <option value="KG" >Kyrgyzstan</option>
						  <option value="LV" >Latvia</option>
						  <option value="LB" >Lebanon</option>
						  <option value="LS" >Lesotho</option>
						  <option value="LR" >Liberia</option>
						  <option value="LY" >Libyan Arab Jamahiriya</option>
						  <option value="LI" >Liechtenstein</option>
						  <option value="LT" >Lithuania</option>
						  <option value="LU" >Luxembourg</option>
						  <option value="MO" >Macao</option>
						  <option value="MK" >Macedonia</option>
						  <option value="MG" >Madagascar</option>
						  <option value="MW" >Malawi</option>
						  <option value="MY" >Malaysia</option>
						  <option value="MV" >Maldives</option>
						  <option value="ML" >Mali</option>
						  <option value="MT" >Malta</option>
						  <option value="MH" >Marshall Islands</option>
						  <option value="MQ" >Martinique</option>
						  <option value="MR" >Mauritania</option>
						  <option value="MU" >Mauritius</option>
						  <option value="YT" >Mayotte</option>
						  <option value="MX" >Mexico</option>
						  <option value="FM" >Micronesia - Federated States of</option>
						  <option value="MD" >Moldova - Republic of</option>
						  <option value="MC" >Monaco</option>
						  <option value="MN" >Mongolia</option>
						  <option value="ME" >Montenegro</option>
						  <option value="MS" >Montserrat</option>
						  <option value="MA" >Morocco</option>
						  <option value="MZ" >Mozambique</option>
						  <option value="MM" >Myanmar</option>
						  <option value="NA" >Namibia</option>
						  <option value="NR" >Nauru</option>
						  <option value="NP" >Nepal</option>
						  <option value="NL" >Netherlands</option>
						  <option value="AN" >Netherlands Antilles</option>
						  <option value="NC" >New Caledonia</option>
						  <option value="NZ" >New Zealand</option>
						  <option value="NI" >Nicaragua</option>
						  <option value="NE" >Niger</option>
						  <option value="NG" >Nigeria</option>
						  <option value="NU" >Niue</option>
						  <option value="NF" >Norfolk Island</option>
						  <option value="NO" >Norway</option>
						  <option value="MP" >Nothern Mariana Islands</option>
						  <option value="OM" >Oman</option>
						  <option value="PK" >Pakistan</option>
						  <option value="PW" >Palau</option>
						  <option value="PA" >Panama</option>
						  <option value="PG" >Papua New Guinea</option>
						  <option value="PY" >Paraguay</option>
						  <option value="PE" >Peru</option>
						  <option value="PH" >Philippines</option>
						  <option value="PN" >Pitcairn</option>
						  <option value="PL" >Poland</option>
						  <option value="PT" >Portugal</option>
						  <option value="PR" >Puerto Rico</option>
						  <option value="QA" >Qatar</option>
						  <option value="RE" >Réunion</option>
						  <option value="RO" >Romania</option>
						  <option value="RU" >Russian Federation</option>
						  <option value="RW" >Rwanda</option>
						  <option value="SH" >Saint Helena</option>
						  <option value="KN" >Saint Kitts and Nevis</option>
						  <option value="LC" >Saint Lucia</option>
						  <option value="PM" >Saint Pierre and Miquelon</option>
						  <option value="VC" >Saint Vincent and the Grenadines</option>
						  <option value="WS" >Samoa</option>
						  <option value="SM" >San Marino</option>
						  <option value="ST" >Sao Tome and Principe</option>
						  <option value="SA" >Saudi Arabia</option>
						  <option value="SN" >Senegal</option>
						  <option value="RS" >Serbia</option>
						  <option value="SC" >Seychelles</option>
						  <option value="SL" >Sierra Leone</option>
						  <option value="SG" >Singapore</option>
						  <option value="SK" >Slovakia</option>
						  <option value="SI" >Slovenia</option>
						  <option value="SB" >Solomon Islands</option>
						  <option value="SO" >Somalia</option>
						  <option value="ZA" >South Africa</option>
						  <option value="ES" >Spain</option>
						  <option value="LK" >Sri Lanka</option>
						  <option value="SD" >Sudan</option>
						  <option value="SR" >Suriname</option>
						  <option value="SJ" >Svalbard and Jan Mayen</option>
						  <option value="SZ" >Swaziland</option>
						  <option value="SE" >Sweden</option>
						  <option value="CH" >Switzerland</option>
						  <option value="SY" >Syrian Arab Republic</option>
						  <option value="TW" >Taiwan - Province Of China</option>
						  <option value="TJ" >Tajikistan</option>
						  <option value="TZ" >Tanzania - United Republic Of</option>
						  <option value="TH" >Thailand</option>
						  <option value="TL" >Timor-Leste</option>
						  <option value="TG" >Togo</option>
						  <option value="TK" >Tokelau</option>
						  <option value="TO" >Tonga</option>
						  <option value="TT" >Trinidad And Tobago</option>
						  <option value="TN" >Tunisia</option>
						  <option value="TR" >Turkey</option>
						  <option value="TM" >Turkmenistan</option>
						  <option value="TC" >Turks and Caicos Islands</option>
						  <option value="TV" >Tuvalu</option>
						  <option value="UG" >Uganda</option>
						  <option value="UA" >Ukraine</option>
						  <option value="AE" >United Arab Emirates</option>
						  <option value="GB" selected>United Kingdom</option>
						  <option value="US" >United States</option>
						  <option value="UY" >Uruguay</option>
						  <option value="UM" >US - Minor Outlying Islands</option>
						  <option value="UZ" >Uzbekistan</option>
						  <option value="VU" >Vanuatu</option>
						  <option value="VA" >Vatican City</option>
						  <option value="VE" >Venezuela</option>
						  <option value="VN" >VietNam</option>
						  <option value="VG" >Virgin Islands - British</option>
						  <option value="VI" >Virgin Islands - U.S.</option>
						  <option value="WF" >Wallis and Futuna Islands</option>
						  <option value="EH" >Western Sahara</option>
						  <option value="YE" >Yemen</option>
						  <option value="ZM" >Zambia</option>
						  <option value="ZW" >Zimbabwe</option>
						  </select></td>
                        <td><p>Country</p></td><td><select name="mpDelCountry" id="mpDelCountry" tabindex="18">
                          <option value="AF" >Afghanistan</option>
                          <option value="AX" >Åland Islands</option>
                          <option value="AL" >Albania</option>
                          <option value="DZ" >Algeria</option>
                          <option value="AS" >American Samoa</option>
                          <option value="AD" >Andorra</option>
                          <option value="AO" >Angola</option>
                          <option value="AI" >Anguilla</option>
                          <option value="AQ" >Antarctica</option>
                          <option value="AG" >Antigua and Barbuda</option>
                          <option value="AR" >Argentina</option>
                          <option value="AM" >Armenia</option>
                          <option value="AW" >Aruba</option>
                          <option value="AU" >Australia</option>
                          <option value="AT" >Austria</option>
                          <option value="AZ" >Azerbaijan</option>
                          <option value="BS" >Bahamas</option>
                          <option value="BH" >Bahrain</option>
                          <option value="BD" >Bangladesh</option>
                          <option value="BB" >Barbados</option>
                          <option value="BY" >Belarus</option>
                          <option value="BE" >Belgium</option>
                          <option value="BZ" >Belize</option>
                          <option value="BJ" >Benin</option>
                          <option value="BM" >Bermuda</option>
                          <option value="BT" >Bhutan</option>
                          <option value="BO" >Bolivia</option>
                          <option value="BA" >Bosnia and Herzegovina</option>
                          <option value="BW" >Botswana</option>
                          <option value="BV" >Bouvet Island</option>
                          <option value="BR" >Brazil</option>
                          <option value="IO" >British Indian Ocean Territory</option>
                          <option value="BN" >Brunei Darussalam</option>
                          <option value="BG" >Bulgaria</option>
                          <option value="BF" >Burkina Faso</option>
                          <option value="BI" >Burundi</option>
                          <option value="CI" >C?te D'Ivoire</option>
                          <option value="KH" >Cambodia</option>
                          <option value="CM" >Cameroon</option>
                          <option value="CA" >Canada</option>
                          <option value="CV" >Cape Verde</option>
                          <option value="KY" >Cayman Islands</option>
                          <option value="CF" >Central African Republic</option>
                          <option value="TD" >Chad</option>
                          <option value="CL" >Chile</option>
                          <option value="CN" >China - Peoples Republic of</option>
                          <option value="CX" >Christmas Island</option>
                          <option value="CC" >Cocos (Keeling) Islands</option>
                          <option value="CO" >Colombia</option>
                          <option value="KM" >Comoros</option>
                          <option value="CG" >Congo</option>
                          <option value="CK" >Cook Islands</option>
                          <option value="CR" >Costa Rica</option>
                          <option value="HR" >Croatia</option>
                          <option value="CU" >Cuba</option>
                          <option value="CY" >Cyprus</option>
                          <option value="CZ" >Czech Republic</option>
                          <option value="DK" >Denmark</option>
                          <option value="DJ" >Djibouti</option>
                          <option value="DM" >Dominica</option>
                          <option value="DO" >Dominican Republic</option>
                          <option value="EC" >Ecuador</option>
                          <option value="EG" >Egypt</option>
                          <option value="SV" >El Salvador</option>
                          <option value="GQ" >Equatorial Guinea</option>
                          <option value="ER" >Eritrea</option>
                          <option value="EE" >Estonia</option>
                          <option value="ET" >Ethiopia</option>
                          <option value="FK" >Falkland Islands (Malvinas)</option>
                          <option value="FO" >Faroe Islands</option>
                          <option value="FJ" >Fiji</option>
                          <option value="FI" >Finland</option>
                          <option value="FR" >France</option>
                          <option value="GF" >French Guiana</option>
                          <option value="PF" >French Polynesia</option>
                          <option value="TF" >French Southern Territories</option>
                          <option value="GA" >Gabon</option>
                          <option value="GM" >Gambia</option>
                          <option value="GE" >Georgia</option>
                          <option value="DE" >Germany</option>
                          <option value="GH" >Ghana</option>
                          <option value="GI" >Gibraltar</option>
                          <option value="GR" >Greece</option>
                          <option value="GL" >Greenland</option>
                          <option value="GD" >Grenada</option>
                          <option value="GP" >Guadeloupe</option>
                          <option value="GU" >Guam</option>
                          <option value="GT" >Guatemala</option>
                          <option value="GG" >Guernsey</option>
                          <option value="GN" >Guinea</option>
                          <option value="GW" >Guinea-Bissau</option>
                          <option value="GY" >Guyana</option>
                          <option value="HT" >Haiti</option>
                          <option value="HN" >Honduras</option>
                          <option value="HK" >Hong Kong</option>
                          <option value="HU" >Hungary</option>
                          <option value="IS" >Iceland</option>
                          <option value="IN" >India</option>
                          <option value="ID" >Indonesia</option>
                          <option value="IR" >Iran - Islamic Republic Of</option>
                          <option value="IQ" >Iraq</option>
                          <option value="IE" >Ireland</option>
                          <option value="IM" >Isle of Man</option>
                          <option value="IL" >Israel</option>
                          <option value="IT" >Italy</option>
                          <option value="JM" >Jamaica</option>
                          <option value="JP" >Japan</option>
                          <option value="JE" >Jersey</option>
                          <option value="JO" >Jordan</option>
                          <option value="KZ" >Kazakhstan</option>
                          <option value="KE" >Kenya</option>
                          <option value="KI" >Kiribati</option>
                          <option value="KR" >Korea - Republic of</option>
                          <option value="KW" >Kuwait</option>
                          <option value="KG" >Kyrgyzstan</option>
                          <option value="LV" >Latvia</option>
                          <option value="LB" >Lebanon</option>
                          <option value="LS" >Lesotho</option>
                          <option value="LR" >Liberia</option>
                          <option value="LY" >Libyan Arab Jamahiriya</option>
                          <option value="LI" >Liechtenstein</option>
                          <option value="LT" >Lithuania</option>
                          <option value="LU" >Luxembourg</option>
                          <option value="MO" >Macao</option>
                          <option value="MK" >Macedonia</option>
                          <option value="MG" >Madagascar</option>
                          <option value="MW" >Malawi</option>
                          <option value="MY" >Malaysia</option>
                          <option value="MV" >Maldives</option>
                          <option value="ML" >Mali</option>
                          <option value="MT" >Malta</option>
                          <option value="MH" >Marshall Islands</option>
                          <option value="MQ" >Martinique</option>
                          <option value="MR" >Mauritania</option>
                          <option value="MU" >Mauritius</option>
                          <option value="YT" >Mayotte</option>
                          <option value="MX" >Mexico</option>
                          <option value="FM" >Micronesia - Federated States of</option>
                          <option value="MD" >Moldova - Republic of</option>
                          <option value="MC" >Monaco</option>
                          <option value="MN" >Mongolia</option>
                          <option value="ME" >Montenegro</option>
                          <option value="MS" >Montserrat</option>
                          <option value="MA" >Morocco</option>
                          <option value="MZ" >Mozambique</option>
                          <option value="MM" >Myanmar</option>
                          <option value="NA" >Namibia</option>
                          <option value="NR" >Nauru</option>
                          <option value="NP" >Nepal</option>
                          <option value="NL" >Netherlands</option>
                          <option value="AN" >Netherlands Antilles</option>
                          <option value="NC" >New Caledonia</option>
                          <option value="NZ" >New Zealand</option>
                          <option value="NI" >Nicaragua</option>
                          <option value="NE" >Niger</option>
                          <option value="NG" >Nigeria</option>
                          <option value="NU" >Niue</option>
                          <option value="NF" >Norfolk Island</option>
                          <option value="NO" >Norway</option>
                          <option value="MP" >Nothern Mariana Islands</option>
                          <option value="OM" >Oman</option>
                          <option value="PK" >Pakistan</option>
                          <option value="PW" >Palau</option>
                          <option value="PA" >Panama</option>
                          <option value="PG" >Papua New Guinea</option>
                          <option value="PY" >Paraguay</option>
                          <option value="PE" >Peru</option>
                          <option value="PH" >Philippines</option>
                          <option value="PN" >Pitcairn</option>
                          <option value="PL" >Poland</option>
                          <option value="PT" >Portugal</option>
                          <option value="PR" >Puerto Rico</option>
                          <option value="QA" >Qatar</option>
                          <option value="RE" >Réunion</option>
                          <option value="RO" >Romania</option>
                          <option value="RU" >Russian Federation</option>
                          <option value="RW" >Rwanda</option>
                          <option value="SH" >Saint Helena</option>
                          <option value="KN" >Saint Kitts and Nevis</option>
                          <option value="LC" >Saint Lucia</option>
                          <option value="PM" >Saint Pierre and Miquelon</option>
                          <option value="VC" >Saint Vincent and the Grenadines</option>
                          <option value="WS" >Samoa</option>
                          <option value="SM" >San Marino</option>
                          <option value="ST" >Sao Tome and Principe</option>
                          <option value="SA" >Saudi Arabia</option>
                          <option value="SN" >Senegal</option>
                          <option value="RS" >Serbia</option>
                          <option value="SC" >Seychelles</option>
                          <option value="SL" >Sierra Leone</option>
                          <option value="SG" >Singapore</option>
                          <option value="SK" >Slovakia</option>
                          <option value="SI" >Slovenia</option>
                          <option value="SB" >Solomon Islands</option>
                          <option value="SO" >Somalia</option>
                          <option value="ZA" >South Africa</option>
                          <option value="ES" >Spain</option>
                          <option value="LK" >Sri Lanka</option>
                          <option value="SD" >Sudan</option>
                          <option value="SR" >Suriname</option>
                          <option value="SJ" >Svalbard and Jan Mayen</option>
                          <option value="SZ" >Swaziland</option>
                          <option value="SE" >Sweden</option>
                          <option value="CH" >Switzerland</option>
                          <option value="SY" >Syrian Arab Republic</option>
                          <option value="TW" >Taiwan - Province Of China</option>
                          <option value="TJ" >Tajikistan</option>
                          <option value="TZ" >Tanzania - United Republic Of</option>
                          <option value="TH" >Thailand</option>
                          <option value="TL" >Timor-Leste</option>
                          <option value="TG" >Togo</option>
                          <option value="TK" >Tokelau</option>
                          <option value="TO" >Tonga</option>
                          <option value="TT" >Trinidad And Tobago</option>
                          <option value="TN" >Tunisia</option>
                          <option value="TR" >Turkey</option>
                          <option value="TM" >Turkmenistan</option>
                          <option value="TC" >Turks and Caicos Islands</option>
                          <option value="TV" >Tuvalu</option>
                          <option value="UG" >Uganda</option>
                          <option value="UA" >Ukraine</option>
                          <option value="AE" >United Arab Emirates</option>
                          <option value="GB" selected>United Kingdom</option>
                          <option value="US" >United States</option>
                          <option value="UY" >Uruguay</option>
                          <option value="UM" >US - Minor Outlying Islands</option>
                          <option value="UZ" >Uzbekistan</option>
                          <option value="VU" >Vanuatu</option>
                          <option value="VA" >Vatican City</option>
                          <option value="VE" >Venezuela</option>
                          <option value="VN" >VietNam</option>
                          <option value="VG" >Virgin Islands - British</option>
                          <option value="VI" >Virgin Islands - U.S.</option>
                          <option value="WF" >Wallis and Futuna Islands</option>
                          <option value="EH" >Western Sahara</option>
                          <option value="YE" >Yemen</option>
                          <option value="ZM" >Zambia</option>
                          <option value="ZW" >Zimbabwe</option>
                        </select></td>
					</tr>					
                   <tr>
						<td colspan="4"><p><%=pcBillingAddress%></p></td>
					</tr>
					<% if pcBillingAddress2<>"" then %>
					<tr>
						<td colspan="4"><p><%=pcBillingAddress2%></p></td>
					</tr>
					<% end if %>
					<tr>
						<td colspan="4" class="pcSpacer"></td>
					</tr>
					<% if pcv_StrProtxTestmode<>0 then %>
						<tr>
							<td colspan="4"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						<tr>
							<td colspan="4" class="pcSpacer"></td>
						</tr>
					<% end if %>
					<tr class="pcSectionTitle">
						<td colspan="4"><p><strong><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></strong></p></td>
					</tr>
					<tr>
						<td colspan="4" class="pcSpacer"></td>
					</tr>
					<tr>
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
						<td colspan="3">
								<select name="ProtxCardTypes" tabindex="19">
									<%=strFormCardTypes%>
								</select>
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td colspan="3"> 
							<input type="text" name="CardNumber" autocomplete="off"  size="18" maxlength="18" value="" tabindex="20">
						</td>
					</tr>
					<% if pcv_StrProtxAmex=1 then %>
					<tr>
						<td><p>Start Date: (mmyy)</p></td>
					<td colspan="3">
						<input name="StartDate" autocomplete="off" type="text" size="6" maxlength="4" tabindex="21">
						<span class="pcSmallText">Required for some Maestro, Solo and Amex; <strong>mmyy</strong> format.</span>
						</td>
					</tr>
					<% end if %>
					<% if pcv_StrProtxMaestro=1 then %>
						<tr>
							<td><p>Issue Number:</p>
							</td>
						<td colspan="3">
							<input name="IssueNumber" autocomplete="off" type="text" size="4" maxlength="2">
							<span class="pcSmallText">Required some Maestro and Solo cards only.</span>
						</td>
						</tr>
					<% end if %>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td colspan="3"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="expMonth" tabindex="22">
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
							<select name="expYear" tabindex="23">
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
						</td>
					</tr>
					<% If pcv_CVV ="1" Then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td colspan="3"> 
								<input name="CVV" autocomplete="off" type="text" id="CVV" value="" size="4" maxlength="4" tabindex="24">
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td colspan="3"><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% End If %>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td colspan="3">&pound;<%=money(mpAmount)%></td>
					</tr>
					
					<tr> 
						<td>&nbsp;</td>
                        <td>
							<input type="submit" value="Make Payment" name="Continue" class="btn product-action pg-green-btn" id="submit" tabindex="25">
						<script type="text/javascript">
                            $(document).ready(function() {
                                $('#submit', this).attr('disabled', false);
                                $('form').submit(function(){
                                    $('#submit', this).attr('disabled', true);
                                    return 
                                });
                            });
                        </script> 
						</td><td colspan="2">&nbsp;</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
   <% end if %>
</table>
					</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->
<% 
'***********************************************
' Useful methods
'***********************************************

function findField( fieldName, postResponse )
  items = split( postResponse, chr( 13 ) )
  for idx = LBound( items ) to UBound( items )
    item = replace( items( idx ), chr( 10 ), "" )
    if InStr( item, fieldName & "=" ) = 1 then
      ' found
      findField = right( item, len( item ) - len( fieldName ) - 1 )
      Exit For
    end if
  next 
end function
%>
<!--#include file="footer_wrapper.asp"-->