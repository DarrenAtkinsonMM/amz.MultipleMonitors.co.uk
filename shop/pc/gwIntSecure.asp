<% 'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Check if this is a post-back
pcv_Response_IdOrder=request("xxxVar1")
if pcv_Response_IdOrder<>"" then
	if session("GWOrderId")="" then
		session("GWOrderId")=pcv_Response_IdOrder
	end if
else
	'//Get Order ID
	if session("GWOrderId")="" then
		session("GWOrderId")=request("idOrder")
	end if
end if
%>

<% '//Retrieve customer data from the database using the current session id	 
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->

<%if pcv_Response_IdOrder<>"" then
	pcv_Response_ApprovalCode=request("ApprovalCode")
	pcv_Response_ReceiptNumber=request("receiptnumber")

	'if pcv_Response_StatusCode="F" OR pcv_Response_StatusCode="0" OR pcv_Response_StatusCode="D" then
	'	Msg=pcv_Response_AuthMessage
	'end if
	session("GWAuthCode")=pcv_Response_ApprovalCode
	session("GWTransId")=pcv_Response_ReceiptNumber
	if session("GWOrderId")="" then
		session("GWOrderId")=pcv_Response_IdOrder
	end if
	session("GWSessionID")=Session.SessionID 
 
	'//Set customer session - we may now be on a different server where this session was lost
	session("idCustomer")=pcIdCustomer
	rtPcBillingAddress=Request("xxxAddress")
	rtPcBillingPostalCode=Request("xxxPostal")
	rtPcBillingProvince=Request("xxxProvince")
	rtPcBillingCity=Request("xxxCity")
	rtPcBillingCountryCode=Request("xxxCountry")
	rtPcBillingCompany = Request("xxxCompany")
	rtPcCustomerEmail = request("xxxEmail")
	rtPcBillingPhone = request("xxxPhone") 
	rtPcCustName = Request("xxxName")
	
	if Ucase(trim(rtPcCustName)) <> Ucase(trim(pcBillingFirstName&" "&pcBillingLastName)) or Ucase(trim(rtPcBillingAddress )) <> Ucase(trim( PcBillingAddress)) or Ucase(trim( rtPcBillingPostalCode)) <> Ucase(trim( PcBillingPostalCode)) or Ucase(trim( rtPcBillingProvince )) <> Ucase(trim( pcBillingState)) or Ucase(trim( rtPcBillingCity )) <> Ucase(trim( PcBillingCity)) or Ucase(trim( rtPcBillingCountryCode )) <> Ucase(trim( PcBillingCountryCode)) or Ucase(trim( rtPcCustomerEmail )) <> Ucase(trim( PcCustomerEmail)) or Ucase(trim( rtPcBillingPhone )) <> Ucase(trim( PcBillingPhone))  Then
		if Ucase(trim(rtPcCustName)) <> Ucase(trim(pcBillingFirstName&" "&pcBillingLastName)) then
		adminComments ="Name on Card: " & rtpcCustName & " does not match Account Name: "&(pcBillingFirstName&" "&pcBillingLastName)&"."&vbcrlf
		else
		 adminComments ="Name on Card: "&rtpcCustName&"."&vbcrlf
		end if 
		if Ucase(trim(rtPcBillingAddress)) <> Ucase(trim(PcBillingAddress)) then
		adminComments = adminComments &"Billing Address: "& rtPcBillingAddress & " does not match Account Billing Address: "&PcBillingAddress&"."&vbcrlf
		else
adminComments = adminComments &"Billing Address: "&PcBillingAddress&"."&vbcrlf
		end if 
		if Ucase(trim(rtPcBillingPostalCode)) <> Ucase(trim(PcBillingPostalCode)) then
		adminComments = adminComments &"Billing Zip: "& rtPcBillingPostalCode & " does not match Account Billing Zip: "&PcBillingPostalCode&"."&vbcrlf
		else
adminComments = adminComments &"Billing Zip: "&PcBillingPostalCode&"."&vbcrlf
		end if 		
		if Ucase(trim(rtPcBillingProvince)) <> Ucase(trim(pcBillingState)) then
		adminComments = adminComments &"Billing Province/State: "& rtPcBillingProvince & " does not match Account Billing Province/State: "&pcBillingState&"."&vbcrlf
		else
		 adminComments = adminComments &"Billing Province/State: "&pcBillingState&"."&vbcrlf
		end if 
		if Ucase(trim(rtPcBillingCity)) <> Ucase(trim(PcBillingCity)) then
		adminComments = adminComments &"Billing City: "& rtPcBillingCity & " does not match Account Billing City: "&PcBillingCity&"."&vbcrlf
		else
adminComments = adminComments &"Billing City: "&PcBillingCity&"."&vbcrlf
		end if 
		if Ucase(trim(rtPcCustomerEmail)) <> Ucase(trim(PcCustomerEmail)) then
		adminComments = adminComments &"Billing Email: "& rtPcCustomerEmail & " does not match Account Billing Email: "&PcCustomerEmail&"."&vbcrlf
		else
		 adminComments = adminComments &"Billing Email: "&PcCustomerEmail&"."&vbcrlf
		end if 
		if Ucase(trim(rtPcBillingPhone)) <> Ucase(trim(PcBillingPhone)) then
adminComments = adminComments &"Billing Phone: "& rtPcBillingPhone & " does not match Account Billing Phone: "&PcBillingPhone&"."&vbcrlf
		else
		 adminComments = adminComments &"Billing Phone: "&rtPcBillingPhone&"."&vbcrlf
		end if 

    
query="UPDATE orders SET adminComments=N'"& getUserInput(adminComments,0) &"'"
query=query &" WHERE idOrder="&pcGatewayDataIdOrder&";"
			
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)		
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if	
		
    End if
	
    call closedb()
    Response.redirect "gwReturn.asp?s=true&gw=InternetSecure"
end if

'//Set redirect page to the current file name
session("redirectPage")="gwIntSecure.asp"
session("redirectPage2")="https://secure.internetsecure.com/process.cgi"

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
	 
 '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT IsMerchantNumber, IsLanguage, IsCurrency, IsTestmode FROM InternetSecure WHERE IsID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
IsMerchantNumber=rs("IsMerchantNumber")
IsLanguage=rs("IsLanguage")
IsCurrency=rs("IsCurrency")
IsTestmode=rs("IsTestmode")

set rs=nothing

'//Get logo
Select case IsLanguage

	case "EN"
		IsLogo="<A HREF='https://www.internetsecure.com/cgi-bin/certified.mhtml?merchant_number="&IsMerchantNumber&"' target='_blank'><IMG ALIGN=CENTER SRC='" & pcf_getImagePath("https://www.internetsecure.com/images","ismerch.gif") & "' BORDER=0 WIDTH=134 HEIGHT=33></A>"
	case "SP"
		IsLogo="<A HREF='https://www.internetsecure.com/cgi-bin/certified.mhtml?merchant_number="&IsMerchantNumber&"' target='_blank'><IMG ALIGN=CENTER SRC='" & pcf_getImagePath("https://www.internetsecure.com/images","ismerch.gif") & "' BORDER=0 WIDTH=134 HEIGHT=33></A>"
	case "JP"
		IsLogo="<A HREF='https://www.internetsecure.com/cgi-bin/certified.mhtml?merchant_number="&IsMerchantNumber&"' target='_blank'><IMG ALIGN=CENTER SRC='" & pcf_getImagePath("https://www.internetsecure.com/images","ismerch.gif") & "' BORDER=0 WIDTH=134 HEIGHT=33></A>"
	case "FR" 
		IsLogo="<A HREF='https://www.internetsecure.com/cgi-bin/certified.mhtml?merchant_number="&IsMerchantNumber&"&language=FR' target='_blank'><IMG ALIGN=CENTER SRC='" & pcf_getImagePath("https://www.internetsecure.com/images","ismer-fr.gif") & "' BORDER=0 WIDTH=134 HEIGHT=33></A>"
	case "PT"
		IsLogo="<A HREF='https://www.internetsecure.com/cgi-bin/certified.mhtml?merchant_number="&IsMerchantNumber&"&language=PT' target='_blank'><IMG ALIGN=CENTER SRC='" & pcf_getImagePath("https://www.internetsecure.com/images","ismer.gif") & "' BORDER=0 WIDTH=134 HEIGHT=33></A>"
	 
end select

'//Create productcart string
prdString=""
prdString=prdString&money(pcBillingTotal)&"::1::001::Online Sales::"
if IsCurrency="USD" then
	prdString=prdString&"{US}"
end if
if IsTestmode="1" then
	prdString=prdString&"{TEST}{TESTD}"
end if


'"Price::Qty::Code::Description::Flags|9.95::1::T001::Extra Large Green InternetSecure T-shirt.::{GST}{PST}{HST}{US}|10.00::1::shp::Overnight Shipping::{GST}{PST}{HST}{US} 
prdString="Price::Qty::Code::Description::Flags"
%>
<!--#include file="pcPay_InternetSecure_Itemize.asp"-->

<%
pcStrReturnCGIURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwIntSecure.asp"),"//","/")
pcStrReturnCGIURL=replace(pcStrReturnCGIURL,"http:/","http://")
pcStrReturnCGIURL=replace(pcStrReturnCGIURL,"https:/","https://")
%>
<div id="pcMain">
	<div class="pcMainContent">
				<form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="PaymentSubmitted" value="Go">

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                        
			<input type="hidden" name="MerchantNumber" value="<%=IsMerchantNumber%>">
			<input type="hidden" name="language" value="<%=IsLanguage%>">
			<input type="hidden" name="Products" value="<%=prdString %>">
			<input type="hidden" name="xxxVar1" value="<%=session("GWOrderId")%>">
			<input type="hidden" name="ReturnCGI" value="<%=pcStrReturnCGIURL%>">
			<input type="hidden" name="xxxName" value="<%=pcBillingFirstName&" "&pcBillingLastName%>">
			<input type="hidden" name="xxxAddress" value="<%=pcBillingAddress%>">
			<input type="hidden" name="xxxCity" value="<%=pcBillingCity%>">
			<input type="hidden" name="xxxProvince" value="<%=pcBillingState%>">
			<input type="hidden" name="xxxPostal" value="<%=pcBillingPostalCode%>">
			<input type="hidden" name="xxxPhone" value="<%=pcBillingPhone%>">
			<input type="hidden" name="xxxEmail" value="<%=pcCustomerEmail%>">
			<input type="hidden" name="xxxCountry" value="<%=pcBillingCountryCode%>">
            <% call pcs_showBillingAddress %>

            <div class="pcFormItem"> 
                Continue with payment through Internet Secure payment system.
            </div>
            <div class="pcFormItem"> 
                <%=IsLogo%>
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
