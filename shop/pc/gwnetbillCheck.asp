<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwNetBillCheck.asp"

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
query="SELECT NBAccountID,NBCVVEnabled,NBAVS,NBSiteTag FROM netbill Where idNetbill=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
NBAccountID=rs("NBAccountID")
'decrypt
NBAccountID=enDeCrypt(NBAccountID, scCrypPass)
pcv_CVV=rs("NBCVVEnabled")
NBAVS=rs("NBAVS")
NBTranType="S"
NBSiteTag=rs("NBSiteTag")
If NBAVS ="1" Then
	NBAVS="y"
Else 
	NBAVS="n"   
End If

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim objHTTP, strRequest, DLLUrl
	
	DLLUrl="https://secure.netbilling.com:1402/gw/sas/direct3.0"
				
	strRequest ="pay_type="&"K"&_
	"&account_id="&NBAccountID&_
	"&site_tag="&NBSiteTag&_
	"&tran_type="&NBTranType&_
	"&amount="&pcBillingTotal&_   
	"&account_number="&request.Form("bank_aba_code")&"%3A"&request.Form("bank_acct_num")&_ 
	"&bill_name1="&pcBillingFirstName&_
	"&bill_name2="&pcBillingLastName&_
	"&bill_street="&pcBillingAddress&_
	"&bill_city="&pcBillingCity&_
	"&bill_state="&pcBillingState&_
	"&bill_zip="&pcBillingPostalCode&_
	"&bill_country="&pcBillingCountryCode&_
	"&cust_phone="&pcBillingPhone&_
	"&cust_email="&pcCustomerEmail&_
	"&ship_name1"&pcShippingFirstName&" "&pcShippingLastName&_
	"&ship_street="&pcShippingAddress&_
	"&ship_city="&pcShippingCity&_
	"&ship_state="&pcShippingState&_
	"&ship_zip"&pcShippingPostalCode&_
	"&ship_country="&pcShippingCountryCode&_
	"&cust_ip="&pcCustIpAddress
			
	set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP"&scXML)
	objHttp.open "POST", DLLUrl, false
	objHttp.setRequestHeader "Host", "secure.netbilling.com:1402"
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objHttp.setRequestHeader "Content-Length", Len(strRequest)
	objHttp.Send strRequest

	If objHTTP.Status = 200 Then  ' HTTP_STATUS_OK=200 
		Dim objDictResponse, intDelimiterPos, ResponseArray
		Dim strResponse, strNameValuePair, strName, strValue
		strResponse = replace(objHTTP.ResponseText,chr(10)," ")
		strResponse = rtrim(strResponse)
					
		' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
		ResponseArray = Split(strResponse, "&")
		Set objDictResponse = server.createobject("Scripting.Dictionary")
		For each ResponseItem in ResponseArray
			NameValue = Split(ResponseItem, "=")
			objDictResponse.Add NameValue(0), NameValue(1)
		Next
					
		' Parse the response into local vars
		strstatus=objDictResponse.Item("status_code")
		strstatusmsg=objDictResponse.Item("auth_msg")
		strAuthorizationNumber=objDictResponse.Item("auth_date")
		strTransactionID=objDictResponse.Item("trans_id")
			
		If strstatus = "1" and strTransactionID <> "" Then 
			session("GWAuthCode")=strAuthorizationNumber
			session("GWTransId")=strTransactionID
			Response.redirect "gwReturn.asp?s=true&gw=NetbillCheck"
		Else

			if strstatus = "0" then
				strstatus = "Failed transaction"
			end if
			
            call closeDb()
            Session("message") = strstatus&"&nbsp;due to&nbsp;"&lcase(strstatusmsg)
            Session("backbuttonURL") = tempURL & "?psslurl=gwNetBillCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1"

		End If  
	Else   
		strStatus= "Transaction Failed...<BR>"
		strStatus= strStatus&"<BR> Https Response <BR>" 
		strStatus= strStatus&"Status Code= " & objHttp.status       & "<BR>"
		strStatus= strStatus&"Error Status = " & objHttp.statusText   & "<BR>"

        call closeDb()
        Session("message") = strStatus
        Session("backbuttonURL") = tempURL & "?psslurl=gwNetBillCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"
		
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
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_12")%></div>
                <div class="pcFormField">
                    <input name="bank_acct_name" type="text" size="35" maxlength="50">
                </div> 
            </div>
  
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_13")%></div>
                <div class="pcFormField">
                    <input name="bank_aba_code" type="text" size="35">
                </div> 
            </div>
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_14")%></div>
                <div class="pcFormField">
                    <input name="bank_acct_num" type="text" size="35" autocomplete="off">
                </div> 
            </div>     
      
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_27")%></div>
                <div class="pcFormField">
                    <input name="check_num" type="text" size="35">
                </div> 
            </div>   
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_15")%></div>
                <div class="pcFormField">
                    <select name="bank_acct_type">
								<option value="CHECKING">Checking Account</option>
								<option value="SAVINGS">Savings Account</option>
							</select>
                </div> 
            </div>           
                          <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_16")%></div>
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
