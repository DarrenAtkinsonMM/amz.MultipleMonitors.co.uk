<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwAuthorizeCheckConfirm.asp"

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
<%
'//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT x_Type, x_Login, x_Password, x_Curcode, x_Method, x_AIMType, x_testmode, x_eCheck, x_secureSource, x_eCheckPending FROM authorizeNet Where id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_Type=rs("x_Type")
pcv_Login=rs("x_Login")
'decrypt
pcv_Login=enDeCrypt(pcv_Login, scCrypPass)
pcv_Password=rs("x_Password")
'decrypt
pcv_Password=enDeCrypt(pcv_Password, scCrypPass)
pcv_Curcode=rs("x_Curcode")
pcv_Method=rs("x_Method")
pcv_AIMType=rs("x_AIMType")
pcv_testmode=rs("x_testmode")
pcv_eCheck=rs("x_eCheck")
pcv_secureSource=rs("x_secureSource")
pcv_eCheckPending=rs("x_eCheckPending")
session("x_eCheckPending")=pcv_eCheckPending
pcv_TypeArray=Split(pcv_Type,"||")
pcv_Type1=pcv_TypeArray(0)

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	'expdate=expmonth & right(expyear, 2)
	Dim objXMLHTTP, xml
	
	'Send the request to the Authorize.NET processor.
	stext="x_version=3.1"
	stext=stext & "&x_delim_data=True"
	stext=stext & "&x_delim_char=,"
	stext=stext & "&x_method=ECHECK"
	if pcv_testmode="1" then
		stext=stext & "&x_Test_Request=True"
	else
		stext=stext & "&x_Test_Request=False"
	end if
	stext=stext & "&x_relay_response=FALSE"
	stext=stext & "&x_login=" & pcv_Login
	stext=stext & "&x_tran_key=" & pcv_Password
	stext=stext & "&x_amount=" & pcBillingTotal
	'check data
	stext=stext & "&x_bank_acct_name="& request.Form("x_bank_acct_name")
	stext=stext & "&x_bank_aba_code="& request.Form("x_bank_aba_code")
	stext=stext & "&x_bank_acct_num="& request.Form("x_bank_acct_num")
	stext=stext & "&x_bank_acct_type="& request.Form("x_bank_acct_type")
	stext=stext & "&x_bank_name="& request.Form("x_bank_name")
	stext=stext & "&x_customer_tax_id="& request.Form("x_customer_tax_id")
	if request.Form("x_customer_tax_id")="" then
		stext=stext & "&x_drivers_license_num="& request.Form("x_drivers_license_num")
		stext=stext & "&x_drivers_license_state="& request.Form("x_drivers_license_state")
		stext=stext & "&x_drivers_license_dob="& request.Form("x_drivers_license_dob")
	end if
	stext=stext & "&x_customer_ip=" & pcCustIpAddress
	if pcv_secureSource="1" then
		stext=stext & "&x_customer_organization_type=" & request.Form("customer_organization_type")
	end if
	stext=stext & "&x_type=AUTH_CAPTURE"
	stext=stext & "&x_echeck_type=WEB"
	stext=stext & "&x_recurring_billing=NO"
	stext=stext & "&x_Currency_Code=" & pcv_Curcode
	stext=stext & "&x_Description=" & replace(scCompanyName,",","-") & " Order: " & session("GWOrderID")
	stext=stext & "&x_Invoice_Num=" & session("GWOrderID")
	stext=stext & "&x_Cust_ID=" & session("idCustomer")
	stext=stext & "&x_first_name=" & pcBillingFirstName
	stext=stext & "&x_last_name=" & pcBillingLastName
	stext=stext & "&x_company=" & replace(pcBillingCompany,",","||")
	stext=stext & "&x_address=" & replace(pcBillingAddress,",","||")
	stext=stext & "&x_city=" & pcBillingCity
	stext=stext & "&x_state=" & pcBillingState
	stext=stext & "&x_zip=" & pcBillingPostalCode
	stext=stext & "&x_country=" & pcBillingCountryCode
	stext=stext & "&x_phone=" & pcBillingPhone
	stext=stext & "&x_email=" & pcCustomerEmail
	stext=stext & "&x_Ship_To_First_Name=" & pcShippingFirstName
	stext=stext & "&x_Ship_To_Last_Name=" & pcShippingLastName
	stext=stext & "&x_Ship_To_Address=" & replace(pcShippingAddress,",","||")
	stext=stext & "&x_Ship_To_City=" & pcShippingCity
	stext=stext & "&x_Ship_To_State=" & pcShippingState
	stext=stext & "&x_Ship_To_Zip=" & pcShippingPostalCode
	stext=stext & "&x_Ship_To_Country=" & pcShippingCountryCode
	
	'Send the transaction info as part of the querystring
	set xml =  Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.open "POST", "https://secure.authorize.net/gateway/transact.dll?" & stext & "", false
	xml.send ""
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText
	Set xml = Nothing
	
	strArrayVal = split(strRetVal, ",", -1)
	session("x_response_code")=strArrayVal(0)
	session("x_response_subcode")=strArrayVal(1)
	session("x_response_reason_code")=strArrayVal(2)
	session("x_response_reason_text")=strArrayVal(3)
	session("GWAuthCode")=strArrayVal(4)    '6 digit approval code
	session("x_avs_code")=strArrayVal(5)
	session("GWTransId")=strArrayVal(6)    'transaction id
	session("x_invoice_num")=strArrayVal(7)
	session("x_description")=strArrayVal(8)
	session("x_amount")=strArrayVal(9)
	session("x_method")=strArrayVal(10)
	session("x_type")=strArrayVal(11)
	session("x_cust_id")=strArrayVal(12)
	session("x_first_name")=strArrayVal(13)
	session("x_last_name")=strArrayVal(14)
	pcv_company = strArrayVal(15)
	session("x_address")=strArrayVal(16)
	pcv_city = strArrayVal(17)
	pcv_state= strArrayVal(18)
	session("x_zip")=strArrayVal(19)
	
	pcv_country                 = strArrayVal(20)
	pcv_phone                   = strArrayVal(21)
	pcv_fax                     = strArrayVal(22)
	pcv_email                   = strArrayVal(23)
	pcv_ship_to_first_name      = strArrayVal(24)
	pcv_ship_to_last_name       = strArrayVal(25)
	pcv_ship_to_company         = strArrayVal(26)
	pcv_ship_to_address         = strArrayVal(27)
	pcv_ship_to_city            = strArrayVal(28)
	pcv_ship_to_state           = strArrayVal(29)
	pcv_ship_to_zip             = strArrayVal(30)
	pcv_ship_to_country         = strArrayVal(31)

	'Check the ErrorCode to make sure that the component was able to talk to the authorization network
	If (strStatus <> 200) Then
		'Log failed transaction
		call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
		
		Response.Write "An error occurred during processing. Please try again later."
	else
		'save and update order 
		If session("x_response_code") = 1 Then
			'save info in authOrders
			tordnum=(int(session("x_invoice_num"))-scpre)
			
			query="INSERT INTO authorders (idOrder, amount, paymentmethod, transtype, authcode, ccnum, ccexp, idCustomer, fname, lname, address, zip, captured) VALUES ("&tordnum&", "&session("x_amount")&", 'ECHECK', '"&x_Type1&"', '"&session("GWAuthCode")&"','1111111111111111','0000',"&session("x_cust_id")&",'"&replace(session("x_first_name"),"'","''")&"', '"&replace(session("x_last_name"),"'","''")&"', '"&replace(session("x_address"),"'","''")&"', '"&session("x_zip")&"',0);"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			'Log successful transaction
			call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 1)
			
			call closedb()
			Response.redirect "gwReturn.asp?s=true&gw=AIM&c=true"
			
		elseif session("x_response_code")<>1 then
      
			'Log failed transaction
			call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
      
			call closeDb()
			Session("message") = "Error code " & session("x_response_code")&": "& session("x_response_reason_text")
			Session("backbuttonURL") = tempURL & "?psslurl=gwAuthorizeCheck.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
			response.redirect "msgb.asp?back=1"
		End If
	end if

	'*************************************************************************************
	' END
	'*************************************************************************************
end if 
x_bank_acct_name = request.Form("x_bank_acct_name")
x_bank_aba_code = request.Form("x_bank_aba_code")
x_bank_acct_num = request.Form("x_bank_acct_num")
x_bank_acct_type = request.Form("x_bank_acct_type")
x_bank_name = request.Form("x_bank_name")
x_customer_tax_id = request.Form("x_customer_tax_id")
x_drivers_license_num = request.Form("x_drivers_license_num")
x_drivers_license_state = request.Form("x_drivers_license_state")
x_drivers_license_dob = request.Form("x_drivers_license_dob")
if pcv_secureSource="1" then
	x_customer_organization_type = request.Form("customer_organization_type")
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
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_12")%></div>
                <div class="pcFormField">
                    <%=x_bank_acct_name%><input name="x_bank_acct_name" type="hidden" size="35" maxlength="50" value="<%=x_bank_acct_name%>">
                </div> 
            </div>
            
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_13")%></div>
                <div class="pcFormField">
                    <%=x_bank_aba_code%><input name="x_bank_aba_code" type="hidden" size="35" value="<%=x_bank_aba_code%>">
                </div> 
            </div>
            
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_14")%></div>
                <div class="pcFormField">
                    <%=x_bank_acct_num%><input name="x_bank_acct_num" type="hidden" size="35" value="<%=x_bank_acct_num%>">
                </div> 
            </div>

            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_15")%></div>
                <div class="pcFormField">
                    <%=x_bank_acct_type%><input name="x_bank_acct_type" type="hidden" size="35" value="<%=x_bank_acct_type%>">
                </div> 
            </div>
            
            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_16")%></div>
                <div class="pcFormField">
                    <%=x_bank_name%><input name="x_bank_name" type="hidden" size="20" maxlength="20" value="<%=asdf%>">
                </div> 
            </div>
            
            
            <% if x_secureSource="1" then %>

            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_17")%></div>
                <div class="pcFormField">
                    <%=x_customer_organization_type%><input type="hidden" name="x_customer_organization_type" value="<%=x_customer_organization_type%>">
                </div> 
            </div>
                
                
            <% end if %>
            <% if request.Form("x_customer_tax_id")<>"" then %>
            
                    <div class="pcFormItem"> 
                        <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_19")%></div>
                        <div class="pcFormField">
                            <%=x_customer_tax_id%><input name="x_customer_tax_id" type="hidden" size="9" maxlength="9" value="<%=x_customer_tax_id%>">
                        </div> 
                    </div>

              <% else %>
              
                    <div class="pcFormItem"> 
                        <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_21")%></div>
                        <div class="pcFormField">
                            <%=x_drivers_license_num%> <input name="x_drivers_license_num" type="hidden" size="35" maxlength="50" value="<%=x_drivers_license_num%>">
                        </div> 
                    </div>
                    
                    <div class="pcFormItem"> 
                        <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_22")%></div>
                        <div class="pcFormField">
                            <%=x_drivers_license_state%><input name="x_drivers_license_state" type="hidden" size="2" maxlength="2" value="<%=x_drivers_license_state%>">
                        </div> 
                    </div>
                    
                    <div class="pcFormItem"> 
                        <div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_24")%></div>
                        <div class="pcFormField">
                            <%=x_drivers_license_dob%><input name="x_drivers_license_dob" type="hidden" size="10" maxlength="10" value="<%=x_drivers_license_dob%>">
                        </div> 
                    </div>

            <% end if %>
            
            <div class="pcFormItem"> 
			    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
                <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
            </div>
 
            <div class="pcInfoMessage">
                By clicking the &quot;Place Order&quot; button below, I authorize <%= scCompanyName %> to charge my <%=x_bank_acct_type%> account on <%=Now()%> for the amount of <%=scCurSign&money(pcBillingTotal)%>.
            </div> 
                    
            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
