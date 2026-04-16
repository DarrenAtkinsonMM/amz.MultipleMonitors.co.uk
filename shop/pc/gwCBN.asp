<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->

<% if request.QueryString("response")<>"" then
	pCBN_response = request.QueryString("response")
	pCBN_idOrder = request.QueryString("idOrder")
	pCBN_status = request.QueryString("status")
	pCBN_responseArray = split(pCBN_response,",",-1,1)
	pCBN_response = pCBN_responseArray(0)
	pCBN_approval = pCBN_responseArray(1)
	select case pCBN_response
		case "RSP0000" ' Transaction approved
			pCBN_redirect = 1
		case "RSP0002" ' Account is setup as "Verification Only" with CrossCheck and check is recommended
			pCBN_redirect = 1
		case "RSP0003" ' Account is setup as "Verification Only" with CrossCheck and check is not recommeded
			pCBN_redirect = 0
			pCBN_response = "RSP0001"
		case "RSP0010" ' Test complete
			pCBN_redirect = 1
	end select
	if pCBN_redirect = 1 then
		'// go to gwReturn.asp
		session("GWAuthCode")=pCBN_approval
		session("GWTransId")=pCBN_idOrder
		session("GWTransType")=pCBN_status
		session("GWSessionID")=Session.SessionID 
		response.redirect "gwReturn.asp?s=true&gw=CBN"
	else
		if pCBN_response <> "" then
			select case pCBN_response
				case "RSP0001"
					Msg = "The transaction was declined."
				case "RSP0051"
					Msg = "Invalid merchant account. Contact the store administrator and have them review their ChecksByNet settings."
				case "RSP0020"
					Msg = "Duplicate check. The check number your provided has already been used."
				case "RSP1101"
					Msg = "The check number was blank or invalid. Make sure the number is greater than 99."
				case "RSP1102"
					Msg = "The order amount was blank or invalid."
				case "RSP1201"
					Msg = "Your name was missing or invalid."
				case "RSP1202"
					Msg = "Your address was missing or invalid."
				case "RSP1203"
					Msg = "Your city was missing or invalid."
				case "RSP1204"
					Msg = "Your state was missing or invalid"
				case "RSP1205"
					Msg = "Your zip code was missing or invalid"
				case "RSP1301"
					Msg = "The bank name was missing or invalid"
				case "RSP1302"
					Msg = "The bank city was missing or invalid"
				case "RSP1303"
					Msg = "The bank state was missing or invalid"
				case "RSP1304"
					Msg = "The bank zip was missing or invalid"
				case "RSP1311"
					Msg = "The bank routing number was missing or invalid"
				case "RSP1312"
					Msg = "The bank account number was missing or invalid"
				case "RSP1313"
					Msg = "The bank rounting and/or account number was missing or invalid"
				case "RSP1401"
					Msg = "Your Driver License number was missing or invalid"
				case "RSP1501"
					Msg = "Your phone number was missing or invalid"
				case "RSP1502"
					Msg = "Your e-mail address was missing or invalid"
			end select
		end if
		if session("GWOrderId")="" then
			session("GWOrderId")=pCBN_idOrder
		end if
	end if
end if

'//Set redirect page to the current file name
session("redirectPage")="gwCBN.asp"
session("redirectPage2")="https://cross.checksbynet.com/response.asp"

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
query="SELECT pcPay_CBN_merchant, pcPay_CBN_test, pcPay_CBN_status FROM pcPay_CBN WHERE pcPay_CBN_id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcvPay_CBN_merchant = rs("pcPay_CBN_merchant")
pcvPay_CBN_test = rs("pcPay_CBN_test")
pcvPay_CBN_status = rs("pcPay_CBN_status")

If pcvPay_CBN_test = 1 Then
	pcv_checknumber = "123"
	pcv_micr = "S123456780S67890S123"
	pcv_driverl = "12345"
	pcv_driverlst = "ZZ"
	pcBillingTotal = "2"
end if

set rs=nothing

ScriptURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwCBN.asp?idOrder="&session("GWOrderID")&"&status="&pcvPay_CBN_status),"//","/")
ScriptURL=replace(ScriptURL,"http:/","http://")
ScriptURL=replace(ScriptURL,"https:/","https://")

%>
<div id="pcMain">
	<div class="pcMainContent">
    
    <form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
        <input type="hidden" name="paytoid" value="<%=pcvPay_CBN_merchant%>">
        <input type="hidden" name="ScriptURL" value="<%=ScriptURL%>">
        <input type="hidden" name="memo" value="<%=scCompanyName%> - Order #: <%=session("GWOrderID")%>">
        <input type="hidden" name="idCustomer" value="<%=session("idCustomer")%>">
        <input type="hidden" name="idorder" value="<%=session("GWOrderID")%>">
        <input type="hidden" name="email" value="<%=pcCustomerEmail%>">
        <input type="hidden" name="checkamt" value="<%=pcBillingTotal%>">
            
        <input type="hidden" name="writerfirst" value="<%=pcBillingFirstName%>" size="20" maxlength="15">
        <input type="hidden" name="writerlast" value="<%=pcBillingLastName%>" size="20" maxlength="29">
        <input type="hidden" name="writeraddr" value="<%= pcBillingAddress%>" size="30" maxlength="50">
        <input type="hidden" name="writercity" value="<%= pcBillingCity%>" size="20" maxlength="30">
        <input type="hidden" name="writerst" value="<%=pcBillingState%>" size="3" maxlength="2">
        <input type="hidden" name="writerzip" value="<%=pcBillingPostalCode%>" size="6" maxlength="5">
        <input type="hidden" name="phone" value="<%=pcBillingPhone%>" size="14" maxlength="14">
        
        <% if Msg<>"" then %>
            <div class="pcErrorMessage"><%=Msg%></div>
        <% end if %>
        
        <% call pcs_showBillingAddress %>                    
        
        <div class="pcFormItem"> 
            <div class="pcFormLabel"><input name="checknbr" value="<%=pcv_checknumber%>" size="12" maxlength="6"></div>
            <div class="pcFormField">
                Check Number (must be greater than 100).
            </div> 
        </div>
        
        <div class="pcFormItem"> 
            <div class="pcFormLabel"><input size="12" name="idnbr" value="<%=pcv_driverl%>"></div>
            <div class="pcFormField">
                Driver License Number (Do not include dashes or spaces)
            </div> 
        </div>
        
        <div class="pcFormItem"> 
            <div class="pcFormLabel"><input size="12" name="idst" value="<%=pcv_driverlst%>"></div>
            <div class="pcFormField">
                Issuing State (2-digit state code)
            </div> 
        </div>
        
        <div class="pcFormItem"> 
            <div class="pcFormLabel"><input size="12" name="bankname"></div>
            <div class="pcFormField">
                Bank Name
            </div> 
        </div>
        
        <div class="pcFormItem"> 
            <div class="pcFormLabel"><input size="12" name="bankcity"></div>
            <div class="pcFormField">
                Bank City
            </div> 
        </div>
        
        <div class="pcFormItem"> 
            <div class="pcFormLabel"><input size="12" maxlength="2" name="bankst"></div>
            <div class="pcFormField">
                Bank State (2-digit state code)
            </div> 
        </div>
        
        <div class="pcFormItem"> 
            <div class="pcFormLabel"><input size="12" maxlength="5" name="bankzip"></div>
            <div class="pcFormField">
                Bank Zip Code (5-digit zip code)
            </div> 
        </div>

        <div class="pcFormItem"> 
            <p>Bank Rounting Number &amp; Account Number.<br>
            Enter <u>ALL NUMBERS</u> from the bottom of your check starting from left to right. For each symbol encountered, enter &quot;S&quot;. <u>BE SURE TO INCLUDE SPACES</u>.</p>
        </div>
        
        <div class="pcFormItem"> 
            <input size="40" name="micr" maxlength="80" value="<%=pcv_micr%>">
        </div>

        <div class="pcFormItem"> 
            Bank routing number and checking account number may vary. Below are two common examples of placement.
        </div>

        <div class="pcFormItem"> 
            <b>Example 1:</b> Enter as: <b>S987654321S&nbsp;67895432S&nbsp;00250</b><br><img src="<%=pcf_getImagePath("images","CBNexample1.gif")%>" width="302" height="26" vspace="10">
        </div>

        <div class="pcFormItem"> 
            <b>Example 2:</b> Enter as: <b>S987654321S00250&nbsp;6789S5432S</b><br><img src="<%=pcf_getImagePath("images","CBNexample2.gif")%>" width="288" height="29" vspace="10">
        </div>

            <div class="pcFormItem"> 
                <div class="pcFormLabel">Amount:</div>
                <div class="pcFormField">
                    <%=money(pcBillingTotal)%>
                </div> 
            </div>

            <div class="pcInfoMessage">
                <p><strong>Please read and approve the following authorization:</strong><br><br>
                I authorize ChecksByNet to duplicate the preceeding information into a bank draft form. I understand that I will receive by email, a check authorization notice, notifying me that a bank draft has been issued on my behalf for said purchase. I will retain my original check for my record of the transaction.
                <br><br>
                I understand that the Payee or authorized agent of Payee, will sign the bank draft as my agent for this transaction only. This authorization is valid for this transaction only. No other bank drafts will be created without my direct written or verbal authorization. All returned checks are subject to a fee of $22.50 or the maximum allowed by law plus returned bank debit fee.</p>
            </div>
        
            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
