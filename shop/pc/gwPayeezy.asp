<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
Dim pcStrPageName
pcStrPageName = "gwPayeezy.asp"
		
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
		
dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
		
' Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if
		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<%
session("idCustomer")=pcIdCustomer

'//Get the Admin Settings / Payeezy data
query="SELECT pcPEY_MerchantID,pcPEY_MToken,pcPEY_APIKey,pcPEY_APISKey,pcPEY_Mode,pcPEY_TestMode,pcPEY_JSKey,pcPEY_TAToken FROM pcPay_Payeezy;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set Admin Settings / Payeezy data
if not rs.eof then
	pcPEYMerchantID=rs("pcPEY_MerchantID")
	if pcPEYMerchantID <> "" then
		pcPEYMerchantID=enDeCrypt(pcPEYMerchantID, scCrypPass)
	end if
	pcPEYMerchantToken=rs("pcPEY_MToken")
	if pcPEYMerchantToken <> "" then
		pcPEYMerchantToken=enDeCrypt(pcPEYMerchantToken, scCrypPass)
	end if
	pcPEYAPIKey=rs("pcPEY_APIKey")
	if pcPEYAPIKey <> "" then
		pcPEYAPIKey=enDeCrypt(pcPEYAPIKey, scCrypPass)
	end if
	pcPEYAPISKey=rs("pcPEY_APISKey")
	if pcPEYAPISKey <> "" then
		pcPEYAPISKey=enDeCrypt(pcPEYAPISKey, scCrypPass)
	end if
	pcPEYJSKey=rs("pcPEY_JSKey")
	if pcPEYJSKey <> "" then
		pcPEYJSKey=enDeCrypt(pcPEYJSKey, scCrypPass)
	end if
	pcPEYMode=rs("pcPEY_Mode")
	if IsNull(pcPEYMode) OR (pcPEYMode="") then
		pcPEYMode=0
	end if
	pcPEYTestMode=rs("pcPEY_TestMode")
	if IsNull(pcPEYTestMode) OR (pcPEYTestMode="") then
		pcPEYTestMode=0
	end if
	pcPEYTAToken = rs("pcPEY_TAToken")
	if pcPEYTAToken<>"" then
		pcPEYTAToken = enDeCrypt(pcPEYTAToken, scCrypPass)
	else
		pcPEYTAToken=""
	end if
	if pcPEYTAToken="" OR pcPEYTestMode="1" then
		pcPEYTAToken = "NOIW"
	end if
end if
set rs=nothing
%>
<% If pcPEYTestMode="1" Then %>
<script src="payeezy_v3.2-cert.js" type="text/javascript"></script>
<% Else %>
<script src="payeezy_v3.2.js" type="text/javascript"></script>
<% End If %>
<script>
	Payeezy.setApiKey('<%=pcPEYAPIKey%>');
	Payeezy.setJs_Security_Key('<%=pcPEYJSKey%>');
	Payeezy.setTa_token('<%=pcPEYTAToken%>');  
</script>
<%

'*************************************************************************************
' Post_Back
' START
'*************************************************************************************
Function epoch2date(myEpoch)
	epoch2date = DateAdd("s", fix(myEpoch/1000), "01/01/1970 00:00:00")
End Function

Function date2epoch(myDate)
	date2epoch = DateDiff("s", "01/01/1970 00:00:00",myDate)*1000
End Function

Function GenNonce()
    Dim Tn1, w
    
	Tn1=""
	For w=1 to 19
		Randomize
		Tn1=Tn1 & Cstr(Fix(10*Rnd))
	Next
    
	GenNonce=Tn1
    
End Function


if request("action")="go" then
	
	'// Handle the Payeezy Token
	cardType = request.Form("card_type")
	cardholderName = request.Form("cardholder_name")
	cardExpDate = request.Form("card_expdate")
	payeezyToken=Trim(request.Form("payeezyToken"))
	if payeezyToken="" then
		call closedb()
		response.redirect "onepagecheckout.asp"
	end if
	
	if pcPEYTestMode="1" then
		pcPEYAPIURL="https://api-cert.payeezy.com/v1/transactions"
	else
		pcPEYAPIURL="https://api.payeezy.com/v1/transactions"
	end if
	
	pcPEYnonce=GenNonce()
	pcPEYtimestamp=date2epoch(UtcNow())

	pcPEYData="{" & VbLf
	pcPEYData=pcPEYData & "  ""merchant_ref"": ""OrdID" & pcGatewayDataIdOrder & """," & VbLf
	pcPEYData=pcPEYData & "  ""transaction_type"": """
	if pcPEYMode="0" then
		pcPEYData=pcPEYData & "authorize"
	else
		pcPEYData=pcPEYData & "purchase"
	end if
	pcPEYData=pcPEYData & """," & vbLf
	pcPEYData=pcPEYData & "  ""method"": ""token""," & VbLf
	pcPEYData=pcPEYData & "  ""amount"": """ & Fix(pcBillingTotal*100) & """," & VbLf
	pcPEYData=pcPEYData & "  ""currency_code"": ""USD""," & VbLf
	pcPEYData=pcPEYData & "  ""token"": {" & VbLf
	pcPEYData=pcPEYData & "    ""token_type"": ""FDToken""," & VbLf
	pcPEYData=pcPEYData & "    ""token_data"": {" & VbLf
	pcPEYData=pcPEYData & "      ""type"": """ & cardType & """," & VbLf
	pcPEYData=pcPEYData & "      ""value"": """ & payeezyToken & """," & VbLf
	pcPEYData=pcPEYData & "      ""cardholder_name"": """ & cardholderName & """," & VbLf
	pcPEYData=pcPEYData & "      ""exp_date"": """ & cardExpDate & """" & VbLf
	pcPEYData=pcPEYData & "    }" & VbLf
	pcPEYData=pcPEYData & "  }" & VbLf
	pcPEYData=pcPEYData & "}"
	
	pcPEYStr=pcPEYAPIKey & pcPEYnonce & pcPEYtimestamp & pcPEYMerchantToken & pcPEYData
	
	Set sha256 = GetObject("script:" & Server.MapPath("sha256.txt"))

	pcPEYStr=sha256.b64_hmac_sha256(pcPEYAPISKey, pcPEYStr)
	
	Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)

	xml.open "POST", pcPEYAPIURL, False
	xml.setRequestHeader "apikey", pcPEYAPIKey
	xml.setRequestHeader "token", pcPEYMerchantToken
	xml.setRequestHeader "Content-type", "application/json"
	xml.setRequestHeader "Authorization", pcPEYStr
	xml.setRequestHeader "nonce", pcPEYnonce
	xml.setRequestHeader "timestamp", pcPEYtimestamp
	
	xml.send pcPEYData
	strStatus = xml.Status

	'store the response
	strRetVal = xml.responseText

	dim Info : set Info = JSON.parse(strRetVal)

	HaveErrors=0
	ErrorMsg=""
	
	TransID=""
	TransTag=""
	
	For Each Key in Info.keys()
		if UCase(key)="ERROR" then
			HaveErrors=1
		end if
		if UCase(key)="TRANSACTION_ID" then
			TransID=Info.transaction_id
		end if
		if UCase(key)="TRANSACTION_TAG" then
			TransTag=Info.transaction_tag
		end if
	Next
	
	if (HaveErrors=0) then
		mStr=Info.transaction_status
		if UCase(trim(mStr))<>Ucase("Approved") then
			ErrorMsg="Transaction Status: " & Ucase(mStr) & "<br>"
			
			'Log failed transaction
			call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
		else
			session("GWAuthCode")=TransTag
			session("GWTransId")=TransID
			
			if pcPEYMode="1" then
				session("GWTransType")="SALE"
			else
				session("GWTransType")="AUTH"
			end if
			
			if pcPEYMode="0" then
				tmpPStatus=0
			else
				tmpPStatus=1
			end if
			
			query="INSERT INTO pcPayeezyLogs (idOrder,idCustomer,pcPEYLg_Status) VALUES (" & pcGatewayDataIdOrder & "," & pcIdCustomer & "," & tmpPStatus & ");"
			set rs=connTemp.execute(query)
			set rs=nothing
			
			'Log successful transaction
			call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 1)
			call closedb()
			
			response.redirect "gwReturn.asp?s=true&gw=Payeezy"
		end if
	else
	  	'Log failed transaction
		call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
		
		ErrorMsg=dictLanguage.Item(Session("language")&"_Payeezy_2") & "<ul>"
		For Each AField In Info.Error.messages.keys()
			ErrorMsg=ErrorMsg & "<li>" & Info.Error.messages.get(AField).description & "</li>"
		Next
		ErrorMsg=ErrorMsg & "</ul>"
	end if
	if ErrorMsg<>"" then%>
	<div class="pcErrorMessage">
		<%=dictLanguage.Item(Session("language")&"_Payeezy_1")%><br><br>
		<%=ErrorMsg%>
		<br>
		<%=dictLanguage.Item(Session("language")&"_Payeezy_3")%>
	</div>
	<%end if
end if

'*************************************************************************************
' Post_Back
' END
'*************************************************************************************

%>
<div id="pcMain">
	<div class="pcMainContent">

        <form method="POST" action="<%=pcStrPageName%>" name="payment-info-form" id="payment-info-form" class="pcForms">

            <% call pcs_showBillingAddress %>
			<div id="errorArea" style="display:none">
			<div id="errorDiv" class="pcErrorMessage">
			</div>
			</div>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></div>
				<div class="pcFormField">
					<select payeezy-data="card_type">
					<option value="visa">Visa</option>
					<option value="mastercard">Master Card</option>
					<option value="American Express" >American Express</option>
					<option value="discover" >Discover</option>
					</select>
				</div>
			</div>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></div>
				<div class="pcFormField">
					<input type="text" payeezy-data="cc_number" size="30" value="" autocomplete="off"/>
				</div>
			</div>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></div>
				<div class="pcFormField"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
					<select payeezy-data="exp_month">
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
					<select payeezy-data="exp_year">
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
				</div>
			</div>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
				<div class="pcFormField">
					<input type="text" payeezy-data="cvv_code" size="5" value=""/>
				</div>
			</div>
			
			<div class="pcFormItem"> 
			    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
                <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
            </div>
			
			<div class="pcFormItem">
				<input type="hidden" payeezy-data="cardholder_name" value="<%=pcBillingFirstName & " " & pcBillingLastName%>">
				<input type="hidden" name="merchant_ref" id="merchant_ref" value="OrdID<%=pcGatewayDataIdOrder%>"/>
				<input type="hidden" id="transaction_type" name="transaction_type" value="<%if pcPEYMode="0" then%>authorize<%else%>purchase<%end if%>">
				<input type="hidden" name="amount" id="amount" value="<%=Fix(pcBillingTotal*100)%>"/>
			</div>
					
            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
		<form method="POST" action="<%=pcStrPageName%>?action=go" name="postbackform" id="postbackform" class="pcForms">
			<input type="hidden" name="payeezyToken" id="payeezyToken" value=""/>
            <input type="hidden" name="card_type" id="card_type" value=""/>
            <input type="hidden" name="cardholder_name" id="cardholder_name" value=""/>
            <input type="hidden" name="card_expdate" id="card_expdate" value=""/>
		</form>
    </div>
</div>
<script>
jQuery(function($) {
	$('#payment-info-form').submit(function(e) {		
        var $form = $(this);
		$form.find('button').prop('disabled', true);
		$("#card_type").val($form.find('[payeezy-data="card_type"]').val());
		$("#cardholder_name").val($form.find('[payeezy-data="cardholder_name"]').val());
		$("#card_expdate").val($form.find('[payeezy-data="exp_month"]').val() + $form.find('[payeezy-data="exp_year"]').val());      
		Payeezy.createToken(responseHandler);

		return false;
	});
});
var responseHandler = function(status, response) {
	var $form = $('#payment-info-form');
	if (status != 201) {
        var allErrors = '';
        if (response.error && status != 400) {
            var error = response["error"];
            var errormsg = error["messages"];
            var errorcode = JSON.stringify(errormsg[0].code, null, 4);
            var errorMessages = JSON.stringify(errormsg[0].description, null, 4); 
			for (i=0; i<errorMessages.length;i++) {
				allErrors = allErrors + "<li>" + errorMessages[i].description + "</li>";
			}
            document.getElementById("errorArea").style.display="";
            $("#errorDiv").html("<ul>" + allErrors + "</ul>");
        }
        if (status == 400 || status == 401 || status == 500) {
            var errormsg = response.Error.messages;
            var errorMessages = "";
            for(var i in errormsg)
            {
                var ecode = errormsg[i].code;
                var eMessage = errormsg[i].description;
                allErrors = allErrors + "<li>" + eMessage + "</li>";
            }
            document.getElementById("errorArea").style.display="";
            $("#errorDiv").html("<ul>" + allErrors + "</ul>");
        }
		$form.find('button').prop('disabled', false);
        
	} else {
		var token = response.token.value;
		console.log('token: ' + token);
		$("#payeezyToken").val(token);
		$("#postbackform").submit();
	}
};
</script>
<%Function UtcNow()
UtcNow = serverdate.toUTCString()
UtcNow = CDate(Replace(Right(UtcNow, Len(UtcNow) - Instr(UtcNow, ",")), "UTC", ""))
End Function
%>
<script language="JScript" runat="server">
var serverdate=new Date();
</script>
<!--#include file="footer_wrapper.asp"-->