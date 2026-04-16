<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="simlib.asp"-->
<% 'Gateway specific files %>
<%
Dev_Testmode = 0
%>
<div id="pcMain">
	<div class="pcMainContent">
    
		<% if session("GWOrderDone")="YES" then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://")
			session("GWOrderDone")=""
			response.redirect tempURL
		end if

		session("redirectPage")="gwAuthorizeDPM.asp" %>

		<% Dim pcCustIpAddress
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
		pcv_IncreaseCustID=(scCustPre + int(pcIdCustomer)) %>

		<%
		query="SELECT x_Type,x_Login,x_Password,x_Curcode,x_DPMType,x_CVV,x_testmode,x_secureSource FROM pcPay_AuthorizeDPM Where id=1"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		x_Type=rs("x_Type")
		x_Login=rs("x_Login")
		'decrypt
		x_Login=enDeCrypt(x_Login, scCrypPass)
		x_Password=rs("x_Password")
		'decrypt
		x_Password=enDeCrypt(x_Password, scCrypPass)
		x_Curcode=rs("x_Curcode")
		x_DPMType=rs("x_DPMType")
		x_CVV=rs("x_CVV")
		x_testmode=rs("x_testmode")
		x_secureSource=rs("x_secureSource")
		x_TypeArray=Split(x_Type,"||")
		x_TransType=x_TypeArray(0)
		set rs=nothing
		
		if Dev_Testmode = 1 Then
			postURL = "https://test.authorize.net/gateway/transact.dll"
		else
			postURL = "https://secure2.authorize.net/gateway/transact.dll"
		end if
		
		If scSSL="" OR scSSL="0" Then
			x_relayURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwAuthorizeDPM_relay.asp"),"//","/")
			x_relayURL=replace(x_relayURL,"https:/","https://")
			x_relayURL=replace(x_relayURL,"http:/","http://") 
		Else
			x_relayURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwAuthorizeDPM_relay.asp"),"//","/")
			x_relayURL=replace(x_relayURL,"https:/","https://")
			x_relayURL=replace(x_relayURL,"http:/","http://")
		End If

		query="SELECT x_Type,x_CVV FROM authorizeNet Where id=1;"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		orderDescription = replace(scCompanyName,",","-") & " Order: " & session("GWOrderID")
		sequence = Int(1000 * Rnd)
		timeStamp = simTimeStamp()
		fingerprint = HMAC (x_Password, x_login & "^" & sequence & "^" & timeStamp & "^" & pcBillingTotal & "^")
		
		%>
	
       	<form action="<%= postURL %>" method="POST" name="form1" class="pcForms">
        	<input type="hidden" name="x_login" value="<%= x_Login %>" />
            <input type="hidden" name="x_amount" value="<%= pcBillingTotal %>" />
            <input type="hidden" name="x_description" value="<%= orderDescription %>" />
            <input type="hidden" name="x_invoice_num" value="<%= session("GWOrderID") %>" />
            <input type="hidden" name="x_cust_id" value="<%= session("idcustomer") %>" />
            <input type="hidden" name="x_first_name" value="<%= pcBillingFirstName %>" />
            <input type="hidden" name="x_last_name" value="<%= pcBillingLastName %>" />
            <input type="hidden" name="x_company" value="<%= replace(pcBillingCompany,",","||") %>" />
            <input type="hidden" name="x_address" value="<%= replace(pcBillingAddress,",","||") %>" />
            <input type="hidden" name="x_city" value="<%= pcBillingCity %>" />
            <input type="hidden" name="x_state" value="<%= pcBillingState %>" />
            <input type="hidden" name="x_zip" value="<%= pcBillingPostalCode %>" />
            <input type="hidden" name="x_country" value="<%= pcBillingCountryCode %>" />
            <input type="hidden" name="x_phone" value="<%= pcBillingPhone %>" />
            <input type="hidden" name="x_email" value="<%= pcCustomerEmail %>" />
            <input type="hidden" name="x_fp_sequence" value="<%= sequence %>" />
            <input type="hidden" name="x_fp_timestamp" value="<%= timeStamp %>" />
            <input type="hidden" name="x_fp_hash" value="<%= fingerprint %>" />
            <input type="hidden" name="x_exp_date" id="x_exp_date" value="" />
            <input type="hidden" name="x_type" value="<%= x_TransType %>" />
            <input type="hidden" name="x_relay_always" value="true" />
            <input type="hidden" name="x_relay_response" value="true" />
            <input type="hidden" name="x_relay_url" value="<%= x_relayURL %>" />

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
            
					
            <% call pcs_showBillingAddress %>
            

            <div class="pcFormItem"> 
                <div class="pcFormLabel">Card Type:</div>
                <div class="pcFormField">
                    <select name="x_Card_Type">
                    <% 	x_TypeArray=Split(x_Type,"||")
                    If ubound(x_TypeArray)=1 Then
                        x_Type2=x_TypeArray(1)
                        cardTypeArray=split(x_Type2,", ")
                        i=ubound(cardTypeArray)
                        cardCnt=0
                        do until cardCnt=i+1
                            cardVar=cardTypeArray(cardCnt)
                            select case cardVar
                                case "V"
                                    response.write "<option value=""V"" selected>Visa</option>"
                                    cardCnt=cardCnt+1
                                case "M" 
                                    response.write "<option value=""M"">MasterCard</option>"
                                    cardCnt=cardCnt+1
                                case "A"
                                    response.write "<option value=""A"">American Express</option>"
                                    cardCnt=cardCnt+1
                                case "D"
                                    response.write "<option value=""D"">Discover</option>"
                                    cardCnt=cardCnt+1
                            end select
                        loop
                    End If %>
                    </select>
                </div>
            </div>
                    
                    
            <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></div>
                <div class="pcFormField"><input type="text" name="x_card_num" value="" autocomplete="off"></div>
            </div>
                    
                    
            <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></div>
                <div class="pcFormField"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
                    <select name="expMonth" id="expMonth">
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
                    <select name="expYear" id="expYear">
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
            
            <script type="text/javascript">
				$('#expMonth').on('change', function() {
				  $('#x_exp_date').val($('#expMonth option:selected').val() + $('#expYear option:selected').val());
				  console.log($('#x_exp_date').val());
				});
				$('#expYear').on('change', function() {
				  $('#x_exp_date').val($('#expMonth option:selected').val() + $('#expYear option:selected').val());
				  console.log($('#x_exp_date').val());
				});
				$(document).ready(function() {
					$('#x_exp_date').val($('#expMonth option:selected').val() + $('#expYear option:selected').val());
					console.log($('#x_exp_date').val());
				});
			</script>
                    
            <% If x_CVV="1" Then %>
            
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></div>
                </div>

            <% End If %>

            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>

    </div>
</div>
<!--#include file="footer_wrapper.asp"-->