<%
'---Start Amazon MWS gateway---
Function gwAMZPEdit()
		'request gateway variables and insert them into the gwAmazon table
	query= "SELECT gwAMZ_SellerID,gwAMZ_AccessKey,gwAMZ_SecretKey,gwAMZ_ClientID,gwAMZ_ClientSecret FROM gwAmazon;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	x_S2=rs("gwAMZ_SellerID")
	x_S2=enDeCrypt(x_S2, scCrypPass)
	x_Login2=rs("gwAMZ_AccessKey")
	x_Login2=enDeCrypt(x_Login2, scCrypPass)
	x_Key2=rs("gwAMZ_SecretKey")
	x_Key2=enDeCrypt(x_Key2, scCrypPass)
	x_C2=rs("gwAMZ_ClientID")
	x_C2=enDeCrypt(x_C2, scCrypPass)
	x_CS2=rs("gwAMZ_ClientSecret")
	x_CS2=enDeCrypt(x_CS2, scCrypPass)
	set rs=nothing
	
	x_S=request.Form("x_S")
	if x_S="" then
		x_S=x_S2
	end if
	x_S=enDeCrypt(x_S, scCrypPass)
	
	
	x_Login=request.Form("x_Login")
	if x_Login="" then
		x_Login=x_Login2
	end if
	x_Login=enDeCrypt(x_Login, scCrypPass)
	
	x_Key=request.Form("x_Key")
	if x_Key="" then
		x_Key=x_Key2
	end if
	x_Key=enDeCrypt(x_Key, scCrypPass)
	
	x_C=request.Form("x_C")
	if x_C="" then
		x_C=x_C2
	end if
	'encrypt
	x_C=enDeCrypt(x_C, scCrypPass)
	
	x_CS=request.Form("x_CS")
	if x_CS="" then
		x_CS=x_CS2
	end if
	'encrypt
	x_CS=enDeCrypt(x_CS, scCrypPass)
	
	x_URLMethod="gwAmazonMWS.asp"
	x_testmode=request.Form("x_testmode")
	if x_testmode="" then
		x_testmode="0"
	end if
	x_Mode=request.Form("x_Mode")
	if x_Mode="" then
		x_Mode=0
	end if
	
	query="UPDATE gwAmazon SET gwAMZ_SellerID='"&x_S&"',gwAMZ_AccessKey='"&x_Login&"',gwAMZ_SecretKey='" & x_Key & "',gwAMZ_ClientID='"&x_C&"',gwAMZ_ClientSecret='"&x_CS&"',gwAMZ_Mode="&x_Mode&",gwAMZ_TestMode="&x_TestMode&";"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&x_URLMethod&"', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=88"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwAMZP()
	varCheck=1
	'request gateway variables and insert them into the gwAmazon table
	x_URLMethod="gwAmazonMWS.asp"
	x_testmode=request.Form("x_testmode")
	if x_testmode="" then
		x_testmode=0
	end if
	x_Mode=request.Form("x_Mode")
	if x_Mode="" then
		x_Mode=0
	end if
	x_S=request.Form("x_S")
	If x_S="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Amazon MWS as your payment gateway. <b>""Seller ID""</b> is a required field.")
	End If
	x_S=enDeCrypt(x_S, scCrypPass)
	
	x_Login=request.Form("x_Login")
	If x_Login="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Amazon MWS as your payment gateway. <b>""Access Key""</b> is a required field.")
	End If
	'encrypt
	x_Login=enDeCrypt(x_Login, scCrypPass)
	x_Key=request.Form("x_Key")
	If x_Key="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Amazon MWS as your payment gateway. <b>""Secret Key""</b> is a required field.")
	End If
	'encrypt
	x_Key=enDeCrypt(x_Key, scCrypPass)
	x_C=request.Form("x_C")
	If x_C="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Amazon MWS as your payment gateway. <b>""Client ID""</b> is a required field.")
	End If
	x_C=enDeCrypt(x_C, scCrypPass)
	
	x_CS=request.Form("x_CS")
	If x_CS="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Amazon MWS as your payment gateway. <b>""Client Secret Key""</b> is a required field.")
	End If
	x_CS=enDeCrypt(x_CS, scCrypPass)
	priceToAddType=request.Form("priceToAddType")
	if priceToAddType="price" then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
		If priceToAdd="" Then
			priceToAdd="0"
		end if
	else
		priceToAdd="0"
		percentageToAdd=request.Form("percentageToAdd")
		If percentageToAdd="" Then
			percentageToAdd="0"
		end if
	end if
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")
	if paymentNickName="" then
		paymentNickName="Pay With Amazon"
	end if
	
	err.clear
	err.number=0
	
	 
	query="SELECT gwAMZ_ID FROM gwAmazon;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		query="UPDATE gwAmazon SET gwAMZ_SellerID='"&x_S&"',gwAMZ_AccessKey='"&x_Login&"',gwAMZ_SecretKey='" & x_Key & "',gwAMZ_ClientID='"&x_C&"',gwAMZ_ClientSecret='"&x_CS&"',gwAMZ_Mode="&x_Mode&",gwAMZ_TestMode="&x_TestMode&";"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
	else
		query="INSERT INTO gwAmazon (gwAMZ_SellerID,gwAMZ_AccessKey,gwAMZ_SecretKey,gwAMZ_ClientID,gwAMZ_ClientSecret,gwAMZ_Mode,gwAMZ_TestMode) VALUES ('" & x_S & "','" & x_Login & "','" & x_Key & "','" & x_C & "','" & x_CS & "'," & x_Mode & "," & x_TestMode & ");"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		
	end if
	set rs=nothing

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Amazon','"&x_URLMethod&"',-1,0,0,9999,0,9999,0,9999,0,"& priceToAdd &","& percentageToAdd &",88,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	set rs=nothing
	
	
end function
%>
				
<% if request("gwchoice")="88" then
	if request("mode")="Edit" then
				query= "SELECT gwAMZ_SellerID,gwAMZ_AccessKey,gwAMZ_SecretKey,gwAMZ_ClientID,gwAMZ_ClientSecret,gwAMZ_Mode,gwAMZ_TestMode FROM gwAmazon;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if not rs.eof then
			x_S=rs("gwAMZ_SellerID")
			x_S=enDeCrypt(x_S, scCrypPass)
			x_Login=rs("gwAMZ_AccessKey")
			x_Login=enDeCrypt(x_Login, scCrypPass)
			x_Key=rs("gwAMZ_SecretKey")
			x_Key=enDeCrypt(x_Key, scCrypPass)
			x_C=rs("gwAMZ_ClientID")
			x_C=enDeCrypt(x_C, scCrypPass)
			x_CS=rs("gwAMZ_ClientSecret")
			x_CS=enDeCrypt(x_CS, scCrypPass)
			x_Mode=rs("gwAMZ_Mode")
			x_testmode=rs("gwAMZ_TestMode")
		end if
		set rs=nothing

		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=88"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if not rs.eof then
			paymentNickName=rs("paymentNickName")
		end if

		if x_Curcode="" then
			x_Curcode="USD"
		end if
		%>
		<input type="hidden" name="mode" value="Edit">
    <%else
		x_Mode=0
		x_testmode=0
		if IsNull(paymentNickName) OR paymentNickName="" then
			paymentNickName="Pay With Amazon"
		end if
	end if %>
	<input type="hidden" name="addGw" value="88">
	<table width="100%">
	<tr>
		<td align="left" style="font-size:15px;"><img src="Gateways/logos/amazonpayments.png"></td>
		<td align="left" style="font-size:15px;">&nbsp;</td>
	</tr>
	</table>
	<br>
    
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>Pay with Amazon (Amazon MWS)</h4>    
                    <p>
                        <strong>Login and Pay with Amazon</strong>
                    </p>
                    <p>
                        Millions of Amazon buyers can login and pay on your website or mobile site with the information already stored in their Amazon account. It's fast, easy and secure. It can help you add new customers, increase sales, and turn browsers into buyers. Leverage the trust of Amazon to grow your business. It's easy.
                    </p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://payments.amazon.com" target="_blank">Learn More</a>                      
                    </p>
                </div>
                
			</td>
        </tr>
        <tr>
            <td>
            
                <% If scSSLUrl="" Then %>
                    <div class="bs-callout bs-callout-warning">
                        <h4>Important Notice</h4>    
                        <p>
                            Your store must have SSL turned on in order to activate Pay with Amazon (Amazon MWS). <a href="AdminSettings.asp">Configure the SSL settings now</a>.                    
                        </p>
                    </div>
                <% End If %>
                
                <div id="accordion" class="panel-group">
                    
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapseOne">
                                    Step 1: Configure Account
                                </a>
                            </h4>
                        </div>
                        <div id="collapseOne" class="panel-collapse collapse in">
                            <div class="panel-body">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <% if request("mode")="Edit" then %>
                                        <% dim AMZ_KeyCnt,AMZ_KeyEnd,AMZ_KeyStart
                                        AMZ_KeyCnt=(len(x_Key)-2)
                                        AMZ_KeyEnd=right(x_Key,2)
                                        AMZ_KeyStart=""
                                        for c=1 to AMZ_KeyCnt
                                            AMZ_KeyStart=AMZ_KeyStart&"*"
                                        next %>
                                        <tr> 
                                            <td colspan="2">Current Amazon Secret Key:&nbsp;<%=AMZ_KeyStart&AMZ_KeyEnd%></td>
                                        </tr>
                                        <tr> 
                                            <td colspan="2"> For security reasons, your &quot;Amazon Secret Key&quot; is only partially shown on this page. If you need to edit your account information, please re-enter your &quot;Amazon Secret Key&quot; below.</td>
                                        </tr>
                                    <% end if %>
                                    <tr> 
                                        <td> <div align="right">Seller ID (Merchant ID):</div></td>
                                        <td width="479"> <input type="text" name="x_S" size="30"></td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">Access Key:</div></td>
                                        <td width="479"> <input type="text" name="x_Login" size="30"></td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">Secret Key:</div></td>
                                        <td> <input name="x_Key" type="text" size="30"> </td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">Client ID:</div></td>
                                        <td width="479"> <input type="text" name="x_C" size="30"></td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">Client Secret Key:</div></td>
                                        <td> <input name="x_CS" type="text" size="30"> </td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">Transaction Type:</div></td>
                                        <td> <select name="x_Mode">
                                                <option value="0" <% if x_Mode="0" then %>selected<% end if %>>Authorize Only</option>
                                                <option value="1" <% if x_Mode="1" then %>selected<% end if %>>Sale</option>
                                            </select> </td>
                                    </tr>
                                    <tr> 
                                        <td><div align="right"> 
                                                <input name="x_testmode" type="checkbox" class="clearBorder" value="1" <% if x_testmode=1 then%>checked<% end if%>> 
                                            </div></td>
                                        <td><b>Enable Test Mode </b></td>
                                    </tr>
                                </table>
                            </div>
                        </div> 
                    </div>
                    
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapse2">
                                    Step 2: You have the option to charge a processing fee for this payment option.
                                </a>
                            </h4>
                        </div>
                        <div id="collapse2" class="panel-collapse collapse">
                            <div class="panel-body">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td nowrap="nowrap">&nbsp;</td>
                                        <td class="pcSubmenuContent">&nbsp;</td>
                                      </tr>
                                  <tr>
                                        <td width="7%" nowrap="nowrap"><div align="left">Processing Fee:&nbsp;</div></td>
                                        <td>
                                      <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>"></td>
                                    </tr>
                                  <tr>
                                    <td>&nbsp;</td>
                                        <td><input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>Percentage of Order Total&nbsp;&nbsp; &nbsp; %<input name="percentageToAdd" size="6" value="<%=percentageToAdd%>"></td>
                                  </tr>
                                  <tr>
                                    <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                  </tr>
                                </table>
                            </div>
                        </div> 
                    </div>
                    
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapse3">
                                    Step 3: You can change the display name that is shown for this payment type.
                                </a>
                            </h4>
                        </div>
                        <div id="collapse3" class="panel-collapse collapse">
                            <div class="panel-body">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td nowrap="nowrap">&nbsp;</td>
                                      <td>&nbsp;</td>
                                    </tr>
                                    <tr> 
                                        <td width="10%" nowrap="nowrap"><div align="left">Payment Name:&nbsp;</div></td>
                                                <td width="90%"><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                                    </tr>
                                    <tr>
                                      <td>&nbsp;</td>
                                      <td>&nbsp;</td>
                                    </tr>
                                </table>
                            </div>
                        </div> 
                    </div>
                    
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapse4">
                                    Step 4: Order Processing: Order Status and Payment Status
                                </a>
                            </h4>
                        </div>
                        <div id="collapse4" class="panel-collapse collapse">
                            <div class="panel-body">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td>&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=301"></a></td>
                                    </tr>
                                    <tr> 
                                        <td>When orders are placed, set the payment status to:
                                        <select name="pcv_setPayStatus">
                                            <option value="3" selected="selected">Default</option>
                                            <option value="0" <%if pcv_setPayStatus="0" then%>selected<%end if%>>Pending</option>
                                            <option value="1" <%if pcv_setPayStatus="1" then%>selected<%end if%>>Authorized</option>
                                            <option value="2" <%if pcv_setPayStatus="2" then%>selected<%end if%>>Paid</option>
                                        </select>
                                        &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=302"></a>					</td>
                                    </tr>
                                    <tr>
                                      <td>&nbsp;</td>
                                    </tr>
                                </table>
                            </div>
                        </div> 
                    </div>
                    
                </div>
                
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td colspan="2">
                        <% if request("mode")="Edit" then
                            strButtonValue="Save Changes" %>
                            <input type="hidden" name="submitMode" value="Edit">
                        <%  else
                            strButtonValue="Add New Payment Method" %>
                            <input type="hidden" name="submitMode" value="Add Gateway">
                        <% end if %>
                        <input type="submit" value="<%=strButtonValue%>" name="Submit" class="btn btn-primary" <% If scSSLUrl="" Then %>disabled<% End If %>> 
                        &nbsp;
                        <input type="button" class="btn btn-default"  value="Back" onclick="javascript:history.back()">
                        </td>
                    </tr>
				</table>

            </td>
        </tr>
    </table>
<% end if %>