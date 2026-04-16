<%
'---Start Payeezy gateway---
Function gwPayEzyEdit()
		'request gateway variables and insert them into the pcPay_Payeezy table
	query= "SELECT pcPEY_MToken,pcPEY_APIKey,pcPEY_APISKey,pcPEY_Mode,pcPEY_TestMode,pcPEY_JSKey,pcPEY_TAToken FROM pcPay_Payeezy;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	'x_S2=rs("pcPEY_MerchantID")
	'x_S2=enDeCrypt(x_S2, scCrypPass)
	x_M2=rs("pcPEY_MToken")
	if x_M2 <> "" then
		x_M2=enDeCrypt(x_M2, scCrypPass)
	end if
	x_Login2=rs("pcPEY_APIKey")
	if x_Login2 <> "" then
		x_Login2=enDeCrypt(x_Login2, scCrypPass)
	end if
	x_Key2=rs("pcPEY_APISKey")
	if x_Key2 <> "" then
		x_Key2=enDeCrypt(x_Key2, scCrypPass)
	end if
	x_JSKey2=rs("pcPEY_JSKey")
	if x_JSKey2 <> "" then
		x_JSKey2=enDeCrypt(x_JSKey2, scCrypPass)
	end if
	x_TAToken2=rs("pcPEY_TAToken")
	if x_TAToken2 <> "" then
		x_TAToken2=enDeCrypt(x_TAToken2, scCrypPass)
	end if
	set rs=nothing
	
	'x_S=request.Form("x_S")
	'if x_S="" then
	'	x_S=x_S2
	'end if
	'x_S=enDeCrypt(x_S, scCrypPass)
	
	x_M=request.Form("x_M")
	if x_M="" then
		x_M=x_M2
	end if
	if x_M <> "" then
		x_M=enDeCrypt(x_M, scCrypPass)
	end if
	
	x_Login=request.Form("x_Login")
	if x_Login="" then
		x_Login=x_Login2
	end if
	if x_Login <> "" then
		x_Login=enDeCrypt(x_Login, scCrypPass)
	end if
	
	x_Key=request.Form("x_Key")
	if x_Key="" then
		x_Key=x_Key2
	end if
	if x_Key <> "" then
		x_Key=enDeCrypt(x_Key, scCrypPass)
	end if
	
	x_JSKey=request.Form("x_JSKey")
	if x_JSKey="" then
		x_JSKey=x_JSKey2
	end if
	if x_JSKey <> "" then
		x_JSKey=enDeCrypt(x_JSKey, scCrypPass)
	end if
	
	x_TAToken=request.Form("x_TAToken")
	if x_TAToken="" then
		x_TAToken=x_TAToken2
	end if
	if x_TAToken <> "" then
		x_TAToken=enDeCrypt(x_TAToken, scCrypPass)
	end if
	
	x_URLMethod="gwPayeezy.asp"
	x_testmode=request.Form("x_testmode")
	if x_testmode="" then
		x_testmode="0"
	end if
	x_Mode=request.Form("x_Mode")
	if x_Mode="" then
		x_Mode=0
	end if
	
	query="UPDATE pcPay_Payeezy SET pcPEY_MToken='"&x_M&"',pcPEY_APIKey='"&x_Login&"',pcPEY_APISKey='" & x_Key & "',pcPEY_Mode="&x_Mode&",pcPEY_TestMode="&x_TestMode&",pcPEY_JSKey='"&x_JSKey&"',pcPEY_TAToken='"&x_TAToken&"';"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&x_URLMethod&"', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=1101"
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

Function gwPayEzy()
	varCheck=1
	'request gateway variables and insert them into the pcPay_Payeezy table
	x_URLMethod="gwPayeezy.asp"
	
	x_testmode=request.Form("x_testmode")
	if x_testmode="" then
		x_testmode=0
	end if
	
	x_Mode=request.Form("x_Mode")
	if x_Mode="" then
		x_Mode=0
	end if
	
	'x_S=request.Form("x_S")
	'If x_S="" then
	'	call closeDb()
	'	response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Payeezy as your payment gateway. <b>""Merchant Identifier""</b> is a required field.")
	'End If
	'x_S=enDeCrypt(x_S, scCrypPass)
	
	x_M=request.Form("x_M")
	If x_M="" then
		call closeDb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Payeezy as your payment gateway. <b>""Merchant Token""</b> is a required field.")
	End If
	x_M=enDeCrypt(x_M, scCrypPass)
	
	x_Login=request.Form("x_Login")
	If x_Login="" then
		call closeDb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Payeezy as your payment gateway. <b>""API Key""</b> is a required field.")
	End If
	'encrypt
	x_Login=enDeCrypt(x_Login, scCrypPass)
	
	x_Key=request.Form("x_Key")
	If x_Key="" then
		call closeDb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Payeezy as your payment gateway. <b>""API Secret Key""</b> is a required field.")
	End If
	'encrypt
	x_Key=enDeCrypt(x_Key, scCrypPass)
	
	x_JSKey=request.Form("x_JSKey")
	If x_JSKey="" then
		call closeDb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Payeezy as your payment gateway. <b>""API JS Security Key""</b> is a required field.")
	End If
	'encrypt
	x_JSKey=enDeCrypt(x_JSKey, scCrypPass)

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
		paymentNickName="Pay With Payeezy"
	end if
	
	err.clear
	err.number=0
	
	 
	query="SELECT pcPEY_ID FROM pcPay_Payeezy;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		query="UPDATE pcPay_Payeezy SET pcPEY_MToken='"&x_M&"',pcPEY_APIKey='"&x_Login&"',pcPEY_APISKey='" & x_Key & "',pcPEY_Mode="&x_Mode&",pcPEY_TestMode="&x_TestMode&",pcPEY_JSKey='"&x_JSKey&"',pcPEY_TAToken='"&x_TAToken&"';"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
	else
		query="INSERT INTO pcPay_Payeezy (pcPEY_MToken,pcPEY_APIKey,pcPEY_APISKey,pcPEY_Mode,pcPEY_TestMode,pcPEY_JSKey,pcPEY_TAToken) VALUES ('" & x_M & "','" & x_Login & "','" & x_Key & "'," & x_Mode & "," & x_TestMode & ", '" & x_JSKey & "','" & x_TAToken & "');"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		
	end if
	set rs=nothing

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Payeezy','"&x_URLMethod&"',-1,0,0,9999,0,9999,0,9999,0,"& priceToAdd &","& percentageToAdd &",1101,N'"&paymentNickName&"')"
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
				
<% if request("gwchoice")="1101" then
	if request("mode")="Edit" then
		query= "SELECT pcPEY_MToken,pcPEY_APIKey,pcPEY_APISKey,pcPEY_Mode,pcPEY_TestMode,pcPEY_JSKey,pcPEY_TAToken FROM pcPay_Payeezy;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if not rs.eof then
			'x_S=rs("pcPEY_MerchantID")
			'x_S=enDeCrypt(x_S, scCrypPass)
			x_M=rs("pcPEY_MToken")
			x_M=enDeCrypt(x_M, scCrypPass)
			x_Login=rs("pcPEY_APIKey")
			x_Login=enDeCrypt(x_Login, scCrypPass)
			x_Key=rs("pcPEY_APISKey")
			x_Key=enDeCrypt(x_Key, scCrypPass)
			x_JSKey=rs("pcPEY_JSKey")
			x_JSKey=enDeCrypt(x_JSKey, scCrypPass)
			x_TAToken=rs("pcPEY_TAToken")
			x_TAToken=enDeCrypt(x_TAToken, scCrypPass)
			x_Mode=rs("pcPEY_Mode")
			x_testmode=rs("pcPEY_TestMode")
		end if
		set rs=nothing

		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=1101"
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
		if IsNull(paymentNickName) OR paymentNickName="" OR paymentNickName="Credit Card" then
			paymentNickName="Pay With Payeezy"
		end if
	end if %>
	<input type="hidden" name="addGw" value="1101">
	<table width="100%">
	<tr>
		<td align="left" style="font-size:15px;"><img src="Gateways/logos/payeezy.png"></td>
		<td align="left" style="font-size:15px;">&nbsp;</td>
	</tr>
	</table>
	<br>
    
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>Pay with Payeezy</h4>    
                    <p>
                        Smart payments in your own language. Build your business with Payeezy, the simplest way to accept transactions online.
                    </p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.firstdatapartners.com/ecommerce/?partner=product-cart" target="_blank">Learn More</a>                      
                    </p>
                </div>
                
			</td>
        </tr>
        <tr>
            <td>
            
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
                                        <% dim PEY_KeyCnt,PEY_KeyEnd,PEY_KeyStart
                                        PEY_KeyCnt=(len(x_Key)-2)
                                        PEY_KeyEnd=right(x_Key,2)
                                        PEY_KeyStart=""
                                        for c=1 to PEY_KeyCnt
                                            PEY_KeyStart=PEY_KeyStart&"*"
                                        next %>
                                        <tr> 
                                            <td colspan="2">Current Payeezy API Secret Key:&nbsp;<%=PEY_KeyStart&PEY_KeyEnd%></td>
                                        </tr>
                                        <tr> 
                                            <td colspan="2"> For security reasons, your &quot;Payeezy API Secret Key&quot; is only partially shown on this page. If you need to edit your account information, please re-enter your &quot;Payeezy API Secret Key&quot; below.</td>
                                        </tr>
                                    <% end if %>
                                    <!--
                                    <tr> 
                                        <td> <div align="right">Merchant Identifier:</div></td>
                                        <td width="479"> <input type="text" name="x_S" size="30"></td>
                                    </tr>
                                    -->
									<tr> 
                                        <td> <div align="right">Merchant Token:</div></td>
                                        <td width="479"> <input type="text" name="x_M" size="30"></td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">API Key:</div></td>
                                        <td width="479"> <input type="text" name="x_Login" size="30"></td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">API Secret Key:</div></td>
                                        <td> <input name="x_Key" type="text" size="30"> </td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">JS Security Key:</div></td>
                                        <td> <input name="x_JSKey" type="text" size="30"> </td>
                                    </tr>
                                    <tr> 
                                        <td> <div align="right">TA Token:</div></td>
                                        <td> <input name="x_TAToken" type="text" size="30"> </td>
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
                                                <input name="x_testmode" type="radio" class="clearBorder" value="1" <% if x_testmode="1" then%>checked<% end if%>>
                                            </div></td>
                                        <td>Sandbox Mode</td>
                                    </tr>
                                    <tr> 
                                        <td><div align="right"> 
                                                <input name="x_testmode" type="radio" class="clearBorder" value="0" <% if x_testmode<>"1" then%>checked<% end if%>>
                                            </div></td>
                                        <td>Live Mode</td>
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
                        <input type="submit" value="<%=strButtonValue%>" name="Submit" class="btn btn-primary"> 
                        &nbsp;
                        <input type="button" class="btn btn-default"  value="Back" onclick="javascript:history.back()">
                        </td>
                    </tr>
				</table>

            </td>
        </tr>
    </table>
<% end if %>