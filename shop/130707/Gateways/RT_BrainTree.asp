<%
'---Start BrainTree gateway---
Function gwBrainTEdit()
	call opendb()
	'request gateway variables and insert them into the BrainTree table
	query= "SELECT gwBT_MerchantID,gwBT_PublicKey,gwBT_PrivateKey FROM gwBrainTree;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	x_Login2=rs("gwBT_MerchantID")
	'decrypt
	x_Login2=enDeCrypt(x_Login2, scCrypPass)
	x_Password2=rs("gwBT_PublicKey")
	'decrypt
	x_Password2=enDeCrypt(x_Password2, scCrypPass)
	x_Key2=rs("gwBT_PrivateKey")
	'decrypt
	x_Key2=enDeCrypt(x_Key2, scCrypPass)
	set rs=nothing
	x_Login=request.Form("x_Login")
	if x_Login="" then
		x_Login=x_Login2
	end if
	'encrypt
	x_Login=enDeCrypt(x_Login, scCrypPass)
	x_Password=request.Form("x_Password")
	if x_Password="" then
		x_Password=x_Password2
	end if
	'encrypt
	x_Password=enDeCrypt(x_Password, scCrypPass)
	x_Key=request.Form("x_Key")
	if x_Key="" then
		x_Key=x_Key2
	end if
	'encrypt
	x_Key=enDeCrypt(x_Key, scCrypPass)
	x_Curcode=request.Form("x_Curcode")
	if x_Curcode="" then
		x_Curcode="USD"
	end if
	x_URLMethod="gwBrainTree.asp"
	x_CVV=request.Form("x_CVV")
	if x_CVV="" then
		x_CVV=0
	end if
	x_testmode=request.Form("x_testmode")
	if x_testmode="" then
		x_testmode="0"
	end if
	x_Mode=request.Form("x_Mode")
	if x_Mode="" then
		x_Mode=0
	end if
	
	query="UPDATE gwBrainTree SET gwBT_MerchantID='"&x_Login&"',gwBT_PublicKey='"&x_Password&"',gwBT_PrivateKey='" & x_Key & "',gwBT_Curcode='"&x_Curcode&"',gwBT_CVV=" & x_CVV & ",gwBT_Mode="&x_Mode&",gwBT_TestMode="&x_TestMode&";"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&x_URLMethod&"', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=1113"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	call closedb()
end function

Function gwBrainT()
	varCheck=1
	'request gateway variables and insert them into the gwBrainTree table
	x_Curcode=request.Form("x_Curcode")
	if x_Curcode="" then
		x_Curcode="USD"
	end if
	x_URLMethod="gwBrainTree.asp"
	x_CVV=request.Form("x_CVV")
	if x_CVV="" then
		x_CVV=0
	end if
	x_testmode=request.Form("x_testmode")
	if x_testmode="" then
		x_testmode=0
	end if
	x_Mode=request.Form("x_Mode")
	if x_Mode="" then
		x_Mode=0
	end if
	x_Login=request.Form("x_Login")
	If x_Login="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add BrainTree as your payment gateway. <b>""Username""</b> is a required field.")
	End If
	'encrypt
	x_Login=enDeCrypt(x_Login, scCrypPass)
	x_Password=request.Form("x_Password")
	If x_Password="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add BrainTree as your payment gateway. <b>""Password""</b> is a required field.")
	End If
	'encrypt
	x_Password=enDeCrypt(x_Password, scCrypPass)
	x_Key=request.Form("x_Key")
	If x_Key="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add BrainTree as your payment gateway. <b>""Key""</b> is a required field.")
	End If
	'encrypt
	x_Key=enDeCrypt(x_Key, scCrypPass)
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
		paymentNickName="Credit Card"
	end if
	
	err.clear
	err.number=0
	
	call openDb() 
	query="SELECT gwBT_id FROM gwBrainTree;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		query="UPDATE gwBrainTree SET gwBT_MerchantID='"&x_Login&"',gwBT_PublicKey='"&x_Password&"',gwBT_PrivateKey='" & x_Key & "',gwBT_Curcode='"&x_Curcode&"',gwBT_CVV=" & x_CVV & ",gwBT_Mode="&x_Mode&",gwBT_TestMode="&x_TestMode&";"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
	else
		query="INSERT INTO gwBrainTree (gwBT_MerchantID,gwBT_PublicKey,gwBT_PrivateKey,gwBT_CurCode,gwBT_CVV,gwBT_Mode,gwBT_TestMode) VALUES ('" & x_Login & "','" & x_Password & "','" & x_Key & "','" & x_Curcode & "'," & x_CVV & "," & x_Mode & "," & x_TestMode & ");"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		
	end if
	set rs=nothing

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'BrainTree','"&x_URLMethod&"',-1,0,0,9999,0,9999,0,9999,0,"& priceToAdd &","& percentageToAdd &",1113,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	set rs=nothing
	
	call closedb()
end function
%>
				
<% if request("gwchoice")="1113" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT gwBT_MerchantID,gwBT_PublicKey,gwBT_PrivateKey,gwBT_CVV,gwBT_Mode,gwBT_CurCode,gwBT_TestMode FROM gwBrainTree;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if not rs.eof then
			x_Login=rs("gwBT_MerchantID")
			'decrypt
			x_Login=enDeCrypt(x_Login, scCrypPass)
			x_Password=rs("gwBT_PublicKey")
			'decrypt
			x_Password=enDeCrypt(x_Password, scCrypPass)
			x_Key=rs("gwBT_PrivateKey")
			'decrypt
			x_Key=enDeCrypt(x_Key, scCrypPass)
			x_Curcode=rs("gwBT_Curcode")
			x_CVV=rs("gwBT_CVV")
			x_Mode=rs("gwBT_Mode")
			x_testmode=rs("gwBT_TestMode")
		end if
		set rs=nothing

		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=1113"
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
		x_Curcode="USD"
		x_CVV=1
		x_Mode=0
		x_testmode=0
	end if %>
	<input type="hidden" name="addGw" value="1113">
	<table width="100%">
	<tr>
		<td align="left" style="font-size:15px;"><img src="images/pcv4_icon_pg.png" width="48" height="48"></td>
		<td align="left" style="font-size:15px;">&nbsp;</td>
	</tr>
	</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>BrainTree</h4>    
                    <p></p>
                    <p>
                        <a class="btn btn-info btn-xs" href="https://www.braintreepayments.com" target="_blank">Learn More</a>        
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
								<% dim BT_LoginCnt,BT_LoginEnd,BT_LoginStart
								BT_LoginCnt=(len(x_Login)-2)
								BT_LoginEnd=right(x_Login,2)
								BT_LoginStart=""
								for c=1 to BT_LoginCnt
									BT_LoginStart=BT_LoginStart&"*"
								next %>
								<tr> 
									<td colspan="2">Current Username:&nbsp;<%=BT_LoginStart&BT_LoginEnd%></td>
								</tr>
								<tr> 
									<td colspan="2"> For security reasons, your &quot;MerchantID&quot; is only partially shown on this page. If you need to edit your account information, please re-enter your &quot;MerchantID&quot; below.</td>
								</tr>
							<% end if %>
							<tr> 
								<td> <div align="right">Merchant ID:</div></td>
								<td width="479"> <input type="text" name="x_Login" size="30"></td>
							</tr>
							<tr> 
								<td> <div align="right">Public Key:</div></td>
								<td> <input name="x_Password" type="password" size="30"> </td>
							</tr>
							<tr> 
								<td> <div align="right">Private Key:</div></td>
								<td> <input name="x_Key" type="text" size="30"> </td>
							</tr>
							<tr> 
								<td> <div align="right">Transaction Type:</div></td>
								<td> <select name="x_Mode">
										<option value="0" <% if x_Mode="0" then %>selected<% end if %>>Authorize Only</option>
										<option value="1" <% if x_Mode="1" then %>selected<% end if %>>Sale</option>
									</select> </td>
							</tr>
							<tr> 
								<td><div align="right">Currency Code:</div></td>
								<td><input name="x_Curcode" type="text" value="<%=x_Curcode%>" size="6" maxlength="4"> 
									<a href="help_auth_codes.asp" target="_blank">Find Codes</a></td>
							</tr>
							<tr> 
								<td> <div align="right">Require CVV:</div></td>
								<td> <input type="radio" class="clearBorder" name="x_CVV" value="1" checked>
									Yes 
									<input name="x_CVV" type="radio" class="clearBorder" value="0" <% if x_CVV="0" then %>checked<%end if %>>
									No</td>
							</tr>
							<tr> 
								<td><div align="right"> 
										<input name="x_testmode" type="checkbox" class="clearBorder" value="1" <% if x_testmode=1 then%>checked<% end if%>> 
									</div></td>
								<td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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
<!-- New View End --><% end if %>