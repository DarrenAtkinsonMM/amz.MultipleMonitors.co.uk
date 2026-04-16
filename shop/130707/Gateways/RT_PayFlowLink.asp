<%

Function gwPFLGetName()
	gwPFLGetName = "PayPal Payflow Link"
End Function

Function gwPFLGetURL()
	gwPFLGetURL = "https://www.paypal.com/webapps/mpp/payflow-payment-gateway"
End Function

'--- Start Payflow Link ---
Function gwPFLEdit()
	'request gateway variables and insert them into the pcPay_PayPal table
	query="SELECT pcPay_PayPal_Username,pcPay_PayPal_Password,pcPay_PayPal_Vendor FROM pcPay_PayPal where pcPay_PayPal_ID=1;"

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	pcPay_PayPal_Username2=rs("pcPay_PayPal_Username")
	'pcPay_PayPal_Username2=enDeCrypt(pcPay_PayPal_Username2, scCrypPass)
	
	pcPay_PayPal_Password2=rs("pcPay_PayPal_Password")
	'pcPay_PayPal_Password2=enDeCrypt(pcPay_PayPal_Password2, scCrypPass)
	
	pcPay_PayPal_Vendor2=rs("pcPay_PayPal_Vendor")
	'pcPay_PayPal_Vendor2=enDeCrypt(pcPay_PayPal_Vendor2, scCrypPass)
	
	set rs=nothing
	pcPay_PayPal_Username=request.Form("pcPay_PayPal_Username")
	if pcPay_PayPal_Username="" then
		pcPay_PayPal_Username=pcPay_PayPal_Username2
	end if
	'pcPay_PayPal_Username=enDeCrypt(pcPay_PayPal_Username, scCrypPass)
	pcPay_PayPal_Password=request.Form("pcPay_PayPal_Password")
	if pcPay_PayPal_Password="" then
		pcPay_PayPal_Password=pcPay_PayPal_Password2
	end if
	'pcPay_PayPal_Password=enDeCrypt(pcPay_PayPal_Password, scCrypPass)

	pcPay_PayPal_Vendor=request.Form("pcPay_PayPal_Vendor")
	if pcPay_PayPal_Vendor="" then
		pcPay_PayPal_Vendor=pcPay_PayPal_Vendor2
	end if
	'pcPay_PayPal_Vendor=enDeCrypt(pcPay_PayPal_Vendor, scCrypPass)
	
	pcPay_PayPal_Partner=request.Form("pcPay_PayPal_Partner")
	pcPay_PayPal_Sandbox=request.Form("pcPay_PayPal_Sandbox")
	pcPay_PayPal_CVC=request.Form("pcPay_PayPal_CVC")
	if pcPay_PayPal_Sandbox="" then
		pcPay_PayPal_Sandbox=0
	end if
	pcPay_PayPal_TransType=request.Form("pcPay_PayPal_TransType")
	
	pcPay_PayPal_Layout=request.Form("pcPay_PayPale_Layout")
	pcPay_PayPal_Shape=request.Form("pcPay_PayPale_Shape")
	pcPay_PayPal_Size=request.Form("pcPay_PayPale_Size")
	pcPay_PayPal_Color=request.Form("pcPay_PayPale_Color")
	
	query="UPDATE pcPay_PayPal SET pcPay_PayPal_Username='"&pcPay_PayPal_Username&"',pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"',pcPay_PayPal_Password='"&pcPay_PayPal_Password &"',pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"',pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&",pcPay_PayPal_TransType="&pcPay_PayPal_TransType&",pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Layout='"&pcPay_PayPal_Layout&"', pcPay_PayPal_Shape='"&pcPay_PayPal_Shape&"', pcPay_PayPal_Size='"&pcPay_PayPal_Size&"', pcPay_PayPal_Color='"&pcPay_PayPal_Color&"' WHERE pcPay_PayPal_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET paymentDesc='" & gwPFLGetName() & "', pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"',pcPayTypes_ppab=0 WHERE gwCode=99"
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

Function gwPFL()
	varCheck=1
	'request gateway variables and insert them into the pcPay_PayPal table
	pcPay_PayPal_Username=request.Form("pcPay_PayPal_Username")
	'pcPay_PayPal_Username=enDeCrypt(pcPay_PayPal_Username, scCrypPass)

	pcPay_PayPal_Password=request.Form("pcPay_PayPal_Password")
	'pcPay_PayPal_Password=enDeCrypt(pcPay_PayPal_Password, scCrypPass)

	pcPay_PayPal_Vendor=request.Form("pcPay_PayPal_Vendor")
	'pcPay_PayPal_Vendor=enDeCrypt(pcPay_PayPal_Vendor, scCrypPass)

	pcPay_PayPal_Sandbox=request.Form("pcPay_PayPal_Sandbox")
	pcPay_PayPal_Partner=request.Form("pcPay_PayPal_Partner")
	if pcPay_PayPal_Sandbox="" then
		pcPay_PayPal_Sandbox=0
	end if
	pcPay_PayPal_TransType=request.Form("pcPay_PayPal_TransType")
	pcPay_PayPal_CVC=request.Form("pcPay_PayPal_CVC")
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
	
	pcPay_PayPal_Layout=request.Form("pcPay_PayPale_Layout")
	pcPay_PayPal_Shape=request.Form("pcPay_PayPale_Shape")
	pcPay_PayPal_Size=request.Form("pcPay_PayPale_Size")
	pcPay_PayPal_Color=request.Form("pcPay_PayPale_Color")

	query="UPDATE pcPay_PayPal SET pcPay_PayPal_Username='"&pcPay_PayPal_Username&"',pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"' ,pcPay_PayPal_Password='"&pcPay_PayPal_Password &"' ,pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"',pcPay_PayPal_Currency='na',pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&",pcPay_PayPal_TransType="&pcPay_PayPal_TransType&",pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Layout='"&pcPay_PayPal_Layout&"', pcPay_PayPal_Shape='"&pcPay_PayPal_Shape&"', pcPay_PayPal_Size='"&pcPay_PayPal_Size&"', pcPay_PayPal_Color='"&pcPay_PayPal_Color&"' WHERE pcPay_PayPal_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'" & gwPFLGetName() & "','gwPFL.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",99,N'"&paymentNickName&"',0)"
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
				
<% if request("gwchoice")="99" then

	if request("mode")="Edit" then
				
		query= "SELECT pcPay_PayPal_Username,pcPay_PayPal_Partner,pcPay_PayPal_Password,pcPay_PayPal_Vendor,pcPay_PayPal_Sandbox,pcPay_PayPal_TransType,pcPay_PayPal_CVC, pcPay_PayPal.pcPay_PayPal_Layout, pcPay_PayPal.pcPay_PayPal_Shape, pcPay_PayPal.pcPay_PayPal_Size, pcPay_PayPal.pcPay_PayPal_Color FROM pcPay_PayPal WHERE pcPay_PayPal_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_PayPal_Username=rs("pcPay_PayPal_Username")
		'pcPay_PayPal_Username=enDeCrypt(pcPay_PayPal_Username, scCrypPass)

		pcPay_PayPal_Partner=rs("pcPay_PayPal_Partner")
		pcPay_PayPal_Password=rs("pcPay_PayPal_Password")
		'pcPay_PayPal_Password=enDeCrypt(pcPay_PayPal_Password, scCrypPass)

		pcPay_PayPal_Vendor=rs("pcPay_PayPal_Vendor")
		'pcPay_PayPal_Vendor=enDeCrypt(pcPay_PayPal_Vendor, scCrypPass)

		pcPay_PayPal_Sandbox=rs("pcPay_PayPal_Sandbox")
		pcPay_PayPal_TransType=rs("pcPay_PayPal_TransType")
		pcPay_PayPal_CVC=rs("pcPay_PayPal_CVC")
		
		pcPay_PayPal_Layout=rs("pcPay_PayPal_Layout")
		pcPay_PayPal_Shape=rs("pcPay_PayPal_Shape")
		pcPay_PayPal_Size=rs("pcPay_PayPal_Size")
		pcPay_PayPal_Color=rs("pcPay_PayPal_Color")
		
		set rs=nothing
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=99"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName="Credit Card"
		else
			pcv_processOrder=rs("pcPayTypes_processOrder")
			pcv_setPayStatus=rs("pcPayTypes_setPayStatus")
			priceToAdd=rs("priceToAdd")
			percentageToAdd=rs("percentageToAdd")
			paymentNickName=rs("paymentNickName")
			if percentageToAdd<>"0" then
				priceToAddType="percentage"
			end if
			if priceToAdd<>"0" then
				priceToAddType="price"
			end if
		end if

		set rs=nothing
		
		
		dim pcPay_PayPal_UsernameCnt,pcPay_PayPal_UsernameEnd,pcPay_PayPal_UsernameStart
		pcPay_PayPal_UsernameCnt=(len(pcPay_PayPal_Username)-2)
		pcPay_PayPal_UsernameEnd=right(pcPay_PayPal_Username,2)
		pcPay_PayPal_UsernameStart=""
		for c=1 to pcPay_PayPal_UsernameCnt
			pcPay_PayPal_UsernameStart=pcPay_PayPal_UsernameStart&"*"
		next
		%>
    
		<input type="hidden" name="mode" value="Edit">
		<% else
				'//Check if any other PayPal Services are activated.
				query= "SELECT idPayment, gwCode FROM payTypes WHERE gwCode IN (3, 46, 53, 80, 999999)"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				if NOT rs.eof then
					pcConflictIdPayment = rs("idPayment")
					pcConflictID = rs("gwCode")
					select case pcConflictID
						case "3"
							pcConflictDesc = gwPPGetName()
						case "46"
							pcConflictDesc = gwPPPDPGetName()
						case "53"
							pcConflictDesc = gwPPPGetName()
						case "80"
							pcConflictDesc = gwPPAGetName()
						case "999999"
							pcConflictDesc = gwPPExGetName()
					end select 
				%>
        	<div class="pcCPmessage">
        	  <p>You currently have <strong><%=pcConflictDesc%></strong> active for this store. In order to use <strong><%= gwPFLGetName() %></strong> you will need to first disable <strong><%=pcConflictDesc%></strong>.<br />
        	    <br />
        	  </p>
        	  <p><a href="pcConfigurePayment.asp?mode=Del&id=<%=pcConflictIdPayment%>&gwchoice=<%=pcConflictID%>&activate=99">Disable <%=pcConflictDesc%> and enable <%= gwPFLGetName() %></a></p>
        	  <br />
        	  <p><a href="pcPaymentSelection.asp">Back to payment selection</a><br />
        	    <br />
      	    </p>
            </div>
		<% end if
		set rs = nothing
		end if %>
	<% if pcConflictIdPayment = 0 then %>
	<input type="hidden" name="addGw" value="99">
    <div class="pcCPmessageSuccess">
  <% if request("mode")="Edit" then %>
            <p>
                <strong>You're editing <%= gwPFLGetName() %></strong>
                - Embedded Payment Integration<br />
                <br />
                <p><strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
                <br />
            </p>
        <% else %>
            <p><strong>You've selected <%= gwPFLGetName() %></strong> - Embedded Payment Integration<br />
            <br />
            </strong>Connect your merchant account with a
            PCI-compliant gateway. Setup is quick and
            customers pay without leaving your site.<br />
            <br />
            <strong> <a href="<%= gwPFLGetURL() %>" target="_blank">Sign Up and Learn More</a></strong>
            <br />
            <br />
            To start accepting payments, please complete the process below.
            <strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
    	<% end if %>
    </div>
    <br />
	<table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/payflow_logo.png"></td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    <%
		if pcPay_PayPal_Layout="" then
			pcPay_PayPal_Layout="vertical"
		end if
		if pcPay_PayPal_Shape="" then
			pcPay_PayPal_Shape="pill"
		end if
	%>
    <table width="100%">
        <tr>
            <td>
            
                <div id="accordion" class="panel-group">
                    
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapseOne">
                                    Step 1: Payflow Account Information
                                </a>
                            </h4>
                        </div>
                        <div id="collapseOne" class="panel-collapse collapse in">
                            <div class="panel-body">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <% if request("mode")="Edit" then %>
                                    <tr> 
                                        <td>Current User</td>
                                        <td width="83%">:&nbsp;<%=pcPay_PayPal_UsernameStart&pcPay_PayPal_UsernameEnd%></td>
                                    </tr>
                                    <tr> 
                                        <td colspan="2"><br />
                                            For security reasons, your &quot;Login&quot; is only 
                                            partially shown on this page. If you need to edit your 
                                            account information, please re-enter your &quot;Login&quot; 
                                            below.</td>
                                    </tr>
                                    <% else %>
                                    <tr>
                                        <td colspan="2">You must have a PayPal Payflow account to use Payflow Link. If you don't have an account, sign up for one now. <a href="<%= gwPFLGetURL() %>" target="_blank">Sign up now</a>
                                    <br />
                                    <br />
                                    Enter your PayPal Payflow Information You created this information when you signed up for <%= gwPFLGetName() %>. Enter it here to connect your account and allow payments. (Note: This is also your login information for <a href="https://manager.paypal.com/" target="_blank">PayPal Manager</a>.)<br /></td>
                                    </tr>
                                    <% end if %>
                                    <% if pcPay_PayPal_Partner&""="" then
                                        pcPay_PayPal_Partner="PayPal"
                                    end if %>
                                    <tr> 
                                        <td width="17%" align="right">Partner Name:</td>
                                        <td><input type="text" value="<%=pcPay_PayPal_Partner%>" name="pcPay_PayPal_Partner" size="24"></td>
                                    </tr>
                                    <tr> 
                                        <td width="17%" align="right">Merchant Login:</td>
                                        <td><input type="text" value="" name="pcPay_PayPal_Vendor" size="24"></td>
                                    </tr>
                                    <tr>
                                        <td align="right">User:</td>
                                        <td><input type="text" value="" name="pcPay_PayPal_Username" size="24" /></td>
                                    </tr>
                                    <tr>
                                        <td align="right">Password:</td>
                                        <td><input type="password" value="" name="pcPay_PayPal_Password" size="24" /></td>
                                    </tr>
                                    <tr> 
                                        <td width="17%" align="right">Transaction Type:</td>
                                        <td>
                                            <select name="pcPay_PayPal_TransType">
                                                <option value="1" <% if pcPay_PayPal_TransType="1" then response.write "selected"%>>Sale (Authorize &amp; Capture)</option>
                                                <option value="2" <% if pcPay_PayPal_TransType="2" then response.write "selected"%>>Authorize Only</option>
                                            </select>
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td align="right">Require CVC:</td>
                                        <td> 
                                            <input type="radio" class="clearBorder" name="pcPay_PayPal_CVC" value="1" <% If pcPay_PayPal_CVC="1" Then Response.Write "checked" %>> Yes 
                                          <input name="pcPay_PayPal_CVC" type="radio" class="clearBorder" value="0" <% If pcPay_PayPal_CVC<>"1" Then Response.Write "checked" %>> No
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td align="right">Enable Test Mode</td>
                                        <td>
                                            <input type="checkbox" class="clearBorder" name="pcPay_PayPal_Sandbox" value="1" <% If pcPay_PayPal_Sandbox="1" Then Response.Write "Checked" %>>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <tr>
                                        <td>&nbsp;</td>
                                        <td>
                                            <a class="pcCPhelp" href="helpOnline.asp?ref=801">More information on <%= gwPFLGetName() %></a>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>&nbsp;</td>
                                        <td class="pcSubmenuContent">&nbsp;</td>
                                    </tr>
                                </table>
                            </div>
                        </div> 
                    </div>
                    
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h4 class="panel-title">
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapse2">
                                    Step 2: Configure Settings
                                </a>
                            </h4>
                        </div>
                        <div id="collapse2" class="panel-collapse collapse">
                            <div class="panel-body">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td colspan="2">
                                        	<p>You must adjust these settings to accept payments. To make configuration easier, ProductCart will adjust any additional settings not listed here.</p>
                                        	<p>
                                                <ol>
                                                    <li>Log in to <a href="http://manager.paypal.com" target="_blank"><strong>PayPal Manager</strong></a>.</li>
                                                    <li>Select <strong>Service Settings</strong>.</li>
                                                    <li>Select <strong>Hosted Checkout Pages</strong>, select <strong>Set up</strong>.</li>
                                                    <li>Under <strong>Security Options</strong>, please set <strong>Enable Secure Token</strong> to &quot;Yes&quot;.</li>
                                                </ol>
                                            </p>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="2"><p><br />Use the following settings to style how PayPal Express Checkout will show on the cart page.</p></td>
                                    </tr>
                                    <tr>
                                        <td width="127" align="right">Layout:</td>
                                        <td>
                                        	<input type="radio" class="clearBorder" name="pcPay_PayPale_Layout" value="vertical" <% if pcPay_PayPal_Layout="vertical" then%>checked<%end if%>>Vertical&nbsp;&nbsp;&nbsp;
                                            <input type="radio" class="clearBorder" name="pcPay_PayPale_Layout" value="horizontal" <% if pcPay_PayPal_Layout="horizontal" then%>checked<%end if%>>Horizontal
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">Shape:</td>
                                        <td>
                                        	<input type="radio" class="clearBorder" name="pcPay_PayPale_Shape" value="pill" <% if pcPay_PayPal_Shape="pill" then%>checked<%end if%>>Pill&nbsp;&nbsp;&nbsp;
                                            <input type="radio" class="clearBorder" name="pcPay_PayPale_Shape" value="rect" <% if pcPay_PayPal_Shape="rect" then%>checked<%end if%>>Rectangular
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">Size:</td>
                                        <td>
                                        	<select name="pcPay_PayPale_Size">
                                                <option value="medium" <% if pcPay_PayPal_Size="medium" then%>selected<% end if %>>Medium</option>
                                                <option value="large" <% if pcPay_PayPal_Size="large" then%>selected<% end if %>>Large</option>
                                                <option value="responsive" <% if pcPay_PayPal_Size="responsive" then%>selected<% end if %>>Responsive</option>
                                            </select>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">Color:</td>
                                        <td>
                                        	<select name="pcPay_PayPale_Color">
                                                <option value="gold" <% if pcPay_PayPal_Color="gold" then%>selected<% end if %>>Gold</option>
                                                <option value="blue" <% if pcPay_PayPal_Color="blue" then%>selected<% end if %>>Blue</option>
                                                <option value="silver" <% if pcPay_PayPal_Color="silver" then%>selected<% end if %>>Silver</option>
                                                <option value="black" <% if pcPay_PayPal_Color="black" then%>selected<% end if %>>Black</option>
                                            </select>
                                        </td>
                                    </tr>
                                    <tr>
                                    	<td colspan="2">&nbsp;</td>
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
                                        <td colspan="2">&nbsp;</td>
                                    </tr>
                                    <tr> 
                                        <td width="10%" nowrap="nowrap"><div align="left">Payment Name:&nbsp;</div></td>
                                        <td width="90%"><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2">&nbsp;</td>
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
                                        <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=301"></a></td>
                                    </tr>
                                    <tr> 
                                        <td>
                                        	When orders are placed, set the payment status to:
                                            <select name="pcv_setPayStatus">
                                                <option value="3" selected="selected">Default</option>
                                                <option value="0" <%if pcv_setPayStatus="0" then%>selected<%end if%>>Pending</option>
                                                <option value="1" <%if pcv_setPayStatus="1" then%>selected<%end if%>>Authorized</option>
                                                <option value="2" <%if pcv_setPayStatus="2" then%>selected<%end if%>>Paid</option>
                                            </select>
                                        	&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=302"></a>
                                        </td>
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
<!-- New View End --><% end if %>