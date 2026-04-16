<%
	
Function gwPPPGetName()
	gwPPPGetName = "PayPal Payments Pro"
	If myCountry = "UK" Then
		gwPPPGetName = "Website Payments Pro"
	End If
End Function

Function gwPPPGetURL()
	url = ""

	Select Case myCountry
	Case "CA"
		url = "https://www.paypal.com/ca/webapps/mpp/paypal-payments-pro"
	Case "UK"
		url = "https://www.paypal.com/uk/webapps/mpp/pro"
	Case Else
		url = "https://www.paypal.com/webapps/mpp/paypal-payments-pro"
	End Select

	gwPPPGetURL = url
End Function

'--- Start PayPal Payments Advanced ---
Function gwPPPEdit()
	PayPalName = gwPPPGetName()
	PayPalPaymentURL="gwPPP.asp"
	ppGwcode=53			

	pcPay_PayPal_TransType=request.Form("pcPay_PayPal_TransType")
	pcPay_PayPal_Username=request.Form("pcPay_PayPal_Username")
	pcPay_PayPal_Subject=request.Form("pcPay_PayPal_Subject")
	pcPay_PayPal_Password=request.Form("pcPay_PayPal_Password")
	pcPay_PayPal_Sandbox=request.Form("pcPay_PayPal_Sandbox")
	pcPay_PayPal_Vendor=request.Form("pcPay_PayPal_Vendor")
	pcPay_PayPal_Partner=request.Form("pcPay_PayPal_Partner")
	pcPay_PayPal_Signature=request.Form("pcPay_PayPal_Signature")					
	pcPay_PayPal_Currency=request.Form("pcPay_PayPal_Currency")		
	pcPay_PayPal_CVC=request.Form("pcPay_PayPal_CVC")	
	pcPay_PayPal_CardTypes=request.Form("pcPay_PayPal_CardTypes")
	if pcPay_PayPal_Sandbox="YES" then
		pcPay_PayPal_Sandbox=1
	else
		pcPay_PayPal_Sandbox=0
	end if
	if pcPay_PayPal_CVC="" then
		pcPay_PayPal_CVC=0
	end if
	
	pcPay_PayPal_Layout=request.Form("pcPay_PayPale_Layout")
	pcPay_PayPal_Shape=request.Form("pcPay_PayPale_Shape")
	pcPay_PayPal_Size=request.Form("pcPay_PayPale_Size")
	pcPay_PayPal_Color=request.Form("pcPay_PayPale_Color")
	
	ppAB=1

	query="UPDATE pcPay_PayPal SET pcPay_PayPal_TransType="&pcPay_PayPal_TransType&", pcPay_PayPal_Username='"&pcPay_PayPal_Username&"', pcPay_PayPal_Subject='"&pcPay_PayPal_Subject&"', pcPay_PayPal_Password='"&pcPay_PayPal_Password&"', pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&", pcPay_PayPal_Signature='"&pcPay_PayPal_Signature&"', pcPay_PayPal_Currency='"&pcPay_PayPal_Currency&"', pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"', pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"', pcPay_PayPal_CardTypes='"&pcPay_PayPal_CardTypes&"', pcPay_PayPal_Layout='"&pcPay_PayPal_Layout&"', pcPay_PayPal_Shape='"&pcPay_PayPal_Shape&"', pcPay_PayPal_Size='"&pcPay_PayPal_Size&"', pcPay_PayPal_Color='"&pcPay_PayPal_Color&"' WHERE (((pcPay_PayPal_ID)=1));"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentDesc='"&PayPalName&"',sslUrl='" & PayPalPaymentURL & "',priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"',pcPayTypes_ppab=0 WHERE gwCode=53"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	set rs=nothing
	
end function

Function gwPPP()
	PayPalName = gwPPPGetName()
	PayPalPaymentURL="gwPPP.asp"
	ppGwcode=53

	varCheck=1
	'request gateway variables and insert them into the database
	pcPay_PayPal_TransType=request.Form("pcPay_PayPal_TransType")
	pcPay_PayPal_Username=request.Form("pcPay_PayPal_Username")
	pcPay_PayPal_Subject=request.Form("pcPay_PayPal_Subject")
	pcPay_PayPal_Password=request.Form("pcPay_PayPal_Password")
	pcPay_PayPal_Sandbox=request.Form("pcPay_PayPal_Sandbox")
	pcPay_PayPal_Vendor=request.Form("pcPay_PayPal_Vendor")
	pcPay_PayPal_Partner=request.Form("pcPay_PayPal_Partner")
	pcPay_PayPal_Signature=request.Form("pcPay_PayPal_Signature")					
	pcPay_PayPal_Currency=request.Form("pcPay_PayPal_Currency")		
	pcPay_PayPal_CVC=request.Form("pcPay_PayPal_CVC")	
	pcPay_PayPal_CardTypes=request.Form("pcPay_PayPal_CardTypes")
	if pcPay_PayPal_Sandbox="YES" then
		pcPay_PayPal_Sandbox=1
	else
		pcPay_PayPal_Sandbox=0
	end if
	if pcPay_PayPal_CVC="" then
		pcPay_PayPal_CVC=0
	end if
	ppAB=1

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
			
	err.clear
	err.number=0
			
	 

	query="UPDATE pcPay_PayPal SET pcPay_PayPal_TransType="&pcPay_PayPal_TransType&", pcPay_PayPal_Username='"&pcPay_PayPal_Username&"', pcPay_PayPal_Subject='"&pcPay_PayPal_Subject&"', pcPay_PayPal_Password='"&pcPay_PayPal_Password&"', pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&", pcPay_PayPal_Signature='"&pcPay_PayPal_Signature&"', pcPay_PayPal_Currency='"&pcPay_PayPal_Currency&"', pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"', pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"', pcPay_PayPal_CardTypes='"&pcPay_PayPal_CardTypes&"', pcPay_PayPal_Layout='"&pcPay_PayPal_Layout&"', pcPay_PayPal_Shape='"&pcPay_PayPal_Shape&"', pcPay_PayPal_Size='"&pcPay_PayPal_Size&"', pcPay_PayPal_Color='"&pcPay_PayPal_Color&"' WHERE (((pcPay_PayPal_ID)=1));"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'" & PayPalName & "','" & PayPalPaymentURL & "',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",53,N'"&paymentNickName&"',0)"
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
				
<% if request("gwchoice")="53" then
	pcConflictIdPayment = 0
	
	Select Case myCountry
	Case "US"
		pcPay_PayPal_Currency = "USD"
	case "CA"
		pcPay_PayPal_Currency = "CAD"
	Case "UK"
		pcPay_PayPal_Currency = "GBP"
	End Select

	if request("mode")="Edit" then
		
		query="SELECT pcPay_PayPal.pcPay_PayPal_TransType, pcPay_PayPal.pcPay_PayPal_Username, pcPay_PayPal.pcPay_PayPal_Password,  pcPay_PayPal.pcPay_PayPal_Sandbox, pcPay_PayPal.pcPay_PayPal_Signature, pcPay_PayPal.pcPay_PayPal_Currency, pcPay_PayPal.pcPay_PayPal_CVC, pcPay_PayPal.pcPay_PayPal_Partner, pcPay_PayPal.pcPay_PayPal_Vendor, pcPay_PayPal.pcPay_PayPal_Subject, pcPay_PayPal.pcPay_PayPal_CardTypes, pcPay_PayPal.pcPay_PayPal_Layout, pcPay_PayPal.pcPay_PayPal_Shape, pcPay_PayPal.pcPay_PayPal_Size, pcPay_PayPal.pcPay_PayPal_Color FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_PayPal_TransType=rs("pcPay_PayPal_TransType")
		pcPay_PayPal_Username=rs("pcPay_PayPal_Username")
		pcPay_PayPal_Password=rs("pcPay_PayPal_Password")
		pcPay_PayPal_Sandbox=rs("pcPay_PayPal_Sandbox")	
		pcPay_PayPal_Signature=rs("pcPay_PayPal_Signature")	
		pcPay_PayPal_Currency=rs("pcPay_PayPal_Currency")
		pcPay_PayPal_CVC=rs("pcPay_PayPal_CVC")
		pcPay_PayPal_Partner=rs("pcPay_PayPal_Partner")
		pcPay_PayPal_Vendor=rs("pcPay_PayPal_Vendor")
		pcPay_PayPal_Subject=rs("pcPay_PayPal_Subject")
		pcPay_PayPal_CardTypes=rs("pcPay_PayPal_CardTypes")
		If len(pcPay_PayPal_Subject)=0 Then
			pcPay_PayPal_Subject=""
		End If
		if pcPay_PayPal_Partner<>"" AND pcPay_PayPal_Vendor<>"" then
			pcPay_PayPal_Version = "UK"			
		else
			pcPay_PayPal_Version = "US"						
		end if
		if IsNull(pcPay_PayPal_CardTypes) then pcPay_PayPal_CardTypes=""
		
		pcPay_PayPal_Layout=rs("pcPay_PayPal_Layout")
		pcPay_PayPal_Shape=rs("pcPay_PayPal_Shape")
		pcPay_PayPal_Size=rs("pcPay_PayPal_Size")
		pcPay_PayPal_Color=rs("pcPay_PayPal_Color")
		
		set rs=nothing
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=53"
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
		
		%>
		<input type="hidden" name="mode" value="Edit">
    <% else
		'//Check if any other PayPal Services are activated.
				query= "SELECT idPayment, gwCode FROM payTypes WHERE gwCode IN (3, 46, 9, 99, 80, 999999)"
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
				case "9", "99"
					pcConflictDesc = gwPFLGetName()
				case "80"
					pcConflictDesc = gwPPAGetName()
				case "999999"
					pcConflictDesc = gwPPExGetName()
			end select %>
        	<div class="pcCPmessage">
        	  <p>You currently have <strong><%=pcConflictDesc%></strong> active for this store. In order to use <strong><%= gwPPPGetName() %></strong> you will need to first disable <strong><%=pcConflictDesc%></strong>.<br />
        	    <br />
        	  </p>
        	  <p><a href="pcConfigurePayment.asp?mode=Del&id=<%=pcConflictIdPayment%>&gwchoice=<%=pcConflictID%>&activate=53">Disable <%=pcConflictDesc%> and enable <%= gwPPPGetName() %></a></p>
        	  <br />
        	  <p><a href="pcPaymentSelection.asp">Back to payment selection</a><br />
        	    <br />
      	    </p>
            </div>
<% end if
		set rs = nothing
		
	end if %>
<% if pcConflictIdPayment = 0 then %>
	<input type="hidden" name="addGw" value="53">
    <div class="pcCPmessageSuccess">
        <% if request("mode")="Edit" then %>
            <p><strong>You're editing </strong><strong><%= gwPPPGetName() %></strong>
            <br />
            <br />
            <p><strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% else %>
            <p><strong>You've selected <%= gwPPPGetName() %></strong><br />
            <br />Fully customize your checkout pages and
            accept credit cards directly on your site. PayPal
            simplifies PCI compliance for you, plus you get
            Virtual Terminal at no added cost.<strong><br>
            <br>
            <a href="<%= gwPPPGetURL() %>" target="_blank">Sign Up and Learn More</a></strong>
            </p>
            <br />
            <p>To start accepting payments, please complete the process below.        <strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% end if %>
    </div>
	<table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/paypal_logo1.gif" width="253" height="80" /></td>
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
                                    Step 1: Configure Account - <%= gwPPPGetName() %>
                                </a>
                            </h4>
                        </div>
                        <div id="collapseOne" class="panel-collapse collapse in">
                            <div class="panel-body">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2" valign="top">
                                            <p>You must have a PayPal business account to use <%= gwPPPGetName() %>. If you don't have an account, <a href="<%= gwPPPGetURL() %>" target="_blank">sign up for one now</a>.</p><br />
                                            <p>You created this information when you signed up for <%= gwPPPGetName() %>. Enter it here to connect your account and allow payments. (Note: This is also your login information for <a href="http://manager.paypal.com" target="_blank">PayPal Manager</a>.)</p><br />
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td width="127" valign="top"><div align="right">Partner Name:</div></td>
                                        <td>
                                        <input type="text" value="<%=pcPay_PayPal_Partner%>" name="pcPay_PayPal_Partner" size="30" maxlength="250"></td>
                                    </tr>
                                    <tr> 
                                        <td width="127" valign="top"><div align="right">Merchant Login:</div></td>
                                        <td>
                                        <input type="text" value="<%=pcPay_PayPal_Vendor%>" name="pcPay_PayPal_Vendor" size="30" maxlength="250"></td>
                                    </tr>
                                    <tr> 
                                        <td width="127" valign="top"><div align="right">User:</div></td>
                                        <td><input type="text" value="<%=pcPay_PayPal_Username%>" name="pcPay_PayPal_Username" size="30" maxlength="50"></td>
                                    </tr>
                                    
                                    <tr> 
                                        <td valign="top"><div align="right">Password:</div></td>
                                        <td><input type="password" value="<%=pcPay_PayPal_Password%>" name="pcPay_PayPal_Password" size="30" maxlength="50"></td>
                                    </tr>
                                    <tr> 
                                        <td width="127" valign="top" nowrap><div align="right">Currency:</div></td>
                                        <td>
                                        	<select name="pcPay_PayPal_Currency">
                                                <option value="AUD" <% if pcPay_PayPal_Currency="AUD" then%>selected<% end if %>>Australian Dollars ($)</option>
                                                <option value="CAD" <% if pcPay_PayPal_Currency="CAD" then%>selected<% end if %>>Canadian Dollars (C $)</option>
                                                <option value="EUR" <% if pcPay_PayPal_Currency="EUR" then%>selected<% end if %>>Euros (â‚¬)</option>
                                                <option value="GBP" <% if pcPay_PayPal_Currency="GBP" then%>selected<% end if %>>Pounds Sterling (Â£)</option>
                                                <option value="JPY" <% if pcPay_PayPal_Currency="JPY" then%>selected<% end if %>>Yen (Â¥)</option>
                                                <option value="USD" <% if pcPay_PayPal_Currency="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
                                            </select>
                                    	</td>
                                    </tr>
                                    <tr> 
                                        <td valign="top" nowrap><div align="right">Transaction Type:</div></td>
                                        <td> 
                                            <select name="pcPay_PayPal_TransType">
                                                <option value="1" <% if pcPay_PayPal_TransType=1 then%>selected<%end if %>>Sale (Authorize and Capture)</option>
                                                <option value="2" <% if pcPay_PayPal_TransType=2 then%>selected<%end if %>>Authorize Only</option>
                                            </select>
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td>&nbsp;</td>
                                        <td>
                                            <input name="pcPay_PayPal_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if pcPay_PayPal_Sandbox=1 then%>checked<% end if %>>
                                            <strong>Enable Test Mode </strong>(Credit cards will not be charged)
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td>&nbsp;</td>
                                        <td>
                                            <input name="pcPay_PayPal_CVC" type="checkbox" class="clearBorder" value="1" <% if pcPay_PayPal_CVC=1 then%>checked<% end if %>>
                                            <strong>Enable Credit Card Security Code</strong>
                                        </td>
                                    </tr>
                                  <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                  </tr>
                                  <tr>
                                    <td>&nbsp;</td>
                                    <td><a class="pcCPhelp" href="helpOnline.asp?ref=806">More information on <%= gwPPPGetName() %></a></td>
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
<!-- New View End --><% end if %>
<% end if %>