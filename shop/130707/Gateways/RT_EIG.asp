<%
'---Start EIG---
Function gwEIGEdit()
	
		
	'// Select gateway variables
	query= "SELECT pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key FROM pcPay_EIG where pcPay_EIG_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?t=1&error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	If NOT rs.EOF Then	
		x_Login2=rs("pcPay_EIG_Username")
		x_Login2=enDeCrypt(x_Login2, scCrypPass)
		x_Password2=rs("pcPay_EIG_Password")
		x_Password2=enDeCrypt(x_Password2, scCrypPass)
		x_Key2=rs("pcPay_EIG_Key")
		x_Key2=enDeCrypt(x_Key2, scCrypPass)
	End If
	set rs=nothing
	
	'// Request Form Fields
	x_Type=request.Form("x_Type")
	x_Login=request.Form("x_Login")	
	x_Password=request.Form("x_Password")
	x_Key=request.Form("x_Key")
	cardTypes=request.Form("cardTypes")
	x_Curcode=request.Form("x_Curcode")	
	x_CVV=request.Form("x_CVV")
	x_SaveCards=request.Form("x_SaveCards")
	x_UseVault=request.Form("x_UseVault")
	x_testmode="0" ' Test mode can only be set from the payment gateway admin area
	
	'// Apply Form Field Logic
	if x_Curcode="" then
		x_Curcode="USD"
	end if
	x_URLMethod="gwEIGateway.asp"
	if x_Login="" then
		x_Login=x_Login2
	end if
	x_Login=enDeCrypt(x_Login, scCrypPass)
	if x_Password="" then
		x_Password=x_Password2
	end if
	x_Password=enDeCrypt(x_Password, scCrypPass)
	if x_Key="" then
		x_Key=x_Key2
	end if
	x_Key=enDeCrypt(x_Key, scCrypPass)
	
	'// Update EIG Table with Form Field Values
	query="UPDATE pcPay_EIG SET pcPay_EIG_Type='"&x_Type&"||"&cardTypes&"',pcPay_EIG_Username='"&x_Login&"',pcPay_EIG_Password='"&x_Password&"',pcPay_EIG_Key='"&x_Key&"',pcPay_EIG_Version='1.0',pcPay_EIG_Curcode='"&x_Curcode&"',pcPay_EIG_CVV="&x_CVV&",pcPay_EIG_SaveCards="&x_SaveCards&", pcPay_EIG_UseVault="&x_UseVault&", pcPay_EIG_TestMode="&x_testmode&" WHERE pcPay_EIG_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?t=2&error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	'// Request Standard Fields
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
	
	'// Update Standard payTypes Fields
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&x_URLMethod&"', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"', paymentDesc='NetSource Commerce Gateway' WHERE gwCode=67"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?t=3&error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	set rs=nothing	
	
	
end function

Function gwEIG()

	 
	
	varCheck=1

	'// Request Form Fields
	x_Type=request.Form("x_Type")	
	x_Login=request.Form("x_Login")
	x_Password=request.Form("x_Password")
	x_Key=request.Form("x_Key")
	cardTypes=request.Form("cardTypes")
	x_Curcode=request.Form("x_Curcode")	
	x_CVV=request.Form("x_CVV")
	x_SaveCards=request.Form("x_SaveCards")
	x_UseVault=request.Form("x_UseVault")
	x_testmode="0" ' Test mode can only be set from the payment gateway admin area

	'// Apply Form Field Logic
	if x_Curcode="" then
		x_Curcode="USD"
	end if
	x_URLMethod="gwEIGateway.asp"
	if x_Login="" then
		x_Login=x_Login2
	end if
	x_Login=enDeCrypt(x_Login, scCrypPass)
	if x_Password="" then
		x_Password=x_Password2
	end if
	x_Password=enDeCrypt(x_Password, scCrypPass)
	if x_Key="" then
		x_Key=x_Key2
	end if
	x_Key=enDeCrypt(x_Key, scCrypPass)

	'// Update EIG Table with Form Field Values
	query="UPDATE pcPay_EIG SET pcPay_EIG_Type='"&x_Type&"||"&cardTypes&"',pcPay_EIG_Username='"&x_Login&"',pcPay_EIG_Password='"&x_Password&"',pcPay_EIG_Key='"&x_Key&"',pcPay_EIG_Version='1.0',pcPay_EIG_Curcode='"&x_Curcode&"',pcPay_EIG_CVV="&x_CVV&",pcPay_EIG_SaveCards="&x_SaveCards&", pcPay_EIG_UseVault="&x_UseVault&",pcPay_EIG_TestMode="&x_testmode&" WHERE pcPay_EIG_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?t=4&err="&pcStrCustRefID
	end if
	
	'// Request Standard Fields
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

	'// Update Standard payTypes Fields
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'NetSource Commerce Gateway','"&x_URLMethod&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",67,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?t=5&err="&pcStrCustRefID
	end if

	set rs=nothing	
	
	
end function
%>
				
<% if request("gwchoice")="67" then
	if request("mode")="Edit" then

				query= "SELECT pcPay_EIG_Type, pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key,pcPay_EIG_Curcode,pcPay_EIG_CVV,pcPay_EIG_SaveCards,pcPay_EIG_UseVault FROM pcPay_EIG where pcPay_EIG_Id=1"
		
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?t=6&error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		x_Type=rs("pcPay_EIG_Type")
		x_Login=rs("pcPay_EIG_Username")
		x_Password=rs("pcPay_EIG_Password")
		x_Key=rs("pcPay_EIG_Key")
		x_Curcode=rs("pcPay_EIG_Curcode")
		x_CVV=rs("pcPay_EIG_CVV")
		x_SaveCards=rs("pcPay_EIG_SaveCards")
		x_UseVault=rs("pcPay_EIG_UseVault")
		set rs=nothing
		x_Login=enDeCrypt(x_Login, scCrypPass)
		x_Password=enDeCrypt(x_Password, scCrypPass)
		x_Key=enDeCrypt(x_Key, scCrypPass)
		
		x_TypeArray=Split(x_Type,"||")
		x_Type1=x_TypeArray(0)
		M="0"
		V="0"
		A="0"
		D="0"
		if ubound(x_TypeArray)=1 then
			x_Type2=x_TypeArray(1)
			cardTypeArray=split(x_Type2,", ")
			for i=0 to ubound(cardTypeArray)
				select case cardTypeArray(i)
					case "M"
						M="1" 
					case "V"
						V="1"
					case "D"
						D="1"
					case "A"
						A="1"
				end select
			next
		end if
		if x_Curcode="" then
			x_Curcode="USD"
		end if
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=67"
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
    <% end if %>
	<input type="hidden" name="addGw" value="67">
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="images/ei_logo_gradient_payment_gateway_175.jpg" alt="NetSource Commerce Payment Gateway" /></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>NetSource (NMI) Gateway</h4>    
                    <p>
                         DowCommerce is one of the most powerful, feature rich, internet payment gateways on the market.   DowCommerce helps you turn your &quot;internet business idea&quot; into an &quot;internet business reality&quot;. DowCommerce provides everything you need except for your website.
                    </p>
                    <p>

                    </p>
                    <p>
                        <a class="btn btn-info btn-xs" href="https://www.productcart.com/nc-payment-gateway.asp" target="_blank">Sign up</a> <a class="btn btn-default btn-xs" href="https://www.productcart.com/nc-payment-gateway.asp" target="_blank">Log in</a> <a class="btn btn-default btn-xs" href="http://wiki.productcart.com/productcart/early_impact_payment_gateway" target="_blank">Docs</a>        
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
                                    <% 
                                    dim EIG_LoginCnt, EIG_LoginEnd, EIG_LoginStart
                                    
                                    EIG_LoginCnt=(len(x_Key)-2)
                                    EIG_LoginEnd=right(x_Key,2)
                                    EIG_LoginStart=""
                                    for c=1 to EIG_LoginCnt
                                        EIG_LoginStart=EIG_LoginStart&"*"
                                    next %>
                                    <tr> 
                                        <td colspan="2">Current API Key:&nbsp;<%=EIG_LoginStart&EIG_LoginEnd%></td>
                                    </tr>
                                    <tr> 
                                        <td colspan="2"> For security reasons, your credentials are not shown (the security key is partially shown as a reference). If you need to edit your account information, please re-enter your credentials below. If you leave the fields blank, the current credentials will be used.</td>
                                    </tr>
                                <% end if %>
                
                                <tr>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Username:</div></td>
                                    <td width="479"> <input type="text" name="x_Login" size="30" autocomplete="off"></td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Password:</div></td>
                                    <td> <input name="x_Password" type="password" size="30" autocomplete="off"> </td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Security Key:</div></td>
                                    <td> <input name="x_Key" type="text" size="50" autocomplete="off"> </td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Transaction Type:</div></td>
                                    <td> <select name="x_Type" id="x_Type">
                                            <option value="AUTH_CAPTURE" <% if x_Type1="AUTH_CAPTURE" then %>selected<% end if %>>Sale</option>
                                            <option value="AUTH_ONLY" <% if x_Type1="AUTH_ONLY" then %>selected<% end if %>>Authorize Only</option>
                                        </select> 
                                    </td>
                                </tr>
                                <tr id="AUTH_ONLY"> 
                                    <td></td>
                                    <td>
                                            Select where to store credit card information: <a href="http://wiki.productcart.com/productcart/early_impact_payment_gateway#authorizations_credit_card_data_and_pci_compliance" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Help on this feature" /></a>
                                            <div style="margin-bottom: 15px;">
                                            <input name="x_UseVault" id="x_UseVault1" type="radio" class="clearBorder" value="1" checked>
                                            PCI compliant vault. <a href="https://www.productcart.com/nc-payment-gateway.asp#fees" target="_blank">An additional fee applies</a>. <br />
                                            <input name="x_UseVault" id="x_UseVault2" type="radio" class="clearBorder" value="0" <% if x_UseVault="0" then %>checked<%end if %>>
                                            ProductCart database
                                            </div>
                                        <div id="AUTH_CAPTURE">
                                            <div class="pcCPnotes">This option is only applicable to "Authorize Only"  </div>
                                        </div>
                                        <script type=text/javascript>
                                            $pc(document).ready(function() {
                                                toggleChoice(1);
                                                $pc('#x_Type').change(function() {
                                                    toggleChoice(2);									
                                                });
                                                function toggleChoice(a) {
                                                    var TransactionType = $pc('#x_Type').val();
                                                    if (TransactionType=='AUTH_ONLY') {
                                                        $pc('#AUTH_ONLY').show();
                                                        $pc('#AUTH_CAPTURE').hide();
                                                        if (a==2) { $pc("#x_UseVault1").attr("checked", "checked") }; 
                                                    } else {
                                                        $pc('#AUTH_CAPTURE').show();
                                                        $pc('#AUTH_ONLY').hide();
                                                        if (a==2) { $pc("#x_UseVault1").attr("checked", "checked") };  
                                                    }
                                                }
                                            });
                                        </script>
                                    </td>
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
                                    <td> <div align="right">Accepted Cards:</div></td>
                                    <td>
                                        <% if V="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="V" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="V"> 
                                        <% end if %> Visa 
                                        <% if M="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="M" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="M"> 
                                        <% end if %> MasterCard 
                                        <% if A="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="A" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="A"> 
                                        <% end if %>  American Express 
                                        <% if D="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="D" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="D"> 
                                        <% end if %> Discover
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" class="pcCPspacer"><hr /></td>
                                </tr>
                
                                <tr> 
                                    <td colspan="2" align="center">Allow customers to <strong>save their credit card(s)</strong> for use during a future purchase: <a href="http://wiki.productcart.com/productcart/early_impact_payment_gateway" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Help on this feature" /></a></td>
                                </tr>
                                <tr> 
                                    <td></td>
                                    <td>
                                        <input name="x_SaveCards" type="radio" class="clearBorder" value="1" checked>
                                        Yes. Credit card data is saved into a PCI Compliant vault. <a href="https://www.productcart.com/nc-payment-gateway.asp#fees" target="_blank">An additional fee applies</a>.<br />
                                        <input name="x_SaveCards" type="radio" class="clearBorder" value="0" <% if x_SaveCards="0" then %>checked<%end if %>>
                                        No
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
