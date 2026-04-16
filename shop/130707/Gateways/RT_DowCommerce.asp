<%
'---Start DowCommerce---
Function gwDowComEdit()
		'request gateway variables and insert them into the pcPay_DowCom table
	query= "SELECT pcPay_Dow_MerchantID,pcPay_Dow_MerchantPassword FROM pcPay_DowCom where pcPay_Dow_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_Dow_MerchantID2=rs("pcPay_Dow_MerchantID")
	'decrypt
	pcPay_Dow_MerchantID2=enDeCrypt(pcPay_Dow_MerchantID2, scCrypPass)
	pcPay_Dow_MerchantPassword2=rs("pcPay_Dow_MerchantPassword")
	'decrypt
	pcPay_Dow_MerchantPassword2=enDeCrypt(pcPay_Dow_MerchantPassword2, scCrypPass)
	set rs=nothing
	pcPay_Dow_TransType=request.Form("pcPay_Dow_TransType")
	pcPay_Dow_MerchantID=request.Form("pcPay_Dow_MerchantID")
	if pcPay_Dow_MerchantID="" then
		pcPay_Dow_MerchantID=pcPay_Dow_MerchantID2
	end if
	'encrypt
	pcPay_Dow_MerchantID=enDeCrypt(pcPay_Dow_MerchantID, scCrypPass)
	pcPay_Dow_MerchantPassword=request.Form("pcPay_Dow_MerchantPassword")
	if pcPay_Dow_MerchantPassword="" then
		pcPay_Dow_MerchantPassword=pcPay_Dow_MerchantPassword2
	end if
	'encrypt
	pcPay_Dow_MerchantPassword=enDeCrypt(pcPay_Dow_MerchantPassword, scCrypPass)
	pcPay_Dow_CardTypes=request.Form("cardTypes")
	x_URLMethod="gwDowCom.asp"

	pcPay_Dow_CVC=request.Form("pcPay_Dow_CVC")
	pcPay_Dow_eCheck=request.Form("pcPay_Dow_eCheck")
	
	pcPay_Dow_eCheckPending=request.Form("pcPay_Dow_eCheckPending")
	pcPay_Dow_TestMode=request.Form("pcPay_Dow_TestMode")
	if pcPay_Dow_TestMode="YES" then
		pcPay_Dow_TestMode="1"
	else
		pcPay_Dow_TestMode="0"
	end if
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	end if

	
	query="UPDATE pcPay_DowCom SET pcPay_Dow_TransType='"&pcPay_Dow_TransType&"',pcPay_Dow_MerchantID='"&pcPay_Dow_MerchantID&"',pcPay_Dow_MerchantPassword='"&pcPay_Dow_MerchantPassword&"',pcPay_Dow_CardTypes='"&pcPay_Dow_CardTypes&"',pcPay_Dow_CVC="&pcPay_Dow_CVC&",pcPay_Dow_TestMode="&pcPay_Dow_TestMode&",pcPay_Dow_eCheck="&pcPay_Dow_eCheck&",pcPay_Dow_eCheckPending="&pcPay_Dow_eCheckPending&" where pcPay_Dow_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")    
    
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&x_URLMethod&"', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=60"

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	if pcPay_Dow_eCheck="1" then
		query="SELECT * FROM payTypes WHERE gwCode=61"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if rs.eof then
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('Dow Commerce eCheck','gwDowComCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",61,N'"&paymentNickName2&"')"
		else
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentNickName=N'"&paymentNickName2&"' WHERE gwCode=61"
		end if
	else
		query="DELETE FROM payTypes WHERE gwCode=61"
	end if
	

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

Function gwDowCom()
	varCheck=1
	'request gateway variables and insert them into the pcPay_DowCom table

	x_URLMethod="gwDowCom.asp"
	'x_AIMType="KEY"
	pcPay_Dow_CVC=request.Form("pcPay_Dow_CVC")
	pcPay_Dow_eCheck=request.Form("pcPay_Dow_eCheck")
	'x_secureSource=request.Form("x_secureSource")
	pcPay_Dow_eCheckPending=request.Form("pcPay_Dow_eCheckPending")
	pcPay_Dow_TestMode=request.Form("pcPay_Dow_TestMode")
	if pcPay_Dow_TestMode="YES" then
		pcPay_Dow_TestMode="1"
	else
		pcPay_Dow_TestMode="0"
	end if
	pcPay_Dow_TransType=request.Form("pcPay_Dow_TransType")
	pcPay_Dow_MerchantPassword=request.Form("pcPay_Dow_MerchantPassword")
	'encrypt
	pcPay_Dow_MerchantPassword=enDeCrypt(pcPay_Dow_MerchantPassword, scCrypPass)
	pcPay_Dow_CardTypes=request.Form("cardTypes")
	pcPay_Dow_MerchantID=request.Form("pcPay_Dow_MerchantID")
	If pcPay_Dow_MerchantID="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add DowCommerce as your payment gateway. <b>""Login ID""</b> is a required field.")
	End If
	'encrypt
	pcPay_Dow_MerchantID=enDeCrypt(pcPay_Dow_MerchantID, scCrypPass)
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
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	end if
	
		
	err.clear
	err.number=0
	
	 
	
	query="UPDATE pcPay_DowCom SET pcPay_Dow_TransType='"&pcPay_Dow_TransType&"',pcPay_Dow_MerchantID='"&pcPay_Dow_MerchantID&"',pcPay_Dow_MerchantPassword='"&pcPay_Dow_MerchantPassword&"',pcPay_Dow_CardTypes='"& pcPay_Dow_CardTypes&"',pcPay_Dow_CVC="&pcPay_Dow_CVC&",pcPay_Dow_TestMode="&pcPay_Dow_TestMode&", pcPay_Dow_eCheck="&pcPay_Dow_eCheck&",pcPay_Dow_eCheckPending="&pcPay_Dow_eCheckPending&" WHERE pcPay_Dow_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'DowCommerce','"&x_URLMethod&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",60,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if pcPay_Dow_eCheck="1" then
		query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Dow Commerce eCheck','gwDowComCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",61,N'"&paymentNickName2&"')"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

	end if
	
	set rs=nothing
	
	
end function
%>
<% if request("gwchoice")="60" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
		query="select * from pcPay_DowCom"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		
		call closeDb()
response.redirect "upddbDowCom.asp"
	else
		set rs=nothing
		
	end if

	if request("mode")="Edit" then
				query= "SELECT pcPay_Dow_TransType,pcPay_Dow_MerchantID,pcPay_Dow_MerchantPassword,pcPay_Dow_CardTypes,pcPay_Dow_CVC,pcPay_Dow_TestMode,pcPay_Dow_eCheck,pcPay_Dow_eCheckPending FROM pcPay_DowCom where pcPay_Dow_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_Dow_TransType=rs("pcPay_Dow_TransType")
		pcPay_Dow_MerchantID=rs("pcPay_Dow_MerchantID")
		'decrypt
		pcPay_Dow_MerchantID=enDeCrypt(pcPay_Dow_MerchantID, scCrypPass)
		pcPay_Dow_MerchantPassword=rs("pcPay_Dow_MerchantPassword")
		'decrypt
		pcPay_Dow_MerchantPassword=enDeCrypt(pcPay_Dow_MerchantPassword, scCrypPass)
		pcPay_Dow_CardTypes=rs("pcPay_Dow_CardTypes")
		pcPay_Dow_CVC=rs("pcPay_Dow_CVC")
		pcPay_Dow_TestMode=rs("pcPay_Dow_TestMode")
		pcPay_Dow_eCheck=rs("pcPay_Dow_eCheck")
     	pcPay_Dow_eCheckPending=rs("pcPay_Dow_eCheckPending")

		cardTypeArray=Split(pcPay_Dow_CardTypes,", ")
		
		M="0"
		V="0"
		A="0"
		D="0"
		
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
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=60"
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


		
		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=61"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName2="Check"
		else
			paymentNickName2=rs("paymentNickName")
		end if

		
		%>
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="60">
    <table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="Gateways/logos/dowcommerce.png" width="238" height="83"></td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>DowCommerce Payment Gateway & Internet Merchant Account</h4>    
                    <p>
                         DowCommerce is one of the most powerful, feature rich, internet payment gateways on the market.   DowCommerce helps you turn your &quot;internet business idea&quot; into an &quot;internet business reality&quot;. DowCommerce provides everything you need except for your website.
                    </p>
                    <p>
                        ProductCart is integration with DowCommerce using the Direct Post Method.<br />
                        This is the most secure and <u>preferred</u> method 
                        of processing payments.&nbsp; Credit Cards are authorized 
                        on your Web server in real-time. Your customers are 
                        not forwarded to any 3rd party Payment Forms.&nbsp; 
                        Review all security documentation from DowCommerce 
                        before utilizing this method. You must have <u>SSL</u> enabled in ProductCart to use this feature.
                    </p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.dowcommerce.com" target="_blank">Learn More</a>                      
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
                                    <% if request("mode")="Edit" then 
                                        A_LoginCnt=(len(pcPay_Dow_MerchantID)-2)
                                        A_LoginEnd=right(pcPay_Dow_MerchantID,2)
                                        A_LoginStart=""
                                        for c=1 to A_LoginCnt
                                            A_LoginStart=A_LoginStart&"*"
                                        next %>
                                        <tr> 
                                            <td colspan="2">Current Login ID:&nbsp;<%=A_LoginStart&A_LoginEnd%></td>
                                        </tr>
                                        <tr> 
                                            <td colspan="2"> For security reasons, your &quot;Login ID&quot; is 
                                                only partially shown on this page. If you need to edit 
                                                your account information, please re-enter your &quot;Login 
                                                ID&quot; below.</td>
                                        </tr>
                                    <% end if %>
                
                                  <tr>
                                    <td width="128" nowrap>&nbsp;</td>
                                    <td class="pcSubmenuContent">&nbsp;</td>
                                  </tr>
                                <tr> 
                                    <td> <div align="right">Login ID:</div></td>
                                    <td width="609"> <input type="text" name="pcPay_Dow_MerchantID" size="30"></td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Password:</div></td>
                                    <td> <input name="pcPay_Dow_MerchantPassword" type="text" size="30"> </td>
                                </tr>
                                <tr> 
                                    <td nowrap="nowrap"> <div align="right">Transaction Type:</div></td>
                                    <td> <select name="pcPay_Dow_TransType">
                                            <option value="sale" <% if pcPay_Dow_TransType="sale" then %>selected<% end if %>>Sale</option>
                                            <option value="auth" <% if pcPay_Dow_TransType="auth" then %>selected<% end if %>>Authorize Only</option>
                                        </select> </td>
                                </tr>
                                 <tr> 
                                    <td> <div align="right">Require CVV:</div></td>
                                    <td> <input type="radio" class="clearBorder" name="pcPay_Dow_CVC" value="1" checked>
                                        Yes 
                                        <input name="pcPay_Dow_CVC" type="radio" class="clearBorder" value="0" <% if pcPay_Dow_CVC="0" then %>checked<%end if %>>
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
                                    <td> <div align="right">Accept Checks:</div></td>
                                    <td> <input type="radio" class="clearBorder" name="pcPay_Dow_eCheck" value="1" checked>
                                        Yes 
                                        <input name="pcPay_Dow_eCheck" type="radio" class="clearBorder" value="0" <% if pcPay_Dow_eCheck=0 then%>checked<% end if %>>
                                        No</td>
                                </tr>
                                <!--tr> 
                                    <td> <div align="right">SecureSource:</div></td>
                                    <td> <input type="radio" class="clearBorder" name="x_secureSource" value="1" checked> Yes 
                                        <input name="x_secureSource" type="radio" class="clearBorder" value="0" <% if x_secureSource=0 then%>checked<% end if %>> No (Select Yes if you are a Wells Fargo SecureSource Merchant)</td>
                                </tr-->
                                <tr> 
                                    <td colspan="2">All eCheck orders are automatically considered 
                                        &quot;Processed&quot; by ProductCart, as the order amount 
                                        is always debited to the customer's bank account. If 
                                        for any reasons you would like eCheck orders to be considered 
                                        &quot;Pending&quot;, use this option. Should eCheck 
                                        orders be considered &quot;Pending&quot;? 
                                        <input type="radio" class="clearBorder" name="pcPay_Dow_eCheckPending" value="1" checked> Yes 
                                        <input name="pcPay_Dow_eCheckPending" type="radio" class="clearBorder" value="0" <% if pcPay_Dow_eCheckPending=0 then%>checked<% end if %>> No</td>
                                </tr>
                                <tr> 
                                    <td><div align="right"> 
                                            <input name="pcPay_Dow_TestMode" type="checkbox" class="clearBorder" value="YES" <% if pcPay_Dow_TestMode=1 then%>checked<% end if%>> 
                                        </div></td>
                                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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
                                            <td width="10%" nowrap="nowrap"><div align="left">eCheck&nbsp;&nbsp;Payment Name:&nbsp;</div></td>
                                            <td><input name="paymentNickName2" value="<%=paymentNickName2%>" size="35" maxlength="255"></td>
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
