<%
'---Start Ogone---
Function gwOgoneEdit()
		'request gateway variables and insert them into the ogone table
query="SELECT pcPay_OG_MerchantID,pcPay_OG_MerchantPassword,pcPay_OG_TransType,pcPay_OG_Lang,pcPay_OG_CurCode,pcPay_OG_cardTypes, pcPay_OG_CVC,pcPay_OG_AccountID,pcPay_OG_TestMode FROM pcPay_Ogone Where pcPay_OG_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_OG_MerchantID2=rs("pcPay_OG_MerchantID")
	pcPay_OG_MerchantID2=enDeCrypt(pcPay_OG_MerchantID2, scCrypPass)

	pcPay_OG_MerchantID=request.Form("pcPay_OG_MerchantID")
	if pcPay_OG_MerchantID="" then
		pcPay_OG_MerchantID=pcPay_OG_MerchantID2
	end if
	pcPay_OG_MerchantPassword2=rs("pcPay_OG_MerchantPassword")
	pcPay_OG_MerchantPassword2=enDeCrypt(pcPay_OG_MerchantPassword2, scCrypPass)
	pcPay_OG_MerchantPassword=request.Form("pcPay_OG_MerchantPassword")
	if pcPay_OG_MerchantPassword="" then
		pcPay_OG_MerchantPassword=pcPay_OG_MerchantPassword2
	end if
	
	
	pcPay_OG_AccountID2= rs("pcPay_OG_AccountID")
	pcPay_OG_AccountID2=enDeCrypt(pcPay_OG_AccountID2, scCrypPass)
	pcPay_OG_AccountID=request.Form("pcPay_OG_AccountID")
	if pcPay_OG_AccountID="" then
		pcPay_OG_AccountID=pcPay_OG_AccountID2
	end if
	set rs=nothing
	
	pcPay_OG_TransType = request.form("pcPay_OG_TransType")
	pcPay_OG_Lang = request.form("pcPay_OG_Lang")
	pcPay_OG_CurCode = request.form("pcPay_OG_CurCode")		
	pcPay_OG_cardTypes=request.Form("pcPay_OG_cardTypes")
	pcPay_OG_CVC= request.Form("pcPay_OG_CVC")
	pcPay_OG_TestMode=request.Form("pcPay_OG_TestMode")
	if pcPay_OG_TestMode="" then
		pcPay_OG_TestMode=0
	end if

	pcPay_OG_MerchantID=enDeCrypt(pcPay_OG_MerchantID, scCrypPass)
	pcPay_OG_MerchantPassword=enDeCrypt(pcPay_OG_MerchantPassword, scCrypPass)
	pcPay_OG_AccountID=enDeCrypt(pcPay_OG_AccountID, scCrypPass)
	
	query="UPDATE pcPay_Ogone SET pcPay_OG_MerchantID='"&pcPay_OG_MerchantID&"',pcPay_OG_MerchantPassword='"&pcPay_OG_MerchantPassword&"',pcPay_OG_TransType ='" &pcPay_OG_TransType &"',pcPay_OG_Lang ='" & pcPay_OG_Lang &"', pcPay_OG_CurCode ='" & pcPay_OG_CurCode &"',pcPay_OG_cardTypes ='"&pcPay_OG_cardTypes&"',pcPay_OG_CVC=" & pcPay_OG_CVC &",pcPay_OG_AccountID='" & pcPay_OG_AccountID&"',pcPay_OG_TestMode=" & pcPay_OG_TestMode & " WHERE pcPay_OG_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=55"
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

Function gwOgone()
	varCheck=1
	'request gateway variables and insert them into the Ogone table
	pcPay_OG_MerchantID=request.Form("pcPay_OG_MerchantID")
	pcPay_OG_AccountID = request.form("pcPay_OG_AccountID")
	pcPay_OG_MerchantPassword=request.Form("pcPay_OG_MerchantPassword")
	pcPay_OG_TransType = request.form("pcPay_OG_TransType")
	pcPay_OG_Lang = request.form("pcPay_OG_Lang")
	pcPay_OG_CurCode = request.form("pcPay_OG_CurCode")			
	pcPay_OG_cardTypes=request.Form("pcPay_OG_cardTypes")
	pcPay_OG_CVC= request.Form("pcPay_OG_CVC")
	pcPay_OG_Signature= request.Form("pcPay_OG_Signature")
	pcPay_OG_TestMode=request.Form("pcPay_OG_TestMode")
	if pcPay_OG_TestMode="" then
		pcPay_OG_TestMode=0
	end if

	If pcPay_OG_MerchantID="" OR pcPay_OG_MerchantPassword="" or pcPay_OG_AccountID ="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""Merchant PSID""</b>, <b>""Account API ID""</b> and <b>""Account API Password""</b> are required fields.")
	End If
	'encrypt
	pcPay_OG_MerchantID=enDeCrypt(pcPay_OG_MerchantID, scCrypPass)
	pcPay_OG_AccountID=enDeCrypt(pcPay_OG_AccountID, scCrypPass)
	pcPay_OG_MerchantPassword=enDeCrypt(pcPay_OG_MerchantPassword, scCrypPass)
	
	priceToAddType=request.Form("priceToAddType")
	if priceToAddType="price" then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
	else
		priceToAdd="0"
		percentageToAdd=request.Form("percentageToAdd")
	end if
	If priceToAdd="" Then
		priceToAdd="0"
	end if
	If percentageToAdd="" Then
		percentageToAdd="0"
	end if
	
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")
	if paymentNickName="" then
		paymentNickName="Credit Card"
	End If
	
	err.clear
	err.number=0
	 

	query="UPDATE pcPay_Ogone SET pcPay_OG_MerchantID='"&pcPay_OG_MerchantID&"',pcPay_OG_MerchantPassword='"&pcPay_OG_MerchantPassword&"',pcPay_OG_TransType ='" &pcPay_OG_TransType &"',pcPay_OG_Lang ='" & pcPay_OG_Lang &"', pcPay_OG_CurCode ='" & pcPay_OG_CurCode &"',pcPay_OG_cardTypes ='"&pcPay_OG_cardTypes&"',pcPay_OG_CVC=" & pcPay_OG_CVC &",pcPay_OG_AccountID='" & pcPay_OG_AccountID&"',pcPay_OG_TestMode=" & pcPay_OG_TestMode & " WHERE pcPay_OG_ID=1;"
	'Response.write query 
	'Response.end
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Ogone','gwOgone.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",55,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
			
    
end function

if request("gwchoice")="55" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
		query="select * from pcPay_Ogone"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		
		call closeDb()
response.redirect "upddbOgone.asp"
	else
		set rs=nothing
		
	end if
	if request("mode")="Edit" then
	 		
		query="SELECT pcPay_OG_MerchantID,pcPay_OG_MerchantPassword,pcPay_OG_TransType,pcPay_OG_Lang,pcPay_OG_CurCode,pcPay_OG_cardTypes, pcPay_OG_CVC,pcPay_OG_AccountID,pcPay_OG_TestMode FROM pcPay_Ogone Where pcPay_OG_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_OG_MerchantID=rs("pcPay_OG_MerchantID")
		pcPay_OG_MerchantID=enDeCrypt(pcPay_OG_MerchantID, scCrypPass)
		pcPay_OG_MerchantPassword=rs("pcPay_OG_MerchantPassword")
		pcPay_OG_MerchantPassword=enDeCrypt(pcPay_OG_MerchantPassword, scCrypPass)
		pcPay_OG_TransType = rs("pcPay_OG_TransType")
		pcPay_OG_Lang = rs("pcPay_OG_Lang")
		pcPay_OG_CurCode = rs("pcPay_OG_CurCode")
		pcPay_OG_cardTypes=rs("pcPay_OG_cardTypes")
		pcPay_OG_CVC=rs("pcPay_OG_CVC")
		pcPay_OG_AccountID = rs("pcPay_OG_AccountID")
		pcPay_OG_AccountID=enDeCrypt(pcPay_OG_AccountID, scCrypPass)
		pcPay_OG_TestMode=rs("pcPay_OG_TestMode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=55"
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
	<input type="hidden" name="addGw" value="55">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/ogone_logo.jpg" width="250" height="94"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>Ogone</h4>    
                    <p>Ogone is a European Payment Service Provider that enables you to manage, control and in most cases eliminate the exposure to fraud in the card-not-present environment.</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.ogone.com" target="_blank">Learn More</a>        
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
					dim pcPay_OG_MerchantIDCnt,pcPay_OG_MerchantIDEnd,pcPay_OG_MerchantIDStart
					pcPay_OG_MerchantIDCnt=(len(pcPay_OG_MerchantID)-2)
					pcPay_OG_MerchantIDEnd=right(pcPay_OG_MerchantID,2)
					pcPay_OG_MerchantIDStart=""
					for c=1 to pcPay_OG_MerchantIDCnt
					pcPay_OG_MerchantIDStart=pcPay_OG_MerchantIDStart&"*"
					next
					
					dim pcPay_OG_AccountIDCnt,pcPay_OG_AccountIDEnd,pcPay_OG_AccountIDStart
					pcPay_OG_AccountIDCnt=(len(pcPay_OG_AccountID)-2)
					pcPay_OG_AccountIDEnd=right(pcPay_OG_AccountID,2)
					pcPay_OG_AccountIDStart=""
					for c=1 to pcPay_OG_AccountIDCnt
					pcPay_OG_AccountIDStart=pcPay_OG_AccountIDStart&"*"
					next
					%>
					<tr> 
						<td height="31">&nbsp;</td>
						<td height="31">Current Merchant PSID:&nbsp;<%=pcPay_OG_MerchantIDStart&pcPay_OG_MerchantIDEnd%></td>
					</tr>
					<tr> 
						<td height="31">&nbsp;</td>
						<td height="31">Current Account ID:&nbsp;<%=pcPay_OG_AccountIDStart&pcPay_OG_AccountIDEnd%></td>
					</tr>
					<tr> 
						<td>&nbsp;</td>
						<td> For security reasons, your &quot;Account ID and your Merchant PSID&quot; 
							is only partially shown on this page. If you need 
							to edit your account information, please re-enter 
							your &quot;Merchant PSID&quot;,&quot;Account ID&quot; and &quot;Account 
							Password&quot; below.</td>
					</tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td width="111"> <div align="right">Merchant&nbsp;PSID:</div></td>
                  <td width="328"> <input type="text" name="pcPay_OG_MerchantID" size="20"> </td>
                </tr>
                <tr>
                    <td width="111"> <div align="right">Account API ID:</div></td>
                    <td width="328"> <input type="text" name="pcPay_OG_AccountID" size="20"></td>
                </tr>
                <tr>
                  <td width="111"> <div align="right">Account API Password:</div></td>
                  <td width="328"> <input type="text" name="pcPay_OG_MerchantPassword" size="20">								</td>
                </tr>
                <tr> 
                    <td colspan="2"><div align="left">
                      <p>Ogone - Direct Link Account Setup Notes:</p>
                      <ul>
                        <li>Create a new user from the Ogone control panel and select the &quot;Special user for API (no access to the admin) check box.&quot; This will be your &quot;Account API ID&quot; and Ogone will assign you your &quot;Account API Password&quot;</li>
                        <li> Technical Information 1.1: Please check &quot;Post&quot; not &quot;Get&quot;</li>
                        <li>Technical Information 2.1: Please enter the  IP address of the server that will be hosting ProductCart</li>
                        <li>Technical Information 3.1 and 3.2: Please Check &quot;NO&quot;</li>
                      </ul>
                    </div></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="pcPay_OG_TransType">
                            <option value="SAL" <% if pcPay_OG_TransType="SAL" then%>selected<% end if %>>Sale</option>
                            <option value="REF" <% if pcPay_OG_TransType="REF" then%>selected<% end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
                    <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td>
                    	<input name="pcPay_OG_CardTypes" type="checkbox" class="clearBorder" value="VISA" <%if pcPay_OG_CardTypes="VISA" or (instr(pcPay_OG_CardTypes,"VISA,")>0) or (instr(pcPay_OG_CardTypes,", VISA")>0) then%>checked<%end if%>>
                        Visa 
                        <input name="pcPay_OG_CardTypes" type="checkbox" class="clearBorder" value="MAST" <%if pcPay_OG_CardTypes="MAST" or (instr(pcPay_OG_CardTypes,"MAST,")>0) or (instr(pcPay_OG_CardTypes,", MAST")>0) then%>checked<%end if%>>
                        MasterCard 
                        <input name="pcPay_OG_CardTypes" type="checkbox" class="clearBorder" value="AMER" <%if pcPay_OG_CardTypes="AMER" or (instr(pcPay_OG_CardTypes,"AMER,")>0) or (instr(pcPay_OG_CardTypes,", AMER")>0) then%>checked<%end if%>>
                        American Express 
                        <input name="pcPay_OG_CardTypes" type="checkbox" class="clearBorder" value="DINE" <%if pcPay_OG_CardTypes="DINE" or (instr(pcPay_OG_CardTypes,"DINE,")>0) or (instr(pcPay_OG_CardTypes,", DINE")>0) then%>checked<%end if%>>
                        Diner's Club 
                        <input name="pcPay_OG_CardTypes" type="checkbox" class="clearBorder" value="JCB" <%if pcPay_OG_CardTypes="JCB" or (instr(pcPay_OG_CardTypes,"JCB,")>0) or (instr(pcPay_OG_CardTypes,", JCB")>0) then%>checked<%end if%>>
                        JCB
                        <input name="pcPay_OG_cardTypes" type="checkbox" class="clearBorder" value="AURORA" <%if pcPay_OG_CardTypes="AURORA" or (instr(pcPay_OG_CardTypes,"AURORA,")>0) or (instr(pcPay_OG_CardTypes,", AURORA")>0) then%>checked<%end if%> >
                        Aurora
                        <input name="pcPay_OG_cardTypes" type="checkbox" class="clearBorder" value="AURORE" <%if pcPay_OG_CardTypes="AURORE" or (instr(pcPay_OG_CardTypes,"AURORE,")>0) or (instr(pcPay_OG_CardTypes,", AURORE")>0) then%>checked<%end if%> >
                        Aurore 
                        <input name="pcPay_OG_cardTypes" type="checkbox" class="clearBorder" value="SOLO" <%if pcPay_OG_CardTypes="SOLO" or (instr(pcPay_OG_CardTypes,"SOLO,")>0) or (instr(pcPay_OG_cardTypes,", SOLO")>0) then%>checked<%end if%> >
                        Solo
                        <input name="pcPay_OG_cardTypes" type="checkbox" class="clearBorder" value="MaestroUK" <%if pcPay_OG_CardTypes="MaestroUK" or (instr(pcPay_OG_CardTypes,"MaestroUK,")>0) or (instr(pcPay_OG_cardTypes,", MaestroUK")>0) then%>checked<%end if%> >
                        MaestroUK
                        <input name="pcPay_OG_cardTypes" type="checkbox" class="clearBorder" value="EUROCARD" <%if pcPay_OG_CardTypes="EUROCARD" or (instr(pcPay_OG_CardTypes,"EUROCARD,")>0) or (instr(pcPay_OG_cardTypes,"EUROCARD")>0) then%>checked<%end if%> >
                        Euro Card </td>
                </tr>
                 <tr> 
                    <td> <div align="right">Languages:</div></td>
                    <td> <select name="pcPay_OG_Lang">
                            <option value="es_US" selected>English</option>
                            <option value="nl_NL" <% if pcPay_OG_Lang="nl_NL" then%>selected<% end if %>>Dutch</option>
                            <option value="fr_FR" <% if pcPay_OG_Lang="fr_FR" then%>selected<% end if %>>French</option>                                      
                            <option value="de_DE" <% if pcPay_OG_Lang="de_DE" then%>selected<% end if %>>German</option>
                            <option value="no_No" <% if pcPay_OG_Lang="no_No" then%>selected<% end if %>>Norwegian</option>
                            <option value="es_ES" <% if pcPay_OG_Lang="es_ES" then%>selected<% end if %>>Spanish</option>
                        </select> </td>
                </tr>
                <% if pcPay_OG_CurCode="" then
					pcPay_OG_CurCode="USD"
				end if %>
                <tr> 
                    <td><div align="right">Currency Code:</div></td>
                    <td><input name="pcPay_OG_CurCode" type="text" value="<%=pcPay_OG_CurCode%>" size="6" maxlength="4"> 
                        <a href="help_auth_codes.asp" target="_blank">Find Codes</a></td>
                </tr>
                    <tr> 
                        <td> <div align="right">Require CVC:</div></td>
                        <td> <input type="radio" class="clearBorder" name="pcPay_OG_CVC" value="1" checked> Yes 
                            <input name="pcPay_OG_CVC" type="radio" class="clearBorder" value="0" <% if pcPay_OG_CVC=0 then%>checked<% end if %>> No<font color="#FF0000"></font></td>
                    </tr>
                    <tr> 
                    <td> <div align="right"> 
                            <input name="pcPay_OG_TestMode" type="checkbox" class="clearBorder" id="pcPay_TD_TestMode" value="1" <% if pcPay_OG_TestMode=1 then%>checked<%end if %>>
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
