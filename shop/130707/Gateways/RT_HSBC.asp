<%
'---Start HSBC---
Function gwHSBCEdit()
		'request gateway variables and insert them into the HSBC table
	query="SELECT pcPay_HSBC_UserID,pcPay_HSBC_Password,pcPay_HSBC_ClientID FROM pcPay_HSBC Where pcPay_HSBC_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_HSBC_UserID2=rs("pcPay_HSBC_UserID")
	pcPay_HSBC_Password2=rs("pcPay_HSBC_Password")
	pcPay_HSBC_ClientID2=rs("pcPay_HSBC_ClientID")
	pcPay_HSBC_Password2=enDeCrypt(pcPay_HSBC_Password2, scCrypPass)
	
	pcPay_HSBC_UserID=request.Form("pcPay_HSBC_UserID")
	if pcPay_HSBC_UserID="" then
		pcPay_HSBC_UserID=pcPay_HSBC_UserID2
	end if

	pcPay_HSBC_Password=request.Form("pcPay_HSBC_Password")
	if pcPay_HSBC_Password="" then
		pcPay_HSBC_Password=pcPay_HSBC_Password2
	end if
	set rs=nothing
	pcPay_HSBC_ClientID=request.Form("pcPay_HSBC_ClientID")
	if pcPay_HSBC_ClientID="" then
		pcPay_HSBC_ClientID=pcPay_HSBC_ClientID2
	end if
	pcPay_HSBC_TransType=request.Form("pcPay_HSBC_TransType")
	pcPay_HSBC_CVV=request.Form("pcPay_HSBC_CVV")
	if pcPay_HSBC_CVV="" then
		pcPay_HSBC_CVV=0
	end if
	pcPay_HSBC_Currency=request.Form("pcPay_HSBC_Currency")
	if pcPay_HSBC_Currency="" then
		pcPay_HSBC_Currency="978"
	end if
	pcPay_HSBC_TestMode=request.Form("pcPay_HSBC_TestMode")
	if pcPay_HSBC_TestMode="" then
		pcPay_HSBC_TestMode=0
	end if

	pcPay_HSBC_Password=enDeCrypt(pcPay_HSBC_Password, scCrypPass)
	
	query="UPDATE pcPay_HSBC SET pcPay_HSBC_UserID='"&pcPay_HSBC_UserID&"',pcPay_HSBC_Password='"&pcPay_HSBC_Password&"',pcPay_HSBC_ClientID='"&pcPay_HSBC_ClientID&"',pcPay_HSBC_TransType='"&pcPay_HSBC_TransType&"',pcPay_HSBC_CVV="&pcPay_HSBC_CVV&",pcPay_HSBC_Currency='" & pcPay_HSBC_Currency & "',pcPay_HSBC_TestMode=" & pcPay_HSBC_TestMode & " WHERE pcPay_HSBC_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=44"
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

Function gwHSBC()
	varCheck=1
	'request gateway variables and insert them into the HSBC table
	pcPay_HSBC_UserId=request.Form("pcPay_HSBC_UserId")
	pcPay_HSBC_Password=request.Form("pcPay_HSBC_Password")
	pcPay_HSBC_ClientID=request.Form("pcPay_HSBC_ClientID")
	pcPay_HSBC_TransType=request.Form("pcPay_HSBC_TransType")
	pcPay_HSBC_Currency=request.Form("pcPay_HSBC_Currency")
	if pcPay_HSBC_Currency="" then
		pcPay_HSBC_Currency="826"
	end if
	pcPay_HSBC_CVV=request.Form("pcPay_HSBC_Cvv")
	if pcPay_HSBC_CVV="" then
		pcPay_HSBC_CVV=0
	end if
	pcPay_HSBC_TestMode=request.Form("pcPay_HSBC_TestMode")
	if pcPay_HSBC_TestMode="" then
		pcPay_HSBC_TestMode=0
	end if

	If pcPay_HSBC_UserID="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add HSBC as your payment gateway. <b>""User ID""</b> is a required field.")
	End If
	If pcPay_HSBC_Password="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add HSBC as your payment gateway. <b>""Password""</b> is a required field.")
	End If
	If pcPay_HSBC_ClientID="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add HSBC as your payment gateway. <b>""Client ID""</b> is a required field.")
	End If

	'encrypt
	pcPay_HSBC_Password=enDeCrypt(pcPay_HSBC_Password, scCrypPass)
	
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
	 

	query="UPDATE pcPay_HSBC SET pcPay_HSBC_UserID='"&pcPay_HSBC_UserID&"',pcPay_HSBC_Password='"&pcPay_HSBC_Password&"',pcPay_HSBC_ClientID='"&pcPay_HSBC_ClientID&"',pcPay_HSBC_TransType='"&pcPay_HSBC_TransType&"',pcPay_HSBC_CVV="&pcPay_HSBC_CVV&",pcPay_HSBC_Currency='" & pcPay_HSBC_Currency & "',pcPay_HSBC_TestMode=" & pcPay_HSBC_TestMode & " WHERE pcPay_HSBC_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'HSBC','gwHSBC.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",44,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
			
	
end function

if request("gwchoice")="44" then
	if request("mode")="Edit" then
				query="SELECT pcPay_HSBC_UserId,pcPay_HSBC_Password,pcPay_HSBC_ClientId,pcPay_HSBC_TransType,pcPay_HSBC_CVV,pcPay_HSBC_Currency,pcPay_HSBC_TestMode FROM pcPay_HSBC Where pcPay_HSBC_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_HSBC_UserID=rs("pcPay_HSBC_UserId")
		pcPay_HSBC_Password=rs("pcPay_HSBC_Password")
		pcPay_HSBC_Password=enDeCrypt(pcPay_HSBC_Password, scCrypPass)
		pcPay_HSBC_ClientID=rs("pcPay_HSBC_ClientId")
		pcPay_HSBC_TransType=rs("pcPay_HSBC_TransType")
		pcPay_HSBC_CVV=rs("pcPay_HSBC_CVV")
		pcPay_HSBC_Currency=rs("pcPay_HSBC_Currency")
		pcPay_HSBC_TestMode=rs("pcPay_HSBC_TestMode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=44"
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
	<input type="hidden" name="addGw" value="44">
<!-- END HSBC -->

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/hsbc_globalpayments_logo.JPG" width="198" height="72"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>HSBC Global Payments</h4>    
                    <p></p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.globalpaymentsinc.com/USA/aboutUs/companyOverview.html" target="_blank">Learn More</a>        
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
					<% dim pcPay_HSBC_UserIDCnt,pcPay_HSBC_UserIDEnd,pcPay_HSBC_UserIDStart
                    pcPay_HSBC_UserIDCnt=(len(pcPay_HSBC_UserID)-2)
                    pcPay_HSBC_UserIDEnd=right(pcPay_HSBC_UserID,2)
                    pcPay_HSBC_UserIDStart=""
                    for c=1 to pcPay_HSBC_UserIDCnt
                    pcPay_HSBC_UserIDStart=pcPay_HSBC_UserIDStart&"*"
                    next
                    %>
                    <tr> 
                        <td height="31">&nbsp;</td>
                        <td height="31">Current User ID:&nbsp;<%=pcPay_HSBC_UserIDStart&pcPay_HSBC_UserIDEnd%></td>
                    </tr>
                    <tr> 
                        <td>&nbsp;</td>
                        <td> For security reasons, your &quot;User ID&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;User ID&quot;, &quot;Password&quot; and 
                            &quot;Client ID&quot; below.</td>
                    </tr>
				<% end if %>
                <tr>
                    <td width="111"> <div align="right">User ID:</div></td>
                    <td width="360"> <input type="text" name="pcPay_HSBC_UserID" size="20"></td>
                </tr>
                <tr>
                    <td width="111"> <div align="right">Password:</div></td>
                    <td width="360"> 
                    <input type="password" name="pcPay_HSBC_Password" size="20"></td>
                </tr>
                <tr>
                    <td width="111"> <div align="right">Client ID:</div></td>
                    <td width="360"> <input type="text" name="pcPay_HSBC_ClientID" size="20"></td>
                </tr>
                <tr>
                    <td width="111"> <div align="right">Transaction Type:</div></td>
                    <td width="360"> <select name="pcPay_HSBC_TransType">
                            <option value="Auth" <%if pcPay_HSBC_TransType="Auth" then%>selected<%end if %>>Sale</option>
                            <option value="PreAuth" <%if pcPay_HSBC_TransType="PreAuth" then%>selected<%end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
                <tr> 
                    <td width="111"> <div align="right">Store Currency:</div></td>
                    <td width="360">
                        <select name="pcPay_HSBC_Currency">
                            <option value="978" selected>Euro</option>
                            <option value="344" <%if pcPay_HSBC_Currency="344" then%>selected<%end if%>>Hong Kong Dollar</option>
                            <option value="392" <%if pcPay_HSBC_Currency="392" then%>selected<%end if%>>Japanese Yen</option>
                            <option value="826" <%if pcPay_HSBC_Currency="826" then%>selected<%end if%>>Pound Sterling</option>
                            <option value="840" <%if pcPay_HSBC_Currency="840" then%>selected<%end if%>>US Dollar</option>
                        </select></td>
                </tr>
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_HSBC_Cvv" value="1" checked>
                        Yes 
                        <input name="pcPay_HSBC_Cvv" type="radio" class="clearBorder" value="0" <%if pcPay_HSBC_CVV="0" then%>checked<%end if%>>
                        No</td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="pcPay_HSBC_TestMode" type="checkbox" class="clearBorder" id="pcPay_HSBC_TestMode" value="1" <% if pcPay_HSBC_Testmode="1" then %>checked<% end if %>>
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
                    <td width="18%" nowrap><span class="pcSubmenuHeader">Processing fee:</span><br /></td>
                            <td width="82%">
                              <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>"></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                            <td><input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                                Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                                <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>"></td>
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
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapse3">
                                    Step 3: You can change the display name that is shown for this payment type.
                                </a>
                            </h4>
                        </div>
                        <div id="collapse3" class="panel-collapse collapse">
                            <div class="panel-body">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="18%"><div align="left"><strong>Payment Name:&nbsp;</strong></div></td>
                                <td width="82%"><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
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
                                <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=301"></a></td>
                    </tr>
                    <tr> 
                        <td>&nbsp;</td>
                        <td>When orders are placed, set the payment status to:
                        <select name="pcv_setPayStatus">
                            <option value="3" selected="selected">Default</option>
                                    <option value="0" <%if pcv_setPayStatus="0" then%>selected<%end if%>>Pending</option>
                                    <option value="1" <%if pcv_setPayStatus="1" then%>selected<%end if%>>Authorized</option>
                                    <option value="2" <%if pcv_setPayStatus="2" then%>selected<%end if%>>Paid</option>
                        </select>
                        &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=302"></a>					</td>
                    </tr>
				</table>
                            </div>
                        </div> 
                    </div>
                    
                </div>
                
                
                
                
			    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td colspan="2">
                            <input type="submit" value="Add Selected Options" name="Submit" class="btn btn-primary"> 
		                        &nbsp;
		                    <input type="button" class="btn btn-default"  value="Back" onclick="javascript:history.back()"></td>
                        </td>
                    </tr>
                </table>


            </td>
        </tr>
    </table>
<!-- New View End --><% end if %>
