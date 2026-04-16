<% 

'---Start eProcessingNetwork---
Function gwEPNEdit()
		'request gateway variables and insert them into the EPN table
	query="SELECT pcPay_EPN_Account,pcPay_EPN_RestrictKey FROM pcPay_EPN Where pcPay_EPN_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
	end If
	pcPay_EPN_Account2=rs("pcPay_EPN_Account")
	pcPay_EPN_Account2=enDeCrypt(pcPay_EPN_Account2, scCrypPass)
	pcPay_EPN_RestrictKey2=rs("pcPay_EPN_RestrictKey")
	pcPay_EPN_RestrictKey2=enDeCrypt(pcPay_EPN_RestrictKey2, scCrypPass)
	set rs=nothing
	
	pcPay_EPN_Account=request.Form("pcPay_EPN_Account")
	pcPay_EPN_RestrictKey=request.Form("pcPay_EPN_RestrictKey")
	if pcPay_EPN_Account="" then
		pcPay_EPN_Account=pcPay_EPN_Account2
	end if
	if pcPay_EPN_RestrictKey="" then
		pcPay_EPN_RestrictKey=pcPay_EPN_RestrictKey2
	end if
	pcPay_EPN_CVV=request.Form("pcPay_EPN_CVV")
	pcPay_EPN_TestMode=request.Form("pcPay_EPN_TestMode")
	pcPay_EPN_TranType=request.Form("pcPay_EPN_TranType")
	if pcPay_EPN_TestMode="1" then
		pcPay_EPN_TestMode=1
	else
		pcPay_EPN_TestMode=0
	end if
	pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)
	pcPay_EPN_RestrictKey=enDeCrypt(pcPay_EPN_RestrictKey, scCrypPass)

	query="UPDATE pcPay_EPN SET pcPay_EPN_Account='"&pcPay_EPN_Account&"',pcPay_EPN_CVV="&pcPay_EPN_CVV&",pcPay_EPN_TestMode="&pcPay_EPN_TestMode&",pcPay_EPN_RestrictKey='"&pcPay_EPN_RestrictKey&"',pcPay_EPN_TranType='"&pcPay_EPN_TranType&"' WHERE pcPay_EPN_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=42"
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

Function gwEPN()
	varCheck=1
	'request gateway variables and insert them into the EPN table
	pcPay_EPN_Account=request.Form("pcPay_EPN_Account")
	pcPay_EPN_CVV=request.Form("pcPay_EPN_CVV")
	pcPay_EPN_testmode=request.Form("pcPay_EPN_testmode")
	if pcPay_EPN_testmode="" then
		pcPay_EPN_testmode="0"
	end if
	pcPay_EPN_RestrictKey=request.Form("pcPay_EPN_RestrictKey")

	If pcPay_EPN_Account="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""Account Number""</b> is a required field.")
	End If
	'encrypt
	pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)

	If pcPay_EPN_RestrictKey="" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""RestrictKey""</b> is a required field.")
	End If
	'encrypt
	pcPay_EPN_RestrictKey=enDeCrypt(pcPay_EPN_RestrictKey, scCrypPass)
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
	
	query="UPDATE pcPay_EPN SET pcPay_EPN_Account='"&pcPay_EPN_Account&"',pcPay_EPN_CVV="&pcPay_EPN_CVV&",pcPay_EPN_TestMode=" & pcPay_EPN_TestMode & ",pcPay_EPN_RestrictKey='"&pcPay_EPN_RestrictKey&"' WHERE pcPay_EPN_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'EPN','gwEPN.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",42,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	
end function
%>

<% if request("gwchoice")="42" then
	'//Check if restrict key Field Exists
	on error resume next
	err.clear
		query="SELECT * FROM pcPay_EPN;"
	set rsChkObj = Server.CreateObject("ADODB.Recordset")
	set rsChkObj = conntemp.execute(query)
	chkRestrictKey = rsChkObj("pcPay_EPN_RestrictKey")
	if err.number<>0 then
		set rsChkObj=nothing
		
		call closeDb()
response.redirect "upddbEPN.asp?mode=Edit&id=42"
	else
		set rsChkObj=nothing
		
	end if

	if request("mode")="Edit" then
				query="SELECT pcPay_EPN_Account,pcPay_EPN_CVV,pcPay_EPN_testmode,pcPay_EPN_RestrictKey,pcPay_EPN_TranType FROM pcPay_EPN Where pcPay_EPN_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
		end If
		pcPay_EPN_Account=rs("pcPay_EPN_Account")
		pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)
		pcPay_EPN_CVV=rs("pcPay_EPN_CVV")
		pcPay_EPN_testmode=rs("pcPay_EPN_testmode")
		pcPay_EPN_RestrictKey=rs("pcPay_EPN_RestrictKey")
		pcPay_EPN_TranType=rs("pcPay_EPN_TranType")
		
		pcPay_EPN_RestrictKey=enDeCrypt(pcPay_EPN_RestrictKey, scCrypPass)
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=42"
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
	<input type="hidden" name="addGw" value="42">
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/eprocessingnetwork_logo.png" width="240" height="120"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>eProcessing Network</h4>    
                    <p>Transparent Database Engine Template</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.eprocessingnetwork.com" target="_blank">Learn More</a>        
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
					<% dim pcPay_EPN_AccountCnt,pcPay_EPN_AccountEnd,pcPay_EPN_AccountStart
					pcPay_EPN_AccountCnt=(len(pcPay_EPN_Account)-2)
					pcPay_EPN_AccountEnd=right(pcPay_EPN_Account,2)
					pcPay_EPN_AccountStart=""
					for c=1 to pcPay_EPN_AccountCnt
					pcPay_EPN_AccountStart=pcPay_EPN_AccountStart&"*"
					next %>
					<tr>
						<td height="31" colspan="2">            
            Current Store ID:&nbsp;<%=pcPay_EPN_AccountStart&pcPay_EPN_AccountEnd%></td>
					</tr>
					<tr>
						<td colspan="2"> For security reasons, your &quot;Account Number&quot;
							is only partially shown on this page. If you need
							to edit your account information, please re-enter
							your &quot;Account Number&quot; below.</td>
					</tr>
				<% end if %>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  </tr>
				<tr>
					<td width="111"> <div align="right">Account Number:</div></td>
					<td width="360"> <input type="text" name="pcPay_EPN_Account" size="20"></td>
				</tr>
				<tr>
					<td width="111"> <div align="right">RestrictKey:</div></td>
					<td width="360"> <input type="text" name="pcPay_EPN_RestrictKey" size="20"></td>
				</tr>
				<tr>
					<td width="111"> <div align="right">Require CVV:</div></td>
					<td width="360">
						<input type="radio" class="clearBorder" name="pcPay_EPN_CVV" value="1" checked>
						Yes
						<input name="pcPay_EPN_CVV" type="radio" class="clearBorder" value="0" <% if pcPay_EPN_CVV="0" then %>checked<% end if %> />
						No</td>
				</tr>
					<td width="111"> <div align="right">Transaction Type:</div></td>
					<td width="360">
          	<select name="pcPay_EPN_TranType">
            	<option value="Sale" <% If pcPay_EPN_TranType="Sale" Then Response.Write "selected" %>>Sale</option>
              <option value="AuthOnly" <% If pcPay_EPN_TranType="AuthOnly" Then Response.Write "selected" %>>Authorize Only</option>
            </select>
          </td>
				<tr>
					<td> <div align="right">
							<input name="pcPay_EPN_TestMode" type="checkbox" class="clearBorder" id="pcPay_EPN_TestMode" value="1" <% if pcPay_EPN_testmode=1 then%>checked<% end if%>>
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