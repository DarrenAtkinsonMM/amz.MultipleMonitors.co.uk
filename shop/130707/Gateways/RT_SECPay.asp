<%
'---Start SecPay---
Function gwSECPayEdit()
		'request gateway variables and insert them into the SECPay table
	query= "SELECT pcPay_SecPay_Username,pcPay_SecPay_TestMode,pcPay_SecPay_TransType,pcPay_SecPay_Password FROM pcPay_SecPay WHERE pcPay_SecPay_ID=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_SecPay_Username2=rstemp("pcPay_SecPay_Username")
	pcPay_SecPay_Username=request.Form("pcPay_SecPay_Username")
	if pcPay_SecPay_Username="" then
		pcPay_SecPay_Username=pcPay_SecPay_Username2
	end if
	
	pcPay_SecPay_TestMode=request.Form("pcPay_SecPay_TestMode")
	if pcPay_SecPay_TestMode="" then
		pcPay_SecPay_TestMode=0
	end if
	
	pcPay_SecPay_TransType2=rstemp("pcPay_SecPay_TransType")
	pcPay_SecPay_TransType=request.Form("pcPay_SecPay_TransType")
	If pcPay_SecPay_TransType="" then
		pcPay_SecPay_TransType=pcPay_SecPay_TransType2
	end if
	
	pcPay_SecPay_Password2=rstemp("pcPay_SecPay_Password")
	pcPay_SecPay_Password=request.Form("pcPay_SecPay_Password")
	if pcPay_SecPay_Password="" then
		pcPay_SecPay_Password=pcPay_SecPay_Password2
	end if
	
	pcPay_SecPay_Cvc2=request.Form("pcPay_SecPay_Cvc2")
	
	query="UPDATE pcPay_SecPay SET pcPay_SecPay_Username='"&pcPay_SecPay_Username&"',pcPay_SecPay_TestMode="&pcPay_SecPay_TestMode&",pcPay_SecPay_TransType='"&pcPay_SecPay_TransType&"',pcPay_SecPay_Password='"&pcPay_SecPay_Password&"',pcPay_SecPay_Cvc2='"&pcPay_SecPay_Cvc2&"' WHERE pcPay_SecPay_ID=1;"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=48"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwSECPay()
	varCheck=1
	'request gateway variables and insert them into the SecPay table
	pcPay_SecPay_Username=request.Form("pcPay_SecPay_Username")
	pcPay_SecPay_TestMode=request.Form("pcPay_SecPay_TestMode")
	if pcPay_SecPay_TestMode="" then
		pcPay_SecPay_TestMode=0
	end if
	pcPay_SecPay_TransType=request.Form("pcPay_SecPay_TransType")
	pcPay_SecPay_Password=request.Form("pcPay_SecPay_Password")
	pcPay_SecPay_Cvc2=request.Form("pcPay_SecPay_Cvc2")
	
	If pcPay_SecPay_Username="" then
		
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add SecPay as your payment gateway. <b>""SECPay Username""</b> is a required field.")
	End If
	
	If pcPay_SecPay_Password="" then
		
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add SecPay as your payment gateway. <b>""VPN Password""</b> is a required field.")
	End If
	
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
	 

	query="UPDATE pcPay_SecPay SET pcPay_SecPay_Username='"&pcPay_SecPay_Username&"', pcPay_SecPay_TestMode="&pcPay_SecPay_TestMode&", pcPay_SecPay_Password='"&pcPay_SecPay_Password&"', pcPay_SecPay_TransType='"&pcPay_SecPay_TransType&"',pcPay_SecPay_Cvc2='"&pcPay_SecPay_Cvc2&"' WHERE pcPay_SecPay_id=1;"
				
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'SecPay','gwSecPay.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",48,N'"&paymentNickName&"')"
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

if request("gwchoice")="48" then
	if request("mode")="Edit" then
				query= "SELECT pcPay_SecPay_Username,pcPay_SecPay_TestMode,pcPay_SecPay_TransType,pcPay_SecPay_Password,pcPay_SecPay_Cvc2 FROM pcPay_SecPay WHERE pcPay_SecPay_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_SecPay_Username=rs("pcPay_SecPay_Username")
		pcPay_SecPay_TestMode=rs("pcPay_SecPay_TestMode")
		pcPay_SecPay_TransType=rs("pcPay_SecPay_TransType")
		pcPay_SecPay_Password=rs("pcPay_SecPay_Password")
		pcPay_SecPay_Cvc2=rs("pcPay_SecPay_Cvc2")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=48"
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
	<input type="hidden" name="addGw" value="48">

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/paypoint_logo.gif" width="248" height="63"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>Gateway Freedom Integration</h4>    
                    <p>SECPay is now PayPoint.net, one of the UK's leading providers of payment gateway solutions. Whether you are starting out, a small business or a large enterprise, PayPoint.net has a payment gateway to suit your business needs.</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.paypoint.net/secpay-payment-gateway" target="_blank">Learn More</a>        
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
					dim pcPay_SecPay_UsernameCnt,pcPay_SecPay_UsernameEnd,pcPay_SecPay_UsernameStart
					pcPay_SecPay_UsernameCnt=(len(pcPay_SecPay_Username)-2)
					pcPay_SecPay_UsernameEnd=right(pcPay_SecPay_Username,2)
					pcPay_SecPay_UsernameStart=""
					for c=1 to pcPay_SecPay_UsernameCnt
						pcPay_SecPay_UsernameStart=pcPay_SecPay_UsernameStart&"*"
					next %>
					<tr> 
                                <td height="31" colspan="2">Current SECPay Merchant Code:&nbsp;<%=pcPay_SecPay_UsernameStart&pcPay_SecPay_UsernameEnd%></td>
                            </tr>
                            <tr> 
                                <td colspan="2"> For security reasons, your &quot;SECPay Account Username&quot; is only partially shown on this page. If 
                                    you need to edit your account information, please 
                                    re-enter your &quot;Account Username&quot; below.</td>
                            </tr>
                        <% end if %>
                        <tr> 
                            <td width="166"> <div align="right">Account Name :</div></td>
                            <td width="483"> <input type="text" name="pcPay_SecPay_Username" size="30"></td>
                        </tr>
                        <tr> 
                            <td> <div align="right">VPN Password:</div></td>
                            <td> <input type="text" name="pcPay_SecPay_Password" size="30"></td>
                        </tr>
                        <tr> 
                            <td> <div align="right">Transaction Type:</div></td>
                            <td> <select name="pcPay_SecPay_TransType">
                                    <option value="SALE" <% if pcPay_SecPay_TransType="SALE" then %>Selected<%end if%>>Sale</option>
                                    <option value="AUTH" <% if pcPay_SecPay_TransType="AUTH" then %>Selected<%end if%>>Authorize Only</option>
                                </select></td>
                        </tr>
                        <tr>
            <td><div align="right">Require CVV:</div></td>
                          <td><input type="radio" class="clearBorder" name="pcPay_SecPay_Cvc2" value="1" checked>
                            Yes
                            <input name="pcPay_SecPay_Cvc2" type="radio" class="clearBorder" value="0" <%if clng(pcPay_SecPay_Cvc2)=0 then%>checked<%end if%>>
                            No</td>
                          </tr>
                        <tr> 
                            <td> <div align="right"> 
                                    <input name="pcPay_SecPay_TestMode" type="checkbox" class="clearBorder" value="1" <% if pcPay_SecPay_TestMode=1 then %>checked<% end if %> />
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