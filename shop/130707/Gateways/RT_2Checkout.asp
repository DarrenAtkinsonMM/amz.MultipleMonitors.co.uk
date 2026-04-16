<% 
'---Start 2Checkout---
Function gw2CheckoutEdit()
		'request gateway variables and insert them into the twoCheckout table
	query= "SELECT store_id FROM TwoCheckout WHERE id_twocheckout=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	store_id2=rstemp("store_id")
	store_id=request.Form("twoCheckout_store_id")
	v2co=request.Form("v2co")
	if (v2co="") then
		v2co="2"
	end if
	str2Checkout="gw2checkout2.asp"
	v2co_testmode=request.Form("v2co_testmode")
	if v2co_testmode="YES" then
		v2co_testmode=1
	else
		v2co_testmode=0
	end if
	if store_id="" then
		store_id=store_id2
	end if
	query="UPDATE twoCheckout SET store_id='"&store_id&"', v2co="&v2co&", v2co_testmode="&v2co_testmode&" WHERE id_twoCheckout=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&str2Checkout&"', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=13"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	
end function

Function gw2Checkout()
	varCheck=1
	'request gateway variables and insert them into the 2Checkout table
	store_id=request.form("twoCheckout_store_id")
	v2co=request.Form("v2co")
	if (v2co="") then
		v2co="2"
	end if
	str2Checkout="gw2Checkout2.asp"
	v2co_TestMode=request.Form("v2co_TestMode")
	if v2co_TestMode="YES" then
		v2co_TestMode=1
	else
		v2co_TestMode=0
	end if
		
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
	End If
	
	err.clear
	err.number=0
	 

	query="UPDATE twocheckout SET store_id='"&store_id&"', v2co="&v2co&", v2co_TestMode="&v2co_TestMode&" WHERE id_twocheckout=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'2Checkout','"&str2Checkout&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",13,N'"&paymentNickName&"')"
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
				
<% if request("gwchoice")="13" then 
	if request("mode")="Edit" then
		    	query= "SELECT store_id, v2co, v2co_testmode FROM TwoCheckout WHERE id_twocheckout=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		twoCheckout_store_id=rs("store_id")
		intv2co_testmode=rs("v2co_testmode")
		v2co=rs("v2co")
		set rs=nothing
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=13"
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

		
		
		dim twoCheckout_store_idCnt,twoCheckout_store_idEnd,twoCheckout_store_idStart
		twoCheckout_store_idCnt=(len(twoCheckout_store_id)-2)
		twoCheckout_store_idEnd=right(twoCheckout_store_id,2)
		twoCheckout_store_idStart=""
		for c=1 to twoCheckout_store_idCnt
			twoCheckout_store_idStart=twoCheckout_store_idStart&"*"
		next %>
		<input type="hidden" name="mode" value="Edit">
	<% end if %>
    <input type="hidden" name="addGw" value="13">
    <table width="100%">
        <tr>
            <td width="10%" align="left" style="font-size:15px;"><img src="Gateways/logos/2CO.png" width="210" height="80"></td>
            <td width="90%" align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>2Checkout</h4>    
                    <p>
                        2Checkout simplifies credit card processing by <strong>bundling the merchant account and payment gateway</strong> together. You don't have to worry about maintaining and paying for multiple payment processing services, 2Checkout packages it all together for one flat rate price.
                    </p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.2checkout.com" target="_blank">Learn More</a>                      
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
                                        <tr> 
                                            <td colspan="2" height="31">Current Store ID:&nbsp;<%=twoCheckout_store_idStart&twoCheckout_store_idEnd%></td>
                                        </tr>
                                        <tr> 
                                            <td colspan="2"> For security reasons, your account number is only partially shown on this page. The password is not shown. If you need to edit your account information, please re-enter your account name and password below.<br /></td>
                                        </tr>
                                    <% end if %>
                                    <tr>
                                      <td nowrap>&nbsp;</td>
                                      <td>&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td width="7%" nowrap>Store ID:</td>
                                        <td width="93%"><input type="text" value="" name="twoCheckout_store_id" size="24"></td>
                                    </tr>
                                    <tr valign="top">
                                        <td width="7%" nowrap>Checkout Option:</td>
                                        <td width="93%">
                                            <input type="radio" value="1" name="v2co" class="clearBorder" <%if v2co="1" then%>checked<%end if%>> Direct Checkout<br />
                                            <input type="radio" value="2" name="v2co" class="clearBorder" <%if v2co<>"1" then%>checked<%end if%>> Dynamic Checkout
                                        </td>
                                    </tr>
                                  <tr>
                                    <td width="7%">&nbsp;</td>
                                    <td><input name="v2co_testmode" type="checkbox" class="clearBorder" value="YES" <% if intv2co_testmode=1 then%>checked<% end if%> /><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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
<% end if %>