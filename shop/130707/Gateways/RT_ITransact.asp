<%
'---Start iTransact---
Function gwitEdit()
		'request gateway variables and insert them into the ITransact table
	query="SELECT Gateway_ID FROM ITransact WHERE id=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrDescription=err.description
		
	  	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrDescription) 
	end If
	Gateway_ID2=rstemp("Gateway_ID")
	Gateway_ID=request.Form("Gateway_ID")
	TransType = request.form("TransType")
	if Gateway_ID="" then
		Gateway_ID=Gateway_ID2
	end if
	URL=request.Form("ITURL")
	it_amex=request.Form("it_amex")
	if it_amex="" then
		it_amex=0
	end if
	it_diner=request.Form("it_diner")
	if it_diner="" then
		it_diner=0
	end if
	it_disc=request.Form("it_disc")
	if it_disc="" then
		it_disc=0
	end if
	it_mc=request.Form("it_mc")
	if it_mc="" then
		it_mc=0
	end if
	it_visa=request.Form("it_visa")
	if it_visa="" then
		it_visa=0
	end if
	if it_amex=0 AND it_diner=0 AND it_disc=0 AND it_mc=0 AND it_visa=0 then
		'at least one card must be selected for Itransact to be active
	end if
	ReqCVV = request.form("ReqCVV") 
	query="UPDATE ITransact SET Gateway_ID='"&Gateway_ID&"',URL='"&URL&"',it_amex="&it_amex&",it_diner="&it_diner&",it_disc="&it_disc&",it_mc="&it_mc&",it_visa="&it_visa&",ReqCVV="&ReqCVV&",transType="&TransType&" where id=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrDescription=err.description
		
	  	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode= 5"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrDescription=err.description
		
	  	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrDescription) 
	end If
	
end function

Function gwit()
	varCheck=1
	'request gateway variables and insert them into the ITransact table
	Gateway_ID=request.Form("Gateway_ID")
	URL=request.Form("ITURL")
	TransType = request.form("TransType")
	it_amex=request.Form("it_amex")
	if it_amex="" then
		it_amex=0
	end if
	it_diner=request.Form("it_diner")
	if it_diner="" then
		it_diner=0
	end if
	it_disc=request.Form("it_disc")
	if it_disc="" then
		it_disc=0
	end if
	it_mc=request.Form("it_mc")
	if it_mc="" then
		it_mc=0
	end if
	it_visa=request.Form("it_visa")
	if it_visa="" then
		it_visa=0
	end if
	if it_amex=0 AND it_diner=0 AND it_disc=0 AND it_mc=0 AND it_visa=0 then
		'at least one card must be selected for Itransact to be active
	end if
	
	ReqCVV = request.form("ReqCVV")  
	
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
	 

	query="UPDATE ITransact SET Gateway_ID='"&Gateway_ID&"',URL='"&URL&"',it_amex="&it_amex&",it_diner="&it_diner&",it_disc="&it_disc&",it_mc="&it_mc&",it_visa="&it_visa&", ReqCVV="&ReqCVV&", TransType="&TransType&" WHERE id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'ITransact','gwit.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",5,N'"&paymentNickName&"')"
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

<% if request("gwchoice")="5" then
	if request("mode")="Edit" then
				query="SELECT URL,Gateway_ID,it_amex,it_diner,it_disc,it_mc,it_visa, ReqCVV, TransType  FROM ITransact WHERE id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		URL=rs("URL")
		Gateway_ID=rs("Gateway_ID")
		it_amex=rs("it_amex")
		it_diner=rs("it_diner")
		it_disc=rs("it_disc")
		it_mc=rs("it_mc")
		it_visa=rs("it_visa")
		ReqCVV = rs("ReqCVV")
		TransType = rs("TransType")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=5"
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
	<input type="hidden" name="addGw" value="5">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/itransact-logo-header.png" width="250" height="50"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>iTransact</h4>    
                    <p>The iTransact payment gateway is compatible with merchant accounts obtained virtually anywhere.</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.itransact.com" target="_blank">Learn More</a>        
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
					<% dim Gateway_IDCnt,Gateway_IDEnd,Gateway_IDStart
                    Gateway_IDCnt=(len(Gateway_ID)-2)
                    Gateway_IDEnd=right(Gateway_ID,2)
                    Gateway_IDStart=""
                    for c=1 to Gateway_IDCnt
                        Gateway_IDStart=Gateway_IDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Gateway ID:&nbsp;<%=Gateway_IDStart&Gateway_IDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Gateway ID&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Gateway ID&quot; below.</td>
                    </tr>
                <% end if
				if URL="" then
					URL="https://secure.paymentclearing.com/cgi-bin/rc/ord.cgi"
				end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="111"> <div align="right">URL:</div></td>
                    <td width="476"> <input type="text" value="<%=URL%>" name="ITURL" size="40"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Gateway ID:</div></td>
                    <td> <input type="text" value="" name="Gateway_ID" size="24"></td>
                </tr>
                <tr>
                  <td>Transaction Type:</td>
                  <td><select name="TransType">
                            <option value="1" <% if TransType="1" then%>selected<% end if %>>Sale</option>
                            <option value="0" <% if TransType="0" then%>selected<% end if %>>Authorize Only</option>
                        </select></td>
                </tr>
                <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td>
                    <% if it_visa="1" then %> <input name="it_visa" type="checkbox" class="clearBorder" id="it_visa" value="1" checked> 
					<% else %> <input name="it_visa" type="checkbox" class="clearBorder" id="it_visa" value="1"> 
                    <% end if %>
                    Visa 
                    <% if it_mc="1" then %> <input name="it_mc" type="checkbox" class="clearBorder" id="it_mc" value="1" checked> 
                    <% else %> <input name="it_mc" type="checkbox" class="clearBorder" id="it_mc" value="1"> 
                    <% end if %>
                    MasterCard 
                    <% if it_amex="1" then %> <input name="it_amex" type="checkbox" class="clearBorder" id="it_amex" value="1" checked> 
                    <% else %> <input name="it_amex" type="checkbox" class="clearBorder" id="it_amex" value="1"> 
                    <% end if %>
                    American Express 
                    <% if it_disc="1" then %> <input name="it_disc" type="checkbox" class="clearBorder" id="it_disc" value="1" checked> 
                    <% else %> <input name="it_disc" type="checkbox" class="clearBorder" id="it_disc" value="1"> 
                    <% end if %>
                    Discover 
                    <% if it_diner="1" then %> <input name="it_diner" type="checkbox" class="clearBorder" id="it_diner" value="1" checked> 
                    <% else %> <input name="it_diner" type="checkbox" class="clearBorder" id="it_diner" value="1"> 
                    <% end if %>
                    Diners Club
                    </td>
                </tr>
                <tr>
                  <td align="right" >Require CVV:</td>
                  <td><input type="radio" class="clearBorder" name="ReqCVV" value="1" <% if ReqCVV="1" then%>checked<% end if %> >
                  Yes
                    <input type="radio" class="clearBorder" name="ReqCVV" value="0" <% if ReqCVV="0" then%>checked<% end if %>>
                  No</td>
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
