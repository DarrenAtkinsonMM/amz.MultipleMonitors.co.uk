<% dim pcv_BeagleNotAvailable
pcv_BeagleNotAvailable=1

'//Check if Beagle Field Exists
on error resume next
err.clear
query="SELECT * FROM eWay;"
set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query)
eWay_BeagleActive=rstemp("eWayBeagleActive")
if err.number<>0 then
	pcv_BeagleNotAvailable=0
end if
set rstemp=nothing

err.clear

'---Start eWay---
Function gwEwayEdit()
		
	pcv_BeagleNotAvailable=1
	
	query="SELECT * FROM eWay;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	eWay_BeagleActive=rstemp("eWayBeagleActive")
	if err.number<>0 then
		pcv_BeagleNotAvailable=0
	end if

	'request gateway variables and insert them into the eWay table
	query="SELECT eWayCustomerId, eWayPostMethod FROM eWay WHERE eWayID=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	eWayCustomerId2=rstemp("eWayCustomerId")
	eWayCustomerId=request.Form("eWayCustomerId")
	If eWayCustomerId="" then
		eWayCustomerId=eWayCustomerId2
	end if
	eWayPostMethod2=rstemp("eWayPostMethod")
	eWayPostMethod=request.Form("eWayPostMethod")
	If eWayPostMethod="" then
		eWayPostMethod=eWayPostMethod2
	end if
	eWayTestmode=request.Form("eWayTestmode")
	if eWayTestmode="YES" then
		eWayTestmode="1"
	else
		eWayTestmode="0"
	end if
	eWay_CVV = request.form("eWay_CVV")
	if pcv_BeagleNotAvailable=1 then
		eWay_BeagleActive = request.form("eWay_BeagleActive")
	end if
	query="UPDATE eWay SET eWayCustomerId='"&eWayCustomerId&"', eWayPostMethod='"&eWayPostMethod&"', eWayTestmode="&eWayTestmode&",eWayCVV=" & eWay_CVV
	if pcv_BeagleNotAvailable=1 then
		query=query &", eWayBeagleActive=" & eWay_BeagleActive
	end if
	query=query &" WHERE eWayID=1;"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=31"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwEway()
		pcv_BeagleNotAvailable=1
	
	query="SELECT * FROM eWay;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	eWay_BeagleActive=rstemp("eWayBeagleActive")
	if err.number<>0 then
		pcv_BeagleNotAvailable=0
	end if

	varCheck=1
	'request gateway variables and insert them into the eWay table
	eWayCustomerId=request.Form("eWayCustomerId")
	eWayPostMethod=request.Form("eWayPostMethod")
	eWayTestmode=request.Form("eWayTestmode")
	eWay_CVV = request.form("eWay_CVV")
	eWay_BeagleActive = request.form("eWay_BeagleActive")
	if eWayTestmode="YES" then
		eWayTestmode="1"
	else
		eWayTestmode="0"
	end if
	If eWayCustomerId="" then
		
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eWay as your payment gateway. <b>""Customer ID""</b> is a required field.")
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
	 

	query="UPDATE eWay SET eWayCustomerId='"&eWayCustomerId&"',eWayPostMethod='"&eWayPostMethod&"',eWayTestmode="&eWayTestmode&",eWayCVV=" & eWay_CVV
	if pcv_BeagleNotAvailable=1 then
		query=query &",eWayBeagleActive=" & eWay_BeagleActive
	end if
	query=query &" WHERE eWayID=1"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'eWay','gwEway.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",31,N'"&paymentNickName&"')"
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

if request("gwchoice")="31" then
	tmp_id=request("id")
	tmp_mode=request("mode")

	'Check to see if fields exists in table, if not, add
	err.clear
		query="SELECT eWayCVV FROM eWay"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		
		call closeDb()
response.redirect "upddbEway.asp?mode="&tmp_mode&"&id="&tmp_id
	else
		set rs=nothing
		
	end if
	
	if request("mode")="Edit" then
        		query="SELECT eWayCustomerId, eWayPostMethod, eWayTestmode, eWayCVV"
		if pcv_BeagleNotAvailable=1 then
			query=query&", eWayBeagleActive"
		end if
		query=query&" FROM eWay WHERE eWayID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		eWayCustomerId=rs("eWayCustomerId")
		eWayPostMethod=rs("eWayPostMethod")
		eWayTestmode=rs("eWayTestmode")
		eWay_CVV = rs("eWayCVV")
		if pcv_BeagleNotAvailable=1 then
			eWay_BeagleActive = rs("eWayBeagleActive")
		end if
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=31"
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
	<input type="hidden" name="addGw" value="31">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/eway_logo.png" width="180" height="83"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>eWay</h4>    
                    <p>Payments made easy.</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.eway.com.au" target="_blank">Learn More</a>        
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
                <tr>
                  <td align="right" valign="top">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td align="right" valign="top"><input name="eWayPostMethod" type="radio" class="clearBorder" value="SHARED" checked></td>
                    <td><strong>Shared Payment</strong><b><br>
                </b>Process credit card payments via eWAY's own secure 
                Shared Payment Page in real time. You POST purchase 
                information from your web site to the eWAY secured site.</td>
                </tr>
                <tr> 
                    <td align="right" valign="top"><input name="eWayPostMethod" type="radio" class="clearBorder" value="XML" <% if eWayPostMethod="XML" then%>checked<% end if %>></td>
                    <td valign="top"><b>XML Payment<font color="#FF0000"> 
                        </font></b><em>(recommended)</em><b><br>
                        </b>Process credit card payments directly through your 
                        own website in real time. Using the eWAY XML Solution, 
                        your web site appears as the payment gateway, with the 
                        transactions POSTed in the background. </td>
                </tr>
                <% if request("mode")="Edit" then %>
					<% dim eWayCustomerIdCnt,eWayCustomerIdEnd,eWayCustomerIdStart
                    eWayCustomerIdCnt=(len(eWayCustomerId)-2)
                    eWayCustomerIdEnd=right(eWayCustomerId,2)
                    eWayCustomerIdStart=""
                    for c=1 to eWayCustomerIdCnt
                        eWayCustomerIdStart=eWayCustomerIdStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current InternetSecure Merchant Number:&nbsp;<%=eWayCustomerIdStart&eWayCustomerIdEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;eWay Customer 
                            ID&quot; is only partially shown on this page. If 
                            you need to edit your account information, please 
                            re-enter your &quot;Customer ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td> <div align="right">Customer ID:</div></td>
                    <td width="1203"> <input type="text" name="eWayCustomerId" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="eWayTestmode" type="checkbox" class="clearBorder" value="YES" <% if eWayTestmode=1 then%>checked<% end if%>> 
                        </div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>
				<TR>
				  <td><div align="right">Real-Time CVN:</div></td>
			      <td>
				        <input type="radio" class="clearBorder" name="eWay_CVV" value="1" <%if eWay_CVV=1 then%> Checked <%end if %> /> Yes 
                        <input name="eWay_CVV" type="radio" class="clearBorder" value="0" <%if eWay_CVV=0 then%> Checked <%end if %> /> 
                       No
(Real-Time CVN is an optional eWay subscription) </td>
				</TR>
                <% if pcv_BeagleNotAvailable=1 then %> 
                    <TR>
                      <td nowrap="nowrap"><div align="right">Beagle (Geo-IP Anti Fraud):</div></td>
                      <td>
                            <input type="radio" class="clearBorder" name="eWay_BeagleActive" value="1" <%if eWay_BeagleActive=1 then%> Checked <%end if %> /> Yes 
                            <input name="eWay_BeagleActive" type="radio" class="clearBorder" value="0" <%if eWay_BeagleActive=0 then%> Checked <%end if %> /> 
                           No
    (Beagle Fraud Prevention is an optional eWay subscription) </td>
                    </TR> 
                <% else %>
                    <TR>
                      <td nowrap="nowrap" valign="top"><div align="right">Beagle (Geo-IP Anti Fraud):</div></td>
                      <td>
                        <input name="eWay_BeagleActive" type="hidden" value="0" />
                        Beagle Fraud Prevention is an optional eWay subscription. This feature is not available in ProductCart until you update your database. <a href="upddbEway.asp">Click here to update your database now.</a></td>
                    </TR> 
                <% end if %>
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
