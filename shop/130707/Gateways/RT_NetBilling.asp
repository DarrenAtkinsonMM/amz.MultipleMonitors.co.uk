<%
'---Start Netbilling---
Function gwnetbillEdit()
		'request gateway variables and insert them into the netbill table
	query="SELECT NBAccountID,NBCVVEnabled,NBTranType, NBAVS, NetbillCheck, NBSiteTag FROM netbill WHERE idNetbill=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	NBAccountID2=rstemp("NBAccountID")
	'decrypt
	NBAccountID2=enDeCrypt(NBAccountID2, scCrypPass)
	NBAccountID=request.Form("NBAccountID")
	if NBAccountID="" then
		NBAccountID=NBAccountID2
	end if
	'encrypt
	NBAccountID=enDeCrypt(NBAccountID, scCrypPass)
	NBCVVEnabled=request.Form("NBCVVEnabled")
	NBAVS=request.Form("NBAVS")
	NBTranType2=rstemp("NBTranType")
	NBTranType=request.Form("NBTranType")
	If NBTranType="" then
		NBTranType=NBTranType2
	end if
	NBSiteTag=request.Form("NBSiteTag")
	NetbillCheck=request.Form("NetbillCheck")
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	query="UPDATE netbill SET NBAccountID='"&NBAccountID&"',NBCVVEnabled="&NBCVVEnabled&",NBTranType='"&NBTranType&"', NBAVS="&NBAVS&", NetbillCheck="&NetbillCheck&",NBSiteTag='"&NBSiteTag&"' WHERE idNetbill=1;"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=27"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	if NetbillCheck="1" then
		query="SELECT * FROM payTypes WHERE gwCode=28"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)
		if rstemp.eof then
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('Netbill eCheck','gwnetbillCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",28,N'"&paymentNickName2&"')"
		else
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentNickName=N'"&paymentNickName2&"' WHERE gwCode=28"
		end if
	else
		query="DELETE FROM payTypes WHERE gwCode=28"
	end if
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwnetbill()
	varCheck=1
	'request gateway variables and insert them into the netbilling table
	NBAccountID=request.Form("NBAccountID")
	'encrypt
	NBAccountID=enDeCrypt(NBAccountID, scCrypPass)
	NBSiteTag=request.Form("NBSiteTag")
	NBCVVEnabled=request.Form("NBCVVEnabled")
	NBAVS=request.Form("NBAVS")
	NBTranType=request.Form("NBTranType")
	NetbillCheck=request.Form("NetbillCheck")
	if NOT isNumeric(NBCVVEnabled) or NBCVVEnabled="" then
		NBCVVEnabled=0
	end if
	If NBAccountID="" then
		
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Netbilling as your payment gateway. <b>""Account ID""</b> is a required field.")
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
	end if
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	End If
	
	err.clear
	err.number=0
	 

	query="UPDATE netbill SET NBAccountID='"&NBAccountID&"',NBCVVEnabled="&NBCVVEnabled&",NBTranType='"&NBTranType&"', NetbillCheck="&NetbillCheck&", NBAVS="&NBAVS&", NBSiteTag='"&NBSiteTag&"' WHERE idNetbill=1"
	
	set rs=Server.CreateObject("ADODB.Recordset")  

	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Netbill','gwnetbill.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",27,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NetbillCheck="1" then
		query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'NetBill eCheck','gwnetbillCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",28,N'"&paymentNickName2&"')"
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

'27 = Credit
'28 = Checks

if request("gwchoice")="27" then
	if request("mode")="Edit" then
				query= "SELECT NBAccountID,NBCVVEnabled,NBTranType, NBAVS, NetbillCheck, NBSiteTag FROM netbill WHERE idNetbill=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		NBAccountID=rs("NBAccountID")
			'decrypt
		NBAccountID=enDeCrypt(NBAccountID, scCrypPass)
		NBCVVEnabled=rs("NBCVVEnabled")
		NBAVS=rs("NBAVS")
		NBTranType=rs("NBTranType")
		NetbillCheck=rs("NetbillCheck")
		NBSiteTag=rs("NBSiteTag")
		set rs=nothing
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=27"
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


		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=28"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName2="Check"
		else
			paymentNickName2=rs("paymentNickName")
		end if

		set rs=nothing
		
		%>
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="27">
<!-- END NETBILLING -->
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/netbilling_logo.JPG" width="276" height="69"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>NETbilling</h4>    
                    <p>NETbilling offers the most flexible &amp; powerful system, software, rates, &amp; customer support in the processing industry.</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.netbilling.com" target="_blank">Learn More</a>        
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
					<% dim NBAccountIDCnt,NBAccountIDEnd,NBAccountIDStart
                    NBAccountIDCnt=(len(NBAccountID)-2)
                    NBAccountIDEnd=right(NBAccountID,2)
                    NBAccountIDStart=""
                    for c=1 to NBAccountIDCnt
                        NBAccountIDStart=NBAccountIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Account Number:&nbsp;<%=NBAccountIDStart&NBAccountIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Account Number&quot; 
                            are only partially shown on this page. If you need 
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
                    <td> <input name="NBAccountID" type="TEXT" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Site Tag:</div></td>
                    <td><input type="text" name="NBSiteTag" size="30" value="<%=NBSiteTag%>"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="NBTranType">
                            <option value="S" <% if NBTranType="S" then%>selected<% end if %>>Sale</option>
                            <option value="A" <% if NBTranType="A" then%>selected<% end if %>>Authorize Only</option>
                        </select>								</td>
                </tr>
                <tr> 
                    <td><div align="right">Enable Address Verification (AVS):</div></td>
                    <td><input type="radio" class="clearBorder" name="NBAVS" value="1" checked>Yes 
                        <input type="radio" class="clearBorder" name="NBAVS" value="0" <% if NBAVS="0" then%>checked<% end if %>>No</td>
                </tr>
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="NBCVVEnabled" value="1" checked>Yes 
                        <input type="radio" class="clearBorder" name="NBCVVEnabled" value="0" <% if NBCVVEnabled="0" then%>checked<% end if %>>No</td>
                </tr>
                <tr> 

                    <td><div align="right">Accept Checks:</div></td>
                    <td> <input type="radio" class="clearBorder" name="NetbillCheck" value="1" checked>
                        Yes 
                        <input name="NetbillCheck" type="radio" class="clearBorder" value="0" <% if NetbillCheck="0" then%>checked<% end if %>>
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
