<%	
'---Start InternetSecure---
Function gwIntSecureEdit()
		'request gateway variables and insert them into the BluePay table
	query="SELECT IsMerchantNumber,IsLanguage,IsCurrency,IsTestmode FROM InternetSecure WHERE IsID=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	IsMerchantNumber2=rstemp("IsMerchantNumber")
	IsMerchantNumber=request.Form("IsMerchantNumber")
	If IsMerchantNumber="" then
		IsMerchantNumber=IsMerchantNumber2
	end if
	IsLanguage2=rstemp("IsLanguage")
	IsLanguage=request.Form("IsLanguage")
	If IsLanguage="" then
		IsLanguage=IsLanguage2
	end if
	IsCurrency2=rstemp("IsCurrency")
	IsCurrency=request.Form("IsCurrency")
	if IsCurrency="" then
		IsCurrency=IsCurrency2
	end if
	IsTestmode=request.Form("IsTestmode")
	if IsTestmode="1" then
		IsTestmode=1
	else
		IsTestmode=0
	end if

	query="UPDATE InternetSecure SET IsMerchantNumber='"&IsMerchantNumber&"',IsLanguage='"&IsLanguage&"',IsCurrency='"&IsCurrency&"',IsTestmode="&IsTestmode&" WHERE isID=1;"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=30"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwIntSecure()
	varCheck=1
	'request gateway variables and insert them into the InterNetSecure table
	IsMerchantNumber=request.Form("IsMerchantNumber")
	IsLanguage=request.Form("IsLanguage")
	IsCurrency=request.Form("IsCurrency")
	IsTestmode=request.Form("IsTestmode")
	if IsTestmode="1" then
		IsTestmode=1
	else
		IsTestmode=0
	end if
	If IsMerchantNumber="" then
		
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add InternetSecure as your payment gateway. <b>""Merchant""</b> is a required field.")
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
	 

	query="UPDATE InternetSecure SET IsMerchantNumber='"&IsMerchantNumber&"',IsLanguage='"&IsLanguage&"',IsCurrency='"&IsCurrency&"',IsTestmode="&IsTestmode&" WHERE IsID=1"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'InternetSecure','gwIntSecure.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",30,N'"&paymentNickName&"')"
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

<% if request("gwchoice")="30" then
	if request("mode")="Edit" then
				query="SELECT IsMerchantNumber,IsLanguage,IsCurrency,IsTestmode FROM InternetSecure WHERE IsID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		IsMerchantNumber=rs("IsMerchantNumber")
		IsLanguage=rs("IsLanguage")
		IsCurrency=rs("IsCurrency")
		IsTestmode=rs("IsTestmode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=30"
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
	<input type="hidden" name="addGw" value="30">
<!-- END INTERNETSECURE -->
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/internetsecure_logo.JPG" width="291" height="74"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>InternetSecure</h4>    
                    <p>Everything your business needs to accept major credit cards.</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.internetsecure.com" target="_blank">Learn More</a>        
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
					<% dim IsMerchantNumberCnt,IsMerchantNumberEnd,IsMerchantNumberStart
                    IsMerchantNumberCnt=(len(IsMerchantNumber)-2)
                    IsMerchantNumberEnd=right(IsMerchantNumber,2)
                    IsMerchantNumberStart=""
                    for c=1 to IsMerchantNumberCnt
                        IsMerchantNumberStart=IsMerchantNumberStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current InternetSecure Merchant Number:&nbsp;<%=IsMerchantNumberStart&IsMerchantNumberEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;InternetSecure 
                            merchant number&quot; is only partially shown on 
                            this page. If you need to edit your account information, 
                            please re-enter your &quot;Merchant Number&quot; 
                            below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="111"> <div align="right">Merchant Number:</div></td>
                    <td width="360"> <input type="text" name="IsMerchantNumber" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Language:</div></td>
                    <td width="360">
                    	<select name="IsLanguage">
                            <option value="EN" Selected>English</option>
                            <option value="FR" <% if IsLanguage="FR" then %>Selected<%end if%>>French</option>
                            <option value="SP" <% if IsLanguage="SP" then %>Selected<%end if%>>Spanish</option>
                            <option value="PT" <% if IsLanguage="PT" then %>Selected<%end if%>>Portuguese</option>
                            <option value="JP" <% if IsLanguage="JP" then %>Selected<%end if%>>Japanese</option>
                        </select></td>
                </tr>
                <tr> 
                    <td> <div align="right">Currency:</div></td>
                    <td width="360">
                    	<select name="IsCurrency">
                            <option value="CND" Selected>CND</option>
                            <option value="USD" <% if IsCurrency="USD" then %>Selected<%end if%>>USD</option>
                        </select> Note: USD only available if you have a $CDN and $US Account.</td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="IsTestmode" type="checkbox" class="clearBorder" value="1" <% if IsTestmode=1 then%>checked<% end if %>>
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
