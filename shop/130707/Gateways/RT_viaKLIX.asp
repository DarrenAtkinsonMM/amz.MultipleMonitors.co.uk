<%
'---Start viaKLIX---
Function gwklixEdit()
		'request gateway variables and insert them into the Klix table
	query="SELECT ssl_merchant_id,ssl_pin,testmode,CVV,ssl_user_id FROM klix WHERE idKlix=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	ssl_merchant_id2=rstemp("ssl_merchant_id")
	ssl_merchant_id=request.Form("ssl_merchant_id")
	if ssl_merchant_id="" then
		ssl_merchant_id=ssl_merchant_id2
	end if
	ssl_pin2=rstemp("ssl_pin")
	'decrypt
	ssl_pin2=enDeCrypt(ssl_pin2, scCrypPass)
	ssl_pin=request.Form("ssl_pin")
	if ssl_pin="" then
		ssl_pin=ssl_pin2
	end if
	'encrypt
	ssl_pin=enDeCrypt(ssl_pin, scCrypPass)
	ssl_user_id2=rstemp("ssl_user_id")
	ssl_user_id=request.Form("ssl_user_id")
	if ssl_user_id="" then
		ssl_user_id=ssl_user_id2
	end if
	CVV=request.Form("CVV")
	testmode=request.Form("testmode")
	query="UPDATE klix SET ssl_merchant_id='"&ssl_merchant_id&"',ssl_pin='"&ssl_pin&"',CVV="&CVV&",ssl_user_id='"&ssl_user_id&"' WHERE idKlix=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="&priceToAdd&" , percentageToAdd="&percentageToAdd&", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=23"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwklix()
	varCheck=1
	'request gateway variables and insert them into the klix table
	ssl_merchant_id=request.Form("ssl_merchant_id")
	ssl_user_id=request.Form("ssl_user_id")
	ssl_pin=request.Form("ssl_pin")
	'encrypt
	ssl_pin=enDeCrypt(ssl_pin, scCrypPass)
	CVV=request.Form("CVV")
	testmode=request.Form("ssl_testmode")
	if testmode="YES" then
		testmode="1"
	else
		testmode="0"
	end if
	if NOT isNumeric(CVV) or CVV="" then
		CVV="0"
	end if
	If ssl_merchant_id="" OR ssl_user_id="" then
		
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add viaKlix as your payment gateway. <b>""Merchant ID""</b> and <b>""User ID""</b> are required fields.")
	End If
	If ssl_pin="" then
		
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add viaKlix as your payment gateway. <b>""PIN""</b> is a required field.")
	End If
	'encrypt
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
	 

	query="UPDATE klix SET ssl_merchant_id='"&ssl_merchant_id&"',ssl_pin='"&ssl_pin&"',CVV="&CVV&", testmode="&testmode&",ssl_user_id='"&ssl_user_id&"' WHERE idKlix=1"
	set rs=Server.CreateObject("ADODB.Recordset")  

	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'viaKLIX','gwklix.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",23,N'"&paymentNickName&"')"
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

if request("gwchoice")="23" then
	if request("mode")="Edit" then
				query= "SELECT ssl_merchant_id,ssl_pin,CVV,testmode,ssl_user_id FROM klix WHERE idKlix=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		ssl_merchant_id=rs("ssl_merchant_id")
		ssl_pin=rs("ssl_pin")
			'decrypt
			ssl_pin=enDeCrypt(ssl_pin, scCrypPass)
		CVV=rs("CVV")
		testmode=rs("testmode")
		ssl_user_id=rs("ssl_user_id")
		set rs=nothing
		
		%>
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="23">

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="images/pcv4_icon_pg.png" width="48" height="48"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>viaKLIX</h4>    
                    <p></p>
                    <p>
                        <a class="btn btn-info btn-xs" href="https://www2.viaklix.com/Admin/main.asp" target="_blank">Learn More</a>        
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
					<% dim ssl_merchant_idCnt,ssl_merchant_idEnd,ssl_merchant_idStart,ssl_user_idCnt,ssl_user_idEnd,ssl_user_idStart
                    ssl_merchant_idCnt=(len(ssl_merchant_id)-2)
                    ssl_merchant_idEnd=right(ssl_merchant_id,2)
                    ssl_merchant_idStart=""
                    for c=1 to ssl_merchant_idCnt
                        ssl_merchant_idStart=ssl_merchant_idStart&"*"
                    next
                    ssl_user_idCnt=(len(ssl_user_id)-2)
                    ssl_user_idEnd=right(ssl_user_id,2)
                    ssl_user_idStart=""
                    for c=1 to ssl_user_idCnt
                        ssl_user_idStart=ssl_user_idStart&"*"
                    next %>
                    <tr> 
                        <td colspan="2">Current Merchant ID:&nbsp;<%=ssl_merchant_idStart&ssl_merchant_idEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2">Current User ID:&nbsp;<%=ssl_user_idStart&ssl_user_idEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Merchant ID&quot; 
                            and &quot;User ID&quot; are only partially shown 
                            on this page. If you need to edit your account information, 
                            please re-enter your &quot;Merchant ID&quot; and 
                            &quot;User ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td width="140px"><div align="right">Merchant ID:</div></td>
                    <td> <input type="text" name="ssl_merchant_id" size="20"></td>
                </tr>
                <tr> 
                    <td><div align="right">User ID:</div></td>
                    <td> <input type="text" name="ssl_user_id" size="20"></td>
                </tr>
                <tr> 
                    <td><div align="right">PIN # :</div></td>
                    <td><input name="ssl_pin" type="password" size="20"> </td>
                </tr>
                <tr> 
                    <td><div align="right">Require CVV:</div></td>
                    <td><input type="radio" class="clearBorder" name="CVV" value="1" checked>
                        Yes 
                        <input type="radio" class="clearBorder" name="CVV" value="0" <% if CVV="0" then %>checked<% end if %>>
                        No</td>
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
                    <td width="111"> <div align="right">Processing fee: </div></td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>
                        Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
                    </td>
                </tr>
                <tr> 
                    <td>&nbsp;</td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                        Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                        <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">
                    </td>
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
						<td><div align="right">Payment Name:&nbsp;</div></td>
						<td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
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
                        <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1">&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=301"></a></td>
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
                    <tr>
                      <td>&nbsp;</td>
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
