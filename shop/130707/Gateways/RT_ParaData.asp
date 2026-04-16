<%
'--- Start ParaData ---
Function gwParaDataEdit()
		'request gateway variables and insert them into the ParaData table
	query="SELECT pcPay_ParaData_Key FROM pcPay_ParaData Where pcPay_ParaData_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_ParaData_Key2=rs("pcPay_ParaData_Key")

	pcPay_ParaData_Key=request.Form("pcPay_ParaData_Key")
	if pcPay_ParaData_Key="" then
		pcPay_ParaData_Key=pcPay_ParaData_Key2
	end if
	set rs=nothing
	
	pcPay_ParaData_TransType=request.Form("pcPay_ParaData_TransType")
	
	pcPay_ParaData_TestMode=request.Form("pcPay_ParaData_TestMode")
	if pcPay_ParaData_TestMode="" then
		pcPay_ParaData_TestMode="0"
	end if

	pcPay_ParaData_CVC=request.Form("pcPay_ParaData_CVC")
	if pcPay_ParaData_CVC="1" then
		pcPay_ParaData_CVC=1
	else
		pcPay_ParaData_CVC=0
	end if

	pcPay_ParaData_AVS=request.Form("pcPay_ParaData_AVS")
	if pcPay_ParaData_AVS="YES" then
		pcPay_ParaData_AVS="1"
	else
		pcPay_ParaData_AVS="0"
	end if
	
	query="UPDATE pcPay_ParaData SET pcPay_ParaData_TransType='"&pcPay_ParaData_TransType&"', pcPay_ParaData_Key='"&pcPay_ParaData_Key&"', pcPay_ParaData_TestMode="&pcPay_ParaData_TestMode&", pcPay_ParaData_CVC="&pcPay_ParaData_CVC&", pcPay_ParaData_AVS="&pcPay_ParaData_AVS&" WHERE pcPay_ParaData_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName=N'"&paymentNickName&"' WHERE gwCode=45"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwParaData()
	varCheck=1
	pcPay_ParaData_TransType=request.Form("pcPay_ParaData_TransType")

	pcPay_ParaData_TestMode=request.Form("pcPay_ParaData_TestMode")
	if pcPay_ParaData_TestMode="" then
		pcPay_ParaData_TestMode="0"
	end if

	pcPay_ParaData_Key=request.Form("pcPay_ParaData_Key")
	if pcPay_ParaData_Key="" AND pcPay_ParaData_TestMode="0" then
		call closeDb()
response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Paradata as your payment gateway. <b>""Transaction Key""</b> is a required field.")
	End If

	pcPay_ParaData_CVC=request.Form("pcPay_ParaData_CVC")
	if pcPay_ParaData_CVC="1" then
		pcPay_ParaData_CVC=1
	else
		pcPay_ParaData_CVC=0
	end if

	pcPay_ParaData_AVS=request.Form("pcPay_ParaData_AVS")
	if pcPay_ParaData_AVS="YES" then
		pcPay_ParaData_AVS="1"
	else
		pcPay_ParaData_AVS="0"
	end if

	err.clear
	err.number=0
	   

	query="UPDATE pcPay_ParaData SET pcPay_ParaData_TransType='"&pcPay_ParaData_TransType&"', pcPay_ParaData_Key='"&pcPay_ParaData_Key&"', pcPay_ParaData_TestMode="&pcPay_ParaData_TestMode&", pcPay_ParaData_CVC="&pcPay_ParaData_CVC&", pcPay_ParaData_AVS="&pcPay_ParaData_AVS&" WHERE pcPay_ParaData_ID=1;"

	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
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
	end if
	query = ""
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Paradata','gwParaData.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",45,N'"&paymentNickName&"')"
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

if request("gwchoice")="45" then
	'//Check for required components and that they are working 
	reDim strComponent(0)
	reDim strClass(0,1)

	'The component names
	strComponent(0) = "ParaData"
	
	'The component class names
	strClass(0,0) = "Paygateway.EClient.1"
	
	isComErr = Cint(0)
	strComErr = Cstr()
	
	For i=0 to UBound(strComponent)
		strErr = IsObjInstalled(i)
		If strErr <> "" Then
			strComErr = strComErr & strErr
			isComErr = 1
		End If
	Next

	if request("mode")="Edit" then
				query="SELECT pcPay_ParaData_TransType, pcPay_ParaData_Key, pcPay_ParaData_TestMode, pcPay_ParaData_AVS, pcPay_ParaData_CVC FROM pcPay_ParaData WHERE pcPay_ParaData_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If 
		pcPay_ParaData_TransType=rs("pcPay_ParaData_TransType") ' auth or sale
		pcPay_ParaData_Key=rs("pcPay_ParaData_Key") ' private key
		pcPay_ParaData_TestMode=rs("pcPay_ParaData_TestMode")  ' test mode or live mode
		pcPay_ParaData_AVS=rs("pcPay_ParaData_AVS") ' avs "on" or "off"
		pcPay_ParaData_CVC=rs("pcPay_ParaData_CVC") ' cvc "on" or "off"
		set rs=nothing
		
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="45">
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
                    <h4>Paradata</h4>    
                    <p></p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.ParaData.com" target="_blank">Learn More</a>        
                    </p>
                </div>
                
			</td>
        </tr>
        <tr>
            <td>
        <div id="acc2">
			<% if isComErr = 1 then
			   intDoNotApply = 1 %>

            
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
                <table width="100%" border="0" cellspacing="0" cellpadding="4">
                	<tr>
                	  <td><img src="images/red_x.png" alt="Unable to add Paradata" width="12" height="12" /> <strong>Paradata cannot be activated</strong> - the required third party components are not installed and working on the web server.<br />                	    <br /></td>
              	  </tr>
                	<tr>
                    	<td>
                        	<center>
                        	<strong>Required components for Paradata:</strong><br />
                       	  <i><%= strComErr %></i><br /><br />
                        	<input type="button" class="btn btn-default"  value="Back" onclick="javascript:history.back()"></center></td>
                  	</tr>
              	</table>
                            </div>
                        </div> 
                    </div>
                    
                </div>
			<% else %>

            
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
					<% dim pcPay_ParaData_KeyCnt,pcPay_ParaData_KeyEnd,pcPay_ParaData_KeyStart
                    pcPay_ParaData_KeyCnt=(len(pcPay_ParaData_Key)-2)
                    pcPay_ParaData_KeyEnd=right(pcPay_ParaData_Key,2)
                    pcPay_ParaData_KeyStart=""
                    for c=1 to pcPay_ParaData_KeyCnt
                        pcPay_ParaData_KeyStart=pcPay_ParaData_KeyStart&"*"
                    next
                    %>
                    <tr> 
                        <td height="31" colspan="2">Current Account Token:&nbsp;<%=pcPay_ParaData_KeyStart&pcPay_ParaData_KeyEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Account Token&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Account Token&quot; below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td width="24%"><div align="right">Account Token:</div></td>
                    <td width="76%"> <div align="left"><input type="text" value="" name="pcPay_ParaData_Key" size="30"></div></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> 
                    <select name="pcPay_ParaData_TransType">
                    <option value="SALE" <% if pcPay_ParaData_TransType="SALE" then%>selected<% end if %>>Sale</option>
                    <option value="AUTH" <% if pcPay_ParaData_TransType="AUTH" then%>selected<% end if %>>Authorize Only</option>
                    </select></td>
                </tr>
                
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_ParaData_CVC" value="1" checked>Yes 
                    <input name="pcPay_ParaData_CVC" type="radio" class="clearBorder" value="0" <% if pcPay_ParaData_CVC=0 then%>checked<% end if %>>No
                    <font color="#FF0000">&nbsp;&nbsp;*Required if you are accepting Discover cards.</font></td>
                </tr>
                
                <tr> 
                    <td><div align="right"> 
                    <input name="pcPay_ParaData_TestMode" type="checkbox" class="clearBorder" id="pcPay_ParaData_TestMode" value="1" <% if pcPay_ParaData_TestMode=1 then %>checked<% end if %> />
                    </div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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
                
                
		<%end if%>



            </td>
        </tr>
    </table>
<!-- New View End --><% end if %>
