<%
'---Start FastTransAct---
Function gwfastEdit()
		'request gateway variables and insert them into the fasttransact table
	query= "SELECT AccountID, SiteTag FROM fasttransact WHERE id=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	AccountID2=rstemp("AccountID")
	SiteTag2=rstemp("SiteTag")
	AccountID=request.Form("AccountID")
	If AccountID="" then
		AccountID=AccountID2
	end if
	SiteTag=request.Form("SiteTag")
	if SiteTag="" then
		SiteTag=SiteTag2
	end if
	tran_type=request.Form("tran_type")
	card_types=request.Form("card_types")
	CVV2=request.Form("CVV2")
	query="UPDATE fasttransact SET AccountID='"&AccountID&"',SiteTag='"&SiteTag&"',tran_type='"&tran_type&"',card_types='"&card_types&"',CVV2="&CVV2&" WHERE id=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=15"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
end function

Function gwfast()
	varCheck=1
	'request gateway variables and insert them into the fasttransact table
	AccountId=request.form("AccountID")
	SiteTag=request.form("SiteTag")
	tran_type=request.Form("tran_type")
	card_types=request.Form("card_types")
	CVV2=request.Form("CVV2")
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
	
		
	query="UPDATE fasttransact SET AccountID='"&AccountID&"',SiteTag='"&SiteTag&"',tran_type='"&tran_type&"',card_types='"&card_types&"',CVV2="&CVV2&" WHERE id=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'FastTransact','gwfast.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",15,N'"&paymentNickName&"')"
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

if request("gwchoice")="15" then
	if request("mode")="Edit" then
				query= "SELECT AccountID, SiteTag, tran_type, card_types, CVV2 FROM fasttransact WHERE id=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		AccountID=rs("AccountID")
		SiteTag=rs("SiteTag")
		tran_type=rs("tran_type")
		card_types=rs("card_types")
		CVV2=rs("CVV2")
		
		cardTypeArray=split(card_types,", ")
		for i=0 to ubound(cardTypeArray)
			select case cardTypeArray(i)
				case "M"
					M="1" 
				case "V"
					V="1"
				case "D"
					D="1"
				case "A"
					A="1"
			end select
		next
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=15"
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
	<input type="hidden" name="addGw" value="15">
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
                    <h4>Fast Transact</h4>    
                    <p></p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://www.fasttransact.com" target="_blank">Learn More</a>        
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
					<% dim FT_AccountIDCnt,FT_AccountIDEnd,FT_AccountIDStart
                    FT_AccountIDCnt=(len(AccountID)-2)
                    FT_AccountIDEnd=right(AccountID,2)
                    FT_AccountIDStart=""
                    for c=1 to FT_AccountIDCnt
                        FT_AccountIDStart=FT_AccountIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Account ID:&nbsp;<%=FT_AccountIDStart&FT_AccountIDEnd%></td>
                    </tr>
                    <% dim FT_SiteTagCnt,FT_SiteTagEnd,FT_SiteTagStart
                    FT_SiteTagCnt=(len(SiteTag)-2)
                    FT_SiteTagEnd=right(SiteTag,2)
                    FT_SiteTagStart=""
                    for c=1 to FT_SiteTagCnt
                        FT_SiteTagStart=FT_SiteTagStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Site Tag:&nbsp;<%=FT_SiteTagStart&FT_SiteTagEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Account ID&quot; 
                            and &quot;Site Tag&quot; is only partially shown 
                            on this page. If you need to edit your account information, 
                            please re-enter your &quot;Account ID&quot; and 
                            &quot;Site Tag&quot; below.</td>
                    </tr>
                <% end if %>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Account ID :</div></td>
                    <td> <input type="text" name="AccountID" size="20"> </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Site Tag :</div></td>
                    <td width="440"> <input name="SiteTag" type="text" size="30"></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="tran_type">
                            <option value="SALE" <% if tran_type="SALE" then %>selected<% end if %>>Sale</option>
                            <option value="PREAUTH" <% if tran_type="PREAUTH" then %>selected<% end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="CVV2" value="1" checked>
                        Yes 
                        <input type="radio" class="clearBorder" name="CVV2" value="0" <% if CVV2="0" then %>checked<% end if %>>
                        No</td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td>
                    <% if V="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="V" checked> 
					<% else %> <input name="card_types" type="checkbox" class="clearBorder" value="V"> 
                    <% end if %>
                    Visa 
                    <% if M="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="M" checked> 
                    <% else %> <input name="card_types" type="checkbox" class="clearBorder" value="M"> 
                    <% end if %>
                    MasterCard 
                    <% if A="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="A" checked> 
                    <% else %> <input name="card_types" type="checkbox" class="clearBorder" value="A"> 
                    <% end if %>
                    American Express 
                    <% if D="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="D" checked> 
                    <% else %> <input name="card_types" type="checkbox" class="clearBorder" value="D"> 
                    <% end if %>
                    Discover
                    </td>
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
                  <tr bgcolor="#FFFFFF"> 
                    <td width="111"> <div align="right">Processing fee: </div></td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>
                        Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
                    </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
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
                <tr bgcolor="#FFFFFF"> 
                    <td><div align="right">Payment Name:&nbsp;</div></td>
                    <td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
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
