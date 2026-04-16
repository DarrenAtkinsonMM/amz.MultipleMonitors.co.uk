<%
'---Start Moneris---
Function gwmonerisEdit()

		'request gateway variables and insert them into the Moneris table
	query= "SELECT pcPay_Moneris_StoreId,pcPay_Moneris_Key, pcPay_Moneris_TransType, pcPay_Moneris_Lang, pcPay_Moneris_Testmode, pcPay_Moneris_CVVEnabled, pcPay_Moneris_Meth, pcPay_Moneris_Interac,pcPay_Moneris_Interac_MerchantID FROM pcPay_Moneris WHERE pcPay_Moneris_Id=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)

	if err.number <> 0 then
			
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	

	pcPay_Moneris_StoreId2=rs("pcPay_Moneris_StoreId")
	pcPay_Moneris_Key2 = rs("pcPay_Moneris_Key")
	
	pcPay_Moneris_StoreId2=enDeCrypt(pcPay_Moneris_StoreId2, scCrypPass)
	
	pcPay_Moneris_StoreId=request.Form("pcPay_Moneris_StoreId")
	if pcPay_Moneris_StoreId="" then
		pcPay_Moneris_StoreId=pcPay_Moneris_StoreId2
	end if
	if pcPay_Moneris_Key2<>"" OR NOT isNull(pcPay_Moneris_Key2) then
		pcPay_Moneris_Key2=enDeCrypt(pcPay_Moneris_Key2, scCrypPass)
	end if
	pcPay_Moneris_Key=request.Form("pcPay_Moneris_Key")
	if pcPay_Moneris_Key="" then
		pcPay_Moneris_Key=pcPay_Moneris_Key2
	end if
	pcPay_Moneris_Interac_MerchantID2 = rs("pcPay_Moneris_Interac_MerchantID")
	if pcPay_Moneris_Interac_MerchantID2<>"" OR NOT isNull(pcPay_Moneris_Interac_MerchantID2) then
		pcPay_Moneris_Interac_MerchantID2 = enDeCrypt(pcPay_Moneris_Interac_MerchantID2, scCrypPass)
	end if
	pcPay_Moneris_Interac_MerchantID = request.form("pcPay_Moneris_Interac_MerchantID")
	if pcPay_Moneris_Interac_MerchantID="" then
		pcPay_Moneris_Interac_MerchantID=pcPay_Moneris_Interac_MerchantID2
	end if
	
	set rs=nothing
	
	pcPay_Moneris_TransType=request.Form("pcPay_Moneris_TransType")
	pcPay_Moneris_Lang=request.Form("pcPay_Moneris_Lang")
	pcPay_Moneris_Testmode=request.Form("pcPay_Moneris_Testmode")
	if pcPay_Moneris_Testmode="" then
		pcPay_Moneris_Testmode=0
	end if
	
   	 pcPay_Moneris_Ver = Request.form("pcPay_Moneris_Ver")
	  if pcPay_Moneris_Ver ="US" Then
	  	mFileV = "gwMonerisUS.asp"
	  else
	  	mFileV = "gwMoneris2.asp" 
	 End if
	 
	pcPay_Moneris_CVVEnabled=request.Form("pcPay_Moneris_CVVEnabled")
	pcPay_Moneris_Meth = request.form("pcPay_Moneris_Meth")

	pcPay_Moneris_Interac = request.form("pcPay_Moneris_Interac")
	if pcPay_Moneris_Interac="" then
		pcPay_Moneris_Interac=0
	end if
	pcPay_Moneris_StoreId=enDeCrypt(pcPay_Moneris_StoreId, scCrypPass)
	if pcPay_Moneris_Key&""<>"" then
		pcPay_Moneris_Key=enDeCrypt(pcPay_Moneris_Key, scCrypPass)
	end if
	if pcPay_Moneris_Interac_MerchantID&""<>"" then
		pcPay_Moneris_Interac_MerchantID=enDeCrypt(pcPay_Moneris_Interac_MerchantID, scCrypPass)
	end if

	paymentNickName2 = request.form("paymentNickName2")
	if paymentNickName2="" then
		paymentNickName2="INTERAC&reg; Online Accepted Here"
	end if
		
	query="UPDATE pcPay_Moneris SET pcPay_Moneris_StoreId='"&pcPay_Moneris_StoreId&"',pcPay_Moneris_Key='"&pcPay_Moneris_Key&"',pcPay_Moneris_TransType='"&pcPay_Moneris_TransType&"',pcPay_Moneris_Lang='"&pcPay_Moneris_Lang&"',pcPay_Moneris_Testmode="&pcPay_Moneris_Testmode&", pcPay_Moneris_CVVEnabled="&pcPay_Moneris_CVVEnabled&", pcPay_Moneris_Meth="&pcPay_Moneris_Meth&",pcPay_Moneris_Interac="&pcPay_Moneris_Interac &",pcPay_Moneris_Interac_MerchantID='"&pcPay_Moneris_Interac_MerchantID&"' WHERE pcPay_Moneris_Id=1;"

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",sslURL='"&mFileV&"', pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName=N'"&paymentNickName&"' WHERE gwCode=11"
	
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if pcPay_Moneris_Interac="1" then
		query="SELECT * FROM payTypes WHERE gwCode=66"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if rs.eof then
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('Moneris','gwMonerisInterac.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",66,N'"&paymentNickName2&"')"
		else
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentNickName=N'"&paymentNickName2&"' WHERE gwCode=66"
		end if
	else
		query="DELETE FROM payTypes WHERE gwCode=66"
	end if

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	set rs = nothing
	
end function

Function gwmoneris()
	varCheck=1
	'request gateway variables and insert them into the Moneris table
	pcPay_Moneris_StoreId=request.form("pcPay_Moneris_StoreId")
	pcPay_Moneris_StoreId=enDeCrypt(pcPay_Moneris_StoreId, scCrypPass)
	pcPay_Moneris_Key=request.Form("pcPay_Moneris_Key")
	pcPay_Moneris_Key=enDeCrypt(pcPay_Moneris_Key, scCrypPass)
	pcPay_Moneris_Interac_MerchantID = request.form("pcPay_Moneris_Interac_MerchantID")
	pcPay_Moneris_Interac_MerchantID=enDeCrypt(pcPay_Moneris_Interac_MerchantID, scCrypPass)
	pcPay_Moneris_TransType=request.form("pcPay_Moneris_TransType")
	if pcPay_Moneris_TransType="" then
		pcPay_Moneris_TransType="0"
	end if
	pcPay_Moneris_Lang=request.Form("pcPay_Moneris_Lang")
	pcPay_Moneris_TestMode=request.Form("pcPay_Moneris_TestMode")
	pcPay_Moneris_CVVEnabled=request.Form("pcPay_Moneris_CVVEnabled")
	pcPay_Moneris_Meth = request.form("pcPay_Moneris_Meth")
	pcPay_Moneris_Interac = request.form("pcPay_Moneris_Interac")
	if pcPay_Moneris_TestMode="" then
		pcPay_Moneris_TestMode=0
	end if
    pcPay_Moneris_Ver = Request.form("pcPay_Moneris_Ver")
	  if pcPay_Moneris_Ver ="US" Then
	  	mFileV = "gwMonerisUS.asp"
	  else
	  	mFileV = "gwMoneris2.asp" 
	 End if
	 
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
	paymentNickName2 = request.form("paymentNickName2")
	if paymentNickName2="" then
		paymentNickName2="Interac&reg; Online"
	end if
	err.clear
	err.number=0
	 

	query="UPDATE pcPay_Moneris SET pcPay_Moneris_StoreId='"&pcPay_Moneris_StoreId&"',pcPay_Moneris_Key='"&pcPay_Moneris_Key&"',pcPay_Moneris_TransType='"&pcPay_Moneris_TransType&"',pcPay_Moneris_Lang='"&pcPay_Moneris_Lang&"',pcPay_Moneris_Testmode="&pcPay_Moneris_Testmode&", pcPay_Moneris_CVVEnabled="&pcPay_Moneris_CVVEnabled&", pcPay_Moneris_Meth="&pcPay_Moneris_Meth&",pcPay_Moneris_Interac="&pcPay_Moneris_Interac &",pcPay_Moneris_Interac_MerchantID='"&pcPay_Moneris_Interac_MerchantID&"' WHERE pcPay_Moneris_Id=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Moneris','"&mFileV&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",11,N'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
    if pcPay_Moneris_Interac="1" then		
		query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('Moneris Interac&reg; Online','gwMonerisInterac.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",66,N'"&paymentNickName2&"')"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		set rs=nothing
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		end if
	set rs=nothing
	
end function

if request("gwchoice")="11" then

    tmp_id=request("id")
	tmp_mode=request("mode")

	'Check to see if fields exists in table, if not, add
	err.clear
		query="SELECT pcPay_Moneris_Meth FROM pcPay_Moneris"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		
		call closeDb()
response.redirect "upddbMoneris.asp?mode="&tmp_mode&"&id="&tmp_id
	else
		set rs=nothing
		
	end if
	if request("mode")="Edit" then
				query= "SELECT pcPay_Moneris_StoreId,pcPay_Moneris_Key, pcPay_Moneris_TransType, pcPay_Moneris_Lang, pcPay_Moneris_Testmode, pcPay_Moneris_CVVEnabled, pcPay_Moneris_Meth, pcPay_Moneris_Interac, pcPay_Moneris_Interac_MerchantID FROM pcPay_Moneris WHERE pcPay_Moneris_Id=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_Moneris_StoreId=rs("pcPay_Moneris_StoreId")
		pcPay_Moneris_StoreId=enDeCrypt(pcPay_Moneris_StoreId, scCrypPass)
		pcPay_Moneris_Key=rs("pcPay_Moneris_Key")
		if pcPay_Moneris_Key<>"" OR NOT isNull(pcPay_Moneris_Key) then
			pcPay_Moneris_Key=enDeCrypt(pcPay_Moneris_Key, scCrypPass)
		end if
		pcPay_Moneris_TransType=rs("pcPay_Moneris_TransType")
		pcPay_Moneris_Lang=rs("pcPay_Moneris_Lang")
		pcPay_Moneris_Testmode=rs("pcPay_Moneris_Testmode")
		pcPay_Moneris_CVVEnabled=rs("pcPay_Moneris_CVVEnabled")
		pcPay_Moneris_Meth = rs("pcPay_Moneris_Meth")
		pcPay_Moneris_Interac = rs("pcPay_Moneris_Interac")
		pcPay_Moneris_Interac_MerchantID = rs("pcPay_Moneris_Interac_MerchantID")
		if pcPay_Moneris_Interac_MerchantID<>"" OR NOT isNull(pcPay_Moneris_Interac_MerchantID) then
			pcPay_Moneris_Interac_MerchantID =enDeCrypt(pcPay_Moneris_Interac_MerchantID, scCrypPass)
		end if
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=11"
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


		query = "select paymentNickName from  payTypes where gwCode=66"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			
			call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		paymentNickName2 = rs("paymentNickName") 
		if paymentNickName2="" then
			paymentNickName2="INTERAC&reg; Online Accepted Here"
		end if
		set rs=nothing
		
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="11">
<!-- END MONERIS -->

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/moneris_logo.JPG" width="185" height="82"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
    
	<table width="100%" border="0" cellspacing="0" cellpadding="12">
        <tr>
            <td>
                <div class="bs-callout bs-callout-info">
                    <h4>Moneris</h4>    
                    <p>Moneris provides the solutions to process debit and credit card payments online in a secure, real-time environment. ProductCart is integrated with eSelect Plus Direct Post. Direct Post requires that the merchant have an SSL certificate. All the   cardholder information is taken on the merchant's website.</p>
                    <p>
                        <a class="btn btn-info btn-xs" href="http://eselectplus.moneris.com" target="_blank">Learn More</a>        
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
					<% dim pcPay_Moneris_StoreIdCnt,pcPay_Moneris_StoreIdEnd,pcPay_Moneris_StoreIdStart
                    pcPay_Moneris_StoreIdCnt=(len(pcPay_Moneris_StoreId)-2)
                    pcPay_Moneris_StoreIdEnd=right(pcPay_Moneris_StoreId,2)
                    pcPay_Moneris_StoreIdStart=""
                    if isNULL(pcPay_Moneris_StoreIdCnt) OR pcPay_Moneris_StoreIdCnt="" then
                        pcPay_Moneris_StoreIdCnt=1
                    end if
                    for c=1 to pcPay_Moneris_StoreIdCnt
                        pcPay_Moneris_StoreIdStart=pcPay_Moneris_StoreIdStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current DirectPost ID:&nbsp;<%=pcPay_Moneris_StoreIdStart&pcPay_Moneris_StoreIdEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;DirectPost ID&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;DirectPost ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="306"> <div align="right">DirectPost ID:</div></td>
                    <td width="651"> <input type="text" name="pcPay_Moneris_StoreId" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right">DirectPost Token:</div></td>
                    <td width="651"> <input type="text" name="pcPay_Moneris_Key" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="pcPay_Moneris_TransType">
                            <option value="1" <% if pcPay_Moneris_TransType="1" then %> selected <% end if %>>Sale</option>
                            <option value="0" <% if pcPay_Moneris_TransType="0" then %> selected <% end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
           
                    <td> <div align="right">Language:</div></td>
                    <td> <select name="pcPay_Moneris_Lang">
                            <option value="en-ca" selected>English</option>
                            <option value="fr-ca" <% if pcPay_Moneris_Lang="fr-ca" then%>selected<% end if %>>French</option>
                        </select> </td>
                </tr>
				<tr>
  					<td><div align="right">Country Version:</div></td>
                <td><input type="radio" class="clearBorder" name="pcPay_Moneris_Ver" value="CA" checked onClick="document.getElementById('interac').style.display='block'; ">
                  Canada
                    <input type="radio" class="clearBorder" name="pcPay_Moneris_Ver" value="US" <% if sslUrl="gwMonerisUS.asp" then%>checked<% end if %> onClick="document.getElementById('interac').style.display='none';document.getElementById('interYes').checked = false;document.getElementById('interNo').checked = true;">
                  US</td>
                </tr>
				 <tr>
  <td><div align="right">Response Method:</div></td>
                <td><input type="radio" class="clearBorder" name="pcPay_Moneris_Meth" value="1" checked>
                  Post
                  <input type="radio" class="clearBorder" name="pcPay_Moneris_Meth" value="0" <% if pcPay_Moneris_Meth=0 then%>checked<% end if %>>
                  Get</td>
                </tr>
				 <tr> 
				   <td><div align="right">Enable Test Mode:</div></td>
                    <td > <div align="left"><input name="pcPay_Moneris_TestMode" type="checkbox" class="clearBorder" value="1" <% if pcPay_Moneris_TestMode=1 then %>checked<% end if %> />
                                      <!--<b>Enable Live Mode </b><input name="pcPay_Moneris_TestMode" type="radio" class="clearBorder" value="1" <%' if pcPay_Moneris_TestMode=1 then %>checked<% 'end if %> />
                                      <b>Enable Test Mode </b>(Credit cards will not be charged) <input name="pcPay_Moneris_TestMode" type="radio" class="clearBorder" value="2" <% 'if pcPay_Moneris_TestMode=2 then %>checked<%' end if %> />
                                      <b>Enable INTERAC&reg; Certification Mode </b>--></div></td>
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
                                    Step 2: eSelect Plus eFraud Settings
                                </a>
                            </h4>
                        </div>
                        <div id="collapse2" class="panel-collapse collapse">
                            <div class="panel-body">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="2"><p>Moneris eSelect Plus eFraud settings will allow you to add an additional layer of fraud protection. You can now activate Card Validation Digits information and Address Verification Service for Moneris. In order to use the eFraud services you will need to contact the eSelect Plus Integration Support Team at <a href="mailto:eselectplus@moneris.com">eselectplus@moneris.com</a> to have the eFraud feature added to your profile.</p>
                    <br>
                    <p>To Acitvate eFraud for Moneris eSelect Plus, select Yes below. When activated, the customer will be prompted to enter in their credit card security code and Moneris will then authenticate the card information. </p>
                    <br> <p>In order to use the Verify By Visa services you will need to contact the eSelect Plus Integration Support Team at <a href="mailto:eselectplus@moneris.com">eselectplus@moneris.com</a> to have the Verify By Visa feature added to your profile.</p>
			</td>
                </tr>
              <tr>
  <td width="32%"><div align="right">Activate  eFraud for Moneris:</div></td>
                <td width="68%"><input type="radio" class="clearBorder" name="pcPay_Moneris_CVVEnabled" value="1" checked>
                  Yes
                  <input type="radio" class="clearBorder" name="pcPay_Moneris_CVVEnabled" value="0" <% if pcPay_Moneris_CVVEnabled=0 then%>checked<% end if %>>
                  No</td>
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
                                    Step 3: Enter INTERAC&reg; Online Settings 
                                </a>
                            </h4>
                        </div>
                        <div id="collapse3" class="panel-collapse collapse">
                            <div class="panel-body">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
			   <tr >
			     <td colspan="2">&nbsp;</td>
			     </tr>
			   <tr >
				   <td colspan="2">In order to use the Interac&reg; Online services you will need to contact the eSelect Plus Integration Support Team at <a href="mailto:eselectplus@moneris.com">eselectplus@moneris.com</a> to have the  Interac&reg; Online feature added to your profile.</td>
		      </tr>
                <% if request("mode")="Edit" then %>
					<% dim pcPay_Moneris_Interac_MerchantIDCnt,pcPay_Moneris_Interac_MerchantIDEnd,pcPay_Moneris_Interac_MerchantIDStart
                    pcPay_Moneris_Interac_MerchantIDCnt=(len(pcPay_Moneris_Interac_MerchantID)-2)
                    pcPay_Moneris_Interac_MerchantIDEnd=right(pcPay_Moneris_Interac_MerchantID,2)
                    pcPay_Moneris_Interac_MerchantIDStart=""
                    if isNULL(pcPay_Moneris_Interac_MerchantIDCnt) OR pcPay_Moneris_Interac_MerchantIDCnt="" then
                        pcPay_Moneris_Interac_MerchantIDCnt=1
                    end if
                    for c=1 to pcPay_Moneris_Interac_MerchantIDCnt
                        pcPay_Moneris_Interac_MerchantIDStart=pcPay_Moneris_Interac_MerchantIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Interac&reg; Online Merchant ID:&nbsp;<%=pcPay_Moneris_Interac_MerchantIDStart&pcPay_Moneris_Interac_MerchantIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Merchant Online ID&quot; 
                            is only partially shown on this page. </td>
                    </tr>
                <% end if %>
				 <tr  >
  <td width="32%"><div align="right">Activate  INTERAC&reg; Online:</div></td>
                <td width="68%"><input type="radio" class="clearBorder" name="pcPay_Moneris_Interac" id="interYes" value="1" <% if pcPay_Moneris_Interac=1 then%>checked<% end if %> >
                  Yes
                  <input type="radio" class="clearBorder" name="pcPay_Moneris_Interac" id="interNo"  value="0" <% if pcPay_Moneris_Interac=0 then%>checked<% end if %> >
                  No</td>
                </tr>
				 <tr >
				   <td><div align="right">INTERAC&reg; Online Merchant Number:</div></td>
				   <td><input type="text" name="pcPay_Moneris_Interac_MerchantID" size="20" value="" ></td>
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
                                    Step 4: You have the option to charge a processing fee for this payment option.
                                </a>
                            </h4>
                        </div>
                        <div id="collapse4" class="panel-collapse collapse">
                            <div class="panel-body">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapse6">
                                    Step 5: You can change the display name that is shown for this payment type.
                                </a>
                            </h4>
                        </div>
                        <div id="collapse6" class="panel-collapse collapse">
                            <div class="panel-body">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                                <a data-toggle="collapse" data-parent="#accordion" href="#collapse7">
                                    Step 6: Order Processing: Order Status and Payment Status.
                                </a>
                            </h4>
                        </div>
                        <div id="collapse7" class="panel-collapse collapse">
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
