<%

if pcgwAuthCode<>"" AND isNULL(pcgwAuthCode)=False then
	
	Dim trxAmt, trxAuthCode, trxType, trxTender, trxTypeId, trxCardType, trxCardNum, trxExpDate, trxGatewayEnabled
	
	trxAmt = amt
	trxAuthCode = pcgwAuthCode

    If varGWInfo="P" Then
        query="SELECT amt, tender, trxtype, origid, captured FROM pfporders WHERE idOrder = " & pidorder & ";"
        set rs=server.CreateObject("ADODB.RecordSet")
	    set rs=conntemp.execute(query)
	    if not rs.eof then
		    trxAmt=rs("amt")
		    trxTender=rs("tender")
		    trxAuthCode=rs("origid")
		    trxTypeId=rs("trxtype")
		    captured=rs("captured")
		    gwCode=2

            If trxTender = "C" Then trxTender = "CC"
	    end if
	    set rs = nothing
    Else
	    query="SELECT amount, paymentmethod, transtype, authcode, captured, fraudcode, gwCode FROM pcPay_PFL_Authorize WHERE idOrder = " & pidorder & ";"
	    set rs=server.CreateObject("ADODB.RecordSet")
	    set rs=conntemp.execute(query)
	    if not rs.eof then
		    trxAmt=rs("amount")
		    trxTender=rs("paymentmethod")
		    trxAuthCode=rs("authcode")
		    trxTypeId=rs("transtype")
		    captured=rs("captured")
		    fraudcode=rs("fraudcode")
		    gwCode=rs("gwCode")
	    end if
	    set rs = nothing
    End If

	'// Check if the payment gateway used for the transaction is still enabled
	If gwCode > 0 Then
		trxGatewayEnabled = false
		query = "SELECT idPayment FROM payTypes WHERE gwCode = " & gwCode & ";"
		set rs=conntemp.execute(query)
		if not rs.eof then
			trxGatewayEnabled = true
		end if
		set rs = nothing
	Else
		trxGatewayEnabled = true
	End If

	'// Get information on this transaction from the API
	If trxGatewayEnabled Then
        If gwCode = 2 Then
            query="SELECT v_Partner, v_Vendor, v_User, v_Password,pfl_testmode,pfl_transtype,pfl_CSC FROM verisign_pfp WHERE id=1;"
			set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

		    pcPay_PayPal_Partner=rs("v_Partner")
		    pcPay_PayPal_Vendor=rs("v_Vendor")		
		    pcPay_PayPal_Username=rs("v_User")				
		    pcPay_PayPal_Password=rs("v_Password")				
		    pcPay_PayPal_TransType=rs("pfl_transtype")			
		    pcPay_PayPal_CVC=rs("pfl_CSC")
		    pcPay_PayPal_Sandbox=rs("pfl_testmode")

            if pcPay_PayPal_Sandbox="YES" then
			    pcPay_PayPal_Method = "sandbox"
		    else
			    pcPay_PayPal_Method = "live"
		    end if

            if pcPay_PayPal_CVC="YES" then
			    pcPay_PayPal_CVC = 1
		    else
			    pcPay_PayPal_CVC = 0
		    end if
        Else
		    '//Retrieve any gateway specific data from database or hard-code the variables
		    query="SELECT pcPay_PayPal_Partner, pcPay_PayPal_Vendor, pcPay_PayPal_Username, pcPay_PayPal_Password, pcPay_PayPal_TransType, pcPay_PayPal_CVC, pcPay_PayPal_Sandbox FROM pcPay_PayPal WHERE pcPay_PayPal_ID=1;"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=connTemp.execute(query)
	
		    if err.number<>0 then
			    call LogErrorToDatabase()
			    set rs=nothing
		
			    call closeDb()
			    response.redirect "techErr.asp?err="&pcStrCustRefID
		    end if
	
		    '// Set gateway specific variables
		    pcPay_PayPal_Partner=rs("pcPay_PayPal_Partner")
		    pcPay_PayPal_Vendor=rs("pcPay_PayPal_Vendor")		
		    pcPay_PayPal_Username=rs("pcPay_PayPal_Username")				
		    pcPay_PayPal_Password=rs("pcPay_PayPal_Password")				
		    pcPay_PayPal_TransType=rs("pcPay_PayPal_TransType")			
		    pcPay_PayPal_CVC=rs("pcPay_PayPal_CVC")
		    pcPay_PayPal_Sandbox=rs("pcPay_PayPal_Sandbox")
		    set rs = nothing
        End If

		If pcPay_PayPal_Username&""="" Then
			pcPay_PayPal_Username = pcPay_PayPal_Vendor
		End If
	
		'get transaction information
		nvpstr = "TRXTYPE=I"
		nvpstr = nvpstr &"&TENDER=C"
		nvpstr = nvpstr &"&PARTNER="& pcPay_PayPal_Partner
		nvpstr = nvpstr &"&VENDOR="& pcPay_PayPal_Vendor
		nvpstr = nvpstr &"&USER="& pcPay_PayPal_Username
		nvpstr = nvpstr &"&PWD="& pcPay_PayPal_Password
		nvpstr = nvpstr &"&ORIGID="& trxAuthCode
		nvpstr = nvpstr &"&VERBOSITY=HIGH"

		'Send the transaction info as part of the querystring
		set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
		'SB S
		if pcPay_PayPal_Sandbox = "1" then
			xml.open "POST", "https://pilot-payflowpro.paypal.com", false
		else
			xml.open "POST", "https://payflowpro.paypal.com", false
		end if
	
		xml.Send nvpstr
		strStatus = xml.Status
	
		'store the response
		strRetVal = xml.responseText
		Set xml = Nothing
	
		split_resultXML = split(strRetVal,"&")
		j=0
		for each item in split_resultXML
			split_param = split(split_resultXML(j),"=")
			formname = split_param(0)
			formvalue = split_param(1)
			if ucase(formname)  = "RESULT" then tmpRESULT = formvalue
			if ucase(formname)  = "PNREF" then tmpPNREF = formvalue
			if ucase(formname)  = "TRANSSTATE" then tmpTRANSSTATE = formvalue
			if ucase(formname)  = "TRANSTIME" then tmpTRANSTIME = formvalue
			if ucase(formname)  = "AMT" then tmpAMT = formvalue
			if ucase(formname)  = "ACCT" then tmpCARDNUM = formvalue
			if ucase(formname)  = "EXPDATE" then tmpEXPDATE = formvalue
			if ucase(formname)  = "CARDTYPE" then tmpCARDTYPE = formvalue
			if ucase(formname)	= "FIRSTNAME" then tmpFIRSTNAME = formvalue
			if ucase(formname)	= "LASTNAME" then tmpLASTNAME = formvalue
			j = j + 1
		next
	
		'// Set the Payor Name
		If tmpFIRSTNAME <> "" And tmpLASTNAME <> "" Then
			trxPayorName = tmpFIRSTNAME & " " & tmpLASTNAME
		End If

		'// Setup Card Number
		If tmpCARDNUM <> "" Then
			trxCardNum = "****" & tmpCARDNUM
		End If
	
		'// Setup Card Expiration Date
		If tmpEXPDATE <> "" Then
			trxExpDate = Left(tmpEXPDATE, 2) & "/" & Right(tmpEXPDATE, 2)
		End If
	
		'// Generate Card Type string
		Select Case tmpCARDTYPE
		Case "0"
			trxCardType = "Visa"
		Case "1"
			trxCardType = "Master Card"
		Case "2"
			trxCardType = "Discover"
		Case "3"
			trxCardType = "American Express"
		Case "4"
			trxCardType = "Diner's Club"
		Case "5"
			trxCardType = "JCB"
		End Select

		'// Set Transaction Date string
		If IsDate(tmpTRANSTIME) Then
			trxTime = FormatDateTime(tmpTRANSTIME)
		End If
	End If
	
	'// Generate Payment Method string
	If trxTender = "CC" Then
		paymentMethod = "Credit Card"
	Else
		paymentMethod = "PayPal"
	End if
	
	'// Generate Transaction Type string
	Select Case trxTypeId
	Case "D"
		trxType = "Delayed Capture"
	Case "A"
		trxType = "Authorization"
	Case "S"
		trxType = "Sale"
	Case "V"
		trxType = "Void"
	Case "C"
		trxType = "Refund"
	End Select
		
	'// Generate Status Note for PPA
	If pcv_PaymentStatus = 6 And trxType = "Refund" Then
		pcv_PPAStatusNote="The payment for this order has been refunded. "
	ElseIf pcv_PaymentStatus = 8 And trxType = "Void" Then
		pcv_PPAStatusNote="The payment for this order has been voided. "	
	End If
	%>
    
    <% if Len(paymentMethod) > 1 then %>
      <tr>
        <td colspan="2">
          Payment Type: <strong><%=paymentMethod%></strong>
        </td>
      </tr> 
    <% end if %>
    
    <% If trxTypeId="A" And (Cdbl(trxAmt)<>Cdbl(tmpAMT)) Then %>
    <tr>
      <td colspan="2">
        Authorized amount: <strong><%=scCurSign&money(tmpAMT) %></strong>
      </td>
    </tr>
    <% End If %>
    
    <tr>
      <td colspan="2">
        Payment Amount: <strong><%=scCurSign&money(trxAmt) %></strong>
      </td>
    </tr>

		<% If Len(trxPayorName) > 0 Then %>
			<tr>
				<td colspan="2">
					Payor Name: <strong><%= trxPayorName %></strong>
				</td>
			</tr>    
		<% End If %>

		<% if Len(trxCardType) > 0 then %>
      <tr>
        <td colspan="2">
          Card Type: <strong><%=trxCardType%>&nbsp;<%= trxCardNum %></strong>
        </td>
      </tr>
    <% end if %>
    <% if Len(trxExpDate) > 1 then %>
      <tr>
        <td colspan="2">
          Expiration Date: <strong><%=trxExpDate%></strong>
        </td>
      </tr>
    <% end if %>
    
    <tr>
      <td colspan="2" class="pcCPspacer"></td>
    </tr>	
    <tr>
      <td colspan="2">
        <span class="pcCPsectionTitle">Last Transaction Details</span>
      </td>
    </tr> 
    <tr>
      <td colspan="2">Transaction ID: <strong><%= trxAuthCode %></strong></td>
    </tr>	
    
    <% if Len(trxType) > 1 then %>
      <tr>
        <td colspan="2">
          Transaction Type: <strong><%=trxType%></strong>
        </td>
      </tr>
    <% end if %>
    <% if Len(trxTime) > 1 then %>
      <tr>
        <td colspan="2">
          Transaction Date: <strong><%=trxTime%></strong>
        </td>
      </tr>
    <% end if %>
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: PAYPAL - Display Risk Managment if its available.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If isNULL(pcv_strAVSRespond)=True Then pcv_strAVSRespond=""
If isNULL(pcv_strCVNResponse)=True Then pcv_strCVNResponse=""
%>
<% 
if (pcv_strAVSRespond<>"" OR pcv_strCVNResponse<>"") then 
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>	
<tr>
<td colspan="2">
  <span class="pcCPsectionTitle">Risk Management</span>
</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<% end if %>
<% if pcv_strAVSRespond<>"" then %>
<tr>
	<td colspan="2">
		<%
		select case pcv_strAVSRespond
            case "N": pcv_strAVSRespond="No Match"
			case "Y": pcv_strAVSRespond="Match"
			case else: pcv_strAVSRespond="Not Available"						
		end select							
		%>
		AVS Response: <strong><%=pcv_strAVSRespond%></strong> 
	</td>
</tr>
<% end if %>

<% if pcv_strCVNResponse<>"" then %>
<tr>
	<td colspan="2">
		<%
		select case pcv_strCVNResponse
			case "N": pcv_strCVNResponse="No Match"
            case "Y": pcv_strCVNResponse="Match"			
			case else: pcv_strCVNResponse="Not Available"							
		end select							
		%>
		CVV2 Response: <strong><%=pcv_strCVNResponse%></strong>
	</td>
</tr>
<% end if %>

<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: PAYPAL - Display Risk Managment if its available.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
    
      
		<% If pcv_PaymentStatus <> 8 And pcv_PaymentStatus <> 6 And trxGatewayEnabled Then %>
      <tr>
        <td colspan="2" class="pcCPspacer"></td>
      </tr>	
    
      <tr>
        <td colspan="2">
          <span class="pcCPsectionTitle">PayPal Actions</span>
        </td>
      </tr> 
        
      <tr>
        <td colspan="2">
          <input type="hidden" value="<%= gwCode %>" name="SubmitPFApiGWCode" />

          <div style="padding-bottom:4px;">
        	<% if trxType = "Authorization" And pcv_PaymentStatus <> 2 then
          	PayPalbtns=1 %>
            <input type="submit" name="SubmitPFApi" value=" Capture "  class="btn btn-primary">&nbsp;&nbsp;
          <% end if %>
          
          <% if trxType = "Delayed Capture" or trxType = "Sale" then
						PayPalbtns=1 %>
            <input type="submit" name="SubmitPFApi" value=" Refund "  onClick="javascript: if (confirm('This action will NOT cancel the order, it will refund the payment via PayPal. Are you sure you want to continue?')) return true ; else return false ;" class="btn btn-primary">&nbsp;&nbsp;	
          <% end if %>
          
          <% if trxType = "Authorization" Or (trxType = "Delayed Capture" and captured <> 1) then
          	PayPalbtns=1 %>												
            <input type="submit" name="SubmitPFApi" value="  Void  " onClick="javascript: if (confirm('This action will cancel the PayPal authorization and mark the order status as canceled.  Are you sure you want to continue?')) return true ; else return false ;" class="btn btn-primary">					
          <% end if %>
          </div>
          
					<%if PayPalbtns=1 then%>							
          <div style="padding:4px;">						
            <a class="pcCPhelp" href="helpOnline.asp?ref=442">Help with these buttons</a>
          </div>
          <%else%>
          <div style="padding-bottom:4px;">						
            no actions available
          </div>
          <%end if%>	
        </td>
      </tr>
      <tr>
        <td colspan="2" class="pcCPspacer"></td>
      </tr>	
    <% end if %>
    
		<% if pcv_PPAStatusNote<>"" then %>
			<tr>
				<td colspan="2">
					<div style="padding-bottom:4px;" class="pcCPnotes">
					<strong>Note</strong>: <%=pcv_PPAStatusNote %>
          </div>
				</td>
			</tr>
		<% end if %>
    
<% 
end if 
%>