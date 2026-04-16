<%
Dim pcPEYMerchantID,pcPEYMerchantToken,pcPEYAPIKey,pcPEYAPISKey,pcPEYMode,pcPEYTestMode,pcPEYTAToken
Dim pcPEYmsg
pcPEYMerchantID=""
pcPEYMerchantToken=""
pcPEYAPIKey=""
pcPEYAPISKey=""
pcPEYMode=""
pcPEYTestMode=""
pcPEYTAToken=""
pcPEYmsg=""

Function epoch2date(myEpoch)
	epoch2date = DateAdd("s", fix(myEpoch/1000), "01/01/1970 00:00:00")
End Function

Function date2epoch(myDate)
	date2epoch = DateDiff("s", "01/01/1970 00:00:00",myDate)*1000
End Function

Function GenNonce()
    Dim Tn1, w
    
	Tn1=""
	For w=1 to 19
		Randomize
		Tn1=Tn1 & Cstr(Fix(10*Rnd))
	Next
    
	GenNonce=Tn1
    
End Function

Sub getPayeezySettings()
Dim rs,query
'//Get the Admin Settings / Payeezy data
query="SELECT pcPEY_MerchantID,pcPEY_MToken,pcPEY_APIKey,pcPEY_APISKey,pcPEY_Mode,pcPEY_TestMode,pcPEY_JSKey,pcPEY_TAToken FROM pcPay_Payeezy;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

'// Set Admin Settings / Payeezy data
if not rs.eof then
	pcPEYMerchantID=rs("pcPEY_MerchantID")
	if pcPEYMerchantID<>"" then
		pcPEYMerchantID=enDeCrypt(pcPEYMerchantID, scCrypPass)
	end if
	pcPEYMerchantToken=rs("pcPEY_MToken")
	if pcPEYMerchantToken<>"" then
		pcPEYMerchantToken=enDeCrypt(pcPEYMerchantToken, scCrypPass)
	end if
	pcPEYAPIKey=rs("pcPEY_APIKey")
	if pcPEYAPIKey<>"" then
		pcPEYAPIKey=enDeCrypt(pcPEYAPIKey, scCrypPass)
	end if
	pcPEYAPISKey=rs("pcPEY_APISKey")
	if pcPEYAPISKey<>"" then
		pcPEYAPISKey=enDeCrypt(pcPEYAPISKey, scCrypPass)
	end if
	pcPEYJSKey=rs("pcPEY_JSKey")
	if pcPEYJSKey<>"" then
		pcPEYJSKey=enDeCrypt(pcPEYJSKey, scCrypPass)
	end if
	pcPEYMode=rs("pcPEY_Mode")
	if IsNull(pcPEYMode) OR (pcPEYMode="") then
		pcPEYMode=0
	end if
	pcPEYTestMode=rs("pcPEY_TestMode")
	if IsNull(pcPEYTestMode) OR (pcPEYTestMode="") then
		pcPEYTestMode=0
	end if
	pcPEYTAToken=rs("pcPEY_TAToken")
	if pcPEYTAToken<>"" then
		pcPEYTAToken=enDeCrypt(pcPEYTAToken, scCrypPass)
	else
		pcPEYTAToken=""
	end if
	if pcPEYTAToken="" OR pcPEYTestMode="1" then
		pcPEYTAToken="NOIW"
	end if
end if
set rs=nothing

End Sub

Function CaptureVoidPayeezy(tmpIdOrder,mType)
Dim rs,query,TransTag,TransID,OrderTotal,mStr,xml
Dim pcPEYStr,pcPEYnonce,pcPEYtimestamp,pcPEYData,pcPEYAPIURL,strStatus,strRetVal,HaveErrors

	CaptureVoidPayeezy=false
	TransTag=""
	TransID=""
	HaveErrors=0
	
	If pcPEYMerchantToken="" then
		call getPayeezySettings()
	end if

	query="SELECT total,gwAuthCode,gwTransID FROM Orders WHERE idOrder=" & tmpIdOrder & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		OrderTotal=rs("total")
		TransTag=rs("gwAuthCode")
		TransID=rs("gwTransID")
	end if
	set rs=nothing
	
	IF TransTag<>"" THEN

	if pcPEYTestMode="1" then
		pcPEYAPIURL="https://api-cert.payeezy.com/v1/transactions" & "/" & TransID
	else
		pcPEYAPIURL="https://api.payeezy.com/v1/transactions" & "/" & TransID
	end if
	
	pcPEYnonce=GenNonce()
	pcPEYtimestamp=date2epoch(UtcNow())

	pcPEYData="{" & VbLf
	pcPEYData=pcPEYData & "  ""merchant_ref"": ""OrdID" & (scpre+int(tmpIdOrder)) & """," & VbLf
	pcPEYData=pcPEYData & "  ""transaction_tag"": """ & TransTag & """," & vbLf
	if mType="1" then
		pcPEYData=pcPEYData & "  ""transaction_type"": ""capture""," & vbLf
	else
		pcPEYData=pcPEYData & "  ""transaction_type"": ""void""," & vbLf
	end if
	pcPEYData=pcPEYData & "  ""method"": ""credit_card""," & VbLf
	pcPEYData=pcPEYData & "  ""amount"": """ & Fix(OrderTotal*100) & """," & VbLf
	pcPEYData=pcPEYData & "  ""currency_code"": ""USD""" & VbLf
	pcPEYData=pcPEYData & "}"
	
	pcPEYStr=pcPEYAPIKey & pcPEYnonce & pcPEYtimestamp & pcPEYMerchantToken & pcPEYData
	
	Set sha256 = GetObject("script:" & Server.MapPath("../pc/sha256.txt"))

	pcPEYStr=sha256.b64_hmac_sha256(pcPEYAPISKey, pcPEYStr)
	
	Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
	
	xml.open "POST", pcPEYAPIURL, False
	xml.setRequestHeader "apikey", pcPEYAPIKey
	xml.setRequestHeader "token", pcPEYMerchantToken
	xml.setRequestHeader "Content-type", "application/json"
	xml.setRequestHeader "Authorization", pcPEYStr
	xml.setRequestHeader "nonce", pcPEYnonce
	xml.setRequestHeader "timestamp", pcPEYtimestamp
	
	xml.send pcPEYData
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText
	
	dim Info : set Info = JSON.parse(strRetVal)

	HaveErrors=0
	ErrorMsg=""
	
	TransID=""
	TransTag=""
	
	For Each Key in Info.keys()
		if UCase(key)="ERROR" then
			HaveErrors=1
		end if
		if UCase(key)="TRANSACTION_STATUS" then
			mStr=Info.transaction_status
		end if
	Next
	
	if (HaveErrors=0) then
		mStr=Info.transaction_status
		if UCase(trim(mStr))<>Ucase("Approved") then
			if pcPEYmsg<>"" then
				pcPEYmsg=pcPEYmsg & "<br>"
			end if
			pcPEYmsg=pcPEYmsg & "Order ID#: " & (scpre+int(tmpIdOrder)) & " error. Transaction Status: " & Ucase(mStr) & "<br>"
			CaptureVoidPayeezy=false
		else
			TransID=Info.transaction_id
			TransTag=Info.transaction_tag
			query="UPDATE pcPayeezyLogs SET pcPEYLg_Status=" & mType & ",pcPEYLg_TransID='" & TransID & "',pcPEYLg_TransTag='" & TransTag & "' WHERE idOrder=" & tmpIdOrder & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			
			if mType="1" then
				query="UPDATE Orders SET pcOrd_PaymentStatus=2 WHERE idOrder=" & tmpIdOrder & ";"
			else
				query="UPDATE Orders SET pcOrd_PaymentStatus=8 WHERE idOrder=" & tmpIdOrder & ";"
			end if
			set rs=connTemp.execute(query)
			set rs=nothing
			
			CaptureVoidPayeezy=true
		end if
	else
	  	CaptureVoidPayeezy=false
		if pcPEYmsg<>"" then
			pcPEYmsg=pcPEYmsg & "<br>"
		end if
		pcPEYmsg=pcPEYmsg & "Order ID#: " & (scpre+int(tmpIdOrder)) & " error.<ul>"
		For Each AField In Info.Error.messages.keys()
			pcPEYmsg=pcPEYmsg & "<li>" & Info.Error.messages.get(AField).description & "</li>"
		Next
		pcPEYmsg=pcPEYmsg & "</ul>"
	end if
	
	ELSE
		CaptureVoidPayeezy=false
		if pcPEYmsg<>"" then
			pcPEYmsg=pcPEYmsg & "<br>"
		end if
		pcPEYmsg=pcPEYmsg & "Order ID#: " & (scpre+int(tmpIdOrder)) & " error. Cannot find Payeezy Transaction Tag"
	END IF

End Function

Function UtcNow()
UtcNow = serverdate.toUTCString()
UtcNow = CDate(Replace(Right(UtcNow, Len(UtcNow) - Instr(UtcNow, ",")), "UTC", ""))
End Function
%>
<script language="JScript" runat="server">
var serverdate=new Date();
</script>