# MultipleMonitors.co.uk - Remove ProductCart Vendor Server Dependencies

## Background

This site runs **ProductCart v5.3.00** (classic ASP / VBScript) on Windows Server 2016 (AWS), with SQL Server database `stagemm`. Payments are handled by SagePay (completely independent of ProductCart licensing).

**NetSource Commerce** (the ProductCart vendor) is sunsetting the product and **shutting down all licensing/service infrastructure on January 11, 2027**. After that date, any code that phones home to their servers will timeout, causing delays or failures. See: `https://productcart-kb.nsource.com/article/productcart-sunset-faq/`

**Goal:** Remove ALL dependencies on vendor servers so the site operates indefinitely as a self-maintained installation. The customer-facing storefront, admin panel, and SagePay payment processing must all keep working.

**Key finding:** The customer-facing storefront has **zero** license phone-home calls. All external calls are in the admin panel and the shared `ErrorHandler.asp` include. When the vendor servers die:
- The storefront will still work but new sessions will have a ~2 sec delay (ErrorHandler.asp timeout)
- Admin logins will work but with ~30-60 sec timeout delays
- No functionality is permanently lost - all comm errors default to "pass"

---

## Files That Need Changes (9 external HTTP calls across 5 files)

| # | External Server URL | File | When Called | Priority |
|---|---|---|---|---|
| 1 | `https://www.productcart.com/verify/vCPRequest.asp` | AdminLoginInclude.asp | First admin login each day | HIGH |
| 2 | `http://www.productcart.com/verify/pcKeyVerifyURL2.asp` | AdminLoginInclude.asp | Admin login (URL mismatch) | HIGH |
| 3 | `ws.productcart.com/api/cfd2` | security.asp (called from AdminLoginInclude.asp) | Every admin login | HIGH |
| 4 | `service.productcartlive.com/antihack/getInjSigs.asp` | ErrorHandler.asp | Every new session | HIGH |
| 5 | `service.productcartlive.com/antihack/logAttack.asp` | ErrorHandler.asp | On detected attacks | MEDIUM |
| 6 | `service.productcartlive.com/antihack/config.asp` | ErrorHandler.asp | On first error per session | MEDIUM |
| 7 | `www.productcart.com/productcart-errorLog.asp` | ErrorHandler.asp | Anti-piracy trace (disabled) | LOW |
| 8 | `service.productcartlive.com/auth/*` | webservices.asp | Marketplace (disabled) | LOW |
| 9 | `service.productcartlive.com/v1/api/Clients` | productcartlive.asp | Admin "Check for Updates" | LOW |

## Files That Need NO Changes

- All customer storefront pages (`shop/pc/`)
- SagePay payment integration (`gwProtx.asp`)
- Admin session validation (`shop/130707/adminv.asp`) - purely local session checks
- All database operations - local SQL Server
- Email sending (`sendmail.asp`) - local SMTP
- Password hashing (`hash.aspx`) - local .NET service
- All settings files, status `.inc` files, theme files
- Apps in `shop/includes/apps/` - none are active

---

## STEP 0: BACKUP (DO THIS FIRST)

Before ANY changes, take a full backup of:
- The entire `shop/` directory
- The SQL Server database (`stagemm`)

---

## STEP 1: Replace AdminLoginInclude.asp (CRITICAL - MOST IMPORTANT)

### Files to modify:
- `shop/130707/AdminLoginInclude.asp`
- `shop/130707/login.asp`
- `shop/includes/pcSurlLvs.asp`

### Context

`AdminLoginInclude.asp` is encrypted with VBScript.Encode. It has been successfully decoded. The file is included from `login.asp` at line 75. `login.asp` line 1 declares `<%@ LANGUAGE = VBScript.Encode %>` which tells IIS to decode the included file at runtime.

The decoded file contains the admin login authentication logic plus 3 external HTTP calls. Below is the **complete decoded source** with annotations showing what to keep and what to remove.

### The admin login flow:
```
1. Call pcs_TestCompatibility() - LOCAL (keep)
2. Call pcs_UpgradeToHash() - LOCAL (keep)
3. Query DB for admin credentials
4. Verify password hash with pcf_CheckPassH() - LOCAL (keep)
5. Get last login date from pcStoreSettings
6. IF first login today:
     → Collect store telemetry data
     → POST to www.productcart.com/verify/vCPRequest.asp ← REMOVE
7. Normalize store URL, compare with stored fingerprint in pcSurlLvs.asp
8. IF URL matches:
     → Call pcs_updateDefinitions() ← REMOVE
     → Set session vars, redirect to menu.asp (keep)
9. ELSE (URL mismatch - CURRENT STATE):
     → POST to www.productcart.com/verify/pcKeyVerifyURL2.asp ← REMOVE
     → If PASS/comm error: set session vars, rewrite pcSurlLvs.asp, redirect
     → If FAIL: block login
```

### Complete decoded source of AdminLoginInclude.asp

The file between `<%` `%>` tags contains (with non-encoded `<!--#include file="../includes/sendAlarmEmail.asp" -->` lines between encoded blocks):

```asp
<%
Public Sub pcs_UpgradeToHash()
    Dim query, rs, tmpArr, i, tmpHash, intCount, tmpID

    '// Upgrade Passwords
	query="SELECT idAdmin, adminPassword FROM admins WHERE NOT (adminPassword LIKE 'NSPC:%');"
	set rs=connTemp.execute(query)
	if not rs.eof then

		tmpArr=rs.getRows()
		set rs=nothing
		intCount = ubound(tmpArr,2)
		For i=0 to intCount
			tmpID=tmpArr(0,i)
			tmpHash=tmpArr(1,i)
			if pcf_ValidPassH(tmpHash)=0 then
				tmpHash=enDeCrypt(tmpHash, scCrypPass)
				tmpHash=pcf_PasswordHash(tmpHash)
                If instr(tmpHash, "NSPC:")>0 Then
                    query="UPDATE admins SET adminPassword='" & tmpHash & "' WHERE idAdmin=" & tmpID & ";"
                    set rs=connTemp.execute(query)
                    set rs=nothing
                End If
			end if
		Next
	end if
	set rs=nothing

End Sub

'// Perform Compatibility Testing (During Every Login)
call pcs_TestCompatibility()

'// Upgrade Admin Passwords
call pcs_UpgradeToHash()

'// Authenticated and charge session
query="SELECT ID, IDAdmin, AdminLevel, adminPassword FROM admins WHERE idAdmin=" & pIdAdmin & " And AdminLevel<>'';"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number>0 then
	call closeDb()
	If (scSecurity=1) and (scAdminLogin=1) then
		if session("AttackCount")="" then
			session("AttackCount")=0
		end if
		session("AttackCount")=session("AttackCount")+1
		if session("AttackCount")>=scAttackCount then
			session("AttackCount")=0
			if (scAlarmMsg=1) then%>
			<!--#include file="../includes/sendAlarmEmail.asp" -->
            <%end if
			response.write dictLanguage.Item(Session("language")&"_security_2")
			response.end()
		end if
	End if
	response.redirect "msg.asp?message=1"
	response.end()
end if

if rstemp.eof then

	call closeDb()
	If (scSecurity=1) and (scAdminLogin=1) then
		if session("AttackCount")="" then
			session("AttackCount")=0
		end if
		session("AttackCount")=session("AttackCount")+1
		if session("AttackCount")>=scAttackCount then
			session("AttackCount")=0
			if (scAlarmMsg=1) then%>
			<!--#include file="../includes/sendAlarmEmail.asp" -->
            <%end if
			response.write dictLanguage.Item(Session("language")&"_security_2")
			response.end()
		end if
	End if
 	response.redirect "msg.asp?message=1"
	response.End()

else

	tmpHash = rstemp("adminPassword")
	tmpResult = pcf_CheckPassH(pAdminPassword, tmpHash)

	If Ucase(""&tmpResult)="TRUE" Then
    Else
		call closeDb()
		response.redirect "msg.asp?message=1"
		response.End()
	End If

	'// ============================================================
	'// EVERYTHING BELOW HERE UNTIL "Check that licensing is valid"
	'// IS THE TELEMETRY + LICENSE CHECK SECTION
	'// ============================================================

	Function encodeString(input)
		Dim newStr : newStr = ""
		for i = 1 to len(input)
			newStr = newStr & chr((asc(mid(input,i,1))+8))
		next
		encodeString = newStr
	End Function

	Function decodeString(input)
		Dim oldStr : oldStr = ""
		for i = 1 to len(input)
			oldStr = oldStr & chr((asc(mid(input,i,1))-8))
		next
		decodeString = oldStr
	End Function

	'// Get last login
	pcAdminLastLogin = DateValue(now)
	pcAdminCurrSign = ""

	query="SELECT pcStoreSettings_AdminLastLogin, pcStoreSettings_Cursign FROM pcStoreSettings"
	set rstemp2=server.CreateObject("ADODB.RecordSet")
	set rstemp2=conntemp.execute(query)
	If Not rstemp2.EOF Then
		pcAdminLastLogin = rstemp2(0)
		pcAdminCurrSign = rstemp2(1)
	End If
	Set rstemp2 = Nothing

	'// Update last login
	query="UPDATE pcStoreSettings SET pcStoreSettings_AdminLastLogin = '" & now & "'"
	conntemp.execute(query)

	Session("pcCPCheckCode") = ""
	Session("pcCPCheckText") = ""

	'// ============================================================
	'// REMOVE: Daily telemetry call (vCPRequest.asp)
	'// Everything from the DateDiff check through to storing
	'// pcCPCheckCode/pcCPCheckText in session
	'// ============================================================
	If DateDiff("d", DateValue(pcAdminLastLogin), DateValue(now)) <> 0 Then

        '// Get current url
		prot = "http"
	    https = lcase(request.ServerVariables("HTTPS"))
        If https <> "off" Then
            prot = "https"
        End if

        domainname = Request.ServerVariables("SERVER_NAME")
        filename = Request.ServerVariables("SCRIPT_NAME")
        querystring = Request.ServerVariables("QUERY_STRING")
        scCurrentURL = prot & "://" & domainname & filename

        If pcDCToggle <> "OFF" Then

            'get shipping providers
            scShipping = ""
            query="SELECT serviceDescription FROM shipService WHERE serviceActive <> 0 ORDER BY serviceDescription"
            set rstemp2=server.CreateObject("ADODB.RecordSet")
            set rstemp2=conntemp.execute(query)

            If Not rstemp2.EOF Then
                allData = rstemp2.GetRows()
                For i = 0 to UBound(allData, 2)
                    scShipping = scShipping &Trim(Replace(Replace(allData(0, i), "|", ""), "$", ""))
                    If i < UBound(allData, 2) Then scShipping = scShipping &"$$"
                Next
            End If

            set rstemp2 = Nothing

            'get payment gateways
            scGateways = ""
            query="SELECT paymentDesc FROM payTypes WHERE active = -1 ORDER BY paymentDesc"
            set rstemp2=server.CreateObject("ADODB.RecordSet")
            set rstemp2=conntemp.execute(query)

            If Not rstemp2.EOF Then
                allData = rstemp2.GetRows()
                For i = 0 to UBound(allData, 2)
                    scGateways = scGateways &Trim(Replace(Replace(allData(0, i), "|", ""), "$", ""))
                    If i < UBound(allData, 2) Then scGateways = scGateways &"$$"
                Next
            End If

            set rstemp2 = Nothing

            'get sales totals
            lastOrderTotal = 0
            lastLineItemTotal = 0
            lastItemTotal = 0
            lastSalesTotal = 0

            lastMonthStart = Month(DateAdd("m", -1, now)) &"/1/" &Year(DateAdd("m", -1, now))
            lastMonthEnd = DateAdd("m", 1, lastMonthStart)

            '# orders
            query="SELECT Count(*) As TotalOrders, Sum(orders.total) FROM orders WHERE ((orders.orderStatus>=2 AND orders.orderStatus<5) OR (orders.orderStatus>=6)) AND orderdate>='" & lastMonthStart & "' AND orderdate<'" & lastMonthEnd & "';"
            set rs=connTemp.execute(query)
            If Not rs.EOF Then
                lastOrderTotal = rs(0)
                lastSalesTotal = rs(1)
            End If
            Set rs = Nothing

            '# line items
            query="SELECT Count(idProductOrdered) As TotalLineItems, SUM(ProductsOrdered.quantity) AS TotalProducts FROM ProductsOrdered INNER JOIN orders ON orders.idOrder = ProductsOrdered.idOrder WHERE ((orders.orderStatus>=2 AND orders.orderStatus<5) OR (orders.orderStatus>=6)) AND orderdate>='" & lastMonthStart & "' AND orderdate<'" & lastMonthEnd & "';"
            set rs=connTemp.execute(query)
            If Not rs.EOF Then
                lastLineItemTotal = rs(0)
                lastItemTotal = rs(1)
            End If
            Set rs = Nothing

        End If

		If lastOrderTotal = "" Then lastOrderTotal = 0
		If lastLineItemTotal = "" Then lastLineItemTotal = 0
		If lastItemTotal = "" Then lastItemTotal = 0
		If lastSalesTotal = "" Then lastSalesTotal = 0
		If lastMonthStart = "" Then lastMonthStart = "1/1/1999"

		'// Get store owner email
		pcOwnerEmail = ""
		query="SELECT ownerEmail FROM emailSettings"
		set rstemp2=server.CreateObject("ADODB.RecordSet")
		set rstemp2=conntemp.execute(query)
		If Not rstemp2.EOF Then
            pcOwnerEmail = rstemp2(0)
		End If
		Set rstemp2 = Nothing

        'must match pc.com!
        keyStr = "KJ4H87BN24G781CZER99PL3837HMN1R83JKQWW28JND9PL32XC87HJE8398KMQYT"

		If IsNull(scVersion) Then scVersion = ""
		If IsNull(scSubVersion) Then scSubVersion = ""
		If IsNull(scCrypPass) Then scCrypPass = ""
		If IsNull(scDB) Then scDB = ""
		If IsNull(scStoreURL) Then scStoreURL = ""
		If IsNull(scCurrentURL) Then scCurrentURL = ""
		If IsNull(scCompanyName) Then scCompanyName = ""
		If IsNull(scShipping) Then scShipping = ""
		If IsNull(lastMonthStart) Then lastMonthStart = ""
		If IsNull(lastOrderTotal) Then lastOrderTotal = ""
		If IsNull(lastLineItemTotal) Then lastLineItemTotal = ""
		If IsNull(lastItemTotal) Then lastItemTotal = ""
		If IsNull(lastSalesTotal) Then lastSalesTotal = ""
		If IsNull(scGateways) Then scGateways = ""
		If IsNull(pcOwnerEmail) Then pcOwnerEmail = ""
		If IsNull(pcAdminCurrSign) Then pcAdminCurrSign = ""

	    stext = Replace(scVersion, "|", "")
        stext = stext & "|" & Replace(scSubVersion, "|", "")
        stext = stext & "|" & Replace(scCrypPass, "|", "")
        stext = stext & "|" & Replace(scDB, "|", "")
        stext = stext & "|" & Replace(scStoreURL, "|", "")
        stext = stext & "|" & Replace(scCurrentURL, "|", "")
        stext = stext & "|" & Replace(scCompanyName, "|", "")
        stext = stext & "|" & scShipping
        stext = stext & "|" & Replace(lastMonthStart, "|", "")
        stext = stext & "|" & Replace(lastOrderTotal, "|", "")
        stext = stext & "|" & Replace(lastLineItemTotal, "|", "")
        stext = stext & "|" & Replace(lastItemTotal, "|", "")
        stext = stext & "|" & Replace(lastSalesTotal, "|", "")
        stext = stext & "|" & scGateways
        stext = stext & "|" & Replace(pcOwnerEmail, "|", "")
        stext = stext & "|" & Replace(pcAdminCurrSign, "|", "")
        stext = stext & "|" & Replace(Request.ServerVariables("HTTP_USER_AGENT"), "|", "")
		stext = SafeEncrypt(stext, keyStr)

		'Send the transaction info as part of the querystring
		set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
		xml.open "POST", "https://www.productcart.com/verify/vCPRequest.asp", false
		xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=ISO-8859"
		xml.send "vCPPostData=" & stext

		if intCommunicationErr=0 then
			strStatus = xml.Status
			if strStatus<>200 then
				'There is a communication issue
				Set xml = Nothing
			else
				'store the response
				strRetVal = xml.responseText
				Set xml = Nothing
				strArrayVal = split(strRetVal, "|", -1)
				pcCPCheckCode = strArrayVal(0)
				pcCPCheckText = strArrayVal(1)

				Session("pcCPCheckCode") = pcCPCheckCode
				Session("pcCPCheckText") = pcCPCheckText
			end if
		else
			'There is a communication issue
		end if
	End If
	'// ============================================================
	'// END OF TELEMETRY BLOCK TO REMOVE
	'// ============================================================


	'// ============================================================
	'// REMOVE: URL check + license verification + pcSurlLvs rewrite
	'// Replace with direct session setup and redirect
	'// ============================================================

	'// Check that licensing is valid
	pcv_strcheck=lcase(scStoreURL)
	pcv_strcheck=replace(pcv_strcheck,".","")
	pcv_strcheck=replace(pcv_strcheck,"http://","")
	pcv_strcheck=replace(pcv_strcheck,"https://","")
	pcv_strcheck=replace(pcv_strcheck,":","")
	pcv_strcheck=replace(pcv_strcheck,"\","")
	pcv_strcheck=replace(pcv_strcheck,"/","")
	pcv_strcheck=replace(pcv_strcheck,"y","")
	pcv_strcheck=replace(pcv_strcheck,"x","")
	pcv_strcheck=replace(pcv_strcheck,"z","")

	pcv_strcheck2=pcv_SURLResponse
	pcv_strcheck2=decodeString(pcv_strcheck2)

	dim intCommunicationErr
	intCommunicationErr=0

	if lcase(pcv_strcheck)=lcase(pcv_strcheck2) AND pcv_ITCResponse="0" then

        '// Update Definitions
        Call pcs_updateDefinitions()

        session("admin")=-1
		session("IDAdmin")=rstemp("ID")
		session("CUID")=rstemp("IDAdmin")
		session("PmAdmin")=rstemp("AdminLevel")
		session("admin." & pcf_getAdminToken()) = Session.SessionID

		call closedb()
		if Session("RedirectURL")<>"" then
			RedirectURL=Session("RedirectURL")
			Session("RedirectURL")=""
			response.redirect RedirectURL
		else
			response.redirect "menu.asp"
		end if
	else

		'// Send detection message and update IE
		stext="pcVersion="&scVersion
		stext=stext & "&pcSubVersion=" & scSubVersion
		stext=stext & "&pcKeyId=" & scCrypPass
		stext=stext & "&pcDBType=" & scDB
		stext=stext & "&pcStoreURL=" & scStoreURL
		stext=stext & "&pcStorePWD=" & pcv_StorePWD
		stext=stext & "&pcStoreUID=" & pcv_StoreUID
		stext=stext & "&pcSessionID=" & session.SessionID
		stext=stext & "&pcRegisteredNumber=" & scRegistered
		stext=stext & "&pcCompanyName=" & scCompanyName
		stext=stext & "&pcCompanyAddress=" & scCompanyAddress
		stext=stext & "&pcCompanyZip=" & scCompanyZip
		stext=stext & "&pcCompanyCity=" & scCompanyCity
		stext=stext & "&pcCompanyState=" & scCompanyState
		stext=stext & "&pcCompanyCountry=" & scCompanyCountry
		stext=stext & "&pcAlertType=ALTERURL"
		if pcv_ITCResponse="" then
			stext=stext & "&pcFlag=YES"
		else
			stext=stext & "&pcFlag=NO"
		end if

		'// Send the transaction info as part of the querystring
		set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
		xml.open "POST", "http://www.productcart.com/verify/pcKeyVerifyURL2.asp?"& stext & "", false
		xml.send ""

		'// Check for connection issues
		if err.number<>0 then
			'Check for Communication Error
			intCommunicationErr=1
			Set xml = Nothing
		end if

		if intCommunicationErr=0 then
			strStatus = xml.Status
			if strStatus<>200 then
				'There is a communication issue - Register the cart as is.
				intCommunicationErr=1
				Set xml = Nothing
			else
				'store the response
				strRetVal = xml.responseText
				Set xml = Nothing
				strArrayVal = split(strRetVal, "|", -1)
				pcv_Status = strArrayVal(0)
			end if
		else
			'There is a communication issue - Register the cart as is.
			pcv_Status="PASS"
		end if

		if pcv_Status="FAIL" then
			call closedb()
			response.write strArrayVal(3)
			response.write pcvErrorMessageContact
		end if

		if pcv_Status="PASS" OR pcv_Status="SILENTPASS" then

            session("admin")=-1
			session("IDAdmin")=rstemp("ID")
			session("CUID")=rstemp("IDAdmin")
			session("PmAdmin")=rstemp("AdminLevel")
			session("admin." & pcf_getAdminToken()) = Session.SessionID

			call closedb()
			'// write pcSurlLvs.asp file
			pcv_strNew=encodeString(pcv_strcheck)
			'//Overwrite existing file
			Dim objFS
			Dim objFile

			Set objFS = Server.CreateObject ("Scripting.FileSystemObject")

			if PPD="1" then
				pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/pcSurlLvs.asp")
			else
				pcStrFileName=Server.Mappath ("../includes/pcSurlLvs.asp")
			end if

			Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
			objFile.WriteLine CHR(60)&CHR(37)& vbCrLf
			objFile.WriteLine "private const pcv_SURLResponse = """&pcv_strNew&"""" & vbCrLf
			objFile.WriteLine "private const pcv_ITCResponse = """&intCommunicationErr&"""" & vbCrLf
			objFile.WriteLine CHR(37)&CHR(62)& vbCrLf

			objFile.Close
			set objFS=nothing
			set objFile=nothing

			if Session("RedirectURL")<>"" then
				RedirectURL=Session("RedirectURL")
				Session("RedirectURL")=""
			else
				RedirectURL = "menu.asp"
			end if

			if intCommunicationErr=1 OR strArrayVal(4)="AUTOREDIRECT" OR pcv_Status="SILENTPASS" then
				response.redirect RedirectURL
			else
				response.write strArrayVal(3)&"<br /><br /><br />"
				response.write "<a href="""&RedirectURL&""">"&strArrayVal(4)&"</a>"
			end if
		end if
	end if
end if

Function SafeEncrypt(stext, keyStr)
    pc_CodePage = Session.CodePage
    Session.CodePage = 1252
    stext = Encrypt(stext, keyStr)
    stext = Server.URLEncode(stext)
    Session.CodePage = pc_CodePage
    SafeEncrypt = stext
End Function

Private Function pcf_getAdminToken()
	pcv_strLocalAddress = Request.ServerVariables("LOCAL_ADDR")
	pcv_strLocalSessionID = Session.SessionID
	pcv_strAdminToken = pcv_strLocalAddress & "." & pcv_strLocalSessionID
	pcf_getAdminToken = pcv_strAdminToken
End Function
%>
```

### What AdminLoginInclude.asp should become after modification

Replace the entire encrypted file with this plaintext version. This keeps all local authentication logic and removes all 3 external HTTP calls:

```asp
<%
Public Sub pcs_UpgradeToHash()
    Dim query, rs, tmpArr, i, tmpHash, intCount, tmpID

    '// Upgrade Passwords
	query="SELECT idAdmin, adminPassword FROM admins WHERE NOT (adminPassword LIKE 'NSPC:%');"
	set rs=connTemp.execute(query)
	if not rs.eof then

		tmpArr=rs.getRows()
		set rs=nothing
		intCount = ubound(tmpArr,2)
		For i=0 to intCount
			tmpID=tmpArr(0,i)
			tmpHash=tmpArr(1,i)
			if pcf_ValidPassH(tmpHash)=0 then
				tmpHash=enDeCrypt(tmpHash, scCrypPass)
				tmpHash=pcf_PasswordHash(tmpHash)
                If instr(tmpHash, "NSPC:")>0 Then
                    query="UPDATE admins SET adminPassword='" & tmpHash & "' WHERE idAdmin=" & tmpID & ";"
                    set rs=connTemp.execute(query)
                    set rs=nothing
                End If
			end if
		Next
	end if
	set rs=nothing

End Sub

'// Perform Compatibility Testing (During Every Login)
call pcs_TestCompatibility()

'// Upgrade Admin Passwords
call pcs_UpgradeToHash()

'// Authenticated and charge session
query="SELECT ID, IDAdmin, AdminLevel, adminPassword FROM admins WHERE idAdmin=" & pIdAdmin & " And AdminLevel<>'';"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number>0 then
	call closeDb()
	If (scSecurity=1) and (scAdminLogin=1) then
		if session("AttackCount")="" then
			session("AttackCount")=0
		end if
		session("AttackCount")=session("AttackCount")+1
		if session("AttackCount")>=scAttackCount then
			session("AttackCount")=0
			if (scAlarmMsg=1) then%>
			<!--#include file="../includes/sendAlarmEmail.asp" -->
            <%end if
			response.write dictLanguage.Item(Session("language")&"_security_2")
			response.end()
		end if
	End if
	response.redirect "msg.asp?message=1"
	response.end()
end if

if rstemp.eof then

	call closeDb()
	If (scSecurity=1) and (scAdminLogin=1) then
		if session("AttackCount")="" then
			session("AttackCount")=0
		end if
		session("AttackCount")=session("AttackCount")+1
		if session("AttackCount")>=scAttackCount then
			session("AttackCount")=0
			if (scAlarmMsg=1) then%>
			<!--#include file="../includes/sendAlarmEmail.asp" -->
            <%end if
			response.write dictLanguage.Item(Session("language")&"_security_2")
			response.end()
		end if
	End if
 	response.redirect "msg.asp?message=1"
	response.End()

else

	tmpHash = rstemp("adminPassword")
	tmpResult = pcf_CheckPassH(pAdminPassword, tmpHash)

	If Ucase(""&tmpResult)="TRUE" Then
    Else
		call closeDb()
		response.redirect "msg.asp?message=1"
		response.End()
	End If

	'// Update last login timestamp
	query="UPDATE pcStoreSettings SET pcStoreSettings_AdminLastLogin = '" & now & "'"
	conntemp.execute(query)

	Session("pcCPCheckCode") = ""
	Session("pcCPCheckText") = ""

	'// Set admin session variables and redirect
	session("admin")=-1
	session("IDAdmin")=rstemp("ID")
	session("CUID")=rstemp("IDAdmin")
	session("PmAdmin")=rstemp("AdminLevel")
	session("admin." & pcf_getAdminToken()) = Session.SessionID

	call closedb()
	if Session("RedirectURL")<>"" then
		RedirectURL=Session("RedirectURL")
		Session("RedirectURL")=""
		response.redirect RedirectURL
	else
		response.redirect "menu.asp"
	end if

end if

Private Function pcf_getAdminToken()
	pcv_strLocalAddress = Request.ServerVariables("LOCAL_ADDR")
	pcv_strLocalSessionID = Session.SessionID
	pcv_strAdminToken = pcv_strLocalAddress & "." & pcv_strLocalSessionID
	pcf_getAdminToken = pcv_strAdminToken
End Function
%>
```

### Also modify login.asp line 1

**File:** `shop/130707/login.asp`

Change line 1 from:
```asp
<%@ LANGUAGE = VBScript.Encode %>
```
To:
```asp
<%@ LANGUAGE = VBScript %>
```

### Also update pcSurlLvs.asp (defensive)

**File:** `shop/includes/pcSurlLvs.asp`

This file is still included via `<!--#include file="../includes/pcSurlLvs.asp" -->` in other places. Set to empty values so the constants are defined but unused:

```asp
<%
private const pcv_SURLResponse = ""
private const pcv_ITCResponse = "0"
%>
```

---

## STEP 2: Modify ErrorHandler.asp (HIGHEST CUSTOMER-FACING IMPACT)

**File:** `shop/includes/ErrorHandler.asp`

This file is included via `common.asp` (line 22) on **every page load** - both storefront and admin.

### 2a. Replace `checkSigs()` function (lines ~16-85)

**Current behaviour:** On new session, calls `https://service.productcartlive.com/antihack/getInjSigs.asp` to fetch SQL injection signatures. Has a 2-second timeout. On failure, sets signatures to a nonsense default string. Also calls `logAttack.asp` when an injection is detected.

**New behaviour:** Set `session("pcInjectionStringsQP")` and `session("pcInjectionStringsFP")` directly with hardcoded injection signatures. Keep the session caching, scanning logic, and banning behaviour. Remove all HTTP calls.

Hardcoded signatures to use (using the existing `^*` delimiter):
```
QueryString patterns (pcInjectionStringsQP):
UNION SELECT^*UNION ALL SELECT^*' OR 1=1^*' OR '1'='1^*'; DROP ^*'; DELETE ^*xp_cmdshell^*EXEC(^*EXECUTE(^*sp_executesql^*INTO OUTFILE^*INTO DUMPFILE^*LOAD_FILE^*BENCHMARK(^*SLEEP(^*WAITFOR DELAY^*<script^*javascript:^*onload=^*onerror=^*eval(^*expression(^*url(^*import(

Form/POST patterns (pcInjectionStringsFP):
UNION SELECT^*UNION ALL SELECT^*'; DROP ^*'; DELETE ^*xp_cmdshell^*EXEC(^*EXECUTE(^*sp_executesql^*INTO OUTFILE^*INTO DUMPFILE^*LOAD_FILE^*<script^*javascript:^*onload=^*onerror=
```

### 2b. Remove Netsparker remote logging (lines ~239-253)

Remove the `MSXML2.serverXMLHTTP` call to `logAttack.asp` in the Netsparker detection block. Keep the local session ban + redirect.

### 2c. Replace `antiInjection()` function (lines ~257-334)

Remove HTTP calls to `config.asp` (rate-limit config) and `logAttack.asp`. Use hardcoded defaults: `scPCDefLimit=10` (max errors), `scPCDefTimespan=1` (minute). Keep all local banning logic.

### 2d. Neutralize `TraceInit()` (lines ~336-413)

Replace with:
```asp
Sub TraceInit()
    Call TraceClose()
End Sub
```
This is anti-piracy tracing controlled by `scEnforceAdmin` (undefined = disabled). Gutting it removes any risk.

### 2e. Neutralize `TraceXML()` (lines ~485-500)

Replace with:
```asp
Sub TraceXML()
    Call TraceClose()
End Sub
```
This phones home to `www.productcart.com/productcart-errorLog.asp`.

**Total: 6 external HTTP calls removed from this file.**

---

## STEP 3: No-op `pcs_updateDefinitions()` in security.asp

**File:** `shop/includes/coreMethods/security.asp` (lines ~408-527)

**Current behaviour:** Calls `http://ws.productcart.com/api/cfd2` sending the license key, receives XML security definitions, and updates the `pcDefinitions` database table.

**Called from:** AdminLoginInclude.asp on every login (removed in Step 1), and `upddb*.asp` database upgrade scripts.

**New behaviour:** Replace the function body with an early exit:
```asp
Public Sub pcs_updateDefinitions()
    '// Remote definition updates disabled - vendor servers decommissioned
    '// Current definitions (build 6) are already stored in pcDefinitions table
    Exit Sub
End Sub
```

---

## STEP 4: Blank vendor URLs in webservices.asp

**File:** `shop/includes/coreMethods/webservices.asp` (lines 5-11)

These services are already disabled (`scPCWS_IsActive = "0"` in settings), but as a defensive measure:

**Current:**
```asp
Const pcv_tokeURL = "https://service.productcartlive.com/auth/oauth/token"
Const pcv_baseURL = "https://service.productcartlive.com/auth/api"
Const pcv_marketURL = "https://service.productcartlive.com/v1/"
```

**New:**
```asp
Const pcv_tokeURL = ""
Const pcv_baseURL = ""
Const pcv_marketURL = ""
```

Also add early-exit guards at the top of `pcf_GetToken()`, `pcf_VerifyClaim()`, and `pcf_VerifyClaimByCode()`:
```asp
If pcv_tokeURL = "" Then Exit Function
```
(or similar, returning empty/false as appropriate for each function)

---

## STEP 5: Disable remote call in productcartlive.asp

**File:** `shop/130707/productcartlive.asp` (lines ~15-91)

This is an admin-only "Check for Updates" page.

**Current behaviour:** Builds a JSON payload with package ID and email, POSTs to `http://service.productcartlive.com/v1/api/Clients`, parses response to determine add-on status.

**New behaviour:** Replace the HTTP call section with hardcoded values (all add-ons disabled, matching current state):
```asp
IsApparel = False
IsConfig = False
IsConfigPlus = False
IsQBWC = False

'// Remote update check disabled - vendor servers decommissioned
'// Add-on status preserved at current values (all disabled)

call pcf_upgradeDowngrade(IsApparel, IsConfig, IsConfigPlus, IsQBWC)
```

Remove lines 19-91 (the pcv_baseURL, JSON build, HTTP POST, response parsing, and for-each loop). Keep the `pcf_upgradeDowngrade` function definition (lines 98-132) as-is.

---

## Implementation Order

1. **Step 0** - Full backup (shop/ directory + database)
2. **Step 1** - AdminLoginInclude.asp + login.asp + pcSurlLvs.asp (test admin login immediately after)
3. **Step 3** - No-op pcs_updateDefinitions() in security.asp (small, safe)
4. **Step 4** - Blank webservices.asp URLs (already disabled, defensive)
5. **Step 5** - Disable productcartlive.asp remote call (admin-only page)
6. **Step 2** - ErrorHandler.asp (biggest change, affects all pages - do last so other changes are verified first)

---

## Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| Admin login breaks | Low | High | Step 1 most critical. Test immediately. Backup allows instant rollback. |
| Injection protection weaker | Low | Medium | Hardcoded signatures cover all common SQL injection and XSS patterns |
| Storefront affected | Very Low | High | Storefront has NO external calls - Step 2 only removes remote parts |
| pcSurlLvs.asp still referenced | Very Low | Low | File kept with empty values for compatibility |
| Database corruption | None | N/A | No database schema changes. Only one UPDATE query preserved (last login timestamp) |

---

## Verification Checklist

After all changes:
1. **Customer storefront:** Browse products, add to cart, proceed through checkout to SagePay
2. **Admin login:** Log in at `/shop/130707/login_1.asp` - must work instantly (no timeout delay)
3. **Admin functions:** Create/edit products, view orders, manage categories
4. **IIS logs:** Check for any failed HTTP calls to `productcart.com`, `productcartlive.com`, or `ws.productcart.com`
5. **Performance:** New sessions should load instantly (no 2-sec timeout). Admin login should be instant (no 30-sec timeout)
6. **Injection protection:** Test with a `' OR 1=1` query string parameter - should trigger ban/redirect
7. **Rollback:** If anything fails, restore from Step 0 backup
