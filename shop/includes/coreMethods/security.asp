<%
Function openDefenderDB()
    On Error Resume Next
    Set connTempPD = server.createobject("adodb.connection")
    connTempPD.Open scDSN  
End Function


Function closeDefenderDB()
    On Error Resume Next
    connTempPD.close
    Set connTempPD = nothing
    err.clear
End Function


Public Function pcf_getTruePath()
	SPath1=Request.ServerVariables("PATH_INFO")
	mycount1=0
	do while mycount1<2
		if mid(SPath1,len(SPath1),1)="/" then
			mycount1=mycount1+1
		end if
		if mycount1<2 then
			SPath1=mid(SPath1,1,len(SPath1)-1)
		end if
	loop
	if Ucase(Request.ServerVariables("HTTPS"))="ON" then
		SPathInfo="https://"
	else
		SPathInfo="http://"
	end if
	pcf_getTruePath=SPathInfo & Request.ServerVariables("HTTP_HOST") & SPath1
End Function


Public Sub pcs_TestCompatibility()

    '// Test for .NET, SSL, and Web.config Error(s)  
    pcv_boolIsSafeSSL = pcf_IsSafeSSL()
    If (pcv_boolIsSafeSSL = False) Then
        set rs=nothing
        response.Redirect("msg.asp?message=47")
        response.End()
    End If    
    pcv_boolIsNET = False
    tmpResult = pcf_PasswordHash("!2500LmB!..")
    If (instr(tmpResult, "NSPC") = 0) Then
        response.Redirect("msg.asp?message=47")
        response.End()
    End If
        
End Sub


Public Function pcf_IsSafeSSL()
    On Error Resume Next
    
    Dim objXML
    Dim SPathInfo

	SPathInfo = pcf_getTruePath()	
	if Right(SPathInfo, 1)="/" then
		SPathInfo=SPathInfo & "pc/service/api/hash.aspx"
	else
		SPathInfo=SPathInfo & "/pc/service/api/hash.aspx"
	end if
    
    pcv_strTemp = pcf_PostForm("ac=G&reqPW=" & server.urlencode("test"), SPathInfo, "")

	If Err.Number <> 0 OR pcv_strTemp="" Then
        pcf_IsSafeSSL = False
    Else
        pcf_IsSafeSSL = True
    End If
    Err.Clear()
	
	Set objXML = nothing
End Function



Public Sub pcs_checkFailedPaymentAttempts(idCustomer)
    If scSecurity=1 Then
		query="SELECT pcCust_FailedPaymentCount FROM customers WHERE idCustomer = " & idCustomer
		Set rs = server.CreateObject("ADODB.RecordSet")
		Set rs = conntemp.execute(query)
		If err.number <> 0 Then
			call LogErrorToDatabase()
			Set rs = Nothing
			call closedb()
			response.redirect "techErr.asp?err= " & pcStrCustRefID
		End If
		
		pcv_FailedPaymentCount = rs("pcCust_FailedPaymentCount")
		
		Set rs = Nothing

		If scGWSecurity = 1 And (pcv_FailedPaymentCount >= (scGWLockAttempts)) Then
            call pcs_lockCustomerAccount(idCustomer)
            session("SFClearCartURL") = "msg.asp?message=56" '// Redirect to display message after logout
            response.Redirect("CustLO.asp") '// Log out
        End If
    End If
End Sub

Public Sub pcs_lockCustomerAccount(idCustomer)
    query="UPDATE customers SET pcCust_Locked = 1 WHERE idCustomer = " & idCustomer
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	If err.number <> 0 Then
  	    call LogErrorToDatabase()
  	    Set rs = Nothing
  	    call closedb()
  	    response.redirect "techErr.asp?err= " & pcStrCustRefID
    End If
	Set rs = Nothing  
End Sub

Public Sub pcs_LogTransaction(idCustomer, orderId, isSuccess)
	Dim pcv_DateTime, pcv_CustIpAddress
	datetime = Now()
	if SQL_Format="1" then
		pcv_DateTime=Day(pcv_DateTime)&"/"&Month(pcv_DateTime)&"/"&Year(pcv_DateTime)
	else
		pcv_DateTime=Month(pcv_DateTime)&"/"&Day(pcv_DateTime)&"/"&Year(pcv_DateTime)
	end if	
	pcv_DateTime=pcv_DateTime & " " & Time()
	
	pcv_CustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If pcv_CustIpAddress="" Then pcv_CustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
    pcv_strGateWayId = Session("DefaultIdPayment")
    If pcv_strGateWayId="" Then
        pcv_strGateWayId = pcGatewayDataIdOrder
    End If
    
	query = "INSERT INTO pcTransactionLogs (datetime, IP, customerId, orderId, isSuccess, gatewayId) VALUES ('" & datetime & "', '" & pcv_CustIpAddress & "', " & idCustomer &", " & orderId & ", " & isSuccess & ", " & pcv_strGateWayId & ")"

	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	If err.number <> 0 Then
		call LogErrorToDatabase()
		Set rs = Nothing
		call closedb()
		response.redirect "techErr.asp?err= " & pcStrCustRefID
	End If
	Set rs = Nothing
	
	If isSuccess = 0 Then
		pcs_countFailedPaymentAttempt(idCustomer)
	End If
End Sub

Public Sub pcs_countFailedPaymentAttempt(idCustomer)
	query="UPDATE customers SET pcCust_FailedPaymentCount = pcCust_FailedPaymentCount + 1 WHERE idCustomer = " & idCustomer
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	If err.number <> 0 Then
		call LogErrorToDatabase()
		Set rs = Nothing
		call closedb()
		response.redirect "techErr.asp?err= " & pcStrCustRefID
	End If
	Set rs = Nothing
	
	pcs_checkFailedPaymentAttempts(idCustomer)
End Sub

Public Sub pcs_clearFailedPaymentAttempt(idCustomer)
  query="UPDATE customers SET pcCust_FailedPaymentCount = 0 WHERE idCustomer = " & idCustomer
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	If err.number <> 0 Then
		call LogErrorToDatabase()
		Set rs = Nothing
		call closedb()
		response.redirect "techErr.asp?err= " & pcStrCustRefID
	End If
	Set rs = Nothing
End Sub
%>



<%
'// START: ProductCart Defender
Public Function pcf_LoadDefinitions()
    On Error Resume Next

    Dim connTempPD
    
    Set connTempPD = server.createobject("adodb.connection")
    connTempPD.Open scDSN
    
    Dim securityObj
    
    If len(Session("DefinitionsObj"))=0 Then
    
        Set securityObj = JSON.parse("{}")

        '// All Active Rules by Priority
        query="SELECT pcDef_Desc, pcDef_Pattern, pcDef_Replace, pcDef_IgnoreCase, pcDef_IsGlobal, pcDef_ContinueOnError, pcDef_Type, pcDef_Priority, pcDef_Active, pcDef_Key FROM pcDefinitions WHERE pcDef_Active=1 Order By pcDef_Priority Asc"
        Set rsDef = Server.CreateObject("ADODB.Recordset")
        set rsDef = connTempPD.execute(query)
        If Not rsDef.Eof Then

            Set definitionObj = JSON.parse("{}") 
            pcv_arrayDefinition = rsDef.getRows()
            pcv_intDefinitionTotal = UBound(pcv_arrayDefinition, 2)
            For definitionCounter = 0 to pcv_intDefinitionTotal
                
                '// Add Definition Item
                set definitionItem = JSON.parse("{}")   
                definitionItem.Set "description", pcv_arrayDefinition(0, definitionCounter)
                 
                definitionItem.Set "pattern", pcv_arrayDefinition(1, definitionCounter) 
                definitionItem.Set "replaceString", pcv_arrayDefinition(2, definitionCounter) 
                definitionItem.Set "ignoreCase", pcf_IsTrueorFalse(pcv_arrayDefinition(3, definitionCounter))
                definitionItem.Set "isGlobal", pcf_IsTrueorFalse(pcv_arrayDefinition(4, definitionCounter) )    
                definitionItem.Set "continueOnError", pcf_IsTrueorFalse(pcv_arrayDefinition(5, definitionCounter))
                definitionItem.Set "type", pcv_arrayDefinition(6, definitionCounter)
                definitionItem.Set "priority", pcv_arrayDefinition(7, definitionCounter)
                definitionItem.Set "isActive", pcf_IsTrueorFalse(pcv_arrayDefinition(8, definitionCounter))
                definitionItem.Set "key", pcv_arrayDefinition(9, definitionCounter)
                definitionObj.Set definitionCounter, definitionItem
                
            Next
            securityObj.set "definitions", definitionObj
            Set definitionObj = Nothing
            
        End If    
        Set rsDef = Nothing

        Dim pcv_strSecurityJSON
        pcv_strSecurityJSON = JSON.stringify(securityObj, null, 2)  
        Session("DefinitionsObj") = pcv_strSecurityJSON

    Else
    
        Set securityObj = JSON.parse(Session("DefinitionsObj"))
      
    End If
    
    connTempPD.close
    Set connTempPD = nothing
    err.clear

    Set pcf_LoadDefinitions = securityObj

End Function


Public Function pcf_IsBuildExpired(pcv_intBuild)
    On Error Resume Next
    If cdbl(pcv_intBuild) > cdbl(scDefenderBuild) Then
        pcf_IsBuildExpired = True
    Else
        pcf_IsBuildExpired = False
    End If
End Function


Public Function pcf_IsTrueorFalse(switch)
    On Error Resume Next
    If switch = 1 Then
        pcf_IsTrueorFalse = true
    Else
        pcf_IsTrueorFalse = false 
    End If
End Function


Public Function pcf_IsOnOrOff(switch)
    On Error Resume Next
    If switch = true Then
        pcf_IsOnOrOff = 1
    Else
        pcf_IsOnOrOff = 0 
    End If
End Function


Public Function pcf_LoadDefaultDefinitions()
    On Error Resume Next
    
    Dim securityObj
    
    if PPD="1" then
        pcStrFolder = "/"&scPcFolder&"/includes/library"
    else
        pcStrFolder = "../includes/library"
    end if

    pcv_strDefaultDefinitions = pcf_OpenUTF8(pcStrFolder & "\definitions.asp", pcStrFolder & "\definitions.asp")
    
    If len(pcv_strDefaultDefinitions)=0 Then
        pcStrFolder = "../../library"
        pcv_strDefaultDefinitions = pcf_OpenUTF8(pcStrFolder & "\definitions.asp", pcStrFolder & "\definitions.asp")
    End If

    Session("DefinitionsObj") = pcv_strDefaultDefinitions
    Set securityObj = JSON.parse(Session("DefinitionsObj"))

    Set pcf_LoadDefaultDefinitions = securityObj
    
End Function


Public Sub pcs_UpdateDefinition(pcv_strDesc, pcv_strPattern, pcv_strReplaceString, pcv_intIgnoreCase, pcv_intIsGlobal, pcv_intContinueOnError, pcv_strType, pcv_intPriority, pcv_intActive, pcv_strKey)
    On Error Resume Next
    
    Dim connTempPD
    
    Set connTempPD = server.createobject("adodb.connection")
    connTempPD.Open scDSN
    
    '// Database
    pcv_boolIsExists = False
    query="SELECT pcDef_Id FROM pcDefinitions WHERE [pcDef_Key] = '" & pcv_strKey & "'"
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = connTempPD.execute(query)
    If Not rs.Eof Then
        pcv_boolIsExists = True
    End If
	Set rs = Nothing  
    
    If pcv_boolIsExists = True Then
        query = "UPDATE pcDefinitions SET pcDef_Desc = '" & pcv_strDesc & "', "
        query = query & "pcDef_Pattern = '" & pcv_strPattern & "', "
        query = query & "pcDef_Replace = '" & pcv_strReplaceString & "', "
        query = query & "pcDef_IgnoreCase = " & pcv_intIgnoreCase & ", "
        query = query & "pcDef_IsGlobal = " & pcv_intIsGlobal & ", "
        query = query & "pcDef_ContinueOnError = " & pcv_intContinueOnError & ", "
        query = query & "pcDef_Type = '" & pcv_strType & "', "
        query = query & "pcDef_Priority = " & pcv_intPriority & ", "
        query = query & "pcDef_Active = " & pcv_intActive & " "
        query = query & "WHERE pcDef_Key = '" & pcv_strKey & "' "
        Set rs = server.CreateObject("ADODB.RecordSet")
        Set rs = connTempPD.execute(query)
        Set rs = Nothing  
    Else
        query = "INSERT INTO pcDefinitions (pcDef_Desc, pcDef_Pattern, "
        query = query & "pcDef_Replace, pcDef_IgnoreCase, pcDef_IsGlobal, "
        query = query & "pcDef_ContinueOnError, pcDef_Type, pcDef_Priority, pcDef_Active, pcDef_Key) "
        query = query & "VALUES ("
        query = query & "'" & pcv_strDesc & "', "
        query = query & "'" & pcv_strPattern & "', "
        query = query & "'" & pcv_strReplaceString & "', "
        query = query & "" & pcv_intIgnoreCase & ", "
        query = query & "" & pcv_intIsGlobal & ", "
        query = query & "" & pcv_intContinueOnError & ", "
        query = query & "'" & pcv_strType & "', "
        query = query & "" & pcv_intPriority & ", "
        query = query & "" & pcv_intActive & ", "
        query = query & "'" & pcv_strKey & "' "
        query = query & ") "
        Set rs = server.CreateObject("ADODB.RecordSet")
        Set rs = connTempPD.execute(query)
        Set rs = Nothing  
    End If
    
    connTempPD.close
    Set connTempPD = nothing
    err.clear

End Sub


Sub pcf_SaveDefaultDefinitions(cfuResult)
    On Error Resume Next
    
    '// Text File
    Dim strtext1
    
    if PPD="1" then
        pcStrFolder = "/"&scPcFolder&"/includes/library"
    else
        pcStrFolder = "../includes/library"
    end if

    strtext1 = trim(cfuResult)
    
    call pcs_SaveUTF8(pcStrFolder & "\definitions.asp", pcStrFolder & "\definitions.asp", strtext1)

End Sub


Sub pcf_SaveDefenderBuild(pcv_intBuild)
    On Error Resume Next
    
    '// Text File
    Dim strtext1
    
    if PPD="1" then
        pcStrFolder = "/"&scPcFolder&"/includes"
    else
        pcStrFolder = "../includes"
    end if

    strtext1 = trim("<" & Chr(37) & vbNewLine & "private const scDefenderBuild =" & pcv_intBuild & " " & vbNewLine & Chr(37) & ">")
    
    call pcs_SaveUTF8(pcStrFolder & "\defenderSettings.asp", pcStrFolder & "\defenderSettings.asp", strtext1)

End Sub


Public Sub pcs_updateDefinitions()
    On Error Resume Next
    
    IsApparel = False
    IsConfig = False
    IsConfigPlus = False
    
    pcv_baseURL = "http://ws.productcart.com/api/cfd2"

    '// START: QUICKBOOKS
    qbsv = ""
    Set fs=server.CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(Server.Mappath("QB_Version.asp")) Then
        Set qbf = fs.OpenTextFile(Server.Mappath("QB_Version.asp"), 1, false)
        Do Until qbf.AtEndOfStream
            tempData = qbf.ReadLine
            If InStr(1, tempData, "scQuickBooksVersion", 1) > 0 Then
                tempDataArr = Split(tempData, "=")
                If UBound(tempDataArr, 1) > 0 Then qbsv = Trim(Replace(tempDataArr(1), chr(34), ""))
                Exit Do
            End If
        Loop	
        Set qbf = Nothing
    End If
    '// END: QUICKBOOKS

    '// START: JSON
    dim jsonService : set jsonService = JSON.parse("{}")
        
    jsonService.set "username", scVersion 
    jsonService.set "password", ""
    jsonService.set "domain", ""
    jsonService.set "licenseKey", scCrypPass 
    jsonService.set "storeVersion", scVersion 
    jsonService.set "subVersion", scSubVersion 
    jsonService.set "servicePack", scSP 
    jsonService.set "qbVersion", qbsv 
    jsonService.set "referrer", currentURL
    jsonService.set "themeFolder", scThemePath
    jsonService.set "directory", ""
    jsonService.set "packageId", ""
    jsonService.set "StatusCode", ""
    jsonService.set "MessageUID", ""
    jsonService.set "ServerName", ""
    jsonService.set "DatabaseUsername", ""
    jsonService.set "DatabasePassword", ""
    jsonService.set "DatabaseName", ""
    jsonService.set "DataFileName", ""
    jsonService.set "DataPathName", ""
    jsonService.set "DataFileSize", ""
    jsonService.set "DataFileMaxSize", ""
    jsonService.set "DataFileGrowth", ""
    jsonService.set "LogFileName", ""
    jsonService.set "LogPathName", ""
    jsonService.set "LogFileGrowth", ""
    jsonService.set "EmailPartner", ""
    jsonService.set "KeyID", ""
        
    Dim jsonObj
    jsonObj = JSON.stringify(jsonService, null, 2)
    '// END: JSON
    
    
    '// START: POST
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "POST", pcv_baseURL, false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    objXMLhttp.send jsonObj
    cfuResult = objXMLhttp.responseText
    
    Set objXMLhttp = Nothing
    '// END: POST


    '// START: PARSE RESULTS
    dim Info : set Info = JSON.parse(cfuResult)
    cfuResult = JSON.stringify(Info, null, 2)  

    pcv_intBuild = Info.information.build

    If len(pcv_strIsOverRide)=0 Then
        pcv_strIsOverRide = False
    End If


    If (pcf_IsBuildExpired(pcv_intBuild) And pcf_IsDbScriptRun()) Or (pcv_strIsOverRide) Then

        pcv_intKeyCounter = 0
        for each key in Info.definitions.keys()
            pcv_intKeyCounter = pcv_intKeyCounter + 1
            If Not Info.definitions.get(key) Is Nothing Then
    
                pcv_strKeyId = Info.definitions.get(key).key        
                pcv_strDesc = Info.definitions.get(key).description 
                pcv_strPattern = Info.definitions.get(key).pattern
                pcv_strReplaceString = Info.definitions.get(key).replaceString
                pcv_intIgnoreCase = pcf_IsOnOrOff(Info.definitions.get(key).ignoreCase)
                pcv_intIsGlobal = pcf_IsOnOrOff(Info.definitions.get(key).isGlobal)
                pcv_intContinueOnError = pcf_IsOnOrOff(Info.definitions.get(key).continueOnError)
                pcv_strType = Info.definitions.get(key).type 
                pcv_intPriority = Info.definitions.get(key).priority
                pcv_intActive = pcf_IsOnOrOff(Info.definitions.get(key).isActive)
                
                pcv_strKeyId = replace(pcv_strKeyId, "'","''")     
                pcv_strDesc = replace(pcv_strDesc, "'","''")   
                pcv_strPattern = replace(pcv_strPattern, "'","''")  
                pcv_strReplaceString = replace(pcv_strReplaceString, "'","''")  
                pcv_strType = replace(pcv_strType, "'","''")  
             
                call pcs_UpdateDefinition(pcv_strDesc, pcv_strPattern, pcv_strReplaceString, pcv_intIgnoreCase, pcv_intIsGlobal, pcv_intContinueOnError, pcv_strType, pcv_intPriority, pcv_intActive, pcv_strKeyId) 
                
            End If
        next
        call pcf_SaveDefaultDefinitions(cfuResult)
        call pcf_SaveDefenderBuild(pcv_intBuild)
    End If
    '// END: PARSE RESULTS

End Sub 

Public Function pcf_IsDbScriptRun()
    On Error Resume Next     
    Dim connTempPD 
    Set connTempPD = server.createobject("adodb.connection")
    connTempPD.Open scDSN
    err.clear    
    query="SELECT pcDef_Key FROM pcDefinitions"
    Set rsDef = Server.CreateObject("ADODB.Recordset")
    set rsDef = connTempPD.execute(query)    
    Set rsDef = Nothing    
    If err.number <>  0 Then
        pcf_IsDbScriptRun = False
    Else
        pcf_IsDbScriptRun = True
    End If  
    connTempPD.close
    Set connTempPD = nothing
    err.clear     
End Function
'// END: ProductCart Defender
%>




<%
'// Password Hash Functions

Public Function pcf_ValidPassH(tmpvalue)
Dim tmp1
	if tmpvalue<>"" then
		tmp1=split(tmpvalue,":")
		if ubound(tmp1)<>2 then
			pcf_ValidPassH=0
		else
			if (ubound(tmp1)=2) AND (tmp1(2)<>"") AND (tmp1(0)="NSPC") then
				pcf_ValidPassH=1
			else
				pcf_ValidPassH=0
			end if
		end if
	else
		pcf_ValidPassH=0
	end if
End Function

Public Function pcf_CheckNewPassH(tmpid,tmpemail)
Dim query,rs,tmpHash,tmpResult

	pcf_CheckNewPassH=0
	
	if tmpid<>"" then
		query="SELECT idcustomer,[password] FROM Customers WHERE idCustomer=" & tmpid & ";"
	else
		query="SELECT idcustomer,[password] FROM Customers WHERE email LIKE '" & tmpemail & "';"
	end if
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpHash=rs("password")
		pcf_CheckNewPassH=pcf_ValidPassH(tmpHash)
	end if
	set rs=nothing

End Function


Public Function pcf_PasswordHash(tmpvalue)
    Dim objXML
    Dim SPath1,mycount1,SPathInfo

	SPathInfo = pcf_getTruePath()	
	if Right(SPathInfo,1)="/" then
		SPathInfo=SPathInfo & "pc/service/api/hash.aspx"
	else
		SPathInfo=SPathInfo & "/pc/service/api/hash.aspx"
	end if

    pcf_PasswordHash = pcf_PostForm("ac=G&reqPW=" & server.urlencode(tmpvalue), SPathInfo, "")

End Function



Public Function pcf_CheckPassH(tmpvalue,tmphash)
    Dim objXML
    Dim SPath1,mycount1,SPathInfo

	SPathInfo = pcf_getTruePath()	
	if Right(SPathInfo,1)="/" then
		SPathInfo=SPathInfo & "pc/service/api/hash.aspx"
	else
		SPathInfo=SPathInfo & "/pc/service/api/hash.aspx"
	end if
    SPathInfo = replace(SPathInfo, "includes/apps/", "")

	pcf_CheckPassH = pcf_PostForm("ac=T&reqPW=" & server.urlencode(tmpvalue) & "&hash=" & server.urlencode(tmphash), SPathInfo, "")

End Function

Public Function pcf_CheckUsedPassH(tmpid, tmpemail, tmpvalue)
    Dim rs, query, tmpArr, intCount, i, tmphash, tmpCustID, tmpResult

	pcf_CheckUsedPassH = 0
	
    IF scCheckSamePass = "1" THEN
	
        if tmpid<>"" then
            query="SELECT idcustomer FROM Customers WHERE idCustomer=" & tmpid & ";"
        else
            query="SELECT idcustomer FROM Customers WHERE email LIKE '" & tmpemail & "';"
        end if
        tmphash=""
        set rs=connTemp.execute(query)
        if not rs.eof then
            tmpCustID=rs("idcustomer")
        end if
        set rs=nothing
        
        query="SELECT idCustomer, pcUP_UsedPass FROM pcUsedPassHistory WHERE idCustomer=" & tmpCustID & ";"
        set rs=connTemp.execute(query)
        if not rs.eof then
            tmpArr=rs.getRows()
            set rs=nothing
            intCount=ubound(tmpArr,2)
            For i=0 to intCount
                tmphash=tmpArr(1,i)
                if tmphash<>"" then
                    tmpResult=pcf_CheckPassH(tmpvalue,tmphash)
                    if Ucase(""&tmpResult)="TRUE" then
                        pcf_CheckUsedPassH=1
                        exit function
                    end if
                end if
            Next
        end if
        set rs=nothing
	
    END IF
	
End Function

Public Function pcf_CheckCommonPass(tmpvalue)
Dim fso,f,tmpStr,fname,findit,Flines,i,intCount,ALines

	pcf_CheckCommonPass=0

	fname="../includes/utilities/pass.inc"
	findit = Server.MapPath(fname)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	Flines = f.ReadAll
	f.close
	
	if Flines<>"" then
		ALines=split(Flines,vbcrlf)
		intCount=ubound(ALines)
	
		For i=0 to intCount
			if UCase(tmpvalue)=UCase(ALines(i)) then
				pcf_CheckCommonPass=1
				exit function
			end if
		Next
	end if

End Function

Public Function pcf_CheckStrongPass(tmpvalue,tmpemail)
Dim tmp1,i,tmpR,tmpR1,tmpR2
Dim upperStr,lowerStr,numStr

upperStr="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
lowerStr="abcdefghijklmnopqrstuvwxyz"
numStr="0123456789"

	pcf_CheckStrongPass=0
	
	if Ucase(tmpemail)=UCase(tmpvalue) then
		pcf_CheckStrongPass=1
		exit function
	end if
	
	if len(tmpvalue)<8 then
		pcf_CheckStrongPass=2
		exit function
	end if
	
	if len(tmpvalue)>=127 then
		pcf_CheckStrongPass=3
		exit function
	end if
	
	tmpR=0
	For i=1 to len(tmpvalue)
		tmp1=Mid(tmpvalue,i,1)
		If Instr(1,upperStr,tmp1,0)>0 then
			tmpR=1
			exit for
		End if
	Next
	
	if tmpR=0 then
		pcf_CheckStrongPass=4
		exit function
	end if
	
IF scStrongPass="1" THEN

	if tmpemail<>"" then
		if (Instr(Ucase(tmpemail),Ucase(tmpvalue))>0) OR (Instr(Ucase(tmpvalue),Ucase(tmpemail))>0) then
			pcf_CheckStrongPass=5
			exit function
		end if
	end if
	
	tmp1=Mid(tmpvalue,1,1)
	If Instr(1,upperStr,tmp1,1)=0 then
		pcf_CheckStrongPass=6
		Exit Function
	End if
	
	tmpR=0
	tmpR1=0
	tmpR2=0
	For i=1 to len(tmpvalue)
		tmp1=Mid(tmpvalue,i,1)
		
		If Instr(1,lowerStr,tmp1,0)>0 then
			tmpR=1
		Else
			If Instr(1,numStr,tmp1,0)>0 then
				tmpR1=1
			Else
				If Instr(1,upperStr,tmp1,0)=0 then
					tmpR2=1
				End if
			End if
		End if
	Next
	
	if tmpR=0 then
		pcf_CheckStrongPass=7
		exit function
	end if
	
	if tmpR1=0 then
		pcf_CheckStrongPass=8
		exit function
	end if
	
	if tmpR2=0 then
		pcf_CheckStrongPass=9
		exit function
	end if
	
END IF

End Function

Public Sub pcs_SaveUsedPass(tmpId,tmpValue)
Dim query,rs,tmpDateTime
Dim strIPAddress

if scCheckSamePass="1" then

	strIPAddress=pcf_GetIPAddress()

	tmpDateTime=Date()
    if SQL_Format="1" then
        tmpDateTime=Day(tmpDateTime)&"/"&Month(tmpDateTime)&"/"&Year(tmpDateTime)
    else
        tmpDateTime=Month(tmpDateTime)&"/"&Day(tmpDateTime)&"/"&Year(tmpDateTime)
    end if
	
	query="INSERT INTO pcUsedPassHistory (IdCustomer,pcUP_UsedPass,pcUP_IPAddress,pcUP_CreatedDate) VALUES (" & tmpID & ",'" & tmpValue & "','" & strIPAddress & "','" & tmpDateTime & "');"
	set rs=connTemp.execute(query)
	set rs=nothing

end if

End Sub


'// Admin GUID
Public Function pcf_CreatePRGuidAdmin(tmpId, tmpEmail)
    Dim query, rs, tmpAdminID, pcStrName, pcStrLastName, pcStrEmail
    Dim Tn1,dd,myC,ReqExist
    Dim ResetTimeOut,tmpDateTime
    Dim SPath1,mycount1,SPathInfo
    Dim pcStrSubject,pcStrBody

	ResetTimeOut = 15 'minutes

	SPathInfo = pcf_getTruePath()	
	if Right(SPathInfo,1)="/" then
		SPathInfo=SPathInfo & scAdminFolderName & "/passwordreset.asp"
	else
		SPathInfo=SPathInfo & "/" & scAdminFolderName & "/passwordreset.asp"
	end if

	query="SELECT idadmin, [adminname], [adm_ContactName], [adm_ContactEmail] FROM [admins] WHERE [idadmin] = " & tmpId & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpAdminID=rs("idadmin")
		pcStrName=rs("adm_ContactName")
        If len(pcStrName)=0 Or IsNull(pcStrName) Then
    	    pcStrName = rs("adminname")
        End If
        If len(pcStrName)=0 Then
    	    pcStrName = "Control Panel User"
        End If
		pcStrEmail = tmpEmail
	else
		set rs=nothing
		pcf_CreatePRGuid=0
		exit function
	end if
	set rs=nothing
	
	DO
		Tn1=""
		For dd=1 to 24
			Randomize
			myC=Fix(3*Rnd)
			Select Case myC
				Case 0:
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
				Case 1:
					Randomize
					Tn1=Tn1 & Cstr(Fix(10*Rnd))
				Case 2:
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+97)
			End Select
		Next

		ReqExist=0

		query="SELECT pcPassID FROM pcPassResetHistory WHERE pcPassResetGuid LIKE '" & Tn1 & "';"
		set rs=connTemp.execute(query)
		if not rs.eof then
			ReqExist=1
		end if
		set rs=nothing
	LOOP UNTIL ReqExist=0
	
	tmpDateTime = DateAdd("n", ResetTimeOut, Now())
	
	query="INSERT INTO pcPassResetHistory (IdCustomer, pcPassResetGuid, pcPassResetTimeOut) VALUES (" & tmpAdminID & ",'" & Tn1 & "','" & tmpDateTIme & "');"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	SPathInfo=SPathInfo & "?aid=" & Server.URLEncode(tmpAdminID) & "&GUID=" & Server.URLEncode(Tn1)
	
	pcStrSubject=dictLanguage.Item(Session("language")&"_resetpasswordmailsubject")
    pcStrBody=dictLanguage.Item(Session("language")&"_resetpasswordadminmailbody")
    pcStrBody=replace(pcStrBody, "#reseturl", SPathInfo)  
    pcStrBody=replace(pcStrBody, "#admin", pcStrName)     

    call sendmail (scEmail, scEmail, pcStrEmail, pcStrSubject, pcStrBody)
    
    pcf_CreatePRGuidAdmin = 1
    
End Function


'// Validate Admin GUID
Public Function pcf_CheckPRGuidAdmin(tmpAdminID, tmpGuid)
    Dim query, rs, tmpDateTime

	pcf_CheckPRGuidAdmin=0
	
	query="SELECT idadmin FROM [admins] WHERE [idadmin] = " & tmpAdminID & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then

	else
		set rs=nothing
		pcf_CheckPRGuid=1
		exit function
	end if
	set rs=nothing
	
    tmpDateTime=Date()
    if SQL_Format="1" then
        tmpDateTime=Day(tmpDateTime)&"/"&Month(tmpDateTime)&"/"&Year(tmpDateTime)
    else
        tmpDateTime=Month(tmpDateTime)&"/"&Day(tmpDateTime)&"/"&Year(tmpDateTime)
    end if
	
	query="SELECT pcPassId FROM pcPassResetHistory WHERE idCustomer=" & tmpAdminID & " AND pcPassResetGuid LIKE '" & tmpGuid & "' AND pcPassResetTimeOut>='" & tmpDateTime & "' AND pcPassResetSuccess=0;"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcf_CheckPRGuidAdmin=2
		call pcs_UpdatePRGuid(tmpGuid, "2")
	end if
	set rs=nothing
	
End Function


'// Customer GUID
Public Function pcf_CreatePRGuid(tmpId, tmpEmail)
    Dim query, rs, tmpCustID, pcStrName, pcStrLastName, pcStrEmail
    Dim Tn1,dd,myC,ReqExist
    Dim ResetTimeOut,tmpDateTime
    Dim SPath1,mycount1,SPathInfo
    Dim pcStrSubject,pcStrBody

	ResetTimeOut = 15 'minutes

	SPathInfo = pcf_getTruePath()	
	if Right(SPathInfo,1)="/" then
		SPathInfo=SPathInfo & "pc/passwordreset.asp"
	else
		SPathInfo=SPathInfo & "/pc/passwordreset.asp"
	end if

	if tmpid<>"" then
		query="SELECT idcustomer,[name],lastname,email FROM Customers WHERE idCustomer=" & tmpid & ";"
	else
		query="SELECT idcustomer,[name],lastname,email FROM Customers WHERE email LIKE '" & tmpemail & "';"
	end if
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpCustID=rs("idcustomer")
		pcStrName=rs("name")
    	pcStrLastName=rs("lastname")
		pcStrEmail=rs("email")
	else
		set rs=nothing
		pcf_CreatePRGuid=0
		exit function
	end if
	set rs=nothing
	
	DO
		Tn1=""
		For dd=1 to 24
			Randomize
			myC=Fix(3*Rnd)
			Select Case myC
				Case 0:
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
				Case 1:
					Randomize
					Tn1=Tn1 & Cstr(Fix(10*Rnd))
				Case 2:
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+97)
			End Select
		Next

		ReqExist=0

		query="SELECT pcPassID FROM pcPassResetHistory WHERE pcPassResetGuid LIKE '" & Tn1 & "';"
		set rs=connTemp.execute(query)
		if not rs.eof then
			ReqExist=1
		end if
		set rs=nothing
	LOOP UNTIL ReqExist=0
	
	tmpDateTime=DateAdd("n",ResetTimeOut,Now())
	
	query="INSERT INTO pcPassResetHistory (IdCustomer,pcPassResetGuid,pcPassResetTimeOut) VALUES (" & tmpCustID & ",'" & Tn1 & "','" & tmpDateTIme & "');"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	SPathInfo=SPathInfo & "?email=" & Server.URLEncode(pcStrEmail) & "&GUID=" & Server.URLEncode(Tn1)
	
	pcStrSubject=dictLanguage.Item(Session("language")&"_resetpasswordmailsubject")
    pcStrBody=dictLanguage.Item(Session("language")&"_resetpasswordmailbody")
    pcStrBody=replace(pcStrBody,"#reseturl",SPathInfo)  
    pcStrBody=replace(pcStrBody,"#firstname",pcStrName)      
    pcStrBody=replace(pcStrBody,"#lastname",pcStrLastName)
    call sendmail (scEmail, scEmail, pcStrEmail, pcStrSubject, pcStrBody)
    pcf_CreatePRGuid=1
    
End Function

'// Validate Customer GUID
Public Function pcf_CheckPRGuid(tmpemail, tmpGuid)
    Dim query, rs, tmpDateTime
    Dim tmpCustID

	pcf_CheckPRGuid=0
	
	query="SELECT idcustomer FROM Customers WHERE email LIKE '" & tmpemail & "';"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpCustID=rs("idcustomer")
	else
		set rs=nothing
		pcf_CheckPRGuid = 1
		exit function
	end if
	set rs=nothing
	
    tmpDateTime=Date()
    if SQL_Format="1" then
        tmpDateTime=Day(tmpDateTime)&"/"&Month(tmpDateTime)&"/"&Year(tmpDateTime)
    else
        tmpDateTime=Month(tmpDateTime)&"/"&Day(tmpDateTime)&"/"&Year(tmpDateTime)
    end if
	
	query="SELECT pcPassId FROM pcPassResetHistory WHERE pcPassResetGuid LIKE '" & tmpGuid & "' AND pcPassResetTimeOut>='" & tmpDateTime & "' AND pcPassResetSuccess=0;"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcf_CheckPRGuid=2
		call pcs_UpdatePRGuid(tmpGuid,"2")
	end if
	set rs=nothing
	
End Function


Public Sub pcs_UpdatePRGuid(tmpGuid,tmpSuccess)
    Dim query,rs,tmpDateTime
    Dim strIPAddress

	strIPAddress=pcf_GetIPAddress()
	
    tmpDateTime=Date()
    if SQL_Format="1" then
        tmpDateTime=Day(tmpDateTime)&"/"&Month(tmpDateTime)&"/"&Year(tmpDateTime)
    else
        tmpDateTime=Month(tmpDateTime)&"/"&Day(tmpDateTime)&"/"&Year(tmpDateTime)
    end if
	
	query="UPDATE pcPassResetHistory SET pcPassResetIPAddress='" & strIPAddress & "', pcPassResetTime='" & tmpDateTime & "', pcPassResetSuccess=" & tmpSuccess & " WHERE pcPassResetGuid LIKE '" & tmpGuid & "';"
	set rs=connTemp.execute(query)
	set rs=nothing

End Sub


Public Sub pcs_SendResetPassMail(tmpId,tmpEmail)
    Dim query, rs, tmpCustID, pcStrName, pcStrLastName, pcStrEmail
    Dim pcStrSubject, pcStrBody
    
    IF scResetPassMail="1" THEN
        if tmpid<>"" then
            query="SELECT idcustomer,[name],lastname,email FROM Customers WHERE idCustomer=" & tmpid & ";"
        else
            query="SELECT idcustomer,[name],lastname,email FROM Customers WHERE email LIKE '" & tmpemail & "';"
        end if
        set rs=connTemp.execute(query)
        if not rs.eof then
            tmpCustID=rs("idcustomer")
            pcStrName=rs("name")
            pcStrLastName=rs("lastname")
            pcStrEmail=rs("email")
        end if
        set rs=nothing
        
        pcStrSubject=dictLanguage.Item(Session("language")&"_passwordchangedmailsubject")
        pcStrBody=dictLanguage.Item(Session("language")&"_passwordchangedmailbody")
        pcStrBody=replace(pcStrBody,"#firstname",pcStrName)      
        pcStrBody=replace(pcStrBody,"#lastname",pcStrLastName)
        call sendmail (scEmail, scEmail, pcStrEmail, pcStrSubject, pcStrBody)
        
    END IF
	
End Sub


Public Function pcf_SaveLoginLockFailed(tmpID,tmpFailed)
Dim query,rs,strIPAddress
Dim tmpDateTime,tmpLocked,tmpCount,tmpMin,tmpTime,tmpTotal
Dim tmpFromDate,tmpToDate

IF (scSaveLogins="1") OR (((scLockFailedUser="1") OR (scLockFailedIP="1")) AND (tmpFailed="1")) THEN
	
	strIPAddress=pcf_GetIPAddress()
		
    tmpDateTime=Date()
    if SQL_Format="1" then
        tmpDateTime=Day(tmpDateTime)&"/"&Month(tmpDateTime)&"/"&Year(tmpDateTime)
    else
        tmpDateTime=Month(tmpDateTime)&"/"&Day(tmpDateTime)&"/"&Year(tmpDateTime)
    end if

	query="INSERT INTO pcLoginHistory (IdCustomer,pcLH_IPAddress,pcLH_DateTime,pcLH_Failed) VALUES (" & tmpID & ",'" & strIPAddress & "','" & tmpDateTime & "'," & tmpFailed & ");"
	set rs=connTemp.execute(query)
	set rs=nothing

END IF

pcf_SaveLoginLockFailed=0

IF (tmpFailed="1") THEN

	tmpLocked=0

	IF (scLockFailedUser="1") THEN
		tmpCount=scLockFailedCount
		if tmpCount="" OR tmpCount="0" then
			tmpCount=1
		end if
		tmpMin=scLockFailedMin
		if tmpMin="" OR tmpMin="0" then
			tmpMin=1
		end if
		tmpFromDate=DateAdd("n",-1*Clng(tmpMin),Now())
		tmpToDate=Now()
			
		query="SELECT Count(*) As TotalFailed FROM pcLoginHistory WHERE idCustomer=" & tmpID & " AND pcLH_DateTime>='" & tmpFromDate & "' AND pcLH_DateTime<='" & tmpToDate & "';"
		set rs=connTemp.execute(query)
		tmpTotal=0
		if not rs.eof then
			tmpTotal=rs("TotalFailed")
		end if
		set rs=nothing
		
		if Clng(tmpTotal)>=tmpCount then
			tmpTime=scLockFailedTime
			if tmpTime="" OR tmpTime="0" then
				tmpTime=5
			end if
			tmpDateTime=DateAdd("n",tmpTime,Now())
			
			query="UPDATE Customers SET pcCust_Locked=1, pcCust_LockUntil='" & tmpDateTime & "', pcCust_LockMinutes=" & tmpTime & " WHERE idCustomer=" & tmpID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			tmpLocked=1
			pcf_SaveLoginLockFailed=1
			session("pcSFLockMinutes")=tmpTime
		end if
	END IF
	
	IF (scLockFailedIP="1") AND (tmpLocked=0) THEN
		tmpCount=scLockFailedIPCount
		if tmpCount="" OR tmpCount="0" then
			tmpCount=1
		end if
		tmpMin=scLockFailedIPMin
		if tmpMin="" OR tmpMin="0" then
			tmpMin=1
		end if
		tmpFromDate=DateAdd("n",-1*Clng(tmpMin),Now())
		tmpToDate=Now()
			
		query="SELECT Count(*) As TotalFailed FROM pcLoginHistory WHERE pcLH_IPAddress='" & strIPAddress & "' AND pcLH_DateTime>='" & tmpFromDate & "' AND pcLH_DateTime<='" & tmpToDate & "';"
		set rs=connTemp.execute(query)
		tmpTotal=0
		if not rs.eof then
			tmpTotal=rs("TotalFailed")
		end if
		set rs=nothing
		
		if Clng(tmpTotal)>=tmpCount then
			tmpTime=scLockFailedIPTime
			if tmpTime="" OR tmpTime="0" then
				tmpTime=5
			end if
			tmpDateTime=DateAdd("n",tmpTime,Now())
			
			query="UPDATE Customers SET pcCust_Locked=1, pcCust_LockUntil='" & tmpDateTime & "', pcCust_LockMinutes=" & tmpTime & " WHERE idCustomer=" & tmpID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			
			tmpLocked=1
			pcf_SaveLoginLockFailed=2
			session("pcSFLockMinutes")=tmpTime
		end if
	END IF
	
END IF

End Function

Public Function pcf_CheckUnlockUser(tmpID,tmpemail)
Dim query,rs,tmpDateTime
Dim tmpCustID,tmpLockUntil

    tmpDateTime=Date()
    if SQL_Format="1" then
        tmpDateTime=Day(tmpDateTime)&"/"&Month(tmpDateTime)&"/"&Year(tmpDateTime)
    else
        tmpDateTime=Month(tmpDateTime)&"/"&Day(tmpDateTime)&"/"&Year(tmpDateTime)
    end if
	
	pcf_CheckUnlockUser=0

	if tmpID>"0" then
		query="SELECT idCustomer,pcCust_LockUntil, pcCust_LockMinutes FROM Customers WHERE idCustomer=" & tmpID & " AND pcCust_Locked=1 AND pcCust_LockMinutes>0;"
	else
		query="SELECT idCustomer,pcCust_LockUntil, pcCust_LockMinutes FROM Customers WHERE email LIKE '" & tmpemail & "' AND pcCust_Locked=1 AND pcCust_LockMinutes>0;"
	end if
	set rs=connTemp.execute(query)

	if not rs.eof then
		pcf_CheckUnlockUser=1
		tmpCustID=rs("idCustomer")
		tmpLockUntil=rs("pcCust_LockUntil")
		tmpLockMinutes=rs("pcCust_LockMinutes")
		session("pcSFLockMinutes")=tmpLockMinutes
		set rs=nothing
		
		if (Not IsNull(tmpLockUntil)) AND (tmpLockUntil<>"") then
			if CDate(tmpLockUntil)<=tmpDateTime then
				query="UPDATE Customers SET pcCust_Locked=0, pcCust_LockUntil='', pcCust_LockMinutes=0 WHERE idCustomer=" & tmpCustID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				pcf_CheckUnlockUser=0
				session("pcSFLockMinutes")=""
			end if
		end if
	end if
	
	set rs=nothing

End Function

Public Function pcf_GetIPAddress()
Dim strIPAddress

	strIPAddress=Request.ServerVariables("HTTP_CLIENT_IP")
	if strIPAddress<>"" then
		if Left(strIPAddress,3)="10." OR Left(strIPAddress,8)="192.168." then
			strIPAddress=""
		end if
	end if
	if strIPAddress = "" then
		strIPAddress=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		if strIPAddress<>"" then
			if Left(strIPAddress,3)="10." OR Left(strIPAddress,8)="192.168." then
				strIPAddress=""
			end if
		end if
	end if
	if strIPAddress = "" then
		strIPAddress = Request.ServerVariables( "REMOTE_ADDR" )
	end if
	
	pcf_GetIPAddress=strIPAddress

End Function

%>