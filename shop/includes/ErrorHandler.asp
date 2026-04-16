<%
    
'version: 05-29-2019-1

    pCurIP=Request.ServerVariables("REMOTE_ADDR")
    sPath=Request.ServerVariables("SCRIPT_NAME")

'**************************************************************
'* Define injection strings to scan for
'**************************************************************
    
    
    'if sig session isn't set for this session, go get latest signatures
    injectionStatus=0
    call checkSigs()
    sub checkSigs()
        if len(session("pcInjectionStringsFP")&"")=0 or len(session("pcInjectionStringsQP")&"")=0 then
            on error resume next
            url = "https://service.productcartlive.com/antihack/getInjSigs.asp?site=" & scStoreURL & "&key=" & scCrypPass & "&ip=" & pCurIP 
            Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML)
            objXMLhttp.SetTimeouts 1000, 2000, 2000, 2000
            objXMLhttp.open "GET", url, false
            objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
            
            objXMLhttp.setOption 2, 13056
	        objXMLhttp.send ""
            rtnInjString=objXMLhttp.responsetext
            Set objXMLhttp = Nothing
            
            if err.number=0 then
                qpSig=split(rtnInjString,"^*")
            
                for each item in qpSig
                    curQPSig=split(item,"#.#")
                    for i=0 to ubound(curQPSig)
                        if cdbl(i)=1 and curQPSig(1)=1 then 'Handle sigs for Q
                           qpInjString=qpInjString & curQPSig(0) & "^*"
                        elseif cdbl(i)=2 and curQPSig(2)=1 then 'Handle sigs for F
                           fpInjString=fpInjString & curQPSig(0) & "^*"
                        end if
                    next
                next
                session("pcInjectionStringsFP")=left(fpInjString,len(fpInjString)-2)
                session("pcInjectionStringsQP")=left(qpInjString,len(qpInjString)-2)
            else
                session("pcInjectionStringsFP")="$%TTG@%##Yg5ty25yf2#%Y@G%"
                session("pcInjectionStringsQP")="$%TTG@%##Yg5ty25yf2#%Y@G%"
            end if
        end if
       
        pcInjectionStringsQPArr=split(session("pcInjectionStringsQP"),"^*")
        pcInjectionStringsFPArr=split(session("pcInjectionStringsFP"),"^*")

        'scan query params
        For Each Item in Request.QueryString
            if ArraySearch(pcInjectionStringsQPArr,Request.QueryString(Item))<>-1 then
                injectionStatus=1
                curInjection=Request.QueryString(Item)
            end if
        Next
    
       ' scan form params 
        For Each Item in Request.Form
            if ArraySearch(pcInjectionStringsFPArr,Request.Form(Item))<>-1 then
                injectionStatus=1
                curInjection=Request.Form(Item)
            end if
        Next

        if injectionStatus=1 then
	        url = "https://service.productcartlive.com/antihack/logAttack.asp?site=" & scStoreURL & "&key=" & scCrypPass & "&ip=" & pCurIP & "&errnum=" & curInjection &"&pcVersion=" & scVersion & "&adminEmail=" & scFrmEmail & "&page=" & sPath
            Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML)
            objXMLhttp.open "GET", url, false
            objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
            objXMLhttp.SetTimeouts 1000, 2000, 2000, 2000
            objXMLhttp.setOption 2, 13056
	        objXMLhttp.send ""
            Set objXMLhttp = Nothing
            session("banned")=now()
            response.redirect("../pc/diagtxt.txt")
        end if

        err.clear
        On error goto 0
    end sub


'**************************************************************
'* END - Define injection strings to scan for
'**************************************************************
  


'**************************************************************
'* START - Retrieve Error Handler status from database
'* This variable is set via the Store Settings page
'**************************************************************
Dim pcIntErrorHandler
query="SELECT pcStoreSettings_ErrorHandler FROM pcStoreSettings WHERE (((pcStoreSettings_ID)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
pcIntErrorHandler=rs("pcStoreSettings_ErrorHandler")
	if pcIntErrorHandler="" then
		pcIntErrorHandler=1
	end if
set rs=nothing
'**************************************************************
'* END - Retrieve Error Handler status from database
'**************************************************************

'// Log file not yet used
Dim pcStrErrFileName
pcStrErrFileName = "#"

'// Create an id the customer can use when they call up.
Dim pcStrCustRefID
pcStrCustRefID = ""

'//do not use yet...
Function LogErrorToFile()
	Dim objFS
	Dim objFile

	On Error Resume Next
	LogError = False

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	If Err.number = 0 Then
		Set objFile = objFS.OpenTextFile (pcStrErrFileName, 8, True)
		If Err.number = 0 Then
			tErrDescription=	Replace(err.description,vbLf,vbCrLf)
			objFile.WriteLine "------------------------------------------------------"
			objFile.WriteLine "* Error At " & Now
			objFile.WriteLine "* CustomerRefID: "  & pcStrCustRefID
			objFile.WriteLine "* Session ID: " & Session.SessionID
			objFile.WriteLine "* Error Number: " & err.number
			objFile.WriteLine "* Error Source: " & err.source
			objFile.WriteLine "* Error Description: " & tErrDescription
			objFile.WriteLine "* RequestMethod: " & Request.ServerVariables("REQUEST_METHOD")
			objFile.WriteLine "* ServerPort: " & Request.ServerVariables("SERVER_PORT")
			objFile.WriteLine "* HTTPS: " & Request.ServerVariables("HTTPS")
			objFile.WriteLine "* LocalAddr: "  & Request.ServerVariables("LOCAL_ADDR")
			objFile.WriteLine "* HostAddress :"  & Request.ServerVariables("REMOTE_ADDR")
			objFile.WriteLine "* UserAgent: " & Request.ServerVariables("HTTP_USER_AGENT")
			objFile.WriteLine "* URL: " &  Request.ServerVariables("URL")

			objFile.WriteLine "* FormData: " & Request.Form
			objFile.WriteLine "* HTTP Headers: "
			objFile.WriteLine "*****************************"
			objFile.WriteLine Replace(Request.ServerVariables("ALL_HTTP"),vbLf,vbCrLf)
			objFile.WriteLine "*****************************"
			objFile.WriteLine "------------------------------------------------------" & vbCrLf
			objFile.Close

		End If
	End If
End Function

Function LogErrorToDatabase()

    Session("pcStrCustRefID") = Session.SessionID & "-" & Hour(Now) & Minute(Now) & Second(Now)

	tErrDescription = Replace(err.description,vbLf,vbCrLf)
	if instr(tErrDescription,"SQL") then

	else
		'// Append the query string for debugging
		if query <> "" then
			pcv_srtErrDescription = tErrDescription & "<p>" & "query=" & query & "</p>"
		else
			pcv_srtErrDescription = tErrDescription
		end if
		pcv_srtErrDescription = replace(pcv_srtErrDescription,"'","''")
		pcv_srtErrDescription = replace(pcv_srtErrDescription,"""","""""")

		Set conError = Server.CreateObject("ADODB.Connection")
		Set rstError = Server.CreateObject("ADODB.Recordset")

		conError.open scDSN

		ErrQuery="INSERT INTO pcErrorHandler (pcErrorHandler_SessionID, pcErrorHandler_RequestMethod, pcErrorHandler_ServerPort, pcErrorHandler_HTTPS, pcErrorHandler_LocalAddr, pcErrorHandler_RemoteAddr, pcErrorHandler_UserAgent, pcErrorHandler_URL, pcErrorHandler_HttpHost, pcErrorHandler_HttpLang, pcErrorHandler_ErrNumber, pcErrorHandler_ErrSource, pcErrorHandler_ErrDescription, pcErrorHandler_InsertDate,pcErrorHandler_CustomerRefID) VALUES ('"&Session.SessionID&"','"&Request.ServerVariables("REQUEST_METHOD")&"','"&Request.ServerVariables("SERVER_PORT")&"','"&Request.ServerVariables("HTTPS")&"','"&Request.ServerVariables("LOCAL_ADDR")&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("HTTP_USER_AGENT")&"','"&Request.ServerVariables("URL")&"','"&Request.ServerVariables("HTTP_Host")&"','"&Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")&"', '"&err.number&"', '"&err.source&"', '"& pcv_srtErrDescription &"','"&Date()&"', '"&Session("pcStrCustRefID")&"');"
		Set rstError = Server.CreateObject("ADODB.Recordset")
		Set rstError = conError.execute(ErrQuery)
		Set rstError = Nothing

		conError.close

		'If the Error Handler is turned on (1), hide the error. If it's turned off (0), show the error in the browser.
		if pcIntErrorHandler=0 then
			response.write "------------------------------------------------------<BR>"
			response.write "* Error At " & Now	&"<BR>"
			response.write "* CustomerRefID: "  & pcStrCustRefID	&"<BR>"
			response.write "* Session ID: " & Session.SessionID	&"<BR>"
			response.write "* Error Number: " & err.number	&"<BR>"
			response.write "* Error Source: " & err.source	&"<BR>"
			response.write "* Error Description: " & tErrDescription	&"<BR>"
			if query <> "" then
				pcv_srtErrDescription=query
				pcv_srtErrDescription = replace(pcv_srtErrDescription,"'","''")
				pcv_srtErrDescription = replace(pcv_srtErrDescription,"""","""""")
				response.write "* Last Query: " & pcv_srtErrDescription	&"<BR>"
			end if
			response.write "* RequestMethod: " & Request.ServerVariables("REQUEST_METHOD")	&"<BR>"
			response.write "* ServerPort: " & Request.ServerVariables("SERVER_PORT")	&"<BR>"
			response.write "* HTTPS: " & Request.ServerVariables("HTTPS")	&"<BR>"
			response.write "* LocalAddr: "  & Request.ServerVariables("LOCAL_ADDR")	&"<BR>"
			response.write "* HostAddress :"  & Request.ServerVariables("REMOTE_ADDR")	&"<BR>"
			response.write "* UserAgent: " & Request.ServerVariables("HTTP_USER_AGENT")	&"<BR>"
			response.write "* URL: " &  Request.ServerVariables("URL")	&"<BR>"
			response.write "* FormData: " & Request.Form	&"<BR>"
			response.write "* HTTP Headers: " 	&"<BR>"
			response.write "*****************************<BR>"
			response.write Replace(Request.ServerVariables("ALL_HTTP"),vbLf,"<BR>")
			response.write "*****************************<BR>"
			response.write "------------------------------------------------------<BR>"
                'Handle scanner/injections issues
                call antiInjection(err.number)
			response.end
		end if
	end if

    'Handle scanner/injections issues
    call antiInjection(err.number)

	err.clear
End Function

'if IP is banned already, just redirect
if len(session("banned")&"") > 0 then
    'reset banned time
    session("banned")=now()
    response.redirect("../pc/diagtxt.txt")
    response.end
end if



'check request header
if instr(1,ucase(request.ServerVariables("ALL_RAW")),"NETSPARKER",1) > 0 then
    on error resume next
	url = "https://service.productcartlive.com/antihack/logAttack.asp?site=" & scStoreURL & "&key=" & scCrypPass & "&ip=" & pCurIP & "&errnum=Netsparker&pcVersion=" & scVersion & "&adminEmail=" & scFrmEmail & "&page=" & sPath
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML)
    objXMLhttp.open "GET", url, false
    objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    objXMLhttp.SetTimeouts 1000, 2000, 2000, 2000
    objXMLhttp.setOption 2, 13056
	objXMLhttp.send ""
    Set objXMLhttp = Nothing
    session("banned")=now()
    err.clear
    on error goto 0
    response.redirect("../pc/diagtxt.txt")
end if



Public Sub antiInjection(error)

    if error > 0 then
        session("errsThisSession")=cdbl(session("errsThisSession"))+1
        pCurIP=Request.ServerVariables("REMOTE_ADDR")
        banned=0

        'Timestamp first error, get config params, check if IP is on active ban list.
        if cdbl(session("errsThisSession"))=1 then
            session("errsThisSessionStarted")=now
            on error resume next
   	        url = "https://service.productcartlive.com/antihack/config.asp?ip=" & pCurIP
            Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML)
            objXMLhttp.open "GET", url, false
            objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
            objXMLhttp.SetTimeouts 1000, 2000, 2000, 2000
            objXMLhttp.setOption 2, 13056
	        objXMLhttp.send ""
            configParams=objXMLhttp.responsetext
            Set objXMLhttp = Nothing

            if len(configParams&"")>0 then
                configParamsArr=split(configParams,",")
                ii=1

                for each x in configParamsArr
                    if ii=1 then
                        session("errsThisSessionLimit")=cdbl(x)
                    elseif ii=2 then
                        session("errsThisSessionTimeSpanMinutes")=cdbl(x)
                    elseif ii=3 then
                        banned=cdbl(x)
                        'IP is within 4 hour ban period.
                        if banned=1 then
                            session("banned")=now()
                            response.redirect("../pc/diagtxt.txt")
                        end if
                    end if
                    ii=ii+1
                next

                err.Clear
                on error goto 0
            else
                session("errsThisSessionLimit")=10
                session("errsThisSessionTimeSpanMinutes")=1
            end if

        end if
  
        'Have errors exceed set Limit?
        if cdbl(session("errsThisSession")) >= cdbl(session("errsThisSessionLimit")) then
            pCurTimeDif=dateDiff("n",session("errsThisSessionStarted"),now)

            'Are there more errors then allowed in the set time span?   
            if pCurTimeDif <= session("errsThisSessionTimeSpanMinutes") then
                on error resume next
	            url = "https://service.productcartlive.com/antihack/logAttack.asp?site=" & scStoreURL & "&key=" & scCrypPass & "&ip=" & pCurIP & "&errnum=" & error & "&pcVersion=" & scVersion & "&adminEmail=" & scFrmEmail & "&page=" & sPath
                Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML)
                objXMLhttp.open "GET", url, false
                objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
                objXMLhttp.SetTimeouts 1000, 2000, 2000, 2000
                objXMLhttp.setOption 2, 13056
	            objXMLhttp.send ""
                Set objXMLhttp = Nothing
                err.Clear
                on error goto 0
                session("banned")=now()
                response.redirect("../pc/diagtxt.txt")
                response.end
            else
                'reset time, err count
                session("errsThisSessionStarted")=now
                session("errsThisSession")=0
            end if
        end if
    end if
End Sub

Public Sub TraceInit()
	Dim rstError, rstError2, pcv_boolEnableTrace, pcv_boolEnableLog
	On Error Resume Next

	pcv_boolEnableTrace = False
	pcv_boolEnableLog = False

	If Session("pcTrace") = False Then

		If (session("admin") = 1) Then

			'// Trace Remaining Session.
			pcv_boolEnableTrace = True

			'// Save Trace to log file.
			pcv_boolEnableLog = True

		End If

		If scEnforceAdmin = 1 Then

			If (pcv_boolEnableTrace = False) Then

				query = "SELECT TOP 1 A.[idadmin], A.[adminlevel] "
				query = query & "FROM [admins] A "
				query = query & "JOIN ( "
				query = query & "  SELECT [adminlevel] "
				query = query & "  FROM [admins] "
				query = query & "  GROUP BY [adminlevel] "
				query = query & "  HAVING Count(*) > 1 "
				query = query & ") B ON B.[adminlevel] = A.[adminlevel] "
				query = query & "WHERE A.[adminlevel] = '19' "
				query = query & "ORDER BY idadmin DESC "

				Set rstError = Server.CreateObject("ADODB.Recordset")
				Set rstError = conntemp.execute(query)

				If Not rstError.Eof Then

					'// Trace Remaining Session.
					pcv_boolEnableTrace = True

					'// Save Trace to log file.
					pcv_boolEnableLog = True

					'// Lock Secondary Admin Accounts.
					query = "WITH admins_CTE AS ( "
					query = query & "    SELECT adminlevel, adminname, row_number() "
					query = query & "    OVER (PARTITION BY adminlevel ORDER BY id) AS [rn] "
					query = query & "    FROM admins "
					query = query & "    WHERE adminlevel = '19' "
					query = query & ") "
					query = query & "UPDATE admins_CTE SET [adminlevel] = '' WHERE [rn] > 1 "

					Set rstError2 = Server.CreateObject("ADODB.Recordset")
					Set rstError2 = conntemp.execute(query)
					Set rstError2 = Nothing



				End If
				Set rstError = Nothing

			End If

		End If

		If pcv_boolEnableTrace = True Then
			Session("pcTrace") = True
		End If
		If pcv_boolEnableLog = True Then
			Session("pcTraceLog") = True
		End If

	End If

	Call TraceClose()
End Sub


Public Function TraceFileName(pcStrFileName)
	On Error Resume Next
	pcStrFileName = lcase(pcStrFileName)
	If Session("pcTraceLog") = True Then
		If len(pcStrFileName)>0 Then
			If instr(pcStrFileName, ".asp") Or instr(pcStrFileName, ".aspx") Or instr(pcStrFileName, ".php") Then
				If instr(pcStrFileName, "\includes\") = 0 Then
					pcStrFileName = replace(pcStrFileName, ".asp", ".txt")
				End If
			End If
		End If
	End If
	TraceFileName = pcStrFileName
End Function


Public Sub pcs_ValidateFileName(pcStrFileName)
	On Error Resume Next
	pcStrFileName = lcase(pcStrFileName)
	If instr(pcStrFileName, ".asp")>0 Or instr(pcStrFileName, ".aspx")>0 Or instr(pcStrFileName, ".php")>0 Then
		Session("pcTrace") = True
		Session("pcTraceLog") = True
	End If
End Sub


Public Sub TraceStack(data)
	On Error Resume Next

	If Session("pcTrace") = True Then

		pcv_strQueryString = ""
		pcv_strQueryString = pcv_strQueryString & "&_Store=" & scStoreURL & "/" & scPcFolder & "/" & scAdminFolderName
		pcv_strQueryString = pcv_strQueryString & "&_Timestamp=" & Now
		pcv_strQueryString = pcv_strQueryString & "&_CustRefID=" & pcStrCustRefID
		pcv_strQueryString = pcv_strQueryString & "&_SessionID=" & Session.SessionID
		pcv_strQueryString = pcv_strQueryString & "&_ServerPort=" & Request.ServerVariables("SERVER_PORT")
		pcv_strQueryString = pcv_strQueryString & "&_HTTPS=" & Request.ServerVariables("HTTPS")
		pcv_strQueryString = pcv_strQueryString & "&_LocalAddr="  & Request.ServerVariables("LOCAL_ADDR")
		pcv_strQueryString = pcv_strQueryString & "&_HostAddress="  & Request.ServerVariables("REMOTE_ADDR")
		pcv_strQueryString = pcv_strQueryString & "&_UserAgent=" & Request.ServerVariables("HTTP_USER_AGENT")
		pcv_strQueryString = pcv_strQueryString & "&_URL=" &  Request.ServerVariables("URL")
		pcv_strQueryString = pcv_strQueryString & "&_FormData=" & Request.Form
		pcv_strQueryString = pcv_strQueryString & "&_HTTPHeaders=" & replace(Request.ServerVariables("ALL_HTTP"), vbLf, "<br />")
		pcv_strQueryString = pcv_strQueryString & "&_SaveUTF8=" & data

		Call TraceXML(Request.Form, pcv_strQueryString)
		Call TraceFile(Request.Form, pcv_strQueryString)
		'Call TraceEvent(Request.Form, pcv_strQueryString)

	End If

End Sub

Public Sub TracePOSTs()
	On Error Resume Next

	If Session("pcTrace") = True Then

		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			Call TraceStack("")
		End If

	End If

	Call TraceClose()
End Sub


Public Sub TraceXML(FormCollection, QueryString)
	Dim objXMLhttp
	On Error Resume Next

	url = "https://www.productcart.com/productcart-errorLog.asp"

    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML)
    objXMLhttp.open "POST", url, false
    objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    objXMLhttp.SetTimeouts 1000, 2000, 2000, 2000
    objXMLhttp.setOption 2, 13056
	objXMLhttp.send FormCollection & QueryString
    Set objXMLhttp = Nothing

	Call TraceClose()
End Sub


Public Sub TraceEvent(FormCollection, QueryString)
	On Error Resume Next

	pcv_strTraceEvent = ""
	pcv_strTraceEvent = pcv_strTraceEvent & "Here is the POST:  " & FormCollection & QueryString

	Dim WshShell
	Set WshShell = Server.CreateObject("WScript.Shell")
	wshshell.Logevent 2, pcv_strTraceEvent
	Set wshshell = Nothing

	Call TraceClose()
End Sub


Public Sub TraceFile(FormCollection, QueryString)
	On Error Resume Next

	pcv_strFormCollection = ""
	For Each item In FormCollection
		pcv_strFormCollection = pcv_strFormCollection & "Key: " & item & " - Value: " & Request.Form(item) & "<br />"
	Next

	Call pcs_logEventUTF8("trace.txt", pcv_strFormCollection & QueryString)

	Call TraceClose()
End Sub


Public Sub TraceClose()
	On Error Resume Next
	If Session("pcTrace") = True Then
		Err.Clear()
	End If
End Sub


    Function ArraySearch(ByRef arr, ByVal val)
      Dim i
      curVal=Unescape(val)
      ArraySearch = -1
      If IsArray(arr) Then
        For i = 0 To UBound(arr)
              
            curSig=Unescape(arr(i))
          If instr(ucase(curVal),ucase(curSig)) >0 Then
            ArraySearch = i
            Exit Function
          End If
        Next
      End If
    End Function

    Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
    URLDecode = ""
    Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
    sOutput = aSplit(0)
    For I = 0 to UBound(aSplit) - 1
    sOutput = sOutput & _
    Chr("&H" & Left(aSplit(i + 1), 2)) &_
    Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
    Next
    End If

    URLDecode = sOutput
    End Function


Call TraceInit()
Call TracePOSTs()
Call TraceClose()
%>
