<%
Dim pcv_AuthToken

Dim pcv_tokeURL
pcv_tokeURL = "https://service.productcartlive.com/auth/oauth/token"

Dim pcv_baseURL
pcv_baseURL = "https://service.productcartlive.com/auth/api" 

Dim pcv_marketURL
pcv_marketURL = "https://service.productcartlive.com/v1/"



Public Function pcf_VerifyClaim(url, uid, token)

    pcv_boolIsClaimValid = False
    
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "GET", url & uid , false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    objXMLhttp.setRequestHeader "Authorization","Bearer " & token
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send ""
    pcv_strResponse = objXMLhttp.responseText
    Set objXMLhttp = Nothing

    If len(pcv_strResponse)>0 Then
    
        dim responseObj : set responseObj = JSON.parse(pcv_strResponse)
        
        If instr(pcv_strResponse, "message")>0 Then
            pcv_strMessage = responseObj.message
        End If
        
        If len(pcv_strMessage)=0 Then
            dim claimsObj : set claimsObj = responseObj.claims
            for each claim in claimsObj.keys()
                pcv_strType = claimsObj.get(claim).type
                If pcv_strType = "Feature" Then
                    pcv_strFeatureCode = claimsObj.get(claim).value
                    If lcase(pcv_strFeatureCode) = lcase(pcv_strThisFeatureCode) Then
                        pcv_boolIsClaimValid = True
                    End If
                End If
            next
        End If 
    
    End If
    
    pcf_VerifyClaim = pcv_boolIsClaimValid

End Function


Public Function pcf_UpdateToken()

    '// Get New Token
    query="SELECT pcPCWS_Uid, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
    Set rs=connTemp.execute(query)
    If Not rs.eof Then
        pcv_strUid = rs("pcPCWS_Uid") 
        pcv_strUsername = rs("pcPCWS_Username")  
        pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
    End If
    Set rs=nothing

    pcv_strAuthToken = pcf_GetToken(pcv_strUsername, pcv_strPassword)
    
    '// Save Token
    call pcs_SaveToken(pcv_strAuthToken)

    '// Return Token    
    pcf_UpdateToken =  pcv_strAuthToken
    
End Function


Public Function pcf_GetTokenFromDatabase()

    query = "SELECT [pcPCWS_AuthToken] FROM pcWebServiceSettings "
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then
        pcv_strAuthToken = rs2("pcPCWS_AuthToken")
    End IF
    Set rs2 = Nothing 
    
    pcf_GetTokenFromDatabase =  pcv_strAuthToken
    
End Function


Public Function pcf_VerifyClaimByCode(code)
    on error resume next
    
    '// Get Data
    pcv_baseURL = pcv_baseURL & "/accounts/user/"
    pcv_strToken = pcf_GetTokenFromDatabase()
    
    data = ""

    pcv_boolIsClaimValid = False
    
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "GET", pcv_baseURL & scPCWS_Uid , false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    objXMLhttp.setRequestHeader "Authorization","Bearer " & pcv_strToken
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send data
    pcv_strResponse = objXMLhttp.responseText
    Set objXMLhttp = Nothing

    If len(pcv_strResponse)>0 Then
    
        dim responseObj : set responseObj = JSON.parse(pcv_strResponse)
        
        If instr(pcv_strResponse, "message")>0 Then
            pcv_strMessage = responseObj.message
        End If
        
        If len(pcv_strMessage)=0 Then

            dim claimsObj : set claimsObj = responseObj.claims
            for each claim in claimsObj.keys()
                pcv_strType = claimsObj.get(claim).type
                If lcase(pcv_strType) = lcase(code) Then
                    pcv_strValue = claimsObj.get(claim).value
                    If len(pcv_strValue)>0 Then
                        pcv_boolIsClaimValid = True
                    End If
                End If
            next
        End If 
    
    End If
    
    pcf_VerifyClaimByCode = pcv_boolIsClaimValid

End Function


Public Function pcf_PostRequest(data, url, token)

    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "POST", url, false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    If len(token)>0 Then
        objXMLhttp.setRequestHeader "Authorization","Bearer " & token
    End If
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send data
    pcv_strResponse = objXMLhttp.responseText
    Set objXMLhttp = Nothing

    pcf_PostRequest = pcv_strResponse
    
End Function


Public Function pcf_PostForm(data, url, token)
    On Error Resume Next
    Err.Clear()

    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "POST", url, false
    objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    If len(token)>0 Then
        objXMLhttp.setRequestHeader "Authorization","Bearer " & token
    End If
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send data
    pcv_strResponse = objXMLhttp.responseText
    Set objXMLhttp = Nothing
    
    '// Try compatiblity mode...
    If Err.Number <> 0 Then
        Set objXMLhttp = Server.CreateObject("MSXML2.XMLHTTP" & scXML) 
        objXMLhttp.open "POST", url, false
        objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
        If len(token)>0 Then
            objXMLhttp.setRequestHeader "Authorization","Bearer " & token
        End If
        objXMLhttp.setOption 2, 13056
        objXMLhttp.send data
        pcv_strResponse = objXMLhttp.responseText
        Set objXMLhttp = Nothing
        Err.Clear()
    End If

    pcf_PostForm = pcv_strResponse
    
End Function


Public Function pcf_GetRequest(url, token)

    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "GET", url, false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    If len(token)>0 Then
        objXMLhttp.setRequestHeader "Authorization","Bearer " & token
    End If
    objXMLhttp.setOption 2, 13056
    objXMLhttp.Send("")
    pcv_strResponse = objXMLhttp.responseText
    Set objXMLhttp = Nothing

    pcf_GetRequest = pcv_strResponse
    
End Function


Public Function pcf_DeleteRequest(data, url, token)

    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "DELETE", url, false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    If len(token)>0 Then
        objXMLhttp.setRequestHeader "Authorization","Bearer " & token
    End If
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send data
    pcv_strResponse = objXMLhttp.responseText
    Set objXMLhttp = Nothing

    pcf_DeleteRequest = pcv_strResponse
    
End Function


Public Function pcf_PutRequest(data, url, token)

    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "PUT", url, false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    If len(token)>0 Then
        objXMLhttp.setRequestHeader "Authorization","Bearer " & token
    End If
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send data
    If objXMLhttp.status = 200 Then
        pcv_strResponse = objXMLhttp.responseText
    Else
        pcv_strResponse = ""
    End If
    Set objXMLhttp = Nothing

    pcf_PutRequest = pcv_strResponse
    
End Function


Public Function pcf_GetToken(username, password)
    'on error resume next
    
    pcv_strAuthToken = ""
    
    pcv_postObj = "username=" & Server.URLEncode(username) & "&password=" & Server.URLEncode(password) & "&grant_type=password"

    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "POST", pcv_tokeURL, false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send pcv_postObj
    pcv_strResponse = objXMLhttp.responseText

    Set objXMLhttp = Nothing
    
    'response.Write(pcv_strResponse)
    'response.End()

    If len(pcv_strResponse)>0 Then
    
        dim responseObj : set responseObj = JSON.parse(pcv_strResponse)
        
        If instr(pcv_strResponse, "message")>0 Or instr(pcv_strResponse, "error")>0 Then
            pcv_strMessage = responseObj.message
            If len(pcv_strMessage)=0 Then
                pcv_strMessage = responseObj.error_description
            End If
        End If
        
        If len(pcv_strMessage)=0 Then
            pcv_strAuthToken = responseObj.access_token
        End If 
    
    End If
    
    pcf_GetToken = pcv_strAuthToken
    
End Function


Public Sub pcs_UpdateFeature(code, val)

    query = "UPDATE pcWebServiceFeatures SET "
    query = query & "[pcPCWS_IsActive]=" & val & " "
    query = query & "WHERE [pcPCWS_FeatureCode] = '" & code & "'"
    Set rs3 = server.CreateObject("ADODB.RecordSet")
    Set rs3 = connTemp.execute(query)
    Set rs3 = Nothing 
             
End Sub


Public Sub pcs_RemoveFeatureByCode(code)

    query="DELETE FROM pcWebServiceFeatures WHERE [pcPCWS_FeatureCode]='" & code & "'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing 
             
End Sub


Public Sub pcs_AddFeatureByCode(code)

    query="SELECT [pcPCWS_FeatureCode] FROM pcWebServiceFeatures WHERE [pcPCWS_FeatureCode]='" & code & "'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2. Eof Then    
        call pcs_UpdateFeature(code, 1)    
    Else    
        call pcs_AddFeature(code, 1)    
    End If
    Set rs2 = Nothing 
             
End Sub


Public Sub pcs_AddFeature(code, val)

    query="INSERT INTO pcWebServiceFeatures ([pcPCWS_FeatureCode], [pcPCWS_IsActive]) VALUES ('" & code & "', " & val & ");"
    Set rs3 = server.CreateObject("ADODB.RecordSet")
    Set rs3 = connTemp.execute(query)
    Set rs3 = Nothing 
             
End Sub


Public Sub pcs_SaveToken(token)

    query = "UPDATE pcWebServiceSettings SET "
    query = query & "pcPCWS_AuthToken='" & token & "' "
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  
             
End Sub 


Public Function pcf_IsFeatureActiveByCode(code)

    query="SELECT pcPCWS_IsActive FROM pcWebServiceFeatures WHERE pcPCWS_FeatureCode='" & code & "';"
    Set rs = server.CreateObject("ADODB.RecordSet")
    Set rs=connTemp.execute(query)
    If Not rs.Eof Then
        pcv_intIsActive = rs("pcPCWS_IsActive") 
    Else
        pcv_intIsActive = "0"
    End If
    Set rs=nothing
    
    pcf_IsFeatureActiveByCode = pcv_intIsActive

End Function


Public Sub pcs_UpdateFeatureStatusByCode(code, newStatus)

    query="UPDATE pcWebServiceFeatures "
    query = query & "SET pcPCWS_IsEnabled=" & newStatus & " "
    query = query & "WHERE [pcPCWS_FeatureCode]='" & code & "'"
    'response.Write(query)
    'response.End()
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  

End Sub


Public Function pcf_GetFeatureStatusByCode(code)

    query="SELECT pcPCWS_IsEnabled FROM pcWebServiceFeatures WHERE pcPCWS_FeatureCode='" & code & "';"
    Set rs = server.CreateObject("ADODB.RecordSet")
    Set rs=connTemp.execute(query)
    If Not rs.Eof Then
        pcv_intIsEnabled = rs("pcPCWS_IsEnabled") 
    Else
        pcv_intIsEnabled = "0"
    End If
    Set rs=nothing
    
    pcf_GetFeatureStatusByCode = pcv_intIsEnabled

End Function


Public Sub pcs_displayPrice(url, price)

    If len(url)>0 Then
        %>
        <a href="<%=url%>" target="_blank">Pricing Info</a>
        <%
    Else
        %>
        <%=price%>
        <%
    End IF

End Sub


Public Sub pcs_UpdateProvisionStatusByCode(code, newStatus)

    query="UPDATE pcWebServiceFeatures "
    query = query & "SET pcPCWS_IsProvisioned=" & newStatus & " "
    query = query & "WHERE [pcPCWS_FeatureCode]='" & code & "'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  

End Sub


Public Sub pcs_GenGlobalAppInclude()

    Dim strtext1
    
	If instr(pcPageName, "pcws_")>0 Then
		pcStrFolder = "../includes"
	else
		pcStrFolder = "../.."
	end if

    strtext1 = ""

    query="SELECT [pcPCWS_FeatureCode], pcPCWS_IsActive FROM pcWebServiceFeatures;"
    Set rsApps = connTemp.execute(query)
    If Not rsApps.eof Then
        Do While Not rsApps.eof
        
            pcv_strFeatureCode = rsApps("pcPCWS_FeatureCode")
            pcv_intIsActive = rsApps("pcPCWS_IsActive")
            
            If pcv_intIsActive = 1 Then
                pcv_boolIsActive = true
            Else
                pcv_boolIsActive = false
            End If

            '// Handle variant paths
            If instr(pcPageName, "pcws_")>0 Then
                strtext1 = strtext1 & "<" & "!--#include file=""../apps/" & pcv_strFeatureCode & "/methods.asp""--" & ">" & vbNewLine
            Else
                strtext1 = strtext1 & "<" & "!--#include file=""../apps/" & pcv_strFeatureCode & "/methods.asp""--" & ">" & vbNewLine
            End IF

            rsApps.movenext
        Loop
    End If
    Set rsApps=nothing

    call pcs_SaveUTF8(pcStrFolder & "\extendedMethods\apps.asp", pcStrFolder & "\extendedMethods\apps.asp", strtext1)

End Sub


Public Sub pcs_GenGlobalWebServiceSettings()

    Dim strtext1
    
	If instr(pcPageName, "pcws_")>0 Then
		pcStrFolder = "../includes"
	else
		pcStrFolder = "../.."
	end if

    strtext1 = ""
    
    strtext1 = strtext1 & "<" & Chr(37) & vbNewLine 
    
    '// ProductCart Apps Service
    strtext1 = strtext1 & pcf_GenServiceSettings
    
    query="SELECT [pcPCWS_FeatureCode], pcPCWS_IsActive FROM pcWebServiceFeatures;"
    Set rsApps = connTemp.execute(query)
    If Not rsApps.eof Then
        Do While Not rsApps.eof
        
            pcv_strFeatureCode = rsApps("pcPCWS_FeatureCode")
            pcv_intIsActive = rsApps("pcPCWS_IsActive")
            
            If pcv_intIsActive = 1 Then
                pcv_boolIsActive = true
            Else
                pcv_boolIsActive = false
            End If

            '// Run the product modified scripts / handle variant paths
            If instr(pcPageName, "pcws_")>0 Then
                pcv_strGlobalPath = "../includes/apps/" & pcv_strFeatureCode & "/global.asp"
            Else
                pcv_strGlobalPath = "../" & pcv_strFeatureCode & "/global.asp"
            End IF

            execute(pcf_dynamicInclude(pcf_getMappedFileAsString(pcv_strGlobalPath)))

            rsApps.movenext
        Loop
    End If
    Set rsApps=nothing

    strtext1 = strtext1 & Chr(37) & ">"

    call pcs_SaveUTF8(pcStrFolder & "\settingsPCWS.asp", pcStrFolder & "\settingsPCWS.asp", strtext1)

End Sub


Public Function pcf_GenServiceSettings()

    pcv_intStoreOn = "0"
    pcv_intIsActive = "0"

    '// ProductCart Apps Service
    query="SELECT [pcPCWS_IsActive], [pcPCWS_TurnOnOff], [pcPCWS_Uid] FROM pcWebServiceSettings"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        pcv_intStoreOn = rs2("pcPCWS_TurnOnOff")
        pcv_intIsActive = rs2("pcPCWS_IsActive")
        pcv_intUid = rs2("pcPCWS_Uid")
    End If
    Set rs2 = Nothing 
    
    If pcv_intIsActive = "1"  Then
        strtext1 = strtext1 & "private const scPCWS_IsActive = " & Chr(34) & 1 & Chr(34) & vbNewLine
    Else
        strtext1 = strtext1 & "private const scPCWS_IsActive = " & Chr(34) & 0 & Chr(34) & vbNewLine
    End If
    
    If pcv_intStoreOn = "1"  Then
        strtext1 = strtext1 & "private const scPCWS_StoreOn = " & Chr(34) & 1 & Chr(34) & vbNewLine
    Else
        strtext1 = strtext1 & "private const scPCWS_StoreOn = " & Chr(34) & 0 & Chr(34) & vbNewLine
    End If
    
    If len(pcv_intUid)>0 Then
        strtext1 = strtext1 & "private const scPCWS_Uid = " & Chr(34) & pcv_intUid & Chr(34) & vbNewLine
    End If
    
    pcf_GenServiceSettings = strtext1

End Function








Public Function RandomStrongPassword()

    Randomize()

    dim CharacterSetArray
    CharacterSetArray = Array(_
        Array(5, "abcdefghijklmnopqrstuvwxyz"), _
        Array(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), _
        Array(1, "0123456789"), _
        Array(1, "!@#$+-*?:") _
    )

    dim i
    dim j
    dim Count
    dim Chars
    dim Index
    dim Temp

    for i = 0 to UBound(CharacterSetArray)

        Count = CharacterSetArray(i)(0)
        Chars = CharacterSetArray(i)(1)

        for j = 1 to Count

            Index = Int(Rnd() * Len(Chars)) + 1
            Temp = Temp & Mid(Chars, Index, 1)

        next

    next

    dim TempCopy

    do until Len(Temp) = 0

        Index = Int(Rnd() * Len(Temp)) + 1
        TempCopy = TempCopy & Mid(Temp, Index, 1)
        Temp = Mid(Temp, 1, Index - 1) & Mid(Temp, Index + 1)

    loop

    RandomStrongPassword = TempCopy

end function



Public Function pcs_displayImage(path)
    pcs_displayImage = "<img class=""img-responsive"" src=""" & path & """ />"
End Function



Public Function pcf_displayMarket(params, token)

    query = Request.ServerVariables("QUERY_STRING")

    endpoint = pcv_marketURL & "/MarketViews/" & params
    endpoint = replace(endpoint,"[]","")
    endpoint = replace(endpoint,"%5B%5D","")
    
    'response.Write(endpoint)
    'response.End()

    '// START: POST
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "GET", endpoint, false
    objXMLhttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objXMLhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    If len(token)>0 Then
        objXMLhttp.setRequestHeader "Authorization","Bearer " & token
    End If    
    objXMLhttp.setOption 2, 13056
    objXMLhttp.send query
    cfuResult = objXMLhttp.responseText
    Set objXMLhttp = Nothing
    '// END: POST

    pcf_displayMarket = trim(cfuResult)

End Function


Public Function pcf_ValidateURL(url)
    on error resume next
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "GET", url, false
    objXMLhttp.send ""        
    pcf_ValidateURL = objXMLhttp.status       
    Set objXMLhttp = Nothing
End Function


Public Function pcf_GenGlobalSettings(var)
    Dim rs2
    
    query="SELECT [pcPCWS_IsActive], [pcPCWS_IsEnabled], [pcPCWS_IsProvisioned] FROM [pcWebServiceFeatures] WHERE [pcPCWS_FeatureCode]='pcGallery'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        pcv_strIsActive = rs2("pcPCWS_IsActive")
        pcv_strIsEnabled = rs2("pcPCWS_IsEnabled")
        pcv_strIsProvisioned = rs2("pcPCWS_IsProvisioned")
    End If
    Set rs2 = Nothing

    '// ProductCart Gallery Settings
    If (pcv_strIsActive="1")  And (pcv_strIsEnabled="1") And (pcv_strIsProvisioned="1") Then        
        strtext1 = strtext1 & "private const " & var & " = " & Chr(34) & 1 & Chr(34) & vbNewLine 
    Else    
        strtext1 = strtext1 & "private const " & var & " = " & Chr(34) & 0 & Chr(34) & vbNewLine     
    End If
    
    pcf_GenGlobalSettings = strtext1

End Function


Public Sub pcs_AddEventHook(pcv_strShortCode, pcv_strDesc, pcv_strType, pcv_strUri, pcv_strMethod, pcv_strEvent)
    On Error Resume Next
    Dim rs, query    
    
    query="SELECT hook_Shortcode FROM pcHooks WHERE hook_Shortcode='" & pcv_strShortCode & "';"
    Set rs=connTemp.execute(query)
    If rs.eof Then
        query="INSERT INTO pcHooks (hook_Shortcode, hook_Desc, hook_Type, hook_Uri, hook_Method, hook_Lang, hook_Event) VALUES ('" & pcv_strShortCode & "', '" & pcv_strDesc & "', '" & pcv_strType & "', '" & pcv_strUri & "', '" & pcv_strMethod & "', 'ASP', '" & pcv_strEvent & "');"
        set rs2 = server.CreateObject("ADODB.RecordSet")
        set rs2 = conntemp.execute(query)
        set rs2 = Nothing
    End If
    Set rs = Nothing
    
End Sub

Public Sub pcs_RemoveEventHookByCode(pcv_strShortCode)
    Dim rs2, query
    
    query="DELETE FROM pcHooks WHERE [hook_Shortcode]='" & pcv_strShortCode & "'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing 
             
End Sub


Public Sub pcs_InstallWidget(pcv_strShortCode, pcv_strDesc, pcv_strType, pcv_strUri, pcv_strMethod)
    On Error Resume Next
    Dim rs, query    
    
    query="SELECT widget_Shortcode FROM pcWidgets WHERE widget_Shortcode='" & pcv_strShortCode & "';"
    Set rs=connTemp.execute(query)
    If rs.eof Then
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('" & pcv_strShortCode & "', '" & pcv_strDesc & "', '" & pcv_strType & "', '" & pcv_strUri & "', '" & pcv_strMethod & "', 'ASP');"
            set rs2 = server.CreateObject("ADODB.RecordSet")
        set rs2 = conntemp.execute(query)
        set rs2 = Nothing
    End If
    Set rs = Nothing
    
End Sub

Public Sub pcs_RemoveWidgetByCode(pcv_strShortCode)
    Dim rs2, query
    
    query="DELETE FROM pcWidgets WHERE [widget_Shortcode]='" & pcv_strShortCode & "'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing 
             
End Sub
%>