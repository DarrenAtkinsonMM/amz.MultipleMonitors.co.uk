<%
strtext1 = strtext1 & pcf_GenSettings

Public Function pcf_GenSettings()

    '// ProductCart CartStack Enabled
    query="SELECT [pcPCWS_IsActive], [pcPCWS_IsEnabled], [pcPCWS_IsProvisioned] FROM [pcWebServiceFeatures] WHERE [pcPCWS_FeatureCode]='pcCartStack'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        pcv_strIsActive = rs2("pcPCWS_IsActive")
        pcv_strIsEnabled = rs2("pcPCWS_IsEnabled")
        pcv_strIsProvisioned = rs2("pcPCWS_IsProvisioned")
    End If
    Set rs2 = Nothing

    '// ProductCart CartStack Settings
    If (pcv_strIsActive="1")  And (pcv_strIsEnabled="1") And (pcv_strIsProvisioned="1") Then
    
        query="SELECT [pcCS_AccountId], [pcCS_SiteId], [pcCS_APIKey] FROM pcWebServiceCartStack"
        Set rs2 = server.CreateObject("ADODB.RecordSet")
        Set rs2 = connTemp.execute(query)
        If Not rs2.Eof Then    
            pcv_strAccountId = rs2("pcCS_AccountId") 
            pcv_strSiteId = rs2("pcCS_SiteId") 
            pcv_strStackAPIKey = rs2("pcCS_APIKey")
            
            strtext1 = strtext1 & "private const scCartStack_AccountId = " & Chr(34) & pcv_strAccountId & Chr(34) & vbNewLine 
            strtext1 = strtext1 & "private const scCartStack_SiteId = " & Chr(34) & pcv_strSiteId & Chr(34) & vbNewLine  
            trtext1 = strtext1 & "private const scCartStack_APIKey = " & Chr(34) & pcv_strStackAPIKey & Chr(34) & vbNewLine             
              
        End If
        Set rs2 = Nothing 
        
        strtext1 = strtext1 & "private const scCartStack_IsEnabled = " & Chr(34) & 1 & Chr(34) & vbNewLine 
        
    Else    
        strtext1 = strtext1 & "private const scCartStack_IsEnabled = " & Chr(34) & 0 & Chr(34) & vbNewLine     
    End If
    
    pcf_GenSettings = strtext1

End Function
%>