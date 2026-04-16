<%
strtext1 = strtext1 & pcf_GenSettings

Public Function pcf_GenSettings()

    '// ProductCart CDN Enabled
    query="SELECT [pcPCWS_IsActive], [pcPCWS_IsEnabled], [pcPCWS_IsProvisioned] FROM [pcWebServiceFeatures] WHERE [pcPCWS_FeatureCode]='pcCDN'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        pcv_strIsActive = rs2("pcPCWS_IsActive")
        pcv_strIsEnabled = rs2("pcPCWS_IsEnabled")
        pcv_strIsProvisioned = rs2("pcPCWS_IsProvisioned")
    End If
    Set rs2 = Nothing

    '// ProductCart CDN Settings
    If (pcv_strIsActive="1")  And (pcv_strIsEnabled="1") And (pcv_strIsProvisioned="1") Then
    
        query="SELECT [pcCDN_Domain], [pcCDN_Distribution] FROM pcWebServiceCDN"
        Set rs2 = server.CreateObject("ADODB.RecordSet")
        Set rs2 = connTemp.execute(query)
        If Not rs2.Eof Then    
            pcv_strDomain = rs2("pcCDN_Domain") 
            pcv_strDistribution = rs2("pcCDN_Distribution") 
            
            strtext1 = strtext1 & "private const scCDN_Domain = " & Chr(34) & pcv_strDomain & Chr(34) & vbNewLine 
            strtext1 = strtext1 & "private const scCDN_Distribution = " & Chr(34) & pcv_strDistribution & Chr(34) & vbNewLine            
              
        End If
        Set rs2 = Nothing 
        
        strtext1 = strtext1 & "private const scCDN_IsEnabled = " & Chr(34) & 1 & Chr(34) & vbNewLine 
        
    Else    
        strtext1 = strtext1 & "private const scCDN_IsEnabled = " & Chr(34) & 0 & Chr(34) & vbNewLine     
    End If
    
    pcf_GenSettings = strtext1

End Function
%>