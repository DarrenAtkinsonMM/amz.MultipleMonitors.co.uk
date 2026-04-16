<%
strtext1 = strtext1 & pcf_GenSettings

Public Function pcf_GenSettings()

    '// ProductCart CDN Enabled
    query="SELECT [pcPCWS_IsActive], [pcPCWS_IsEnabled], [pcPCWS_IsProvisioned] FROM [pcWebServiceFeatures] WHERE [pcPCWS_FeatureCode]='pcBackInStock'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        pcv_strIsActive = rs2("pcPCWS_IsActive")
        pcv_strIsEnabled = rs2("pcPCWS_IsEnabled")
        pcv_strIsProvisioned = rs2("pcPCWS_IsProvisioned")
    End If
    Set rs2 = Nothing

    '// Settings
    If (pcv_strIsActive="1")  And (pcv_strIsEnabled="1") And (pcv_strIsProvisioned="1") Then
    
        query="SELECT [pcBIS_Id], [pcBIS_Msg], [pcBIS_Auto], [pcBIS_ButtonText]  FROM pcWebServiceBackInStock"
        Set rs2 = server.CreateObject("ADODB.RecordSet")
        Set rs2 = connTemp.execute(query)
        If Not rs2.Eof Then    

            pcv_strAuto = rs2("pcBIS_Auto")
            pcv_strButtonText = rs2("pcBIS_ButtonText")

            strtext1 = strtext1 & "private const scNM_Auto = " & Chr(34) & pcv_strAuto & Chr(34) & vbNewLine
            strtext1 = strtext1 & "private const scNM_ButtonText = " & Chr(34) & pcv_strButtonText & Chr(34) & vbNewLine          
              
        End If
        Set rs2 = Nothing 
        
        strtext1 = strtext1 & "private const scNM_IsEnabled = " & Chr(34) & 1 & Chr(34) & vbNewLine 
        
    Else    
        strtext1 = strtext1 & "private const scNM_IsEnabled = " & Chr(34) & 0 & Chr(34) & vbNewLine     
    End If
    
    pcf_GenSettings = strtext1

End Function
%>