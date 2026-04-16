<%
strtext1 = strtext1 & pcf_GenSettings

Public Function pcf_GenSettings()

    '// ProductCart Fraud Enabled
    query="SELECT [pcPCWS_IsActive], [pcPCWS_IsEnabled], [pcPCWS_IsProvisioned] FROM [pcWebServiceFeatures] WHERE [pcPCWS_FeatureCode]='pcFraud'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        pcv_strIsActive = rs2("pcPCWS_IsActive")
        pcv_strIsEnabled = rs2("pcPCWS_IsEnabled")
        pcv_strIsProvisioned = rs2("pcPCWS_IsProvisioned")
    End If
    Set rs2 = Nothing

    '// ProductCart Fraud Settings
    If (pcv_strIsActive="1")  And (pcv_strIsEnabled="1") And (pcv_strIsProvisioned="1") Then
    
        query="SELECT [pcPay_FA_LicenseKey] FROM pcWebServiceFraud"
        Set rs2 = server.CreateObject("ADODB.RecordSet")
        Set rs2 = connTemp.execute(query)
        If Not rs2.Eof Then    
            pcv_strAccountId = rs2("pcPay_FA_LicenseKey") 

            
            strtext1 = strtext1 & "private const scFraud_AccountId = " & Chr(34) & pcv_strAccountId & Chr(34) & vbNewLine          
              
        End If
        Set rs2 = Nothing 
        
        strtext1 = strtext1 & "private const scFraud_IsEnabled = " & Chr(34) & 1 & Chr(34) & vbNewLine 
        
    Else    
        strtext1 = strtext1 & "private const scFraud_IsEnabled = " & Chr(34) & 0 & Chr(34) & vbNewLine     
    End If
    
    pcf_GenSettings = strtext1

End Function
%>