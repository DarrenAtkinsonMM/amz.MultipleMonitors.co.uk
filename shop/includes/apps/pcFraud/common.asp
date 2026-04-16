<!--#include file="../../adminv.asp"-->
<!--#include file="../../common.asp"-->
<%
Dim pcv_appURL
pcv_appURL = pcv_marketURL & "api/provision/fraud/"


pcv_strThisFeatureCode = request("code") '// "pcFraud"


Public Sub pcs_provision(pcUrl, uid, key, code)

    '// 1) Generate Request
    Dim jsonService : Set jsonService = JSON.parse("{}")
    jsonService.Set "domain", pcUrl
    jsonService.Set "uid", uid
    jsonService.Set "key", key

    Dim jsonObj
    jsonObj = JSON.stringify(jsonService, null, 2) 


    '// 2) Send Request
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "POST", pcv_appURL, false
    objXMLhttp.setRequestHeader "Content-type","application/json"
    objXMLhttp.setRequestHeader "Accept","application/json"
    If len(pcv_AuthToken)>0 Then
        objXMLhttp.setRequestHeader "Authorization","Bearer " & pcv_AuthToken
    End If
    objXMLhttp.send jsonObj
    pcv_strResponse = objXMLhttp.responseText
    Set objXMLhttp = Nothing
    

    '// 3) Parse Response 
    If len(pcv_strResponse)>0 Then
    
        dim responseObj : set responseObj = JSON.parse(pcv_strResponse)
        
        If instr(pcv_strResponse, "message")>0 Then
            pcv_strMessage = responseObj.message
        End If
        
        If len(pcv_strMessage)=0 Then
            pcv_strAccountId = responseObj.id
        End If 
    
    Else
        response.Write("We could not provision service. Please contact support.")
        response.End()
    End If
    
    If len(pcv_strAccountId)>0 Then

        '// 4) Parse Response and Save Data
        call pcs_Save(pcv_strAccountId)
        
    
        '// 5) Update Provisioning Status
        call pcs_UpdateProvisionStatusByCode(code, 1)
    
    Else
    
        '// 4) Parse Response and Save Data
        call pcs_UpdateProvisionStatusByCode(code, 0)
    
    End If

End Sub


Public Sub pcs_Save(pcv_strAccountId)

    query="SELECT [pcPay_FA_Id] FROM pcWebServiceFraud"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        call pcs_Update(pcv_strAccountId)    
    Else    
        call pcs_Add(pcv_strAccountId)    
    End If
    Set rs2 = Nothing 

End Sub


Public Sub pcs_Add(pcv_strAccountId)

    query="INSERT INTO pcWebServiceFraud ([pcPay_FA_LicenseKey]) VALUES ('" & pcv_strAccountId & "');"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing 

End Sub


Public Sub pcs_Update(pcv_strAccountId)

    query="UPDATE pcWebServiceFraud SET "
    query = query & "pcPay_FA_LicenseKey='" & pcv_strAccountId & "' "
    'query = query & "WHERE [pcPay_FA_Id]=1"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  

End Sub
%>