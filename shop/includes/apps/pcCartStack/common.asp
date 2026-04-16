<!--#include file="../../adminv.asp"-->
<!--#include file="../../common.asp"-->
<%
Dim pcv_appURL
pcv_appURL = pcv_marketURL & "api/provision/cartStack/"


pcv_strThisFeatureCode = request("code") '// "pcCartStack"


Public Sub pcs_provision(pcUrl, uid, key, code)

    '// 1) Generate Request
    Dim jsonService : Set jsonService = JSON.parse("{}")
    jsonService.Set "url", pcUrl
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
            pcv_strAccountId = responseObj.accountid
            pcv_strSiteId = responseObj.siteid
            pcv_strStackAPIKey = responseObj.apikey
        End If 
    
    Else
        response.Write("We could not provision service. Please contact support.")
        response.End()
    End If
    
    If len(pcv_strAccountId)>0 Then

        '// 4) Parse Response and Save Data
        call pcs_Save(pcv_strAccountId, pcv_strSiteId, pcv_strStackAPIKey)
        
    
        '// 5) Update Provisioning Status
        call pcs_UpdateProvisionStatusByCode(code, 1)
    
    Else
    
        '// 4) Parse Response and Save Data
        call pcs_UpdateProvisionStatusByCode(code, 0)
    
    End If

End Sub


Public Sub pcs_Save(pcv_strAccountId, pcv_strSiteId, pcv_strStackAPIKey)

    query="SELECT [pcCS_AccountId] FROM pcWebServiceCartStack"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        call pcs_Update(pcv_strAccountId, pcv_strSiteId, pcv_strStackAPIKey)    
    Else    
        call pcs_Add(pcv_strAccountId, pcv_strSiteId, pcv_strStackAPIKey)    
    End If
    Set rs2 = Nothing 

End Sub


Public Sub pcs_Add(pcv_strAccountId, pcv_strSiteId, pcv_strStackAPIKey)

    query="INSERT INTO pcWebServiceCartStack ([pcCS_AccountId], [pcCS_SiteId], [pcCS_APIKey]) VALUES ('" & pcv_strAccountId & "', '" & pcv_strSiteId & "', '" & pcv_strStackAPIKey & "');"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing 

End Sub


Public Sub pcs_Update(pcv_strAccountId, pcv_strSiteId, pcv_strStackAPIKey)

    query="UPDATE pcWebServiceCartStack SET "
    query = query & "pcCS_AccountId='" & pcv_strAccountId & "', "
    query = query & "pcCS_SiteId='" & pcv_strSiteId & "', "
    query = query & "pcCS_APIKey='" & pcv_strStackAPIKey & "' "
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  

End Sub
%>