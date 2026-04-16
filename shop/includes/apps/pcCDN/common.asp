<!--#include file="../../adminv.asp"-->
<!--#include file="../../common.asp"-->
<%
Dim pcv_appURL
pcv_appURL = pcv_marketURL & "api/provision/cdn/"


pcv_strThisFeatureCode = request("code") '// "pcCDN"


Public Sub pcs_provision(domain, uid, key, code)

    '// 1) Generate Request
    Dim jsonService : Set jsonService = JSON.parse("{}")
    jsonService.Set "domain", domain
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
            pcv_strAmazonDomain = responseObj.domain
            pcv_strAmazonId = responseObj.id
        End If 
    
    Else
        response.Write("We could not provision service. Please contact support.")
        response.End()
    End If
    
    If len(pcv_strAmazonDomain)>0 Then

        '// 4) Parse Response and Save Data
        call pcs_Save(pcv_strAmazonDomain, pcv_strAmazonId)
        
    
        '// 5) Update Provisioning Status
        call pcs_UpdateProvisionStatusByCode(code, 1)
    
    Else
    
        '// 4) Parse Response and Save Data
        call pcs_UpdateProvisionStatusByCode(code, 0)
    
    End If

End Sub


Public Sub pcs_Save(domain, id)

    query="SELECT [pcCDN_Id] FROM pcWebServiceCDN"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        call pcs_Update(domain, id)    
    Else    
        call pcs_Add(domain, id)    
    End If
    Set rs2 = Nothing 

End Sub


Public Sub pcs_Add(domain, id)

    query="INSERT INTO pcWebServiceCDN ([pcCDN_Domain], [pcCDN_Distribution]) VALUES ('" & domain & "', '" & id & "');"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing 

End Sub


Public Sub pcs_Update(domain, id)

    query="UPDATE pcWebServiceCDN SET "
    query = query & "pcCDN_Domain='" & domain & "', "
    query = query & "pcCDN_Distribution='" & id & "' "
    query = query & "WHERE [pcCDN_Id]=1"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  

End Sub
%>