<!--#include file="../../adminv.asp"-->
<!--#include file="../../common.asp"-->
<%
Dim pcv_appURL
pcv_appURL = pcv_marketURL & "api/provision/backinstock/"
'pcv_appURL = "http://localhost:14211/api/provision/backinstock"


pcv_strThisFeatureCode = request("code") '// "pcBackInStock"


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
            pcv_strVal = 1
        End If

    Else
        response.Write("We could not provision service. Please contact support.")
        response.End()
    End If

    If len(pcv_strVal)>0 Then

        '// 4) Parse Response and Save Data
        call pcs_Save("", "", "")


        '// 5) Update Provisioning Status
        call pcs_UpdateProvisionStatusByCode(code, 1)

    Else

        '// 4) Parse Response and Save Data
        call pcs_UpdateProvisionStatusByCode(code, 0)

    End If

End Sub


Public Sub pcs_Save(Msg, Auto, ButtonText)

    query="SELECT [pcBIS_Id] FROM pcWebServiceBackInStock"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then
        call pcs_Update(Msg, Auto, ButtonText)
    Else
        call pcs_Add(Msg, Auto, ButtonText)
    End If
    Set rs2 = Nothing

End Sub


Public Sub pcs_Add(Msg, Auto, ButtonText)

    query="INSERT INTO pcWebServiceBackInStock ([pcBIS_Msg], [pcBIS_Auto], [pcBIS_ButtonText]) VALUES ('" & Msg & "', '" & Auto & "', '" & ButtonText & "');"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing

End Sub


Public Sub pcs_Update(Msg, Auto, ButtonText)

    query="UPDATE pcWebServiceBackInStock SET "
    query = query & "pcBIS_Msg='" & Msg & "', "
    query = query & "pcBIS_Auto='" & Auto & "', "
    query = query & "pcBIS_ButtonText='" & ButtonText & "' "
    'query = query & "WHERE [pcBIS_Id]=1"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing

End Sub
%>
