<!--#include file="../../common.asp"-->
<%
Dim CheckHeader
CheckHeader=Request.ServerVariables("HTTP_APP")
If CheckHeader<>"NetSource ProductCart" Then
	response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_3")
	response.End()
End If

nmMethod=getUserInput(request("action"),0)
if nmMethod="" OR ((nmMethod<>"send") AND (nmMethod<>"add") AND (nmMethod<>"rmv")) then
	response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_3")
	response.End()
end if


Select Case nmMethod

    Case "send":

		if session("admin")="0" OR session("admin")="" then
			response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_3")
			response.End()
		end if
    
        pcv_intSegmentId = getUserInput(request("idproduct"),0) 
        If (pcv_intSegmentId = "") Then
            response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_3")
            response.End()
        End If   
        
        '// Load Settings
        query="SELECT * FROM pcWebServiceBackInStock"
        Set rs2 = server.CreateObject("ADODB.RecordSet")
        Set rs2 = connTemp.execute(query)
        If Not rs2.Eof Then    
            pcv_strMsg = rs2("pcBIS_Msg")
            pcv_strAuto = rs2("pcBIS_Auto")
            pcv_strButtonText = rs2("pcBIS_ButtonText")
            pcv_strSubject = rs2("pcBIS_Subject")
            pcv_strFromEmail = rs2("pcBIS_FromEmail")
            pcv_strFromName = rs2("pcBIS_FromName")
        End If
        Set rs2 = Nothing 
        
        nmSubject = pcv_strSubject 
        nmFromName = pcv_strFromName 
        nmFromEmail = pcv_strFromEmail 
        
        If len(nmFromName)=0 Or IsNull(nmFromName) Then
            nmFromName = scCompanyName
        End If
        If len(nmFromEmail)=0 Or IsNull(nmFromEmail) Then
            nmFromEmail = scFrmEmail
        End If
        If len(nmSubject)=0 Or IsNull(nmSubject) Then
            nmSubject = "Back in Stock {productname}"
        End If
        
        '// Create JSON
        Dim jsonService : Set jsonService = JSON.parse("{}")
        jsonService.Set "handle", pcv_intSegmentId
        jsonService.Set "body", pcv_strMsg
        jsonService.Set "subject", nmSubject
        jsonService.Set "fromEmail", nmFromEmail
        jsonService.Set "fromName", nmFromName
        
        Dim BackInStockAddRequest
        BackInStockAddRequest = JSON.stringify(jsonService, null, 2)       
        
        pcv_strToken = pcf_GetTokenFromDatabase()   

        strRetVal = pcf_PostRequest(BackInStockAddRequest, BackInStockSendUrl & "/" & pcv_intSegmentId, pcv_strToken)
        if (strRetVal="") then
            response.write "ERROR||" & "Could not send."
            response.End()
        end if
    
        If InStr(UCase(strRetVal),"SUCCESS")=0 Then
            response.write "ERROR||Cannot send any notification e-mails. Please try later."
            response.End()            
        End If
        
        query="UPDATE pcBIS_ListEmails SET Sent=1, SentTime='" & Now() & "' WHERE idProduct=" & pcv_intSegmentId & ";"
        set rs=connTemp.execute(query)
        set rs=nothing
    
        'call pcs_RmvWaitList(pcv_intSegmentId)
        
        response.write "SUCCESS||" & tmpSuccess & "All Back in Stock emails have been sent and should arrive within minutes."


    Case "add":


        nmEmail = getUserInput(request("nmEmail"),0)
        nmPrdID = getUserInput(request("idproduct"),0)
        nmQty = getUserInput(request("quantity"),0)

        If nmEmail="" Or nmPrdID="" Or nmQty="" Then
            response.Write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_3")
            response.End()
        End If

        pcvParentPrd=0
        If statusAPP="1" Then
            query="SELECT pcProd_ParentPrd FROM Products WHERE idProduct=" & nmPrdID & ";"
            Set rs = connTemp.execute(query)
            If Not rs.Eof Then
                pcvParentPrd = rs("pcProd_ParentPrd")
                If IsNull(pcvParentPrd) Or pcvParentPrd="" Then
                    pcvParentPrd=0
                End If
            End If
            Set rs = Nothing
        End If
        
        BackInStockAddRequest="{" & vbcrlf
        BackInStockAddRequest=BackInStockAddRequest & """handle"": """ & nmPrdID & """," & vbcrlf
        BackInStockAddRequest=BackInStockAddRequest & """email"": """ & nmEmail & """" & vbcrlf
        BackInStockAddRequest=BackInStockAddRequest & "}" & vbcrlf        
        
        pcv_strToken = pcf_GetTokenFromDatabase()        

        strRetVal = pcf_PostRequest(BackInStockAddRequest, BackInStockAddUrl, pcv_strToken)
        If (strRetVal="") Then
            response.Write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_2")
            response.End()
        End If
    
        '// SUCCESS: Process Results
        If InStr(UCase(strRetVal),"SUCCESS")=0 Then      
            If InStr(UCase(strRetVal),"""ERROR""")>0 Then
                If InStr(UCase(strRetVal),"""UNAUTHORIZED""")>0 Then
                    response.write "ERROR||" & "Please try again later."
                    response.End()
                ElseIf InStr(UCase(strRetVal),"""SUBSCRIBED""")>0 then
                    Response.write "SUCCESS||" & "Good news! You are already subscribed."
                    nmGUID = pcf_getGuidBIS(nmEmail, nmPrdID)
                    call pcs_BISAddCookie(nmPrdID, nmGUID)
                Else
                    response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_2")
                    response.End()
                End If                
            Else
                If InStr(UCase(strRetVal),"""MESSAGE""")>0 Then
                    response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_2")
                    response.End()
                End If
            End If
            
        Else
        
            set Info = JSON.parse(strRetVal)         
            nmGUID = createGuid()
            
            '// Keep Local Record
            query="INSERT INTO pcBIS_ListEmails (Guid, idProduct, Email, parentProductId, AddedTime, Quantity, Sent) VALUES ('" & nmGUID & "', " & nmPrdID & ", '" & nmEmail & "', " & pcvParentPrd & ", '" & Now() & "', " & nmQty & ", 0);"
            set rs=connTemp.execute(query)
            set rs=nothing
            
            Session("pcSFFromEmail")=nmEmail
            
            '// Set Cookie
            call pcs_BISAddCookie(nmPrdID, nmGUID) 
            
            Response.write "SUCCESS||" & dictLanguage.Item(Session("language")&"_BackInStock_1a")  
                    
        End If        
        %>        
        &nbsp;&nbsp;<input name="nmButton" value="<%=dictLanguage.Item(Session("language")&"_css_cancel")%>" class="btn btn-default" onclick="javascript:rmvBackInStock('<%=nmGUID %>');" type="button">

    <%
    Case "rmv":

        pcv_strGuid = getUserInput(request("nmGUID"),0)
        nmPrdID = getUserInput(request("idproduct"),0)
        
        query="SELECT idProduct, Email FROM pcBIS_ListEmails WHERE Guid='" & pcv_strGuid & "'"
        set rs=conntemp.execute(query)
        if not rs.eof then
            pcv_intSegmentId=rs("idProduct")
            pcv_strEmail=rs("Email")
        else
            call pcs_BISRemoveCookie(nmPrdID)
        end if
        set rs=nothing

        if pcv_intSegmentId="" OR pcv_strEmail="" OR pcv_strGuid="" then
            response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_3")
            response.End()
        end if

        BackInStockRemoveRequest="{" & vbcrlf
        BackInStockRemoveRequest=BackInStockRemoveRequest & """email"": """ & pcv_strEmail & """," & vbcrlf
        BackInStockRemoveRequest=BackInStockRemoveRequest & """handle"": """ & nmPrdID & """" & vbcrlf
        BackInStockRemoveRequest=BackInStockRemoveRequest & "}" & vbcrlf   

        pcv_strToken = pcf_GetTokenFromDatabase()

        strRetVal = pcf_DeleteRequest(BackInStockRemoveRequest, BackInStockRmvUrl & "/" & pcv_intSegmentId, pcv_strToken)
        if (strRetVal="") then
            response.write "ERROR||" & dictLanguage.Item(Session("language")&"_BackInStock_3")
            response.End()
        end if
    
        if InStr(UCase(strRetVal),"SUCCESS")=0 then 
            set Info = JSON.parse(strRetVal)
            if InStr(UCase(strRetVal),"""ERROR""")>0 then
                response.write "ERROR||" & Info.error
            else
                if InStr(UCase(strRetVal),"""MESSAGE""")>0 then
                    response.write "ERROR||" & Info.message
                end if
            end if
            response.End()
        end if

        query="DELETE FROM pcBIS_ListEmails WHERE Guid='" & pcv_strGuid & "';"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call pcs_BISRemoveCookie(nmPrdID)
        
        Response.write "SUCCESS||" & dictLanguage.Item(Session("language")&"_BackInStock_4")
        
End Select

Public Sub pcs_BISRemoveCookie(nmPrdID)  
    Response.Cookies("BackInStockPrdID" & nmPrdID) = ""
    Response.Cookies("BackInStockPrdID" & nmPrdID).Expires = Now()   
End Sub

Public Sub pcs_BISAddCookie(nmPrdID, nmGUID)
    Response.Cookies("BackInStockPrdID" & nmPrdID) = nmGUID
    Response.Cookies("BackInStockPrdID" & nmPrdID).Expires = Date()+365  
End Sub


Public Function pcf_getGuidBIS(nmEmail, nmPrdID)
    Dim query, rs
    
    nmGUID=""
    
    query="SELECT Guid FROM pcBIS_ListEmails WHERE Email='" & nmEmail & "' And idProduct=" & nmPrdID & ";"
    Set rs = connTemp.execute(query)
    If Not rs.Eof Then
        nmGUID = rs("Guid")
        If IsNull(nmGUID) Or nmGUID="" Then
            nmGUID=""
        End If
    End If
    Set rs = Nothing
    
    pcf_getGuidBIS = nmGUID
End Function
%>