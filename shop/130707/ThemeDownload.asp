<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcMobileSettings.asp"-->
<% 
pageTitle="Theme Downloads"
pageIcon="pcv4_icon_settings.png"
section="layout"
%>
<%
Dim pcv_strPageName
pcv_strPageName="ThemeDownload.asp"
%>
<!--include file="AdminHeader.asp"-->
<%
'// START  Set Capture Values
Dim pcv_strImageId
pcv_strImageId = getUserInput(request("id"), 0)

Dim pcv_strThemName
pcv_strThemName = getUserInput(request("theme"), 0)

Dim pcv_strImagePath
pcv_strImagePath = "../pc/theme/" & pcv_strThemName

Dim pcv_strImageUrl
pcv_strImageUrl = "https://api.grabz.it/services/getjspicture.ashx?id=" & pcv_strImageId
'// END  Set Capture Values

If len(pcv_strImageId)>0 And len(pcv_strImageUrl)>0 And len(pcv_strThemName)>0 Then
    Call SaveOnServer(pcv_strImageUrl, pcv_strThemName & ".jpg")
End If

Public Sub SaveOnServer(url, strFileName)

    Dim strRawData, objFSO, objFile
    Dim strFilePath, strFolderPath, strError

    strRawData = GetBinarySource(url, strError)
    
    If Len(strError)>0 Then
        Response.Write("<span style=""color: red;"">Failed to get binary source. Error:<br />" & strError & "</span>")
    Else  
        strFolderPath = Server.MapPath(pcv_strImagePath)
        Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
        If Not(objFSO.FolderExists(strFolderPath)) Then
            objFSO.CreateFolder(strFolderPath)
        End If

        If Len(strFileName)=0 Then
            strFileName = GetCleanName(url)
        End If

        strFilePath = Server.MapPath(pcv_strImagePath & "/" & strFileName)
        Set objFile = objFSO.CreateTextFile(strFilePath)
        objFile.Write(RSBinaryToString(strRawData))
        objFile.Close
        Set objFile = Nothing
        Set objFSO = Nothing
        
        response.Write("OK")
        response.End()

    End If
    
End Sub

Function RSBinaryToString(xBinary)
    Dim Binary
    If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary
    Dim RS, LBinary
    Const adLongVarChar = 201
    Set RS = CreateObject("ADODB.Recordset")
    LBinary = LenB(Binary)
    If LBinary>0 Then
        RS.Fields.Append "mBinary", adLongVarChar, LBinary
        RS.Open
        RS.AddNew
        RS("mBinary").AppendChunk Binary 
        RS.Update
        RSBinaryToString = RS("mBinary")
    Else  
        RSBinaryToString = ""
    End If
End Function

Function MultiByteToBinary(MultiByte)
    Dim RS, LMultiByte, Binary
    Const adLongVarBinary = 205
    Set RS = CreateObject("ADODB.Recordset")
    LMultiByte = LenB(MultiByte)
    If LMultiByte>0 Then
        RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
        RS.Open
        RS.AddNew
        RS("mBinary").AppendChunk MultiByte & ChrB(0)
        RS.Update
        Binary = RS("mBinary").GetChunk(LMultiByte)
    End If
    MultiByteToBinary = Binary
End Function

Function GetBinarySource(url, ByRef strError)
    Dim objXML
    Set objXML=Server.CreateObject("Microsoft.XMLHTTP")
    GetBinarySource=""
    strError = ""
    On Error Resume Next
        objXML.Open "GET", url, False
        objXML.Send
        If Err.Number<>0 Then
            Err.Clear
            Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
            objXML.Open "GET", url, False
            objXML.Send
            If Err.Number<>0 Then
                strError = "Error " & Err.Number & ": " & Err.Description
                Err.Clear
                Exit Function
            End If
         End If
    On Error Goto 0
    GetBinarySource=objXML.ResponseBody
    Set objXML=Nothing
End Function

Function GetCleanName(s)
    Dim result, x, c
    Dim arrTemp
    arrTemp = Split(s, "/")
    If UBound(arrTemp)>0 Then
        For x=0 To UBound(arrTemp)-1
            result = result & GetCleanName(arrTemp(x)) & "_"
        Next
        result = result & GetPageName(s)
    Else  
        For x=1 To Len(s)
            c = Mid(s, x, 1)
            If IsValidChar(c) Then
                result = result & c
            Else  
                result = result & "_"
            End If
        Next
    End If
    Erase arrTemp
    GetCleanName = result
End Function

Function IsValidChar(c)
    IsValidChar = (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or (IsNumeric(c))
End Function

Function GetPageName(strUrl)
    If Len(strUrl)>0 Then
        GetPageName=Mid(strUrl, InStrRev(strUrl, "/")+1, Len(strUrl))
    Else  
        GetPageName=""
    End If
End Function
%>
<!--include file="AdminFooter.asp"-->
