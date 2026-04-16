<%
'//Specify UTF-8 codepage 
Public Sub pcs_SetCodePage
	Session.LCID = 1033
	Session.CodePage = 65001
	Response.CharSet = "utf-8"
End Sub

Public Function pcf_basicURLDecode(ByVal str)
    Dim intI, strChar, strRes
    str = Replace(str, "+", " ")
    For intI = 1 To Len(str)
        strChar = Mid(str, intI, 1)
        If strChar = "%" Then
            If intI + 2 < Len(str) Then
                strRes = strRes & Chr(CLng("&H" & Mid(str, intI+1, 2)))
                intI = intI + 2
            End If
        Else
            strRes = strRes & strChar
        End If
    Next
    pcf_basicURLDecode = strRes
End Function

Public Function createGuid()
	Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
	tg = TypeLib.Guid
	createGuid = left(tg, len(tg)-2)
	createGuid = replace(createGuid,"{","")
	createGuid = replace(createGuid,"}","")
	Set TypeLib = Nothing
End Function

Public Sub createCookie(n, v, d)
    Response.Cookies(n) = v
    Response.Cookies(n).Expires = Now + d
End Sub
%>