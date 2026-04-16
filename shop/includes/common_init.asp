<%
Dim query, conntemp, rs, rstemp

If ((scThemeFolder<>"") OR (scStoreOff="1")) _
    And ( _
        (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "msg.asp") = 0) _
        And _
        (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), LCase(scAdminFolderName)) = 0) _
    ) _
Then
	
    response.redirect "msg.asp?message=83"
    
End If
%>