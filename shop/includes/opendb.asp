<% 
'// Open Database Connection
Function openDB()
    On Error Resume Next
    Set connTemp = server.createobject("adodb.connection")
    connTemp.Open scDSN  
    If err.number <> 0 Then
	    response.redirect "dbError.asp"
	    response.End()
    End If
End Function

'// Close Database Connection
Function closeDB()
    On Error Resume Next
    connTemp.close
    Set connTemp = nothing
End Function

call openDB()

If scErrorHandler = 1 Then
    On Error Resume Next
Else
    On Error Goto 0  
End If
%>