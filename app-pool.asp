<%
On Error Resume Next
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 5
conn.Open "Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=stagemm;Integrated Security=SSPI;"

If Err.Number = 0 Then
    Response.Write "Connection SUCCESS<br>"
    Dim rs
    Set rs = conn.Execute("SELECT SYSTEM_USER AS LoggedInAs")
    Response.Write "Connected as: " & rs("LoggedInAs") & "<br>"
    rs.Close
    conn.Close
Else
    Response.Write "Connection FAILED<br>"
    Response.Write "Error: " & Err.Description & "<br>"
End If

Set conn = Nothing
%>