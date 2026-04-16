<%
On Error Resume Next
Dim providers(2)
providers(0) = "MSOLEDBSQL"
providers(1) = "SQLNCLI11"
providers(2) = "sqloledb"

Dim p, conn
For Each p In providers
    Err.Clear
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 5
    conn.Open "Provider=" & p & ";Data Source=localhost;Initial Catalog=stagemm;User ID=testuser;Password=TestPass123!;"
    
    If Err.Number = 0 Then
        Response.Write p & ": SUCCESS<br>"
        conn.Close
    Else
        Response.Write p & ": FAILED - " & Err.Description & "<br>"
    End If
    Set conn = Nothing
Next
%>