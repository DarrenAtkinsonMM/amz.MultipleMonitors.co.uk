<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->

<%

bid = Request.QueryString("bid")

query = "DELETE FROM mod_bannermanagement WHERE bannerid = " & bid
set rs=server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number<>0 then
  call LogErrorToDatabase()
  set rs=nothing
  call closedb()
  response.redirect "techErr.asp?err="&pcStrCustRefID
else
  response.redirect "bannermanagement.asp?action=del"
end if

%>
