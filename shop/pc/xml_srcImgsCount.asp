<%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<!--#include file="../includes/common.asp"-->
<!--#include file="inc_srcImgsQuery.asp"-->
<%totalrecords=0

Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if not rs.eof then
	totalrecords=clng(rs.RecordCount)
end if
set rs=nothing
%>
<count><%=totalrecords%></count>
