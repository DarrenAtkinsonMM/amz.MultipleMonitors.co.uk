<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_srcSDSQuery.asp"-->
<%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<%
totalrecords=0

Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open query, connTemp, adOpenStatic, adLockReadOnly, adCmdText
if not rs.eof then
	totalrecords=clng(rs.RecordCount)
end if
set rs=nothing

%>
<count><%=totalrecords%></count>
