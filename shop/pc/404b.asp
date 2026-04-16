<!--#include file="../includes/common.asp"-->
<%
	Dim pcStrSeoURLs404
	pcStrSeoURLs404="/"&scSeoURLs404
	pcStrSeoURLs404=replace(pcStrSeoURLs404,"//","/")
	if trim(pcStrSeoURLs404) = "" then
		pcStrSeoURLs404="404c.asp"
	end if
	if trim(pcStrSeoURLs404) = "/" then
		pcStrSeoURLs404="404c.asp"
	end if
	Response.Buffer = "True"
	Response.Status = "404 Not Found"
	Server.Transfer(pcStrSeoURLs404)
%>