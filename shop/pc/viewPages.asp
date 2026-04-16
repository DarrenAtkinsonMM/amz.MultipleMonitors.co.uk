<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
dim pcv_IDPage

pcv_IDPage=session("idContentPageRedirect")
if pcv_IDPage = "" then
	pcv_IDPage=getUserInput(request("idpage"),10)
end if

if pcv_IDPage = "" then
	pcf_do301Redirect("viewcontent.asp")
else
	pcf_do301Redirect("viewcontent.asp?idpage=" & pcv_IDPage)
end if
%>
