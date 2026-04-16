<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove XML Partner" %>
<% section="layout"%>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
pidPartner=trim(request("idPartner"))

If Not IsNumeric(pidPartner) then
	call closeDb()
    response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

if request("action")<>"del" then
	call closeDb()
    response.redirect "menu.asp"
end if

query="UPDATE pcXMLPartners SET pcXP_Removed=1 WHERE pcXP_ID=" & pidPartner & ";"
set rs=connTemp.execute(query)
set rs=nothing

call closeDb()
response.redirect "AdminManageXMLPartner.asp?s=1&msg=deleted"
%>
