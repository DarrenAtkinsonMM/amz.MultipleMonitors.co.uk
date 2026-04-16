<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<%
if session("admin") = 0 then
	response.clear
	response.write "SECURITY"
	response.End
end if

if request("duration") = 1 then
	Response.Cookies("pcHideExpressSignUp")="Agreed"
	Response.Cookies("pcHideExpressSignUp").Expires=Date() + 365
	MyCookiePath=Request.ServerVariables("PATH_INFO")
	do while not (right(MyCookiePath,1)="/")
		MyCookiePath=mid(MyCookiePath,1,len(MyCookiePath)-1)
	loop
	Response.Cookies("pcHideExpressSignUp").Path=MyCookiePath
	response.Write("OK")
	response.End()
else
	session("pcPayPalExpressCookie") = 1
	response.Write("OK")
	response.End()
end if


%>
