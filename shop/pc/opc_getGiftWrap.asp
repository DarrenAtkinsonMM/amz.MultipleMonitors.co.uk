<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

response.Buffer=true	
%>
<!--#include file="../includes/common.asp"-->

<!--#include file="opc_contentType.asp" -->
<%
Call pcs_CheckLoggedIn()

tmpList=request("list")
if tmpList="" then
	DontHavePrds="1"
end if

IF DontHavePrds="1" THEN   
	response.clear
	Call SetContentType()
	response.write "OK"
	response.End	  
 ELSE   
	response.clear
	Call SetContentType()
	response.write "LOAD"
	response.End	  
END IF
%>
