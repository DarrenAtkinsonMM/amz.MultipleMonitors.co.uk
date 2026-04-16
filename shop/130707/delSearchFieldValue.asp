<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Special Customer Fields" %>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
pcv_ID=request("idSearchData")
idSearchField=request("idSearchField")

if pcv_ID="" or pcv_ID="0" then
	call closeDb()
    response.redirect "ManageSearchValues.asp?idSearchField=" & idSearchField & ";"
end if

	query="DELETE FROM pcSearchFields_Products WHERE idSearchData=" & pcv_ID & ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	query="DELETE FROM pcSearchData WHERE idSearchData=" & pcv_ID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing

call closeDb()
response.redirect "ManageSearchValues.asp?idSearchField=" & idSearchField & "&s=1&msg=" & Server.URLEncode("The selected search field value was successfully deleted.")
%>
