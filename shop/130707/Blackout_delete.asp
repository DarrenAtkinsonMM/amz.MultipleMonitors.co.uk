<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Add New Blackout Date"
pageIcon="pcv4_icon_calendar.png"
section="layout"
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
	Blackout_Date=request("Blackout_Date")

	if session("PmAdmin")<>"19" then
		call closeDb()
response.redirect "Blackout_main.asp?r=1&msg=You don't have permissions to delete this blackout date."
	end if

	query="delete from Blackout where Blackout_Date="
	query=query & "'" & Blackout_Date  & "'"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
	
	
	call closeDb()
response.redirect "Blackout_main.asp?s=1&msg=This Blackout Date was deleted successfully!"

%>
