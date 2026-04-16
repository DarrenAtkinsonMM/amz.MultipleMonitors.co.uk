<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
	requestSTR=request("requestSTR")
	pidOrder=request("orderID")
	pTodaysDate=Date()
	if SQL_Format="1" then
		pTodaysDate=(day(pTodaysDate)&"/"&month(pTodaysDate)&"/"&year(pTodaysDate))
	else
		pTodaysDate=(month(pTodaysDate)&"/"&day(pTodaysDate)&"/"&year(pTodaysDate))
	end if
	query="update DPRequests set StartDate='" & pTodaysDate & "' where RequestStr='" & requestSTR & "'"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	set rstemp=nothing

call closeDb()
response.redirect "OrdDetails.asp?id=" & pidOrder & "&s=1&msg=" & Server.URLEncode("The URL expiration date was successfully reset.")
%>
