<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<% On Error Resume Next

Session("pcCustomerRegAgreed")="1"

if session("idCustomer")<>"" AND  session("idCustomer")<>"0" then
	'query="UPDATE Customers SET pcCust_AgreeTerms=1 WHERE idCustomer="&session("idCustomer")&";"
	'set rs=connTemp.execute(query)
	'set rs=nothing
end if

OKmsg="OK"
response.write OKmsg
%>


