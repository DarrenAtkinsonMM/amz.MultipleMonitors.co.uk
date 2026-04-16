<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
response.Buffer=true
pageTitle="Control Panel - Message" %>
<%PmAdmin=0%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<div class="pcCPmessage">
	<% 
	msg = Session("message")
    %>
	<%= msg %>
</div>

<% If Request("back") = 1 Then %>
	<a href="javascript: window.history.go(-1);" class="btn btn-default">Back</a><br /><br />
<% End If %>

<!--#include file="AdminFooter.asp"-->
