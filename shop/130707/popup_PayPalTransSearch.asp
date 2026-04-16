<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!DOCTYPE html>
<html>
<head>
<title>PayPal - Transaction Search</title>
<!--#include file="inc_header.asp"-->
</head>
<body style="background-image: none;">
<%
pcv_strAdminPrefix="1"

If Request("TransID")<>"" then
	pcv_strTransID=trim(Request("TransID"))
End if
%>
<form name="form1" method="post" action="popup_PayPalSearchResults.asp?action=src" class="pcForms">
<table class="pcCPcontent" width="100%">
<tr>
	<td colspan="2"><div class="title">PayPal - Transaction Search</div></td>
</tr>
<tr>
	<td colspan="2" class="pcSpacer">&nbsp;</td>
</tr>
<tr>
   	<th colspan="2">Enter a date range and/or enter the Transaction ID:</th>
</tr>
<tr>
	<td>From Date:</td>
	<td>
		<input type="text" name="startDate" maxlength="10" size="10" value="<%=Date()-6%>"> <i>(format: mm/dd/yyyy)</i>
	</td>
</tr>
<tr>
	<td>To Date:</td>
	<td>
		<input type="text" name="endDate" maxlength="10" size="10" value="<%=Date()%>"> <i>(format: mm/dd/yyyy)</i>
	</td>
</tr>
<tr>
	<td>Transaction ID:</td>
	<td>
		<input type="text" name="transactionID" value="<%=pcv_strTransID%>">
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input type="Submit" value="Submit" class="btn btn-primary">&nbsp;<input type="button" class="btn btn-default"  name="close" value=" Close window " onClick="javascript:window.close();">
	</td>
</tr>
</table>
</form>
</body>
</html>
