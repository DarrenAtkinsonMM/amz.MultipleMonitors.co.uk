<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_PayeezyFunctions.asp"-->
<%
Dim IdOrder,tmpResult

IdOrder=request("Id")
If Not isNumeric(IdOrder) then
	call closeDb()
response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

If request("action")="1" then
	tmpResult=CaptureVoidPayeezy(IdOrder,1)
	if tmpResult=true then
		msg="Captured the Payeezy payment successfully! This Order payment status was changed to <strong>PAID</strong>"
		msgType="1"
	else
		msg="<strong>Cannot catpure the Payeezy payment!</strong><br>Returned error(s) from Payeezy Gateway:<br>" & pcPEYmsg
		msgType="2"
	end if
Else
	If request("action")="2" then
		tmpResult=CaptureVoidPayeezy(IdOrder,2)
		if tmpResult=true then
			msg="Voided the Payeezy payment successfully!  This Order payment status is changed to <strong>VOIDED</strong>"
			msgType="1"
		else
			msg="<strong>Cannot void the Payeezy payment!</strong><br>Returned error(s) from Payeezy Gateway:<br>" & pcPEYmsg
			msgType="2"
		end if
	End if
End if
%>
<%
Dim pageTitle, Section
If request("action")="1" then
	pageTitle="Capture Payeezy Payment"
Else
	pageTitle="Void Payeezy Payment"
End if
pageIcon="pcv4_icon_orders.gif"
Section="orders"
%>
<!--#include file="AdminHeader.asp"-->
<br>
<!--#include file="pcv4_showMessage.asp"-->
<br><br>
<div><input type="button" class="btn btn-default"  name="BackButton" value="Back to Order Details" onClick="location='OrdDetails.asp?id=<%=IdOrder%>';"></div>

<!--#include file="AdminFooter.asp"-->