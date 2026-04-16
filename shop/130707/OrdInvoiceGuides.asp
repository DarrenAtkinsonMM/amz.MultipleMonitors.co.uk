<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/ShipFromSettings.asp" -->
<!--#include file="../pdf/fpdf.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Order Invoice or Packing Slip</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin: 10px; background-image: none; font-size:14px;">

    
    <table class="pcCPcontent">
	<tr>
		<td>
			<div class="pcCPmessageSuccess" style="font-size:14px; font-family: 'Helvetica Neue',Helvetica,Arial,sans-serif;">
				Please select which guides should be sent
                <ul>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=pc">PC Only</a></li>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=stand">Synergy Stand Only</a></li>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=pcstand">PC & Synergy Stand</a></li>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=bundle">Bundle (1 delivery)</a></li>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=bundle2">Bundle (2 deliveries)</a></li>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=array">Array (1 delivery)</a></li>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=array2">Array (2 deliveries)</a></li>
                <li><a href="/shop/130707/OrdInvoicePDFemail.asp?id=<%= request.querystring("id")%>&guide=none">Invoice Only</a></li>
                </ul>
				Select an option and an email with the invoice and correct attachements will be sent to the customer.
			</div>
		</td>
	</tr>
	<tr> 
		<td valign="top">&nbsp;</td>
	</tr>
	<tr> 
		<td valign="top">&nbsp;</td>
	</tr>
	</table>
<%

call closedb()
%>
</body>
</html>