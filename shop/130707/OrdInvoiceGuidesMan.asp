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
				Enter Details for manual email send:
				<form method="post" name="contact" action="OrdInvoiceGuidesManSend.asp" class="form">
				<table>
				<tr>
				<td>Customer Name:</td>
				<td><input class="form-control" type="text" id="tbCustName" name="tbCustName" size="35" maxlength="170" ></td>
				</tr>
				<tr>
				<td>Customer Email:</td>
				<td><input class="form-control" type="text" id="tbCustEmail" name="tbCustEmail" size="35" maxlength="170" ></td>
				</tr>
				<tr>
				<td>Tracking Number:</td>
				<td><input class="form-control" type="text" id="tbTracking" name="tbTracking" size="35" maxlength="170" ></td>
				</tr>
				<tr>
				<td>Order Type:</td>
				<td><select id="ddlOrdType" name="ddlOrdType">
				<option value="pc">PC Only</option>
				<option value="stand">Stand Only</option>
				</select>
				</td>
				</tr>
				<tr>
				<td>&nbsp;</td>
				<td>
				<button class="pcButton pcButtonContinue btn btn-skin btn-wc btn-contact" id="FormSubmit" name="FormSubmit">
                            <span class="pcButtonText">Send Email</span>
                </button>
				</td>
				</tr>
				</table>
				
				</form>
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