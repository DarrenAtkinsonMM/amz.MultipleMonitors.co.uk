<!DOCTYPE html>
<%@Language="VBScript"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<html>
<head>
<title>SSL Certificates</title>
<!--#include file="inc_header.asp"-->
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<table class="pcCPcontent">
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	<tr>
		<th colspan="2">Purchase a dedicated SSL certificate</th>
	</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	<tr>
    	<td width="100%">
			Here is a partial list of companies that provide SSL certificates. We are an official Digicert reseller.<br>
			<ul>
			<li><a href="http://www.thawte.com/ssl/index.html" target="_blank">Thawte</a></li>
			<li><a href="http://www.verisign.com/products/site/secure/index.html" target="_blank">VeriSign</a></li>
			<li><a href="http://www.digicert.com/" target="_blank">DigiCert</a></li>
			<li><a href="http://www.geotrust.com/web_security/index.htm" target="_blank">GeoTrust</a></li>
			</ul>
		</td>
  </tr>
</table> 
		
<p>&nbsp;</p>
<p align="center"><a href=# onClick="self.close();">Close Window</a></p>
</div>
</body>
</html>
<% call closeDb() %>
