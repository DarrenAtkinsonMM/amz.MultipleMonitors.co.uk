<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="About ProductCart&reg; - Terms and Conditions" %>
<% Section="about" %>
<%PmAdmin=0%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<form class="pcForms">                
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th>Terms &amp; Conditions</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td><p>Use of this software indicates acceptance of the following End User License Agreement.</p></td>
	</tr>
	<tr>
		<td align="center">
		<!--#include file="inc_EULA.asp"-->
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>	
</table>
</form>
<!--#include file="AdminFooter.asp"-->
