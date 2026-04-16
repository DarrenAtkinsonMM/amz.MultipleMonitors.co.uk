<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageTitle="Remove A Sale"
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<%
pcSaleID=request("id")

if pcSaleID="" then
	call closeDb()
response.redirect "sm_manage.asp"
else
	if Not (IsNumeric(pcSaleID)) then
		call closeDb()
response.redirect "sm_manage.asp"
	end if
end if



query="UPDATE pcSales SET pcSales_Removed=1 WHERE pcSales_ID=" & pcSaleID & ";"
set rs=connTemp.execute(query)
set rs=nothing



%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessageSuccess">
			The SALE has been removed successfully!
		</div>
	</td>
</tr>
<tr>
	<td class="pcCPspacer">&nbsp;</td>
</tr>
<tr>
	<td class="pcCPspacer">&nbsp;</td>
</tr>
<tr>
	<td>
		<input type="button" class="btn btn-default"  name="Go" value=" View & Edit Sales " onclick="location='sm_manage.asp';" class="btn btn-primary">	
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->
