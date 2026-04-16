<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Synchronize the Amazon.com inventory levels" %>
<% section="specials" %>
<%PmAdmin=3%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
if request("action")<>"upd" then
	response.redirect "menu.asp"
end if

'APP-S
query="SELECT products.idproduct,products.stock,products.pcprod_Apparel FROM products INNER JOIN pcAmazon ON products.idproduct=pcAmazon.idproduct;"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	set rs=nothing
	For i=0 to intCount
		'APP-S
		if pcArr(2,i)="1" then
			query="SELECT sum(stock) As TotalStock,pcprod_ParentPrd As ParentID FROM Products WHERE active=0 AND pcProd_SPInActive=0 AND removed=0 AND pcprod_ParentPrd=" & pcArr(0,i) & " GROUP BY pcprod_ParentPrd;"
			set rs=connTemp.execute(query)
			if not rs.eof then
				pcArr(1,i)=rs("TotalStock")
			end if
			set rs=nothing
		end if
		'APP-E
		query="UPDATE pcAmazon SET pcAmz_quantity=" & pcArr(1,i) & " WHERE idproduct=" & pcArr(0,i)
		set rs=connTemp.execute(query)
		set rs=nothing
	Next
end if
set rs=nothing
%>

<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<div class="pcCPmessage">
			The Amazon.com inventory levels were synchronized successfully!
		</div>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input type="button" name="Back" value=" Back to Main page " onclick="location='pcAmz_main.asp';" class="ibtnGrey">
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->