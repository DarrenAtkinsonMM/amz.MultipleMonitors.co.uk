<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Deleting Orders" %>
<% Section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
msg=""
if request.form("action")="del" then
	call opendb()
	query="UPDATE Products SET Removed=-1 WHERE pcprod_ParentPrd>0;"
	set rs=connTemp.execute(query)
	set rs=nothing
	query="SELECT DISTINCT idproduct FROM Products WHERE pcprod_ParentPrd>0 AND (idproduct NOT IN (SELECT DISTINCT idProduct FROM wishlist)) AND (idproduct NOT IN (SELECT DISTINCT idProduct FROM ProductsOrdered));"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
		set rs=nothing
		For i=0 to intCount
			query="DELETE FROM Products WHERE idproduct=" & pcArr(0,i) & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		Next
	end if
	set rs=nothing
	
	'APP-S
	' Update parent products inventory levels if necessary%>
	<!--#include file="../pc/app-updstock.asp"-->
	<%' End update parent products inventory levels
	'APP-E
	
	call closedb()
	msg="Sub-products not in previous orders and wish lists were permanently removed from the database. All other sub-products were removed from the store, but they still exist in the 'Products' table with the 'removed' field set to '-1'."
end if%>
<table class="pcCPcontent">
<%if msg<>"" then%>
<tr>
	<td>
		<div class="pcCPmessage">
			<%=msg%>
		</div>
	</td>
</tr>
<tr>
	<td>
		<br>
		<input name="back" type="button" value=" Back to Main menu " onclick="javascript:location='menu.asp';">
	</td>
</tr>
<%else%>
<form name="form1" method="post" action="delAllSubPrds.asp" class="pcForms">
<tr>
	<td>
		<input type="hidden" name="action" value="del">
		<p><b>****Warning****</b><br>
		You may use this form to remove sub-products from your database. This action is permanent and cannot be reversed. The main purpose of this form is to purge sub-products that were entered into your database for testing purposes.</p>
	</td>
</tr>
<tr>
	<td>
		<br>
		<input name="submit1" type="submit" value=" Delete All Sub-Products "onclick="javascript:if (confirm('Are you sure you want to complete this action?')) {return(true);} else {return(false);}" class="ibtnGrey">
	</td>
</tr>
</form>
<%end if%>
</table>
<!--#include file="AdminFooter.asp"-->

	
	
	
