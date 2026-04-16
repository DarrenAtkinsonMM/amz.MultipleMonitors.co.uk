<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2
section="products"
pageTitle="Remove Product From A Pending Sale"
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<% 
pIDProduct=request("id")

if pIDProduct="" then
	call closeDb()
response.redirect "menu.asp"
else
	if Not (IsNumeric(pIDProduct)) then
		call closeDb()
response.redirect "menu.asp"
	end if
end if

pcSaleID=request("saleid")



IF Clng(pcSaleID)>0 THEN
	query="DELETE FROM pcSales_Pending WHERE idProduct=" & pIDProduct & " AND pcSales_ID=" & pcSaleID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
END IF



%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
<%IF Clng(pcSaleID)>0 THEN%>
<tr>
	<td>
		<div class="pcCPmessageSuccess">
			This Product has been successfully removed from the pending SALE!
		</div>
	</td>
</tr>
<%ELSE%>
<tr>
	<td>
		<div class="pcCPmessage">
			No Sale was found for this Product (it cannot be removed).
		</div>
	</td>
</tr>
<%END IF%>
<tr>
	<td class="pcCPspacer">&nbsp;</td>
</tr>
<tr>
	<td class="pcCPspacer">&nbsp;</td>
</tr>
<tr>
	<td>
		<%IF Clng(pcSaleID)>0 THEN%>
		<input type="button" class="btn btn-default"  name="Go" value=" View & Modify Product " onclick="location='FindProductType.asp?id=<%=pIDProduct%>';" class="btn btn-primary">
		<%END IF%>	
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->
