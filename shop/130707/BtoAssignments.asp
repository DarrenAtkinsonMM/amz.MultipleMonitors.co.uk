<%@ LANGUAGE="VBSCRIPT" %>
<% 'CONFIGURATOR ONLY FILE %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Dim pageTitle, Section
pageTitle = "Configurable Product Assignments"
Section = "services" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
		
<% pcPageName="BTOAssignments.asp"
'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////

' Get product ID
pidProduct=request("idProduct")
	if pidProduct="" or not validNum(pidProduct) then
		response.write "locateProduct.asp"
	end if
	
dim pcAssignments, intCount, i
	
	query="SELECT description FROM products WHERE idproduct="&pidProduct&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_ProductName = rs("description")
	set rs=nothing
%>
	<table class="pcCPcontent">
<%
	query="SELECT DISTINCT products.idproduct, products.description FROM products INNER JOIN configSpec_products ON (products.idproduct=configSpec_products.specProduct) WHERE products.removed=0 AND configSpec_products.configProduct="&pidProduct&" UNION (SELECT DISTINCT products.idproduct, products.description FROM products INNER JOIN configSpec_Charges ON (products.idproduct=configSpec_Charges.specProduct) WHERE configSpec_Charges.configProduct="&pidProduct&") ORDER BY products.Description ASC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error retrieving configurable product assignments") 
	end if
	IF rs.EOF THEN
%>
		<tr>
			<td colspan="2"><a href="FindProductType.asp?id=<%=pidProduct%>"><%=pcv_ProductName%></a> has not yet been assigned to any configurable product.</td>
		</tr>
<%
	ELSE
%>
		<tr>
			<td colspan="2"><a href="FindProductType.asp?id=<%=pidProduct%>"><%=pcv_ProductName%></a> is currently assigned to the following configurable products:</td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
<%
		pcAssignments	= rs.getRows()
		set rs=nothing
		
		intCount=ubound(pcAssignments,2)
		i = 0
		For i=0 to intCount
%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td width="70%"><a href="FindProductType.asp?id=<%=pcAssignments(0,i)%>" target="_blank"><%=pcAssignments(1,i)%></a></td>
				<td colspan="3" class="cpLinksList" nowrap><a href="modBTOconfiga.asp?idproduct=<%=pcAssignments(0,i)%>" target="_blank">View Configuration</a> | <a href="FindProductType.asp?id=<%=pcAssignments(0,i)%>" target="_blank">Edit Configurable Product</a> | <a href="../pc/configureprd.asp?idproduct=<%=pcAssignments(0,i)%>&adminpreview=1" target="_blank">Preview</a></td>
			</tr>
<%
		Next
	END IF
%>
		<tr>
			<td colspan="2" align="right" style="padding-top: 20px;">&lt;&lt; Back to <a href="FindProductType.asp?id=<%=pidProduct%>"><%=pcv_ProductName%></a></td>
		</tr>
</table>
<!--#include file="adminfooter.asp"-->
