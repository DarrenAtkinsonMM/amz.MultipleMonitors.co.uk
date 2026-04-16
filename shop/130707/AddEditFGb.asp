<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
tmpFGid=getUserInput(request("id"),0)
if tmpFGid="" then
	tmpFGid=0
end if

tmpFGName=getUserInput(request("pcFGName"),0)

if tmpFGid="0" then
	query="INSERT INTO pcFacetGroups (pcFG_Name) VALUES (N'" &tmpFGName& "')"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	msg="The Facet Group has been added successfully!"
	msgType=1
else
	query="UPDATE pcFacetGroups SET pcFG_Name=N'" &tmpFGName& "' WHERE pcFG_ID=" & tmpFGID & ";"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	msg="The Facet Group has been updated successfully!"
	msgType=1
end if
%>
<%
if tmpFGid="0" then
	pageTitle="Add New Facet Group"
else
	pageTitle="View/Edit Facet Group"
end if
%>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
<br>
<!--#include file="pcv4_showMessage.asp"-->
<br><br>
<div><input type="button" class="btn btn-default"  name="BackButton" value="Back to Manage Facet Groups" onClick="location='ManageFacetGroups.asp';"></div>

<!--#include file="AdminFooter.asp"-->
