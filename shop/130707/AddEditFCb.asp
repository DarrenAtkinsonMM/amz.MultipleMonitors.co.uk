<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
tmpFCid=getUserInput(request("id"),0)
if tmpFCid="" then
	tmpFCid=0
end if

tmpFGid=getUserInput(request("pcFGid"),0)
if tmpFGid="" then
	tmpFGid=0
end if

tmpFCName=getUserInput(request("pcFCName"),0)
tmpFCCode=getUserInput(request("pcFCCode"),0)
tmpFCImg=getUserInput(request("pcFCImg"),0)
tmpFCOrder=getUserInput(request("pcFCOrder"),0)
if tmpFCOrder="" then
	tmpFCOrder=0
end if

if tmpFCid="0" then
	query="INSERT INTO pcFacets (pcFG_ID,pcFC_Name,pcFC_Code,pcFC_Img,pcFC_Order) VALUES (" & tmpFGid & ",N'" &tmpFCName& "',N'" &tmpFCCode& "','" &tmpFCImg& "'," & tmpFCOrder & ")"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	msg="The Facet has been added successfully!"
	msgType=1
else
	query="UPDATE pcFacets SET pcFG_ID=" & tmpFGid & ",pcFC_Name=N'" &tmpFCName& "',pcFC_Code=N'" & tmpFCCode & "',pcFC_Img='" & tmpFCImg & "',pcFC_Order=" & tmpFCOrder & " WHERE pcFC_ID=" & tmpFCID & ";"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	msg="The Facet has been updated successfully!"
	msgType=1
end if
%>
<%
if tmpFCid="0" then
	pageTitle="Add New Facet"
else
	pageTitle="View/Edit Facet"
end if
%>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
<br>
<!--#include file="pcv4_showMessage.asp"-->
<br><br>
<div><input type="button" class="btn btn-default"  name="BackButton" value="Back to Manage Facets" onClick="location='ManageFacets.asp?id=<%=tmpFGid%>';"></div>

<!--#include file="AdminFooter.asp"-->
