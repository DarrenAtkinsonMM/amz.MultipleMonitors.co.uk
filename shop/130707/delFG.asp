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

if tmpFGid="0" then
	response.redirect "ManageFacetGroups.asp"
end if

'//Delete relationship between product attributes and facets of this facet group
query="DELETE FROM pcFCAttr WHERE pcFC_ID IN (SELECT pcFC_ID FROM pcFacets WHERE pcFG_ID=" & tmpFGID & ");"
set rs=connTemp.execute(query)
set rs=nothing

'//Delete relationship between this facet group and product option group
query="DELETE FROM pcFGOG WHERE pcFG_ID=" & tmpFGID & ";"
set rs=connTemp.execute(query)
set rs=nothing

'//Delete facets of this facet group
query="DELETE FROM pcFacets WHERE pcFG_ID=" & tmpFGID & ";"
set rs=connTemp.execute(query)
set rs=nothing

'//Delete this facet group
query="DELETE FROM pcFacetGroups WHERE pcFG_ID=" & tmpFGID & ";"
set rs=connTemp.execute(query)
set rs=nothing

msg="The Facet Group has been removed successfully!"
msgType=1

pageTitle="Remove Facet Group"
%>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
<br>
<!--#include file="pcv4_showMessage.asp"-->
<br><br>
<div><input type="button" class="btn btn-default"  name="BackButton" value="Back to Manage Facet Groups" onClick="location='ManageFacetGroups.asp';"></div>

<!--#include file="AdminFooter.asp"-->
