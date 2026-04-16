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

tmpFGid=getUserInput(request("idgroup"),0)
if tmpFGid="" then
	tmpFGid=0
end if

if tmpFCid="0" then
	response.redirect "ManageFacets.asp?id=" & tmpFGID
end if

'//Delete relationship between product attributes and this facet
query="DELETE FROM pcFCAttr WHERE pcFC_ID=" & tmpFCID & ";"
set rs=connTemp.execute(query)
set rs=nothing

'//Delete this facet from the facet group
query="DELETE FROM pcFacets WHERE pcFC_ID=" & tmpFCID & ";"
set rs=connTemp.execute(query)
set rs=nothing

msg="The Facet has been removed successfully!"
msgType=1

pageTitle="Remove Facet"
%>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
<br>
<!--#include file="pcv4_showMessage.asp"-->
<br><br>
<div><input type="button" class="btn btn-default"  name="BackButton" value="Back to Manage Facets" onClick="location='ManageFacets.asp?id=<%=tmpFGID%>';"></div>

<!--#include file="AdminFooter.asp"-->
