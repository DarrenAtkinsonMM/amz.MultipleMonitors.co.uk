<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
tmpSMid=getUserInput(request("id"),0)
if tmpSMid="" then
	tmpSMid=0
end if

if tmpSMid="0" then
	response.redirect "manageShipMap.asp"
end if

'//Delete relationship between shipping methods and this shipping filter
query="DELETE FROM pcSMRel WHERE pcSM_ID=" & tmpSMid & ";"
set rs=connTemp.execute(query)
set rs=nothing

'//Delete this shipping filter
query="DELETE FROM pcShippingMap WHERE pcSM_ID=" & tmpSMid & ";"
set rs=connTemp.execute(query)
set rs=nothing

msg="The filter has been removed successfully!"
msgType=1

pageTitle="Remove Shipping Filter"
%>
<% Section="shipOpt" %>
<!--#include file="AdminHeader.asp"-->
<br>
<!--#include file="pcv4_showMessage.asp"-->
<br><br>
<div><input type="button" class="btn btn-default"  name="BackButton" value="Back to Manage Shipping Filters" onClick="location='manageShipMap.asp';"></div>

<!--#include file="AdminFooter.asp"-->
