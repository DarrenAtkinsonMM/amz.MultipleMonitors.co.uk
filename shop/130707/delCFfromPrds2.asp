<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove custom field from products" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
if (request("action")="apply") and (request("idcustom")<>"") then

	idcustom=mid(request("idcustom"),2,len(request("idcustom")))  
	  
	prdlist=request("prdlist")
	
	pcArr=split(prdlist,",")
	intCount=ubound(pcArr)
	 
	RSu=0
	RFa=0

	for i=0 to intCount
		if trim(pcArr(i))<>"" then
			if Left(request("idcustom"),1)="C" then
				query="DELETE FROM pcPrdXFields WHERE idProduct=" & trim(pcArr(i)) & " AND IdXField=" & idcustom & ";"
				Set rstemp=conntemp.execute(query)
				set rstemp = nothing
			else
				query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & trim(pcArr(i)) & " AND (idSearchData IN (SELECT DISTINCT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & "))"
				Set rstemp=conntemp.execute(query)
				set rstemp = nothing
			end if
			RSu=RSu+1	
		end if
	next 
	
	set rstemp = nothing
	
end if
%>
<!--#include file="AdminHeader.asp"-->

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<div class="pcCPmessageSuccess">
    The selected custom field was deleted from <b><%=RSu%></b> products. <a href="ManageCFields.asp">Manage Custom Fields</a>.
</div>               
<!--#include file="AdminFooter.asp"-->
