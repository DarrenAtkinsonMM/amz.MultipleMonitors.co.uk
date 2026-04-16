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

AddNewSM=0

tmpSMName=getUserInput(request("pcSMName"),0)
tmpSMType=getUserInput(request("pcSMType"),0)
if tmpSMType="" then
	tmpSMType="0"
end if
tmpSMOrder=getUserInput(request("pcSMOrder"),0)
if tmpSMOrder="" then
	tmpSMOrder="0"
end if

tmpSMMap=trim(request("pcSMMap"))


if tmpSMid="0" then
	AddNewSM=1
	query="INSERT INTO pcShippingMap (pcSM_Name,pcSM_Type,pcSM_Order) VALUES (N'" &tmpSMName& "'," & tmpSMType & "," & tmpSMOrder & ")"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	set rs=nothing
	
	query="SELECT TOP 1 pcSM_ID FROM pcShippingMap ORDER BY pcSM_ID DESC;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpSMid=rs("pcSM_ID")
	end if
	set rs=nothing
	
	msg="The filter has been added successfully!"
	msgType=1
else
	query="UPDATE pcShippingMap SET pcSM_Name=N'" &tmpSMName& "', pcSM_Type=" & tmpSMType & ", pcSM_Order=" & tmpSMOrder & " WHERE pcSM_ID=" & tmpSMid & ";"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	msg="The filter has been updated successfully!"
	msgType=1
end if

if tmpSMMap<>"" then
	query="DELETE FROM pcSMRel WHERE pcSM_ID=" & tmpSMid & ";" 
	set rs=conntemp.execute(query)
	set rs=nothing
	
	tmpArr=split(tmpSMMap,",")
	For ik=lbound(tmpArr) to ubound(tmpArr)
		if trim(tmpArr(ik))<>"" then
			query="INSERT INTO pcSMRel (pcSM_ID,idshipservice) VALUES (" & tmpSMid & "," & trim(tmpArr(ik)) & ");"
			set rs=conntemp.execute(query)
			set rs=nothing
		end if
	Next
end if

%>
<%
if AddNewSM=1 then
	pageTitle="Add New Shipping Filter"
else
	pageTitle="View/ Edit Shipping Filters"
end if
%>
<% Section="shipOpt" %>
<!--#include file="AdminHeader.asp"-->
<br>
<!--#include file="pcv4_showMessage.asp"-->
<br><br>
<div><input type="button" class="btn btn-default"  name="BackButton" value="Back to Manage Shipping Filters" onClick="location='manageShipMap.asp';"></div>

<!--#include file="AdminFooter.asp"-->
