<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Add New Blackout Date"
pageIcon="pcv4_icon_calendar.png"
section="layout"
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
if request("action")="add" then
	
	Blackout_Date=getUserInput(request("Blackout_Date"),0)
	Blackout_Message=getUserInput(request("Blackout_Message"),1400)
	if scDateFrmt = "DD/MM/YY" AND SQL_Format="0" then
		Blackout_DateArry=split(Blackout_Date,"/")
		Blackout_Date=Blackout_DateArry(1)&"/"&Blackout_DateArry(0)&"/"&Blackout_DateArry(2)
	end if
	query="select * from Blackout where Blackout_Date="
	query=query&"'" & Blackout_Date  & "'"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			
			call closeDb()
response.redirect "Blackout_add.asp?r=1&msg=This Blackout Date is already in use"
		end if

	query="insert into Blackout (Blackout_Date,Blackout_Message) values ("
	query=query&"'" & Blackout_Date  & "'"
	query = query & ",N'" & Blackout_Message & "')"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	call closeDb()
response.redirect "Blackout_main.asp?s=1&msg=New Blackout Date was added successfully!"
end if

%>
	
<!--#include file="AdminHeader.asp"-->
<form name="addnew" method="post" action="Blackout_add.asp?action=add" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr> 
		<td width="20%">Blackout Date:</td>
    <td width="80%"><input type="text" id="Blackout_Date" class="datepicker" name="Blackout_Date" size="20">
		</td>
	</tr>
	<tr> 
		<td valign="top" nowrap="nowrap">Blackout Message:</td>
    <td><textarea cols="60" rows="6" name="Blackout_Message"></textarea>
	</td>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" align="center">
		<input type="submit" name="submit" value="Add New" class="btn btn-primary">
		&nbsp;
		<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
