<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Edit Blackout Date"
pageIcon="pcv4_icon_calendar.png"
section="layout"
%>
<%PmAdmin=1%> 
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Blackout_Date=getUserInput(request("Blackout_Date"),0)
oldBlackout_Date=getUserInput(request("oldBlackout_Date"),0)
Blackout_Message=getUserInput(request("Blackout_Message"),1400)

if request("action")="update" then
	if scDateFrmt = "DD/MM/YY" AND SQL_Format="0" then
		Blackout_DateArry=split(Blackout_Date,"/")
		Blackout_Date=Blackout_DateArry(1)&"/"&Blackout_DateArry(0)&"/"&Blackout_DateArry(2)
	end if
	
	if SQL_Format="1" then
		oldBlackout_Date=Day(oldBlackout_Date)&"/"&Month(oldBlackout_Date)&"/"&Year(oldBlackout_Date)
	else
		oldBlackout_Date=Month(oldBlackout_Date)&"/"&Day(oldBlackout_Date)&"/"&Year(oldBlackout_Date)
	end if

	query="update Blackout set Blackout_Date="
	query=query & "'" & Blackout_Date  & "'"

	query=query & ",Blackout_Message=N'" & Blackout_Message & "' where Blackout_Date=" 
	query=query & "'" & oldBlackout_Date  & "'"
	
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
	
	
	call closeDb()
response.redirect "Blackout_main.asp?s=1&msg=Blackout Date updated successfully!"

end if

%>
	
<!--#include file="AdminHeader.asp"-->

<%
if SQL_Format="1" then
	Blackout_Date=Day(Blackout_Date)&"/"&Month(Blackout_Date)&"/"&Year(Blackout_Date)
else
	Blackout_Date=Month(Blackout_Date)&"/"&Day(Blackout_Date)&"/"&Year(Blackout_Date)
end if
query="select * from Blackout where Blackout_Date="
query=query & "'" & Blackout_Date  & "'"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)
Blackout_Date=rs("Blackout_Date")
Blackout_Message=rs("Blackout_Message")
set rs=nothing

%>
            
<form name="updateform" method="post" action="Blackout_edit.asp?action=update" class="pcForms">
<input type="hidden" name="oldBlackout_Date" value="<%=Blackout_Date%>">
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
		<td width="75%">
      <input type="text" id="Blackout_date" name="Blackout_Date" size="20" class="datepicker" value="<%=showdateFrmt(Blackout_Date)%>">
 		</td>
	</tr>
	<tr> 
		<td width="20%" valign="top">Blackout Message:</td>
		<td width="75%"><textarea cols="60" rows="6" name="Blackout_Message"><%=Blackout_Message%></textarea></td>
	</tr>           
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" align="center">
		<input type="submit" name="submit" value="Update" class="btn btn-primary">
		&nbsp;
		<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
