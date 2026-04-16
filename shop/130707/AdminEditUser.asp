<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Users - View/Edit Control Panel User" %>
<% section="layout" %>
<%PmAdmin=19%> 
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Dim pcvAdminName, pcvAdminEmail

IDAdmin=request("ID")
	if not validNum(IDAdmin) then
		call closeDb()
		response.redirect "AdminEditUser.asp?r=1&msg=" & Server.Urlencode("The user ID is not valid.") & "&ID=" & IDAdmin
	end if

if request("action")="update" then

	AdminUser=request("AdminUser")
		if not validNum(AdminUser) then
			call closeDb()
			response.redirect "AdminEditUser.asp?r=1&msg=" & Server.Urlencode("The user ID is not valid.") & "&ID=" & IDAdmin
		end if
		
	
	query="select * from Admins where IDAdmin=" & AdminUser & " and ID<>" & IDAdmin
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	
		if not rstemp.eof then
			set rstemp=nothing
			
			call closeDb()
			response.redirect "AdminEditUser.asp?r=1&msg=" & Server.Urlencode("This User Id is already in use in this store.") & "&ID=" & IDAdmin
		end if
	
	password=request("AdminPassword")
	tmpUpdPass=""
	if password<>"" then
		password=pcf_PasswordHash(password)
		tmpUpdPass=",AdminPassword='" & password & "'"
	end if
	
	pcvAdminName = request("adminName")
	pcvAdminName = pcf_ReplaceCharacters(pcvAdminName)
	pcvAdminEmail = request("adminEmail")
	pcvAdminEmail = pcf_ReplaceCharacters(pcvAdminEmail)
	
	Count=request("Count")
	Permissions=""
	For i=1 to Count
		if request("C" & i)="1" then
			Permissions=Permissions & request("ID" & i) & "*"
		end if
	Next

	if Permissions="" then
		set rstemp=nothing
		
		call closeDb()
		response.redirect "AdminEditUser.asp?r=1&msg=" & Server.Urlencode("You must choose at least one permission.") & "&ID=" & IDAdmin
	end if
	
	query="UPDATE Admins SET IDadmin='" & AdminUser & "'" & tmpUpdPass & ",AdminLevel='" & permissions & "',adm_ContactName=N'" & pcvAdminName & "',adm_ContactEmail='" & pcvAdminEmail & "' WHERE ID=" & IDAdmin
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
	
	call closeDb()
response.redirect "AdminUserManager.asp?s=1&msg=" & Server.Urlencode("User updated successfully!")
end if

%>
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>

function isDigit(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}

function testLen(s,tmpValue)
	{
		var test=""+s ;
		if (test.length<tmpValue)
		{
				return (false);
		}
		return (true);
	}
	
function Form1_Validator(theForm)
{

	if (theForm.AdminUser.value == "")
 	{
		    alert("Please enter a value for the User Name. It must be a number and it must contain a minimum of 5 digits.");
		    theForm.AdminUser.focus();
		    return (false);
	}
	else
	{
	if (testLen(theForm.AdminUser.value,5) == false)
 	{
		    alert("The User Name must contain at least 5 numbers.");
		    theForm.AdminUser.focus();
		    return (false);
	}
	}
	
	if (allDigit(theForm.AdminUser.value) == false)
	{
		    alert("The User Name must be numeric.");
		    theForm.AdminUser.focus();
		    return (false);
	}	
	
	if (theForm.AdminPassword.value != "")
 	{
		if (testLen(theForm.AdminPassword.value,8) == false)
		{
				alert("The Password must contain at least 8 characters.");
				theForm.AdminPassword.focus();
				return (false);
		}
	}
	
	if (theForm.C11.checked == true && theForm.C12.checked == true)
			{
		    alert("Please select only one of the two Manage Pages permissions.");
		    theForm.C11.focus();
		    return (false);
		    }	
  
return (true);
}
</script>

<form name="updateform" method="post" action="AdminEditUser.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" name="ID" value="<%=IDAdmin%>">  
<table class="pcCPcontent">
	<tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
	</tr>
	<tr>
		<td colspan="2">
		Use this feature to edit an existing store manager. For details, <a href="http://wiki.productcart.com/productcart/settings-manage-users" target="_blank">see the ProductCart documentation</a>. </td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>

<%
		
		query="SELECT * FROM Admins WHERE ID=" & IDAdmin
		set rstemp=server.CreateObject("ADODB.Recordset")
		set rstemp=connTemp.execute(query)
	
		if rstemp.eof then
			set rstemp=nothing
			
			call closeDb()
			response.redirect "AdminUserManager.asp?msg=" & Server.Urlencode("This user doesn't exist in the database.")
		end if

		AdminUser=rstemp("IDAdmin")
		AdminPassword=rstemp("AdminPassword")
		Permissions=rstemp("adminlevel")
		pcvAdminName=rstemp("adm_ContactName")
		pcvAdminEmail=rstemp("adm_ContactEmail")

%>
	<tr> 
		<td width="20%" align="right" nowrap>User ID:</td>
		<td width="80%"><input name="AdminUser" type="text" value="<%=AdminUser%>" size="20" maxlength="9">&nbsp;&nbsp;<i>Must be numeric, at least 5 numbers.</i></td>
	</tr>
	<tr> 
		<td align="right">New Password:</td>
		<td><input name="AdminPassword"  type="password" size="20" value="" maxlength="20">&nbsp;&nbsp;<i>Must be at least 8 characters.</i></td>
	</tr>
	<tr> 
		<td width="20%" align="right" nowrap>Contact Name:</td>
		<td width="80%"><input name="AdminName" type="text" value="<%=pcvAdminName%>" size="30"> (<em>optional</em>)</td>
	</tr>
	<tr> 
		<td width="20%" align="right" nowrap>Contact Email:</td>
		<td width="80%"><input name="AdminEmail" type="text" value="<%=pcvAdminEmail%>" size="30"> (<em>optional</em>)</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>  
	<tr> 
		<td align="right" valign="top">Permissions:</td>
		<td valign="top">
			<%
            query="SELECT * FROM Permissions ORDER BY IDPm"
            set rstemp=Server.CreateObject("ADODB.Recordset")
            set rstemp=connTemp.execute(query)
			
			Dim myArr,permissionID,pcv_intCount
			myArr=Split(Permissions,"*")
			pcv_intCount=ubound(myArr)-1
			
			Function findValue(ByRef arr, ByVal val)
				findValue=Null
				For i=0 To pcv_intCount
						If CLng(val) = CLng(arr(i)) Then
							findValue=i
							Exit Function
						End If
				Next
			End Function
			
			Count=0
			do while not rstemp.eof
			permissionID=CLng(rstemp("IDPM"))			
			Count=Count+1
			%>
			<input type="hidden" name="ID<%=Count%>" value="<%=permissionID%>">
			<input type="checkbox" name="C<%=Count%>" value="1" <% if not isNull(findValue(myArr,permissionID)) then %>checked<%end if%> class="clearBorder">
			&nbsp;<%=rstemp("PMName")%><br>
			<%
			rstemp.MoveNext
			loop
			set rstemp = nothing
			
			%>
			<input type="hidden" name="Count" value="<%=Count%>">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td>&nbsp;</td>
		<td>
			<input name="submit" type="submit" value="Update" class="btn btn-primary">
			&nbsp;
			<input name="back" type="button" class="btn btn-default"  onClick="javascript:history.back()" value="Back"> 
		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
