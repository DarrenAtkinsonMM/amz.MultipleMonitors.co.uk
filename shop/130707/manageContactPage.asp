<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Contact Page" %>
<% section="specials" %>
<%PmAdmin=3%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!-- #include file="../htmleditor/editor.asp" -->
<% 
Dim pcStrPageDesc

if request("action")="upd" then
	pcStrPageDesc=replace(request("pcStrPageDesc"),"'","''")

	query="SELECT pcCPage_ID FROM pcContactPageSettings;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		dim rsUpdObj
		query="UPDATE pcContactPageSettings SET pcCPage_PageDesc=N'" & pcStrPageDesc & "';"
		set rsUpdObj=server.CreateObject("ADODB.RecordSet")
		set rsUpdObj=connTemp.execute(query)
		set rsUpdObj=nothing
	else
		dim rsInsObj
		query="INSERT INTO pcContactPageSettings (pcCPage_PageDesc) VALUES (N'" & pcStrPageDesc & "');"
		set rsInsObj=server.CreateObject("ADODB.RecordSet")
		set rsInsObj=connTemp.execute(query)
		set rsInsObj=nothing
	end if
	
	set rs=nothing
	
	msg="Contact Page Settings were updated successfully!"
	msgtype=1
end if

pcStrPageDesc=""

query="SELECT pcCPage_PageDesc FROM pcContactPageSettings;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcStrPageDesc=rs("pcCPage_PageDesc")
	if pcStrPageDesc<>"" then
		pcStrPageDesc=replace(pcStrPageDesc,"""","&quot;")
		pcStrPageDesc=replace(pcStrPageDesc,"<","&lt;")
		pcStrPageDesc=replace(pcStrPageDesc,">","&gt;")
	end if
end if

set rs=nothing
%>
<script type=text/javascript>
function newWindow(file,window)
{
	msgWindow=open(file,window,'resizable=no,width=400,height=500');
	if (msgWindow.opener == null) msgWindow.opener = self;
}
</script>
<form action="manageContactPage.asp?action=upd" method="post" name="manageContactPage" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
			<td colspan="2">This page controls the way <a href="../pc/contact.asp" target="_blank">Contact Page</a> is shown in the storefront.</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Display Settings</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr valign="top">
			<td width="25%">
			Page Description:
				<br />
				<br />
				<span class="pcCPnotes">If empty, no page description is shown at the top of the page</span>
			</td>
			<td width="75%">
				<textarea class="htmleditor" id="pcStrPageDesc" name="pcStrPageDesc" cols="65" rows="6"><%=pcStrPageDesc%></textarea>
				<br />
			</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> 
			<input type="submit" name="modify" value="Update Settings" class="btn btn-primary">
            &nbsp;
            <input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
