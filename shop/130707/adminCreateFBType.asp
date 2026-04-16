<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% Dim pageTitle, Section
pageTitle="Create Feedback Type"
Section="layout" %>
<!--#include file="AdminHeader.asp" -->
<%
if request("action")="create" then
	
	Dim strPName,strPImg,intShowImg
	strPName=getUserInput(request("PName"),0)
	strPImg=getUserInput(request("PImg"),0)
	query="select pcFType_ShowImg from pcFTypes order by pcFType_IDType asc"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
		if not rs.eof then
			intShowImg=rs("pcFType_ShowImg")
		else
			intShowImg="0"
		end if
	query="Insert Into pcFTypes (pcFType_Name,pcFType_Img,pcFType_ShowImg) values (N'" & strPName &"','" & strPImg &"'," & intShowImg & ")"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	call closeDb()
response.redirect "adminFBTypeManager.asp?s=1&msg=New Feedback Type added successfully!"
end if
%>
<script type=text/javascript>
function Form1_Validator(theForm)
{

	if (theForm.PName.value == "")
 	{
		    alert("Please enter a value for the Feedback Type name.");
		    theForm.PName.focus();
		    return (false);
	}
  
return (true);
}
</script>
<form name="search" method="post" action="adminCreateFBType.asp?action=create" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
  <tr>
    <td colspan="2" class="pcCPspacer"></td>
  </tr>
  <tr>
    <td align="right" width="10%">Name:</td>
    <td width="90%"><input name="PName" type="text" value="" size="25"></td>
  </tr>
  <tr>
    <td align="right" nowrap="nowrap">Image File:</td>
    <td valign="top" nowrap="nowrap"><input name="PImg" type="text" value="" size="25"> Type in the file name. To upload an image <a href="#" onClick="window.open('adminimageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a>
    </td>
  </tr>
  <tr>
    <td colspan="2" class="pcCPspacer"></td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td>
    	<input type="submit" name="Submit" value="Create" class="btn btn-primary">
        &nbsp;<input type="button" class="btn btn-default"  name="back" value="Back" onClick="location='adminFBTypeManager.asp';"></td>
  </tr>
</table>
</form>
<!--#include file="Adminfooter.asp" -->
