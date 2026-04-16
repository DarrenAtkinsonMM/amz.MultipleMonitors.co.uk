<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
Dim pageTitle, Section
Section="layout" 

Dim lngIDPro,strPName,strPImg,intShowMessage

lngIDPro=getUserInput(request("IDPro"),0)

if request("action")="update" then
	strPName=getUserInput(request("PName"),0)
	strPImg=getUserInput(request("PImg"),0)
	query="Update pcFTypes set pcFType_Name=N'" & strPName &"', pcFType_Img='" & strPImg &"' where pcFType_IDType=" & lngIDPro
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	intShowMessage=1
end if

query="select * from pcFTypes where pcFType_IDType=" & lngIDPro
set rs=connTemp.execute(query)
strPName=rs("pcFType_name")
strPImg=rs("pcFType_img")
set rs=nothing


pageTitle="Edit Feedback Type: " & strPName & " (" &  lngIDPro & ")"
%>
<!--#include file="AdminHeader.asp" -->

<script type=text/javascript>
function Form1_Validator(theForm)
{
	if (theForm.PName.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.PName.focus();
		    return (false);
	}
return (true);
}
</script>
<form name="search" method="post" action="adminEditFBType.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
<Input type="hidden" name="IDPro" value="<%=lngIDPro%>">
<table class="pcCPcontent">
  <tr>
    <td colspan="2" class="pcCPspacer" align="center">
			<% if intShowMessage=1 then %>
			<div class="pcCPmessageSuccess">This Feedback Type was updated successfully!</div>
			<% end if %>
		</td>
  </tr>
  <tr>
    <td align="right" width="10%">Name:</td>
    <td width="90%"><input name="PName" type="text" value="<%=strPName%>" size="25"></td>
  </tr>
  <tr>
    <td align="right" nowrap> Image File:</td>
    <td valign="top"><input name="PImg" type="text" value="<%=strPImg%>" size="25"> Type in the file name. To upload an image <a href="#" onClick="window.open('adminimageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a></td>
  </tr>
  <tr>
    <td colspan="2" class="pcCPspacer"></td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td>
        <input type="submit" name="Submit" value="Update" class="btn btn-primary">
        &nbsp;<input type="button" class="btn btn-default"  name="back" value="Back" onClick="location='adminFBTypeManager.asp';">
	</td>
  </tr>
</table>
</form>
<!--#include file="Adminfooter.asp" -->
