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

Dim lngIDPro
Dim strPName, strPImg, strPBgColor, intShowMessage
intShowMessage=0


lngIDPro=getUserInput(request("IDPro"),0)

if request("action")="update" then
	strPName=getUserInput(request("PName"),0)
	strPImg=getUserInput(request("PImg"),0)
	strPBgColor=getUserInput(request("PBgColor"),0)
	query="Update pcFStatus set pcFStat_Name=N'" & strPName &"', pcFStat_Img='" & strPImg &"',pcFStat_BgColor='" & strPBgColor & "' where pcFStat_IDStatus=" & lngIDPro
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	intShowMessage=1
end if

query="select * from pcFStatus where pcFStat_IDStatus=" & lngIDPro
set rs=connTemp.execute(query)
strPName=rs("pcFStat_name")
strPImg=rs("pcFStat_img")
strPBgColor=rs("pcFStat_BgColor")
set rs=nothing


pageTitle="Edit Message Status: " & strPName
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
<form name="form1" method="post" action="adminEditFBStatus.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
<Input type="hidden" name="IDPro" value="<%=lngIDPro%>">
<table class="pcCPcontent">
  <tr>
    <td colspan="2" class="pcCPspacer" align="center">
			<% if intShowMessage=1 then %>
			<div class="pcCPmessageSuccess">This Feedback Status was updated successfully!</div>
			<% end if %>
		</td>
  </tr>
  <tr>
    <td width="10%" align="right" valign="top">Name:</td>
    <td width="90%"><%if (lngIDPro="1") or (lngIDPro="2") then%><%=strPName%><input type="hidden" name="PName" value="<%=strPName%>"><br><span class="pcSmallText">Default status: cannot be edited or removed.</span><%else%><input name="PName" type="text" value="<%=strPName%>" size="25"><%end if%></td>
  </tr>
  <tr>
    <td align="right" valign="top" nowrap>Image File:</td>
    <td valign="top"><input name="PImg" type="text" value="<%=strPImg%>" size="25"> Type in the file name. To upload an image <a href="#" onClick="window.open('adminimageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a></td>
  </tr>
  <tr>
    <td align="right" nowrap>Background Color:</td>
    <td><input name="PBgColor" id="PBgColor" type="text" value="<%=strPBgColor%>" size="25"> <input type="button" class="btn btn-default"  value="Choose" id="Choose" onClick="PcjsColorChooser('Choose','PBgColor','value')" name="button">  
    </td>
  </tr>
  <tr>
    <td colspan="2" class="pcCPspacer"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
	<input type="submit" name="Submit" value="Update" class="btn btn-primary">
    <input type="button" class="btn btn-default"  name="back" value="Back" onClick="location='adminFBStatusManager.asp';">
	</td>
  </tr>
</table>
</form>
<!--#include file="AdminFooter.asp" -->
