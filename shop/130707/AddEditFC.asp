<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<%

tmpFCid=getUserInput(request("id"),0)
if tmpFCid="" then
	tmpFCid=0
end if

tmpFGid=getUserInput(request("idgroup"),0)
if tmpFGid="" then
	tmpFGid=0
end if

tmpFCName=""
tmpFCCode=""
tmpFCImg=""
tmpFCOrder=0
pmode="add"

if tmpFCid<>"0" then
	query="SELECT pcFG_ID,pcFC_Code,pcFC_Name,pcFC_Img,pcFC_Order FROM pcFacets WHERE pcFC_ID=" & tmpFCid & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpFGid=rs("pcFG_ID")
		tmpFCCode=rs("pcFC_Code")
		tmpFCName=rs("pcFC_Name")
		tmpFCImg=rs("pcFC_Img")
		tmpFCOrder=rs("pcFC_Order")
		pmode="edit"
	end if
	set rs=nothing
end if

%>
<%
if pmode="add" then
	pageTitle="Add New Facet"
else
	pageTitle="View/Edit Facet"
end if
%>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>
function Form1_Validator(theForm)
{
	if (theForm.pcFCName.value == "")
  	{
			alert("Please enter value for the Facet Name");
		    theForm.pcFCName.focus();
		    return (false);
	}
	if ((theForm.pcFCCode.value == "") && (theForm.pcFCImg.value == ""))
  	{
			alert("Please enter value for the Facet Code or Facet Image");
		    theForm.pcFCCode.focus();
		    return (false);
	}
	if (theForm.pcFCOrder.value == "")
  	{
			alert("Please enter value for the Sort Order");
		    theForm.pcFCOrder.focus();
		    return (false);
	}
	return (true);
}
</script>
<form method="post" action="AddEditFCb.asp?action=go" name="addFC" id="addFC" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" name="id" value="<%=tmpFCid%>">
<table class="pcCPcontent"> 
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="20%" align="right" nowrap>Facet Type:</td>
<td width="80%">
	<%query="SELECT pcFG_ID,pcFG_Name FROM PcFacetGroups ORDER BY pcFG_Name ASC;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		%>
		<select name="pcFGID">
		<%For i=0 to intCount%>
			<option value="<%=tmpArr(0,i)%>" <%if Clng(tmpArr(0,i))=Clng(tmpFGid) then%>selected<%end if%> ><%=tmpArr(1,i)%></option>
		<%Next%>
		</select>
	<%end if
	set rs=nothing%>
</td>
</tr>
<tr> 
<td width="20%" align="right" nowrap>Facet Name:</td>
<td width="80%">
	<input type="text" name="pcFCName" id="pcFCName" size="30" value="<%=tmpFCName%>">
</td>
</tr>
<tr> 
<td width="20%" align="right" nowrap>Facet Code:</td>
<td width="80%">
	<input type="text" name="pcFCCode" id="pcFCCode" size="30" value="<%=tmpFCCode%>">
</td>
</tr>
<tr>
<td width="20%" align="right" nowrap>Facet Image:</td>
<td width="80%">
	<input type="text" name="pcFCImg" id="pcFCImg" size="30" value="<%=tmpFCImg%>"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=pcFCImg&fid=addFC','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
	<% if tmpFCImg <> "" then %>
		<img src="../pc/catalog/<%=tmpFCImg%>"  border=0 align=absbottom>
	<% end if %>
</td>
</tr>
<tr>
<td width="20%" align="right" nowrap>Sort Order:</td>
<td width="80%">
	<input type="text" name="pcFCOrder" id="pcFCOrder" size="4" value="<%=tmpFCOrder%>">
</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td></td>
	<td>
		<input type="submit" name="Submit" value="Save" class="btn btn-primary">&nbsp;
		<%if tmpFGid>"0" then%>
			<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:location='ManageFacets.asp?id=<%=tmpFGid%>';">
		<%else%>
			<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:location='ManageFacetGroups.asp';">
		<%end if%>
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
