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

tmpFGid=getUserInput(request("id"),0)
if tmpFGid="" then
	tmpFGid=0
end if

tmpFGName=""
pmode="add"

if tmpFGid<>"0" then
	query="SELECT pcFG_Name FROM pcFacetGroups WHERE pcFG_ID=" & tmpFGid & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpFGName=rs("pcFG_Name")
		pmode="edit"
	end if
	set rs=nothing
end if

%>
<%
if pmode="add" then
	pageTitle="Add New Facet Group"
else
	pageTitle="View/Edit Facet Group"
end if
%>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>
function Form1_Validator(theForm)
{
	if (theForm.pcFGName.value == "")
  	{
			alert("Please enter value for the Facet Group name");
		    theForm.pcFGName.focus();
		    return (false);
	}
	return (true);
}
</script>
<form method="post" action="AddEditFGb.asp?action=go" name="addFG" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" name="id" value="<%=tmpFGid%>">
<table class="pcCPcontent"> 
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="20%" align="right" nowrap>Facet group name:</td>
<td width="80%">
	<input type="text" name="pcFGName" id="pcFGName" size="30" value="<%=tmpFGName%>">
</td>
</tr>
<tr> 
	<td>&nbsp;</td>
	<td>Example: &quot;Size&quot;, &quot;Color&quot;, &quot;Style&quot;, etc.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td></td>
	<td>
		<input type="submit" name="Submit" value="Save" class="btn btn-primary">&nbsp;
		<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:location='ManageFacetGroups.asp';">
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
