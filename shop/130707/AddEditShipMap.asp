<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<%

tmpSMid=getUserInput(request("id"),0)
if tmpSMid="" then
	tmpSMid=0
end if

tmpSMName=""
tmpSMType="0"
tmpSMOrder="0"
pmode="add"

if tmpSMid<>"0" then
	query="SELECT pcSM_Name,pcSM_Type,pcSM_Order FROM pcShippingMap WHERE pcSM_ID=" & tmpSMid & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpSMName=rs("pcSM_Name")
		tmpSMType=rs("pcSM_Type")
		tmpSMOrder=rs("pcSM_Order")
		pmode="edit"
	end if
	set rs=nothing
end if

%>
<%
if pmode="add" then
	pageTitle="Add Shipping Filter"
else
	pageTitle="View/ Edit Shipping Filters"
end if
%>
<% Section="shipOpt" %>
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>
function Form1_Validator(theForm)
{
	if (theForm.pcSMName.value == "")
  	{
			alert("Please enter value for the filter name");
		    theForm.pcSMName.focus();
		    return (false);
	}
	if (theForm.pcSMMap.value == "")
  	{
			alert("Please select shipping methods to map");
		    theForm.pcSMMap.focus();
		    return (false);
	}
	return (true);
}
</script>
<form method="post" action="AddEditShipMapB.asp?action=go" name="addFG" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" name="id" value="<%=tmpSMid%>">
<table class="pcCPcontent"> 
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="20%" align="right" nowrap>Filter name:</td>
<td width="80%">
	<input type="text" name="pcSMName" id="pcSMName" size="30" value="<%=tmpSMName%>">
</td>
</tr>
<tr> 
	<td>&nbsp;</td>
	<td><i>Generic shipping name. Example: &quot;2 days shipping&quot;, etc.</i></td>
</tr>
<%
queryQ="SELECT idshipservice FROM pcSMRel WHERE pcSM_ID=" & tmpSMid & ";"
set rsQ=connTemp.execute(queryQ)
intH=-1
if not rsQ.eof then
	HArr=rsQ.getRows()
	intH=ubound(HArr,2)
end if
set rsQ=nothing

tmpRel=""
if tmpSMid<>"0" then
	tmpQ=" WHERE pcSM_ID<>" & tmpSMid
else
	tmpQ=""
end if
queryQ="SELECT idshipservice,serviceDescription FROM shipService  WHERE (shipService.serviceActive<>0) AND (shipService.idshipservice NOT IN (SELECT idshipservice FROM pcSMRel " & tmpQ & ")) ORDER BY idshipservice ASC;"
set rsQ=connTemp.execute(queryQ)
if not rsQ.eof then
	RelArr=rsQ.getRows()
	intR=ubound(RelArr,2)
	For iR=0 to intR
		if tmpRel<>"" then
			tmpRel=tmpRel & vbcrlf
		end if
		tmpH=0
		For iH=0 to intH
			if Clng(HArr(0,iH))=Clng(RelArr(0,iR)) then
				tmpH=1
				exit for
			end if
		Next
		tmpRel=tmpRel & "<option value=""" & RelArr(0,iR) & """"
		if tmpH=1 then
			tmpRel=tmpRel & " selected"
		end if
		tmpRel=tmpRel & ">" & RelArr(1,iR) & "</option>"
	Next
end if
set rsQ=nothing
%>
<%if tmpRel<>"" then%>
<tr valign="top"> 
<td width="20%" align="right" nowrap>Map to Shipping Methods:</td>
<td width="80%">
	<select name="pcSMMap" id="pcSMMap" multiple size="5">
		<%=tmpRel%>
	</select>
	<br>
	<font color="#666666"><i>Notes: To select more than one method, keep the CTRL button down</i></font>
</td>
</tr>
<%end if%>
<tr> 
<td width="20%" align="right" nowrap>Display Type:</td>
<td width="80%">
	<select name="pcSMType" id="pcSMType">
		<option value="0" <%if tmpSMType<>"1" then%>selected<%end if%>>Lowest Shipping Rate Method</option>
		<option value="1" <%if tmpSMType="1" then%>selected<%end if%>>Highest Shipping Rate Method</option>
	</select>
</td>
</tr>
<tr> 
<td width="20%" align="right" nowrap>Order:</td>
<td width="80%">
	<input type="text" name="pcSMOrder" id="pcSMOrder" size="4" value="<%=tmpSMOrder%>">
</td>
</tr><tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td></td>
	<td>
		<input type="submit" name="Submit" value="Save" class="btn btn-primary">&nbsp;
		<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:location='manageShipMap.asp';">
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
