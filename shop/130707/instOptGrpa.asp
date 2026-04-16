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
dim  prdFrom, AssignID

'get product ID of the referring product page, if any
prdFrom = request.QueryString("prdFrom")
 if prdFrom = "" then
  prdFrom = 0
 end if
AssignID = request.QueryString("AssignID")
 if AssignID = "" then
  AssignID = 0
 end if
 
validateForm "instOptGrpb.asp"
%>
<% pageTitle="Add New Option Group" %>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
<form method="post" action="instOptGrpa.asp" name="addOpt" class="pcForms">
<% validateError %>
<table class="pcCPcontent"> 
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="20%" align="right" nowrap>Name this new option group:</td>
<td width="80%">  
<%textbox "optionGroupDesc", "", 30, "textbox"%>
<%validate "optionGroupDesc", "required"%>
</td>
</tr>
<tr> 
	<td>&nbsp;</td>
	<td>Example: &quot;Size&quot;, &quot;Color&quot;, &quot;Style&quot;, etc.</td>
</tr>
<%
If scSearch_IsEnabled = True Then
    query="SELECT pcFG_ID,pcFG_Name FROM pcFacetGroups ORDER BY pcFG_Name ASC;"
    set rs=connTemp.execute(query)
    if not rs.eof then
        tmpArr=rs.getRows()
        set rs=nothing
        intCount=ubound(tmpArr,2)
        %>
        <tr>
            <td width="20%" align="right" nowrap>Facet Group:</td>
            <td width="80%"> 
                <select name="pcFGID" id="pcFGID">
                    <option value="0"></option>
                    <% For i=0 to intCount %>
                        <option value="<%=tmpArr(0,i)%>"><%=tmpArr(1,i)%></option>
                    <% Next %>
                </select>
            </td>
        </tr>
        <%
    end if
    set rs=nothing
End If
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td></td>
	<td>
		<!-- send product ID of the referring product page, if any -->
		<input type="hidden" name="prdFrom" value="<%=prdFrom%>">
		<input type="hidden" name="AssignID" value="<%=AssignID%>">
					 
		<input type="submit" name="Submit" value="Save" class="btn btn-primary">&nbsp;
		<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:history.back()">
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
