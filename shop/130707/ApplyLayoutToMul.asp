<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Apply Product Layout to multiple products" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% on error resume next
Dim rsOrd, pid

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="description"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If

if request("action")="go" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				response.clear()
				call closeDb()
				overwrite=request("overwrite")
				if overwrite="" then
					overwrite="0"
				end if
				response.redirect "ApplyLayoutToPrds.asp?idproduct=" & id & "&overwrite=" & request("overwrite")
			End if
		Next
	end if
end if
	


if request("action")="apply" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				exit for
			End if
		Next
	end if
	If (id="0") and (id="") then
		pcMessage="Please select product before copying Product Layout"
		success=0
	end if
end if
%>

<!--#include file="AdminHeader.asp"-->

<% ' START show message, if any
If pcMessage <> "" Then %>
<div <%if success=1 then%>class="pcCPmessageSuccess"<%else%>class="pcCPmessageInfo"<%end if%>>
	<%=pcMessage%>
</div>
<%else
if request("action")="apply" then%>
	<form name="form1" method="post" action="ApplyLayoutToMul.asp?action=go" class="pcForms">
	<input type="hidden" name="prdlist" value="<%=request("prdlist")%>">
	<div class="pcCPmessageInfo">
		NOTE: If you copy a product layout to multiple products, <strong>ALL</strong> custom HTML areas of the target products will either be overwritten or cleared by the options below.
	</div>
	<div>
	<br>
	</div>
	<div>
		When copying layout from the source product:<br>
		<input type="radio" name="overwrite" value="0" class="clearBorder" checked> Clear custom HTML content when copying from source product tabs.<br>
		<input type="radio" name="overwrite" value="1" class="clearBorder"> Also copy custom HTML content from source products tabs.<br>
		<input type="radio" name="overwrite" value="2" class="clearBorder"> Don't copy product tab layout and custom HTML content (safest option).<br><br>
	</div>
	<div>
		<input type="submit" name="submit" value="Continue" class="btn btn-primary">
	</div>
	</form>
	<!--#include file="AdminFooter.asp"-->
<%response.end
end if
end if
' END show message %>

<table id="FindProducts" class="pcCPcontent">
	<tr>
		<td>
		<%
			src_FormTitle1="Find Product"
			src_FormTitle2="Apply Product Layout to multiple products"
			src_FormTips1="Use the following filters to locate products which use a custom Product Page Layout, and then apply it to other products."
			src_FormTips2="Select which product you would like to copy Product Layout from."
			src_IncNormal=1
			src_IncBTO=1
			src_IncItem=0
			src_DisplayType=2
			src_ShowLinks=0
			src_FromPage="ApplyLayoutToMul.asp"
			src_ToPage="ApplyLayoutToMul.asp?action=apply"
			src_Button1=" Search "
			src_Button2=" Copy Layout from Selected Product "
			src_Button3=" Back "
			src_PageSize=15
			UseSpecial=1
			session("srcprd_from")=""
			session("srcprd_where")=" AND (pcprod_DisplayLayout='t') AND ((pcProd_Top<>'') OR (pcProd_TopLeft<>'') OR (pcProd_TopRight<>'') OR (pcProd_Middle<>'') OR (pcProd_Tabs<>'') OR (pcProd_Bottom<>''))"
		%>
			<!--#include file="inc_srcPrds.asp"-->
		</td>
	</tr>
</table>
	
<!--#include file="AdminFooter.asp"-->
