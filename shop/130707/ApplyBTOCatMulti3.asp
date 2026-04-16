<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%> 
<% pageTitle = "Change category settings across multiple configurable products" %>
<% section = "services" %>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<% 
dim f, pidProduct
%>
<!--#include file="AdminHeader.asp"-->
<%IF request("action")="upd" THEN
		pcv_idcategory=session("cp_bto_ar2_idcategory")
		pcv_btolist=session("cp_bto_ar2_btolist")
		pcArr1=split(pcv_btolist,",")
		
		'Get Category Settings
		pcv_catsort=request("pcv_catsort")
		if not validNum(pcv_catsort) then pcv_catsort="0"

		pcv_showInfo=request("pcv_showInfo")
		if not validNum(pcv_showInfo) then pcv_showInfo="0"

		pcv_requiredCategory=request("pcv_requiredCategory")
		if not validNum(pcv_requiredCategory) then pcv_requiredCategory="0"

		pcv_displayQF=request("pcv_displayQF")
		if not validNum(pcv_displayQF) then pcv_displayQF="0"

		pcv_notes=request("pcv_Notes")
		if pcv_notes<>"" then
			pcv_notes=replace(pcv_notes,"'","''")
		end if
		
		pcv_ShowDesc=request("pcv_ShowDesc")
		if not validNum(pcv_ShowDesc) then pcv_ShowDesc="0"

		pcv_ShowImg=request("pcv_ShowImg")
		if not validNum(pcv_ShowImg) then pcv_ShowImg="0"

		pcv_ImgWidth=request("pcv_ImgWidth")
		if not validNum(pcv_ImgWidth) then pcv_ImgWidth="0"

		pcv_ShowSKU=request("pcv_ShowSKU")
		if not validNum(pcv_ShowSKU) then pcv_ShowSKU="0"

		pcv_UseRadio=request("pcv_UseRadio")
		if not validNum(pcv_UseRadio) then pcv_UseRadio="0"
		If pcv_UseRadio="2" then
			pcv_UseRadio="0"
			pcv_multiSelect="1"
		else
			pcv_multiSelect="0"
		end If
				
		For j=lbound(pcArr1) to ubound(pcArr1)
			if pcArr1(j)<>"" then
					query="UPDATE configSpec_products SET showInfo=" & pcv_showInfo & ", requiredCategory=" & pcv_requiredCategory & ", multiSelect=" & pcv_multiSelect & ",catSort=" & pcv_catsort & ",displayQF=" & pcv_displayQF & ",notes=N'" & pcv_Notes & "',pcConfPro_ShowDesc=" & pcv_ShowDesc & ",pcConfPro_ShowImg=" & pcv_ShowImg & ",pcConfPro_ImgWidth=" & pcv_ImgWidth & ",pcConfPro_ShowSKU=" & pcv_ShowSKU & ",pcConfPro_UseRadio=" & pcv_UseRadio & " WHERE specProduct=" & pcArr1(j) & " AND configProductCategory=" & pcv_idcategory & ";" 
					set rs=Server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
					set rs=nothing
					
					call updPrdEditedDate(pcArr1(j))
					
				end if
			Next
	set rs=nothing
	
	msg="Category settings were successfully applied to the selected configurable products"

	session("cp_bto_ar2_idcategory")=""
	session("cp_bto_ar2_btolist")=""%>
	<table class="pcCPcontent">
	<tr>
		<td>
			<div class="pcCPmessageSuccess">
				<%=msg%>
			</div>
		</td>
	</tr>
	<tr>
		<td class="pcSpacer">&nbsp;</td>
	</tr>
	<tr>
		<td>
			<ul class="pcListIcon">
				<li><a href="LocateProducts.asp?cptype=1">Locate a configurable product</a></li>
				<li><a href="ApplyBTOCatMulti1.asp">Update another category across multiple configurable products</a></li>
			</ul>
		</td>
	</tr>
	</table>
<%
ELSE
If request("action")="add" then
	session("cp_bto_ar2_btolist")=request("prdlist")
	query="SELECT categoryDesc FROM categories WHERE idcategory=" & session("cp_bto_ar2_idcategory")
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_CatDesc=rs("categoryDesc")
	end if
	set rs=nothing
%>
<form method="post" name="modifyProduct" action="ApplyBTOCatMulti3.asp?action=upd" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th colspan="2">Category: <%=pcv_CatDesc%></th>
	</tr>
    <tr> 
		<td colspan="2"><strong>Display Options</strong></td>
	</tr>
    <tr>
		<td>&nbsp;</td>
		<td>
			<input type="radio" name="pcv_UseRadio<%=oCnt%>" value="0" checked class="clearBorder"> Display choices using radio buttons<br>
			<input type="radio" name="pcv_UseRadio<%=oCnt%>" value="1" class="clearBorder"> Display choices using drop down menus<br>
			<input type="radio" name="pcv_UseRadio<%=oCnt%>" value="2" class="clearBorder"> Display choices using check boxes
		</td>
	</tr>
	<tr>
		<td>
			<input type="checkbox" name="pcv_showInfo" value="1" class="clearBorder">
		</td>
		<td>
			Show Details
		</td>
	</tr>
	<tr>
		<td>
			<input type="checkbox" name="pcv_requiredCategory" value="1" class="clearBorder">
		</td>
		<td>
			Required Category
		</td>
	</tr>
	<tr>
		<td>
			<input type="checkbox" name="pcv_displayQF" value="1" class="clearBorder">
		</td>
		<td>
			Display Quantity Field
		</td>
	</tr>
	<tr>
		<td>
			<input type="checkbox" name="pcv_ShowDesc" value="1" class="clearBorder">
		</td>
		<td>
			Show Product Description
		</td>
	</tr>
	<tr>
		<td>
			<input type="checkbox" name="pcv_ShowImg" value="1" class="clearBorder">
		</td>
		<td>
			Show Item Image
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			Image Width = <input type="text" name="pcv_ImgWidth" value="0" size="4">
		</td>
	</tr>
	<tr>
		<td>
			<input type="checkbox" name="pcv_ShowSKU" value="1" class="clearBorder">
		</td>
		<td>
			Show Item SKU
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			Category Order: <input type="text" name="pcv_catsort" value="0" size="4">
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			Configuration instructions:<br>
			<textarea rows="3" name="pcv_Notes" cols="80"></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Apply to Configurable Products:</th>
	</tr>
	<% 
	pcv_btolist=session("cp_bto_ar2_btolist")
	pcArr=split(pcv_btolist,",")
	For i=lbound(pcArr) to ubound(pcArr)
	if pcArr(i)<>"" then
		query="SELECT description FROM products WHERE idproduct=" & pcArr(i) & " AND removed=0;"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			pcv_PrdDesc=rstemp("description")
			set rstemp=nothing %>
			<tr> 
				<td width="2%" align="right">&nbsp;</td>
				<td><%=pcv_PrdDesc%></td>
			</tr>
		<%end if
		set rstemp=nothing
	end if
	Next
	%>
	<tr> 
		<td colspan="2" align="center">
		<input type="submit" name="Submit1" value="Apply Settings" class="btn btn-primary">&nbsp;
		<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
</table>
</form>
<%
end if
END IF

%>
<!--#include file="Adminfooter.asp"-->
