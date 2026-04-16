<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Assign/Remove Configurable-Only Items To/From Multiple Products" %>
<% section = "services" %>
<%PmAdmin=2%> 
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
	IF request("Submit2")<>"" THEN
		pcv_idcategory=session("cp_bto_ar1_idcategory")
		pcv_itemlist=session("cp_bto_ar1_itemlist")
		pcArr=split(pcv_itemlist,",")
		pcv_btolist=session("cp_bto_ar1_btolist")
		pcArr1=split(pcv_btolist,",")
		For i=lbound(pcArr) to ubound(pcArr)
			if pcArr(i)<>"" then
				For j=lbound(pcArr1) to ubound(pcArr1)
					if pcArr1(j)<>"" then
						query="DELETE FROM configSpec_products WHERE specProduct=" & pcArr1(j) & " AND configProduct=" & pcArr(i) & " AND configProductCategory=" & pcv_idcategory & ";"
						set rs=connTemp.execute(query)
						set rs=nothing
					end if
				Next
			end if
		Next
		
		call updPrdEditedDate(pcArr1(j))
		
		msg="Configurable-Only Items were successfully removed from the selected Configurable Products!"
		msgtype=1
	ELSE
		pcv_idcategory=session("cp_bto_ar1_idcategory")
		pcv_btolist=session("cp_bto_ar1_btolist")
		pcArr1=split(pcv_btolist,",")
		Count=request("iCnt")
		For i=1 to Count
			pcv_Price=0
			pcv_WPrice=0
			pcv_configPrd=request("configProduct" & i)
			if request("rPrice" & i)<>"" then
				pcv_Price=request("rPrice" & i)
			else
				if request("Price" & i)<>"" then
					pcv_Price=request("Price" & i)
				end if
				if request("WPrice" & i)<>"" then
					pcv_WPrice=request("WPrice" & i)
				end if
			end if
			
			pcv_Price = replacecomma(pcv_Price)
			pcv_WPrice = replacecomma(pcv_WPrice)
			
			For j=lbound(pcArr1) to ubound(pcArr1)
				if pcArr1(j)<>"" then
					'Get Category Settings
					pcv_catsort="0"
					pcv_showInfo="0"
					pcv_requiredCategory="0"
					pcv_multiSelect="0"
					pcv_displayQF="0"
					pcv_notes=""
					pcv_ShowDesc="0"
					pcv_ShowImg="0"
					pcv_ImgWidth="0"
					pcv_ShowSKU="0"
					query="SELECT catSort,showInfo, requiredCategory, multiSelect,displayQF,notes,pcConfPro_ShowDesc,pcConfPro_ShowImg,pcConfPro_ImgWidth,pcConfPro_ShowSKU FROM configSpec_products WHERE specProduct=" & pcArr1(j) & " AND configProductCategory=" & pcv_idcategory & ";"
					set rstemp=connTemp.execute(query)
					if not rstemp.eof then
						pcv_catsort=rstemp("catSort")
						pcv_showInfo=cint(rstemp("showInfo"))
						pcv_requiredCategory=cint(rstemp("requiredCategory"))
						pcv_multiSelect=cint(rstemp("multiSelect"))
						pcv_displayQF=cint(rstemp("displayQF"))
						pcv_notes=rstemp("notes")
						pcv_ShowDesc=rstemp("pcConfPro_ShowDesc")
						pcv_ShowImg=rstemp("pcConfPro_ShowImg")
						pcv_ImgWidth=rstemp("pcConfPro_ImgWidth")
						pcv_ShowSKU=rstemp("pcConfPro_ShowSKU")
					end if
					set rstemp=nothing
					
					query="SELECT configProduct FROM configSpec_products WHERE specProduct=" & pcArr1(j) & " AND configProduct=" & pcv_configPrd & ";"
					set rs=connTemp.execute(query)
					if rs.eof then
						query="INSERT INTO configSpec_products (specProduct, configProduct, price, Wprice, cdefault, showInfo, requiredCategory, multiSelect,prdSort,catSort,configProductCategory,displayQF,notes,pcConfPro_ShowDesc,pcConfPro_ShowImg,pcConfPro_ImgWidth,pcConfPro_ShowSKU) VALUES ("&pcArr1(j)&","&pcv_configPrd&","&pcv_Price&","&pcv_WPrice&",0," & pcv_showInfo & "," & pcv_requiredCategory & "," & pcv_multiSelect & ",0,"&pcv_catSort&","&pcv_idcategory&"," & pcv_displayQF & ",N'" & pcv_notes & "'," & pcv_ShowDesc & "," & pcv_ShowImg & "," & pcv_ImgWidth & "," & pcv_ShowSKU & ");"
						set rs=conntemp.execute(query)
					else
						query="UPDATE configSpec_products SET price=" & pcv_Price & ",Wprice=" & pcv_WPrice & " WHERE specProduct=" & pcArr1(j) & " AND configProduct=" & pcv_configPrd & ";"
						set rs=connTemp.execute(query)
					end if
					set rs=nothing
					
					call updPrdEditedDate(pcArr1(j))
					
				end if
			Next
		
		Next
		
		msg="Configurable-Only Items were successfully assigned to the selected Configurable Products!"
		msgtype=1
	END IF
	session("cp_bto_ar1_idcategory")=""
	session("cp_bto_ar1_itemlist")=""
	session("cp_bto_ar1_btolist")=""%>
	<table class="pcCPcontent">
	<tr>
		<td>
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
		</td>
	</tr>
	<tr>
		<td class="pcSpacer">&nbsp;</td>
	</tr>
	<tr>
		<td>
			<ul class="pcListIcon">
				<li><a href="LocateProducts.asp?cptype=1">Locate a configurable product</a></li>
				<li><a href="AddRmvBTOItemsMulti1.asp">Assign/Remove configurable-only items to/from multiple products</a></li>
			</ul>
		</td>
	</tr>
	</table>
<%
ELSE
If request("action")="add" then
	session("cp_bto_ar1_btolist")=request("prdlist")
	query="SELECT categoryDesc FROM categories WHERE idcategory=" & session("cp_bto_ar1_idcategory")
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_CatDesc=rs("categoryDesc")
	end if
	set rs=nothing
%>
<form method="post" name="modifyProduct" action="AddRmvBTOItemsMulti4.asp?action=upd" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th colspan="2">Category: <%=pcv_CatDesc%></th>
		<th nowrap align="center">Use Online Price</th>
		<th align="center" nowrap>New Price</th>
		<th align="center" nowrap>Wholesale</th>
	</tr>
	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="5">The prices that you specify here are used only if you are assigning items to the selected configurable products. If the items already exist, they are ignored. If you are removing items, disregard the price fields.</td>
	</tr>
	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
						
	<% 
	pcv_itemlist=session("cp_bto_ar1_itemlist")
	pcArr=split(pcv_itemlist,",")
	For i=lbound(pcArr) to ubound(pcArr)
	if pcArr(i)<>"" then
		query="SELECT description,price FROM products WHERE idproduct=" & pcArr(i) & " AND removed=0;"
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			iCnt = iCnt + 1
			pcv_PrdDesc=rstemp("description")
			pcv_PrdPrice=rstemp("price")
			set rstemp=nothing %>
			<tr> 
				<td width="2%" align="right">  
					<input type="hidden" name="configProduct<%=iCnt%>" value=<%=pcArr(i)%>>
				</td>
				<td width="60%"><%=pcv_PrdDesc%></td>
				<td nowrap><input type="checkbox" name="rPrice<%=iCnt%>" value="<%=money(pcv_PrdPrice)%>" class="clearBorder"><% response.write scCurSign&money(pcv_PrdPrice)%></td>
				<td nowrap><%=scCurSign%> <input type="text" name="price<%=iCnt%>" size="4" maxlength="10"></td>
				<td nowrap><%=scCurSign%> <input type="text" name="Wprice<%=iCnt%>" size="4" maxlength="10"></td>
				</tr>
		<%end if
		set rstemp=nothing
	end if
	Next
	%>
	<tr>
		<td colspan="5" class="pcCPspacer"><input type="hidden" name="iCnt" value=<%=iCnt%>></td>
	</tr>
	<tr>
		<td colspan="5"><hr></td>
	</tr>
	<tr>
		<td colspan="5">These changes will apply to the <strong>following configurable products</strong>:</td>
	</tr>
	<% 
	pcv_btolist=session("cp_bto_ar1_btolist")
	pcArr=split(pcv_btolist,",")
	For i=lbound(pcArr) to ubound(pcArr)
	if pcArr(i)<>"" then
		query="SELECT description FROM products WHERE idproduct=" & pcArr(i) & " AND removed=0;"
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			pcv_PrdDesc=rstemp("description")
			set rstemp=nothing %>
			<tr> 
				<td colspan="5"><p><%=pcv_PrdDesc%></p></td>
			</tr>
		<%end if
		set rstemp=nothing
	end if
	Next
	%>
	<tr>
		<td colspan="5"><hr></td>
	</tr>
	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td colspan="5" align="center">
		<input type="submit" name="Submit1" value="Assign to these Configurable Products" class="btn btn-primary">&nbsp;
		<input type="submit" name="Submit2" value="Remove from these Configurable Products" class="btn btn-primary">&nbsp;
		<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
</table>
</form>
<%end if
END IF
%>
<!--#include file="Adminfooter.asp"-->
