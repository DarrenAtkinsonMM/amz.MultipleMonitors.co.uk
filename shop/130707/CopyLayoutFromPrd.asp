<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Copy Product Layout From Another Product" %>
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
	


idproduct=request("idproduct")

if idproduct<>"" then
session("CLPidproduct")=idproduct
else
idproduct=session("CLPidproduct")
end if

if request("action")="go" then
	if idproduct<>"" then
			if (request("prdlist")<>"") and (request("prdlist")<>",") then
				prdlist=split(request("prdlist"),",")
				For i=lbound(prdlist) to ubound(prdlist)
					id=prdlist(i)
					If (id<>"0") and (id<>"") then
						query="SELECT pcProd_Top,pcProd_TopLeft,pcProd_TopRight,pcProd_Middle,pcProd_Bottom,pcProd_Tabs FROM products WHERE idproduct="&id&";"
						set rs=conntemp.execute(query)
						if not rs.eof then
							ppTop=rs("pcProd_Top")
							ppTopLeft=rs("pcProd_TopLeft")
							ppTopRight=rs("pcProd_TopRight")
							ppMiddle=rs("pcProd_Middle")
							ppTabs=rs("pcProd_Tabs")
							ppBottom=rs("pcProd_Bottom")
							overwrite=request("overwrite")
							if overwrite="" then
								overwrite="0"
							end if
							if overwrite="0" then
								'Clear HTML Content
								if ppTabs<>"" then
									ppTabsNew=""
									tmpStr1=split(ppTabs,"||")
									For k=lbound(tmpStr1) to ubound(tmpStr1)
										if tmpStr1(k)<>"" then
											tmpStr2=split(tmpStr1(k),"``")
											tmpStr2(2)=""
											ppTabsNew=ppTabsNew & tmpStr2(0) & "``" & tmpStr2(1) & "``" & tmpStr2(2) & "||"
										end if
									next
									ppTabs=ppTabsNew
								end if
							end if
							query="UPDATE products SET pcprod_DisplayLayout='t',pcProd_Top='" & replace(ppTop, "'", "''") & "',pcProd_TopLeft='" & replace(ppTopLeft, "'", "''") & "',pcProd_TopRight='" & replace(ppTopRight, "'", "''") & "',pcProd_Middle='" & replace(ppMiddle, "'", "''") & "'"
                            if overwrite<>"2" then
                                query=query&",pcProd_Tabs=N'" & replace(ppTabs, "'", "''") & "'"
                            end if
                            query=query&",pcProd_Bottom='" & replace(ppBottom, "'", "''") & "' WHERE idproduct=" & idproduct
							set rstemp1=connTemp.execute(query)
							call pcs_hookProductModified(idproduct, "")
						end if
						set rs=nothing
						Exit For
					End if
				Next
			end if
			set rs=nothing
			
			pcMessage="Copied product layout successfully!"
			success=1
	else
		pcMessage="Please select a product before copying the product layout."
		success=0
	end if
end if

if request("action")="apply" then
	if idproduct<>"" then
			if (request("prdlist")<>"") and (request("prdlist")<>",") then
				prdlist=split(request("prdlist"),",")
				For i=lbound(prdlist) to ubound(prdlist)
					id=prdlist(i)
					If (id<>"0") and (id<>"") then
						Exit For
					End if
				Next
				If (id="0") and (id="") then
					pcMessage="Please select a product before copying the product layout."
					success=0
				End if
			end if
	else
		pcMessage="Please select a product before copying the product layout."
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
<br>
<%
    query="SELECT sku,description FROM products WHERE idProduct = " & idProduct & ";"
    set rstemp=conntemp.execute(query)
    if not rstemp.eof then
        productName = """" & rstemp("sku") & " - " & rstemp("description") & """"
    else
        productName = "Product"
    end if
    set rstemp = nothing
%>
<a class="btn btn-default" href="modifyProduct.asp?idProduct=<%= idproduct %>#tab-7">Return to <%= productName %></a>
<br>
<!--#include file="AdminFooter.asp"-->
<%response.end
else
if request("action")="apply" then%>
	<form name="form1" method="post" action="CopyLayoutFromPrd.asp?action=go" class="pcForms">
	<input type="hidden" name="idproduct" value="<%=idproduct%>">
	<input type="hidden" name="prdlist" value="<%=request("prdlist")%>">
	<div class="pcCPmessageInfo">
		NOTE: If you copy a product layout from another product, <strong>ALL</strong> custom HTML areas of target product will either be overwritten or cleared by the options below.
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
			src_FormTitle2="Copy Product Layout from another product"
			src_FormTips1="Use the following filters to look for product in your store."
			src_FormTips2="Select which product you would like to copy Product Layout from."
			src_IncNormal=1
			src_IncBTO=1
			src_IncItem=0
			src_DisplayType=2
			src_ShowLinks=0
			src_FromPage="CopyLayoutFromPrd.asp"
			src_ToPage="CopyLayoutFromPrd.asp?action=apply"
			src_Button1=" Search "
			src_Button2=" Copy Layout from Selected Product "
			src_Button3=" Back "
			src_PageSize=15
			UseSpecial=1
			session("srcprd_from")=""
			session("srcprd_where")=" AND idProduct<>" & idproduct & " AND (pcprod_DisplayLayout='t') AND ((pcProd_Top<>'') OR (pcProd_TopLeft<>'') OR (pcProd_TopRight<>'') OR (pcProd_Middle<>'') OR (pcProd_Tabs<>'') OR (pcProd_Bottom<>''))"
		%>
			<!--#include file="inc_srcPrds.asp"-->
		</td>
	</tr>
</table>
	
<!--#include file="AdminFooter.asp"-->
