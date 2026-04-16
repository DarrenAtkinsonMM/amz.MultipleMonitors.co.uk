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
	


idproduct=request("idproduct")
overwrite=request("overwrite")

if idproduct<>"" and overwrite="" then
	response.redirect "ApplyLayoutToMul.asp?action=apply&prdlist=" & idproduct
end if

if overwrite<>"" then
	session("ALPoverwrite")=overwrite
else
	overwrite=session("ALPoverwrite")
end if

if idproduct<>"" then
	session("ALPidproduct")=idproduct
else
	idproduct=session("ALPidproduct")
end if

if request("action")="apply" then
	if idproduct<>"" then
		query="SELECT pcProd_Top,pcProd_TopLeft,pcProd_TopRight,pcProd_Middle,pcProd_Bottom,pcProd_Tabs FROM products WHERE idproduct="&idproduct&";"
		set rs=conntemp.execute(query)
		if not rs.eof then
			ppTop=rs("pcProd_Top")
			ppTopLeft=rs("pcProd_TopLeft")
			ppTopRight=rs("pcProd_TopRight")
			ppMiddle=rs("pcProd_Middle")
			ppTabs=rs("pcProd_Tabs")
			ppBottom=rs("pcProd_Bottom")
			if overwrite="0" then
				'Clear HTML Content
				if ppTabs<>"" then
					ppTabsNew=""
					tmpStr1=split(ppTabs,"||")
					For i=lbound(tmpStr1) to ubound(tmpStr1)
						if tmpStr1(i)<>"" then
							tmpStr2=split(tmpStr1(i),"``")
							tmpStr2(2)=""
							ppTabsNew=ppTabsNew & tmpStr2(0) & "``" & tmpStr2(1) & "``" & tmpStr2(2) & "||"
						end if
					next
					ppTabs=ppTabsNew
				end if
			end if

			if (request("prdlist")<>"") and (request("prdlist")<>",") then
				prdlist=split(request("prdlist"),",")
				For i=lbound(prdlist) to ubound(prdlist)
					id=prdlist(i)
					If (id<>"0") and (id<>"") then
						query="UPDATE products SET pcprod_DisplayLayout='t',pcProd_Top='" & replace(ppTop, "'", "''") & "',pcProd_TopLeft='" & replace(ppTopLeft, "'", "''") & "',pcProd_TopRight='" & replace(ppTopRight, "'", "''") & "',pcProd_Middle='" & replace(ppMiddle, "'", "''")& "'"

                        if overwrite<>"2" then
                            query=query&",pcProd_Tabs=N'" & replace(ppTabs, "'", "''") & "'"
                        end if

                        query=query&",pcProd_Bottom='" & replace(ppBottom, "'", "''") & "' WHERE idproduct=" & id
						set rstemp1=connTemp.execute(query)
						call pcs_hookProductModified(id, "")
					end if
				Next
			end if
			set rs=nothing
			
			pcMessage="Applied product layout to selected products successfully!"
			success=1
		end if
		set rs=nothing
		
	else
		pcMessage="Please select a source product before applying the product layout."
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
end if
' END show message %>

<table id="FindProducts" class="pcCPcontent">
	<tr>
		<td>
		<%
			src_FormTitle1="Find Products"
			src_FormTitle2="Apply Product Layout to multiple products"
			src_FormTips1="Use the following filters to look for products in your store."
			src_FormTips2="Select which products you would like to apply Product Layout to."
			src_IncNormal=1
			src_IncBTO=1
			src_IncItem=0
			src_DisplayType=1
			src_ShowLinks=0
			src_FromPage="ApplyLayoutToPrds.asp"
			src_ToPage="ApplyLayoutToPrds.asp?action=apply"
			src_Button1=" Search "
			src_Button2=" Add Product Layout to Selected Products "
			src_Button3=" Back "
			src_PageSize=15
			UseSpecial=0
			session("srcprd_from")=""
			session("srcprd_where")=""
		%>
			<!--#include file="inc_srcPrds.asp"-->
		</td>
	</tr>
</table>
	
<!--#include file="AdminFooter.asp"-->
