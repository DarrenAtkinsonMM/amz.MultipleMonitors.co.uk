<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Dim pcCatArr,intCatCount

intCatCount=-1
query="SELECT categoryDesc, idCategory, idParentcategory, pccats_BreadCrumbs FROM categories WHERE idCategory>1 ORDER BY priority, categoryDesc ASC"
set rs=connTemp.execute(query)
if not rs.eof then
	pcCatArr=rs.getRows()
	set rs=nothing
	intCatCount=ubound(pcCatArr,2)
end if
set rs=nothing

if intCatCount>=0 then
	For ik=0 to intCatCount
	if IsNull(pcCatArr(3,ik)) OR pcCatArr(3,ik)="" then
	
		indexCategories=0
		redim arrCategories(999,4)
		pUrlString=Cstr("")
		pIdCategory=pcCatArr(1,ik)
		pIdCategory2=pcCatArr(1,ik)

		' load category array with all categories until parent
		do while pIdCategory2>1
			
			For j=0 to intCatCount
				if Clng(pcCatArr(1,j))=Clng(pIdCategory2) then
					pCategoryName=pcCatArr(0,j)
					intIdCategory=pcCatArr(1,j)
					intIdParentCategory=pcCatArr(2,j)
					pIdCategory3=intIdParentCategory 
					arrCategories(indexCategories,0)=pCategoryName
					arrCategories(indexCategories,1)=intIdCategory
					arrCategories(indexCategories,2)=intIdParentCategory
					pIdCategory2=pIdCategory3
					indexCategories=indexCategories + 1
					Exit For
				end if
			Next
		loop
		
		'create new breadcrumb and enter it into database
		strDBBreadCrumb=""
		for f1=indexCategories-1 to 0 step -1
			If arrCategories(f1,2)="1" Then
				strDBBreadCrumb=strDBBreadCrumb&arrCategories(f1,1)&"||"&arrCategories(f1,0)
			Else
				strDBBreadCrumb=strDBBreadCrumb&"|,|"&arrCategories(f1,1)&"||"&arrCategories(f1,0)
			End If
		next
		'enter BreadCrumb into database
		query="UPDATE categories SET pccats_BreadCrumbs=N'"&replace(strDBBreadCrumb,"'","''")&"' WHERE idCategory="&pIdCategory&";"
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		
	End if
	Next
End if
set rs=nothing

%>
<% pageTitle="Update Categories BreadCrumbs" %>
<% section="products" %>

<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr> 
    <td><div class="pcCPmessageSuccess">
		<p>&nbsp;</p>
		<p>Categories BreadCrumbs were updated successfully!</p>
		<p>&nbsp;</p>
    </div></td>
	</tr>
	<tr>
		<td>
      		<input type="button" value=" Manage Product Categories " onclick="javascript:location='manageCategories.asp';" class="btn btn-default">
		</td>
	</tr>   
</table>
<!--#include file="AdminFooter.asp"-->