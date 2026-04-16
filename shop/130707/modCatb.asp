<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<% 
dim f
Dim pcCatArr,CatRecords,tmpCatList

pIdcategory=request("idcategory")
parent=request.form("parent")
top=request.form("top")
pIdParentCategory=request.form("idParentCategory")
pCategoryDesc=replace(request.form("categoryDesc"),"'","''")
pCategoryDesc=replace(pCategoryDesc,"&amp;","&")
pCategoryDesc=replace(pCategoryDesc,"&","&amp;")
pImage=request.form("image")
if pImage="" then
	pImage="no_image.gif"
end if
plargeImage=request.form("largeimage")
if plargeImage="" then
	plargeImage="no_image.gif"
end if

NotImg=request("NotImg")
if NotImg="" then
	NotImg="0"
end if

if NotImg="1" then
	pImage=""
	plargeImage=""
end if

pIntSubCategoryView=request.form("intSubCategoryView")
pIntCategoryColumns=request.form("intCategoryColumns")
pIntCategoryRows=request.form("intCategoryRows")
pStrPageStyle=request.form("strPageStyle")
pStrProductOrder=request.form("strProductOrder")
pIntProductColumns=request.form("intProductColumns")
pIntProductRows=request.form("intProductRows")
pIntFeaturedCategory=request.form("intFeaturedCategory")
pIntFeaturedCategoryImage=request.form("intFeaturedCategoryImage")
if NOT validNum(pIntSubCategoryView) then pIntSubCategoryView=0
if NOT validNum(pIntCategoryColumns) then pIntCategoryColumns=0
if NOT validNum(pIntCategoryRows) then pIntCategoryRows=0
if NOT validNum(pIntProductColumns) then pIntProductColumns=0
if NOT validNum(pIntProductRows) then pIntProductRows=0
if NOT validNum(pIntFeaturedCategory) then pIntFeaturedCategory=0
if NOT validNum(pIntFeaturedCategoryImage) then pIntFeaturedCategoryImage=0
if NOT validNum(HideDesc) then HideDesc=0
if NOT validNum(pcv_intRetailHide) then pcv_intRetailHide=0


SDesc=replace(request.form("SDesc"),"'","''")
LDesc=replace(request.form("LDesc"),"'","''")
HideDesc=request.form("HideDesc")

if not HideDesc<>"" then
	HideDesc="0"
end if

pBoton=request.form("modify")
piBTOhide=request.form("iBTOhide")
if piBTOhide="" then
	piBTOhide="0"
end if
	
pcv_intRetailHide=request.form("RetailHide")
if pcv_intRetailHide="" then
	pcv_intRetailHide="0"
end if

runSubCats=request.form("runSubCats")
if runSubCats="" then
	runSubCats=0
end if


'//Retrieve Category Level Product Display Setting
pcv_StrCatDisplayLayout=getUserInput(request.Form("CatDisplayLayout"),4)
if pcv_StrCatDisplayLayout="D" then pcv_StrCatDisplayLayout=""

'//Retrieve new Meta Tag related fields
pcv_StrCatMetaTitle=getUserInput(request.Form("CatMetaTitle"), 0)
pcv_StrCatMetaDesc=getUserInput(request.Form("CatMetaDesc"), 0)
pcv_StrCatMetaKeywords=getUserInput(request.Form("CatMetaKeywords"), 0)

pcv_StrAvalaraTaxCode=request("AvalaraTaxCode")
if pcv_StrAvalaraTaxCode<>"" then
	pcv_StrAvalaraTaxCode=replace(pcv_StrAvalaraTaxCode,"'","''")
end if


sub UpdateSubCats(tmpParent,CType,tmpValue)
	Dim rstemp,query,pcArr,i,intCount,tmpStr
	if CType="0" then
		tmpStr="iBTOhide="& tmpValue
	else
		tmpStr="pccats_RetailHide="& tmpValue
	end if
	
	query="UPDATE categories SET "& tmpStr &" WHERE idParentCategory=" & tmpParent
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
	
	call UpdCatEditedDate(tmpParent," idParentCategory=" & tmpParent)
	
	query="SELECT idcategory FROM categories WHERE idParentCategory=" & tmpParent
	set rstemp=connTemp.execute(query)
	if not rstemp.eof then
		pcArr=rstemp.GetRows()
		intCount=ubound(pcArr,2)
		set rstemp=nothing
		For i=0 to intCount
			call UpdateSubCats(pcArr(0,i),CType,tmpValue)
		Next
	end if
	set rstemp=nothing
end sub

sub UpdateSubCatsCSF(tmpParent)
	Dim rstemp,query,pcArr,i,intCount	
	SFData=request("SFData")
	query="DELETE FROM pcSearchFields_Categories WHERE idCategory=" & tmpParent & ";"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	call UpdCatEditedDate(tmpParent,"")
	if SFData<>"" then
		tmp1=split(SFData,"||")
		For i=0 to ubound(tmp1)
			if tmp1(i)<>"" then
				tmp2=split(tmp1(i),"^^^")
				idSearchData=tmp2(1)			
				query="INSERT INTO pcSearchFields_Categories (idCategory,idSearchData) VALUES (" & tmpParent & "," & idSearchData & ");"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
			end if
		Next
	end if		
	query="SELECT idcategory FROM categories WHERE idParentCategory=" & tmpParent
	set rstemp=connTemp.execute(query)
	if not rstemp.eof then
		pcArr=rstemp.GetRows()
		intCount=ubound(pcArr,2)
		set rstemp=nothing
		For i=0 to intCount
			call UpdateSubCatsCSF(pcArr(0,i))
		Next
	end if
	set rstemp=nothing
end sub

sub UpdateSubCatsDisplay(tmpParent)
	Dim rstemp,query,pcArr,i,intCount,tmpStr
	query="UPDATE categories SET HideDesc=" & HideDesc & ", pcCats_SubCategoryView="&pIntSubCategoryView&",  pcCats_CategoryColumns="&pIntCategoryColumns&", pcCats_CategoryRows="&pIntCategoryRows&", pcCats_PageStyle='"&pStrPageStyle&"', pcCats_ProductOrder='"&pStrProductOrder&"', pcCats_ProductColumns="&pIntProductColumns&", pcCats_ProductRows="&pIntProductRows&", pcCats_DisplayLayout='"&pcv_StrCatDisplayLayout&"' WHERE idParentCategory=" &tmpParent
  set rstemp=connTemp.execute(query)
	set rstemp=nothing
	
	call UpdCatEditedDate(tmpParent," idParentCategory=" & tmpParent)
	
	query="SELECT idcategory FROM categories WHERE idParentCategory=" & tmpParent
	set rstemp=connTemp.execute(query)
	if not rstemp.eof then
		pcArr=rstemp.GetRows()
		intCount=ubound(pcArr,2)
		set rstemp=nothing
		For i=0 to intCount
			call UpdateSubCatsDisplay(pcArr(0,i))
		Next
	end if
	set rstemp=nothing
end sub


'// START MODIFY category
IF trim(pboton)="Save" THEN
	
    '// identify tier of parent category and set tier + 1
    If pIdParentCategory>0 Then
    
        query="SELECT tier,iBTOhide from categories WHERE idCategory="& pIdParentCategory
        set rstemp=conntemp.execute(query)
        ptier=rstemp("tier")+1
        pcv_ParentiBTOhide=rstemp("iBTOhide")
        if IsNull(pcv_ParentiBTOhide) or pcv_ParentiBTOhide="" then
            pcv_ParentiBTOhide=0
        end if
        set rstemp=nothing
        
    Else
        pIdParentCategory = 1
        ptier = 0
    End If
    
	if pcv_ParentiBTOhide="1" then
		piBTOhide=pcv_ParentiBTOhide
	end if

	query="UPDATE categories set SDesc=N'" & SDesc & "', LDesc=N'" & LDesc & "', HideDesc=" & HideDesc & ", [image]='"& pImage &"', largeimage='"& plargeImage &"', categoryDesc=N'" &pCategoryDesc& "', idParentCategory="& pIdParentCategory &" , tier="& ptier &", iBTOhide="& piBTOhide&", pccats_RetailHide=" & pcv_intRetailHide & ", pcCats_SubCategoryView="&pIntSubCategoryView&",  pcCats_CategoryColumns="&pIntCategoryColumns&", pcCats_CategoryRows="&pIntCategoryRows&", pcCats_PageStyle='"&pStrPageStyle&"', pcCats_ProductOrder='"&pStrProductOrder&"', pcCats_ProductColumns="&pIntProductColumns&", pcCats_ProductRows="&pIntProductRows&", pcCats_FeaturedCategory="&pIntFeaturedCategory&", pcCats_FeaturedCategoryImage="&pIntFeaturedCategoryImage&", pcCats_DisplayLayout='"&pcv_StrCatDisplayLayout&"', pcCats_MetaTitle=N'"&pcv_StrCatMetaTitle&"', pcCats_MetaDesc=N'"&pcv_StrCatMetaDesc&"', pcCats_MetaKeywords=N'"&pcv_StrCatMetaKeywords&"', pcCats_NotImg=" & NotImg & ", pcCats_AvalaraTaxCode='" & pcv_StrAvalaraTaxCode & "' WHERE idCategory=" &pIdCategory
	set rstemp=conntemp.execute(query)
	
	call UpdCatEditedDate(pIdCategory,"")

	if err.number <> 0 then
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modCatb in Update: "&Err.Description) 
	end If

	if runSubCats="1" then
		call UpdateSubCatsDisplay(pIdCategory)
		'Update iBTOhide for all sub-categories
		call UpdateSubCats(pIdCategory,0,piBTOhide)
		call UpdateSubCats(pIdCategory,1,pcv_intRetailHide)
	end if

	'// Remove any categories that contain a breadcrumb for this category
	query="UPDATE categories SET pccats_BreadCrumbs='' WHERE pccats_BreadCrumbs LIKE '%"&pIdCategory&"||%';"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	
	'--------------------------------------------------------------
	' START - Update breadcrumb navigation in case the category was moved
	'--------------------------------------------------------------
	dim arrCategories(999,4)
	indexCategories=0
	pUrlString=Cstr("")
	pIdCategory2=pidCategory

	' load category array with all categories until parent
	do while pIdCategory2>1
		query="SELECT categoryDesc, idCategory, idParentcategory, largeimage, SDesc, LDesc, HideDesc FROM categories WHERE idCategory=" & pIdCategory2 &" ORDER BY priority, categoryDesc ASC"
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)

		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rs=nothing
			'//close any connections
			
			'//redirect to error page
			call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
 
		if rs.eof then
			set rs=nothing
			
			call closeDb()
response.redirect "msg.asp?message=86"           
		end if
		
		'categoryDesc, idCategory, idParentcategory, largeimage, SDesc, LDesc, HideDesc
		if pIdCategory2=pidCategory then
			pCategoryName=rs("categoryDesc")
			intIdCategory=rs("idCategory")
			intIdParentCategory=rs("idParentCategory")
			plargeImage=rs("largeimage")
			if pLargeImage = "no_image.gif" then
				pLargeImage = ""
			end if
			SDesc=rs("SDesc")
			LDesc=rs("LDesc")
			HideDesc=rs("HideDesc")
			if isNULL(HideDesc) OR HideDesc="" then
				HideDesc="0"
			end if
		else
			pCategoryName=rs("categoryDesc")
			intIdCategory=rs("idCategory")
			intIdParentCategory=rs("idParentCategory")
		end if
		
		pIdCategory3=intIdParentCategory 
		arrCategories(indexCategories,0)=pCategoryName
		arrCategories(indexCategories,1)=intIdCategory
		arrCategories(indexCategories,2)=intIdParentCategory
		pIdCategory2=pIdCategory3
		indexCategories=indexCategories + 1   
	loop
	set rs=nothing
	
	'create new breadcrumb and enter it into database
	strBreadCrumb=""
	for f=indexCategories-1 to 0 step -1
		If arrCategories(f,2)="1" Then
			strDBBreadCrumb=strDBBreadCrumb&arrCategories(f,1)&"||"&arrCategories(f,0)
			strBreadCrumb=strBreadCrumb & "<a href='viewCategories.asp?idCategory=" &arrCategories(f,1) & "'>" & arrCategories(f,0) &"</a>"
		Else
			strDBBreadCrumb=strDBBreadCrumb&"|,|"&arrCategories(f,1)&"||"&arrCategories(f,0)
			strBreadCrumb=strBreadCrumb & " > " & "<a href='viewCategories.asp?idCategory=" &arrCategories(f,1) & "'>" & arrCategories(f,0) &"</a>"
		End If
	next
	'enter BreadCrumb into database
	query="UPDATE categories SET pccats_BreadCrumbs=N'"&replace(strDBBreadCrumb,"'","''")&"' WHERE idCategory="&pIdCategory&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	'--------------------------------------------------------------
	' END - Update breadcrumb
	'--------------------------------------------------------------

	
	'// Custom Search Fields
	if runSubCats="1" then
		'// Update Category Search Fields for all sub-categories
		call UpdateSubCatsCSF(pIdCategory)
	else
		'// Update Category Search Fields
		SFData=request("SFData")
		query="DELETE FROM pcSearchFields_Categories WHERE idCategory=" & pIdCategory & ";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		if SFData<>"" then
			tmp1=split(SFData,"||")
			For i=0 to ubound(tmp1)
				if tmp1(i)<>"" then
					tmp2=split(tmp1(i),"^^^")
					idSearchData=tmp2(1)			
					query="INSERT INTO pcSearchFields_Categories (idCategory,idSearchData) VALUES (" & pIdCategory & "," & idSearchData & ");"
					set rsQ=connTemp.execute(query)
					set rsQ=nothing
				end if
			Next
		end if		
	end if
	
	'Update Category Tree XML Cache
	%>
	<!--#include file="inc_genCatXML.asp"-->
	<%
	
	call closeDb()
response.redirect "modCata.asp?idcategory="&pIdCategory&"&top="&top&"&parent="&parent&"&update=1&s=1&message=OK1"
	
'// END MODIFY

ELSE

'// START DELETE category
	tmpCatList=pIdCategory
	
	query="SELECT idcategory,idParentCategory FROM categories;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcCatArr=rs.GetRows()
		set rs=nothing
		CatRecords=ubound(pcCatArr,2)
		Call LoopCats(pIdCategory)
	end if
	set rs=nothing

	' Verify assignment products
	query="SELECT TOP 1 products.idProduct FROM products INNER JOIN categories_products ON Products.idProduct=categories_products.idProduct WHERE products.removed=0 AND categories_products.idCategory IN (" & tmpCatList & ");"
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		set rstemp = nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in modCatb: "&Err.Description) 
	end If
	
	if not rstemp.eof then
		set rstemp = nothing
		
		call closeDb()

		Session("message") = "The category and/or its subcategories are not empty. You must remove all products from a category before deleting it. <br><br><A href=editCategories.asp?prdType=1&lid="& pIdCategory &">View products</a> | <a href=manageCategories.asp>Manage other categories</a>"
		response.redirect "msgb.asp"
	end if
	
	tmpCatList=split(tmpCatList,",")
	For i=0 to ubound(tmpCatList)
		if trim(tmpCatList(i))<>"" then
			DelCat(Clng(tmpCatList(i)))
		end if
	Next
	
	'Update Category Tree XML Cache
	%>
	<!--#include file="inc_genCatXML.asp"-->
	<%
	
	set rstemp = nothing
	
	call closeDb()
response.redirect "managecategories.asp?s=1&msg=" & Server.URLEncode("The category has been successfully deleted.")
	
'// END DELETE

END IF

Sub DelCat(tmpID)
Dim rs,query

' delete from categories_products
	query="DELETE FROM categories_products WHERE idCategory=" &tmpID
	
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		pcvErrDescription = Err.Description
		set rstemp = nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modCatb: "& pcvErrDescription) 
	end If
	
	' delete from categories
	query="DELETE FROM categories WHERE idCategory=" &tmpID
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		pcvErrDescription = Err.Description
		set rstemp = nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modCatb: "& pcvErrDescription) 
	end If
	
	'// Remove the cateogry from any search filters
	query="DELETE FROM pcSearchFields_Categories WHERE idCategory=" & tmpID & ";"
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		pcvErrDescription = Err.Description
		set rstemp = nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modCatb: "& pcvErrDescription) 
	end If	
	
	'// Remove the cateogry from any electronic coupon filter
	query="DELETE FROM pcDFCats WHERE pcFCat_IDCategory=" & tmpID & ";"
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		pcvErrDescription = Err.Description
		set rstemp = nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modCatb: "& pcvErrDescription) 
	end If	
	
	'// Remove any categories that contain a breadcrumb for this category
	query="UPDATE categories SET pccats_BreadCrumbs='' WHERE pccats_BreadCrumbs LIKE '%"&tmpID&"||%';"
	set rstemp=conntemp.execute(query)
	
End Sub

Sub LoopCats(IDParent)
	Dim m
	For m=0 to CatRecords
		if Clng(pcCatArr(1,m))=Clng(IDParent) then
			tmpCatList=tmpCatList & "," & pcCatArr(0,m)
			Call LoopCats(pcCatArr(0,m))
		end if
	Next
End Sub
%>
