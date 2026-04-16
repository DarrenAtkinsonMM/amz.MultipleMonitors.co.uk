<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>

<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<%
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "viewCategories.asp"

'*******************************
' Page Settings
'*******************************
Dim pcCategoryClass, pcCategoryHover, pcProductHover
pcCategoryClass 	= "pcShowCategory"
pcCategoryHover 	= "pcShowCategoryBgHover"
pcProductHover		= "pcShowProductBgHover"

'*******************************
' START: Check store on/off, start PC session, check affiliate ID
'*******************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*******************************
' END: Check store on/off, start PC session, check affiliate ID
'*******************************

dim pTempIntSubCategory
%>
<!--#include file="prv_getSettings.asp"-->
<%

pTempIntSubCategory=session("idCategoryRedirect")
if pTempIntSubCategory = "" then
	pTempIntSubCategory=getUserInput(request("idCategory"),10)
end if

'// Validate Category ID
	if not validNum(pTempIntSubCategory) then
		pTempIntSubCategory=""
	end if
	if pTempIntSubCategory="" or pTempIntSubCategory="0" then
		pTempIntSubCategory=1
	end if
intIdCategory=pTempIntSubCategory

'// Wholesale-only categories
If Session("customerType")=1 Then
	pcv_strTemp=""
else
	pcv_strTemp=" AND pccats_RetailHide<>1"
end if

'*******************************
' START Display Settings
'*******************************

pFeaturedCategory=0
pFeaturedCategoryImage=0

If validNum(pTempIntSubCategory) and pTempIntSubCategory<>1 then
	query="SELECT pcCats_SubCategoryView, pcCats_CategoryColumns, pcCats_CategoryRows, pcCats_PageStyle, pcCats_ProductOrder, pcCats_ProductColumns, pcCats_ProductRows, pcCats_FeaturedCategory, pcCats_FeaturedCategoryImage FROM categories WHERE (((idCategory)="&pTempIntSubCategory&")" & pcv_strTemp &");"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	if rs.EOF then
		set rs=nothing
		call closeDb()
		response.redirect "msg.asp?message=86"
	end if	
	
	Dim pIntSubCategoryView
	Dim pIntCategoryColumns
	Dim pIntCategoryRows
	Dim pIntProductColumns
	Dim pIntProductRows
	
	pIntSubCategoryView=rs("pcCats_SubCategoryView")
	pIntCategoryColumns=rs("pcCats_CategoryColumns")
	pIntCategoryRows=rs("pcCats_CategoryRows")
	pStrPageStyle=rs("pcCats_PageStyle")
	pStrProductOrder=rs("pcCats_ProductOrder")
	pIntProductColumns=rs("pcCats_ProductColumns")
	pIntProductRows=rs("pcCats_ProductRows")
	pFeaturedCategory=rs("pcCats_FeaturedCategory")
	pFeaturedCategoryImage=rs("pcCats_FeaturedCategoryImage")
	
	set rs=nothing
	
	Session("pStrPageStyle")=pStrPageStyle
End if
	
' START Load category-specific values. If empty, use storewide settings

' How sub-categories are displayed
' 	0 = in a list, with images
'	1 = in a list, without images
'	2 = drop-down
'	3 = default
'	4 = thumbnail only
if NOT validNum(pIntSubCategoryView) OR pIntSubCategoryView=3 then
	 pIntSubCategoryView=scCatImages
end if

' How many per row: number of columns
if NOT validNum(pIntCategoryColumns) OR pIntCategoryColumns=0 then
	pIntCategoryColumns=scCatRow
end if

' How many rows per page
if NOT validNum(pIntCategoryRows) OR pIntCategoryRows=0 then
	pIntCategoryRows=scCatRowsPerPage
end if

' How many products per row
if NOT validNum(pIntProductColumns) OR pIntProductColumns=0 then
	pIntProductColumns=scPrdRow
end if

' How many rows per page
if NOT validNum(pIntProductRows) OR pIntProductRows=0 then
	pIntProductRows=scPrdRowsPerPage
end if

' END Load category-specific values


' OVERRIDE page style: check to see if a querystring or a form is sending the page style.
Dim pcPageStyle, strSeoQueryString

pcPageStyle = LCase(getUserInput(Request("pageStyle"),1))

'// Check querystring saved to session by 404.asp
if pcPageStyle = "" then
	strSeoQueryString=lcase(session("strSeoQueryString"))
	if strSeoQueryString<>"" then
		if InStr(strSeoQueryString,"pagestyle")>0 then
			pcPageStyle=left(replace(strSeoQueryString,"pagestyle=",""),1)
		end if
	end if
end if

'// Category Level Settings
if pcPageStyle = "" then
	pcPageStyle = pStrPageStyle
end if

'// Global Settings
if isNULL(pcPageStyle) OR trim(pcPageStyle) = "" then
	pcPageStyle = LCase(bType)
end if

if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

' OTHER display settings
' These variables show/hide information when products are shown with Page Style = L or M
Dim pShowSKU, pShowSmallImg
pShowSKU = scShowSKU ' If 0, then the SKU is hidden
pShowSmallImg = scShowSmallImg ' If 0, then the small image is not shown
' Note: the size of the small image is set via the css/pcStorefront.css stylesheet

'FB-S
if (session("Facebook")="1") AND (session("pcFBS_CustomDisplay")="1") then
	pIntSubCategoryView=session("pcFBS_CatImages")
	pIntCategoryColumns=session("pcFBS_CatRow")
	pIntCategoryRows=session("pcFBS_CatRowsperPage")
	pIntProductColumns=session("pcFBS_PrdRow")
	pIntProductRows=session("pcFBS_PrdRowsPerPage")
	pcPageStyle = session("pcFBS_BType")
	pShowSKU = session("pcFBS_ShowSKU")
	pShowSmallImg = session("pcFBS_ShowSmallImg")
end if
'FB-E

'// Check For Mobile Storefront Overrides
If session("Mobile")="1" Then
	pIntSubCategoryViewBAK=pIntSubCategoryView
	pIntSubCategoryView=0
	pIntCategoryColumns=1
	pIntCategoryRows=10
	pIntProductColumns=1
	pIntProductRows=10
	pcPageStyle = "h"
End If

'*******************************
' END Display Settings
'*******************************


if pFeaturedCategory<>0 then
	pcv_strTemp=pcv_strTemp&" AND idCategory<>"&pFeaturedCategory & " "
end if

dim pIdCategory, pCategoryDesc, pcStrViewAll

rMode=server.HTMLEncode(request.querystring("mode"))
if rMode="" then
	iPageSize=(pIntProductColumns*pIntProductRows)
	iCatPageSize=(pIntCategoryColumns*pIntCategoryRows)
	If Request("page")="" Then
		iPageCurrent=1
	Else
		iPageCurrent=CInt(Request("page"))
	End If
end if

'// View All
pcStrViewAll = Lcase(getUserInput(Request("viewAll"),3))
if pcStrViewAll = "yes" then
	iPageSize = 9999
end if	

if NOT validNum(iPageSize) OR iPageSize=0 then
	iPageSize=5
end if

pIdCategory=session("idCategoryRedirect")
mIdCategory=session("idCategoryRedirect")
if pIdCategory="" then
	pIdCategory=getUserInput(request.querystring("idCategory"),10)
	mIdCategory=getUserInput(request.querystring("idCategory"),10)
	'// Validate Category ID
	if not validNum(pIdCategory) then
		pIdCategory=""          
	end if
	if not validNum(mIdCategory) then
		mIdCategory=""          
	end if
	
	if pIdCategory="" then
		pIdCategory=1
		mIdCategory=1
	end if
end if
session("idCategoryRedirect")=""

'*******************************
' get category tree array
'*******************************
if pIdCategory<>1 then %>
	<!--#include file="pcBreadCrumbs.asp"-->
<% end if

'*******************************
' End get category tree array
'*******************************

'*******************************
' Get sub-categories array
'*******************************
Dim intSubCatExist
Dim iCategoriesPageCount
intSubCatExist=0

IF pIdCategory=1 THEN
	scCatTotal=(pIntCategoryColumns*pIntCategoryRows)
	if pIntSubCategoryView="2" then
		scCatTotal=999999
	end if
	iCategoriesPageSize=scCatTotal
	if pcStrViewAll = "yes" then
		iCategoriesPageSize = 9999
	end if
	
	Dim pcInt_CategoriesPage
	pcInt_CategoriesPage=getUserInput(request("CategoriesPage"),10)
	if not validNum(pcInt_CategoriesPage) then
		iCategoriesPageCurrent=1
	Else
		iCategoriesPageCurrent=Cint(pcInt_CategoriesPage)
	End If

	query = "SELECT idCategory,categoryDesc,[image],idParentCategory,SDesc,HideDesc FROM Categories WHERE idParentCategory=1 AND idCategory<>1 AND iBTOhide=0 " & pcv_strTemp & " ORDER BY priority, categoryDesc ASC;"
	SET rs=Server.CreateObject("ADODB.RecordSet")

	rs.PageSize=iCategoriesPageSize
	pcv_strPageSize=iCategoriesPageSize
	rs.CacheSize=iCategoriesPageSize
		
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
	'// Page Count
	iCategoriesPageCount=rs.PageCount
	If Cint(iCategoriesPageCurrent) > Cint(iCategoriesPageCount) Then iCategoriesPageCurrent=Cint(iCategoriesPageCount)
	If Cint(iCategoriesPageCurrent) < 1 Then iCategoriesPageCurrent=1	
ELSE
	scCatTotal=(pIntCategoryColumns*pIntCategoryRows)
	if pIntSubCategoryView="2" then
		scCatTotal=999999
	end if
	iCategoriesPageSize=scCatTotal
	if pcStrViewAll = "yes" then
		iCategoriesPageSize = 9999
	end if
	
	pcInt_CategoriesPage=getUserInput(request("CategoriesPage"),10)
	if not validNum(pcInt_CategoriesPage) then
		iCategoriesPageCurrent=1
	else
		iCategoriesPageCurrent=Cint(pcInt_CategoriesPage)
	end if
	
	query = "SELECT idCategory, categoryDesc FROM Categories WHERE idParentCategory = " & pIdCategory & " AND idCategory<>1 AND iBTOhide=0 " & pcv_strTemp & " ORDER BY priority, categoryDesc ASC;"
	set rs=Server.CreateObject("ADODB.RecordSet")

	rs.PageSize=iCategoriesPageSize
	pcv_strPageSize=iCategoriesPageSize
	rs.CacheSize=iCategoriesPageSize
		
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
	'// Page Count
	iCategoriesPageCount=rs.PageCount
	If Cint(iCategoriesPageCurrent) > Cint(iCategoriesPageCount) Then iCategoriesPageCurrent=Cint(iCategoriesPageCount)
	If Cint(iCategoriesPageCurrent) < 1 Then iCategoriesPageCurrent=1	
END IF

If NOT rs.EOF Then
	rs.AbsolutePage=iCategoriesPageCurrent
	intSubCatExist=1
	SubCatArray=rs.GetRows(iCategoriesPageSize)
	intSubCatCount=ubound(SubCatArray,2)
End If

SET rs=nothing
'*******************************
' End get sub-categories array
'*******************************
%>

<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->

<div id="pcMain" class="pcViewCategories">
	<div class="pcMainContent">
    
		<% if pIdCategory=1 then %>
			<h1><% response.write dictLanguage.Item(Session("language")&"_titles_9")%></h1>
		<% else
			'*******************************
			' Show current category info
			'*******************************
			' Show BreadCrumbs - current category name and location - If subcategory z %>
			<h1><%=pCategoryName%></h1>
			<%
			' Display promotion message if any
			pcs_CategoryPromotionMsg
			' End display promotion
	
			'// SEO-S
			pMainCategoryName=pCategoryName
			'// SEO-E
			%>
			<div class="pcPageNav" id="pcBreadCrumbs" itemscope itemtype="http://schema.org/BreadcrumbList">
				<% 
				response.write dictLanguage.Item(Session("language")&"_viewCat_P_2")
				response.write strBreadCrumb
				intIdCategory=pIdCategory
		
				'// Load category discount icon
				%>
				 <!--#include file="pcShowCatDiscIcon.asp" -->
			</div>
			<% ' End Show BreadCrumbs
		end if
	
		' Show large category image
		'/////CSS NOTE: This will need a class for the large category
		if pLargeImage<>"" then %>
			<div class="pcShowCategoryLargeImage">
				<img src="<%=pcf_getImagePath("catalog",pLargeImage)%>" alt="<%=pCategoryName%>">
			</div>
            <div class="pcClear"></div>
		<% end if
		' End show large category image

		' Start Show long category description
		If pcf_HasHTMLContent(LDesc) And HideDesc <> "1" Then %>
			<div class="pcPageDesc"><%=pcf_FixHTMLContentPaths(LDesc)%></div>
		<% End If
		' End Show Categories Description
			
		'*******************************
		' Show subcategories, if any
		'*******************************
		if intSubCatExist=1 then
		
			if pIdCategory<>1 then %>
				<h3><%=dictLanguage.Item(Session("language")&"_viewCategories_2")%>&quot;<%=pCategoryName%>&quot;</h3>
			<% end if %>
			
			<% ' FIRST subcategory display option = Drop-down
			if pIntSubCategoryView="2" then %>
            
				<% if pFeaturedCategory<>0 then %>
					<!--#include file="pcShowCategoryFeatured.asp" -->
				<% end if %>
                
                <div class="pcShowContent">
                  <form class="pcForms">
                                <% if trim(pCategoryName)<>"" then %>
                        <%=dictLanguage.Item(Session("language")&"_viewCategories_3")%>&quot;<%=pCategoryName%>&quot;:&nbsp;
                    <% else %>
                      <%=dictLanguage.Item(Session("language")&"_viewCategories_6")%>
                    <% end if %>
                    <select onChange="window.location.href=this.options[selectedIndex].value" name="CatDropSelect">
                      <option>Browse Subcategories</option>
                        <% 	
                        if pIdCategory=1 then
                            pcv_mc=0 
                            Do While (pcv_mc < iCategoriesPageSize) And (pcv_mc < (intSubCatCount+1))
                                intIdCategory=SubCatArray(0,pcv_mc)
                                pcStrCategoryDesc=SubCatArray(1,pcv_mc)
                                '// Call SEO Routine
                                pcGenerateSeoLinks
                                '//
                                query="SELECT categories_products.idProduct FROM categories_products WHERE categories_products.idCategory = " & intIdCategory
                                
                                %>
                                <option value="<%=pcStrCatLink%>"><%=pcStrCategoryDesc%></option>
                                <% 	pcv_mc=pcv_mc+1
                            Loop
                        else
                            For pcv_mc=0 to intSubCatCount
                                intIdCategory=SubCatArray(0,pcv_mc)
                                pcStrCategoryDesc=SubCatArray(1,pcv_mc)
                                '// Call SEO Routine
                                pcGenerateSeoLinks
                                '//							
                                query="SELECT categories_products.idProduct FROM  categories_products WHERE categories_products.idCategory = " & intIdCategory
                                
                                %>
                                <option value="<%=Server.HtmlEncode(pcStrCatLink)%>"><%=pcStrCategoryDesc%></option>
                            <% Next
                        end if 
                        %>
                    </select>
                  </form>
                </div>
			<% end if 
			' SECOND & THIRD subcategory display options
			if pIntSubCategoryView<>"2" then
				if pFeaturedCategory<>0 then
					'// Call SEO Routine
					pcGenerateSeoLinks
					'//
					%>
					<!--#include file="pcShowCategoryFeatured.asp" -->
				<% end if %>
                
                <div class="pcCategoriesWrapper">
            
                    <% if pIdCategory=1 then %>
                        <% 
                        iCurOGNum=0
                        pcv_mc=0 
                        Do While pcv_mc < iCategoriesPageSize And pcv_mc<intSubCatCount+1
                            intIdCategory=SubCatArray(0,pcv_mc)
                            strCategoryDesc=SubCatArray(1,pcv_mc)
                            pcStrCategoryDesc=SubCatArray(1,pcv_mc)
                            ' SECOND display option: rich display
                            ' Thumbnail only view
                            pcCategoryClasses =  pcCategoryClass & " " & pcCategoryHover
                            %>
                            <div class="pccol-fluid pccol-fluid-<%= pIntCategoryColumns %>">
                                <div class="<%= pcCategoryClasses %>">
                                    <%
                                    if pIntSubCategoryView=4 then %>
                                            <!--#include file="pcShowCategoryT.asp" -->
                                    <%
                                    elseif pIntSubCategoryView="0" then
                                            if pIntCategoryColumns > 1 then %>
                                                <!--#include file="pcShowCategoryH.asp" -->
                                            <% else %>
                                                <!--#include file="pcShowCategoryP.asp" -->
                                            <% end if
                                    else
                                            '// Show categories as text links only
                                            '// Call SEO Routine
                                            pcGenerateSeoLinks
                                            '//
                                            %>
                                            <a href="<%=Server.HtmlEncode(pcStrCatLink)%>"><%=pcStrCategoryDesc%></a>
                                            <!--#include file="pcShowCatDiscIcon.asp" -->
                                    <% end if %>
                                </div>
                            </div>
                            <%
                            iCurOGNum = iCurOGNum + 1
                            pcv_mc=pcv_mc+1
                            
                            If ( iCurOGNum = pIntCategoryColumns ) Then
                                iCurOGNum = 0

                            End If
                        Loop '// Do While pcv_mc < iCategoriesPageSize And pcv_mc<intSubCatCount+1
                        %>
    
                    <% else %>
                    
                        <% pcv_mc=0 
                        Do While pcv_mc < iCategoriesPageSize And pcv_mc<intSubCatCount+1
                            intIdCategory=SubCatArray(0,pcv_mc)
                            strCategoryDesc=SubCatArray(1,pcv_mc)
                            pcStrCategoryDesc=SubCatArray(1,pcv_mc)
                            ' SECOND display option: rich display
                            ' Thumbnail only view
                            
                            pcCategoryClasses = pcCategoryClass & " " & pcCategoryHover
                            %>
                            <div class="pccol-fluid-<%= pIntCategoryColumns %>">
                                <div class="<%= pcCategoryClasses %>">
                                    <%
                                    if pIntSubCategoryView=4 then %>
                                        <!--#include file="pcShowCategoryT.asp" -->
                                    <%
                                    elseif pIntSubCategoryView="0" then
                                        if pIntCategoryColumns > 1 then %>
                                                <!--#include file="pcShowCategoryH.asp" -->
                                        <% else %>
                                                <!--#include file="pcShowCategoryP.asp" -->
                                        <% end if
                                    else 
                                        '// Call SEO Routine
                                        pcGenerateSeoLinks
                                        '//
                                        %>
                                        <a href="<%=Server.HtmlEncode(pcStrCatLink)%>"><%=strCategoryDesc%></a>
                                        <!--#include file="pcShowCatDiscIcon.asp" -->
                                    <% end if %>
                                </div>
                            </div>
                            <%
                            iCurOGNum = iCurOGNum + 1
                            pcv_mc=pcv_mc+1
                            If ( iCurOGNum = pIntCategoryColumns ) Then
                                iCurOGNum = 0

                            End If
                        Loop %>
                        
                    <% end if %>
                    <div class="pcClear"></div>
          	    </div>
			
				<% call PageCategoriesNav(iCategoriesPageCount) %>
			<% end if %>

		<% End If
		'*******************************
		' END show subcategories
		'*******************************
		
		'*******************************
		' START show products
		'*******************************
	
		'Query order	
		Dim UONum, pcIntProductOrder
		query="SELECT POrder FROM categories_products WHERE idCategory="& pIdCategory &";"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		UONum=0
		do while not rs.eof
			pcIntProductOrder=rs("POrder")
			if not validNum(pcIntProductOrder) then pcIntProductOrder=0
			if pcIntProductOrder>0 then
				UONum=UONum+CLng(pcIntProductOrder)
			end if
			rs.MoveNext
		loop
		SET rs=nothing
		
		'Decide Order By
		Dim ProdSort 
		ProdSort=trim(getUserInput(request("prodsort"),2))
		if NOT validNum(ProdSort) then
			ProdSort=""
		end if
		if ProdSort="" then
			if UONum>0 then
				ProdSort="19"
			elseif pStrProductOrder <> "" then
        ProdSort=CInt(pStrProductOrder)
      else
				ProdSort=PCOrd
			end if
		end if
	
		select case ProdSort
			Case "19": query1 = " ORDER BY categories_products.POrder Asc"
			Case "0": query1 = " ORDER BY products.SKU Asc"
			Case "1": query1 = " ORDER BY products.description Asc" 	
			Case "2": 
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
					else
						query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) DESC"
					end if
				else
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
					else
						query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) DESC"
					end if
				End if
			Case "3":
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
					else
						query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) ASC"
					end if
				else
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
					else
						query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) ASC"
					end if
				End if	
		end select
		
		'////////////////////////////////////////////////////////////////
		'// START: Category Seach Fields 
		'////////////////////////////////////////////////////////////////
        If SRCH_CSFON = "1" Then 
			pcv_strCSFieldQuery = Session("pcv_strCSFieldQuery")
            pcv_strCSFilters = Session("pcv_strCSFilters")
		End If
		'////////////////////////////////////////////////////////////////
		'// END: Category Seach Fields
		'////////////////////////////////////////////////////////////////
		
		%>
    	<!--#include file="pcShowProducts.asp" -->
    <%
		
		'// Query Products of current category
		query="SELECT products.idProduct, products.sku, products.description, products.price, products.listhidden, products.listprice, products.serviceSpec, products.bToBPrice, products.smallImageUrl,products.noprices,products.stock, products.noStock,products.pcprod_HideBTOPrice,products.pcProd_BackOrder,products.FormQuantity,products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& mIdCategory&" AND active=-1 AND configOnly=0 and removed=0 " & pcv_strCSFilters & query1
		set rs=Server.CreateObject("ADODB.Recordset")   
		rs.CacheSize=iPageSize
		rs.PageSize=iPageSize
		pcv_strPageSize=iPageSize
			
		rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		dim iPageCount, pcv_intProductCount
		iPageCount=rs.PageCount
		If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
		If Cint(iPageCurrent) < 1 Then iPageCurrent=1
		
		if NOT rs.eof then
			rs.AbsolutePage=Cint(iPageCurrent)
			pcArray_Products = rs.getRows(iPageSize)
			pcv_intProductCount = UBound(pcArray_Products,2)+1
		end if
	
		set rs = nothing
	
		if pcv_intProductCount<1 then ' START IF-1: check if there are no products in this category...
			if intSubCatExist <> 1 then ' ... and there are no sub-categories, then show a message  %>
				<p><%=dictLanguage.Item(Session("language")&"_viewCat_P_4")%></p>
			<% end if
		
		else ' ELSE IF-1: there are products or sub-categories
			
			if intSubCatExist = 1 then
				' If there are products AND subcategories, then products are considered
				' "featured" products within the category and are shown above the subcategories
				%>
				<hr />
				<% if pIdCategory<>1 then %>
					<h3><%=dictLanguage.Item(Session("language")&"_viewCategories_1")%>&quot;<%=pCategoryName%>&quot;</h3>
				<% else %>
					<h3><%=dictLanguage.Item(Session("language")&"_viewCategories_1b")%></h3>
				<% end if %>
			<%	end if ' The category contains products, but not subcategories %>
			<%
        	Dim pcv_strFacetContent, pcv_boolIsFacetContent 
			pcv_boolIsFacetContent = False

			If scSearch_IsEnabled = "1" Then
        		pcv_strFacetContent = pcs_SolrCatalog(pIdCategory)        		
        		If len(pcv_strFacetContent)>0 Then
            		pcv_boolIsFacetContent = True
        		Else
            		pcv_boolIsFacetContent = False
        		End If
			End If

			'// If SORT BY drop-down does not exist, show page nav still
        	If pcv_boolIsFacetContent = False Then
            	call PageNav(iPagecount, "Top")
        	End If
        	%>
            
      <%
      '=================================
      'show SORT BY drop-down
      '=================================
      If HideSortPro<>"1" And pcv_boolIsFacetContent = False Then %>
        <div class="pcSortProducts">
          <form action="<%= Server.HtmlEncode("viewCategories.asp?pageStyle=" & pcPageStyle & "&idcategory=" & pidcategory & pcv_strCSFieldQuery)%>" method="post" class="pcForms">
            <%=dictLanguage.Item(Session("language")&"_viewCatOrder_5")%>
            <select id="pcSortBox" class="form-control" name="prodSort" onChange="javascript:if (this.value != '') {this.form.submit();}">
            <%if UONum>0 then%>          
                <option value="19" <%if ProdSort="19" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_6")%></option>
            <%end if%>                        
                <option value="0"<%if ProdSort="0" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_1")%></option>
                <option value="1"<%if ProdSort="1" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_2")%></option>
                <option value="2"<%if ProdSort="2" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_3")%></option>
                <option value="3"<%if ProdSort="3" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_4")%></option>
            </select>
          </form>
        </div>
      <% End If 
      '=================================
      'end SORT BY drop-down
      '=================================
      %>    
      
		<div class="pcClear"></div>
        <% If pcv_boolIsFacetContent = True Then %>
        
            <script>
                var category = <%=pIdCategory %>;
            </script>
            <div data-ng-controller="solrSearchCtrl">
                <htmldiv content="myhtml">
                    <%= pcs_SolrCatalog(pIdCategory) %>
                </htmldiv>
            </div>
            
        <% Else %>

			<%
			call pcShowProducts(iPageSize, 0)
			%>

			<!--#include file="atc_viewprd.asp"-->
    
			<% call PageNav(iPagecount, "Bottom") %>
        <% End If %>
            
        <% End If %>
  
        <%=(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_viewCategories_22"), "viewAll", 200)) %>

    </div>
</div>
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Category Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CategoryPromotionMsg
Dim rs,query,tmpStr	
	query="SELECT pcCatPro_PromoMsg FROM pcCatPromotions WHERE idcategory=" & pIdCategory & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpStr=rs("pcCatPro_PromoMsg")
		set rs=nothing
		' Display long product description if there is a short description
		if tmpStr <> "" then %>
      <div class="pcPromoMessage">
      	<%=tmpStr%>
      </div>
		<% end if
	end if
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Category Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>

<% 
'====================
' Page Navigation
'==================== 

Sub PageNav(thepagecount, topBottom)
	'// SEO-S
	intIdCategory=mIdCategory
	pcStrCategoryDesc=pMainCategoryName
	'// Call SEO Routine
	pcGenerateSeoLinks
	'// SEO-E
	iRecSize=10
	If thepagecount>1 then %>
		<div id="pcPagination<%= topBottom %>" class="pcPagination">
		<span>
			<%=(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & thepagecount)%>
			<% if thepagecount>iRecSize then %>
                <% if cint(iPageCurrent)>iRecSize then %>
                	<%
										url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&page=1&idCategory=" & mIdCategory & pcv_strCSFieldQuery
									%>
                    <a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_15")%></a>&nbsp;
                <% end if %>
                <% if cint(iPageCurrent)>1 then
                    if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
                        iPagePrev=cint(iPageCurrent)-1
                    else
                        iPagePrev=iRecSize
                    end if %>
                    <%
											url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&page=" & cint(iPageCurrent)-iPagePrev & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
										%>
                    <a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17a")%><%=iPagePrev%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
                <% end if %>
                <% 
                if cint(iPageCurrent)+1>1 then
                    intPageNumber=cint(iPageCurrent)
                else
                    intPageNumber=1
                end if
            else
                intPageNumber=1
            end if
            
            if (cint(thepagecount)-cint(iPageCurrent))<iRecSize then
                iPageNext=cint(thepagecount)-cint(iPageCurrent)
            else
                iPageNext=iRecSize
            end if
	
			%>
      	&nbsp;-&nbsp;
      <%
			For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
				If Cint(pageNumber)=Cint(iPageCurrent) Then %>
					<b><%=pageNumber%></b> 
				<% Else %>
        	<%
						url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&page=" & pageNumber & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
					%>
					<a href="<%= Server.HtmlEncode(url) %>"><%=pageNumber%></a>
				<% End If 
			Next
		
			if (cint(iPageNext)+cint(iPageCurrent))=thepagecount then
			else
				if thepagecount>(cint(iPageCurrent) + (iRecSize-1)) then %>
        	<%
						url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&page=" & cint(intPageNumber)+iPageNext & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
					%>
					&nbsp;<a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17")%><%=iPageNext%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
				<% end if
			
				if cint(thepagecount)>iRecSize AND (cint(iPageCurrent)<>cint(thepagecount)) then %>
        	<%
						url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&page=" & cint(thepagecount) & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
					%>
					<a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_16")%></a>
				<% end if 
			end if %>
      	<%
					url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&page=" & cint(thepagecount) & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery & "&viewAll=yes"
				%>
            &nbsp;<a href="<%= Server.HtmlEncode(url) %>" onClick="pcf_Open_viewAll();"><%=dictLanguage.Item(Session("language")&"_viewCategories_21")%></a>
      </span>
    </div>
	<% end if
end Sub

Sub PageCategoriesNav(thepagecount)
	'// SEO-S
	intIdCategory=mIdCategory
	pcStrCategoryDesc=pMainCategoryName
	'// Call SEO Routine
	pcGenerateSeoLinks
	'// SEO-E

	iRecSize=10
	If thepagecount>1 then %>
		<div id="pcPagination" class="pcPagination">
			<span>
			
			<%=(dictLanguage.Item(Session("language")&"_advSrcb_4") & iCategoriesPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & thepagecount)%>
        <% if thepagecount>iRecSize then %>
			<% if cint(iCategoriesPageCurrent)>iRecSize then %>
						<%
              url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&page=1&idCategory=" & mIdCategory & pcv_strCSFieldQuery
            %>
            <a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_15")%></a>&nbsp;
         <% end if %>
            <% if cint(iCategoriesPageCurrent)>1 then
                if cint(iCategoriesPageCurrent)<iRecSize AND cint(iCategoriesPageCurrent)<iRecSize then
                    iPagePrev=cint(iCategoriesPageCurrent)-1
                else
                    iPagePrev=iRecSize
                end if %>
                	<%
										url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&CategoriesPage=" & cint(iCategoriesPageCurrent)-iPagePrev & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
									%>
                	&nbsp;<a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17a")%><%=iPagePrev%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
            <% end if
			if cint(iCategoriesPageCurrent)+1>1 then
				intPageNumber=cint(iCategoriesPageCurrent)
			else
				intPageNumber=1
			end if
		else
			intPageNumber=1
		end if
		if (cint(thepagecount)-cint(iCategoriesPageCurrent))<iRecSize then
			iPageNext=cint(thepagecount)-cint(iCategoriesPageCurrent)
		else
			iPageNext=iRecSize
		end if
	
		%>
			&nbsp;-&nbsp;
		<%
			
		For pageNumber=intPageNumber To (cint(iCategoriesPageCurrent) + (iPageNext))
			If Cint(pageNumber)=Cint(iCategoriesPageCurrent) Then %>
				<b><%=pageNumber%></b> 
			<% Else %>
      	<%
					url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&CategoriesPage=" & pageNumber & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
				%>
				<a href="<%= Server.HtmlEncode(url) %>"><%=pageNumber%></a>
			<% End If 
		Next
		
		if (cint(iPageNext)+cint(iCategoriesPageCurrent))=thepagecount then
		else
			if thepagecount>(cint(iCategoriesPageCurrent) + (iRecSize-1)) then %>
      	<%
					url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&CategoriesPage=" & cint(intPageNumber)+iPageNext & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
				%>
				<a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17")%><%=iPageNext%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
			<% end if
			
			url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&CategoriesPage=" & cint(thepagecount) & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery
				
			if cint(thepagecount)>iRecSize AND (cint(iCategoriesPageCurrent)<>cint(thepagecount)) then %>
				<a href="<%= Server.HtmlEncode(url) %>"><%=dictLanguage.Item(Session("language")&"_viewCategories_16")%></a>
			<% end if 
		end if %>
    <%
			url = pcStrCatLink2 & "?pageStyle=" & pcPageStyle & "&ProdSort=" & ProdSort & "&idCategory=" & mIdCategory & pcv_strCSFieldQuery & "&viewAll=yes"
		%>
        &nbsp;<a href="<%= Server.HtmlEncode(url) %>" onClick="pcf_Open_viewAll();"><%=dictLanguage.Item(Session("language")&"_viewCategories_21")%></a>
			</span>
  	</div>
	<% end if
end Sub

'====================
' END Page Navigation
'==================== 
%>

<!--#include file="footer_wrapper.asp"-->
<!--#include file="bulkAddToCart.asp"-->
