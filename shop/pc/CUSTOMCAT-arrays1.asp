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

pTempIntSubCategory=5

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
'DA - EDIT
pIdCategory=pTempIntSubCategory
mIdCategory=pTempIntSubCategory
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

	<!-- Header: pagetitle -->
    <header id="stands-header" class="stands-header">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-6">
                         <div class="wow fadeInDown pt-headtext maxW-fix marginT-xlfix" data-wow-offset="0" data-wow-delay="0">
							<h1 class="">Monitor Arrays</h1>
							<h2 class="text-uppercase">MONITOR ARRAYS ARE <span class="color-sb">THE COMPLETE DISPLAY PACKAGE</span>, GIVING YOU AN INSTANT PRODUCTIVITY BOOST AND LOOKING AMAZING ALL IN ONE</h2>
						 </div>
						 <div class="wow fadeInUp maxW-fix" data-wow-offset="0" data-wow-delay="0">
							<p class="home-head-text text-white text-justify">A Monitor Array is made up of a multi screen stand and a set of monitors all perfectly aligned together, with our wide range of stands you can achieve the perfect system for your needs.</p>
						    <p class="home-head-text text-white text-justify">Take a look through the most popular arrays, view our special designs or scroll further down to create your ultimate monitor display setup.</p>
						</div>
                    </div>
					<div class="col-md-6">
                         <div class="wow fadeInRight text-center" data-wow-offset="0" data-wow-delay="0.1s">
							  <img src="/images/banners/arrays.png">
						 </div>
                    </div>
					
				</div>		
			</div>		
		</div>	
    </header>

	<section id="product-stands" class="product-stands bg-smog product-grid paddingtop-20 paddingbot-40">
		<div class="container">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-check-square-o green-link"></i> <span>Top Selling</span> Monitor Arrays</h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20">View the most popular Monitor Arrays or <a href="javascript:arraycustomjump();void(0);">scroll down</a> to create your own</h5>
						</div>			
			 <div class="row">
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Dual 21.5 inch Monitor Array"><a class="" href="/display-systems-3/?sid=287&mid=304">Dual 21.5" Monitor Array</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/display-systems-3/?sid=287&mid=304"><img src="/images/bundles/s2h-a22-atn.jpg" alt="Dual 21.5 Monitor Array" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Dual Monitor Array complete with two AOC 21.5" widescreen monitors.</p>
								
								<h4>Price: <span>&pound;295.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/display-systems-3/?sid=287&mid=304">View Full Array Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Triple 24inch Monitor Array"><a class="" href="/display-systems-3/?sid=312&mid=317">Triple 24" Monitor Array</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/display-systems-3/?sid=312&mid=317"><img src="/images/bundles/s3h-a24-atn.jpg" alt="Triple 24 Monitor Array" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Triple Horizontal Monitor Array using three Acer 24" widescreen monitors.</p>
								
								<h4>Price: <span>&pound;445.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/display-systems-3/?sid=312&mid=317">View Full Array Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Quad Square 24inch Monitor Array"><a class="" href="/display-systems-3/?sid=313&mid=320">Quad Square 24" Monitor Array</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/display-systems-3/?sid=313&mid=320"><img src="/images/bundles/s4s-i22-atn.jpg" alt="Quad Square 24 Monitor Array" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Quad Square Monitor Array using four Iiyama IPS 24" widescreen monitors.</p>
								
								<h4>Price: <span>&pound;735.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/display-systems-3/?sid=313&mid=320">View Full Array Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Quad Pyramid 24inch Monitor Array"><a class="" href="/display-systems-3/?sid=325&mid=317">Quad Pyramid 24" Monitor Array</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/display-systems-3/?sid=325&mid=317"><img src="/images/bundles/s4p-a24-atn.jpg" alt="Quad Pyramid 24 Monitor Array" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Quad Pyramid Monitor Array using four Acer 24" widescreen monitors.</p>
								
								<h4>Price: <span>&pound;585.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/display-systems-3/?sid=325&mid=317">View Full Array Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Five Pyramid 24inch Monitor Array"><a class="" href="/display-systems-3/?sid=318&mid=320">Five Pyramid 24" Monitor Array</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/display-systems-3/?sid=318&mid=320"><img src="/images/bundles/s5p-i22-atn.jpg" alt="Five Pyramid 24 Monitor Array" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Five Pyramid Monitor Array using five Iiyama IPS 24" widescreen monitors.</p>
								
								<h4>Price: <span>&pound;895.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/display-systems-3/?sid=318&mid=320">View Full Array Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Six 24inch Monitor Array"><a class="" href="/display-systems-3/?sid=338&mid=317">Six 24" Monitor Array</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/display-systems-3/?sid=338&mid=317"><img src="/images/bundles/s6r-a24-atn.jpg" alt="Six 24 Monitor Array" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Six Monitor Array using six Acer 24" widescreen monitors.</p>
								
								<h4>Price: <span>&pound;815.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/display-systems-3/?sid=338&mid=317">View Full Array Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
		<a name="arraycustom"></a></div>
	</section>
    <!-- /Section: Welcome -->

<!--#include file="array-breadcrumb.asp"-->
	<section id="product-stands" class="product-stands bg-smog product-grid paddingtop-70 paddingbot-40">
		<div class="container">
			<div class="row">
			<% 
	
		
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
		
		'DA -EDIT
		'// Query Products of current category
		query="SELECT products.idProduct, products.sku, products.description, products.price, products.listhidden, products.listprice, products.serviceSpec, products.bToBPrice, products.smallImageUrl,products.noprices,products.stock, products.noStock,products.pcprod_HideBTOPrice,products.pcProd_BackOrder,products.FormQuantity,products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage, products.pcUrl FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& mIdCategory&" AND active=-1 AND configOnly=0 and removed=0 " & pcv_strCSFilters & query1
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

        	%>
            
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
    
        <% End If %>
            
        <% End If %>
  
		</div>
	</section>
    <!-- /Section: Welcome -->

<!--#include file="footer_wrapper.asp"-->
<!--#include file="bulkAddToCart.asp"-->
