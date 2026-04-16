<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="prv_incFunctions.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<% 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim rsProducts, rsDisc, pDiscountPerQuantity, pcStrPageName
pcStrPageName = "shownewarrivals.asp"

%>
<!--#include file="prv_getSettings.asp"-->
<%

'*******************************
' GET PAGE SETTINGS FROM DB
'*******************************
Dim pcIntNewArrCount, pagesize, pcStrNewArrDesc, pcIntNewArrNFS, queryNFS, pcIntNewArrInStock, queryInStock, pcNDays, pShowSKU, pShowSmallImg, pcPageStyle

query="SELECT pcNAS_NewArrCount, pcNAS_Style, pcNAS_PageDesc, pcNAS_NDays, pcNAS_NotForSale, pcNAS_OutOfStock, pcNAS_SKU, pcNAS_ShowImg FROM pcNewArrivalsSettings;"
set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntNewArrCount=rs("pcNAS_NewArrCount")
	pcPageStyle=rs("pcNAS_Style")
	pcStrNewArrDesc=rs("pcNAS_PageDesc")
	pcNDays=rs("pcNAS_NDays")
	pcIntNewArrNFS=rs("pcNAS_NotForSale")
	pcIntNewArrInStock=rs("pcNAS_OutOfStock")
	pShowSKU=rs("pcNAS_SKU")
	pShowSmallImg=rs("pcNAS_ShowImg")
end if

set rs=nothing

if isNULL(pcIntNewArrCount) or (pcIntNewArrCount="0") then
	pcIntNewArrCount = 14
end if
pagesize = pcIntNewArrCount

if len(pcNDays)<1 then
	pcNDays=15
end if

if pcIntNewArrNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 then
	queryNFS = "((products.formQuantity)=0) AND"
else
	queryNFS = " "
end if

if isNULL(pShowSKU) OR (pShowSKU="") then
	pShowSKU=0
end if

if isNULL(pShowSmallImg) OR (pShowSmallImg="") then
	pShowSmallImg=0
end if

if pcPageStyle = "" then
	pcPageStyle = LCase(Request.QueryString("pageStyle"))
	if pcPageStyle = "" then
		pcPageStyle = LCase(Request.Form("pageStyle"))
	end if
end if

if pcPageStyle = "" then
	pcPageStyle = LCase(bType)
end if
		
if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

'FB-S
pIntProductColumns=scPrdRow
pIntProductRows=scPrdRowsPerPage
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
	pIntSubCategoryView=0
	pIntCategoryColumns=1
	pIntCategoryRows=10
	pIntProductColumns=1
	pIntProductRows=10
	pcPageStyle = "h"
End If
%>
	<!--#include file="pcShowProducts.asp" -->
<%

'*******************************
' GET new arrivals from DB
'*******************************
if session("CustomerType")<>"1" then
	query1= " AND ((categories.pccats_RetailHide)=0)"
else
	query1=""
end if

pcTodayDate=Date()
if SQL_Format="1" then
	pcTodayDate=Day(pcTodayDate)&"/"&Month(pcTodayDate)&"/"&Year(pcTodayDate)
else
	pcTodayDate=Month(pcTodayDate)&"/"&Day(pcTodayDate)&"/"&Year(pcTodayDate)
end if

y="'"

if pcIntNewArrInStock <> 0 and scOutOfStockPurchase<>0 then
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.pcProd_BackOrder, products.formQuantity,  products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.stock)>0) AND ((products.formQuantity)=0) AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-convert(datetime, [products].[pcprod_EnteredOn],101))<="& pcNDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) OR (((products.noStock)=-1) AND ((products.formQuantity)=0) AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-convert(datetime, [products].[pcprod_EnteredOn],101))<="& pcNDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) ORDER BY products.pcprod_EnteredOn DESC;"
else
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.pcProd_BackOrder, products.formQuantity,  products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE ("&queryNFS&" ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-convert(datetime, [products].[pcprod_EnteredOn],101))<="& pcNDays &") AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.pcprod_EnteredOn DESC;"
end if

set rsProducts=server.CreateObject("ADODB.Recordset")
set rsProducts=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsProducts=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if NOT rsProducts.eof then
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
else
	set rsProducts = nothing
	call closeDb()
	response.redirect "msg.asp?message=93"
end if
set rsProducts = nothing


'*******************************
' Build the page
'*******************************
%>
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<!--#include file="inc_AddThis.asp"-->

<div id="pcMain" class="pcShowNewArrivals">
  <div class="pcMainContent">
		<%
      '// PC v4.5 AddThis integration
      if scAddThisDisplay=1 then pcs_AddThis
    %>
    <h1><%= dictLanguage.Item(Session("language")&"_viewNewArrivals_2")%></h1>
  
    <% ' Show New Arrival description, if any
      If pcf_HasHTMLContent(pcStrNewArrDesc) Then
      %>
        <div class="pcPageDesc"><%= pcf_FixHTMLContentPaths(pcStrNewArrDesc) %></div>
      <%
      End If
    %>
      
    <%
    call pcShowProducts(pagesize, 0)
    %>
    <!--#include file="atc_viewprd.asp"-->
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->
