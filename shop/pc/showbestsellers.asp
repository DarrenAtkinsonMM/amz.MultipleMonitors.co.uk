<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%pcStrPageName="showbestsellers.asp"%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="prv_incFunctions.asp"-->

<%
Dim iAddDefaultPrice, iAddDefaultWPrice
%>
<% 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' End: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim rsProducts, rsDisc, pDiscountPerQuantity, pcStrPageName
%>
<!--#include file="prv_getSettings.asp"-->
<%

'*******************************
' GET PAGE SETTINGS FROM DB
'*******************************
Dim pcIntBestSellCount, pcStrBestSellDesc, pcIntBestSellNFS, queryNFS, pcIntBestSellInStock, queryInStock, pcIntBestSellSales, pShowSKU, pShowSmallImg, pcPageStyle, pagesize

pcIntBestSellSales=0
pcIntBestSellCount=0

query="SELECT pcBSS_BestSellCount,pcBSS_Style,pcBSS_PageDesc,pcBSS_NSold,pcBSS_NotForSale,pcBSS_OutOfStock,pcBSS_SKU,pcBSS_ShowImg FROM pcBestSellerSettings;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)
If not rs.eof Then
	pcIntBestSellCount=rs("pcBSS_BestSellCount")
	pcPageStyle=rs("pcBSS_Style")
	pcStrBestSellDesc=rs("pcBSS_PageDesc")
	pcIntBestSellSales=rs("pcBSS_NSold")
	pcIntBestSellNFS=rs("pcBSS_NotForSale")
	pcIntBestSellInStock=rs("pcBSS_OutOfStock")
	pShowSKU=rs("pcBSS_SKU")
	pShowSmallImg=rs("pcBSS_ShowImg")
End If
set rs=nothing

If isNULL(pcIntBestSellCount) or (pcIntBestSellCount="0") Then
	pcIntBestSellCount= 14
End If
pagesize = pcIntBestSellCount

If isNULL(pcIntBestSellSales) or (pcIntBestSellSales="0") Then
	pcIntBestSellSales=2
End If

If pcIntBestSellNFS<> 0 and NotForSaleOverride(session("customerCategory"))=0 Then
	queryNFS = " AND ((products.formQuantity)=0)"
Else
	queryNFS = " "
End If

If isNULL(pShowSKU) OR (pShowSKU="") Then
	pShowSKU=0
End If

If isNULL(pShowSmallImg) OR (pShowSmallImg="") Then
	pShowSmallImg=0
End If

If pcPageStyle = "" Then
	pcPageStyle = LCase(Request.QueryString("pageStyle"))
	If pcPageStyle = "" Then
		pcPageStyle = LCase(Request.Form("pageStyle"))
	End If
End If

If pcPageStyle = "" Then
	pcPageStyle = LCase(bType)
End If
		
If pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" Then
	pcPageStyle = LCase(bType)
End If

'FB-S
pIntProductColumns=scPrdRow
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
' GET Best Sellers from DB
'*******************************
If session("CustomerType")<>"1" Then
	query1= " AND ((categories.pccats_RetailHide)=0)"
Else
	query1=""
End If

If pcIntBestSellInStock<> 0 and scOutOfStockPurchase<>0 Then
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice,products.pcProd_BackOrder,products.formQuantity,  products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.stock)>0 OR (products.pcProd_Apparel=1)) AND ((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") OR (((products.noStock)=-1 OR (products.pcProd_Apparel=1)) AND ((products.sales)>"&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
Else
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice,products.pcProd_BackOrder,products.formQuantity,products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
End If

set rsProducts=server.CreateObject("ADODB.Recordset")
set rsProducts=conntemp.execute(query)
If err.number<>0 Then
	call LogErrorToDatabase()
	set rsProducts=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
End If
If NOT rsProducts.eof Then
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
Else
	set rsProducts = nothing
	call closeDb()
	response.redirect "msg.asp?message=94"
End If
set rsProducts = nothing

'*******************************
' Build the page
'*******************************
%>
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<!--#include file="inc_AddThis.asp"-->

<div id="pcMain" class="pcShowBestSellers">
  <div class="pcMainContent">
    <%
      '// PC v4.5 AddThis integration
      If scAddThisDisplay=1 Then pcs_AddThis
    %>
    <h1><%=dictLanguage.Item(Session("language")&"_viewBestSellers_2")%></h1>
  
    <%
    '// Show New Best Sellers description, If any
    If pcf_HasHTMLContent(pcStrBestSellDesc) Then
    %>
      <div class="pcPageDesc"><%= pcf_FixHTMLContentPaths(pcStrBestSellDesc) %></div>
    <%
    End If
    %>
    
    <%
    call pcShowProducts(pageSize, 0)
    %>

    <!--#include file="atc_viewprd.asp"-->
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->