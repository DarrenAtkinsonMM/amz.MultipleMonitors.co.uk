<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="prv_incfunctions.asp"-->
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

Dim rsProducts, rsDisc, pDiscountPerQuantity, pcStrPageNameOR
pcStrPageNameOR="showRecentlyReviewed.asp"

%>
<!--#include file="prv_getsettings.asp"-->
<%

'*******************************
' GET PAGE SETTINGS FROM DB
'*******************************

Dim pcIntRecentRevCount, pagesize, pcStrRecentRevDesc, pcIntRecentRevNFS, queryNFS, pcIntRecentRevInStock, queryInStock, pShowSKU, pShowSmallImg, pcPageStyle, pcintReviewsPerProduct, pcPageStyleOR


pcIntRecentRevCount=0


query="SELECT pcRR_RecentRevCount,pcRR_Style,pcRR_PageDesc,pcRR_RevDays,pcRR_NotForSale,pcRR_OutOfStock,pcRR_SKU,pcRR_ShowImg, pcRR_ReviewsPerProduct FROM pcRecentRevSettings;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)
if not rs.eof then
	pcIntRecentRevCount=rs("pcRR_RecentRevCount")
	pcPageStyle=rs("pcRR_Style")
	pcStrRecentRevDesc=rs("pcRR_PageDesc")
	pcIntRevDays=rs("pcRR_RevDays")
	pcIntRecentRevNFS=rs("pcRR_NotForSale")
	pcIntRecentRevInStock=rs("pcRR_OutOfStock")
	pShowSKU=rs("pcRR_SKU")
	pShowSmallImg=rs("pcRR_ShowImg")
	' PRV41 begin
	pcintReviewsPerProduct = rs("pcRR_ReviewsPerProduct")
	' PRV41 end
end if
set rs=nothing

if isNULL(pcIntRecentRevCount) or (pcIntRecentRevCount="0") then
	pcIntRecentRevCount= 14
end If

' PRV41 begin
if isNULL(pcintReviewsPerProduct) or (pcintReviewsPerProduct="0") then
	pcintReviewsPerProduct= 3
end if
' PRV41 end
pagesize = pcIntRecentRevCount

if isNULL(pcIntRevDays) or (pcIntRevDays="0") or (pcIntRevDays="") then
	pcIntRevDays=30
end if

if pcIntRecentRevNFS<> 0 and NotForSaleOverride(session("customerCategory"))=0 then
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

pcPageStyleOR = trim(LCase(Request.QueryString("pageStyle")))
if pcPageStyleOR <> "" then pcPageStyle = pcPageStyleOR

if pcPageStyle = "" then
	pcPageStyle = LCase(bType)
end if
		
if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end If

'// Forse page style (could be reviewed in the future)
pcPageStyle = "l"

%>
	<!-- #include file="pcShowProducts.asp" -->
<%

'*******************************
' GET Best Sellers from DB
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

if pcIntRecentRevInStock<> 0 and scOutOfStockPurchase<>0 then
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.pcProd_BackOrder, products.formQuantity,  products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM ((products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory) INNER JOIN pcReviews ON products.idProduct = pcReviews.pcRev_IDProduct WHERE (((products.stock)>0) AND ((products.formQuantity)=0) AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-[products].[pcprod_EnteredOn])<="& pcIntRevDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) OR (((products.noStock)=-1) AND ((products.formQuantity)=0) AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-[products].[pcprod_EnteredOn])<="& pcIntRevDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) AND ((pcReviews.pcRev_Active = 1)) ORDER BY products.pcprod_EnteredOn DESC;"
else
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.pcProd_BackOrder, products.formQuantity,  products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM ((products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory) INNER JOIN pcReviews ON products.idProduct = pcReviews.pcRev_IDProduct WHERE ("&queryNFS&" ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-[pcReviews].[pcRev_Date])<="& pcIntRevDays &") AND ((categories.iBTOhide)=0)"&query1&") AND ((pcReviews.pcRev_Active = 1)) ORDER BY products.pcprod_EnteredOn DESC;"
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
	
	'// START Troubleshooting: print number of products with recent reviews
	'response.write "pcv_intProductCount=" & pcv_intProductCount
	'response.End()
	'// END
	
else
	set rsProducts = nothing
	call closeDb()
	response.redirect "msg.asp?message=301"
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

<div id="pcMain">
  <div class="pcMainContent">
		<%
      '// PC v4.5 AddThis integration
      if scAddThisDisplay=1 then pcs_AddThis
    %>
    <h1><%=dictLanguage.Item(Session("language")&"_ShowRecentRev_1")%></h1>
        
		<%
      ' Show Recently Reviewed Products description, if any
      If pcf_HasHTMLContent(pcStrRecentRevDesc) Then
      %>
        <div class="pcPageDesc"><%= pcf_FixHTMLContentPaths(pcStrRecentRevDesc) %></div>
      <%
      End If
    %>
    
    <%
    call pcShowProducts(pagesize, 0)
    %>
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->
