<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"--> 
<!--#include file="../includes/CashbackConstants.asp"--> 
<!--#include file="HomeCode.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<% 
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "home.asp"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'-------------------------------
' declare local variables
'-------------------------------

dim pcStrHPStyle, pcStrHPDesc, pcIntHPFirst, pcIntHPShowSKU, pcIntHPShowImg, pcIntHPFeaturedCount, pcIntHPFeaturedOrder
dim pcIntHPSpcCount, pcIntHPSpcOrder, pcIntHPNewCount, pcIntHPSNewOrder, pcIntHPBestCount, pcIntHPBestOrder
Dim rsProducts, rsDisc, pDiscountPerQuantity, pTotalCount
%>
<!--#include file="prv_getSettings.asp"-->
<%

'*******************************
' LOAD HOMEPAGE SETTINGS
'*******************************
' Refer to "pcadmin/manageHomePage.asp" to see features added to this page

query = "SELECT pcHPS_FeaturedCount, pcHPS_Style, pcHPS_PageDesc, pcHPS_First, pcHPS_ShowSKU, pcHPS_ShowImg," &_
        "pcHPS_SpcCount, pcHPS_SpcOrder, pcHPS_NewCount, pcHPS_NewOrder, pcHPS_BestCount, pcHPS_BestOrder FROM pcHomePageSettings"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntHPFeaturedCount=rs("pcHPS_FeaturedCount")
	pcStrHPStyle=rs("pcHPS_Style")	
	pcStrHPDesc=replace(rs("pcHPS_PageDesc"),"''","'")
	pcIntHPFirst=rs("pcHPS_First")
	pcIntHPShowSKU=rs("pcHPS_ShowSKU")
	pcIntHPShowImg=rs("pcHPS_ShowImg")
	pcIntHPSpcCount=rs("pcHPS_SpcCount")
	pcIntHPSpcOrder=rs("pcHPS_SpcOrder")
	pcIntHPNewCount=rs("pcHPS_NewCount")
	pcIntHPNewOrder=rs("pcHPS_NewOrder")
	pcIntHPBestCount=rs("pcHPS_BestCount")
	pcIntHPBestOrder=rs("pcHPS_BestOrder")
end if

if pcIntHPFeaturedCount = "" or not validNum(pcIntHPFeaturedCount) or pcIntHPFeaturedCount < 0 then
	pcIntHPFeaturedCount = 3
end if

' // Note: 0 is an acceptable value and it indicates that Specials should not be shown
if pcIntHPSpcCount = "" or not validNum(pcIntHPSpcCount) or pcIntHPSpcCount < 0 then
	pcIntHPSpcCount = 4
end if

' // Note: 0 is an acceptable value and it indicates that New Arrivals should not be shown
if pcIntHPNewCount = "" or not validNum(pcIntHPNewCount) or pcIntHPNewCount < 0 then
	pcIntHPNewCount = 4
end if

' // Note: 0 is an acceptable value and it indicates that Best Sellers should not be shown
if pcIntHPBestCount = "" or not validNum(pcIntHPBestCount) or pcIntHPBestCount < 0 then
	pcIntHPBestCount = 4
end if

set rs=nothing

pShowSKU = pcIntHPShowSKU
if pShowSKU = "" or isNull(pShowSKU) then
	pShowSKU = -1 ' If 0, then the SKU is hidden
end if

pShowSmallImg = pcIntHPShowImg
if pShowSmallImg = "" or isNull(pShowSmallImg) then
	pShowSmallImg = -1 ' If 0, then the Image is hidden
end if

'*******************************
' END LOAD HOMEPAGE SETTINGS
'*******************************
%>

<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->			
<!--#include file="pcValidateQty.asp"-->

<!--#include file="../includes/javascripts/pcValidateFormViewPrd.asp"-->

<%

'*******************************
' GET page style
'*******************************
' Load the page style: check to see if a querystring
' or a form is sending the page style.
Dim pcPageStyle

pcPageStyle = LCase(Request.QueryString("pageStyle"))
if pcPageStyle = "" then
	pcPageStyle = LCase(Request.Form("pageStyle"))
end if

if pcPageStyle = "" then
    pcPageStyle = pcStrHPStyle
end if

if pcPageStyle = "" then
	pcPageStyle = LCase(bType)
end if

if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

Dim pcArray_Products, pcv_intProductCount, pIntProductColumns
%>
<!--#include file="pcShowProducts.asp" -->
<%
'*******************************
' GET Featured Products from DB
'*******************************
if session("CustomerType")<>"1" then
	query1= " AND ((categories.pccats_RetailHide)=0)"
else
	query1=""
end if

'// START v4.1 - Not For Sale override
	if NotForSaleOverride(session("customerCategory"))=1 then
		queryNFSO=""
	else
		queryNFSO="AND formQuantity = 0 "
	end if
'// END v4.1
query = "SELECT products.idProduct,products.sku,products.description,products.price,products.listHidden,products.listPrice,products.serviceSpec,products.bToBPrice,products.smallImageUrl,products.noprices,products.stock,products.noStock, "
query = query & "products.pcprod_HideBTOPrice,products.pcProd_BackOrder,products.formQuantity,products.pcProd_BTODefaultPrice, products.sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, "
query = query & "products.hotdeal, products.pcProd_SkipDetailsPage "
query = query & "FROM products "
query = query & "WHERE "
query = query & "products.idProduct IN "
query = query & "( "
query = query & "	SELECT categories_products.idProduct FROM categories_products INNER JOIN categories ON categories.idCategory = categories_products.idCategory "
query = query & "	WHERE categories.iBTOhide=0 " & query1 & " "
query = query & ") "
query = query & "AND products.active=-1 AND products.showInHome=-1 AND products.configOnly=0 AND products.removed=0 " & queryNFSO & " "
query = query & "order by pcprod_OrdInHome ASC "

set rsProducts=server.CreateObject("ADODB.Recordset")
set rsProducts=conntemp.execute(query)
if Err.number <> 0 then
    call LogErrorToDatabase()
    set rsProducts = Nothing
    call closeDb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
end If
if NOT rsProducts.eof then
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
end if
set rsProducts = nothing

'*******************************
' Set Total Count
'*******************************
pTotalCount=pcv_intProductCount
pIntProductColumns=scPrdRow

'*******************************
' Build the page
'*******************************
%>

<div id="pcMain" class="pcHome">
	<div class="pcMainContent">
		<%
			If Not Session("HideSlideShow")="1" Then
				If Session("Mobile")="1" Then
					pcs_ShowSlideShowMobile
				Else
					pcs_ShowSlideShow
				End If
			Else
				Session("HideSlideShow")="0"
			End If
		%>
        
    <%
      'If there are no featured products and no page description, hide the table row
      if pcIntHPFeaturedCount > 0 or pcStrHPDesc <> "" then
          ' If there are featured products, show that message, otherwise hide it
          if pcIntHPFeaturedCount > 0 then %>
            <h1><%=dictLanguage.Item(Session("language")&"_mainIndex_11")%></h1>
          <% end if 
          
          ' Show Home Page description, if any
          if pcf_HasHTMLContent(pcStrHPDesc) then %>
            <div class="pcPageDesc"><%=pcf_FixHTMLContentPaths(pcStrHPDesc)%></div>
          <% end if
      end if
  
      '*****************************************************************************************************
      ' 1) PRODUCT OF THE MONTH
      '*****************************************************************************************************
      pcs_ProductOfTheMonth
      '*****************************************************************************************************
      ' END PRODUCT OF THE MONTH
      '*****************************************************************************************************
  
      '*****************************************************************************************************
      ' 2) FEATURED PRODUCTS
      '*****************************************************************************************************
      pcs_FeaturedProducts
      '*****************************************************************************************************
      ' END FEATURED PRODUCTS
      '*****************************************************************************************************
  
      '*****************************************************************************************************
      ' 3) Best sellers, new arrivals, specials
      '*****************************************************************************************************
      pcs_ShowProducts
      '*****************************************************************************************************
      ' END Best sellers, new arrivals, specials 
      '*****************************************************************************************************
    %>
  </div>
      
  <!--#include file="atc_viewprd.asp"-->
</div>

<%	  
set rsProducts=Nothing
set iPageCurrent=Nothing
%>
<!--#include file="orderCompleteTracking.asp"-->
<!--#include file="inc-Cashback.asp"-->
<!--#include file="footer_wrapper.asp"-->