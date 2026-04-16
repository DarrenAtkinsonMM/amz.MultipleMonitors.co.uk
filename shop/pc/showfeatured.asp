<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%pcStrPageName="showfeatured.asp"%>
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
pcStrPageName = "showfeatured.asp"

%>
<!--#include file="prv_getSettings.asp"-->
<%

'*******************************
' LOAD SETTINGS (same as home page)
'*******************************

query="SELECT pcHPS_Style,pcHPS_ShowSKU,pcHPS_ShowImg FROM pcHomePageSettings;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

if not rs.eof then
	pcStrHPStyle=rs("pcHPS_Style")	
	pcIntHPShowSKU=rs("pcHPS_ShowSKU")
	pcIntHPShowImg=rs("pcHPS_ShowImg")
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

' START - Not For Sale visibility
' This variable controls whether NOT FOR SALE items should be shown
' PC v4.1: copy from Best Sellers
	Dim pcIntFeaturedNFS, queryNFS
	query="SELECT pcBSS_NotForSale FROM pcBestSellerSettings;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcIntFeaturedNFS=rs("pcBSS_NotForSale")
	end if
	set rs=nothing

	' Or you can override the value manually by uncommenting one of the lines below
	'	pcIntFeaturedNFS = 0 ' Not for sale items are shown
	'	pcIntFeaturedNFS = -1 ' Not for sale items are not shown
		
	if pcIntFeaturedNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 then
		queryNFS = " AND formQuantity = 0 "
		else
		queryNFS = " "
	end if
'// END - Not For Sale visibility

'*******************************
' END LOAD PAGE SETTINGS
'*******************************

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
'*******************************
' GET page size
'*******************************
	Dim pcv_ViewAllVar
	pcv_ViewAllVar=getUserInput(request("VA"),1)
	if NOT isNumeric(pcv_ViewAllVar) OR pcv_ViewAllVar="" then
		pcv_ViewAllVar=0
	end if
	
	
	Dim iPageSize
	'FB-S
	iPageSize=(pIntProductColumns*pIntProductRows)
	'FB-E
	if request.queryString("iPageCurrent")="" then
		if request.queryString("page")="" then
			iPageCurrent=1
		else
			iPageCurrent=server.HTMLEncode(request.queryString("page"))
		end if
	else
		iPageCurrent=server.HTMLEncode(request.queryString("iPageCurrent"))
	end if

'*******************************
' GET sorting criteria
'*******************************

 	Dim ProdSort, querySort
	ProdSort="" & request("prodsort")
 	if not validNum(ProdSort) then
		ProdSort="" & PCOrd
 	end if

 	if ProdSort="" then
		ProdSort="0"
 	end if
 	
 	select case ProdSort
		Case "0": querySort = " ORDER BY pcprod_OrdInHome asc"
		Case "1": querySort = " ORDER BY products.description Asc" 	
		Case "2": 
		If Session("customerType")=1 then
		querySort = " ORDER BY products.btoBprice desc, products.price Desc"
		else
		querySort = " ORDER BY products.price Desc"
		End if 	
		Case "3":
		If Session("customerType")=1 then
		querySort = " ORDER BY products.bToBprice Asc, products.price Asc" 	
		else
		querySort = " ORDER BY products.price Asc" 	
		end if 	
 	end select

'*******************************
' GET Featured Items from DB
'*******************************

%>
	<!--#include file="pcShowProducts.asp" -->
<%

if session("CustomerType")<>"1" then
	query1= " AND categories.pccats_RetailHide=0"
else
	query1=""
end if

query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.noprices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.pcProd_BackOrder, products.formQuantity,  products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM products, categories_products, categories WHERE products.active=-1 AND products.showInHome=-1 AND products.configOnly=0 AND products.removed=0 " & queryNFS & " AND categories_products.idProduct=products.idProduct AND categories.idCategory=categories_products.idCategory AND categories.iBTOhide=0 " & query1 & querySort
set rsProducts=Server.CreateObject("ADODB.Recordset")     
rsProducts.CursorLocation=adUseClient
rsProducts.CacheSize=iPageSize
rsProducts.Open query, conntemp
	
if err.number<>0 then
	call LogErrorToDatabase()
	set rsProducts=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
dim iPageCount, count
if NOT rsProducts.eof then	
	
	rsProducts.MoveFirst
	rsProducts.PageSize=iPageSize
	pcv_strPageSize=iPageSize
	iPageCount=rsProducts.PageCount

	rsProducts.AbsolutePage=Cint(iPageCurrent)
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1

else
	set rsProducts = nothing
	call closeDb()
  	response.redirect "msg.asp?message=89"         
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

<div id="pcMain" class="pcShowFeatured">
	<div class="pcMainContent">
		<%
      '// PC v4.5 AddThis integration
      if scAddThisDisplay=1 then pcs_AddThis
    %>
    <h1><%= dictLanguage.Item(Session("language")&"_mainIndex_7")%></h1>
    
    <% if pcv_ViewAllVar=0 then %>
			<!--#include file="pcPageNavigation.asp"-->
    <% end if %>
        
		<%if HideSortPro<>"1" then%>
			<div class="pcSortProducts">
			<form action="<%=pcStrPageName%>" method="post" class="pcForms">
			<%=dictLanguage.Item(Session("language")&"_viewCatOrder_5")%> <select id="pcSortBox" class="form-control" name="prodSort" onChange="javascript:if (this.value != '') {this.form.submit();}">
					<option value="0" <%if ProdSort="0" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_1")%></option>
					<option value="1" <%if ProdSort="1" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_2")%></option>
					<option value="2" <%if ProdSort="2" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_3")%></option>
					<option value="3" <%if ProdSort="3" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_4")%></option>
							</select>
					<input type="hidden" value="<%=pcPageStyle%>" name="PageStyle">
                      <input type="hidden" value="<%=pcv_ViewAllVar%>" name="VA">
			</form>
			</div>
		<%end if%>
    
    <%
    call pcShowProducts(iPageSize, 0)
    %>
  </div>
</div>

<!--#include file="atc_viewprd.asp"-->
<!--#include file="footer_wrapper.asp"-->
