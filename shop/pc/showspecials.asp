<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%pcStrPageName="showspecials.asp"%>
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
' Display Settings
'*******************************

' This variable allows the admin to show some text above Specials -> move to Control Panel
Dim pcStrSpecialsDesc
pcStrSpecialsDesc = dictLanguage.Item(Session("language")&"_viewSpc_5")

' START - Not For Sale visibility
' This variable controls whether NOT FOR SALE items should be shown
' PC v4.1: copy from Best Sellers
	Dim pcIntSpecialsNFS, queryNFS
	query="SELECT pcBSS_NotForSale FROM pcBestSellerSettings;"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	If Not rs.eof Then
		pcIntSpecialsNFS=rs("pcBSS_NotForSale")
	End If
	Set rs=Nothing

	' Or you can override the value manually by uncommenting one of the lines below
	'	pcIntSpecialsNFS = 0 ' Not for sale items are shown
	'	pcIntSpecialsNFS = -1 ' Not for sale items are not shown
		
	If pcIntSpecialsNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 Then
		queryNFS = " AND formQuantity = 0 "
		Else
		queryNFS = " "
	End If
'// End - Not For Sale visibility

' These variables show/hide information when products are shown with Page Style = L or M
Dim pShowSKU, pShowSmallImg
pShowSKU = scShowSKU ' If 0, Then the SKU is hidden
pShowSmallImg = scShowSmallImg ' If 0, Then the small image is not shown
	' Note: the size of the small image is Set via the css/pcStorefront.css stylesheet

'*******************************
' GET page style
'*******************************
	' Load the page style: check to see If a querystring
	' or a form is sEnding the page style.
	Dim pcPageStyle
	pcPageStyle = LCase(Request.QueryString("pageStyle"))
		If pcPageStyle = "" Then
			pcPageStyle = LCase(Request.Form("pageStyle"))
		End If

		If pcPageStyle = "" Then
			pcPageStyle = LCase(bType)
		End If

		If pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" Then
			pcPageStyle = LCase(bType)
		End If
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
	If Not isNumeric(pcv_ViewAllVar) OR pcv_ViewAllVar="" Then
		pcv_ViewAllVar=0
	End If
	

	Dim iPageSize
	'FB-S
	iPageSize=(pIntProductColumns*pIntProductRows)
	'FB-E
		If request.queryString("iPageCurrent")="" Then
			If request.queryString("page")="" Then
				iPageCurrent=1
			Else
				iPageCurrent=server.HTMLEncode(request.queryString("page"))
			End If
		Else
			iPageCurrent=server.HTMLEncode(request.queryString("iPageCurrent"))
	End If

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
		Case "0": querySort = " ORDER BY products.SKU Asc"
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
	
%>
	<!--#include file="pcShowProducts.asp" -->
<%

'*******************************
' GET Specials from DB
'*******************************

If session("CustomerType")<>"1" Then
	query1= " AND categories.pccats_RetailHide=0"
Else
	query1=""
End If

query="SELECT distinct products.idProduct,products.sku,products.description,products.price,products.listHidden,products.listPrice,products.serviceSpec,products.bToBPrice,products.smallImageUrl,products.noprices,products.stock,products.noStock,products.pcprod_HideBTOPrice,products.pcProd_BackOrder,products.formQuantity,products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM products,categories_products,categories WHERE products.active=-1 AND products.hotdeal=-1 AND products.configOnly=0 AND products.removed=0 " & queryNFS & " AND categories_products.idProduct=products.idProduct AND categories.idCategory=categories_products.idCategory AND categories.iBTOhide=0 " & query1 & querySort
Set rsProducts=Server.CreateObject("ADODB.Recordset")     
rsProducts.CursorLocation=adUseClient
rsProducts.CacheSize=iPageSize
rsProducts.Open query, conntemp
	
If err.number<>0 Then
	call LogErrorToDatabase()
	Set rsProducts=Nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
End If
dim iPageCount, count
If Not rsProducts.eof Then	
	
	rsProducts.MoveFirst
	rsProducts.PageSize=iPageSize
	pcv_strPageSize=iPageSize
	iPageCount=rsProducts.PageCount

	rsProducts.AbsolutePage=Cint(iPageCurrent)
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
	HaveProducts=1
Else
	Set rsProducts = Nothing
	HaveSales=0
	HaveProducts=0
	If UCase(scDB)="SQL" Then
	query="SELECT pcSC_ID,pcSC_SaveName,pcSC_SaveDesc FROM pcSales_Completed WHERE pcSC_Status=2 ORDER BY pcSC_SaveName ASC;"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	If Not rs.eof Then
	HaveSales=1
	End If
	Set rs=Nothing
	End If
	If HaveSales=0 Then
	call closeDb()
  	response.redirect "msg.asp?message=89"
	End If       
End If
Set rsProducts = Nothing

'*******************************
' Build the page
'*******************************
%>
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<!--#include file="inc_AddThis.asp"-->

<div id="pcMain" class="pcShowSpecials">
  <div class="pcMainContent">
		<%
      '// PC v4.5 AddThis integration
      If scAddThisDisplay=1 Then pcs_AddThis
    %>
		<h1><%= dictLanguage.Item(Session("language")&"_viewSpc_2")%></h1>
    
    <% ' Show Specials description, If any
      If pcf_HasHTMLContent(pcStrSpecialsDesc) Then
      %>
        <div class="pcPageDesc"><%=pcf_FixHTMLContentPaths(pcStrSpecialsDesc)%></div>
      <%
      End If
      %>
      
    	
		<%if HideSortPro<>"1" then%>
			<% if pcv_ViewAllVar=0 then %>
				<% pcPageNavTopBottom = "Top" %>
				<!--#include file="pcPageNavigation.asp"-->
			<% end if %>
      
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
    If UCase(scDB)="SQL" Then
      tmpTargetType=0
      If session("customerCategory")<>"" AND session("customerCategory")<>"0" Then
        tmpTargetType=session("customerCategory")
      Else
        If session("customerType")="1" Then
          tmpTargetType="-1"
        End If
      End If
      query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveDesc,pcSales_Completed.pcSC_SaveIcon FROM pcSales_Completed INNER JOIN pcSales ON pcSales_Completed.pcSales_ID=pcSales.pcSales_ID WHERE pcSales_Completed.pcSC_Status=2 AND pcSales.pcSales_TargetPrice=" & tmpTargetType & " ORDER BY pcSC_SaveName ASC;"
      Set rs=Server.CreateObject("ADODB.Recordset")
      Set rs=connTemp.execute(query)
      If Not rs.eof Then
        saleArr=rs.getRows()
        intSale=ubound(saleArr,2)%>
        <div class="pcSectionTitle">
          <%= dictLanguage.Item(Session("language")&"_SaleSpecials_1")%>
        </div>
        <%For k=0 to intSale%>
          <div class="pcSaleDesc">
            <div class="pcSaleDescTitle">
              <% If trim(saleArr(3,k))<>"" Then%>
                <img src="<%=pcf_getImagePath("../pc/catalog",saleArr(3,k))%>" alt="SALE">
              <% End If %>
              <a href="showsearchresults.asp?incSale=1&IDSale=<%=saleArr(0,k)%>"><%=saleArr(1,k)%></a>
            </div>
            <div class="pcSaleDescContent"><%=saleArr(2,k)%></div>
          </div>
        <%Next%>
        <%If HaveProducts=1 Then%>
        <div class="pcSectionTitle">
          <%= dictLanguage.Item(Session("language")&"_SaleSpecials_2")%>
        </div>
        <%End If%>
      <%End If
      Set rs=Nothing
    End If%>
    
    <%
      If HaveProducts = "1" Then
	  
		If pcv_ViewAllVar=0 Then
			call pcShowProducts(iPageSize, 0)
		Else
			call pcShowProducts(pcv_intProductCount, 0)
		End if

        'Insert page navigation
        If pcv_ViewAllVar=0 Then %>
					<% pcPageNavTopBottom = "Bottom" %>
					<!--#include file="pcPageNavigation.asp"-->
		   <% End If
        
      End If
    %>
  <!--#include file="atc_viewprd.asp"-->
  </div>
</div>

<!--#include file="footer_wrapper.asp"-->
