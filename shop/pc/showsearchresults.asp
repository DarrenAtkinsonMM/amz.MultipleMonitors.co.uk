<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<%
Dim pcv_strHideCatSearch
pcv_strHideCatSearch = False '// Set to "True" to disable category search

Dim pcv_strHideSubSearch
If statusAPP="1" Then
	pcv_strHideSubSearch = False '// Set to "True" to disable sub-product search
Else
	pcv_strHideSubSearch = True
End If

'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "showSearchResults.asp"

'*******************************
' Query results
'*******************************
dim pSearchSKU, pKeywords, pPriceFrom, pPriceUntil, pIdCategory, pIdSupplier, pWithStock, strSearch, pCValues, tKeywords, tIncludeSKU, IDBrand, strPrdOrd, strOrderBy

pSearchSKU=getUserInput(request.querystring("SKU"),150)
pKeywords=getUserInput(request.querystring("keyWord"),100)
pCValues=getUserInput(request.querystring("SearchValues"),0)
tKeywords=pKeywords
tIncludeSKU=getUserInput(request.querystring("includeSKU"),10)
	if tIncludeSKU = "" then
		tIncludeSKU = "true"
	end if
pPriceFrom=getUserInput(request.querystring("priceFrom"),20)
if Instr(pPriceFrom,",")>Instr(pPriceFrom,".") then
	pPriceFrom=replace(pPriceFrom,",",".")
end if
if NOT isNumeric(pPriceFrom) then
	pPriceFrom=0
end if
pPriceUntil=getUserInput(request.querystring("priceUntil"),20)
if Instr(pPriceUntil,",")>Instr(pPriceUntil,".") then
	pPriceUntil=replace(pPriceUntil,",",".")
end if
if NOT isNumeric(pPriceUntil) then
	pPriceUntil=999999999
end if
pIdCategory=getUserInput(request.querystring("idCategory"),4)
pIdSupplier=getUserInput(request.querystring("idSupplier"),4)
if NOT validNum(pIdSupplier) or trim(pIdSupplier)="" then
	pIdSupplier=0
end if
pWithStock=getUserInput(request.querystring("withStock"),2)	
IDBrand=getUserInput(request.querystring("IDBrand"),20)
if NOT validNum(IDBrand) or trim(IDBrand)="" then
	IDBrand=0
end if
incSale=getUserInput(request("incSale"),4)
if NOT validNum(incSale) or trim(incSale)="" then
	incSale=0
end if
tmpIDSale=getUserInput(request("IDSale"),4)
if NOT validNum(tmpIDSale) or trim(tmpIDSale)="" then
	tmpIDSale=0
end if
strPrdOrd=getUserInput(request.querystring("order"),4)
	if NOT validNum(strPrdOrd) or trim(strPrdOrd)="" then strPrdOrd=PCOrd
	if NOT validNum(strPrdOrd) or trim(strPrdOrd)="" then strPrdOrd=1
	Select Case strPrdOrd
		Case "0": strOrderBy="A.sku ASC, A.idproduct DESC"
		Case "1": strOrderBy="A.description ASC"
		Case "2":
			If Session("customerType")=1 then
				strOrderBy = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) DESC"
			else
				strOrderBy = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) DESC"
			End if
		Case "3":
			If Session("customerType")=1 then
				strOrderBy = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) ASC"
			else
				strOrderBy = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) ASC"
			End if
		Case Else: strOrderBy="A.description ASC"
	End Select
	strORD=strPrdOrd

intExact=getUserInput(request.querystring("exact"),4)
if NOT validNum(intExact) or trim(intExact)="" then
	intExact=0
end if

'*******************************
' START - Don't allow empty searches
'*******************************
Dim pcIntNullSearch
if pKeywords="" AND pIdCategory="" then
	pcIntNullSearch=1
end if
if NOT validNum(pIdCategory) or trim(pIdCategory)="" then
	pIdCategory=0
end if

if pIdCategory="0" AND (pIdSupplier="" OR pIdSupplier="0") AND pPriceFrom="0" AND pPriceUntil="999999999" AND pSearchSKU="" AND IDBrand="0" AND pKeywords="" AND (pCValues="" OR pCValues="0" OR pCValues="||") AND trim(pWithStock)="" then
	pcIntNullSearch=1
end if

'// Let price-based searches go through
if (pPriceFrom<>"0" OR pPriceUntil<>"999999999") then
	pcIntNullSearch=0
end if

'// Let brand-based searches go through
if IDBrand<>"0" then
	pcIntNullSearch=0
end if

if incSale<>"0" then
	pcIntNullSearch=0
end if

'// Let custom search field queries go through
if (pCValues<>"0" AND pCValues<>"" AND pCValues<>"||") then
	pcIntNullSearch=0
end if

if pcIntNullSearch=1 then
	response.redirect "search.asp"
end if

'*******************************
' END - Don't allow empty searches
'*******************************

%>
<!--#include file="pcStartSession.asp"-->
<%

%>
<!--#include file="prv_getSettings.asp"-->
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
			pcPageStyle = LCase(bType)
		end if

		if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
			pcPageStyle = LCase(bType)
		end if


'//===========================
'// BRAND Information - Start
'//===========================

if IDBrand>0 then

	query="SELECT BrandName, pcBrands_Description, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg FROM Brands WHERE pcBrands_Active=1 AND idBrand="&IDBrand
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		set rstemp=nothing
		call closeDb()
		response.redirect "msg.asp?message=85"       
	End if
	pcvBrandName=pcf_PrintCharacters(rstemp("BrandName"))
	pcvBrandsDescription=pcf_PrintCharacters(rstemp("pcBrands_Description"))
	pcvIntBrandsParent=rstemp("pcBrands_Parent")
	pcvBrandLogoLg=rstemp("pcBrands_BrandLogoLg")

	pcv_DefaultTitle=rstemp("pcBrands_MetaTitle")
	if isNull(pcv_DefaultTitle) or trim(pcv_DefaultTitle)="" then
		pcv_DefaultTitle=ClearHTMLTags2(pcvBrandName,0)
	end if
	pcv_DefaultTitle = pcv_DefaultTitle & " - " & scCompanyName
	pcv_DefaultDescription=rstemp("pcBrands_MetaDesc")
	pcv_DefaultKeywords=rstemp("pcBrands_MetaKeywords")
	
	set rstemp=nothing

	if not validNum(parentIntBrandsParent) then parentIntBrandsParent=0
	
end if

'//===========================
'// BRAND Information - Start
'//===========================

		
' OTHER display settings
' These variables show/hide information when products are shown with Page Style = L or M
Dim pShowSKU, pShowSmallImg
pShowSKU = scShowSKU ' If 0, then the SKU is hidden
pShowSmallImg = scShowSmallImg ' If 0, then the small image is not shown
' Note: the size of the small image is set via the css/pcStorefront.css stylesheet

%>
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<%

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
' Set page size and get current page
'*******************************
	Dim pcv_ViewAllVar
	pcv_ViewAllVar=getUserInput(request("VA"),1)
	if NOT isNumeric(pcv_ViewAllVar) OR pcv_ViewAllVar="" then
		pcv_ViewAllVar=0
	end if
	
	dim iPageSize
	iPageSize=getUserInput(request("resultCnt"),10)
	if iPageSize="" then
		iPageSize=getUserInput(request("iPageSize"),0)
	end if
	if NOT validNum(iPageSize) then
	'FB-S
	iPageSize=(pIntProductColumns*pIntProductRows)
	'FB-E
	end if
	
	dim iPageCurrent
	if request.queryString("iPageCurrent")="" then
		iPageCurrent=1 
	else
		iPageCurrent=server.HTMLEncode(request.querystring("iPageCurrent"))
		if NOT validNum(iPageCurrent) then
			iPageCurrent=1
		end if
	end if

'*******************************
' Create Search Query
'*******************************
Dim strSQL, tmpSQL, tmpSQL2, tmp_StrQuery, pcv_strMaxResults

tmp_StrQuery=""
if session("customerCategory")="" or session("customerCategory")=0 then
	If session("customerType")=1 then
		tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultWPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultWPrice<=" &pPriceUntil&")"
	else
		tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultPrice<=" &pPriceUntil&")"
	end if
else
	tmp_StrQuery="(A.serviceSpec<>0 AND A.idproduct IN (SELECT DISTINCT idproduct FROM pcBTODefaultPriceCats WHERE pcBTODefaultPriceCats.idCustomerCategory=" & session("customerCategory") & " AND pcBTODefaultPriceCats.pcBDPC_Price>="&pPriceFrom&" AND pcBTODefaultPriceCats.pcBDPC_Price<=" &pPriceUntil&"))"
end if

pcv_strMaxResults=SRCH_MAX
If pcv_strMaxResults>"0" Then
	pcv_strLimitPhrase="TOP " & pcv_strMaxResults
Else
	pcv_strLimitPhrase=""
End If

strSQLP= "SELECT "& pcv_strLimitPhrase &" idProduct, sku, description, price, listHidden, listPrice, serviceSpec, bToBPrice, smallImageUrl, noprices, stock, noStock, pcprod_HideBTOPrice, pcProd_BackOrder, FormQuantity, pcProd_BTODefaultPrice, sDesc, RowNum, totalRecords, pcprod_OrdInHome, sales, pcprod_EnteredOn, hotdeal, pcProd_SkipDetailsPage FROM (" 

strSQL= "SELECT "& pcv_strLimitPhrase &" A.idProduct, A.sku, A.description, A.price, A.listHidden, A.listPrice, A.serviceSpec, A.bToBPrice, A.smallImageUrl, A.noprices, A.stock, A.noStock, A.pcprod_HideBTOPrice, A.pcProd_BackOrder, A.FormQuantity, A.pcProd_BTODefaultPrice, cast(A.sDesc as varchar(8000)) sDesc " 

strSQLP= strSQLP &strSQL &", ROW_NUMBER() OVER (ORDER BY "&strOrderBy&") AS RowNum, COUNT(idProduct) OVER() AS totalRecords, A.pcprod_OrdInHome, A.sales, A.pcprod_EnteredOn, A.hotdeal, A.pcProd_SkipDetailsPage"

strSQL=strSQL& "FROM products A "
strSQL=strSQL& " WHERE A.idProduct IN (" 

	'// START: Category Sub-Query
	strSQL=strSQL& "SELECT B.idProduct FROM categories_products B INNER JOIN categories C ON "
	strSQL=strSQL & "C.idCategory=B.idCategory WHERE C.iBTOhide=0 "
	if pIdCategory<>"0" then
		if (schideCategory = "1") OR (SRCH_SUBS = "1") then
			Dim TmpCatList
			TmpCatList=""
			call pcs_GetSubCats(pIdCategory) '// get sub cats
			TmpCatList = pIdCategory&TmpCatList
			if len(TmpCatList)>0 then
				strSQL=strSQL & " AND B.idCategory IN ("& TmpCatList &")" '// include sub cats
			else
				strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
			end if
		else
			strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
		end if
	end if
	if session("CustomerType")<>"1" then
		tmpiHide=1
		strSQL=strSQL & " AND C.pccats_RetailHide=0"
	else
		tmpiHide=0
	end if

	tmpHiddenCat=pcf_FindHiddenCatList(tmpiHide,1)
	if tmpHiddenCat<>"" then
		strSQL=strSQL & " AND (NOT C.idCategory IN (" & tmpHiddenCat & ")) "
	end if
	'// END: Category Sub-Query

strSQL=strSQL& ") AND A.active=-1 AND A.removed=0 AND (((" & tmp_StrQuery & " OR (A.serviceSpec=0 AND A.configOnly=0 AND A.price>="&pPriceFrom&" AND A.price<=" &pPriceUntil&")) " 
'((( APP OR

if UCase(scDB)="SQL" then
	if (incSale>"0") then
		if tmpIDSale="0" then
			strSQL=strSQL & " AND A.pcSC_ID>0"
		else
			strSQL=strSQL & " AND A.pcSC_ID=" & tmpIDSale
		end if
	end if
end if

if len(pSearchSKU)>0 then
	strSQL=strSQL & " AND A.sku like '%"&pSearchSKU&"%'"
end if

if pIdSupplier<>"0" then
	strSQL=strSQL & " AND A.idSupplier=" &pIdSupplier
end if

if pWithStock="-1" then
    If statusAPP="1" Then    
	    strSQLSP=strSQLSP & " AND (((A.stock>0 OR A.noStock<>0) AND (A.pcProd_Apparel=0)) "
        strSQLSP=strSQLSP & " OR (((SELECT SUM(SP.stock) FROM products SP WHERE SP.pcProd_ParentPrd=A.idProduct) > 0 OR A.noStock<>0) AND (A.pcProd_Apparel=1))) "   
    Else    
        strSQLSP=strSQLSP & " AND (A.stock>0 OR A.noStock<>0) "        
    End IF    
    strSQL = strSQL & strSQLSP    
end if

if (IDBrand&""<>"") and (IDBrand&""<>"0") then
	strSQL=strSQL & " AND A.IDBrand=" & IDBrand
end if
pKeywords=replace(pKeywords,"''''","''")
TestWord=""
if intExact<>"1" then
	if Instr(pKeywords," AND ")>0 then
		keywordArray=split(pKeywords," AND ")
		TestWord=" AND "
	else
		if Instr(pKeywords," and ")>0 then
			keywordArray=split(pKeywords," and ")
			TestWord=" AND "
		else
			if Instr(pKeywords,",")>0 then
				keywordArray=split(pKeywords,",")
				TestWord=" OR "
			else
				if (Instr(pKeywords," OR ")>0) then
					keywordArray=split(pKeywords," OR ")
					TestWord=" OR "
				else
					if (Instr(pKeywords," or ")>0) then
						keywordArray=split(pKeywords," or ")
						TestWord=" OR "
					else
						if (Instr(pKeywords," ")>0) then
							keywordArray=split(pKeywords," ")
							TestWord=" AND "
						else
							keywordArray=split(pKeywords,"***")	
							TestWord=" OR "
						end if
					end if
				end if
			end if
		end if
	end if
else
	pKeywords=trim(pKeywords)
	if pKeywords<>"" then
		pKeywords="'" & pKeywords & "'***'%[^a-zA-z0-9]" & pKeywords & "[^a-zA-z0-9]%'***'" & pKeywords & "[^a-zA-z0-9]%'***'%[^a-zA-z0-9]" & pKeywords & "'"
	end if
	keywordArray=split(pKeywords,"***")	
	TestWord=" OR "
end if

tmpStrEx=""
if pCValues<>"" AND pCValues<>"0" then
	tmpSValues=split(pCValues,"||")
	For k=lbound(tmpSValues) to ubound(tmpSValues)
		if tmpSValues(k)<>"" then
			sfquery=""
			sfquery = "SELECT pcSearchFields_Products.idproduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData=" & tmpSValues(k)
			set rsSearchFields=Server.CreateObject("ADODB.Recordset")
			set rsSearchFields=connTemp.execute(sfquery)
			If NOT rsSearchFields.eof Then
				SearchFieldArray = pcf_ColumnToArray(rsSearchFields.getRows(),0)
				SearchFieldString = Join(SearchFieldArray,",")		
				If len(SearchFieldString)>0 Then
					tmpStrEx=tmpStrEx & " AND A.idproduct IN ("& SearchFieldString &")"
				End If
			Else
				tmpStrEx=tmpStrEx & " AND A.idproduct IN (0)"				
			End If
			set rsSearchFields = nothing
		end if
	Next
end if

'////////////////////////////////////////////////////////////////
'// START: Category Seach Fields 
'////////////////////////////////////////////////////////////////
If SRCH_CSFRON = "1" Then 

	pcv_strCSFilters=""
	pcs_CSFSetVariables()
	pcv_strCSFieldQuery = pcf_CSFieldQuery()
	if len(pcv_strCValues)>0 then
		queryCSF = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData>0 " 
		tmpSValues3=split(pcv_strCValues,"||")
		For k=lbound(tmpSValues3) to ubound(tmpSValues3)
			if tmpSValues3(k)<>"" then
				SubQuery = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData = " & tmpSValues3(k) & ""
				set rsSubQuery=Server.CreateObject("ADODB.Recordset")  
				set rsSubQuery=connTemp.execute(SubQuery)
				If NOT rsSubQuery.eof Then
					ProductIdArray = pcf_ColumnToArray(rsSubQuery.getRows(),0)
					ProductIdString = Join(ProductIdArray,",")
					tmpStrEx3=tmpStrEx3 & " AND pcSearchFields_Products.idProduct IN "
					tmpStrEx3=tmpStrEx3 & "(" & ProductIdString & ")"
				End If
				set rsSubQuery = nothing				
			end if
		Next
		queryCSF = queryCSF & tmpStrEx3	
		set rsCSF=Server.CreateObject("ADODB.Recordset")  
		set rsCSF=connTemp.execute(queryCSF)
		if NOT rsCSF.eof then
			ProductIdArray = pcf_ColumnToArray(rsCSF.getRows(),0)
			ProductIdString = Join(ProductIdArray,",")
			pcv_strCSFilters = " AND (A.idProduct In ("& ProductIdString &"))"
		else 
			pcv_strCSFilters = " AND (A.idProduct In (0))"
		end if
		set rsCSF = nothing
	end if
	err.clear
End If
tmpStrEx = tmpStrEx & pcv_strCSFilters
'////////////////////////////////////////////////////////////////
'// END: Category Seach Fields
'////////////////////////////////////////////////////////////////

IF intExact<>"1" THEN

	if pKeywords<>"" then
	
		strSQl=strSql & " AND ("
		
		tmpSQL="(A.details LIKE "
		tmpSQL2="(A.description LIKE "
		tmpSQL3="(A.sDesc LIKE "
		tmpSQL5="(A.pcProd_MetaKeywords LIKE "
		if tIncludeSKU="true" then
			tmpSQL4="(A.SKU LIKE "
		end if
		Dim Pos
		Pos=0
		For L=LBound(keywordArray) to UBound(keywordArray)
			if trim(keywordArray(L))<>"" then
			Pos=Pos+1
			if Pos>1 Then
				tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
				tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
				tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
				tmpSQL5=tmpSQL5 & TestWord & " A.pcProd_MetaKeywords LIKE "
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
				end if
			end if
				tmpSQL=tmpSQL  & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL2=tmpSQL2 & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL3=tmpSQL3 & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL5=tmpSQL5 & "'%" & trim(keywordArray(L)) & "%'"
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & "'%" & trim(keywordArray(L)) & "%'"
				end if
			end if
		Next
		tmpSQL=tmpSQL & ")"
		tmpSQL2=tmpSQL2 & ")"
		tmpSQL3=tmpSQL3 & ")"
		tmpSQL5=tmpSQL5 & ")"
		if tIncludeSKU="true" then
			tmpSQL4=tmpSQL4 & ")"
		end if
		
		strSQL=strSQL & tmpSQL
		strSQL=strSQL & " OR " & tmpSQL2
		strSQL=strSQL & " OR " & tmpSQL5
		if tIncludeSKU="true" then
			strSQL=strSQL & " OR " & tmpSQL3
			strSQL=strSQL & " OR " & tmpSQL4 & ")"
		else	
			strSQL=strSQL & " OR " & tmpSQL3 & ")"
		end if
		strSQL=strSQL& ")" & tmpStrEx
	else
		strSQL=strSQL& ")" & tmpStrEx
	end if

ELSE 'Exact=1

	if pKeywords<>"" then
	
		strSQl=strSql & " AND ("
		
		tmpSQL="(A.details LIKE "
		tmpSQL2="(A.description LIKE "
		tmpSQL3="(A.sDesc LIKE "
		tmpSQL5="(A.pcProd_MetaKeywords LIKE "
		if tIncludeSKU="true" then
			tmpSQL4="(A.SKU LIKE "
		end if
		Pos=0
		For L=LBound(keywordArray) to UBound(keywordArray)
			if trim(keywordArray(L))<>"" then
			Pos=Pos+1
			if Pos>1 Then
				tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
				tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
				tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
				tmpSQL5=tmpSQL5 & TestWord & " A.pcProd_MetaKeywords LIKE "
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
				end if
			end if
				tmpSQL=tmpSQL & trim(keywordArray(L))
				tmpSQL2=tmpSQL2 & trim(keywordArray(L))
				tmpSQL3=tmpSQL3 & trim(keywordArray(L))
				tmpSQL5=tmpSQL5 & trim(keywordArray(L))
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & trim(keywordArray(L))
				end if
			end if
		Next
		tmpSQL=tmpSQL & ")"
		tmpSQL2=tmpSQL2 & ")"
		tmpSQL3=tmpSQL3 & ")"
		tmpSQL5=tmpSQL5 & ")"
		if tIncludeSKU="true" then
			tmpSQL4=tmpSQL4 & ")"
		end if
		
		strSQL=strSQL & tmpSQL
		strSQL=strSQL & " OR " & tmpSQL2
		strSQL=strSQL & " OR " & tmpSQL5
		if tIncludeSKU="true" then
			strSQL=strSQL & " OR " & tmpSQL3
			strSQL=strSQL & " OR " & tmpSQL4 & ")"
		else	
			strSQL=strSQL & " OR " & tmpSQL3 & ")"
		end if
		strSQL=strSQL& ")" & tmpStrEx
	else
		strSQL=strSQL& ")" & tmpStrEx
	end if
	query=strSQL
END IF 'Exact


If statusAPP="1" Then
	If (pcv_strHideSubSearch=False) AND (IDBrand=0) Then
	  tmpNewQuery=""
		tmpNewQuery=strSQL
		tmpMarker = instr(strSQL,"AND A.active=-1")
		tmpNewQuery1=mid(strSQL,tmpMarker,len(strSQL))
		tmpNewQuery1=replace(tmpNewQuery1,"AND A.active=-1 AND A.removed=0 AND","")
        tmpNewQuery1=replace(tmpNewQuery1, strSQLSP, " AND (A.stock>0 OR A.noStock<>0) ")
		tmpNewQuery1=replace(tmpNewQuery1,"A.","D.")
                           
		tmpSubQuery = "SELECT D.pcProd_ParentPrd FROM products D WHERE "
		tmpSubQuery = tmpSubQuery & "D.active=0 AND D.pcProd_SPInActive=0 AND D.pcProd_ParentPrd>0 AND D.removed=0 AND "
		tmpSubQuery = tmpSubQuery & tmpNewQuery1 & ") GROUP BY D.pcProd_ParentPrd"
		'APP OR
		
		'----------------------------------------
		'Get Temp Values for Performance improvement, comment out when you don't need these codes
		query=tmpSubQuery
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		
		tmpSubQuery=""
		
		if not rstemp.eof then
			tmpArrQ=rstemp.getRows()
			set rstemp=nothing
			intC=ubound(tmpArrQ,2)
			For ic=0 to intC
				if tmpSubQuery<>"" then
					tmpSubQuery=tmpSubQuery & ","
				end if
				tmpSubQuery=tmpSubQuery & tmpArrQ(0,ic)
			Next
		end if
		set rstemp=nothing
		'----------------------------------------
	                                                                                               
		if tmpSubQuery<>"" then
			tmpSubQuery2= " OR (A.idProduct IN ( "& tmpSubQuery &" ))"
		end if
	End If
End If


tempInt = InStr(1, strSQL, "FROM products A", 1)

If tempInt > 0 Then	strSQL = Right(strSQL, Len(strSQL) - tempInt + 2)


If pcv_ViewAllVar=0 Then

	sqlBeginRecord = ((iPageCurrent-1)*iPageSize)+1
	sqlEndRecord = iPageSize*iPageCurrent

	strSQLP = strSQLP & strSQL & tmpSubQuery2 &" )) AS ProductSearch WHERE ProductSearch.RowNum BETWEEN "&sqlBeginRecord&" AND " &sqlEndRecord
Else
	strSQLP = strSQLP & strSQL &tmpSubQuery2 &" )) AS ProductSearch WHERE ProductSearch.RowNum BETWEEN 0 AND 999999"
End If
')) APP OR


totalrecords=0
session("pcstore_prdlist")=""
session("pcstore_newsrc")="OK"
Set rs=Server.CreateObject("ADODB.Recordset")

rs.Open strSQLP, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end If

dim iPageCount, count
if NOT rs.eof then		
	
	pcArray_Products = rs.getRows()

	pcv_intProductCount = pcArray_Products(18, 0)
	totalrecords = pcv_intProductCount
	
	if pcv_strMaxResults>"0" then
		if Clng(totalrecords)>Clng(pcv_strMaxResults) then
			totalrecords=pcv_strMaxResults
		end if
	end if
		
	iPageCount=Int(totalrecords/iPageSize)
	If iPageCount <> totalrecords/iPageSize Then iPageCount = iPageCount + 1
	
	If Cint(iPageCurrent) > Cint(iPageCount) Then Cint(iPageCurrent)=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
	pcv_strPageSize=iPageSize

	pcv_intProductCount = UBound(pcArray_Products, 2)+1

	'// Next and Previous Buttons
	if session("pcstore_prdlist")="" then
		session("pcstore_prdlist")="*****"
	end if
    
    If pcv_intProductCount>0 Then
        For prdList = 0 to pcv_intProductCount-1
            session("pcstore_prdlist")=session("pcstore_prdlist") & pcArray_Products(0,prdList) & "*****"
        Next
    End If

else
	set rs = nothing
	call closeDb()
	if request("fp")="bnd" then
		response.redirect "msg.asp?message=90"
	else
		response.redirect "msg.asp?message=3"                
	end if        
end if
set rs = nothing
%>

<div id="pcMain" class="pcShowSearchResults">
  <!-- Cart Container Begin -->
  <div class="pcMainContent">
		<% if IDBrand=0 then %>
	    <h1><%= dictLanguage.Item(Session("language")&"_showSearchResults_1")%></h1>
    <% else %>
      <h1><%=pcvBrandName%></h1>
      <% if pcvBrandLogoLg<>"" then %>
        <div class="pcBrandLogo"><img src="<%=pcf_getImagePath("catalog",pcvBrandLogoLg)%>" alt="<%=ClearHTMLTags2(pcvBrandName,0)%>"></div>
      <% end if %>
      <% If pcf_HasHTMLContent(pcvBrandsDescription) Then %>
        <div class="pcPageDesc"><%=pcf_FixHTMLContentPaths(pcvBrandsDescription)%></div>
      <% End If %>
  	<% end if %>    
    
		<%if pIdCategory="0" AND pcv_strHideCatSearch=False then%>
    	<!--#include file="inc_srcPrdsCAT.asp"-->
    <%end if%>
    <%strORD=strPrdOrd%>
    
    <div class="pcSpacer"></div>
          
    <div class="pcSectionTitle">
      <div class="pcColWidth50">  
        <% if pcv_strLimitPhrase="" then %>
          <%=dictLanguage.Item(Session("language")&"_advSrcb_3")%><%=totalrecords%> 
        <% else %>
          <%=dictLanguage.Item(Session("language")&"_advSrca_24")%> <%=totalrecords%> <%=dictLanguage.Item(Session("language")&"_advSrca_25")%>
        <% end if %>
        - <a href="search.asp"><%= dictLanguage.Item(Session("language")&"_ShowSearch_1")%></a>
      </div>
      
			<div class="pcColWidth50">  
				<%
          if HideSortPro<>"1" then
            tKeywords=replace(tKeywords,"''''","''") %>
            <div class="pcSortProducts">
              <form class="pcForms" name="pcSortForm">  
                <%
                  searchUrl = ""
                  searchUrl = searchUrl & "showSearchResults.asp"
                  searchUrl = searchUrl & "?VA=" & pcv_ViewAllVar 
                  searchUrl = searchUrl & "&SearchValues=" & pCValues
                  searchUrl = searchUrl & "&exact=" & intExact
                  searchUrl = searchUrl & "&iPageSize=" & iPageSize
                  searchUrl = searchUrl & "&iPageCurrent=" & iPageCurrent
                  searchUrl = searchUrl & "&pageStyle=" & pcPageStyle
                  searchUrl = searchUrl & "&keyword=" & tKeywords
                  searchUrl = searchUrl & "&priceFrom=" & pPriceFrom
                  searchUrl = searchUrl & "&priceUntil=" & pPriceUntil
                  searchUrl = searchUrl & "&idCategory=" & pIdCategory
                  searchUrl = searchUrl & "&IdSupplier=" & IdSupplier
                  searchUrl = searchUrl & "&withStock=" & pWithStock
                  searchUrl = searchUrl & "&IDBrand=" & IDBrand
                  searchUrl = searchUrl & "&SKU=" & pSearchSKU
                  searchUrl = searchUrl & pcv_strCSFieldQuery
                  searchUrl = searchUrl & "&order="
                  
                %>
              
                <%=dictLanguage.Item(Session("language")&"_advSrca_16")%>
                <select id="pcSortBox" class="form-control" name="order" onChange="javascript: if (document.pcSortForm.order.value!='') location='<%= Server.HtmlEncode(searchUrl) %>' + document.pcSortForm.order.value;">
                  <option value="0" <%if strPrdOrd="0" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_18")%></option>
                  <option value="1" <%if strPrdOrd="1" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_19")%></option>
                  <option value="3" <%if strPrdOrd="3" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_20")%></option>
                  <option value="2" <%if strPrdOrd="2" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_21")%></option>
                </select>
              </form>
            </div>
          <%
          end if 
        %>
    	</div>
      <div class="pcClear"></div>
    </div>

    <!-- Results Pagination Begin -->
    <% 
      if pcv_ViewAllVar=0 then 
        pcPageNavTopBottom = "Top"
    %>
			<div class="pcSpacer"></div>
			
			<div class="pcPageNav">
	      <!--#include file="pcPageNavigation.asp"-->
				<div class="pcClear"></div>
			</div>
    <% end if %>
    <!-- Results Pagination End -->
            
		<%
        if pCValues<>"" AND pCValues<>"0" then
            
            %>
            <div style="padding-top:2px">
				<%
                tmpSValues=split(pCValues,"||")
                For k=lbound(tmpSValues) to ubound(tmpSValues)
                    if tmpSValues(k)<>"" then
                        sfquery=""
                        sfquery = "SELECT pcSearchFields.pcSearchFieldName, pcSearchData.pcSearchDataName FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchFields_Products.idSearchData = pcSearchData.idSearchData) ON pcSearchData.idSearchField = pcSearchFields.idSearchField  WHERE pcSearchFields_Products.idSearchData=" & tmpSValues(k)
                        set rsSearchFields=Server.CreateObject("ADODB.Recordset")
                        set rsSearchFields=connTemp.execute(sfquery)
                        If NOT rsSearchFields.eof Then
                            pcv_strSearchDataName=rsSearchFields("pcSearchDataName")
                            pcv_strSearchFieldName=rsSearchFields("pcSearchFieldName")
                            pTempCValues = replace(pCValues,tmpSValues(k),"")
                            pTempCValues = replace(pTempCValues,"||||","")
                            %>
                            
														<%
                              searchUrl = pcStrPageName
                              searchUrl = searchUrl & "?ProdSort=" & ProdSort 
                              searchUrl = searchUrl & "&iPageCurrent=" & iPageCurrent
                              searchUrl = searchUrl & "&iPageSize=" & iPageSize
                              searchUrl = searchUrl & "&PageStyle=" & pcPageStyle
                              searchUrl = searchUrl & "&SearchValues=" & pTempCValues
                              searchUrl = searchUrl & "&exact=" & intExact
                              searchUrl = searchUrl & "&keyword=" & tKeywords
                              searchUrl = searchUrl & "&priceFrom=" & pPriceFrom
                              searchUrl = searchUrl & "&priceUntil=" & pPriceUntil
                              searchUrl = searchUrl & "&idCategory=" & pIdCategory
                              searchUrl = searchUrl & "&IdSupplier=" & IdSupplier
                              searchUrl = searchUrl & "&withStock=" & pWithStock
                              searchUrl = searchUrl & "&IDBrand=" & IDBrand
                              searchUrl = searchUrl & "&SKU=" & pSearchSKU
                              searchUrl = searchUrl & "&order=" & strORD
                              searchUrl = searchUrl & pcv_strCSFieldQuery
                            %>
                        
                            <%=pcv_strSearchFieldName%>: <%=pcv_strSearchDataName%> <a href="<%= Server.HtmlEncode(searchUrl) %>"><img src="<%=pcf_getImagePath("images","minus.jpg")%>" hspace="2"></a>
                            <%
                            if k<(ubound(tmpSValues)-1) then
                           		response.Write("&nbsp;|&nbsp;")
                            end if
                        End If
                        set rsSearchFields = nothing
                    end if
                Next
                %>
            </div>
					<%
        end if
        %>
        
        <%	
        totalPrds = iPageSize
        if pcv_ViewAllVar = 1 then
            totalPrds = pcv_intProductCount
        end if
					
      	call pcShowProducts(totalPrds, 0)
		%>
        
        <!-- Results Pagination Begin -->
        <% 
					if pcv_ViewAllVar=0 then
						pcPageNavTopBottom = "Bottom"
					%>
					<div class="pcPageNav">
						<!--#include file="pcPageNavigation.asp"-->
						<div class="pcClear"></div>
					</div>
        <% end if %>
        <!-- Results Pagination End -->

				<%	  
        set rs=Nothing
        set iPageCurrent=Nothing       
        %>

        <!--#include file="atc_viewprd.asp"-->
    </div>
</div>
<!--#include file="footer_wrapper.asp"-->