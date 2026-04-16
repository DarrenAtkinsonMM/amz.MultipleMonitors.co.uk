<%@  language="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "pcSyndication_GetItems.asp"
' This page outputs a JSON representation of the shopping cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp" -->
<!--#include file="../includes/dateinc.asp" -->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<!--#include file="inc_sb.asp"-->
<% 
response.Clear()
Response.ContentType = "application/json"
Response.Charset = "UTF-8"
%>
<%  
dim jsonService : set jsonService = JSON.parse("{}")

'// Category ID
pIdCategory=SNW_CATEGORY
if pIdCategory="" OR isNULL(pIdCategory) then
	pIdCategory=0
end if

'// Affiliate ID
idaffiliate=Request("idaffiliate")

'// Sort
if ProdSort="" then
	ProdSort="19"
end if

select case ProdSort
	Case "19": query1 = " ORDER BY categories_products.POrder Asc, products.description Asc"
	Case "0": query1 = " ORDER BY products.SKU Asc"
	Case "1": query1 = " ORDER BY products.description Asc" 	
	Case "2": 
		If Session("customerType")=1 then
			query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
		else
			query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
		End if
	Case "3":
		If Session("customerType")=1 then
			query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
		else
			query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
		End if
end select

'// Query Products
query="SELECT TOP " & SNW_MAX & " products.idProduct, products.price, products.smallImageUrl, products.description FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& pIdCategory &" AND active=-1 AND configOnly=0 and removed=0 " & query1
set rs=server.CreateObject("ADODB.RecordSet")

set rs=conntemp.execute(query)		
if NOT rs.EOF then
	pcArray_Products = rs.getRows()
	set rs=nothing
	pcv_intProductCount = UBound(pcArray_Products,2)
	
    redim syndicationitemrows(pcv_intProductCount)
	
	if scSSL = "1" then
		widgetURL = scSslURL &"/"& scPcFolder & "/pc/"
	else
		widgetURL = scStoreURL &"/"& scPcFolder & "/pc/"
	end if
	
	pcv_URL=replace(widgetURL, "//", "/")
	pcv_URL=replace(pcv_URL,"http:/","http://")
	pcv_URL=replace(pcv_URL,"https:/","https://")
    
	For pCnt=0 to pcv_intProductCount
	
		pidProduct=""
		pDescription=""   
		pPrice=""
		pSmallImageUrl="" 
	
		pidProduct=pcArray_Products(0,pCnt) '// rs("idProduct")
		pPrice=pcArray_Products(1,pCnt) '// rs("price")
		pSmallImageUrl=pcArray_Products(2,pCnt) '// rs("smallImageUrl")
		pDescription=pcArray_Products(3,pCnt) '// rs("description")
		
		if pSmallImageUrl="" OR isNULL(pSmallImageUrl) then
			pSmallImageUrl="no_image.gif"
		end if
	
		pDescription=ClearHTMLTags2(pDescription,2)
		pDescription=replace(pDescription,"&quot;","""")
		If 44<len(pDescription) then
			pDescription=trim(left(pDescription,44)) & "..."
		End If

        Dim syndicationitemrow : set syndicationitemrow = JSON.parse("{}")
        
        syndicationitemrow.set "description", pDescription
        syndicationitemrow.set "price", money(pPrice)
        syndicationitemrow.set "image", pcv_URL & "catalog/" & pSmallImageUrl
        syndicationitemrow.set "url", pcv_URL & "viewPrd.asp?idproduct=" & pidProduct

        set syndicationitemrows(pCnt) = syndicationitemrow
        
	Next
    
    jsonService.set "syndicationitemrow", syndicationitemrows
    jsonService.set "totalItems", pcv_intProductCount + 1

end If	
set rs=nothing

response.Clear()
Response.write( JSON.stringify(jsonService, null, 2) & vbNewline )
set Info = nothing

call closeDb()
response.End()
%>
