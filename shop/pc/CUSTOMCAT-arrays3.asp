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

pTempIntSubCategory=67

'// Validate Category ID
	if not validNum(pTempIntSubCategory) then
		pTempIntSubCategory=""
	end if
	if pTempIntSubCategory="" or pTempIntSubCategory="0" then
		pTempIntSubCategory=1
	end if
intIdCategory=pTempIntSubCategory
%>

<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->

<!--#include file="array-breadcrumb.asp"-->

<!--#include file="footer_wrapper.asp"-->
<!--#include file="bulkAddToCart.asp"-->
