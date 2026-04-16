<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="services" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
dim lngBTOProduct,lngIDCategory,intCatType

lngBTOProduct=request("BtOProduct")
lngIDCategory=request("IDCategory")
intCatType=request("CatType")

Select Case intCatType
	Case "1": TempStr="configSpec_products"
	Case "2": TempStr="configSpec_Charges"
End Select

if validNum(lngBTOProduct) and validNum(lngIDCategory) then
	query="delete from " & TempStr & " where specProduct=" & lngBTOProduct & " and configProductCategory=" & lngIDCategory
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
end if	



call closeDb()
response.redirect "modBTOconfiga.asp?idProduct="& lngBTOProduct & "&s=1&msg=" & Server.URLEncode("The selected category was successfully removed from the product configuration.")

%>
