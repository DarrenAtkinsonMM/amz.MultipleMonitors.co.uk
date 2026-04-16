<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% pageTitle="Remove Product from Control Panel" %>
<%
' form parameters
pIdProduct=request.Querystring("idProduct")

if not ValidNum(pIdProduct) then
  	call closeDb()
	response.redirect "msg.asp?message=2"
end if



' delete from taxPrd
query="DELETE FROM taxPrd WHERE idProduct=" &pidproduct
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
  call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from configSpec_products
query="DELETE FROM configSpec_products WHERE configProduct=" &pIdProduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
  call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from cs_relationships
query="DELETE FROM cs_relationships WHERE idProduct=" &pIdProduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
  call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from categories_products
query="DELETE FROM categories_products WHERE idProduct=" &pIdProduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
  call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from products table
query="UPDATE products SET active=0, removed=-1 WHERE idproduct=" &pidproduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
  call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

set rs=nothing

call pcs_hookProductRemoved(pidproduct, "")

If statusAPP="1" OR scAPP=1 Then

	query="UPDATE products SET active=0, removed=-1 WHERE pcprod_ParentPrd=" &pidproduct
	set rs=conntemp.execute(query)	
	set rs=nothing

End If


call closeDb()
response.redirect "srcPrds.asp"
%>