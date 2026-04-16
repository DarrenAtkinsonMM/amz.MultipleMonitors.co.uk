<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Delivery Zip Codes" %>
<% section="shipOpt" %>
<%PmAdmin="1*4*"%> 
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
if request("action")="add" then

	zipcode=request("zipcode")
	if zipcode="" then
		call closeDb()
response.redirect "DeliveryZipCodes_main.asp?r=1&msg=The Zip Code cannot be blank."
	end if

	
	query="select * from ZipCodeValidation where zipcode='" & zipcode & "'"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			
			call closeDb()
response.redirect "DeliveryZipCodes_main.asp?r=1&msg=This Zip Code is already in use."
		end if
	query="insert into ZipCodeValidation (zipcode) values ('" & zipcode & "')"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	call closeDb()
response.redirect "DeliveryZipCodes_main.asp?s=1&msg=The zip code was added successfully!"

end if

%>
