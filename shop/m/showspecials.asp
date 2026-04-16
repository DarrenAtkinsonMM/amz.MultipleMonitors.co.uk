<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"--> 
<%
Dim originalurl, rooturl, homepageurl

pcv_URLPrefix = scStoreURL & "/" & scPcFolder
pcv_URLPrefix = replace(pcv_URLPrefix,"//","/")
pcv_URLPrefix = replace(pcv_URLPrefix,"http:/","http://")
pcv_URLPrefix = replace(pcv_URLPrefix,"https:/","https://")

rooturl = pcv_URLPrefix & "/pc/"

newurl = rooturl & "showspecials.asp"
        
call closeDb()
Response.Status = "301 Moved Permanently" 
Response.AddHeader "Location", newurl
Response.End
%>
