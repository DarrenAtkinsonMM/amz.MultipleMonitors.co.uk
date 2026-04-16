<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwMonerisInterac.asp"
	
'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
	
IDEBIT_INVOICE = request("IDEBIT_INVOICE") 
IDEBIT_ISSLANG = request("IDEBIT_ISSLANG")
IDEBIT_ISSCONF = request("IDEBIT_ISSCONF")
IDEBIT_ISSNAME = request("IDEBIT_ISSNAME")
IDEBIT_TRACK2 = request("IDEBIT_TRACK2") 
IDEBIT_VERSION = request("IDEBIT_VERSION")
IDEBIT_VERSION = request("IDEBIT_VERSION") 

call closeDb()
Session("message") = "The INTERAC&reg; Online transaction was declined."
Session("backbuttonURL") = tempURL & "?psslurl=gwMonerisInterac.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
response.redirect "msgb.asp?back=1"
%>
