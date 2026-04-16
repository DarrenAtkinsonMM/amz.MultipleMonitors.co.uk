<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Upload Images" %>
<% Section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<%
on error resume next
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent" style="height: 400px;">
<tr>
	<td>
  	<iframe style="border: none; width: 100%; height: 480px;" src="../htmleditor/addons/assetmanager/asset.asp?ffilter=image">
    	<p>Your browser does not support iframes. Please <a href="ImageUploada.asp">use this file</a> to upload images.</p>
  	</iframe>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->
