<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<% pageTitle = "Change category settings across multiple configurable products" %>
<% section = "services" %>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f, pidProduct
%>
<!--#include file="AdminHeader.asp"-->
<% If request("catlist")&""="" Then
	call closeDb()
response.redirect "ApplyBTOCatMulti1.asp"
End If %>
<% pIdcategory=request("catlist")
pcArr=split(pIdcategory,",")
pIdcategory=pcArr(0)
session("cp_bto_ar2_idcategory")=pIdcategory
%>
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
				src_FormTitle1="Find Configurable Products"
				src_FormTitle2="Change category settings across multiple configurable products"
				src_FormTips1="Use the following filters to look for configurable products in your store for which you would like to change category settings."
				src_FormTips2="Select one or more configurable products for which you would like to change category settings."
				src_IncNormal=0
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ApplyBTOCatMulti2.asp?action=add&catlist=" & request("catlist")
				src_ToPage="ApplyBTOCatMulti3.asp?action=add"
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=0
				session("srcprd_from")=""
				session("srcprd_where")=""
				%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="Adminfooter.asp"-->
