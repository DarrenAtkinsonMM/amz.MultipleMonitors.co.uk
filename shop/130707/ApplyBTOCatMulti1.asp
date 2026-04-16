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

session("cp_bto_ar2_idcategory")=""
session("cp_bto_ar2_btolist")="" 
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
					src_FormTitle1="Find a category"
					src_FormTips1="Use the following filters to look for a category in your store whose settings you want to update."
					src_FormTitle2="Change configurable category settings across multiple configurable products"
					src_FormTips2="Select the category whose settings you want to update across multiple configurable products."
					src_DisplayType=2
					src_ShowLinks=0
					src_FromPage="ApplyBTOCatMulti1.asp"
					src_ToPage="ApplyBTOCatMulti2.asp?action=add"
					src_Button1=" Search "
					src_Button2="Continue"
					src_Button3="Back"
					src_ParentOnly=0
				%>
				<!--#include file="inc_srcCATs.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="Adminfooter.asp"-->
