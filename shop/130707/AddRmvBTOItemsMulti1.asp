<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Assign/Remove Configurable-Only Items To/From Multiple Products" %>
<% section = "services" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f, pidProduct

session("cp_bto_ar1_idcategory")=""
session("cp_bto_ar1_itemlist")=""
session("cp_bto_ar1_btolist")="" 
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
					src_FormTitle1="STEP 1 - Choose a category"
					src_FormTips1="Use the following filters to look for a category that contains configurable-only items. You will then be able to select which of the items included in the category should be added/removed to/from the configurable products that you will select later."
					src_FormTitle2="Assign/Remove Configurable-Only Items To/From Multiple Products"
					src_FormTips2="Choose a category."
					src_DisplayType=2
					src_ShowLinks=0
					src_FromPage="AddRmvBTOItemsMulti1.asp"
					src_ToPage="AddRmvBTOItemsMulti2.asp?action=add"
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
