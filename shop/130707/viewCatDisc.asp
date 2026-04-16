<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% pageTitle="Quantity Discounts by Category" %>
<% section="specials" %>
<!--#include file="AdminHeader.asp"-->
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Categories"
				src_FormTitle2="Quantity Discounts by Category - Search Results"
				src_FormTips1="Use the following filters to look for categories in your store. <b>Important Note</b>: Categories that contain <b>both</b> Products <b>and</b> Sub Categories are not shown in the Search Results and are not eligible for Category-based Quantity Discounts. I.E. Only Categories that do not have Sub Categories can have Quantity Discounts."
				src_FormTips2="Quantity discount by categories allow you to apply discounts when multiple products from the same categories are purchased (e.g. 5 music CD's, regardless of which CD's are purchased). Only categories to which products have been assigned are displayed on this page. Once you have assigned discounts to a category, you can apply the same discounts to multiple other categories at once. Just click on the &quot;Modify&quot; icon, then select &quot;Apply to Other Categories&quot;."
				src_DisplayType=0
				src_ShowLinks=0
				src_ParentOnly=2
				src_FromPage="viewCatDisc.asp"
				src_ToPage=""
				src_Button1=" Search "
				src_Button2=""
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcCat_DiscArea")="1"
				src_ShowDiscTypes="1"
				session("srcCat_from")=""
				session("srcCat_where")=""
			%>
				<!--#include file="inc_srcCATs.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->
