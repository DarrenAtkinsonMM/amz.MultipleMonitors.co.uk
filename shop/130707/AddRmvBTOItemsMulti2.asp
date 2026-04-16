<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% if request("action")<>"add" then
	call closeDb()
response.redirect "AddRmvBTOItemsMulti1.asp"
	response.End()
end if %>

<% pageTitle = "Assign/Remove Configurable-Only Items To/From Multiple Products" %>
<% section = "services" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f, pidProduct 
%>
<!--#include file="AdminHeader.asp"-->
<%
pIdcategory=request("catlist")
pcArr=split(pIdcategory,",")
pIdcategory=pcArr(0)
session("cp_bto_ar1_idcategory")=pIdcategory
	
query="SELECT categoryDesc FROM categories WHERE idcategory=" & pIdcategory
set rs=connTemp.execute(query)
if not rs.eof then
	pcv_CatDesc=rs("categoryDesc")
end if
set rs=nothing
cat_HadItem=pcv_CatDesc & " <input type=hidden name=idcategory value=""" & pIdcategory & """>"
%>
<table class="pcCPcontent">
    <tr>
        <td>
            <table id="FindProducts" class="pcCPcontent">
                <tr>
                    <td>
						<%
                        src_FormTitle1="STEP 2 - Pick items within the category"
                        src_FormTitle2="Assign/Remove Configurable-Only Items To/From Multiple Products"
                        src_FormTips1="Use the following filters to look for the items within this category that you would like to assign/remove to/from multiple configurable products."
                        src_FormTips2="Select one or more configurable-only items that you would like to assign/remove to/from multiple configurable products."
                        src_IncNormal=1
                        src_IncBTO=0
                        src_IncItem=1
                        src_DisplayType=1
                        src_ShowLinks=0
                        src_FromPage="AddRmvBTOItemsMulti2.asp?action=add&catlist=" & request("catlist")
                        src_ToPage="AddRmvBTOItemsMulti3.asp?action=add"
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
