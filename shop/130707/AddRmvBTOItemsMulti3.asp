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
<% session("cp_bto_ar1_itemlist")=request("prdlist") %>
<table class="pcCPcontent">
    <tr>
        <td>
            <table id="FindProducts" class="pcCPcontent">
                <tr>
                    <td>
						<%
                        src_FormTitle1="STEP 3 - Find the configurable products"
                        src_FormTitle2="Assign/Remove Configurable-Only Items To/From Multiple Products"
                        src_FormTips1="Use the following filters to look for the configurable products to/from which you would like to assign/remove the items selected on the previous page."
                        src_FormTips2="Select one or more configurable products to/from which you would like to assign/remove configurable-only items."
                        src_IncNormal=0
                        src_IncBTO=1
                        src_IncItem=0
                        src_DisplayType=1
                        src_ShowLinks=0
                        src_FromPage="AddRmvBTOItemsMulti3.asp?action=add&prdlist=" & request("prdlist")
                        src_ToPage="AddRmvBTOItemsMulti4.asp?action=add"
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
