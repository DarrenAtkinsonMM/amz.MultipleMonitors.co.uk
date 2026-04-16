<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->

<% 
if request.form("submit")<>"" then
	CPServer=request.form("CPServer")
	Session("ship_CP_Server")=CPServer
	CPUsername=request.form("CPUsername")
	Session("ship_CP_Username")=CPUsername
	CPPassword=request.form("CPPassword")
	Session("ship_CP_Password")=enDeCrypt(CPPassword, scCrypPass)
	CPCustNo=request.form("CPCustNo")
	Session("ship_CP_CustNo")=CPCustNo
	
	if CPServer="" or CPUsername="" or CPPassword="" or CPCustNo="" then
		call closeDb()
response.redirect "ConfigureOption4.asp?msg="&Server.URLEncode("All fields are required.")
		response.end
	end if
	call closeDb()
response.redirect "4_Step2.asp"
	response.end
else %>
<form name="form1" method="post" action="ConfigureOption4.asp" class="pcForms">
    <table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr> 
        <td colspan="2"><h2>Enable Canada Post - <a href="https://www.canadapost.ca/cpc/en/business/ecommerce.page?ecid=murl12008339" target="_blank">Web site</a></h2></td>
        </tr>
        <tr> 
            <td colspan="2">
            ProductCart utilizes Canada Post's Sell Online XML Direct Connection. To enable Canada Post in ProductCart, you will need an Online Profile:
<br><br>
        <a href="https://www.canadapost.ca/web/en/kb/details.page?article=online_profiles_for_&cattype=kb&cat=accountinformation&subcat=myaccount" target="_blank"><b><u>Click Here to Obtain your Online Profile</u></b></a>
            </td>
        </tr>
        <tr>
      		<td colspan="2">Once you have your Online Profile, enter your credentials in the fields below and specify whether your store is 'live' (Production Mode), or still in testing (Development Mode). If you enable Development Mode, remember to switch to Production when your store is operational. Click 'Continue' to save your information.</td>
    	</tr>
        <% 
        pcv_boolIsDevelopment = False
        If Session(Ship_CPServer) = "https://ct.soa-gw.canadapost.ca/rs/ship/price" Then 
            pcv_boolIsDevelopment = True
        End If
        %>
	    <tr>
      		<td width="20%">&nbsp;</td>
      		<td width="80%">
                <input type="radio" name="CPServer" value="https://ct.soa-gw.canadapost.ca/rs/ship/price" class="clearBorder" <% If pcv_boolIsDevelopment Then %>checked<% End If %> /> <strong>Development</strong> - https://ct.soa-gw.canadapost.ca/rs/ship/price
      		</td>
    	</tr>
    	<tr>
      		<td>&nbsp;</td>
      		<td><input type="radio" name="CPServer" value="https://soa-gw.canadapost.ca/rs/ship/price" class="clearBorder" <% If Not pcv_boolIsDevelopment Then %>checked<% End If %> /> <strong>Production</strong> - https://soa-gw.canadapost.ca/rs/ship/price</td>
    	</tr>
        <tr>
            <td align="right">Username:</td>
            <td><input type="text" name="CPUsername" size="30" value="<%=Session(Ship_CP_Username)%>"></td>
        </tr>
        <tr>
            <td align="right">Password:</td>
            <td><input type="password" name="CPPassword" size="30" value="<%=Session(Ship_CP_Password)%>" class="pcAutoCompleteOff" autocomplete="new-password"></td>
        </tr>
        <tr>
            <td align="right">Customer Number:</td>
            <td><input type="text" name="CPCustNo" size="30" value="<%=Session(Ship_CP_CustNo)%>"></td>
        </tr>
        <tr>
            <td colspan="2">&nbsp;</td>
        </tr>
        <tr>
            <td>&nbsp;</td>
            <td> 
                <input type="submit" name="Submit" value="Continue" class="btn btn-primary">
                &nbsp;
                <input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
            </td>
        </tr>
    </table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->
