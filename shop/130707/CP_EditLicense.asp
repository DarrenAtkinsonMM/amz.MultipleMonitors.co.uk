<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration - Edit License" %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
if request.form("submit")<>"" then
	CPServer=request.form("CPServer")
	CPUsername=trim(request.form("CPUsername"))
	CPPassword=enDeCrypt(trim(request.form("CPPassword")), scCrypPass)
	CPCustNo=trim(request.form("CPCustNo"))
	
	if CPServer="" or CPUsername="" or CPPassword="" or CPCustNo="" then
		call closeDb()
response.redirect "CP_EditLicense.asp?msg="&Server.URLEncode("All fields are required.")
		response.end
	end if
	'update db
	mySQL="UPDATE ShipmentTypes SET shipserver='"&CPServer&"', userID='"&CPUsername&"', password='"&CPPassword&"', AccessLicense='"&CPCustNo&"' WHERE idShipment=7"
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(mySQL)
	set rs=nothing
	
	call closeDb()
response.redirect "viewShippingOptions.asp#CP"
	response.end
else 

	mySQL="SELECT shipserver, userID, password, AccessLicense FROM ShipmentTypes WHERE idShipment=7"
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(mySQL)
	if not rs.eof then
		CPServer=rs("shipserver")
		CPUsername=rs("userID")
		CPPassword=enDeCrypt(rs("password"), scCrypPass)
		CPCustNo=rs("AccessLicense")
	end if
%>
    <form name="form1" method="post" action="CP_EditLicense.asp" class="pcForms">
        <table class="pcCPcontent">
            <tr>
                <td colspan="2" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
            <tr> 
              <td colspan="2"><h2>Canada Post - <a href="https://www.canadapost.ca/cpc/en/business/ecommerce.page?ecid=murl12008339" target="_blank">Web site</a></h2></td>
                This Page contains your existing Canada Post Online Profile credentials. You should not need to edit or change your credentials unless you have been assigned a new Profile or Customer Number.
                
             </td>
            </tr>
            <tr>
                <td colspan="2">Here, you can specify whether your store is 'live' (Production Mode), or still in testing (Development Mode). If you enable Development Mode, remember to switch to Production when your store is operational. Click 'Continue' to save your information.</td>
            </tr>
            <%
            pcv_boolIsDevelopment = False
            If CPServer = "https://ct.soa-gw.canadapost.ca/rs/ship/price" Then 
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
                <td><input type="text" name="CPUsername" size="30" value="<%=CPUsername%>"></td>
            </tr>
            <tr>
                <td align="right">Password:</td>
                <td><input type="password" name="CPPassword" size="30" value="<%=CPPassword%>" class="pcAutoCompleteOff" autocomplete="new-password"></td>
            </tr>
            <tr>
                <td align="right">Customer Number:</td>
                <td><input type="text" name="CPCustNo" size="30" value="<%=CPCustNo%>"></td>
            </tr>
            <tr>
                <td colspan="2" class="pcCPspacer"></td>
            </tr>
            <tr> 
            <td></td>
            <td>
            <input type="submit" name="Submit" value="Continue" class="btn btn-primary"></td>
            </tr>
        </table>
    </form>
    <% 
    set rs=nothing
end if 
%>
<!--#include file="AdminFooter.asp"-->
