<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcMobileSettings.asp"-->
<% 
pageTitle="Mobile Commerce Settings"
pageIcon="pcv4_icon_settings.png"
section="layout"
%>
<%
Dim pcv_strPageName
pcv_strPageName="MobileSettings.asp"

msg = getUserInput(Request("msg"), 0)

If request("action")="add" Then
	session("adm_MobileOn")						= getUserInput(Request("pcvOn"), 0)
	session("adm_MobilePay")					= getUserInput(Request("pcvPay"), 0)
	session("adm_MobileLogo")					= getUserInput(Request("CompanyLogo"), 0)
	session("adm_MobileShowHomeNav")	= getUserInput(Request("pcvMobileShowHomeNav"), 0)
	session("adm_MobileShowHomeSP")		= getUserInput(Request("pcvMobileShowHomeSP"), 0)
	session("adm_MobileShowHomeNA")		= getUserInput(Request("pcvMobileShowHomeNA"), 0)
	session("adm_MobileShowHomeBS")		= getUserInput(Request("pcvMobileShowHomeBS"), 0)
	session("adm_MobileShowHomeFP")		= getUserInput(Request("pcvMobileShowHomeFP"), 0)
	session("adm_MobileShowNavTop")		= getUserInput(Request("pcvMobileShowNavTop"), 0)
	session("adm_MobileShowNavBot")		= getUserInput(Request("pcvMobileShowNavBot"), 0)
	session("adm_MobileIsApparelAddOn")= getUserInput(Request("pcvMobileIsApparelAddOn"), 0)
	tmpCTypes=""
	if request("pcvVisa")="1" then
		tmpCTypes=tmpCTypes & "V"
	end if
	if request("pcvMaster")="1" then
		if tmpCTypes<>"" then
			tmpCTypes=tmpCTypes & ", "
		end if
		tmpCTypes=tmpCTypes & "M"
	end if
	if request("pcvAM")="1" then
		if tmpCTypes<>"" then
			tmpCTypes=tmpCTypes & ", "
		end if
		tmpCTypes=tmpCTypes & "A"
	end if
	if request("pcvDR")="1" then
		if tmpCTypes<>"" then
			tmpCTypes=tmpCTypes & ", "
		end if
		tmpCTypes=tmpCTypes & "D"
	end if
	session("adm_MobilePayPalCardTypes")=tmpCTypes

	call closeDb()
response.redirect("../includes/PageCreateMobileSettings.asp")

End If
%>
<!--#include file="AdminHeader.asp"-->

<% If msg="success" Then %>

<%
	session("adm_MobileOn")=""
	session("adm_MobilePay")=""
	session("adm_MobileLogo")=""
	session("adm_MobileShowHomeNav")=""
	session("adm_MobileShowHomeSP")=""
	session("adm_MobileShowHomeNA")=""
	session("adm_MobileShowHomeBS")=""
	session("adm_MobileShowHomeFP")=""
	session("adm_MobileShowNavTop")=""
	session("adm_MobileShowNavBot")=""
	session("adm_MobileIsApparelAddOn")=""
	session("adm_MobilePayPalCardTypes")=""
%>

    <div class="pcCPmessageSuccess">
		Mobile Commerce Settings Saved Successfully!
        <br />
        <a href="mobileSettings.asp">Edit</a> them again or return to the <a href="menu.asp">Control Panel start page</a>.   
	</div>

<% Else
pcvMobileOn=scMobileOn
pcvMobilePay=scMobilePay
pcvMobileLogo=scMobileLogo
pcvMobileShowHomeNav=scMobileShowHomeNav
pcvMobileShowHomeSP=scMobileShowHomeSP
pcvMobileShowHomeNA=scMobileShowHomeNA
pcvMobileShowHomeBS=scMobileShowHomeBS
pcvMobileShowHomeFP=scMobileShowHomeFP
pcvMobileShowNavTop=scMobileShowNavTop
pcvMobileShowNavBot=scMobileShowNavBot
pcvMobileIsApparelAddOn=scMobileIsApparelAddOn 
pcvMobilePayPalCardTypes=scMobilePayPalCardTypes



query="SELECT pcMS_TurnOn,pcMS_Pay,pcMS_Logo,pcMS_ShowHomeNav,pcMS_ShowHomeSP,pcMS_ShowHomeNA,pcMS_ShowHomeBS,pcMS_ShowHomeFP,pcMS_ShowNavTop,pcMS_ShowNavBot, pcMS_IsApparelAddOn, pcMS_PayPalCardTypes FROM pcMobileSettings;"
set rs=connTemp.execute(query)

if not rs.eof then
	pcvMobileOn=rs("pcMS_TurnOn")
	pcvMobilePay=rs("pcMS_Pay")
	pcvMobileLogo=rs("pcMS_Logo")
	pcvMobileShowHomeNav=rs("pcMS_ShowHomeNav")
	pcvMobileShowHomeSP=rs("pcMS_ShowHomeSP")
	pcvMobileShowHomeNA=rs("pcMS_ShowHomeNA")
	pcvMobileShowHomeBS=rs("pcMS_ShowHomeBS")
	pcvMobileShowHomeFP=rs("pcMS_ShowHomeFP")
	pcvMobileShowNavTop=rs("pcMS_ShowNavTop")
	pcvMobileShowNavBot=rs("pcMS_ShowNavBot")
	pcvMobileIsApparelAddOn=rs("pcMS_IsApparelAddOn")
	pcvMobilePayPalCardTypes=rs("pcMS_PayPalCardTypes")
end if
set rs=nothing


if IsNull(pcvMobileOn) OR pcvMobileOn="" then
	pcvMobileOn=0
end if

%>

<form method="post" name="form1" action="<%=pcv_strPageName%>?action=add" class="pcForms">
	<table class="pcCPcontent">
	<%if msg<>"" then%>
	<tr>
		<td colspan="2">
			<div class="pcCPmessage"><%=msg%></div>
     	</td>
	</tr>
	<%end if%>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>
		<td colspan="2">For information on this feature, please see the <a href="http://wiki.productcart.com/mobile/mobile-commerce-settings" target="_blank">Mobile Commerce documentation</a>.</td>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>
		<th colspan="2">Turn Mobile Site On & Off</th>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2">
			<input type="radio" name="pcvOn" value="1" <%if pcvMobileOn="1" then%>checked<%end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_7")%>
			<input type="radio" name="pcvOn" value="0" <% if pcvMobileOn<>"1" then%>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_8")%>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>
		<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_29")%></th>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>       
		<td nowrap><%=dictLanguageCP.Item(Session("language")&"_cpCommon_312")%>:</td>
		<td> 
			<input type="text" name="CompanyLogo" value="<%=pcvMobileLogo%>" size="20">
			&nbsp;(e.g.: <i>mylogo.gif</i>) <a href="http://wiki.productcart.com/mobile/mobile-commerce-settings#company_logo" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="More information" border="0"></a></td>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>
		<th colspan="2">Display Settings</th>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>
		<td colspan="2"><strong>Home Page</strong>:</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="pcvMobileShowHomeSP" value="1" class="clearBorder" <%if pcvMobileShowHomeSP="1" then%>checked<%end if%>> Show the "Specials" text link in the slide menu</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="pcvMobileShowHomeNA" value="1" class="clearBorder" <%if pcvMobileShowHomeNA="1" then%>checked<%end if%>> Show the "New" arrivals text link in the slide menu</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="pcvMobileShowHomeBS" value="1" class="clearBorder" <%if pcvMobileShowHomeBS="1" then%>checked<%end if%>> Show the "Best Sellers" text link in the slide menu</td>
	</tr>
    <!--
	<tr>
		<td></td>
		<td><input type="checkbox" name="pcvMobileShowHomeFP" value="1" class="clearBorder" <%if pcvMobileShowHomeFP="1" then%>checked<%end if%>> Show the "Featured" products text link in the slide menu</td>
	</tr>
    -->
	<tr>
		<td class="pcCPspacer" colspan="2"><hr></td>
	</tr>
	<tr>
		<td colspan="2"><strong>Search Results</strong>:</td>
	</tr>
	<tr>
		<td align="right"><input type="checkbox" name="pcvMobileShowNavTop" value="1" class="clearBorder" <%if pcvMobileShowNavTop="1" then%>checked<%end if%>></td>
		<td>Show page navigation at the top of the search results</td>
	</tr>
	<tr>
		<td align="right"><input type="checkbox" name="pcvMobileShowNavBot" value="1" class="clearBorder" <%if pcvMobileShowNavBot="1" then%>checked<%end if%>></td>
		<td>Show page navigation at the bottom of the search results</td>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr> 
		<td colspan="2" style="text-align: center;">
			<input name="submit" type="submit" class="submit2" value="Save Settings">&nbsp;
    </td>
	</tr>
	</table>
</form>
<%
End If
%>
<!--#include file="AdminFooter.asp"-->
