<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="CustLIv.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
set rstemp = Server.CreateObject("ADODB.Recordset")
mySQL = "UPDATE customers set iRewardPointsAccrued=0 where irewardPointsAccrued is Null"
rstemp.Open mySQL, connTemp, adOpenStatic
mySQL = "UPDATE customers set iRewardPointsUsed=0 where irewardPointsUsed is Null"
rstemp.Open mySQL, connTemp, adOpenStatic
mySQL = "SELECT iRewardPointsAccrued,iRewardPointsUsed FROM customers WHERE idCustomer=" & session("idCustomer")

rstemp.Open mySQL, connTemp, adOpenStatic

iRewardPointsAccrued = rstemp("iRewardPointsAccrued")
iRewardPointsUsed = rstemp("iRewardPointsUsed")
if iRewardPointsAccrued="" then
	iRewardPointsAccrued=0
end if
if iRewardPointsUsed="" then
	iRewardPointsUsed=0
end if
iBalance = INT(iRewardPointsAccrued) - Int(iRewardPointsUsed)
if iBalance=0 then
	iDollarValue =0
else
	iDollarValue = iBalance * (RewardsPercent / 100)
end if
iRewardPointsHistory=iRewardPointsAccrued
rstemp.Close
%>
		
	<div id="pcMain">
		<div class="pcMainContent">
			<h1><%= RewardsLabel%></h1>
      
      <div class="pcShowContent">
        <p><%= ship_dictLanguage.Item(Session("language")&"_custRewards_a") & RewardsLabel & ship_dictLanguage.Item(Session("language")&"_custRewards_b") & "<strong>" & iBalance & "</strong>" %>.</p>
        <p><%= ship_dictLanguage.Item(Session("language")&"_custRewards_c") & "<strong>" & scCurSign&money(iDollarValue) & "</strong>" & ship_dictLanguage.Item(Session("language")&"_custRewards_d")%></p>
        <br/>
        <p><%= ship_dictLanguage.Item(Session("language")&"_custRewards_e") & RewardsLabel & ship_dictLanguage.Item(Session("language")&"_custRewards_f")%></p>
        <br/>
        <p><%= ship_dictLanguage.Item(Session("language")&"_custRewards_g") & iRewardPointsHistory & " " & RewardsLabel %>.</p>
        <p><%= ship_dictLanguage.Item(Session("language")&"_custRewards_h") & iRewardPointsUsed & " " & RewardsLabel %>.</p>
      </div>
      
      <hr>
      
			<a class="pcButton" href="javascript:history.go(-1)">
        <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
      </a>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
