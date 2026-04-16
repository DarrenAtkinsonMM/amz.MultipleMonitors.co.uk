<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="AffLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<%
' Load affiliate ID
affVar=session("pc_idaffiliate")
if not validNum(affVar) then
	response.redirect "AffiliateLogin.asp"
end if
%>
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<div class="pcMainContent">
		<h1><%=dictLanguage.Item(Session("language")&"_AffMain_1")%></h1>

		<% 
			msg = ""
			code = getUserInput(Request.QueryString("msg"), 0)
			Select Case code
			Case "1" : msg = dictLanguage.Item(Session("language")&"_ModAffb_1")
			End Select

			If msg<>"" Then
				%><div class="pcErrorMessage"><%= msg %></div><% 
			End If 
		%>
		
		<ul>
			<li><a href="pcmodAffA.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_2")%></a></li>
			<li><a href="Affgenlinks.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_3")%></a></li>
			<li><a href="AffCommissions.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_4")%></a></li>
			<li><a href="AffLO.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_5")%></a></li>
		</ul> 
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
