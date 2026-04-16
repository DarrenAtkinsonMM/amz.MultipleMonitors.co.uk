<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp" -->
<%
validateForm "Affretreivepassword.asp"%>
<!--#include file="header_wrapper.asp"-->
<%
pcPageName = Request.ServerVariables("SCRIPT_NAME")
pcRequestRedirect=getUserInput(request("redirectUrl"),250)
if len(pcRequestRedirect)>0 then
	Session("pcSF_redirectUrl")=pcRequestRedirect
end if
pcfrUrl=getUserInput(request("frUrl"),250)
if len(pcfrUrl)>0 then
	Session("pcSF_pcfrUrl")=pcfrUrl
end if
%>
<div id="pcMain">
	<div class="pcMainContent">
		<form method="post" name="auth" action="<%= pcPageName %>"class="pcForms">
			<h2><%= dictLanguage.Item(Session("language")&"_AffLogin_8")%></h2>

			<div class="pcSpacer"></div>

			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_AffLogin_3")%></div>
				<div class="pcFormField">
					<%textbox "Email", "", 30, "textbox"%>
					<%validate "email", "email"%>
					<div class="pcSpacer"></div>
					<%validateError%>
				</div>
			</div>
			
			<div class="pcFormButtons">
				<button class="pcButton pcButtonSubmit" name="Submit" id="submit" value="Submit">
					<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
				</button>
			</div>
		</form>
	</div>
</div>
<%call clearLanguage()%>
<!--#include file="footer_wrapper.asp"-->
