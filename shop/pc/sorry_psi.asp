<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<div class="pcMainContent">
    <p><%= dictLanguage.Item(Session("language")&"_sorry_1")%></p>
    <hr>
    <p><%= Server.HTMLEncode(request.querystring("ErrMsg")) %></p>
    <br>
    <div class="pcFormButtons">
      <a class="pcButton pcButtonBack" href="javascript:history.back(-1);">
        <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
      </a>
    </div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
