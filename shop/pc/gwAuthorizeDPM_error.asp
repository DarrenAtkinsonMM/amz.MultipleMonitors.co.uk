<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="header_wrapper.asp"-->

<div id="pcMain" class="pcMsgb">
	<div class="pcMainContent">
        <div class="pcErrorMessage">
            <%= Request("msg") %>
            <br /><br /><br />
            <a href="gwSubmit.asp?psslurl=gwAuthorizeDPM.asp&idCustomer=<%= Request("idcust")%>&idOrder=<%= Request("idorder")%>&ordertotal=<%= Request("amount")%>" class="pcButton pcButtonBack">
                <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
            </a>
        </div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->