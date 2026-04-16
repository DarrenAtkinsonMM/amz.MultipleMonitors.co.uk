<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<div id="pcMain" class="pcTechErr">
	<div class="pcMainContent">
		<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_techErr_1")%></div>
		<div class="pcSpacer"></div>
    
    <p><strong><%=dictLanguage.Item(Session("language")&"_techErr_6")%></strong></p>
    <p><%=dictLanguage.Item(Session("language")&"_techErr_7")%><b><%=Session("pcStrCustRefID")%></b><%=dictLanguage.Item(Session("language")&"_techErr_8")%></p>
    <p>&nbsp;</p>
    <%=dictLanguage.Item(Session("language")&"_techErr_9")%>
    
		<div class="pcSpacer"></div>
    
	</div>
</div>
<%
Session("pcStrCustRefID") = ""
call clearLanguage()
%>
<!--#include file="footer_wrapper.asp"-->
