<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
%>
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
  <div class="pcMainContent">
    <h1><%=dictLanguage.Item(Session("language")&"_sdsMain_1")%></h1>
        
		<%
		msg = ""
		code = getUserInput(request.QueryString("msg"),0)
		Select Case code
			Case "1" : msg = dictLanguage.Item(Session("language")&"_ModsdsB_1")
		End Select
		If msg<>"" then
		%>
			<div class="pcErrorMessage"><%=msg%></div>
		<%
		end if
		%>
      
    <ul>
      <li><a href="pcmodsdsA.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_2")%></a></li>
      <li><a href="sds_viewPast.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_3")%></a></li>
      <li><a href="contact.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_4")%></a></li>
      <li><a href="sds_LO.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_5")%></a></li>
    </ul> 
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->
