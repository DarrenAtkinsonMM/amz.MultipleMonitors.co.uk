<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->

<!--#include file="inc_headerV5.asp"-->

<html>
<head>
<title><%= dictLanguage.Item(Session("language")&"_catering_9")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","pcStorefront.css")%>" />
</head>
<body id="pcPopup">
<div id="pcMain">
	<div class="pcMainTable">
		<div class="pcSectionTitle"><%= dictLanguage.Item(Session("language")&"_catering_9")%></div>
		<div class="pcFormItem">
			<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_catering_10")%></div>
		</div>

		<hr />
		 
		<%
		query="select * from ZipCodeValidation order by zipcode asc"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=connTemp.execute(query)
		If rstemp.eof Then %>

		<div class="pcErrorMessage"><%= dictLanguage.Item(Session("language")&"_catering_11")%></div>
								
	<% Else %>
			<ul id="pcDeliveryZipCodesList">

			<%
			strClass="alt"
			Do While NOT rstemp.EOF
				zipcode=rstemp("zipcode")
				If strClass = "alt" Then
					strClass = ""
				Else
					strClass = "alt"
				End If 
			%>
			
			<li class="<%= strClass %>"><%=zipcode%></li>
								
			<% rstemp.MoveNext
				 Loop
			%>
			</ul>
			<%
				 End If
			%>
			
	</div>
</div>
</body>
</html>
<% 
set rstemp = nothing
call closeDB()
%>
