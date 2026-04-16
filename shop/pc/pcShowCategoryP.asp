<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'Get more category details
	query = "SELECT categoryDesc,image,SDesc,pcCats_NotImg FROM Categories WHERE idCategory = " & intIdCategory
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	pcStrCategoryDesc=rs("categoryDesc")
	pcStrCategoryImg=rs("image")
	pcStrCategorySDesc=rs("SDesc")
	NotImg=rs("pcCats_NotImg")
	SET rs=nothing
	
	if (session("Mobile")="1") AND (pIntSubCategoryViewBAK="1" OR pIntSubCategoryViewBAK="2") AND ((Instr(Ucase(pcStrCategoryImg),"NO_IMAGE.GIF")>0) OR (IsNull(pcStrCategoryImg) OR (NOT pcStrCategoryImg<>""))) then
		NotImg=1
	end if
	
'// Call SEO Routine
pcGenerateSeoLinks
'//
%>
<div class="pcRow">
	<div class="pcShowCategoryP">
		<div class="pcShowCategoryImage">
			<%if NotImg<>"1" then%>
				<%if pcStrCategoryImg<>"" then%>
					<a href='<%=pcStrCatLink%>' data-idCategory="<%= intIdCategory %>" title="<%=pcStrCategoryDesc%>"><img src="<%=pcf_getImagePath("catalog",pcStrCategoryImg)%>" alt="<%=pcStrCategoryDesc%>"></a>
				<%else%>
					<a class="pcShowCategoryNoImage" href='<%=pcStrCatLink%>' data-idCategory="<%= intIdCategory %>" title="<%=pcStrCategoryDesc%>"><img src="<%=pcf_getImagePath("catalog","no_image.gif")%>" alt="<%=pcStrCategoryDesc%>"></a>
				<%end if%>
			<%end if%>
		</div>

		<div class="pcShowCategoryInfoP">
			<div class="pcShowCategoryName">
				<a href="<%=Server.HtmlEncode(pcStrCatLink)%>" data-idCategory="<%= intIdCategory %>" title="<%=pcStrCategoryDesc%>"><%=pcStrCategoryDesc%></a>
			</div>
			<!-- Load category discount icon -->
			<!--#include file="pcShowCatDiscIcon.asp" -->
			<%		
			' Show short category description
			If pcf_HasHTMLContent(pcStrCategorySDesc) Then %>
				<div class="pcShowCategorySDesc">
					<%= pcf_FixHTMLContentPaths(pcStrCategorySDesc) %>
				</div>
			<% End If %>
		</div>
	</div>
</div>

<div class="pcClear"></div>