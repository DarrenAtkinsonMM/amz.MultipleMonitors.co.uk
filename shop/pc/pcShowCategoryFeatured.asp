<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

if validNum(pFeaturedCategory) then
	' Get data about the featured subcategory
	query="SELECT categoryDesc, [image], largeimage, SDesc FROM categories WHERE idCategory=" &pFeaturedCategory&";"
	SET rsTemp=Server.CreateObject("ADODB.RecordSet")
	SET rsTemp=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsTemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	pcStrCategoryDesc=replace(rsTemp("categoryDesc"), """", "&quot;")
	pcStrCategoryDesc=replace(pcStrCategoryDesc, "&amp;", "&")
	pImage=rsTemp("image")
	plargeImage=rsTemp("largeimage")
	if pFeaturedCategoryImage=0 then
		pFeaturedCatImage=pImage
		else
		pFeaturedCatImage=plargeImage
	end if
	pcStrCategorySDesc=rsTemp("SDesc")							
	set rsTemp=nothing
	'// Call SEO Routine
	pcGenerateSeoLinks
	'//
%>

<div id="pcShowCategoryFeatured">
	<p><%=dictLanguage.Item(Session("language")&"_viewCategories_4")%>&quot;<%=pCategoryName%>&quot;<%=dictLanguage.Item(Session("language")&"_viewCategories_5")%></p>

	<div class="pcRow">
		<div class="pcShowCategoryP">
			<div class="pcShowCategoryImage">
				<%if pFeaturedCatImage<>"" then%>
					<a href="<%=Server.HtmlEncode(pcStrFeaturedCatLink)%>" data-idCategory="<%= intIdCategory %>"><img src="<%=pcf_getImagePath("catalog",pFeaturedCatImage)%>" alt="<%=pcStrCategoryDesc%>"></a>
				<%end if%>
			</div>
			<div class="pcShowCategoryInfoP">
				<div class="pcShowCategoryName"><a href="<%=Server.HtmlEncode(pcStrFeaturedCatLink)%>"><%=pcStrCategoryDesc%></a></div>
				<!-- Load category discount icon -->
				<%intIdCategory=pFeaturedCategory%>
				<!--#include file="pcShowCatDiscIcon.asp" -->
				<%		
				' Show short category description
				if not pcStrCategorySDesc="" then%>
					<%=pcStrCategorySDesc%>
				<%end if%>
			</div>
		</div>
	</div>

	<div class="pcClear"></div>
</div>
<% end if %>
