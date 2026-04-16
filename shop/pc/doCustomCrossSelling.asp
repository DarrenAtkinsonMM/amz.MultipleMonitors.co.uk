<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
	csw_status=0
	query= "SELECT csw_status FROM crossSelldata WHERE id=1"
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)
	if not rs.eof then
		csw_status=rs("csw_status")
	end if
	set rs=nothing
if csw_status=-1 then

    'Randomly show 5 products from the same brand
    query="SELECT TOP (5) products.idProduct, products.description, products.price, products.sales, products.smallImageUrl, Brands.BrandName, products.sku "
    query=query & "FROM products WITH (NOLOCK)  LEFT OUTER JOIN Brands WITH (NOLOCK) ON products.IDBrand = Brands.IdBrand "
    query=query & "WHERE (products.stock > 0) AND (products.formQuantity <> - 1) AND (products.active = - 1) AND (products.smallImageUrl<>'')  AND  (products.IDBrand = " & pidbrand & " and products.IDBrand <> 0) and (products.idProduct<>"&pidproduct&") "
    query=query & "ORDER BY products.sales, NEWID() DESC"

    set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)	

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NOT rs.eof then
        moreInBrandArry=rs.getRows()
	end if
	set rs=nothing


    'Randomly show 5 products from the current category 
    query="SELECT top(5) * from ( select products.idproduct, products.description, products.sales, products.price, products.smallImageUrl, sku, ROW_NUMBER() OVER(PARTITION BY products.idbrand ORDER BY sales DESC) rn "
    query=query &" FROM  products WITH (NOLOCK) INNER JOIN  categories_products WITH (NOLOCK) ON products.idProduct = categories_products.idProduct  "
    query=query &" WHERE (products.stock > 0) AND (products.formQuantity <> - 1) AND (products.active = - 1) AND (products.smallImageUrl<>'') AND (categories_products.idCategory = " & pidCategory & ") and (products.idproduct<>"& pidproduct & ") ) a "
    'query=query &" where rn=1  "
    query=query &" order by sales, NEWID() "

    set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NOT rs.eof then
        moreInCatArry=rs.getRows()
	end if
	set rs=nothing



    'Build structured data
			storePath=scStoreUrl
			if right(storePath,1)="/" then
				storePath=Left(storePath,len(storePath)-1)
			end if 
		if isArray(moreIncatArry) or isArray(moreInBrandArry)	then		
		    %>
		    <script type="application/ld+json">
		    {
		        "@context":"http://schema.org",
		        "@type":"ItemList",
		        "itemListElement":[ 

		    <% 
			if isArray(moreInBrandArry) then             
			        if ubound(moreInBrandArry,2) > 0 then
			            for x=0 to ubound(moreInBrandArry,2)
			                pcStrPrdLinkBrand = pcGenerateSeoProductLink(moreInBrandArry(1,x), "",moreInBrandArry(0,x))
			                %>  
			                {

			                    "@type":"ListItem",
			                    "position":<%=x+1%>,
			                    "url":"<%=storePath & "/" & scPcFolder & "/pc/" & pcStrPrdLinkBrand%>"
			                }
			                <%if x<ubound(moreInBrandArry,2) then %>,<%end if%>

			            <%
			            next
			        end if
				end if
     
		   	if isArray(moreIncatArry) then
		        if ubound(moreInCatArry,2) > 1 then
		            response.write ","
		                for x=0 to ubound(moreInCatArry,2)
		                    pcStrPrdLinkBrand = pcGenerateSeoProductLink(moreInCatArry(1,x), "",moreInCatArry(0,x))
		                 %>  
		                    {
		                      "@type":"ListItem",
		                      "position":<%=x+(2+ubound(moreInCatArry,2))%>,
		                      "url":"<%=storePath & "/" & scPcFolder & "/pc/" & pcStrPrdLinkBrand%>"
		                    }
		                    <% 
							if x<ubound(moreInCatArry,2) then %>,<%end if%>

		                <%
		                next
		        end if
			end if		
		        %>  
		        ]
		    }
		    </script>
		<%
		end if

		if  isArray(moreIncatArry) then
		    if ubound(moreInCatArry,2)>1 then 
			cwidth=round(100/(ubound(moreInCatArry,2)+1))
			%>
			<div class="pcSectionTitle">More great Brands in this category...</div>
		    <div class="pcSectionContents">
		    	<div class="pcShowProducts">
			
		            <%'output brand listing 
		             for x=0 to ubound(moreInCatArry,2)
		                 pcStrPrdLinkBrand = pcGenerateSeoProductLink(moreInCatArry(1,x), "",moreInCatArry(0,x))
		            %>	
		             <table class="pcShowProductsHCS" style="width:<%=cwidth%>%; ">            
		            <tr><td class="pcShowProductImageH" style="height:170px">
						<p><a href="<%=pcStrPrdLinkBrand%>" ><img  style="max-height:150px"  alt="<%=moreInCatArry(1,x)%>; " src="catalog/<%=moreInCatArry(4,x)%>" alt="<%=moreInCatArry(1,x)%> "></a></p>
						</td></tr>
					<tr><td class="pcShowProductInfoH"  style="vertical-align:top">
						<p class="pcShowProductName"><a href="<%=pcStrPrdLinkBrand%>" ><%=moreInCatArry(1,x)%> </a><p>
					    <p class="pcShowProductPrice">Price $<%=moreInCatArry(3,x)%></p>
						</td></tr>
		            </table>
		        <%next %>
		  
		        </div>
		</div>		<%
			end if
   		end if
		if  isArray(moreInBrandArry) then     
			if ubound(moreInBrandArry,2)>0 then 
			cwidth=round(100/(ubound(moreInBrandArry,2)+1))
			%>
			<div class="pcSectionTitle">More great products from <%=moreInBrandArry(5,0) %>...</div>
		    <div class="pcSectionContents">
		    	<div class="pcShowProducts">
		            <%'output brand listing 
		             for x=0 to ubound(moreInBrandArry,2)
		                 pcStrPrdLinkBrand = pcGenerateSeoProductLink(moreInBrandArry(1,x), "",moreInBrandArry(0,x))
		            %>	
		             <table class="pcShowProductsHCS" style="width:<%=cwidth%>%;">            
		            <tr><td class="pcShowProductImageH"  style="height:170px">
						<p><a href="<%=pcStrPrdLinkBrand%>" ><img style="max-height:150px" alt="<%=moreInBrandArry(1,x)%>; <%=moreInBrandArry(6,x)%>" src="//resources.partspak.com/productcart/pc/catalog/<%=moreInBrandArry(4,x)%>" alt="<%=moreInBrandArry(1,x)%>; <%=moreInBrandArry(6,x)%>"></a></p>
						</td></tr>
					<tr><td class="pcShowProductInfoH" style="vertical-align:top">
						<p class="pcShowProductName"><a href="<%=pcStrPrdLinkBrand%>" ><%=moreInBrandArry(1,x)%>;</a><p>
					    <p class="pcShowProductPrice">Price $<%=moreInBrandArry(2,x)%></p>
						</td></tr>
		            </table>
		        <%next %>

		        </div>
		</div>
		<%end if
		end if
		%>
		
<%end if%>