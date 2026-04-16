<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<%
dim pcIntBrandID, pcvBrandsDescription, pcvBrandsSDescription, pcIntBrandsActive, pcIntSubBrandsView, pcvProductsView, pcIntBrandsParent, pcvBrandsMetaTitle, pcvBrandsMetaDesc, pcvBrandsMetaKeywords, pcvBrandsBrandLogoLg, pcIntCurrentPage, iPageSize, iPPageSize, pcIntIDBrand

dim pShowSKU, pShowSmallImg, pcProductHover

pShowSKU=1
pShowSmallImg=1
pcProductHover		= "pcShowProductBgHover"
		
Function GetParamFromSEOQueryStr(ParamName)
    Dim tmpValue1,tmpValue2    
    tmpValue2=""    
    if strSeoQueryString<>"" then
        if InStr(strSeoQueryString,ParamName & "=")>0 then
            tmpValue1=split(strSeoQueryString,ParamName & "=")
            if Instr(tmpValue1(1),"&")>0 then
                tmpValue2=Left(tmpValue1(1),Instr(tmpValue1(1),"&")-1)
            else
                tmpValue2=trim(tmpValue1(1))
            end if
        end if
    end if    
    GetParamFromSEOQueryStr=tmpValue2
End Function

'Number of brands to show
iPageSize=(scCatRow*scCatRowsPerPage)

'Number of products displayed on the brands page
iPPageSize=(scPrdRow*scPrdRowsPerPage)

strSeoQueryString=lcase(session("strSeoQueryString"))

'// View All
pcStrViewAll = Lcase(getUserInput(Request("viewall"),3))
if pcStrViewAll = "" then
	pcStrViewAll=GetParamFromSEOQueryStr("viewall")
end if
if pcStrViewAll = "yes" then
	iPPageSize = 9999
end if	

pcIntCurrentPage=getUserInput(Request("page"),10)
if pcIntCurrentPage = "" then
	pcIntCurrentPage=GetParamFromSEOQueryStr("page")
end if
If pcIntCurrentPage="" Then
	iPageCurrent=1
Else
	iPageCurrent=CInt(pcIntCurrentPage)
End If

'// Load Parent Brand Information - START
pcIntBrandID=trim(request("idbrand"))
pcIntBrandIDm=trim(request("idbrand"))
if pcIntBrandID = "" then
	pcIntBrandID=session("idBrandRedirectSF")
    pcIntBrandIDm=session("idBrandRedirectSF")
	if pcIntBrandID="" then
		pcIntBrandID=GetParamFromSEOQueryStr("idbrand")
        pcIntBrandIDm=GetParamFromSEOQueryStr("idbrand")
	end if
end if
session("idBrandRedirectSF")=""
if not validNum(pcIntBrandID) then pcIntBrandID=""
if not validNum(pcIntBrandIDm) then pcIntBrandIDm=""

iRecSize=10
pcStrPageName="viewbrands.asp"

%>
<!--#include file="pcStartSession.asp"-->
<%

'Decide Order By
Dim ProdSort 
ProdSort=trim(getUserInput(request("prodsort"),2))
if ProdSort = "" then
	ProdSort=GetParamFromSEOQueryStr("prodsort")
end if
if NOT validNum(ProdSort) then
	ProdSort=""
end if
if ProdSort="" then
	if UONum>0 then
		ProdSort="1"
	else
		ProdSort=PCOrd
end if
end if
%>
<!--#include file="prv_getSettings.asp"-->
<%

if pcIntBrandID<>"" then

	query="SELECT BrandName, pcBrands_SDescription, pcBrands_Description, pcBrands_SubBrandsView, pcBrands_ProductsView, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg FROM Brands WHERE pcBrands_Active=1 AND idBrand="&pcIntBrandID
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		set rstemp=nothing
		call closeDb()
		response.redirect "msg.asp?message=85"       
	End if
	parentBrandName=pcf_PrintCharacters(rstemp("BrandName"))
	parentBrandsSDescription=rstemp("pcBrands_SDescription")
	parentBrandsDescription=rstemp("pcBrands_Description")
	parentIntSubBrandsView=rstemp("pcBrands_SubBrandsView")
	parentIntProductsView=rstemp("pcBrands_ProductsView")
	parentIntBrandsParent=rstemp("pcBrands_Parent")
	parentBrandLogoLg=rstemp("pcBrands_BrandLogoLg")

	pcv_DefaultTitle=rstemp("pcBrands_MetaTitle")
	if isNull(pcv_DefaultTitle) or trim(pcv_DefaultTitle)="" then
		pcv_DefaultTitle=ClearHTMLTags2(parentBrandName,0)
	end if
	pcv_DefaultTitle = pcv_DefaultTitle & " - " & scCompanyName
	pcv_DefaultDescription=rstemp("pcBrands_MetaDesc")
	pcv_DefaultKeywords=rstemp("pcBrands_MetaKeywords")
	
	set rstemp=nothing

	if not validNum(parentIntSubBrandsView) then parentIntSubBrandsView=0
	if not validNum(parentIntBrandsParent) then parentIntBrandsParent=0
	
end if
'// Load Parent Brand Information - END

' OVERRIDE page style: check to see if a querystring or a form is sending the page style.
Dim pcPageStyle, strSeoQueryString

pcPageStyle = LCase(getUserInput(Request("pageStyle"),1))

'// Check querystring saved to session by 404.asp
if pcPageStyle = "" then
	pcPageStyle=GetParamFromSEOQueryStr("pagestyle")
end if

if pcPageStyle = "" then
	pcPageStyle = parentIntProductsView
end if

if pcPageStyle = "" Or IsNull(pcPageStyle) then
	pcPageStyle = lcase(bType)
end if

Dim pIntProductColumns, pIntProductRows
' How many products per row
pIntProductColumns=scPrdRow

' How many rows per page
pIntProductRows=scPrdRowsPerPage

%>
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<div id="pcMain" class="pcViewBrands">
	<div class="pcMainContent">
  
	<% if pcIntBrandID="" then %>
		<h1><%= dictLanguage.Item(Session("language")&"_titles_8")%></h1> 
	<% else %>
		<h1><%=parentBrandName%></h1>
		<% if parentBrandLogoLg<>"" then %>
			<div class="pcShowBrandLargeImage"><img src="<%=pcf_getImagePath("catalog",parentBrandLogoLg)%>" alt="<%=ClearHTMLTags2(parentBrandName,0)%>"></div>
		<% end if %>
		
		<% If pcf_HasHTMLContent(parentBrandsDescription) Then %>
			<div class="pcPageDesc"><%=pcf_FixHTMLContentPaths(parentBrandsDescription)%></div>
		<% End If %>
	<% end if %>
		
<%
'// Load data from Existing Brands - START

	'// Look for subBrand
	if pcIntBrandID<>"" then
		query1=" AND pcBrands_Parent="&pcIntBrandID
		else
		query1=" AND pcBrands_Parent=0"
	end if
	
	intSubBrandExist=0

	query="SELECT idbrand, BrandName, BrandLogo, pcBrands_Description, pcBrands_SDescription, pcBrands_SubBrandsView, pcBrands_ProductsView, pcBrands_Active, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg FROM Brands WHERE pcBrands_Active=1"&query1&" ORDER BY pcBrands_Order, BrandName ASC;"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.PageSize=iPageSize
	rs.CacheSize=iPageSize
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	dim iPageCount
	iPageCount=rs.PageCount
	If iPageCurrent > iPageCount Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1

	If not rs.eof then
	
	rs.AbsolutePage=iPageCurrent

'// Load data from Existing Brand - END
%>

		<% If pcIntBrandID <> "" Then %>
			<h3><%= dictLanguage.Item(Session("language")&"_viewBrand_2") & """" & parentBrandName & """" %></h3>
		<% End If %>
		
		<%
			Dim pcvBrandsLink,pcStrBrandLink,pcStrBrandLink2

			pcBrandDisplaySetting = parentIntSubBrandsView

			If pcBrandDisplaySetting = "3" Or IsNull(pcBrandDisplaySetting) Or pcBrandDisplaySetting = "" Then
				pcBrandDisplaySetting = scCatImages
			End If

			If pcBrandDisplaySetting="2" Then
			%>
				<div class="pcShowContent">
					<form class="pcForms">
						<% if trim(parentBrandName)<>"" then %>
            	<%=dictLanguage.Item(Session("language")&"_viewBrand_4")%>&quot;<%=parentBrandName%>&quot;:&nbsp;
            <% else %>
              <%=dictLanguage.Item(Session("language")&"_viewBrand_5")%>
            <% end if %>
            <select id="pcSortBox" class="form-control" onChange="window.location.href=this.options[selectedIndex].value" name="BrandDropSelect">
              <option>Browse Sub-Brands</option>
							<%
								iRecordsShown=0 
								Do While iRecordsShown < iPageSize And NOT rs.EOF
									intSubBrandExist=1
                                    
									pcIntIDBrand=rs("IdBrand")
									BrandName=pcf_PrintCharacters(rs("BrandName"))
									pcvProductsView=rs("pcBrands_ProductsView")
        
									'// Call SEO Routine
									pcGenerateSeoLinks
									'// SEO-E
          
									pcvBrandsLink=pcStrBrandLink2
									
									If Len(pcvProductsView) > 0 Then
										if scSeoURLs<>1 then
											pcvBrandsLink=pcvBrandsLink & "&"
										else
											pcvBrandsLink=pcvBrandsLink & "?"
										end if
          
										pcvBrandsLink=pcvBrandsLink & "pagestyle=" & pcvProductsView
									End If
									%>
										<option value="<%= pcvBrandsLink %>"><%= BrandName %></option>
									<%
									iRecordsShown=iRecordsShown + 1
									rs.movenext
								loop
      
								set rs=nothing
							%>
						</select>
					</form>
				</div>
			<%
			Else
			%>
    		<div class="pcBrandsWrapper">
					<%
					i=0 
					iRecordsShown=0 
					Do While iRecordsShown < iPageSize And NOT rs.EOF
						intSubBrandExist=1
                        
						pcIntIDBrand=rs("IdBrand")                        
						BrandName=pcf_PrintCharacters(rs("BrandName"))
						BrandLogo=rs("BrandLogo")
						pcvBrandsDescription=rs("pcBrands_Description")
						pcvBrandsSDescription=rs("pcBrands_SDescription")
						pcIntSubBrandsView=rs("pcBrands_SubBrandsView")
						pcvProductsView=rs("pcBrands_ProductsView")
						pcIntBrandsActive=rs("pcBrands_Active")
						pcIntBrandsParent=rs("pcBrands_Parent")
							' Check for SubBrands
							Dim pcIntSubBrandsExist
							pcIntSubBrandsExist=0
							query="SELECT idbrand FROM brands WHERE pcBrands_Parent="&pcIntIDBrand
							set rstemp=Server.CreateObject("ADODB.RecordSet")
							set rstemp=conntemp.execute(query)
							if not rstemp.EOF then
								pcIntSubBrandsExist=1
							end if
							set rstemp=nothing
						pcvBrandsMetaTitle=rs("pcBrands_MetaTitle")
						pcvBrandsMetaDesc=rs("pcBrands_MetaDesc")
						pcvBrandsMetaKeywords=rs("pcBrands_MetaKeywords")
						pcvBrandsBrandLogoLg=rs("pcBrands_BrandLogoLg")
        
						If BrandLogo="" Then
							BrandLogo="no_image.gif"
						End if
						if not validNum(pcIntSubBrandsView) then pcIntSubBrandsView=0
						if not validNum(pcIntBrandsActive) then pcIntBrandsActive=1
						if not validNum(pcIntBrandsParent) then pcIntBrandsParent=0
          
						'// Call SEO Routine
						pcGenerateSeoLinks
						'// SEO-E
          
						pcvBrandsLink=pcStrBrandLink2
						
						If Len(pcvProductsView) > 0 Then
							if scSeoURLs<>1 then
								pcvBrandsLink=pcvBrandsLink & "&"
							else
								pcvBrandsLink=pcvBrandsLink & "?"
							end if
          
							pcvBrandsLink=pcvBrandsLink & "pagestyle=" & pcvProductsView
						End If
					%>
  
					<div class="pcColCount<%= scCatRow %>">
  
						<%

						pcShowBrandClass = "pcShowBrand"
						If scCatRow = 1 And pcBrandDisplaySetting="0" Then
							pcShowBrandClass = "pcShowBrandP"
						End If
						%>
							<div class="<%= pcShowBrandClass %> pcShowBrandBgHover">
							<%
							if pcBrandDisplaySetting<>"2" then %>
								<% '// List with Images OR Thumnails Only %>
								<% If (pcBrandDisplaySetting="0" Or pcBrandDisplaySetting="4") And Not (pcIntBrandID="" And sBrandLogo<>"1") Then %>
									
									<div class="pcShowBrandImage">
										<a href="<%=pcvBrandsLink%>"><img src="<%=pcf_getImagePath("catalog",BrandLogo)%>" alt="<%=ClearHTMLTags2(BrandName,0)%>"></a>
									</div>
								<% End If %>
							
								<% '// All Except Thumbnails Only %>
								<% If pcBrandDisplaySetting<>"4" Then %>
								<div class="pcShowBrandInfo">
									<div class="pcShowBrandName">
										<a href="<%=pcvBrandsLink%>"><%=BrandName%></a>
									</div>
									<% If Len(pcvBrandsSDescription) > 2 And pcShowBrandClass = "pcShowBrandP" Then %>
										<div class="pcShowBrandSDesc">
											<%= pcvBrandsSDescription %>
										</div>
									<% End If %>
								</div>
								<% End If %>
								<div class="pcClear"></div>
							<% else %>
							<% end if %>
						</div>
					</div>
				<% i=i + 1
				If i > (scCatRow-1) then 
					response.write "<div class='pcRowClear'></div>"
					i=0
				End If
				iRecordsShown=iRecordsShown + 1
				rs.movenext
				loop
      
				set rs=nothing
				%>
				</div>
			<% End If %>
			<div class="pcClear"></div>
		<%end if 'Have Sub-Brands
		set rs=nothing
		%>
		<% 
        '*******************************
        ' START Page Navigation
        '*******************************
        
        pcIntBrandID = pcIntBrandIDm
        pcIntIDBrand = pcIntBrandIDm
        
		'// Call SEO Routine
		pcGenerateSeoLinks
		'// SEO-E
		
		if scSeoURLs<>1 then
			pcvBrandsLink=pcStrBrandLink & "&"
		else
			pcvBrandsLink=pcStrBrandLink & "?"
		end if
		
        If iPageCount>1 then %>
        
            <div class="pcPageNav">
                <%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
                &nbsp;-&nbsp;
                <% if iPageCount>iRecSize then %>
                    <% if cint(iPageCurrent)>iRecSize then %>
                        <a href="<%=pcvBrandsLink%>page=1&ppage=1&prodsort=<%=prodsort%>">First</a>&nbsp;
                    <% end if %>
                    <% if cint(iPageCurrent)>1 then
                        if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
                            iPagePrev=cint(iPageCurrent)-1
                        else
                            iPagePrev=iRecSize
                        end if %>
                        <a href="<%=pcvBrandsLink%>page=<%=cint(iPageCurrent)-iPagePrev%>&ppage=1&prodsort=<%=prodsort%>">Previous <%=iPagePrev%> Pages</a>
                    <% end if
                    if cint(iPageCurrent)+1>1 then
                        intPageNumber=cint(iPageCurrent)
                    else
                        intPageNumber=1
                    end if
                else
                    intPageNumber=1
                end if
                
                if (cint(iPageCount)-cint(iPageCurrent))<iRecSize then
                    iPageNext=cint(iPageCount)-cint(iPageCurrent)
                else
                    iPageNext=iRecSize
                end if
            
                For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
                    If Cint(pageNumber)=Cint(iPageCurrent) Then %>
                        <strong><%=pageNumber%></strong> 
                    <% Else %>
                        <a href="<%=pcvBrandsLink%>page=<%=pageNumber%>&ppage=1&prodsort=<%=prodsort%>"><%=pageNumber%></a>
                    <% End If 
                Next
                
                if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
                else
                    if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
                        <a href="<%=pcvBrandsLink%>page=<%=cint(intPageNumber)+iPageNext%>&ppage=1&prodsort=<%=prodsort%>">Next <%=iPageNext%> Pages</a>
                    <% end if
                
                    if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
                        &nbsp;<a href="<%=pcvBrandsLink%>page=<%=cint(iPageCount)%>&ppage=1&prodsort=<%=prodsort%>">Last</a>
                    <% end if 
                end if %>
            </div>

        <% end if
        '*******************************
        ' END Page Navigation
        '*******************************
		%>


		<%
		'*******************************
		' START show products
		'*******************************
        If len(pcIntBrandID)>0 Then
            If cint(pcIntBrandID)>0 Then            
                %>
                <!--#include file="pcShowProducts.asp" -->
                <%                
                select case ProdSort
                    Case "0": query1 = " ORDER BY products.SKU Asc"
                    Case "1": query1 = " ORDER BY products.description Asc" 	
                    Case "2": 
                        If Session("customerType")=1 then
                            query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
                        else
                            query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
                        End if
                    Case "3":
                        If Session("customerType")=1 then
                            query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
                        else
                            query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
                        End if	
                end select
        
                ' If customer is not wholesale, disallow wholesale-only categories
				if NOT session("customerType")="1" then
					queryW = " AND categories.pccats_RetailHide<>1"
				end if
				' If admin preview, ignore hidden categories
				if session("pcv_intAdminPreview")<>1 then
					queryHC = " AND categories.iBTOhide<>1" & queryW
				end if
				
				'// Query Products of current Brand
                query="SELECT DISTINCT products.idProduct, products.sku, products.description, products.price, products.listhidden, products.listprice, products.serviceSpec, products.bToBPrice, products.smallImageUrl,products.noprices,products.stock, products.noStock,products.pcprod_HideBTOPrice,products.pcProd_BackOrder, products.FormQuantity,products.pcProd_BTODefaultPrice,cast(products.sDesc as varchar(8000)) sDesc, 0, 0, products.pcprod_OrdInHome, products.sales, products.pcprod_EnteredOn, products.hotdeal, products.pcProd_SkipDetailsPage FROM products JOIN categories_products ON products.idproduct=categories_products.idproduct JOIN categories ON categories.idCategory=categories_products.idCategory WHERE products.idBrand=" & pcIntBrandID & " AND active=-1 AND configOnly=0 and removed=0 " & queryHC & query1
                set rs=Server.CreateObject("ADODB.Recordset")   
                rs.CacheSize=iPPageSize
                rs.PageSize=iPPageSize
                pcv_strPageSize=iPPageSize

                rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
            
                if err.number<>0 then
                    call LogErrorToDatabase()
                    set rs=nothing
                    call closedb()
                    response.redirect "techErr.asp?err="&pcStrCustRefID
                end if
                
                pcIntPCurrentPage=getUserInput(Request("ppage"),10)
                If pcIntPCurrentPage = "" Then
                    pcIntPCurrentPage=GetParamFromSEOQueryStr("ppage")
                End If
                If pcIntPCurrentPage="" Then
                    iPageCurrent=1
                Else
                    iPageCurrent=CInt(pcIntPCurrentPage)
                End If
            
                dim pcv_intProductCount
                iPageCount=rs.PageCount

                If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
                If Cint(iPageCurrent) < 1 Then iPageCurrent=1
                
                if NOT rs.eof then
                    rs.AbsolutePage=Cint(iPageCurrent)
                    pcArray_Products = rs.getRows()
                    pcv_intProductCount = UBound(pcArray_Products,2)+1
                end if
            
                set rs = nothing
            
                if pcv_intProductCount<1 then 	' START IF-1: check if there are no products in this category...
                    if intSubBrandExist <> 1 then ' ... and there are no sub-categories, then show a message
                    %>
                        <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_viewBrand_1")%></div>
                        <br>
                    <% 
                    end if
                else 
                    
                    'if SORT BY drop-down does not exist, show page nav still %>
        
                    <hr />
                
                    <h3><%= dictLanguage.Item(Session("language")&"_viewBrand_3") & """" & parentBrandName & """" %></h3>
        
                    <div class="pcPageNav">
                        <% call PageNav(iPagecount, "Top") %>
                        <%
                        '=================================
                        'show SORT BY drop-down
                        '=================================
                        if HideSortPro<>"1" then
                        
                        '// Call SEO Routine
                        pcGenerateSeoLinks
                        '// SEO-E
                        
                        if scSeoURLs<>1 then
                        pcvBrandsLink=pcStrBrandLink & "&"
                        else
                        pcvBrandsLink=pcStrBrandLink & "?"
                        end if %>			
                        <div class="pcSortProducts">	
                          <form action="<%=pcvBrandsLink%>page=1&ppage=1" method="post" class="pcForms">
                            <%=dictLanguage.Item(Session("language")&"_viewCatOrder_5")%>
                            <select id="pcSortBox" class="form-control" name="prodsort" onChange="javascript:if (this.value != '') {this.form.submit();}">
                              <option value="0"<%if ProdSort="0" OR ProdSort="" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_1")%></option>
                              <option value="1"<%if ProdSort="1" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_2")%></option>
                              <option value="2"<%if ProdSort="2" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_3")%></option>
                              <option value="3"<%if ProdSort="3" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_4")%></option>
                            </select>
                          </form>
                                </div>
                                <% end if 
                        '=================================
                        'end SORT BY drop-down
                        '=================================
                        %>
                    <div class="pcClear"></div>
                </div>
              
                <%
                call pcShowProducts(iPPageSize, 0)
                %>
                <div class="pcPageNav">
                    <% call PageNav(iPagecount, "Bottom") %>
                    <div class="pcClear"></div>
                </div>
                <%
            End If ' END IF-1
            %>
        <!--#include file="atc_viewprd.asp"-->
        <%
        End If
    End If
	'*** END OF SHOW PRODUCTS	
	%>

	<%	
	Function PageNav(iPageCount, TopBottom)
		'*******************************
		' START Product Page Navigation
		'*******************************
		pcIntBrandID = pcIntBrandIDm
        
		'// Call SEO Routine
		pcGenerateSeoLinks
		'// SEO-E
		
		if scSeoURLs<>1 then
			pcvBrandsLink=pcStrBrandLink & "&"
		else
			pcvBrandsLink=pcStrBrandLink & "?"
		end if

		If iPageCount>1 then %>
			<div id="pcPagination<%= TopBottom %>" class="pcPagination">
			<%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
			&nbsp;-&nbsp;
			<% if iPageCount>iRecSize then %>
				<% if cint(iPageCurrent)>iRecSize then %>
					<a href="<%=pcvBrandsLink%>page=1&ppage=1&prodsort=<%=prodsort%>">First</a>&nbsp;
				<% end if %>
				<% if cint(iPageCurrent)>1 then
					if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
						iPagePrev=cint(iPageCurrent)-1
					else
						iPagePrev=iRecSize
					end if %>
					<a href="<%=pcvBrandsLink%>page=1&ppage=<%=cint(iPageCurrent)-iPagePrev%>&prodsort=<%=prodsort%>">Previous <%=iPagePrev%> Pages</a>
				<% end if
				if cint(iPageCurrent)+1>1 then
					intPageNumber=cint(iPageCurrent)
				else
					intPageNumber=1
				end if
			else
				intPageNumber=1
			end if
			
			if (cint(iPageCount)-cint(iPageCurrent))<iRecSize then
				iPageNext=cint(iPageCount)-cint(iPageCurrent)
			else
				iPageNext=iRecSize
			end if
		
			For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
				If Cint(pageNumber)=Cint(iPageCurrent) Then %>
					<strong><%=pageNumber%></strong> 
				<% Else %>
					<a href="<%=pcvBrandsLink%>page=1&ppage=<%=pageNumber%>&prodsort=<%=prodsort%>"><%=pageNumber%></a>
				<% End If 
			Next
			
			if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
			else
				if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
					<a href="<%=pcvBrandsLink%>page=1&ppage=<%=cint(intPageNumber)+iPageNext%>&prodsort=<%=prodsort%>">Next <%=iPageNext%> Pages</a>
				<% end if
			
				if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
					&nbsp;<a href="<%=pcvBrandsLink%>page=1&ppage=1&<%=cint(iPageCount)%>&prodsort=<%=prodsort%>">Last</a>
				<% end if 
			end if%>
			&nbsp;<a href="<%=pcvBrandsLink%>page=1&prodsort=<%=prodsort%>&viewall=yes" onClick="pcf_Open_viewAll();"><%=dictLanguage.Item(Session("language")&"_viewCategories_21")%></a>
			</div>
		<% end if
		'*******************************
		' END Product Page Navigation
		'*******************************	
        
	End Function
		
	if pcIntBrandID<>"" then
	%>
	<div class="pcSpacer"></div>
	
	<div class="pcFormButtons">    
    <a class="pcButton pcButtonBack" href="javascript:window.history.back();">
      <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
      <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
    </a>

  </div>
	<%
	end if        
	%>

    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
