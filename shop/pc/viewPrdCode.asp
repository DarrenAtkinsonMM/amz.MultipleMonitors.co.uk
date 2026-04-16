<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
Dim strShowBTO
strShowBTO=""
Dim showAddtoCart,showCustomize
showAddtoCart=0
showCustomize=0
Dim bCounter

Dim delEstProdAds, funDABundlesCalcs, funDAArrayCalcs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Enhanced Views Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim pcv_strUseEnhancedViews, pcv_strHighSlide_Align, pcv_strHighSlide_Template
Dim pcv_strHighSlide_Eval, pcv_strHighSlide_Effects, pcv_strHighSlide_MinWidth, pcv_strHighSlide_MinHeight
Dim pcv_strImageClickEvent, pcv_strCurrentClass

pcv_strUseEnhancedViews = True '// Turn Enhanced Views ON or OFF
pcv_strCurrentClass = "current"
pcv_strImageClickEvent = "if(document.readyState=='complete') {if ($(this).hasClass('" & pcv_strCurrentClass & "') || $('.highslide-image').is(':visible')) return hs.expand(this);} else {return(false);}"

pcv_strHighSlide_Align = "center" '// Align Images from anchor or screen
pcv_strHighSlide_Template = "rounded-white" '// Template
pcv_strHighSlide_Eval = "this.thumb.alt"
pcv_strHighSlide_Effects = "'expand', 'fade'"
pcv_strHighSlide_MinWidth = 250
pcv_strHighSlide_MinHeight = 250
pcv_strHighSlide_Fade = "true"
pcv_strHighSlide_Dim = 0.3
pcv_strHighSlide_Interval = 3500
pcv_strHighSlide_Heading = "highslide-caption" '// "highslide-heading"
pcv_strHighSlide_Hide = "true"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Enhanced Views Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="app-ViewPrdFuncs.asp"-->
<%

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PRODUCT ID - Retrieve and validate product ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
pIdProduct=session("idProductRedirect")
if not validNum(pIdProduct) then

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' DA EDIT - ReWrite means that we have to add a db call to grab correct product id based on url field 			
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
dim queryDa, rsDa

		queryDa="Select idProduct from products where pcUrl='"&request.QueryString("url")&"';"
		
		set rsDa=server.CreateObject("ADODB.RecordSet")
		set rsDa=conntemp.execute(queryDa)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsDa=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		'CHECKS IF RECORDS FOUND, IF NOT IT MEANS PRODUCT DELETED SO REDIRECT TO HOMEPAGE
		If Not rsDa.eof Then
			pIdProduct=rsDa("idProduct")
			Session("darrenpid")=rsDa("idProduct")
		Else
			Response.Status = "301 Moved Permanently"
			Response.AddHeader "Location", "https://www.multiplemonitors.co.uk/"
			Response.End
		End If

		set rsDa=nothing
		
	'pIdProduct=request("idProduct")
	if not validNum(pIdProduct) then
		'// Set Privacy Settings Test Cookie		
		Response.Cookies("pcC_detect") = "PASS"
		Response.Cookies("pcC_detect").Expires = Date() + 1
		call closedb()
		response.redirect "msg.asp?message=207"
	end if
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
' ADMIN PREVIEW: Check to see if this is a store manager preview
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim pcv_intAdminPreview
pcv_intAdminPreview=0
pcv_intAdminPreview=getUserInput(Request("adminPreview"),10)
	if validNum(pcv_intAdminPreview) and session("admin") <> 0 then
		session("pcv_intAdminPreview")=pcv_intAdminPreview
	else
		session("pcv_intAdminPreview")=0
	end if
pcv_IDProduct=pIdProduct

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' CATEGORY ID - Retrieve, validate and lookup category ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	'// Retrieve category ID from querystring and validate
	pIdCategory=session("intTempCatId")
	session("intTempCatId")=""
	if not validNum(pIdCategory) then
	pIdCategory=request.QueryString("idCategory")
		if not validNum(pIdCategory) then
			pIdCategory=0
		end if
	end if
	pcv_IDCategory=pIdCategory
		
	'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
	if pIdCategory=0 then		
		' If customer is not wholesale, disallow wholesale-only categories
		if not session("customerType")="1" then
			queryW = " AND categories.pccats_RetailHide<>1"
		end if
		' If admin preview, ignore hidden categories
		if session("pcv_intAdminPreview")<>1 then
			queryHC = " AND categories.iBTOhide<>1" & queryW
		end if		
		query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& pIdProduct & queryHC &";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.EOF then
			pIdCategory=rs("idCategory")
		else
			set rs=nothing
			call closeDb()
			response.redirect "msg.asp?message=86"   
		end if
		set rs=nothing
	end if
	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Previous and Next Buttons
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_newNextButtons
Dim i,pcArr,pcv_strPreviousPage,pcv_strNextPage

IF scDisplayPNButtons="1" THEN

	pcArr=split(session("pcstore_prdlist"),"*****")
	pcv_strPreviousPage=0
	pcv_strNextPage=0
	
    ' Feedback ID: #13290 - begin
    ' More than one product, then only show Prev & Next buttons
    if(ubound(pcArr) > 2) then	
	    For i=lbound(pcArr) to ubound(pcArr)
		    if trim(pcArr(i))<>"" then
			    if clng(pcArr(i))=clng(pIDProduct) then
				    pcv_strPreviousPage=i-1
				    if pcv_strPreviousPage=0 then
					    pcv_strPreviousPage=ubound(pcArr)-1
				    end if
				    pcv_strPreviousPage=pcArr(pcv_strPreviousPage)
				    pcv_strNextPage=i+1
				    if pcv_strNextPage=ubound(pcArr) then
					    pcv_strNextPage=1
				    end if
				    pcv_strNextPage=pcArr(pcv_strNextPage)
				    exit for
			    end if
		    end if
	    Next
	    %>
	    <div class="pcShowProductNav">
		    <a class="pcButton pcButtonPrevious" rel="nofollow" href="viewPrd.asp?idcategory=&idproduct=<%=pcv_strPreviousPage%>&frmsrc=1" data-idProduct="<%= pcv_strPreviousPage %>">
        	<img src="<%=pcf_getImagePath("",rslayout("pcLO_previous"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_pcLO_previous")%>">
          <span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_pcLO_previous")%></span>
        </a>
		    &nbsp;
		    <a class="pcButton pcButtonNext" rel="nofollow" href="viewPrd.asp?idcategory=&idproduct=<%=pcv_strNextPage%>&frmsrc=1" data-idProduct="<%= pcv_strNextPage %>">
        	<img src="<%=pcf_getImagePath("",rslayout("pcLO_next"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_pcLO_next")%>">
          <span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_pcLO_next")%></span>
        </a>
		    <br/>
	    </div>
        <div class="pcClear"></div>
	    <%
	end if
	' Feedback ID: #13290 - end
	
END IF
	
End Sub

Public Sub pcs_NextButtons

IF scDisplayPNButtons="1" THEN

IF ((session("pcstore_newsrc")="OK") or (request("frmsrc")="1")) AND (session("pcstore_prdlist")<>"") THEN
	session("pcstore_newsrc")=""
	call pcs_newNextButtons
ELSE
	session("pcstore_newsrc")=""
	session("pcstore_prdlist")=""
	'// We can only display this section if the category is greater than 0
	If pIdCategory>1 Then		
		'// Get our array
		'// Unfortunately we have to generate this everytime, since the admin may deactivate a product at any time.
		'// We can NOT use a session or save to the database for this reason.
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Decide Order By
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		query="Select POrder from categories_products where idCategory="& pIdCategory &";"
		set rsCatOrder=Server.CreateObject("ADODB.Recordset")     
		set rsCatOrder=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsCatOrder=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		UONum=0
		do while not rsCatOrder.eof
			pcv_strCatOrder=rsCatOrder("POrder")
			if pcv_strCatOrder<>"" AND isNULL(pcv_strCatOrder)=False then
				UONum=UONum+CLng(pcv_strCatOrder)
			end if
			rsCatOrder.MoveNext
		loop
		SET rsCatOrder=nothing		
		
		ProdSort=""
		if UONum>0 then
			ProdSort="19"
		else
			ProdSort="" & PCOrd
		end if			
		if ProdSort="" then
			ProdSort="0"
		end if
		
		select case ProdSort
			Case "19": query1 = " ORDER BY categories_products.POrder Asc"
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
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Decide Order By
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		' SELECT DATA SET
		' TABLES: products, categories_products
		query = 		"SELECT products.idProduct, products.description, products.pcProd_BTODefaultWPrice, products.bToBprice, products.pcProd_BTODefaultPrice, categories_products.idCategory, categories_products.POrder "
		query = query & "FROM products "
		query = query & "INNER JOIN categories_products "
		query = query & "ON products.idProduct = categories_products.idProduct "
		query = query & "WHERE categories_products.idCategory=" & pIdCategory &" "
		query = query & "AND products.active=-1 AND products.removed=0 AND products.configOnly=0 "
		query = query & "" & query1
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)			
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		pcv_strNextProductID = ""		
			
		if NOT rs.eof then
			Do until rs.eof
				'response.write rs("idProduct") & " " & rs("description") & "<br />"
				'// We need to form our Array
				xProductArrayCount = xProductArrayCount + 1
				pcv_strTmpProductID = rs("idProduct")
				if pcv_strTmpProductID <> "" then			
					pcv_strNextProductID = pcv_strNextProductID & pcv_strTmpProductID & chr(124)	
				end if	
			rs.movenext
			Loop
			
			'// Trim the last pipe if there is one
			xStringLength = len(pcv_strNextProductID)
			pcv_strShowButtons = 0
			if xStringLength>0 then
				pcv_strNextProductID = left(pcv_strNextProductID,(xStringLength-1))
				'// If there are no other pipes left then we only have one product in this category, so we can exit.
				if instr(pcv_strNextProductID,chr(124))>0 then
					pcv_strShowButtons = 1 '// show buttons
				else
					pcv_strShowButtons = 0
				end if
			end if
			
		end if
		set rs=nothing
		
		If pcv_strShowButtons = 1 Then
		
			'// Set Up Our Array
			pcArrayNextProductID = split(pcv_strNextProductID,chr(124))		
			pcv_intLBound = LBound(pcArrayNextProductID)
			pcv_intUBound = UBound(pcArrayNextProductID)
			
			'// Now find our place in the array
			For i = pcf_IDMaximum(pcv_intLBound, intStartIndex) To pcv_intUBound
				If CStr(pcArrayNextProductID(i)) = CStr(pIdProduct) Then
					pcv_intCurrentPosition = i
					Exit For
				End If
			Next
			
			'// Previous Product	
			if (pcv_intCurrentPosition-1) < pcv_intLBound then
				pcv_strPreviousPage=pcArrayNextProductID(pcv_intUBound)
			else
				pcv_strPreviousPage=pcArrayNextProductID(pcv_intCurrentPosition-1)
			end if
			
			'// Next Product
			if (pcv_intCurrentPosition+1) > pcv_intUBound then
				pcv_strNextPage=pcArrayNextProductID(pcv_intLBound)
			else
				pcv_strNextPage=pcArrayNextProductID(pcv_intCurrentPosition+1)
			end if

			'// Generate SEO Links
			'// Get product description
			query = "SELECT description FROM products WHERE idProduct="&pcv_strPreviousPage
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pcv_strPreviousPageDesc=rs("description")
			query = "SELECT description FROM products WHERE idProduct="&pcv_strNextPage
			set rs=conntemp.execute(query)			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pcv_strNextPageDesc=rs("description")
			set rs=nothing
			Call pcGenerateSeoLinks
			%>
			<div class="pcShowProductNav">
        <a class="pcButton pcButtonPrevious" href="<%=Server.HtmlEncode(pcStrPrdPreLink)%>" data-idproduct="<%= pcv_strPreviousPage %>" >
          <img src="<%=pcf_getImagePath("",rslayout("pcLO_previous"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_pcLO_previous")%>">
          <span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_pcLO_previous")%></span>
        </a>
			&nbsp;
        <a class="pcButton pcButtonNext" href="<%=Server.HtmlEncode(pcStrPrdNextLink)%>" data-idproduct="<%= pcv_strNextPage %>">
        	<img src="<%=pcf_getImagePath("",rslayout("pcLO_next"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_pcLO_next")%>">
          <span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_pcLO_next")%></span>
        </a>
			</div>
      <div class="pcSpacer"></div>
			<%
		End If '// If pcv_strShowButtons = 1 Then
	End If
END IF

END IF

End Sub


Function pcf_IDMaximum(ByVal x, ByVal y) 
  If x > y Then 
    pcf_IDMaximum = x 
  Else 
    pcf_IDMaximum = y 
  End If 
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Previous and Next Buttons
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Check if we should hide the options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Function pcf_VerifyShowOptions
	If pserviceSpec<>0 AND (pPrice=0 OR scConfigPurchaseOnly=1) Then
		pcf_VerifyShowOptions = false
	Else
		pcf_VerifyShowOptions = true
	End If
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Check if we should hide the options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Start SDBA
' START:  Display Back-Order Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_DisplayBOMsg
	If (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=1) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=1) Then
		If clng(pcv_intShipNDays)>0 then
			response.write "<div>"&dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b") & "</div>"
		End if
	End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Display Back-Order Message
'End SDBA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  BTOisConfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Function pcf_BTOisConfig

	query="SELECT categories.categoryDesc, products.description, configSpec_products.configProductCategory, configSpec_products.price, categories_products.idCategory, categories_products.idProduct, products.weight FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "tech2Err.asp?err="&pcStrCustRefID
	end if
	if NOT rstemp.eof then
		pcf_BTOisConfig = true
	else
		pcf_BTOisConfig = false
	end if 
	Set rstemp = nothing
	
	query="SELECT * FROM configSpec_Charges WHERE specProduct="&pIdProduct
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if	
	BTOCharges=0
	if not rstemp.eof then
		BTOCharges=1
	end if
	set rstemp=nothing

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  BTOisConfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  CATEGORY TREE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_CategoryTree
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: get category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pIdCategory > 0 then
	%>  <!--#include file="pcBreadCrumbs.asp"-->  <% 
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  get category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  show breadcrumbs - category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If strBreadCrumb<>"" then %>
		<div class="pcPageNav" itemscope itemtype="http://schema.org/BreadcrumbList">
			<%=dictLanguage.Item(Session("language")&"_viewCat_P_2") %>
			<%=strBreadCrumb %>
			<%
			intIdCategory=pIdCategory
			'// Load category discount icon
			%>
			<!--#include file="pcShowCatDiscIcon.asp" -->
		</div>
	<% 
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  show breadcrumbs - category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  CATEGORY TREE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show product name 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductName 
	If scAddThisDisplay=1 And Not pcv_IsQuickView = True Then 
		pcs_AddThis
	End If
%><h1 itemprop="name" <% If pcv_IsQuickView=True Then %>class="pcQVShowPrdName"<% End If %>><%=pDescription%></h1><%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show product name 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// <!-- DA - EDIT -->
Public Sub pcs_ProductTitle 
%>
<h3 class="color marginbot-0"><%=pDescription%></h3>
<p class="p-code marginbot-0">Product Code: <strong class="medium"><%=pSku%></strong></p>
<%
End Sub

Public Sub pcs_ShowDetailsTop 

	response.write(pDetailsTop)

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show SKU
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowSKU
IF pHideSKU<>"1" THEN%>
	<div class="pcShowProductSku">
		<%=dictLanguage.Item(Session("language")&"_viewCat_P_8")%>: <span itemprop="sku" id="sku"><%=pSku%></span>
	</div>
<%ELSE%>
	<input id="sku" name="sku" type="hidden" value="<%=pSku%>">
<%END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show SKU
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' PRV41 start
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Average Rating (from reviews)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowRating
IF pRSActive And pcv_ShowRatSum And pNumRatings>0 THEN%>
	<div class="pcShowProductRating">
  <% if pcv_RatingType="0" then
				query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pcv_IDProduct
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
		
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
	
				pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)
	
				query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " AND pcRev_Active=1 AND pcRev_MainRate>0"
				set rs=connTemp.execute(query)
	
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
	
				intCount = clng(rs("ct"))
	
				set rs=Nothing
				%>
				<%if pcv_tmpRating>"0" then%><a href="#productReviews" style="text-decoration: none;"><%=dictLanguage.Item(Session("language")&"_prv_2")%></a><img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%=dictLanguage.Item(Session("language")&"_prv_7")%>"><span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating"><meta itemprop="ratingValue" content="<%=pcv_tmpRating%>">%<meta itemprop="bestRating" content="100" /> <%=pcv_MainRateTxt1%> (<span itemprop="ratingCount"><%=intCount%></span>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)</span><%end if%>
				<%
			ELSE
				if pcv_CalMain="1" then     ' Can be set independently of sub-ratings 
					query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pcv_IDProduct
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)

					pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)
					
					query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " AND pcRev_Active=1 AND pcRev_MainDRate>0"
					set rs=connTemp.execute(query)
					if not rs.eof then
						intCount = clng(rs("ct"))
					end if
					set rs=nothing
	
					if CDbl(pcv_tmpRating)>0 then 
				    %>
					    <a href="#productReviews" style="text-decoration: none;"><%=dictLanguage.Item(Session("language")&"_prv_39")%></a> 
						<% Call WriteStar(pcv_tmpRating,1)%>
						<span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating">
						<meta itemprop="worstRating" content = "1" />
						<meta itemprop="ratingValue" content="<%=pcv_tmpRating%>" />
						<meta itemprop="bestRating" content="<%=pcv_MaxRating%>" />
                        <meta itemprop="ratingCount" content="<%=intCount%>" />
						</span>
			        <%end if
				    %>
		        <% else 'Will be calculated automatically by averaging sub-ratings
					Call CreateList()
				    pcv_tmpRating=CalRating()
					if CDbl(pcv_tmpRating)>0 then %>
				    <a href="#productReviews" style="text-decoration: none;"><%=dictLanguage.Item(Session("language")&"_prv_2")%></a>
				    <% Call WriteStar(pcv_tmpRating,1)
					end if %>
					<span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating">
					<meta itemprop="worstRating" content = "1" />
					<meta itemprop="ratingValue" content="<%=pcv_tmpRating%>" />
					<meta itemprop="bestRating" content="<%=pcv_MaxRating%>" />
                    <meta itemprop="ratingCount" content="<%=intCount%>" />
					</span>
		        <% end if
		    END IF 'Main Rating
		 %>

	</div>
<%END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Average Rating
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PRV41 end


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Custom Search Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_CustomSearchFields
Dim query,rs,pcArr,intCount,i
	query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pIdProduct & " AND pcSearchFieldShow=1 ORDER BY pcSearchFields.pcSearchFieldOrder ASC,pcSearchFields.pcSearchFieldName ASC;"
	set rs=connTemp.execute(query)
	IF not rs.eof THEN
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		response.Write("<div style='padding-top: 5px;'></div>")
		For i=0 to intCount
			searchFieldLink = "showsearchresults.asp?customfield="&pcArr(0,i)&"&SearchValues="&Server.URLEncode(pcArr(2,i))
				response.write "<div class='pcShowProductCustSearch'>"&pcArr(1,i)&": <a href='" & Server.HtmlEncode(searchFieldLink) & "'>"&pcArr(3,i)&"</a></div>"
		Next
	END IF
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Custom Search Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Weight (If admin turned on)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_DisplayWeight

Dim query,rs,totalSubWeight
query="SELECT sum(weight) As TotalWeight FROM Products WHERE pcProd_ParentPrd=" & pidProduct & " AND removed=0 and pcProd_SPInActive=0 GROUP BY pcProd_ParentPrd;"
set rs=connTemp.execute(query)
if not rs.eof then
	totalSubWeight=rs("TotalWeight")
else
	totalSubWeight=0
end if
set rs=nothing

if scShowProductWeight="-1" then

		if (int(pWeight)>0) OR (totalSubWeight>0) then
			response.write "<div class='pcShowProductWeight'>"
			response.write ship_dictLanguage.Item(Session("language")&"_viewCart_c")
			if scShipFromWeightUnit="KGS" then
				pKilos=Int(pWeight/1000)
				pWeight_g=pWeight-(pKilos*1000)
				pWeightUnit=pKilos
				if pWeight_g>0 then
					response.write dictLanguage.Item(Session("language")&"_viewCart_c") 
                    %>
                    <span itemprop="weight" id="appw1"><%=pWeightUnit%></span>
                    <%=" kg " %> 
                    <span itemprop="weight" id="appw2"><%=pWeight_g%></span>
                    <%=" g " %> 
                    <%
				else
					response.write dictLanguage.Item(Session("language")&"_viewCart_c")
                    %>
                    <span itemprop="weight" id="appw1"><%=pWeightUnit%></span> 
                    <%=" kg " %>
                    <span id="appw2" style="display:none"></span>
                    <%
				end if
			else
				pPounds=Int(pWeight/16)
				pWeight_oz=pWeight-(pPounds*16)
				pWeightUnit=pPounds
				if pWeight_oz>0 then
					response.write dictLanguage.Item(Session("language")&"_viewCart_c")
                    %>
                    <span itemprop="weight" id="appw1"><%=pWeightUnit%></span>
                    <%=" lbs " %>
                    <span itemprop="weight" id="appw2"><%=pWeight_oz%></span>
                    <%=" ozs " %> 
                    <%
				else
					response.write dictLanguage.Item(Session("language")&"_viewCart_c") %>
                    <span itemprop="weight" id="appw1"><%=pWeightUnit%></span>
                    <%=" lbs " %>
                    <span id="appw2" style="display:none"></span>
                    <%
				end if
			end if
			response.write "</div>"
		else%>
			<span style="display:none" id="appw1"></span>
			<span style="display:none" id="appw2"></span>
		<%end if
        
else%>
	<span style="display:none" id="appw1"></span>
	<span style="display:none" id="appw2"></span>
<%end if
'APP-E

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Weight
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Brand (If assigned)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowBrand
	if sBrandPro="1" then
		if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then

        pcIntBrandID = pIDBrand
        pcIntIDBrand = pIDBrand

        Call pcGenerateSeoLinks

		response.write "<div class='pcShowProductBrand'>"
		response.write dictLanguage.Item(Session("language")&"_viewPrd_brand")
		%>
			<a href="<%= pcStrBrandLink2 %>">
				<span itemprop="brand"><%=BrandName%></span>
			</a>
		<% 
		response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Brand
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Units in Stock (if on, show the stock level here)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_UnitsStock
	if scdisplayStock=-1 AND pNoStock=0 then
		if pstock > 0 then
			response.write "<div class=""pcShowProductStock"">"
			response.write dictLanguage.Item(Session("language")&"_viewPrd_19") & " " & pStock
			response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Units in Stock (if on, show the stock level here)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductDescription
	If pcf_HasHTMLContent(psDesc) Then
		response.write "<div class='pcShowProductSDesc'>"
		response.Write "<span itemprop=""description"">" & pcf_FixHTMLContentPaths(psDesc) & "</span>"
		response.write "</div>"
	End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_GetBTOConfiguration
Dim query,rs
	pcv_BTORP=Clng(0)
	strShowBTO=""		
	if pserviceSpec=true then
	 '// Product is BTO
		
		' Get data
		query="SELECT categories.categoryDesc, products.description, products.iRewardPoints,configSpec_products.configProductCategory, configSpec_products.price, configSpec_products.Wprice, categories_products.idCategory, categories_products.idProduct, products.weight, products.pcprod_minimumqty FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		iAddDefaultPrice=Cdbl(0)
		iAddDefaultWPrice=Cdbl(0)
		iAddDefaultPrice1=Cdbl(0)
		iAddDefaultWPrice1=Cdbl(0)
		
		if NOT rs.eof then 
			Dim FirstCnt
			FirstCnt=0
			if intpHideDefConfig="0" then
				strShowBTO= strShowBTO & "<div class='pcShowProductBTOConfig' style='padding-top: 10px; padding-bottom: 2px;'>"
				strShowBTO= strShowBTO & "<b>"&dictLanguage.Item(Session("language")&"_viewPrd_25")&"</b>"
				strShowBTO= strShowBTO & "</div>"
			end if
			do until rs.eof
				FirstCnt=FirstCnt+1
				strCategoryDesc=rs("categoryDesc")
				strDescription=rs("description")
				strConfigProductCategory=rs("configProductCategory")
				dblPrice=rs("price")
				dblWPrice=rs("Wprice")
				intIdCategory=rs("idCategory")
				intIdProduct=rs("idProduct")
				intReward=rs("iRewardPoints")
				if (intReward<>"") and (intReward<>"0") then
				else
				intReward=0
				end if

				if intReward="0" then
					query="SELECT pcprod_ParentPrd FROM Products WHERE idproduct=" & intIdProduct & " AND pcprod_ParentPrd>0;"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						pcv_tmpParent=rsQ("pcprod_ParentPrd")
						set rsQ=nothing
						query="SELECT iRewardPoints FROM Products WHERE idproduct=" & pcv_tmpParent & " AND active<>0;"
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							intReward=rsQ("iRewardPoints")
							set rsQ=nothing
							if intReward="" OR IsNull(intReward) then
								intReward=0
							end if
						end if
					end if
				end if

				intWeight=rs("weight")
				if Not ((intWeight<>"") and (intWeight<>"0")) then
					intWeight=0
				end if

				if intWeight="0" then
					query="SELECT pcprod_ParentPrd FROM Products WHERE idproduct=" & intIdProduct & " AND pcprod_ParentPrd>0;"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						pcv_tmpParent=rsQ("pcprod_ParentPrd")
						set rsQ=nothing
						query="SELECT weight FROM Products WHERE idproduct=" & pcv_tmpParent & " AND active<>0;"
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							intWeight=rsQ("weight")
							set rsQ=nothing
							if intWeight="" OR IsNull(intWeight) then
								intWeight=0
							end if
						end if
					end if
				end if

				pcv_iminqty=rs("pcprod_minimumqty")
				if IsNull(pcv_iminqty) or pcv_iminqty="" then
					pcv_iminqty=1
				end if
				if pcv_iminqty="0" then
					pcv_iminqty=1
				end if
				pcv_BTORP=pcv_BTORP+clng(intReward*pcv_iminqty)
				
				dblPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,0)
				dblWPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,1)
				iAddDefaultPrice=Cdbl(iAddDefaultPrice+dblPrice*pcv_iminqty)
				iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+dblWPrice*pcv_iminqty)
				iAddDefaultPrice1=Cdbl(iAddDefaultPrice1+dblPrice1*pcv_iminqty)
				iAddDefaultWPrice1=Cdbl(iAddDefaultWPrice1+dblWPrice1*pcv_iminqty)
				ItemPrice=0
				if Session("CustomerType")=1 then
					if (dblWPrice<>0) then
						ItemPrice=dblWPrice1
					else
						ItemPrice=dblPrice1
					end if
				else
					ItemPrice=dblPrice1
				end if
				if intpHideDefConfig="0" then
					strShowBTO= strShowBTO & "<div class='pcShowProductBTOConfig'>"
					strShowBTO= strShowBTO & "<b>"&strCategoryDesc&"</b>: "&strDescription
					strShowBTO= strShowBTO & "</div>"
				end if
				response.write "<input name=""CAT"&FirstCnt&""" type=""HIDDEN"" value=""CAG"&intIdCategory&""">"
				response.write "<input name=""CAG"&intIdCategory&"QF"" type=""HIDDEN"" value=""" & pcv_iminqty & """>"
				response.write "<input type=""hidden"" name=""CAG"&intIdCategory&""" value="""&intIdProduct&"_0_"&intWeight&"_" & ItemPrice & """>"
				rs.moveNext
			loop			
			response.write "<input type=""hidden"" name=""FirstCnt"" value="""&FirstCnt&""">"
		end if 
		set rs=nothing
	end if
End Sub

Public Sub pcs_BTOConfiguration
	if strShowBTO<>"" then
		response.write strShowBTO
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Reward Points
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_RewardPoints

	If RewardsActive=1 then
		If Not IsNumeric(iRewardPoints) Then iRewardPoints = 0
		If Not IsNumeric(pcv_BTORP) Then pcv_BTORP = 0
		
		' Show Reward Points associated with this product, if any
		' By default, Reward Points are not shown to Wholesale Customers
		rewardsStr = dictLanguage.Item(Session("language")&"_viewPrd_50") & Clng(iRewardPoints+clng(pcv_BTORP)) & "&nbsp;" & RewardsLabel & dictLanguage.Item(Session("language")&"_viewPrd_51")

		if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")<>"1" then
		%>
			<div class="pcShowProductRewards"><%= rewardsStr %></div>
    <%
		else
			' If the system is setup to include Wholesale Customers, then show Reward Points to them too
			if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")="1" and RewardsIncludeWholesale=1 then
			%>
				<div class="pcShowProductRewards"><%= rewardsStr %></div>
      <%
			end if 
		end If
	Else
		response.write "<script>DefaultReward=0;</script>"
	End If

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Reward Points
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show product prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductPrices
Dim rs,query,pcTestPrice,pcHidePricesIfNFS
Dim ShowSaleIcon,rsS,pcSCID,pcSCName,pcSCIcon,pcTargetPrice

ShowSaleIcon=0
pcTestPrice=0

	'// If product is "Not for Sale", should prices be hidden or shown?
	'// Set pcHidePricesIfNFS = 1 to hide, 0 to show.
	'// Here we leverage the "pnoprices" variable to change the behavior (a Control Panel setting could be added in the future)
	pcHidePricesIfNFS = 0
	if (pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0) and pcHidePricesIfNFS=1 then
		pnoprices=2
	end if

	' Don't show prices if the BTO product has been set up to hide prices (pnoprices)
	If pnoprices<2 Then
	
			query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & pidproduct & " AND Products.pcSC_ID>0;"
			set rsS=Server.CreateObject("ADODB.Recordset")
			set rsS=conntemp.execute(query)
					
			if not rsS.eof then
				ShowSaleIcon=1
				pcSCID=rsS("pcSC_ID")
				pcSCName=rsS("pcSC_SaveName")
				pcSCIcon=rsS("pcSC_SaveIcon")
				pcTargetPrice=rsS("pcSales_TargetPrice")
				
				query="SELECT pcSB_Price FROM pcSales_BackUp WHERE idProduct=" & pIdProduct & " AND pcSC_ID=" & pcSCID & ";"
				set rsQ=connTemp.execute(query)
				pcOrgPrice=0
				if not rsQ.eof then
					pcOrgPrice=rsQ("pcSB_Price")
				end if
				set rsQ=nothing
				%>
			<%end if
			set rsS=nothing
	
		' If this is a BTO product, calculate the base price as the sum of price + default prices
		pPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
		pBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
		if pserviceSpec=true then
			pPrice=Cdbl(pPrice+iAddDefaultPrice)
			pBtoBPrice=Cdbl(pBtoBPrice+iAddDefaultWPrice)
			pPrice1=Cdbl(pPrice1+iAddDefaultPrice1)
			pBtoBPrice1=Cdbl(pBtoBPrice1+iAddDefaultWPrice1)
		end if
		
		'START - Bing Cashback Gleaming
		If LSCB_STATUS = "1" AND LSCB_KEY <>"" Then
			pcv_strIsCashback = getUserInput(request("cashback"),1)
			If len(pcv_strIsCashback)>0 Then
				response.Write("<script  type=""text/javascript"" ")
				response.Write("src=""http://search.live.com/cashback/products/gleam/javascript.ashx")
				response.Write("?merchantId="& LSCB_KEY &"&type=1&bgcolor=FFFFFF&version=1.00""")
				response.Write("></script>")
			End If
		End If
		'END - Bing Cashback Gleaming
		
		'START - Visually separate prices from other information. Don't use if layout is One Column.
		if pcv_strViewPrdStyle <> "o" then
			'response.write "<div class='pcShowPrices'>"
		end if
	 
		' Display the online price if it's not zero
		if ((pPrice>Cdbl(0)) OR ((pcv_Apparel="1") AND (HaveDiffPrice=0))) and (pcv_intHideBTOPrice<>"1") then
		
			' If the List Price is not zero and higher than the online price, display striken through
			if ((pListPrice-pPrice)>0) and (pcv_intHideBTOPrice<>"1") then
				response.write "<div class='pcShowProductPrice'>"
				response.write dictLanguage.Item(Session("language")&"_viewPrd_20")

				if pcv_Apparel="1" then %>
					<span id="lprice" class='pcShowProductListPrice'><%=scCurSign & money(pListPrice)%></span>
				<% else
					response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pListPrice) & "</span>"
				end if

				response.write "</div>"
			end if
			
			if (ShowSaleIcon=1) AND (pcTargetPrice="0") then
				response.write "<div class='pcShowProductPrice'>"
				response.write dictLanguage.Item(Session("language")&"_Sale_3")
				response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pcOrgPrice) & "</span>"
				response.write "</div>"
			end if
		
			' DA Edit	
		
			' Display online price
			'response.write "<div class=""pcShowProductPrice"" itemprop=""offers"" itemscope itemtype=""http://schema.org/Offer""><span class=""pcShowProductMainPrice"">"
			'response.write dictLanguage.Item(Session("language")&"_viewPrd_3")
			response.write scCurSign & money(pPrice)
			
			' Adding Google structured data for "Availability" property
			if scdisplayStock=-1 AND pNoStock=0 AND pstock > 0 then
				response.write "<link itemprop=""availability"" href=""http://schema.org/InStock"" />"
			elseif (scShowStockLmt=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=0) OR (pserviceSpec<>0 AND scShowStockLmt=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=0) then
				response.write "<link itemprop=""availability"" href=""http://schema.org/OutOfStock"" />"
			elseif (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=1) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=1) AND clng(pcv_intShipNDays)>0 then
				response.write "<link itemprop=""availability"" href=""http://schema.org/PreOrder"" />"
			end if
			
			'response.write "<meta itemprop=""priceCurrency"" content=""USD"" />"
            %>
            <%
			if (ShowSaleIcon=1) AND (pcTargetPrice="0") then
				response.write " <span class=""pcSaleIcon""><a href=""javascript:openbrowser('sm_showdetails.asp?id=" & pcSCID & "')""><img src=""" & pcf_getImagePath("catalog",pcSCIcon) & """ title=""" &  pcSCName & """ alt=""" & pcSCName & """></a></span>"
			end if
			'response.write "</div>"
			
			' If the product is setup to use the Show Savings feature, show the savings if they exist and the customer is retail
			if ((pListPrice-pPrice)>0) AND (plistHidden<0) AND (session("customerType")<>1) and (pcv_intHideBTOPrice<>"1") then
				'response.write " - "
				response.write "<div class='pcShowProductSavings'>"
				response.write dictLanguage.Item(Session("language")&"_viewPrd_4")
			
				if pcv_Apparel="1" then %>
					<span id="psavings"><%=scCurSign & money((pListPrice-pPrice))%></span>
				<%else
					response.write scCurSign & money((pListPrice-pPrice))
				end if

				if pcv_Apparel="1" then%>
					<span id="savingspercent">(<%=round(((pListPrice-pPrice)/pListPrice)*100)%>%)</span>
				<%else
					response.write " (" & round(((pListPrice-pPrice)/pListPrice)*100) & "%)"
				end if
		
				response.write "</div>"
			end if

			' If the store is using and showing VAT, show the VAT included message and price without VAT
			if ptaxVAT="1" and ptaxdisplayVAT="1" and pnotax <> "-1" then
				if session("customerType")="1" AND ptaxwholesale="0" then
				else
					response.write "<div class='pcSmallText'>"
					response.write dictLanguage.Item(Session("language")&"_viewPrd_26") & "<br>"
					response.write dictLanguage.Item(Session("language")&"_viewPrd_27") 
					%><span id="vatspace"><%=scCurSign & money(pcf_RemoveVAT(pPrice,pIdProduct))%></span>
					<%
					response.write "</div>"
				end if
			end if
		
		end if 'this is the IF statement regarding the online price being > zero
	
		' If this is a wholesale customer and the wholesale price is > zero, display it here
		if pcv_intHideBTOPrice<>"1" then
			if session("customertype")=1 and pBtoBPrice1>0 then
				pPrice1=pBtoBPrice1
			end if
			if session("customerCategory")<>0 then
				if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=clng(session("customerCategory"))) then
					response.write "<div class='pcShowProductPriceW'>"
					response.write session("customerCategoryDesc") & " " & dictLanguage.Item(Session("language")&"_Sale_3")
					response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pcOrgPrice) & "</span>"
					response.write "</div>"
				end if
				response.write "<div class='pcShowProductPriceW'>"
				response.write session("customerCategoryDesc")&": "
			
				if pcv_Apparel="1" then %>
					<span id="wprice"><%=scCurSign & money(pPrice1)%></span>
				<%else
					response.write scCurSign & money(pPrice1)
				end if
			
				if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=clng(session("customerCategory"))) then
					response.write " <span class=""pcSaleIcon""><a href=""javascript:openbrowser('sm_showdetails.asp?id=" & pcSCID & "')""><img src=""" & pcf_getImagePath("catalog",pcSCIcon) & """ title=""" &  pcSCName & """ alt=""" & pcSCName & """></a></span>"
				end if
				response.write "</div>"
			else
				if ((pBtoBPrice1>"0") OR (pcv_Apparel="1")) and (session("customerType")=1) then 
					if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then
						response.write "<div class='pcShowProductPrice'>"
						response.write dictLanguage.Item(Session("language")&"_Sale_4")
						response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pcOrgPrice) & "</span>"
						response.write "</div>"
					end if 
					response.write "<div class='pcShowProductPrice'>"
					response.write dictLanguage.Item(Session("language")&"_viewPrd_15") &" "
				
					if pcv_Apparel="1" then %>
						<span id="wprice"><%=scCurSign & money(pBtoBPrice1)%></span>
					<%else
						response.write scCurSign & money(pBtoBPrice1)
					end if
				
					if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then
						response.write " <span class=""pcSaleIcon""><a href=""javascript:openbrowser('sm_showdetails.asp?id=" & pcSCID & "')""><img src=""" & pcf_getImagePath("catalog",pcSCIcon) & """ title=""" &  pcSCName & """ alt=""" & pcSCName & """></a></span>"
					end if
					response.write "</div>"
				end if
			end if
		end if
		
		' END - Visually separate prices from rest of product information
		if pcv_strViewPrdStyle <> "o" then
			'response.write "</div>"
		end if
		
	end if 'this is the IF statement regarding the BTO product being setup not to show prices	
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show product prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show product prices without VAT
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductPricesNoVat
	response.write scCurSign & money(pPrice/1.2)
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show product prices without VAT
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: DA Edit Bundles Lookup
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub funBundlesCalcs
	' Get stand data first
	query = "SELECT price, description FROM products WHERE idProduct=" & request.querystring("sid") &";"
	'query = query & ""
	'query = query & "" & request.querystring("sid") &";"
	set rs8=server.createobject("adodb.recordset")
	set rs8=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs8=nothing
		call closedb()
		response.redirect "tech8Err.asp?err="&pcStrCustRefID
	end if
	' If we have data	
	if NOT rs8.eof then
		pStandPrice = rs8("price")
		pStandName = rs8("description")
	end if
	set rs8=nothing
	
	query = "SELECT price, description FROM products WHERE idProduct=" & request.querystring("mid") &";"
	set rs9=server.createobject("adodb.recordset")
	set rs9=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs9=nothing
		call closedb()
		response.redirect "tech9Err.asp?err="&pcStrCustRefID
	end if
	' If we have data	
	if NOT rs9.eof then
		pMonitorPrice = rs9("price")
		pMonitorName = rs9("description")
	end if
	set rs9=nothing
	
	pMonitorNumber = 1
	
	'Work out no of screens based on id of stand
	if InStr(pStandName,"Dual") > 0 then
		pMonitorNumber = 2
	end if
	
	if InStr(pStandName,"Triple") > 0 then
		pMonitorNumber = 3
	end if
	
	if InStr(pStandName,"Quad") > 0 then
		pMonitorNumber = 4
	end if
	
	if InStr(pStandName,"Five") > 0 then
		pMonitorNumber = 5
	end if
	
	if InStr(pStandName,"Six") > 0 then
		pMonitorNumber = 6
	end if
	
	if InStr(pStandName,"Eight") > 0 then
		pMonitorNumber = 8
	end if
	
	
	funDABundlesCalcs = pStandPrice & "," & pMonitorPrice & "," & pMonitorNumber & "," & pMonitorName & "," & pStandName
	
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Bundle Calcs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: DA Edit Array Stand Lookup
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub funArrayCalcs
	' Get stand data first
	query = 		"SELECT description "
	query = query & "FROM products "
	query = query & "WHERE idProduct=" & request.querystring("sid") &";"
	set rs7=server.createobject("adodb.recordset")
	set rs7=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs7=nothing
		call closedb()
		response.redirect "tech5Err.asp?err="&pcStrCustRefID
	end if
	' If we have data	
	if NOT rs7.eof then
		pStandName = rs7("description")
	end if
	set rs7=nothing
		
	'Work out no of screens based on id of stand
	if InStr(pStandName,"Dual") > 0 then
		pMonitorNumber = 2
	end if
	
	if InStr(pStandName,"Triple") > 0 then
		pMonitorNumber = 3
	end if
	
	if InStr(pStandName,"Quad") > 0 then
		pMonitorNumber = 4
	end if
	
	if InStr(pStandName,"Five") > 0 then
		pMonitorNumber = 5
	end if
	
	if InStr(pStandName,"Six") > 0 then
		pMonitorNumber = 6
	end if
	
	if InStr(pStandName,"Eight") > 0 then
		pMonitorNumber = 8
	end if
	
	
	funDAArrayCalcs = pMonitorNumber
	
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Array Calcs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get Additional Images Array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
function pcf_GetAdditionalImages
	' Check if additional images are enabled
	If pcv_HideAdditionalImages = 0 Then
		' // SELECT DATA SET
		' TABLES: pcProductsImages
		' COLUMNS ORDER: pcProductsImages.pcProdImage_Url, pcProductsImages.pcProdImage_LargeUrl, pcProductsImages.pcProdImage_Order
		
		query = "SELECT pcProductsImages.pcProdImage_Url, pcProductsImages.pcProdImage_LargeUrl, pcProductsImages.pcProdImage_Order, pcProdImage_AltTagText "
		query = query & "FROM pcProductsImages "
		query = query & "WHERE pcProductsImages.idProduct=" & pidProduct &" "
		query = query & "ORDER BY pcProductsImages.pcProdImage_Order;"	
		set rs=server.createobject("adodb.recordset")
		set rs=conntemp.execute(query)	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "tech3Err.asp?err="&pcStrCustRefID
		end if
		
		If rs.EOF Then
			pcf_GetAdditionalImages = ""
		Else
			pcf_GetAdditionalImages = ""
			Dim xCounter '// declare a temporary counter
			xCounter = 0
			do while NOT rs.EOF
			
				pcv_strProdImage_Url = ""
				pcv_strProdImage_LargeUrl = ""
				pcv_strProdImage_Url = rs("pcProdImage_Url")
				pcv_strProdImage_LargeUrl = rs("pcProdImage_LargeUrl")
				pcv_strAltTagText = rs("pcProdImage_AltTagText")
				If pcv_strAltTagText = "" OR IsNull(pcv_strAltTagText) Then
					pcv_strAltTagText = pAltTagText
				End If
				
				if len(pcv_strProdImage_Url)>0 then
					xCounter = xCounter + 1
					if xCounter > 1 then
						pcf_GetAdditionalImages = pcf_GetAdditionalImages & "|"
					end if
					'// Add a sorted item onto the end of the string
					pcf_GetAdditionalImages = pcf_GetAdditionalImages & pcv_strProdImage_Url & "|" & pcv_strProdImage_LargeUrl & "|" & pcv_strAltTagText
				end if
	
				rs.movenext 
			loop		
		End If
		set rs=nothing
	End If
end function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Get Additional Images Array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductImage

	'// If display option is One Column (ideal for products without images), and there is no image don't show anything
	If Not (pcv_strViewPrdStyle = "o" And (len(pImageUrl) = 0 Or pImageUrl="no_image.gif")) Then  
    
        If len(pImageUrl) > 0 Then

		'// A)  The image exists
		%>
    	<div id="mainimgdiv" class="pcShowMainImage <% If pcv_IsQuickView = True Then %>QVShowMainImage <% End If %>">
				<%
					'// If this is the pop window swap out the image for the selection
					pcv_strVariableImage = pImageUrl

					Dim pcv_strZoomLink, pcv_strZoomLocation  			
	 
					pcv_strZoomLink = pcf_getImagePath(pcv_tmpNewPath&"/shop/pc/catalog",pLgimageURL)
					pcv_strZoomLocation = "onclick=""return mainImgClick(this);"""
	
					if pcv_strUseEnhancedViews = True then
						pcv_strZoomLocation = pcv_strZoomLocation & " class=""highslide"""
					end if
                %>
                
                <% if pcv_Apparel="1" then %>

                        <a href="<%=pcv_strZoomLink%>" <%=pcv_strZoomLocation%>>
                            <img id="mainimg" itemprop="image" src='<%=pcf_getImagePath(pcv_tmpNewPath & "/shop/pc/catalog",pImageUrl)%>' alt="<% Response.Write Replace(pDescription, """", "&quot;")%>"/>
                        </a>

                <% else %>

                    <% if len(trim(pLgimageURL))>0 then %>    
                                
                     <%
					'DA Edit for stand pages to display video as main image
				
						if pcv_strViewPrdStyle = "stand" then
					%>  
                    
                        <a data-toggle="lightbox" href="<%=daVimeoUrl%>" <%'=pcv_strZoomLocation%> data-title="<% Response.Write Replace(pDescription, """", "&quot;")%>">
                            <img id="img01" src="/images/ss-videos/<%=lcase(pSku)%>.jpg" alt="<% Response.Write Replace(pDescription, """", "&quot;")%>"/>
                            <% if pcv_strUseEnhancedViews = False then %>
                                <span id="zoombutton" class="pcShowAdditionalZoom"><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("zoom"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_5")%>"></span>
                            <% end if %>
                        </a>
					<%
					'Not a stand page
					else
					%> 
                                           <a data-toggle="lightbox" href="<%=pcv_strZoomLink%>" <%'=pcv_strZoomLocation%> data-title="<% Response.Write Replace(pDescription, """", "&quot;")%>">
                            <img id="img01" src='<%=pcf_getImagePath(pcv_tmpNewPath & "/shop/pc/catalog",pImageUrl)%>' alt="<% Response.Write Replace(pDescription, """", "&quot;")%>"/>
                            <% if pcv_strUseEnhancedViews = False then %>
                                <span id="zoombutton" class="pcShowAdditionalZoom"><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("zoom"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_5")%>"></span>
                            <% end if %>
                        </a>
                       <% end if %> 
                    <% else %>
				<a href="<%=pcf_getImagePath(pcv_tmpNewPath&"catalog",pImageUrl)%>" <%=pcv_strZoomLocation%>>
					<img itemprop="image" id='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pImageUrl)%>' alt="<%=pAltTagText%>" />
				</a>
                    <% end if %>

                	<% If pcv_strUseEnhancedViews = True Then %>
                    	    <div class="<%=pcv_strHighSlide_Heading %>"><%=replace(pDescription,"""","&quot;")%></div>
                	<% End If %>

                <% end if %>
                
                <% if (len(pLgimageURL)>0 and pcv_strUseEnhancedViews = False) or (pcv_Apparel="1") then %>
                    <div <% if pcv_Apparel="1" then %>id="show_10"<% end if %> style="width:100%; text-align:right; <%if (len(pLgimageURL)=0) and (pcv_Apparel="1") then%>display:none;<%end if%>">
                    </div>
                <% end if %>

                
            </div>
            
        <% Else '// If len(pImageUrl) > 0 Then %>
    
            <%
            '// B)  The image DOES NOT exist (show no_image.gif)
            %>	
                <div id="mainimgdiv" class="pcShowMainImage">
                    <img itemprop="image" id='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog","no_image.gif")%>' alt="<%=replace(pDescription,"""","&quot;")%>">
                </div>
                
                <% if (len(pLgimageURL)>0) or (pcv_Apparel="1") then %>
                    <div id="show_10" style="width:100%; text-align:right; <%if (len(pLgimageURL)=0) and (pcv_Apparel="1") then%>display:none;<%end if%>">
                        <% if InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") OR InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Firefox") then %>
                            <a id="zoombutton" href="javascript:enlrge('catalog/'+LargeImg)">
                        <% elseif InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Safari") then %>
                            <a href="javascript:open_win('catalog/'+LargeImg);">
                        <% else %>
                            <a href="javascript:open_win('catalog/'+LargeImg);" target="_blank">
                        <% end if %>
                        <img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("zoom"))%>" border="0" hspace="10">
                    </a>
                    </div>
                <% end if %>

        <%
        END IF
        %>
        
        <%
        if (pcv_Apparel="1") then
            call ColorSwatches()
        end if
        %>

    <%
    End If
	%>

	<%

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_MakeAdditionalImage

	'// Make the popup link, but dont set large image preference if the large image doesnt exist
	If len(pcv_strShowImage_LargeUrl)<1 Then
		pcv_strShowImage_LargeUrl = pcv_strShowImage_Url '// we dont have one, show the regular size
	End If
	pcv_strLargeUrlPopUp= "javascript:enlrge('" & pcf_getImagePath("catalog",pcv_strShowImage_LargeUrl)&"')" 

	'// Use Enhanced Views
	If pcv_strUseEnhancedViews = True Then
		addClass = ""
		addStyle = ""
		If bcounter = 1 Then 
			addClass = pcv_strCurrentClass
		Else
			addStyle = "cursor: pointer;"
		End If

		%>
			<a href="<%=pcf_getImagePath("/shop/pc/catalog",pcv_strShowImage_Url)%>" data-zoom="<%=pcf_getImagePath("/shop/pc/catalog",pcv_strShowImage_LargeUrl)%>" rel="lightbox"><img src='<%=pcf_getImagePath("/shop/pc/catalog",pcv_strShowImage_Url)%>' alt="<%=pcv_strShowImage_AltTagText%>" /></a>
			<% if pcv_strUseEnhancedViews = True then %>
				<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
			<% end if %>
		<%
	'// Use Pop Window 
	Else
		imageChangeEvent = ""
		imageChangeEvent = imageChangeEvent & "setMainImg('" & pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_Url) & "', '" & pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_LargeUrl) & "');"
		%>	
			<a href="<%=pcv_strLargeUrlPopUp%>"><img onClick='<%= imageChangeEvent %>' src='<%=pcf_getImagePath("catalog",pcv_strShowImage_Url)%>' alt="<%=pcv_strShowImage_AltTagText%>" /></a> 
		<% 
	End If		
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Additional Product Images (If there are any)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_AdditionalImages

if len(pImageUrl) > 0 And (Not pcv_IsQuickView = True) then ' // only if there is a main image can there be additional images.
	pcv_strAdditionalImages = pcf_GetAdditionalImages '// set variable to array of images, if there are any
	if len(pcv_strAdditionalImages)>0 then '// there is a main, are there additionals?
	%>    
	
		<%
        '// the main image to the first place in the image set
        pcv_strAdditionalImages = pImageUrl & "|" & pLgimageURL & "|" & pAltTagText & "|" & pcv_strAdditionalImages
		
        Dim pcArray_AdditionalImages '// declare a temporary array
        pcArray_AdditionalImages = Split(pcv_strAdditionalImages,"|")
        
        bCounter = 1
        
        '// When the product has additional images, this variable defines how many thumbnails are shown per row, below the main product image
        if pcv_intProdImage_Columns="" then
            pcv_intProdImage_Columns = 3
        end if
        
        modnum = pcv_intProdImage_Columns '// Get this from the db
		
		'DA Edit - Add first image as video thumbnail to stand pages
	if pcv_strViewPrdStyle = "stand" then
	%>
			<a class="act-thumb" href="/images/ss-videos/<%=lcase(pSku)%>.jpg" data-zoom="<%=daVimeoUrl%>" data-gallery="mixedgallery" rel="lightbox">	
				<img id="img_01" src='/images/ss-videos/play-thumb.jpg' alt="<% Response.Write Replace(pDescription, """", "&quot;")%>"  />		
			</a> 
    <%
	'Increase counter to avoid setting 2nd image as active
	bCounter = bCounter + 1
	end if
		
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' START Loop
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        For cCounter = LBound(pcArray_AdditionalImages) TO UBound(pcArray_AdditionalImages)

            '// Check if we have a normal image
            Dim pcv_strTempAssignment	
            pcv_strTempAssignment = ""
            pcv_strTempAssignment = pcArray_AdditionalImages(cCounter)
            pcv_strShowImage_Url = pcv_strTempAssignment '// we have one, set it
			
            '// Do Not generate an additional image if there is not one
            If len(pcv_strShowImage_Url)>0 Then
                    
				'// Check if we have a large image
				pcv_strTempAssignment = ""	
				pcv_strTempAssignment = pcArray_AdditionalImages(cCounter+1)
				pcv_strShowImage_LargeUrl = pcv_strTempAssignment '// we have one
				
				'// Check if the additional images have an alt tag text
				pcv_strTempAssignment = ""	
				pcv_strTempAssignment = pcArray_AdditionalImages(cCounter+2)
				pcv_strShowImage_AltTagText = pcv_strTempAssignment '// we have one
				
				if not bCounter mod modnum = 0 then %>
					<%pcs_MakeAdditionalImage%>
				<% Else %>
					<%pcs_MakeAdditionalImage%>
				<% end if		
				bCounter = bCounter + 1 %>
            <% End If
			
			cCounter = cCounter + 2
        
        Next
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' END Loop
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
        %>
    
    
	<%
	end if '// end if len(pcf_GetAdditionalImages)>0 then
end if	'// end if len(pImageUrl) > 0 then
%>
	<script type=text/javascript>	
		var pcv_hasAdditionalImages = <% If len(pcv_strAdditionalImages)>0 Then Response.Write "true" Else Response.Write "false" End If %>
		var pcv_strIsMojoZoomEnabled = <% If pcv_IntMojoZoom="1" And (Not pcv_IsQuickView = True And Not Session("Mobile") = "1") Then Response.Write "true" Else Response.Write "false" End If %>;
		var pcv_strMojoZoomOrientation = "<%= pcv_strMojoZoomOrientation %>";
		var pcv_strUseEnhancedViews = <% If pcv_strUseEnhancedViews Then Response.Write "true" Else Response.Write "false" End If %>;
		<% if pcv_strUseEnhancedViews = True then %>
			var CurrentImg=1;
		<% End If %>

		$pc(document).ready(function() {
			<% if pcv_strUseEnhancedViews = True then %>

				// Init Highslide
				hs.align = '<%=pcv_strHighSlide_Align%>';
				hs.transitions = [<%=pcv_strHighSlide_Effects%>];
				hs.outlineType = '<%=pcv_strHighSlide_Template%>';
				hs.fadeInOut = <%=pcv_strHighSlide_Fade%>;
				hs.dimmingOpacity = <%=pcv_strHighSlide_Dim%>;
				hs.expandCursor = null;
				
				hs.numberPosition = 'caption';
				<% if bCounter>0 then %>
						if (hs.addSlideshow) hs.addSlideshow({
							interval: <%=pcv_strHighSlide_Interval%>,
							repeat: true,
							useControls: true,
							fixedControls: false,
							overlayOptions: {
								opacity: .75,
								position: 'top center',
								hideOnMouseOut: <%=pcv_strHighSlide_Hide%>
								}
						});	
						<% end if %>

			<% end if %>
		});
		
		$pc(window).on('load', function() {
			<% If pcv_IntMojoZoom="1" And (Not pcv_IsQuickView = True And Not Session("Mobile") = "1") Then %>
				mainImgMakeZoomable("<%=pcf_getImagePath(pcv_tmpNewPath&"catalog",pLgimageURL)%>");
			<% End If %>

			$(".pcShowAdditional a").click(function(e) {
				if ($(this).hasClass('<%= pcv_strCurrentClass %>')) {
			} else {
					CurrentImg = $(this).attr("id");
					setMainImg($(this).find("img").attr("src"), $(this).attr("href"), $(this).find("img").attr("alt"));

					$(".pcShowAdditional a").removeClass('<%= pcv_strCurrentClass %>').css("cursor", "pointer");
					$(this).addClass('<%= pcv_strCurrentClass %>');

					e.preventDefault();

					$(this).css("cursor", "");
			}
			});
		});
	</script>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Additional Product Images (If there are any)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Free Shipping Text
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_NoShippingText
	if scorderlevel <> "0" then
	else
		' Check to see if the product is set as a Non-Shipping Item and display message if product is for sale
		if pnoshipping="-1" and (pFormQuantity <> "-1" or NotForSaleOverride(session("customerCategory"))=1) and pnoshippingtext="-1" then 
			response.write "<div class='pcShowProductNoShipping'>"
			response.write dictLanguage.Item(Session("language")&"_viewPrd_8")
			response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Free Shipping Text
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  CONFIGURATOR ADDON
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_BTOADDON
	if pserviceSpec<>0 then
		if Cdbl(pBtoBPrice)>0 and session("customerType")="1" then
			response.write "<input type=""hidden"" name=""GrandTotal"" value="""&scCurSign&money(pBtoBPrice1)&""">"
			response.write "<input type=""hidden"" name=""TLPriceDefault"" value="""&money(pBtoBPrice1)&""">"
			response.write "<input type=""hidden"" name=""TLPriceDefaultVP"" value="""">"
		else
			response.write "<input type=""hidden"" name=""GrandTotal"" value="""&scCurSign&money(pPrice1)&""">"
			response.write "<input type=""hidden"" name=""TLPriceDefault"" value="""">"
			response.write "<input type=""hidden"" name=""TLPriceDefaultVP"" value="""&money(pPrice1)&""">"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  CONFIGURATOR ADDON
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Out of Stock Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_OutStockMessage
	' if out of stock and show message is enabled (-1) then show message unless stock is ignored
	if (scShowStockLmt=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=0) OR (pserviceSpec<>0 AND scShowStockLmt=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=0) then
		response.write "<div class=""pcShowProductOutOfStock"">"&dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_viewPrd_7")& "</div>"
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Out of Stock Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show quantity discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_QtyDiscounts

	Dim DispSubDisc
	DispSubDisc=0
	'=0 - Link to another page that displays sub-products discounts (Default Value)
	'=1 - Display sub-products discounts directly on the View Product details page
	
	'--> check for discount per quantity
	if (pDiscountPerQuantity=0) AND (pcv_Apparel="1") then
		
		query="SELECT idDiscountperquantity FROM discountsperquantity INNER JOIN Products ON discountsperquantity.idProduct=Products.idProduct WHERE Products.removed=0 AND Products.pcProd_ParentPrd=" &pidProduct
		set rsQ=server.CreateObject("ADODB.RecordSet")
		set rsQ=conntemp.execute(query)
	
		if not rsQ.eof then
			pDiscountPerQuantity=-1
		else
			pDiscountPerQuantity=0
		end if
		set rsQ=nothing

	end if

	if pDiscountPerQuantity=-1 then
    
		'if customer is retail, check if there are discounts with retail <> 0
		VardiscGo=0
		if session("customerType")=1 then
			query="SELECT discountPerWUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerWUnit>0"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if rs.eof then
				VardiscGo=1
			end if
			set rs=nothing
		else
			query="SELECT discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerUnit>0"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if rs.eof then
				VardiscGo=1
			end if
			set rs=nothing
		end if
	
		if VardiscGo=0 then
			
            query="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" ORDER BY num"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query) 
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if NOT rs.eof then '// Quick Loop - there will not be too many discounts
				pcv_intTotalDiscounts = 0
				do until rs.eof
					pcv_intTotalDiscounts=pcv_intTotalDiscounts+1
				rs.moveNext		
				loop
				rs.moveFirst
			end if
			%>
			<div class="pcShowProductQtyDiscounts">
				<div class="pcCartLayout pcShowList container-fluid">
					<div class="pcTableHeader row">
						<div class="col-xs-8"><%=dictLanguage.Item(Session("language")&"_pricebreaks_1")%></div>
						<div class="col-xs-4"><%=dictLanguage.Item(Session("language")&"_pricebreaks_2")%>&nbsp;<a href="javascript:openbrowser('<%=pcv_tmpNewPath%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&amp;SIArray=<%=pIdProduct%>')"><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_6")%>"></a></div>
					</div>

					<% 
					pc_intCounterQ = 0
					qString1 = Cstr("")
					qString2 = Cstr("")
					do until rs.eof
						pc_intCounterQ = pc_intCounterQ + 1 '// count Discount Rows
						dblQuantityFrom=rs("quantityFrom")
						dblQuantityUntil=rs("quantityUntil")
						dblPercentage=rs("percentage")
						dblDiscountPerWUnit=rs("discountPerWUnit")
						dblDiscountPerUnit=rs("discountPerUnit")
						%>
							<% 
            
							if dblQuantityFrom=dblQuantityUntil then
								qString1 = dblQuantityUntil&"&nbsp;"&dictLanguage.Item(Session("language")&"_pricebreaks_4")
							else 
								qString1 = dblQuantityFrom&" "&dictLanguage.Item(Session("language")&"_pricebreaks_3")&" "&dblQuantityUntil&" "&dictLanguage.Item(Session("language")&"_pricebreaks_4")
							end if 
                      
							If session("customerType")=1 Then
								If dblPercentage="0" then
									qString2 = scCurSign & money(dblDiscountPerWUnit)
								else
									qString2 = dblDiscountPerWUnit & "%"
								End If
								else
								If dblPercentage="0" then
									qString2 = scCurSign & money(dblDiscountPerUnit)
								else
									qString2 = dblDiscountPerUnit & "%"
								End If
							end If
							if pc_intCounterQ = 6 then '// limit to 6 Rows
								exit do
							end if							
              
						%>
						<div class="row">
							<div class="col-xs-8"><%= qString1 %></div>
							<div class="col-xs-4"><%= qString2 %></div>
						</div>
						<%
						rs.moveNext		
					loop
					set rs=nothing
					%> 
					<% '// Display link to full chart
					if pcv_intTotalDiscounts > pc_intCounterQ then	%>
						<div class="pcTableRow">
							<div class="pcQtyDiscLink">
            		<a href="javascript:openbrowser('<%=pcv_tmpNewPath%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&amp;SIArray=<%=pIdProduct%>')"><%=dictLanguage.Item(Session("language")&"_mainIndex_9")%></a>
							</div>
						</div>
					<% end if %>
				</div>
				<div class="pcClear"></div>
			</div> 

		<% else '// if VardiscGo=0 then
        
            if pcv_Apparel="1" then

                if (DispSubDisc=0) OR (DispSubDisc="") then
                %>
                    <div class="pcShowProductQtyDiscounts">
                        <div class="pcTable pcShowList">
                            <div class="pcTableHeader">
                                <div class="pcQtyDiscQuantity"><%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg8") %></div>
                            </div>
                            <div class="pcTableRow">
                                <div class="pcQtyDiscLink">
                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>" hspace="2"><a href="javascript:openbrowser('app-subPrdDiscount.asp?idproduct=<%=pIdProduct%>');"><%response.write dictLanguage.Item(Session("language")&"_viewPrd_spmsg8a")%></a>
                                </div>
                            </div>
                        </div>
                    </div>
				<%
                else
                
					query="select idproduct,description from Products where Products.removed=0 AND pcprod_ParentPrd=" & pIDProduct
					set rsA=connTemp.execute(query)
					
					if not rsA.eof then %>
                        
                        <!--
                        <div class="pcTableRow">
                            <%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg8")%>
                        </div>
                        -->
                        
					<% end if
                    
					do while not rsA.eof
					
                        pcv_sprdID=rsA("idproduct")
					    pcv_prdName=rsA("description")
				
					    queryq="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pcv_sprdID &" ORDER BY num"
					    set rsTmp=conntemp.execute(queryq)
					    if not rsTmp.eof then
                        %>

                        <div class="pcShowProductQtyDiscounts">
                            <div class="pcTable pcShowList">
                                <div class="pcTableHeader">
                                    <div><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>" hspace="2"><%=pcv_prdName%></div>
                                </div>
                                <div class="pcTableHeader">
                                    <div class="pcQtyDiscQuantity"><%=dictLanguage.Item(Session("language")&"_pricebreaks_1")%></div>
                                    <div class="pcQtyDiscSave">
                                        <%=dictLanguage.Item(Session("language")&"_pricebreaks_2")%>
                                    </div>
                                </div>

                            <% do until rstmp.eof %>
                            
                                <div class="pcTableRow">
                                    <% if rstmp("quantityFrom")=rstmp("quantityUntil") then %>
                                        <div class="pcQtyDiscQuantity"><%=rstmp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%></div>
                                    <% else %>
                                        <div class="pcQtyDiscQuantity"><%=rstmp("quantityFrom")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_3")&"&nbsp;"&rstmp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%></div>
                                    <% end if %>
                                    <div class="pcQtyDiscSave">
                                        <% If (request.querystring("Type")="1")  or (session("CustomerType")="1") Then %>
                                            <% If rstmp("percentage")="0" then %>
                                                <%=scCurSign & money(rstmp("discountPerWUnit"))%> 
                                            <% else %>
                                                <%=rstmp("discountPerWUnit")%>%
                                            <% End If %>
                                        <% else %>
                                            <% If rstmp("percentage")="0" then %>
                                                <%=scCurSign & money(rstmp("discountPerUnit"))%> 
                                            <% else %>
                                                <%=rstmp("discountPerUnit")%>%
                                            <% End If %>
                                        <% end If %>
                                    </div>
                                </div>

                                <% 
                                rstmp.moveNext
                            loop 
                            %>
                            </div>
                            </div>
					        <%
                        end if '// if not rsTmp.eof then
					    set rstmp=nothing
					    rsA.MoveNext
					loop
					set rsA=nothing
                    
				end if '// if (DispSubDisc=0) OR (DispSubDisc="") then
                
			end if '// if pcv_Apparel="1" then


		end if '// if VardiscGo=0 then
        
	end if '// if pDiscountPerQuantity=-1 then
    
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show quantity discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: INPUT FIELDS (X)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_OptionsX
Dim i

							xrequired="0"
							xfieldCnt=0
							xfieldArrCnt=0
							reqstring="" 

							dim isArrCount,tmpCount
							isArrCount=0
							tmpCount=0

				            if tIndex<>0 then ' Check they are updating the product after adding it to the shopping cart
					            pcCartArray=session("pcCartSession")
					            tempIdOpt = ""
								tempIdxName=""
					            tempIdOpt = pcCartArray(tIndex,21)
					            if tempIdOpt = "" then
					            else
									tempIdxName=tempIdOpt
						            tempIdOpt = Split(trim(tempIdOpt),"<br>")
									tempIdxName = Split(trim(tempIdxName),"<br>")
						            xfieldArrCnt=Ubound(tempIdOpt)
						            isArrCount=xfieldArrCnt
						            if xfieldArrCnt=0 then isArrCount=1
					                for xfieldCnter = 0 to Ubound(tempIdOpt)
					                    tempIdOpt(xfieldCnter) = mid(tempIdOpt(xfieldCnter),instr(1,tempIdOpt(xfieldCnter),": ")+2)
					                    tempIdOpt(xfieldCnter) = replace(tempIdOpt(xfieldCnter),"''","'")
					                    tempIdOpt(xfieldCnter) = replace(tempIdOpt(xfieldCnter),"<BR>",vbcrlf)
										tempIdxName(xfieldCnter) = Left(tempIdxName(xfieldCnter),instr(1,tempIdxName(xfieldCnter),": ")-1)
										tempIdxName(xfieldCnter) = replace(tempIdxName(xfieldCnter),"''","'")
					                    tempIdxName(xfieldCnter) = replace(tempIdxName(xfieldCnter),"<BR>",vbcrlf)
					                next
					            end if
						
							else

								'// Only execute during a control panel session. Never run this code on the storefront
								if pcv_strAdminPrefix="1" then									
									tempIdxName=""
									tempIdxName=tempIdOpt
						            tempIdOpt = Split(trim(tempIdOpt),"|")
									tempIdxName = Split(trim(tempIdxName),"|")
						            xfieldArrCnt=Ubound(tempIdOpt)
						            isArrCount=xfieldArrCnt
						            if xfieldArrCnt=0 then isArrCount=1
					                for xfieldCnter = 0 to Ubound(tempIdOpt)
					                    tempIdOpt(xfieldCnter) = mid(tempIdOpt(xfieldCnter),instr(1,tempIdOpt(xfieldCnter),": ")+2)
					                    tempIdOpt(xfieldCnter) = replace(tempIdOpt(xfieldCnter),"''","'")
					                    tempIdOpt(xfieldCnter) = replace(tempIdOpt(xfieldCnter),"<BR>",vbcrlf)
										tempIdxName(xfieldCnter) = Left(tempIdxName(xfieldCnter),instr(1,tempIdxName(xfieldCnter),": ")-1)
										tempIdxName(xfieldCnter) = replace(tempIdxName(xfieldCnter),"''","'")
					                    tempIdxName(xfieldCnter) = replace(tempIdxName(xfieldCnter),"<BR>",vbcrlf)
					                next
									tIndex=1
								end if

					        end if

							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start pxfield Array
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							IF intXFCount>=0 THEN
							For i=0 to intXFCount
								'select from the database more info 
								query= "SELECT xfield,textarea,widthoffield,rowlength,maxlength FROM xfields WHERE idxfield="&pcXFArr(0,i)
								set rs=server.createobject("adodb.recordset")
								set rs=conntemp.execute(query)

								if not rs.EOF then '// Check for no field in DB, although the field is referenced by the product
									xField=rs("xfield")
									TextArea=rs("textarea")
									widthoffield=rs("widthoffield")
									rowlength=rs("rowlength")
									maxlength=rs("maxlength")
									set rs=nothing
									tmpCount=tmpCount+1
									pxreq=pcXFArr(1,i)
									
									if pxreq="-1" then
										xfieldCnt=xfieldCnt+1
										xrequired="1"
										if reqstring<>"" then
											reqstring=reqstring & ","
										end if
										reqstring=reqstring&"additem.xfield" & tmpCount & ".value,'"&replace(xfield,"'","\'")&"'"
									end if
									
									XValue=""
									if tIndex<>0 and (xfieldArrCnt > 0 or isArrCount>0) then
										for xfieldCnter = 0 to Ubound(tempIdOpt)
											if tempIdxName(xfieldCnter)=xField then
					                    		XValue=tempIdOpt(xfieldCnter)
												exit for
											end if
					                	next
									end if
									%>
									
                                    <input type="hidden" name="xf<%=tmpCount%>" value="<%=pcXFArr(0,i)%>">
                                    <label for="xfield<%=tmpCount%>"><%=xField%></label>
                                    
                                    <% if TextArea="-1" then %>
                                       
                                       
                                        <textarea name="xfield<%=tmpCount%>" class="form-control" cols="<%=widthoffield%>" rows="<%=rowlength%>" style="margin-top: 6px" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>><%if tIndex<>0 and (xfieldArrCnt > 0 or isArrCount>0) then response.write XValue end if %></textarea>
                                        <% if maxlength>"0" then %>
                                            <span class="help-block">
                                                <%=dictLanguage.Item(Session("language")&"_GiftWrap_5a")%>
                                                <span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> 
                                                <%=dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
                                            </span>
                                        <% end if %>
                                        
                                    <% else %>
                                       
                                        <input type="text" class="form-control" name="xfield<%=tmpCount%>" size="<%=widthoffield%>" maxlength="<%=maxlength%>" style="margin-top: 6px" <%if tIndex<>0 and (xfieldArrCnt > 0 or isArrCount>0) then%> value="<% response.write XValue %>" <%end if %> <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>>
                                        <% if maxlength>"0" then %>                                        
                                            <span class="help-block">
                                                <%=dictLanguage.Item(Session("language")&"_GiftWrap_5a")%>
                                                <span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> 
                                                <%=dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
                                            </span>
                                        <% end if %>
                                        
                                    <% end if %>
							
								<% 
								end if ' rs.eof
							Next
							END IF
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// End pxfield Array
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
					%>
					<input type="hidden" name="XFCount" value="<%=tmpCount%>" />		
					<%if tmpCount>0 then%>
					<script type=text/javascript>
					function testchars(tmpfield,idx,maxlen)
					{
						var tmp1=tmpfield.value;
						if (tmp1.length>maxlen)
						{
							alert("<%response.write dictLanguage.Item(Session("language")&"_CheckTextField_1")%>" + maxlen + "<%response.write dictLanguage.Item(Session("language")&"_CheckTextField_1a")%>");
							tmp1=tmp1.substr(0,maxlen);
							tmpfield.value=tmp1;
							document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
							tmpfield.focus();
						}
						document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
					}
					</script>
					<%end if		
End Sub

Public Sub pcs_OptionsXTab
Dim i

							xrequired="0"
							xfieldCnt=0
							reqstring="" 

							dim tmpCount
							tmpCount=0

							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start pxfield Array
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							IF intXFCount>=0 THEN
							For i=0 to intXFCount
								'select from the database more info 
								query= "SELECT xfield,textarea,widthoffield,rowlength,maxlength FROM xfields WHERE idxfield="&pcXFArr(0,i)
								set rs=server.createobject("adodb.recordset")
								set rs=conntemp.execute(query)

								if not rs.EOF then '// Check for no field in DB, although the field is referenced by the product
									xField=rs("xfield")
									TextArea=rs("textarea")
									widthoffield=rs("widthoffield")
									rowlength=rs("rowlength")
									maxlength=rs("maxlength")
									set rs=nothing
									tmpCount=tmpCount+1
									pxreq=pcXFArr(1,i)
									
									if pxreq="-1" then
										xfieldCnt=xfieldCnt+1
										xrequired="1"
										if reqstring<>"" then
											reqstring=reqstring & ","
										end if
										reqstring=reqstring&"additem.xfield" & tmpCount & ".value,'"&replace(xfield,"'","\'")&"'"
									end if
								end if ' rs.eof
							Next
							END IF
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// End pxfield Array
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: INPUT FIELDS (X)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show WishList
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_WishList
	if scWL=-1 then
		if pserviceSpec=0 then 
		%>
			<div class="pcShowWishlist">
			<% if pcv_Apparel="1" then
				pcv_strWishListLink =  "location='Custwl.asp?OptionGroupCount=0&idproduct='+document.additem.idproduct.value+'"
			else
		
				'// Form the link that gets attached to the wishlist button
				pcv_strWishListLink =  "location='Custwl.asp?OptionGroupCount="&pcv_intOptionGroupCount&"&idproduct="&pIdProduct
				Dim bCounter
				Do until bCounter = pcv_intOptionGroupCount
					bCounter = bCounter + 1
					pcv_strWishListLink = pcv_strWishListLink & "&idOption"&bCounter&"='+"
					'if optionN <> "" then
					pcv_strWishListLink = pcv_strWishListLink & "document.additem.idOption"&bCounter&".value+'"
					'end if	
				Loop

			end if
			pcv_strWishListLink = pcv_strWishListLink & "';"
			pcv_strFuntionCall = "cdDynamic"
			if xOtionrequired = "1" then '// If there are any required options at all.
			'// figure some stuff out
			%>
				<a class="pcButton pcButtonAddToWishlist" href="javascript: if (checkproqty(document.additem.quantity)) { if (<%=pcv_strFuntionCall%>(<%=pcv_strReqOptString%>,1)) {<%=Server.HtmlEncode(pcv_strWishListLink)%>;}};">
        	<img src="<%=pcf_getImagePath("",rslayout("addtowl"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_addtowl")%>">
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtowl") %></span>
        </a>
			<% else %>
				<a class="pcButton pcButtonAddToWishlist" href="javascript:<%=Server.HtmlEncode(pcv_strWishListLink)%>">
        	<img src="<%=pcf_getImagePath("",rslayout("addtowl"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_addtowl")%>">
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtowl") %></span>
        </a>
			<%end if %>
			</div>
		<%
		end if 
	end if 
    %>
    <!--#include file="inc_addPinterest.asp"-->
    <%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show WishList
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Customize Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CustomizeButton
Dim rsQ,queryQ
	queryQ="SELECT TOP 1 configProduct FROM configSpec_products WHERE specProduct=" & pIdProduct & ";"
	set rsQ=connTemp.execute(queryQ)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsQ=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rsQ.eof then
		showCustomize=1
        
        If pcf_BTOisConfig OR scConfigPurchaseOnly=1 Then
			if pcv_lngMinimumQty <> 0 then
				pcv_strConfigClickEvent = "javascript:parent.location='configurePrd.asp?idproduct=" & pIdProduct & "&qty=" & pcv_lngMinimumQty & "'; return false;"
			else
				pcv_strConfigClickEvent = "javascript:parent.location='configurePrd.asp?idproduct=" & pIdProduct & "&qty='+document.additem.quantity.value; return false;"
			end if
        Else
            pcv_strConfigClickEvent = "javascript:document.additem.action='configurePrd.asp?idproduct=" & pIdProduct & "&qty='+document.additem.quantity.value;"
        End If
        %>
		<button class="pcButton pcButtonCustomize" onclick="<%=pcv_strConfigClickEvent%>">        
			<img src="<%=pcf_getImagePath("",rslayout("customize"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_customize")%>" />
			<span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_customize")%></span>
		</button>
        <input class="form-control form-control-inline" type="hidden" name="quantity" value="1">
	<%end if
	set rsQ=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Customize Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Required Cross Selling Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim cs_imageheight, cs_imagewidth, cs_ViewCnt, pcv_strHaveResults, pcv_intProductCount, pcArray_CSRelations, pcv_intCategoryActive, pcv_intAccessoryActive, cs_showNFS

Public Sub pcs_RequiredCrossSelling

    xCSCnt = 0
    pcv_strCSString = ""
    pcv_strReqCSString = ""
    cs_RequiredIds = ""
    pcv_strPrdDiscounts = ""
    pcv_strCSDiscounts = ""
	
	Dim pcv_strGetSitewide
	Dim pcv_strIsBundleActiveFlag

	'// Get Cross Sell Settings - Product Level
	query= "SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,cs_ImageHeight,cs_ImageWidth,crossSellText,cs_ProductViewCnt,cs_showNFS, csw_status FROM crossSelldata WHERE id=" & pIdProduct
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	pcv_strGetSitewide=0
	if NOT rs.eof then
		scCS=rs("cs_status")
		cs_showprod=rs("cs_showprod")
		cs_showcart=rs("cs_showcart")
		cs_showimage=rs("cs_showimage")
		cs_imageheight=rs("cs_imageheight")
		cs_imagewidth=rs("cs_imagewidth")
		crossSellText=rs("crossSellText")
		cs_ViewCnt=rs("cs_ProductViewCnt")
		cs_showNFS=rs("cs_showNFS")
		csw_status=rs("csw_status")
	else
		pcv_strGetSitewide=1			
	end if	
	set rs=nothing	
    
    
	
	'// Get Cross Sell Settings - Sitewide 
	If pcv_strGetSitewide=1 Then			
		query= "SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,cs_ImageHeight,cs_ImageWidth,crossSellText,cs_ProductViewCnt,cs_showNFS FROM crossSelldata WHERE id=1;"
		set rs=server.createobject("adodb.recordset")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if NOT rs.eof then
			scCS=rs("cs_status")
			cs_showprod=rs("cs_showprod")
			cs_showcart=rs("cs_showcart")
			cs_showimage=rs("cs_showimage")
			cs_imageheight=rs("cs_imageheight")
			cs_imagewidth=rs("cs_imagewidth")
			crossSellText=rs("crossSellText")
			cs_ViewCnt=rs("cs_ProductViewCnt")
			cs_showNFS=rs("cs_showNFS")
		end if
		set rs=nothing
	 End If 
     
     If IsNull(cs_showNFS) Or (cs_showNFS="") Then
        cs_showNFS = -1
     End If

	If scCS=-1 AND cs_showProd="-1" Then		
	
		If cs_ViewCnt < 1 then
			cs_ViewCnt = 2
		End if			
		
		query="SELECT cs_relationships.idproduct, cs_relationships.idrelation, cs_relationships.cs_type, cs_relationships.discount, cs_relationships.ispercent,cs_relationships.isRequired, products.servicespec, products.price, products.description, products.bToBprice, products.serviceSpec, products.noprices FROM cs_relationships INNER JOIN products ON cs_relationships.idrelation=products.idProduct WHERE (((cs_relationships.idproduct)="&pidproduct&") AND ((products.active)=-1) AND ((products.removed)=0)) ORDER BY cs_relationships.num,cs_relationships.idrelation;"		
		set rs=server.createobject("adodb.recordset")
		set rs=conntemp.execute(query)	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		pcv_strHaveResults=0
		if NOT rs.eof then
			pcArray_CSRelations = rs.getRows()
			pcv_intProductCount = UBound(pcArray_CSRelations,2)+1
			pcv_strHaveResults=1
		end if
		set rs=nothing		
		
		tCnt=Cint(0)	
		
		if pcv_strHaveResults=1 then
			do while (tCnt < pcv_intProductCount)					

				pidrelation=pcArray_CSRelations(1,tCnt) '// rs("idrelation")
				pcsType=pcArray_CSRelations(2,tCnt) '// rs("cs_type")			
				pDiscount=pcArray_CSRelations(3,tCnt) '// rs("discount")
				pIsPercent=pcArray_CSRelations(4,tCnt) '// rs("isPercent")
				pcv_strIsRequired=pcArray_CSRelations(5,tCnt) '// rs("isRequired")
				cs_pserviceSpec=pcArray_CSRelations(6,tCnt) '// rs("servicespec")
				
				ppPrice=pcArray_CSRelations(7,tCnt) '// rs("price")
				
				if pcArray_CSRelations(9,tCnt)>"0" then
					ppBPrice=pcArray_CSRelations(9,tCnt)
				else
					ppBPrice=ppPrice
				end if
				
				cs_pserviceSpec=pcArray_CSRelations(10,tCnt)
				if cs_pserviceSpec="" OR IsNull(cs_pserviceSpec) then
					cs_pserviceSpec=0
				end if
				cs_pnoprices=pcArray_CSRelations(11,tCnt)
				if cs_pnoprices="" OR IsNull(cs_pnoprices) then
					cs_pnoprices=0
				end if
				
				pPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
				pBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
				if session("customertype")=1 and pBtoBPrice1>0 then
					pPrice1=pBtoBPrice1
				end if
				
				tmp_pidProduct=pidProduct
				tmp_pPrice=pPrice
				tmp_pPrice1=pPrice1
				tmp_pBtoBPrice=pBtoBPrice
				tmp_pBtoBPrice1=pBtoBPrice1
				tmp_pnoprices=pnoprices
				tmp_pserviceSpec=pserviceSpec
				
				pidProduct=pidrelation
				pPrice=ppPrice
				pBtoBPrice=ppBPrice
				pnoprices=cs_pnoprices
				pserviceSpec=cs_pserviceSpec
				
				ppPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
				ppBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
				if session("customertype")=1 and ppBtoBPrice1>0 then
					ppPrice1=ppBtoBPrice1
				end if
				
				pidProduct=tmp_pidProduct
				pPrice=tmp_pPrice
				pPrice1=tmp_pPrice1
				pBtoBPrice=tmp_pBtoBPrice
				pBtoBPrice1=tmp_pBtoBPrice1
				pnoprices=tmp_pnoprices
				pserviceSpec=tmp_pserviceSpec
				
				tCnt=tCnt+1
				
				'Store ALL Ids
				pcv_strCSString = pcv_strCSString & pidrelation & ","
				xCSCnt = xCSCnt + 1
				
				if pIsPercent<>0 then
					pcv_strPrdDiscounts = pcv_strPrdDiscounts & CDbl(pPrice1*(pDiscount/100)) & ","
					pcv_strCSDiscounts = pcv_strCSDiscounts & CDbl(ppPrice1*(pDiscount/100)) & ","
				else
					pcv_strCSDiscounts = pcv_strCSDiscounts & CDbl(pDiscount) & ","
					pcv_strPrdDiscounts = pcv_strPrdDiscounts & "0,"
				end if
				
				cs_RequiredIds = cs_RequiredIds & pcv_strIsRequired & ","
				
				'// Clear Variables
				cs_pserviceSpec=""
				pidrelation=""
				pcsType=""
			
			loop		
			
			
			if len(pcv_strCSString) > 0 then
				pcv_strCSString = left(pcv_strCSString,len(pcv_strCSString)-1)
			end if
			if len(pcv_strPrdDiscounts) > 0 then
				pcv_strPrdDiscounts = left(pcv_strPrdDiscounts,len(pcv_strPrdDiscounts)-1)
			end if
			if len(pcv_strCSDiscounts) > 0 then
				pcv_strCSDiscounts = left(pcv_strCSDiscounts,len(pcv_strCSDiscounts)-1)
			end if
			if len(cs_RequiredIds) > 0 then
				cs_RequiredIds = left(cs_RequiredIds,len(cs_RequiredIds)-1)
			end if
            %>
            <input name="pCSCount" type="hidden" value="<%=xCSCnt%>">
            <input name="pCrossSellIDs" type="hidden" value="<%=pcv_strCSString%>">
            <input name="pPrdDiscounts" type="hidden" value="<%=pcv_strPrdDiscounts%>">
            <input name="pCSDiscounts" type="hidden" value="<%=pcv_strCSDiscounts%>">
            <input name="pRequiredIDs" type="hidden" value="<%=cs_RequiredIds%>">
            <%	
		end if
	
	End if '// If cint(cs_pOptCnt) <> cint(cs_pCnt) Then
%>	
<!--#include file="doCustomCrossSelling.asp"-->
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Required Cross Selling Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Cross Selling With Discounts (Bundles)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CrossSellingDiscounts
	
	'// Only run when having "Add to Cart" button
	IF scCS=-1 AND cs_showProd="-1" THEN 
								
		dim cs_count,cs_pCnt, cs_pOptCnt,cs_pAddtoCart		
		
		tCnt=Cint(0)	
		if pcv_strHaveResults=1 then
			
			cs_pCnt=Cint(0)
			cs_pOptCnt=Cint(0)
			cs_pAddtoCart=Cint(0)
			pcv_intCategoryActive=2	'// set bundle group to inactive
			pcv_intAccessoryActive=2 '// set accessories group to inactive
			cs_count=Cint(0)
			session("listcross")=""
			
			do while ( (tCnt < pcv_intProductCount) AND (tCnt < cs_ViewCnt))				
				
				pidrelation=pcArray_CSRelations(1,tCnt) '// rs("idrelation")
				pcsType=pcArray_CSRelations(2,tCnt) '// rs("cs_type")			
				pDiscount=pcArray_CSRelations(3,tCnt) '// rs("discount")
				cs_pserviceSpec=pcArray_CSRelations(6,tCnt)				
				pcArray_CSRelations(8,tCnt) = 1

				'// Only show when the product is not an Apparel product
				If (pcsType="Accessory") OR ((pcv_Apparel<>"1") AND ((pcsType="Bundle") AND (pDiscount>0))) Then

					'// CHECK IF BUNDLES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY		
					'// CHECK IF ACCESSORIES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY  						
					If Session("customerType")=1 Then
						pcv_strCSTemp=""
					else
						pcv_strCSTemp=" AND pccats_RetailHide<>1 "
					end if									
					query="SELECT categories_products.idProduct "
					query=query+"FROM categories_products " 
					query=query+"INNER JOIN categories "
					query=query+"ON categories_products.idCategory = categories.idCategory "
					query=query+"WHERE categories_products.idProduct="& pidrelation &" AND iBTOhide=0 " & pcv_strCSTemp & " "
					query=query+"ORDER BY priority, categoryDesc ASC;"	
					set rsCheckCategory=server.CreateObject("ADODB.RecordSet")
					set rsCheckCategory=conntemp.execute(query)									
					If NOT rsCheckCategory.eof Then
						If pcsType="Accessory" Then
							pcv_intAccessoryActive=1
						End If
						If pcsType="Bundle" Then							
							pcv_intCategoryActive=1
						End If	
					Else
						session("listcross")=session("listcross") & "," & pidrelation					
					End If	
					set rsCheckCategory=nothing
					
				End If '// If (pcsType="Bundle") AND (pDiscount>0) Then	

				pcv_intOptionsExist=0
				
				'// CHECK FOR REQUIRED OPTIONS							
				pcv_intOptionsExist=pcf_CheckForReqOptions(pidrelation) '// check options function (1=YES, 2=NO)			


				'// CHECK FOR REQUIRED INPUT FIELDS
				if pcv_intOptionsExist=2 then
					pcv_intOptionsExist=pcf_CheckForReqInputFields(pidrelation)
				end if				


				'// VALIDATE
				if (cs_pserviceSpec=true) OR (pcv_intOptionsExist = 1) then
					If pcsType<>"Accessory" Then
						cs_pOptCnt=cs_pOptCnt+1
					End If
					pcArray_CSRelations(8,tCnt) = 0					
				End If	
				If pcsType<>"Accessory" Then
					cs_pCnt=cs_pCnt+1 
				End If
				tCnt=tCnt+1				
			loop
		
		end if '// if pcv_strHaveResults=1 then		

					
		'// If ALL items are either BTO or have options or inactive, do not show items
		if (cint(cs_pOptCnt) <> cint(cs_pCnt)) AND (pcv_intCategoryActive=1) then
			
			cs_DisplayCheckBox=-1
			cs_Bundle=-1
			
			%>
      <div class="pcSectionTitle">
        <%=dictLanguage.Item(Session("language")&"_viewPrd_cs1")%>
      </div>
      <div class="pcSectionContents">
	      <%=dictLanguage.Item(Session("language")&"_viewPrd_cs2")%>
      </div>
      <% if cs_showImage="-1" then %>
        <!--#include file="cs_img.asp"-->
      <% else %>
        <!--#include file="cs.asp"-->
      <% end if %>
			<div class="pcSpacer"></div>
		<% end if
	
	END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Cross Selling With Discounts (Bundles)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Cross Selling With Accessories
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CrossSellingAccessories
	
	'// Only run when having "Add to Cart" button
	IF scCS=-1 AND cs_showProd="-1" THEN 
								
		dim cs_pAddtoCart
		
		if pcv_strHaveResults=1 then

			cs_DisplayCheckBox=-1
			cs_Bundle=0
			
			if pcv_intAccessoryActive=1 then %>
				<div class="pcSectionTitle">
						<%=crossSellText%>
				</div>
				<div class="pcSectionContents">
					<%
						if showAddtoCart=1 then
							response.write dictLanguage.Item(Session("language")&"_viewPrd_cs3")
						%>&nbsp;(<img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("requiredicon"))%>">)<%=dictLanguage.Item(Session("language")&"_viewPrd_cs4")%>
						<%
						else
							if not (pserviceSpec<>0 AND showCustomize=1) then
								response.write pDescription & dictLanguage.Item(Session("language")&"_viewPrd_cs10")
							end if
						end if
					%>
				</div>
			
				<div class="pcShowCrossSellProducts">
					<% if cs_showImage="-1" then %>
						<!--#include file="cs_img.asp"-->
					<% else %>
						<!--#include file="cs.asp"-->
					<% end if %>
				</div>
        <div class="pcSpacer"></div>
			<% end if
		end if
	END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Cross Selling With Accessories
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Long Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_LongProductDescription 
    '// Display long product description if it isn't empty
	'If len(pDetails)>0 Then 
    %>
    <%=pDetails%>
	<% 
    'End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Long Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>



<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Options (N)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_OptionsN

	pcv_TotalOpts=0
	' SELECT DATA SET
	' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
	query = "SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
	query = query & "FROM products "
	query = query & "INNER JOIN ( "
	query = query & "pcProductsOptions INNER JOIN ( "
	query = query & "optionsgroups "
	query = query & "INNER JOIN options_optionsGroups "
	query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
	query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
	query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
	query = query & "WHERE products.idProduct=" & pidProduct &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order, optionsGroups.OptionGroupDesc;"
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	' If we have data	
	if NOT rs.eof then %>
    <div id="spec-options" class=" wow fadeInLeft" data-wow-delay="0.1s">
			<% 
				pcv_intOptionGroupCount = 0 '// keeps count of the number of options
				xOptionsCnt = 0 '// keeps count of the number of required options
				do until rs.eof	

          'if pcv_intOptionGroupCount <= 5  then ' // start limit to 5 options
          '// Get the Group Name
          pcv_strOptionGroupDesc=rs("OptionGroupDesc")
          '// Get the Group ID
          pcv_strOptionGroupID=rs("idOptionGroup")
          '// Is it required
          pcv_strOptionRequired=rs("pcProdOpt_Required")			
      
          '// Start: Do Option Count
          pcv_intOptionGroupCount = pcv_intOptionGroupCount + 1 
          '// End: Do Option Count
          
          '// Get the number of the Option Group
          pcv_strOptionGroupCount = pcv_intOptionGroupCount
          
          '// Start: Do Required Option Count AND generate validation string
          if IsNull(pcv_strOptionRequired) OR pcv_strOptionRequired="" then
              pcv_strOptionRequired=0 '// not required // else it is "1"
          end if			
          if pcv_strOptionRequired=1 then
            
            ' Keep Tally
            xOptionsCnt = xOptionsCnt + 1
            
            ' Generate String
            if xOtionrequired="1" then
              pcv_strReqOptString = pcv_strReqOptString & ","
            end if
          
            xOtionrequired="1"
            pcv_strOptionGroupDesc2=pcv_strOptionGroupDesc
            pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"'","")
            pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"""","\'\'")
            pcv_strReqOptString = pcv_strReqOptString & "document.additem.idOption" & pcv_strOptionGroupCount & ".selectedIndex,'"& pcv_strOptionGroupDesc2 &"'"
          
          end if
          '// End: Do Required Option Count
      
          '// Make the Option Box
          If pcv_IsQuickView = True Then
            pcs_makeOptionBoxQV	
          Else
            pcs_makeOptionBox
          End If						
      
          'end if ' // end limit to 5 options
      rs.movenext
    loop	
    %>
	<% end if
	set rs=nothing
    tmpQVOptions=tmpQVOptions & "<input type=""hidden"" name=""OptionGroupCount"" value=""" & pcv_intOptionGroupCount & """>"
	%>
    <input type="hidden" name="OptionGroupCount" value="<%=pcv_intOptionGroupCount%>">
									</div>
								</div>
<%
If statusAPP="1" Then
	if pcv_TotalOpts>0 then
	    call CreateStockMsgArea()
	end if
End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Options (N)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Options Box
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_makeOptionBox
	' SELECT DATA SET
	' TABLES: options_optionsGroups, options
	query = 		"SELECT options_optionsGroups.InActive, options_optionsGroups.price, options_optionsGroups.Wprice, "
	query = query & "options_optionsGroups.idoptoptgrp, options_optionsGroups.sortOrder, options.idoption, options.optiondescrip "
	query = query & "FROM options_optionsGroups "
	query = query & "INNER JOIN options "
	query = query & "ON options_optionsGroups.idOption = options.idOption "
	query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_strOptionGroupID &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	'query = query & "ORDER BY options_optionsGroups.sortOrder, options.optiondescrip;"	
	query = query & "ORDER BY options_optionsGroups.sortOrder, options_optionsGroups.price, options.optiondescrip;"	
	set rs2=server.createobject("adodb.recordset")
	set rs2=conntemp.execute(query)	
	'if err.number<>0 then
		'call LogErrorToDatabase()
		'set rs2=nothing
		'call closedb()
		'response.redirect "techErr.asp?err="&pcStrCustRefID
	'end if
	
	' If we have data
	if NOT rs2.eof then
	
	'DA Edit to try and hide DisplayPort Row
	if pcv_strOptionGroupDesc = "DisplayPort Adapters" then
		strDPRowID = " id=""trGA"""
	else
		strDPRowID = ""
	end if
	%>
    
    
<div<%=strDPRowID%> class="row specb-row"><!-- Row Start -->
<div class="col-sm-3 specb-title">
<label>
                <%
		'// clean up the option group description
		if pcv_strOptionGroupDesc<>"" then
			pcv_strOptionGroupDesc=replace(pcv_strOptionGroupDesc,"""","&quot;")
		end if 
		
		'// START SELECT
		pcv_isOptionSelected="" '// Is this option box selected? Fill variable to "1" during the following loop.

		'DA Edit - to add an extra JS function to the onchange for graphics cards only
		daJSFunction = ""
		
		Select Case pcv_strOptionGroupDesc 
			Case "MS Office" 
				Response.Write "Microsoft Office:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Microsoft Office"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-office2013.htm"">Learn More</a></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Security Software:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Microsoft Windows Defender AntiVirus</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Bonus Software:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Display Fusion - (Multi Screen Management Tools)</p><a data-toggle=""lightbox"" data-title=""Display Fusion Software"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-displayfusion.htm"">Learn More</a>"
			Case "CPU" 
				Response.Write "CPU / Processor:"
				if not InStr(pSku, "EXT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""CPU / Processor"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-cpu-ext.htm"">Learn More</a></div></div><!-- Row end -->"
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Motherboard:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optMotherboard""></span></p><a data-toggle=""lightbox"" data-title=""Motherboard"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-mboard-x79.htm"">Learn More</a>"
					daJSFunction = "reCalcColourEXT();"
				end if 
				if not InStr(pSku, "ULT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""CPU / Processor"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-cpu-ultra.htm"">Learn More</a></div></div><!-- Row end -->"
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Motherboard:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optMotherboard""></span></p><a data-toggle=""lightbox"" data-title=""Motherboard"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-mboard-z170.htm"">Learn More</a>"
				end if 
				if not InStr(pSku, "PRO1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""CPU / Processor"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-cpu-pro.htm"">Learn More</a></div></div><!-- Row end -->"
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Motherboard:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optMotherboard""></span></p><a data-toggle=""lightbox"" data-title=""Motherboard"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-mboard-z170.htm"">Learn More</a>"
				end if 
			Case "Keyb. / Mouse"
				Response.Write "Keyboard &amp; Mouse:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Keyboard &amp; Mouse"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-keyboards.htm"">Learn More</a>"
			Case "Speakers"
				Response.Write "Speakers:"
				strEndQM = ""
			Case "2nd Hard Drive"
				Response.Write "2nd Hard Drive:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""2nd (Storage) Hard Drive"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-storage-drive.htm"">Learn More</a>"
			Case "OS"
				Response.Write "Operating System:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Microsoft Windows Versions"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-os.htm"">Learn More</a>"
			Case "Power Leads"
				Response.Write "Power Leads:"
				strEndQM = ""
			Case "RAM"
				Response.Write "System RAM:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""RAM / Memory"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-ram.htm"">Learn More</a>"
			Case "Wireless Card"
				Response.Write "Wireless Network Card:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Wireless Network Card (WiFi)"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-wifi.htm"">Learn More</a></div></div><!-- Row end -->"
				if not InStr(pSku, "PRO1") = 0 Then
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Fractal Design Core 1100 (W:175mm, H:355mm, D:420mm)</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case Cooling:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCaseCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Power Supply:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optPSU""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>CPU Cooler:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCPUCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>USB Ports:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">4 x USB 3 (4 Rear) 4 x USB 2 (2 Rear & 2 Front)</p>"
				end if 
				if not InStr(pSku, "ULT1") = 0 Then
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Antec P7 (W:210mm, H:470mm, D:445mm)</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case Cooling:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCaseCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Power Supply:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optPSU""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>CPU Cooler:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCPUCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>USB Ports:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">4 x USB 3 (2 Front & 2 Rear) 4 x USB 2 (4 Rear)</p>"
				end if 
				if not InStr(pSku, "EXT1") = 0 Then
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Corsair 3000D Case (W:230mm, H:466mm, D:462mm)</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case Cooling:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCaseCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Power Supply:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optPSU""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>CPU Cooler:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCPUCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>USB Ports:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">4 x USB 3 (2 Front & 2 Rear) 8 x USB 2 (8 Rear)</p>"
				end if 
			Case "Optical Drive" 
				Response.Write "Optical Drive:"
				strEndQM = "</div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Network / LAN Port:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Integrated Gigabit Ethernet LAN Adapter</p>"
			Case "Boot Hard Drive"
				Response.Write "Boot Hard Drive:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Boot Hard Drive"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-boot-ssd.htm"">Learn More</a>"
			Case "Warranty Cover"
				Response.Write "Warranty Cover:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""On-Site Warranty Cover"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-warranty.htm"">Learn More</a>"
			Case "DisplayPort Adapters"
				Response.Write "DisplayPort Adapters:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Card DisplayPort Adapters"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-adapters.htm"">Learn More</a></div></div><!-- Row end -->"
				strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Sound Card:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Integrated 6 Channel HD Audio Sound Card</p><a data-toggle=""lightbox"" data-title=""Sound Card / Speakers"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-sound.htm"">Learn More</a>"
			Case "Graphics Cards"
				Response.Write "Graphics Card Setup:"
				if not InStr(pSku, "PRO1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Cards Setup Options"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-graphics-210.htm"">Learn More</a>"	
					daJSFunction = "reCalcColour();"				
				end if 
				if not InStr(pSku, "ULT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Cards Setup Options"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-graphics.htm"">Learn More</a>"
					daJSFunction = "reCalcColour();"
				end if 
				if not InStr(pSku, "EXT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Cards Setup Options"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-graphics.htm"">Learn More</a>"
					daJSFunction = ""
				end if 
			Case "Backup System"
				Response.Write "Backup System:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Bootable Backup Hard Drive"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-backup.htm"">Learn More</a>"
			Case Else
				Response.Write pcv_strOptionGroupDesc
				strEndQM = ""
		End Select
		%>
		</label>
        </div>
        <div class="col-sm-9 specb-field">
		<div class="specb-dd">
        <select id="idOption<%=pcv_strOptionGroupCount%>" name="idOption<%=pcv_strOptionGroupCount%>" class="spec-dd" onchange="reCalc();<%=daJSFunction%>">
		<%
			'Attempt to stop 2nd free options saying included
			icount = 0
		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			do until rs2.eof			
			
			OptInActive=rs2("InActive") ' Is it active?
			if IsNull(OptInActive) OR OptInActive="" then
				OptInActive="0"
			end if
			
			dblOptPrice=rs2("price") '// Price
			dblOptWPrice=rs2("Wprice") '// WPrice
			intIdOptOptGrp=rs2("idoptoptgrp") '// The Id of the Option Group
			intIdOption=rs2("idoption") '// The Id of the Option
			strOptionDescrip=rs2("optiondescrip") '// A description of the Option
			strOptSortOrder=rs2("sortOrder") '// Sort order set in Admin
	
			'**************************************************************************************************
			' START: Dispay the Options
			'**************************************************************************************************
			if OptInActive="0" then
				If session("customerType")=1 then 
					optPrice=dblOptWPrice
				Else
					optPrice=dblOptPrice
				End If 
				
				'' DA Edit to ignore wifi card option for bundles
				if intIdOption = 148 then
				%>
                <%
				elseif intIdOption = 149 then
				%>
                <%
				elseif intIdOption = 198 then
				%>
                <%
				else
				
				'DA Edit to insert dotted line between GPU Options
				if intIdOption = 256 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Traditional Hard Drives:</option>")
				elseif intIdOption = 254 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Fast &amp; Silent SSDs:</option>")
				elseif intIdOption = 262 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel 10th Generation CPUs:</option>")
				elseif intIdOption = 318 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>AMD Ryzen CPUs:</option>")
				elseif intIdOption = 273 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Dual Monitor Capable:</option>")
				elseif intIdOption = 297 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Quad Monitor Capable:</option>")
				elseif intIdOption = 298 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Six Monitor Capable:</option>")
				elseif intIdOption = 299 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Eight Monitor Capable:</option>")
				elseif intIdOption = 300 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 3 Monitor Capable:</option>")
				elseif intIdOption = 333 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 4 Monitor Capable:</option>")
				elseif intIdOption = 361 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 4 Monitor Capable:</option>")
				elseif intIdOption = 335 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 6 Monitor Capable:</option>")
				elseif intIdOption = 334 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 8 Monitor Capable:</option>")
				elseif intIdOption = 362 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 8 Monitor Capable:</option>")
				elseif intIdOption = 336 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 10 Monitor Capable:</option>")
				elseif intIdOption = 330 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 12 Monitor Capable:</option>")
				elseif intIdOption = 342 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel 14th Generation CPUs:</option>")
				 elseif intIdOption = 337 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel 14th Generation CPUs:</option>")
				 elseif intIdOption = 356 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel Core Ultra CPUs:</option>")
                end if


				%>
				<option class="spec-dd-option" value="<%=intIdOptOptGrp%>" id="<%=intIdOption%>" title="<%=Round(optPrice/1.2,2)%>"
					<% 
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Check if Option should be Selected
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' DA Edit to force selection if sort order is set to 999
					if strOptSortOrder=999 then
						response.write " selected"
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					Dim xIdOptCounter
					
					if tIndex<>0 then ' Check they are updating the product after adding it to the shopping cart
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
						pcCartArray=session("pcCartSession")
						tempIdOpt = ""
						tempIdOpt = pcCartArray(tIndex,11)
						
						if tempIdOpt = "" then
							response.write ">"
						else
							tempIdOpt = Split(trim(tempIdOpt),chr(124))							
							for xIdOptCounter = 0 to Ubound(tempIdOpt)
								if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
									response.write " selected"								
								end if
							next
							response.write ">"
						end if						

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					else
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						tempIdOpt = ""
						tempIdOpt = request.querystring("idOptionArray")
						
						if tempIdOpt = "" then
						'' DA Edit to ignore wifi card option for bundles
						if intIdOption = 148 then
						elseif intIdOption = 149 then
						else
							response.write ">"
							end if
						else
							tempIdOpt = Split(trim(tempIdOpt),chr(124))
							for xIdOptCounter = 0 to Ubound(tempIdOpt)
								if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
									response.write " selected"
									pcv_isOptionSelected="1"								
								end if
							next
							response.write ">"
						end if						
						
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: Check if Option should be Selected
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Display Option Name
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'' DA Edit to ignore wifi card option for bundles
				if intIdOption = 148 then
				elseif intIdOption = 149 then
				else
					response.write strOptionDescrip
				end if					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: Display Option Name
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Display Pricing
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					if optPrice>0 then  
					'// If there is a price thats greater than zero
				'' DA Edit to ignore wifi card option for bundles
				if intIdOption = 148 then
				elseif intIdOption = 149 then
				else
					%>
						<%=" + " & scCurSign& money(optPrice/1.2)%>  
					<% 
				end if
					end if %>
					<% 
					if optPrice<0 then 
					'// If there is not a price
					%>
						<%=" - "&scCurSign& money(optPrice/1.2)%> 
					<% 
					end if 
					
					
					
					if optPrice=0 then 
					'// If there is not a price
					%>
						<%
						if strOptionDescrip = "None" then
							response.write "&nbsp;&nbsp;"
						else
							If iCount = 0 Then
								if intIdOption = 148 then
								elseif intIdOption = 149 then
								else
								response.write " - Included"
								end if
							Else
							response.write " - No Charge"
							End If
							iCount = iCount + 1
						end if
						%> 
					<% 
					end if 
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: Display Pricing
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if intIdOption = 148 then
					elseif intIdOption = 149 then
					else
					%>										
				</option>
				<% 
					end if
				end if
				'**************************************************************************************************
				' END: Dispay the Options
				'**************************************************************************************************

				 'DA Edit to 128GB RAM option on Extreme PC
				'if intIdOption = 324 then
					'Response.write("<option value=""title"" class=""spec-dd-option"" disabled>128GB DDR5 - Call Us</option>")
                'end if
 				%>
				   <%
			rs2.movenext 
			loop
			
				
			'// Only execute when the Remove Option Feature is activated.
			if pcv_strAdminPrefix="1" AND pcv_strRemoveFeature="1" then %>		
				<% if pcv_isOptionSelected="1" then %>
					<option value=""></option>
					<option value="">----- Remove Option -----</option>
				<% else %>
					<option value="" selected><%=dictLanguage.Item(Session("language")&"_viewPrd_61")%></option>
				<% end if %>
			<% end if 	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		


	
	
	set rs2=nothing	
	end if
	%>
</select>
</div>
<%=strEndQM%></div>
</div><!-- Row end -->
<%
'//Check if UltimatePC option is checked, if so end current table, add line of text then restart table
if boolUltimatePC = "1" then
%>
</table>
<p><span style="font-weight:bold;">Upgrade your new Ultimate Trading PC even further here:</span></p><br />
<table class="upgoptions">
<%
boolUltimatePC = "0"
end if
%>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Options Box
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Options (BUN) PC Bundles
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_OptionsBUN
	' SELECT DATA SET
	' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
	query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
	query = query & "FROM products "
	query = query & "INNER JOIN ( "
	query = query & "pcProductsOptions INNER JOIN ( "
	query = query & "optionsgroups "
	query = query & "INNER JOIN options_optionsGroups "
	query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
	query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
	query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
	query = query & "WHERE products.idProduct=" & pidProduct &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order, optionsGroups.OptionGroupDesc;"
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)	
	'if err.number<>0 then
		'call LogErrorToDatabase()
		'set rs=nothing
		'call closedb()
		'response.redirect "techErr.asp?err="&pcStrCustRefID
	'end if
	' If we have data	
	if NOT rs.eof then
				%>
                <div id="spec-options" class=" wow fadeInLeft" data-wow-delay="0.1s">
                <%
		pcv_intOptionGroupCount = 0 '// keeps count of the number of options
		xOptionsCnt = 0 '// keeps count of the number of required options
		do until rs.eof				
			
			'if pcv_intOptionGroupCount <= 5  then ' // start limit to 5 options
				'// Get the Group Name
				pcv_strOptionGroupDesc=rs("OptionGroupDesc")
				'// Get the Group ID
				pcv_strOptionGroupID=rs("idOptionGroup")
				'// Is it required
				pcv_strOptionRequired=rs("pcProdOpt_Required")			
		
				'// Start: Do Option Count
				pcv_intOptionGroupCount = pcv_intOptionGroupCount + 1 
				'// End: Do Option Count
				
				'// Get the number of the Option Group
				pcv_strOptionGroupCount = pcv_intOptionGroupCount
				
				'// Start: Do Required Option Count AND generate validation string
				if IsNull(pcv_strOptionRequired) OR pcv_strOptionRequired="" then
						pcv_strOptionRequired=0 '// not required // else it is "1"
				end if			
				if pcv_strOptionRequired=1 then
					
					' Keep Tally
					xOptionsCnt = xOptionsCnt + 1
					
					' Generate String
					if xOtionrequired="1" then
						pcv_strReqOptString = pcv_strReqOptString & ","
					end if
				
					xOtionrequired="1"
					pcv_strOptionGroupDesc2=pcv_strOptionGroupDesc
					pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"'","")
					pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"""","\'\'")
					pcv_strReqOptString = pcv_strReqOptString & "document.additem.idOption" & pcv_strOptionGroupCount & ".selectedIndex,'"& pcv_strOptionGroupDesc2 &"'"
				
				end if
				'// End: Do Required Option Count
				'// Make the Option Box
				pcs_makeOptionBoxBUN							
			'end if ' // end limit to 5 options
		rs.movenext
		loop		
				%>
                </table>
                <%
	end if
	set rs=nothing
%>
<input type="hidden" name="OptionGroupCount" value="<%=pcv_intOptionGroupCount%>">
									</div>
								</div>

<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Options (BUN) PC Bundles
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Options Box PC Bundles
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_makeOptionBoxBUN
	' SELECT DATA SET
	' TABLES: options_optionsGroups, options
	query = 		"SELECT options_optionsGroups.InActive, options_optionsGroups.price, options_optionsGroups.Wprice, "
	query = query & "options_optionsGroups.idoptoptgrp, options_optionsGroups.sortOrder, options.idoption, options.optiondescrip "
	query = query & "FROM options_optionsGroups "
	query = query & "INNER JOIN options "
	query = query & "ON options_optionsGroups.idOption = options.idOption "
	query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_strOptionGroupID &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	'query = query & "ORDER BY options_optionsGroups.sortOrder, options.optiondescrip;"	
	query = query & "ORDER BY options_optionsGroups.sortOrder, options_optionsGroups.price, options.optiondescrip;"	
	set rs2=server.createobject("adodb.recordset")
	set rs2=conntemp.execute(query)	
	'if err.number<>0 then
		'call LogErrorToDatabase()
		'set rs2=nothing
		'call closedb()
		'response.redirect "techErr.asp?err="&pcStrCustRefID
	'end if
	
	' If we have data
	if NOT rs2.eof then
	
	'DA Edit to try and hide DisplayPort Row
	if pcv_strOptionGroupDesc = "DisplayPort Adapters" then
		strDPRowID = " id=""trGA"""
	else
		strDPRowID = ""
	end if
	%>
    
    
<div<%=strDPRowID%> class="row specb-row"><!-- Row Start -->
<div class="col-sm-3 specb-title">
<label>
                <%
		'// clean up the option group description
		if pcv_strOptionGroupDesc<>"" then
			pcv_strOptionGroupDesc=replace(pcv_strOptionGroupDesc,"""","&quot;")
		end if 
		
		'// START SELECT
		pcv_isOptionSelected="" '// Is this option box selected? Fill variable to "1" during the following loop.

		'DA Edit - to add an extra JS function to the onchange for graphics cards only
		daJSFunction = ""
		
		Select Case pcv_strOptionGroupDesc 
			Case "MS Office" 
				Response.Write "Microsoft Office:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Microsoft Office"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-office2013.htm"">Learn More</a></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Security Software:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Microsoft Windows Defender AntiVirus</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Bonus Software:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Display Fusion - (Multi Screen Management Tools)</p><a data-toggle=""lightbox"" data-title=""Display Fusion Software"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-displayfusion.htm"">Learn More</a>"
			Case "CPU" 
				Response.Write "CPU / Processor:"
				if not InStr(pSku, "EXT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""CPU / Processor"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-cpu-ext.htm"">Learn More</a></div></div><!-- Row end -->"
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Motherboard:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optMotherboard""></span></p><a data-toggle=""lightbox"" data-title=""Motherboard"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-mboard-x79.htm"">Learn More</a>"
					daJSFunction = "reCalcColourEXT();"
				end if 
				if not InStr(pSku, "ULT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""CPU / Processor"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-cpu-ultra.htm"">Learn More</a></div></div><!-- Row end -->"
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Motherboard:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optMotherboard""></span></p><a data-toggle=""lightbox"" data-title=""Motherboard"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-mboard-z170.htm"">Learn More</a>"
				end if 
				if not InStr(pSku, "PRO1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""CPU / Processor"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-cpu-pro.htm"">Learn More</a></div></div><!-- Row end -->"
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Motherboard:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optMotherboard""></span></p><a data-toggle=""lightbox"" data-title=""Motherboard"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-mboard-z170.htm"">Learn More</a>"
				end if 
			Case "Keyb. / Mouse"
				Response.Write "Keyboard &amp; Mouse:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Keyboard &amp; Mouse"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-keyboards.htm"">Learn More</a>"
			Case "Speakers"
				Response.Write "Speakers:"
				strEndQM = ""
			Case "2nd Hard Drive"
				Response.Write "2nd Hard Drive:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""2nd (Storage) Hard Drive"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-storage-drive.htm"">Learn More</a>"
			Case "OS"
				Response.Write "Operating System:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Microsoft Windows Versions"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-os.htm"">Learn More</a>"
			Case "Power Leads"
				Response.Write "Power Leads:"
				strEndQM = ""
			Case "RAM"
				Response.Write "System RAM:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""RAM / Memory"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-ram.htm"">Learn More</a>"
			Case "Wireless Card"
				Response.Write "Wireless Network Card:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Wireless Network Card (WiFi)"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-wifi.htm"">Learn More</a></div></div><!-- Row end -->"
				if not InStr(pSku, "PRO1") = 0 Then
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Fractal Design Core 1100 (W:175mm, H:355mm, D:420mm)</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case Cooling:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCaseCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Power Supply:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optPSU""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>CPU Cooler:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCPUCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>USB Ports:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">4 x USB 3 (4 Rear) 4 x USB 2 (2 Rear & 2 Front)</p>"
				end if 
				if not InStr(pSku, "ULT1") = 0 Then
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Antec P7 (W:210mm, H:470mm, D:445mm)</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case Cooling:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCaseCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Power Supply:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optPSU""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>CPU Cooler:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCPUCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>USB Ports:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">4 x USB 3 (2 Front & 2 Rear) 4 x USB 2 (4 Rear)</p>"
				end if 
				if not InStr(pSku, "EXT1") = 0 Then
					strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Corsair 3000D Case (W:230mm, H:466mm, D:462mm)</p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Case Cooling:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCaseCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Power Supply:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optPSU""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>CPU Cooler:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text""><span id=""optCPUCool""></span></p></div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>USB Ports:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">4 x USB 3 (2 Front & 2 Rear) 8 x USB 2 (8 Rear)</p>"
				end if 
			Case "Optical Drive" 
				Response.Write "Optical Drive:"
				strEndQM = "</div></div><!-- Row end --><div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Network / LAN Port:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Integrated Gigabit Ethernet LAN Adapter</p>"
			Case "Boot Hard Drive"
				Response.Write "Boot Hard Drive:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Boot Hard Drive"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-boot-ssd.htm"">Learn More</a>"
			Case "Warranty Cover"
				Response.Write "Warranty Cover:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""On-Site Warranty Cover"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-warranty.htm"">Learn More</a>"
			Case "DisplayPort Adapters"
				Response.Write "DisplayPort Adapters:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Card DisplayPort Adapters"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-adapters.htm"">Learn More</a></div></div><!-- Row end -->"
				strEndQM = strEndQM & "<div class=""row specb-row""><!-- Row Start --><div class=""col-sm-3 specb-title""><label>Sound Card:</label></div><div class=""col-sm-9 specb-field""><p class=""specb-text"">Integrated 6 Channel HD Audio Sound Card</p><a data-toggle=""lightbox"" data-title=""Sound Card / Speakers"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-sound.htm"">Learn More</a>"
			Case "Graphics Cards"
				Response.Write "Graphics Card Setup:"
				if not InStr(pSku, "PRO1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Cards Setup Options"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-graphics-210.htm"">Learn More</a>"
					daJSFunction = "reCalcColour();"				
				end if 
				if not InStr(pSku, "ULT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Cards Setup Options"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-graphics.htm"">Learn More</a>"
					daJSFunction = "reCalcColour();"
				end if 
				if not InStr(pSku, "EXT1") = 0 Then
					strEndQM = "<a data-toggle=""lightbox"" data-title=""Graphics Cards Setup Options"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-graphics.htm"">Learn More</a>"
					daJSFunction = ""
				end if 
			Case "Backup System"
				Response.Write "Backup System:"
				strEndQM = "<a data-toggle=""lightbox"" data-title=""Bootable Backup Hard Drive"" title=""Learn more"" class=""spec-more-link"" href=""/pop-pages/custpc-backup.htm"">Learn More</a>"
			Case Else
				Response.Write pcv_strOptionGroupDesc
				strEndQM = ""
		End Select
		%>
		</label>
        </div>
        <div class="col-sm-9 specb-field">
		<div class="specb-dd">
		<select id="idOption<%=pcv_strOptionGroupCount%>" name="idOption<%=pcv_strOptionGroupCount%>"  class="spec-dd" onchange="reCalc();<%=daJSFunction%>">
		<%
			'Attempt to stop 2nd free options saying included
			icount = 0
		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			do until rs2.eof			
			
			OptInActive=rs2("InActive") ' Is it active?
			if IsNull(OptInActive) OR OptInActive="" then
				OptInActive="0"
			end if
			
			dblOptPrice=rs2("price") '// Price
			dblOptWPrice=rs2("Wprice") '// WPrice
			intIdOptOptGrp=rs2("idoptoptgrp") '// The Id of the Option Group
			intIdOption=rs2("idoption") '// The Id of the Option
			strOptionDescrip=rs2("optiondescrip") '// A description of the Option
			strOptSortOrder=rs2("sortOrder") '// Sort order set in Admin
	
			'**************************************************************************************************
			' START: Dispay the Options
			'**************************************************************************************************
			if OptInActive="0" then
				If session("customerType")=1 then 
					optPrice=dblOptWPrice
				Else
					optPrice=dblOptPrice
				End If 
				
				'' DA Edit to ignore wifi card option for bundles
				if intIdOption = 29 then
				%>
                <%
				elseif intIdOption = 26 then
				%>
                <%
				elseif intIdOption = 197 then
				%>
                <%
				else
				
				'DA Edit to insert dotted line between GPU Options
				if intIdOption = 256 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Traditional Hard Drives:</option>")
				elseif intIdOption = 254 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Fast &amp; Silent SSDs:</option>")
				elseif intIdOption = 262 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel 10th Generation CPUs:</option>")
				elseif intIdOption = 318 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>AMD Ryzen CPUs:</option>")
				elseif intIdOption = 273 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Dual Monitor Capable:</option>")
				elseif intIdOption = 297 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Quad Monitor Capable:</option>")
				elseif intIdOption = 298 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Six Monitor Capable:</option>")
				elseif intIdOption = 299 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Eight Monitor Capable:</option>")
				elseif intIdOption = 300 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 3 Monitor Capable:</option>")
				elseif intIdOption = 333 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 4 Monitor Capable:</option>")
				elseif intIdOption = 361 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 4 Monitor Capable:</option>")
				elseif intIdOption = 335 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 6 Monitor Capable:</option>")
				elseif intIdOption = 334 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 8 Monitor Capable:</option>")
				 elseif intIdOption = 362 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 8 Monitor Capable:</option>")
				elseif intIdOption = 336 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 10 Monitor Capable:</option>")
				elseif intIdOption = 330 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Up To 12 Monitor Capable:</option>")
				elseif intIdOption = 342 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel 14th Generation CPUs:</option>")
				elseif intIdOption = 337 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel 14th Generation CPUs:</option>")
				elseif intIdOption = 356 then
					Response.write("<option value=""title"" class=""spec-dd-dis"" disabled>Intel Core Ultra CPUs:</option>")
                end if

				%>
				<option class="spec-dd-option" value="<%=intIdOptOptGrp%>" id="<%=intIdOption%>" title="<%=Round(optPrice/1.2,2)%>"
					<% 	
				end if				
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Check if Option should be Selected
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' DA Edit to force selection if sort order is set to 999
					if strOptSortOrder=999 then
						response.write " selected"
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					Dim xIdOptCounter
					
					if tIndex<>0 then ' Check they are updating the product after adding it to the shopping cart
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
						pcCartArray=session("pcCartSession")
						tempIdOpt = ""
						tempIdOpt = pcCartArray(tIndex,11)
						
						if tempIdOpt = "" then
							response.write ">"
						else
							tempIdOpt = Split(trim(tempIdOpt),chr(124))							
							for xIdOptCounter = 0 to Ubound(tempIdOpt)
								if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
									response.write " selected"								
								end if
							next
							response.write ">"
						end if						

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					else
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						tempIdOpt = ""
						tempIdOpt = request.querystring("idOptionArray")
						
						if tempIdOpt = "" then
						'' DA Edit to ignore wifi card option for bundles
						if intIdOption = 29 then
						elseif intIdOption = 26 then
						else
							response.write ">"
							end if
						else
							tempIdOpt = Split(trim(tempIdOpt),chr(124))
							for xIdOptCounter = 0 to Ubound(tempIdOpt)
								if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
									response.write " selected"
									pcv_isOptionSelected="1"								
								end if
							next
							response.write ">"
						end if						
						
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: Check if Option should be Selected
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Display Option Name
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					'' DA Edit to ignore wifi card option for bundles
				if intIdOption = 29 then
				elseif intIdOption = 26 then
				else
					response.write strOptionDescrip
				end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: Display Option Name
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Display Pricing
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					if optPrice>0 then  
					'// If there is a price thats greater than zero
					
				'' DA Edit to ignore wifi card option for bundles
				if intIdOption = 29 then
				elseif intIdOption = 26 then
				else
					%>
						<%=" + " & scCurSign& money(optPrice/1.2)%>  
					<% 
				end if

					end if %>
					<% 
					if optPrice<0 then 
					'// If there is not a price
					%>
						<%=" - "&scCurSign& money(optPrice/1.2)%> 
					<% 
					end if 
					
					
					
					if optPrice=0 then 
					'// If there is not a price
					%>
						<%
						if strOptionDescrip = "None" then
							response.write "&nbsp;&nbsp;"
						else
							If iCount = 0 Then
							response.write " - Included"
							Else
							response.write " - No Charge"
							End If
							iCount = iCount + 1
						end if
						%> 
					<% 
					end if 
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: Display Pricing
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if intIdOption = 29 then
					elseif intIdOption = 26 then
					else
					%>										
				</option>
				<% 
					end if
				end if
				'**************************************************************************************************
				' END: Dispay the Options
				'**************************************************************************************************
				 'DA Edit to 128GB RAM option on Extreme PC
				if intIdOption = 324 then
					Response.write("<option value=""title"" class=""spec-dd-option"" disabled>128GB DDR5 - Call Us</option>")
                end if
				%>
                <%
			rs2.movenext 
			loop
			
				
			'// Only execute when the Remove Option Feature is activated.
			if pcv_strAdminPrefix="1" AND pcv_strRemoveFeature="1" then %>		
				<% if pcv_isOptionSelected="1" then %>
					<option value=""></option>
					<option value="">----- Remove Option -----</option>
				<% else %>
					<option value="" selected><%=dictLanguage.Item(Session("language")&"_viewPrd_61")%></option>
				<% end if %>
			<% end if 	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		


	
	
	set rs2=nothing	
	end if
	%>
</select>
</div>
<%=strEndQM%></div>
</div><!-- Row end -->
<%
'//Check if UltimatePC option is checked, if so end current table, add line of text then restart table
if boolUltimatePC = "1" then
%>
</table>
<p><span style="font-weight:bold;">Upgrade your new Ultimate Trading PC even further here:</span></p><br />
<table class="upgoptions">
<%
boolUltimatePC = "0"
end if
%>

<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Options Box - PC Bundles
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Add to Cart (Dynamic)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddtoCart
	pcv_strFuntionCall = "cdDynamic"

	'// Check in-stock sub-product(s)
	IF pcv_Apparel=1 THEN
		Dim query,rs
		query="SELECT TOP 1 idProduct FROM Products WHERE pcProd_ParentPrd=" & pIdProduct & " AND ((stock>0) OR (pcProd_BackOrder<>0) OR (noStock<>0)) AND pcProd_SPInActive=0 AND removed=0;"
		set rs=connTemp.execute(query)
		if rs.eof then
			set rs=nothing%>
			<input  type="hidden" name="quantity" value="1">
			<%exit sub
		end if
		set rs=nothing
	END IF

	%>
    <!-- Cart -->
    <div class="pcShowAddToCart">
		<%
		if tIndex<>0 then '// Check they are updating the product after adding it to the shopping cart
				pcCartArray=session("pcCartSession")
				tempQty = ""
				tempQty = pcCartArray(tIndex,2)
				if tempQty<>"" then
					pcv_intQuantityField=tempQty
				else
					pcv_intQuantityField=1
				end if	
			else
				if pcv_lngMinimumQty <> 0 then
					pcv_intQuantityField=pcv_lngMinimumQty
				else
					pcv_intQuantityField=1
				end if
			end if
			%>
			<%
			'SB S
			if pSubscriptionID > 0 and pSubType <> 2 Then%>
				<input class="form-control form-control-inline" type="hidden" name="quantity" value="1"> 
			<%else%>

			<input type="text" class="form-control form-control-inline" name="quantity" size="10" maxlength="10" value="<%=pcv_intQuantityField%>" <%if pcv_Apparel="1" then%>onblur="checkproqty(document.additem.quantity);"<%end if%>>

			<%
			End if 
			'SB E %>
      
		<input type="hidden" name="idproduct" value="<%=pidProduct%>">
		<meta itemprop="productID" content="<%=pidProduct%>" />
		<meta itemprop="url" content="<%=scmURL%>" />

		
<% 
'// there is at least one reuqired custom field	
if xrequired="1" then
%>

		<% 
		If BTOCharges=0 then
			if xOtionrequired = "1" then '// If there are any required options at all.
			'// figure some stuff out
			%>
				<a class="pcButton pcButtonAddToCart" href="#" onClick="javascript: if (CheckRequiredCS('<%=pcv_strReqCSString%>')) {if (checkproqty(document.additem.quantity)) {<%=pcv_strFuntionCall%>(<%=pcv_strReqOptString%>,<%=reqstring%>,0);}} return false;"><%showAddtoCart=1%>
        	<img src="<%=pcf_getImagePath(pcv_tmpNewPath,rslayout("addtocart"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_2b")%>">
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtocart") %></span>
        </a>
			<%
			else ' There are no required options at all.
			%>
        <a class="pcButton pcButtonAddToCart" href="#" onClick="javascript: if (CheckRequiredCS('<%=pcv_strReqCSString%>')) {if (checkproqty(document.additem.quantity)) {<%=pcv_strFuntionCall%>(<%=reqstring%>,0);}} return false;"><%showAddtoCart=1%>
          <img src="<%=pcf_getImagePath(pcv_tmpNewPath,rslayout("addtocart"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_2b")%>">
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtocart") %></span>
        </a> 
			<% 
			end if
		end if 
		%>
		
<% 
' There are no required custom fields.
else
%>
		
		<% 
		If BTOCharges=0 then
			if xOtionrequired = "1" then '// If there are any required options at all.
			'// figure some stuff out
			%>
        <a class="pcButton pcButtonAddToCart" href="#" onClick="javascript: if (CheckRequiredCS('<%=pcv_strReqCSString%>')) {if (checkproqty(document.additem.quantity)) {<%=pcv_strFuntionCall%>(<%=pcv_strReqOptString%>,0);}} return false"><%showAddtoCart=1%>
        	<img src="<%=pcf_getImagePath(pcv_tmpNewPath,rslayout("addtocart"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_2b")%>">
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtocart") %></span>
        </a>
			<% else %>
			<%showAddtoCart=1%>
			<%
			'SB S
			'// Don't show if subscription and subscriptions disabled
			if (pSubscriptionID = "0" OR scSBStatus="1") OR (pcv_strAdminPrefix="1") then %> 
      <button class="pcButton pcButtonAddToCart" id="submit" name="add">
				<img src="<%=pcf_getImagePath(pcv_tmpNewPath,rslayout("addtocart"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_2b")%>" />
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtocart") %></span>
      </button>
			<% 
			end if
			'SB E
			%>
			<% 
			end if
		End if 
		%>
		

<% 
'// End ADD TO CART SECTION
end if 
%>
		
		<% If pserviceSpec<>0 then
		Dim rsQ,queryQ
		queryQ="SELECT TOP 1 configProduct FROM configSpec_products WHERE specProduct=" & pIdProduct & ";"
		set rsQ=connTemp.execute(queryQ)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsQ=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rsQ.eof then
			showCustomize=1
            
            If pcf_BTOisConfig Then
                pcv_strConfigClickEvent = "javascript:parent.location='configurePrd.asp?idproduct=" & pIdProduct & "&qty='+document.additem.quantity.value; return false;"
            Else
                pcv_strConfigClickEvent = "javascript:document.additem.action='configurePrd.asp?idproduct=" & pIdProduct & "&qty='+document.additem.quantity.value;"
            End If            
            %>
			<button class="pcButton pcButtonCustomize" onclick="<%=pcv_strConfigClickEvent%>">
				<img src="<%=pcf_getImagePath(pcv_tmpNewPath,rslayout("customize"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_customize")%>" />
				<span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_customize")%></span>
			</button>
		<%End if
		set rsQ=nothing
		End If %>		
	</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Add to Cart (Dynamic)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'SB S
Public Sub pcs_SubscriptionProduct

	If pSubscriptionID <> 0  then
		
	  	If pIsLinked="1" Then
			%> <!--#include file="inc_sb_widget.asp"--> <%
		End If	  

	 	response.write "<input type=""hidden"" name=""pSubscriptionID"" id=""pcSubId"" value="""&pSubscriptionID&""">"
		
	End If
	
End Sub
'SB S


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Product Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_ProductPromotionMsg
	Dim rs,rsQ,query,tmpStr

	query="SELECT pcPrdPro_id,idproduct,pcPrdPro_QtyTrigger,pcPrdPro_DiscountType,pcPrdPro_DiscountValue,pcPrdPro_ApplyUnits,pcPrdPro_PromoMsg,pcPrdPro_ConfirmMsg,pcPrdPro_SDesc,pcPrdPro_IncExcCust,pcPrdPro_IncExcCPrice,pcPrdPro_RetailFlag,pcPrdPro_WholesaleFlag FROM pcPrdPromotions WHERE pcPrdPro_Inactive=0 AND idproduct=" & pIDProduct & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		pcv_HavePrdPromotions=1
		PrdPromoArr=rsQ.getRows()
		set rsQ=nothing
		PrdPromoCount=ubound(PrdPromoArr,2)
		
		tmpIDCode=PrdPromoArr(0,0)
		tmpIDProduct=PrdPromoArr(1,0)
		tmpQtyTrigger=clng(PrdPromoArr(2,0))
		tmpDiscountType=PrdPromoArr(3,0)
		tmpDiscountValue=PrdPromoArr(4,0)
		tmpApplyUnits=PrdPromoArr(5,0)
		tmpConfirmMsg=PrdPromoArr(7,0)
		tmpDescMsg=PrdPromoArr(8,0)
		pcIncExcCust=PrdPromoArr(9,0)
		pcIncExcCPrice=PrdPromoArr(10,0)
		pcv_retail=PrdPromoArr(11,0)
		pcv_wholeSale=PrdPromoArr(12,0)
		
		pcv_Filters=0
		pcv_FResults=0
		'Filter by Customers
		pcv_CustFilter=0
		query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			pcv_Filters=pcv_Filters+1
			pcv_CustFilter=1
		end if
		set rs=nothing
		
		if pcv_CustFilter=1 then
				
		query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode & " and IDCustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			if (pcIncExcCust="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCust="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing
		
		end if
		'End of Filter by Customers
		
		
		'Filter by Customer Categories
		pcv_CustCatFilter=0
		
		query="select idCustomerCategory from pcPPFCustPriceCats where pcPrdPro_id=" & tmpIDCode
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			pcv_Filters=pcv_Filters+1
			pcv_CustCatFilter=1
		end if
		set rs=nothing
		
		if pcv_CustCatFilter=1 then
				
		query="select pcPPFCustPriceCats.idCustomerCategory from pcPPFCustPriceCats, Customers where pcPPFCustPriceCats.pcPrdPro_id=" & tmpIDCode & " and pcPPFCustPriceCats.idCustomerCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			if (pcIncExcCPrice="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCPrice="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing
		
		end if
		'End of Filter by Customer Categories
		
		' Check to see if promotion is filtered by reatil or wholesale.
		if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
			pcv_Filters=pcv_Filters+1
			if pcv_wholeSale = "1" and session("customertype") = 1 then
				pcv_FResults=pcv_FResults+1		
			end if 
			if pcv_retail = "1" and session("customertype") <> 1 Then
				pcv_FResults=pcv_FResults+1
			end if    
		end if
		
		if (pcv_Filters=pcv_FResults) AND PrdPromoArr(6,0)<>"" then%>
			<div class="pcPromoMessage">
				<%=PrdPromoArr(6,0)%>
	    	</div>
		<%end if
	end if
	set rsQ=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Product Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: QuickView Functions
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function pcf_ProductDescriptionQV
    '// Display short product description if it isn't empty
	If len(psDesc)>0 Then 
		pcf_ProductDescriptionQV= "<div class='pcShowProductSDesc' style='padding-top: 5px'>"
		pcf_ProductDescriptionQV=pcf_ProductDescriptionQV & psDesc 
		pcf_ProductDescriptionQV=pcf_ProductDescriptionQV & "</div>"
	End If
End Function

Public Function pcf_CustomSearchFieldsQV
Dim query,rs,pcArr,intCount,i
	query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pIdProduct & " AND pcSearchFieldShow=1 ORDER BY pcSearchFields.pcSearchFieldOrder ASC,pcSearchFields.pcSearchFieldName ASC;"
	set rs=connTemp.execute(query)
	IF not rs.eof THEN
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		pcf_CustomSearchFieldsQV="<div style='padding-top: 5px;'></div>"
		For i=0 to intCount
			searchFieldLink = "showsearchresults.asp?customfield=" & pcArr(0,i) & "&SearchValues=" & Server.URLEncode(pcArr(2,i))
			
			pcf_CustomSearchFieldsQV=pcf_CustomSearchFieldsQV & "<div class='pcShowProductCustSearch'>"&pcArr(1,i)&": <a href='" & Server.HtmlEncode(searchFieldLink) & "'>"&pcArr(3,i)&"</a></div>"
		Next
	END IF
	set rs=nothing
End Function

Public Function pcf_ShowBrandQV
	if sBrandPro="1" then
		if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
			pcf_ShowBrandQV="<div class='pcShowProductBrand'>"
			pcf_ShowBrandQV=pcf_ShowBrandQV & dictLanguage.Item(Session("language")&"_viewPrd_brand")
			pcf_ShowBrandQV=pcf_ShowBrandQV & BrandName
			pcf_ShowBrandQV=pcf_ShowBrandQV & "</div>"
		end if
	end if
End Function

Public Function pcf_DisplayWeightQV

Dim query,rs,totalSubWeight
query="SELECT sum(weight) As TotalWeight FROM Products WHERE pcProd_ParentPrd=" & pidProduct & " AND removed=0 and pcProd_SPInActive=0 GROUP BY pcProd_ParentPrd;"
set rs=connTemp.execute(query)
if not rs.eof then
	totalSubWeight=rs("TotalWeight")
else
	totalSubWeight=0
end if
set rs=nothing

if scShowProductWeight="-1" then
		if (int(pWeight)>0) OR (totalSubWeight>0) then
			pcf_DisplayWeightQV="<div class='pcShowProductWeight'>"
			pcf_DisplayWeightQV=pcf_DisplayWeightQV & ship_dictLanguage.Item(Session("language")&"_viewCart_c")
			if scShipFromWeightUnit="KGS" then
				pKilos=Int(pWeight/1000)
				pWeight_g=pWeight-(pKilos*1000)
				pWeight=pKilos
				pcf_DisplayWeightQV=pcf_DisplayWeightQV &  dictLanguage.Item(Session("language")&"_viewCart_c")
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & "<span id=""appw1"" name=""appw1"">" & pWeight & "</span>"
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & " kg "
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & "<span id=""appw2"" name=""appw2"">" & pWeight_g & "</span>"
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & " g" & "<br />"
			else
				pPounds=Int(pWeight/16)
				pWeight_oz=pWeight-(pPounds*16)
				pWeight=pPounds
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & dictLanguage.Item(Session("language")&"_viewCart_c")
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & "<span id=""appw1"" name=""appw1"">" & pPounds & "</span>"
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & " lbs "
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & "<span id=""appw2"" name=""appw2"">" & pWeight_oz & "</span>"
				pcf_DisplayWeightQV=pcf_DisplayWeightQV & " ozs" & "<br />"
			end if
			pcf_DisplayWeightQV=pcf_DisplayWeightQV & "</div>"
		else
			pcf_DisplayWeightQV="<span style=""display:none"" id=""appw1"" name=""appw1""></span>"
			pcf_DisplayWeightQV=pcf_DisplayWeightQV & "<span style=""display:none"" id=""appw2"" name=""appw2""></span>"
		end if
else
	pcf_DisplayWeightQV="<span style=""display:none"" id=""appw1"" name=""appw1""></span>"
	pcf_DisplayWeightQV=pcf_DisplayWeightQV & "<span style=""display:none"" id=""appw2"" name=""appw2""></span>"
end if

End Function

Public Function pcf_RewardPointsQV
	
	If RewardsActive=1 then
		' Show Reward Points associated with this product, if any
		' By default, Reward Points are not shown to Wholesale Customers
		if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")<>"1" then
			pcf_RewardPointsQV="<div style='padding-top: 5px;'>"&dictLanguage.Item(Session("language")&"_viewPrd_50")&Clng(iRewardPoints+clng(pcv_BTORP))&"&nbsp;"&RewardsLabel&dictLanguage.Item(Session("language")&"_viewPrd_51")&"</div>"
			pcf_RewardPointsQV=pcf_RewardPointsQV & "<script>DefaultReward=" & Clng(iRewardPoints+clng(pcv_BTORP)) & ";</script>"
		else
			' If the system is setup to include Wholesale Customers, then show Reward Points to them too
			if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")="1" and RewardsIncludeWholesale=1 then
				pcf_RewardPointsQV="<div style='padding-top: 5px;'>"&dictLanguage.Item(Session("language")&"_viewPrd_50")&Clng(iRewardPoints+clng(pcv_BTORP))&"&nbsp;"&RewardsLabel&dictLanguage.Item(Session("language")&"_viewPrd_51")&"</div>"
				pcf_RewardPointsQV=pcf_RewardPointsQV & "<script>DefaultReward=" & Clng(iRewardPoints+clng(pcv_BTORP)) & ";</script>"
			else
				pcf_RewardPointsQV= "<script>DefaultReward=0;</script>"
			end if 
		end If
		
	Else
		pcf_RewardPointsQV= "<script>DefaultReward=0;</script>"
	End If

End Function

Public Function pcf_UnitsStockQV
	if scdisplayStock=-1 AND pNoStock=0 then
		if pstock > 0 then
			pcf_UnitsStockQV="<div class='pcShowProductStock'>"
			pcf_UnitsStockQV=pcf_UnitsStockQV & dictLanguage.Item(Session("language")&"_viewPrd_19") & " " & pStock
			pcf_UnitsStockQV=pcf_UnitsStockQV &"</div>"
		end if
	end if
End Function

Public Function pcf_DisplayBOMsgQV
	If (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=1) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=1) Then
		If clng(pcv_intShipNDays)>0 then
			pcf_DisplayBOMsgQV="<div>"&dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b") & "</div>"
		End if
	End If
End Function

Public Function pcf_ProductPromotionMsgQV
	Dim rs,rsQ,query,tmpStr

	query="SELECT pcPrdPro_id,idproduct,pcPrdPro_QtyTrigger,pcPrdPro_DiscountType,pcPrdPro_DiscountValue,pcPrdPro_ApplyUnits,pcPrdPro_PromoMsg,pcPrdPro_ConfirmMsg,pcPrdPro_SDesc,pcPrdPro_IncExcCust,pcPrdPro_IncExcCPrice,pcPrdPro_RetailFlag,pcPrdPro_WholesaleFlag FROM pcPrdPromotions WHERE pcPrdPro_Inactive=0 AND idproduct=" & pIDProduct & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		pcv_HavePrdPromotions=1
		PrdPromoArr=rsQ.getRows()
		set rsQ=nothing
		PrdPromoCount=ubound(PrdPromoArr,2)
		
		tmpIDCode=PrdPromoArr(0,0)
		tmpIDProduct=PrdPromoArr(1,0)
		tmpQtyTrigger=clng(PrdPromoArr(2,0))
		tmpDiscountType=PrdPromoArr(3,0)
		tmpDiscountValue=PrdPromoArr(4,0)
		tmpApplyUnits=PrdPromoArr(5,0)
		tmpConfirmMsg=PrdPromoArr(7,0)
		tmpDescMsg=PrdPromoArr(8,0)
		pcIncExcCust=PrdPromoArr(9,0)
		pcIncExcCPrice=PrdPromoArr(10,0)
		pcv_retail=PrdPromoArr(11,0)
		pcv_wholeSale=PrdPromoArr(12,0)
		
		pcv_Filters=0
		pcv_FResults=0
		'Filter by Customers
		pcv_CustFilter=0
		query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			pcv_Filters=pcv_Filters+1
			pcv_CustFilter=1
		end if
		set rs=nothing
		
		if pcv_CustFilter=1 then
				
		query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode & " and IDCustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			if (pcIncExcCust="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCust="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing
		
		end if
		'End of Filter by Customers
		
		
		'Filter by Customer Categories
		pcv_CustCatFilter=0
		
		query="select idCustomerCategory from pcPPFCustPriceCats where pcPrdPro_id=" & tmpIDCode
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			pcv_Filters=pcv_Filters+1
			pcv_CustCatFilter=1
		end if
		set rs=nothing
		
		if pcv_CustCatFilter=1 then
				
		query="select pcPPFCustPriceCats.idCustomerCategory from pcPPFCustPriceCats, Customers where pcPPFCustPriceCats.pcPrdPro_id=" & tmpIDCode & " and pcPPFCustPriceCats.idCustomerCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			if (pcIncExcCPrice="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCPrice="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing
		
		end if
		'End of Filter by Customer Categories
		
		' Check to see if promotion is filtered by retail or wholesale.
		if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
			pcv_Filters=pcv_Filters+1
			if pcv_wholeSale = "1" and session("customertype") = 1 then
				pcv_FResults=pcv_FResults+1		
			end if 
			if pcv_retail = "1" and session("customertype") <> 1 Then
				pcv_FResults=pcv_FResults+1
			end if    
		end if
		
		if (pcv_Filters=pcv_FResults) AND PrdPromoArr(6,0)<>"" then
			pcf_ProductPromotionMsgQV="<div class='pcPromoMessage'>"
			pcf_ProductPromotionMsgQV=pcf_ProductPromotionMsgQV & PrdPromoArr(6,0)
	    	pcf_ProductPromotionMsgQV=pcf_ProductPromotionMsgQV & "</div>"
		end if
	end if
	set rsQ=nothing
End Function

Public Function pcf_NoShippingTextQV
	if scorderlevel <> "0" then
	else
		' Check to see if the product is set for Free Shipping and display message if product is for sale
		if pnoshipping="-1" and (pFormQuantity <> "-1" or NotForSaleOverride(session("customerCategory"))=1) and pnoshippingtext="-1" then 
			pcf_NoShippingTextQV= "<div class='pcShowProductNoShipping'>"
			pcf_NoShippingTextQV=pcf_NoShippingTextQV & dictLanguage.Item(Session("language")&"_viewPrd_8")
			pcf_NoShippingTextQV=pcf_NoShippingTextQV & "</div>"
		end if
	end if
End Function

Dim strShowBTOQV,strShowBTOQV1

strShowBTOQV=""
strShowBTOQV1=""

Public Sub pcs_GetBTOConfigurationQV
Dim query,rs
	pcv_BTORP=Clng(0)
	strShowBTO=""		
	if pserviceSpec=true then
	 '// Product is BTO
		
		' Get data
		query="SELECT categories.categoryDesc, products.description, products.iRewardPoints,configSpec_products.configProductCategory, configSpec_products.price, configSpec_products.Wprice, categories_products.idCategory, categories_products.idProduct, products.weight, products.pcprod_minimumqty FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		iAddDefaultPrice=Cdbl(0)
		iAddDefaultWPrice=Cdbl(0)
		iAddDefaultPrice1=Cdbl(0)
		iAddDefaultWPrice1=Cdbl(0)
		
		if NOT rs.eof then 
			Dim FirstCnt
			FirstCnt=0
			if intpHideDefConfig="0" then
				strShowBTO= strShowBTO & "<div class='pcShowProductBTOConfig' style='padding-top: 10px; padding-bottom: 2px;'>"
				strShowBTO= strShowBTO & "<b>"&dictLanguage.Item(Session("language")&"_viewPrd_25")&"</b>"
				strShowBTO= strShowBTO & "</div>"
			end if
			do until rs.eof
				FirstCnt=FirstCnt+1
				strCategoryDesc=rs("categoryDesc")
				strDescription=rs("description")
				strConfigProductCategory=rs("configProductCategory")
				dblPrice=rs("price")
				dblWPrice=rs("Wprice")
				intIdCategory=rs("idCategory")
				intIdProduct=rs("idProduct")
				intReward=rs("iRewardPoints")
				if (intReward<>"") and (intReward<>"0") then
				else
				intReward=0
				end if
			
				if intReward="0" then
					query="SELECT pcprod_ParentPrd FROM Products WHERE idproduct=" & intIdProduct & " AND pcprod_ParentPrd>0;"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						pcv_tmpParent=rsQ("pcprod_ParentPrd")
						set rsQ=nothing
						query="SELECT iRewardPoints FROM Products WHERE idproduct=" & pcv_tmpParent & " AND active<>0;"
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							intReward=rsQ("iRewardPoints")
							set rsQ=nothing
							if intReward="" OR IsNull(intReward) then
								intReward=0
							end if
						end if
					end if
				end if
			
				intWeight=rs("weight")
				if Not ((intWeight<>"") and (intWeight<>"0")) then
					intWeight=0
				end if
			
				if intWeight="0" then
					query="SELECT pcprod_ParentPrd FROM Products WHERE idproduct=" & intIdProduct & " AND pcprod_ParentPrd>0;"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						pcv_tmpParent=rsQ("pcprod_ParentPrd")
						set rsQ=nothing
						query="SELECT weight FROM Products WHERE idproduct=" & pcv_tmpParent & " AND active<>0;"
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							intWeight=rsQ("weight")
							set rsQ=nothing
							if intWeight="" OR IsNull(intWeight) then
								intWeight=0
							end if
						end if
					end if
				end if
				
				pcv_iminqty=rs("pcprod_minimumqty")
				if IsNull(pcv_iminqty) or pcv_iminqty="" then
					pcv_iminqty=1
				end if
				if pcv_iminqty="0" then
					pcv_iminqty=1
				end if
				pcv_BTORP=pcv_BTORP+clng(intReward*pcv_iminqty)
				
				dblPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,0)
				dblWPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,1)
				iAddDefaultPrice=Cdbl(iAddDefaultPrice+dblPrice*pcv_iminqty)
				iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+dblWPrice*pcv_iminqty)
				iAddDefaultPrice1=Cdbl(iAddDefaultPrice1+dblPrice1*pcv_iminqty)
				iAddDefaultWPrice1=Cdbl(iAddDefaultWPrice1+dblWPrice1*pcv_iminqty)
				ItemPrice=0
				if Session("CustomerType")=1 then
					if (dblWPrice<>0) then
						ItemPrice=dblWPrice1
					else
						ItemPrice=dblPrice1
					end if
				else
					ItemPrice=dblPrice1
				end if
				if intpHideDefConfig="0" then
					strShowBTOQV= strShowBTOQV & "<div class='pcShowProductBTOConfig'>"
					strShowBTOQV= strShowBTOQV & "<b>"&strCategoryDesc&"</b>: "&strDescription
					strShowBTOQV= strShowBTOQV & "</div>"
				end if
				strShowBTOQV1= strShowBTOQV1 & "<input name=""CAT"&FirstCnt&""" type=""HIDDEN"" value=""CAG"&intIdCategory&""">"
				strShowBTOQV1= strShowBTOQV1 & "<input name=""CAG"&intIdCategory&"QF"" type=""HIDDEN"" value=""" & pcv_iminqty & """>"
				strShowBTOQV1= strShowBTOQV1 & "<input type=""hidden"" name=""CAG"&intIdCategory&""" value="""&intIdProduct&"_0_"&intWeight&"_" & ItemPrice & """>"
				rs.moveNext
			loop			
			strShowBTOQV1= strShowBTOQV1 & "<input type=""hidden"" name=""FirstCnt"" value="""&FirstCnt&""">"
		end if 
		set rs=nothing
	end if
End Sub

Public Function pcf_BTOConfigurationQV
	if (strShowBTOQV<>"") then
		pcf_BTOConfigurationQV = strShowBTOQV
	end if
End Function

Dim tmpQVOptions

tmpQVOptions=""

Public Sub pcs_OptionsNQV

	pcv_TotalOpts=0

	' SELECT DATA SET
	' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
	query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
	query = query & "FROM products "
	query = query & "INNER JOIN ( "
	query = query & "pcProductsOptions INNER JOIN ( "
	query = query & "optionsgroups "
	query = query & "INNER JOIN options_optionsGroups "
	query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
	query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
	query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
	query = query & "WHERE products.idProduct=" & pidProduct &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order, optionsGroups.OptionGroupDesc;"
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	' If we have data	
	if NOT rs.eof then
		pcv_intOptionGroupCount = 0 '// keeps count of the number of options
		xOptionsCnt = 0 '// keeps count of the number of required options
		do until rs.eof				
			
			'if pcv_intOptionGroupCount <= 5  then ' // start limit to 5 options
				'// Get the Group Name
				pcv_strOptionGroupDesc=rs("OptionGroupDesc")
				'// Get the Group ID
				pcv_strOptionGroupID=rs("idOptionGroup")
				'// Is it required
				pcv_strOptionRequired=rs("pcProdOpt_Required")			
		
				'// Start: Do Option Count
				pcv_intOptionGroupCount = pcv_intOptionGroupCount + 1 
				'// End: Do Option Count
				
				'// Get the number of the Option Group
				pcv_strOptionGroupCount = pcv_intOptionGroupCount
				
				'// Start: Do Required Option Count AND generate validation string
				if IsNull(pcv_strOptionRequired) OR pcv_strOptionRequired="" then
						pcv_strOptionRequired=0 '// not required // else it is "1"
				end if			
				if pcv_strOptionRequired=1 then
					
					' Keep Tally
					xOptionsCnt = xOptionsCnt + 1
					
					' Generate String
					if xOtionrequired="1" then
						pcv_strReqOptString = pcv_strReqOptString & ","
					end if
				
					xOtionrequired="1"
					pcv_strOptionGroupDesc2=pcv_strOptionGroupDesc
					pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"'","")
					pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"""","\'\'")
					pcv_strReqOptString = pcv_strReqOptString & "document.additem.idOption" & pcv_strOptionGroupCount & ".selectedIndex,'"& pcv_strOptionGroupDesc2 &"'"
				
				end if
				'// End: Do Required Option Count
			
				'// Make the Option Box
				pcs_makeOptionBoxQV							
		
			'end if ' // end limit to 5 options
		rs.movenext
		loop		
	end if
	set rs=nothing
    tmpQVOptions=tmpQVOptions & "<input type=""hidden"" name=""OptionGroupCount"" value=""" & pcv_intOptionGroupCount & """>"

    if pcv_TotalOpts>0 then
        call CreateStockMsgArea()
    end if

End Sub


Public Sub pcs_makeOptionBoxQV

    Dim rsT,pcv_HaveSPs,pcv_CountOpt
    Dim defaultOpts,tmp1,ik
    
    defaultOpts=""
    query="SELECT pcProd_Relationship FROM Products WHERE pcProd_ParentPrd=" & pidProduct & " AND removed=0 AND active=0 AND pcProd_SPInActive=0 AND pcProd_AppDefault=1;"
    Set rs2=server.createobject("adodb.recordset")
    Set rs2=conntemp.execute(query)
    If Not rs2.eof Then
        tmp1 = rs2("pcProd_Relationship")
        tmp1 = split(tmp1,"_")
        For ik=1 To ubound(tmp1)
            defaultOpts=defaultOpts & tmp1(ik) & "$$"
        Next
    End If
    Set rs2 = Nothing

	' SELECT DATA SET
	' TABLES: options_optionsGroups, options
	query = 		"SELECT options_optionsGroups.InActive, options_optionsGroups.price, options_optionsGroups.Wprice, "
	query = query & "options_optionsGroups.idoptoptgrp, options.idoption, options.optiondescrip "
	query = query & "FROM options_optionsGroups "
	query = query & "INNER JOIN options "
	query = query & "ON options_optionsGroups.idOption = options.idOption "
	query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_strOptionGroupID &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	query = query & "ORDER BY options_optionsGroups.sortOrder, options.optiondescrip;"	
	Set rs2=server.createobject("adodb.recordset")
	Set rs2=conntemp.execute(query)	
	If err.number<>0 Then
		call LogErrorToDatabase()
		set rs2=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	End If

	' If we have data
	If Not rs2.Eof Then
		
		'// clean up the option group description
		If pcv_strOptionGroupDesc<>"" Then
			pcv_strOptionGroupDesc=replace(pcv_strOptionGroupDesc, """", "&quot;")
		End If 
		
		'// START SELECT
		pcv_isOptionSelected="" '// Is this option box selected? Fill variable to "1" during the following loop.
		tmpQVOptions=tmpQVOptions & "<div>" & pcv_strOptionGroupDesc & ":</div>"
        %>

		<% If (pcv_Apparel="1") And (pcv_ApparelRadio="1") Then
			tmpQVOptions=tmpQVOptions & "<div style=""display:none;"">"
			tmpQVOptions=tmpQVOptions & "<div><input type=radio onclick=""javascript:new_CheckOptGroup(" & pcv_strOptionGroupCount & ",0);"" name=""idOption" & pcv_strOptionGroupCount & """ value="""" checked class=""clearBorder""></div>"
			tmpQVOptions=tmpQVOptions & "<div><input name=""idOption" & pcv_strOptionGroupCount & "_0_TXT"" id=""idOption" & pcv_strOptionGroupCount & "_0_TXT"" type=""TEXT"" value=""" & dictLanguage.Item(Session("language")&"_viewPrd_61") & """ size=""" & len(dictLanguage.Item(Session("language")&"_viewPrd_61")) & """ readonly=""readonly"" class=""transparentField""></div>"
			tmpQVOptions=tmpQVOptions & "</div>"
			%>

		<% Else %>

			<%
			tmpQVOptions=tmpQVOptions & "<select autocomplete=""off"" name=""idOption" & pcv_strOptionGroupCount & """ class=""form-control"" style=""margin-top: 3px;"""			
			if pcv_Apparel="1" then
				tmpQVOptions=tmpQVOptions & " onchange=""javascript:new_CheckOptGroup(" & pcv_strOptionGroupCount & ",0);"""
			end if
			tmpQVOptions=tmpQVOptions & ">"
		 	'// Only execute when the Remove Option Feature is activated.
            If pcv_strRemoveFeature<>"1" Then 
                tmpQVOptions=tmpQVOptions & "<option value="""">" & dictLanguage.Item(Session("language")&"_viewPrd_61") & "</option>"
           End If

        End If

        
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
			pcv_CountOpt=0
	
			do until rs2.eof			
				
				OptInActive=rs2("InActive") ' Is it active?
				if IsNull(OptInActive) OR OptInActive="" then
					OptInActive="0"
				end if
				
				dblOptPrice=rs2("price") '// Price
				dblOptWPrice=rs2("Wprice") '// WPrice
				intIdOptOptGrp=rs2("idoptoptgrp") '// The Id of the Option Group
				intIdOption=rs2("idoption") '// The Id of the Option
				strOptionDescrip=rs2("optiondescrip") '// A description of the Option
		
				'**************************************************************************************************
				' START: Dispay the Options
				'**************************************************************************************************
				if OptInActive="0" then
					
                    If session("customerType")=1 then 
						optPrice=dblOptWPrice
					Else
						optPrice=dblOptPrice
					End If %>

                    <%
                    pcv_HaveSPs=0
                    if (pcv_Apparel="1") then
                        
                        query="SELECT idproduct From Products where pcProd_ParentPrd=" & pidProduct & " AND ((pcProd_Relationship like '%[_]" & intIdOptOptGrp & "') OR (pcProd_Relationship like '%[_]" & intIdOptOptGrp & "[_]%')) AND removed=0"
                        query=query & " AND pcProd_SPInActive=0;"
                        set rsT=connTemp.execute(query)
                    
                        if not rsT.eof then
                            pcv_HaveSPs=1
                        end if
                        set rsT=nothing
                        
                    end if '// if (pcv_Apparel="1") then
 
				    if ((pcv_Apparel="1") and (((pcv_ShowStockMsg<>"1") and (pcv_HaveSPs=1)) or ((pcv_ShowStockMsg="1") and (pcv_HaveSPs=1)))) or (pcv_Apparel="0") then
				
                        pcv_CountOpt=pcv_CountOpt+1
				
                        If (pcv_Apparel="1") And (pcv_ApparelRadio="1") Then
							
							tmpQVOptions=tmpQVOptions & "<div id=""Opt_" & intIdOptOptGrp & "_TABLE"">"
				            tmpQVOptions=tmpQVOptions & "<span class=""QVOptL""><input type=radio onclick=""javascript:var isChecked = $(this).attr('is_che');if (isChecked) {$(this).removeAttr('checked');$(this).removeAttr('is_che');document.additem.idOption" & pcv_strOptionGroupCount & ".value='';} else {$(this).attr('checked', 'checked');$(this).attr('is_che', 'true');} new_CheckOptGroup(" & pcv_strOptionGroupCount & ",0);"" name=""idOption" & pcv_strOptionGroupCount & """ value=""" & intIdOptOptGrp & """ class=""clearBorder"" "
                                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                        ' START: Check if Option should be Selected
                                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    
                                        Dim xIdOptCounter
                    
                                        if (tIndex<>0) OR (defaultOpts<>"") then ' Check they are updating the product after adding it to the shopping cart
                    
                                            pcCartArray=session("pcCartSession")
                                            tempIdOpt = ""
                                            if (tIndex<>0) then
                                                tempIdOpt = pcCartArray(tIndex,32)
                                            else
                                                tempIdOpt = defaultOpts
                                            end if
                                            
                                            if tempIdOpt = "" then
												tmpQVOptions=tmpQVOptions &  ">"
                                            else
                                                tempIdOpt = Split(trim(tempIdOpt),"$$")							
                                                for xIdOptCounter = 0 to Ubound(tempIdOpt)-1
                                                    if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
														tmpQVOptions=tmpQVOptions &  " checked"								
													end if
												next
												tmpQVOptions=tmpQVOptions &  ">&nbsp;</span>"
											end if						

                                        else '// if (tIndex<>0) OR (defaultOpts<>"") then
                    
                                            if tIndex="0" then
                    
                                                tempIdOpt = ""
                                                tempIdOpt = request.querystring("idOptionArray")
						
												if tempIdOpt = "" then
													tmpQVOptions=tmpQVOptions &  ">&nbsp;</span>"
												else '// if tempIdOpt = "" then
													tempIdOpt = Split(trim(tempIdOpt),chr(124))
													for xIdOptCounter = 0 to Ubound(tempIdOpt)
														if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then														
                                                            tmpQVOptions=tmpQVOptions &  " checked"								
														end if
													next
													tmpQVOptions=tmpQVOptions &  ">&nbsp;</span>"
                                                end if	'// if tempIdOpt = "" then					
                    
                                            end if '// if tIndex="0" then
                    
                                        end if '// if (tIndex<>0) OR (defaultOpts<>"") then
                                    
                                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                        ' END: Check if Option should be Selected
                                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										tmpQVOptions=tmpQVOptions & "<span class=""QVOptR""><input name=""Opt_" & intIdOptOptGrp & "_TXT"" id=""Opt_" & intIdOptOptGrp & "_TXT"" type=""TEXT"" value=""" & strOptionDescrip & """ size=""" & len(strOptionDescrip) & """ readonly=""readonly"" class=""transparentField""></span>"
										tmpQVOptions=tmpQVOptions & "</div>"
										%>

                        <% ELSE '// IF (pcv_Apparel="1") AND (pcv_ApparelRadio="1") THEN %>
			
							<%
							tmpQVOptions=tmpQVOptions & "<option value=""" & intIdOptOptGrp & """"
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            ' START: Check if Option should be Selected
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        
                            'APP-S
                            if (tIndex<>0) OR (defaultOpts<>"") then ' Check they are updating the product after adding it to the shopping cart
                            'APP-E
                        
                                pcCartArray=session("pcCartSession")
                                tempIdOpt = ""
                            
                                'APP-S
                                if pcv_Apparel="1" then
                                    if (tIndex<>0) then
                                        tempIdOpt = pcCartArray(tIndex,32)
                                    else
                                        tempIdOpt = defaultOpts
                                    end if
                                else
                                    tempIdOpt = pcCartArray(tIndex,11)
                                end if
                                'APP-E
                            
                                if tempIdOpt = "" then                                
									tmpQVOptions=tmpQVOptions & ">"                                    
                                else '// if tempIdOpt = "" then
                                
                                    'APP-S
                                    if pcv_Apparel="1" then
                                        tempIdOpt = Split(trim(tempIdOpt),"$$")
                                    else
                                        tempIdOpt = Split(trim(tempIdOpt),chr(124))
                                    end if
                                    'APP-E
                        
                                    'APP-S
                                    if pcv_Apparel="1" then
                                        for xIdOptCounter = 0 to Ubound(tempIdOpt)-1
                                            if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
												tmpQVOptions=tmpQVOptions & " selected"		
                                            end if
                                        next
                                    else
                                        for xIdOptCounter = 0 to Ubound(tempIdOpt)
                                            if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
												tmpQVOptions=tmpQVOptions & " selected"		
                                            end if
                                        next
                                    end if
                                    'APP-E
									tmpQVOptions=tmpQVOptions & ">"
                                end if '// if tempIdOpt = "" then					
                        
                        
                            else '// if (tIndex<>0) OR (defaultOpts<>"") then
                            
                                'APP-S
                                if (tIndex="0") OR (tIndex="") then
                                'APP-E
                        
                                    tempIdOpt = ""
                                    tempIdOpt = request.querystring("idOptionArray")
						
                                    if (tempIdOpt = "") OR (tempIdOpt = "NULL") then
										tmpQVOptions=tmpQVOptions & ">"
                                    else
                                        tempIdOpt = Split(trim(tempIdOpt),chr(124))
                                        for xIdOptCounter = 0 to Ubound(tempIdOpt)
                                            if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
												tmpQVOptions=tmpQVOptions & " selected"		
                                                pcv_isOptionSelected="1"
                                            end if
                                        next
										tmpQVOptions=tmpQVOptions & ">"
                                    end if '// if (tempIdOpt = "") OR (tempIdOpt = "NULL") then				
                            
                                'APP-S
                                end if  '// if (tIndex="0") OR (tIndex="") then
                                'APP-E
                        
                            'APP-S
                            end if  '// if (tIndex<>0) OR (defaultOpts<>"") then
                            'APP-E
                            
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            ' END: Check if Option should be Selected		
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						

                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: Display Option Name				
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						tmpQVOptions=tmpQVOptions &  strOptionDescrip & "&nbsp;"
						

                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: Display Option Name
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						

                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Display Pricing
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
					if (pcv_Apparel="1") then
						optPrice=0
					end if
                            
                            if optPrice>0 then '// If there is a price thats greater than zero %>
								<%tmpQVOptions=tmpQVOptions & " - " &dictLanguage.Item(Session("language")&"_prodOpt_1")&" "&scCurSign& money(optPrice)%>  
                            <% end if %>
                            
                            <% if optPrice<0 then  '// If there is not a price %>
								<%tmpQVOptions=tmpQVOptions & " - " &dictLanguage.Item(Session("language")&"_prodOpt_2")&" "&scCurSign& money(optPrice)%> 
                            <% end if 
                            
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            ' END: Display Pricing						
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							tmpQVOptions=tmpQVOptions & "</option>"
                        
                        End If '// IF (pcv_Apparel="1") AND (pcv_ApparelRadio="1") THEN


				    end if '// if ((pcv_Apparel="1") and (((pcv_ShowStockMsg="0") and (pcv_HaveSPs=1)) or ((pcv_ShowStockMsg="1")and (pcv_HaveSPs=1)))) or (pcv_Apparel="0") then
				 
                
				end if '// if OptInActive="0" then
				'**************************************************************************************************
				' END: Dispay the Options
				'**************************************************************************************************
				
                rs2.movenext 
			loop
			set rs2=nothing	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			pcv_TotalOpts = pcv_TotalOpts + pcv_CountOpt
			if (pcv_Apparel="1") and (pcv_CountOpt=0) then
				IF (pcv_ApparelRadio="1") THEN
					tmpQVOptions=tmpQVOptions & "<div style=""clear:both"">" & pcv_StockMsg & "</div>"
				ELSE
					tmpQVOptions=tmpQVOptions & "<option """">" & pcv_StockMsg & "</option>"
				END IF
			end if
            %>
        <%
		tmpQVOptions=tmpQVOptions & "</select>"
     
	End If '// If Not rs2.Eof Then         
    %>      
<% End Sub


Public Function pcf_OptionsXQV(CFCount)

    Dim i

    xrequired="0"
    xfieldCnt=0
    xfieldArrCnt=0
    reqstring="" 
    
    dim isArrCount,tmpCount
    isArrCount=0
    tmpCount=0
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '// Start pxfield Array
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    IF intXFCount>=0 THEN
        For i=0 to intXFCount
            'select from the database more info 
            query= "SELECT xfield,textarea,widthoffield,rowlength,maxlength FROM xfields WHERE idxfield="&pcXFArr(0,i)
            set rs=conntemp.execute(query)
        
            if not rs.EOF then '// Check for no field in DB, although the field is referenced by the product
                xField=rs("xfield")
                TextArea=rs("textarea")
                widthoffield=rs("widthoffield")
                rowlength=rs("rowlength")
                maxlength=rs("maxlength")
                set rs=nothing
                
                tmpCount=tmpCount+1
                pxreq=pcXFArr(1,i)
                
                if pxreq="-1" then
                    xfieldCnt=xfieldCnt+1
                    xrequired="1"
                    if reqstring<>"" then
                        reqstring=reqstring & ","
                    end if
                    reqstring=reqstring&"additem.xfield" & tmpCount & ".value,'"&replace(xfield,"'","\'")&"'"
                end if
                %>
                
                <%pcf_OptionsXQV=pcf_OptionsXQV & "<input type=""hidden"" name=""xf" & tmpCount & """ value=""" & pcXFArr(0,i) & """>"%>
                <%pcf_OptionsXQV=pcf_OptionsXQV & xField%>
                
                <% if TextArea="-1" then
                    pcf_OptionsXQV=pcf_OptionsXQV & "<br>"
                    pcf_OptionsXQV=pcf_OptionsXQV & "<textarea name=""xfield" & tmpCount & """ cols=""" & widthoffield & """ rows=""" & rowlength & """ style=""margin-top: 6px"""
                    if maxlength>"0" then
                        pcf_OptionsXQV=pcf_OptionsXQV & "onkeyup=""javascript:testchars(this,'" & tmpCount & "'," & maxlength & ");"""
                    end if
                    pcf_OptionsXQV=pcf_OptionsXQV & ">"
                    pcf_OptionsXQV=pcf_OptionsXQV & "</textarea>"
                    if maxlength>"0" then
                        pcf_OptionsXQV=pcf_OptionsXQV & "<br>"
                        pcf_OptionsXQV=pcf_OptionsXQV & dictLanguage.Item(Session("language")&"_GiftWrap_5a")
                        pcf_OptionsXQV=pcf_OptionsXQV & "<span id=""countchar" & tmpCount & """ name=""countchar" & tmpCount & """ style=""font-weight: bold"">" & maxlength & "</span>"
                        pcf_OptionsXQV=pcf_OptionsXQV & " " & dictLanguage.Item(Session("language")&"_GiftWrap_5b")
                        pcf_OptionsXQV=pcf_OptionsXQV & "<br><br>"
                    end if%>
                <% else 
                    pcf_OptionsXQV=pcf_OptionsXQV & "<br>"
                    pcf_OptionsXQV=pcf_OptionsXQV & "<input type=""text"" name=""xfield" & tmpCount & """ size=""" & widthoffield & """ maxlength=""" & maxlength & """ style=""margin-top: 6px"""
                    pcf_OptionsXQV=pcf_OptionsXQV & " value="""""
                    if maxlength>"0" then
                        pcf_OptionsXQV=pcf_OptionsXQV & " onkeyup=""javascript:testchars(this,'" & tmpCount & "'," & maxlength & ");"""
                    end if
                    pcf_OptionsXQV=pcf_OptionsXQV & ">"
                    if maxlength>"0" then
                        pcf_OptionsXQV=pcf_OptionsXQV & "<br>"
                        pcf_OptionsXQV=pcf_OptionsXQV & dictLanguage.Item(Session("language")&"_GiftWrap_5a")
                        pcf_OptionsXQV=pcf_OptionsXQV & "<span id=""countchar" & tmpCount & """ name=""countchar" & tmpCount & """ style=""font-weight: bold"">" & maxlength & "</span>"
                        pcf_OptionsXQV=pcf_OptionsXQV & " " & dictLanguage.Item(Session("language")&"_GiftWrap_5b")
                        pcf_OptionsXQV=pcf_OptionsXQV & "<br><br>"
                    end if%>
                <% end if %>
            <%end if ' rs.eof
        Next
    END IF
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '// End pxfield Array
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    CFCount = tmpCount
    
    pcf_OptionsXQV=pcf_OptionsXQV & "<input type=""hidden"" name=""XFCount"" value=""" & tmpCount & """ />" & vbcrlf
    
    if tmpCount>0 then
    pcf_OptionsXQV=pcf_OptionsXQV & "<br><br>"
    pcf_OptionsXQV=pcf_OptionsXQV & "<script type=text/javascript>" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "function testchars(tmpfield,idx,maxlen)" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "{" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "var tmp1=tmpfield.value;" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "if (tmp1.length>maxlen)" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "{" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "alert(""" & dictLanguage.Item(Session("language")&"_CheckTextField_1") & """ + maxlen + """ & dictLanguage.Item(Session("language")&"_CheckTextField_1a") & """);" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "tmp1=tmp1.substr(0,maxlen);" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "tmpfield.value=tmp1;" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "document.getElementById(""countchar"" + idx).innerHTML=maxlen-tmp1.length;" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "tmpfield.focus();" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "}" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "document.getElementById(""countchar"" + idx).innerHTML=maxlen-tmp1.length;" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "}" & vbcrlf
    pcf_OptionsXQV=pcf_OptionsXQV & "</script>"
    end if		
End Function


Public Function pcf_QtyDiscountsQV

	Dim DispSubDisc
	DispSubDisc=1
	
	'--> check for discount per quantity
	if (pDiscountPerQuantity=0) AND (pcv_Apparel="1") then
		query="SELECT idDiscountperquantity FROM discountsperquantity INNER JOIN Products ON discountsperquantity.idProduct=Products.idProduct WHERE Products.pcProd_ParentPrd=" &pidProduct
		set rsQ=server.CreateObject("ADODB.RecordSet")
		set rsQ=conntemp.execute(query)
	
		if not rsQ.eof then
			pDiscountPerQuantity=-1
		else
			pDiscountPerQuantity=0
		end if
		set rsQ=nothing
	end if

	if pDiscountPerQuantity=-1 then
		'if customer is retail, check if there are discounts with retail <> 0
		VardiscGo=0
		if session("customerType")=1 then
			query="SELECT discountPerWUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerWUnit>0"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if rs.eof then
				VardiscGo=1
			end if
			set rs=nothing
		else
			query="SELECT discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerUnit>0"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if rs.eof then
				VardiscGo=1
			end if
			set rs=nothing
		end if
	
		if VardiscGo=0 then
			query="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" ORDER BY num"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query) 
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if NOT rs.eof then '// Quick Loop - there will not be too many discounts
				pcv_intTotalDiscounts = 0
				do until rs.eof
					pcv_intTotalDiscounts=pcv_intTotalDiscounts+1
				rs.moveNext		
				loop
				rs.moveFirst
			end if
					
						pcf_QtyDiscountsQV="<div class=""QVDiscContainer"">"
                        pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""pcClear""></div>"
                        pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderWrapper"">"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderLeft"">" & dictLanguage.Item(Session("language")&"_pricebreaks_1") & "</div>"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderRight"">" & dictLanguage.Item(Session("language")&"_pricebreaks_2") & "<img align=""right"" style=""vertical-align: middle"" src=""" & pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount")) & """ alt=""" & dictLanguage.Item(Session("language")&"_altTag_6") & """></div>"
                        pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
							
							pc_intCounterQ = 0
							do until rs.eof
								pc_intCounterQ = pc_intCounterQ + 1 '// count Discount Rows
								dblQuantityFrom=rs("quantityFrom")
								dblQuantityUntil=rs("quantityUntil")
								dblPercentage=rs("percentage")
								dblDiscountPerWUnit=rs("discountPerWUnit")
								dblDiscountPerUnit=rs("discountPerUnit")
								%>
									<% if dblQuantityFrom=dblQuantityUntil then
										pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscCellLeft"">"
										pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	dblQuantityUntil & " " & dictLanguage.Item(Session("language")&"_pricebreaks_4")
										pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
									else
										pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscCellLeft"">"
										pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	dblQuantityFrom & " " & dictLanguage.Item(Session("language")&"_pricebreaks_3") & " " & dblQuantityUntil & " " & dictLanguage.Item(Session("language")&"_pricebreaks_4")
										pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
									end if
									pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscCellRight"">"
											If session("customerType")=1 Then
												If dblPercentage="0" then
													pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	scCurSign & money(dblDiscountPerWUnit)
												else
													pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	dblDiscountPerWUnit & "%"
												End If
												else
												If dblPercentage="0" then
													pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	scCurSign & money(dblDiscountPerUnit)
												else
													pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	dblDiscountPerUnit & "%"
												End If
												end If
											
										pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
			
						
							rs.moveNext		
							loop
							set rs=nothing
                        pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""pcClear""></div>"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
						


		else
			if pcv_Apparel="1" then
				if (DispSubDisc=0) OR (DispSubDisc="") then
					pcf_QtyDiscountsQV="<div>"
					pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderAll"">" & dictLanguage.Item(Session("language")&"_viewPrd_spmsg8") & "</div>"
					pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & "<div class=""QVDiscCellAll""><img src=""" & pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount")) & """ hspace=""2""><a href=""javascript:openbrowser('app-subPrdDiscount.asp?idproduct=" & pIdProduct & "');"">" & dictLanguage.Item(Session("language")&"_viewPrd_spmsg8a") & "</a></div>"
					pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
				else
					query="select idproduct,description from Products where pcprod_ParentPrd=" & pIDProduct
					set rsA=connTemp.execute(query)
					
					if not rsA.eof then
						pcf_QtyDiscountsQV="<div>"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderAll"">" & dictLanguage.Item(Session("language")&"_viewPrd_spmsg8") & "</div>"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
					end if
					do while not rsA.eof
					pcv_sprdID=rsA("idproduct")
					pcv_prdName=rsA("description")
				
					queryq="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pcv_sprdID &" ORDER BY num"
					set rsTmp=conntemp.execute(queryq)
					if not rsTmp.eof then
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & "<div>"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderAll"">" & pcv_prdName & "</div>"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderLeft"">"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	dictLanguage.Item(Session("language")&"_pricebreaks_1")
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscHeaderRight"">"
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	dictLanguage.Item(Session("language")&"_pricebreaks_2")
						pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
						do until rstmp.eof
								if rstmp("quantityFrom")=rstmp("quantityUntil") then
									pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscCellLeft"">"
									pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	rstmp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")
									pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
								else
									pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscCellLeft"">"
									pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	rstmp("quantityFrom")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_3")&"&nbsp;"&rstmp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")
									pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"</div>"
								end if
								pcf_QtyDiscountsQV=pcf_QtyDiscountsQV &	"<div class=""QVDiscCellRight"">"
								If (request.querystring("Type")="1")  or (session("CustomerType")="1") Then %>
									<% If rstmp("percentage")="0" then %>
										<%pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & scCurSign & money(rstmp("discountPerWUnit"))%> 
									<% else %>
										<%pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & rstmp("discountPerWUnit") & "%"%>
									<% End If %>
								<% else %>
									<% If rstmp("percentage")="0" then %>
										<%pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & scCurSign & money(rstmp("discountPerUnit"))%> 
									<% else %>
										<%pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & rstmp("discountPerUnit") & "%"%>
									<% End If %>
								<% end If%>
								<%pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & "</div>"%>
							<% rstmp.moveNext
						loop%>
						<%pcf_QtyDiscountsQV=pcf_QtyDiscountsQV & "</div>"%>
					<%end if
					set rstmp=nothing
					rsA.MoveNext
					loop
					set rsA=nothing
				end if
			end if
		end if
	end if
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: QuickView Functions
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>