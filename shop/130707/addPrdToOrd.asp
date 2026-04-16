<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
response.Buffer=true 
pageTitle="Add Product to an Existing Order"
%>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../pc/inc_AddThis.asp"-->
<% 
Dim tIndex, tUpdPrd, pIdCategory, strBreadCrumb, pIdProduct, dblpcCC_Price
Dim pcv_strViewPrdStyle, pcv_strFormAction, pcv_intValidationFile, pcv_blnBTOisConfig, iRewardPoints, pDescription, pMainProductName
Dim pSku, pconfigOnly, pserviceSpec, pPrice, pBtoBPrice, pDetails, pListPrice, plistHidden, pimageUrl, pLgimageURL
Dim pArequired, pBrequired, pStock, pWeight, pEmailText, pFormQuantity, pnoshipping, pcustom1, pcontent1
Dim pcustom2, pcontent2, pcustom3, pcontent3, pnoprices
Dim pIDBrand, psDesc, pNoStock, pnoshippingtext, intIdProduct, intWeight, optionA, optionB
Dim pcv_intHideBTOPrice, pcv_intQtyValidate, pcv_lngMinimumQty, intpHideDefConfig, pnotax, BrandName
Dim FirstCnt, strDescription, intReward, pcv_BTORP, strConfigProductCategory, dblPrice, dblWPrice, intIdCategory
Dim VardiscGo, dblQuantityFrom, dblQuantityUntil, dblPercentage, dblDiscountPerWUnit, dblDiscountPerUnit
Dim intIdOptOptGrp, intIdOption, strOptionDescrip, OptInActive, optPrice, tempIdOptA, tempIdOptB
Dim xrequired, xfieldCnt, reqstring, TextArea, widthoffield, rowlength
Dim scCS, cs_showprod, cs_showcart, cs_showimage, crossSellText
Dim pcv_strOptionGroupDesc, pcv_intOptionGroupCount, pcv_strOptionGroupCount, pcv_strOptionGroupID, pcv_strOptionRequired
Dim xOptionsCnt, pcv_strNumberValidations, pcv_strFuntionCall, pcv_strReqOptString, xOtionrequired, pcv_strCSDiscounts , pcv_strPrdDiscounts 
Dim pcv_strProdImage_Url, pcv_strProdImage_LargeUrl, pcv_intProdImage_Columns, pcv_strShowImage_LargeUrl, pcv_strShowImage_Url, pcv_strCurrentUrl
Dim pcv_strAdditionalImages, cCounter, pcv_strWishListLink, BTOCharges, pcv_strCSString, pcv_strReqCSString, cs_RequiredIds, xCSCnt

Dim pcv_ReorderLevel,pcv_Apparel,pcv_ShowStockMsg,pcv_StockMsg,pcv_SizeLink,pcv_SizeInfo,pcv_SizeImg,pcv_SizeURL,pcv_ApparelRadio,pcv_HaveSPs
Dim pcv_TotalOpts
Dim popUpAPP

pIdProduct=request.QueryString("idProduct")
pIdOrder=request.QueryString("ido")

if trim(pIdProduct)="" or IsNumeric(pIdProduct)=false then
   call closeDb()
response.redirect "msg.asp?message=85"
end if


'Change paths of images
pcv_tmpNewPath="../pc/"

'--> open database connection


	query="SELECT customers.customerType,customers.idCustomerCategory FROM customers INNER JOIN orders ON customers.idcustomer=orders.idcustomer WHERE idorder=" & pIdOrder & ";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if not rs.eof then
		session("customerType")=rs("customerType")
		idcustomerCategory=rs("idcustomerCategory")
		if IsNull(idcustomerCategory) or idcustomerCategory="" then
			idcustomerCategory=0
		end if
	end if
	set rs=nothing

	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory="&idcustomerCategory&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then
		session("customerCategory")=rs("idcustomerCategory")
		strpcCC_Name=rs("pcCC_Name")
		session("customerCategoryDesc")=strpcCC_Name
		strpcCC_Description=rs("pcCC_Description")
		session("customerCategoryType")=rs("pcCC_CategoryType")
		if session("customerCategoryType")="ATB" then
			session("ATBCustomer")=1
			session("ATBPercentage")=rs("pcCC_ATB_Percentage")
			intpcCC_ATB_Off=rs("pcCC_ATB_Off")
			if intpcCC_ATB_Off="Retail" then
				session("ATBPercentOff")=0
			else
				session("ATBPercentOff")=1
			end if
		else
			session("ATBCustomer")=0
			session("ATBPercentage")=0
			session("ATBPercentOff")=0
		end if
	end if
	set rs=nothing
%>
<!--#include file="../pc/viewPrdCode.asp"-->
<%
' --> check for discount per quantity
query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct

set rs=conntemp.execute(query)

if err.number <> 0 then
	call LogErrorToDatabase()
	set rs=nothing
	
	call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

dim pDiscountPerQuantity
if not rs.eof then
 pDiscountPerQuantity=-1
else
 pDiscountPerQuantity=0
end if

' --> gets product details from db

query="SELECT iRewardPoints, description, sku, configOnly, serviceSpec, price, btobprice, listprice, listHidden, imageurl, largeImageURL, stock, weight, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3, noprices,IDBrand,noshippingtext,nostock,pcprod_HideDefConfig, notax, pcProd_BackOrder,pcProd_ShipNDays,pcProd_SkipDetailsPage,pcProd_ReorderLevel,pcprod_Apparel,pcprod_ShowStockMsg,pcprod_StockMsg,pcprod_SizeLink,pcprod_SizeInfo,pcprod_SizeImg,pcprod_SizeURL,pcProd_ApparelRadio,details, sDesc FROM products WHERE idProduct=" &pidProduct& " AND active=-1"
set rs=conntemp.execute(query)
if err.number <> 0 then
	call LogErrorToDatabase()
	set rs=nothing	
	call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
	set rs = nothing
	 
  	call closeDb()
response.redirect "msg.asp?message=34"
end if

' --> set product variables <---
iRewardPoints=rs("iRewardPoints")
pDescription=rs("description")
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pDetails=rs("details")
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
pLgimageURL=rs("largeImageURL")
pStock=rs("stock")
pWeight=rs("weight")
pEmailText=rs("emailText")
pFormQuantity=rs("formQuantity")
pnoshipping=rs("noshipping")
pcustom1=rs("custom1")
pcontent1=rs("content1")
pcustom2=rs("custom2")
pcontent2=rs("content2")
pcustom3=rs("custom3")
pcontent3=rs("content3")
pnoprices=rs("noprices")
pIDBrand=rs("IDBrand")
pnoshippingtext=rs("noshippingtext")
pnostock=rs("nostock")
intHideDefConfig=rs("pcprod_HideDefConfig")

pcv_ReorderLevel=rs("pcProd_ReorderLevel")
if IsNull(pcv_ReorderLevel) or pcv_ReorderLevel="" then
	pcv_ReorderLevel=0
end if
pcv_Apparel=rs("pcprod_Apparel")
if pcv_Apparel="1" then
	pcv_ShowStockMsg=rs("pcprod_ShowStockMsg")
	if IsNull(pcv_ShowStockMsg) or pcv_ShowStockMsg="" then
		pcv_ShowStockMsg="0"
	end if
	pcv_StockMsg=rs("pcprod_StockMsg")
	if IsNull(pcv_StockMsg) or pcv_StockMsg="" then
		pcv_StockMsg=dictLanguage.Item(Session("language")&"_viewPrd_7")
	end if
	pcv_SizeLink=rs("pcprod_SizeLink")
	pcv_SizeInfo=rs("pcprod_SizeInfo")
	pcv_SizeImg=rs("pcprod_SizeImg")
	pcv_SizeURL=rs("pcprod_SizeURL")
	if pcv_SizeURL<>"" then
		if ucase(pcv_SizeURL)="HTTP://" then
			pcv_SizeURL=""
		end if
	end if
	pcv_ApparelRadio=rs("pcprod_ApparelRadio")
	if IsNull(pcv_ApparelRadio) or pcv_ApparelRadio="" then
		pcv_ApparelRadio="0"
	end if
end if

pDetails=replace(rs("details"),"&quot;",chr(34))
psDesc=rs("sDesc")

'Disregard Stock option is always checked
if pcv_Apparel="1" then
	pNoStock=1
end if

set rs=nothing

'// Check sub-products discounts
IF (pcv_Apparel="1") then

	if session("customerType")=1 then
		query="SELECT discountsperquantity.discountPerWUnit FROM products,discountsperquantity WHERE products.pcprod_ParentPrd=" & pIDProduct & " and products.pcProd_SPInActive=0 and discountsperquantity.idProduct=products.idproduct AND discountsperquantity.discountPerWUnit>0"
		set rsTempQ=conntemp.execute(query)
		if not rsTempQ.eof then
			pDiscountPerQuantity=-1
		end if
	else
		query="SELECT discountsperquantity.discountPerUnit FROM products,discountsperquantity WHERE products.pcprod_ParentPrd=" & pIDProduct & " and products.pcProd_SPInActive=0 and discountsperquantity.idProduct=products.idproduct AND discountsperquantity.discountPerUnit>0"
		set rsTempQ=conntemp.execute(query)
		if not rsTempQ.eof then
			pDiscountPerQuantity=-1
		end if
	end if

END IF

if intHideDefConfig<>"" then
else
	intHideDefConfig="0"
end if	

if pnoprices=1 then
	pPrice=0
	pBtoBPrice=0
	pListPrice=0
end if

'Get XFields
Dim pcXFArr,intXFCount
intXFCount=-1
query="SELECT IdXField,pcPXF_XReq FROM pcPrdXFields WHERE idProduct=" & pidProduct & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	pcXFArr=rs.getRows()
	intXFCount=ubound(pcXFArr,2)
end if
set rs=nothing

if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
	query="select BrandName from Brands where IDBrand=" & pIDBrand
	set rstemp4=connTemp.execute(query)
	if not rstemp4.eof then
		BrandName=rstemp4("BrandName")
	end if
end if
%>

<!--#include file="adminheader.asp"-->
<!-- 58eed21a4b2adf0316e95c5c4ee68f13 -->
<link type="text/css" rel="stylesheet" href="../pc/css/pcStorefront.css" />
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<!--#include file="../includes/javascripts/pcValidateViewPrd.asp"-->
<!-- Start Form -->
<% If statusAPP="1" OR scAPP=1 Then %>
	<script>
	<%if pImageUrl<>"" then%>
	var GeneralImg="<%=pImageUrl%>";
	<%else%>
	var GeneralImg="no_image.gif";
	<%end if%>
	<%if pLgimageURL<>"" then%>
	var DefLargeImg="<%=pLgimageURL%>";
	var LargeImg=DefLargeImg
	<%else%>
	var DefLargeImg="";
	var LargeImg=DefLargeImg
	<%end if%>
	</script>
	<%
	IF (pcv_Apparel="1") then
		call GenApparelSubProducts()
	END IF
End If
%>
<% 
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' GENERATE FORM																						//
' > Configurable / Standard Product															//
' > Each uses a different form action and JavaScript validation function                            //
'/////////////////////////////////////////////////////////////////////////////////////////////////////

'********************************************************************
' VALIDATION FILE
' pcv_intValidationFile = 1 // Configurable
' pcv_intValidationFile = 2 // Standard
'
' FORM ACTION
' pcv_strFormAction = "instConfiguredPrd.asp" // Configurable (setup)
' pcv_strFormAction = "instPrd.asp" // Configurable NON setup and Standard
'********************************************************************
pcv_blnBTOisConfig = pcf_BTOisConfig '// returns true or false for Configured

If pserviceSpec = "False" Then
	pserviceSpec = 0
End If

If pserviceSpec<>0 Then '// If its Configurable Then
	if pcv_blnBTOisConfig then '// if its configured then
		pcv_strFormAction = "bto_instConfiguredPrd.asp"
		pcv_intValidationFile = 1
	else '// Its not configured
		pcv_strFormAction = "instPrdToOrd.asp"
		pcv_intValidationFile = 1
	end if
else '// Its standard
	pcv_strFormAction = "instPrdToOrd.asp"
	pcv_intValidationFile = 2
end if
%>
<div id="pcMain">
<div class="pcMainContent" itemscope itemtype="http://schema.org/Product">
<form method="post" action="<%=pcv_strFormAction%>" name="additem" onsubmit="return checkproqty(document.additem.quantity);" class="pcForms">
<input name="idorder" type="hidden" value="<%=pIdOrder%>">
<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->
		<%
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  Show product name 
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcs_ProductName
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END:  Show product name 
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>
<div id="pcViewProductC" class="pcViewProduct">
	<div class="pcViewProductLeft">
	<%
	'*****************************************************************************************************
	' 2) GENERAL INFORMATION
	'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show SKU
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ShowSKU	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show SKU
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Weight (If admin turned on)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_DisplayWeight
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Weight
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Brand (If assigned)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ShowBrand
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Brand
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Units in Stock (if on, show the stock level here)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_UnitsStock
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Units in Stock
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'*****************************************************************************************************
	' END GENERAL INFORMATION
	'*****************************************************************************************************
	%>
	
	<br />
	
	<%	
	
	'*****************************************************************************************************
	' 5) DESCRIPTION
	'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Product Description
	'   >  If there is a short description, show it and link to the long description below.
	'   >  Otherwise show the long description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ProductDescription
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Product Description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'*****************************************************************************************************
	' END DESCRIPTION
	'*****************************************************************************************************
	%>
	
	<br />
	
	<%
	'*****************************************************************************************************
	' 6) DEFAULT CONFIGURATION (CONFIGURATOR)
	'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Default Configuration
	'   >  If this is a configurable product, then gather information about default configuration
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_BTOConfiguration
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Defualt Configuration
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'*****************************************************************************************************
	' END DEFAULT CONFIGURATION (CONFIGURATOR)
	'*****************************************************************************************************
	%>
	
	<%
	'*****************************************************************************************************
	' 8) CUSTOM SEARCH FIELDS
	'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Custom Search Fields
	'   >  Check to see if the product has been assigned Custom Search Fields. If so, display the values
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_CustomSearchFields
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Custom Search Fields
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'*****************************************************************************************************
	' END CUSTOM SEARCH FIELDS
	'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Reward Points
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_RewardPoints
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Reward Points
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show product prices
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
	pcs_ProductPrices
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show product prices
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Free Shipping Text
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	pcs_NoShippingText
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Free Shipping Text
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  CONFIGURATOR ADDON S 0r E
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_BTOADDON
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  CONFIGURATOR ADDON S 0r E
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Out of Stock Message
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_OutStockMessage
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Out of Stock Message
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	%>
	
	<%'Start SDBA%>
	<!-- Start Back-Order Message -->
	<%pcs_DisplayBOMsg%>
	<!-- End Back-Order Message -->
	<%'End SDBA%>
	
	<%
	popUpAPP=0

	'*****************************************************************************************************
	' 7) PRODUCT OPTIONS
	'*****************************************************************************************************
		response.write "<br />"
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Options A,B
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if pcf_VerifyShowOptions then '// IF [price =0 and configurable] DO NOT show Options	
			
			'/////////////////////////////////////////////////////////////
			'//      ORDERING OPTIONS									//
			'/////////////////////////////////////////////////////////////
			
			'*************************************************************
			' START: Options
			'*************************************************************
			pcs_OptionsN
			'*************************************************************
			' END: Options
			'*************************************************************
					
		end if  
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Options A,B
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
	'*****************************************************************************************************
	' END PRODUCT OPTIONS
	'*****************************************************************************************************
	
	
	'*****************************************************************************************************
	' 9) CUSTOM INPUT FIELDS
	'*****************************************************************************************************
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Options X
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pcf_VerifyShowOptions then '// IF [price =0 and configurable] DO NOT show Custom Fields	
		
		'/////////////////////////////////////////////////////////////
		'//      CUSTOM INPUT FIELDS								//
		'/////////////////////////////////////////////////////////////
		If statusAPP="1" OR scAPP=1 Then
			Dim rsDetailsObj
			pIdProductOrdered=request("pIdPrdOrd")
			if len(pIdProductOrdered)>0 then
				query="SELECT ProductsOrdered.xfdetails FROM ProductsOrdered WHERE idProductOrdered=" & pIdProductOrdered
				set rsDetailsObj=server.CreateObject("ADODB.RecordSet")
				set rsDetailsObj=connTemp.execute(query)
				if NOT rsDetailsObj.eof then
					tempIdOpt=rsDetailsObj("xfdetails")
				end if
				set rsDetailsObj=nothing
			end if
		End If
		
		'*************************************************************
		' START: Options X
		'*************************************************************
		pcs_OptionsX							
		'*************************************************************
		' END: Options X
		'*************************************************************
		
	end if  
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Options x
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'*****************************************************************************************************
	' END CUSTOM INPUT FIELDS
	'*****************************************************************************************************
	%>
	
	</div><!--end left-->
	<div class="pcViewProductRight">

	<% 

	'*****************************************************************************************************
	' 4) PRODUCT IMAGES
	'*****************************************************************************************************

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Product Image (If there is one)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	pcs_ProductImage%>
	<script>
		var pcv_hasAdditionalImages = false;
		var pcv_strIsMojoZoomEnabled = <% If pcv_IntMojoZoom="1" And (Not pcv_IsQuickView = True And Not Session("Mobile") = "1") Then Response.Write "true" Else Response.Write "false" End If %>;
		var pcv_strMojoZoomOrientation = "<%= pcv_strMojoZoomOrientation %>";
		var pcv_strUseEnhancedViews = <% If pcv_strUseEnhancedViews Then Response.Write "true" Else Response.Write "false" End If %>;
		<% if pcv_strUseEnhancedViews = True then %>
			var CurrentImg=1;
		<% End If %>
	</script>
	<%if pcv_strUseEnhancedViews = True then
	%>
		<script type=text/javascript>	
			hs.align = '<%=pcv_strHighSlide_Align%>';
			hs.transitions = [<%=pcv_strHighSlide_Effects%>];
			hs.outlineType = '<%=pcv_strHighSlide_Template%>';
			hs.fadeInOut = <%=pcv_strHighSlide_Fade%>;
			hs.dimmingOpacity = <%=pcv_strHighSlide_Dim%>;			
			//hs.numberPosition = 'caption';
			<% if bCounter>0 then %>
				if (hs.addSlideshow) hs.addSlideshow({
					slideshowGroup: 'slides',
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
			function pcf_initEnhancement(ele,img) {
				if (document.getElementById('1')==null) {
					hs.expand(ele, { src: img, minWidth: <%=pcv_strHighSlide_MinWidth%>, minHeight: <%=pcv_strHighSlide_MinHeight%> }); 
				} else {
					document.getElementById('1').onclick();			
				}
			}
		</script>
		
	<% end if   
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Product Image (If there is one)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'*****************************************************************************************************
	' END PRODUCT IMAGES
	'*****************************************************************************************************
	
	response.write "<br />"
	response.write "<br />"
	
	'*****************************************************************************************************
	' 15) QUANTITY DISCOUNTS ZONE
	'*****************************************************************************************************

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show quantity discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	err.clear
	pcs_QtyDiscounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show quantity discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'*****************************************************************************************************
	' END QUANTITY DISCOUNTS ZONE
	'*****************************************************************************************************
	%>
	</div><!--right-->
	<div class="pcClear"></div>     

	<div class="pcViewProductBottom">
		

	<!-- Start Quantity and Add to Cart -->
		<% 	
	'*****************************************************************************************************
	' 3) QUANTITY AND ADD TO CART
	'*****************************************************************************************************
		if pFormQuantity="-1" then
		'/////////////////////////////////////
		'// Product NOT For Sale 			//
		'/////////////////////////////////////
			if pEmailText<>"" then 
				response.write "<div class=pcShowProductNFS>" 
				response.write pEmailText '// reason why it's not for sale
				response.write "</div>" 
			end if
					
		else 
		'/////////////////////////////////////
		'// Product For Sale				//
		'/////////////////////////////////////
		
			' 2a) Check for order level permission		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' NOTES:
			' Check for order level permission "scorderlevel".
			' scorderlevel = 0 // everybody
			' scorderlevel = 1 // wholesale only
			' scorderlevel = 2 // catalog only
			
			' Also check what level the current customer is classified.
			' session("customerType") = "" // not logged in
			' session("customerType") = 1  // wholesale
			' session("customerType") = 0  // retail
			
			' Verify level is 0 OR is 1 with a custmer type of 1	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
			' 2b) If out of stock AND out of stock purchase is allowed show button.
				
				tIndex=0
				if pcf_OutStockPurchaseAllow then
				' // out of stock purchase is allowed
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: Show CUSTOMIZE BUTTON or ADD TO CART
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						If ((pserviceSpec<>0) AND ((pnoprices>0) OR (pPrice=0) OR (scConfigPurchaseOnly=1))) or ((iBTOQuoteSubmitOnly=1) and (pserviceSpec<>0)) then 
						' // customize button only						
						
								'/////////////////////////////////////////////////////////////
								'//      CUSTOMIZE BUTTON									//
								'/////////////////////////////////////////////////////////////							 
								'*************************************************************
								' START: Customize Button Only
								'*************************************************************
								pcs_CustomizeButton
								'*************************************************************
								' END: Customize Button Only
								'*************************************************************
								
						else 
						' // show add to cart
						
								'/////////////////////////////////////////////////////////////
								'//      ADD TO CART										//
								'/////////////////////////////////////////////////////////////							 
								'*************************************************************
								' START: Add to Cart
								'*************************************************************
								pcs_AddtoCart
								'*************************************************************
								' END: Add to Cart
								'*************************************************************
						end if 
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: Show CUSTOMIZE BUTTON or ADD TO CART
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				end if ' end 2b			
		end if ' end if pFormQuantity="-1" then 	
	'*****************************************************************************************************
	' END QUANTITY AND ADD TO CART
	'*****************************************************************************************************
		%>
	<!-- End Quantity and Add to Cart -->

<%
'*****************************************************************************************************
' 12) LONG DESCRIPTION
'*****************************************************************************************************

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Display long product description if there is a short description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_LongProductDescription
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Display long product description if there is a short description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END LONG DESCRIPTION
'*****************************************************************************************************
	%>
	
	</div>
</div>
<!--#include file="../includes/javascripts/pcValidateFormViewPrd.asp"-->
<% if pcv_Apparel="1" then %>
	<input type=hidden name="Apparel" value="1">
	<input type=hidden name="SavedList" value="">
	<input type=hidden name="SavedQtyList" value="">
	<script>
		var firstRun=0;
		function addEvent( obj, type, fn )
		{ 
			if ( obj.attachEvent ) { 
			obj['e'+type+fn] = fn; 
			obj[type+fn] = function(){obj['e'+type+fn]( window.event );} 
			obj.attachEvent( 'on'+type, obj[type+fn] ); 
			} else 
			obj.addEventListener( type, fn, false ); 
		} 

		function myInitFunction() {
			if (firstRun==0)
			{
			new_CheckOptGroup(1,0);
			firstRun=1;
			}
		}
		
		addEvent(window,'load',myInitFunction);
	</script>
	<%if request.querystring("SubPrd")<>"" then%>
	<input type=hidden name="SubPrd" value="<%=request.querystring("SubPrd")%>">
	<%
		tmpSubID=getUserInput(request.querystring("SubPrd"),0)
		if IsNumeric(tmpSubID) then
			call opendb()
			query="SELECT pcProd_Relationship FROM Products WHERE idproduct=" & tmpSubID
			set rs=connTemp.execute(query)
			if not rs.eof then
				tmpRelationship=rs("pcProd_Relationship")
				if (tmpRelationship<>"") and (instr(tmpRelationship,"_")>0) then
					tmpArray=split(tmpRelationship,"_")
					%>
					<script>
					<%if pcv_ApparelRadio="1" then%>
					function SetRadioValue(tmpID,tmpValue)
					{
						var els = document.additem.elements; 
						for(i=0; i<els.length; i++)
						{ 
							if ((els[i].name==tmpID) && (eval(els[i].value)==eval(tmpValue)) && (els[i].type=="radio"))
							{
								els[i].checked=true;
								break;
							}
						}
					}
					<%end if%>
					<%For mk=1 to ubound(tmpArray)%>
						<%if pcv_ApparelRadio="1" then%>
							SetRadioValue("idOption<%=mk%>","<%=tmpArray(mk)%>");
						<%else%>
							document.additem.idOption<%=mk%>.value=<%=tmpArray(mk)%>;
						<%end if%>
					<%Next%>
						firstRun=1;
						new_CheckOptGroup(1,0);
					</script>
					<%
				end if
			end if
			set rs=nothing
		end if
	else
		if request.querystring("index")<>"" then%>
		<script>
			firstRun=1;
			new_CheckOptGroup(1,0);
		</script>
		<%end if
	end if%>
<% end if %>
</form>
<!-- End Form -->
</div>
</div>
<!--#include file="adminfooter.asp"-->
<%   
call clearLanguage()

set conntemp=Nothing
set rs=Nothing
set pIdproduct=Nothing
set pDescription=Nothing
set pDetails=Nothing
set pListPrice=Nothing
set pImageUrl=Nothing
set pWeight=Nothing
%>