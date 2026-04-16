<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "viewPrd.asp"
' This page is handles and displays all product-level info
' All product info is retreived from the database and
' displayed in its corresponding display zone.
'
'/////////////////////////////////////////////////////////////////
' NOTES:														//
'																//
' The "viewPrdCode.asp" include will hold the routines that 
' display the product information. Each segment of product 
' information has been divided into zone.
'
' PRODUCT INFORMATION DISPLAY ZONES
'
'1)		CATGEORY BREADCRUMBS
'2)		GENERAL INFORMATION
'3)		QUANTITY AND ADD TO CART
'4)		PRODUCT IMAGES
'5)		DESCRIPTION
'6)		DEFAULT CONFIGURATION (BTO)
'7)		PRODUCT OPTIONS
'8)		CUSTOM SEARCH FIELDS
'9)		CUSTOM INPUT FIELDS
'10)	ACCESSORIES // coming soon!
'11)	QUANTITY AND ADD TO CART (2)
'12)	LONG DESCRIPTION
'13)	CROSS SELLING ZONE
'14)	PRODUCT REVIEWS ZONE
'15)	QUANTITY DISCOUNTS ZONE
'
' View the commented sections of this page to
' find a particular zone.
'
' ZONE RULES
'
' 1) "add-to-cart" must be placed below "options" and "custom fields"
' 2) "wishlist" must be places below "options"
'
'
'/////////////////////////////////////////////////////////////////
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/CashbackConstants.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<%
Response.Buffer = True
'-------------------------------
' declare local variables
'-------------------------------

Dim tIndex, tUpdPrd, pIdCategory, strBreadCrumb, pIdProduct, dblpcCC_Price
Dim pcv_strViewPrdStyle, pcv_strFormAction, pcv_intValidationFile, pcv_blnBTOisConfig, iRewardPoints, pDescription, pMainProductName
Dim pSku, pconfigOnly, pserviceSpec, pPrice, pBtoBPrice, pDetails, pListPrice, plistHidden, pimageUrl, pLgimageURL
Dim pArequired, pBrequired, pStock, pWeight, pEmailText, pFormQuantity, pnoshipping, pcustom1, pcontent1
Dim pcustom2, pcontent2, pcustom3, pcontent3, pnoprices
Dim psDesc, pNoStock, pnoshippingtext, intIdProduct, intWeight, optionA, optionB
Dim pcv_intHideBTOPrice, pcv_intQtyValidate, pcv_lngMinimumQty, intpHideDefConfig, pnotax
Dim FirstCnt, strDescription, intReward, pcv_BTORP, strConfigProductCategory, dblPrice, dblWPrice
Dim VardiscGo, dblQuantityFrom, dblQuantityUntil, dblPercentage, dblDiscountPerWUnit, dblDiscountPerUnit
Dim intIdOptOptGrp, intIdOption, strOptionDescrip, OptInActive, optPrice, tempIdOptA, tempIdOptB
Dim xrequired, xfieldCnt, reqstring, TextArea, widthoffield, rowlength
Dim scCS, cs_showprod, cs_showcart, cs_showimage, crossSellText
Dim pcv_strOptionGroupDesc, pcv_intOptionGroupCount, pcv_strOptionGroupCount, pcv_strOptionGroupID, pcv_strOptionRequired
Dim xOptionsCnt, pcv_strNumberValidations, pcv_strFuntionCall, pcv_strReqOptString, xOtionrequired, pcv_strCSDiscounts , pcv_strPrdDiscounts 
Dim pcv_strProdImage_Url, pcv_strProdImage_LargeUrl, pcv_intProdImage_Columns, pcv_strShowImage_LargeUrl, pcv_strShowImage_Url, pcv_strCurrentUrl, pcv_strShowImage_AltTagText
Dim pcv_strAdditionalImages, cCounter, pcv_strWishListLink, BTOCharges, pcv_strCSString, pcv_strReqCSString, cs_RequiredIds, xCSCnt
Dim iAddDefaultWPrice, iAddDefaultPrice, pcv_intActive
Dim pHideSKU, pcv_IntMojoZoom, pcv_HideAdditionalImages
'APP-S
Dim pcv_ReorderLevel,pcv_Apparel,pcv_ShowStockMsg,pcv_StockMsg,pcv_SizeLink,pcv_SizeInfo,pcv_SizeImg,pcv_SizeURL,pcv_ApparelRadio,pcv_HaveSPs
Dim pcv_TotalOpts
Dim popUpAPP

popUpAPP=0

Dim HaveDiffPrice

HaveDiffPrice=0
'APP-E
Dim pShowAvgRating, pAvgRating, pNumRatings, pcRS_Active, pRSActive, pcv_RatingType, pcv_Img1
Dim ppTop, ppTopLeft, ppTopRight, ppMiddle, ppTabs, ppBottom
Dim pIDBrand, pcIntBrandID, pcIntIDBrand, BrandName, pcStrBrandLink, pcStrBrandLink2
%>
<!--#include file="prv_getSettings.asp"-->
<!--#include file="viewPrdCode.asp"-->
<!--#include file="inc_addThis.asp"-->
<%
'// When the product has additional images, this variable defines how many thumbnails are shown per row, below the main product image
pcv_intProdImage_Columns = 3

'// When this variable is set to 1, ProductCart will up by 1 the "views" count when a product is viewed by store visitors. 
'// This feature can negatively affect performance and a good Web statistics system will provide better information. So the feature is OFF by default.
pcv_IncreaseVisitsOn = 0  'Change to 1 if you wish to utilize this feature (not recommended)

'*****************************************************************************************************
' START PAGE ON-LOAD
'*****************************************************************************************************

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID, check referral
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID, check referral
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Check to see if the user is updating the product after adding it to the shopping cart
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
tIndex=0
tUpdPrd=getUserInput(request.QueryString("imode"),50)
if tUpdPrd="updOrd" then
	tIndex=getUserInput(request.QueryString("index"),10)
	if not validNum(tIndex) then
		tIndex=0
	end if
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Check to see if the user is updating the product after adding it to the shopping cart
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Product
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Get Page Style
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'AVAILABLE LAYOUTS
' pcv_strViewPrdStyle = c // classic product cart layout (image right)
' pcv_strViewPrdStyle = l // two column layout (image left)
' pcv_strViewPrdStyle = o // one column layout

pcv_strViewPrdStyle = ""

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' STEP 1:  Check querystring
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if pcv_strViewPrdStyle = "" then
    pcv_strViewPrdStyle = LCase(getUserInput(Request.QueryString("ViewPrdStyle"),10))
	'// Check querystring saved to session by 404.asp
	if pcv_strViewPrdStyle = "" then
		strSeoQueryString=lcase(session("strSeoQueryString"))
		if strSeoQueryString<>"" then
			if InStr(strSeoQueryString,"viewprdstyle")>0 then
				pcv_strViewPrdStyle=left(replace(strSeoQueryString,"viewprdstyle=",""),1)
			end if
		end if
	end if
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' STEP 2:  Check form
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if pcv_strViewPrdStyle = "" then
	pcv_strViewPrdStyle = LCase(getUserInput(Request.Form("ViewPrdStyle"),10))
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' STEP 3:  Check Product Table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if pcv_strViewPrdStyle = "" and validNum(pIdProduct) then
    query="SELECT pcprod_DisplayLayout FROM products WHERE idProduct=" &pIdProduct
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
    if err.number<>0 then
	    call LogErrorToDatabase()
	    set rs=nothing
	    call closedb()
	    response.redirect "techErr.asp?err="&pcStrCustRefID
    end if
    
    if NOT rs.eof then
        pcv_strViewPrdStyle=LCase(rs("pcprod_DisplayLayout"))
        if isNull(pcv_strViewPrdStyle) OR pcv_strViewPrdStyle="" then
	        pcv_strViewPrdStyle=""
        end if
    end if
	set rs = nothing
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' STEP 4:  Check Categories Table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
	if pcv_strViewPrdStyle = "" AND pIdCategory>0 then
			query="SELECT pcCats_DisplayLayout FROM categories WHERE idCategory=" &pIdCategory
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rs=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			if NOT rs.eof then
					pcv_strViewPrdStyle=LCase(rs("pcCats_DisplayLayout"))
					if isNull(pcv_strViewPrdStyle) OR pcv_strViewPrdStyle="" then
						pcv_strViewPrdStyle=""
					end if
			end if
			set rs = nothing
	end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' STEP 5:  Check Global Settings
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
prodLayout = scViewPrdStyle
if pcv_strViewPrdStyle = "" then
    pcv_strViewPrdStyle = LCase(prodLayout)
end if


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' STEP 6:  Set default layout - no valid layout found
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if pcv_strViewPrdStyle <> "custom" and pcv_strViewPrdStyle <> "stand" and pcv_strViewPrdStyle <> "computer" and pcv_strViewPrdStyle <> "monitor" and pcv_strViewPrdStyle <> "traderpc" and pcv_strViewPrdStyle <> "traderpropc" and pcv_strViewPrdStyle <> "charterpc" then
	pcv_strViewPrdStyle = "custom" 
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Get Page Style
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'--> Check if this customer is logged in with a customer category
dblpcCC_Price=0
if session("customerCategory")<>0 then
	query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rs=nothing
		'//close any connections
		call closedb()
		'//redirect to error page
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if NOT rs.eof then
		strcustomerCategory="YES"
		dblpcCC_Price=rs("pcCC_Price")
		dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
	else
		strcustomerCategory="NO"
	end if
	set rs=nothing
end if

'--> check for discount per quantity
query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rs=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if not rs.eof then
	pDiscountPerQuantity=-1
else
	pDiscountPerQuantity=0
end if
set rs=nothing

'SHW-S
call GetSHWSettings()

if shwOnOff=1 then
	query="SELECT sku FROM Products WHERE IDProduct=" & pidProduct & ";"
	set rs=connTemp.execute(query)
	shwQty=-1
	SHWSync=0
	if not rs.eof then
		tmpSKU=rs("sku")
		set rs=nothing
		shwQty=SHWGetInventoryStatus(tmpSKU)
		if clng(shwQty)>=0 then
			query="UPDATE Products SET stock=" & shwQty & " WHERE IDProduct=" & pidProduct & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			SHWSync=1
			call pcs_hookStockChanged(pidProduct, "")
		end if
	end if
	set rs=nothing
end if
'SHW-E

'--> gets product details from db

query="SELECT active,iRewardPoints, description, sku, configOnly, serviceSpec, price, btobprice, listprice, listHidden, imageurl, largeImageURL, Arequired, Brequired, stock, weight, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3, noprices, IDBrand, noStock, noshippingtext,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcProd_multiQty,pcprod_HideDefConfig, notax, pcProd_BackOrder,pcProd_ShipNDays,pcProd_SkipDetailsPage,pcProd_ReorderLevel,pcprod_Apparel,pcprod_ShowStockMsg,pcprod_StockMsg,pcprod_SizeLink,pcprod_SizeInfo,pcprod_SizeImg,pcprod_SizeURL,pcProd_ApparelRadio,pcProd_HideSKU, pcPrd_MojoZoom, details, sDesc, pcProd_AvgRating, pcProd_Top,pcProd_TopLeft,pcProd_TopRight,pcProd_Middle,pcProd_Bottom,pcProd_Tabs,pcProd_AdditionalImages,pcProd_AltTagText,detailstop FROM products WHERE idProduct=" & pidProduct & " AND configOnly=0 AND removed=0 "
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

'// Check to see if a product with that ID exists in the database
if rs.eof then
	set rs=nothing
	call closeDb()
  	response.redirect "msg.asp?message=88"
end if

'// Load product status (active or inactive)
pcv_intActive=rs("active")
	
'// If inactive and not previewed, redirect to "product inactive" message
if pcv_intActive<>-1 AND session("pcv_intAdminPreview")<>1 then
	set rs=nothing
	call closeDb()
  	response.redirect "msg.asp?message=95"
end if

'// increase visits for product
if pcv_IncreaseVisitsOn=1 then
	query="UPDATE products SET visits=visits+1 WHERE idProduct="& pIdProduct
	set rsVisits=server.CreateObject("ADODB.RecordSet")
	set rsVisits=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsVisits=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rsVisits=nothing
End if

'//save product to Viewed Products List
pcv_intIdProduct = pcf_GetParentId(pIdProduct)
ViewedPrdList=getUserInput2(Request.Cookies("pcfront_visitedPrds"),0)
if Instr(ViewedPrdList,"*" & pcv_intIdProduct & "*")>0 then
	ViewedPrdList = Replace(ViewedPrdList, "*" & pcv_intIdProduct & "*", "*")
end if
if ViewedPrdList="" then
	ViewedPrdList="*"
end if
ViewedPrdList="*" & pcv_intIdProduct & ViewedPrdList
'APP-E

Response.Cookies("pcfront_visitedPrds")=ViewedPrdList
Response.Cookies("pcfront_visitedPrds").Expires=Date() + 365

'// Assign variable values
iRewardPoints=rs("iRewardPoints")
pDescription=replace(rs("description"),"&quot;",chr(34))
pMainProductName=pDescription
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
pLgimageURL=rs("largeImageURL")
pArequired=rs("Arequired")
pBrequired=rs("Brequired")
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
if isNull(pnoprices) OR pnoprices="" then
	pnoprices=0
end if
pIDBrand=rs("IDBrand")
pNoStock=rs("noStock")
pnoshippingtext=rs("noshippingtext")
pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
if isNull(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
	pcv_intHideBTOPrice="0"
end if
pcv_intQtyValidate=rs("pcprod_QtyValidate")
if isNull( pcv_intQtyValidate) OR  pcv_intQtyValidate="" then
	pcv_intQtyValidate="0"
end if				
pcv_lngMinimumQty=rs("pcprod_MinimumQty")
if isNull(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
	pcv_lngMinimumQty="0"
end if
pcv_lngMultiQty=rs("pcProd_multiQty")
if isNull(pcv_lngMultiQty) OR pcv_lngMultiQty="" then
	pcv_lngMultiQty="0"
end if
intpHideDefConfig=rs("pcprod_HideDefConfig")
if isNull(intpHideDefConfig) OR intpHideDefConfig="" then
	intpHideDefConfig="0"
end if
pnotax=rs("notax")

'Start SDBA
pcv_intBackOrder = rs("pcProd_BackOrder")
if isNull(pcv_intBackOrder) OR pcv_intBackOrder="" then
	pcv_intBackOrder = 0
end if
pcv_intShipNDays = rs("pcProd_ShipNDays")
if isNull(pcv_intShipNDays) OR pcv_intShipNDays="" then
	pcv_intShipNDays = 0
end if
'End SDBA

pcv_intSkipDetailsPage=rs("pcProd_SkipDetailsPage")
if isNull(pcv_intSkipDetailsPage) or pcv_intSkipDetailsPage="" then
	pcv_intSkipDetailsPage=0
end if

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

pHideSKU=rs("pcProd_HideSKU")
if IsNull(pHideSKU) or pHideSKU="" then
	pHideSKU=0
end if

pcv_IntMojoZoom=rs("pcPrd_MojoZoom")
if not validNum(pcv_IntMojoZoom) then
	pcv_IntMojoZoom=0
end if

pcv_HideAdditionalImages=rs("pcProd_AdditionalImages")
if not validNum(pcv_HideAdditionalImages) then
	pcv_HideAdditionalImages=0
end if

pAltTagText=rs("pcProd_AltTagText")
If pAltTagText="" Then
	pAltTagText=replace(pDescription,"""","&quot;")
End If

pDetails=replace(rs("details"),"&quot;",chr(34))
psDesc=rs("sDesc")
pDetailsTop=rs("detailstop")

' PRV41 start
pAvgRating = rs("pcProd_AvgRating")
' PRV41 end
ppTop=rs("pcProd_Top")
ppTopLeft=rs("pcProd_TopLeft")
ppTopRight=rs("pcProd_TopRight")
ppMiddle=rs("pcProd_Middle")
ppTabs=rs("pcProd_Tabs")
ppBottom=rs("pcProd_Bottom")
if ppTopLeft="" OR IsNull(ppTopLeft) then
	ppTopLeft=ppTopRight
	ppTopRight=""
end if

set rs=Nothing

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

'// Apparel Product Disregard Stock option is always checked
if pcv_Apparel="1" then
	pNoStock=1
end if
%>
<!--#include file="inc_CheckReqItemStock.asp"-->
<%

'PRV41 start
query = "SELECT COUNT(*) FROM pcReviews WHERE pcRev_IDProduct=" & pidProduct & " AND pcRev_Active<>0"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
If NOT rs.EOF Then
	pNumRatings = CLng(rs(0))
Else
	pNumRatings = 0
End If
Set rs = nothing
' PRV41 end

'--> Check to see if the product has been assigned to a brand. If so, get the brand name
if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
 	query="select BrandName from Brands where IDBrand=" & pIDBrand
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if not rs.eof then
		BrandName=rs("BrandName")
	end if
	set rs=nothing
end if
'SB S
Dim objSB 
Set objSB = New pcARBClass
pSubscriptionID = objSB.getSubscriptionID(pidProduct)
if isNull(pSubscriptionID) OR pSubscriptionID="" then
	pSubscriptionID = "0"
end if
%>
<!--#include file="../includes/pcSBDataInc.asp" --> 
<% 
'SB E
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Product
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'-----------------------------------
' START: Skip Product Details Page
'-----------------------------------
If pcv_intSkipDetailsPage="1" then
	If pserviceSpec<>0 Then
		'// SEO URLs START
		if scSeoURLs=1 then
			session("idProductTransfer")=pidProduct
			Server.Transfer("configureprd.asp") 
		else
			response.redirect "configurePrd.asp?idProduct="&pidProduct
		end if
		'// SEO URLs END
	End if
End if
'-----------------------------------
' END: Skip Product Details Page
'-----------------------------------

'*****************************************************************************************************
' END PAGE ON-LOAD
'*****************************************************************************************************
%>

<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="../includes/javascripts/pcValidateViewPrd.asp"-->

<!-- Link to MojoZoom image magnifier -->
<%
	If (pcv_strViewPrdStyle = "o") OR (pcv_strViewPrdStyle = "c") OR ((pcv_strViewPrdStyle = "t" OR pcv_strViewPrdStyle = "d") AND InStr(ppTopRight, "PrdImg") > 0) Then
		pcv_strMojoZoomOrientation = "left"
	Else
		pcv_strMojoZoomOrientation = "right"
	End If
%>

<% If statusAPP="1" Then %>
	<script type="text/javascript">
	
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
	    
		function open_win(url)
		{
		    var newWin = window.open(url,'','scrollbars=yes,resizable=yes');
		    newWin.focus();
		}
	</script>
	<%
	IF (pcv_Apparel="1") then
		call GenApparelSubProducts()
	END IF
End If
%>



<!-- Start Form -->
<% 
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' GENERATE FORM																						//
' > BTO / BTO Configured / Standard Product															//
' > Each uses a different form action and JavaScript validation function                            //
'/////////////////////////////////////////////////////////////////////////////////////////////////////

'********************************************************************
' VALIDATION FILE
' pcv_intValidationFile = 1 // BTO
' pcv_intValidationFile = 2 // Standard
'
' FORM ACTION
' pcv_strFormAction = "instConfiguredPrd.asp" // BTO configured
' pcv_strFormAction = "instPrd.asp" // BTO NON configured and Standard
'********************************************************************
pcv_blnBTOisConfig = pcf_BTOisConfig '// returns true or false for Configured BTO

If pserviceSpec<>0 Then '// If its BTO Then
	if pcv_blnBTOisConfig then '// if its configured then
		pcv_strFormAction = "instConfiguredPrd.asp"
		pcv_intValidationFile = 1
	else '// Its not configured
		pcv_strFormAction = "instPrd.asp"
		pcv_intValidationFile = 1
	end if
else '// Its standard
	pcv_strFormAction = "instPrd.asp"
	pcv_intValidationFile = 2
end if
%>
<div id="pcMain" class="pcViewPrd">
	<div class="pcMainContent" itemscope itemtype="http://schema.org/Product">
		<!-- Start Form -->
    <form autocomplete="off" method="post" action="/shop/pc/<%=pcv_strFormAction%>" name="additem" class="pcForms" onSubmit="return checkproqty(document.additem.quantity);">
		<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->
		<%
        If NOT pcv_blnBTOisConfig Then
            if tIndex<>0 then '// Check to see if the user is updating a product after adding it to the shopping cart
            %>
            <input name="index" type="hidden" value="<%=tIndex%>">
            <input name="imode" type="hidden" value="<%=tUpdPrd%>">
            <% 
            end if
        End If
        
        set rs=nothing
        
        '************************
        ' GET BTO Config Infor
        '************************
        pcs_GetBTOConfiguration
		
		'// <!-- DA - EDIT -->
		Select Case pcv_strViewPrdStyle
			Case "stand"
				%>  <!--#include file="viewPrd-Stands.asp" -->  <%
			Case "monitor"
				%>  <!--#include file="viewPrd-Monitors.asp" -->  <%
			Case "computer"
				'Check for bundle
				if request.querystring("sid") = "" Then
					%>  <!--#include file="viewPrd-Computers.asp" -->  <%
				else
					%>  <!--#include file="viewPrd-ComputerBundles.asp" -->  <%
				end if
			Case "traderpc"
				%>  <!--#include file="viewPrd-TraderPC.asp" -->  <%
			Case "traderpropc"
				%>  <!--#include file="viewPrd-TraderProPC.asp" -->  <%
			Case "charterpc"
				%>  <!--#include file="viewPrd-CharterPC.asp" -->  <%
		End Select

        '// Customized Layout
        if pcv_strViewPrdStyle = "t" OR pcv_strViewPrdStyle = "d" then
			if pcv_strViewPrdStyle = "d" then
				query="SELECT pcDPL_Top,pcDPL_TopLeft,pcDPL_TopRight,pcDPL_Middle,pcDPL_Bottom,pcDPL_Tabs FROM pcDefaultPrdLayout;"
				set rs=connTemp.execute(query)
				if not rs.eof then
					ppTop=rs("pcDPL_Top")
					ppTopLeft=rs("pcDPL_TopLeft")
					ppTopRight=rs("pcDPL_TopRight")
					ppMiddle=rs("pcDPL_Middle")
					ppBottom=rs("pcDPL_Bottom")
					ppTabs=rs("pcDPL_Tabs")
				end if
				set rs=nothing
			end if
        %>  <!--#include file="viewPrdT.asp" -->  <%
        end if
        %>
        <!--#include file="../includes/javascripts/pcValidateFormViewPrd.asp"-->

<% if pcv_Apparel="1" then %>

	<%if pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0 then%>
	<input type="hidden" name="idproduct" value="<%=pidProduct%>">
	<%end if%>
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
			<% if pcv_ApparelRadio="1" then %>
			    new_GenRadioList(1,"",0);
			<%else%>
			    new_GenDropDown(1,"",0);
			<%end if%>            
			new_CheckOptGroup(1,0);
			firstRun=1;
			}
		}
		
		addEvent(window,'load',myInitFunction);
	</script>
	<%if request.querystring("SubPrd")<>"" then
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

					</script>
					<%
				end if
			end if
			set rs=nothing
		end if
	end if%>

<% end if %>
        </form>
        <!-- End Form -->
          
					<!--#include file="atc_viewprd.asp"-->

        <script type=text/javascript>
            function stopRKey(evt)
            {
                var evt  = (evt) ? evt : ((event) ? event : null);
                var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
                if ((evt.keyCode == 13) && (node.type != "textarea") && (node.getAttribute("name") != "keyword")) { return false; }
            }
            document.onkeypress = stopRKey;
        </script>
    	
  	</div>
    
	<div class="pcClear"></div>
</div>
<!--#include file="orderCompleteTracking.asp"-->
<!--#include file="inc-Cashback.asp"-->
<!--#include file="footer_wrapper.asp"-->