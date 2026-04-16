<%@ LANGUAGE = VBScript.Encode %>
<% Response.AddHeader "X-XSS-Protection","0" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Store Settings"
pageIcon="pcv4_icon_settings.png"
%>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<% pcPageName="AdminSettings.asp"

'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
%>
<!--#include file="pcAdminRetrieveSettings.asp"-->

<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer" align="center">
			<%
			msg=getUserInput(request.querystring("msg"),0)
			if msg<>"" then %>
				<div class="pcCPmessage"><%=msg%></div>
			<% end if %>
		</td>
	</tr>
</table>
<%

Function pcf_catalogItemExists(fileName)
  Set FS = Server.CreateObject("Scripting.FileSystemObject")
  pcv_tmpCatalogPath = Server.MapPath("../pc/catalog/") & "/"
  pcf_catalogItemExists = FS.FileExists(pcv_tmpCatalogPath & fileName)
  Set FS = Nothing
End Function

pcv_isCompanyNameRequired=false
pcv_isCompanyAddressRequired=false
pcv_isCompanyPhoneNumberRequired=false
pcv_isCompanyFaxNumberRequired=false
pcv_isCompanyZipRequired=false
pcv_isCompanyCityRequired=false
pcv_isCompanyStateRequired=false
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isCompanyStateRequired=pcv_strStateCodeRequired
end if
pcv_isCompanyProvinceRequired=false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isCompanyProvinceRequired=pcv_strProvinceCodeRequired
end if
pcv_isCompanyCountryRequired=false
pcv_isCompanyLogoRequired=false
pcv_isMetaTitleRequired=true
pcv_isMetaDescriptionRequired=false
pcv_isMetaKeywordsRequired=false
pcv_isQtyLimitRequired=true
pcv_isAddLimitRequired=true
pcv_isPreRequired=true
pcv_isCustPreRequired=true
pcv_isCatImagesRequired=true
pcv_isShowStockLmtRequired=true
pcv_isOutOfStockPurchaseRequired=true
pcv_isCurSignRequired=false
pcv_isDecSignRequired=false
pcv_isDateFrmtRequired=false
pcv_isMinPurchaseRequired=true
pcv_isWholesaleMinPurchaseRequired=true
pcv_isURLredirectRequired=false
pcv_isSSLRequired=false
pcv_isSSLUrlRequired=false
pcv_isIntSSLPageRequired=false
pcv_isPrdRowRequired=true
pcv_isPrdRowsPerPageRequired=true
pcv_isCatRowRequired=true
pcv_isCatRowsPerPageRequired=true
pcv_isBTypeRequired=true
pcv_isStoreOffRequired=false
pcv_isStoreMsgRequired=false
pcv_isWLRequired=false
pcv_isorderLevelRequired=false
pcv_isDisplayStockRequired=false
pcv_isHideCategoryRequired=false
pcv_isPCOrdRequired=true
pcv_isHideSortProRequired=false
pcv_isViewPrdStyleRequired=false
pcv_isOrderNameRequired=false
pcv_isAllowCheckoutWRRequired=false
pcv_isHideDiscFieldRequired=false
pcv_isDispDiscCartRequired=false
pcv_isAllowSeparateRequired=false
pcv_isDisableDiscountCodesRequired=false
pcv_isShowSKURequired=false
pcv_isShowSmallImgRequired=false
pcv_isHideRMARequired=false
pcv_isShowHDRequired=false
pcv_isErrorHandlerRequired=false
pcv_isDisableGiftRegistryRequired=false
pcv_isBrandProRequired=false
pcv_isBrandLogoRequired=false
pcv_isSeoURLsRequired=false
pcv_isSeoURLs404Required=false
pcv_isDisplayQuickViewRequired=false
pcv_isnewStoreURLRequired=false
pcv_isThemeFolderRequired=false
pcv_isConURLRequired=false
pcv_isDisplayPNButtonsRequired=false
pcv_isQuickBuyRequired=false
pcv_isATCEnabledRequired=false
pcv_isRestoreCartRequired=false
pcv_isAddThisDisplayRequired=false
pcv_isAddThisCodeRequired=false
pcv_isPinterestDisplayRequired=false
pcv_isPinterestCounterRequired=false
pcv_isGoogleAnalyticsRequired=false
pcv_isGATypeRequired=false
pcv_isCartStackRequired=false
pcv_isCSSiteIdRequired=false
pcv_isEnableBundlingRequired=false
pcv_isOptimizeJavascriptRequired=false
pcv_isKeepSessionRequired=false
pcv_isGoogleTagManagerRequired=false
pcv_isEnableGCTRequired=false

if request("updateSettings")<>"" then
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions
	'/////////////////////////////////////////////////////

	'// set errors to none
	pcv_intErr=0

	'// generic error for page
	pcv_strGenericPageError = dictLanguageCP.Item(Session("language")&"_cpCommon_403")

	'// validate all fields
	pcs_ValidateTextField	"CompanyName", pcv_isCompanyNameRequired, 150
	pcs_ValidateTextField	"CompanyAddress", pcv_isCompanyAddressRequired, 250
	pcs_ValidateTextField	"CompanyZip", pcv_isCompanyZipRequired, 20
	pcs_ValidateTextField	"CompanyPhoneNumber", pcv_isCompanyPhoneNumberRequired, 20
	pcs_ValidateTextField	"CompanyFaxNumber", pcv_isCompanyFaxNumberRequired, 20
	pcs_ValidateTextField	"CompanyCity", pcv_isCompanyCityRequired, 50
	pcs_ValidateTextField	"CompanyState", pcv_isCompanyProvinceRequired, 50
	pcs_ValidateTextField	"CompanyProvince", pcv_isCompanyStateRequired, 50
	pcs_ValidateTextField	"CompanyCountry", pcv_isCompanyCountryRequired, 50
	pcs_ValidateTextField	"CompanyLogo", pcv_isCompanyLogoRequired, 250
	pcs_ValidateTextField	"MetaTitle", pcv_isMetaTitleRequired, 250
	pcs_ValidateTextField	"MetaDescription", pcv_isMetaDescriptionRequired, 250
	pcs_ValidateTextField	"MetaKeywords", pcv_isMetaKeywordsRequired, 250
	pcs_ValidateTextField	"QtyLimit", pcv_isQtyLimitRequired, 6
	pcs_ValidateTextField	"AddLimit", pcv_isAddLimitRequired, 6
	pcs_ValidateTextField	"Pre", pcv_isPreRequired, 15
	pcs_ValidateTextField	"CustPre", pcv_isCustPreRequired, 15
	pcs_ValidateTextField	"CatImages", pcv_isCatImagesRequired, 2
	pcs_ValidateTextField	"ShowStockLmt", pcv_isShowStockLmtRequired, 2
	pcs_ValidateTextField	"OutOfStockPurchase", pcv_isOutOfStockPurchaseRequired, 2
	pcs_ValidateTextField	"CurSign", pcv_isCurSignRequired, 10
	pcs_ValidateTextField	"DecSign", pcv_isDecSignRequired, 4
	pcs_ValidateTextField	"DateFrmt", pcv_isDateFrmtRequired, 10
	pcs_ValidateTextField	"MinPurchase", pcv_isMinPurchaseRequired, 20
	pcs_ValidateTextField	"WholesaleMinPurchase", pcv_isWholesaleMinPurchaseRequired, 20
	pcs_ValidateTextField	"URLredirect", pcv_isURLredirectRequired, 250
	pcs_ValidateTextField	"SSL", pcv_isSSLRequired, 4
	pcs_ValidateTextField	"SSLUrl", pcv_isSSLUrlRequired, 250
	pcs_ValidateTextField	"IntSSLPage", pcv_isIntSSLPageRequired, 4
	pcs_ValidateTextField	"PrdRow", pcv_isPrdRowRequired, 20
	pcs_ValidateTextField	"PrdRowsPerPage", pcv_isPrdRowsPerPageRequired, 20
	pcs_ValidateTextField	"CatRow", pcv_isCatRowRequired, 20
	pcs_ValidateTextField	"CatRowsPerPage", pcv_isCatRowsPerPageRequired, 20
	pcs_ValidateTextField	"BType", pcv_isBTypeRequired, 4
	pcs_ValidateTextField	"StoreOff", pcv_isStoreOffRequired, 4
	pcs_ValidateHTMLField	"StoreMsg", pcv_isStoreMsgRequired, 0
	pcs_ValidateTextField	"WL", pcv_isWLRequired, 2
	pcs_ValidateTextField	"orderLevel", pcv_isorderLevelRequired, 2
	pcs_ValidateTextField	"DisplayStock", pcv_isDisplayStockRequired, 2
	pcs_ValidateTextField	"HideCategory", pcv_isHideCategoryRequired, 2
	pcs_ValidateTextField	"PCOrd", pcv_isPCOrdRequired, 10
	pcs_ValidateTextField	"HideSortPro", pcv_isHideSortProRequired, 10
	pcs_ValidateTextField	"ViewPrdStyle",  pcv_isViewPrdStyleRequired, 10
	pcs_ValidateTextField	"OrderName", pcv_isOrderNameRequired, 4
	pcs_ValidateTextField	"AllowCheckoutWR", pcv_isAllowCheckoutWRRequired, 4
	pcs_ValidateTextField	"HideDiscField", pcv_isHideDiscFieldRequired, 4
	pcs_ValidateTextField	"DispDiscCart", pcv_isDispDiscCartRequired, 4
	pcs_ValidateTextField	"AllowSeparate", pcv_isAllowSeparateRequired, 4
	pcs_ValidateTextField	"DisableDiscountCodes", pcv_isDisableDiscountCodesRequired, 4
	pcs_ValidateTextField	"ShowSKU", pcv_isShowSKURequired, 4
	pcs_ValidateTextField	"ShowSmallImg", pcv_isShowSmallImgRequired, 4
	pcs_ValidateTextField	"HideRMA", pcv_isHideRMARequired, 4
	pcs_ValidateTextField	"ShowHD", pcv_isShowHDRequired, 4
	pcs_ValidateTextField	"ErrorHandler", pcv_isErrorHandlerRequired, 4
	pcs_ValidateTextField	"DisableGiftRegistry", pcv_isDisableGiftRegistryRequired, 4
	pcs_ValidateTextField	"BrandPro", pcv_isBrandProRequired, 2
	pcs_ValidateTextField	"BrandLogo", pcv_isBrandLogoRequired, 2
	pcs_ValidateTextField	"SeoURLs", pcv_isSeoURLsRequired, 2
	pcs_ValidateTextField	"SeoURLs404", pcv_isSeoURLs404Required, 50
	pcs_ValidateTextField	"DisplayQuickView", pcv_isDisplayQuickViewRequired, 4
	pcs_ValidateTextField	"DisplayPNButtons", pcv_isDisplayPNButtonsRequired, 4
	pcs_ValidateTextField	"newStoreURL", pcv_isnewStoreURLRequired, 0
	pcs_ValidateTextField	"ThemeFolder", pcv_isThemeFolderRequired, 0
	pcs_ValidateTextField	"ConURL", pcv_isConURLRequired, 4
	pcs_ValidateTextField	"QuickBuy", pcv_isQuickBuyRequired, 4
	pcs_ValidateTextField	"ATCEnabled", pcv_isATCEnabledRequired, 4
	pcs_ValidateTextField	"RestoreCart", pcv_isRestoreCartRequired, 4
	pcs_ValidateTextField	"AddThisDisplay", pcv_isAddThisDisplayRequired, 2
	pcs_ValidateHtmlField	"AddThisCode", pcv_isAddThisCodeRequired, 0
	pcs_ValidateTextField	"PinterestDisplay", pcv_isPinterestDisplayRequired, 2
	pcs_ValidateTextField	"PinterestCounter", pcv_isPinterestCounterRequired, 15
	pcs_ValidateTextField	"GoogleAnalytics", pcv_isGoogleAnalyticsRequired, 50
	pcs_ValidateTextField	"GAType", pcv_isGATypeRequired, 1
	pcs_ValidateTextField	"CartStack", pcv_isCartStackRequired, 1
	pcs_ValidateTextField	"CSSiteId", pcv_isCSSiteIdRequired, 20
	pcs_ValidateTextField	"EnableBundling", pcv_isEnableBundlingRequired, 1
	pcs_ValidateTextField	"OptimizeJavascript", pcv_isOptimizeJavascriptRequired, 1
	pcs_ValidateTextField	"KeepSession", pcv_isKeepSessionRequired, 1
	pcs_ValidateTextField	"GoogleTagManager", pcv_isGoogleTagManagerRequired, 50
	pcs_ValidateTextField	"EnableGCT", pcv_isEnableGCTRequired, 1
	
	'// Validate social link fields
	Dim pcSocialLink_Items
	Set pcSocialLink_Items = Server.CreateObject("Scripting.Dictionary")
	
	for i = 0 to pcSocialLinksCnt - 1
		slId = pcSocialLinksArr(0, i)
		slImage = pcSocialLinksArr(3, i)
		slUrl = pcSocialLinksArr(4, i)
		slAlt = pcSocialLinksArr(5, i)
		slOrder = pcSocialLinksArr(6, i)
		
		pcs_ValidateTextField "SocialLink_Image" & slId, false, 500
		pcs_ValidateTextField "SocialLink_Url" & slId, false, 500
		pcs_ValidateTextField "SocialLink_Alt" & slId, false, 200
		pcs_ValidateTextField "SocialLink_Order" & slId, false, 4
		
		pcSocialLink_Items.Add slId, Array(Session("pcAdminSocialLink_Image" & slId), _
		 Session("pcAdminSocialLink_Url" & slId), _
		 Session("pcAdminSocialLink_Alt" & slId), _
		 Session("pcAdminSocialLink_Order" & slId))
	next
	
	'// Validate social link fields
	Dim pcAcceptedPayments_Items
	Set pcAcceptedPayments_Items = Server.CreateObject("Scripting.Dictionary")
	
	for i = 0 to pcAcceptedPaymentsCnt - 1
		paymentId = pcAcceptedPayments(0, i)
		
		pcs_ValidateTextField "AcceptedPayment_Image" & paymentId, false, 500
		pcs_ValidateTextField "AcceptedPayment_Alt" & paymentId, false, 200
		pcs_ValidateTextField "AcceptedPayment_Active" & paymentId, false, 4
		pcs_ValidateTextField "AcceptedPayment_Order" & paymentId, false, 4
		
		pcAcceptedPayments_Items.Add paymentId, Array(Session("pcAdminAcceptedPayment_Image" & paymentId), _
		Session("pcAdminAcceptedPayment_Alt" & paymentId), _
		Session("pcAdminAcceptedPayment_Active" & paymentId), _
		Session("pcAdminAcceptedPayment_Order" & paymentId))
	next
	
		if NOT validNum(Session("pcAdminCatRow")) OR Session("pcAdminCatRow")= "" OR Session("pcAdminCatRow")= "0" then
			Session("pcAdminCatRow")= 3
		end if

		if NOT validNum(Session("pcAdminCatRowsPerPage")) OR Session("pcAdminCatRowsPerPage")= "" OR Session("pcAdminCatRowsPerPage")= "0" then
			Session("pcAdminCatRowsPerPage")= 3
		end if
		if NOT validNum(Session("pcAdminPrdRow")) OR Session("pcAdminPrdRow")= "" OR Session("pcAdminPrdRow")= "0" then
			Session("pcAdminPrdRow")= 3
		end if

		if NOT validNum(Session("pcAdminPrdRowsPerPage")) OR Session("pcAdminPrdRowsPerPage")= "" OR Session("pcAdminPrdRowsPerPage")= "0" then
			Session("pcAdminPrdRowsPerPage")= 3
		end if

		if NOT validNum(Session("pcAdminBrandPro")) OR Session("pcAdminBrandPro")<>"1" then
			Session("pcAdminBrandPro")="0"
		end if
		if NOT validNum(Session("pcAdminBrandLogo")) OR Session("pcAdminBrandLogo")<>"1" then
			Session("pcAdminBrandLogo")="0"
		end if
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Set specific errors and default values instead of generic error message.
		' response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		if Session("pcAdminQtyLimit")= "" or not validNum(Session("pcAdminQtyLimit")) then
			Session("pcAdminQtyLimit")= 100
		end if

		if Session("pcAdminAddLimit")= "" or not validNum(Session("pcAdminAddLimit")) then
			Session("pcAdminAddLimit")= 10
		end if

		if Session("pcAdminPre")= "" then
			Session("pcAdminPre")= 0
		end if

		if Session("pcAdminCustPre")= "" then
			Session("pcAdminCustPre")= 0
		end if

		if Session("pcAdminCatImages")= "" then
			Session("pcAdminCatImages")= 0
		end if

		if Session("pcAdminShowStockLmt")= "" then
			Session("pcAdminShowStockLmt")= 0
		end if

		if Session("pcAdminOutOfStockPurchase")= "" then
			Session("pcAdminOutOfStockPurchase")= 0
		end if

		if Session("pcAdminMinPurchase")= "" then
			Session("pcAdminMinPurchase")= 0
		end if

		if Session("pcAdminWholesaleMinPurchase")= "" then
			Session("pcAdminWholesaleMinPurchase")= 0
		end if

		if Session("pcAdminPrdRow")= "" or not validNum(Session("pcAdminPrdRow")) then
			Session("pcAdminPrdRow")= 3
		end if

		if Session("pcAdminPrdRowsPerPage")= "" or not validNum(Session("pcAdminPrdRowsPerPage")) then
			Session("pcAdminPrdRowsPerPage")= 3
		end if

		if NOT validNum(Session("pcAdminCatRow")) OR Session("pcAdminCatRow")= "" then
			Session("pcAdminCatRow")= 3
		end if

		if NOT validNum(Session("pcAdminCatRowsPerPage")) OR Session("pcAdminCatRowsPerPage")= "" then
			Session("pcAdminCatRowsPerPage")= 3
		end if

		if Session("pcAdminBType")= "" then
			Session("pcAdminBType")= "H"
		end if

		if Session("pcAdminPCOrd")= "" then
			Session("pcAdminPCOrd")= 0
		end if

	End If

	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	pcStrCompanyName = Session("pcAdminCompanyName")
	pcStrCompanyName=Replace(pcStrCompanyName,"&quot;","""""")
	pcStrCompanyName=Replace(pcStrCompanyName,"""","""""")
	pcStrCompanyAddress = Session("pcAdminCompanyAddress")
	pcStrCompanyZip = Session("pcAdminCompanyZip")
	pcStrCompanyPhoneNumber = Session("pcAdminCompanyPhoneNumber")
	pcStrCompanyFaxNumber = Session("pcAdminCompanyFaxNumber")
	pcStrCompanyCity = Session("pcAdminCompanyCity")
	if Session("pcAdminCompanyProvince")<>"" then
		pcStrCompanyState = Session("pcAdminCompanyProvince")
	else
		pcStrCompanyState = Session("pcAdminCompanyState")
	end if
	pcStrCompanyCountry = Session("pcAdminCompanyCountry")
	pcStrCompanyLogo = Session("pcAdminCompanyLogo")
	pcStrMetaTitle = Session("pcAdminMetaTitle")
	pcStrMetaDescription = Session("pcAdminMetaDescription")
	pcStrMetaKeywords = Session("pcAdminMetaKeywords")
	pcIntQtyLimit = Session("pcAdminQtyLimit")
		if pcIntQtyLimit=0 or not validNum(pcIntQtyLimit) then
			pcIntQtyLimit=50
		end if
	pcIntAddLimit = Session("pcAdminAddLimit")
		if pcIntAddLimit=0 or not validNum(pcIntAddLimit) then
			pcIntAddLimit=1000
		end if
	pcIntPre = Session("pcAdminPre")
		if not validNum(pcIntPre) then
			pcIntPre=0
		end if
	pcIntCustPre = Session("pcAdminCustPre")
		if not validNum(pcIntCustPre) then
			pcIntCustPre=0
		end if
	pcIntCatImages = Session("pcAdminCatImages")
	pcIntShowStockLmt = Session("pcAdminShowStockLmt")
	pcIntOutOfStockPurchase = Session("pcAdminOutOfStockPurchase")
	pcStrCurSign = Session("pcAdminCurSign")
	pcStrDecSign = Session("pcAdminDecSign")
	pcStrDateFrmt = Session("pcAdminDateFrmt")

	'// Alert that integers are required
	if not validNum(Session("pcAdminMinPurchase")) or not validNum(Session("pcAdminWholesaleMinPurchase")) then
		call closeDb()
response.redirect "AdminSettings.asp?tab=2&msg=" & Server.URLEncode("The &quot;Minimum Order Amount&quot; and &quot;Wholesale Minimum Order Amount&quot; fields must be integers.")
	end if
	pcIntMinPurchase = pcf_ReplaceChars(Session("pcAdminMinPurchase"))
	pcIntWholesaleMinPurchase = pcf_ReplaceChars(Session("pcAdminWholesaleMinPurchase"))

	pcStrURLredirect = Session("pcAdminURLredirect")
	pcStrSSLUrl = Session("pcAdminSSLUrl")
	if pcStrSSLUrl = "" OR Left(pcStrSSLUrl, 8)<>"https://" then
		pcStrSSL = 0
	else
		pcStrSSL = Session("pcAdminSSL")
	end if
	pcStrIntSSLPage = Session("pcAdminIntSSLPage")
	pcIntPrdRow = Session("pcAdminPrdRow")
	pcIntPrdRowsPerPage = Session("pcAdminPrdRowsPerPage")
	pcIntCatRow = Session("pcAdminCatRow")
	pcIntCatRowsPerPage = Session("pcAdminCatRowsPerPage")
	pcStrBType = Session("pcAdminBType")
	if len(pcStrBType)<1 then
		pcStrBType="H"
	end if
	pcStrStoreOff = Session("pcAdminStoreOff")
	pcStrStoreMsg = Session("pcAdminStoreMsg")
	if pcStrStoreMsg="" then
		pcStrStoreMsg=scStoreMsg
	end if
	pcStrStoreMsg=Replace(pcStrStoreMsg, vbCrLf, "<BR>")
	pcStrStoreMsg=Replace(pcStrStoreMsg,"""","""""")
	pcIntWL = Session("pcAdminWL")
	pcIntTF = Session("pcAdminTF")
	pcStrorderLevel = Session("pcAdminorderLevel")
	pcIntDisplayStock = Session("pcAdminDisplayStock")
	pcIntHideCategory = Session("pcAdminHideCategory")
	pcIntPCOrd = Session("pcAdminPCOrd")
	pcIntHideSortPro = Session("pcAdminHideSortPro")
	if pcIntHideSortPro="" then
		pcIntHidesortPro=0
	end if
	pcStrViewPrdStyle = Session("pcAdminViewPrdStyle")
	if len(pcStrViewPrdStyle)<1 then
		pcStrViewPrdStyle="L"
	end if
	pcStrOrderName = Session("pcAdminOrderName")
	pcIntAllowCheckoutWR = Session("pcAdminAllowCheckoutWR")
	if pcIntAllowCheckoutWR="" then
		pcIntAllowCheckoutWR=0
	end if
	pcStrHideDiscField = Session("pcAdminHideDiscField")
	pcStrDispDiscCart = Session("pcAdminDispDiscCart")
	pcStrAllowSeparate = Session("pcAdminAllowSeparate")
	if pcStrAllowSeparate="" then
		pcStrAllowSeparate=0
	end if
	pcIntDisableDiscountCodes = Session("pcAdminDisableDiscountCodes")
	if pcIntDisableDiscountCodes="" then
		pcIntDisableDiscountCodes=0
	end if
	pcIntShowSKU = Session("pcAdminShowSKU")
	pcIntShowSmallImg = Session("pcAdminShowSmallImg")
	pcIntHideRMA = Session("pcAdminHideRMA")
	pcIntShowHD = Session("pcAdminShowHD")
	pcIntErrorHandler = Session("pcAdminErrorHandler")
	pcIntDisableGiftRegistry = Session("pcAdminDisableGiftRegistry")
	pcIntBrandPro = Session("pcAdminBrandPro")
	pcIntBrandLogo = Session("pcAdminBrandLogo")
	pcIntSeoURLs = Session("pcAdminSeoURLs")
	if pcIntSeoURLs="" then
		pcIntSeoURLs=0
	end if
	pcStrSeoURLs404 = Session("pcAdminSeoURLs404")
	
	pcIntDisplayQuickView = Session("pcAdminDisplayQuickView")
	pcIntDisplayPNButtons = Session("pcAdminDisplayPNButtons")
	pcStrnewStoreURL = Session("pcAdminnewStoreURL")
	if pcStrnewStoreURL="" then
		pcStrnewStoreURL=scStoreURL
		Session("pcAdminnewStoreURL")=pcStrnewStoreURL
	end if
	pcStrThemeFolder = Session("pcAdminThemeFolder")
	if pcStrThemeFolder="" then
		pcStrThemeFolder="theme/basic_blue"
		Session("pcAdminThemeFolder")=pcStrThemeFolder
	end if
	pcIntConURL = Session("pcAdminConURL")
	pcIntQuickBuy = Session("pcAdminQuickBuy")
	pcIntATCEnabled = Session("pcAdminATCEnabled")
	pcIntRestoreCart = Session("pcAdminRestoreCart")
	pcIntAddThisDisplay = Session("pcAdminAddThisDisplay")
	if pcIntAddThisDisplay="" then
		pcIntAddThisDisplay=0
	end if
	pcStrAddThisCode = Session("pcAdminAddThisCode")
	pcIntPinterestDisplay = Session("pcAdminPinterestDisplay")
	if pcIntPinterestDisplay&""="" then
		pcIntPinterestDisplay="0"
	end if
	pcStrPinterestCounter = Session("pcAdminPinterestCounter")

	pcStrGoogleAnalytics = Session("pcAdminGoogleAnalytics")
	pcIntGAType = Session("pcAdminGAType")
	if pcIntGAType="" then
		pcIntGAType="0"
	end if
	
	pcIntEnableGCT = Session("pcAdminEnableGCT")
	pcStrGCTCode = Trim(Request.Form("GCTCode"))
	if pcIntEnableGCT="" OR pcStrGCTCode="" then
		pcIntEnableGCT="0"
	end if

    pcIntEnableBulkAdd = request("enableBulkAdd")


	
	pcIntCartStack = Session("pcAdminCartStack")
	if pcIntCartStack = "" then
		pcIntCartStack = "0"
	end if
	pcStrCSSiteId = Session("pcAdminCSSiteId")
	if pcStrCSSiteId = "" then
		pcIntCartStack = "0"
	end if
	
	pcIntEnableBundling = Session("pcAdminEnableBundling")
	if pcIntEnableBundling = "" then
		pcIntEnableBundling = "0"
	end if
	
	pcIntOptimizeJavascript = Session("pcAdminOptimizeJavascript")
	if pcIntOptimizeJavascript = "" then
		pcIntOptimizeJavascript = "0"
	end if
	
	pcIntKeepSession = Session("pcAdminKeepSession")
	if pcIntKeepSession = "" then
		pcIntKeepSession = "0"
	end if
	
	pcStrGoogleTagManager = Session("pcAdminGoogleTagManager")

	pcs_ClearAllSessions()
	if pcStrDecSign="," then
		pcStrDivSign="."
	else
		pcStrDivSign=","
	end if

	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
	%>
	<!--#include file="pcAdminSaveSettings.asp"-->
	<!--#include file="pcAdminRetrieveSettings.asp"-->
<% end if %>

<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>

<script type="text/javascript">
function Form1_Validator(theForm)
{

<%
StrGenericJSError=dictLanguageCP.Item(Session("language")&"_cpCommon_403")

'// Panel 0
pcs_JavaTextField	"CurSign", pcv_isCurSignRequired, StrGenericJSError, "0"
pcs_JavaDropDownList "DecSign", pcv_isDecSignRequired, StrGenericJSError
pcs_JavaTextField	"StoreMsg", pcv_isStoreMsgRequired, StrGenericJSError, "0"
pcs_JavaDropDownList "PinterestCounter", pcv_isPinterestCounterRequired, StrGenericJSError
pcs_JavaDropDownList "DateFrmt", pcv_isDateFrmtRequired, StrGenericJSError
pcs_JavaTextField	"URLredirect", pcv_isURLredirectRequired, StrGenericJSError, "0"
pcs_JavaCheckedBox "SSL", pcv_isSSLRequired, StrGenericJSError
pcs_JavaTextField	"SSLUrl", pcv_isSSLUrlRequired, StrGenericJSError, "0"
%>

if(theForm.SSLUrl.value!="") {
	if(!checkSSLUrl(theForm.SSLUrl.value)) {
		alert("Please enter a valid SSL URL beginning with \"https://\"");
		$pc('#TabbedPanels2 li:eq(0) a').tab('show');
		theForm.SSLUrl.focus();
		return (false);
	}
}

<%
'// Panel 1
pcs_JavaTextField	"CompanyName", pcv_isCompanyNameRequired, StrGenericJSError, "1"
pcs_JavaTextField	"CompanyAddress", pcv_isCompanyAddressRequired, StrGenericJSError, "1"
pcs_JavaTextField	"CompanyZip", pcv_isCompanyZipRequired, StrGenericJSError, "1"
pcs_JavaTextField	"CompanyPhoneNumber", pcv_isCompanyPhoneNumberRequired, StrGenericJSError, "1"
pcs_JavaTextField	"CompanyFaxNumber", pcv_isCompanyFaxNumberRequired, StrGenericJSError, "1"
pcs_JavaTextField	"CompanyCity", pcv_isCompanyCityRequired, StrGenericJSError, "1"
pcs_JavaTextField	"CompanyState", pcv_isCompanyStateRequired, StrGenericJSError, "1"
pcs_JavaDropDownList "CompanyCountry", pcv_isCompanyCountryRequired, StrGenericJSError
pcs_JavaTextField	"CompanyLogo", pcv_isCompanyLogoRequired, StrGenericJSError, "1"
pcs_JavaTextField	"MetaTitle", pcv_isMetaTitleRequired, StrGenericJSError, "1"
pcs_JavaTextField	"MetaDescription", pcv_isMetaDescriptionRequired, StrGenericJSError, "1"
pcs_JavaTextField	"MetaKeywords", pcv_isMetaKeywordsRequired, StrGenericJSError, "1"

'// Panel 2
pcs_JavaTextField	"QtyLimit", pcv_isQtyLimitRequired, StrGenericJSError, "2"
pcs_JavaTextField	"AddLimit", pcv_isAddLimitRequired, StrGenericJSError, "2"
pcs_JavaTextField	"MinPurchase", pcv_isMinPurchaseRequired, StrGenericJSError, "2"
pcs_JavaTextField	"WholesaleMinPurchase", pcv_isWholesaleMinPurchaseRequired, StrGenericJSError, "2"
pcs_JavaTextField	"Pre", pcv_isPreRequired, StrGenericJSError, "2"
pcs_JavaTextField	"CustPre", pcv_isCustPreRequired, StrGenericJSError, "2"

'// Panel 3
pcs_JavaTextField	"PrdRow", pcv_isPrdRowRequired, StrGenericJSError, "3"
pcs_JavaTextField	"PrdRowsPerPage", pcv_isPrdRowsPerPageRequired, StrGenericJSError, "3"
pcs_JavaTextField	"CatRow", pcv_isCatRowRequired, StrGenericJSError, "3"
pcs_JavaTextField	"CatRowsperPage", pcv_isCatRowsPerPageRequired, StrGenericJSError, "3"
'pcs_JavaCheckedBox	"AllowCheckoutWR", pcv_isAllowCheckoutWR, StrGenericJSError
pcs_JavaCheckedBox "HideSortPro", pcv_isHideSortProRequired, StrGenericJSError
pcs_JavaCheckedBox	"AllowSeparate", pcv_isAllowSeparateRequired, StrGenericJSError
pcs_JavaCheckedBox	"DisableDiscountCodes", pcv_isDisableDiscountCodesRequired, StrGenericJSError

pcs_JavaTextField	"ShowSKU", pcv_isShowSKURequired, StrGenericJSError, "4"
pcs_JavaTextField	"ShowSmallImg", pcv_isShowSmallImgRequired, StrGenericJSError, "4"
pcs_JavaTextField	"DisplayQuickView", pcv_isDisplayQuickViewRequired, StrGenericJSError, "4"
pcs_JavaTextField	"DisplayPNButtons", pcv_isDisplayPNButtonsRequired, StrGenericJSError, "4"
pcs_JavaTextField	"newStoreURL", pcv_isnewStoreURLRequired, StrGenericJSError, "4"
pcs_JavaTextField	"ThemeFolder", pcv_isThemeFolderRequired, StrGenericJSError, "4"
pcs_JavaTextField	"ConURL", pcv_isConURLRequired, StrGenericJSError, "4"
pcs_JavaTextField	"QuickBuy", pcv_isQuickBuyRequired, StrGenericJSError, "4"
pcs_JavaTextField	"ATCEnabled", pcv_isATCEnabledRequired, StrGenericJSError, "4"
pcs_JavaTextField	"RestoreCart", pcv_isRestoreCartRequired, StrGenericJSError, "4"
%>

return (true);
}

function checkSSLUrl(value) {
	var urlregex = new RegExp(/^(https):\/\/[a-z0-9]+([\-\.]{1}[a-z0-9]+)*\.[a-z]{2,5}(:[0-9]{1,5})?(\/.*)?$/i);
    if (!urlregex.test(value)) {
    	return (false);
    }
	return (true);
}
</script>
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' End Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

%>

<form name="form1" method="post" action="<%=pcPageName%>" onSubmit="return Form1_Validator(this);" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td valign="top">
				<div id="TabbedPanels2" class="tabbable">
				  <ul class="nav nav-tabs">
					<li class="active"><a href="#tabs-1" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_1")%></a></li>
					<li><a href="#tabs-2" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_2")%></a></li>
					<li><a href="#tabs-3" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_3")%></a></li>
					<li><a href="#tabs-4" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_4")%></a></li>
					<li><a href="#tabs-5" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_77")%></a></li>
					<li><a href="#tabs-6" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_5")%></a></li>
				  </ul>

                  <div class="tab-content">
						<div id="tabs-1" class="tab-pane active">
						<table class="pcCPcontent">
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_6")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="StoreOff" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_7")%>
									<input type="radio" name="StoreOff" value="1" <% if pcStrStoreOff="1" then%>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_8")%>
									 &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=412"></a>
								</td>
							</tr>
							<tr>
								<td align="right" valign="top" width="20%"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_9")%></td>
								<%
								pcStrStoreMsg=Replace(pcStrStoreMsg, "<BR>", vbCrLf)
								pcStrStoreMsg=Replace(pcStrStoreMsg, "<br>", vbCrLf)
								pcStrStoreMsg=Replace(pcStrStoreMsg, """""", """")
								%>
								<td><textarea name="StoreMsg" cols="60" rows="6"><%=pcStrStoreMsg%></textarea></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_10")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_11")%>:</td>
								<td align="left">
								<input type="text" name="CurSign" value="<%=pcStrCurSign%>" size="20">
								<% pcs_RequiredImageTag "CurSign", pcv_isCurSignRequired %>
								</td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_12")%>: </td>
								<td align="left">
									<select name="DecSign">
									<option value=""></option>
									<option value="." <% if (pcStrDecSign=".") and (pcStrDivSign=",") then %>selected<%end if%>>1,234,567.89</option>
									<option value="," <% if (pcStrDecSign=",") and (pcStrDivSign=".") then %>selected<% end if %>>1.234.567,89</option>
									</select>
									<% pcs_RequiredImageTag "DecSign", pcv_isDecSignRequired %>
							  </td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_13")%>:</td>
								<td align="left">
									<select name="DateFrmt">
									<option value="MM/DD/YY" selected><%=dictLanguageCP.Item(Session("language")&"_cpCommon_235")%></option>
									<option value="DD/MM/YY" <% if pcStrDateFrmt="DD/MM/YY" then %>selected<% end if %>><%=dictLanguageCP.Item(Session("language")&"_cpCommon_234")%></option>
									</select>
									<% pcs_RequiredImageTag "DateFrmt", pcv_isDateFrmtRequired %>
							  </td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_78")%></th>
							</tr>
							<tr>
							<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
							<td class="pcCPspacer" colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_79")%></td>
							</tr>
							<%if pcStrnewStoreURL="" then
								pcStrnewStoreURL=scStoreURL
							end if%>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_78")%>:</td>
								<td>
									<input type="text" name="newStoreURL" size="50" maxlength="250" value="<%=pcStrnewStoreURL%>">
								</td>
							</tr>
							<%
							'///////////////////////////////////////////////////////////////////////
							'// START: ENFORCE URL
							'//
							'// Note: Only if scStoreURL contains "http://" or "https://"
							'//
							'///////////////////////////////////////////////////////////////////////
							strDomain = getDomainFromURL(scStoreURL)
							If len(strDomain)>0 Then
                            %>
							<tr>
								<td valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_82")%></td>
								<td>
									<input type="radio" name="ConURL" value="1" <% If pcIntConURL="1" then %>checked<%end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="ConURL" value="0" <% If pcIntConURL<>"1" then %>checked<%end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=480"></a><br />
									<span class="pcSmallText"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_83")%> http://<%= strDomain %>/</span>
								</td>
							</tr>
                            <% Else %>
                            <input type="hidden" name="ConURL" value="1" />

                            <%
                            End If
                            '///////////////////////////////////////////////////////////////////////
                            '// END: ENFORCE URL
                            '///////////////////////////////////////////////////////////////////////
                            %>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_14")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_15")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% if pcStrSSL="1" then %>                                
									<input type="checkbox" name="ssl" value="1" checked class="clearBorder">
								<% else %>
									<input id="sslCheckbox" type="checkbox" name="ssl" value="1" data-target=".ssl" class="clearBorder">
                                    <div id="sslModal" class="modal fade ssl">
                                      <div class="modal-dialog">
                                        <div class="modal-content">
                                          <div class="modal-header">
                                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                            <h4 class="modal-title"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_14")%></h4>
                                          </div>
                                          <div class="modal-body">
                                            <p>
                                                It is important to test checkout on your storefront after you enable SSL.  If the shopping cart becomes empty when your browser switches to SSL, then you need to contact your web host to disable the setting <strong>New ID On Secure Connection</strong>.  You may view the <a href="http://wiki.productcart.com/developers/timeout-issues?s[]=session&s[]=loss#iis7" target="_blank">ProductCart Wiki</a> for more information about the setting.  If you need further assistance please contact your web host directly.
                                            </p>
                                          </div>
                                          <div class="modal-footer">
                                            <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
								<% end if %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_16")%></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_17")%>:
								<input type="text" id="SSLUrl" name="sslURL" size="50" value="<%=pcStrSSLURL%>">
								<% pcs_RequiredImageTag "sslURL", pcv_issslURLRequired %>
								</td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_18")%></td>
							</tr>
							<tr>
							<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_19")%>:</td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;&nbsp;
								<input name="intSSLPage" type="radio" value="1" <% if pcStrIntSSLPage="1" then %>checked<%end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_20")%>&nbsp;<a href="JavaScript:win('AdminSettingsSSL.asp')"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_21")%> &gt;&gt;</a></td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;&nbsp;
								<input name="intSSLPage" type="radio" value="0" <% if pcStrIntSSLPage="0" OR pcStrIntSSLPage="" then %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_22")%></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_23")%></th>
							</tr>
							<tr>
							<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
							<td class="pcCPspacer" colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_24")%></td>
							</tr>
							<tr>
								<td colspan="3"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_25")%>:
								<input type="text" name="URLredirect" size="50" maxlength="250" value="<%=pcStrURLredirect%>">
								<% pcs_RequiredImageTag "URLredirect", pcv_isURLredirectRequired %>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
						</table>
					</div>

					<div id="tabs-2" class="tab-pane">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_3")%></th>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td nowrap width="20%"><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_2")%>:</p></td>
								<td width="80%">
								<%
								pcStrCompanyName=Replace(pcStrCompanyName, """""", "&quot;")
								%>
								<p>
								<input type="text" name="CompanyName" value="<%=pcStrCompanyName%>" size="40">
								<% pcs_RequiredImageTag "CompanyName", pcv_isCompanyNameRequired %>
								</p>
								</td>
							</tr>
							<%
							
							'///////////////////////////////////////////////////////////
							'// START: COUNTRY AND STATE/ PROVINCE CONFIG
							'///////////////////////////////////////////////////////////
							'
							' 1) Place this section ABOVE the Country field
							' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
							' 3) Additional Required Info

							'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
							pcv_isStateCodeRequired = pcv_isstateRequired '// determines if validation is performed (true or false)
							pcv_isProvinceCodeRequired = pcv_isprovinceRequired '// determines if validation is performed (true or false)
							pcv_isCountryCodeRequired = pcv_iscountryRequired '// determines if validation is performed (true or false)

							'// #3 Additional Required Info
							pcv_strTargetForm = "form1" '// Name of Form
							pcv_strCountryBox = "CompanyCountry" '// Name of Country Dropdown
							pcv_strTargetBox = "CompanyState" '// Name of State Dropdown
							pcv_strProvinceBox =  "CompanyProvince" '// Name of Province Field

							'// Set local Country to Session
							if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrCompanyCountry
							end if

							'// Set local State to Session
							if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrCompanyState
							end if

							'// Set local Province to Session
							if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pcStrCompanyState
							end if
							%>
							<!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
							<%
							'///////////////////////////////////////////////////////////
							'// END: COUNTRY AND STATE/ PROVINCE CONFIG
							'///////////////////////////////////////////////////////////
							%>
							<%
							'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
							pcs_CountryDropdown
							%>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_3")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyAddress" value="<%=pcStrCompanyAddress%>" size="40">
								<% pcs_RequiredImageTag "CompanyAddress", pcv_isCompanyAddressRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_22")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyCity" value="<%=pcStrCompanyCity%>" size="40">
								<% pcs_RequiredImageTag "CompanyCity", pcv_isCompanyCityRequired %>
								</p>
								</td>
							</tr>
							<%
							'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
							pcs_StateProvince
							%>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_25")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyZip" value="<%=pcStrCompanyZip%>" size="40">
								<% pcs_RequiredImageTag "CompanyZip", pcv_isCompanyZipRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_13")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyPhoneNumber" value="<%=pcStrCompanyPhoneNumber%>" size="40">
								<% pcs_RequiredImageTag "CompanyPhoneNumber", pcv_isCompanyPhoneNumberRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_14")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyFaxNumber" value="<%=pcStrCompanyFaxNumber%>" size="40">
								<% pcs_RequiredImageTag "CompanyFaxNumber", pcv_isCompanyFaxNumberRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_29")%>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=472"></a></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_312")%>:</td>
								<td>
								<input type="text" name="CompanyLogo" value="<%=pcStrCompanyLogo%>" size="40">
								<% pcs_RequiredImageTag "CompanyLogo", pcv_isCompanyLogoRequired %>
								<a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=CompanyLogo&fid=form1','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
								<a href="javascript:;" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')"><img src="images/sortasc_blue.gif" alt="Upload Image"></a>
							&nbsp;(e.g.: <i>mylogo.gif</i>)
								</td>
							</tr>
							<tr>
								<td></td>
								<td>
								<% if trim(pcStrCompanyLogo)<>"" then %>
								<hr>
								Currently using:
								<div style="padding: 15px 0;"><img src="../pc/catalog/<%=pcStrCompanyLogo%>"></div>
								<% end if %>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2">Default Meta Tags&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=473"></a></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td valign="top">Title:</td>
								<td>
								<textarea id="MetaTitle" name="MetaTitle" rows="2" cols="60" onKeyUp="javascript:testchars(this,'1',250); javascript:document.getElementById('MetaTitleCounter').style.display='';"><%=pcStrMetaTitle%></textarea>
								<% pcs_RequiredImageTag "MetaTitle", pcv_isMetaTitleRequired %>
								<div id="MetaTitleCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar1" name="countchar1" style="font-weight: bold"><%=maxlength%></span> characters left. Recommended length: around 60 characters.</div>
								</td>
							</tr>
							<tr>
								<td valign="top">Description:</td>
								<td>
								<textarea id="MetaDescription" name="MetaDescription" rows="2" cols="60" onKeyUp="javascript:testchars(this,'2',250); javascript:document.getElementById('MetaDescriptionCounter').style.display='';"><%=pcStrMetaDescription%></textarea>
								<% pcs_RequiredImageTag "MetaTitle", pcv_isMetaDescriptionRequired %>
								<div id="MetaDescriptionCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar2" name="countchar2" style="font-weight: bold"><%=maxlength%></span> characters left. Recommended length: around 150 characters.</div>
								</td>
							</tr>
							<tr>
								<td valign="top">Keywords:</td>
								<td>
								<textarea id="MetaKeywords" name="MetaKeywords" rows="2" cols="60" onKeyUp="javascript:testchars(this,'3',250); javascript:document.getElementById('MetaKeywordsCounter').style.display='';"><%=pcStrMetaKeywords%></textarea>
								<% pcs_RequiredImageTag "MetaKeywords", pcv_isMetaKeywordsRequired %>
								<div id="MetaKeywordsCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar3" name="countchar3" style="font-weight: bold"><%=maxlength%></span> characters left.</div>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
						</table>
					</div>

					<div id="tabs-3" class="tab-pane">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_31")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
							<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_32")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=415"></a></td>
							</tr>
							<tr>
								<td align="right">
								<input name="orderlevel" type="radio" value="0" checked class="clearBorder">
								</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_33")%></td>
							</tr>
							<tr>
								<td align="right">
								<input type="radio" name="orderlevel" value="1" <% if pcStrOrderlevel="1" then%>checked<%end if%> class="clearBorder">
								</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_34")%></td>
							</tr>
							<tr>
								<td align="right">
								<input type="radio" name="orderlevel" value="2" <% if pcStrOrderlevel="2" then%>checked<%end if%> class="clearBorder">
								</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_35")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_36")%>:</td>
								<td align="left">
								<input type="text" name="QtyLimit" value="<%=pcIntQtyLimit%>" size="20">
								<% pcs_RequiredImageTag "QtyLimit", pcv_isQtyLimitRequired %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=416"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2" align="right"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_37")%>:</td>
								<td align="left">
								<input type="text" name="AddLimit" value="<%=pcIntAddLimit%>" size="20">
								<% pcs_RequiredImageTag "AddLimit", pcv_isAddLimitRequired %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=417"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_38")%>:</td>
								<td align="left">
								<input type="text" name="MinPurchase" value="<%=pcIntMinPurchase%>" size="20">
								<% pcs_RequiredImageTag "MinPurchase", pcv_isMinPurchaseRequired %>
								</td>
							</tr>
							<tr>
								<td align="right">&nbsp;</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_39")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_40")%>:</td>
								<td align="left">
								<input type="text" name="WholesaleMinPurchase" value="<%=pcIntWholesaleMinPurchase%>" size="20">
								<% pcs_RequiredImageTag "WholesaleMinPurchase", pcv_isWholesaleMinPurchaseRequired %>
								</td>
							</tr>
							<tr>
								<td align="right">&nbsp;</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_41")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_42")%>:</td>
								<td>
								<input name="Pre" type="text" id="Pre" value="<%=pcIntPre%>" size="10">
								<% pcs_RequiredImageTag "Pre", pcv_isPreRequired %>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_43")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_44")%>:</td>
								<td>
								<input name="CustPre" type="text" id="CustPre" value="<%=pcIntCustPre%>" size="10">
								<% pcs_RequiredImageTag "CustPre", pcv_isCustPreRequired %>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_45")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_47")%>:
									<% If pcStrOrderName="1" then %>
									<input type="radio" name="OrderName" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="OrderName" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									<% else %>
									<input type="radio" name="OrderName" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="OrderName" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									<% end if %>
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=434"></a>
							 </td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_48")%>: <input type="checkbox" name="AllowSeparate" value="1" <%if pcStrAllowSeparate="1" then%>checked<%end if%> class="clearBorder">&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=108"></a></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_76")%>: <% If pcIntDisableDiscountCodes="1" then %>
									<input type="radio" name="DisableDiscountCodes" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisableDiscountCodes" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
									<input type="radio" name="DisableDiscountCodes" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisableDiscountCodes" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=220"></a></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_49")%>:
								<% If pcIntOutofstockpurchase="0" then %>
								<input type="radio" name="outofstockpurchase" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="outofstockpurchase" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
								<input type="radio" name="outofstockpurchase" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="outofstockpurchase" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=411"></a>
								</td>
							</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						</table>
					</div>

					<div id="tabs-4" class="tab-pane">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_80")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<%
							query = "SELECT pcThemes_Name FROM pcThemes WHERE pcThemes_Active = 1;"
							set rs=connTemp.execute(query)
							If Not rs.eof Then
							  pcStrTheme = rs("pcThemes_Name")
							Else
							  pcStrTheme = Mid(scThemePath, 7, Len(scThemePath))
							End If
                            Set rs = Nothing
							%>
							<tr>
								<td colspan="2">
								    <%
									If PPD="1" Then
										ThemePath = Server.MapPath("/" & scPcFolder & "/pc/theme/") & "/" & pcStrTheme
									Else
										ThemePath = Server.MapPath("../pc/theme/") & "/" & pcStrTheme
									End If

									Dim ThemeFS
									Set ThemeFS = Server.CreateObject("Scripting.FileSystemObject")

									If ThemeFS.FolderExists(ThemePath) Then
								    %>
									<div class="row">
                                        <div class="col-xs-4">
                                            <div class="pcThemeIconMain">
                                            <%
                                            ThumbnailImage = ThemePath & "\"& pcStrTheme & ".jpg"										
                                            If ThemeFS.FileExists(ThumbnailImage) Then
                                                %>
                                                <img src="../pc/theme/<%=pcStrTheme%>/<%=pcStrTheme%>.jpg" class="img-thumbnail" />
                                                <% else %>
                                                <img src="../pc/theme/_common/images/icon.png" class="img-thumbnail" />
                                                <%
                                            End If
                                            %>
                                            </div>
                                        </div>
                                        <div class="col-xs-8">
                                            <div class="pcThemeDetail">
											    <h3><%=pcf_displayThemeName(pcStrTheme)%></h3>
                                                <a href="ThemeSettings.asp" class="btn btn-default">Change Theme</a>
                                                <% If session("PmAdmin")="19" Then %>
                                                    &nbsp;&nbsp;
                                                    <a href="ThemeEditor.asp" class="btn btn-default">Edit Theme</a>
                                                <% End If %>
                                                
                                                &nbsp;&nbsp;
                                                <a href="../pc/default.asp" class="btn btn-default" target="_blank">View Storefront</a>
                                            </div>
                                        </div>
  									</div>
								<% Else %>
									<input type="hidden" name="ThemeFolder" value="theme/<%=pcStrTheme%>">
								<% End If %>
                                <%
                                Set ThemeFS = Nothing
                                Set ThemeDir = Nothing
                                %>
								</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_50")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_51")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=427"></a></td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_505")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="0" <% If trim(pcIntCatImages)="0" then  %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_506")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="4" <% If trim(pcIntCatImages)="4" then  %>checked<% end if %> class="clearBorder">Thumbnails only
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="2" <% If trim(pcIntCatImages)="2" then  %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_507")%>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_52")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=428"></a></td>
							</tr>
							<tr>
								<td align="right" nowrap="nowrap"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_508")%>:</td>
								<td align="left">
                                <input type="number" min="1" max="12" name="CatRow" value="<%=pcIntCatRow%>">
								<% pcs_RequiredImageTag "CatRow", pcv_isCatRowRequired %>
								</td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
								<td width="556" align="left">
								<input type="number" min="1" name="CatRowsperPage" value="<%=pcIntCatRowsPerPage%>">
								<% pcs_RequiredImageTag "CatRowsperPage", pcv_isCatRowsperPageRequired %>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_53")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=429"></a></td>
							</tr>
							<tr>
								<td colspan="2">
									<% If ucase(trim(pcStrBType))="H" then  %>
									 <input type="radio" name="BType" value="H" checked class="clearBorder">
									<% Else %>
									 <input type="radio" name="BType" value="H" class="clearBorder">
									<% End If %>
								 <%=dictLanguageCP.Item(Session("language")&"_cpCommon_510")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="P" then  %>
								 <input type="radio" name="BType" value="P" checked class="clearBorder">
								<% Else %>
								 <input type="radio" name="BType" value="P" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_511")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="L" then  %>
									<input type="radio" name="BType" value="L" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="BType" value="L" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_512")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="M" then  %>
									<input type="radio" name="BType" value="M" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="BType" value="M" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_513")%></td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_54")%>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=430"></a></td>
							</tr>
							<tr>
								<td align="right" width="20%" nowrap="nowrap"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_514")%>:</td>
								<td align="left" width="80%" nowrap="nowrap">
                                <input type="number" min="1" max="12" name="PrdRow" value="<%=pcIntPrdRow%>">
								<% pcs_RequiredImageTag "PrdRow", pcv_isPrdRowRequired %></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
								<td align="left">
								<input type="number" min="1" name="PrdRowsPerPage" value="<%=pcIntPrdRowsPerPage%>">
								<% pcs_RequiredImageTag "PrdRowsPerPage", pcv_isPrdRowsPerPageRequired %>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_55")%>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=431"></a></td>
							</tr>
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_515")%>: <input type="radio" name="ShowSKU" value="-1" <%If pcIntShowSKU="-1" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="radio" name="ShowSKU" value="0" <%If pcIntShowSKU="0" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_516")%>: <input type="radio" name="ShowSmallImg" value="-1" <%If pcIntShowSmallImg="-1" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="radio" name="ShowSmallImg" value="0" <%If pcIntShowSmallImg="0" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><a name="brandSettings"></a>Brands Display Settings</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2">
									These settings apply to top level brands. <a href="BrandsManage.asp" target="_blank">Additional display settings</a> are available for second level brands (sub-brands) and products displayed within a brand. 
									<strong>NOTE: Items per row and rows per page will default to the Category Display settings set above.</strong> 
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<input name="BrandPro" type="checkbox" id="BrandPro" value="1" <%if pcIntBrandPro=1 then%>checked<%end if%>>
								<% pcs_RequiredImageTag "BrandPro", pcv_isBrandProRequired %>
								Show brand on product details page
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<input type="checkbox" name="BrandLogo" value="1" <%if pcIntBrandLogo=1 then%>checked<%end if%>>
								<% pcs_RequiredImageTag "BrandLogo", pcv_isBrandLogoRequired %>
								Show brand logo on &quot;<a href="../pc/viewbrands.asp" target="_blank">Browse By Brand</a>&quot; page</td>
							</tr>

							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_56")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_57")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=432"></a></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="") or (pcIntPCOrd="0") then  %>
									<input type="radio" name="PCOrd" value="0" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="0" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_58")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="1") then  %>
									<input type="radio" name="PCOrd" value="1" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="1" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_59")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="2") then  %>
									<input type="radio" name="PCOrd" value="2" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="2" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_60")%>&nbsp;</td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="3") then  %>
									<input type="radio" name="PCOrd" value="3" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="3" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_61")%></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_65")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=433"></a></td>
							</tr>
							<tr>
								<td colspan="2">
								<input type="checkbox" name="HideSortPro" value="1" <% If (pcIntHideSortPro="1") then  %>checked<%end if%> class="clearBorder">
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_64")%></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_62")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_63")%>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=424"></a></td>
							</tr>
							<tr>
								<td colspan="2">
									<% If ucase(trim(pcStrViewPrdStyle))="C" then  %>
									 <input type="radio" name="ViewPrdStyle" value="C" checked class="clearBorder">
									<% Else %>
									 <input type="radio" name="ViewPrdStyle" value="C" class="clearBorder">
									<% End If %>
								 <%=dictLanguageCP.Item(Session("language")&"_cpCommon_502")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrViewPrdStyle))="L" then  %>
								 <input type="radio" name="ViewPrdStyle" value="L" checked class="clearBorder">
								<% Else %>
								 <input type="radio" name="ViewPrdStyle" value="L" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_503")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrViewPrdStyle))="O" then  %>
									<input type="radio" name="ViewPrdStyle" value="O" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="ViewPrdStyle" value="O" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_504")%></td>
							</tr>
                            <tr>
								<td colspan="2">
								<% If ucase(trim(pcStrViewPrdStyle))="D" then  %>
									<input type="radio" name="ViewPrdStyle" value="D" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="ViewPrdStyle" value="D" class="clearBorder">
								<% End If %>
								Use Custom Layout</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
						</table>
					</div>

          <div id="tabs-5" class="tab-pane">            
            <table class="pcCPcontent">
              <tr>
                <th colspan="2">Social Links</th>
              </tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
              <tr>
              	<td colspan="2">
                	<p>
                    Add a link to any of the items below to display them as social icons on your storefront. If no link is included, no icon will be displayed.
                    <b>NOTE: </b> You can drag-and-drop items to change the display order.
                    &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=900"></a>
                  </p>
                  <br/>
                  <p>To upload custom social icons to specify below, <a href="pcv4_image_upload.asp" target="_blank">click here</a>.</p>
                </td>
              </tr>
              <tr>
                <td colspan="2"><hr></td>
              </tr>
              <tr>
              	<td colspan="2">
                	<% if pcSocialLinksCnt > 0 then %>
                    <ul id="pcCPsocialLinks" class="pcCPsortable">
                      <%
                        for i = 0 to pcSocialLinksCnt - 1
                          slId = pcSocialLinksArr(0, i)
                          slName = pcSocialLinksArr(1, i)
                          slImage = pcSocialLinksArr(2, i)
                          slCustomImage = pcSocialLinksArr(3, i)
                          slUrl = pcSocialLinksArr(4, i)
                          slAlt = pcSocialLinksArr(5, i)
                          slOrder = pcSocialLinksArr(6, i)
                          
                          if not isnumeric(slOrder) then
                            slOrder = i + 1
                          end if

                          %>
                            <li>
                              <div class="pcCPsortableHandle pcCPsocialImage">
                                <img src="images/social/<%= slImage %>" alt="<%= slName %>">
                                <input type="hidden" class="pcCPsortableOrder" name="SocialLink_Order<%= slId %>" value="<%= slOrder %>" />
                              </div>
                              
                              <div class="pcCPsocialLink">
                                <label>Link: </label>
                                <br/>
                                <input type="text" name="SocialLink_Url<%= slId %>" value="<%= slUrl %>"/>
                              </div>
                              
                              <div class="pcCPsocialMoreOptions">
	                              <br/>
                              	<a class="pcCPsocialMoreLink" href="#">More Options</a>
                              </div>
                              
                              <div class="pcCPsocialMore">                                
                              	<div class="pcCPsocialAlt">
                                  <label>Alt/Title: </label>
                                  <br/>
                                  <input type="text" name="SocialLink_Alt<%= slId %>" value="<%= slAlt %>"/>
                                </div>
                                
                                <div class="pcCPsocialCustImage">
                                  <label>Custom Icon (optional): </label>
                                  <br/>
                                  <input type="text" name="SocialLink_Image<%= slId %>" value="<%= slCustomImage %>" placeholder="<%= slImage %>" />
                                	<a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=SocialLink_Image<%= slId %>&fid=form1','window2')">
                                  	<img src="images/search.gif" alt="Locate uploaded images." class="pcCPchooseImage">
                                  </a>
                                </div>
                                
                                <% If Len(slCustomImage) > 0 And pcf_catalogItemExists(slCustomImage) Then %>
                                  <div class="pcCPsocialCustIcon">
                                  	<img src="../pc/catalog/<%= slCustomImage %>" alt="Custom <%= slName %> Icon" />
                                  </div>
                                <% End If %>
                              </div>
                              
                            </li>
                          <%
                        next
                      %>
                    </ul>
                  <% end if %>
                </td>
              </tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
              <tr>
              	<th colspan="2">Sharing Settings</th>
              </tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td valign="top" width="20%">
									<a href="https://www.addthis.com/get/original-sharing-buttons" target="_blank"><img src="images/AddThisLogo.png" alt="AddThis" border="0"></a>
								</td>
								<td>
									<div style="margin-bottom: 10px;">Show <a href="https://www.addthis.com/get/original-sharing-buttons" target="_blank">AddThis</a> Buttons:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=901"></a></div>
									<input type="radio" name="AddThisDisplay" value="0"<% If pcIntAddThisDisplay="0" then %> checked<% end if %> class="clearBorder"> Never &nbsp;
									<input type="radio" name="AddThisDisplay" value="1"<% If pcIntAddThisDisplay="1" then %> checked<% end if %> class="clearBorder"> Right of Page Title &nbsp;
									<input type="radio" name="AddThisDisplay" value="2"<% If pcIntAddThisDisplay="2" then %> checked<% end if %> class="clearBorder"> Below 'Add to Cart' section
								</td>
							</tr>
							<tr>
								<td></td>
								<td>
									<% if trim(pcStrAddThisCode)<>"" then %>
									<div style="margin-top: 10px; padding-top: 10px; border-top: 1px dashed #CCC;">
									You are currently using:
									<div style="margin: 10px 0;"><%=pcStrAddThisCode%></div>
									</div>
									<% end if %>
									<div style="margin-top: 10px; padding-top: 10px; border-top: 1px dashed #CCC;"><a href="https://www.addthis.com/get/original-sharing-buttons" target="_blank">Get the AddThis code</a> that best fits your needs, then <a href="JavaScript:;" onClick="document.getElementById('AddThisCodeDiv').style.display='';">paste it here</a>.</div>
									<div style="margin-top: 10px; display: none;" id="AddThisCodeDiv">
										<textarea name="AddThisCode" id="AddThisCode" cols="50" rows="6" onClick="selectFieldContent('AddThisCode')"><%=pcStrAddThisCode%></textarea>
									</div>
								</td>
							</tr>
              <tr>
                <td colspan="2"><hr></td>
              </tr>
              <tr>
                <td><img src="images/pinterest.JPG" width="112" height="38" alt="Pinterest"></td>
                <td>                  
                  Show Pinterest's &quot;Pin It&quot; Button:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=902"></a>
                  <br/>
                  <input type="radio" name="PinterestDisplay" value="1"<% If pcIntPinterestDisplay="1" then %> checked<% end if %> class="clearBorder"> 
      Enable &nbsp;
                  <input type="radio" name="PinterestDisplay" value="0"<% If pcIntPinterestDisplay="0" then %> checked<% end if %> class="clearBorder"> 
      Disable
                </td>
              </tr>
              <tr>
                <td></td>
                <td>
                  <div style="margin-top: 5px; padding-top: 10px; border-top: 1px dashed #CCC;">
                    Pin Counter: 	
                    <select name="PinterestCounter">
                    <option value="none" selected>Don't Show</option>
                    <option value="horizontal" <% if (pcStrPinterestCounter="horizontal") then %>selected<%end if%>>Beside</option>
                    <option value="vertical" <% if (pcStrPinterestCounter="vertical") then %>selected<% end if %>>Above</option>
                    </select>
                    <% pcs_RequiredImageTag "PinterestCounter", pcv_isPinterestCounterRequired %>
                  </div>
                </td>
              </tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
              <tr>
                <th colspan="2">Accepted Payment Types</th>
              </tr>
              <tr> 
                <td colspan="2" class="pcCPspacer"></td>
              </tr>
              <tr>
                <td colspan="2">
                  <p>
                    Select the accepted payment types you would like to display on the storefront.
                    <b>NOTE:</b> Drag-and-drop the payment icons below to change the display order.
                    &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=903"></a>
									</p>
                  <br/>
                  <p>To upload custom payment icons <a href="pcv4_image_upload.asp" target="_blank">click here</a>.</p>
                </td>
              </tr>
              <tr>
              	<td colspan="2"><hr></td>
              </tr>
              <tr>      
                <td colspan="2">
                	<% If pcAcceptedPaymentsCnt > 0 Then %>
                    <ul id="pcCPpaymentTypes" class="pcCPsortable">
                      <%
                        For i = 0 To pcAcceptedPaymentsCnt - 1
                          paymentID = pcAcceptedPayments(0, i)
                          paymentName = pcAcceptedPayments(1, i)
                          paymentImage = pcAcceptedPayments(2, i)
                          paymentCustomImage = pcAcceptedPayments(3, i)
                          paymentAlt = pcAcceptedPayments(4, i)
                          paymentActive = pcAcceptedPayments(5, i)
                          paymentOrder = pcAcceptedPayments(6, i)
													
													If IsNull(paymentActive) Then
														paymentActive = "False"
													End If
													
													If IsNull(paymentOrder) Then
														paymentOrder = i + 1
													End If
												%>
												<li>
													<div class="pcCPsortableHandle pcCPpaymentImage">
														<img src="images/payment/<%= paymentImage %>" alt="<%= paymentName %>" title="<%= paymentName %>"/>
                            <input type="hidden" class="pcCPsortableOrder" name="AcceptedPayment_Order<%= paymentID %>" value="<%= paymentOrder %>" />
													</div>
													
													<div class="pcCPpaymentEnable">
														<span><input type="radio" name="AcceptedPayment_Active<%= paymentID %>" value="1" class="clearBorder" <% If paymentActive = "True" Then Response.Write "checked" %>> Enable</span>
														<span><input type="radio" name="AcceptedPayment_Active<%= paymentID %>" value="0" class="clearBorder" <% If paymentActive <> "True" Then Response.Write "checked" %>> Disable</span>
													</div>
													
													<div class="pcCPpaymentLinks">
														<a href="#" class="pcCPpaymentCustomizeLink">Customize &raquo;</a>
													</div>

                          <br />
													
													<div class="pcCPpaymentCustom">
														<div class="pcCPpaymentAlt">
															<label>Alt/Title (optional): </label>
															<input type="text" name="AcceptedPayment_Alt<%= paymentID %>" value="<%= paymentAlt %>" />
														</div>
                            
														<div class="pcCPpaymentCustImage">
															<label>Custom Image (optional): </label>
															<input type="text" name="AcceptedPayment_Image<%= paymentID %>" value="<%= paymentCustomImage %>" placeholder="<%= paymentImage %>" />
                              <a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=AcceptedPayment_Image<%= paymentID %>&fid=form1','window2')">
                                <img src="images/search.gif" alt="Locate uploaded images." class="pcCPchooseImage">
                              </a>
                            </div>
                            
														<% If Len(paymentImage) > 0 And pcf_catalogItemExists(paymentImage) Then %>
                              <div class="pcCPpaymentCustIcon">
                                <img src="../pc/catalog/<%= paymentImage %>" alt="Custom <%= slName %> Icon" />
                              </div>
                            <% end if %>
													</div>
												</li>
                    		<%
												Next
											%>
                    </ul>
                  <% End If %>
                </td>
              </tr>
            </table>
          </div>
            
					<div id="tabs-6" class="tab-pane">
						<table class="pcCPcontent">
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2">Show/Hide Storefront Elements</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2">
								On a store with hundreds of categories and many category levels, you can use the following setting to improve the loading time for the storefront's <a href="../pc/search.asp" target="_blank">advanced search page</a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=410"></a>:
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<input name="hideCategory" type="radio" value="0" class="clearBorder" <% If pcIntHideCategory="0" or pcIntHideCategory="" then %>checked<%end if%>> Show all categories in the drop-down <br />
								<input name="hideCategory" type="radio" value="1" class="clearBorder" <% If pcIntHideCategory="1" then %>checked<%end if%>> Only show top level categories (<u>recommended</u> for stores with large category trees) <br />
								<input name="hideCategory" type="radio" value="-1" class="clearBorder" <% If pcIntHideCategory="-1" then %>checked<%end if%>> Hide the categories drop-down completely
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_68")%>:</td>
								<td>
								<% If pcIntDisplayStock="-1" then %>
								<input type="radio" name="displayStock" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="displayStock" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
								<input type="radio" name="displayStock" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="displayStock" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
							 </td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_69")%>:</td>
								<td>
								<% If pcIntShowStockLmt="-1" then %>
								<input type="radio" name="ShowStockLmt" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="ShowStockLmt" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
								<input type="radio" name="ShowStockLmt" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="ShowStockLmt" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_70")%>: </td>
								<td>
									<input type="radio" name="HideDiscField" value="0" <% If pcStrHideDiscField <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="HideDiscField" value="1" <% If pcStrHideDiscField = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=418"></a>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_70a")%>: </td>
								<td>
									<input type="radio" name="DispDiscCart" value="1" <% If pcStrDispDiscCart = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DispDiscCart" value="0" <% If pcStrDispDiscCart <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
									</a>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Enable &quot;Product Quick View&quot; Feature:</td>
								<td>
								<% If pcIntDisplayQuickView="1" then %>
									<input type="radio" name="DisplayQuickView" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisplayQuickView" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="DisplayQuickView" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisplayQuickView" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Enable &quot;Quick Buy&quot; Feature:</td>
								<td>
								<% If pcIntQuickBuy="1" then %>
									<input type="radio" name="QuickBuy" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="QuickBuy" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="QuickBuy" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="QuickBuy" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=464"></a>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Enable &quot;Stay on Page when Adding To Cart&quot; Feature:</td>
								<td>
								<% If pcIntATCEnabled="1" then %>
									<input type="radio" name="ATCEnabled" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="ATCEnabled" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="ATCEnabled" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="ATCEnabled" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=465"></a>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_27")%>:</td>
								<td>
								<% If pcIntWL=0 then %>
									<input type="radio" name="WL" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="WL" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="WL" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="WL" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								 &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=413"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Show "Previous" & "Next" item buttons on the product details page:</td>
								<td>
								
								<input type="radio" name="DisplayPNButtons" value="1" <% If pcIntDisplayPNButtons="1" then %>checked<%end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="DisplayPNButtons" value="0" <% If pcIntDisplayPNButtons<>"1" then %>checked<%end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
							 </td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Restore saved shopping cart on next visit:</td>
								<td>
								<% If pcIntRestoreCart="1" then %>
									<input type="radio" name="RestoreCart" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="RestoreCart" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="RestoreCart" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="RestoreCart" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=468"></a>
								</td>
							</tr>
                            <tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Enable Bulk Add on Category Pages:</td>
								<td>
								<input type="radio" name="enableBulkAdd" value="1" <% If pcIntEnableBulkAdd="1" then %>checked<%end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="enableBulkAdd" value="0" <% If pcIntEnableBulkAdd<>"1" then %>checked<%end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
							 </td>
							</tr>
							
							<tr>
								<td class="pcCPspacer" colspan="2" style="padding-top: 20px;"></td>
							</tr>
							<tr>
								<th colspan="2">Turn Features On/Off</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<%
							pcv_SEOURLCheck = scStoreURL&"/"&scPcFolder&"/pc/pcSEOTest.asp"
							pcv_SEOURLCheck = replace(pcv_SEOURLCheck,"//","/")
							pcv_SEOURLCheck = replace(pcv_SEOURLCheck,"http:/","http://")
							pcv_SEOURLCheck = replace(pcv_SEOURLCheck,"https:/","https://")
							%>
							<tr>
								<td>Use Keyword-rich URLs:
								<% If pcIntSeoURLs="1" then %>
									<input type="radio" name="SeoURLs" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="SeoURLs" value="0" class="clearBorder" onClick="JavaScript:alert('Turning off this feature can cause links to your store pages to return 404 Page Not Found errors. Make sure to update all navigation links so that they no longer use the keyword rich URLs as those URLs will no longer be working. Among other things, remember to re-generate the storefront navigation if you are using that feature, review header.asp and footer.asp, and any other page on your Web site that was linking to storefront pages.')"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
									<input type="radio" name="SeoURLs" value="1" class="clearBorder" onClick="JavaScript:alert('Make sure that you have changed the 404 error handler in your Web hosting account or directly on your dedicated Web server before enabling this feature. See the documentation for more information.')"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="SeoURLs" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=460"></a>
                                </td>
                                <td>
                                	<a href="<%=pcv_SEOURLCheck%>" class="btn btn-default" role="button" target="_blank">Test SEO Settings</a>
								</td>
							</tr>
							<tr>
								<td align="left">File name of &quot;Page Not Found&quot; page:</td>
								<td align="left">
								<input type="text" name="SeoURLs404" value="<%=pcStrSeoURLs404%>">
								<% pcs_RequiredImageTag "SeoURLs404", pcv_isSeoURLs404Required %>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=461"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><div class="pcCPmessageInfo">This feature requires that you <a href="http://wiki.productcart.com/productcart/seo-urls" target="_blank">carefully review the related documentation</a>.</div></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td valign="top">
									<a href="https://www.google.com/analytics/" target="_blank"><img src="images/ga_logo.png" alt="Google Analytics" border="0"></a>
								</td>
								<td>
									<div id="showGA" style="margin-top: 10px; width: 370px;<% If pcIntGAType = "2" Then%> display: none;<%end if%>">
                                        Enter your <a href="https://www.google.com/analytics/" target="_blank">Google Analytics Profile ID</a> to activate the integration:
                                        <div style="margin-top: 10px; margin-bottom: 15px;">Web site profile ID: <input type="text" name="GoogleAnalytics" id="GoogleAnalytics" value="<%=pcStrGoogleAnalytics%>" onClick="selectFieldContent('GoogleAnalytics')">&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=470"></a></div>
                                    </div>
									<div style="margin-top: 10px;">
										<input type="radio" name="GAType" value="1" <% If pcIntGAType = "1" Then%>checked<%end if%> class="clearBorder" onClick="$('#showGA').show(); $('#showGTM').hide();">Google Universal Analytics<br />
										<input type="radio" name="GAType" value="0" <% If pcIntGAType = "0" Then%>checked<%end if%> class="clearBorder" onClick="$('#showGA').show(); $('#showGTM').hide();">Google Analytics
                                        <hr />
                                        <input type="radio" name="GAType" value="2" <% If pcIntGAType = "2" Then%>checked<%end if%> class="clearBorder" onClick="$('#showGTM').show(); $('#showGA').hide();">Google Tag Manager
									</div>
                                    
                                    <div id="showGTM" style="margin-top: 10px; width: 370px;<% If pcIntGAType <> "2" Then%> display: none;<%end if%>">
                                    	Enter your <a href="https://tagmanager.google.com/" target="_blank">Google Tag Manager Container ID</a> to activate the integration:
                                        <div style="margin-top: 10px;">Container ID: <input type="text" name="GoogleTagManager" id="GoogleTagManager" value="<%=pcStrGoogleTagManager%>" onClick="selectFieldContent('GoogleTagManager')"></div>
                                    </div>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
                            <tr>
                            	<td valign="top">Enable Google Conversion Tracking:</td>
                                <td>
									<input type="radio" name="EnableGCT" value="1" <% If pcIntEnableGCT = "1" Then Response.Write "checked" %> class="clearBorder"onClick="$('#showGCT').show();"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="EnableGCT" value="0" <% If pcIntEnableGCT <> "1" Then Response.Write "checked" %> class="clearBorder" onClick="$('#showGCT').hide();"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
                                    <div id="showGCT" style="margin-top: 10px; width: 370px;<% If pcIntEnableGCT <> "1" Then%> display: none;<%end if%>">
                                    	Enter your <a href="https://support.google.com/adwords/answer/1722022?hl=en" target="_blank">Google Conversion Tracking</a> code to activate the integration:
                                        <div style="margin-top: 10px;">Google Conversion Tracking code: <textarea type="text" name="GCTCode" id="GCTCode" rows="15" style="margin: 10px 0px;" onClick="selectFieldContent('GCTCode')"><%=Server.HTMLEncode(pcStrGCTCode)%></textarea></div>
                                        <p><strong><font color="#FF0000">&lt;ORDER_TOTAL&gt;</font></strong> tag will be replaced by order total amount.</p>
                                    </div>
								</td>
                            </tr>
                            <tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_75")%>:</td>
								<td>
									<input type="radio" name="DisableGiftRegistry" value="0" <% If pcIntDisableGiftRegistry <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisableGiftRegistry" value="1" <% If pcIntDisableGiftRegistry = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_71")%>:</td>
								<td>
									<input type="radio" name="HideRMA" value="0" <% If pcIntHideRMA <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="HideRMA" value="1" <% If pcIntHideRMA = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=421"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_72")%>:</td>
								<td>
									<input type="radio" name="ShowHD" value="1" <% If pcIntShowHD = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="ShowHD" value="0" <% If pcIntShowHD <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=422"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_74")%>:</td>
								<td>
									<input type="radio" name="ErrorHandler" value="1" <% If pcIntErrorHandler = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="ErrorHandler" value="0" <% If pcIntErrorHandler <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=420"></a>
								</td>
							</tr>

							<tr>
								<td colspan="2"><hr></td>
							</tr>
                            <tr>
								<td>Enable 'Keep Admin Session Alive' feature:</td>
								<td>
									<input type="radio" name="KeepSession" value="1" <% If pcIntKeepSession = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="KeepSession" value="0" <% If pcIntKeepSession <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=483"></a>
								</td>
							</tr>
                            <tr>
								<td colspan="2"><hr></td>
							</tr>
                            <tr>
								<td>Enable "Combine &amp; Minify CSS / JavaScript":</td>
								<td>
									<input type="radio" name="EnableBundling" value="1" <% If pcIntEnableBundling = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="EnableBundling" value="0" <% If pcIntEnableBundling <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=481"></a>
								</td>
							</tr>
                            <input type="hidden" name="OptimizeJavascript" value="0">
                            <!--
                            <tr>
                                <td class="pcCPspacer" colspan="2"></td>
                            </tr>
                            <tr>
								<td colspan="2"><hr></td>
							</tr>                            
                            <tr>
								<td>Enable JavaScript Optimization:</td>
								<td>
									<input type="radio" name="OptimizeJavascript" value="1" <% If pcIntOptimizeJavascript = "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="OptimizeJavascript" value="0" <% If pcIntOptimizeJavascript <> "1" Then Response.Write "checked" %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=482"></a>
								</td>
							</tr>
                            -->
                            <tr>
                                <td class="pcCPspacer" colspan="2"></td>
                            </tr>                            
						</table>
						</div>
            
                    </div>
            
				</div>
				<script type=text/javascript>
                    $pc( "#TabbedPanels2" ).tab('show')
					<%if request("tab")<>"" and IsNumeric(request("tab")) then%>
                        $pc('#TabbedPanels2 li:eq(<%=request("tab")-1%>) a').tab('show');
					<%end if%>
				</script>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
				<p>
				  <input type="submit" name="updateSettings" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_107")%>" class="btn btn-primary">
				</p>
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->