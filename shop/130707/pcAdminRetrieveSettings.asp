<%
If scUpgrade = 1 Then
	pcStrScVersion=scVersion
	pcStrScSubVersion=scSubVersion
	pcStrScSP=scSP
	pcIntScUpgrade=scUpgrade
	pcStrCompanyName=scCompanyName
	pcStrCompanyAddress=scCompanyAddress
	pcStrCompanyZip=scCompanyZip
	pcStrCompanyCity=scCompanyCity
	pcStrCompanyState=scCompanyState
	pcStrCompanyCountry=scCompanyCountry
	pcStrCompanyLogo=scCompanyLogo
	pcIntQtyLimit=scQtyLimit
	pcIntAddLimit=scAddLimit
	pcIntPre=scPre
	pcIntCustPre=scCustPre
	pcIntCatImages=scCatImages
	pcIntShowStockLmt=scShowStockLmt
	pcIntOutOfStockPurchase=scOutOfStockPurchase
	pcStrCurSign=scCurSign
	pcStrDecSign=scDecSign
	pcStrDivSign=scDivSign
	pcStrDateFrmt=scDateFrmt
	pcIntMinPurchase=scMinPurchase
	pcIntWholesaleMinPurchase=scWholesaleMinPurchase
	pcStrURLredirect=scURLredirect
	pcStrSSL=scSSL
	pcStrSSLUrl=scSSLUrl
	pcStrIntSSLPage=scIntSSLPage
	pcIntPrdRow=scPrdRow
	pcIntPrdRowsPerPage=scPrdRowsPerPage
	pcIntCatRow=scCatRow
	pcIntCatRowsPerPage=scCatRowsPerPage
	pcStrBType=bType
	pcStrStoreOff=scStoreOff
	pcStrStoreMsg=scStoreMsg
	pcIntWL=scWL
	pcStrorderLevel=scorderLevel
	pcIntDisplayStock=scDisplayStock
	pcIntHideCategory=scHideCategory
	pcIntAllowNews=AllowNews
	pcIntNewsCheckOut=NewsCheckOut
	pcIntNewsReg=NewsReg
	pcStrNewsLabel=NewsLabel
	pcIntPCOrd=PCOrd
	pcIntHideSortPro=HideSortPro
	pcStrDFLabel=DFLabel
	pcStrDFShow=DFShow
	pcStrDFReq=DFReq
	pcStrTFLabel=TFLabel
	pcStrTFShow=TFShow
	pcStrTFReq=TFReq
	pcStrDTCheck=DTCheck
	pcStrDeliveryZip=DeliveryZip
	pcStrOrderName=scOrderName
	pcStrHideDiscField=scHideDiscField
	pcStrDispDiscCart=scDispDiscCart
	pcStrAllowSeparate=scAllowSeparate
	pcIntDisableDiscountCodes=scDisableDiscountCodes
	pcStrReferLabel=ReferLabel
	pcIntViewRefer=ViewRefer
	pcIntRefNewCheckout=RefNewCheckout
	pcIntRefNewReg=RefNewReg
	pcIntBrandLogo=sBrandLogo
	pcIntBrandPro=sBrandPro
	pcIntRewardsActive=RewardsActive
	pcIntRewardsIncludeWholesale=RewardsIncludeWholesale
	pcIntRewardsPercent=RewardsPercent
	pcStrRewardsLabel=RewardsLabel
	pcStrXML=scXML
	pcIntQDiscounttype=pcQDiscountType
	pcIntBTODisplayType=iBTODisplayType
	pcIntBTOOutofStockPurchase=iBTOOutofStockPurchase
	pcIntBTOShowImage=iBTOShowImage
	pcIntBTOShowMustIm=iBTOShowMustIm
	pcIntBTOQuote=iBTOQuote
	pcIntBTOQuoteSubmit=iBTOQuoteSubmit
	pcIntBTOQuoteSubmitOnly=iBTOQuoteSubmitOnly
	pcIntBTODetLinkType=iBTODetLinkType
	pcStrBTODetTxt=vBTODetTxt
	pcIntBTOPopWidth=iBTOPopWidth
	pcIntBTOPopHeight=iBTOPopHeight
	pcIntBTOPopImage=iBTOPopImage
	pcIntConfigPurchaseOnly=scConfigPurchaseOnly
	pcIntTerms=scTerms
	pcStrTermsLabel=scTermsLabel
	pcStrTermsCopy=scTermsCopy
	pcIntTermsShown=scTermsShown
	pcIntShowSKU=scShowSKU
	pcIntShowSmallImg=scShowSmallImg
	pcIntHideRMA=scHideRMA
	pcIntShowHD=scShowHD
	pcIntErrorHandler=scErrorHandler
	pcIntAllowCheckoutWR=scAllowCheckoutWR
	pcStrViewPrdStyle=scViewPrdStyle
	pcStrCustomerIPAlert=CustomerIPAlert
	pcStrCompanyPhoneNumber=scCompanyPhoneNumber
	pcStrCompanyFaxNumber=scCompanyFaxNumber
	pcIntDisableGiftRegistry=scDisableGiftRegistry
	'// Hard code to 0
	pcIntSeoURLs=0
	pcStrSeoURLs404=scSeoURLs404
	pcIntQuickBuy=scQuickBuy
	pcIntATCEnabled=scATCEnabled
	pcIntRestoreCart=scRestoreCart
	pcIntGuestCheckoutOpt=scGuestCheckoutOpt
	pcIntAddThisDisplay=scAddThisDisplay
	pcStrAddThisCode=scAddThisCode
	pcStrGoogleAnalytics=scGoogleAnalytics
	pcIntGAType=scGAType
	pcIntEnableGCT=scEnableGCT
	pcStrMetaTitle=scMetaTitle
	pcStrMetaDescription=scMetaDescription
	pcStrMetaKeywords=scMetaKeywords
	pcIntPinterestDisplay=scPinterestDisplay
	pcStrPinterestCounter=scPinterestCounter
	pcStrScRegistered = scRegistered
	pcIntDisplayQuickView=scDisplayQuickView
	pcIntDisplayPNButtons=scDisplayPNButtons
	pcStrUpgradeValue = scUpgradeValue
	pcIntConURL = scConURL
	pcIntCartStack = scCartStack
	pcStrCSSiteId = scCSSiteId
    pcIntEnableBundling = scEnableBundling
	pcIntOptimizeJavascript = scOptimizeJavascript
	pcIntKeepSession = scKeepSession
	pcStrGoogleTagManager = scGoogleTagManager
	pcStrSPhoneReg=SPhoneReq

	query="SELECT pcStoreSettings_TermsLabel, pcStoreSettings_TermsCopy FROM pcStoreSettings WHERE (((pcStoreSettings_ID)=1));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	pcStrTermsLabel=rs("pcStoreSettings_TermsLabel")
	pcStrTermsCopy=rs("pcStoreSettings_TermsCopy")
    set rs=nothing

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
Else
	query="SELECT pcStoreVersion_Num, pcStoreVersion_Sub, pcStoreVersion_SP FROM pcStoreVersions WHERE pcStoreVersion_ID=1;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if NOT rs.eof then

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
		
			call closeDb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		pcStrScVersion=rs("pcStoreVersion_Num")
		pcStrScSubVersion=rs("pcStoreVersion_Sub")
		pcStrScSP=rs("pcStoreVersion_SP")
	end if 

	if isNull(pcStrScVersion) OR pcStrScVersion&""="" then
		pcStrScVersion = scVersion
		pcStrScSubVersion = scSubVersion
		pcStrScSP = 0
	end if

	query="SELECT pcStoreSettings_CompanyName, pcStoreSettings_CompanyAddress, pcStoreSettings_CompanyZip, pcStoreSettings_CompanyCity, pcStoreSettings_CompanyState, pcStoreSettings_CompanyCountry, pcStoreSettings_CompanyLogo, pcStoreSettings_QtyLimit, pcStoreSettings_AddLimit, pcStoreSettings_Pre, pcStoreSettings_CustPre, pcStoreSettings_CatImages, pcStoreSettings_ShowStockLmt, pcStoreSettings_OutOfStockPurchase, pcStoreSettings_Cursign, pcStoreSettings_DecSign, pcStoreSettings_DivSign, pcStoreSettings_DateFrmt, pcStoreSettings_MinPurchase, pcStoreSettings_WholesaleMinPurchase, pcStoreSettings_URLredirect, pcStoreSettings_SSL, pcStoreSettings_SSLUrl, pcStoreSettings_IntSSLPage, pcStoreSettings_PrdRow, pcStoreSettings_PrdRowsPerPage, pcStoreSettings_CatRow, pcStoreSettings_CatRowsPerPage, pcStoreSettings_BType, pcStoreSettings_StoreOff, pcStoreSettings_StoreMsg, pcStoreSettings_WL, pcStoreSettings_orderLevel, pcStoreSettings_DisplayStock, pcStoreSettings_HideCategory, pcStoreSettings_AllowNews, pcStoreSettings_NewsCheckOut, pcStoreSettings_NewsReg, pcStoreSettings_NewsLabel, pcStoreSettings_PCOrd, pcStoreSettings_HideSortPro, pcStoreSettings_DFLabel, pcStoreSettings_DFShow, pcStoreSettings_DFReq, pcStoreSettings_TFLabel, pcStoreSettings_TFShow, pcStoreSettings_TFReq, pcStoreSettings_DTCheck, pcStoreSettings_DeliveryZip, pcStoreSettings_OrderName, pcStoreSettings_HideDiscField, pcStoreSettings_DispDiscCart, pcStoreSettings_AllowSeparate, pcStoreSettings_DisableDiscountCodes, pcStoreSettings_ReferLabel, pcStoreSettings_ViewRefer, pcStoreSettings_RefNewCheckout, pcStoreSettings_RefNewReg, pcStoreSettings_BrandLogo, pcStoreSettings_BrandPro, pcStoreSettings_RewardsActive, pcStoreSettings_RewardsIncludeWholesale, pcStoreSettings_RewardsPercent, pcStoreSettings_RewardsLabel, pcStoreSettings_XML, pcStoreSettings_QDiscounttype, pcStoreSettings_BTODisplayType, pcStoreSettings_BTOOutofStockPurchase, pcStoreSettings_BTOShowImage, pcStoreSettings_BTOShowMustIm, pcStoreSettings_BTOQuote, pcStoreSettings_BTOQuoteSubmit, pcStoreSettings_BTOQuoteSubmitOnly, pcStoreSettings_BTODetLinkType, pcStoreSettings_BTODetTxt, pcStoreSettings_BTOPopWidth, pcStoreSettings_BTOPopHeight, pcStoreSettings_BTOPopImage, pcStoreSettings_ConfigPurchaseOnly, pcStoreSettings_Terms, pcStoreSettings_TermsLabel, pcStoreSettings_TermsCopy, pcStoreSettings_TermsShown, pcStoreSettings_ShowSKU, pcStoreSettings_ShowSmallImg, pcStoreSettings_HideRMA, pcStoreSettings_ShowHD, pcStoreSettings_ErrorHandler, pcStoreSettings_AllowCheckoutWR, pcStoreSettings_ViewPrdStyle, pcStoreSettings_CustomerIPAlert, pcStoreSettings_CompanyPhoneNumber, pcStoreSettings_CompanyFaxNumber, pcStoreSettings_DisableGiftRegistry,  pcStoreSettings_SeoURLs, pcStoreSettings_SeoURLs404, pcStoreSettings_QuickBuy, pcStoreSettings_ATCEnabled, pcStoreSettings_RestoreCart, pcStoreSettings_GuestCheckoutOpt, pcStoreSettings_AddThisDisplay, pcStoreSettings_AddThisCode, pcStoreSettings_GoogleAnalytics, pcStoreSettings_MetaTitle, pcStoreSettings_MetaDescription, pcStoreSettings_MetaKeywords, pcStoreSettings_PinterestDisplay, pcStoreSettings_PinterestCounter, pcStoreSettings_DisplayQuickView, pcStoreSettings_PNButtons, pcStoreSettings_ThemeFolder, pcStoreSettings_ConURL, pcStoreSettings_GAType, pcStoreSettings_CartStack, pcStoreSettings_CSSiteId, pcStoreSettings_EnableBundling, pcStoreSettings_OptimizeJavascript, pcStoreSettings_GoogleTagManager, pcStoreSettings_KeepSession, pcStoreSettings_SPhoneReq, pcStoreSettings_EnableGCT, pcEnableBulkAdd FROM pcStoreSettings WHERE (((pcStoreSettings_ID)=1));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
	
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	pcStrCompanyName=rs("pcStoreSettings_CompanyName")
	pcStrCompanyAddress=rs("pcStoreSettings_CompanyAddress")
	pcStrCompanyZip=rs("pcStoreSettings_CompanyZip")
	pcStrCompanyCity=rs("pcStoreSettings_CompanyCity")
	pcStrCompanyState=rs("pcStoreSettings_CompanyState")
	pcStrCompanyCountry=rs("pcStoreSettings_CompanyCountry")
	pcStrCompanyLogo=rs("pcStoreSettings_CompanyLogo")
	pcIntQtyLimit=rs("pcStoreSettings_QtyLimit")
	pcIntAddLimit=rs("pcStoreSettings_AddLimit")
	pcIntPre=rs("pcStoreSettings_Pre")
	pcIntCustPre=rs("pcStoreSettings_CustPre")
	pcIntCatImages=rs("pcStoreSettings_CatImages")
	pcIntShowStockLmt=rs("pcStoreSettings_ShowStockLmt")
	pcIntOutOfStockPurchase=rs("pcStoreSettings_OutOfStockPurchase")
	pcStrCurSign=rs("pcStoreSettings_CurSign")
	pcStrDecSign=rs("pcStoreSettings_DecSign")
	pcStrDivSign=rs("pcStoreSettings_DivSign")
	pcStrDateFrmt=rs("pcStoreSettings_DateFrmt")
	pcIntMinPurchase=rs("pcStoreSettings_MinPurchase")
	pcIntWholesaleMinPurchase=rs("pcStoreSettings_WholesaleMinPurchase")
	pcStrURLredirect=rs("pcStoreSettings_URLredirect")
	pcStrSSL=rs("pcStoreSettings_SSL")
	pcStrSSLUrl=rs("pcStoreSettings_SSLUrl")
	pcStrIntSSLPage=rs("pcStoreSettings_IntSSLPage")
	pcIntPrdRow=rs("pcStoreSettings_PrdRow")
	pcIntPrdRowsPerPage=rs("pcStoreSettings_PrdRowsPerPage")
	pcIntCatRow=rs("pcStoreSettings_CatRow")
	pcIntCatRowsPerPage=rs("pcStoreSettings_CatRowsPerPage")
	pcStrBType=rs("pcStoreSettings_BType")
	pcStrStoreOff=rs("pcStoreSettings_StoreOff")
	pcStrStoreMsg=rs("pcStoreSettings_StoreMsg")
	pcIntWL=rs("pcStoreSettings_WL")
	pcStrorderLevel=rs("pcStoreSettings_orderLevel")
	pcIntDisplayStock=rs("pcStoreSettings_DisplayStock")
	pcIntHideCategory=rs("pcStoreSettings_HideCategory")
	pcIntAllowNews=rs("pcStoreSettings_AllowNews")
	pcIntNewsCheckOut=rs("pcStoreSettings_NewsCheckOut")
	pcIntNewsReg=rs("pcStoreSettings_NewsReg")
	pcStrNewsLabel=rs("pcStoreSettings_NewsLabel")
	pcIntPCOrd=rs("pcStoreSettings_PCOrd")
	pcIntHideSortPro=rs("pcStoreSettings_HideSortPro")
	pcStrDFLabel=rs("pcStoreSettings_DFLabel")
	pcStrDFShow=rs("pcStoreSettings_DFShow")
	pcStrDFReq=rs("pcStoreSettings_DFReq")
	pcStrTFLabel=rs("pcStoreSettings_TFLabel")
	pcStrTFShow=rs("pcStoreSettings_TFShow")
	pcStrTFReq=rs("pcStoreSettings_TFReq")
	pcStrDTCheck=rs("pcStoreSettings_DTCheck")
	pcStrDeliveryZip=rs("pcStoreSettings_DeliveryZip")
	pcStrOrderName=rs("pcStoreSettings_OrderName")
	pcStrHideDiscField=rs("pcStoreSettings_HideDiscField")
	pcStrDispDiscCart=rs("pcStoreSettings_DispDiscCart")
	pcStrAllowSeparate=rs("pcStoreSettings_AllowSeparate")
	pcIntDisableDiscountCodes=rs("pcStoreSettings_DisableDiscountCodes")
	pcStrReferLabel=rs("pcStoreSettings_ReferLabel")
	pcIntViewRefer=rs("pcStoreSettings_ViewRefer")
	pcIntRefNewCheckout=rs("pcStoreSettings_RefNewCheckout")
	pcIntRefNewReg=rs("pcStoreSettings_RefNewReg")
	pcIntBrandLogo=rs("pcStoreSettings_BrandLogo")
	pcIntBrandPro=rs("pcStoreSettings_BrandPro")
	pcIntRewardsActive=rs("pcStoreSettings_RewardsActive")
	pcIntRewardsIncludeWholesale=rs("pcStoreSettings_RewardsIncludeWholesale")
	pcIntRewardsPercent=rs("pcStoreSettings_RewardsPercent")
	pcStrRewardsLabel=rs("pcStoreSettings_RewardsLabel")
	pcStrXML=rs("pcStoreSettings_XML")
	pcIntQDiscounttype=rs("pcStoreSettings_QDiscountType")
	pcIntBTODisplayType=rs("pcStoreSettings_BTODisplayType")
	pcIntBTOOutofStockPurchase=rs("pcStoreSettings_BTOOutofStockPurchase")
	pcIntBTOShowImage=rs("pcStoreSettings_BTOShowImage")
	pcIntBTOShowMustIm=rs("pcStoreSettings_BTOShowMustIm")
	pcIntBTOQuote=rs("pcStoreSettings_BTOQuote")
	pcIntBTOQuoteSubmit=rs("pcStoreSettings_BTOQuoteSubmit")
	pcIntBTOQuoteSubmitOnly=rs("pcStoreSettings_BTOQuoteSubmitOnly")
	pcIntBTODetLinkType=rs("pcStoreSettings_BTODetLinkType")
	pcStrBTODetTxt=rs("pcStoreSettings_BTODetTxt")
	pcIntBTOPopWidth=rs("pcStoreSettings_BTOPopWidth")
	pcIntBTOPopHeight=rs("pcStoreSettings_BTOPopHeight")
	pcIntBTOPopImage=rs("pcStoreSettings_BTOPopImage")
	pcIntConfigPurchaseOnly=rs("pcStoreSettings_ConfigPurchaseOnly")
	pcIntTerms=rs("pcStoreSettings_Terms")
	pcStrTermsLabel=rs("pcStoreSettings_TermsLabel")
	pcStrTermsCopy=rs("pcStoreSettings_TermsCopy")
	pcIntTermsShown=rs("pcStoreSettings_TermsShown")
	pcIntShowSKU=rs("pcStoreSettings_ShowSKU")
	pcIntShowSmallImg=rs("pcStoreSettings_ShowSmallImg")
	pcIntHideRMA=rs("pcStoreSettings_HideRMA")
	pcIntShowHD=rs("pcStoreSettings_ShowHD")
	pcIntErrorHandler=rs("pcStoreSettings_ErrorHandler")
	pcIntAllowCheckoutWR=rs("pcStoreSettings_AllowCheckoutWR")
	pcStrViewPrdStyle=rs("pcStoreSettings_ViewPrdStyle")
	pcStrCustomerIPAlert=rs("pcStoreSettings_CustomerIPAlert")
	pcStrCompanyPhoneNumber=rs("pcStoreSettings_CompanyPhoneNumber")
	pcStrCompanyFaxNumber=rs("pcStoreSettings_CompanyFaxNumber")
	pcIntDisableGiftRegistry=rs("pcStoreSettings_DisableGiftRegistry")
	pcIntSeoURLs=rs("pcStoreSettings_SeoURLs")
	pcStrSeoURLs404=rs("pcStoreSettings_SeoURLs404")
	pcIntQuickBuy=rs("pcStoreSettings_QuickBuy")
	pcIntATCEnabled=rs("pcStoreSettings_ATCEnabled")
	pcIntRestoreCart=rs("pcStoreSettings_RestoreCart")
	pcIntGuestCheckoutOpt=rs("pcStoreSettings_GuestCheckoutOpt")
	pcIntAddThisDisplay=rs("pcStoreSettings_AddThisDisplay")
	pcStrAddThisCode=rs("pcStoreSettings_AddThisCode")
	pcStrGoogleAnalytics=rs("pcStoreSettings_GoogleAnalytics")
	pcIntGAType=rs("pcStoreSettings_GAType")
	pcIntEnableGCT=rs("pcStoreSettings_EnableGCT")
	pcStrMetaTitle=rs("pcStoreSettings_MetaTitle")
	pcStrMetaDescription=rs("pcStoreSettings_MetaDescription")
	pcStrMetaKeywords=rs("pcStoreSettings_MetaKeywords")
	pcIntPinterestDisplay=rs("pcStoreSettings_PinterestDisplay")
	pcStrPinterestCounter=rs("pcStoreSettings_PinterestCounter")
	pcStrScRegistered = scRegistered
	pcIntDisplayQuickView=rs("pcStoreSettings_DisplayQuickView")
	pcIntDisplayPNButtons=rs("pcStoreSettings_PNButtons")
	pcStrThemeFolder = rs("pcStoreSettings_ThemeFolder")
	pcIntConURL = rs("pcStoreSettings_ConURL")
	pcIntCartStack = rs("pcStoreSettings_CartStack")
	pcStrCSSiteId = rs("pcStoreSettings_CSSiteId")
	pcIntEnableBundling = rs("pcStoreSettings_EnableBundling")
	pcIntOptimizeJavascript = rs("pcStoreSettings_OptimizeJavascript")
	pcStrGoogleTagManager = rs("pcStoreSettings_GoogleTagManager")
	pcIntKeepSession = rs("pcStoreSettings_KeepSession")
	pcStrSPhoneReq=rs("pcStoreSettings_SPhoneReq")
    pcIntEnableBulkAdd=rs("pcEnableBulkAdd")
	set rs=nothing
End If

if pcStrXML="" then
	pcStrXML=scXML
end if
if IsNull(pcIntGAType) OR pcIntGAType="" then
	pcIntGAType="0"
end if
if IsNull(pcIntEnableGCT) OR pcIntEnableGCT="" then
	pcIntEnableGCT="0"
end if

'// First Upgrade / Check for NULL
if pcIntPinterestDisplay&""="" then
	pcIntPinterestDisplay="0"
end if
if pcStrScRegistered&""="" then
	pcStrScRegistered="124587"
end if
if pcIntDisplayQuickView&""="" then
	pcIntDisplayQuickView=0
end if
if pcIntDisplayPNButtons&""="" then
	pcIntDisplayPNButtons=1
end if
if pcStrThemeFolder&""="" then
	pcStrThemeFolder="theme/basic_blue"
end if
if pcIntConURL&""="" then
	pcIntConURL=0
end if

'// Get Social Links
query = "SELECT pcSocialLink_ID, pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_CustomImage, pcSocialLink_Url, pcSocialLink_Alt, pcSocialLink_Order FROM pcSocialLinks ORDER BY pcSocialLink_Order, pcSocialLink_Name"
set rs=conntemp.execute(query)
if not rs.eof then
	pcSocialLinksArr = rs.GetRows()
	pcSocialLinksCnt = UBound(pcSocialLinksArr, 2) + 1
else
	pcSocialLinksCnt = 0
end if
set rs = nothing

'// Get Accepted Payments items
query = "SELECT pcAcceptedPayment_ID, pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_CustomImage, pcAcceptedPayment_Alt, pcAcceptedPayment_Active, pcAcceptedPayment_Order FROM pcAcceptedPayments ORDER BY pcAcceptedPayment_Order, pcAcceptedPayment_Name"
set rs=conntemp.execute(query)
if not rs.eof then
	pcAcceptedPayments = rs.GetRows()
	pcAcceptedPaymentsCnt = UBound(pcAcceptedPayments, 2) + 1
else
	pcAcceptedPaymentsCnt = 0
end if
set rs = nothing

'// Read Google Conversion Tracking code from file
pcStrFolder = "../pc/theme/_common"
pcStrGCTCode = pcf_OpenUTF8(pcStrFolder & "\GCT.inc", pcStrFolder & "\GCT.inc")
%>
