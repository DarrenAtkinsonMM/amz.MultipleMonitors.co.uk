<% 
if pcStrPageName<>"RPSettings.asp" then
	if RewardsPercent<>0 then
		pcIntRewardsPercent=RewardsPercent
	end if
end if


'// START ProductCart Sub-Version
Dim pcStrScSubVersion
pcStrScSubVersion = pcf_GetSubVersions()
'// END ProductCart Sub-Version	


If pcIntScUpgrade <> 1 Then

	query = "UPDATE pcStoreVersions SET pcStoreVersion_Num='"&removeReplaceSQ(pcStrScVersion)&"', pcStoreVersion_Sub='"&removeReplaceSQ(pcStrScSubVersion)&"' ,pcStoreVersion_SP="&removeReplaceSQ(pcStrScSP)&" WHERE pcStoreVersion_ID=1;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
	
		call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	set rs=nothing

	query="UPDATE pcStoreSettings SET pcStoreSettings_CompanyName=N'"&removeReplaceSQ(pcStrCompanyName)&"', pcStoreSettings_CompanyAddress=N'"&removeReplaceSQ(pcStrCompanyAddress)&"', pcStoreSettings_CompanyZip='"&removeReplaceSQ(pcStrCompanyZip)&"', pcStoreSettings_CompanyCity=N'"&removeReplaceSQ(pcStrCompanyCity)&"', pcStoreSettings_CompanyState=N'"&removeReplaceSQ(pcStrCompanyState)&"', pcStoreSettings_CompanyCountry='"&removeReplaceSQ(pcStrCompanyCountry)&"', pcStoreSettings_CompanyLogo='"&removeReplaceSQ(pcStrCompanyLogo)&"', pcStoreSettings_QtyLimit="&pcIntQtyLimit&", pcStoreSettings_AddLimit="&pcIntAddLimit&", pcStoreSettings_Pre="&pcIntPre&", pcStoreSettings_CustPre="&pcIntCustPre&", pcStoreSettings_CatImages="&pcIntCatImages&", pcStoreSettings_ShowStockLmt="&pcIntShowStockLmt&", pcStoreSettings_OutOfStockPurchase="&pcIntOutOfStockPurchase&", pcStoreSettings_Cursign=N'"&removeReplaceSQ(pcStrCurSign)&"', pcStoreSettings_DecSign='"&removeReplaceSQ(pcStrDecSign)&"', pcStoreSettings_DivSign='"&removeReplaceSQ(pcStrDivSign)&"', pcStoreSettings_DateFrmt='"&removeReplaceSQ(pcStrDateFrmt)&"', pcStoreSettings_MinPurchase="&pcIntMinPurchase&", pcStoreSettings_WholesaleMinPurchase="&pcIntWholesaleMinPurchase&", pcStoreSettings_URLredirect='"&removeReplaceSQ(pcStrURLredirect)&"', pcStoreSettings_SSL='"&removeReplaceSQ(pcStrSSL)&"', pcStoreSettings_SSLUrl='"&removeReplaceSQ(pcStrSSLUrl)&"', pcStoreSettings_IntSSLPage='"&removeReplaceSQ(pcStrIntSSLPage)&"', pcStoreSettings_PrdRow="&pcIntPrdRow&", pcStoreSettings_PrdRowsPerPage="&pcIntPrdRowsPerPage&",  pcStoreSettings_CatRow="&pcIntCatRow&", pcStoreSettings_CatRowsPerPage="&pcIntCatRowsPerPage&", pcStoreSettings_BType='"&removeReplaceSQ(pcStrBType)&"', pcStoreSettings_StoreOff='"&removeReplaceSQ(pcStrStoreOff)&"', pcStoreSettings_StoreMsg=N'"&removeReplaceSQ(pcStrStoreMsg)&"', pcStoreSettings_WL="&pcIntWL&", pcStoreSettings_orderLevel='"&removeReplaceSQ(pcStrorderLevel)&"', pcStoreSettings_DisplayStock="&pcIntDisplayStock&", pcStoreSettings_HideCategory="&pcIntHideCategory&", pcStoreSettings_AllowNews="&pcIntAllowNews&", pcStoreSettings_NewsCheckOut="&pcIntNewsCheckOut&", pcStoreSettings_NewsReg="&pcIntNewsReg&", pcStoreSettings_NewsLabel=N'"&removeReplaceSQ(pcStrNewsLabel)&"', pcStoreSettings_PCOrd="&pcIntPCOrd&", pcStoreSettings_HideSortPro="&pcIntHideSortPro&", pcStoreSettings_DFLabel=N'"&removeReplaceSQ(pcStrDFLabel)&"', pcStoreSettings_DFShow='"&removeReplaceSQ(pcStrDFShow)&"', pcStoreSettings_DFReq='"&removeReplaceSQ(pcStrDFReq)&"', pcStoreSettings_TFLabel=N'"&removeReplaceSQ(pcStrTFLabel)&"', pcStoreSettings_TFShow='"&removeReplaceSQ(pcStrTFShow)&"', pcStoreSettings_TFReq='"&removeReplaceSQ(pcStrTFReq)&"', pcStoreSettings_DTCheck='"&removeReplaceSQ(pcStrDTCheck)&"', pcStoreSettings_DeliveryZip='"&removeReplaceSQ(pcStrDeliveryZip)&"', pcStoreSettings_OrderName=N'"&removeReplaceSQ(pcStrOrderName)&"', pcStoreSettings_HideDiscField='"&removeReplaceSQ(pcStrHideDiscField)&"', pcStoreSettings_DispDiscCart='"&removeReplaceSQ(pcStrDispDiscCart)&"', pcStoreSettings_AllowSeparate='"&removeReplaceSQ(pcStrAllowSeparate)&"', pcStoreSettings_ReferLabel=N'"&removeReplaceSQ(pcStrReferLabel)&"', pcStoreSettings_ViewRefer="&pcIntViewRefer&", pcStoreSettings_RefNewCheckout="&pcIntRefNewCheckout&", pcStoreSettings_RefNewReg="&pcIntRefNewReg&", pcStoreSettings_BrandLogo="&pcIntBrandLogo&", pcStoreSettings_BrandPro="&pcIntBrandPro&", pcStoreSettings_RewardsActive="&pcIntRewardsActive&", pcStoreSettings_RewardsIncludeWholesale="&pcIntRewardsIncludeWholesale&", pcStoreSettings_RewardsPercent="&pcIntRewardsPercent&", pcStoreSettings_RewardsLabel=N'"&removeReplaceSQ(pcStrRewardsLabel)&"', pcStoreSettings_XML='"&removeReplaceSQ(pcStrXML)&"', pcStoreSettings_QDiscounttype="&pcIntQDiscountType&", pcStoreSettings_BTODisplayType="&pcIntBTODisplayType&", pcStoreSettings_BTOOutofStockPurchase="&pcIntBTOOutofStockPurchase&", pcStoreSettings_BTOShowImage="&pcIntBTOShowImage&", pcStoreSettings_BTOQuote="&pcIntBTOQuote&", pcStoreSettings_BTOQuoteSubmit="&pcIntBTOQuoteSubmit&", pcStoreSettings_BTOQuoteSubmitOnly="&pcIntBTOQuoteSubmitOnly&", pcStoreSettings_BTODetLinkType="&pcIntBTODetLinkType&", pcStoreSettings_BTODetTxt=N'"&removeReplaceSQ(pcStrBTODetTxt)&"', pcStoreSettings_BTOPopWidth="&pcIntBTOPopWidth&", pcStoreSettings_BTOPopHeight="&pcIntBTOPopHeight&", pcStoreSettings_BTOPopImage="&pcIntBTOPopImage&", pcStoreSettings_ConfigPurchaseOnly="&pcIntConfigPurchaseOnly&", pcStoreSettings_Terms="&pcIntTerms&", pcStoreSettings_TermsLabel=N'"&removeReplaceSQ(pcStrTermsLabel)&"', pcStoreSettings_TermsCopy=N'"&removeReplaceSQ(pcStrTermsCopy)&"', pcStoreSettings_TermsShown="&pcIntTermsShown&", pcStoreSettings_ShowSKU="&pcIntShowSKU&", pcStoreSettings_ShowSmallImg="&pcIntShowSmallImg&", pcStoreSettings_HideRMA="&pcIntHideRMA&", pcStoreSettings_ShowHD="&pcIntShowHD&", pcStoreSettings_ErrorHandler="&pcIntErrorHandler&", pcStoreSettings_AllowCheckoutWR="&pcIntAllowCheckoutWR&", pcStoreSettings_ViewPrdStyle='"&removeReplaceSQ(pcStrViewPrdStyle)&"', pcStoreSettings_CustomerIPAlert='"&removeReplaceSQ(pcStrCustomerIPAlert)&"', pcStoreSettings_CompanyPhoneNumber='"&removeReplaceSQ(pcStrCompanyPhoneNumber)&"', pcStoreSettings_CompanyFaxNumber='"&removeReplaceSQ(pcStrCompanyFaxNumber)&"', pcStoreSettings_DisableGiftRegistry="&pcIntDisableGiftRegistry&", pcStoreSettings_DisableDiscountCodes="&pcIntDisableDiscountCodes&", pcStoreSettings_SeoURLs="&pcIntSeoURLs&", pcStoreSettings_SeoURLs404='"&removeReplaceSQ(pcStrSeoURLs404)&"', pcStoreSettings_QuickBuy="&pcIntQuickBuy&", pcStoreSettings_ATCEnabled="&pcIntATCEnabled&", pcStoreSettings_RestoreCart="&pcIntRestoreCart&",pcStoreSettings_GuestCheckoutOpt=" & pcIntGuestCheckoutOpt &",pcStoreSettings_AddThisDisplay="&pcIntAddThisDisplay&",pcStoreSettings_AddThisCode='"&removeReplaceSQ(pcStrAddThisCode)&"',pcStoreSettings_PinterestDisplay="&pcIntPinterestDisplay&", pcStoreSettings_PinterestCounter='"&removeReplaceSQ(pcStrPinterestCounter)&"',pcStoreSettings_GoogleAnalytics='"&removeReplaceSQ(pcStrGoogleAnalytics)&"', pcStoreSettings_MetaTitle=N'"&removeReplaceSQ(pcStrMetaTitle)&"', pcStoreSettings_MetaDescription=N'"&removeReplaceSQ(pcStrMetaDescription)&"', pcStoreSettings_MetaKeywords=N'"&removeReplaceSQ(pcStrMetaKeywords)&"',pcStoreSettings_DisplayQuickView=" & pcIntDisplayQuickView & ",pcStoreSettings_PNButtons=" & pcIntDisplayPNButtons & ",pcStoreSettings_ThemeFolder='" & pcStrThemeFolder & "',pcStoreSettings_ConURL=" & pcIntConURL & ",pcStoreSettings_GAType=" & pcIntGAType & ", pcStoreSettings_CartStack=" & pcIntCartStack & ", pcStoreSettings_CSSiteId='" & pcStrCSSiteId & "', pcStoreSettings_EnableBundling=" & pcIntEnableBundling & ", pcStoreSettings_OptimizeJavascript=" & pcIntOptimizeJavascript & ", pcStoreSettings_GoogleTagManager='" & pcStrGoogleTagManager & "', pcStoreSettings_KeepSession=" & pcIntKeepSession & ", pcStoreSettings_BTOShowMustIm=" & pcIntBTOShowMustIm & ", pcStoreSettings_SPhoneReq='"&removeReplaceSQ(pcStrSPhoneReq)&"', pcStoreSettings_EnableGCT="&pcIntEnableGCT&", pcEnableBulkAdd=" & pcIntEnableBulkAdd & " WHERE (((pcStoreSettings_ID)=1));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
	
		call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	set rs=nothing
End If


'// Only Update this Section on AdminSettings.asp
If pcPageName="AdminSettings.asp" Then
    
    '// Save Social Links
    for i = 0 to pcSocialLinksCnt - 1
        slId = pcSocialLinksArr(0, i)
        
        itemArray = pcSocialLink_Items.Item(slId)
        
        slImage = itemArray(0)
        slUrl = itemArray(1)
        slAlt = itemArray(2)
        slOrder = itemArray(3)
        
        if len(slOrder) > 0 then
            query = "UPDATE pcSocialLinks SET pcSocialLink_CustomImage = '" & slImage & "', pcSocialLink_Url = '" & slUrl & "', pcSocialLink_Alt = '" & slAlt & "', pcSocialLink_Order = " & slOrder & " WHERE pcSocialLink_ID = " & slId & ";" 
        end if
        conntemp.execute(query)
    next
    
    '// Save Payment Options
    for i = 0 to pcAcceptedPaymentsCnt - 1
        paymentId = pcAcceptedPayments(0, i)
        
        itemArray = pcAcceptedPayments_Items.Item(paymentId)
        
        paymentImage = itemArray(0)
        paymentAlt = itemArray(1)
        paymentActive = itemArray(2)
        paymentOrder = itemArray(3)
        
        If paymentActive = "" Then
            paymentActive = 0
        End If
        
        query = "UPDATE pcAcceptedPayments SET pcAcceptedPayment_CustomImage = '" & paymentImage & "', pcAcceptedPayment_Alt = '" & paymentAlt & "', pcAcceptedPayment_Active = " & paymentActive & ", pcAcceptedPayment_Order = " & paymentOrder & " WHERE pcAcceptedPayment_ID = " & paymentId & ";" 
        conntemp.execute(query)
    next

End If

	
'/////////////////////////////////////////////////////
'// Write all changes to Settings.asp file
'/////////////////////////////////////////////////////
set StringBuilderObj = new StringBuilder

StringBuilderObj.append CHR(60)&CHR(37)&"'// Storewide Settings //" & vbCrLf
StringBuilderObj.append "private const scVersion = """&pcStrScVersion&"""" & vbCrLf
StringBuilderObj.append "private const scSubVersion = """&pcStrScSubVersion&"""" & vbCrLf
StringBuilderObj.append "private const scSP = """&pcStrScSP&"""" & vbCrLf
StringBuilderObj.append "private const scUpgrade = " & pcIntScUpgrade & vbCrLf
StringBuilderObj.append "private const scRegistered = """&pcStrScRegistered&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyName = """&removeSQ(pcStrCompanyName)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyAddress = """&removeSQ(pcStrCompanyAddress)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyZip = """&removeSQ(pcStrCompanyZip)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyCity = """&removeSQ(pcStrCompanyCity)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyState = """&removeSQ(pcStrCompanyState)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyCountry = """&removeSQ(pcStrCompanyCountry)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyPhoneNumber = """&removeSQ(pcStrCompanyPhoneNumber)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyFaxNumber = """&removeSQ(pcStrCompanyFaxNumber)&"""" & vbCrLf
StringBuilderObj.append "private const scCompanyLogo = """&removeSQ(pcStrCompanyLogo)&"""" & vbCrLf
StringBuilderObj.append "private const scMetaTitle = """&ClearHTMLTags2(removeSQ(pcStrMetaTitle),0)&"""" & vbCrLf
StringBuilderObj.append "private const scMetaDescription = """&ClearHTMLTags2(removeSQ(pcStrMetaDescription),0)&"""" & vbCrLf
StringBuilderObj.append "private const scMetaKeywords = """&ClearHTMLTags2(removeSQ(pcStrMetaKeywords),0)&"""" & vbCrLf
StringBuilderObj.append "private const scQtyLimit = "&pcIntQtyLimit&"" & vbCrLf
StringBuilderObj.append "private const scAddLimit = "&pcIntAddLimit&"" & vbCrLf
StringBuilderObj.append "private const scPre = "&pcIntPre&"" & vbCrLf
StringBuilderObj.append "private const scCustPre = "&pcIntCustPre&"" & vbCrLf
StringBuilderObj.append "private const scBTO = "&pcIntBTO&"" & vbCrLf
StringBuilderObj.append "private const scAPP = "&pcIntAPP&"" & vbCrLf
StringBuilderObj.append "private const scCM = "&pcIntCM&"" & vbCrLf
StringBuilderObj.append "private const scMS = "&pcIntMS&"" & vbCrLf
StringBuilderObj.append "private const scCatImages = "&pcIntCatImages&"" & vbCrLf
StringBuilderObj.append "private const scShowStockLmt = "&pcIntShowStockLmt&"" & vbCrLf
StringBuilderObj.append "private const scOutOfStockPurchase = "&pcIntOutOfStockPurchase&"" & vbCrLf
StringBuilderObj.append "private const scCurSign = """&removeSQ(pcStrCurSign)&"""" & vbCrLf
StringBuilderObj.append "private const scDecSign = """&removeSQ(pcStrDecSign)&"""" & vbCrLf
StringBuilderObj.append "private const scDivSign = """&removeSQ(pcStrDivSign)&"""" & vbCrLf
StringBuilderObj.append "private const scDateFrmt = """&removeSQ(pcStrDateFrmt)&"""" & vbCrLf
StringBuilderObj.append "private const scMinPurchase = "&pcIntMinPurchase&"" & vbCrLf
StringBuilderObj.append "private const scWholesaleMinPurchase = "&pcIntWholesaleMinPurchase&"" & vbCrLf
StringBuilderObj.append "private const scURLredirect = """&removeSQ(pcStrURLredirect)&"""" & vbCrLf
StringBuilderObj.append "private const scSSL = """&removeSQ(pcStrSSL)&"""" & vbCrLf
StringBuilderObj.append "private const scSSLUrl = """&removeSQ(pcStrSSLUrl)&"""" & vbCrLf
StringBuilderObj.append "private const scIntSSLPage = """&removeSQ(pcStrIntSSLPage)&"""" & vbCrLf
StringBuilderObj.append "private const scPrdRow = "&pcIntPrdRow&"" & vbCrLf
StringBuilderObj.append "private const scPrdRowsPerPage = "&pcIntPrdRowsPerPage&"" & vbCrLf
StringBuilderObj.append "private const scCatRow = "&pcIntCatRow&"" & vbCrLf
StringBuilderObj.append "private const scCatRowsPerPage = "&pcIntCatRowsPerPage&"" & vbCrLf
StringBuilderObj.append "private const bType = """&removeSQ(pcStrBType)&"""" & vbCrLf
StringBuilderObj.append "private const scStoreOff = """&removeSQ(pcStrStoreOff)&"""" & vbCrLf
StringBuilderObj.append "private const scStoreMsg = """&removeSQ(pcStrStoreMsg)&"""" & vbCrLf
StringBuilderObj.append "private const scWL = "&pcIntWL&"" & vbCrLf
StringBuilderObj.append "private const scorderLevel = """&removeSQ(pcStrorderLevel)&"""" & vbCrLf
StringBuilderObj.append "private const scDisplayStock = "&pcIntDisplayStock&"" & vbCrLf
StringBuilderObj.append "private const scHideCategory = "&pcIntHideCategory&"" & vbCrLf
StringBuilderObj.append "private const AllowNews = "&pcIntAllowNews&"" & vbCrLf
StringBuilderObj.append "private const NewsCheckOut = "&pcIntNewsCheckOut&"" & vbCrLf
StringBuilderObj.append "private const NewsReg = "&pcIntNewsReg&"" & vbCrLf
StringBuilderObj.append "private const NewsLabel = """&removeSQ(pcStrNewsLabel)&"""" & vbCrLf
StringBuilderObj.append "private const PCOrd = "&pcIntPCOrd&"" & vbCrLf
StringBuilderObj.append "private const HideSortPro = "&pcIntHideSortPro&"" & vbCrLf
StringBuilderObj.append "private const scViewPrdStyle = """&removeSQ(pcStrViewPrdStyle)&"""" & vbCrLf
StringBuilderObj.append "private const SPhoneReq = """&removeSQ(pcStrSPhoneReq)&"""" & vbCrLf
StringBuilderObj.append "private const DFLabel = """&removeSQ(pcStrDFLabel)&"""" & vbCrLf
StringBuilderObj.append "private const DFShow = """&removeSQ(pcStrDFShow)&"""" & vbCrLf
StringBuilderObj.append "private const DFReq = """&removeSQ(pcStrDFReq)&"""" & vbCrLf
StringBuilderObj.append "private const TFLabel = """&removeSQ(pcStrTFLabel)&"""" & vbCrLf
StringBuilderObj.append "private const TFShow = """&removeSQ(pcStrTFShow)&"""" & vbCrLf
StringBuilderObj.append "private const TFReq = """&removeSQ(pcStrTFReq)&"""" & vbCrLf
StringBuilderObj.append "private const DTCheck = """&removeSQ(pcStrDTCheck)&"""" & vbCrLf
StringBuilderObj.append "private const DeliveryZip = """&removeSQ(pcStrDeliveryZip)&"""" & vbCrLf
StringBuilderObj.append "private const CustomerIPAlert = """&removeSQ(pcStrCustomerIPAlert)&"""" & vbCrLf
StringBuilderObj.append "private const scOrderName = """&removeSQ(pcStrOrderName)&"""" & vbCrLf
StringBuilderObj.append "private const scHideDiscField = """&removeSQ(pcStrHideDiscField)&"""" & vbCrLf
StringBuilderObj.append "private const scDispDiscCart = """&removeSQ(pcStrDispDiscCart)&"""" & vbCrLf
StringBuilderObj.append "private const scAllowSeparate = """&removeSQ(pcStrAllowSeparate)&"""" & vbCrLf
StringBuilderObj.append "private const ReferLabel = """&removeSQ(pcStrReferLabel)&"""" & vbCrLf
StringBuilderObj.append "private const ViewRefer = "&pcIntViewRefer&"" & vbCrLf
StringBuilderObj.append "private const RefNewCheckout = "&pcIntRefNewCheckout&"" & vbCrLf
StringBuilderObj.append "private const RefNewReg = "&pcIntRefNewReg&"" & vbCrLf
StringBuilderObj.append "private const sBrandLogo = "&pcIntBrandLogo&"" & vbCrLf
StringBuilderObj.append "private const sBrandPro = "&pcIntBrandPro&"" & vbCrLf
StringBuilderObj.append "private const RewardsActive = "&pcIntRewardsActive&"" & vbCrLf
StringBuilderObj.append "private const RewardsIncludeWholesale = "&pcIntRewardsIncludeWholesale&"" & vbCrLf
StringBuilderObj.append "private const RewardsPercent = "&pcIntRewardsPercent&"" & vbCrLf
StringBuilderObj.append "private const RewardsLabel = """&removeSQ(pcStrRewardsLabel)&"""" & vbCrLf
StringBuilderObj.append "private const pcQDiscountType = "&pcIntQDiscountType&"" & vbCrLf
StringBuilderObj.append "private const iBTODisplayType = "&pcIntBTODisplayType&"" & vbCrLf
StringBuilderObj.append "private const iBTOOutofStockPurchase = "&pcIntBTOOutofStockPurchase&"" & vbCrLf
StringBuilderObj.append "private const iBTOShowImage = "&pcIntBTOShowImage&"" & vbCrLf
StringBuilderObj.append "private const iBTOShowMustIm = "&pcIntBTOShowMustIm&"" & vbCrLf
StringBuilderObj.append "private const iBTOQuote = "&pcIntBTOQuote&"" & vbCrLf
StringBuilderObj.append "private const iBTOQuoteSubmit = "&pcIntBTOQuoteSubmit&"" & vbCrLf
StringBuilderObj.append "private const iBTOQuoteSubmitOnly = "&pcIntBTOQuoteSubmitOnly&"" & vbCrLf
StringBuilderObj.append "private const iBTODetLinkType = "&pcIntBTODetLinkType&"" & vbCrLf
StringBuilderObj.append "private const vBTODetTxt = """&removeSQ(pcStrBTODetTxt)&"""" & vbCrLf
StringBuilderObj.append "private const iBTOPopWidth = "&pcIntBTOPopWidth&"" & vbCrLf
StringBuilderObj.append "private const iBTOPopHeight = "&pcIntBTOPopHeight&"" & vbCrLf
StringBuilderObj.append "private const iBTOPopImage = "&pcIntBTOPopImage&"" & vbCrLf
StringBuilderObj.append "private const scConfigPurchaseOnly = "&pcIntConfigPurchaseOnly&"" & vbCrLf
StringBuilderObj.append "private const scTerms = "&pcIntTerms&"" & vbCrLf
StringBuilderObj.append "private const scTermsShown = "&pcIntTermsShown&"" & vbCrLf
StringBuilderObj.append "private const scShowSKU = "&pcIntShowSKU&"" & vbCrLf
StringBuilderObj.append "private const scShowSmallImg = "&pcIntShowSmallImg&"" & vbCrLf
StringBuilderObj.append "private const scHideRMA = "&pcIntHideRMA&"" & vbCrLf
StringBuilderObj.append "private const scShowHD = "&pcIntShowHD&"" & vbCrLf
StringBuilderObj.append "private const scErrorHandler = "&pcIntErrorHandler&"" & vbCrLf
StringBuilderObj.append "private const scDisableGiftRegistry = """&pcIntDisableGiftRegistry&"""" & vbCrLf
StringBuilderObj.append "private const scDisableDiscountCodes = """&pcIntDisableDiscountCodes&"""" & vbCrLf
StringBuilderObj.append "private const scAllowCheckoutWR = "&pcIntAllowCheckoutWR&"" & vbCrLf
StringBuilderObj.append "private const scSeoURLs = "&pcIntSeoURLs&"" & vbCrLf
StringBuilderObj.append "private const scSeoURLs404 = """&removeSQ(pcStrSeoURLs404)&"""" & vbCrLf
StringBuilderObj.append "private const scDisplayQuickView = "&pcIntDisplayQuickView&"" & vbCrLf
StringBuilderObj.append "private const scDisplayPNButtons = "&pcIntDisplayPNButtons&"" & vbCrLf
StringBuilderObj.append "private const scConURL = "&pcIntConURL&"" & vbCrLf
StringBuilderObj.append "private const scQuickBuy = "&pcIntQuickBuy&"" & vbCrLf
StringBuilderObj.append "private const scATCEnabled = "&pcIntATCEnabled&"" & vbCrLf
StringBuilderObj.append "private const scRestoreCart = "&pcIntRestoreCart&"" & vbCrLf
StringBuilderObj.append "private const scXML = """&removeSQ(pcStrXML)&"""" & vbCrLf
StringBuilderObj.append "private const scGuestCheckoutOpt = "&pcIntGuestCheckoutOpt&"" & vbCrLf
StringBuilderObj.append "private const scAddThisDisplay = "&pcIntAddThisDisplay&"" & vbCrLf
StringBuilderObj.append "private const scPinterestDisplay = "&pcIntPinterestDisplay&"" & vbCrLf
StringBuilderObj.append "private const scPinterestCounter = """&removeSQ(pcStrPinterestCounter)&"""" & vbCrLf
StringBuilderObj.append "private const scGoogleAnalytics = """&removeSQ(pcStrGoogleAnalytics)&"""" & vbCrLf
StringBuilderObj.append "private const scGAType = """&pcIntGAType&"""" & vbCrLf
StringBuilderObj.append "private const scEnableGCT = """&pcIntEnableGCT&"""" & vbCrLf
StringBuilderObj.append "private const scCartStack = """&pcIntCartStack&"""" & vbCrLf
StringBuilderObj.append "private const scCSSiteId = """&pcStrCSSiteId&"""" & vbCrLf
StringBuilderObj.append "private const scEnableBundling = """&pcIntEnableBundling&"""" & vbCrLf
StringBuilderObj.append "private const scOptimizeJavascript = """&pcIntOptimizeJavascript&"""" & vbCrLf
StringBuilderObj.append "private const scKeepSession = """&pcIntKeepSession&"""" & vbCrLf
StringBuilderObj.append "private const scGoogleTagManager = """&removeSQ(pcStrGoogleTagManager)&"""" & vbCrLf
StringBuilderObj.append "private const enableBulkAdd = "&pcIntEnableBulkAdd&"" & vbCrLf
StringBuilderObj.append "'// Storewide Settings // " &CHR(37)&CHR(62)& vbCrLf

call pcs_SaveUTF8("/"&scPcFolder&"/includes/settings.asp","../includes/settings.asp",StringBuilderObj.toString)

set StringBuilderObj=nothing

' Save Google Conversion Tracking code to file
if len(pcStrGCTCode)>0 Then
    pcStrGCTCode = pcf_ReplaceCharacters(pcStrGCTCode)
	call pcs_SaveUTF8("/"&scPcFolder&"/pc/theme/_common/GCT.inc","../pc/theme/_common/GCT.inc", pcStrGCTCode)
End If

if (pcStrnewStoreURL<>scStoreURL) AND (pcStrnewStoreURL<>"") then
	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	
	if PPD="1" then
		pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/storeconstants.asp")
	else
		pcStrFileName=Server.Mappath ("../includes/storeconstants.asp")
	end if
	
	Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
	objFile.WriteLine CHR(60)&CHR(37)
	objFile.WriteLine "private const scCrypPass=""" & scCrypPass & """"
	objFile.WriteLine "private const scDSN=""" & scDSN & """"
	objFile.WriteLine "private const scDB=""" & scDB & """"
	objFile.WriteLine "private const scStoreURL=""" & pcStrnewStoreURL & """"
	objFile.WriteLine CHR(37)&CHR(62)
	objFile.Close
	set objFS=nothing
	set objFile=nothing
end if

If pcIntScUpgrade = 1 And pcPageName="AdminSettings.asp" Then 
	Response.Redirect pcPageName
End If
%>
