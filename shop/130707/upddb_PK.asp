<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next 
PmAdmin=19
%>
<% pageTitle = "Database Update" %>
<% Section = "" %>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Sub CheckClustered(TableName,PrimaryField)
Dim rs,query

	query="SELECT [name],[TYPE],object_id FROM sys.indexes WHERE (([type]=2) OR ([type]=1)) AND object_id=OBJECT_ID('dbo." & TableName & "');"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		tmpName=rs("name")
		tmpType=rs("type")
		tmpOID=rs("object_id")
		if tmpType="2" then
			on error resume next
			query="ALTER TABLE dbo." & TableName & " DROP CONSTRAINT " & tmpName & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			if err.number=0 then
				query="ALTER TABLE dbo." & TableName & " ADD CONSTRAINT PK_" & TableName & " PRIMARY KEY CLUSTERED (" & PrimaryField & ");"
				set rs=connTemp.execute(query)
				set rs=nothing
			else
				err.number=0
				err.description=""
			end if
			on error goto 0
		end if
	else
		query="ALTER TABLE dbo." & TableName & " ADD CONSTRAINT PK_" & TableName & " PRIMARY KEY CLUSTERED (" & PrimaryField & ");"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	set rs=nothing

End Sub

Sub CreatePK(TableName,PrimaryField)
Dim rs,query
	on error resume next
	query="ALTER TABLE [dbo].[" & TableName & "] ADD [" & PrimaryField & "] [int] IDENTITY(1,1) NOT NULL;"
	set rs=connTemp.execute(query)
	set rs=nothing
	if err.number<>0 then
		err.number=0
		err.description=""
	end if
	on error goto 0
End Sub

call openDb()

'//Table used_discounts
call CreatePK("used_discounts","idUsedDisc")
call CheckClustered("used_discounts","idUsedDisc")

'//Table ups_license
call CreatePK("ups_license","idUPSLi")
call CheckClustered("ups_license","idUPSLi")

'//Table twoCheckout
call CheckClustered("twoCheckout","id_twocheckout")

'//Table tclink
call CreatePK("tclink","idTCL")
call CheckClustered("tclink","idTCL")

'//Table taxPrd
call CheckClustered("taxPrd","idTaxPerProduct")

'//Table taxLoc
call CheckClustered("taxLoc","idTaxPerPlace")

'//Table suppliers
call CreatePK("suppliers","suppliers_id")
call CheckClustered("suppliers","suppliers_id")

'//Table states
call CreatePK("states","idState")
call CheckClustered("states","idState")

'//Table shipService
call CheckClustered("shipService","idshipservice")

'//Table ShipmentTypes
call CheckClustered("ShipmentTypes","idShipment")

'//Table shipAlert
call CheckClustered("shipAlert","idShipAlert")

'//Table SB_Settings
call CheckClustered("SB_Settings","Setting_ID")

'//Table SB_Packages
call CheckClustered("SB_Packages","SB_PackageID")

'//Table SB_Orders
call CheckClustered("SB_Orders","SB_OrderID")

'//Table Referrer
call CheckClustered("Referrer","IdRefer")

'//Table recipients
call CheckClustered("recipients","idRecipient")

'//Table PSIGate
call CreatePK("PSIGate","idPSIG")
call CheckClustered("PSIGate","idPSIG")

'//Table protx
call CheckClustered("protx","idProtx")

'//Table ProductsOrdered
call CheckClustered("ProductsOrdered","idProductOrdered")

'//Table Products
call CheckClustered("Products","idProduct")

'//Table pfporders
call CheckClustered("pfporders","idpfporder")

'//Table Permissions
call CheckClustered("Permissions","IdPm")

'//Table pcXMLSettings
call CheckClustered("pcXMLSettings","pcXMLSet_ID")

'//Table pcXMLPartners
call CheckClustered("pcXMLPartners","pcXP_ID")

'//Table pcXMLLogs
call CheckClustered("pcXMLLogs","pcXL_id")

'//Table pcXMLIPs
call CheckClustered("pcXMLIPs","pcXIP_id")

'//Table pcXMLExportLogs
call CheckClustered("pcXMLExportLogs","pcXEL_ID")

'//Table pcVATRates
call CheckClustered("pcVATRates","pcVATRate_ID")

'//Table pcVATCountries
call CheckClustered("pcVATCountries","pcVATCountry_ID")

'//Table pcUPSPreferences
call CheckClustered("pcUPSPreferences","pcUPSPref_ID")

'//Table pcUploadFiles
call CheckClustered("pcUploadFiles","pcUpld_IDFile")

'//Table pcTaxZonesGroups
call CheckClustered("pcTaxZonesGroups","pcTaxZonesGroup_ID")

'//Table pcPay_PayPalAdvanced
call CheckClustered("pcPay_PayPalAdvanced","pcPay_PayPalAd_ID")

'//Table pcPay_PFL_Authorize
call CheckClustered("pcPay_PFL_Authorize","idPFL_Authorize")

'//Table pcTaxZones
call CheckClustered("pcTaxZones","pcTaxZone_ID")

'//Table pcTaxZoneRates
call CheckClustered("pcTaxZoneRates","pcTaxZoneRate_ID")

'//Table pcTaxZoneDescriptions
call CheckClustered("pcTaxZoneDescriptions","pcTaxZoneDesc_ID")

'//Table pcTaxGroups
call CheckClustered("pcTaxGroups","pcTaxGroup_ID")

'//Table pcTaxEptCust
call CheckClustered("pcTaxEptCust","pcTaxEptCust_ID")

'//Table pcTaxEpt
call CreatePK("pcTaxEpt","idTaxE")
call CheckClustered("pcTaxEpt","idTaxE")

'//Table pcTaskManager
call CheckClustered("pcTaskManager","pcTaskManager_id")

'//Table pcSuppliers
call CheckClustered("pcTaskManager","pcSupplier_ID")

'//Table pcStoreVersions
call CheckClustered("pcStoreVersions","pcStoreVersion_ID")

'//Table pcStoreSettings
call CheckClustered("pcStoreSettings","pcStoreSettings_ID")

'//Table pcSecurityKeys
call CheckClustered("pcSecurityKeys","pcSecurityKeyID")

'//Table pcSearchFields_Products
call CheckClustered("pcSearchFields_Products","idSearchFieldProduct")

'//Table pcSearchFields_Mappings
call CheckClustered("pcSearchFields_Mappings","idSearchFieldMapping")

'//Table pcSearchFields_Categories
call CheckClustered("pcSearchFields_Categories","idSearchFieldCategory")

'//Table pcSavedPrdStats
call CheckClustered("pcSavedPrdStats","pcSPS_ID")

'//Table pcRevSettings
call CreatePK("pcRevSettings","pcRS_ID")
call CheckClustered("pcRevSettings","pcRS_ID")

'//Table pcRevLists
call CreatePK("pcRevLists","pcRL_ID")
call CheckClustered("pcRevLists","pcRL_ID")

'//Table pcReviewSpecials
call CreatePK("pcReviewSpecials","pcRSP_ID")
call CheckClustered("pcReviewSpecials","pcRSP_ID")

'//Table pcReviewsData
call CreatePK("pcReviewsData","pcRD_ID")
call CheckClustered("pcReviewsData","pcRD_ID")

'//Table pcReviews
call CheckClustered("pcReviews","pcRev_IDReview")

'//Table pcReviewPoints
call CheckClustered("pcReviewPoints","pcRP_ID")

'//Table pcReviewNotifications
call CreatePK("pcReviewNotifications","pcRN_ID")
call CheckClustered("pcReviewNotifications","pcRN_ID")

'//Table pcRevFields
call CheckClustered("pcRevFields","pcRF_IDField")

'//Table pcRevExc
call CreatePK("pcRevExc","pcRE_ID")
call CheckClustered("pcRevExc","pcRE_ID")

'//Table pcRevBadWords
call CreatePK("pcRevBadWords","pcRBW_ID")
call CheckClustered("pcRevBadWords","pcRBW_ID")

'//Table PCReturns
call CheckClustered("PCReturns","idRMA")

'//Table pcProductsVATRates
call CheckClustered("pcProductsVATRates","pcProductsVATRates_ID")

'//Table pcProductsOrderedOptions
call CheckClustered("pcProductsOrderedOptions","ProdOrdOpt_ID")

'//Table pcProductsOptions
call CheckClustered("pcProductsOptions","pcProdOpt_ID")

'//Table pcProductsImages
call CheckClustered("pcProductsImages","pcProdImage_ID")

'//Table pcProductsExc
call CreatePK("pcProductsExc","pcPE_ID")
call CheckClustered("pcProductsExc","pcPE_ID")

'//Table pcPriority
call CheckClustered("pcPriority","pcPri_IDPri")

'//Table pcPPFCusts
call CheckClustered("pcPPFCusts","pcPPFCusts_id")

'//Table pcPPFCustPriceCats
call CheckClustered("pcPPFCustPriceCats","pcPPFCustPriceCats_id")

'//Table pcPay_USAePay_Orders
call CheckClustered("pcPay_USAePay_Orders","idePayOrder")

'//Table pcPay_USAePay
call CreatePK("pcPay_USAePay","pcPayUEP_ID")
call CheckClustered("pcPay_USAePay","pcPayUEP_ID")

'//Table pcPay_TripleDeal
call CreatePK("pcPay_TripleDeal","pcPayTD_ID")
call CheckClustered("pcPay_TripleDeal","pcPayTD_ID")

'//Table pcPay_SkipJack
call CreatePK("pcPay_SkipJack","pcPaySJ_ID")
call CheckClustered("pcPay_SkipJack","pcPaySJ_ID")

'//Table pcPay_SecPay
call CreatePK("pcPay_SecPay","pcPaySP_ID")
call CheckClustered("pcPay_SecPay","pcPaySP_ID")

'//Table pcPay_PayPal_Authorize
call CheckClustered("pcPay_PayPal_Authorize","idPayPal_Authorize")

'//Table pcPay_PayPal
call CreatePK("pcPay_PayPal","pcPayPP_ID")
call CheckClustered("pcPay_PayPal","pcPayPP_ID")

'//Table pcPay_PaymentExpress
call CreatePK("pcPay_PaymentExpress","pcPayPE_ID")
call CheckClustered("pcPay_PaymentExpress","pcPayPE_ID")

'//Table pcPay_Paymentech
call CheckClustered("pcPay_Paymentech","pcPay_PT_Id")

'//Table pcPay_ParaData
call CheckClustered("pcPay_ParaData","pcPay_ParaData_ID")


'//Table pcPay_OrdersMoneris
call CheckClustered("pcPay_OrdersMoneris","pcPay_MOrder_ID")

'//Table pcPay_NETOne
call CheckClustered("pcPay_NETOne","pcPay_NETOne_ID")

'//Table pcPay_Moneris
call CheckClustered("pcPay_Moneris","pcPay_Moneris_ID")

'//Table pcPay_LinkPointAPI
call CheckClustered("pcPay_LinkPointAPI","pcPay_LPAPI_ID")

'//Table pcPay_HSBC
call CheckClustered("pcPay_HSBC","pcPay_HSBC_ID")

'//Table pcPay_GestPay_Response
call CheckClustered("pcPay_GestPay_Response","ID")

'//Table pcPay_GestPay_OTP
call CheckClustered("pcPay_GestPay_OTP","pcPay_GestPay_OTP_id")

'//Table pcPay_GestPay
call CheckClustered("pcPay_GestPay","pcPay_GestPay_Id")

'//Table pcPay_FastCharge
call CheckClustered("pcPay_FastCharge","pcPay_FAC_ID")

'//Table pcPay_EPN
call CreatePK("pcPay_EPN","pcPayEPN_ID")
call CheckClustered("pcPay_EPN","pcPayEPN_ID")

'//Table pcPay_eMerchant
call CheckClustered("pcPay_eMerchant","pcPay_eMerch_ID")

'//Table pcPay_eMerch_Orders
call CheckClustered("pcPay_eMerch_Orders","pcPay_eMerch_Ord_ID")

'//Table pcPay_EIG_Vault
call CheckClustered("pcPay_EIG_Vault","pcPay_EIG_Vault_ID")

'//Table pcPay_EIG_Authorize
call CheckClustered("pcPay_EIG_Authorize","idauthorder")

'//Table pcPay_EIG
call CreatePK("pcPay_EIG","pcPayEIG_ID")
call CheckClustered("pcPay_EIG","pcPayEIG_ID")

'//Table pcPay_CyberSource
call CreatePK("pcPay_CyberSource","pcPayCys_Id")
call CheckClustered("pcPay_CyberSource","pcPayCys_Id")

'//Table pcPay_Chronopay
call CheckClustered("pcPay_Chronopay","CP_Id")

'//Table pcPay_Centinel_Orders
call CheckClustered("pcPay_Centinel_Orders","pcPay_CentOrd_ID")

'//Table pcPay_Centinel
call CreatePK("pcPay_Centinel","pcPayCent_ID")
call CheckClustered("pcPay_Centinel","pcPayCent_ID")

'//Table pcPay_CBN
call CreatePK("pcPay_CBN","pcPayCBN_id")
call CheckClustered("pcPay_CBN","pcPayCBN_id")

'//Table pcPay_ACHDirect
call CheckClustered("pcPay_ACHDirect","pcPay_ACH_ID")

'//Table pcPackageInfo
call CheckClustered("pcPackageInfo","pcPackageInfo_ID")

'//Table pcNewArrivalsSettings
call CheckClustered("pcNewArrivalsSettings","pcNAS_ID")

'//Table pcImageDirectory
call CheckClustered("pcImageDirectory","pcImgDir_ID")

'//Table pcHomePageSettings
call CheckClustered("pcHomePageSettings","pcHPS_ID")

'//Table pcGWSettings
call CheckClustered("pcGWSettings","pcGWSet_ID")

'//Table pcGWOptions
call CheckClustered("pcGWSettings","pcGW_IDOpt")

'//Table pcGCOrdered
call CreatePK("pcGCOrdered","pcGO_ID")
call CheckClustered("pcGCOrdered","pcGO_ID")

'//Table pcGC
call CreatePK("pcGC","pcGC_ID")
call CheckClustered("pcGC","pcGC_ID")

'//Table pcFTypes
call CheckClustered("pcFTypes","pcFType_IDType")

'//Table pcFStatus
call CheckClustered("pcFStatus","pcFStat_IDStatus")

'//Table pcExportGoogle
call CheckClustered("pcExportGoogle","pcExpG_ID")

'//Table pcExportCashback
call CheckClustered("pcExportCashback","pcExpCB_ID")

'//Table pcEvProducts
call CheckClustered("pcEvProducts","pcEP_ID")

'//Table pcEvents
call CheckClustered("pcEvents","pcEv_IDEvent")

'//Table pcErrorHandler
call CheckClustered("pcErrorHandler","pcErrorHandler_ID")

'//Table pcEDCTrans
call CheckClustered("pcEDCTrans","pcET_ID")

'//Table pcEDCSettings
call CheckClustered("pcEDCSettings","pcES_ID")

'//Table pcEDCLogs
call CheckClustered("pcEDCLogs","pcELog_ID")

'//Table pcDropShippersSuppliers
call CheckClustered("pcDropShippersSuppliers","pcDS_ID")

'//Table pcDropShippersOrders
call CheckClustered("pcDropShippersOrders","pcDropShipO_ID")

'//Table pcDropshippers
call CheckClustered("pcDropshippers","pcDropShipper_ID")

'//Table pcDFShip
call CreatePK("pcDFShip","pcFShip_ID")
call CheckClustered("pcDFShip","pcFShip_ID")

'//Table pcDFProds
call CreatePK("pcDFProds","pcFPro_ID")
call CheckClustered("pcDFProds","pcFPro_ID")

'//Table pcDFCusts
call CreatePK("pcDFCusts","pcFCust_ID")
call CheckClustered("pcDFCusts","pcFCust_ID")

'//Table pcDFCustPriceCats
call CheckClustered("pcDFCustPriceCats","pcFCPCat_ID")

'//Table pcDFCats
call CreatePK("pcDFCats","pcFCat_ID")
call CheckClustered("pcDFCats","pcFCat_ID")

'//Table pcCustomerTermsAgreed
call CheckClustered("pcCustomerTermsAgreed","pcCustomerTermsAgreed_ID")

'//Table pcCustomerSessions
call CheckClustered("pcCustomerSessions","idDbSession")

'//Table pcCustomerFieldsValues
call CheckClustered("pcCustomerFieldsValues","pcCFV_ID")

'//Table pcCustomerFields
call CheckClustered("pcCustomerFieldsValues","pcCField_ID")

'//Table pcCustomerCategories
call CheckClustered("pcCustomerCategories","idCustomerCategory")

'//Table pcCustFieldsPricingCats
call CheckClustered("pcCustFieldsPricingCats","pcCFPC_ID")

'//Table pcContents
call CheckClustered("pcContents","pcCont_IDPage")

'//Table pcComments
call CheckClustered("pcComments","pcComm_IdFeedback")

'//Table pcCC_Pricing
call CheckClustered("pcCC_Pricing","idCC_Price")

'//Table pcCC_BTO_Pricing
call CheckClustered("pcCC_BTO_Pricing","idCC_BTO_Price")

'//Table pcCatDiscounts
call CheckClustered("pcCatDiscounts","pcCD_IDDiscount")

'//Table pcCartArray
call CheckClustered("pcCartArray","pcCartArray_ID")

'//Table pcBTODefaultPriceCats
call CheckClustered("pcBTODefaultPriceCats","pcBDPC_id")

'//Table pcBestSellerSettings
call CheckClustered("pcBestSellerSettings","pcBSS_ID")

'//Table pcAmazonSettings
call CheckClustered("pcAmazonSettings","pcAmzSet_id")

'//Table pcAmazon
call CheckClustered("pcAmazon","pcAmz_id")

'//Table pcAffiliatesPayments
call CheckClustered("pcAffiliatesPayments","pcAffpay_idpayment")

'//Table pcAdminComments
call CheckClustered("pcAdminComments","pcACOM_ID")

'//Table pcAdminAuditLog
call CheckClustered("pcAdminAuditLog","pcAdminAuditLogID")

'//Table payTypes
call CheckClustered("payTypes","idPayment")

'//Table paypal
call CreatePK("paypal","paypal_id")
call CheckClustered("paypal","paypal_id")

'//Table orders
call CheckClustered("orders","idOrder")

'//Table optionsGroups
call CheckClustered("optionsGroups","idOptionGroup")

'//Table options_optionsGroups
call CheckClustered("options_optionsGroups","idoptoptgrp")

'//Table options
call CheckClustered("options","idOption")

'//Table optGrps
call CreatePK("optGrps","pcOG_ID")
call CheckClustered("optGrps","pcOG_ID")

'//Table offlinepayments
call CreatePK("offlinepayments","pcOffPay_ID")
call CheckClustered("offlinepayments","pcOffPay_ID")

'//Table News
call CheckClustered("News","idnews")

'//Table netbillorders
call CheckClustered("netbillorders","idnetbillorder")

'//Table netbill
call CreatePK("netbill","netbill_id")
call CheckClustered("netbill","netbill_id")

'//Table moneris
call CreatePK("moneris","pcMR_id")
call CheckClustered("moneris","pcMR_id")

'//Table linkpoint
call CreatePK("linkpoint","linkpoint_id")
call CheckClustered("linkpoint","linkpoint_id")

'//Table layout
call CheckClustered("layout","id")

'//Table klix
call CheckClustered("layout","idKlix")

'//Table ITransact
call CreatePK("ITransact","idIT_ID")
call CheckClustered("ITransact","idIT_ID")

'//Table InternetSecure
call CreatePK("InternetSecure","InterSec_ID")
call CheckClustered("InternetSecure","InterSec_ID")

'//Table icons
call CreatePK("icons","icons_id")
call CheckClustered("icons","icons_id")

'//Table FlatShipTypes
call CheckClustered("FlatShipTypes","idFlatShipType")

'//Table FlatShipTypeRules
call CheckClustered("FlatShipTypeRules","idFlatShipTypeRule")

'//Table FedExWSAPI
call CheckClustered("FedExWSAPI","FedExAPI_ID")

'//Table FedExAPI
call CheckClustered("FedExAPI","FedExAPI_ID")

'//Table fasttransact
call CheckClustered("fasttransact","id")

'//Table eWay
call CreatePK("eWay","eWay_ID")
call CheckClustered("eWay","eWay_ID")

'//Table emailSettings
call CheckClustered("emailSettings","id")

'//Table echo
call CreatePK("echo","echo_ID")
call CheckClustered("echo","echo_ID")

'//Table DProducts
call CreatePK("DProducts","DP_ID")
call CheckClustered("DProducts","DP_ID")

'//Table DPRequests
call CreatePK("DPRequests","DPR_ID")
call CheckClustered("DPRequests","DPR_ID")

'//Table DPLicenses
call CreatePK("DPLicenses","DPL_ID")
call CheckClustered("DPLicenses","DPL_ID")

'//Table discountsPerQuantity
call CheckClustered("discountsPerQuantity","idDiscountPerQuantity")

'//Table discounts
call CheckClustered("discounts","iddiscount")

'//Table customfields
call CheckClustered("customfields","idcustom")

'//Table customers
call CheckClustered("customers","idcustomer")

'//Table customCardTypes
call CheckClustered("customCardTypes","idCustomCardType")

'//Table customCardRules
call CheckClustered("customCardRules","idCustomCardRules")

'//Table customCardOrders
call CheckClustered("customCardOrders","idCCOrder")

'//Table cs_relationships
call CheckClustered("cs_relationships","idcrosssell")

'//Table crossSelldata
call CreatePK("crossSelldata","pcCSD_id")
call CheckClustered("crossSelldata","pcCSD_id")

'//Table creditCards
call CreatePK("creditCards","CCO_ID")
call CheckClustered("creditCards","CCO_ID")

'//Table countries
call CreatePK("countries","Country_ID")
call CheckClustered("countries","Country_ID")

'//Table configWishlistSessions
call CheckClustered("configWishlistSessions","idconfigWishlistSession")

'//Table configSpec_products
call CreatePK("configSpec_products","pcConfPro_ID")
call CheckClustered("configSpec_products","pcConfPro_ID")

'//Table configSpec_Charges
call CreatePK("configSpec_Charges","pcConfCha_ID")
call CheckClustered("configSpec_Charges","pcConfCha_ID")

'//Table configSpec_categories
call CreatePK("configSpec_categories","pcConfCat_ID")
call CheckClustered("configSpec_categories","pcConfCat_ID")

'//Table configSessions
call CheckClustered("configSessions","idconfigSession")

'//Table concord
call CheckClustered("concord","idConcord")

'//Table CCTypes
call CheckClustered("CCTypes","idCCType")

'//Table categories_products
call CreatePK("categories_products","CatPrd_ID")
call CheckClustered("categories_products","CatPrd_ID")

'//Table categories
call CheckClustered("categories","idCategory")

'//Table Brands
call CheckClustered("Brands","IdBrand")

'//Table BluePay
call CheckClustered("BluePay","idBluePay")

'//Table Blackout
call CreatePK("Blackout","Blackout_ID")
call CheckClustered("Blackout","Blackout_ID")

'//Table authorizeNet
call CreatePK("authorizeNet","pcAuNet_id")
call CheckClustered("authorizeNet","pcAuNet_id")

'//Table authorders
call CheckClustered("authorders","idauthorder")

'//Table affiliates
call CheckClustered("affiliates","idAffiliate")

'//Table admins
call CheckClustered("admins","id")

'//Table ZipCodeValidation
call CreatePK("ZipCodeValidation","ZCV_ID")
call CheckClustered("ZipCodeValidation","ZCV_ID")

'//Table xfields
call CheckClustered("xfields","idxfield")

'//Table WorldPay
call CreatePK("WorldPay","WorldPay_id")
call CheckClustered("WorldPay","WorldPay_id")

'//Table wishList
call CheckClustered("wishList","IDQuote")

'//Table verisign_pfp
call CheckClustered("verisign_pfp","id")

call closedb()

%>
<!--#include file="Adminheader.asp"-->
	<table class="pcCPcontent" width="100%">
	<tr>
		<td>
			<div class="pcCPmessageSuccess">
				Your database has been updated successfully!
			</div>
		</td>
	</tr>
	</table>
<!--#include file="AdminFooter.asp"-->