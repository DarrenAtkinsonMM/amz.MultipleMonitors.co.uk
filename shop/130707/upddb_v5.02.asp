<% 
PmAdmin=19
pageTitle = "ProductCart v5.x to v5.02 - Database Update" 
Section = "" 
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="fixedNTextConst.asp"-->
<% 
On Error Resume Next
dim conntemp1

IF request("action")="sql" then
	if request("hmode")="2" then
		SSIP=request("SSIP")
		UID=request("UID")
		PWD=request("PWD")
		SSDB=request("SSDB")
		if SSIP="" or UID="" or PWD="" then
			call closeDb()
            response.redirect "upddb_v5.02.asp?mode=3"
			response.End
		end if
		set connTemp=server.createobject("adodb.connection")
		connTemp.Open "Provider=sqloledb;Data Source="&SSIP&";Initial Catalog="&SSDB&";User Id="&UID&";Password="&PWD
		if err.number <> 0 then
			call closeDb()
            response.redirect "techErr.asp?error="&Server.Urlencode("Error while opening database")
		end if
	else
		if instr(ucase(scDSN),"DSN=") then
			call closeDb()
            response.redirect "upddb_v5.02.asp?mode=1"
			response.End
		end if
		
	end if
	
	iCnt=0
	ErrStr=""
	
	'========================================================================
	'// START OF DB UPDATES FOR v4.1
	'========================================================================
	
        '// ALTER EXISTING TABLES
        call AlterTableSQL("admins", "ADD", "adm_ContactName", "[NVarChar](250)", 0, "", "0")	
        call AlterTableSQL("admins", "ADD", "adm_ContactEmail", "[NVarChar](250)", 0, "", "0")	
        call AlterTableSQL("customers","ADD","pcCust_AllowReviewEmails","[INT]","1","1","0") '//Reviews
        call AlterTableSQL("pcContents", "ADD", "pcCont_HideBackButton", "[INT]", 1, "0", "0")
        call AlterTableSQL("pcContents", "ADD", "pcCont_Draft", "[ntext]", 0, "", "0")
        call AlterTableSQL("pcContents", "ADD", "pcCont_DraftStatus", "[INT]", 1, "0", "0")
        call AlterTableSQL("pcGWSettings", "ADD", "pcGWSet_OverviewCart", "[INT]", 1, "0", "0")
        call AlterTableSQL("pcGWSettings", "ADD", "pcGWSet_HTMLCart", "[ntext]", 0, "", "0")
        call AlterTableSQL("pcRecentRevSettings","ADD","pcRR_ReviewsPerProduct","[INT]","1","1","0") '//Reviews
        call AlterTableSQL("pcReviews","ADD","pcRev_IDCustomer","[INT]","1","1","0") '//Reviews
        call AlterTableSQL("pcReviews","ADD","pcRev_IDOrder","[INT]","1","0","0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_SendReviewReminder","[INT]", 1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_sendReviewReminderDays","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_sendReviewReminderType","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_sendReviewReminderFormat","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_sendReviewReminderTemplate","[nvarchar] (255)","0","","0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_RewardForReview","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_RewardForReviewURL","[nvarchar] (255)","0","","0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_RewardForReviewFirstPts","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_RewardForReviewAdditionalPts","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_RewardForReviewMinLength","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_RewardForReviewMaxPts","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_DisplayRatings","[INT]",1, "0", "0") '//Reviews
        call AlterTableSQL("pcRevSettings","ADD","pcRS_LastRunDate","[datetime]","0","","0") '//Reviews
        call AlterTableSQL("pcStoreSettings", "ADD", "pcStoreSettings_GuestCheckoutOpt", "[INT]", 1, "0", "0")
        call AlterTableSQL("pcStoreSettings", "ADD", "pcStoreSettings_RestoreCart", "[INT]", 1, "1", "0")
        call AlterTableSQL("products", "ADD", "pcPrd_MojoZoom", "[INT]", 1, "0", "0")
        call AlterTableSQL("products","ADD","pcProd_AvgRating","[float]", 1, "0", "0") '//Reviews
        call AlterTableSQL("layout", "ADD", "pcLO_Update", "[NVarChar](155)", 2, "images/sample/pc_button_update.gif", "0")
        call AlterTableSQL("pcCustomerCategories", "ADD", "pcCC_NFSoverride","[INT]",1, "0", "0") '//Not for sale override (Private shopping club)

		'// START - ENDICIA
        call AlterTableSQL("pcPackageInfo", "ADD", "pcPackageInfo_Endicia", "[INT]", 1, "0", "0")
        call AlterTableSQL("pcPackageInfo", "ADD", "pcPackageInfo_EndiciaLabelFile", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcPackageInfo", "ADD", "pcPackageInfo_EndiciaIsPIC", "[INT]", 1, "0", "0")
        call AlterTableSQL("pcPackageInfo", "ADD", "pcPackageInfo_EndiciaExp", "[DateTime]", 0, "", "0")
		'// END - ENDICIA

		'// PAY TYPES
        call AlterTableSQL("payTypes", "ADD", "pcPayTypes_Subscription", "[INT]", 1, "0", "1")	
        
        '// SUB LEVEL
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_LinkID", "[nvarchar] (250)", 0, "", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcSubscription_ID", "[INT]", 1, "0", "1")			
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubFrequency", "[INT]", 1, "0", "1")		
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubPeriod", "[nvarchar] (20)", 0, "", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubCycles", "[INT]", 1, "0", "1")	
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubTrialFrequency", "[INT]", 1, "0", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubTrialPeriod", "[nvarchar] (20)", 0, "", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubTrialCycles", "[INT]", 1, "0", "1")		
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_IsTrial", "[INT]", 1, "0", "1")		
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubAmount", "[float]", 1, "0", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubTrialAmount", "[INT]", 1, "0", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubAgree", "[INT]", 1, "0", "1")		
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubType", "[nvarchar] (20)", 0, "", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_NoShipping", "[INT]", 1, "0", "1")
    
        '// Not used yet
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubStartDate", "[nvarchar] (50)", 0, "", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubDetails", "[nvarchar] (250)", 0, "", "1")			
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubUPDStartDate", "[nvarchar] (50)", 0, "", "1")
        call AlterTableSQL("ProductsOrdered", "ADD", "pcPO_SubActive", "[INT]", 1, "0", "1")
        
        '// ORDER LEVEL
        call AlterTableSQL("orders", "ADD", "pcOrd_SubTax", "[float]", 1, "0", "1")
        call AlterTableSQL("orders", "ADD", "pcOrd_SubTrialTax", "[float]", 1, "0", "1")
        call AlterTableSQL("orders", "ADD", "pcOrd_SubShipping", "[float]", 1, "0", "1")
        call AlterTableSQL("orders", "ADD", "pcOrd_SubTrialShipping", "[float]", 1, "0", "1") '// not for immediate use
        
        '// CUSTOMER SESSION
        call AlterTableSQL("pcCustomerSessions", "ADD", "pcCustSession_SB_taxAmount", "[float]", 1, "0", "1")
        
        '// SAVED CART ARRAY
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray36", "[nvarchar] (200)", 0, "", "1")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray37", "[nvarchar] (200)", 0, "", "1")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray38", "[nvarchar] (200)", 0, "", "1")
        '// END - SubscriptionBridge Integration
        
        '// SAVED CART ARRAY
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray39", "[nvarchar] (200)", 0, "", "0")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray40", "[nvarchar] (200)", 0, "", "0")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray41", "[nvarchar] (200)", 0, "", "0")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray42", "[nvarchar] (200)", 0, "", "0")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray43", "[nvarchar] (200)", 0, "", "0")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray44", "[nvarchar] (200)", 0, "", "0")
        call AlterTableSQL("pcSavedCartArray", "ADD", "SCArray45", "[nvarchar] (200)", 0, "", "0")
        
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_36", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_37", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_38", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_39", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_40", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_41", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_42", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_43", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_44", "[nvarchar] (250)", 0, "", "0")
        call AlterTableSQL("pcCartArray", "ADD", "pcCartArray_45", "[nvarchar] (250)", 0, "", "0")
        
        '// START - PAYPAL 
        call AlterTableSQL("pcPay_PayPal", "ADD", "pcPay_PayPal_CardTypes", "[NVarChar](250)", 0, "", "0")
        call AlterTableSQL("pcPay_Centinel", "ADD", "pcPay_Cent_Password", "[NVarChar](250)", 0, "", "0")
        '// END - PAYPAL
            
        '// Alter field size
        call AlterTableSQL("orders", "ALTER COLUMN", "pcOrd_shippingPhone", "[NVarChar](30)", 0, "", "0")
        call AlterTableSQL("products", "ALTER COLUMN", "emailText", "[NVarChar](250)", 0, "", "0")
        
        '// Changes already included in SP3. Included again in case this is a direct update from v4.
        call AlterTableSQL("pcPay_PayPal", "ADD", "pcPay_PayPal_Subject", "[NVarChar](250)", 0, "", "0")
        call AlterTableSQL("payTypes", "ADD", "pcPayTypes_ppab", "[INT]", 1, "0", "1")	

        '// Start converting ntext to nvarchar(MAX)
        call AlterTableSQL("ProductsOrdered","ALTER COLUMN","xfdetails","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("ProductsOrdered","ALTER COLUMN","pcPrdOrd_SelectedOptions","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("ProductsOrdered","ALTER COLUMN","pcPrdOrd_OptionsPriceArray","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("ProductsOrdered","ALTER COLUMN","pcPrdOrd_OptionsArray","[nvarchar](max)", 0, "","0")
        
        call AlterTableSQL("pcSavedCartArray","ALTER COLUMN","SCArray21","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcSavedCartArray","ALTER COLUMN","SCArray1","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcSavedCartArray SET SCArray21=SCArray21,SCArray1=SCArray1;"
        set rs=connTemp.execute(query)
        set rs=nothing        
        
        query="UPDATE ProductsOrdered SET xfdetails=xfdetails,pcPrdOrd_SelectedOptions=pcPrdOrd_SelectedOptions,pcPrdOrd_OptionsPriceArray=pcPrdOrd_OptionsPriceArray,pcPrdOrd_OptionsArray=pcPrdOrd_OptionsArray;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("Products","ALTER COLUMN","details","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("Products","ALTER COLUMN","sDesc","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("Products","ALTER COLUMN","pcProd_MetaDesc","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("Products","ALTER COLUMN","pcProd_MetaKeywords","[nvarchar](max)", 0, "","0")
        
        query="UPDATE Products SET details=details,sDesc=sDesc,pcProd_MetaDesc=pcProd_MetaDesc,pcProd_MetaKeywords=pcProd_MetaKeywords;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcXMLLogs","ALTER COLUMN","pcXL_RequestXML","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcXMLLogs","ALTER COLUMN","pcXL_ResponseXML","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcXMLLogs SET pcXL_RequestXML=pcXL_RequestXML,pcXL_ResponseXML=pcXL_ResponseXML;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcUploadFiles","ALTER COLUMN","pcUpld_FileName","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcUploadFiles SET pcUpld_FileName=pcUpld_FileName;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcTaxEpt","ALTER COLUMN","pcTEpt_ProductList","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcTaxEpt","ALTER COLUMN","pcTEpt_CategoryList","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcTaxEpt SET pcTEpt_ProductList=pcTEpt_ProductList,pcTEpt_CategoryList=pcTEpt_CategoryList;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcSuppliers","ALTER COLUMN","pcSupplier_NoticeMsg","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcSuppliers SET pcSupplier_NoticeMsg=pcSupplier_NoticeMsg;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcStoreSettings","ALTER COLUMN","pcStoreSettings_StoreMsg","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcStoreSettings","ALTER COLUMN","pcStoreSettings_TermsCopy","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcStoreSettings SET pcStoreSettings_StoreMsg=pcStoreSettings_StoreMsg,pcStoreSettings_TermsCopy=pcStoreSettings_TermsCopy;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcReviewSpecials","ALTER COLUMN","pcRS_FieldList","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcReviewSpecials","ALTER COLUMN","pcRS_FieldOrder","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcReviewSpecials","ALTER COLUMN","pcRS_Required","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcReviewSpecials SET pcRS_FieldList=pcRS_FieldList,pcRS_FieldOrder=pcRS_FieldOrder,pcRS_Required=pcRS_Required;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcReviewsData","ALTER COLUMN","pcRD_Comment","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcReviewsData SET pcRD_Comment=pcRD_Comment;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("PCReturns","ALTER COLUMN","rmaReturnReason","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("PCReturns","ALTER COLUMN","rmaReturnStatus","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("PCReturns","ALTER COLUMN","rmaIdProducts","[nvarchar](max)", 0, "","0")
        
        query="UPDATE PCReturns SET rmaReturnReason=rmaReturnReason,rmaReturnStatus=rmaReturnStatus,rmaIdProducts=rmaIdProducts;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcRecentRevSettings","ALTER COLUMN","pcRR_PageDesc","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcRecentRevSettings SET pcRR_PageDesc=pcRR_PageDesc;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcPay_GestPay_Response","ALTER COLUMN","CUSTOMINFO","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcPay_GestPay_Response SET CUSTOMINFO=CUSTOMINFO;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcPackageInfo","ALTER COLUMN","pcPackageInfo_UPSNotifyEmailMsg","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcPackageInfo","ALTER COLUMN","pcPackageInfo_Comments","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcPackageInfo SET pcPackageInfo_UPSNotifyEmailMsg=pcPackageInfo_UPSNotifyEmailMsg,pcPackageInfo_Comments=pcPackageInfo_Comments;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcNewArrivalsSettings","ALTER COLUMN","pcNAS_PageDesc","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcNewArrivalsSettings SET pcNAS_PageDesc=pcNAS_PageDesc;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcHomePageSettings","ALTER COLUMN","pcHPS_PageDesc","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcHomePageSettings SET pcHPS_PageDesc=pcHPS_PageDesc;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcGWSettings","ALTER COLUMN","pcGWSet_HTML","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcGWSettings","ALTER COLUMN","pcGWSet_HTMLCart","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcGWSettings SET pcGWSet_HTML=pcGWSet_HTML,pcGWSet_HTMLCart=pcGWSet_HTMLCart;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcEvProducts","ALTER COLUMN","pcEP_xdetails","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcEvProducts SET pcEP_xdetails=pcEP_xdetails;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcErrorHandler","ALTER COLUMN","pcErrorHandler_ErrDescription","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcErrorHandler SET pcErrorHandler_ErrDescription=pcErrorHandler_ErrDescription;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcEDCLogs","ALTER COLUMN","pcELog_Request","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcEDCLogs","ALTER COLUMN","pcELog_Response","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcEDCLogs SET pcELog_Request=pcELog_Request,pcELog_Response=pcELog_Response;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcDropshippers","ALTER COLUMN","pcDropShipper_NoticeMsg","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcDropshippers SET pcDropShipper_NoticeMsg=pcDropShipper_NoticeMsg;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcCustomerSessions","ALTER COLUMN","pcCustSession_ShippingArray","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcCustomerSessions","ALTER COLUMN","pcCustSession_Comment","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcCustomerSessions SET pcCustSession_ShippingArray=pcCustSession_ShippingArray,pcCustSession_Comment=pcCustSession_Comment;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcCustomerFieldsValues","ALTER COLUMN","pcCFV_Value","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcCustomerFieldsValues SET pcCFV_Value=pcCFV_Value;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcCustomerFields","ALTER COLUMN","pcCField_Description","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcCustomerFields SET pcCField_Description=pcCField_Description;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcContents","ALTER COLUMN","pcCont_Description","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcContents","ALTER COLUMN","pcCont_Draft","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcContents SET pcCont_Description=pcCont_Description,pcCont_Draft=pcCont_Draft;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcComments","ALTER COLUMN","pcComm_Description","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("pcComments","ALTER COLUMN","pcComm_Details","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcComments SET pcComm_Description=pcComm_Description,pcComm_Details=pcComm_Details;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcBestSellerSettings","ALTER COLUMN","pcBSS_PageDesc","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcBestSellerSettings SET ppcBSS_PageDesc=pcBSS_PageDesc;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcAffiliatesPayments","ALTER COLUMN","pcAffpay_Status","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcAffiliatesPayments SET pcAffpay_Status=pcAffpay_Status;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("pcAdminComments","ALTER COLUMN","pcACOM_Comments","[nvarchar](max)", 0, "","0")
        
        query="UPDATE pcAdminComments SET pcACOM_Comments=pcACOM_Comments;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("payTypes","ALTER COLUMN","emailText","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("payTypes","ALTER COLUMN","terms","[nvarchar](max)", 0, "","0")
        
        query="UPDATE payTypes SET emailText=emailText,terms=terms;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("orders","ALTER COLUMN","details","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("orders","ALTER COLUMN","comments","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("orders","ALTER COLUMN","taxDetails","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("orders","ALTER COLUMN","adminComments","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("orders","ALTER COLUMN","pcOrd_GcReMsg","[nvarchar](max)", 0, "","0")
        
        query="UPDATE orders SET details=details,comments=comments,taxDetails=taxDetails,adminComments=adminComments,pcOrd_GcReMsg=pcOrd_GcReMsg;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("News","ALTER COLUMN","msgbody","[nvarchar](max)", 0, "","0")
        
        query="UPDATE News SET msgbody=msgbody;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("emailSettings","ALTER COLUMN","ConfirmEmail","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("emailSettings","ALTER COLUMN","PayPalEmail","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("emailSettings","ALTER COLUMN","ReceivedEmail","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("emailSettings","ALTER COLUMN","ShippedEmail","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("emailSettings","ALTER COLUMN","CancelledEmail","[nvarchar](max)", 0, "","0")
    
        query="UPDATE emailSettings SET ConfirmEmail=ConfirmEmail,PayPalEmail=PayPalEmail,ReceivedEmail=ReceivedEmail,ShippedEmail=ShippedEmail,CancelledEmail=CancelledEmail;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("DProducts","ALTER COLUMN","AddToMail","[nvarchar](max)", 0, "","0")
        
        query="UPDATE DProducts SET AddToMail=AddToMail;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("customCardOrders","ALTER COLUMN","strFormValue","[nvarchar](max)", 0, "","0")
        
        query="UPDATE customCardOrders SET strFormValue=strFormValue;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("creditCards","ALTER COLUMN","comments","[nvarchar](max)", 0, "","0")
        
        query="UPDATE creditCards SET comments=comments;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringProducts","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringValues","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringCategories","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringOptions","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","xfdetails","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringQuantity","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringPrice","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringCProducts","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringCValues","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configWishlistSessions","ALTER COLUMN","stringCCategories","[nvarchar](max)", 0, "","0")
        
        query="UPDATE configWishlistSessions SET stringProducts=stringProducts,stringValues=stringValues,stringCategories=stringCategories,stringOptions=stringOptions,xfdetails=xfdetails,stringQuantity=stringQuantity,stringPrice=stringPrice,stringCProducts=stringCProducts,stringCValues=stringCValues,stringCCategories=stringCCategories;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("configSpec_products","ALTER COLUMN","Notes","[nvarchar](max)", 0, "","0")
        
        query="UPDATE configSpec_products SET Notes=Notes;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("configSpec_Charges","ALTER COLUMN","Notes","[nvarchar](max)", 0, "","0")
        
        query="UPDATE configSpec_Charges SET Notes=Notes;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("configSessions","ALTER COLUMN","stringProducts","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringValues","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringCategories","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringOptions","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringQuantity","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringPrice","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringCProducts","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringCValues","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("configSessions","ALTER COLUMN","stringCCategories","[nvarchar](max)", 0, "","0")
        
        query="UPDATE configSessions SET stringProducts=stringProducts,stringValues=stringValues,stringCategories=stringCategories,stringOptions=stringOptions,stringQuantity=stringQuantity,stringPrice=stringPrice,stringCProducts=stringCProducts,stringCValues=stringCValues,stringCCategories=stringCCategories;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("categories","ALTER COLUMN","details","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("categories","ALTER COLUMN","SDesc","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("categories","ALTER COLUMN","LDesc","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("categories","ALTER COLUMN","pcCats_BreadCrumbs","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("categories","ALTER COLUMN","pcCats_MetaKeywords","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("categories","ALTER COLUMN","pcCats_MetaDesc","[nvarchar](max)", 0, "","0")
        
        query="UPDATE categories SET details=details,SDesc=SDesc,LDesc=LDesc,pcCats_BreadCrumbs=pcCats_BreadCrumbs,pcCats_MetaKeywords=pcCats_MetaKeywords,pcCats_MetaDesc=pcCats_MetaDesc;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("Blackout","ALTER COLUMN","Blackout_Message","[nvarchar](max)", 0, "","0")
        
        query="UPDATE Blackout SET Blackout_Message=Blackout_Message;"
        set rs=connTemp.execute(query)
        set rs=nothing
        
        call AlterTableSQL("wishList","ALTER COLUMN","pcwishList_OptionsArray","[nvarchar](max)", 0, "","0")
        
        query="UPDATE wishList SET pcwishList_OptionsArray=pcwishList_OptionsArray;"
        set rs=connTemp.execute(query)
        set rs=nothing

		'// Create table gwAmazon
        if not TableExists("gwAmazon") then
		    query="CREATE TABLE gwAmazon ("
		    query=query&"gwAMZ_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
		    query=query&"gwAMZ_SellerID [nvarchar] (250) NULL, "
		    query=query&"gwAMZ_AccessKey [nvarchar] (250) NULL, "
		    query=query&"gwAMZ_SecretKey [nvarchar] (250) NULL, "
		    query=query&"gwAMZ_ClientID [nvarchar] (250) NULL, "
		    query=query&"gwAMZ_ClientSecret [nvarchar] (250) NULL, "
		    query=query&"gwAMZ_Mode [int] NULL DEFAULT(0) ,"
		    query=query&"gwAMZ_TestMode [int] NULL DEFAULT(0)"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("gwAmazon")
		    end if
		    set rs=nothing
        end if
		
		'// Create table pcShipwireSettings
        if not TableExists("pcShipwireSettings") then
		    query="CREATE TABLE pcShipwireSettings ("
		    query=query&"pcSWS_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
		    query=query&"pcSWS_UserName [nvarchar] (250) NULL ,"
		    query=query&"pcSWS_Password [nvarchar] (250) NULL ,"
		    query=query&"pcSWS_OnOff [int] NULL DEFAULT(0) ,"
		    query=query&"pcSWS_Mode [int] NULL DEFAULT(0) ,"
		    query=query&"pcSWS_SyncDate [datetime] NULL "
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcShipwireSettings")
		    end if
		    set rs=nothing
        end if
	
        if not TableExists("pcShipwireOrders") then
		    query="CREATE TABLE pcShipwireOrders ("
		    query=query&"pcSWO_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
		    query=query&"idOrder [int] NULL DEFAULT(0) ,"
		    query=query&"pcSWO_ShipwireID [nvarchar] (250) NULL ,"
		    query=query&"pcSWO_ShipwireDetails [varchar] (8000) NULL "
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcShipwireOrders")
		    end if
		    set rs=nothing
        end if
		
        if not TableExists("pcContactPageSettings") then
		    query="CREATE TABLE pcContactPageSettings ("
		    query=query&"pcCPage_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
		    query=query&"pcCPage_PageDesc [nvarchar](max) NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcContactPageSettings")
		    end if
        end if
		
        if not TableExists("pcPrdXFields") then
		    query="CREATE TABLE pcPrdXFields ("
		    query=query&"pcPXF_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
		    query=query&"IdProduct [int] NULL DEFAULT(0) ,"
		    query=query&"IdXfield [int] NULL DEFAULT(0) ,"
		    query=query&"pcPXF_XReq [int] NULL DEFAULT(0)"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcPrdXFields")
		    end if
		    set rs=nothing
        end if

		call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_DisplayQuickView","[int]",1,"0","0")
		call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_AdminLastLogin","[datetime]",2,"1/1/2013","0")
    
		call AlterTableSQL("categories","ADD","pcCats_ProductOrder","[nvarchar] (4)", 0, "","0")

		call AlterTableSQL("Products","ADD","pcProd_Top","[nvarchar] (800)", 0, "","0")
		call AlterTableSQL("Products","ADD","pcProd_TopLeft","[nvarchar] (800)", 0, "","0")
		call AlterTableSQL("Products","ADD","pcProd_TopRight","[nvarchar] (800)", 0, "","0")
		call AlterTableSQL("Products","ADD","pcProd_Middle","[nvarchar] (800)", 0, "","0")
		call AlterTableSQL("Products","ADD","pcProd_Bottom","[nvarchar] (800)", 0, "","0")
		call AlterTableSQL("Products","ADD","pcProd_Tabs","[nvarchar] (max)", 0, "","0")

		call AlterTableSQL("paypal","ADD","PP_PaymentAction","[int]", 1, "1","0")

		call AlterTableSQL("pcPay_PFL_Authorize","ADD","gwCode","[int]", 1, "1","0")
        call AlterTableSQL("pcPay_PFL_Authorize","ADD","fraudcode","[int]", 1, "1","0")
		
		'// Default Product Layout
        if not TableExists("pcDefaultPrdLayout") then
		    query="CREATE TABLE pcDefaultPrdLayout ("
		    query=query&"pcDPL_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
		    query=query&"pcDPL_idProduct [INT] NULL ,"
		    query=query&"pcDPL_Name [nvarchar] (255) NULL,"
		    query=query&"pcDPL_Top [nvarchar] (800) NULL,"
		    query=query&"pcDPL_TopLeft [nvarchar] (800) NULL,"
		    query=query&"pcDPL_TopRight [nvarchar] (800) NULL,"
		    query=query&"pcDPL_Middle [nvarchar] (800) NULL,"
		    query=query&"pcDPL_Bottom [nvarchar] (800) NULL,"
		    query=query&"pcDPL_Tabs [nvarchar] (max) NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcDefaultPrdLayout")

			    '// Add new column if the table already exists
			    call AlterTableSQL("pcDefaultPrdLayout","ADD","pcDPL_Middle","[nvarchar] (800)", 0, "","0")
			    call AlterTableSQL("pcDefaultPrdLayout","ADD","pcDPL_Name","[nvarchar] (255)", 0, "","0")
			    call AlterTableSQL("pcDefaultPrdLayout","ADD","pcDPL_idProduct","[int]", 0, "","0")
		    end if
        end if
		
		'// Slideshow Feature
        if not TableExists("pcSlideShow") then
		    query="CREATE TABLE pcSlideShow ("
		    query=query&"idSlide [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		    query=query&"slideImage [nvarchar](255) NOT NULL,"
		    query=query&"slideCaption [nvarchar](MAX) NULL,"
		    query=query&"slideUrl [nvarchar](500) NULL,"
		    query=query&"slideAlt [nvarchar](255) NULL,"
		    query=query&"slideOrder [int] NULL,"
		    query=query&"slideDateUploaded datetime NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcSlideShow")
		    else
		    end if
		    set rs=nothing
        end if

		call AlterTableSQL("pcSlideShow","ADD","idSetting","[int]",0,"","0")

		'// Slideshow Settings
        if not TableExists("pcSlideShowSettings") then
		    query="CREATE TABLE pcSlideShowSettings ("
		    query=query&"id [int] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		    query=query&"slideWidth [int] NOT NULL,"
		    query=query&"slideHeight [int] NOT NULL,"
		    query=query&"effect [nvarchar](50) NULL,"
		    query=query&"pauseTime [int] NULL,"
		    query=query&"animSpeed [int] NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcSlideShowSettings")
		    else
		    end if
		    set rs=nothing
        end if

		call AlterTableSQL("pcSlideShowSettings","ADD","idSetting","[int]",0,"","0")
		call AlterTableSQL("pcSlideShowSettings","ADD","useDefault","[int]",1,"0","0")

		'// Add default slideshow configuration
		query="SELECT COUNT(*) AS count FROM pcSlideShowSettings;"
		set rs=conntemp.execute(query)
		if not rs.eof then
			if rs("count") = 0 then
				query="INSERT INTO pcSlideShowSettings (idSetting, slideWidth, slideHeight, effect, pauseTime, animSpeed) VALUES (1, 1280, 458, 'random', 5000, 100);"
				conntemp.execute(query)
			elseif rs("count") = 1 then
				query="UPDATE pcSlideShowSettings SET idSetting = 1;"
				conntemp.execute(query)
			end if
			if err.number <> 0 then
				err.number=0
				Err.Description=""
			end if
		end if
		set rs=nothing
		
		'// Add slideshow config for the mobile
		query="SELECT COUNT(*) AS count FROM pcSlideShowSettings WHERE idSetting = 2;"
		set rs=conntemp.execute(query)
		if not rs.eof then
			if rs("count") = 0 then
				query="INSERT INTO pcSlideShowSettings (idSetting, slideWidth, slideHeight, effect, pauseTime, animSpeed) VALUES (2, 1024, 600, 'random', 5000, 250);"
				conntemp.execute(query)
				if err.number <> 0 then
					err.number=0
					Err.Description=""
				end if
			end if
		end if
		set rs=nothing
        
		'// Accepted Payments
        if not TableExists("pcAcceptedPayments") then
		    query="CREATE TABLE pcAcceptedPayments ("
		    query=query&"pcAcceptedPayment_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		    query=query&"pcAcceptedPayment_Name [nvarchar](50) NOT NULL,"
		    query=query&"pcAcceptedPayment_Image [nvarchar](200) NOT NULL,"
		    query=query&"pcAcceptedPayment_CustomImage [nvarchar](200) NULL,"
		    query=query&"pcAcceptedPayment_Alt [nvarchar](255) NULL,"
		    query=query&"pcAcceptedPayment_Active [bit] NULL,"
		    query=query&"pcAcceptedPayment_Order [int] NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcAcceptedPayments")
		    end if
		    set rs=nothing
        end if


		'// Add default accepted payments
		query="SELECT COUNT(*) AS count FROM pcAcceptedPayments;"
		set rs=conntemp.execute(query)
		if not rs.eof then
			if rs("count") = 0 then
				query=""
				query=query&"INSERT INTO pcAcceptedPayments (pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_Active, pcAcceptedPayment_Order) VALUES ('PayPal', 'paypal.png', 0, 1);"
				query=query&"INSERT INTO pcAcceptedPayments (pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_Active, pcAcceptedPayment_Order) VALUES ('Amazon', 'amazon.png', 0, 2);"
				query=query&"INSERT INTO pcAcceptedPayments (pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_Active, pcAcceptedPayment_Order) VALUES ('American Express', 'amex.png', 0, 3);"
				query=query&"INSERT INTO pcAcceptedPayments (pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_Active, pcAcceptedPayment_Order) VALUES ('Discover', 'discover.png', 0, 4);"
				query=query&"INSERT INTO pcAcceptedPayments (pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_Active, pcAcceptedPayment_Order) VALUES ('MasterCard', 'mastercard.png', 0, 5);"
				query=query&"INSERT INTO pcAcceptedPayments (pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_Active, pcAcceptedPayment_Order) VALUES ('Visa', 'visa.png', 0, 6);"
				query=query&"INSERT INTO pcAcceptedPayments (pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_Active, pcAcceptedPayment_Order) VALUES ('Cirrus', 'cirrus.png', 0, 7);"
				conntemp.execute(query)
				if err.number <> 0 then
					err.number=0
					Err.Description=""
				end if
			end if
		end if
		set rs=nothing
		
		'// Google Trusted Store
        if not TableExists("pcGoogleTS") then
		    query="CREATE TABLE pcGoogleTS ("
		    query=query&"pcGTS_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
			query=query&"pcGTS_TurnOn [INT] NULL DEFAULT(0),"
			query=query&"pcGTS_AccNo [nvarchar](50) NULL,"
		    query=query&"pcGTS_PageLang [nvarchar](50) NULL,"
			query=query&"pcGTS_ShopAccID [nvarchar](50) NULL,"
			query=query&"pcGTS_ShopCountry [nvarchar](50) NULL,"
			query=query&"pcGTS_ShopLang [nvarchar](50) NULL,"
			query=query&"pcGTS_Currency [nvarchar](5) NULL,"
			query=query&"pcGTS_ShipDays [INT] NULL DEFAULT(0)"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcGoogleTS")
		    end if
		    set rs=nothing
        end if
        
		'// Social Links
        if not TableExists("pcSocialLinks") then
		    query="CREATE TABLE pcSocialLinks ("
		    query=query&"pcSocialLink_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		    query=query&"pcSocialLink_Name [nvarchar](50) NOT NULL,"
		    query=query&"pcSocialLink_Image [nvarchar](200) NOT NULL,"
		    query=query&"pcSocialLink_CustomImage [nvarchar](200) NULL,"
		    query=query&"pcSocialLink_Url [nvarchar](500) NULL,"
		    query=query&"pcSocialLink_Alt [nvarchar](255) NULL,"
		    query=query&"pcSocialLink_Order [int] NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcSocialLinks")
		    end if
		    set rs=nothing
        end if


		'// Add default social links
		query="SELECT COUNT(*) AS count FROM pcSocialLinks;"
		set rs=conntemp.execute(query)
		if not rs.eof then
			if rs("count") = 0 then
				query=""
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Facebook', 'facebook.png', 'Like us on Facebook!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Google+', 'googleplus.png', 'Add us to your circles on Google+!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Instagram', 'instagram.png', 'Follow us on Instagram!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Twitter', 'twitter.png', 'Follow us on Twitter!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('YouTube', 'youtube.png', 'Follow us on YouTube!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('LinkedIn', 'linkedin.png', 'Connect to us on LinkedIn!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Pinterest', 'pinterest.png', 'Visit our Pinterest page!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Blogger', 'blogger.png', 'Visit our Blog!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Tumblr', 'tumblr.png', 'Subscribe to us on Tumblr!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('Wordpress', 'wordpress.png', 'Visit our Wordpress page!');"
				query=query&"INSERT INTO pcSocialLinks (pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_Alt) VALUES ('RSS', 'rss.png', 'Subscribe to our RSS feed!');"
				conntemp.execute(query)
				if err.number <> 0 then
					err.number=0
					Err.Description=""
				end if
			end if
		end if
		set rs=nothing
		
		'// Create table pcFacebookSettings
        if not TableExists("pcFacebookSettings") then
		    query="CREATE TABLE [pcFacebookSettings] ("
		    query=query&"[pcFBS_id] [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
		    query=query&"[pcFBS_TurnOnOff] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_OffMsg] [nvarchar] (400) NULL ,"
		    query=query&"[pcFBS_AppID] [nvarchar] (100) NULL ,"
		    query=query&"[pcFBS_RedirectURL] [nvarchar] (250) NULL ,"
		    query=query&"[pcFBS_Header] [nvarchar] (max) NULL ,"
		    query=query&"[pcFBS_Footer] [nvarchar] (max) NULL ,"
		    query=query&"[pcFBS_PageWidth] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_CustomDisplay] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_CatImages] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_CatRow] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_CatRowsPerPage] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_BType] [nvarchar] (5) NULL ,"
		    query=query&"[pcFBS_PrdRow] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_PrdRowsPerPage] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_ShowSKU] [int] NULL DEFAULT(0) ,"
		    query=query&"[pcFBS_ShowSmallImg] [int] NULL DEFAULT(0)"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcFacebookSettings")
		    end if
		    set rs=nothing
        end if

		call AlterTableSQL("Customers","ADD","pcCust_FBId","[nvarchar] (100)", 0, "","0")
		call AlterTableSQL("Customers","ADD","pcCust_AmazonId","[nvarchar] (200)", 0, "","0")
		
		call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_PNButtons","[int]",1,"1","0")
		call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_ConURL","[int]",1,"0","0")
		call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_GAType","[int]",1,"0","0")
		call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_ThemeFolder","[nvarchar] (100)",0,"","0")

		'// Add shipping service shipment ID
		call AlterTableSQL("shipService","ADD","idShipment","[int]",0,"","1")
        
		'// Create table pcPackageLabel
        if not TableExists("pcPackageLabel") then
		    query="CREATE TABLE pcPackageLabel ("
		    query=query&"pcPackageLabel_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL,"
		    query=query&"pcPackageInfo_ID [INT] NOT NULL ,"
		    query=query&"pcPackageLabel_Name [nvarchar] (100) NOT NULL ,"
		    query=query&"pcPackageLabel_File [nvarchar] (255) NOT NULL ,"
		    query=query&"pcPackageLabel_FileType [nvarchar] (50) NULL ,"
		    query=query&"pcPackageLabel_Resolution [INT] NULL ,"
		    query=query&"pcPackageLabel_Type [nvarchar] (50) NULL ,"
		    query=query&"pcPackageLabel_Date [datetime] NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcPackageLabel")
		    end if
		    set rs=nothing
        end if

		'// UPS
		query = "UPDATE shipService SET idShipment = 3 WHERE serviceCode IN ('01','02','03','07','08','11','12','13','14','54','59','65')"
		set rs=conntemp.execute(query)
		if err.number<>0 then
			err.number=0
			Err.Description=""
		end if
		set rs=nothing

		'// USPS
		query = "UPDATE shipService SET idShipment = 4 WHERE serviceCode IN ('9901','9902','9903','9904','9905','9906','9907','9908','9909','9910','9911','9912','9914','9915','9916','9917')"
		set rs=conntemp.execute(query)
		if err.number<>0 then
			err.number=0
			Err.Description=""
		end if
		set rs=nothing
		
		'// Change "USPS Parcel" to "USPS Standard Post"
		query="UPDATE shipService SET serviceDescription = 'USPS Standard Post<sup>&reg;</sup>' WHERE serviceCode = '9903';"
		set rs=conntemp.execute(query)
		if err.number<>0 then
			err.number=0
			Err.Description=""
		end if
		set rs=nothing

		'// Canada Post
		query = "UPDATE shipService SET idShipment = 7 WHERE serviceCode IN ('1010','1020','1130','1030','1040','1120','1220','1230','2010','2020','2030','2040','2050','3010','3020','3040','2005','2015','2025','3005','3015','3025','3050')"
		set rs=conntemp.execute(query)
		if err.number<>0 then
			err.number=0
			Err.Description=""
		end if
		set rs=nothing


		'// Update icon resources
		call UpdateTableIfValue("pcRevSettings", "pcRS_Img1", "", "smileygreen.gif", "smileygreen.png")
		call UpdateTableIfValue("pcRevSettings", "pcRS_Img2", "", "smileyred.gif", "smileyred.png")
		call UpdateTableIfValue("pcRevSettings", "pcRS_Img3", "", "fullstar.gif", "fullstar.png")
		call UpdateTableIfValue("pcRevSettings", "pcRS_Img4", "", "halfstar.gif", "halfstar.png")
		call UpdateTableIfValue("pcRevSettings", "pcRS_Img5", "", "emptystar.gif", "emptystar.png")

		call UpdateTableIfValue("icons", "discount", "where id=1", "images/sample/pc_icon_discount.gif", "images/sample/pc_icon_discount.png")
		call UpdateTableIfValue("icons", "erroricon", "where id=1", "images/sample/pc_icon_error.gif", "images/sample/pc_icon_error.png")
		call UpdateTableIfValue("icons", "zoom", "where id=1", "images/sample/pc_icon_zoom.gif", "images/sample/pc_icon_zoom.png")

		'// Create table pcUpdateLog
        if not TableExists("pcUpdateLog") then
		    query="CREATE TABLE pcUpdateLog ("
		    query=query&"id [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL,"
		    query=query&"name [nvarchar] (1000) NULL ,"
		    query=query&"filename [nvarchar] (1000) NULL ,"
		    query=query&"date_installed [datetime] NULL DEFAULT (GETDATE()) ,"
		    query=query&"notes [nvarchar] (max) NULL"
		    query=query&");"
		    set rs=server.CreateObject("ADODB.RecordSet")
		    set rs=conntemp.execute(query)
		    if err.number <> 0 then
			    TrapSQLError("pcUpdateLog")
		    end if
		    set rs=nothing
	    end if

	'========================================================================
	'// END OF DB UPDATES FOR v5.01
	'========================================================================

	'========================================================================
	'// START OF DB UPDATES FOR v5.02
	'========================================================================

    call AlterTableSQL("emailSettings","ADD","FontSize","[nvarchar](10)", 2, "13px","0")

	'========================================================================
	'// END OF DB UPDATES FOR v5.02
	'========================================================================
	set rs=nothing
	%>
		<!-- #include file="pcAdminRetrieveSettings.asp" -->
	<%
	pcIntScUpgrade = 0
	%>
		<!-- #include file="pcAdminSaveSettings.asp" -->
	<%

	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if

END IF
%>
<!--#include file="AdminHeader.asp"-->
<form action="upddb_v5.02.asp" method="post" name="form1" id="form1" class="pcForms">
<%
if mode="complete" then
	call closeDb()
	response.redirect "upddb_v5.02_complete.asp?CanUpd=" & CanUpd
	response.end()	
else 
%>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

		<% if mode="errors" then %>
			<tr>
				<td align="center">
					<div class="pcCPmessage">The following errors occurred while updating the store database. Try running the database update script again. If the errors persist, please open a support ticket:
                        <br><br>
					    <%=ErrStr%>
                    </div>
				</td>
			</tr>
		<% end if %>
		<% if request("s")="88" then %>
			<tr>
				<td align="center">
					<div class="pcCPmessageSuccess">Updated SQL database successfully to use the data type: 'Nvarchar(Max)' instead of 'NText'</div>
				</td>
			</tr>
		<% end if %>
		<%IF scFixedNText = 0 then%>
			<tr>
				<td align="center">
					<p><strong>From ProductCart v5.0, we don't use the field data type: 'NText' anymore for store database because the next versisons of MS SQL Server won't support it.<br>
					You need to update store database to use the data type: 'Nvarchar(Max)' instead of 'NText'.</p>
					<br><br>
					<input name="fixntext" type="button" class="btn btn-default"  id="fixntext" value="Update Your ProductCart MS SQL Database" class="btn btn-primary" onclick="javascript:location='upddb_fixNtext.asp';">
					<br><br>					
				</td>
			</tr>
		<%ELSE%>
		<tr>
			<td>
            
                <h1 class="page-header">Welcome to ProductCart 5.02</h1>
                <p class="lead">
                    ProductCart 5.02 is a cumulative patch containing miscellaneous bug fixes and improvements for ProductCart v5.0. 
                    Be sure to read the <a href="https://productcart.desk.com/customer/portal/articles/1604177-updating-productcart-v5-x-to-v5-02" target="_blank">v5.02 Update Guide</a>.
                </p>

				
					<% 
                    dim findit
                    if PPD="1" then
                        PageName="/"&scPcFolder&"/includes/diagtxt.txt"
                    else
                        PageName="../includes/diagtxt.txt"
                    end if
                    findit=Server.MapPath(PageName)
                    
                    Dim fso, f, errpermissions, errdelete_includes, errwrite_includes, errwrite_others
                    errpermissions=0
                    errdelete_includes=0
                    errwrite_includes=0
                    errwrite_others=0
                    Set fso=server.CreateObject("Scripting.FileSystemObject")
                    Set f=fso.GetFile(findit)
                    Err.number=0
                    f.Delete
                    if Err.number>0 then
                        errdelete_includes=1
                        errpermissions=1
                        Err.number=0
                    end if
                    'Set f=nothing
                                
                    Set f=fso.OpenTextFile(findit, 2, True)
                    f.Write "test done"
                    if Err.number>0 then
                        errwrite_includes=1
                        errpermissions=1
                        Err.number=0
                    end if
                    
                    if PPD="1" then
                        PageName="/"&scPcFolder&"/pc/diagtxt.txt"
                    else			
                        PageName="../pc/diagtxt.txt"
                    end if
                    findit=Server.MapPath(PageName)
                    Set f=fso.OpenTextFile(findit, 2, True)
                    f.Write "test done"
                    if Err.number>0 then
                        errwrite_others=1
                        errpermissions=1
                        Err.number=0
                    end if
                                
                    f.Close
                    Set fso=nothing
                    Set f=nothing
                    if errpermissions=0 then %>
 
					<% else %>
                    
                        <div class="pcCPmessageWarning">
                        
                        <h2>Please correct these issues before you begin:</h2>

                        <% if scDB<>"SQL" then %> 
                            <table>
                                <tr> 
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">ProductCart v5 only works with MS SQL databases.  The Access database is been deprecated for security and performance reasons.  <a href="https://www.productcart.com/support/v5/article.asp?id=3" target="_blank">Click here</a> to ask for a quote to convert your Access database to SQL.</font></td>
                                </tr>
                            </table>
                        <% end if %>
                        
					    <% if errwrite_others=1 or errwrite_includes=1 then %> 
                            <table>
                                <tr> 
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">You need to assign 'read/write' permissions to the 'productcart' folder and all of its subfolders.</font></td>
                                </tr>
                            </table>
						<% end if

                            if errdelete_includes=1 then 
                                %>
                                <table>
                                <tr> 
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">You need to assign 'read/write/delete' permissions to the 'productcart/includes' folder and all of its subfolders.</font></td>
                                </tr> 
                            </table>
                            <% 
                            end if
                            %>
                            </div>
                            <%
				    end if 
                    %>
                
                    <div class="bs-callout bs-callout-info">
                        <h4>Read Me</h4>
                        <p>
                            Click "Update Now" to update your MS SQL Database to v5.02.  
                        </p>
                    </div>              
    
                    <div class="bs-callout bs-callout-warning">
                        <h4>Backup Your Database</h4>
                        <p>
                            Although we have tested this update script in a variety of environments, there is always the possibility of something going wrong. 
                            Make sure to <span style="font-weight: bold">backup your database</span> prior to executing this update.
                            Depending on how the database has been setup, you may be able to either perform the backup yourself or have your Web hosting company do it for you. 
                            Note: Your SQL database is likely being automatically backed up every day: confirm that this is the case by asking your Web host when the last back up occurred.
                        </p>
                    </div>

			<table class="pcCPcontent" width="80%">
			<% if request.querystring("mode")="1" OR request.querystring("mode")="3" then %>
				<tr>
					<td>
						It appears that you are using a DSN connection to connect to your SQL server. In order to complete this update, please enter your SQL Server Information below:
						<% if request.querystring("mode")="1" then %>
							<br>
							<strong>*All fields are required.</strong>
						<% end if %>

						<input name="hmode" type="hidden" id="hmode" value="2">	
					</td>
				</tr>
				<tr>
					<td>Server Domain/IP: <input name="SSIP" type="text" id="SSIP" size="30"></td>
				</tr>
				<tr>
					<td>Database Name: <input name="SSDB" type="text" id="SSDB" size="30"></td>
				</tr>
				<tr>
					<td>User ID: <input name="UID" type="text" id="UID" size="30"></td>
				</tr>
				<tr>
					<td>Password: <input name="PWD" type="password" id="PWD" size="30"></td>
				</tr>

			<% end if %>
				<tr>
					<td align="center">
						<input name="action" type="hidden" id="action" value="sql">

                        <% if errpermissions=0 then %>
                                <input type="button" name="access2" value=" Update Now " onClick="$pc('#form1').submit();" class="btn btn-primary">
                        <% else %>
                                <input type="button" name="access2" value=" Update Now " class="btn btn-primary disabled" disabled>
                        <% end if %>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<%END IF%>
	</table>
<% end if %>
</form>
<!--#include file="AdminFooter.asp"-->
