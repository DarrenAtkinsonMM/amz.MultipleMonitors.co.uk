<%@LANGUAGE="VBSCRIPT"%>
<%
Server.ScriptTimeout = 5400
On Error Resume Next
PmAdmin=19
pageTitle = "ProductCart v5.0 - Switch NText to NVarChar(Max)" 
Section = "" 
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="fixedNTextConst.asp"-->
<%
if scFixedNText=1 then
	call closeDb()
    response.redirect "upddb_v50.asp"
	response.End()
end if

	iCnt=0
	ErrStr=""
	
	'========================================================================
	'// SQL DB Update	
	'========================================================================
	
	call AlterTableSQL("ProductsOrdered","ALTER COLUMN","xfdetails","[nvarchar](max)", 0, "","0")
	call AlterTableSQL("ProductsOrdered","ALTER COLUMN","pcPrdOrd_SelectedOptions","[nvarchar](max)", 0, "","0")
	call AlterTableSQL("ProductsOrdered","ALTER COLUMN","pcPrdOrd_OptionsPriceArray","[nvarchar](max)", 0, "","0")
	call AlterTableSQL("ProductsOrdered","ALTER COLUMN","pcPrdOrd_OptionsArray","[nvarchar](max)", 0, "","0")
	
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
	
	'call AlterTableSQL("pcBestSellerSettings","ALTER COLUMN","pcBSS_PageDesc","[nvarchar](max)", 0, "","0")
	
	'query="UPDATE pcBestSellerSettings SET ppcBSS_PageDesc=pcBSS_PageDesc;"
	'set rs=connTemp.execute(query)
	'set rs=nothing
	
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

	'========================================================================
	'// END OF DB UPDATES
	'========================================================================


	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if


%>
<!--#include file="AdminHeader.asp"-->
<%
if mode="complete" then
	SaveFile = "fixedNTextConst.asp"
	findit = Server.MapPath(Savefile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 2)
	f.WriteLine CHR(60)&CHR(37)
	f.WriteLine "private const scFixedNText = 1"
	f.WriteLine CHR(37)&CHR(62)
	f.close
	call closeDb()
    response.redirect "upddb_v5.1.00.asp?s=88"
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
					<%=ErrStr%></div>
				</td>
			</tr>
		<% end if %>
	</table>
<% end if %>
<!--#include file="AdminFooter.asp"-->
