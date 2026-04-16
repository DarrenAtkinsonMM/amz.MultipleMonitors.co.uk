<%@LANGUAGE="VBSCRIPT"%>
<% 
On Error Resume Next
PmAdmin=19
pageTitle = "ProductCart v5.2.10 - Database Update Completed"
Section = "" 
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<% dim f, iCnt %>
<!--#include file="AdminHeader.asp"-->
<!--#include file="pcAdminRetrieveSettings.asp"-->
<%
'// Is this an update, or upgrade?
pcv_boolIsUpdate = false
If instr(scVersion, "5.2.00")>0 Or instr(scVersion, "5.2.10")>0 Then
    pcv_boolIsUpdate = true
End If

if NOT isNULL(pcStrCompanyName) then pcStrCompanyName=replace(pcStrCompanyName,"'","''")
if NOT isNULL(pcStrCompanyAddress) then pcStrCompanyAddress=replace(pcStrCompanyAddress,"'","''")
if NOT isNULL(pcStrCompanyZip) then pcStrCompanyZip=replace(pcStrCompanyZip,"'","''")
if NOT isNULL(pcStrCompanyCity) then pcStrCompanyCity=replace(pcStrCompanyCity,"'","''")
if NOT isNULL(pcStrCompanyState) then pcStrCompanyState=replace(pcStrCompanyState,"'","''")
if NOT isNULL(pcStrCompanyCountry) then pcStrCompanyCountry=replace(pcStrCompanyCountry,"'","''")
if NOT isNULL(pcStrCompanyLogo) then pcStrCompanyLogo=replace(pcStrCompanyLogo,"'","''")
if NOT isNULL(pcStrCurSign) then pcStrCurSign=replace(pcStrCurSign,"'","''")
if NOT isNULL(pcStrDecSign) then pcStrDecSign=replace(pcStrDecSign,"'","''")
if NOT isNULL(pcStrDivSign) then pcStrDivSign=replace(pcStrDivSign,"'","''")
if NOT isNULL(pcStrDateFrmt) then pcStrDateFrmt=replace(pcStrDateFrmt,"'","''")
if NOT isNULL(pcStrURLredirect) then pcStrURLredirect=replace(pcStrURLredirect,"'","''")
if NOT isNULL(pcStrSSL) then pcStrSSL=replace(pcStrSSL,"'","''")
if NOT isNULL(pcStrSSLUrl) then pcStrSSLUrl=replace(pcStrSSLUrl,"'","''")
if NOT isNULL(pcStrIntSSLPage) then pcStrIntSSLPage=replace(pcStrIntSSLPage,"'","''")
if NOT isNULL(pcStrBType) then pcStrBType=replace(pcStrBType,"'","''")
if NOT isNULL(pcStrStoreOff) then pcStrStoreOff=replace(pcStrStoreOff,"'","''")
if NOT isNULL(pcStrStoreMsg) then pcStrStoreMsg=replace(pcStrStoreMsg,"'","''")
if NOT isNULL(pcStrorderLevel) then pcStrorderLevel=replace(pcStrorderLevel,"'","''")
if NOT isNULL(pcStrNewsLabel) then pcStrNewsLabel=replace(pcStrNewsLabel,"'","''")
if NOT isNULL(pcStrDFLabel) then pcStrDFLabel=replace(pcStrDFLabel,"'","''")
if NOT isNULL(pcStrDFShow) then pcStrDFShow=replace(pcStrDFShow,"'","''")
if NOT isNULL(pcStrDFReq) then pcStrDFReq=replace(pcStrDFReq,"'","''")
if NOT isNULL(pcStrTFLabel) then pcStrTFLabel=replace(pcStrTFLabel,"'","''")
if NOT isNULL(pcStrTFShow) then pcStrTFShow=replace(pcStrTFShow,"'","''")
if NOT isNULL(pcStrTFReq) then pcStrTFReq=replace(pcStrTFReq,"'","''")
if NOT isNULL(pcStrDTCheck) then pcStrDTCheck=replace(pcStrDTCheck,"'","''")
if NOT isNULL(pcStrDeliveryZip) then pcStrDeliveryZip=replace(pcStrDeliveryZip,"'","''")
if NOT isNULL(pcStrOrderName) then pcStrOrderName=replace(pcStrOrderName,"'","''")
if NOT isNULL(pcStrHideDiscField) then pcStrHideDiscField=replace(pcStrHideDiscField,"'","''")
if NOT isNULL(pcStrAllowSeparate) then pcStrAllowSeparate=replace(pcStrAllowSeparate,"'","''")
if NOT isNULL(pcStrReferLabel) then pcStrReferLabel=replace(pcStrReferLabel,"'","''")
if NOT isNULL(pcStrRewardsLabel) then pcStrRewardsLabel=replace(pcStrRewardsLabel,"'","''")
if NOT isNULL(pcStrXML) then pcStrXML=replace(pcStrXML,"'","''")
if NOT isNULL(pcStrBTODetTxt) then pcStrBTODetTxt=replace(pcStrBTODetTxt,"'","''")
if NOT isNULL(pcStrTermsLabel) then pcStrTermsLabel=replace(pcStrTermsLabel,"'","''")
if NOT isNULL(pcStrTermsCopy) then pcStrTermsCopy=replace(pcStrTermsCopy,"'","''")
if NOT isNULL(pcStrViewPrdStyle) then pcStrViewPrdStyle=replace(pcStrViewPrdStyle,"'","''")
if NOT isNULL(pcStrCustomerIPAlert) then pcStrCustomerIPAlert=replace(pcStrCustomerIPAlert,"'","''")
if NOT isNULL(pcStrCompanyPhoneNumber) then pcStrCompanyPhoneNumber=replace(pcStrCompanyPhoneNumber,"'","''")
if NOT isNULL(pcStrCompanyFaxNumber) then pcStrCompanyFaxNumber=replace(pcStrCompanyFaxNumber,"'","''")
if NOT isNULL(pcStrseoURLs404) then pcStrseoURLs404=replace(pcStrseoURLs404,"'","''")

'// New v5 Fields
if len(pcIntShowSKU)<1 then
	pcIntShowSKU=0
end if
if len(pcIntShowSmallImg)<1 then
	pcIntShowSmallImg=-1
end if
if len(pcStrViewPrdStyle)<1 then
	pcStrViewPrdStyle="c"
end if
pcIntErrorHandler=1
pcIntEnableBundling=0
pcIntKeepSession=0

'/////////////////////////////////////////////////////
'// Update version number
'/////////////////////////////////////////////////////

'//Version 3 & UP only - change for any new version updates
if scBTO=1 then
	pcStrScVersion="5.2.10b"
else
	pcStrScVersion="5.2.10"
end if

'// Subversion
pcStrScSubVersion = ""

'//Service Pack Number
pcStrScSP = "0"
%>
<!--#include file="pcAdminSaveSettings.asp"-->
<% 'Detect Add-on %>
<!--#include file="pcAddOnDetection.asp"-->
<style>
li {
	padding-bottom: 8px;
}
h2 {
	font-size: 12px;
}
</style>
<table class="pcCPcontent">
    <tr>
		<td>
            <div class="pcCPmessageSuccess">
                Your store database was successfully updated.                
            </div> 
            
            <% If pcv_boolIsUpdate <> true Then %>
            
                <div class="bs-callout bs-callout-info upgrade">
                    <h4>Post-Update Review</h4>
                    <strong>You're almost done!</strong> Please take a moment to go over the following post-update action items.
                    <ul class="list-group">
                        <li class="list-group-item ">                        
                            <h3><span class="glyphicon glyphicon-check checklist"></span>Tax Settings</h3>
                            ProductCart v5.2 includes an integration with Avalara Tax services. Due to the new features we recommend that you <a href="AdminTaxSettings.asp" target="_blank">review and resave your tax settings</a>.
                            
                            <div style="float: right; position: absolute; right: 15px; top: 15px;">
                                <a href="AdminTaxSettings.asp" target="_blank" class="btn btn-default btn-xs">Review</a>
                            </div>
                        </li>                    
                        <li class="list-group-item">                        
                            <h3><span class="glyphicon glyphicon-check checklist"></span>Theme Settings</h3>
                            ProductCart v5.2 includes an improved theme editor.  Due to the new features we recommend that you <a href="ThemeSettings.asp" target="_blank">review and resave your theme settings</a>.  Note that we will create preview snapshots for each theme on the first page load.  Please allow the page 30 seconds to a minute to generate the previews.
                            <div style="float: right; position: absolute; right: 15px; top: 15px;">
                                <a href="ThemeSettings.asp" target="_blank" class="btn btn-default btn-xs">Review</a>
                            </div>
                        </li>
                        <li class="list-group-item">
                            <h3><span class="glyphicon glyphicon-check checklist"></span>Store Settings</h3>
                            There are several new store settings, such as the following:
                            <ul>
                                <li><a href="https://productcart.desk.com/customer/portal/articles/2492905-google-tag-manager" target="_blank">Google Tag Manager</a></li>
                                <li><a href="https://productcart.desk.com/customer/portal/articles/2492927-combine-minify-css" target="_blank">Combine & Minify CSS / JavaScript</a></li>
                            </ul>
                            Due to the new settings we recommend that you <a href="AdminSettings.asp" target="_blank">review and resave your Store Settings</a>.
                            <div style="float: right; position: absolute; right: 15px; top: 15px;">
                                <a href="AdminSettings.asp" target="_blank" class="btn btn-default btn-xs">Review</a>
                            </div>
                        </li>
                        <li class="list-group-item">
                            <h3><span class="glyphicon glyphicon-check checklist"></span>Security Settings</h3>
                            There are several new security settings, such as the following:
                            <ul>
                                <li><a href="https://productcart.desk.com/customer/portal/articles/2324205-advanced-security-settings" target="_blank">Advanced Gateway Security</a></li>
                                <li><a href="https://productcart.desk.com/customer/portal/articles/2288120-google-recaptcha" target="_blank">Google reCapatcha</a></li>
                            </ul>
                            Due to the new settings we recommend that you <a href="AdminSecuritySettings.asp" target="_blank">review and resave your Store Settings</a>.
                            <div style="float: right; position: absolute; right: 15px; top: 15px;">
                                <a href="AdminSecuritySettings.asp" target="_blank" class="btn btn-default btn-xs">Review</a>
                            </div>
                        </li>
                        <li class="list-group-item">
                            <h3><span class="glyphicon glyphicon-check checklist"></span>ProductCart Apps</h3>
                            Create a FREE <a href="pcws_MyAccount.asp" target="_blank">ProductCart Apps</a> account and add the power of the cloud to your ProductCart store.
                            <div style="float: right; position: absolute; right: 15px; top: 15px;">
                                <a href="pcws_MyAccount.asp" target="_blank" class="btn btn-default btn-xs">Review</a>
                            </div>
                        </li>
                    </ul>
                </div> 
            
            <% End If %> 
              
            <div style="padding: 5px; text-align: center">
                <input type="button" name="continue" value=" Main Menu " onClick="location='menu.asp'" class="btn btn-primary">
            </div>      
        </td>
    </tr>
    <tr>
  	    <td class="pcCPspacer"></td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->