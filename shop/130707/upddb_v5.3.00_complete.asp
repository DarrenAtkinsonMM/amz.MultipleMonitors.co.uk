<%@LANGUAGE="VBSCRIPT"%>
<% 
On Error Resume Next
PmAdmin=19
pageTitle = "ProductCart v5.3.00 - Database Update Completed"
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
	pcStrScVersion="5.3.00b"
else
	pcStrScVersion="5.3.00"
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
            
            <div class="bs-callout bs-callout-info upgrade">
                <h4>Post-Update Review</h4>
                <strong>You're almost done!</strong> Please take a moment to go over the following post-update action items.
                <ul class="list-group">
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Home Page Slideshow Scheduler</h3>
                        ProductCart v5.3 adds the ability to provide Start dates and End Dates to your home page slider images.  Edit your slider images as usual, and you'll see this new feature.
                        
                    </li>   
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Change Slider Engine to Use Flickety</h3>
                        ProductCart v5.3 updates the slider engine to Flickety.  The Flickety engine is much more robust and works great with the newest version of JQuery.
                        
                    </li>					
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Banner Manager</h3>
                        ProductCart v5.3 adds the ability to create a banner and display it on any page(s) you choose.  For complete instructions on how to use and implement this new feature, please see: <a href="https://productcart.desk.com/customer/portal/articles/2959590-banner-manager---v5-3-" target="_blank">Banner Manager</a>
                        
                    </li>   
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Cross Selling Widgets</h3>
                        ProductCart v5.3 provides a couple of widgets that will aid in your SEO and increase your cross selling. It's located at the top of Control Panel > Marketing > Manage Cross Selling > Cross Selling Settings.
						This quick widget will allow you to benefit from cross selling until you can set up full cross selling features.  Even once that's done, it will still be beneficial!
                    </li> 
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Updated JQuery Library</h3>
                        ProductCart v5.3 updates the version of JQuery to 3.3.1.
                    </li>                     
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Quick Order Entry</h3>
                        ProductCart v5.3 adds a slideout form for quick bulk entries based on SKU.  Turn this setting on by going to Control Panel > Settings > Store & Display Settings > Miscellaneous Tab - "Enable Bulk Add on Category Pages".
                    </li>
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Offline Credit Card Feature Removed</h3>
                        ProductCart v5.3 removes the offline credit card feature to help you better meet PCI requirements.  While the ability to create/edit/use the offline feature
						has been removed, existing stored credit cards have not been removed.  We recommend that you go to Control Panel > Orders > Purge Credit Cart Numbers and 
						delete any that may be stored.
                    </li>					
                    <li class="list-group-item ">                        
                        <h3><span class="glyphicon glyphicon-check checklist"></span>Configurator - Multiple Defaults</h3>
                        ProductCart v5.3 adds a feature in our Configurator version where you now can set multiple defaults in a configuration.
                    </li>					
                </ul>
            </div>  
              
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