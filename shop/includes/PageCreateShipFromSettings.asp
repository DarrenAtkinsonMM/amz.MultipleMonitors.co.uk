<!--#include file="adminv.asp"-->
<!--#include file="common.asp"-->
<!--#include file="utilities.asp"-->
<!--#include file="shipFromSettings.asp"-->
<% 
' form parameters		
pShipFromPersonName=Session("pcAdminpShipFromPersonName")
pShipFromPersonName=replace(pShipFromPersonName,"''","'")
pShipFromName=Session("pcAdminpShipFromName")
pShipFromName=replace(pShipFromName,"''","'")
pShipFromDepartment=Session("pcAdminpShipFromDepartment")
pShipFromDepartment=replace(pShipFromDepartment,"''","'")
pShipFromEmail=Session("pcAdminpShipFromEmail")
pShipFromPhone=Session("pcAdminpShipFromPhone")
pShipFromPage=Session("pcAdminpShipFromPage")
pShipFromFax=Session("pcAdminpShipFromFax")
pShipFromAddress1=Session("pcAdminpShipFromAddress1")
pShipFromAddress2=Session("pcAdminpShipFromAddress2")
pShipFromAddress3=Session("pcAdminpShipFromAddress3")
pShipFromCity=Session("pcAdminpShipFromCity")
pShipFromCity=replace(pShipFromCity,"''","'")
pShipFromPostalCode=Session("pcAdminpShipFromPostalCode")
pShipFromZip4=Session("pcAdminpShipFromZip4")
if Session("pcAdminpShipFromProvince") <> "" then
	pShipFromState=Session("pcAdminpShipFromProvince")
else
	pShipFromState=Session("pcAdminpShipFromState")
end if
pShipFromPostalCountry=Session("pcAdminpShipFromPostalCountry")
pPackageWeightLimit=Session("pcAdminpackageWeightLimit")

if NOT isNumeric(pPackageWeightLimit) then
	pPackageWeightLimit=0
end if

pDefaultProvider=Session("pcAdminDefaultProvider")

pAlwAltShipAddress=Session("pcAdminAlwAltShipAddress")
if pAlwAltShipAddress="" then
	pAlwAltShipAddress="0"
end if

If pAlwAltShipAddress=0 Then
	pHideShipAddress="1"
End If

pComResShipAddress=Session("pcAdminComResShipAddress")
if pComResShipAddress="" then
	pComResShipAddress="0"
end if

pAlwNoShipRates=Session("pcAdminAlwNoShipRates")
if pAlwNoShipRates="" then
	pAlwNoShipRates="0"
end if
pUseShipMap=Session("pcAdminUseShipMap")
if pUseShipMap="" then
	pUseShipMap="0"
end if
pShowProductWeight=Session("pcAdminpShowProductWeight")
if pShowProductWeight="" then
	pShowProductWeight="0"
end if

'Start SDBA
tShipNotifySeparate=Session("pcAdminsds_NotifySeparate")
if tShipNotifySeparate="" then
	tShipNotifySeparate="0"
end if
'End SDBA

pShowCartWeight=Session("pcAdminpShowCartWeight")
if pShowCartWeight="" then
	pShowCartWeight="0"
end if
pShowEstimateLink=Session("pcAdminpShowEstimateLink")
if pShowEstimateLink="" then
	pShowEstimateLink="0"
end if
pHideProductPackage=Session("pcAdminpHideProductPackage")
if pHideProductPackage="" then
	pHideProductPackage="0"
end if

pHideEstimateDeliveryTimes=Session("pcAdminpHideEstimateDeliveryTimes")
if pHideEstimateDeliveryTimes="" then
	pHideEstimateDeliveryTimes="0"
end if

pShipFromWeightUnit=scShipFromWeightUnit
if pShipFromWeightUnit="" then
	pShipFromWeightUnit="LBS"
end if
pSectionShow=Session("pcAdminsectionShow")
select case pSectionShow
	case "NA"
		pRatesOnly="NO"
		pShipDetailTitle=""
		pShipDetails=""
	case "TOP"
		pRatesOnly="NO"
		pShipDetailTitle=Session("pcAdminshipDetailTitle")
		pShipDetailTitle=replace(pShipDetailTitle,"''","'")
		pShipDetailTitle=replace(pShipDetailTitle,"""","&quot;")
		if pShipDetailTitle="" then
			strErr="Ship Details Title is a required field."
		end if
		pShipDetails=Session("pcAdminshipDetails")
		pShipDetails=replace(pShipDetails,vbCrlF,"<BR>")
		pShipDetails=replace(pShipDetails,"""","""""")
		if pShipDetails="" then
			if strErr<>"" then
				strErr=strErr&"<BR>"
			else
				strErr=strErr&"Ship Details is a required field."
			end if
		end if
	case "BTM"
		pRatesOnly=Session("pcAdminratesOnly")
		if pRatesOnly="YES" then
		else
			pRatesOnly="NO"
		end if
		pShipDetailTitle=Session("pcAdminshipDetailTitle")
		pShipDetailTitle=replace(pShipDetailTitle,"''","'")
		pShipDetailTitle=replace(pShipDetailTitle,"""","&quot;")
		if pShipDetailTitle="" then
			strErr="Ship Details Title is a required field."
		end if
		pShipDetails=Session("pcAdminshipDetails")
		pShipDetails=replace(pShipDetails,vbCrlF,"<BR>")
		pShipDetails=replace(pShipDetails,"""","""""")
		if pShipDetails="" then
			if strErr<>"" then
				strErr=strErr&"<BR>"
			else
				strErr=strErr&"Ship Details is a required field."
			end if
		end if
end select

If (len(pShipFromPersonName)=0) Then
	response.End() '// Not a valid form.
End If

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="shipFromSettings.asp"

set StringBuilderObj = new StringBuilder

StringBuilderObj.append CHR(60)&CHR(37)&"private const scShipFromName="&q&pShipFromName&q&CHR(10)
StringBuilderObj.append "private const scOriginPersonName="&q&pShipFromPersonName&q&CHR(10)
StringBuilderObj.append "private const scOriginDepartment="&q&pShipFromDepartment&q&CHR(10)
StringBuilderObj.append "private const scOriginEmailAddress="&q&pShipFromEmail&q&CHR(10)
StringBuilderObj.append "private const scOriginPhoneNumber="&q&pShipFromPhone&q&CHR(10)
StringBuilderObj.append "private const scOriginPagerNumber="&q&pShipFromPage&q&CHR(10)
StringBuilderObj.append "private const scOriginFaxNumber="&q&pShipFromFax&q&CHR(10)
StringBuilderObj.append "private const scShipFromAddress1="&q&pShipFromAddress1&q&CHR(10)
StringBuilderObj.append "private const scShipFromAddress2="&q&pShipFromAddress2&q&CHR(10)
StringBuilderObj.append "private const scShipFromAddress3="&q&pShipFromAddress3&q&CHR(10)
StringBuilderObj.append "private const scShipFromCity="&q&pShipFromCity&q&CHR(10)
StringBuilderObj.append "private const scShipFromState="&q&pShipFromState&q&CHR(10)
StringBuilderObj.append "private const scShipFromPostalCode="&q&pShipFromPostalCode&q&CHR(10)
StringBuilderObj.append "private const scShipFromZip4="&q&pShipFromZip4&q&CHR(10)
StringBuilderObj.append "private const scAlwAltShipAddress="&q&pAlwAltShipAddress&q&CHR(10)
StringBuilderObj.append "private const scComResShipAddress="&q&pComResShipAddress&q&CHR(10)
StringBuilderObj.append "private const scAlwNoShipRates="&q&pAlwNoShipRates&q&CHR(10)
StringBuilderObj.append "private const scUseShipMap="&q&pUseShipMap&q&CHR(10)
StringBuilderObj.append "private const scShipFromPostalCountry="&q&pShipFromPostalCountry&q&CHR(10)
StringBuilderObj.append "private const scShowProductWeight="&q&pShowProductWeight&q&CHR(10)
StringBuilderObj.append "private const scPackageWeightLimit="&q&pPackageWeightLimit&q&CHR(10)
StringBuilderObj.append "private const scShowCartWeight="&q&pShowCartWeight&q&CHR(10)
StringBuilderObj.append "private const scShowEstimateLink="&q&pShowEstimateLink&q&CHR(10)
StringBuilderObj.append "private const scHideProductPackage="&q&pHideProductPackage&q&CHR(10)

StringBuilderObj.append "private const scHideEstimateDeliveryTimes="&q&pHideEstimateDeliveryTimes&q&CHR(10)

StringBuilderObj.append "private const scShipFromWeightUnit="&q&pShipFromWeightUnit&q&CHR(10)
StringBuilderObj.append "private const scDefaultProvider="&q&pDefaultProvider&q&CHR(10)
StringBuilderObj.append "private const scHideShipAddress="&q&pHideShipAddress&q&CHR(10)

'Start SDBA
StringBuilderObj.append "private const scShipNotifySeparate="&q&tShipNotifySeparate&q&CHR(10)
'End SDBA

StringBuilderObj.append "private const PC_SECTIONSHOW="&q&pSectionShow&q&CHR(10)
StringBuilderObj.append "private const PC_RATESONLY="&q&pRatesOnly&q&CHR(10)
StringBuilderObj.append "private const PC_SHIP_DETAIL_TITLE="&q&pShipDetailTitle&q&CHR(10)
StringBuilderObj.append "private const PC_SHIP_DETAILS="&q&pShipDetails&q&CHR(10)&CHR(37)&CHR(62)


call pcs_SaveUTF8(PageName, PageName, StringBuilderObj.toString)

set StringBuilderObj=nothing

if trim(strErr)<>"" then
	response.redirect "../"&scAdminFolderName&"/modFromShipper.asp?msg="&strErr
	else
	response.redirect "../"&scAdminFolderName&"/modFromShipper.asp?s=1&message="&Server.URLEncode("Shipping settings updated successfully")
end if
%>