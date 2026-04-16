<!--#include file="adminv.asp"-->
<!--#include file="common.asp"-->
<%
Dim PageName, Body
Dim FS, f, findit

' request values
q=Chr(34)
if request.queryString("sa")<>"" then
	tTaxonCharges=pTaxonCharges
	if tTaxonCharges="" then
		tTaxonCharges=0
	end if
	tTaxonFees=pTaxonFees
	if tTaxonFees="" then
		tTaxonFees=0
	end if
	ttaxfile=ptaxfile
	if ttaxfile="" then
		ttaxfile=0
	end if
	ttaxshippingaddress=ptaxshippingaddress
	ttaxseparate=ptaxseparate
	ttaxwholesale=ptaxwholesale
	tshowVatID = pshowVatID
	tshowSSN = pshowSSN
	tVatIDReq = pVatIDReq
	tSSNReq = pSSNReq
	ttaxfilename=ptaxfilename
	ttaxCanada=ptaxCanada
	if ttaxfile="1" then
		ttaxCanada="0"
	end if
	PageName="taxsettings.asp"
	refpage="AdminTaxSettings_file.asp"
	strDelete=request.queryString("sa")
	rateArray=split(ptaxRateDefault,", ")
	stateArray=split(ptaxRateState,", ")
	taxSNHArray=split(ptaxSNH,", ")
	ttaxRateDefault=""
	ttaxRateState=""
	ttaxSNH=""
	for i=0 to ubound(stateArray)-1
		if stateArray(i)<>strDelete then
			ttaxRateDefault=ttaxRateDefault&rateArray(i)&", "
			ttaxRateState=ttaxRateState&stateArray(i)&", "
			ttaxSNH=ttaxSNH&taxSNHArray(i)&", "
		end if
	next
else
	if request.Form("RateOnly")="1" OR request("ActivateZone")="1" then
		tTaxonCharges=pTaxonCharges
		if tTaxonCharges="" then
			tTaxonCharges=0
		end if
		tTaxonFees=pTaxonFees
		if tTaxonFees="" then
			tTaxonFees=0
		end if
		ttaxfile=ptaxfile
		if ttaxfile="" then
			ttaxfile=0
		end if
		pcv_DefaultCheck=replace(trim(ptaxRateState),",","")
		if pcv_DefaultCheck="" then
			instAddComma=""
			ttaxRateDefault=""
			ttaxRateState=""
			ttaxSNH=""
		else
			stateArray=split(ptaxRateState,", ")
			if ptaxRateState<>"" and ubound(stateArray)=0 then
				instAddComma=", "
			else
				instAddComma=""
			end if
			ttaxRateDefault=ptaxRateDefault
			ttaxRateState=ptaxRateState
			ttaxSNH=ptaxSNH
		end if
		if request("RateOnly")="1" then
			ttaxRateDefault=ttaxRateDefault&instAddComma&replace(getUserInput(request.form("taxRateDefault"), 0),"%","")&", "
			if request.Form("PopForm")="YES" then
				ttaxSNH=ttaxSNH&instAddComma&getUserInput(request.Form("taxSNH"),0)&", "
			else
				ttaxSNH=ttaxSNH&instAddComma&getUserInput(request.Form("taxSNH"&stateArray(0)),0)&", "
			end if
			ttaxRateState=ttaxRateState&instAddComma&getUserInput(request.form("taxRateState"),0)&", "
			PageName=getUserInput(request.form("page_name"),0)
			refpage=getUserInput(request.form("refpage"),0)
			ttaxCanada=ptaxCanada
			if ttaxCanada="" then
				ttaxCanada="0"
			end if
			if ttaxfile="1" then
				ttaxCanada="0"
			end if
		else
			ttaxRateDefault=ptaxRateDefault
			ttaxSNH=ptaxSNH
			ttaxRateState=ptaxRateState
			ttaxCanada="1"
			PageName="taxsettings.asp"
			refpage="AddTaxPerZone.asp"
		end if
		ttaxshippingaddress=ptaxshippingaddress
		ttaxseparate=ptaxseparate
		ttaxwholesale=ptaxwholesale
		tshowVatID = pshowVatID
		tshowSSN = pshowSSN
		tVatIDReq = pVatIDReq
		tSSNReq = pSSNReq
		ttaxVATrate=ptaxVATrate
		ttaxVATRate_Code=ptaxVATRate_Code
		ttaxVAT=ptaxVAT
		ttaxdisplayVAT=ptaxdisplayVAT
		ttaxfilename=ptaxfilename
	else
		tTaxonCharges=getUserInput(request.form("TaxonCharges"),0)
		if tTaxonCharges="" then
			tTaxonCharges="0"
		end if
		tTaxonFees=getUserInput(request.form("TaxonFees"),0)
		if tTaxonFees="" then
			tTaxonFees="0"
		end if
		ttaxfile=getUserInput(request.form("taxfile"),0)
		if ttaxfile="" then
			ttaxfile="0"
		end if
		ttaxRateState=getUserInput(request.form("taxRateState"),0)&", "
		tempRateStateArray=split(ttaxRateState,", ")
		if request.Form("PopForm")="YES" then
			ttaxSNH=ttaxSNH&getUserInput(request.Form("taxSNH"),0)&", "
		else
			for j=0 to ubound(tempRateStateArray)-1
				ttaxSNH=ttaxSNH&getUserInput(request.Form("taxSNH"&tempRateStateArray(j)),0)&", "
			next
		end if
		ttaxshippingaddress=getUserInput(request.form("taxshippingaddress"),0)
		if ttaxshippingaddress="" then
			ttaxshippingaddress="0"
		end if
		ttaxseparate=getUserInput(request.form("taxseparate"),0)
		if ttaxseparate="" then
			ttaxseparate="0"
		end if
		ttaxwholesale=getUserInput(request.form("taxwholesale"),0)
		if ttaxwholesale="" then
			ttaxwholesale="0"
		end if		
		tshowVatID=getUserInput(request.form("showVatID"),0)
		if tshowVatID="" then
			tshowVatID="0"
		end if		
		tshowSSN=getUserInput(request.form("showSSN"),0)
		if tshowSSN="" then
			tshowSSN="0"
		end if
		tVatIDReq=getUserInput(request.form("VatIDReq"),0)
		if tVatIDReq="" then
			tVatIDReq="0"
		end if
		tSSNReq=getUserInput(request.form("SSNReq"),0)
		if tSSNReq="" then
			tSSNReq="0"
		end if
		ttaxVATrate=replace(getUserInput(request.form("taxVATrate"),0),"%","")
		if ttaxVATrate="" then
			ttaxVATrate="0"
		end if
		ttaxVATRate_Code=getUserInput(request.form("taxVATRate_Code"),0)
		ttaxVAT=getUserInput(request.form("taxVAT"),0)
		if ttaxVAT="" then
			ttaxVAT="0"
		end if
		ttaxdisplayVAT=getUserInput(request.form("taxdisplayVAT"),0)
		if ttaxdisplayVAT="" then
			ttaxdisplayVAT="0"
		end if
		ttaxRateDefault=replace(getUserInput(request.form("taxRateDefault"),0),"%","")&", "
		if ttaxRateState="" then
			ttaxRateDefault="0"
		end if
		If NOT isNumeric(ttaxRateDefault) then
			'ttaxRateDefault="0"
		End If
		ttaxfilename=getUserInput(request.form("taxfilename"),0)
		ttaxCanada=ptaxCanada
		if ttaxfile="1" then
			ttaxCanada="0"
		end if
		if ttaxfile="1" then

			query="DELETE FROM pcTaxZoneRates;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZonesGroups;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxGroups;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZoneDescriptions;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZones;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing

		end if
		PageName=getUserInput(request.form("page_name"),0)
		refpage=getUserInput(request.form("refpage"),0)
		if ttaxfile="1" and ttaxfilename="" then
			response.redirect "../"&scAdminFolderName&"/AdminTaxSettings_file.asp?nofilename=1"
		end if
		
		ttaxAvalara=getUserInput(request.form("taxAvalara"),0)
		if ttaxAvalara="" then
			ttaxAvalara=0
		end If
		ttaxAvalaraAccount=getUserInput(request.form("AvalaraAccount"),0)
		ttaxAvalaraLicense=getUserInput(request.form("AvalaraLicense"),0)
		ttaxAvalaraCode=getUserInput(request.form("AvalaraCode"),0)
		ttaxAvalaraProductCode=getUserInput(request.form("AvalaraProductCode"),0)
		ttaxAvalaraShippingCode=getUserInput(request.form("AvalaraShippingCode"),0)
		ttaxAvalaraHandlingCode=getUserInput(request.form("AvalaraHandlingCode"),0)
		ttaxAvalaraURL=getUserInput(request.form("AvalaraURL"),0)
		ttaxAvalaraEnabled=getUserInput(request.form("AvalaraEnabled"),0)
		if ttaxAvalaraEnabled="" then
			ttaxAvalaraEnabled=0
		end if
		ttaxAvalaraLog=getUserInput(request.form("AvalaraLog"),0)
		if ttaxAvalaraLog="" then
			ttaxAvalaraLog=0
		end if
		ttaxAvalaraAddressValidation=getUserInput(request.form("AvalaraAddressValidation"),0)
		if ttaxAvalaraAddressValidation="" then
			ttaxAvalaraAddressValidation=0
		end if
		ttaxAvalaraCommit=getUserInput(request.form("AvalaraCommit"),0)
		if ttaxAvalaraCommit="" then
			ttaxAvalaraCommit=0
		end if
		ttaxAvalaraReason=getUserInput(request.form("AvalaraReason"),0)
	end if
end if

on error resume next
if ttaxfilename <> "" then
	if PPD="1" then
		findit=Server.MapPath("/"&scPcFolder&"/pc/tax/"&ttaxfilename)
	else
		findit=Server.MapPath("../pc/tax/"&ttaxfilename)
	end if
	
	Set fso=server.CreateObject("Scripting.FileSystemObject")
	fileLocate=findit
	Set f=fso.GetFile(fileLocate)
	if err.number>0 then
		nofile=1
	else
		nofile=0
	end if
	err.number=0
end if

'// Set Avalara Defaults
if ttaxAvalara="" then
    ttaxAvalara=0
end if
if ttaxAvalaraAccount="" then
    ttaxAvalaraAccount=""
end if
if ttaxAvalaraLicense="" then
    ttaxAvalaraLicense=""
end if
if ttaxAvalaraCode="" then
    ttaxAvalaraCode=""
end if
if ttaxAvalaraProductCode="" then
    ttaxAvalaraProductCode=""
end if
if ttaxAvalaraShippingCode="" then
    ttaxAvalaraShippingCode=""
end if
if ttaxAvalaraHandlingCode="" then
    ttaxAvalaraHandlingCode=""
end if
if ttaxAvalaraEnabled="" then
    ttaxAvalaraEnabled=0
end if
if ttaxAvalaraLog="" then
    ttaxAvalaraLog=0
end if
if ttaxAvalaraAddressValidation="" then
    ttaxAvalaraAddressValidation=0
end if
if ttaxAvalaraCommit="" then
    ttaxAvalaraCommit=0
end if
if ttaxAvalaraReason="" then
    ttaxAvalaraReason=""
end if

findit=Server.MapPath(PageName)
Body=CHR(60)&CHR(37)&CHR(10)&"private const pTaxonCharges="&tTaxonCharges&CHR(10)
Body=Body & "private const pTaxonFees="&tTaxonFees&CHR(10)
Body=Body & "private const ptaxfile="&ttaxfile&CHR(10)
Body=Body & "private const ptaxsetup=1"&CHR(10)
Body=Body & "private const ptaxshippingaddress="&q&ttaxshippingaddress&q&CHR(10)
Body=Body & "private const ptaxseparate="&q&ttaxseparate&q&CHR(10)
Body=Body & "private const ptaxwholesale="&q&ttaxwholesale&q&CHR(10)
Body=Body & "private const pshowVatID="&q&tshowVatID&q&CHR(10)
Body=Body & "private const pVatIdReq="&q&tVatIdReq&q&CHR(10)
Body=Body & "private const pshowSSN="&q&tshowSSN&q&CHR(10)
Body=Body & "private const pSSNReq="&q&tSSNReq&q&CHR(10)
Body=Body & "private const ptaxVATrate="&q&ttaxVATrate&q&CHR(10)
Body=Body & "private const ptaxVATRate_Code="&q&ttaxVATRate_Code&q&CHR(10)
Body=Body & "private const ptaxVAT="&q&ttaxVAT&q&CHR(10)
Body=Body & "private const ptaxdisplayVAT="&q&ttaxdisplayVAT&q&CHR(10)
Body=Body & "private const ptaxRateDefault="&q&ttaxRateDefault&q&CHR(10)
Body=Body & "private const ptaxRateState="&q&ttaxRateState&q&CHR(10)
Body=Body & "private const ptaxSNH="&q&ttaxSNH&q&CHR(10)
Body=Body & "private const ptaxCanada="&q&ttaxCanada&q&CHR(10)
Body=Body & "private const ptaxfilename="&q&ttaxfilename&q&CHR(10)
Body=Body & "private const ptaxAvalara="&ttaxAvalara&CHR(10)
Body=Body & "private const ptaxAvalaraAccount="&q&ttaxAvalaraAccount&q&CHR(10)
Body=Body & "private const ptaxAvalaraLicense="&q&ttaxAvalaraLicense&q&CHR(10)
Body=Body & "private const ptaxAvalaraCode="&q&ttaxAvalaraCode&q&CHR(10)
Body=Body & "private const ptaxAvalaraProductCode="&q&ttaxAvalaraProductCode&q&CHR(10)
Body=Body & "private const ptaxAvalaraShippingCode="&q&ttaxAvalaraShippingCode&q&CHR(10)
Body=Body & "private const ptaxAvalaraHandlingCode="&q&ttaxAvalaraHandlingCode&q&CHR(10)
Body=Body & "private const ptaxAvalaraURL="&q&ttaxAvalaraURL&q&CHR(10)
Body=Body & "private const ptaxAvalaraEnabled="&ttaxAvalaraEnabled&CHR(10)
Body=Body & "private const ptaxAvalaraLog="&ttaxAvalaraLog&CHR(10)
Body=Body & "private const ptaxAvalaraAddressValidation="&ttaxAvalaraAddressValidation&CHR(10)
Body=Body & "private const ptaxAvalaraCommit="&ttaxAvalaraCommit&CHR(10)
Body=Body & "private const ptaxAvalaraReason="&q&ttaxAvalaraReason&q&CHR(10)&CHR(37)&CHR(62)

' create the file using the FileSystemObject

Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&scAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Tax")
end if
Set f=nothing

Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body
f.Close

Set fso=nothing
Set f=nothing

if request.Form("RateOnly")="1" then
	response.redirect "../"&scAdminFolderName&"/"&refpage&"?ro=1&rstate="&getUserInput(request.form("taxRateState"),0)&"&nofile="&nofile
else
	response.redirect "../"&scAdminFolderName&"/"&refpage&"?nofile="&nofile
end if
%>