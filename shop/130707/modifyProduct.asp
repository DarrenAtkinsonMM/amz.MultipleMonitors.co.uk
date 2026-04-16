<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<% 
dim pIdProduct, pIdSupplier, pDescription, pDetails, pPrice, pImageUrl, pListPrice, pstock, psku, plisthidden,pweight,pserviceSpec,pconfigOnly,pnotax,pnoshipping,pBToBPrice,pCost,pSmallImageUrl,pLargeImageUrl,pDeliveringTime,pHotDeal,pActive,pShowInHome,pEmailText,pFormQuantity,pIdOptionGroupA,pArequired,pIdOptionGroupB,pBrequired,pNoStock,pnoshippingtext,pcv_ProductType,pcv_IntMojoZoom,pcv_IntAdditionalImages,pAltTagText,pExemptDisc

'// Load and validate product ID
pIdProduct=request.Querystring("idProduct")
if trim(pIdProduct)="" or not validNum(pIdProduct) then
	call closeDb()
	response.redirect "msg.asp?message=2"
end if
	
'//save product to Viewed Products List
ViewedPrdList=getUserInput2(Request.Cookies("pcfront_visitedPrdsCP"),0)
if Instr(ViewedPrdList,"*" & pIdProduct & "*")>0 then
	ViewedPrdList = Replace(ViewedPrdList, "*" & pIdProduct & "*", "*")
end if
if ViewedPrdList="" then
	ViewedPrdList="*"
end if
ViewedPrdList="*" & pIdProduct & ViewedPrdList

Response.Cookies("pcfront_visitedPrdsCP")=ViewedPrdList
Response.Cookies("pcfront_visitedPrdsCP").Expires=Date() + 365
	
'// Determine product type: std, bto, item, app
'// std = "Standard" product
'// bto = "Configurable" product
'// item = "Configurable-Only Item"
'// app = "Apparel" product
pcv_ProductType=lcase(trim(request.Querystring("prdType")))
'// If not an accepted Product Type, go get it
if pcv_ProductType="" or (pcv_ProductType<>"std" and pcv_ProductType<>"bto" and pcv_ProductType<>"item") then
	tab = request.QueryString("tab")
	call closeDb()
	response.redirect "FindProductType.asp?id="&pIdProduct&"&tab="&tab
end if


query="UPDATE Products SET pcprod_SentNotice=0 WHERE IDProduct=" & pIdProduct & " AND removed=0;"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)
set rstemp=nothing

'SHW-S
call GetSHWSettings()

if shwOnOff=1 then
	query="SELECT sku FROM Products WHERE IDProduct=" & pIdProduct & ";"
	set rs=connTemp.execute(query)
	shwQty=-1
	SHWSync=0
	if not rs.eof then
		tmpSKU=rs("sku")
		set rs=nothing
		shwQty=SHWGetInventoryStatus(tmpSKU)
		if clng(shwQty)>=0 then
			query="UPDATE Products SET stock=" & shwQty & " WHERE IDProduct=" & pIdProduct & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			SHWSync=1
			call pcs_hookStockChanged(pIdProduct, "")
		end if
	end if
	set rs=nothing
end if
'SHW-E

'retrieve product details from db

query="SELECT idProduct, description, configOnly, serviceSpec, price, listPrice, bToBPrice, imageUrl, smallImageUrl, largeImageURL, sku, stock, listHidden, weight, deliveringTime, active, hotDeal,  visits, sales, emailText, formQuantity, showInHome, notax, noshipping, noprices, iRewardPoints, IDBrand, OverSizeSpec, downloadable,noStock,iRewardPoints,noshippingtext,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcprod_QtyToPound,pcprod_HideDefConfig,cost,pcProd_BackOrder,pcProd_ShipNDays,pcProd_NotifyStock,pcProd_ReorderLevel,pcSupplier_ID,pcProd_IsDropShipped,pcDropShipper_ID,pcprod_GC,pcProd_multiQty,pcProd_SkipDetailsPage,pcProd_HideSKU,pcProd_MaxSelect, pcProd_showBtoCmMsg,pcprod_DisplayLayout, pcprod_MetaTitle, pcprod_Apparel,pcprod_ShowStockMsg,pcprod_SizeImg,pcProd_ApparelRadio,pcprod_StockMsg,pcprod_SizeLink,pcprod_SizeInfo,pcprod_SizeURL,pcProd_Surcharge1, pcProd_Surcharge2, pcPrd_MojoZoom,pcProd_GoogleCat,pcProd_GoogleGender,pcProd_GoogleAge,pcProd_GoogleSize,pcProd_GoogleColor,pcProd_GooglePattern,pcProd_GoogleMaterial, details, sDesc, pcprod_MetaDesc, pcprod_MetaKeywords,pcProd_PrdNotes, pcSC_ID , pcProd_Top,pcProd_TopLeft,pcProd_TopRight,pcProd_Middle,pcProd_Bottom,pcProd_Tabs,pcProd_AvalaraTaxCode,pcProd_AdditionalImages,pcProd_AltTagText FROM products WHERE products.idProduct=" &pIdProduct &" AND products.removed=0;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error on modifyProduct in main SQL query") 
end if

if rs.EOF then
	set rs=nothing
	
	call closeDb()
	Session("message") = "The product that you are trying to load either does not exist or has been removed from the Control Panel."
	response.redirect "msgb.asp"
end if

pIdProduct=rs("idProduct")
pDescription=rs("description")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pListPrice=rs("listPrice")
pBToBPrice=rs("bToBPrice")
pImageUrl=rs("imageUrl")
pSmallImageUrl=rs("smallImageUrl")
pLargeImageUrl=rs("largeImageUrl")
pSku=rs("sku")
pStock=rs("stock")
pListhidden=rs("listhidden")
tWeight=rs("weight")
pPounds=Int(tWeight/16)
pWeight_oz=tWeight-(pPounds*16)
pWeight=pPounds
pKilos=Int(tWeight/1000)
pWeight_g=tWeight-(pKilos*1000)
pWeight_kg=pKilos
pDeliveringTime=rs("deliveringTime")
pActive=rs("active")
pHotDeal=rs("hotDeal")
pEmailText=rs("emailText")
pFormQuantity=rs("formQuantity")
pShowInHome=rs("showInHome")
pnotax=rs("notax") 
pnoshipping=rs("noshipping")
pnoprices=rs("noprices")
iRewardPoints=rs("iRewardPoints")
pIDBrand=rs("IDBrand") 
pOverSizeSpec=rs("OverSizeSpec")
	if pOverSizeSpec="" or isNull(pOverSizeSpec) then
		pOverSizeSpec="NO"
	end if
pdownloadable=rs("downloadable")
if NOT pdownloadable<>"" then
	pdownloadable="0"
end if
pNoStock=rs("noStock")
pnoshippingtext=rs("noshippingtext")

pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
if pcv_intHideBTOPrice<>"" then
else
pcv_intHideBTOPrice=0
end if

pcv_intQtyValidate=rs("pcprod_QtyValidate")
if pcv_intQtyValidate<>"" then
else
pcv_intQtyValidate="0"
end if

pcv_lngMinimumQty=rs("pcprod_MinimumQty")
if pcv_lngMinimumQty<>"" then
else
pcv_lngMinimumQty="0"
end if

pcv_QtyToPound=rs("pcprod_QtyToPound")

phideDefConfig=rs("pcprod_HideDefConfig")
if IsNull(phideDefConfig) or (phideDefConfig="") then
	phideDefConfig="0"
end if	

'Start SDBA
pCost=rs("cost")
if isNULL(pCost) OR pCost="" then
	pCost="0"
end If

pcbackorder=rs("pcProd_backorder")
if (pcbackorder="") or (not IsNumeric(pcbackorder)) then
	pcbackorder="0"
end If

pcShipNDays=rs("pcProd_ShipNDays")
if (pcShipNDays="") or (not IsNumeric(pcShipNDays)) then
	pcShipNDays="0"
end If

pcnotifystock=rs("pcProd_notifystock")
if (pcnotifystock="") or (not IsNumeric(pcnotifystock)) then
	pcnotifystock="0"
end If

pcreorderlevel=rs("pcProd_reorderlevel")
if (pcreorderlevel="") or (not IsNumeric(pcreorderlevel)) then
	pcreorderlevel="0"
end If

pcIDSupplier=rs("pcSupplier_ID")
if (pcIDSupplier="") or (not IsNumeric(pcIDSupplier)) then
	pcIDSupplier="0"
end If

pcIsdropshipped=rs("pcProd_IsDropShipped")
if (pcIsdropshipped="") or (not IsNumeric(pcIsdropshipped)) then
	pcIsdropshipped="0"
end If

pcIDDropShipper=rs("pcDropShipper_ID")
if (pcIDDropShipper="") or (not IsNumeric(pcIDDropShipper)) then
	pcIDDropShipper="0"
end If
'End SDBA

pcv_intSkipDetailsPage=rs("pcProd_SkipDetailsPage")
if IsNull(pcv_intSkipDetailsPage) or (pcv_intSkipDetailsPage="") then
	pcv_intSkipDetailsPage="0"
end if

'GGG add-on start
pGC=rs("pcprod_GC")
if NOT pGC<>"" then
	pGC="0"
end if
'GGG add-on end

pcv_multiQty=rs("pcProd_multiQty")
if IsNull(pcv_multiQty) or pcv_multiQty="" then
	pcv_multiQty=0
end if
pHideSKU=rs("pcProd_HideSKU")
if IsNull(pHideSKU) or pHideSKU="" then
	pHideSKU=0
end if

pMaxSelect=rs("pcProd_MaxSelect")
if pMaxSelect="" OR IsNull(pMaxSelect) then
	pMaxSelect=0
end if

pcv_showBtoCmMsg=rs("pcProd_showBtoCmMsg")
if IsNull(pcv_showBtoCmMsg) or pcv_showBtoCmMsg="" then
	pcv_showBtoCmMsg=0
end if

pDisplayLayout=LCase(rs("pcprod_DisplayLayout"))
pStrPrdMetaTitle=rs("pcprod_MetaTitle")

pcv_Apparel=rs("pcprod_Apparel")
pcv_ShowStockMsg=rs("pcprod_ShowstockMsg")
pcv_SizeImg=rs("pcprod_SizeImg")
pcv_ApparelRadio=rs("pcProd_ApparelRadio")
if pcv_ApparelRadio<>"" then
else
	pcv_ApparelRadio="0"
end if
pcv_StockMsg=rs("pcprod_StockMsg")
pcv_SizeLink=rs("pcprod_SizeLink")
pcv_SizeInfo=rs("pcprod_SizeInfo")

pcv_SizeURL=rs("pcprod_SizeURL")
if pcv_SizeURL="" then
	pcv_SizeURL="http://"
end if

pcv_Surcharge1=rs("pcProd_Surcharge1")
pcv_Surcharge2=rs("pcProd_Surcharge2")
pcv_IntMojoZoom=rs("pcPrd_MojoZoom")
pcv_IntAdditionalImages=rs("pcProd_AdditionalImages")

pAltTagText=rs("pcProd_AltTagText")

'//Get Google Shopping Settings
pcv_GCat=rs("pcProd_GoogleCat")
if pcv_GCat<>"" then
	pcv_GCat=replace(pcv_GCat,"&","&amp;")
	pcv_GCat=replace(pcv_GCat,">","&gt;")
end if
pcv_GGen=rs("pcProd_GoogleGender")
if pcv_GGen<>"" then
	pcv_GGen=replace(pcv_GGen,"&","&amp;")
	pcv_GGen=replace(pcv_GGen,">","&gt;")
end if
pcv_GAge=rs("pcProd_GoogleAge")
if pcv_GAge<>"" then
	pcv_GAge=replace(pcv_GAge,"&","&amp;")
	pcv_GAge=replace(pcv_GAge,">","&gt;")
end if
pcv_GSize=rs("pcProd_GoogleSize")
if pcv_GSize<>"" then
	pcv_GSize=replace(pcv_GSize,"&","&amp;")
	pcv_GSize=replace(pcv_GSize,">","&gt;")
end if
pcv_GColor=rs("pcProd_GoogleColor")
if pcv_GColor<>"" then
	pcv_GColor=replace(pcv_GColor,"&","&amp;")
	pcv_GColor=replace(pcv_GColor,">","&gt;")
end if
pcv_GPat=rs("pcProd_GooglePattern")
if pcv_GPat<>"" then
	pcv_GPat=replace(pcv_GPat,"&","&amp;")
	pcv_GPat=replace(pcv_GPat,">","&gt;")
end if
pcv_GMat=rs("pcProd_GoogleMaterial")
if pcv_GMat<>"" then
	pcv_GMat=replace(pcv_GMat,"&","&amp;")
	pcv_GMat=replace(pcv_GMat,">","&gt;")
end if

' NTEXT fields
pDetails=rs("details")
psDesc=rs("sDesc")
pStrPrdMetaDesc=rs("pcprod_MetaDesc")
pStrPrdMetaKeywords=rs("pcprod_MetaKeywords")
pcv_PrdNotes=rs("pcProd_PrdNotes")

pcSCID=rs("pcSC_ID")
IF pcSCID="" OR (IsNull(pcSCID)) then
	pcSCID=0
END IF

ppTop=rs("pcProd_Top")
ppTopLeft=rs("pcProd_TopLeft")
ppTopRight=rs("pcProd_TopRight")
ppMiddle=rs("pcProd_Middle")
ppTabs=rs("pcProd_Tabs")
ppBottom=rs("pcProd_Bottom")

pAvalaraTaxCode=rs("pcProd_AvalaraTaxCode")

set rs=nothing

' replace characters from details and not for sale field
pDetails=pcf_PrintCharacters(pDetails)
psDesc=pcf_PrintCharacters(psDesc)
pEmailText=pcf_PrintCharacters(pEmailText)

'GGG add-on start
if (pGC<>"") and (pGC="1") then
	query="select pcGC_Exp,pcGC_ExpDate,pcGC_ExpDays,pcGC_EOnly,pcGC_CodeGen,pcGC_GenFile from pcGC where pcGC_idproduct=" & pIDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error on modifyProduct around line 205") 
	end if
	
	if not rs.eof then
		pGCExp=rs("pcGC_Exp")
		pGCExpDate=rs("pcGC_ExpDate")
		if pGCExpDate<>"" then
			if year(pGCExpDate)=1900 then
				pGCExpDate=""
			else
				if scDateFrmt="DD/MM/YY" then
					pGCExpDate=day(pGCExpDate) & "/" & month(pGCExpDate) & "/" & year(pGCExpDate)
				else
					pGCExpDate=month(pGCExpDate) & "/" & day(pGCExpDate) & "/" & year(pGCExpDate)
				end if
			end if
		end if
		
		pGCExpDay=rs("pcGC_ExpDays")
		pGCEOnly=rs("pcGC_EOnly")
		pGCGen=rs("pcGC_CodeGen")
		pGCGenFile=rs("pcGC_GenFile")
	end if
	set rs=nothing
end if

'GGG add-on end

if (pdownloadable<>"") and (pdownloadable="1") then
	query="select ProductURL, URLExpire, ExpireDays, License, localLG, RemoteLG, LicenseLabel1, LicenseLabel2, LicenseLabel3, LicenseLabel4, LicenseLabel5, AddtoMail from DProducts where idproduct=" & pIDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error on modifyProduct around line 241") 
	end if
	
	if not rs.eof then
		pProductURL=rs("ProductURL")
		pURLExpire=rs("URLExpire")
		pExpireDays=rs("ExpireDays")
		pLicense=rs("License")
		pLocalLG=rs("localLG")
		pRemoteLG=rs("RemoteLG")
		if pRemoteLG<>"" then
		else
			pRemoteLG="http://"
		end if
		pLicenseLabel1=rs("LicenseLabel1")
		pLicenseLabel2=rs("LicenseLabel2")
		pLicenseLabel3=rs("LicenseLabel3")
		pLicenseLabel4=rs("LicenseLabel4")
		pLicenseLabel5=rs("LicenseLabel5")
		pAddtoMail=rs("AddtoMail")
	else
		pProductURL=""
		pURLExpire="0"
		pExpireDays=""
		pLicense="0"
		pLocalLG=""
		pRemoteLG="http://"
		pLicenseLabel1=""
		pLicenseLabel2=""
		pLicenseLabel3=""
		pLicenseLabel4=""
		pLicenseLabel5=""
		pAddtoMail=""
	end if
	set rs=nothing
else
	pProductURL=""
	pURLExpire="0"
	pExpireDays=""
	pLicense="0"
	pLocalLG=""
	pRemoteLG="http://"
	pLicenseLabel1=""
	pLicenseLabel2=""
	pLicenseLabel3=""
	pLicenseLabel4=""
	pLicenseLabel5=""
	pAddtoMail=""
end if

'query="SELECT pcFPE_IdProduct FROM pcDFProdsExempt WHERE pcFPE_IdProduct=" & pIDProduct
'set rs=server.CreateObject("ADODB.RecordSet")
'set rs=connTemp.execute(query)
'if not rs.eof then
'	pExemptDisc = 1
'end if

Dim pageTitle, pageIcon, section
pageTitle="Modify Product: <strong>" & pDescription & "</strong>"
pageIcon="pcv4_icon_inventoryAdded.gif"
section="products" 
%>
<!--#include file="AdminHeader.asp"-->
<!-- #include file="../htmleditor/editor.asp" -->
<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->

<script type=text/javascript>
<%' GGG add-on start%>
function check_date(field){
var checkstr = "0123456789";
var DateField = field;
var Datevalue = "";
var DateTemp = "";
var seperator = "/";
var day;
var month;
var year;
var leap = 0;
var err = 0;
var i;
   err = 0;
   DateValue = DateField.value;
   /* Delete all chars except 0..9 */
   for (i = 0; i < DateValue.length; i++) {
	  if (checkstr.indexOf(DateValue.substr(i,1)) >= 0) {
	     DateTemp = DateTemp + DateValue.substr(i,1);
	  }
	  else
	  {
	  if (DateTemp.length == 1)
		{
    	  DateTemp = "0" + DateTemp
		}
	  else
	  {
	  	if (DateTemp.length == 3)
	  	{
	  	DateTemp = DateTemp.substr(0,2) + '0' + DateTemp.substr(2,1);
	  	}
	  }
	 }
   }
   DateValue = DateTemp;
   /* Always change date to 8 digits - string*/
   /* if year is entered as 2-digit / always assume 20xx */
   if (DateValue.length == 6) {
      DateValue = DateValue.substr(0,4) + '20' + DateValue.substr(4,2); }
   if (DateValue.length != 8) {
      return(false);}
   /* year is wrong if year = 0000 */
   year = DateValue.substr(4,4);
   if (year == 0) {
      err = 20;
   }
   /* Validation of month*/
   <%if scDateFrmt="DD/MM/YY" then%>
   month = DateValue.substr(2,2);
   <%else%>
   month = DateValue.substr(0,2);
   <%end if%>
   if ((month < 1) || (month > 12)) {
      err = 21;
   }
   /* Validation of day*/
   <%if scDateFrmt="DD/MM/YY" then%>
   day = DateValue.substr(0,2);
   <%else%>
   day = DateValue.substr(2,2);
   <%end if%>
   if (day < 1) {
     err = 22;
   }
   /* Validation leap-year / february / day */
   if ((year % 4 == 0) || (year % 100 == 0) || (year % 400 == 0)) {
      leap = 1;
   }
   if ((month == 2) && (leap == 1) && (day > 29)) {
      err = 23;
   }
   if ((month == 2) && (leap != 1) && (day > 28)) {
      err = 24;
   }
   /* Validation of other months */
   if ((day > 31) && ((month == "01") || (month == "03") || (month == "05") || (month == "07") || (month == "08") || (month == "10") || (month == "12"))) {
      err = 25;
   }
   if ((day > 30) && ((month == "04") || (month == "06") || (month == "09") || (month == "11"))) {
      err = 26;
   }
   /* if 00 ist entered, no error, deleting the entry */
   if ((day == 0) && (month == 0) && (year == 00)) {
      err = 0; day = ""; month = ""; year = ""; seperator = "";
   }
   /* if no error, write the completed date to Input-Field (e.g. 13.12.2001) */
   if (err == 0) {
	<%if scDateFrmt="DD/MM/YY" then%>
	DateField.value = day + seperator + month + seperator + year;
    <%else%>
	DateField.value = month + seperator + day + seperator + year;   
    <%end if%>
	return(true);
   }
   /* Error-message if err != 0 */
   else {
	return(false);   
   }
}
<%' GGG add-on end%>

function isDigit(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}
	
function CountStr(tmpStr,subStr)
{
   var substrings = tmpStr.split(subStr);
   count=substrings.length - 1;
   return count;
}
	
function Form1_Validator(theForm)
{	
	// InnovaStudio HTML Editor Workaround for this keyword
	theForm = document.hForm;

	SavePPToFields();
	
	if (theForm.sku.value == "")
  	{
		alert("Please enter a SKU or part number.");
	    return (false);
	}
	if (theForm.description.value == "")
  	{
		alert("Please enter a name for the product.");
	    return (false);
	}
	if (theForm.details.value == "")
  	{
		alert("Please enter a description for the product.");
	    return (false);
	}
	
	<%if pcv_ProductType<>"item" then%>

	if (theForm.downloadable1.value == "1")
  	{
  	
  		if (theForm.producturl.value == "")
  		{
			alert("Please enter the full physical path to the downloadable file");
		    return (false);
		}

	  	if (theForm.urlexpire1.value == "1")
	  	{
  	
		  	if (theForm.expiredays.value == "")
		  	{
				alert("Please enter a number for this field.");
			    return (false);
			}
	
			if (allDigit(theForm.expiredays.value) == false)
			{
			    alert("Please enter a number for this field.");
			    return (false);
			}
	
			if (theForm.expiredays.value == "0")
			{
			    alert("Please enter a number greater than zero for this field.");
			    return (false);
			}
		}
	
		if (theForm.license1.value == "1")
	  	{
  	
		  	if ((theForm.locallg.value == "") && ((theForm.remotelg.value == "") || (theForm.remotelg.value == "http://")) )
		  	{
				alert("Please fill out one of License Generator fields.");
			    return (false);
			}
	
		  	if ((theForm.locallg.value != "") && (theForm.remotelg.value != "") && (theForm.remotelg.value != "http://") )
		  	{
				alert("Please fill out only one of License Generator fields.");
			    return (false);
			}
	
			if ((theForm.licenselabel1.value == "") && (theForm.licenselabel2.value == "") && (theForm.licenselabel3.value == "") && (theForm.licenselabel4.value == "") && (theForm.licenselabel5.value == ""))
		  	{
				alert("Please fill out at least one of License Label fields.");
			    return (false);
			}
		}
	}
	
	<%end if%>
	
	<%' GGG add-on start
	if pcv_ProductType="std" then %>
	
	if (theForm.GC[0].checked == true)
  	{
  		if (theForm.GCExp[1].checked == true)
	  	{
	  		if (theForm.GCExpDate.value == "")
		  	{
				alert("Please enter a valid date for this field.");
			    return (false);
			}
			if (check_date(theForm.GCExpDate) == false)
		  	{
				alert("Please enter a valid date for this field.");
			    return (false);
			}
	  	}
		if (theForm.GCExp[2].checked == true)
	  	{
		  	if (theForm.GCExpDay.value == "")
		  	{
				alert("Please enter a number for this field.");
			    return (false);
			}
	
			if (allDigit(theForm.GCExpDay.value) == false)
			{
			    alert("Please enter a number for this field.");
			    return (false);
			}
	
			if (theForm.GCExpDay.value == "0")
			{
			    alert("Please enter a number greater than zero for this field.");
			    return (false);
			}

		}
		
		if (theForm.GCGen[1].checked == true)
	  	{
		  	if (theForm.GCGenFile.value == "")
		  	{
				alert("Please fill out the File Name field.");
			    return (false);
			}
	  	}
	}
	
	<% end if
	' GGG add-on end%>
	
	<% 'Check duplicated Custom Input Fields in Product Page Layout%>
	if (theForm.displayLayout.value=="t")
	{
	var CIF=0;
	if (theForm.ppTop.value != "") CIF=CIF+CountStr(theForm.ppTop.value,"PrdInput");
	if (theForm.ppTopLeft.value != "") CIF=CIF+CountStr(theForm.ppTopLeft.value,"PrdInput");
	if (theForm.ppTopRight.value != "") CIF=CIF+CountStr(theForm.ppTopRight.value,"PrdInput");
	if (theForm.ppMiddle.value != "") CIF=CIF+CountStr(theForm.ppMiddle.value,"PrdInput");
	if (theForm.ppTabs.value != "") CIF=CIF+CountStr(theForm.ppTabs.value,"PrdInput");
	if (theForm.ppBottom.value != "") CIF=CIF+CountStr(theForm.ppBottom.value,"PrdInput");
	
	if (CIF>1)
	{
		alert("You cannot add 'Custom Input Fields' into multiple areas of Product Page Layout.");
		return (false);
	}
	}
		
	try
	{
		document.hForm.pcIDDropShipper.disabled=false;
		document.hForm.pcIDSupplier.disabled=false;
	}
	catch(err)
	{
		//Do nothing
	}
	return (true);
}


function CheckWindow() {
options = "toolbar=0,status=0,menubar=0,scrollbars=0,resizable=0,width=600,height=400";
myloc='testurl.asp?file1=' + document.hForm.producturl.value + '&file2=' + document.hForm.locallg.value + '&file3=' + document.hForm.remotelg.value;
newcheckwindow=window.open(myloc,"mywindow", options);
}

function TestWindow() {
options = "toolbar=0,status=0,menubar=0,scrollbars=0,resizable=0,width=600,height=400";
myloc="testlg.asp?idproduct=<%=pIdProduct%>";
newtestwindow=window.open(myloc,"mywindow1", options);
}

function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
}

function newWindow2(file,window) {
		catWindow=open(file,window,'resizable=no,width=500,height=600,scrollbars=1');
		if (catWindow.opener == null) catWindow.opener = self;
}

//Open Sale Details
function winSale(fileName)
	{
		myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=650,height=300')
		myFloater.location.href=fileName;
	}

// Set mouse cursor focus on page load
function setCursorFocus(){
document.hForm.sku.focus();
}
onload = function() {setCursorFocus()}
</script>
<%
'// CONFIGURABLE-ONLY
'// Find out if product has been assigned to any configurable products
IF scBTO = 1 THEN
	Dim pcBtoAssignmnt
	query="SELECT DISTINCT products.idproduct, products.description FROM products INNER JOIN configSpec_products ON (products.idproduct=configSpec_products.specProduct) WHERE configSpec_products.configProduct="&pidProduct&" UNION (SELECT DISTINCT products.idproduct, products.description FROM products INNER JOIN configSpec_Charges ON (products.idproduct=configSpec_Charges.specProduct) WHERE configSpec_Charges.configProduct="&pidProduct&") ORDER BY products.Description ASC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
		if err.number <> 0 then
			set rs=nothing
			
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error retrieving configurable product assignments") 
		end if
	if rs.EOF then
		pcBtoAssignmnt=0
	else
		pcBtoAssignmnt=1
	end if
	set rs=nothing
ELSE
	pcBtoAssignmnt=0
END IF
%>
<%'SHW-S
IF SHWSync=1 THEN%>
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessageInfo">
			This Product Inventory Status has been synchronized with SHIPWIRE
		</div>
	</td>
</tr>
</table>
<%END IF
'SHW-E%>
<form method="post" name="hForm" action="modifyProductB.asp" onSubmit="return Form1_Validator(this);" class="pcForms">
	
	<% '// PRODUCT NAME & TEXT NAVIGATION - Start %>
		<div class="cpOtherLinks">
			<% if pcv_ProductType<>"item" then %>
            <a href="../pc/viewPrd.asp?idproduct=<%=pIdProduct%>&adminPreview=1" target="_blank">Preview</a></font>
            <% end if %>
            <% if pcv_ProductType="std" then %>
				<% if pcv_Apparel="1" then %>
					: <a href="app-subPrdsMngAll.asp?idproduct=<%=pIdProduct%>">Sub-products</a>
				<% end if %>
             : <a href="modPrdOpta.asp?idproduct=<%=pIdProduct%>">Options</a>
            <% elseif pcv_ProductType="bto" then %>
             : <a href="modBTOconfiga.asp?idproduct=<%=pIdProduct%>">Configuration</a>
            <% end if %>
            <% if pcv_ProductType<>"item" then %>
             : <a href="AdminCustom.asp?idproduct=<%=pIdProduct%>">Custom fields</a>
             : <a href="crossSellEdit.asp?idmain=<%=pIdProduct%>">Cross-selling</a>
             : <a href="ModPromotionPrd.asp?idproduct=<%=pIdProduct%>&iMode=start">Promotion</a>
            <% end if %>
			<% if pcv_ProductType<>"item" then %> :<% end if %>
            <a href="FindProductQtyDisc.asp?idproduct=<%=pIdProduct%>">Qty Discounts</a>
             <%
			 ' START - Check whether Product Reviews are active, show links
				query = "SELECT pcRS_Active FROM pcRevSettings;"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				pcv_Active=rs("pcRS_Active")
				if isNull(pcv_Active) or pcv_Active="" then
					pcv_Active="0"
				end if
				Set rs=Nothing
				if pcv_Active<>"0" then
			 %>
                 : Reviews: <a href="prv_ManageReviews.asp?IDProduct=<%=pIdProduct%>&nav=2">Live</a>&nbsp;<a href="prv_ManageReviews.asp?IDProduct=<%=pIdProduct%>&nav=1">Pending</a>
			 <%
				end if
			 ' END - Product Reviews links
			 %>
            <% if pcBtoAssignmnt=1 then %>
             : <a href="BTOAssignments.asp?idproduct=<%=pIdProduct%>">Product Assignments</a>
            <% end if %>
             : <a href="FindDupProductType.asp?idproduct=<%=pIdProduct%>">Clone</a>
             : <a href="LocateProducts.asp">Search</a>
             | Type:&nbsp;
            <% if pcv_ProductType="std" then %>
                Standard
            <% elseif pcv_ProductType="app" then %>
                Apparel
            <% elseif pcv_ProductType="bto" then %>
                Configurable Product
            <% else %>
                Configurable-Only Item
            <% end if %>
		</div>
	<% '// PRODUCT NAME & TEXT NAVIGATION - End %>
    
    <%
	'// START - Promotion + Quantity Discounts Test (Rare scenario)
	Dim pcProductTest1, pcProductTest2, pcProductTest3
	pcProductTest1=0
	pcProductTest2=0
	pcProductTest3=0
		'// Test for promotions
		query="SELECT DISTINCT idproduct FROM pcPrdPromotions WHERE idproduct=" & pIDProduct & ";"
		set rsTest=Server.CreateObject("ADODB.Recordset")
		set rsTest=connTemp.execute(query)
		if not rsTest.eof then
			pcProductTest1=1
		end if
		'// Test for quantity discounts
		query="SELECT DISTINCT idproduct FROM discountsPerQuantity WHERE idproduct=" & pIDProduct & ";"
		set rsTest=connTemp.execute(query)
		if not rsTest.eof then
			pcProductTest2=1
		end if
		'// Test for category-based quantity discounts
		query="SELECT DISTINCT pcCD_idcategory FROM pcCatDiscounts WHERE pcCD_idcategory IN (SELECT DISTINCT idcategory FROM categories_products WHERE idproduct=" & pidProduct & ");"
		set rsTest=connTemp.execute(query)
		if not rsTest.eof then
			pcProductTest3=1
		end if
		set rsTest=nothing
		
	if pcProductTest1=1 and pcProductTest2=1 then ' Promotion and product-based quantity discount
	%>
		<div class="pcCPmessage">	
		There is a problem: there are both a Promotion and Quantity Discounts running on this product. You need to <a href="javascript:if (confirm('Quantity discounts will be removed. Do you wish to continue?')) location='moddctQtyPrd.asp?Delete=Yes&idproduct=<%=pIDProduct%>'">remove the quantity discounts</a> or <a href="javascript:if (confirm('The promotion will be removed. Do you wish to continue?')) location='ModPromotionPrd.asp?Delete=Yes&idproduct=<%=pIdProduct%>'">remove the promotion</a>.
		</div>
    <%
	elseif pcProductTest1=1 and pcProductTest3=1 then
	%>
		<div class="pcCPmessage">	
		There is a problem: there are both a Promotion and Category-based Quantity Discounts running on this product. You need to <a href="javascript:if (confirm('The promotion will be removed. Do you wish to continue?')) location='ModPromotionPrd.asp?Delete=Yes&idproduct=<%=pIdProduct%>'">remove the promotion</a> or remove the category-based quantity discounts (which may affect other products).
		</div>
    <%
	end if
	'// END - Promotion + Quantity Discounts Test
	%>
	
	<%'SM-START
	if pcSCID="0" then
		query="SELECT pcSales_Pending.pcSales_ID,pcSales.pcSales_Name FROM pcSales INNER JOIN pcSales_Pending ON pcSales.pcSales_ID=pcSales_Pending.pcSales_ID WHERE (pcSales_Pending.idProduct=" & pIdProduct & ") AND (pcSales_Pending.pcSales_ID NOT IN (SELECT pcSales_ID FROM pcSales_Completed));"
		set rsS=Server.CreateObject("ADODB.Recordset")
		set rsS=conntemp.execute(query)
					
		if not rsS.eof then
			tmpArr=rsS.getRows()
			intC=ubound(tmpArr,2)
			%>
			<div class="pcCPmessageInfo">	
				This product is currently included in the pending sale: &nbsp;
				<select id="psaleid">
					<%For k=0 to intC%>
					<option value="<%=tmpArr(0,k)%>"><%=tmpArr(1,k)%></option>
					<%Next%>
				</select>
				<br>
				<a href="javascript:if (confirm('This product will be removed from the selected pending sale. Do you wish to continue?')) location='sm_RmvPrd.asp?id=<%=pIdProduct%>&saleid='+document.hForm.psaleid.value;">Click here</a> to remove this product from the selected sale.
			</div>
		<%end if
		set rsS=nothing
	else
		if pcSCID>"0" then
			query="SELECT pcSC_SaveName,pcSC_SaveIcon FROM pcSales_Completed WHERE pcSC_ID=" & pcSCID & ";"
			set rsS=connTemp.execute(query)
			if not rsS.eof then
				pcSCName=rsS("pcSC_SaveName")
				pcSCIcon=rsS("pcSC_SaveIcon")
				%>
				<div class="pcCPmessageInfo">This product is currently included in the running Sale: <a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><%=pcSCName%></a></div>
				<%
			end if
			set rsS=nothing
			query="SELECT pcSales_TargetPrice,pcSB_Price FROM pcSales_BackUp WHERE pcSC_ID=" & pcSCID & " AND IDProduct=" & pIdProduct & ";"
			set rsS=connTemp.execute(query)
			pcTPrice=0
			pcBUPrice=0
			if not rsS.eof then
				pcTPrice=rsS("pcSales_TargetPrice")
				pcBUPrice=rsS("pcSB_Price")
				if (pcTPrice="-1") AND (pcBUPrice=0) then
					pcBUPrice=pPrice
				end if
			end if
			set rsS=nothing
		end if
	end if
	'SM-END%>
	

		<%
		'// TABBED PANELS - MAIN DIV START
		%>
	  <div id="TabbedPanels1" class="tabbable-left">

		<%
		'// TABBED PANELS - START NAVIGATION
		%>
		<div class="col-xs-3">
			<ul class="nav nav-tabs tabs-left">
				<li class="active">
					<a href="#tab-1" data-toggle="tab">
						Product Details
						<div class="pcCPextraInfo"><span class="pcSmallText">SKU, Descriptions, and Meta Tags</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-2" data-toggle="tab">
						Prices
						<div class="pcCPextraInfo"><span class="pcSmallText">Online, Retail, &amp; Wholesale Prices</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-3" data-toggle="tab">
						Categories
						<div class="pcCPextraInfo"><span class="pcSmallText">Assign to Multiple Categories</span></div>
					</a>
				</li>				
				<li>
					<a href="#tab-4" data-toggle="tab">
						Images
						<div class="pcCPextraInfo"><span class="pcSmallText">Primary &amp; Additional Images</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-5" data-toggle="tab">
						Product Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Brand, Tax, &amp; Display Options</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-6" data-toggle="tab">
						Shipping Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Inventory, Weight, &amp; Shipping</span></div>
					</a>
				</li>
				<% if pcv_ProductType<>"item" then %>
				<li>
					<a href="#tab-7" data-toggle="tab">
						Product Page Layout
						<div class="pcCPextraInfo"><span class="pcSmallText">Preset, Custom, &amp; Tabbed Layouts</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-8" data-toggle="tab">
						Downloadable Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Make it a Digital Product</span></div>
					</a>
				</li>
				<% end if %>
				<% if pcv_ProductType="std" then %>
				<li>
					<a href="#tab-9" data-toggle="tab">
						Gift Certificate Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Make it a Gift Certificate</span></div>
					</a>
				</li>
				<% If statusAPP="1" OR scAPP=1 Then %>
				<li>
					<a href="#tabs-13" data-toggle="tab">
						Apparel Product Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Make it an Apparel Product</span></div>
					</a>
				</li>
				<% End If %>
				<% end if %>
				<% if pcv_ProductType<>"item" then %>
				<li>
					<a href="#tab-10" data-toggle="tab">
						Custom Fields
						<div class="pcCPextraInfo"><span class="pcSmallText">Manage Custom Search Fields</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-11" data-toggle="tab">
						Google Shopping Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Setup with Google Shopping</span></div>
					</a>
				</li>
				<% end if %>
				<li>
					<div style="margin-top:10px; margin-bottom: 10px; margin-right: 5px; text-align: center">
						<input type="hidden" name="idproduct" value="<%=pIdProduct%>">
						<input type="hidden" name="idsupplier" value="10">
						<input type="hidden" name="re1" value="0">
						<input type="hidden" name="prdType" value="<%=pcv_ProductType%>">
										<input type="hidden" name="tab" id="tab"  value="<%=Request("tab")%>">
						<input type="submit" name="Submit" value="Save" class="btn btn-primary" onClick="document.hForm.re1.value='1';SavePPToFields();">
                
										<% 
						Dim varMonth, varDay, varYear, varMonthStart, varDayStart, varYearStart, dtInputStrStart, dtInputStr
						varMonth=Month(Date)
						varDay=Day(Date)
						varYear=Year(Date)
						dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
						if scDateFrmt="DD/MM/YY" then
							dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
						end if
						varMonthStart=Month(Date()-29)
						varDayStart=Day(Date()-29)
						varYearStart=Year(Date()-29)
						dtInputStrStart=(varMonthStart&"/"&varDayStart&"/"&varYearStart)
						if scDateFrmt="DD/MM/YY" then
							dtInputStrStart=(varDayStart&"/"&varMonthStart&"/"&varYearStart)
						end if
						%>
        				<input type="button" class="btn btn-default"  name="recentSales" value="Recent Sales" onClick="window.open('PrdsalesReport.asp?FromDate=<%=replace(dtInputStrStart,"/","\%2F")%>&ToDate=<%=replace(dtInputStr,"/","\%2F")%>&basedon=1&IDProduct=<%=pIdProduct%>&submit=Search')">
					</div>
						</li>
			</ul>
		</div>
		<%
		'// TABBED PANELS - END NAVIGATION
		%>
		
		<%
		'// TABBED PANELS - START PANELS
		%>
        <div class="col-xs-9">
            <div class="tab-content">
		
			<%
			'// =========================================
			'// FIRST PANEL - START - Name, SKU, descriptions
			'// =========================================
			%>
				<div id="tab-1" class="tab-pane active">
				
					<table class="pcCPcontent">				
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_4")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td>Product ID:</td>
							<td><%=pidProduct%></td>
						</tr>
						<tr> 
							<td nowrap>SKU (Part Number): <img src="images/pc_required.gif" alt="required field" width="9" height="9"> </td>
							<td>  
								<input type="text" name="sku" value="<%=pSku%>" size="30" tabindex="101">
								<input type="hidden" name="origsku" value="<%=pSku%>">
							</td>
						</tr>
						<tr> 
							<td>Name: <img src="images/pc_required.gif" alt="required field" width="9" height="9"> </td>
								<td>
								<input type="text" name="description" value="<%=pDescription%>" size="40" tabindex="102">
							</td>
						</tr>
						<tr> 
							<td valign="top">Description: <img src="images/pc_required.gif" alt="required field" width="9" height="9"></td>
							<td>								
								<textarea class="htmleditor" name="details" id="details" rows="6" cols="60" tabindex="103"><%=pDetails%></textarea>                
							</td>
						</tr>
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr> 
							<td valign="top">Short Description:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=401"></a></td>
							<td valign="top">
								<textarea name="sdesc" rows="6" cols="56" tabindex="104"><%=psDesc%></textarea>
							</td>
						</tr>
						<% end if %>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<th colspan="2">SEO/Meta Tags&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=204"></a></th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td align="right" valign="top">Title: </td>
							<td><textarea name="PrdMetaTitle" cols="50" rows="3" tabindex="1001"><%=pStrPrdMetaTitle%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Description: </td>
							<td><textarea name="PrdMetaDesc" cols="50" rows="6" tabindex="1002"><%=pStrPrdMetaDesc%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Keywords: </td>
							<td><textarea name="PrdMetaKeywords" cols="50" rows="4" tabindex="1003"><%=pStrPrdMetaKeywords%></textarea>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
					</table>
					
				</div>
			<%
			'// =========================================
			'// FIRST PANEL - END
			'// =========================================
			
			'// =========================================
			'// SECOND PANEL - START - Prices
			'// =========================================
			%>
				<div id="tab-2" class="tab-pane">
				
					<table class="pcCPcontent">
					
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Product Prices</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<% if pcv_ProductType="std" then %>
							<td width="30%">Online Price:</td>
							<% else %>
							<td>Base Price:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=500"></a></td>
							<% end if %>
							<td width="70%"><%=scCurSign%>&nbsp;<input type="text" name="price" value="<%=money(pPrice)%>" size="10" tabindex="201"><%if pcSCID>0 then%><%if pcTPrice="0" then%>&nbsp;<a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="../pc/catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>" style="vertical-align: middle"></a>&nbsp;(Original Price: <%=scCurSign & money(pcBUPrice)%>)<%end if%><%end if%></td>
						</tr>
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr> 
							<td>List Price:</td>
							<td><%=scCurSign%>&nbsp;<input type="text" name="listPrice" value="<%=money(pListPrice)%>" size="10" tabindex="202"></td>
						</tr>
						<tr> 
							<td>Show savings:</td>
							<td>Yes 
							<% If pListhidden="-1" then %>
								<input type="checkbox" name="listhidden" value="-1" checked  class="clearBorder" tabindex="203">
							<% else %>
								<input type="checkbox" name="listhidden" value="-1" class="clearBorder" tabindex="203">
							<% end if %>
							</td>
						</tr>
						<% end if ' Hide if it's a Configurable-Only Item
						
						'// START CT ADD
						'// if there are PBP customer type categories - List them here
						query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories;"
						SET rs=Server.CreateObject("ADODB.RecordSet")
						SET rs=conntemp.execute(query)
						if NOT rs.eof then 
							do until rs.eof 
								intIdcustomerCategory=rs("idcustomerCategory")
								strpcCC_Name=rs("pcCC_Name")
								strpcCC_CategoryType=rs("pcCC_CategoryType")
								intpcCC_ATBPercentage=rs("pcCC_ATB_Percentage")
								intpcCC_ATB_Off=rs("pcCC_ATB_Off")
								if intpcCC_ATB_Off="Retail" then
									intpcCC_ATBPercentOff=0
								else
									intpcCC_ATBPercentOff=1
								end if	
								
								query="SELECT pcCC_Pricing.idcustomerCategory, pcCC_Pricing.idProduct, pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE (((pcCC_Pricing.idcustomerCategory)="&intIdcustomerCategory&") AND ((pcCC_Pricing.idProduct)="&pIdProduct&"));"
								SET rsPriceObj=server.CreateObject("ADODB.RecordSet")
								SET rsPriceObj=conntemp.execute(query)
								if rsPriceObj.eof then
									dblpcCC_Price=0
								else
									dblpcCC_Price=rsPriceObj("pcCC_Price")
									dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
								end if
								SET rsPriceObj=nothing
								%>
								<tr valign="top">
									<td><%=strpcCC_Name%></td>
									<td><%=scCurSign%>&nbsp;<input type="text" name="pcCC_<%=intIdcustomerCategory%>" value="<%=money(dblpcCC_Price)%>" size="10">
									<%if pcSCID>0 then%><%if Clng(pcTPrice)=Clng(intIdcustomerCategory) then%>&nbsp;<a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="../pc/catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>" style="vertical-align: middle"></a>&nbsp;(Original Price: <%=scCurSign & money(pcBUPrice)%>)<br><%end if%><%end if%>
									<%
									' Find out if there is a wholesale price
									if (pBtoBPrice>"0") then
										tempPrice=pBtoBPrice
									else
										tempPrice=pPrice
									end if
									' Calculate the "across the board" price
									if strpcCC_CategoryType="ATB" then
										if intpcCC_ATBPercentOff=0 then
											ATBPrice=pPrice-(pcf_Round(pPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
										else
											ATBPrice=tempPrice-(pcf_Round(tempPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
										end if					
									%>
									Default price for this pricing category: <% response.write(scCurSign & money(ATBPrice))%>&nbsp;&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=308"></a>
									<%
									end if
									%>
									</td>
								</tr>
							<% rs.moveNext
							loop
						end if
						SET rs=nothing
						'// END CT ADD 
						%>
						<tr> 
							<td>Wholesale Price:</td>
							<td><%=scCurSign%>&nbsp;<input type="text" name="bToBprice" value="<%response.write money(pBtoBPrice)%>" size="10" tabindex="204"><%if pcSCID>0 then%><%if pcTPrice="-1" then%>&nbsp;<a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="../pc/catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>" style="vertical-align: middle"></a>&nbsp;(Original Price: <%=scCurSign & money(pcBUPrice)%>)<%end if%><%end if%></td>
						</tr>
						<%'Start SDBA%>
						<tr> 
							<td>Cost:</td>
							<td><%=scCurSign%>&nbsp;<input type="text" name="cost" value="<%=money(pCost)%>" size="10" tabindex="205"></td>
						</tr>
						<%'End SDBA%>
						
						<% if pcv_ProductType="bto" then %>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<tr> 
							<td>Hide Default Price:</td>
							<td>Yes <input type="checkbox" name="hidebtoprice" value="1" <%if pcv_intHideBTOPrice=1 then%>checked<%end if%>>&nbsp;<font color="#666666">When the default price is very small, use this option to hide it</font></td>
						</tr>
						<tr> 
							<td>Hide default configuration:</td>
							<td>Yes <input type="checkbox" name="hidedefconfig" value="1" <%if cint(phideDefConfig)=1 then%>checked<%end if%> class="clearBorder"></td>
						</tr>
						<tr> 
							<td valign="bottom">Skip Product Details Page:</td>
							<td>Yes <input type="checkbox" name="pcv_intSkipDetailsPage" value="1" <%if pcv_intSkipDetailsPage="1" then%>checked<%end if%> class="clearBorder">
							</td>
						</tr>
						<tr>
							<td valign="top">Disallow purchasing<br />(quoting only):</td>
							<td>
							<input type="radio" name="noprices" value="0" <%if (pnoprices="") or (pnoprices="0") then%>checked<%end if%> class="clearBorder">No&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="noprices" value="1" <%if cint(pnoprices)=1 then%>checked<%end if%> class="clearBorder">Yes - Show Prices&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="noprices" value="2" <%if cint(pnoprices)=2 then%>checked<%end if%> class="clearBorder">Yes - Hide Prices
						</td>
						</tr>
						<% If statusCM="0" OR scCM=1 Then %>
						<tr> 
							<td>Show Configurator+ conflict alert messages<br />(Conflict Management):</td>
							<td>Yes <input type="checkbox" name="showBtoCmMsg" value="1" class="clearBorder" <%if pcv_showBtoCmMsg="1" then%>checked<%end if%>></td>
						</tr>
						<% End If %>
						<tr>
							<td valign="top">Maximum number of selections:</td>
							<td>
							<input type="text" size="5" name="maxselect" value="<%=pMaxSelect%>"><br>
							<i>(The number of total items selected on the product configuration page)</i>
							</td>
						</tr>
						<% end if %>
						
					</table>
					
				</div>
			<%
			'// =========================================
			'// SECOND PANEL - END
			'// =========================================

			'// =========================================
			'// THIRD PANEL - START - Categories
			'// =========================================
			%>
				<div id="tab-3" class="tab-pane">
				
					<table class="pcCPcontent">				
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Categories</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td colspan="2">
								<%'Begin Categories%>                  
								<table class="pcCPcontent">
									
									<%
									query="SELECT idProduct, idCategory, POrder FROM categories_products WHERE idproduct="&pIdProduct&" ORDER BY POrder"
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=conntemp.execute(query)
									
									if err.number <> 0 then
										set rs=nothing
										
										call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modifyProduct.asp loading product categories") 
									end if
					
									if rs.eof then 
										intTotCat=0 %>
										<tr>
											<td colspan="2">This product is not assigned to any categories.</td>
										</tr>
									<% else
										intTotCat=1 %>
										<tr>
											<td><strong>Category</strong></td>
											<td><strong>Parent</strong></td>
										</tr>
										<% dim parent
										
										do until rs.eof
											tempidCategory=rs("idCategory")
											query="SELECT idCategory, idparentCategory, categoryDesc FROM categories WHERE idCategory="&tempidCategory&";"
											set rstemp=server.CreateObject("ADODB.RecordSet")
											set rstemp=conntemp.execute(query)
											if not rstemp.eof then
												idparentCategory=rstemp("idparentCategory")
												categoryDesc=rstemp("categoryDesc")
												set rstemp=nothing
												if idparentCategory=1 then %>
													<tr>
														<td>
															<a href="modcata.asp?idcategory=<%=tempidCategory%>" target="_blank" style="text-decoration: none;"><%=categoryDesc%></a>
														</td>
														<td>&nbsp;</td>
													</tr>
												<% else 
													parent=""
													query="SELECT idCategory, idparentCategory, categoryDesc FROM categories WHERE idCategory="&idparentCategory&";"
													set rsTemp=server.CreateObject("ADODB.RecordSet")
													set rsTemp=conntemp.execute(query)
													idparentCategory=rsTemp("idparentCategory")
													parent=rsTemp("categoryDesc")
													set rsTemp=nothing
													if idparentCategory<>1 then
														Call GetParent()
													end if
													%>
													<tr>
														<td>
															<a href="modcata.asp?idcategory=<%=tempidCategory%>" target="_blank" style="text-decoration: none;"><%=categoryDesc%></a>
														</td>
														<td><%=parent%></td>
													</tr>
												<% end if
											end if
											rs.movenext
										loop 
										set rs=nothing
									end if
									
									function GetParent() 
										query="SELECT idparentCategory, categoryDesc FROM categories WHERE idCategory=" & idparentCategory
										set rsTemp=server.CreateObject("ADODB.RecordSet")
										set rsTemp=conntemp.execute(query)
										idparentCategory=rsTemp("idparentCategory")
										parent=parent & "/" & rsTemp("categoryDesc")
										set rsTemp=nothing
										If idparentCategory<>1 then
											call GetParent() 
										end if
									End function %>
                                    <tr>
                                    	<td colspan="2" class="pcCPspacer"></td>
                                    </tr>
									<tr>
										<td colspan="2" >
										<input type="button" class="btn btn-default"  name="Update" value="Edit Category Assignment" onClick="newWindow2('cat_popup.asp?intTotCat=<%=intTotCat%>&idproduct=<%=pidproduct%>','window2')">
										&nbsp;</td>
									</tr>
								</table>
								<%'End Categories%>
							</td>
						</tr>
					</table>
			
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================

			'// =========================================
			'// FOURTH PANEL - START - Product images
			'// =========================================
			%>
				<div id="tab-4" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Product Images</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">Type in the file name, not the file path. All images must be located in the 'pc/catalog' folder.
							<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
							<%If HaveImgUplResizeObjs=1 then%>
								To upload and resize an image <a href="javascript:;" onClick="pcCPWindow('uploadresize/productResizea.asp', 450, 450); return false;">click here</a>.
							<% Else %>
								To upload an image <a href="javascript:;" onClick="pcCPWindow('imageuploada_popup.asp', 400, 360)">click here</a>.
							<% End If %>
							</td>
						</tr>
						<tr>
							<td>Thumbnail Image:</td>
							<td valign=bottom>  
									<input type="text" name="smallImageUrl" value="<%response.write pSmallImageUrl%>" size="30" tabindex="401"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=smallImageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>  
									<% if pSmallImageUrl <> "" then %>
											<a href="javascript:enlrge('../pc/catalog/<%=pSmallImageUrl%>')">
													<img src="../pc/catalog/<%=pSmallImageUrl%>" border=0 align=absbottom class="pcShowProductImageM">
											</a>
									<% else %>
											<img src="../pc/catalog/no_image.gif" border=0 align=absbottom class="pcShowProductImageM">
									<% end if %>
							</td>
						</tr>
						<tr> 
							<td>General Image:</td>
							<td valign=bottom>  
								<input type="text" name="imageUrl" value="<%response.write pImageUrl%>" size="30" tabindex="402"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=imageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
									<% if pImageUrl <> "" then %>
											<a href="javascript:enlrge('../pc/catalog/<%=pImageUrl%>')">
													<img src="../pc/catalog/<%=pImageUrl%>"  border=0 align=absbottom class="pcShowProductImageM">
											</a>
									<% else %>
											<img src="../pc/catalog/no_image.gif"  border=0 align=absbottom class="pcShowProductImageM">
									<% end if %>
							</td>
						</tr>
						<tr> 
							<td>Detail View Image:</td>
							<td valign=bottom>  
								<input type="text" name="largeImageUrl" value="<%response.write pLargeImageUrl%>" size="30" tabindex="403"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=largeImageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
									<% if pLargeImageUrl <> "" then %>
											<a href="javascript:enlrge('../pc/catalog/<%=pLargeImageUrl%>')">		    
													<img src="../pc/catalog/<%=pLargeImageUrl%>"  border=0 align=absbottom class="pcShowProductImageM">
											</a>
									<% else %>
											<img src="../pc/catalog/no_image.gif"  border=0 align=absbottom class="pcShowProductImageM">
									<% end if %>
							</td>
						</tr>
                        <tr> 
							<td>Alt Tag Text (optional) :</td>
							<td>
								<input type="text" name="altTagText" value="<%response.write pAltTagText%>" size="30" tabindex="403">
							</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
                        <tr>
							<td>Enable Image Magnifier:</td>
							<td>
								<input type="checkbox" name="MojoZoom" value="1" <%if cint(pcv_IntMojoZoom)=1 then%>checked<%end if%> class="clearBorder" tabindex="404">
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=467"></a>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr>
							<th colspan="2">Additional Product Views - <a href="javascript:;" onClick="javascript:newAddWindow('addImg_popup.asp?idproduct=<%=pidproduct%>','addwindow')">Add New</a></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
                        <tr>
							<td>Hide Additional Images:</td>
							<td>
								<input type="checkbox" name="AdditionalImages" value="1" <%if cint(pcv_IntAdditionalImages)=1 then%>checked<%end if%> class="clearBorder" tabindex="404">
							</td>
						</tr>
						<tr>
							<td colspan="2">
									<!--#include file="modPrdAddImg.asp"-->
							</td>
						</tr>
						<% end if ' Hide if it's a Configurable-Only Item %>
						
					</table>
					
				</div>
			<%
			'// =========================================
			'// FOURTH PANEL - END
			'// =========================================
			%>
							
			<%
			'// =========================================			
			'// FIFTH PANEL - START - Product settings
			'// =========================================
			%>
			
				<div id="tab-5" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<th colspan="2">Product Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item
						
						'// Brands - Start
						query="Select IDBrand, BrandName from Brands order by BrandName asc"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if not rs.eof then %>
						<tr> 
							<td>Brand:</td>
							<td>
								<select name="IDBrand" tabindex="701">
									<option value="0" selected></option>
									<% do while not rs.eof
										intIDBrand=rs("IDBrand")
										strBrandName=rs("BrandName") %>
										<option value="<%=intIDBrand%>" <%if pIDBrand & ""=intIDBrand & "" then%>selected<%end if%>><%=strBrandName%></option>
										<%
										rs.MoveNext
									loop
									set rs=nothing
									%>
								</select>
							</td>
						</tr>
						<%
							else
							set rs=nothing
						%>
						<tr>
							<td colspan="2">
								<input type="hidden" name="IDBrand" value="<%=pIDBrand%>">
						</td>
						</tr>
						<% end if
						'// Brands - End
						end if ' Hide if it's a Configurable-Only Item %>
						
						<tr> 
							<td>Active:</td>
							<td>Yes 
								<input type="checkbox" name="active" value="-1" <%
								if pactive=-1 then
									response.write "checked"
								end if
								%> 
								 class="clearBorder" tabindex="702">
							</td>
						</tr>
						
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr> 
							<td>Special:</td>
							<td>Yes&nbsp;<input type="checkbox" name="hotDeal" value="-1" <%
							if photDeal=-1 then
								response.write "checked"
							end if
							%>
							 class="clearBorder" tabindex="703"></td>
						</tr>
						<tr> 
							<td>Featured Product:</td>
							<td>Yes&nbsp;<input type="checkbox" name="showInHome" value="-1" <%
							if pShowInHome=-1 then
								response.write "checked"
							end if
							%> class="clearBorder" tabindex="704">
							</td>
						</tr>
						<% end if ' Hide if it's a Configurable-Only Item

						'RP ADDON-S
						If RewardsActive <> 0 Then %>
							<tr> 
								<td><%=RewardsLabel%>:</td>
								<td><input type="text" name="iRewardPoints" value=<%=iRewardPoints%> size="20" tabindex="705"></td>
							</tr>
						<% 
						End If 
						'RP ADDON-E
						%>
						<tr> 
							<td>Non-taxable:</td>
							<td>Yes <input type="checkbox" name="notax" value="-1" <% if pnotax="-1" then%>checked<%end if %> class="clearBorder" tabindex="707"></td>
						</tr>
                        <% If ptaxAvalara = 1 Then %>
                        <tr> 
							<td>Avalara Tax Code:</td>
							<td><input type="text" name="AvalaraTaxCode" value="<%=pAvalaraTaxCode%>"></td>
						</tr>
                        <% End If %>
						<%if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item%>
						<tr>
							<td nowrap><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_96")%>:</td>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="checkbox" name="hideSKU" value="1" class="clearBorder" tabindex="709" <%if pHideSKU="1" then%>checked<%end if%>></td>
						</tr>
                        <!--
                        <tr>
							<td>Exempt from ALL Discount/Coupon codes:</td>
							<td>Yes <input type="checkbox" name="exemptDisc" value="1" <% if pExemptDisc=1 then%>checked<%end if %> class="clearBorder" tabindex="709"></td>
						</tr>
                        -->
						<tr>
							<td>Not for Sale:</td>
							<td>Yes <input type="checkbox" name="formQuantity" value="-1" <% if pFormQuantity="-1" then%>checked<%end if %> class="clearBorder" tabindex="710"></td>
						</tr>
						<tr> 
							<td valign="top">Not for Sale Message:<br /><span class="pcSmallText">(e.g. &quot;Coming Soon&quot;)</span></td>
							<td>
								<textarea name="emailText" rows="4" cols="40" tabindex="711" onKeyUp="javascript:testchars(this,'1',250); javascript:document.getElementById('emailTextCounter').style.display='';"><%=pEmailText%></textarea>
                <div id="emailTextCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar1" name="countchar1" style="font-weight: bold"><%=maxlength%></span> characters left.</div>
							</td>
						</tr>
						<% end if ' Hide if it's a Configurable-Only Item %>
						<tr>
							<td valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_87")%></td>
							<td><textarea name="prdnotes" rows="6" cols="60" tabindex="712"><%=pcv_prdnotes%></textarea></td>
						</tr>
					</table>
	
				</div>
			<%
			'// =========================================			
			'// FIFTH PANEL - END - Product settings
			'// =========================================
			%>

			<%
			'// =========================================			
			'// SIXTH PANEL - START - Inventory settings
			'// =========================================
			%>
				<div id="tab-6" class="tab-pane">
				
					<table class="pcCPcontent">

						<%'Start SDBA%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Inventory Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<script type=text/javascript>
							$pc(document).ready(function() {
								$pc("#noStock").click(function(e) {
									if ($pc(this).is(":checked")) {
										$pc("#stockOptions").hide();
									} else {
										$pc("#stockOptions").show();
									}
								});
							});	
						</script>
						<tr>
							<td width="30%">Disregard Stock:</td>
							<td><input type="checkbox" name="noStock" id="noStock" value="-1" <% if pNoStock<>0 then response.write "checked" %> class="clearBorder" tabindex="601"></td>
						</tr>
						<tr> 
							<td>Stock:</td>
							<td>  
								<%'SHW-S
								if SHWSync=1 then%><b><%=pStock%></b>&nbsp;&nbsp;&nbsp;<font color=blue><i>(Has been synchronized with SHIPWIRE)</i></font>
									<input type="hidden" name="stock" value="<%=pStock%>">
								<%else%>
									<input type="text" name="stock" id="stock" value="<%=pStock%>" size="4" tabindex="602">
								<%end if
								'SHW-E%>
								<%'BackInStock-S%>
								<input type="hidden" name="nmstock" id="nmstock" value="<%=pStock%>">
								<%'BackInStock-E%>
								<input type="hidden" name="deliveringTime" value="<%response.write pDeliveringTime%>"> 
							</td>
						</tr>
						<%
							pcv_strDisplayStyle = ""
							If pNoStock <> 0 Then
								pcv_strDisplayStyle	= "style='display: none'"
							End If
						%>
						<tbody id="stockOptions" <%= pcv_strDisplayStyle %>>
							<tr> 
								<td valign="top">Allow back-ordering:</td>
								<td>
									<input type="radio" name="pcbackorder" value="1" <%if pcbackorder="1" then%>checked<%end if%> class="clearBorder" tabindex="603"> Yes 
									&nbsp;<input type="radio" name="pcbackorder" value="0" <%if pcbackorder<>"1" then%>checked<%end if%> class="clearBorder" tabindex="603"> No<br>
									When back-ordered, typically ships within <input type="text" size="5" value="<%=pcShipNDays%>" name="pcShipNDays" tabindex="604"> days </td>
							</tr>
							<tr> 
								<td valign="top">Low inventory notification:</td>
								<td><input type="radio" name="pcnotifystock" value="1" <%if pcnotifystock="1" then%>checked<%end if%> class="clearBorder" tabindex="605"> Yes 
									&nbsp;<input type="radio" name="pcnotifystock" value="0" <%if pcnotifystock="0" then%>checked<%end if%> class="clearBorder" tabindex="605"> No 
									<br />
									<span class="pcSmallText"><i>(Store admin is notified when inventory drops below the Reorder Level)</i></span></td>
							</tr>
							<tr> 
								<td>Reorder Level:</td>
								<td>
								<input name="pcreorderlevel" type="text" value="<%=pcreorderlevel%>" size="5" maxlength="10" tabindex="606"></td>
							</tr>
						</tbody>
						<tr> 
							<td>Minimum Quantity to Buy:</td>
							<td><input name="minimumqty" type="text" value="<%=pcv_lngMinimumQty%>" size="5" maxlength="10" tabindex="607">
								&nbsp;&nbsp;&nbsp;&nbsp;          
									<input type="checkbox" name="qtyvalidate" value="1" <%if pcv_intQtyValidate=1 then%>checked<%end if%> class="clearBorder" tabindex="608"> Force purchase of multiples of:&nbsp;<input name="multiQty" type="text" value="<%=pcv_multiQty%>" size="5" maxlength="10" tabindex="609">
							</td>
						</tr>
						<%'End SDBA%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Weight Settings</th>
						</tr>
						
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<%
						'// WEIGHTS - Start
						if scShipFromWeightUnit="KGS" then %>
						<tr> 
							<td width="30%">Weight:</td>
							<td width="70%">
								<input type="text" name="weight_kg" value="<%=pWeight_kg%>" size="4" tabindex="620"> kg 
								<input type="text" name="weight_g" value="<%=pWeight_g%>" size="4" tabindex="621"> g
							</td>
						</tr>
						<tr>
							<td colspan="2">If this product weighs less than one gram, use the field below to specify how many units of this product it takes to weigh 1 KG. For more information, see the User Guide.</td>
						</tr>
						<tr>
							<td>Units to make 1 KG:</td>
							<td><input name="QtyToPound" type="text" value="<%=pcv_QtyToPound%>" size="10" maxlength="10" tabindex="622"></td>
						</tr>
						<% else %>
						<tr> 
							<td width="30%">Weight:</td>
							<td width="70%">
								<input type="text" name="weight" value="<%=pWeight%>" size="4" tabindex="620"> lbs. 
								<input type="text" name="weight_oz" value="<%=pWeight_oz%>" size="4" tabindex="621"> ozs.</td>
						</tr>
						<tr>
							<td colspan="2">If this product weighs less than one ounce, use the field below to specify how many units of this product it takes to weigh 1 pound. For more information, see the User Guide.</td>
						</tr>
						<tr>
							<td>Units to make 1 lb:</td>
							<td><input name="QtyToPound" type="text" value="<%=pcv_QtyToPound%>" size="10" maxlength="10" tabindex="622"></td>
						</tr>
						<% end if
						'// WEIGHTS - End
						%>

						<%
						if pOverSizeSpec<>"NO" then
							pOSArray=split(pOverSizeSpec,"||")
							if ubound(pOSArray)>2 then
								tOS_width=pOSArray(0)
								tOS_height=pOSArray(1)
								tOS_length=pOSArray(2)
							else
								tOS_width=0
								tOS_height=0
								tOS_length=0
							end if
						end if
						%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Shipping Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<%
						if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item
						%>
						<% 
							pcv_noShippingStyle = "display: none"
							if pnoshipping="-1" then 
								pcv_noShippingStyle = ""
							end if
						%>
						<tr> 
							<td>Non-Shipping Item:</td>
							<td> 
								<input type="checkbox" name="noshipping" id="noShipping" value="-1" class="clearBorder" tabindex="630" <% if pnoshipping="-1" then response.write "checked" %>>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=449"></a>
							</td>
						</tr>
						<tbody id="noShippingSettings" style="<%= pcv_noShippingStyle %>">
							<tr>
								<td>Display Non-Shipping Text:</td>
								<td>
									<input type="checkbox" name="noshippingtext" value="-1" class="clearBorder" tabindex="631" <% if pnoshippingtext="-1" then response.write "checked" %>>
								</td>
							</tr>
						</tbody>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<%

						end if ' Hide if it's a Configurable-Only Item
						%>

						<tr> 
							<td colspan="2"><strong>Oversized</strong> products shipped via <strong>UPS, USPS, or FedEx</strong></td>
						</tr>
						<tr>
							<td colspan="2">This product will be shipped as an oversized product.
								<input name="OverSizeSpec" type="radio" value="YES" <% if pOverSizeSpec<>"NO" then Response.Write "checked" %> class="clearBorder" tabindex="632">&nbsp;Yes
								<input name="OverSizeSpec" type="radio" value="NO" <% if pOverSizeSpec="NO" then Response.Write "checked" %> class="clearBorder" tabindex="632">&nbsp;No
								<br>
								If 'Yes', set the size below in inches. NOTE: Oversized products will always be shipped separately.
							</td>
						</tr>
						<tr> 
							<td colspan="2">
								<table class="pcCPcontent">
									<tr> 
										<td>Length:</td>
										<td width="15%"> 
											<input name="os_length" type="text" id="os_length" size="3" maxlength="3" value="<%=tOS_length%>" tabindex="633">
										</td>
										<td rowspan="3" align="left" valign="top">
												<table width="100%" border="0" cellpadding="6" cellspacing="0">
													<tr>
														<td>
															Notes about shipping oversized packages with UPS, USPS, or FedEx:
															<ul>
																<li><strong>Length</strong> should always be the longest side.</li>
																<li><strong>Girth</strong> is defined as (width * 2) + (height * 2).</li>
															</ul>
														</td>
													</tr>
													<tr>
														<td><strong>Please refer to the <a href="http://wiki.productcart.com/productcart/shipping-oversized_items" target="_blank">Wiki</a> or the shipping provider's documentation for more information on oversized packages.</strong></td>
													</tr>
												</table>
                                       	</td>
                                    </tr>
                                    <tr> 
                                        <td>Width:</td>
                                        <td width="15%"> 
                                            <input name="os_width" type="text" id="os_width" size="3" maxlength="3" value="<%=tOS_width%>" tabindex="634"></td>
                                    </tr>
                                    <tr> 
														<td width="11%">Height:</td>
														<td width="15%"> 
															<input name="os_height" type="text" id="os_height" size="3" maxlength="3" value="<%=tOS_height%>" tabindex="635">
														</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<%
						if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item

						'Start SDBA
						'Get Suppliers List
						query="Select pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers order by pcSupplier_Company asc"
						set rs=connTemp.execute(query)
						if not rs.eof then
							pcArray=rs.getRows()
							intCount=ubound(pcArray,2)
							%>
						<tr>
							<td>Supplier:</td>
							<td>
							<select name="pcIDSupplier" onChange="javascript:TestDropShipper();" tabindex="636">
							<option value="0" selected></option>
							<%For i=0 to intCount%>
								<option value="<%=pcArray(0,i)%>" <%if clng(pcIDSupplier)=clng(pcArray(0,i)) then%>selected<%end if%>><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
							<%Next%>
							</select>
							</td>
						</tr>
						<%else%>
							<input type=hidden name="pcIDSupplier" value="0">
						<%end if
						set rs=nothing

						'Get Drop-Shippers List
						query="SELECT pcDropShipper_ID,pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName,0 FROM pcDropShippers UNION (SELECT pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName,1 FROM pcSuppliers WHERE pcSupplier_IsDropShipper=1) ORDER BY pcDropShipper_Company ASC"
						set rs=connTemp.execute(query)
						dim pcv_ShowDSFields
						pcv_ShowDSFields=0
						if not rs.eof then
							pcv_ShowDSFields=1
						'// Allow selection only if drop-shippers exist
						%>
						
						<tr>
							<td>This product is drop-shipped:</td>
							<td> 
								<input type="radio" name="pcIsdropshipped" value="1" <%if pcIsdropshipped="1" then%>checked<%end if%> class="clearBorder" onClick="javascript:TurnOnDropShipper();" tabindex="637"> Yes 
								&nbsp;<input type="radio" name="pcIsdropshipped" value="0" <%if pcIsdropshipped<>"1" then%>checked<%end if%> class="clearBorder" onClick="javascript:TurnOffDropShipper();" tabindex="637"> No
							</td>
						</tr>
						
						<%
						'// Get list of drop-shippers
						
							pcArray=rs.getRows()
							intCount=ubound(pcArray,2)
							set rs=nothing
							
							'Drop-Shipper is also a Supplier or not
							query="SELECT pcDS_ID FROM pcDropShippersSuppliers WHERE idproduct=" & pIdProduct & " AND pcDS_IsDropShipper=1;"
							set rs=connTemp.execute(query)
							pcDropShipperSupplier=0
							if not rs.eof then
								pcDropShipperSupplier=1
							end if
							set rs=nothing
							%>
						<tr>
							<td>Drop-Shipper:</td>
							<td>
							<select name="pcIDDropShipper" onChange="javascript:TestSupplier()" tabindex="638">
							<option value="0" selected></option>
							<%For i=0 to intCount%>
								<option value="<%=pcArray(0,i)%>_<%=pcArray(4,i)%>" <%if (clng(pcIDDropShipper)=clng(pcArray(0,i))) AND (clng(pcArray(4,i))=pcDropShipperSupplier) then%>selected<%end if%>><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
							<%Next%>
							</select>
							</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<%else%>
						<tr> 
							<td colspan="2">
								<input type="hidden" name="pcIDDropShipper" value="0">
							</td>
						</tr>
						<%end if
						set rs=nothing%>
						<script type=text/javascript>
							function TestDropShipper()
							{
								var tmp1=document.hForm.pcIDSupplier.value;
								try
								{
									var j=document.hForm.pcIDDropShipper.length;
									var i=0;
									var test=0;
									do
									{
										i=j-1;
										if (tmp1 + "_1" == document.hForm.pcIDDropShipper.options[i].value)
										{
											document.hForm.pcIDDropShipper.options[i].selected=true;
											document.hForm.pcIDDropShipper.disabled=true;
											document.hForm.pcIsdropshipped[0].checked=true;
											test=1;
											break;
										}
									}
									while (--j);
									if (test==0)
									{
										if (document.hForm.pcIsdropshipped[0].checked==true)
										{
											document.hForm.pcIDDropShipper.disabled=false;
										}
										var tmp1=document.hForm.pcIDDropShipper.value;
										var tmp2=tmp1.split("_");
										if (tmp2[1]==1)
										{
											document.hForm.pcIDDropShipper.options[0].selected=true;
										}
									}
								}
								catch(err)
								{
									return(true);
								}
							}
							function TestSupplier()
							{
								var tmp1=document.hForm.pcIDDropShipper.value;
								var tmp2=tmp1.split("_");
								try
								{
									var test=0;
									if (tmp2[1]=="1")
									{
										var j=document.hForm.pcIDSupplier.length;
										var i=0;
									
										do
										{
											i=j-1;
											if (tmp2[0] == document.hForm.pcIDSupplier.options[i].value)
											{
												document.hForm.pcIDSupplier.options[i].selected=true;
												document.hForm.pcIDSupplier.disabled=true;
												test=1;
												break;
											}
										}
										while (--j);
									}
									if (test==0)
									{
										if (document.hForm.pcIDSupplier.disabled==true)
										{
											document.hForm.pcIDSupplier.disabled=false;
											document.hForm.pcIDSupplier.options[0].selected=true;
										}
									}
								}
								catch(err)
								{
									return(true);
								}
					
							}
						
							function TurnOnDropShipper()
							{
								try
								{
									document.hForm.pcIDDropShipper.disabled=false;
									document.hForm.pcIDSupplier.disabled=false;
								}
								catch(err)
								{
									//Do nothing
								}
							
							}
						
							function TurnOffDropShipper()
							{
								try
								{
									document.hForm.pcIDDropShipper.disabled=true;
									document.hForm.pcIDSupplier.disabled=false;
									var tmp1=document.hForm.pcIDDropShipper.value;
									if (tmp1!="0")
									{
										var tmp2=tmp1.split("_");
										if (tmp2[1]=="1")
										{
											document.hForm.pcIDSupplier.options[0].selected=true;
										}
									}
									document.hForm.pcIDDropShipper.options[0].selected=true;
								}
								catch(err)
								{
									//Do nothing
								}
							
							}
							<% if pcv_ShowDSFields=1 then %>
							TestDropShipper();
							if (document.hForm.pcIsdropshipped[1].checked==true) TurnOffDropShipper();
							<% end if %>
						</script>
						<%
						'End SDBA

						end if ' Hide if it's a Configurable-Only Item
						%>

						<tr> 
							<td colspan="2"><strong>Shipping Surcharge</strong>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=463"></a></td>
						</tr>
						<tr>
							<td>First Unit Surcharge:</td>
							<td><input name="surcharge1" type="text" id="surcharge1" value="<%=money(pcv_Surcharge1)%>" size="10" maxlength="10" tabindex="639"></td>
						</tr>
						<tr>
							<td>Additional Unit(s) Surcharge:</td>
							<td><input name="surcharge2" type="text" id="surcharge2" value="<%=money(pcv_Surcharge2)%>" size="10" maxlength="10" tabindex="640"></td>
						</tr>
					</table>
	
				</div>
			<%
			'// =========================================
			'// SIXTH PANEL - END
			'// =========================================
			%>

			<%
			'// =========================================			
			'// SEVENTH PANEL - START - Product Page Layout
			'// =========================================
			%>
				<%
				if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item
				%>
				<div id="tab-7" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Product Page Layout</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">
								Choose a product details page layout below. 
							</td>
						</tr>
							<tr> 
								<td>Page Layout:</td>
								<td>
										<select name="displayLayout" id="displayLayout" tabindex="706">
											<option value="computer" <%if pDisplayLayout="computer" then%>selected<%end if%>>Standard Computer</option>
											<option value="charterpc" <%if pDisplayLayout="charterpc" then%>selected<%end if%>>Charter PC</option>
											<option value="traderpc" <%if pDisplayLayout="traderpc" then%>selected<%end if%>>Trader PC</option>
											<option value="traderpropc" <%if pDisplayLayout="traderpropc" then%>selected<%end if%>>Trader Pro PC</option>
											<option value="stand" <%if pDisplayLayout="stand" then%>selected<%end if%>>Stand</option>
											<option value="monitor" <%if pDisplayLayout="monitor" then%>selected<%end if%>>Monitor</option>
                                            <option value="custom" <%if pDisplayLayout="custom" then%>selected<%end if%>>Custom</option>
										</select>
								</td>
							</tr>
							<%
								pcv_strDisplayStyle = ""
								pcv_strCustPrdDisplayStyle = ""
								pcv_strCustTabDisplayStyle = ""
								If Left(pDisplayLayout, 1) = "t" Then
									pcv_strDisplayStyle = "display: none"
								Else
									pcv_strCustPrdDisplayStyle = "style='display: none'"
									pcv_strCustTabDisplayStyle = "style='display: none'"
								End If
							%>
							<tr>
								<td></td>
								<td>
									<ul id="customizeButtons" style="list-style-type: none; padding: 0px; margin: 0px; <%= pcv_strDisplayStyle %>">
										<li style="padding: 5px"><a href="#" id="customizeLayout"><span class="glyphicon glyphicon-cog"></span>&nbsp;Customize this Layout</a></li>
										<li style="padding: 5px"><a href="#" id="addTabsToLayout"><span class="glyphicon glyphicon-cog"></span>&nbsp;Customize this Layout w/ Tabs</a></li>
									</ul>
								</td>
							</tr>
							<!--#include file="inc_CustomPrdPage.asp"-->

					</table>
				
				</div>
				<%end if%>
			<%
			'// =========================================			
			'// SEVENTH PANEL - END - Product Page Layout
			'// =========================================
			%>

			
			<%
			'// EIGHTH PANEL - START - Downloadable product
			if pcv_ProductType<>"item" then	 ' Hide for Configurable-Only Items
			%>
				<div id="tab-8" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<th colspan="2">Downloadable Product Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer">
							<input type=hidden name="downloadable1" value="<%=pDownloadable%>">
							<input type=hidden name="urlexpire1" value="<%=pURLExpire%>">
							<input type=hidden name="license1" value="<%=pLicense%>">
							</td>
						</tr>
						<tr> 
							<td colspan="2">This is a downloadable product&nbsp; 
							<% If statusAPP="1" OR scAPP=1 Then %>
								<input name="downloadable" type="radio" value="1" <%if pdownloadable="1" then%>checked<%end if%> onClick="<% if pcv_ProductType="std" then %>document.hForm.apparel[1].checked='true'; document.hForm.GC[1].checked='true'; <% end if %>document.hForm.downloadable1.value='1'; document.getElementById('show_19').style.display='';<% if pcv_ProductType="std" then %> document.getElementById('show_20').style.display='none'; document.getElementById('show_21').style.display='none'<% end if %>" class="clearBorder" tabindex="801">Yes 
							<% Else %>
								<input name="downloadable" type="radio" value="1" <%if pdownloadable="1" then%>checked<%end if%> onClick="<% if pcv_ProductType="std" then %>document.hForm.GC[1].checked='true'; <% end if %>document.hForm.downloadable1.value='1'; document.getElementById('show_19').style.display='';<% if pcv_ProductType="std" then %> document.getElementById('show_20').style.display='none'<% end if %>" class="clearBorder" tabindex="801">Yes 
							<% End If %>
							<input name="downloadable" type="radio" value="0" <%if (pdownloadable="0") or (pdownloadable="") then%>checked<%end if%> onClick="document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none';" class="clearBorder" tabindex="802">No</td>
						</tr>
						<tr>
							<td align="center" colspan="2">                     
								<table id="show_19" <%if (pdownloadable="0") or (pdownloadable="") then%>style="display:none;"<%end if%> class="pcCPcontent">
									<tr>
										<td colspan="2"><p>Downloadable file location. You have <u>two options</u>:</p>
										<ul>
										<li><u>Enter the full physical path to the file</u> (e.g. 
										M:\myAccount\downloads\downloadfile.zip). This option uses the <em>Hide URL </em>feature. On Web servers running IIS 6 or above, this feature only works with files that are less than 4 MB in size. For more information and a technical explanation of this feature, please see the User Guide. <br>
										<img src="images/spacer.gif" height="15" width="1">Current physical path of the root directory: <%=Server.MapPath("/")%></li>
										<li>
										<u>Enter the full HTTP path to the file</u> (e.g. http://www.myStore.com/downloads/downloadfile.zip). This option does not use the <em>Hide URL </em>feature. There is no limitation on the file size, regardless of the version of IIS used on the Web server. For more information, please see the User Guide.</li>
										</ul>
									</td>
									</tr>
									<tr>
										<td colspan="2"><input type="text" name="producturl" value="<%=pProductURL%>" size="70" tabindex="803"></td>
									</tr>
									<tr>
										<td colspan="2">Make download URL expire:&nbsp;
											<input name="urlexpire" type="radio" value="1" <%if pURLExpire="1" then%>checked<%end if%> onClick="document.hForm.urlexpire1.value='1';" class="clearBorder" tabindex="804">&nbsp;Yes 
											<input name="urlexpire" type="radio" value="0" <%if pURLExpire="0" then%>checked<%end if%> onClick="document.hForm.urlexpire1.value='0'; document.hForm.expiredays.value='';" class="clearBorder" tabindex="805">&nbsp;No
										</td>
									</tr>
									<tr>
										<td colspan="2">URL will expire after: <input type="text" name="expiredays" value="<%=pExpireDays%>" size="5" tabindex="806">&nbsp;days</td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td>Deliver license with order confirmation:&nbsp;
											<input name="license" type="radio" value="1" <%if plicense="1" then%>checked<%end if%> onClick="document.hForm.license1.value='1';" class="clearBorder" tabindex="807">&nbsp;Yes 
											<input name="license" type="radio" value="0" <%if plicense="0" then%>checked<%end if%> onClick="document.hForm.license1.value='0'; document.hForm.locallg.value=''; document.hForm.remotelg.value='http://';" class="clearBorder" tabindex="808">&nbsp;No
										</td>
									</tr>
									<tr>
										<td colspan="2"><b>1.</b> Use local license generator. Enter file name:<br>
											<font color="#333333">Note: Upload license generator script to the folder &quot;/productcart/pcadmin/licenses&quot; and enter the filename in the textbox below (e.g. myLicense.asp). <a href="http://wiki.productcart.com/productcart/products_adding_new#downloadable_settings" target="_blank">More information</a></font></td>
									</tr>
									<tr>
										<td colspan="2"><input type="text" name="locallg" size="70" value="<%=pLocalLG%>" tabindex="809"></td>
									</tr>
									<tr>
										<td colspan="2"><b>2.</b> Use remote license generator. Enter URL to file:<br>
											<font color="#333333">Note: enter the full URL, starting with HTTP:// (e.g. http://www.myWeb.com/myLicense.asp).
											<a href="http://wiki.productcart.com/productcart/products_adding_new#downloadable_settings" target="_blank">More information</a></font></td>
									</tr>
									<tr>
										<td colspan="2">
											<input type="text" name="remotelg" value="<%=pRemoteLG%>" size="70" tabindex="810"></td>
									</tr>
									<tr>
										<td colspan="2">Your license generator can return to ProductCart up 5 variables. Enter the descriptive names for those variables. <a href="http://wiki.productcart.com/productcart/products_adding_new#downloadable_settings" target="_blank">More information</a></td>
									</tr>
									<tr>
										<td colspan="2">License Field (1):&nbsp;
										<input type="text" name="licenselabel1" size="36" value="<%=pLicenseLabel1%>" tabindex="811"></td>
									</tr>
									<tr>
										<td colspan="2">License Field (2):&nbsp;
										<input type="text" name="licenselabel2" size="36" value="<%=pLicenseLabel2%>" tabindex="812"></td>
									</tr>
									<tr>
										<td colspan="2">License Field (3):&nbsp;
										<input type="text" name="licenselabel3" size="36" value="<%=pLicenseLabel3%>" tabindex="813"></td>
									</tr>
									<tr>
										<td colspan="2">License Field (4):&nbsp;
										<input type="text" name="licenselabel4" size="36" value="<%=pLicenseLabel4%>" tabindex="814"></td>
									</tr>
									<tr>
										<td colspan="2">License Field (5):&nbsp;
										<input type="text" name="licenselabel5" size="36" value="<%=pLicenseLabel5%>" tabindex="815"></td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2">Additional copy for confirmation e-mail (e.g. installation instructions)</td>
									</tr>
									<tr>
										<td colspan="2"><textarea name="addtomail" rows="9" cols="65" tabindex="816"><%=pAddtoMail%></textarea></td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2" align="center">
										<input type="button" class="btn btn-default"  name="checkbutton" value=" Verify Download URL " onClick="javascript:CheckWindow();" tabindex="817">
										&nbsp;
										<input type="button" class="btn btn-default"  name="checkbutton" value=" Test license generator " onClick="javascript:TestWindow();" tabindex="818"><br>
										&nbsp;<br>
										<font color="#333333">Note: Please save all your updates for this product (if any) before testing license generator</font></td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
	
				</div>
			<%
			end if ' Hide for Configurable-Only Items
			'// EIGHTH PANEL - END
			
			'// NINTH PANEL - START - Gift certificate
			if pcv_ProductType="std" then ' Hide if this is not a standard product
			%>
				<div id="tab-9" class="tab-pane">
					
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<th colspan="2">Gift Certificate Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td colspan="2">This is a Gift Certificate&nbsp;
								<% If statusAPP="1" OR scAPP=1 Then %>
									<input name="GC" type="radio" value="1" <%if pGC="1" then%>checked<%end if%> onClick="document.hForm.apparel[1].checked='true'; document.hForm.downloadable[1].checked='true'; document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none';document.getElementById('show_20').style.display='';document.getElementById('show_21').style.display='none';" class="clearBorder" tabindex="901">Yes 
								<% Else %>
									<input name="GC" type="radio" value="1" <%if pGC="1" then%>checked<%end if%> onClick="document.hForm.downloadable[1].checked='true'; document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none';document.getElementById('show_20').style.display=''" class="clearBorder" tabindex="901">Yes 
								<% End If %>
								<input name="GC" type="radio" value="0" <%if pGC="0" then%>checked<%end if%> onClick="document.getElementById('show_20').style.display='none'" class="clearBorder" tabindex="902">No
							</td>
						</tr>
						<tr>
							<td colspan="2">                       
							<table id="show_20" <%if (pGC="0") or (pGC="") then%>style="display: none;"<%end if%> class="pcCPcontent">
							<tr>
								<td colspan="2">Expiration:</td>
							</tr>
							<tr>
								<td align="right">
									<input name="GCExp" type="radio" value="0" <%if pGCExp="0" then%>checked<%end if%> class="clearBorder" tabindex="903">
								</td>
								<td>Does not expire</td>
							</tr>
							<tr>
								<td align="right" valign="top">
									<input name="GCExp" type="radio" value="1" <%if pGCExp="1" then%>checked<%end if%> class="clearBorder" tabindex="904">
								</td>
								<td>Expires on:<br>
									Expiration Date: <input type="text" name="GCExpDate" size="25" value="<%=pGCExpDate%>" tabindex="905">&nbsp;(<i></i>Format: <%if scDateFrmt="DD/MM/YY" then%>DD/MM/YY<%else%>MM/DD/YY<%end if%></i>)
								</td>
							</tr>
							<tr>
								<td align="right" valign="top">
									<input name="GCExp" type="radio" value="2" <%if pGCExp="2" then%>checked<%end if%> class="clearBorder" tabindex="906">
								</td>
								<td>Expires N days after purchase<br>
									Numbers of Days: <input type="text" name="GCExpDay" size="5" value="<%=pGCExpDay%>" tabindex="907"> days
								</td>
							</tr>
							<tr>
								<td colspan="2">Electronic Only:&nbsp;
									<input name="GCEOnly" type="checkbox" value="1" <%if pGCEOnly="1" then%>checked<%end if%> class="clearBorder" tabindex="908">
								</td>
							</tr>
							<tr>
								<td colspan="2">Code Generator:</td>
							</tr>
							<tr>
								<td align="right">
									<input name="GCGen" type="radio" value="0" <%if pGCGen="0" then%>checked<%end if%> class="clearBorder" tabindex="909">
								</td>
								<td>Use Default</td>
							</tr>
							<tr>
								<td align="right" valign="top">
									<input name="GCGen" type="radio" value="1" <%if pGCGen="1" then%>checked<%end if%> class="clearBorder" tabindex="910">
								</td>
								<td>Use custom generator<br>
									File name: <input type="text" name="GCGenFile" size="53" value="<%=pGCGenFile%>" tabindex="911">
									<div class="pcCPnotes">Note: Upload your custom gift certificate code generator script to the folder &quot;pcadmin/licenses&quot; and enter the file name in the input field above (e.g. myGiftCode.asp)</div>
								</td>
							</tr>
							</table>
							</td>
						</tr>					
					</table>

				</div>
			<%
			end if ' Hide if this is not a standard product
			'// NINTH PANEL - END		


			'// APPAREL PANEL - START
			If statusAPP="1" OR scAPP=1 Then
				if pcv_ProductType="std" then ' Hide if this is not a standard product	
				%>
					<div id="tabs-13" class="tab-pane">
					
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr> 
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_70")%></th>
							</tr>
							<tr> 
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_71")%>&nbsp; 
									<input name="apparel" type="radio" value="1" <%if pcv_Apparel="1" then%>checked<%end if%> onClick="document.hForm.downloadable[1].checked='true'; document.hForm.GC[1].checked='true'; document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none'; document.getElementById('show_20').style.display='none';document.getElementById('show_21').style.display='';" class="clearBorder">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>
									<input name="apparel" type="radio" value="0" <%if (pcv_Apparel="0") or (pcv_Apparel="") then%>checked<%end if%> onClick="document.getElementById('show_21').style.display='none';" class="clearBorder">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">                       
								<table id="show_21" <%if (pcv_Apparel="0") or (pcv_Apparel="") then%>style="display:none"<%end if%> class="pcCPcontent">
									<tr>
										<td width="100%" nowrap colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_72")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="0" style="float: right" <%if (pcv_ShowStockMsg="0") or (pcv_ShowStockMsg="") then%>checked<%end if%> class="clearBorder">
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_73")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="2" style="float: right" <%if (pcv_ShowStockMsg="2") then%>checked<%end if%> class="clearBorder">
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_73a")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="3" style="float: right" <%if (pcv_ShowStockMsg="3") then%>checked<%end if%> class="clearBorder">
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_73b")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="1" style="float: right" class="clearBorder" <%if (pcv_ShowStockMsg="1") then%>checked<%end if%>>
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_74")%><br>
											<input type="text" name="stockmsg" size="40" value="<%=pcv_StockMsg%>"><br>
											<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_75")%>
										</td>
									</tr>
									<tr>
										<td colspan="2"><hr /></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_76")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="middle">
											<input type="radio" name="pcv_ApparelRadio" value="0" style="float: right" checked class="clearBorder" <%if pcv_ApparelRadio<>"1" then%>checked<%end if%>>
										</td>
										<td width="70%" valign="middle"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_77")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="middle">
											<input type="radio" name="pcv_ApparelRadio" value="1" style="float: right" class="clearBorder" <%if pcv_ApparelRadio="1" then%>checked<%end if%>>
										</td>
										<td width="70%" valign="middle"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_78")%></td>
									</tr>
									<tr>
										<td colspan="2"><hr /></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_79")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_80")%></td>
										<td width="70%" valign="top"><input type="text" name="sizelink" cols="40" value="<%=pcv_SizeLink%>"></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_81")%></td>
										<td width="70%" valign="top"><textarea rows="9" name="sizeinfo" cols="40"><%=pcv_SizeInfo%></textarea></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_82")%></td>
										<td width="70%" valign="top"><input type="text" name="sizeimg" size="40" value="<%=pcv_SizeImg%>"><br>
										<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_83")%>&nbsp;<a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_84")%></a></td>
									</tr>
									<tr>
										<td width="30%" nowrap align="right" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_85")%></td>
										<td width="70%" valign="top">
											<input type="text" name="sizeurl" size="40" value="<%=pcv_SizeURL%>">
											<br><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_86")%>
										</td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						
					</div>
				<%
				end if
			End IF
			'// APPAREL PANEL - END


			'// ELEVENTH PANEL - START - Custom fields
			if pcv_ProductType<>"item" then	 ' Hide for Configurable-Only Items
			%>
				<div id="tab-10" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Custom Search Fields</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">This tab will allow the store manager to view, add, and edit custom search fields associated with this product.</td>
						</tr>
						<tr>
							<td colspan="2">
								<%query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pIdProduct & ";"
								set rs=connTemp.execute(query)
								tmpJSStr=""
								tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
								intCount=-1
								IF not rs.eof THEN
									pcArr=rs.getRows()
									set rs=nothing
									intCount=ubound(pcArr,2)
									For i=0 to intCount
										tmpJSStr=tmpJSStr & "SFID[" & i & "]=" & pcArr(0,i) & ";" & vbcrlf
										tmpJSStr=tmpJSStr & "SFNAME[" & i & "]='" & replace(pcArr(1,i),"'","\'") & "';" & vbcrlf
										tmpJSStr=tmpJSStr & "SFVID[" & i & "]=" & pcArr(2,i) & ";" & vbcrlf
										tmpJSStr=tmpJSStr & "SFVALUE[" & i & "]='" & replace(pcArr(3,i),"'","\'") & "';" & vbcrlf
										tmpJSStr=tmpJSStr & "SFVORDER[" & i & "]=" & pcArr(4,i) & ";" & vbcrlf
									Next
								END IF
								set rs=nothing
								tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf%>
								<script type=text/javascript>
									<%=tmpJSStr%>
									function CreateTable()
									{
										var tmp1="";
										var tmp2="";
										var i=0;
										var found=0;
										tmp1='<table class="pcCPcontent"><tr><td></td><td nowrap><strong>Text to display</strong></td><td><strong>Value</strong></td></tr>';
										for (var i=0;i<=SFCount;i++)
										{
											found=1;
											tmp1=tmp1 + '<tr><td align="right"><a href="javascript:ClearSF(SFID['+i+']);"><img src="../pc/images/minus.jpg" alt="Remove" border="0"></a></td><td width="275" nowrap>'+SFNAME[i]+'</td><td width="100%">'+SFVALUE[i]+'</td></tr>';
											if (tmp2=="") tmp2=tmp2 + "||";
											tmp2=tmp2 + "^^^" + SFID[i] + "^^^" + SFVID[i] + "^^^" + SFVALUE[i] + "^^^" + SFVORDER[i] + "^^^||"
										}
										tmp1=tmp1+'</table>';
										if (found==0) tmp1="<br><b>No search fields are assigned to this product</b><br><br>";
										document.getElementById("stable").innerHTML=tmp1;
										document.hForm.SFData.value=tmp2;
									}
									function ClearSF(tmpSFID)
									{
										var i=0;
										for (var i=0;i<=SFCount;i++)
										{
											if (SFID[i]==tmpSFID)
											{
												removedArr = SFID.splice(i,1);
												removedArr = SFNAME.splice(i,1);
												removedArr = SFVID.splice(i,1);
												removedArr = SFVALUE.splice(i,1);
												removedArr = SFVORDER.splice(i,1);
												SFCount--;
												break;
											}
										}
										CreateTable();
									}
					
									function AddSF(tmpSFID,tmpSFName,tmpSVID,tmpSValue,tmpSOrder)
									{
										if (tmpSValue!="")
										{
											var i=0;
											var found=0;
											for (var i=0;i<=SFCount;i++)
											{
												if (SFID[i]==tmpSFID)
												{
													SFVID[i]=tmpSVID;
													SFVALUE[i]=tmpSValue;
													SFVORDER[i]=tmpSOrder;
													found=1;
													break;
												}
											}
											if (found==0)
											{
												SFCount++;
												SFID[SFCount]=tmpSFID;
												SFNAME[SFCount]=tmpSFName;
												SFVID[SFCount]=tmpSVID;
												SFVALUE[SFCount]=tmpSValue;
												SFVORDER[SFCount]=tmpSOrder;
											}
											CreateTable();
										}
									}
								</script>
								<span id="stable" name="stable"></span>
								<input type="hidden" name="SFData" value="">
								<%query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields WHERE pcSearchFieldCPShow=1 ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
								set rs=Server.CreateObject("ADODB.Recordset")
								set rs=conntemp.execute(query)
								if not rs.eof then
									set pcv_tempFunc = new StringBuilder
									pcv_tempFunc.append "<script type=text/javascript>" & vbcrlf
									pcv_tempFunc.append "function CheckList(cvalue) {" & vbcrlf
									pcv_tempFunc.append "if (cvalue==0) {" & vbcrlf
									pcv_tempFunc.append "var SelectA = document.hForm.SearchValues;" & vbcrlf
									pcv_tempFunc.append "SelectA.options.length = 0; }" & vbcrlf
					
									set pcv_tempList = new StringBuilder
									pcv_tempList.append "<select name=""customfield"" onchange=""javascript:document.hForm.newvalue.value='';document.hForm.neworder.value='0';CheckList(document.hForm.customfield.value);"">" & vbcrlf
					
									pcArray=rs.getRows()
									intCount=ubound(pcArray,2)
									set rs=nothing
					
									For i=0 to intCount
										pcv_tempList.append "<option value=""" & pcArray(0,i) & """>" & replace(pcArray(1,i),"""","&quot;") & "</option>" & vbcrlf
										query="SELECT idSearchData,pcSearchDataName FROM pcSearchData WHERE idSearchField=" & pcArray(0,i) & " ORDER BY pcSearchDataOrder ASC,pcSearchDataName ASC;"
										set rs=connTemp.execute(query)
										if not rs.eof then
											tmpArr=rs.getRows()
											LCount=ubound(tmpArr,2)
											pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
											pcv_tempFunc.append "var SelectA = document.hForm.SearchValues;" & vbcrlf
											pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
											For j=0 to LCount
												pcv_tempFunc.append "SelectA.options[" & j & "]=new Option(""" & replace(tmpArr(1,j),"""","\""") & """,""" & tmpArr(0,j) & """);" & vbcrlf
											Next
											pcv_tempFunc.append "}" & vbcrlf
										else
											pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
											pcv_tempFunc.append "var SelectA = document.hForm.SearchValues;" & vbcrlf
											pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
											pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option("""",""""); }" & vbcrlf
										end if
									Next
			
									pcv_tempList.append "</select>" & vbcrlf
									pcv_tempFunc.append "}" & vbcrlf
									pcv_tempFunc.append "</script>" & vbcrlf
									
									pcv_tempList=pcv_tempList.toString
									pcv_tempFunc=pcv_tempFunc.toString
									%>
									<br><br>
									<hr>
									<table class="pcCPcontent" style="width:auto;">
										<tr>
											<td colspan="2"><a name="2"></a><b><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_91")%></b></td>
										</tr>
										<tr>
											<td width="20%"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_92")%></td>
											<td width="80%">
											<%=pcv_tempList%>&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_93")%>&nbsp;
											<select name="SearchValues">
											</select>
											<%=pcv_tempFunc%>
											<script type=text/javascript>
												CheckList(document.hForm.customfield.value);
											</script>
											&nbsp;<a href="javascript:AddSF(document.hForm.customfield.value,document.hForm.customfield.options[document.hForm.customfield.selectedIndex].text,document.hForm.SearchValues.value,document.hForm.SearchValues.options[document.hForm.SearchValues.selectedIndex].text,0);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
											</td>
										</tr>
										<tr>
											<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_94")%></td>
											<td>
												<input type="text" value="" name="newvalue" size="30">
                        						<input type="hidden" value="0" name="neworder">
												&nbsp;<a href="javascript:AddSF(document.hForm.customfield.value,document.hForm.customfield.options[document.hForm.customfield.selectedIndex].text,-1,document.hForm.newvalue.value,document.hForm.neworder.value);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
											</td>
										</tr>
										<tr>
											<td colspan="2">
												<b><u><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_88")%></u></b> <i><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_90")%></i>
												<br><br>
											</td>
										</tr>
							  </table>
								<%else
									query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
									set rs=Server.CreateObject("ADODB.Recordset")
									set rs=conntemp.execute(query)
									if not rs.eof then%>
										<a href="ManageSearchFields.asp">Click here</a> to manage custom search fields.</a>
									<%else%>
										<a href="ManageSearchFields.asp">Click here</a> to add new product custom search field.</a>
									<%end if
									set rs=nothing%>
								<%end if%>
								<script type=text/javascript>CreateTable();</script>
							</td>
						</tr>
					</table>
				
				</div>
				
				<div id="tab-11" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Google Shopping Settings</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2"><b>Google Product Category</b></td>
						</tr>
						<tr>
							<td><input type="radio" name="pcv_GPC" value="0" <%if (pcv_GCat="") OR IsNull(pcv_GCat) then%>checked<%end if%> class="clearBorder"></td>
							<td>Use the Product’s current category assignment for Google Shopping. (Set by default)</td>
						</tr>
						<tr>
							<td><input type="radio" name="pcv_GPC" value="1" <%if (pcv_GCat<>"") then%>checked<%end if%> class="clearBorder"></td>
							<td>Use a Google Product Category Attribute</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td>
								<%GCatPre=0%>
								<select name="pcv_GCat">
									<option value="" selected>Select one... </option>
									<option value="Apparel &amp; Accessories" <%if (pcv_GCat="Apparel &amp; Accessories") then%><%GCatPre=1%>selected<%end if%>>Apparel &amp; Accessories</option>
									<option value="Apparel &amp; Accessories &gt; Clothing" <%if (pcv_GCat="Apparel &amp; Accessories &gt; Clothing") then%><%GCatPre=1%>selected<%end if%>>Apparel &amp; Accessories &gt; Clothing</option>
									<option value="Apparel &amp; Accessories &gt; Shoes" <%if (pcv_GCat="Apparel &amp; Accessories &gt; Shoes") then%><%GCatPre=1%>selected<%end if%>>Apparel &amp; Accessories &gt; Shoes</option>
									<option value="Media &gt; Books" <%if (pcv_GCat="Media &gt; Books") then%><%GCatPre=1%>selected<%end if%>>Media &gt; Books</option>
									<option value="Media &gt; DVDs &amp; Movies" <%if (pcv_GCat="Media &gt; DVDs &amp; Movies") then%><%GCatPre=1%>selected<%end if%>>Media &gt; DVDs &amp; Movies</option>
									<option value="Media &gt; Music" <%if (pcv_GCat="Media &gt; Music") then%><%GCatPre=1%>selected<%end if%>>Media &gt; Music</option>
									<option value="Software &gt; Video Game Software" <%if (pcv_GCat="Software &gt; Video Game Software") then%><%GCatPre=1%>selected<%end if%>>Software &gt; Video Game Software</option>
								</select>
						<tr>
							<td>&nbsp;</td>
							<td>
								Or using other: <input type="text" name="pcv_GCatO" size="35" value="<%if GCatPre=0 then%><%=pcv_GCat%><%end if%>"><br>
								<i><u>Note:</u> To get correct Google's Product Taxonomy, <a href="http://support.google.com/merchants/bin/answer.py?hl=en&answer=1705911" target="_blank">click here</a></i>
						 	</td>
						</tr>
						<tr>
							<td colspan="2"><hr width="95%"></td>
						</tr>
						<tr>
							<td colspan="2"><b>Google Apparel Product Attributes</b></td>
						</tr>
						<tr>
							<td>Gender:</td>
							<td>
								<select name="pcv_GGen">
									<option value="" selected>Select one... </option>
									<option value="male" <%if ucase(pcv_GGen)="MALE" then%>selected<%end if%>>Male</option>
									<option value="female" <%if ucase(pcv_GGen)="FEMALE" then%>selected<%end if%>>Female</option>
									<option value="unisex" <%if ucase(pcv_GGen)="UNISEX" then%>selected<%end if%>>Unisex</option>
								</select>
							</td>
						</tr>
						<tr>
							<td>Age Group:</td>
							<td>
								<select name="pcv_GAge">
									<option value="" selected>Select one... </option>
									<option value="adult" <%if ucase(pcv_GAge)="ADULT" then%>selected<%end if%>>Adult</option>
									<option value="kids" <%if ucase(pcv_GAge)="KIDS" then%>selected<%end if%>>Kids</option>
								</select>
							</td>
						</tr>
						<tr>
							<td>Size:</td>
							<td>
								<input type="text" name="pcv_GSize" size="35" value="<%=pcv_GSize%>">
							</td>
						</tr>
						<tr>
							<td>Color:</td>
							<td>
								<input type="text" name="pcv_GColor" size="35" value="<%=pcv_GColor%>">
							</td>
						</tr>
						<tr>
							<td>Pattern:</td>
							<td>
								<input type="text" name="pcv_GPat" size="35" value="<%=pcv_GPat%>">
							</td>
						</tr>
						<tr>
							<td>Material:</td>
							<td>
								<input type="text" name="pcv_GMat" size="35" value="<%=pcv_GMat%>">
							</td>
						</tr>
					</table>
				
				</div>
				
			<%
			end if	 ' Hide for Configurable-Only Items
			'// ELEVENTH PANEL - END
			%>
            
			</div>
		<%
		'// TABBED PANELS - MAIN DIV END
		%>
        
        </div>  
    </div>
    <div style="clear: both;">&nbsp;</div>

		<script type=text/javascript>
			var tab = window.location.hash;
			$pc('#TabbedPanels1 .tabs-left').on('click', 'a', function() {
				window.location.hash = $pc(this).attr('href');
			});

			$pc('#TabbedPanels1 a[href="' + tab + '"]').tab('show');
		</script>

</form> 
<!--#include file="Adminfooter.asp"-->