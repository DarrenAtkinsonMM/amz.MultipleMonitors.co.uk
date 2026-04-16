<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------

Dim pcStrPageName
pcStrPageName = "quickview.asp"
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/CashbackConstants.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<script type="text/javascript" src="<%=pcf_getJSPath("../includes/mojozoom","mojozoom.min.js")%>"></script>
<%
'Response.Buffer = True
'-------------------------------
' declare local variables
'-------------------------------

Dim tIndex, tUpdPrd, pIdCategory, strBreadCrumb, pIdProduct, dblpcCC_Price
Dim pcv_strViewPrdStyle, pcv_strFormAction, pcv_intValidationFile, pcv_blnBTOisConfig, iRewardPoints, pDescription, pMainProductName
Dim pSku, pconfigOnly, pserviceSpec, pPrice, pBtoBPrice, pDetails, pListPrice, plistHidden, pimageUrl, pLgimageURL
Dim pArequired, pBrequired, pStock, pWeight, pEmailText, pFormQuantity, pnoshipping, pcustom1, pcontent1
Dim pcustom2, pcontent2, pcustom3, pcontent3, pnoprices
Dim pIDBrand, psDesc, pNoStock, pnoshippingtext, intIdProduct, intWeight, optionA, optionB
Dim pcv_intHideBTOPrice, pcv_intQtyValidate, pcv_lngMinimumQty, intpHideDefConfig, pnotax, BrandName
Dim FirstCnt, strDescription, intReward, pcv_BTORP, strConfigProductCategory, dblPrice, dblWPrice, intIdCategory
Dim VardiscGo, dblQuantityFrom, dblQuantityUntil, dblPercentage, dblDiscountPerWUnit, dblDiscountPerUnit
Dim intIdOptOptGrp, intIdOption, strOptionDescrip, OptInActive, optPrice, tempIdOptA, tempIdOptB
Dim xrequired, xfieldCnt, reqstring, TextArea, widthoffield, rowlength
Dim scCS, cs_showprod, cs_showcart, cs_showimage, crossSellText
Dim pcv_strOptionGroupDesc, pcv_intOptionGroupCount, pcv_strOptionGroupCount, pcv_strOptionGroupID, pcv_strOptionRequired
Dim xOptionsCnt, pcv_strNumberValidations, pcv_strFuntionCall, pcv_strReqOptString, xOtionrequired, pcv_strCSDiscounts , pcv_strPrdDiscounts 
Dim pcv_strProdImage_Url, pcv_strProdImage_LargeUrl, pcv_intProdImage_Columns, pcv_strShowImage_LargeUrl, pcv_strShowImage_Url, pcv_strCurrentUrl
Dim pcv_strAdditionalImages, cCounter, pcv_strWishListLink, BTOCharges, pcv_strCSString, pcv_strReqCSString, cs_RequiredIds, xCSCnt
Dim iAddDefaultWPrice, iAddDefaultPrice, pcv_intActive
Dim pHideSKU, pcv_IntMojoZoom

Dim pcv_ReorderLevel,pcv_Apparel,pcv_ShowStockMsg,pcv_StockMsg,pcv_SizeLink,pcv_SizeInfo,pcv_SizeImg,pcv_SizeURL,pcv_ApparelRadio,pcv_HaveSPs
Dim pcv_TotalOpts

Dim pShowAvgRating, pAvgRating, pNumRatings, pcRS_Active, pRSActive, pcv_RatingType, pcv_Img1

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
pcv_IsQuickView = True
%>
<!--#include file="prv_getSettings.asp"-->
<!--#include file="viewPrdCode.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
'--> Check if this customer is logged in with a customer category
dblpcCC_Price=0
if session("customerCategory")<>0 then
	query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rs=nothing
		'//close any connections
		call closedb()
		'//redirect to error page
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if NOT rs.eof then
		strcustomerCategory="YES"
		dblpcCC_Price=rs("pcCC_Price")
		dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
	else
		strcustomerCategory="NO"
	end if
	set rs=nothing
end if

'--> check for discount per quantity
query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rs=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if not rs.eof then
	pDiscountPerQuantity=-1
else
	pDiscountPerQuantity=0
end if
set rs=nothing

'--> gets product details from db

query="SELECT active,iRewardPoints, description, sku, configOnly, serviceSpec, price, btobprice, listprice, listHidden, imageurl, largeImageURL, Arequired, Brequired, stock, weight, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3, noprices, IDBrand, noStock, noshippingtext,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcProd_multiQty,pcprod_HideDefConfig, notax, pcProd_BackOrder,pcProd_ShipNDays,pcProd_SkipDetailsPage,pcProd_ReorderLevel,pcprod_Apparel,pcprod_ShowStockMsg,pcprod_StockMsg,pcprod_SizeLink,pcprod_SizeInfo,pcprod_SizeImg,pcprod_SizeURL,pcProd_ApparelRadio,pcProd_HideSKU,pcPrd_MojoZoom,details, sDesc, pcProd_AvgRating FROM products WHERE idProduct=" & pidProduct & " AND configOnly=0 AND removed=0 "
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rs=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Check to see if a product with that ID exists in the database
if rs.eof then
	set rs=nothing
	call closeDb()
  	response.redirect "msg.asp?message=88"
end if

'// Load product status (active or inactive)
pcv_intActive=rs("active")
	
'// If inactive and not previewed, redirect to "product inactive" message
if pcv_intActive<>-1 AND session("pcv_intAdminPreview")<>1 then
	set rs=nothing
	call closeDb()
  	response.redirect "msg.asp?message=95"
end if

'// Assign variable values
iRewardPoints=rs("iRewardPoints")
pDescription=replace(rs("description"),"&quot;",chr(34))
pMainProductName=pDescription
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
pLgimageURL=rs("largeImageURL")
pArequired=rs("Arequired")
pBrequired=rs("Brequired")
pStock=rs("stock")
pWeight=rs("weight")
pEmailText=rs("emailText")
pFormQuantity=rs("formQuantity")
pnoshipping=rs("noshipping")
pcustom1=rs("custom1")
pcontent1=rs("content1")
pcustom2=rs("custom2")
pcontent2=rs("content2")
pcustom3=rs("custom3")
pcontent3=rs("content3")

pnoprices=rs("noprices")
if isNull(pnoprices) OR pnoprices="" then
	pnoprices=0
end if
pIDBrand=rs("IDBrand")
pNoStock=rs("noStock")
pnoshippingtext=rs("noshippingtext")
pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
if isNull(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
	pcv_intHideBTOPrice="0"
end if
pcv_intQtyValidate=rs("pcprod_QtyValidate")
if isNull( pcv_intQtyValidate) OR  pcv_intQtyValidate="" then
	pcv_intQtyValidate="0"
end if				
pcv_lngMinimumQty=rs("pcprod_MinimumQty")
if isNull(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
	pcv_lngMinimumQty="0"
end if
pcv_lngMultiQty=rs("pcProd_multiQty")
if isNull(pcv_lngMultiQty) OR pcv_lngMultiQty="" then
	pcv_lngMultiQty="0"
end if
intpHideDefConfig=rs("pcprod_HideDefConfig")
if isNull(intpHideDefConfig) OR intpHideDefConfig="" then
	intpHideDefConfig="0"
end if
pnotax=rs("notax")

'Start SDBA
pcv_intBackOrder = rs("pcProd_BackOrder")
if isNull(pcv_intBackOrder) OR pcv_intBackOrder="" then
	pcv_intBackOrder = 0
end if
pcv_intShipNDays = rs("pcProd_ShipNDays")
if isNull(pcv_intShipNDays) OR pcv_intShipNDays="" then
	pcv_intShipNDays = 0
end if
'End SDBA

pcv_intSkipDetailsPage=rs("pcProd_SkipDetailsPage")
if isNull(pcv_intSkipDetailsPage) or pcv_intSkipDetailsPage="" then
	pcv_intSkipDetailsPage=0
end if

pcv_ReorderLevel=rs("pcProd_ReorderLevel")
if IsNull(pcv_ReorderLevel) or pcv_ReorderLevel="" then
	pcv_ReorderLevel=0
end if
pcv_Apparel=rs("pcprod_Apparel")
if pcv_Apparel="1" then
	pcv_ShowStockMsg=rs("pcprod_ShowStockMsg")
	if IsNull(pcv_ShowStockMsg) or pcv_ShowStockMsg="" then
		pcv_ShowStockMsg="0"
	end if
	pcv_StockMsg=rs("pcprod_StockMsg")
	if IsNull(pcv_StockMsg) or pcv_StockMsg="" then
		pcv_StockMsg=dictLanguage.Item(Session("language")&"_viewPrd_7")
	end if
	pcv_SizeLink=rs("pcprod_SizeLink")
	pcv_SizeInfo=rs("pcprod_SizeInfo")
	pcv_SizeImg=rs("pcprod_SizeImg")
	pcv_SizeURL=rs("pcprod_SizeURL")
	if pcv_SizeURL<>"" then
		if ucase(pcv_SizeURL)="HTTP://" then
			pcv_SizeURL=""
		end if
	end if
	pcv_ApparelRadio=rs("pcprod_ApparelRadio")
	if IsNull(pcv_ApparelRadio) or pcv_ApparelRadio="" then
		pcv_ApparelRadio="0"
	end if
end if

pHideSKU=rs("pcProd_HideSKU")
if IsNull(pHideSKU) or pHideSKU="" then
	pHideSKU=0
end if

pcv_IntMojoZoom=rs("pcPrd_MojoZoom")
if not validNum(pcv_IntMojoZoom) then
	pcv_IntMojoZoom=0
end if

pDetails=replace(rs("details"),"&quot;",chr(34))
psDesc=rs("sDesc")

' PRV41 start
pAvgRating = rs("pcProd_AvgRating")
' PRV41 end

set rs=Nothing

'Get XFields
Dim pcXFArr,intXFCount
intXFCount=-1
query="SELECT IdXField,pcPXF_XReq FROM pcPrdXFields WHERE idProduct=" & pidProduct & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	pcXFArr=rs.getRows()
	intXFCount=ubound(pcXFArr,2)
end if
set rs=nothing

'// Disregard Stock option is always checked
if pcv_Apparel="1" then
	pNoStock=1
end if

%>
<!--#include file="inc_CheckReqItemStock.asp"-->
<%
pNumRatings = pcf_TotalReviewCount(pidProduct)

'--> Check to see if the product has been assigned to a brand. If so, get the brand name
if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
 	query="select BrandName from Brands where IDBrand=" & pIDBrand
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if not rs.eof then
		BrandName=rs("BrandName")
	end if
	set rs=nothing
end if
'SB S
Dim objSB 
Set objSB = New pcARBClass
pSubscriptionID = objSB.getSubscriptionID(pidProduct)
if isNull(pSubscriptionID) OR pSubscriptionID="" then
	pSubscriptionID = "0"
end if
%>
<!--#include file="../includes/pcSBDataInc.asp" --> 
<% 
'SB E

'-----------------------------------
' START: Skip Product Details Page
'-----------------------------------
If pcv_intSkipDetailsPage="1" then
	If pserviceSpec<>0 Then
		'// SEO URLs START
		if scSeoURLs=1 then
			session("idProductTransfer")=pidProduct
			Server.Transfer("configureprd.asp") 
		else
			response.redirect "configurePrd.asp?idProduct="&pidProduct
		end if
		'// SEO URLs END
	End if
End if
'-----------------------------------
' END: Skip Product Details Page
'-----------------------------------

%>

<div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
    <h4 class="modal-title" id="myModalLabel">&nbsp;</h4>
</div>
<div id="pcQuickViewBody" class="modal-body">

<div class="pcClear"></div>

<!--#include file="../includes/javascripts/pcValidateViewPrd.asp"-->
<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->

<!-- Start Form -->
<% If statusAPP="1" Then %>
	<script>
	<%if pImageUrl<>"" then%>
	var GeneralImg="<%=pImageUrl%>";
	<%else%>
	var GeneralImg="no_image.gif";
	<%end if%>
	<%if pLgimageURL<>"" then%>
	var DefLargeImg="<%=pLgimageURL%>";
	var LargeImg=DefLargeImg
	<%else%>
	var DefLargeImg="";
	var LargeImg=DefLargeImg
	<%end if%>
	
		function open_win(url)
		{
		    var newWin=window.open(url,'','scrollbars=yes,resizable=yes');
		    newWin.focus();
		}
	</script>
	<%
	IF (pcv_Apparel="1") then
		call GenApparelSubProducts()
	END IF
End If
%>
<% 
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' GENERATE FORM																						//
' > BTO / BTO Configured / Standard Product															//
' > Each uses a different form action and JavaScript validation function                            //
'/////////////////////////////////////////////////////////////////////////////////////////////////////

'********************************************************************
' VALIDATION FILE
' pcv_intValidationFile = 1 // BTO
' pcv_intValidationFile = 2 // Standard
'
' FORM ACTION
' pcv_strFormAction = "instConfiguredPrd.asp" // BTO configured
' pcv_strFormAction = "instPrd.asp" // BTO NON configured and Standard
'********************************************************************
pcv_blnBTOisConfig = pcf_BTOisConfig '// returns true or false for Configured BTO

If pserviceSpec<>0 Then '// If its BTO Then
	if pcv_blnBTOisConfig then '// if its configured then
		pcv_strFormAction = "instConfiguredPrd.asp"
		pcv_intValidationFile = 1
	else '// Its not configured
		pcv_strFormAction = "instPrd.asp"
		pcv_intValidationFile = 1
	end if
else '// Its standard
	pcv_strFormAction = "instPrd.asp"
	pcv_intValidationFile = 2
end if
%>

<div id="pcMain" class="pcQuickViewContent">
	<form method="post" action="<%=pcv_strFormAction%>" id="additem" name="additem" onsubmit="javascript:if (checkproqty(document.additem.quantity)) {return(ajaxsubmit());} else {return(false)};" class="pcForms">
	<script type=text/javascript>
	
	$pc(document).ready(function () {
		try { hs.zIndexCounter = 2050; } catch (err) { /* Ignore errors */ }
	});

	function ajaxsubmit()
	{
		var tmpAction=document.additem.action;
		if (tmpAction.indexOf("configurePrd.asp")>=0)
		{
			window.location=document.additem.action;
			return(false);
		}
		$pc.ajax({
				type: "POST",
				url: "<%=pcv_strFormAction%>",
				data: $pc('#additem').formSerialize(),
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="OK")
					{
                       window.location = '<%=pcf_GetCurrentPage() %>';
					}
					else
					{
						if (data=="CART")
						{
							window.location = 'viewCart.asp?cs=1';
						}
					}
				}
				});
		return(false);
   }
	</script>
	<%
	
	'************************
	' GET BTO Config Infor
	'************************
	pcs_GetBTOConfigurationQV
	
	%>
	<div class="qvmainLeft">
		<% Call pcs_ProductImage() %>
        
        <%
        pcs_AdditionalImages()
        
        pcv_strAdditionalImages = pcf_GetAdditionalImages
        if len(pcv_strAdditionalImages)>0 then
            %><div><a href="viewPrd.asp?idproduct=<%=pIdProduct%>"><%=dictLanguage.Item(Session("language")&"_QV_12")%></a></div><%
        end if
        %>
	</div>
	<div class="qvmainRight">
		<%
        pcs_ProductName
		pcs_ShowSKU
		pcs_ShowRating
		
		'Create Tabs
		pcvNumTabs=0
		ActiveTabId=0
		tabList=""
		tabOverViewTxt=""
		tabOverViewId=0
		tabBTOTxt=""
		tabBTOId=0
		tabOptionsTxt=""
		tabOptionsId=0
		tabCFsTxt=""
		tabCFsId=0
		tabPromoTxt=""
		tabPromoId=0
		tabDiscTxt=""
		tabDiscId=0
        pcv_strActiveTab = "in active"
		
		if pcv_Apparel="1" then
			tabOverViewTxt=pcf_ProductDescriptionQV & pcf_ShowBrandQV & pcf_CustomSearchFieldsQV & pcf_NoShippingTextQV
		else
			tabOverViewTxt=pcf_ProductDescriptionQV & pcf_ShowBrandQV & pcf_CustomSearchFieldsQV & pcf_DisplayWeightQV & pcf_RewardPointsQV & pcf_UnitsStockQV & pcf_DisplayBOMsgQV & pcf_NoShippingTextQV
		end if
		
		tabBTOTxt=pcf_BTOConfigurationQV
		if tabBTOTxt<>"" then
			pcvNumTabs=pcvNumTabs+1
			tabBTOId=pcvNumTabs
			tabBTOTxt="<div class=""tab-pane fade " & pcv_strActiveTab & " "" id='tab-" & pcvNumTabs & "'>" & tabBTOTxt & "</div>"
			tabList=tabList & "<li class=""" & pcv_strActiveTab & """><a data-toggle=""tab"" href='#tab-" & pcvNumTabs & "'>" & dictLanguage.Item(Session("language")&"_QV_6") & "</a></li>"
            pcv_strActiveTab=""
		end if
		
		if pcf_VerifyShowOptions then
			Call pcs_OptionsN()
			tabOptionsTxt=tmpQVOptions
			if pcv_intOptionGroupCount>0 then
				pcvNumTabs=pcvNumTabs+1
				tabOptionsId=pcvNumTabs
			
				if pcv_Apparel="1" then
					tabOptionsTxt="<div class=""tab-pane fade " & pcv_strActiveTab & " "" id='tab-" & pcvNumTabs & "'>" & tabOptionsTxt & pcf_DisplayWeightQV & pcf_RewardPointsQV & "</div>"
				else
					tabOptionsTxt="<div class=""tab-pane fade " & pcv_strActiveTab & " "" id='tab-" & pcvNumTabs & "'>" & tabOptionsTxt & "</div>"
				end if

				tabList=tabList & "<li class=""" & pcv_strActiveTab & """><a data-toggle=""tab"" href='#tab-" & pcvNumTabs & "'>" & dictLanguage.Item(Session("language")&"_QV_7") & "</a></li>"
                pcv_strActiveTab=""
			end if
		end if
		
		if pcf_VerifyShowOptions then
			tabCFCnt = 0
			tabCFsTxt=pcf_OptionsXQV(tabCFCnt)
			if tabCFsTxt<>"" And tabCFCnt > 0 then
				pcvNumTabs=pcvNumTabs+1
				tabCFsId=pcvNumTabs
				tabCFsTxt="<div class=""tab-pane fade " & pcv_strActiveTab & " "" id='tab-" & pcvNumTabs & "'>" & tabCFsTxt & "</div>"
				tabList=tabList & "<li class=""" & pcv_strActiveTab & """><a data-toggle=""tab"" href='#tab-" & pcvNumTabs & "'>" & dictLanguage.Item(Session("language")&"_QV_8") & "</a></li>"
                pcv_strActiveTab=""
			end if
		end if
		
		if tabOverViewTxt<>"" then
			pcvNumTabs=pcvNumTabs+1
			tabOverViewId=pcvNumTabs
			tabOverViewTxt="<div class=""tab-pane fade " & pcv_strActiveTab & " "" id='tab-" & pcvNumTabs & "'>" & tabOverViewTxt & "</div>"
			tabList=tabList & "<li class=""" & pcv_strActiveTab & """><a data-toggle=""tab"" href='#tab-" & pcvNumTabs & "'>" & dictLanguage.Item(Session("language")&"_QV_5") & "</a></li>"
            pcv_strActiveTab=""
		end if
		
		tabPromoTxt=pcf_ProductPromotionMsgQV
		if tabPromoTxt<>"" then
			pcvNumTabs=pcvNumTabs+1
			tabPromoId=pcvNumTabs
			tabPromoTxt="<div class=""tab-pane fade " & pcv_strActiveTab & " "" id='tab-" & pcvNumTabs & "'>" & tabPromoTxt & "</div>"
			tabList=tabList & "<li class=""" & pcv_strActiveTab & """><a data-toggle=""tab"" href='#tab-" & pcvNumTabs & "'>" & dictLanguage.Item(Session("language")&"_QV_9") & "</a></li>"
            pcv_strActiveTab=""
		end if
		
		tabDiscTxt=pcf_QtyDiscountsQV
		if tabDiscTxt<>"" then
			pcvNumTabs=pcvNumTabs+1
			tabDiscId=pcvNumTabs
			tabDiscTxt="<div class=""tab-pane fade " & pcv_strActiveTab & " "" id='tab-" & pcvNumTabs & "'>" & tabDiscTxt & "</div>"
			tabList=tabList & "<li class=""" & pcv_strActiveTab & """><a data-toggle=""tab"" href='#tab-" & pcvNumTabs & "'>" & dictLanguage.Item(Session("language")&"_QV_10") & "</a></li>"
            pcv_strActiveTab=""
		end if


		if pcvNumTabs>0 AND tabList<>"" then%>
			<br />
			<%if pcv_intOptionGroupCount=0 then%>
				<%=tabOptionsTxt%>
				<%tabOptionsTxt=""%>
			<%end if%>
			<%if strShowBTOQV1<>"" then%>
				<%=strShowBTOQV1%>
				<%strShowBTOQV1=""%>
			<%end if%>
			<div id="QuickViewTabArea">
				<div id="qvtabs" class="tabbable">
					<ul class="nav nav-tabs">
						<%=tabList%>
					</ul>
                    <div class="tab-content">
                        <%if tabOverViewTxt<>"" then%>
                            <%=tabOverViewTxt%>
                        <%end if%>
                        <%if tabBTOTxt<>"" then%>
                            <%=tabBTOTxt%>
                        <%end if%>
                        <%if tabOptionsTxt<>"" then%>
                            <%=tabOptionsTxt%>
                        <%end if%>
                        <%if tabCFsTxt<>"" then%>
                            <%=tabCFsTxt%>
                        <%end if%>
                        <%if tabPromoTxt<>"" then%>
                            <%=tabPromoTxt%>
                        <%end if%>
                        <%if tabDiscTxt<>"" then%>
                            <%=tabDiscTxt%>
                        <%end if%>
                    </div>
				</div>                
			</div>
			<%if tabOptionsTxt<>"" then
				ActiveTabId=tabOptionsId
			else
				if tabCFsTxt<>"" then
					ActiveTabId=tabCFsId
				else
					if tabBTOTxt<>"" then
						ActiveTabId=tabBTOId
					else
						if tabPromoTxt<>"" then
							ActiveTabId=tabPromoId
						else
							if tabDiscTxt<>"" then
								ActiveTabId=tabDiscId
							else
								ActiveTabId=tabOverViewId
							end if
						end if
					end if
				end if
			end if
		end if
		'End Tabs
		
		%>
		<div class="QVPriceLeft">
		<%'Display Prices
		Call pcs_ProductPrices()
		Call pcs_BTOADDON()%>
		</div>
		<div class="QVCartRight">
		<%
		'Display "Add to Cart" button
		if pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0 then
			if pEmailText<>"" then 
				response.write "<div class=pcShowProductNFS>" 
				response.write pEmailText '// reason why it's not for sale
				response.write "</div>" 
			end if
			tmp_showQty="1"
			if pcv_lngMinimumQty>0 then
				tmpMinQty=pcv_lngMinimumQty
			else
				tmpMinQty=1
			end if%>
			<input type="hidden" name="quantity" value="<%=tmpMinQty%>">
			<%
		else 
			If scorderlevel = "0" OR pcf_WholesaleCustomerAllowed Then
				if pcf_OutStockPurchaseAllow then
					If ((pserviceSpec<>0) AND ((pnoprices>0) OR (pPrice=0) OR (scConfigPurchaseOnly=1))) or ((iBTOQuoteSubmitOnly=1) and (pserviceSpec<>0)) then
						pcs_CustomizeButton
					else
            			pcs_AddtoCart
					end if  
				end if ' end 2b
		
			else
				tmp_showQty="1"
				if pcv_lngMinimumQty>0 then
					tmpMinQty=pcv_lngMinimumQty
				else
					tmpMinQty=1
				end if%>
				<input type="hidden" name="idproduct" value="<%=pIdProduct%>">
				<input type="hidden" name="quantity" value="<%=tmpMinQty%>">
			<%
			end if ' end 2a
			If (not pcf_OutStockPurchaseAllow) OR (scorderlevel = "2") OR ((pcf_WholesaleCustomerAllowed or scorderlevel = "1") and session("customerType")<>"1") then
		
				If tmp_showQty<>"1" then
				%>
					<input type="hidden" name="quantity" value="1">
					<input type="hidden" name="idproduct" value="<%=pidProduct%>">
				<%
				End if
			End if
		end if
        
        '// Display Custom Widgets
        %>
		</div>	
    <div class="pcClear"></div>
	</div>
	<script type=text/javascript>
		 $pc(function() {             
	        //$pc( "#qvtabs").tabs({active: <%=Clng(ActiveTabId)-1%>, heightStyle: "fill" });
            $pc( "#qvtabs").tab('show');
	    });
	</script>


<!--#include file="../includes/javascripts/pcValidateFormViewPrdQV.asp"-->
<% if pcv_Apparel="1" then %>
	<%if pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0 then%>
	<input type="hidden" name="idproduct" value="<%=pidProduct%>">
	<%end if%>
	<input type=hidden name="Apparel" value="1">
	<input type=hidden name="SavedList" value="">
	<input type=hidden name="SavedQtyList" value="">
	<script>
		var firstRun=0;
		function addEvent( obj, type, fn )
		{ 
			if ( obj.attachEvent ) { 
			obj['e'+type+fn] = fn; 
			obj[type+fn] = function(){obj['e'+type+fn]( window.event );} 
			obj.attachEvent( 'on'+type, obj[type+fn] ); 
			} else 
			obj.addEventListener( type, fn, false ); 
		} 

		function myInitFunction() {
			if (firstRun==0)
			{
			<% if pcv_ApparelRadio="1" then %>
			    new_GenRadioList(1,"",0);
			<%else%>
			    new_GenDropDown(1,"",0);
			<%end if%>
			new_CheckOptGroup(1,0);
			firstRun=1;
			}
		}
		
		addEvent(window,'load',myInitFunction);
	</script>
	<%if request.querystring("SubPrd")<>"" then
		tmpSubID=getUserInput(request.querystring("SubPrd"),0)
		if IsNumeric(tmpSubID) then
			call opendb()
			query="SELECT pcProd_Relationship FROM Products WHERE idproduct=" & tmpSubID
			set rs=connTemp.execute(query)
			if not rs.eof then
				tmpRelationship=rs("pcProd_Relationship")
				if (tmpRelationship<>"") and (instr(tmpRelationship,"_")>0) then
					tmpArray=split(tmpRelationship,"_")
					%>
					<script>
					<%if pcv_ApparelRadio="1" then%>
					function SetRadioValue(tmpID,tmpValue)
					{
						var els = document.additem.elements; 
						for(i=0; i<els.length; i++)
						{ 
							if ((els[i].name==tmpID) && (eval(els[i].value)==eval(tmpValue)) && (els[i].type=="radio"))
							{
								els[i].checked=true;
								break;
							}
						}
					}
					<%end if%>
					<%For mk=1 to ubound(tmpArray)%>
						<%if pcv_ApparelRadio="1" then%>
							SetRadioValue("idOption<%=mk%>","<%=tmpArray(mk)%>");
						<%else%>
							document.additem.idOption<%=mk%>.value=<%=tmpArray(mk)%>;
						<%end if%>
					<%Next%>
						firstRun=1;
						<% if pcv_ApparelRadio="1" then %>
                            new_GenRadioList(1,"",0);
                        <%else%>
                            new_GenDropDown(1,"",0);
                        <%end if%>
						new_CheckOptGroup(1,0);
					</script>
					<%
				end if
			end if
			set rs=nothing
		end if
	else
		if (request.querystring("index")<>"") or (pcv_intOptionGroupCount="1") then%>
		<script>
			firstRun=1;
			<% if pcv_ApparelRadio="1" then %>
				new_GenRadioList(1,"",0);
			<%else%>
				new_GenDropDown(1,"",0);
			<%end if%>
			new_CheckOptGroup(1,0);
		</script>
		<%end if
	end if%>
<% end if %>
<input name="FromQuickView" value="1" type="hidden" />
</form>
<!-- End Form -->
<%
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' CLOSE FORM																						//
'/////////////////////////////////////////////////////////////////////////////////////////////////////
Set RSlayout = nothing
Set rsIconObj = nothing
Set conlayout=nothing
call closedb()
%>
<script type=text/javascript>
	function stopRKey(evt)
	{
		var evt  = (evt) ? evt : ((event) ? event : null);
		var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
		if ((evt.keyCode == 13) && (node.type != "textarea") && (node.getAttribute("name") != "keyword")) { return false; }
	}
	document.onkeypress = stopRKey;
</script>

<div class="pcClear"></div>
      </div>
      <div class="modal-footer">
        <%
        prdLink = pcGenerateSeoProductLink(pDescription, "", pIdProduct) 
        %>
        <a href="<%=prdLink%>"><%=dictLanguage.Item(Session("language")&"_QV_11")%></a>
      </div>