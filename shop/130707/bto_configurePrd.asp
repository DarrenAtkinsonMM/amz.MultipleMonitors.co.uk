<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "bto_configurePrd.asp"
' This page is the setup page for configurable products
'
'/////////////////////////////////////////////////////////////////
' NOTES:														//
'																//
' The "bto_configurePrdCode.asp" include will hold the routines that 
' display the product information. Each segment of product 
' information has been divided into zone.
'
'/////////////////////////////////////////////////////////////////
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim pIdProduct
Dim xrequired, xfieldCnt, reqstring, ProQuantity, pcv_strFormAction

pIdOrder=request("ido")



	query="SELECT customers.customerType,customers.idCustomerCategory FROM customers INNER JOIN orders ON customers.idcustomer=orders.idcustomer WHERE idorder=" & pIdOrder & ";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if not rs.eof then
		session("customerType")=rs("customerType")
		idcustomerCategory=rs("idcustomerCategory")
		if IsNull(idcustomerCategory) or idcustomerCategory="" then
			idcustomerCategory=0
		end if
	end if
	set rs=nothing

	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory="&idcustomerCategory&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then
		session("customerCategory")=rs("idcustomerCategory")
		strpcCC_Name=rs("pcCC_Name")
		session("customerCategoryDesc")=strpcCC_Name
		strpcCC_Description=rs("pcCC_Description")
		session("customerCategoryType")=rs("pcCC_CategoryType")
		if session("customerCategoryType")="ATB" then
			session("ATBCustomer")=1
			session("ATBPercentage")=rs("pcCC_ATB_Percentage")
			intpcCC_ATB_Off=rs("pcCC_ATB_Off")
			if intpcCC_ATB_Off="Retail" then
				session("ATBPercentOff")=0
			else
				session("ATBPercentOff")=1
			end if
		else
			session("ATBCustomer")=0
			session("ATBPercentage")=0
			session("ATBPercentOff")=0
		end if
	end if
	set rs=nothing%>
<!--#include file="../pc/configurePrdCode.asp"-->
<%
Response.Buffer = True


Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

'Change paths of images
pcv_tmpNewPath="../pc/"

'-------------------------------
' declare local variables
'-------------------------------

pIdProduct = getUserInput(request.QueryString("idProduct"),0)
pcv_strFormAction = "bto_instConfiguredPrd.asp"

' check for discount per quantity
'if customer is retail, check if there are discounts with retail <> 0
Dim VardiscGo
VardiscGo=0

if session("customerType")=1 then
	query="SELECT discountPerWUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerWUnit>0"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if rs.eof then
		VardiscGo=1
	end if
	SET rs=nothing
else
	query="SELECT discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerUnit>0"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if rs.eof then
		VardiscGo=1
	end if
	SET rs=nothing
end if

query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct
SET rs=Server.CreateObject("ADODB.RecordSet")
SET rs=conntemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

Dim pDiscountPerQuantity
if not rs.eof then
	if VardiscGo=0 then
		pDiscountPerQuantity=-1
	end if
else
	pDiscountPerQuantity=0
end if 

SET rs=nothing

' gets item details from db
query="SELECT description, sku, configOnly, serviceSpec, price, btobprice, details, listprice, listHidden, imageurl, largeImageURL, stock, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3, noprices, sDesc,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcProd_multiQty,pcProd_MaxSelect FROM products WHERE idProduct=" &pidProduct& " AND active=-1"
SET rs=Server.CreateObject("ADODB.RecordSet")
SET rs=conntemp.execute(query)

if err.number<>0 then
    call LogErrorToDatabase()
    set rs=nothing		
    call closeDb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if rs.eof then 
	Set rs = nothing	
	call closeDb()
    session("message") = dictLanguage.Item(Session("language")&"_viewPrd_2")
    response.redirect "msgb.asp?back=1"     
end if

Dim pserviceSpec, pDescription, psDesc, pPrice, pBtoBPrice, pDetails, pListPrice, plistHidden, pimageUrl, pLgimageURL, pStock, pFormQuantity
Dim pcv_lngMultiQty,pcv_MaxSelect

pDescription=rs("description")
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pPrice=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
pBtoBPrice=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
pDetails=rs("details")
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
	if pimageUrl="" then
		pimageUrl="no_image.gif"
	end if
pLgimageURL=rs("largeImageURL")
pStock=rs("stock")
pEmailText=rs("emailText")
pFormQuantity=rs("formQuantity")
pnoshipping=rs("noshipping")
pcustom1=rs("custom1")
pcontent1=rs("content1")
pcustom2=rs("custom2")
pcontent2=rs("content2")
pcustom3=rs("custom3")
pcontent3=rs("content3")
'pnoprices=rs("noprices")
'if isNULL(pnoprices) or pnoprices="" then
	pnoprices=0
'end if
psDesc=rs("sDesc")
	if trim(psDesc)="" then
		psDesc=pDetails
	end if
pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
if isNULL(pcv_intHideBTOPrice) or pcv_intHideBTOPrice="" then
	pcv_intHideBTOPrice="0"
end if
pcv_intQtyValidate=rs("pcprod_QtyValidate")
if isNull(pcv_intQtyValidate) OR pcv_intQtyValidate="" then
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

pcv_MaxSelect=rs("pcProd_MaxSelect")
if isNull(pcv_MaxSelect) OR pcv_MaxSelect="" then
	pcv_MaxSelect="0"
end if

SET rs=nothing
	
if pserviceSpec <> 0 then
	%>
	<!--#include file="../pc/pcGetPrdPrices.asp"-->
	<%
	pPrice=dblpcCC_Price
	pBtoBPrice=pPrice
end if

query="SELECT specProduct FROM configSpec_Charges WHERE specProduct="&pIdProduct
SET rs=Server.CreateObject("ADODB.RecordSet")
SET rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
BTOCharges=0
if NOT rs.eof then
	BTOCharges=1
end if
SET rs=nothing
	
ProQuantity=getUserInput(request("qty"),0)
if ProQuantity="" then
	if pcv_lngMinimumQty>"0" then
		ProQuantity=pcv_lngMinimumQty
	else
		ProQuantity="1"
	end if
end if	
%>
<!--#include file="AdminHeader.asp"-->
<link type="text/css" rel="stylesheet" href="../pc/css/pcStorefront.css" />
<style>
#pcBTOhideTopPrices {
	display: none;
}
#pcMain .transparentField
{
	background-color: transparent !important;
	border:none !important;
	box-shadow: none !important;
	outline:none !important;
	padding: 0 !important;
}
</style>
<!-- 58eed21a4b2adf0316e95c5c4ee68f13 -->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<script type="text/javascript" src="../includes/formatNumber.js"></script>
<script type="text/javascript" src="../includes/javascripts/ConfigurePrdFuncs.js"></script>
<%
If statusCM="1" OR scCM=1 Then
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Conflict Management Module
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
call ConflictManager()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Conflict Management Module
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If
%>
<div id="pcMain">
<div class="pcMainContent">
		<% 
		Dim strMsg		
		strMsg=getUserInput(Request.QueryString("msg"),0)
		
		If strMsg <> "" Then 
		%><div class="pcShowContent"><div class="pcErrorMessage"><%=strMsg %></div></div><% 
		End If 
		%>
		
		<form method="post" action="<%=pcv_strFormAction%>" name="additem" onSubmit="return chkR();" class="pcForms">
			<h1><% response.write(bto_dictLanguage.Item(Session("language")&"_configurePrd_1")&pDescription) %></h1>
			<div class="pcShowContent">
			<div class="pcTable">
				<div class="pcTableRow">
					<%
					if iBTOShowImage=1 then
						response.write "<div class=""pcTableColumn100"">"
					else 
					%> 
					<div class="pcTableColumn65" valign="bottom"> 
					<% 
					end if 
					%>		
				
					<% if psDesc <> "" then %>
						<div class="pcShowProductSDesc">
							<%=psDesc %>
						</div>
					<% end if %>					

					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  Configurator Prices
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
					pcs_BTOPrices
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  Configurator Prices
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					%>
					</div>
					<% if iBTOShowImage=0 then %>
					<div class="pcTableColumn35" valign="top">			
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  Show Product Image (If there is one)
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
					pcs_ProductImage
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  Show Product Image (If there is one)
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					%>
					<% end if %>
					</div>									
				</div>
				
				<div class="pcTableRowFull">
						
				<!--#include file="../includes/javascripts/pcFunctionsConfigurePrd.asp"-->
				
				<!-- Configurable Items for this product -->
				<%
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START:  Check Quantity Discounts
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				pcs_CheckQTYDiscount
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  Check Quantity Discounts
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				
				If pserviceSpec <> 0 then 'START: If its configurable
				
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  javascript for calculations
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pcs_BTOJavaCalculations	
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  javascript for calculations
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  product configuration table
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pcs_BTOConfigTable
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  product configuration table
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				End If 'END: If its configurable
				%>
						
				</div>
			
				<div class="pcTableRowFull">
					<hr>
				</div>
				
				<%if pnoprices<2 then%>
				<div class="pcTableRowFull">
					<div id="pcBTOfloatPrices">
						<div class="pcTable">
						<div class="pcTableRowFull">
							<div class="pcTableColumn60">
								<b><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_12")%></b>
							</div>
							<div class="pcTableColumn1"></div>
							<div class="pcTableColumn39">
							</div>
						</div>
						</div>
					</div>
				</div>
				<%end if%>
				
				<!-- Start Product Features -->
				<% if pnoprices=2 then %>
					<input name="curPrice" type="hidden" value="<%=scCurSign & money(pPrice) %>">
					<input name="currentValue0" type="hidden" value="<%=pPrice%>">
					<input name="jCnt" type="hidden" value="<%=jCnt%>">
					<input name="total" type="hidden" value="0">
					<input name="GrandTotal" type="hidden" value="<%=scCurSign & money(pPrice)%>">
					<input name="Discounts" type="hidden" value="<%=scCurSign & money(0)%>">
					<input name="QDiscounts" type="hidden" value="<%=scCurSign & money(0)%>">
					<input name="QDiscounts0" type="hidden" value="<%=scCurSign & money(0)%>">
					<input name="TLcurPrice" type="hidden" value="<%=scCurSign & money(pPrice) %>">
					<input name="TLPriceDefault" type="hidden">
					<input name="TLtotal" type="hidden" value="0">
					<input name="TLGrandTotal" type="hidden" value="<%=scCurSign & money(pPrice)%>">
					<input name="CMDefault" type="hidden">
					<input name="CMWQD" type="hidden">
					<input name="UGrandTotal" type="hidden" value="<%=scCurSign & money(pPrice)%>">
					<input name="TotalWithQD" type="hidden" value="<%=scCurSign & money(0)%>">
				<%else%>
				<div class="pcTableRowFull">		
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  Display Totals 
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pcs_DisplayTotals
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  Display Totals
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~					
					%>				
				<div class="pcTableRowFull">
				<% end if %>
				<!-- End Product Features -->
			
				<!-- start discounts -->
				<%
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START:  Discounts
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				if pnoprices<2 then
					pcs_BTODiscounts
				end if
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  Discounts
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%>
				<!-- end discounts -->
				<%if pnoprices<2 then%>
				<div class="pcTableRowFull">
					<hr>
				</div>
				<%end if%>
				<!-- If xfields are present add here -->
				<%
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START:  X Fields
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				pcs_BTOXFields
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  X Fields
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%>
				<!--end of xfields -->
				<%
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: Disallow purchasing. Quote Submission only
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				pcs_BTOPurchasing
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  Disallow purchasing. Quote Submission only
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%>
				
			</div>
			</div>
			
				<script type=text/javascript>
					<%if scDecSign="," then
						pcv_CustomizedPrice=replace(pcv_CustomizedPrice,",",".")
						pcv_ItemDiscounts=replace(pcv_ItemDiscounts,",",".")
					end if%>
					var pcQDiscountType="<%=pcQDiscountType%>";
					Ctotal=<%=pcv_CustomizedPrice%>;
					QD1=<%=pcv_ItemDiscounts%>;
					var save_Ctotal=<%=pcv_CustomizedPrice%>;
					var save_QD1=<%=pcv_ItemDiscounts%>;
					GetItemLocation();
					New_GetField();
					<%if clng(ProQuantity)>1 then%>
						new_GenInforOnLoad();
					<%end if%>
					<% if pcv_HaveRules=1 then %>
					PresetValues();
					<% end if %>
					<%=pcv_ListForGenInfo%>
					var MaxSelectMsg1="<%=bto_dictLanguage.Item(Session("language")&"_configurePrd_21")%>";
					var MaxSelectMsg1a="<%=bto_dictLanguage.Item(Session("language")&"_configurePrd_21a")%>";
					var TotalMaxSelect=<%=pcv_MaxSelect%>;
					New_calculateAll();
				</script>
				<input type="hidden" name="savequantity" value="<%=ProQuantity%>">
			<!--#include file="../includes/javascripts/pcValidateFormViewPrd.asp"-->
			<input type="hidden" name="idOrder" value="<%=pIdOrder%>">
			<input type="hidden" name="idproduct" value="<%=pidProduct%>">
			</form>
			<%
			'// Cached Page or History Page: Refresh()
			pcs_BTOPageReLoader
			%>
</div>
</div>
<!--#include file="AdminFooter.asp"-->
<%
   
call clearLanguage()

set conntemp = Nothing
set rs = Nothing
set pIdproduct = Nothing
set pDescription = Nothing
set pDetails = Nothing
set pListPrice = Nothing
set pImageUrl	= Nothing
set pWeight = Nothing
set arrCategories	= Nothing

%>