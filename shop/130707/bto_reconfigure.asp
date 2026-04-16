<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "bto_Reconfigure.asp"
' This page is the review setup page for configurable products
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

Dim pcCartIndex, f, pidProduct, pConfigSession, pDefaultPrice, pSavedQuantity
Dim xrequired, xfieldCnt, reqstring, ProQuantity, pcv_strFormAction



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Determine if this Reconfigure Options from cart  - OR -  Reconfigure Options from a saved quote
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If request("idconf")="" Then

	'// set the form action
	pcv_strFormAction = "bto_instConfiguredPrd.asp"

	'get these variables from querystring
	pIdProductOrdered=request.QueryString("idp")
	pIDorder=request.QueryString("ido")

	'get these variables from db
	query="SELECT idOrder,idProduct,quantity,unitPrice,xfdetails,idconfigsession FROM productsOrdered WHERE idProductOrdered="&pIdProductOrdered&" and idOrder=" & pIdOrder & ";"
	set rs=conntemp.execute(query)
	pIdOrder=rs("idOrder")
	pidProduct=rs("idProduct")
	pSavedQuantity=rs("quantity")
	ProQuantity=pSavedQuantity
	pDefaultPrice=rs("unitPrice")
	xstr=rs("xfdetails")
	pConfigSession=rs("idconfigsession")
	set rs=nothing

	query="SELECT * FROM configSessions WHERE idconfigSession=" & pConfigSession													
	set rs=conntemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rs=nothing
		'//close any connections
		
		'//redirect to error page
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	query="SELECT idcustomer FROM orders where idorder=" & pIdOrder
	set rs1=connTemp.execute(query)
	pidcustomer=rs1("idcustomer")
	set rs1=nothing
	
Else

	'/////////////////////////////////////////////////////
	'// This is Reconfigure Options from a saved quote
	'/////////////////////////////////////////////////////
	
	'// set the form action
	pcv_strFormAction = "bto_instConfiguredPrd.asp"
	
	'// set variables specific to this action	
	pConfigWishlistSession=getUserInput(request.querystring("idconf"),0)
	
	'// set common variable to tmp values, we can merge them later	
	pIdProduct = getUserInput(request.QueryString("idProduct"),0)
	pDefaultPrice=getUserInput(request.querystring("price"),0)
	pDefaultPrice=replace(pDefaultPrice, scDecSign&"00", "")
	pSavedQuantity=1
	ProQuantity=pSavedQuantity

	
	query="SELECT discountcode,idcustomer FROM wishlist WHERE idconfigWishlistSession=" & pConfigWishlistSession
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	pdiscountcode=rs("discountcode")
	pidcustomer=rs("idcustomer")
	
	if pdiscountcode="0" then
		pdiscountcode=""
	end if
	
	if pdiscountcode<>"" then
		session("DCODE")=pdiscountcode
	else
		session("DCODE")=""
	end if
													
	query="SELECT pcconf_Quantity, xfdetails, stringProducts, stringValues, stringCategories, stringQuantity, stringPrice FROM configWishlistSessions WHERE idconfigWishlistSession=" & pConfigWishlistSession													
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	

End If

	customertype=0

	query="SELECT customers.customerType,customers.idCustomerCategory FROM customers WHERE idcustomer=" & pidcustomer & ";"
	SET rs1=Server.CreateObject("ADODB.RecordSet")
	SET rs1=conntemp.execute(query)
	if not rs1.eof then
		session("customerType")=rs1("customerType")
		customertype=rs1("customertype")
		idcustomerCategory=rs1("idcustomerCategory")
		if IsNull(idcustomerCategory) or idcustomerCategory="" then
			idcustomerCategory=0
		end if
	end if
	set rs1=nothing

	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory="&idcustomerCategory&";"
	SET rs1=Server.CreateObject("ADODB.RecordSet")
	SET rs1=conntemp.execute(query)
	if NOT rs1.eof then
		session("customerCategory")=rs1("idcustomerCategory")
		strpcCC_Name=rs1("pcCC_Name")
		session("customerCategoryDesc")=strpcCC_Name
		strpcCC_Description=rs1("pcCC_Description")
		session("customerCategoryType")=rs1("pcCC_CategoryType")
		if session("customerCategoryType")="ATB" then
			session("ATBCustomer")=1
			session("ATBPercentage")=rs1("pcCC_ATB_Percentage")
			intpcCC_ATB_Off=rs1("pcCC_ATB_Off")
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
	set rs1=nothing
	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Determine if this Reconfigure Options from a save order  - OR -  Reconfigure Options from a saved quote
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="../pc/configurePrdCode.asp"-->
<%

Dim xstr, stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory

if not rs.eof then
	if Request("idconf")<>"" Then 
		xstr = rs("xfdetails") 
	end if
	stringProducts = rs("stringProducts")
	stringValues = rs("stringValues")
	stringCategories = rs("stringCategories")
	stringQuantity=rs("stringQuantity")
	stringPrice=rs("stringPrice")
	If request("idconf")<>"" Then
		pSavedQuantity=rs("pcconf_Quantity")
	end if
else
	xstr = ""
	stringProducts = ""
	stringValues = ""
	stringCategories = ""
	stringQuantity=""
	stringPrice=""
end if

ArrProduct = Split(stringProducts, ",")
ArrValue = Split(stringValues, ",")
ArrCategory = Split(stringCategories, ",")
ArrQuantity = Split(stringQuantity, ",")
ArrPrice = Split(stringPrice, ",")

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

query="SELECT description, sku, configOnly, serviceSpec, price, btobprice, details, listprice, listHidden, imageurl, largeImageURL, stock, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3, noprices, sDesc,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcProd_multiQty,pcProd_MaxSelect,pcProd_BackOrder,pcProd_ShipNDays,pcProd_HideSKU FROM products WHERE idProduct=" &pidProduct& " AND active=-1"
SET rs=Server.CreateObject("ADODB.RecordSet")
SET rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	
	call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then 
	SET rs=nothing
	
  call closeDb()
	response.redirect "msg.asp?message="&Server.Urlencode(dictLanguage.Item(Session("language")&"_viewPrd_2") )  
end if

Dim pserviceSpec, pDescription, psDesc, pPrice, pBtoBPrice, pDetails, pListPrice, plistHidden, pimageUrl, pLgimageURL, pStock, pFormQuantity

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
pnoprices=rs("noprices")
if isNULL(pnoprices) or pnoprices="" then
	pnoprices=0
end if
if pIDorder<>"" then
	pnoprices=0
end if
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

pHideSKU=rs("pcProd_HideSKU")
if IsNull(pHideSKU) or pHideSKU="" then
	pHideSKU=0
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
if session("customerType")=1 AND pBtoBPrice>0 then
	pPrice=pBtoBPrice
end if
%>

<% HRcolor="#e1e1e1" %>
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
<!--#include file="../includes/pcServerSideValidation.asp"-->
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
<!-- Main Table --> 
			<h1><% response.write(bto_dictLanguage.Item(Session("language")&"_configurePrd_1")&pDescription) %></h1>
			<div class="pcShowContent">
			<div class="pcTable">
				<div class="pcTableRow">
					<div <% if iBTOShowImage=1 then %>class="pcTableColumn100"<% else %>class="pcTableColumn65"<% end if %>>		
						<% if psDesc <> "" then %>
							<div class="pcShowProductSDesc">
								<%=psDesc %>
							</div>
						<% end if %>	
					</div>
					<%
					if iBTOShowImage=0 and trim(pimageUrl)<>"no_image.gif" then %>
					<div class="pcTableColumn35">			
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  Show Product Image (If there is one)
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
					pcs_ProductImage
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  Show Product Image (If there is one)
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					%>
					</div>
					<% end if %>					
				</div>
								
				<div class="pcTableRow">
					<div <% if iBTOShowImage=1 then %>class="pcTableColumn100"<% else %>class="pcTableColumn65"<% end if %> valign="bottom">
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
					pcs_BTOReconfigTable
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
				</div>
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
				' START:  X Fields - Reconfigure/ Quote
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				if len(pConfigWishlistSession)>0 then
					pcs_XFieldsQuote
				else
					pcs_XFieldsReconfigure
				end if
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  X Fields - Reconfigure/ Quote
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%>
				<!--end of xfields -->
				<%
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: Disallow purchasing. Quote Submission only
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				pcs_AdmBTOPurchasing
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  Disallow purchasing. Quote Submission only
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%> 
				<div class="pcTableRowFull">
						<%if request("idquote")<>"" then
						pcv_strFuntionCall = "cdDynamic"
						
						if BTOCharges=0 then
							if iBTOQuote=1 then
							%>
								<div>
									<button class="pcButton pcButtonSaveQuote" value="1" <%if pnoprices<2 and scHideDiscField<>"1" then%>onclick="javascript:besubmit(); return(false);"<%else%><%if xrequired="1" then%>onClick="javascript: if (checkproqty(document.additem.quantity)) {if (chkR()) {if (<%=pcv_strFuntionCall%>(<%=reqstring%>,2)) { return(true) };}}; return false;"<%end if%><%end if%> name="<%if pnoprices<2 and scHideDiscField<>"1" then%>iBTOQuoteA<%else%>iBTOQuote<%end if%>">
										<img src="<%=pcv_tmpNewPath%><%if pnoprices="0" or len(f)=0 then%><%=rslayout("savequote")%><%else%><%=rslayout("pcLO_requestQuote")%><%end if%>" alt="<%if pnoprices="0" or len(f)=0 then%>Save Quote<%else%>Request a quote<%end if%>" />
										<span class="pcButtonText"><%if pnoprices="0" or len(f)=0 then%>Save Quote<%else%>Request a quote<%end if%></span>
									</button>
									<div id="show_1" style="display:none">
											<div class="pcFormItem">
												<div class="pcFormLabel"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_50")%></div>
												<div class="pcFormField"><input type="text" name="discountcode" value="" size="20"></div>
											</div>
											<%if pnoprices<2 and scHideDiscField<>"1" then%>
												<button class="pcButton pcButtonSubmit" value="1" id="iBTOQuote" name="iBTOQuote" <%if xrequired="1" then%>onClick="javascript: if (checkproqty(document.additem.quantity)) {if (chkR()) {if (<%=pcv_strFuntionCall%>(<%=reqstring%>,2)) { return(true) };}}; return false;"<%end if%>>
													<img src="<%=pcv_tmpNewPath%><%=rslayout("submit")%>" alt="Next Step" />
													<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
												</button>
											<%end if%>
									</div>
								</div>
							<%
							end if
						end if
						end if
						%>
				</div>
			</div>
			</div>
<% 
NextStep=getUserInput(request("N"),0)

if NextStep="" then
	NextStep="0"
end if
%>

<% 
'// If we are on the Quote Page these values are different
if len(f)>0 then 
%>
<input type=hidden name="NextStep" value="<%=NextStep%>">
<% else %>
<input type=hidden name="idConfigWishlistSession" value="<%=pConfigWishlistSession%>">
<input type=hidden name="NextStep" value="1">
<% end if %>

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
			<input type="hidden" name="idorder" value="<%response.write pidorder%>">
			<input type="hidden" name="IdProductOrdered" value="<%response.write pIdProductOrdered%>">
			<input type="hidden" name="idquote" value="<%=request("idquote")%>">
			<input type="hidden" name="pre_idConfigSession" value="<%=pConfigSession%>">
			<input type="hidden" name="customertype" value="<%=customertype%>">
			<input type="hidden" name="idcustomer" value="<%=pidcustomer%>">
			<input type="hidden" name="idproduct" value="<%=pidProduct%>">
			</form>
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