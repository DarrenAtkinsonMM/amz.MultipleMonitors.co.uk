<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "bto_RePrdAddCharges.asp"
' This page handles configurable products.
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
<!--#include file="../pc/PrdAddChargesCode.asp"-->
<%
Response.Buffer = True

Dim pcCartIndex, f, pidProduct, pConfigSession, pDefaultPrice, pSavedQuantity

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

'Change paths of images
pcv_tmpNewPath="../pc/"

'// set variables specific to this action	
pConfigWishlistSession=getUserInput(request.querystring("idconf"),0)

'get these variables from querystring
pIdProductOrdered=request("idp")
pIDorder=request.QueryString("ido")



if pIdProductOrdered<>"0" and pIdProductOrdered<>"" then
query="SELECT idOrder,idProduct,quantity,unitPrice,xfdetails,idconfigsession FROM productsOrdered WHERE IdProductOrdered="&pIdProductOrdered&" and idOrder=" & pIdOrder & ";"
set rs=conntemp.execute(query)
	pIdOrder=rs("idOrder")
	pidProduct=rs("idProduct")
	pSavedQuantity=rs("quantity")	
	'pDefaultPrice=rs("unitPrice")	
	xstr=rs("xfdetails")
	pConfigSession=rs("idconfigsession")
else
	pIdOrder=request("ido")
	pidProduct=request("idproduct")
	pConfigSession=request("pre_idConfigSession")
end if
set rs=nothing

pPriceDefault=session("DefaultPrice" & pIDProduct)
pDefaultPrice=pPriceDefault
pCMPrice=session("CMPrice" & pIDProduct)
pCMWQD=session("CMWQD"  & pIDProduct)
pDefaultPrice = replace(pDefaultPrice, scCurSign, "")
pDefaultPrice = replace(pDefaultPrice, scDecSign&"00", "")

idquote=request("idquote")
pre_idConfigSession=request("pre_idConfigSession")
customertype=request("customertype")
pidcustomer=request("idcustomer")

if pConfigWishlistSession<>"" and pConfigWishlistSession<>"0" then
	query="SELECT stringCProducts,stringCValues,stringCCategories,pcconf_Quantity FROM configWishlistSessions WHERE idconfigWishlistSession=" & pConfigWishlistSession
else
	query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pConfigSession
end if
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

Dim stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory
if not rs.eof then
stringProducts = rs("stringCProducts")
stringValues = rs("stringCValues")
stringCategories = rs("stringCCategories")
if pConfigWishlistSession<>"" and pConfigWishlistSession<>"0" then
	pSavedQuantity=rs("pcconf_Quantity")
end if
else
stringProducts = ""
stringValues = ""
stringCategories = ""
end if
set rs=nothing
ArrProduct = Split(stringProducts, ",")
ArrValue = Split(stringValues, ",")
ArrCategory = Split(stringCategories, ",")


' check for discount per quantity
query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct
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

Dim pdiscountPerQuantity
if not rs.eof then
 pDiscountPerQuantity=-1
else
 pDiscountPerQuantity=0
end if

' gets item details from db

query="SELECT description, sku, configOnly, serviceSpec, price, btobprice, details, listprice, listHidden, imageurl, largeImageURL, Arequired, Brequired, stock, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3, noprices, sDesc,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty FROM products WHERE idProduct=" &pidProduct& " AND active=-1"
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

if rs.eof then 
  call closeDb()
response.redirect "msg.asp?message=88"
end if

Dim pserviceSpec, pDescription,pPrice,pBtoBPrice,pDetails,pListPrice,plistHidden,pimageUrl,pLgimageURL,pStock,pFormQuantity,pArequired,pBrequired

pDescription=rs("description")
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pDetails=rs("details")
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
pLgimageURL=rs("largeImageURL")
pArequired=rs("Arequired")
pBrequired=rs("Brequired")
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
pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
				if pcv_intHideBTOPrice<>"" then
				  else
				  pcv_intHideBTOPrice="0"
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

if pserviceSpec <> 0 then
	query="SELECT categories.idCategory, categories.categoryDesc, products.idProduct, products.description, configSpec_Charges.configProductCategory, configSpec_Charges.price, configSpec_Charges.Wprice, configSpec_Charges.multiSelect, products.weight FROM (configSpec_Charges INNER JOIN products ON configSpec_Charges.configProduct = products.idProduct) INNER JOIN categories ON configSpec_Charges.configProductCategory = categories.idCategory WHERE (((configSpec_Charges.specProduct)="&pIdProduct&") AND ((configSpec_Charges.cdefault)<>0)) ORDER BY configSpec_Charges.catSort, categories.idCategory, configSpec_Charges.prdSort,products.description;"

	set rsSSObj=conntemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rsSSObj=nothing
		'//close any connections
		
		'//redirect to error page
		call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	
	if NOT rsSSobj.eof then 
		Dim iAddDefaultPrice,	iAddDefaultWPrice
		iAddDefaultPrice=Cdbl(0)
		iAddDefaultWPrice=Cdbl(0)
		do until rsSSobj.eof
			if (rsSSobj("multiSelect")<>"") and (rsSSobj("multiSelect")<>-1) then
			iAddDefaultPrice=Cdbl(iAddDefaultPrice+rsSSobj("price"))
			iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+rsSSobj("Wprice"))
			end if 
		rsSSobj.moveNext
		loop
		set rsSSobj=nothing
		pPrice=Cdbl(pPrice+iAddDefaultPrice)
		pBtoBPrice=Cdbl(pBtoBPrice+iAddDefaultWPrice)
	end if
end if
if session("customerType")=1 AND pBtoBPrice>0 then
	pPrice=pBtoBPrice
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
<script type=text/javascript>
imagename = '';
function enlrge(imgnme) {
	lrgewin = window.open("about:blank","","height=200,width=200")
	imagename = imgnme;
	setTimeout('update()',500)
}
function win(fileName)
	{
	myFloater = window.open('','myWindow','scrollbars=auto,status=no,width=400,height=300')
	myFloater.location.href = fileName;
	}
	function viewWin(file)
	{
	myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=<%=iBTOPopWidth%>,height=<%=iBTOPopHeight%>')
	myFloater.location.href = file;
	}
function update() {
doc = lrgewin.document;
doc.open('text/html');
doc.write('<HTML><HEAD><TITLE>Enlarged Image<\/TITLE><\/HEAD><BODY bgcolor="white" onLoad="if (document.all || document.layers) window.resizeTo((document.images[0].width + 10),(document.images[0].height + 80))" topmargin="4" leftmargin="0" rightmargin="0" bottommargin="0"><table width=""' + document.images[0].width + '" height="' + document.images[0].height +'"border="0" cellspacing="0" cellpadding="0"><tr><td>');
doc.write('<IMG SRC="' + imagename + '"><\/td><\/tr><tr><td><form name="viewn"><A HREF="javascript:window.close()"><img  src="<%=pcv_tmpNewPath%>images/close.gif" align="right" border=0><\/a><\/td><\/tr><\/table>');
doc.write('<\/form><\/BODY><\/HTML>');
doc.close();
}
function checkDropdown(choice, option) {
	if (choice == 0) {
		alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ option + "<%=dictLanguage.Item(Session("language")&"_alert_6")%>.\n");
		return false;
			}
	return true;
	}
	function checkDropdowns(choice1, choice2, option1, option2) {
	if (choice1 == 0) {
		alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ option1 + "<%=dictLanguage.Item(Session("language")&"_alert_6")%>.\n");
		return false;
			}
	if (choice2 == 0) {
		alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ option2 + "<%=dictLanguage.Item(Session("language")&"_alert_6")%>.\n");
		return false;
			}
	return true;
	}
</script>
<script type=text/javascript>
function besubmit()
{
show_1.style.display = '';;
return (false);
}
</script>
<div id="pcMain">
<div class="pcMainContent">
	<% 
	Dim strMsg		
	strMsg=getUserInput(Request.QueryString("msg"),0)
	
	If strMsg <> "" Then 
	%><div class="pcShowContent"><div class="pcCPmessage"><%=strMsg %></div></div><% 
	End If 
	%>
	<form method="post" action="bto_instRePrdCharges.asp" name="additem" onSubmit="return chkR();" class="pcForms">
	<h1><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_1") & pDescription %></h1>
	<div class="pcShowContent">
	<div class="pcTable">
	<div class="pcTableRow">
		<% 
		if iBTOShowImage=0 then
			response.write "<div class=""pcTableColumn65"" valign=""top"">"
		else 
		%> 
			<div class="pcTableColumn100">
		<%
		end if 
		%>

		
		<!-- Short description, if any -->
		<% if psDesc <> "" then %>
			<div class="pcShowProductSDesc">
				<%=psDesc %>
			</div>
		<% end if %>
		<!-- End short description -->
		
		<!-- Product Price -->
		<%
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  Configurator Prices - Reconfig
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		pcs_AddChargesPrices
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END:  Configurator Prices - Reconfig
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\
		%>
		<!-- Product Price -->
		</div>
		<% if iBTOShowImage=0 then %>
			<div class="pcTableColumn35" valign="top"> 
			<%
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START:  Show Product Image (If there is one)
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			pcs_AddChargesImages
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END:  Show Product Image (If there is one)
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			%>
			</div>
		<% end if %>
	</div>
	
	<div class="pcTableRowFull">
		<b><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_13")%></b>	
				
		<!-- start product configuration -->
		<% If pserviceSpec=true then %>
					
		<% 				
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  javascript for calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcs_AddChargesCalculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END:  javascript for calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
		
		query="SELECT categories.categoryDesc, categories.idCategory, products.description, configSpec_Charges.catSort, configSpec_Charges.prdSort, configSpec_Charges.price, configSpec_Charges.Wprice, configSpec_Charges.cdefault, configSpec_Charges.configProductCategory, configSpec_Charges.multiSelect,configSpec_Charges.Notes FROM (configSpec_Charges INNER JOIN products ON configSpec_Charges.configProduct = products.idProduct) INNER JOIN categories ON configSpec_Charges.configProductCategory = categories.idCategory WHERE (((configSpec_Charges.multiSelect)<>3) AND ((configSpec_Charges.specProduct)="&pIdProduct&")) ORDER BY configSpec_Charges.catSort, categories.idCategory, configSpec_Charges.prdSort,products.description;"
				
		SET rsSSObj=Server.CreateObject("ADODB.RecordSet")
		SET rsSSObj=conntemp.execute(query)
		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rsSSObj=nothing
			'//close any connections
			
			'//redirect to error page
			call closeDb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		%>					
						

		<!-- START of Table -->	
		<%
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  Configuration Table
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcs_AddChargesTableReconfig
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END:  Configuration Table
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>
		<!-- END of Table -->

		
		<!-- Start Product Features -->
		<%
		if pnoprices=2 then%>
			<input name="curPrice" type="hidden" value="<%=scCurSign & money(pPrice) %>">
			<input name="currentValue0" type="HIDDEN" value="<%=pPriceDefault%>">
			<input name="jCnt" type="HIDDEN" value="<%=jCnt%>">
			<input name="total" type="hidden" value="None">
			<input name="GrandTotal" type="hidden" value="<%=scCurSign%><%=money(pCMPrice+pPriceDefault)%>">
			<input name="GrandTotalQD" type="hidden" value="<%=scCurSign%><%=money(pCMWQD+pPriceDefault)%>">
			<input name="CMPrice0" type="HIDDEN" value="<%=pCMPrice%>">
			<input name="CMWQD0" type="HIDDEN" value="<%=pCMWQD%>">
			<input name="CMWQD" type="hidden" value="<%=scCurSign & money(pCMWQD)%>">
			<input name="CMPrice" type="hidden" value="<%=scCurSign & money(pCMPrice)%>">
			<input name="CHGTotal" type="hidden" value="0">
		<%else
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Totals  
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			pcs_AddChargesTotals
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: Totals  
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		end if
		%>                              
		<!-- End Product Features -->                      
                       
		<% end if %>
		<!-- end product configuration -->
                  
				  
				  
		<%
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Discounts  
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcs_AddChargesDiscounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Discounts  
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>                      
		<input type="hidden" name="idproduct" value="<%response.write pidProduct%>">
		<input type="hidden" name="pcCartIndex" value="<%response.write pcCartIndex%>">					
		<%
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Disallow purchasing. Quote Submission only 
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcs_SubmissionReconfig
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Disallow purchasing. Quote Submission only 
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>
	</div>
	
	<div class="pcTableRowFull">
		<%if request("idquote")<>"" then
		if iBTOQuote=1 then%>
			<div>
				<button class="pcButton pcButtonSaveQuote" value="1" <%if pnoprices<2 and scHideDiscField<>"1" then%>onclick="javascript:besubmit(); return(false);"<%end if%> name="<%if pnoprices<2 and scHideDiscField<>"1" then%>iBTOQuoteA<%else%>iBTOQuote<%end if%>">
					<img src="<%=pcv_tmpNewPath%><%=rslayout("savequote")%>" alt="Save Quote" />
					<span class="pcButtonText">Save Quote</span>
				</button>
				<div id="show_1" style="display:none">
						<div class="pcFormItem">
							<div class="pcFormLabel"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_50")%></div>
							<div class="pcFormField"><input type="text" name="discountcode" value="" size="20"></div>
						</div>
						<%if pnoprices<2 and scHideDiscField<>"1" then%>
							<button class="pcButton pcButtonSubmit" value="1" id="iBTOQuote" name="iBTOQuote">
								<img src="<%=pcv_tmpNewPath%><%=rslayout("submit")%>" alt="Next Step" />
								<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
							</button>
						<%end if%>
				</div>
			</div>
		<%end if
		end if%>
						
		<input type=hidden name="idConfigWishlistSession" value="<%=pConfigWishlistSession%>">
		<input type="hidden" name="idorder" value="<%response.write pidorder%>">
		<input type="hidden" name="IdProductOrdered" value="<%response.write pIdProductOrdered%>">
		<input type="hidden" name="idquote" value="<%=request("idquote")%>">
		<input type="hidden" name="pre_idConfigSession" value="<%=pConfigSession%>">
		<input type="hidden" name="customertype" value="<%=customertype%>">
		<input type="hidden" name="idcustomer" value="<%=pidcustomer%>">
	</div>
	</div>
	</div>
	</form>
</div>
</div>

<script type=text/javascript>
	<%if scDecSign="," then
		pcv_CustomizedPrice=replace(pcv_CustomizedPrice,",",".")
		pcv_ItemDiscounts=replace(pcv_ItemDiscounts,",",".")
	end if%>
	Ctotal=<%=pcv_CustomizedPrice%>;
	QD1=<%=pcv_ItemDiscounts%>;
	var save_Ctotal=<%=pcv_CustomizedPrice%>;
	var save_QD1=<%=pcv_ItemDiscounts%>;
	GetItemLocation();
	New_GetField();
	<%if clng(ProQuantity)>1 then%>
		new_GenInforOnLoad();
	<%end if%>
	<%=pcv_ListForGenInfo%>
	New_calculateAll();
</script>
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
