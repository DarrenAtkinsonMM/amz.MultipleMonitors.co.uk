<!DOCTYPE html>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<%
Response.Buffer = True

dim total
total = Cint(0)

dim pWeight
pWeight=getUserInput(request("w"),0)

Dim rsCustObj

query="SELECT name, lastName, customerCompany, phone, address, zip, stateCode, state, city, countryCode, email FROM customers WHERE idCustomer=" & session("idCustomer")
Set rsCustObj=Server.CreateObject("ADODB.Recordset")
Set rsCustObj=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsCustObj=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
	
CustomerName=rsCustObj("name")& " " & rsCustObj("lastName")
CustomerCompany=rsCustObj("customerCompany")
pAddress=rsCustObj("address")
pcity=rsCustObj("city")
pStateCode=rsCustObj("stateCode")
if pStateCode="" then
	pStateCode=rsCustObj("state")
end if
pzip=rsCustObj("zip")
pcountry=rsCustObj("countryCode")

customerPhone=rsCustObj("phone")
customerEmail=rsCustObj("email")
set rsCustObj=nothing

%>
<html>
<head>
	<title>Edited Quote - Printable Version</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","pcStorefront.css")%>" />
</head>
<body> 
<div id="pcMain">
		<div class="pcMainContent">
			<div class="pcTable">
			<div class="pcTableRow">
				<div style="width:25%;"><img src="<%=pcf_getImagePath("../pc/catalog",scCompanyLogo)%>"></div>
				<div style="width:50%;text-align:center;">
					<b><%=scCompanyName%></b><br>
					<%=scCompanyAddress%><br>
					<%=scCompanyCity%>, <%=scCompanyState%>&nbsp;<%=scCompanyZip%><br>
					<hr noshade align="center" color=SILVER>
					<%=scStoreURL%>
				</div>
				<div style="width:25%;">&nbsp;</div>
			</div>
			<div class="pcTableRowFull">
				<div class="pcSpacer">&nbsp;</div>
			</div>
			<div class="pcTableRow">
			<div style="width:50%;padding-left:0;">	
			<!-- Start: Billing Info -->
			<div class="invoice">
			<strong><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_2")%></strong>:<br>	
				<%=CustomerName%>
				<br>
				<% if CustomerCompany<>"" then 
					response.write CustomerCompany&"<BR>"
				end if %>
				<%=pAddress%>
				<br>
				<% if pAddress2<>"" then 
					response.write pAddress2&"<BR>"
				end if %>
				<% response.write pcity&", "&pStateCode&" "&pzip %>
				<% if pCountryCode <> scShipFromPostalCountry then
					response.write "<BR>" & pCountryCode
				end if %>
				<br><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_3") & CustomerPhone%>
				<br><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_4") & CustomerEmail%>
			</div>
			</div>
			<div style="width:48%;float:right;padding-right:0;">
				<div class="invoice" align="right">
					<%
					' Retrieve quote number
					pidconfigWishlistSession=getUserInput(request.QueryString("idconf"),0)
					if not validNum(pidconfigWishlistSession) then
						call closeDb()
					   	response.redirect "Custquotesview.asp"
					end if
					
 					query="SELECT idquote FROM wishlist WHERE idconfigWishlistSession=" & pidconfigWishlistSession
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if (not rs.eof) and (err.number = 0) then
						pidquote=rs("idquote")
					end if
					set rs = nothing
					if pidquote <> "" then
					%>
                    <strong><%response.write bto_dictLanguage.Item(Session("language")&"_Custquotesview_4")%> #: <%=pidquote%></strong> <br />
					<% end if %>
					<%response.write bto_dictLanguage.Item(Session("language")&"_Custquotesview_18")%>
					<%
					' Retrieve quote date
 					query="SELECT dtCreated FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
						if (not rs.eof) and (err.number = 0) then
							pqdate=rs("dtCreated")
						end if
					set rs = nothing
						if pqdate <> "" then
							response.write ShowDateFrmt(pqdate)
						else
							response.write "N/A"
						end if
					%>
					<%
					' Retrieve submitted quote date
 					query="SELECT QSubmit,QDate FROM wishlist WHERE idconfigWishlistSession=" & pidconfigWishlistSession
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
						if (not rs.eof) and (err.number = 0) then
							pQSubmit=rs("QSubmit")
							if IsNull(pQSubmit) or pQSubmit="" then
								pQSubmit=0
							end if
							pSqdate=rs("QDate")
						end if
					set rs = nothing
						if pSqdate <> "" then %>
						<br>
						Submitted on: <%= ShowDateFrmt(pSqdate)%>
					<% end if	%>
				</div>
			</div>
			</div>
			<div class="pcTableRowFull">
				<div class="pcSpacer">&nbsp;</div>
			</div>
			<% 
			query="SELECT discountcode FROM Wishlist WHERE idconfigWishlistSession=" & pidconfigWishlistSession
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			if (not rs.eof) and (err.number = 0) then
				pdiscountcode=rs("discountcode")
				if pdiscountcode="0" then
					pdiscountcode=""
				end if
			else
				call closeDb()
				response.write("Error in printableQuote (100): " & err.number & "--" & err.description)
				response.end
			end if
			
			query="SELECT idProduct, dtCreated, fPrice, dPrice, pcconf_Quantity, pcconf_QDiscount, stringProducts, stringValues, stringCategories, stringQuantity, stringPrice, stringCProducts, stringCValues, stringCCategories, xfdetails FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			Dim pIdProduct, pdtCreated, pxfedtails, pfPrice, stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory
			pIdProduct=rs("idProduct")
			QIDProduct=pIdProduct
			pdtCreated=rs("dtCreated")
			pfPrice=rs("fPrice")
			dPrice=rs("dPrice")
			ItemsDiscounts=dPrice
			total = total+pfPrice
			pQuantity=rs("pcconf_Quantity")
			if (pQuantity<>"") then
			else
			pQuantity="1"
			end if
			pQty=pQuantity
			QDiscounts=rs("pcconf_QDiscount")
			if (QDiscounts<>"") then
			else
			QDiscounts=0
			end if
			pstringProducts = rs("stringProducts")
			pstringValues = rs("stringValues")
			pstringCategories = rs("stringCategories")
			pstringQuantity = rs("stringQuantity")
			pstringPrice = rs("stringPrice")
			pstringCProducts = rs("stringCProducts")
			pstringCValues = rs("stringCValues")
			pstringCCategories = rs("stringCCategories")
			pxfdetails=rs("xfdetails") 
			
			ArrProduct = Split(pstringProducts, ",")
			ArrValue = Split(pstringValues, ",")
			ArrCategory = Split(pstringCategories, ",")
			ArrQuantity = Split(pstringQuantity, ",")
			ArrPrice = Split(pstringPrice, ",")
			ArrCProduct = Split(pstringCProducts, ",")
			ArrCValue = Split(pstringCValues, ",")
			ArrCCategory = Split(pstringCCategories, ",")

			query="SELECT sku, description,noprices,price,btoBPrice FROM Products WHERE idProduct=" & trim(pidProduct)
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			psku=rs("sku")
			pname=rs("description")
			pnoprices=rs("noprices")
			if pnoprices<>"" then
			else
			pnoprices=0
			end if
			if pQSubmit=3 then
				pnoprices=0
			end if
			pcv_price=rs("price")
			if session("customertype")=1 then
				if rs("btoBPrice")>"0" then
					pcv_price=rs("btoBPrice")
				end if
			end if
			set rs=nothing
			%> 
        	<div class="pcTableRow"> 
			        <div class="pcTable" style="padding:0;">
					<div class="invoice" style="padding: 0 0 8px 0;">
	                	<div class="pcTableHeader"> 
							<div class="pcCustOrdInvoice-QTY"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_6")%></div>
							<div class="pcCustOrdInvoice-SKU"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_1")%></div>
							<div class="pcCustOrdInvoice-Desc"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_2")%></div> 

							<%if pnoprices<2 then%> 
								<div class="pcCustOrdInvoice-Total"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_3")%></div>
							<%end if%> 
						</div>
						<div class="pcTableRow">
							<div class="pcCustOrdInvoice-QTY"><%=pQuantity%></div>
							<div class="pcCustOrdInvoice-SKU"><%=psku%></div>
							<div class="pcCustOrdInvoice-Desc"><b><%=pname%></b></div>
							<div class="pcCustOrdInvoice-Total"><%=scCurSign & money(pcv_price*pQuantity)%></div>
						</div>
						<div class="pcTableRow">
							<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
							<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
							<div class="pcCustOrdInvoice-Desc" style="background:#F5F5F5"> 
						<% if ArrProduct(0)="na" then %> 
							<%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_4")%>
							</div>
							<div class="pcCustOrdInvoice-Total" style="background:#F5F5F5">&nbsp;</div>
						</div>
						<% else %> 
							<%response.write bto_dictLanguage.Item(Session("language")&"_viewcart_1")%>
							<%response.write "</div>"%>
							<div class="pcCustOrdInvoice-Total" style="background:#F5F5F5">&nbsp;</div>
						<%response.write "</div>"%>
						<%
						'calculate
						for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)

							pcv_intIdProduct = pcf_GetParentId(ArrProduct(i))

							query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="& pcv_intIdProduct &"))"
							set rsObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							if NOT rsObj.eof then
											
								query="SELECT displayQF FROM configSpec_Products WHERE configProduct="& pcv_intIdProduct & " AND specProduct=" & pIdProduct
								set rsObj1=conntemp.execute(query)								
								if (not rsObj1.eof) then
								%> 
								<div class="pcTableRow">
									<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
									<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
									<div class="pcCustOrdInvoice-Desc" style="background:#F5F5F5">
										<%=rsObj("categoryDesc")%>: <%=rsObj("description")%> (<%=rsObj("sku")%>)
										<%if rsObj1("displayQF")=True then%> 
											- QTY: <%=ArrQuantity(i)%> 
										<%end if%> 
									</div>
									<div class="pcCustOrdInvoice-Total" style="background:#F5F5F5">
											<%if pnoprices<2 then%>
												<%=scCurSign & money((ArrValue(i)+(ArrPrice(i)*(ArrQuantity(i)-1)))*pQuantity)%> 
											<%end if%>
									</div>
								</div> 
								<%end if%> 
							<%end if
							set rsObj1=nothing
							set rsObj=nothing
						next 
						end if %>
<%
if pnoprices<2 then
if ItemsDiscounts<>0 then%> 
<div class="pcTableRow">
	<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
	<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
	<div class="pcCustOrdInvoice-Desc"><div align="right">Items Discounts:</div></div> 
	<div class="pcCustOrdInvoice-Total">
		<%if pnoprices<2 then%> 
			<%=scCurSign & money(ItemsDiscounts)%> 
		<%end if%>
	</div>
</div>
<%end if
end if%> 
<%if pnoprices<2 then
if QDiscounts<>0 then%> 
<div class="pcTableRow">
	<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
	<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
	<div class="pcCustOrdInvoice-Desc"><div align="right">Quantity Discounts:</div></div> 
	<div class="pcCustOrdInvoice-Total">
		<%if pnoprices<2 then%> 
			<%=scCurSign & money(-1*QDiscounts)%> 
		<%end if%>
	</div>
</div> 
<%end if
end if%> 
<% if ArrCProduct(0)<>"na" then%> 
<div class="pcTableRowFull">
	<div class="pcSpacer">&nbsp;</div>
</div> 
<div class="pcTableRow">
	<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
	<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
	<div class="pcCustOrdInvoice-Desc" style="background:#F5F5F5"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_5")%></div>
	<%if pnoprices<2 then%>
		<div class="pcCustOrdInvoice-Total" style="background:#F5F5F5">&nbsp;</div>
	<%end if%>
</div>
<% 
'calculate
for i = lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
	query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
	set rsObj=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
								
	if (not rsObj.eof) and (err.number = 0) then %> 
	<div class="pcTableRow">
		<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
		<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
		<div class="pcCustOrdInvoice-Desc" style="background:#F5F5F5"><%=rsObj("categoryDesc")%>: <%=rsObj("description")%> (<%=rsObj("sku")%>)</div>
		<div class="pcCustOrdInvoice-Total" style="background:#F5F5F5">
			<%if pnoprices<2 then%> 
				<%if (CDbl(ArrCValue(i))<>0) then%> 
					<%=scCurSign & money(ArrCValue(i))%> 
				<%end if%> 
			<%end if%>
		</div> 
	</div> 
<% else
	call closeDb()
	response.write("Error in printableQuote (527): " & err.number & "--" & err.description)
	response.end 
	end if%> 
<% set rsObj=nothing
next 
%> 
<%end if%> 
<% if trim(pxfdetails)<>"" then
	xfieldsarray=split(pxfdetails,"||")
	for i=lbound(xfieldsarray)to (UBound(xfieldsarray)-1)
		xfields=split(xfieldsarray(i),"|")
		query="SELECT xfield FROM xfields WHERE idxfield="&xfields(0)
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		if (not rs.eof) and (err.number = 0) then
			xfielddesc=rs("xfield")
			set rs=nothing
%> 
<div class="pcTableRow">
	<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
	<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
	<div class="pcCustOrdInvoice-Desc"><%response.write xfielddesc&": "&xfields(1)%></div>
	<%if pnoprices<2 then%> 
	<div class="pcCustOrdInvoice-Total">&nbsp;</div> 
	<%end if%> 
</div> 
<% else
	call closeDb()
	response.write("Error in printableQuote (561): " & err.number & "--" & err.description)
	response.end 
end if%> 
<% next
end if

pSubTotal=pfPrice
%> 
<%'discounts by categories
CatDiscTotal=0

query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
set rs1=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs1=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

CatSubDiscount=0

Do While (not rs1.eof) and (CatSubDiscount=0)
	CatSubQty=0
	CatSubTotal=0
	CatSubDiscount=0

	query="select idproduct from categories_products where idcategory=" & rs1("IDCat") & " and idproduct=" & QIDProduct
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if not rs.eof then
		CatSubQty=CatSubQty+pQty
		CatSubTotal=CatSubTotal+pfPrice
	end if

if CatSubQty>0 then

query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rs1("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
set rs2=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs2=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if not rs2.eof then

 	' there are quantity discounts defined for that quantity 
 	pDiscountPerUnit=rs2("pcCD_discountPerUnit")
 	pDiscountPerWUnit=rs2("pcCD_discountPerWUnit")
 	pPercentage=rs2("pcCD_percentage")
	pbaseproductonly=rs2("pcCD_baseproductonly")

 	if session("customerType")<>1 then  'customer is a normal user
 		if pPercentage="0" then 
 			CatSubDiscount=pDiscountPerUnit*CatSubQty
		else
			CatSubDiscount=(pDiscountPerUnit/100) * CatSubTotal
		end if
	else  'customer is a wholesale customer
 		if pPercentage="0" then 
 			CatSubDiscount=pDiscountPerWUnit*CatSubQty
		else
			CatSubDiscount=(pDiscountPerWUnit/100) * CatSubTotal
		end if
	end if

	 	
end if
end if

	CatDiscTotal=CatDiscTotal+CatSubDiscount
	rs1.MoveNext
loop

'// Round the Category Discount to two decimals
if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
	CatDiscTotal = Round(CatDiscTotal,2)
end if
							
if CatDiscTotal>0 then
%> 
<div class="pcTableRow">
	<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
	<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
	<div class="pcCustOrdInvoice-Desc"><b><%if pnoprices<2 then%><%=dictLanguage.Item(Session("language")&"_catdisc_2")%><%end if%></b></div>
	<%if pnoprices<2 then%> 
		<div class="pcCustOrdInvoice-Total"><%response.write scCurSign & "-" & money(CatDiscTotal)%></div> 
	<%end if%> 
</div>
<%end if%> 
<% if (pnoprices<2) and (pdiscountcode<>"") and (pdiscountcode<>"-") then%> 
<div class="pcTableRow">
	<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
	<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
	<div class="pcCustOrdInvoice-Desc"> 
	<%if pDiscountError="" then    
		discountTotal=Cdbl(0)
		if pPriceToDiscount>0 or ppercentageToDiscount>0 then 
			discountTotal=pPriceToDiscount + (ppercentageToDiscount*(pfPrice)/100)
		end if
		pSubTotal=pfPrice - discountTotal
		if pSubTotal<0 then
			pSubTotal=0
		end if
	%> 
	<b>Discount code: <%=pdiscountcode%></b><br> 
	Details: <%=pDiscountDesc%><br>
	Amount:	<%=scCurSign & money(-1*discountTotal)%>
	<%
	else
		if pDiscountError<>"-" then%> 
			<b>Discount code: <%=pdiscountcode%></b><br> 
			Error: <%=pDiscountError%>
		<%end if
	end if%>
	</div>
	<%if pnoprices<2 then%> 
		<div class="pcCustOrdInvoice-Total">
			<%if (discountTotal>"0") and (discountTotal<>"") then%> 
			<%=scCurSign & money(-1*discountTotal)%> 
			<%end if%> 
		</div>
	<%end if%> 
</div> 
<%end if %> 
<%if CatDiscTotal>0 then
pSubTotal=pSubTotal - Round(CatDiscTotal,2)
if pSubTotal<0 then
	pSubTotal=0
end if
end if%> 
<%if pnoprices<2 then %> 
<div class="pcTableRow">
	<div class="pcCustOrdInvoice-QTY">&nbsp;</div>
	<div class="pcCustOrdInvoice-SKU">&nbsp;</div>
	<div class="pcCustOrdInvoice-Desc"><div align="right"><b>Total:</b></div></div>
	<div class="pcCustOrdInvoice-Total"><b><%response.write scCurSign & money(pSubTotal)%></b></div>
</div> 
<%end if%> 
<div class="pcTableRowFull">
	<div class="pcSpacer">&nbsp;</div>
</div>
</div>
</div>
</div>
</div>
</div>
</div>
</body>
</html>
<%
	set rs1=nothing
	set rs2=nothing
	call closedb()
%>