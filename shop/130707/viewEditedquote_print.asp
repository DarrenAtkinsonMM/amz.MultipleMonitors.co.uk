<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=10%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
pidCustomer = request.QueryString("idcustomer")
pidconfigWishlistSession = request.QueryString("idconfigWishlistSession")

dim total
total = Cint(0)

dim pWeight
pWeight=request("w")

Dim rsCustObj

query="SELECT name, lastName, customerCompany, phone, address, zip, stateCode, state, city, countryCode, email,customerType FROM customers WHERE idCustomer=" & pidCustomer
Set rsCustObj=Server.CreateObject("ADODB.Recordset")
Set rsCustObj=connTemp.execute(query)

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
customertype=rsCustObj("customertype")
set rsCustObj=nothing

%>
<!DOCTYPE html>
<html>
<head>
<title>Saved Quote - Printable Version</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="inc_header.asp"-->
</head>
<body style="background-image: none;"> 
<table border="0" cellpadding="4" cellspacing="0" align="center" width="100%"> 
	<tr>
	<td valign="top" width="100%">
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
		<tr valign="middle">
		<td width="18%" align="left"><img src="../pc/catalog/<%=scCompanyLogo%>"></td> 
		<td width="39%" height="71" class="invoiceNob">
			<div align="center">
			<b><%=scCompanyName%></b><br>
			<%=scCompanyAddress%><br>
			<%=scCompanyCity%>, <%=scCompanyState%>&nbsp;<%=scCompanyZip%><br>
			<hr width=100 noshade align="center" color=SILVER>
			<%=scStoreURL%>
			</div>
		</td> 
		<td width="43%" valign="bottom"></td>
		</tr>
		<tr>
		<td colspan="3">&nbsp;</td>
		</tr>
		<tr>
		<td colspan="3" valign="top">
			<table width="100%" cellpadding="5" cellspacing="0" class="invoiceNob">
			<tr> 
			<td class="invoiceNob">
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
			</td>
			<td width="50%" class="invoiceNob">
			<div align="right">
			<table width="50%" align="right" cellpadding="5" cellspacing="0">
				<tr> 
					<td class="invoiceNob">
					<div align="right">
					Created on: 
					<%
					' Retrieve quote date
 					query="SELECT dtCreated FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
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
 					query="SELECT QDate FROM wishlist WHERE idconfigWishlistSession=" & pidconfigWishlistSession
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
						if (not rs.eof) and (err.number = 0) then
							pSqdate=rs("QDate")
						end if
					set rs = nothing
						if pSqdate <> "" then %>
						<br>
						Submitted on: <%= ShowDateFrmt(pSqdate)%>
					<% end if	%>
					</div>
					</td>
				</tr>
			</table>
			</div>
			</td>
			</tr>
			</table>
		</td>
		</tr>
		<tr> 
		<td colspan="3">&nbsp;</td>
		</tr>
		</table> 
		
<% query="SELECT discountcode FROM Wishlist WHERE idconfigWishlistSession=" & pidconfigWishlistSession
								set rs=server.CreateObject("ADODB.Recordset")
								set rs=conntemp.execute(query)
								if (not rs.eof) and (err.number = 0) then
									pdiscountcode=rs("discountcode")
									if pdiscountcode="0" then
										pdiscountcode=""
									end if
								else
									
									response.write("Error in printableQuote (100): " & err.number & "--" & err.description)
									response.end
								end if
								
								query="SELECT idProduct, dtCreated, fPrice, dPrice, pcconf_Quantity, pcconf_QDiscount, stringProducts, stringValues, stringCategories,stringQuantity,stringPrice, stringCProducts, stringCValues, stringCCategories, xfdetails FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession

								set rs=server.CreateObject("ADODB.Recordset")
								set rs=conntemp.execute(query)
								if err.number <> 0 then
									
									call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in printableQuote: "&err.description) 
								end if
								Dim pIdProduct, pdtCreated, pxfedtails, pfPrice, stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory
								pIdProduct=rs("idProduct")
								QIDProduct=pIdProduct
								pdtCreated=rs("dtCreated")
								pfPrice=rs("fPrice")
								dPrice=rs("dPrice")
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
								ItemsDiscounts=dPrice
								total = total+pfPrice
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
								if err.number <> 0 then
									
									call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in printableQuote: "&err.description) 
								end if
								psku=rs("sku")
								pname=rs("description")
								pnoprices=0
								pcv_price=rs("price")
								if customertype=1 then
									if rs("btoBPrice")>"0" then
										pcv_price=rs("btoBPrice")
									end if
								end if
								set rs=nothing
								%> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td>
						<table width="100%" cellpadding="5" cellspacing="0" border="1" class="invoice"> 
						<tr bgcolor="<%=AColor%>"> 
							<td nowrap class="invoice">
								<b><font color="<%=AFColor%>"> 
								<%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_6")%> 
								</font></b>
							</td> 
							<td nowrap class="invoice">
								<b><font color="<%=AFColor%>"> 
								<%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_1")%> 
								</font></b>
							</td> 
							<td colspan="2" class="invoice">
								<b><font color="<%=AFColor%>">
								<%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_2")%> 
								</font></b>
							</td>
							<td nowrap class="invoice">
								<div align="right">
									<b><font color="<%=AFColor%>"> 
									<%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_3")%> 
									</font></b>
								</div>
							</td> 
							</tr> 

							<tr> 
								<td class="invoice" valign="top" align="right"><%=pQuantity%></td> 
								<td class="invoice" valign="top" align="left" width="14%"><%=psku%></td> 
								<td colspan="2" align="left" valign="top" class="invoice"><b><%=pname%></b></td>
								<td class="invoice" nowrap>
									<div align="right"><%=scCurSign & money(pcv_price*pQuantity)%></div>
								</td> 
							</tr> 
					<tr> 
					<td class="invoice">&nbsp;</td> 
					<td class="invoice">&nbsp;</td> 
					<td colspan="3" class="invoice">
						<table width="100%" border="0" cellspacing="2" cellpadding="0" bgcolor="#FFFFCC" class="invoiceBto"> 
							<% if ArrProduct(0)="na" then %> 
								<tr> 
									<td colspan="2" class="invoiceNob">
										<%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_4")%>
									</td> 
								</tr> 
							<% else %> 
								<tr> 
									<td colspan="2" class="invoiceNob">
										<%response.write bto_dictLanguage.Item(Session("language")&"_viewcart_1")%>
									</td> 
								</tr> 
								<%for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
							
								query="SELECT pcProd_ParentPrd FROM Products WHERE idproduct=" & ArrProduct(i) & " AND pcProd_ParentPrd>0;"
								set rs99=connTemp.execute(query)
								sp_ParentPrd=0
								if not rs99.eof then
									sp_ParentPrd=rs99("pcProd_ParentPrd")
								else
									sp_ParentPrd=ArrProduct(i)
								end if
								set rs99=nothing
								query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&sp_ParentPrd&"))" 
							
								set rsObj=conntemp.execute(query)
								if NOT rsObj.eof then
								
									query="SELECT displayQF FROM configSpec_Products WHERE configProduct=" & sp_ParentPrd & " AND specProduct=" & pIdProduct 
									set rsObj1=conntemp.execute(query)
									if (not rsObj1.eof) then
										customizations=customizations+ArrValue(i) %>
										<tr valign="top" class="invoice"> 
											<td width="90%" class="invoiceNob">
												<%=rsObj("categoryDesc")%>: <%=rsObj("description")%> (<%=rsObj("sku")%>)
												<%if rsObj1("displayQF")=True then%> 
													- QTY: <%=ArrQuantity(i)%> 
												<%end if%> 
											</td>
											<td width="10%" class="invoiceNob">
												<div align="right">
														<%=scCurSign & money((ArrValue(i)+(ArrPrice(i)*(ArrQuantity(i)-1)))*pQuantity)%>
												</div>
											</td>
										</tr> 
									<%end if%> 
								<%end if
								set rsObj1=nothing
								set rsObj=nothing
								next 
							end if %> 
						</table> 
					</td>
					</tr> 
<%
if pnoprices<2 then
if ItemsDiscounts<>0 then%> 
<tr> 
	<td class="invoice">&nbsp;</td> 
	<td class="invoice">&nbsp;</td> 
	<td colspan="2" class="invoice"><div align="right">Items Discounts:</div></td> 
	<%if pnoprices<2 then%> 
	<td class="invoice">
	<div align="right">
		<%if pnoprices<2 then%> 
			<%=scCurSign & money(ItemsDiscounts)%> 
		<%end if%>
	</div>
	</td> 
	<%end if%> 
</tr> 
<%end if
end if%> 
<%
if pnoprices<2 then
if QDiscounts<>0 then
%> 
<tr> 
	<td class="invoice">&nbsp;</td> 
	<td class="invoice">&nbsp;</td> 
	<td colspan="2" class="invoice"><div align="right">Quantity Discounts:</div></td>
	<%if pnoprices<2 then%>
	<td class="invoice">
		<div align="right">
		<%if pnoprices<2 then%> 
			<%=scCurSign & money(-1*QDiscounts)%> 
		<%end if%>
		</div>
	</td> 
	<%end if%> 
</tr> 
<%end if
end if%> 
<% if ArrCProduct(0)<>"na" then%> 
<tr>
	<td class="invoice">&nbsp;</td> 
	<td class="invoice">&nbsp;</td> 
	<td colspan="3" class="invoice"> 
	<table width="100%" border="0" cellspacing="2" cellpadding="0" bgcolor="#FFFFCC" class="invoiceBto">
	<tr> 
	<td colspan="2" class="invoiceNob"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_5")%></td> 
	</tr> 
	<% 
	'calculate
	for i = lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
	query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
	set rsObj=conntemp.execute(query)
	if (not rsObj.eof) and (err.number = 0) then %> 
	<tr valign="top"> 
	<td width="90%" class="invoiceNob"><%=rsObj("categoryDesc")%>: <%=rsObj("description")%> (<%=rsObj("sku")%>)</td>
	<td width="10%" class="invoiceNob"><%if pnoprices<2 then%> 
		<div align="right">
		<%if (CDbl(ArrCValue(i))<>0) then%> 
		<%=scCurSign & money(ArrCValue(i))%> 
		<%end if%> 
		</div> 
		<%end if%>
	</td> 
	</tr> 
	<%end if%> 
	<% set rsObj=nothing
	next 
	%> 
	</table> 
	</td> 
</tr> 
<%end if%> 
<% if trim(pxfdetails)<>"" then
	xfieldsarray=split(pxfdetails,"||")
	for i=lbound(xfieldsarray)to (UBound(xfieldsarray)-1)
		xfields=split(xfieldsarray(i),"|")
		query="SELECT xfield FROM xfields WHERE idxfield="&xfields(0)
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if (not rs.eof) and (err.number = 0) then
			xfielddesc=rs("xfield")
			set rs=nothing
%> 
<tr> 
<td class="invoice">&nbsp;</td> 
<td class="invoice">&nbsp;</td> 
<td colspan="2" class="invoice">
	<table width="100%" border="0" cellspacing="0" cellpadding="0"> 
	<tr> 
	<td class="invoiceNob"><%response.write xfielddesc&": "&xfields(1)%>
	</td> 
	</tr> 
	</table>
</td> 
<%if pnoprices<2 then%> 
	<td class="invoice">&nbsp;</td> 
<%end if%> 
</tr> 
<% else
	
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

CatSubDiscount=0

Do While (not rs1.eof) and (CatSubDiscount=0)
	CatSubQty=0
	CatSubTotal=0
	CatSubDiscount=0

	query="select idproduct from categories_products where idcategory=" & rs1("IDCat") & " and idproduct=" & QIDProduct
	set rs=connTemp.execute(query)
	if not rs.eof then
		CatSubQty=CatSubQty+pQty
		CatSubTotal=CatSubTotal+pfPrice
	end if

if CatSubQty>0 then

query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rs1("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
set rs2=conntemp.execute(query)

if not rs2.eof then

 	' there are quantity discounts defined for that quantity 
 	pDiscountPerUnit=rs2("pcCD_discountPerUnit")
 	pDiscountPerWUnit=rs2("pcCD_discountPerWUnit")
 	pPercentage=rs2("pcCD_percentage")
	pbaseproductonly=rs2("pcCD_baseproductonly")

 	if customertype<>1 then  'customer is a normal user
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
<tr> 
<td class="invoice">&nbsp;</td> 
<td class="invoice">&nbsp;</td> 
<td colspan="2" class="invoice"><b><%if pnoprices<2 then%><%=dictLanguage.Item(Session("language")&"_catdisc_2")%><%end if%></b></td>
<%if pnoprices<2 then%> 
<td class="invoice" nowrap><div align="right"><%response.write scCurSign & "-" & money(CatDiscTotal)%></div></td> 
<%end if%> 
</tr> 
<%end if%> 
<% if (pnoprices<2) and (pdiscountcode<>"") and (pdiscountcode<>"-") then%> 
<tr> 
<td class="invoice">&nbsp;</td> 
<td class="invoice">&nbsp;</td> 
<td colspan="2" class="invoice">
	<table width="100%" border="0" cellspacing="0" cellpadding="0"> 
<%	if pDiscountError="" then    
		discountTotal=Cdbl(0)
		if pPriceToDiscount>0 or ppercentageToDiscount>0 then 
			discountTotal=pPriceToDiscount + (ppercentageToDiscount*(pfPrice)/100)
		end if
		pSubTotal=pfPrice - discountTotal
		if pSubTotal<0 then
			pSubTotal=0
		end if
%> 
	<tr> 
	<td colspan=2 class="invoiceNob"><b>Discount code: <%=pdiscountcode%></b><br> 
		Details: <%=pDiscountDesc%></font>
	</td> 
	</tr> 
	<tr> 
	<td width="100%" class="invoiceNob">Amount:</td>
	<td nowrap class="invoice"><%=scCurSign & money(-1*discountTotal)%></td> 
	</tr> 
<%
else
if pDiscountError<>"-" then%> 
	<tr> 
	<td colspan=2 class="invoiceNob"><b>Discount code: <%=pdiscountcode%></b><br> 
		<font color=#FF0000>Error: <%=pDiscountError%></font>
	</td> 
	</tr> 
<%
end if
end if%> 
	</table>
</td> 
<%if pnoprices<2 then%> 
<td class="invoice" nowrap>
	<div align="right">
	<%if (discountTotal>"0") and (discountTotal<>"") then%> 
	<%=scCurSign & money(-1*discountTotal)%> 
	<%end if%> 
	</div>
</td> 
<%end if%> 
</tr> 
<%end if %>
<%if CatDiscTotal>0 then
pSubTotal=pSubTotal - Round(CatDiscTotal,2)
if pSubTotal<0 then
	pSubTotal=0
end if
end if%> 
<%if pnoprices<2 then %> 
<tr> 
	<td class="invoice" colspan="4"><div align="right">	<b>Total:</b></div></td>
	<td class="invoice" nowrap><div align="right"><b><%response.write scCurSign &  money(pSubTotal) %></b></div></td> 
</tr> 
<%end if%> 
</table>
</td> 
</tr> 
</table> 
</td> 
</tr> 
<tr> 
<td valign="top">&nbsp;</td> 
</tr> 
</table> 
</body>
</html>