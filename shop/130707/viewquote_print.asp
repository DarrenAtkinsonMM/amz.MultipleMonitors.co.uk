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

query="SELECT name, lastName, customerCompany, phone, address, zip, stateCode, state, city, countryCode, email FROM customers WHERE idCustomer=" & pidCustomer
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
                        <div style="padding-left: 64%;">
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
                            query="SELECT idquote, QDate FROM wishlist WHERE idconfigWishlistSession=" & pidconfigWishlistSession
                            set rs=server.CreateObject("ADODB.Recordset")
                            set rs=conntemp.execute(query)
                            if (not rs.eof) and (err.number = 0) then
                                pIdQuote=rs("idquote")
								pSqdate=rs("QDate")
                            end if
                            set rs = nothing
                            if pSqdate <> "" then %>
                                <br>
                                Submitted on: <%= ShowDateFrmt(pSqdate)%>
                            <% end if %>
                            <br />
                            Quote ID: <%=pIdQuote%>
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
								
								query="SELECT idProduct, dtCreated, fPrice, dPrice, pcconf_Quantity, pcconf_QDiscount, stringProducts, stringValues, stringCategories, stringQuantity, stringPrice, stringCProducts, stringCValues, stringCCategories, xfdetails  FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
								set rs=server.CreateObject("ADODB.Recordset")
								set rs=conntemp.execute(query)
								if err.number <> 0 then
									pcvErrDescription = err.description
									set rs=nothing
									
									call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in printableQuote: "&pcvErrDescription) 
								end if
								Dim pIdProduct, pdtCreated, pxfedtails, pfPrice, stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory
								pIdProduct=rs("idProduct")
								QIDProduct=pIdProduct
								pdtCreated=rs("dtCreated")
								pfPrice=rs("fPrice")
								dPrice=rs("dPrice")
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
		
								query="SELECT sku, description,noprices FROM Products WHERE idProduct=" & trim(pidProduct)
								set rs=conntemp.execute(query)
								if err.number <> 0 then
									pcvErrDescription = err.description
									set rs=nothing
									
									call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in printableQuote: "&pcvErrDescription) 
								end if
								psku=rs("sku")
								pname=rs("description")
								'pnoprices=rs("noprices")
								pnoprices=0
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
							<%if pnoprices<2 then%> 
							<td nowrap class="invoice">
								<div align="right">
									<b><font color="<%=AFColor%>"> 
									<%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_3")%> 
									</font></b>
								</div>
							</td> 
							<%end if%> 
							</tr> 

							<tr> 
								<td class="invoice" valign="top" align="right"><%=pQuantity%></td> 
								<td class="invoice" valign="top" align="left" width="14%"><%=psku%></td> 
								<td colspan="2" align="left" valign="top" class="invoice"><b><%=pname%></b></td>

						<%if pnoprices<2 then%> 
						<!--#include file="../pc/checkDiscount.asp"--> 
						<%
							discountcheck=0
							if pDiscountError="" then
								discountcheck=1
							end if
							itemsDiscounts=0
							
						for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					
						query="SELECT pcProd_ParentPrd FROM Products WHERE idproduct=" & ArrProduct(i) & " AND pcProd_ParentPrd>0;"
						set rs99=connTemp.execute(query)
						sp_ParentPrd=0
						if not rs99.eof then
							sp_ParentPrd=rs99("pcProd_ParentPrd")
						end if
						set rs99=nothing

						query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i) & " OR IDProduct=" & sp_ParentPrd & ";"			
						set rs99=connTemp.execute(query)

						TempDiscount=0
						do while not rs99.eof
										QFrom=rs99("quantityFrom")
										QTo=rs99("quantityUntil")
										DUnit=rs99("discountperUnit")
										QPercent=rs99("percentage")
										DWUnit=rs99("discountperWUnit")
										if (DWUnit=0) and (DUnit>0) then
										DWUnit=DUnit
										end if
										

										TempD1=0
										if (clng(ArrQuantity(i)*pQuantity)>=clng(QFrom)) and (clng(ArrQuantity(i)*pQuantity)<=clng(QTo)) then
										if QPercent="-1" then
										if session("customerType")=1 then
										TempD1=ArrQuantity(i)*pQuantity*ArrPrice(i)*0.01*DWUnit
										else
										TempD1=ArrQuantity(i)*pQuantity*ArrPrice(i)*0.01*DUnit
										end if
										else
										if session("customerType")=1 then
										TempD1=ArrQuantity(i)*pQuantity*DWUnit
										else
										TempD1=ArrQuantity(i)*pQuantity*DUnit
										end if
										end if
										end if
										TempDiscount=TempDiscount+TempD1
										rs99.movenext
						loop
						itemsDiscounts=ItemsDiscounts+TempDiscount
						next			
						
						if ItemsDiscounts>0 then
							pfPrice=pfPrice+ItemsDiscounts
						else
							pfPrice=pfPrice-ItemsDiscounts
						end if
						
						if QDiscounts<>0 then
							pfPrice=pfPrice+QDiscounts
						end if
						
						Charges=0
						for i = lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
							UPrice=ArrCValue(i)
							Charges=Charges+UPrice
						next
						if Charges>0 then
							pfPrice=pfPrice-Charges	
						else
							pfPrice=pfPrice-Charges
						end if%> 
						<td class="invoice" nowrap>
						<% customizations=0
						for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
							
							query="SELECT pcProd_ParentPrd FROM Products WHERE idproduct=" & ArrProduct(i) & " AND pcProd_ParentPrd>0;"
							set rs99=connTemp.execute(query)
							sp_ParentPrd=0
							if not rs99.eof then
								sp_ParentPrd=rs99("pcProd_ParentPrd")
							else
								sp_ParentPrd=ArrProduct(i)
							end if
							set rs99=nothing
							query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & sp_ParentPrd & ";"
						
							set rsQ=connTemp.execute(query)
							tmpMinQty=1
							if not rsQ.eof then
								tmpMinQty=rsQ("pcprod_minimumqty")
								if IsNull(tmpMinQty) or tmpMinQty="" then
									tmpMinQty=1
								else
									if tmpMinQty="0" then
										tmpMinQty=1
									end if
								end if
							end if
							set rsQ=nothing
							tmpDefault=0
						
							query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & sp_ParentPrd & " AND cdefault<>0;"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then
								tmpDefault=rsQ("cdefault")
								if IsNull(tmpDefault) or tmpDefault="" then
									tmpDefault=0
								else
									if tmpDefault<>"0" then
									 	tmpDefault=1
									end if
								end if
							end if
							set rsQ=nothing

							query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&sp_ParentPrd&"))" 
							set rsObj=conntemp.execute(query)
						
							if NOT rsObj.eof then
								if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
									if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
										if tmpDefault=1 then
											UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
										else
											UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
										end if
									else
										UPrice=0
									end if
									UPrice= UPrice + ArrValue(i)
									customizations=customizations+UPrice
								end if
							end if
						next
						
						pfPrice=pfPrice-(pQuantity*customizations)%> 
						<div align="right"><%=money(pfPrice)%></div>
					</td> 
					<%end if%> 
					</tr> 
					<tr> 
					<td class="invoice">&nbsp;</td> 
					<td class="invoice">&nbsp;</td> 
					<td colspan="<%if pnoprices<2 then%>3<%else%>2<%end if%>" class="invoice">
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
						<% 
						'calculate
						for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
						
							query="SELECT pcProd_ParentPrd FROM Products WHERE idproduct=" & ArrProduct(i) & " AND pcProd_ParentPrd>0;"
							set rs99=connTemp.execute(query)
							sp_ParentPrd=0
							if not rs99.eof then
								sp_ParentPrd=rs99("pcProd_ParentPrd")
							else
								sp_ParentPrd=ArrProduct(i)
							end if
							set rs99=nothing
							query="SELECT displayQF FROM configSpec_Products WHERE configProduct="& sp_ParentPrd &" and specProduct=" & pidProduct 
							set rsQ=server.CreateObject("ADODB.RecordSet") 
							set rsQ=conntemp.execute(query)

							btDisplayQF=rsQ("displayQF")
							set rsQ=nothing

							query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & sp_ParentPrd & ";"
							set rsQ=connTemp.execute(query)
							tmpMinQty=1
							if not rsQ.eof then
								tmpMinQty=rsQ("pcprod_minimumqty")
								if IsNull(tmpMinQty) or tmpMinQty="" then
									tmpMinQty=1
								else
									if tmpMinQty="0" then
										tmpMinQty=1
									end if
								end if
							end if
							set rsQ=nothing
							tmpDefault=0

							query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & sp_ParentPrd & " AND cdefault<>0;"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then
								tmpDefault=rsQ("cdefault")
								if IsNull(tmpDefault) or tmpDefault="" then
									tmpDefault=0
								else
									if tmpDefault<>"0" then
									 	tmpDefault=1
									end if
								end if
							end if
							set rsQ=nothing
	
							query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&sp_ParentPrd&"))" 
							set rsObj=conntemp.execute(query)
							if NOT rsObj.eof then
										
								query="SELECT displayQF FROM configSpec_Products WHERE configProduct="& sp_ParentPrd & " AND specProduct=" & pIdProduct 
								set rsObj1=conntemp.execute(query)
								if (not rsObj1.eof) then
									customizations=customizations+ArrValue(i) %> 

							<tr valign="top" class="invoice"> 
							<td width="90%" class="invoiceNob">
								<%=rsObj("categoryDesc")%>: <%=rsObj("description")%> (<%=rsObj("sku")%>)
								<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%> - QTY: <%=ArrQuantity(i)%><%end if%> 
							</font>
							</td>
							<td width="10%" class="invoiceNob">
                            <div align="right">
								<%if pnoprices<2 then%> 
							<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
								if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
									if tmpDefault=1 then
										UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
									else
										UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
									end if
								else
									UPrice=0
								end if
								pfPrice=pfPrice+cdbl((ArrValue(i)+UPrice)*pQuantity) %> 
								<%=scCurSign & money((ArrValue(i)+UPrice)*pQuantity)%>
							<%else
								if tmpDefault=1 then%>
									Included
								<%end if%>
							<%end if%> 
							<% end if %>
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
if ItemsDiscounts<>0 then
pfprice=pfprice-ItemsDiscounts%> 
<tr> 
<td class="invoice">&nbsp;</td> 
<td class="invoice">&nbsp;</td> 
<td colspan="2" class="invoice"><div align="right">Items Discounts:</div></td> 
<%if pnoprices<2 then%> 
<td class="invoice">
	<div align="right">
	<%if pnoprices<2 then%> 
	<%=scCurSign & money(-1*ItemsDiscounts)%> 
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
pfPrice=pfPrice-QDiscounts
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
<% if ArrCProduct(0)<>"na" then
pfprice=pfprice+Charges %> 
<tr> 
<td class="invoice">&nbsp;</td> 
<td class="invoice">&nbsp;</td> 
<td colspan="2" class="invoice"> 
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
<% else
	
	response.write("Error in printableQuote (527): " & err.number & "--" & err.description)
	response.end 
	end if%> 
<% set rsObj=nothing
next 
%> 
	</table> 
</td> 
<%if pnoprices<2 then%> 
<td class="invoice" nowrap>
	<%if Charges<>0 then%>
	<div align="right"><%=scCurSign & money(Charges)%></div>
	<%end if%>
</td> 
<%end if%> 
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
set rs1=nothing


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