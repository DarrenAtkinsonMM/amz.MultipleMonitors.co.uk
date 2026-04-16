<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true %>
<% pageTitle="Reply to a Customer's Quote" %>
<% Section="genRpts" %>
<%PmAdmin=10%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"--> 
<%
dim total
total = Cint(0)

Dim rsCustObj

mySQL="SELECT idcustomer, idproduct, idconfigWishlistSession, qdate,qsubmit FROM wishlist WHERE idquote=" & request("idquote")
set rs=server.CreateObject("ADODB.RecordSet")							
set rs=conntemp.execute(mySQL)
					
if err.number <> 0 then
	pcvErrDescription = err.description
	set rs = nothing	
	call closeDb()
    response.redirect "techErr.asp?error="& Server.Urlencode("Error in CreateReplyQuote: "& pcvErrDescription) 
end if
	
idcustomer=rs("idcustomer")
idproduct=rs("idproduct")
idconf=rs("idconfigWishlistSession")
qdate=rs("qdate")
qsubmit=rs("qsubmit")
idquote=request("idquote")
dim qSubmitDate
if scDateFrmt="DD/MM/YY" then
	qSubmitDate=(day(qdate)&"/"&month(qdate)&"/"&year(qdate))
else
	qSubmitDate=(month(qdate)&"/"&day(qdate)&"/"&year(qdate))
end if

mySQL="SELECT name, lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, shippingaddress,shippingcity,shippingstatecode,shippingstate,shippingcountrycode,shippingzip FROM customers WHERE idCustomer=" & idCustomer
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(mySQL)
CustomerName=rs("name")& " " & rs("lastName")
CustomerCompany=rs("CustomerCompany")
phone=rs("phone")
pemail=rs("email")
pAddress=rs("address")
pcity=rs("city")
pStateCode=rs("stateCode")
if pStateCode="" then
	pStateCode=rs("state")
end if
pzip=rs("zip")
pcountry=rs("countryCode")
pshippingaddress=rs("shippingaddress")
pshippingcity=rs("shippingcity")
pshippingStateCode=rs("shippingStateCode")
pshippingState=rs("shippingState")
pshippingCountryCode=rs("shippingCountryCode")
pshippingZip=rs("shippingZip")
storeAdminEmail=""
storeAdminEmail=storeAdminEmail & "The quote #" & idquote & " was submitted to " & scCompanyName & " on " & qSubmitDate & "." & vbcrlf & vbcrlf
storeAdminEmail=storeAdminEmail & "CUSTOMER DETAILS" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & vbcrlf
storeAdminEmail=storeAdminEmail & "Name: " & CustomerName & vbcrlf
if CustomerCompany<>"" then
storeAdminEmail=storeAdminEmail & "Company: " & CustomerCompany & vbcrlf
end if
if phone<>"" then
storeAdminEmail=storeAdminEmail & "Phone: " & Phone & vbcrlf
end if
if pemail<>"" then
storeAdminEmail=storeAdminEmail & "E-mail: " & Pemail & vbcrlf & vbcrlf
end if
storeAdminEmail=storeAdminEmail & "Billing Information" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & vbcrlf
storeAdminEmail=storeAdminEmail & "Address: " & pAddress & vbcrlf
storeAdminEmail=storeAdminEmail & "City: " & pCity & vbcrlf
storeAdminEmail=storeAdminEmail & "State/Province: " & pStateCode & vbcrlf
storeAdminEmail=storeAdminEmail & "Postal Code: " & pZip & vbcrlf
storeAdminEmail=storeAdminEmail & "Country Code: " & pCountry & vbcrlf & vbcrlf
storeAdminEmail=storeAdminEmail & "Shipping Information" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & vbcrlf
if pshippingAddress<>"" then
storeAdminEmail=storeAdminEmail & "Address: " & pshippingAddress & vbcrlf
storeAdminEmail=storeAdminEmail & "City: " & pshippingCity & vbcrlf
storeAdminEmail=storeAdminEmail & "State/Province: " & pshippingStateCode & pshippingState & vbcrlf
storeAdminEmail=storeAdminEmail & "Postal Code: " & pshippingZip & vbcrlf
storeAdminEmail=storeAdminEmail & "Country Code: " & pshippingCountry & vbcrlf
else
storeAdminEmail=storeAdminEmail & "(Same as the billing address)" & vbcrlf
end if
storeAdminEmail=storeAdminEmail & vbCrLf
storeAdminEmail=storeAdminEmail & "PRODUCT DETAILS" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & vbcrlf & vbcrlf
%>
<% 
pidconfigWishlistSession=idconf
mySQL="SELECT idProduct, dtCreated, fPrice, pcconf_Quantity, stringProducts, stringValues, stringCategories, stringCProducts,  stringCValues, stringCCategories, stringQuantity, stringPrice, xfdetails FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
set rs=server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(mySQL)
if err.number <> 0 then
	pcvErrDescription = err.description
	set rs = nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in printableQuote: "&pcvErrDescription) 
end if
Dim pIdProduct, pdtCreated, pxfedtails, pfPrice, stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory
pIdProduct=rs("idProduct")	
pdtCreated=rs("dtCreated")
pfPrice=rs("fPrice")
total = total+pfPrice
pQuantity=rs("pcconf_Quantity")
if (pQuantity<>"") then
else
pQuantity="1"
end if
pstringProducts = rs("stringProducts")
pstringValues = rs("stringValues")
pstringCategories = rs("stringCategories")
pstringCProducts = rs("stringCProducts")
pstringCValues = rs("stringCValues")
pstringCCategories = rs("stringCCategories")
stringQuantity = rs("stringQuantity")
stringPrice = rs("stringPrice")
pxfdetails=rs("xfdetails") 
ArrProduct = Split(pstringProducts, ",")
ArrValue = Split(pstringValues, ",")
ArrCategory = Split(pstringCategories, ",")
ArrQuantity = Split(stringQuantity, ",")
ArrPrice = Split(stringPrice, ",")
ArrCProduct = Split(pstringCProducts, ",")
ArrCValue = Split(pstringCValues, ",")
ArrCCategory = Split(pstringCCategories, ",")

mySQL="SELECT sku, description,noprices FROM Products WHERE idProduct=" & trim(pidProduct)
set rs=conntemp.execute(mySQL)
if err.number <> 0 then
	pcvErrDescription = err.description
	set rs = nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in printableQuote: "&pcvErrDescription) 
end if
psku=rs("sku")
pname=rs("description")
pnoprices=rs("noprices")
if pnoprices<>"" then
else
pnoprices=0
end if
%>
<%
' Column headings ...
storeAdminEmail=storeAdminEmail & FixedField(10, "L", "QTY")
storeAdminEmail=storeAdminEmail & FixedField(10, "L", "SKU")
storeAdminEmail=storeAdminEmail & FixedField(30, "L", "Description")
if pnoprices<2 then
	storeAdminEmail=storeAdminEmail & FixedField(20, "R", "Price")
end if
storeAdminEmail=storeAdminEmail & vbCrLf
'Column Dividers
storeAdminEmail=storeAdminEmail & FixedField(10, "L", "==========")
storeAdminEmail=storeAdminEmail & FixedField(10, "L", "==========")
storeAdminEmail=storeAdminEmail & FixedField(30, "L", "============================================================")
if pnoprices<2 then
	storeAdminEmail=storeAdminEmail & FixedField(20, "R", "====================")
end if
storeAdminEmail=storeAdminEmail & vbCrLf 								      

storeAdminEmail=storeAdminEmail & FixedField(10, "L", pQuantity)
storeAdminEmail=storeAdminEmail & FixedField(10, "L", psku)
dispStr = pname
tStr = dispStr
wrapPos=30
if len(dispStr) > 30 then
	tStr = WrapString(30, dispStr)
end if
storeAdminEmail=storeAdminEmail & FixedField(30, "L", tStr)


if pnoprices<2 then 'Calculate Product Default Price		
	
	itemsDiscounts=0
	for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
 		mySQL="select quantityFrom, quantityUntil, discountperUnit, percentage, discountperWUnit from discountsPerQuantity where IDProduct=" & ArrProduct(i)
		set rs=connTemp.execute(mySQL)
		TempDiscount=0
		do while not rs.eof
			QFrom=rs("quantityFrom")
			QTo=rs("quantityUntil")
			DUnit=rs("discountperUnit")
			QPercent=rs("percentage")
			DWUnit=rs("discountperWUnit")
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
			rs.movenext
		loop
		itemsDiscounts=ItemsDiscounts+TempDiscount
	next			

	if ItemsDiscounts>0 then
		pfPrice=pfPrice+ItemsDiscounts
	else
		pfPrice=pfPrice-ItemsDiscounts
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
	end if

	customizations=0
	for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
		if (ArrQuantity(i)-1)>=0 then
			UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
		else
			UPrice=0
		end if
		UPrice= UPrice + ArrValue(i)
		customizations=customizations+UPrice
	next
	pfPrice=pfPrice-customizations*pQuantity
	storeAdminEmail=storeAdminEmail & FixedField(20, "R", money(pfPrice))
end if 'End Calculate Product Default Price

storeAdminEmail=storeAdminEmail & vbcrlf

dispStrLen = len(dispStr)-wrapPos
do while dispStrLen > 30
	dispStr = right(dispStr,dispStrLen)
	tStr = WrapString(30, dispStr)
	storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
	storeAdminEmail=storeAdminEmail  & FixedField(30, "L", tStr)
	storeAdminEmail=storeAdminEmail  & vbCrLf					
	dispStrLen = dispStrLen-wrapPos	
loop 
if dispStrLen > 0 then
	dispStr = right(dispStr,dispStrLen)
	storeAdminEmail=storeAdminEmail  & FixedField(20, "L", "")
	storeAdminEmail=storeAdminEmail  & FixedField(30, "L", dispStr)
	storeAdminEmail=storeAdminEmail  & vbCrLf
end if

storeAdminEmail=storeAdminEmail & vbcrlf & vbcrlf
storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
%>  
             
<% 
if ArrProduct(0)="na" then
	storeAdminEmail=storeAdminEmail & bto_dictLanguage.Item(Session("language")&"_printableQuote_4") & vbcrlf
else
	storeAdminEmail=storeAdminEmail & bto_dictLanguage.Item(Session("language")&"_viewcart_1") & vbcrlf
	'calculate
	for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
		customizations=customizations+ArrValue(i)
		mySQL="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
		set rsObj=conntemp.execute(mySQL)
		mySQL="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i) 
		set rsObj1=conntemp.execute(mySQL)
		storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
		TempStr=ClearHTMLTags2(rsObj("categoryDesc"),0) & ": " & ClearHTMLTags2(rsObj("description"),0)
		if not rsObj1.eof then
			if rsObj1("displayQF")=True then
				TempStr=TempStr & "- QTY: " & ArrQuantity(i)
			end if
		end if
		dispStr = TempStr		
		dispStr = replace(dispStr,"&quot;", chr(34))
		tStr = dispStr
		wrapPos=30
		if len(dispStr) > 30 then
			tStr = WrapString(30, dispStr)
		end if
		storeAdminEmail=storeAdminEmail & FixedField(30, "L", tStr)

	
		if pnoprices<2 then
			if (CDbl(ArrValue(i))<>0) or (((ArrQuantity(i)-1)*pQuantity>0) and (ArrPrice(i)>0)) then
				if (ArrQuantity(i)-1)>=0 then
					UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
				else
					UPrice=0
				end if
				pfPrice=pfPrice+cdbl((ArrValue(i)+UPrice)*pQuantity)
				storeAdminEmail=storeAdminEmail & FixedField(20, "R", money((ArrValue(i)+UPrice)*pQuantity)) 
			end if
		end if
                              					
		storeAdminEmail=storeAdminEmail & vbcrlf

		dispStrLen = len(dispStr)-wrapPos
		do while dispStrLen > 30
			dispStr = right(dispStr,dispStrLen)
			tStr = WrapString(30, dispStr)
			storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(30, "L", tStr)
			storeAdminEmail=storeAdminEmail  & vbCrLf					
			dispStrLen = dispStrLen-wrapPos	
		loop 
		if dispStrLen > 0 then
			dispStr = right(dispStr,dispStrLen)
			storeAdminEmail=storeAdminEmail  & FixedField(20, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(30, "L", dispStr)
			storeAdminEmail=storeAdminEmail  & vbCrLf
		end if

		set rsObj=nothing
	next
	storeAdminEmail=storeAdminEmail & vbcrlf 
end if 

if pnoprices<2 then
	if ItemsDiscounts<>0 then
		pfprice=pfprice-ItemsDiscounts
		storeAdminEmail=storeAdminEmail & vbcrlf & FixedField(20, "L", "")
		storeAdminEmail=storeAdminEmail & 	FixedField(30, "L","Items Discounts:")
		if pnoprices<2 then
			storeAdminEmail=storeAdminEmail & 	FixedField(20, "R", money(-1*ItemsDiscounts))
		end if
	end if
end if%>
<% if ArrCProduct(0)<>"na" then
	pfprice=pfprice+Charges
	storeAdminEmail=storeAdminEmail & vbcrlf & vbcrlf
	storeAdminEmail=storeAdminEmail & FixedField(20, "L", "") & "Additional Charges:" & vbcrlf
	'calculate
	for i = lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
		storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
		mySQL="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
		set rsObj=conntemp.execute(mySQL)
		TempStr=ClearHTMLTags2(rsObj("categoryDesc"),0) & ": " & ClearHTMLTags2(rsObj("description"),0)
		
		dispStr = TempStr
		dispStr = replace(dispStr,"&quot;", chr(34))
		tStr = dispStr
		wrapPos=30
		if len(dispStr) > 30 then
			tStr = WrapString(30, dispStr)
		end if
		storeAdminEmail=storeAdminEmail & FixedField(30, "L", tStr)

		if pnoprices<2 then
			if (CDbl(ArrCValue(i))<>0) then
				storeAdminEmail=storeAdminEmail & FixedField(20, "R",money(ArrCValue(i)))
			end if
		end if
		set rsObj=nothing
		storeAdminEmail=storeAdminEmail & vbcrlf

		dispStrLen = len(dispStr)-30
		do while dispStrLen > 30
			dispStr = right(dispStr,dispStrLen)
			storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
			storeAdminEmail=storeAdminEmail & FixedField(30, "L", dispStr)
			storeAdminEmail=storeAdminEmail & vbCrLf					
			dispStrLen = dispStrLen-30					
		loop 
		if dispStrLen > 0 then
			dispStr = right(dispStr,dispStrLen)
			storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
			storeAdminEmail=storeAdminEmail & FixedField(30, "L", dispStr)
			storeAdminEmail=storeAdminEmail & vbCrLf					
		end if
	next
	storeAdminEmail=storeAdminEmail & vbcrlf
end if%>				
<% if trim(pxfdetails)<>"" then
	xfieldsarray=split(pxfdetails,"||")
	for i=lbound(xfieldsarray)to (UBound(xfieldsarray)-1)
		xfields=split(xfieldsarray(i),"|")
		mySQL="SELECT xfield FROM xfields WHERE idxfield="&xfields(0)
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(mySQL)
		xfielddesc=rs("xfield")
		set rs=nothing
		storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
		dispStr = xfielddesc&": "&xfields(1)
		dispStr = replace(dispStr,"&quot;", chr(34))
		tStr = dispStr
		wrapPos=30
		if len(dispStr) > 30 then
			tStr = WrapString(30, dispStr)
		end if
		storeAdminEmail=storeAdminEmail & FixedField(30, "L", tStr)

		dispStrLen = len(dispStr)-wrapPos
		do while dispStrLen > 30
			dispStr = right(dispStr,dispStrLen)
			tStr = WrapString(30, dispStr)
			storeAdminEmail=storeAdminEmail & FixedField(20, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(30, "L", tStr)
			storeAdminEmail=storeAdminEmail  & vbCrLf					
			dispStrLen = dispStrLen-wrapPos	
		loop 
		if dispStrLen > 0 then
			dispStr = right(dispStr,dispStrLen)
			storeAdminEmail=storeAdminEmail  & FixedField(20, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(30, "L", dispStr)
			storeAdminEmail=storeAdminEmail  & vbCrLf
		end if
	next
end if 
 %>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="CreateReplyEmaila.asp?action=post&datefrom=<%=request("datefrom")%>&dateto=<%=request("dateto")%>&idcustomer=<%=request("idcustomer")%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
            <td width="10%"><div align="right">To:</div></td>
            <td width="90%"><input name="toemail" type="text" value="<%=pemail%>" size="70">
            </td>
        </tr>
        <tr>
            <td><div align="right">Subject:</div></td>
            <td><input name="subject" type="text" value="This quote was submitted to <%=scCompanyName%> on <%=qSubmitDate%>" size="70"></td>
        </tr>

        <tr>
            <td valign="top"><p align="right">Message:</td>
            <td>
                <textarea wrap="off" style="background-color:White;" name="messageText" cols="50" rows="15"><%=storeAdminEmail%></textarea>
            </td>
        </tr> 
		<tr> 
        	<td></td>
			<td>
            <br>
			<br>
			<input type="submit" name="Submit" value="Send message" OnClick="document.form1.body.value = frames.message.document.body.innerHTML;" class="btn btn-primary">
			&nbsp; 
            <input type="button" class="btn btn-default"  name="Button" value="Back" onClick="document.location.href='srcQuotesa.asp?datefrom=<%=request("datefrom")%>&dateto=<%=request("dateto")%>&idcustomer=<%=request("idcustomer")%>';">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
