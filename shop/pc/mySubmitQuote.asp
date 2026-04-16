<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "mySubmitQuote.asp"
' This page is handles BTO Product Quotes
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"--> 
<%
Response.Buffer = True

dim total
total = Cint(0)

Dim rsCustObj

pcv_intProductID = getUserInput(request("idProduct"),0)
if not validNum(pcv_intProductID) then
   response.redirect "Custquotesview.asp"
end if

pcv_intConfigID = getUserInput(request("idconf"),0)
if not validNum(pcv_intConfigID) then
   response.redirect "Custquotesview.asp"
end if

pDateOrder=Date()
if SQL_Format="1" then
	pDateOrder=Day(pDateOrder)&"/"&Month(pDateOrder)&"/"&Year(pDateOrder)
else
	pDateOrder=Month(pDateOrder)&"/"&Day(pDateOrder)&"/"&Year(pDateOrder)
end if

query="Update wishlist set qsubmit=1, qdate='" & pDateOrder & "' where idCustomer=" & session("idCustomer") & " and idProduct=" & pcv_intProductID & " and idconfigWishlistSession=" & pcv_intConfigID
Set rstm=Server.CreateObject("ADODB.Recordset")
Set rstm=connTemp.execute(query) 
if err.number<>0 then
	call LogErrorToDatabase()
	set rstm=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

query="Select idquote,discountcode from wishlist where idCustomer=" & session("idCustomer") & " and idProduct=" & pcv_intProductID & " and idconfigWishlistSession=" & pcv_intConfigID
Set rstm=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstm=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

idquote=rstm("idquote")
pdiscountcode=rstm("discountcode")
if pdiscountcode="0" then
	pdiscountcode=""
end if 
Set rstm=nothing
query="SELECT name, lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, shippingaddress,shippingcity,shippingstatecode,shippingstate,shippingcountrycode,shippingzip FROM customers WHERE idCustomer=" & session("idCustomer")
Set rsCustObj=Server.CreateObject("ADODB.Recordset")
Set rsCustObj=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsCustObj=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

CustomerName=rsCustObj("name")& " " & rsCustObj("lastName")
CustomerCompany=rsCustObj("CustomerCompany")
phone=rsCustObj("phone")
pemail=rsCustObj("email")
pAddress=rsCustObj("address")
pcity=rsCustObj("city")
pStateCode=rsCustObj("stateCode")
if pStateCode="" then
	pStateCode=rsCustObj("state")
end if
pzip=rsCustObj("zip")
pcountry=rsCustObj("countryCode")
pshippingaddress=rsCustObj("shippingaddress")
pshippingcity=rsCustObj("shippingcity")
pshippingStateCode=rsCustObj("shippingStateCode")
pshippingState=rsCustObj("shippingState")
pshippingCountryCode=rsCustObj("shippingCountryCode")
pshippingZip=rsCustObj("shippingZip")
storeAdminEmail=""
storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf & "A customer has submitted a quote. Customer and product details are listed below." & "<br>" & vbcrlf & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "Quote ID: #" & idquote & "<br>" & vbcrlf & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "CUSTOMER DETAILS" & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "Name: " & CustomerName & "<br>" & vbcrlf
if CustomerCompany<>"" then
	storeAdminEmail=storeAdminEmail & "Company: " & CustomerCompany & "<br>" & vbcrlf
end if
if phone<>"" then
	storeAdminEmail=storeAdminEmail & "Phone: " & Phone & "<br>" & vbcrlf
end if
if pemail<>"" then
	storeAdminEmail=storeAdminEmail & "E-mail: " & Pemail & "<br>" & vbcrlf & "<br>" & vbcrlf
end if
storeAdminEmail=storeAdminEmail & "Billing Information" & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "Address: " & pAddress & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "City: " & pCity & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "State/Province: " & pStateCode & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "Postal Code: " & pZip & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "Country Code: " & pCountry & "<br>" & vbcrlf & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "Shipping Information" & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & "<br>" & vbcrlf
if pshippingAddress<>"" then
	storeAdminEmail=storeAdminEmail & "Address: " & pshippingAddress & "<br>" & vbcrlf
	storeAdminEmail=storeAdminEmail & "City: " & pshippingCity & "<br>" & vbcrlf
	storeAdminEmail=storeAdminEmail & "State/Province: " & pshippingStateCode & pshippingState & "<br>" & vbcrlf
	storeAdminEmail=storeAdminEmail & "Postal Code: " & pshippingZip & "<br>" & vbcrlf
	storeAdminEmail=storeAdminEmail & "Country Code: " & pshippingCountry & "<br>" & vbcrlf
else
	storeAdminEmail=storeAdminEmail & "(Same as the billing address)" & "<br>" & vbcrlf
end if
set rsCustObj=nothing
storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "PRODUCT DETAILS" & "<br>" & vbcrlf
storeAdminEmail=storeAdminEmail & "====================" & "<br>" & vbcrlf & "<br>" & vbcrlf
%>
<% pidconfigWishlistSession=request.QueryString("idconf")
query="SELECT idProduct, dtCreated, fPrice, dPrice, stringQuantity, stringProducts, stringValues, stringCategories, xfdetails FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
set rs=server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
								
Dim pIdProduct, pdtCreated, pxfedtails, pfPrice, stringQuantity, stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory
pIdProduct=rs("idProduct")
pdtCreated=rs("dtCreated")
pfPrice=rs("fPrice")
dPrice=rs("dPrice")
total = total+pfPrice
pstringQuantity = rs("stringQuantity")
pstringProducts = rs("stringProducts")
pstringValues = rs("stringValues")
pstringCategories = rs("stringCategories")
pxfdetails=rs("xfdetails") 
ArrQuantity = Split(pstringQuantity, ",")
ArrProduct = Split(pstringProducts, ",")
ArrValue = Split(pstringValues, ",")
ArrCategory = Split(pstringCategories, ",")

query="SELECT sku, description,noprices FROM Products WHERE idProduct=" & trim(pidProduct)
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
psku=rs("sku")
pname=rs("description")
pnoprices=cint(rs("noprices"))

' Column headings ...
storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
storeAdminEmail=storeAdminEmail & FixedField(40, "L", "Description")
if pnoprices=0 then
	storeAdminEmail=storeAdminEmail & FixedField(20, "R", "Price")
end if
storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf
'Column Dividers
storeAdminEmail=storeAdminEmail & FixedField(10, "L", "==========")
storeAdminEmail=storeAdminEmail & FixedField(40, "L", "============================================================")
if pnoprices=0 then
	storeAdminEmail=storeAdminEmail & FixedField(20, "R", "====================")
end if
storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf 								      

storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
dispStr = replace(pname & " (" & psku & ")","&quot;", chr(34))
tStr = dispStr
wrapPos=40
if len(dispStr) > 40 then
	tStr = WrapString(40, dispStr)
end if
storeAdminEmail=storeAdminEmail & FixedField(40, "L", tStr)

if pnoprices=0 then
	customizations=0
	for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
		customizations=customizations+ArrValue(i)
	next
	originalprice=pfPrice-customizations
	storeAdminEmail=storeAdminEmail & FixedField(20, "R", money(originalprice))      
end if
storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf
dispStrLen = len(dispStr)-wrapPos
do while dispStrLen > 40
	dispStr = right(dispStr,dispStrLen)
	tStr = WrapString(40, dispStr)
	storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
	storeAdminEmail=storeAdminEmail  & FixedField(40, "L", tStr)
	storeAdminEmail=storeAdminEmail  & "<br>" & vbcrlf					
	dispStrLen = dispStrLen-wrapPos	
loop 
if dispStrLen > 0 then
	dispStr = right(dispStr,dispStrLen)
	storeAdminEmail=storeAdminEmail  & FixedField(10, "L", "")
	storeAdminEmail=storeAdminEmail  & FixedField(40, "L", dispStr)
	storeAdminEmail=storeAdminEmail  & "<br>" & vbcrlf
end if


storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
%>               
<% if ArrProduct(0)="na" then
	storeAdminEmail=storeAdminEmail & bto_dictLanguage.Item(Session("language")&"_printableQuote_4") & "<br>" & vbcrlf
else
	storeAdminEmail=storeAdminEmail & bto_dictLanguage.Item(Session("language")&"_viewcart_1") & "<br>" & vbcrlf

	'calculate
	for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
		query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pidProduct 
		set rsQ=server.CreateObject("ADODB.RecordSet") 
		set rsQ=conntemp.execute(query)
		if not rsQ.eof then						
			btDisplayQF=rsQ("displayQF")
		else
			btDisplayQF=0
		end if
		set rsQ=nothing
				
		customizations=customizations+ArrValue(i)
		query="SELECT products.sku, categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
		set rsObj=Server.CreateObject("ADODB.Recordset")
		set rsObj=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsObj=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		pcv_strBtoItemSku = rsObj("sku")
		pcv_strBtoItemSku=ClearHTMLTags2(pcv_strBtoItemSku,0)
		pcv_strBtoItemName = rsObj("description")
		pcv_strBtoItemName=ClearHTMLTags2(pcv_strBtoItemName,0)
		pcv_strBtoItemCat=rsObj("categoryDesc")
		pcv_strBtoItemCat=ClearHTMLTags2(pcv_strBtoItemCat,0)		
	
		dispStr =""
		storeAdminEmail=storeAdminEmail & FixedField(10, "L", "") 		
		dispStr = pcv_strBtoItemCat&": "&pcv_strBtoItemName
		dispStr = dispStr & " - SKU: " & pcv_strBtoItemSku				
		if btDisplayQF=True then
			if clng(ArrQuantity(i))>1 then
				dispStr = dispStr & " - QTY: " & ArrQuantity(i)
			end if
		end if			
		dispStr = replace(dispStr,"&quot;", chr(34))
		tStr = dispStr
		wrapPos=40
		if len(dispStr) > 40 then
			tStr = WrapString(40, dispStr)
		end if
		storeAdminEmail=storeAdminEmail & FixedField(40, "L", tStr)

		if pnoprices=0 then
			if Ccur(ArrValue(i))>0 then
				storeAdminEmail=storeAdminEmail & FixedField(20, "R", money(ArrValue(i)))
			else
				storeAdminEmail=storeAdminEmail & FixedField(20, "R", "")
			end if
		end if
		storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf

		dispStrLen = len(dispStr)-wrapPos
		do while dispStrLen > 40
			dispStr = right(dispStr,dispStrLen)
			tStr = WrapString(40, dispStr)
			storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(40, "L", tStr)
			storeAdminEmail=storeAdminEmail  & "<br>" & vbcrlf					
			dispStrLen = dispStrLen-wrapPos	
		loop 
		if dispStrLen > 0 then
			dispStr = right(dispStr,dispStrLen)
			storeAdminEmail=storeAdminEmail  & FixedField(10, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(40, "L", dispStr)
			storeAdminEmail=storeAdminEmail  & "<br>" & vbcrlf
		end if
		set rsObj=nothing
	next 
end if %>


<% 
if trim(pxfdetails)<>"" then

	xfieldsarray=split(pxfdetails,"||")
	
	for i=lbound(xfieldsarray)to (UBound(xfieldsarray)-1)
		xfields=split(xfieldsarray(i),"|")
		query="SELECT xfield FROM xfields WHERE idxfield="&xfields(0)
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
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
		
		xfielddesc=rs("xfield")
		set rs=nothing
		dispStr = ""
		storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
		dispStr = xfielddesc&": "&xfields(1)
		dispStr = replace(dispStr,"&quot;", chr(34))
		tStr = dispStr
		wrapPos=50
		if len(dispStr) > 50 then
			tStr = WrapString(50, dispStr)
		end if
		storeAdminEmail=storeAdminEmail & FixedField(50, "L", tStr)& "<br>" & vbcrlf
		
		dispStrLen = len(dispStr)-wrapPos
		do while dispStrLen > 50
			dispStr = right(dispStr,dispStrLen)
			tStr = WrapString(50, dispStr)
			storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(50, "L", tStr)
			storeAdminEmail=storeAdminEmail  & "<br>" & vbcrlf					
			dispStrLen = dispStrLen-wrapPos	
		loop 
		if dispStrLen > 0 then
			dispStr = right(dispStr,dispStrLen)
			storeAdminEmail=storeAdminEmail  & FixedField(10, "L", "")
			storeAdminEmail=storeAdminEmail  & FixedField(50, "L", dispStr)
			storeAdminEmail=storeAdminEmail  & "<br>" & vbcrlf
		end if
		
	next
	
end if
								
discountcheck=0
if pnoprices=0 then		
	 
	IF Pdiscountcode<>"" then
		pDiscountError=Cstr("")
	
		if pDiscountCode="" then
			noCode="1"
		  pDiscountError="-" 
		else
			query="SELECT iddiscount, onetime, expDate, idProduct, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, DiscountDesc, priceToDiscount, percentageToDiscount, pcDisc_StartDate FROM discounts WHERE discountcode='" &pDiscountCode& "' AND active=-1"
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			if rstemp.eof then
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4") 
			else 
				piddiscount=rstemp("iddiscount")
				ponetime=rstemp("onetime")
				pexpDate=rstemp("expDate")
				ptmpidProduct=rstemp("idProduct")
				pquantityFrom=rstemp("quantityFrom")
				pquantityUntil=rstemp("quantityUntil")
				pweightFrom=rstemp("weightFrom")
				pweightUntil=rstemp("weightUntil")
				ppriceFrom=rstemp("priceFrom")
				ppriceUntil=rstemp("priceUntil")
				pDiscountDesc=rstemp("DiscountDesc")
				ppriceToDiscount=rstemp("priceToDiscount")
				ppercentageToDiscount=rstemp("percentageToDiscount")
				pStartDate=rstemp("pcDisc_StartDate")
				set rstemp=nothing
	 
				'check to see if discount has been used for one use only for this customer specified
				If ponetime=true Then
					'check customer's id in database with iddiscount
					query="SELECT A.idcustomer FROM used_discounts A,customers B WHERE A.iddiscount=" & piddiscount & " AND A.idcustomer=B.idcustomer AND B.email IN (SELECT C.email FROM customers C WHERE C.idcustomer="&session("IDCustomer")&");"
					set rsDisObj=Server.CreateObject("ADODB.Recordset")
					set rsDisObj=conntemp.execute(query)
				
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsDisObj=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
				
					if NOT rsDisObj.eof then
						'discount has been used already by the customer
						pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
					end if
					set rsDisObj=nothing
				Else
					'check to see if discount code is expired
					If pexpDate<>"" then
						expDate=pexpDate
						If datediff("d", Now(), expDate) <= 0 Then
							pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
						Else
							' check if the discount has defined the product   
							pVerPrdCode=-1     
							if isNull(ptmpidProduct) or ptmpidProduct=0 then
								' discount is across the board
							else
								' find out if the product is in the cart
								if findProduct(pcCartArray, ppcCartIndex, ptmpidProduct)=0 then
									pVerPrdCode=0
								end if
							end if   
						end if
					end if
			
					'check to see if discount has start date
					If pStartDate<>"" then
						StartDate=pStartDate
						If datediff("d", Now(), StartDate) > 0 Then
							pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_43")
						End If
					end if
				end if
				If pDiscountError="" Then
					if Int(pCartQuantity)>=Int(pquantityFrom) and Int(pCartQuantity)<=Int(pquantityUntil) and Int(pCartTotalWeight)>=Int(pweightFrom) and Int(pCartTotalWeight)<=Int(pweightUntil) and Cdbl(pSubTotal)>=Cdbl(ppriceFrom) and Cdbl(pSubTotal)<=Cdbl(ppriceUntil) then
						pcv_DiscountDesc=pDiscountDesc
						pcv_PriceToDiscount=cdbl(pPriceToDiscount)
						pcv_percentageToDiscount=ppercentageToDiscount
					else
						pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5") 
					end if
				End If
			end if
		end if

		if pdiscountcode<>"" then
			storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf & "Discount code: " & pdiscountcode & "<br>" & vbcrlf
		end if
		

		if pDiscountError="" then
			storeAdminEmail=storeAdminEmail & "+ Details: " & pcv_DiscountDesc & "<br>" & vbcrlf
			storeAdminEmail=storeAdminEmail & "+ Amount: -" & scCurSign & money(dPrice) & "<br>" & vbcrlf   
			discountcheck=1
		else
			if pDiscountError<>"-" then
				storeAdminEmail=storeAdminEmail & "Error: " & pDiscountError & "<br>" & vbcrlf
			end if
		end if
		storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf		
						
						
	END IF
end if%>			
<%if pnoprices=0 then								
	storeAdminEmail=storeAdminEmail & FixedField(50, "R", "Total:")
	if discountcheck=1 then
		storeAdminEmail=storeAdminEmail & 	FixedField(20, "R", scCurSign &  money(pfPrice-dPrice)) & "<br>" & vbcrlf
	else
		storeAdminEmail=storeAdminEmail & 	FixedField(20, "R", scCurSign &  money(pfPrice)) & "<br>" & vbcrlf
	end if
end if
storeAdminEmail=storeAdminEmail & "<br>" & vbcrlf & "<br>" & vbcrlf

storeAdminEmail=replace(storeAdminEmail,"''",chr(39))

session("News_MsgType")="1"
storeAdminEmail = pcf_HtmlEmailWrapper(storeAdminEmail, pcv_HTMLEmailFontFamily)
call sendmail (scCompanyName, scEmail, scFrmEmail, scCompanyName & " - A new quote has been submitted", storeAdminEmail)
response.redirect "Custquotesview.asp?msg=4"
%>
