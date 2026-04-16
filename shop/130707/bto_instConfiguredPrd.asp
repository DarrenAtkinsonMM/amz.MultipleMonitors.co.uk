<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "bto_instConfiguredPrd.asp"
' This page handles Configurable Products
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Response.Buffer = True

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

dim ptotalQuantity



pTotalQuantity = Cint(0)

' check for bound quantity in cart
if countCartRows(pcCartArray, ppcCartIndex) = scQtyLimit then
  	call closeDb()
	response.redirect "msg.asp?message=39"
end if

' get data from viewPrd form
piBTOQuote_rec=getUserInput(request("iBTOQuote_rec.x"),0)
if piBTOQuote_rec="" then
	piBTOQuote_rec=getUserInput(request("iBTOQuote_rec"),0)
end if
pidconf=getUserInput(request("idconf"),10)
save_pidconf=getUserInput(request("idConfigWishlistSession"),10)
if save_pidconf="0" then
	save_pidconf=""
end if
if pidconf="" then
	pidconf=save_pidconf
end if
pBTOQuote = getUserInput(request("iBTOQuote.x"),0)
if pBTOQuote="" then
	pBTOQuote = getUserInput(request("iBTOQuote"),0)
end if
pIdProduct = getUserInput(request("idproduct"),10)

query="SELECT * FROM configSpec_Charges WHERE specProduct="&pIdProduct
set rs99=server.CreateObject("ADODB.RecordSet")
set rs99=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs99=nothing
	
	call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

BTOCharges=0
if not rs99.eof then
	BTOCharges=1
end if

pcv_QDisc=getUserInput(request.Form("QDiscounts0"),10)
if pcv_QDisc="" then
	pcv_QDisc="0"
end if

session("QDisc" & pIDProduct)=pcv_QDisc

if BTOCharges=1 then
	session("DefaultPrice"  & pIDProduct)=cdbl(getUserInput(request.Form("TLPriceDefault"),10))
	session("CMPrice"  & pIDProduct)=cdbl(getUserInput(request.Form("CMDefault"),10))
	session("CMWQD"  & pIDProduct)=cdbl(getUserInput(request.Form("CMWQD"),10))
else
	session("DefaultPrice"  & pIDProduct)=0
	session("CMPrice"  & pIDProduct)=0
	session("CMWQD"  & pIDProduct)=0
end if

pQuantity = getUserInput(request.Form("quantity"),10)
'how many products are part of this configured product
pJcnt = getUserInput(request.Form("jCnt"),4)
'how many categories are multiselect
CB_CatCnt = request.Form("CB_CatCnt")
'price of configured product
ReqDiscounts = request.Form("Discounts")
ReqDiscounts=replace(ReqDiscounts,scCurSign,"")
if scDecSign="." then
	ReqDiscounts=replace(ReqDiscounts,",","")
else
	ReqDiscounts=replace(ReqDiscounts,".","")
	ReqDiscounts=replace(ReqDiscounts,",",".")
end if
if ReqDiscounts="" then
	ReqDiscounts=0
end if
pGrTotal = getUserInput(request.Form("GrandTotal"),10)
pGrTotal=replace(pGrTotal,scCurSign,"")
if scDecSign="." then
	pGrTotal=replace(pGrTotal,",","")
else
	pGrTotal=replace(pGrTotal,".","")
	pGrTotal=replace(pGrTotal,",",".")
end if
pGrTotal1=pGrTotal
pGrandTotal2=getUserInput(request.Form("GrandTotal2"),10)
pGrandTotal2=replace(pGrandTotal2,scCurSign,"")
if scDecSign="." then
	pGrandTotal2=replace(pGrandTotal2,",","")
else
	pGrandTotal2=replace(pGrandTotal2,".","")
	pGrandTotal2=replace(pGrandTotal2,",",".")
end if

pfPrice=getUserInput(request.Form("TLGrandTotal"),10)
pfPrice=replace(pfPrice,scCurSign,"")
if scDecSign="." then
	pfPrice=replace(pfPrice,",","")
else
	pfPrice=replace(pfPrice,".","")
	pfPrice=replace(pfPrice,",",".")
end if

'--> Custom input fields
Dim XFCount,iq,tmpXFStr
tmpXFStr=""
xString=""
tmpXFStr=""
xCnt=0
xfieldsCnt=0
XFCount=getUserInput(request("XFCount"),0)
if (XFCount<>"") and (IsNumeric(XFCount)) then
	session("XFCount")=XFCount
	if Clng(XFCount)>=1 then
		For iq=1 to XFCount
			pxfield = getUserInput(request.Form("xfield" & iq),0)

			'replace line breaks to <br>
			if pxfield<>"" then
				pxfield=replace(pxfield,vbCrlf,"<BR>")
			end if
	
			session("SFxfield" & iq & "_" & pIdProduct)=pxfield

			pxf = getUserInput(request.Form("xf" & iq),10)
			session("SFxfield" & iq & "ID_" & pIdProduct)=pxf
			
			if pxfield<>"" then
				xString=xString&pxf&"|"&pxfield&"||"
				xfieldsCnt=xfieldsCnt+1
				query="SELECT xfield FROM xfields WHERE idxfield="&pxf
				set rstemp=conntemp.execute(query)
				if not rstemp.eof then
					pXfieldDescrip=rstemp("xfield")
				end if
				set rstemp=nothing
	
				if xCnt=1 then
					tmpXFStr=tmpXFStr & "<br>"
				end if
		
				tmpXFStr=tmpXFStr & pXfieldDescrip & ": " & pxfield
				xCnt=1
			end if
		Next
	end if
end if

' if cannot get quantity get quantity 1 (from listing)
if NOT validNum(pQuantity) then
	pQuantity=1
end if

if pQuantity="" OR int(pQuantity)<1 then
 pQuantity=1
end if

' get item details
err.clear
dim noOS
noOS=0
query="SELECT OverSizeSpec FROM products"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number<>0 then
	noOS=1
end if
err.clear

query="SELECT iRewardPoints,description, price, bToBPrice, sku, weight, emailText, deliveringTime, idSupplier, cost, stock, notax, noshipping, OverSizeSpec,noStock,pcprod_QtyToPound,pcProd_BackOrder,pcProd_ShipNDays,pcDropShipper_ID FROM products WHERE idproduct=" &pIdProduct& " AND active=-1"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	
	call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if  rstemp.eof then
	
	call closeDb()
response.redirect "msg.asp?message="&Server.Urlencode(dictLanguage.Item(Session("language")&"_instPrd_D") )
end if

pcv_FromVP=0

pDfPrice = getUserInput(request("TLPriceDefault"),10)
if pDfPrice="" then
	pDfPrice = getUserInput(request("TLPriceDefaultVP"),10)
	if pDfPrice<>"" then
		pDfPrice=pDfPrice*pQuantity
		pcv_FromVP=1
	end if
end if
if pDfPrice="" then
	pDfPrice = 0
end if

pPrice = pGrTotal
pDescription = rstemp("description")
pWeight	= rstemp("weight")
pSku = rstemp("sku")
pEmailText = rstemp("emailText")
pDeliveringTime	= rstemp("deliveringTime")
pIdSupplier	= rstemp("idSupplier")
pcDropShipperID = rstemp("pcDropShipper_ID")
if IsNull(pcDropShipperID) OR (pcDropShipperID="") then
	pcDropShipperID=0
end if
pCost = rstemp("cost")
pStock = rstemp("stock")
pnotax = rstemp("notax")
pnoshipping = rstemp("noshipping")
pOverSizeSpec=rstemp("OverSizeSpec")
iRewardPoints = rstemp("iRewardPoints")
iRewardDollars = pPrice * (RewardsPercent / 100)
pNoStock=rstemp("noStock")
pcv_QtyToPound=rstemp("pcprod_QtyToPound")
if pcv_QtyToPound>0 then
	pWeight=(16/pcv_QtyToPound)
	if scShipFromWeightUnit="KGS" then
		pWeight=(1000/pcv_QtyToPound)
	end if
end if

'Start SDBA
pcv_intBackOrder = rstemp("pcProd_BackOrder")
if isNull(pcv_intBackOrder) OR pcv_intBackOrder="" then
	pcv_intBackOrder = 0
end if
pcv_intShipNDays = rstemp("pcProd_ShipNDays")
if isNull(pcv_intShipNDays) OR pcv_intShipNDays="" then
	pcv_intShipNDays = 0
end if
'End SDBA

set rstemp=nothing

' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function

'insert product configuration into configSessions

'create strings
Dim Pstring, Vstring, Cstring, tempVar, tempCatarray, tempString, strArray
Pstring = ""
Vstring = ""
Cstring = ""
Cweight = 0
Qstring = ""
Pricestring=""
FirstCnt = request.form("FirstCnt")
If FirstCnt<>"" then
	for i = 1 to FirstCnt
		pcv_TempCweight=0
		tempVar = request.form("CAT"&i)
		MS=request.form("MS"&i)
		If MS="" then
			tempCatarray = split(tempVar,"G")
			tempString = request.form(tempVar)
			If tempString<>"" then
			strArray = split(tempString, "_")
			If strArray(0)<>0 then
				Cstring = Cstring & tempCatarray(1) & ","
				Pstring = Pstring & strArray(0) & ","
				Vstring = Vstring & strArray(1) & ","
				if cdbl(strArray(2))>0 then
					pcv_TempCweight = pcv_TempCweight + cdbl(strArray(2))
				else
					query="SELECT pcprod_QtyToPound FROM Products WHERE idproduct=" & strArray(0)
					set rsW=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsW=nothing
						
						call closeDb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if

					itemWeight=0
					if not rsW.eof then
						item_QtyToPound=rsW("pcprod_QtyToPound")
						if item_QtyToPound>0 then
							itemWeight=(16/item_QtyToPound)
							if scShipFromWeightUnit="KGS" then
								itemWeight=(1000/item_QtyToPound)
							end if
						end if
					end if
					pcv_TempCweight = pcv_TempCweight + itemWeight
					set rsW=nothing
				end if
				Pricestring=Pricestring & strArray(3) & ","
				QF = request.form(tempVar & "QF")
				pcv_TempCweight = pcv_TempCweight * QF
				Cweight=Cweight+pcv_TempCweight
				Qstring= Qstring & QF & ","
			End If
			End if 'tempString<>""
		End If
	next
end if
'continue strings with multiselect items
Dim c, p, tempCatVar, tempMSPrd
if CB_CatCnt<>"" then
	for c=1 to CB_CatCnt
		tempCatVar= request.form("CB_CatID"&c)
		tempPrdCntVar= request.form("PrdCnt"&tempCatVar)
		for p=1 to tempPrdCntVar
			pcv_TempCweight=0
			tempMSPrd = request.form("Cat"&tempCatVar&"_"&"Prd"&p)
			tempString = request.form("CAG"&tempCatVar&tempMSPrd)
			if tempString<>"" then
				strArray = split(tempString, "_")
				If strArray(0)<>0 then
					Cstring = Cstring & tempCatVar & ","
					Pstring = Pstring & strArray(0) & ","
					Vstring = Vstring & strArray(1) & ","
					if Cdbl(strArray(2))>0 then
						pcv_TempCweight = pcv_TempCweight + Cdbl(strArray(2))
					else
						query="SELECT pcprod_QtyToPound FROM Products WHERE idproduct=" & strArray(0)
						set rsW=connTemp.execute(query)
						if err.number<>0 then
							'//Logs error to the database
							call LogErrorToDatabase()
							'//clear any objects
							set rsW=nothing
							'//close any connections
							
							'//redirect to error page
							call closeDb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if

						itemWeight=0
						if not rsW.eof then
						item_QtyToPound=rsW("pcprod_QtyToPound")
						if item_QtyToPound>0 then
							itemWeight=(16/item_QtyToPound)
							if scShipFromWeightUnit="KGS" then
								itemWeight=(1000/item_QtyToPound)
							end if
						end if
						end if
						pcv_TempCweight = pcv_TempCweight + itemWeight
						set rsW=nothing
					end if
					Pricestring=Pricestring & strArray(3) & ","
					QF = request.form("CAG"&tempCatVar&tempMSPrd & "QF")
					pcv_TempCweight = pcv_TempCweight * QF
					Cweight=Cweight+pcv_TempCweight
					Qstring= Qstring & QF & ","
				End If

			End IF
		next
	next
end if

if Pstring<>"" then
	query="select * from configSpec_products where specproduct=" & pIDproduct & " order by catSort, prdSort"
	set rs4=connTemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rs4=nothing
		'//close any connections
		
		'//redirect to error page
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	A=split(Cstring,",")
	B=split(PString,",")
	C=split(Vstring,",")
	D=split(Qstring,",")
	E=split(Pricestring,",")
	Count=0
	Dim A1(100),B1(100),C1(100),D1(100),E1(100)

	Do while not rs4.eof

		For k=lbound(B) to ubound(B)
			if B(k)<>"" then
				if (clng(B(k))=clng(rs4("configProduct"))) and (clng(A(k))=clng(rs4("configProductCategory"))) then
					A1(Count)=A(k)
					B1(Count)=B(k)
					C1(Count)=C(k)
					D1(Count)=D(k)
					E1(Count)=E(k)
					Count=Count+1
				
				else

					'// Apparel Sub-Prds
					IF statusAPP="1" THEN
						query="SELECT pcProd_ParentPrd FROM Products WHERE idproduct=" & B(k)
						set rsA=connTemp.execute(query)
						app_ParentPrd=0
						if not rsA.eof then
							app_ParentPrd=rsA("pcProd_ParentPrd")
						end if
						set rsA=nothing
						if (clng(app_ParentPrd)=clng(rs4("configProduct"))) and (clng(A(k))=clng(rs4("configProductCategory"))) then
							A1(Count)=A(k)
							B1(Count)=B(k)
							C1(Count)=C(k)
							D1(Count)=D(k)
							E1(Count)=E(k)
							Count=Count+1
						end if
					END IF
			
				end if
			end if
		Next

		rs4.MoveNext
	Loop
	Cstring=""
	Pstring=""
	Vstring=""
	Qstring=""
	Pricestring=""
	For k=0 to Count-1
		Cstring=Cstring & A1(k) & ","
		Pstring=Pstring & B1(k) & ","
		Vstring=Vstring & C1(k) & ","
		Qstring=Qstring & D1(k) & ","
		Pricestring=Pricestring & E1(k) & ","
	Next
end if

if Cstring="" then
	Cstring="na"
end if
if Pstring="" then
	Pstring="na"
end if
if Vstring="" then
	Vstring="na"
end if

If Pstring="" then
	if scConfigPurchaseOnly=1 then
	call closeDb()
	response.redirect "bto_configurePrd.asp?idproduct="&pIdProduct&"&msg="&server.URLEncode(bto_dictLanguage.Item(Session("language")&"_instConfiguredPrd_1"))
	response.end
	end if
end if

discountcodetemp=request("discountcode")
If trim(discountcodetemp)<>"" then
	query="SELECT onetime,iddiscount FROM discounts WHERE discountcode='"&discountcodetemp&"' AND active=-1"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rs.eof then
		
		call closeDb()
response.redirect "bto_configurePrd.asp?idproduct="&pIdProduct&"&msg="&Server.Urlencode("The discount code you entered is not valid.")
	end if
	If rs("onetime")=true Then
		'check customer's id in database with iddiscount
		query="SELECT A.idcustomer FROM used_discounts A,customers B WHERE A.iddiscount=" & rs("iddiscount") & " AND A.idcustomer=B.idcustomer AND B.email IN (SELECT C.email FROM customers C WHERE C.idcustomer="&session("IDCustomer")&");"
		set rsDisObj=conntemp.execute(query)
		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rsDisObj=nothing
			'//close any connections
			
			'//redirect to error page
			call closeDb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		if rsDisObj.eof then
		else
			
			call closeDb()
response.redirect "bto_configurePrd.asp?idproduct="&pIdProduct&"&msg="&Server.Urlencode("The discount code you entered is no longer valid.")
		end if
	End if
End If

' get discount certificate data
pDiscountError=Cstr("")
pDiscountCode=discountcodetemp

if pDiscountCode="" then
	noCode="1"
	pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_3")
else
	query="SELECT iddiscount, onetime, expDate, idProduct, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, DiscountDesc, priceToDiscount, percentageToDiscount, pcDisc_StartDate FROM discounts WHERE discountcode='" &pDiscountCode& "' AND active=-1"

	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4")
	else
		piddiscount=rstemp("iddiscount")
		ponetime=rstemp("onetime")
		pexpDate=rstemp("expDate")
		pidProduct=rstemp("idProduct")
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
			set rsDisObj=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsDisObj=nothing
				
				call closeDb()
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
					if isNull(pidProduct) or pidProduct=0 then
						' discount is across the board
					else
						' find out if the product is in the cart
						if findProduct(pcCartArray, ppcCartIndex, pidProduct)=0 then
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
				pcv_PriceToDiscount=cdbl(ppriceToDiscount)
				pcv_percentageToDiscount=ppercentageToDiscount
			else
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
			end if
		End If
	end if
end if

discountTotal=Cdbl(0)
' discounts
if pDiscountError="" then
	discountTotal=Cdbl(0)
	' calculate discount. Note: percentage dont affects shipment and payment prices
	if pcv_PriceToDiscount>0 or pcv_percentageToDiscount>0 then
		discountTotal=pcv_PriceToDiscount + (pcv_percentageToDiscount*(pGrTotal1)/100)
	end if
end if

if (pcv_FromVP=1) and (pcQDiscountType<>"1") and (Pstring<>"") then
itemsDiscounts=0
ArrProduct=split(Pstring,",")
ArrValue=Split(Vstring, ",")
ArrCategory=Split(Cstring, ",")
ArrQuantity=Split(Qstring,",")
ArrPrice=split(Pricestring,",")
for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i)
set rs99=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs99=nothing
	
	call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

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
				if request("customertype")=1 then
				TempD1=ArrQuantity(i)*pQuantity*ArrPrice(i)*0.01*DWUnit
				else
				TempD1=ArrQuantity(i)*pQuantity*ArrPrice(i)*0.01*DUnit
				end if
				else
				if request("customertype")=1 then
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
ReqDiscounts=-1*ItemsDiscounts
end if
else
	ItemsDiscounts=ReqDiscounts
end if

discountTotal=discountTotal + cdbl(ReqDiscounts)
session("SPstring")=Pstring
session("SVstring")=Vstring
session("SCstring")=Cstring
session("SxString")=xString
session("SpGrTotal1")=pGrTotal1
session("SdiscountTotal")=cdbl(ReqDiscounts)
session("SQstring")=Qstring
session("SPricestring")=Pricestring

pConfigKey=trim(randomNumber(9999)&randomNumber(9999))

session("pConfigKey")=pConfigKey

'If this is a quote, add to quote sessions and then to wishlist and then redirect to wishlist page
Dim pTodayDate
pTodayDate=Date()
if SQL_Format="1" then
	pTodayDate=Day(pTodayDate)&"/"&Month(pTodayDate)&"/"&Year(pTodayDate)
else
	pTodayDate=Month(pTodayDate)&"/"&Day(pTodayDate)&"/"&Year(pTodayDate)
end if
If pBTOQuote<>"" or piBTOQuote_rec<>"" or save_pidconf<>"" then
	if Pstring="" then
	else
		if (piBTOQuote_rec<>"") or (save_pidconf<>"") then
			pidConf=save_pidconf
			query="UPDATE configWishlistSessions SET stringProducts='"&Pstring&"',stringValues='"&Vstring&"',stringCategories='"&Cstring&"',xfdetails=N'"&xString&"',fPrice="&pfPrice-pcv_QDisc&",dPrice=" & discountTotal & ",stringQuantity='" & Qstring & "',stringPrice='" & Pricestring & "',pcconf_Quantity=" & pQuantity & ",pcconf_QDiscount=" & pcv_QDisc & " WHERE idconfigWishlistSession="&pidConf
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsConf=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pIdConfigWishlistSession=pidConf
		else
			err.clear
			query="INSERT INTO configWishlistSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,xfdetails,dtCreated,fPrice,dPrice,stringQuantity,stringPrice,stringCProducts,stringCValues,stringCCategories,pcconf_Quantity,pcconf_QDiscount) VALUES ("&pConfigKey &","&pIdProduct&",'"&Pstring&"','"&Vstring&"','"&Cstring&"',N'"&xString&"','"&pTodayDate&"',"&pfPrice-pcv_QDisc&"," & discountTotal & ",'" & Qstring & "','" & Pricestring & "','na','na','na'," & pQuantity & "," & pcv_QDisc & ")"
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsConf=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		end if
		if pidConf="" then
			query="SELECT idconfigWishlistSession FROM configWishlistSessions WHERE configKey="&pConfigKey&" AND dtCreated='"&pTodayDate&"';"
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsConf=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			Dim pIdConfigWishlistSession
			if not rsConf.eof then
				pIdConfigWishlistSession = rsConf("idconfigWishlistSession")
			else
				pIdConfigWishlistSession=0
			end if
			set rsConf=nothing
		else
			pIdConfigWishlistSession = pidConf
		end if
	end if

	set rsConf=nothing
	if (pBTOQuote="") and (piBTOQuote_rec="") and (save_pidconf<>"") then
	else
		if request("idquote") then
			call closeDb()
response.redirect "viewquote.asp?action=autoupd&idquote=" & request("idquote") & "&idcustomer=" & request("idcustomer")
		end if
	end if
'else
else
	if Pstring="" then
		pIdConfigSession=""
	else
		if request("pre_idConfigSession")<>"" then
			query="UPDATE configSessions SET idproduct="&pIdProduct&",stringProducts='"&Pstring&"',stringValues='"&Vstring&"',stringCategories='"&Cstring&"',stringQuantity='" & Qstring & "',stringPrice='" & Pricestring & "' WHERE idconfigSession="& request("pre_idConfigSession")
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			set rsConf=nothing
			pIdConfigSession = request("pre_idConfigSession")
		else
			query="INSERT INTO configSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,dtCreated,stringQuantity,stringPrice,stringCProducts,stringCValues,stringCCategories) VALUES ("&pConfigKey &","&pIdProduct&",'"&Pstring&"','"&Vstring&"','"&Cstring&"','"&pTodayDate&"','" & Qstring & "','" & Pricestring & "','na','na','na')"
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsConf=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			query="SELECT idconfigSession FROM configSessions WHERE configKey="&pConfigKey&" AND dtCreated='"&pTodayDate&"';"
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsConf=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pIdConfigSession = rsConf("idconfigSession")
			set rsConf=nothing
		end if
	end if
'end if
end if

pta_Price=pPrice
pTotalQuantity = pQuantity

pta_xfdetails=""
if xfieldsCnt>0 then
	pta_xfdetails=tmpXFStr
end if

' get discount per quantity

query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
set rstemp=conntemp.execute(query)
if err.number<>0 and err.number<>9 then
	call LogErrorToDatabase()
	set rstemp=nothing
	
	call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
end if

tempNum=0
pOrigPrice=cdbl(pPrice)

if pcQDiscountType<>"1" then
pOrigPrice=pOrigPrice+(ReqDiscounts/pTotalQuantity)
else
pOrigPrice=pDfPrice/pTotalQuantity
end if

tunitPrice = pta_Price

if not rstemp.eof then
	' there are quantity discounts defined for that quantity
	pDiscountPerUnit = rstemp("discountPerUnit")
	pDiscountPerWUnit = rstemp("discountPerWUnit")
	pPercentage = rstemp("percentage")

	if customerType<>1 then
		if pPercentage = "0" then
			pta_Price  = pta_Price - pDiscountPerUnit
			tempNum = tempNum + (pDiscountPerUnit * pTotalQuantity)
		else
			pta_Price =pta_Price - ((pDiscountPerUnit/100) * pOrigPrice)
			tempNum = tempNum + ((pDiscountPerUnit/100) * (pOrigPrice * pTotalQuantity))
		end if
	else
		if pPercentage = "0" then
			pta_Price  = pta_Price - pDiscountPerWUnit
			tempNum = tempNum + (pDiscountPerWUnit * pTotalQuantity)
		else
			pta_Price = pta_Price - ((pDiscountPerWUnit/100) * pOrigPrice)
			tempNum = tempNum + ((pDiscountPerWUnit/100) * (pOrigPrice * pTotalQuantity))
		end if
	end if
end if

if request("idquote")<>"" then
	tunitPrice = pta_Price
end if

pIdOrder=request("idorder")
if ItemsDiscounts<0 then
	ItemsDiscounts=-1*ItemsDiscounts
end if

if pta_xfdetails<>"" then
	pta_xfdetails=replace(pta_xfdetails,"<br>","|")
	pta_xfdetails=replace(pta_xfdetails,"'","''")
	pta_xfdetails=replace(pta_xfdetails,"''''","''")
	pta_xfdetails=replace(pta_xfdetails,vbCrlf,"<BR>") '// replace last so the <BR> is not replaced
end if

if request("pre_idConfigSession")<>"" then
	query="UPDATE ProductsOrdered SET quantity="&pQuantity&", unitPrice="&tunitPrice&", xfdetails=N'"&pta_xfdetails&"',QDiscounts=" & tempNum & ",ItemsDiscounts=" & ItemsDiscounts & ", pcDropShipper_ID="&pcDropShipperID&" WHERE idConfigSession="&request("pre_idConfigSession")&";"
	set rstemp=conntemp.execute(query)
else
	if pIdOrder<>"" and pIdOrder<>"0" then
		query="INSERT INTO ProductsOrdered (idOrder, idProduct, quantity, unitPrice, unitCost, xfdetails, idconfigSession,QDiscounts,ItemsDiscounts,pcDropShipper_ID) VALUES ("&pIdOrder&","&pIdProduct&","&pQuantity&","&tunitPrice&",0,N'"&pta_xfdetails&"',"&pIdConfigSession&"," & tempNum & "," & ItemsDiscounts & ","&pcDropShipperID&");"
		set rstemp=conntemp.execute(query)

		query="Select IdProductOrdered from ProductsOrdered order by IdProductOrdered desc"
		set rstemp=conntemp.execute(query)
		pIdProductOrdered=rstemp("IdProductOrdered")
	end if
end if

NextStep=request("NextStep")


call clearLanguage()

conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing

if BTOCharges=1 then
	if NextStep="1" then
		call closeDb()
response.redirect "bto_RePrdAddCharges.asp?idconf=" & getUserInput(request("idConfigWishlistSession"),0) & "&idp=" & request("IdProductOrdered") & "&ido=" & request("idorder") & "&idquote=" & request("idquote") & "&pre_idConfigSession=" & request("pre_idConfigSession") & "&customertype=" & request("customertype") & "&idcustomer=" & request("idcustomer") & "&idproduct=" & pIdProduct
	else
		call closeDb()
response.redirect "bto_PrdAddCharges.asp?idp=" & pIdProductOrdered & "&ido=" & pIDOrder & "&ConfigSession=" & pIdConfigSession
	end if
else
	if request("idquote")<>"" then
		call closeDb()
response.redirect "viewquote.asp?action=autoupd&idquote=" & request("idquote") & "&idcustomer=" & request("idcustomer")
	else
		call closeDb()
response.redirect "AdminEditOrder.asp?ido="&request("idorder")&"&action=upd"
	end if
end if
%>