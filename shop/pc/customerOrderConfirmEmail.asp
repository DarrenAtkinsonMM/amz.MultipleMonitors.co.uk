<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'GGG Add-on start

query="select pcOrd_IDEvent,pcOrd_GWTotal from Orders where idOrder=" & qry_ID
set rs19=server.CreateObject("ADODB.RecordSet")
set rs19=conntemp.execute(query)

gIDEvent=rs19("pcOrd_IDEvent")
if gIDEvent<>"" then
else
gIDEvent="0"
end if

pGWTotal=rs19("pcOrd_GWTotal")
if pGWTotal<>"" then
else
pGWTotal="0"
end if

geHideAddress=0

if gIDEvent<>"0" then

	query="select pcEvents.pcEv_Notify, pcEvents.pcEv_name, pcEvents.pcEv_Date, pcEvents.pcEv_HideAddress, customers.name, customers.lastname, customers.email from pcEvents, Customers where Customers.idcustomer=pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
	set rs1=server.CreateObject("ADODB.RecordSet")
	set rs1=conntemp.execute(query)

	geNotify=rs1("pcEv_Notify")
	if geNotify<>"" then
	else
	geNotify="0"
	end if
	geName=rs1("pcEv_name")
	geDate=rs1("pcEv_Date")
	if year(geDate)="1900" then
	geDate=""
	end if
	if gedate<>"" then
		if scDateFrmt="DD/MM/YY" then
		gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
		else
		gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
		end if
	end if
	geHideAddress=rs1("pcEv_HideAddress")
	if geHideAddress<>"" then
	else
	geHideAddress="0"
	end if
	gReg=rs1("name") & " " & rs1("lastname")
	gRegemail=rs1("email")

end if
'GGG Add-on end

' compile customer email
customerEmail=Cstr("")
' Build body of message ...

pCustomerFullName=pName&" "&pLName

customerEmail=""
'Customized message from store owner
If (scConfirmEmail<>"" and pcv_CustomerReceived=0) or (scReceivedEmail<>"" and pcv_CustomerReceived=1) Then
	todaydate=showDateFrmt(now())
	if pcv_CustomerReceived=1 then
		personalmessage=replace(scReceivedEmail,vbCrlf,"<br>")
	else
		personalmessage=replace(scConfirmEmail,vbCrlf,"<br>")
	end if
	personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
	personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
	personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
	personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",pCustomerFullName)
	personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(pIdOrder)))
	personalmessage=replace(personalmessage,"<ORDER_DATE>",todaydate)
	personalmessage=replace(personalmessage,"''",chr(39))
	personalmessage=replace(personalmessage,"//","/")
	personalmessage=replace(personalmessage,"http:/","http://")
	personalmessage=replace(personalmessage,"https:/","https://")

	customerEmail=customerEmail & "<br>" & vbcrlf & personalmessage & "<br><br>" & vbcrlf
End If

If pcOrderKey<>"" then
	customerEmail=customerEmail & FixedField(80, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbcrlf 
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_storeEmail_30") & pcOrderKey & "<br>" & vbcrlf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_storeEmail_31") & "<br>" & vbcrlf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_storeEmail_32") & "<br>" & vbcrlf
	customerEmail=customerEmail & FixedField(80, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbcrlf 
End If


'GGG Add-on start
	query="select idproduct from ProductsOrdered WHERE idOrder="& qry_ID
	set rstemp=connTemp.execute(query)
	pGCs="0"
	do while not rstemp.eof
		query="select pcprod_GC from products where idproduct=" & rstemp("idproduct")
		set rs=connTemp.execute(query)
		if not rs.eof then
			pGC=rs("pcprod_GC")
			if (pGC<>"") and (pGC="1") then
				pGCs="1"
			end if
		end if
		rstemp.moveNext
	loop

	query="select pcOrd_GCDetails,pcOrd_GCAmount from Orders where idOrder=" & pIDOrder
	set rs19=connTemp.execute(query)

	GCDetails=rs19("pcOrd_GCDetails")
	GCAmountTotal=rs19("pcOrd_GCAmount")
	if GCAmountTotal="" OR IsNull(GCAmountTotal) then
		GCAmountTotal=0
	end if
'GGG Add-on end

customerEmail=customerEmail & "<br>" & vbcrlf & dictLanguage.Item(Session("language")&"_sendMail_2") & "<br>" & vbcrlf
customerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf
customerEmail=customerEmail & pCustomerFullName & "<br>" & vbcrlf

If Trim(pCustomerCompany) <> "" Then
	customerEmail=customerEmail & pCustomerCompany & "<br>" & vbcrlf
End If

customerEmail=customerEmail & paddress & "<br>" & vbcrlf
if paddress2<>"" then
	customerEmail=customerEmail & paddress2 & "<br>" & vbcrlf
end if
customerEmail=customerEmail & pCity & ", "
if pState = "" then
	customerEmail=customerEmail & pStateCode & " "
	else
	customerEmail=customerEmail & pState & " "
end if
customerEmail=customerEmail & pzip & "<br>" & vbcrlf
customerEmail=customerEmail & pCountryCode & "<br>" & vbcrlf
customerEmail=customerEmail & pEmail & "<br>" & vbcrlf & "<br>" & vbcrlf

if geHideAddress=0 then
customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_3") & "<br>" & vbcrlf
customerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf
If Trim(pshippingAddress) <> "" Then
	if pShippingFullName<>"" then
		customerEmail=customerEmail & pShippingFullName & "<br>" & vbcrlf
	end if
	if pShippingCompany<>"" then
		customerEmail=customerEmail & pShippingCompany & "<br>" & vbcrlf
	end if
	customerEmail=customerEmail & pshippingAddress & "<br>" & vbcrlf
	if pshippingAddress2<>"" then
		customerEmail=customerEmail & pshippingAddress2 & "<br>" & vbcrlf
	end if
	customerEmail=customerEmail & pshippingCity & ", "
	if pshippingState = "" then
		customerEmail=customerEmail & pshippingStateCode & " "
		else
		customerEmail=customerEmail & pshippingState & " "
	end if
	customerEmail=customerEmail & pshippingZip & "<br>" & vbcrlf
	customerEmail=customerEmail & pshippingCountryCode & "<br>" & vbcrlf
	customerEmail=customerEmail & trim(pshippingPhone) & "<br>" & vbcrlf
Else
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_4") & "<br>" & vbcrlf
End if
End if

customerEmail=customerEmail & "<br>" & vbcrlf

pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)

boolVAT = CheckVAT(pcv_IsEUMemberState,pshippingCountryCode,pVATID)

'get shipping details...
shipping=split(pshipmentDetails,",")
if ubound(shipping)>1 then
	if NOT isNumeric(trim(shipping(2))) then
		customerEmail=customerEmail & ship_dictLanguage.Item(Session("language")&"_noShip_a") & "<br>" & vbcrlf
		Service=""
		Postage=0
		serviceHandlingFee=0
	else
		Shipper=shipping(0)
		Service=shipping(1)
		Postage=trim(shipping(2))
		if ubound(shipping)=>3 then
			serviceHandlingFee=trim(shipping(3))
			if NOT isNumeric(serviceHandlingFee) then
				serviceHandlingFee=0
			end if
		else
			serviceHandlingFee=0
		end if
	end if
else
	customerEmail=customerEmail & ship_dictLanguage.Item(Session("language")&"_noShip_a") & "<br>" & vbcrlf
	Service=""
	Postage=0
	serviceHandlingFee=0
end if

If DFShow="1" AND pord_DeliveryDate <> "" Then
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_CustviewOrd_39") & pord_DeliveryDate & "<br>" & vbcrlf
end if

if Service<>"" then
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_5") & Service & "<br>" & vbcrlf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_41") & pOrdPackageNum & "<br>" & vbcrlf
end if
customerEmail=customerEmail & "<br>" & vbcrlf

'offline payment details
paymentdetails=split(trim(pPaymentDetails),"||")
if ubound(paymentdetails)>0 then
	paymentCharge=trim(paymentdetails(1))
	If NOT isNumeric(paymentCharge) then
		paymentCharge=0
	End if
else
	paymentCharge=0
end if

'GGG Add-on start

query="select pcOrd_IDEvent,pcOrd_GWTotal from Orders where idOrder=" & qry_ID
set rs19=connTemp.execute(query)

gIDEvent=rs19("pcOrd_IDEvent")
if gIDEvent<>"" then
else
gIDEvent="0"
end if

pGWTotal=rs19("pcOrd_GWTotal")
if pGWTotal<>"" then
else
pGWTotal="0"
end if

if gIDEvent<>"0" then
customerEmail=customerEmail & "<br>" & vbcrlf
customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_55") & geName & "<br>" & vbcrlf
customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_56") & geDate & "<br>" & vbcrlf
customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_57") & gReg & "<br>" & vbcrlf
customerEmail=customerEmail & "<br>" & vbcrlf
end if
'GGG Add-on end

'get discount details...
'Check if more then one discount code was utilized
if instr(pdiscountDetails,",") then
	DiscountDetailsArry=split(pdiscountDetails,",")
	intArryCnt=ubound(DiscountDetailsArry)
else
	intArryCnt=0
end if
pTotalDiscountAmount=0

for k=0 to intArryCnt
	if intArryCnt=0 then
		pTempDiscountDetails=pdiscountDetails
	else
		pTempDiscountDetails=DiscountDetailsArry(k)
	end if
	if instr(pTempDiscountDetails,"- ||") then
		discounts= split(pTempDiscountDetails,"- ||")
		pdiscountDesc=discounts(0)
		pdiscountAmt=trim(discounts(1))
		pIsNumeric=1
		if NOT isNumeric(pdiscountAmt) then
			pdiscountAmt=0
			pIsNumeric=0
		end if
		if (pdiscountAmt>0 OR pdiscountAmt=0) AND pIsNumeric=1 then
			customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_6") & pdiscountDesc & "<br>" & vbcrlf
		end if
	Else
		pdiscountAmt=0
	end if
	pTotalDiscountAmount=pTotalDiscountAmount+pdiscountAmt
Next

If RewardsActive <> 0 And ( (piRewardPointsCustAccrued > 0) Or (piRewardPoints > 0)) Then
	customerEmail=customerEmail & "<br>" & vbcrlf
	'Did we use points or accrue points?
	If piRewardPointsCustAccrued > 0 AND piRewardPoints=0 Then 'Accrued
		iDollarValue=piRewardPointsCustAccrued * (RewardsPercent / 100)
		customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_7") & piRewardPointsCustAccrued & " " & RewardsLabel & dictLanguage.Item(Session("language")&"_sendMail_8") & scCurSign &money(iDollarValue) & "<br>" & vbcrlf
	End If
	If piRewardPoints > 0 Then
		customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_9") & money(piRewardValue) & dictLanguage.Item(Session("language")&"_sendMail_10") & RewardsLabel & "!" & "<br>" & vbcrlf
	End If
	customerEmail=customerEmail & "<br>" & vbcrlf
End If

' Begin order details ...

'GGG Add-on start
'Add bookmarks
customerEmail=customerEmail & "AAAAAAAAAA"
'GGG Add-on end

customerEmail=customerEmail & "<br>" & vbcrlf

customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_adminMail_13") & "<br>" & vbcrlf & "<br>" & vbcrlf
customerEmail=customerEmail & FixedField(10, "L", dictLanguage.Item(Session("language")&"_adminMail_16"))
customerEmail=customerEmail & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_18"))
customerEmail=customerEmail & FixedField(15, "R", dictLanguage.Item(Session("language")&"_adminMail_14"))
customerEmail=customerEmail & FixedField(15, "R", dictLanguage.Item(Session("language")&"_adminMail_15"))
customerEmail=customerEmail & "<br>" & vbcrlf

customerEmail=customerEmail & FixedField(80, "R", "====================================================================================================")
customerEmail=customerEmail & "<br>" & vbcrlf
iSubtotal=0

query="SELECT products.idproduct,products.sku, products.description, ProductsOrdered.pcSC_ID, ProductsOrdered.quantity, ProductsOrdered.unitPrice, xfdetails"
'CONFIGURATOR ADDON-S
if scBTO=1 then
	query=query&" ,ProductsOrdered.idconfigSession"
end if
'CONFIGURATOR ADDON-E
query=query&",ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray,ProductsOrdered.pcPO_GWOpt,ProductsOrdered.pcPO_GWNote,ProductsOrdered.pcPO_GWPrice, pcPrdOrd_BundledDisc FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& pIdOrder

set rsOrderDetails=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsOrderDetails=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

Do While Not rsOrderDetails.EOF
	pidproduct=rsOrderDetails("idproduct")
	psku=rsOrderDetails("sku")
	pdescription=rsOrderDetails("description")
	pdescription=ClearHTMLTags2(pdescription,0)
	pcSCID=rsOrderDetails("pcSC_ID")
	if IsNull(pcSCID) OR len(pcSCID)=0 then
		pcSCID=0
	end if
	pqty=rsOrderDetails("quantity")
	pPrice=rsOrderDetails("unitPrice")
	xfdetails=replace(rsOrderDetails("xfdetails"),"&lt;BR&gt;","<br>")
	'xfdetails=replace(xfdetails,"<BR>",vbcrlf)
	if scBTO=1 then
		pIdConfigSession=rsOrderDetails("idconfigSession")
	end if

	'BTO Items Discounts & Quantity Discounts
	QDiscounts=rsOrderDetails("QDiscounts")
	ItemsDiscounts=rsOrderDetails("ItemsDiscounts")

	'// Product Options Arrays
	pcv_strSelectedOptions = rsOrderDetails("pcPrdOrd_SelectedOptions") ' Column 11
	pcv_strOptionsPriceArray = rsOrderDetails("pcPrdOrd_OptionsPriceArray") ' Column 25
	pcv_strOptionsArray = rsOrderDetails("pcPrdOrd_OptionsArray") ' Column 4

	'GGG Add-on start
	pGWOpt=rsOrderDetails("pcPO_GWOpt")
	if pGWOpt<>"" then
	else
	pGWOpt="0"
	end if
	pGWText=rsOrderDetails("pcPO_GWNote")
	pGWPrice=rsOrderDetails("pcPO_GWPrice")
	if pGWPrice<>"" then
	else
	pGWPrice="0"
	end if
	'GGG Add-on end
	pcPrdOrd_BundledDisc=rsOrderDetails("pcPrdOrd_BundledDisc")

	pExtendedPrice=pPrice*pqty
	customerEmail=customerEmail & FixedField(10, "L", pqty)
	dispStr = replace(pdescription & " (" & psku & ")","&quot;", chr(34))
	tStr = dispStr
	wrapPos=40
	if len(dispStr) > 40 then
		tStr = WrapString(40, dispStr)
	end if
	customerEmail=customerEmail & FixedField(40, "L", tStr)
	customerEmail=customerEmail & "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" & "<br>" & vbcrlf
	dispStrLen = len(dispStr)-wrapPos
	do while dispStrLen > 40
		dispStr = right(dispStr,dispStrLen)
		tStr = WrapString(40, dispStr)
		customerEmail=customerEmail & FixedField(10, "L", "")
		customerEmail=customerEmail & FixedField(40, "L", tStr)
		customerEmail=customerEmail & "<br>" & vbcrlf					
		dispStrLen = dispStrLen-wrapPos	
	loop 
	if dispStrLen > 0 then
		dispStr = right(dispStr,dispStrLen)
		customerEmail=customerEmail & FixedField(10, "L", "")
		customerEmail=customerEmail & FixedField(40, "L", dispStr)
		customerEmail=customerEmail & "<br>" & vbcrlf
	end if
	'CONFIGURATOR ADDON-S
	TotalUnit=0
	if scBTO=1 then
		'Add customizations if there are any
		if pIdConfigSession<>"0" then
			query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsConfigObj=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsConfigObj=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_34A"))
			customerEmail=customerEmail & FixedField(15, "R", " ")
			customerEmail=customerEmail & FixedField(15, "R", " ") & "<br>" & vbcrlf
			stringProducts=rsConfigObj("stringProducts")
			stringValues=rsConfigObj("stringValues")
			stringCategories=rsConfigObj("stringCategories")
			stringQuantity=rsConfigObj("stringQuantity")
			stringPrice=rsConfigObj("stringPrice")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			ArrQuantity=Split(stringQuantity, ",")
			ArrPrice=Split(stringPrice, ",")
			for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
			
				If statusAPP="1" Then

					query="SELECT products.pcProd_ParentPrd FROM products WHERE products.idProduct="&ArrProduct(i)&";" 
					set rsConfigObj=conntemp.execute(query)
					If Not rsConfigObj.Eof Then
						pcv_intIdProduct=rsConfigObj("pcProd_ParentPrd")
					End If
					Set rsConfigObj = Nothing

					if pcv_intIdProduct>"0" then
					else
						pcv_intIdProduct=ArrProduct(i)
					end if

				Else

					pcv_intIdProduct = ArrProduct(i)
				
				End If
				
				query="SELECT displayQF FROM configSpec_Products WHERE configProduct="& pcv_intIdProduct &" and specProduct=" & pidProduct & " AND configProductCategory=" & ArrCategory(i) & ";"
				set rsQ=server.CreateObject("ADODB.RecordSet")
				set rsQ=conntemp.execute(query)
				if not rsQ.eof then
					btDisplayQF=rsQ("displayQF")
				end if
				set rsQ=nothing

				query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & pcv_intIdProduct & ";"
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
				query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & pcv_intIdProduct & " AND cdefault<>0" & " AND configProductCategory=" & ArrCategory(i) & ";"
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

				query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
				set rsConfigObj=conntemp.execute(query)
				pcv_strBtoItemName = rsConfigObj("description")
				pcv_strBtoItemName=ClearHTMLTags2(pcv_strBtoItemName,0)
				pcv_strBtoItemCat=rsConfigObj("categoryDesc")
				pcv_strBtoItemCat=ClearHTMLTags2(pcv_strBtoItemCat,0)
				customerEmail=customerEmail & FixedField(10, "L", "")
				dispStr=""
				dispStr = pcv_strBtoItemCat &": "& pcv_strBtoItemName
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
				customerEmail=customerEmail & FixedField(40, "L", tStr)

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
					TotalUnit=TotalUnit+cdbl((ArrValue(i)+UPrice)*pQty)
					customerEmail=customerEmail & FixedField(30, "R", scCurSign & money((ArrValue(i)+UPrice)*pQty))
				else
					if tmpDefault=1 then
						customerEmail=customerEmail & FixedField(30, "R", dictLanguage.Item(Session("language")&"_defaultnotice_1"))
					end if
				end if

				customerEmail=customerEmail & "<br>" & vbcrlf

				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 40
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(40, dispStr)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(40, "L", tStr)
					customerEmail=customerEmail & "<br>" & vbcrlf
					dispStrLen = dispStrLen-wrapPos
				loop
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(40, "L", dispStr)
					customerEmail=customerEmail & "<br>" & vbcrlf
				end if
				set rsConfigObj=nothing
			next
		end if
	end if
	'CONFIGURATOR ADDON-E

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Add first 40 characters of options on a separate line
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
		pcv_strSelectedOptions = ""
	end if

	If len(pcv_strSelectedOptions)>0 Then

			'// Add the header "OPTIONS"
			customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L","OPTIONS") & "<br>" & vbcrlf

			'#####################
			' START LOOP
			'#####################
			'// Generate Our Local Arrays from our Stored Arrays

			' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers
			pcArray_strSelectedOptions = ""
			pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))

			' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
			pcArray_strOptionsPrice = ""
			pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))

			' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
			pcArray_strOptions = ""
			pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))

			' Get Our Loop Size
			pcv_intOptionLoopSize = 0
			pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)

			' Display Our Options
			For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
				dispStr = ""

				'//There isnt a header after the first one so we indent
				customerEmail=customerEmail & FixedField(10, "L", " ")

				dispStr = pcArray_strOptions(pcv_intOptionLoopCounter)
				dispStr = replace(dispStr,"&quot;", chr(34))
				tStr = dispStr
				wrapPos=40
				if len(dispStr) > 40 then
					tStr = WrapString(40, dispStr)
				end if
				customerEmail=customerEmail & FixedField(40, "L", tStr)

				tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)

				if tempPrice="" or tempPrice=0 then
					customerEmail=customerEmail & FixedField(30, "R", " ")
					customerEmail=customerEmail & "<br>" & vbcrlf
				else
					customerEmail=customerEmail & FixedField(30, "R", "")
					customerEmail=customerEmail & "<br>" & vbcrlf
				end if
				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 40
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(40, dispStr)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(40, "L", tStr)
					customerEmail=customerEmail & "<br>" & vbcrlf
					dispStrLen = dispStrLen-wrapPos
				loop
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(40, "L", dispStr)
					customerEmail=customerEmail & "<br>" & vbcrlf
				end if
			Next

			'#####################
			' END LOOP
			'#####################

			customerEmail=customerEmail & "<br>" & vbcrlf
	End If
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Add first 40 characters of options on a separate line
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	If len(xfdetails)>3 then
		customerEmail=customerEmail & "<br>" & vbcrlf
		xfarray=split(xfdetails,"|")
		for q=lbound(xfarray) to ubound(xfarray)
			customerEmail=customerEmail & FixedField(10, "L", "")
			dispStr = replace(xfarray(q),"&quot;", chr(34))
			tStr = dispStr
			wrapPos=40
			if len(dispStr) > 40 then
				tStr = WrapString(40, dispStr)
			end if
			if Instr(tStr,"<BR>") then
				tStr=replace(tStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
				customerEmail=customerEmail & tStr & "<br>" & vbcrlf
			else
				customerEmail=customerEmail & FixedField(40, "L", tStr) & "<br>" & vbcrlf
			end if
			dispStrLen = len(dispStr)-wrapPos
			if inStr(dispStr,"<BR>")>0 then
				dispStr = right(dispStr,dispStrLen)
				dispStr = FixedField(10, "L", "") & replace(dispStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
				customerEmail=customerEmail & dispStr & "<br>" & vbcrlf
			else
				do while dispStrLen > 40
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(40, dispStr)
					if Instr(tStr,"<BR>") then
						tStr=replace(tStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
					end if
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(40, "L", tStr)
					customerEmail=customerEmail & "<br>" & vbcrlf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					if Instr(dispStr,"<BR>") then
						dispStr=replace(dispStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
					end if
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(40, "L", dispStr)
					customerEmail=customerEmail & "<br>" & vbcrlf
				end if
			end if
		next
	End If

	pPrice1=pPrice
	pExtendedPrice1=pExtendedPrice

	if TotalUnit>0 then
		pExtendedPrice1=pExtendedPrice1-TotalUnit
		pPrice1=Round(pExtendedPrice1/pqty,2)
	end if

	tmpText1=""
	if money(pPrice1)=money(pExtendedPrice1) then
		tmpText1=tmpText1 & FixedField(15, "R","")
	else
		tmpText1=tmpText1 & FixedField(15, "R", scCurSign & money(pPrice1/1.2))
	end if
	tmpText1=tmpText1 & FixedField(15, "R", scCurSign & money(pExtendedPrice1/1.2)) & "<br>" & vbcrlf
	customerEmail=replace(customerEmail,"BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" & "<br>" & vbcrlf,tmpText1)
	if pcPrdOrd_BundledDisc>0 then
		customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_custOrdInvoice_36"))
		customerEmail=customerEmail & FixedField(15, "R", " ")
		customerEmail=customerEmail & FixedField(15, "R","-" & scCurSign & money(pcPrdOrd_BundledDisc))  & "<br>" & vbcrlf
	end if
	customerEmail=customerEmail & "<br>" & vbcrlf

	'CONFIGURATOR ADDON-S
	Charges=0
	if scBTO=1 then
		if pIdConfigSession<>"0" then
			if (ItemsDiscounts<>"") and (ItemsDiscounts<>"0") then
				customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_sendMail_37"))
				customerEmail=customerEmail & FixedField(15, "R", " ")
				customerEmail=customerEmail & FixedField(15, "R","-" & scCurSign & money(ItemsDiscounts))  & "<br>" & vbcrlf
			end if

				'Add customizations if there are any
				if pIdConfigSession<>"0" then
					query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
					set rsConfigObj=conntemp.execute(query)
					stringCProducts=rsConfigObj("stringCProducts")
					stringCValues=rsConfigObj("stringCValues")
					stringCCategories=rsConfigObj("stringCCategories")
					ArrCProduct=Split(stringCProducts, ",")
					ArrCValue=Split(stringCValues, ",")
					ArrCCategory=Split(stringCCategories, ",")
					if ArrCProduct(0)<>"na" then
						for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
							Charges=Charges+Cdbl(ArrCValue(i))
						next

						customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_sendMail_40"))
						customerEmail=customerEmail & FixedField(15, "R", " ")
						customerEmail=customerEmail & FixedField(15, "R", " ") & "<br>" & vbcrlf

						for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
							query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))"
							set rsConfigObj=conntemp.execute(query)
							dispStr =""
							pcv_strBtoItemName = rsConfigObj("description")
							pcv_strBtoItemName=ClearHTMLTags2(pcv_strBtoItemName,0)
							pcv_strBtoItemCat=rsConfigObj("categoryDesc")
							pcv_strBtoItemCat=ClearHTMLTags2(pcv_strBtoItemCat,0)

							customerEmail=customerEmail & FixedField(10, "L", "")
							dispStr = pcv_strBtoItemCat &": "& pcv_strBtoItemName

							dispStr = replace(dispStr,"&quot;", chr(34))
							tStr = dispStr
							wrapPos=40
							if len(dispStr) > 40 then
								tStr = WrapString(40, dispStr)
							end if
							customerEmail=customerEmail & FixedField(40, "L", tStr)

							if ArrCValue(i)<>0 then
								customerEmail=customerEmail & FixedField(30, "R", scCursign & money(ArrCValue(i)))
							end if
							customerEmail=customerEmail & "<br>" & vbcrlf

							dispStrLen = len(dispStr)-wrapPos
							do while dispStrLen > 40
								dispStr = right(dispStr,dispStrLen)
								tStr = WrapString(40, dispStr)
								customerEmail=customerEmail & FixedField(10, "L", "")
								customerEmail=customerEmail & FixedField(40, "L", tStr)
								customerEmail=customerEmail & "<br>" & vbcrlf
								dispStrLen = dispStrLen-wrapPos
							loop
							if dispStrLen > 0 then
								dispStr = right(dispStr,dispStrLen)
								customerEmail=customerEmail & FixedField(10, "L", "")
								customerEmail=customerEmail & FixedField(40, "L", dispStr)
								customerEmail=customerEmail & "<br>" & vbcrlf
							end if

							set rsConfigObj=nothing
						next
					end if
				end if

			iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges)-cdbl(pcPrdOrd_BundledDisc)

		else
			iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(pcPrdOrd_BundledDisc)
		end if
	else
		iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(pcPrdOrd_BundledDisc)
			end if
			'CONFIGURATOR ADDON-E

	'======================================
		if (QDiscounts<>"") and (QDiscounts<>"0") then
			customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_32"))
			customerEmail=customerEmail & FixedField(15, "R", " ")
			customerEmail=customerEmail & FixedField(15, "R", "-" & scCurSign & money(QDiscounts)) & "<br>" & vbcrlf
		end if
	iSubTotal=iSubtotal-cdbl(QDiscounts)

	cdblCmprTmp1=(pPrice*pqty)
	cdblCmprTmp2=(pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges)

	if cdblCmprTmp2<>cdblCmprTmp1 then
		customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_33"))
		customerEmail=customerEmail & FixedField(15, "R", " ")
		customerEmail=customerEmail & FixedField(15, "R", scCurSign & money((pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges))) & "<br>" & vbcrlf
	end if

	'GGG Add-on start
	if pGWOpt<>"0" then
	query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
	set rsG=connTemp.execute(query)
	if not rsG.eof then
		pGWOptName=rsG("pcGW_OptName")
		customerEmail=customerEmail & "<br>" & vbcrlf
		customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_39")) & "<br>" & vbcrlf
		customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", pGWOptName) & FixedField(15, "R", scCurSign & money(pGWPrice)) & "<br>" & vbcrlf
		customerEmail=customerEmail & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_41")) & "<br>" & vbcrlf
		dispStr = pGWText
		customerEmail=customerEmail & FixedField(10, "L", "")
		tStr = dispStr
		wrapPos=40
		if len(dispStr) > 40 then
			tStr = WrapString(40, dispStr)
		end if
		if Instr(tStr,vbcrlf) then
			tStr=replace(tStr,vbcrlf,"<br>" & vbcrlf & FixedField(10, "L", ""))
			customerEmail=customerEmail & tStr & "<br>" & vbcrlf
		else
			customerEmail=customerEmail & FixedField(40, "L", tStr) & "<br>" & vbcrlf
		end if
		dispStrLen = len(dispStr)-wrapPos
		do while dispStrLen > 40
			dispStr = right(dispStr,dispStrLen)
			tStr = WrapString(40, dispStr)
			if Instr(tStr,vbcrlf) then
				tStr=replace(tStr,vbcrlf,"<br>" & vbcrlf & FixedField(10, "L", ""))
			end if
			customerEmail=customerEmail & FixedField(10, "L", "")
			customerEmail=customerEmail & FixedField(40, "L", tStr)
			customerEmail=customerEmail & "<br>" & vbcrlf					
			dispStrLen = dispStrLen-wrapPos	
		loop 
		if dispStrLen > 0 then
			dispStr = right(dispStr,dispStrLen)
			if Instr(dispStr,vbcrlf) then
				dispStr=replace(dispStr,vbcrlf,"<br>" & vbcrlf & FixedField(10, "L", ""))
			end if
			customerEmail=customerEmail & FixedField(10, "L", "")
			customerEmail=customerEmail & FixedField(40, "L", dispStr)
			customerEmail=customerEmail & "<br>" & vbcrlf
		end if
	end if
	end if
	'GGG Add-on end

	customerEmail=customerEmail & "<br>" & vbcrlf

	rsOrderDetails.MoveNext
loop

' Break then start totals ...
customerEmail=customerEmail & "<br>" & vbcrlf
customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_19"))
customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(iSubTotal/1.2))
customerEmail=customerEmail & "<br>" & vbcrlf

'GGG Add-on start
'Add bookmarks
customerEmail=customerEmail & "AAAAAAAAAA"
'GGG Add-on end

' processing fees ...
if paymentCharge<>0 then
	customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_20"))
	customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(paymentCharge))
	customerEmail=customerEmail & "<br>" & vbcrlf
end if

'DiscountCode/Rewards Pts., when applicable...
ptotalDiscounts=pTotalDiscountAmount+piRewardValue+pcOrd_CatDiscounts+GCAmountTotal
if ptotalDiscounts>0 then
	if piRewardValue>0 then
		customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_21")&RewardsLabel&dictLanguage.Item(Session("language")&"_sendMail_22"))
	else
		customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_23"))
	end if
	customerEmail=customerEmail & FixedField(15, "R", "-"&scCurSign & money(ptotalDiscounts/1.2))
	customerEmail=customerEmail & "<br>" & vbcrlf
End If

'GGG Add-on start
If pGWTotal<>"0" Then
	customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_63"))
	customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(pGWTotal))
	customerEmail=customerEmail & "<br>" & vbcrlf
End If
'GGG Add-on end

' Shipping, when applicable ...
If Postage<>0 Then
	customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_25"))
	customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(Postage/1.2))
	customerEmail=customerEmail & "<br>" & vbcrlf
End If

' Shipping, when applicable ...
If serviceHandlingFee>"0" Then
	customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_26"))
	customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(serviceHandlingFee))
	customerEmail=customerEmail & "<br>" & vbcrlf
End If

'DA EDIT - SHOW VAT CALC STUFF
if boolVAT > 0 Then
	customerEmail=customerEmail & FixedField(64, "R", "VAT:")
	customerEmail=customerEmail & FixedField(16, "R", scCurSign & money(ptotal - (ptotal/1.2)))
	customerEmail=customerEmail & "<br />"
else
	customerEmail=customerEmail & FixedField(64, "R", "VAT:")
	customerEmail=customerEmail & FixedField(16, "R", scCurSign & money(0))
	customerEmail=customerEmail & "<br />"
end if

' Sales tax, when applicable ...
if pord_VAT>0 then
	If ptaxAmount>"0" Then
		customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_orderverify_35"))
		customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(pord_VAT))
		customerEmail=customerEmail & "<br>" & vbcrlf
	End If
else
	if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
		If ptaxAmount>"0" Then
			customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_24"))
			customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(ptaxAmount))
			customerEmail=customerEmail & "<br>" & vbcrlf
		End If
	else
		taxArray=split(ptaxDetails,",")
		tempTaxAmount=0
		for i=0 to (ubound(taxArray)-1)
			taxDesc=split(taxArray(i),"|")
			if taxDesc(0)<>"" then
			customerEmail=customerEmail & FixedField(65, "R", taxDesc(0)&":")
			customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(taxDesc(1)))
			customerEmail=customerEmail & "<br>" & vbcrlf
			end if
		next
	end if
end if

' Grand total ...
	customerEmail=customerEmail & FixedField(65, "R", "===============")
	customerEmail=customerEmail & FixedField(15, "R", "===============")
	customerEmail=customerEmail & "<br>" & vbcrlf
	customerEmail=customerEmail & FixedField(65, "R", dictLanguage.Item(Session("language")&"_sendMail_27"))
	customerEmail=customerEmail & FixedField(15, "R", scCurSign & money(ptotal))
	customerEmail=customerEmail & "<br>" & vbcrlf

' Check for comments by customer
If pcomments<>"" then
	customerEmail=customerEmail & "<br>" & vbcrlf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_80") & pcomments
	customerEmail=customerEmail & "<br>" & vbcrlf
End If


' Sign off ...
	customerEmail=customerEmail & "<br>" & vbcrlf

'GGG Add-on start

IF (GCDetails<>"") then
CustomerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf & "<br>" & vbcrlf

CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_46") & "<br>" & vbcrlf & "<br>" & vbcrlf

GCArr=split(GCDetails,"|g|")
intGCCount=ubound(GCArr)
For y=0 to intGCCount
if GCArr(y)<>"" then
	GCInfo=split(GCArr(y),"|s|")
	pGiftCode=GCInfo(0)
	pGiftUsed=GCInfo(2)

	query="select products.IDProduct,products.Description from pcGCOrdered,Products where products.idproduct=pcGCOrdered.pcGO_idproduct and pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
	set rs=connTemp.execute(query)

	if not rs.eof then
		pIdproduct=rs("idproduct")
		pName=rs("Description")
		pCode=pGiftCode
		CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_47") & pName & "<br>" & vbcrlf

		query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
		set rs19=connTemp.execute(query)

		if not rs19.eof then
			CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_48") & rs19("pcGO_GcCode") & "<br>" & vbcrlf
			CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_49") & scCurSign & money(pGiftUsed) & "<br>" & vbcrlf & "<br>" & vbcrlf
			pGCAmount=rs19("pcGO_Amount")
			if cdbl(pGCAmount)<=0 then
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_50") & "<br>" & vbcrlf
			else
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_51") & scCurSign & money(pGCAmount) & "<br>" & vbcrlf
				pExpDate=rs19("pcGO_ExpDate")
				if year(pExpDate)="1900" then
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_52") & "<br>" & vbcrlf
				else
					if scDateFrmt="DD/MM/YY" then
						pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
					else
						pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
					end if
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_53") & pExpDate & "<br>" & vbcrlf
				end if
				pGCStatus=rs19("pcGO_Status")
				if pGCStatus="1" then
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_54") & dictLanguage.Item(Session("language")&"_sendMail_54a") & "<br>" & vbcrlf
				else
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_54") & dictLanguage.Item(Session("language")&"_sendMail_54b") & "<br>" & vbcrlf
				end if
			end if
			CustomerEmail=customerEmail & "<br>" & vbcrlf & "<br>" & vbcrlf
		end if 'not rs19.eof
		set rs19=nothing
	end if 'not rs.eof
	set rs=nothing
end if
Next
	CustomerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf & "<br>" & vbcrlf
END IF

'GGG Add-on end

'********************************************************
'* START - DIGITAL PRODUCTS
'* If the order contains digital products and the order
'* status is "processed", then include license/link info
'********************************************************

	IF DPOrder="1" AND pOrderStatus="3" then
		query="select IdProduct from DPRequests WHERE IdOrder=" & qry_ID & ";"
		pidorder=qry_ID
		set rs11=connTemp.execute(query)
		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rs11=nothing
			'//close any connections
			call closedb()
			'//redirect to error page
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		do while not rs11.eof
			query="select * from Products,DProducts where products.idproduct=" & rs11("idproduct") & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
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
			if not rs.eof then
				pIdproduct=rs11("idproduct")
				pProductName=rs("Description")
				pURLExpire=rs("URLExpire")
				pExpireDays=rs("ExpireDays")
				pLicense=rs("License")
				pLL1=rs("LicenseLabel1")
				pLL2=rs("LicenseLabel2")
				pLL3=rs("LicenseLabel3")
				pLL4=rs("LicenseLabel4")
				pLL5=rs("LicenseLabel5")
				pAddtoMail=rs("AddtoMail")

				query="select RequestSTR from DPRequests where idproduct=" & rs11("idproduct") & " and idorder=" & pidorder & " and idcustomer=" & pidcustomer
				set rs19=connTemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rs19=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				pdownloadStr=rs19("RequestSTR")

				SPath1=Request.ServerVariables("PATH_INFO")
				mycount1=0
				do while mycount1<2
					if mid(SPath1,len(SPath1),1)="/" then
						mycount1=mycount1+1
					end if
					if mycount1<2 then
						SPath1=mid(SPath1,1,len(SPath1)-1)
					end if
				loop
				SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1

				if Right(SPathInfo,1)="/" then
					pdownloadStr=SPathInfo & "pc/pcdownload.asp?id=" & pdownloadStr
				else
					pdownloadStr=SPathInfo & "/pc/pcdownload.asp?id=" & pdownloadStr
				end if

				CustomerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf & "<br>" & vbcrlf

				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_28") & pProductName & "<br>" & vbcrlf & "<br>" & vbcrlf
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_29")
				if (pURLExpire<>"") and (pURLExpire="1") then
					if date()-(CDate(pprocessDate)+pExpireDays)<0 then
						CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_30") & (CDate(pprocessDate)+pExpireDays)-date() & dictLanguage.Item(Session("language")&"_sendMail_31") & "<br>" & vbcrlf & "<br>" & vbcrlf
					else
						if date()-(CDate(pprocessDate)+pExpireDays)=0 then
							CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_32") & "<br>" & vbcrlf & "<br>" & vbcrlf
						else
							CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_33") & "<br>" & vbcrlf & "<br>" & vbcrlf
						end if
					end if
				else
					CustomerEmail=CustomerEmail & ":" & "<br>" & vbcrlf & "<br>" & vbcrlf
				end if
				CustomerEmail=CustomerEmail & pdownloadStr & "<br>" & vbcrlf & "<br>" & vbcrlf
				CustomerEmail=CustomerEmail & dictLanguage.Item(Session("language")&"_DownloadURLNote_1") & "<br>" & vbcrlf & "<br>" & vbcrlf

				if (pLicense<>"") and (pLicense="1") then
					query="select * from DPLicenses where idproduct=" & rs11("idproduct") & " and idorder=" & pidorder
					set rs19=connTemp.execute(query)
					if err.number<>0 then
						'//Logs error to the database
						call LogErrorToDatabase()
						'//clear any objects
						set rs19=nothing
						'//close any connections
						call closedb()
						'//redirect to error page
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					TempLicStr=""
					do while not rs19.eof
						TempLic=""
						Lic1=rs19("Lic1")
						if Lic1<>"" then
							TempLic=TempLic & pLL1 & ": " & Lic1 & "<br>" & vbcrlf
						end if
						Lic2=rs19("Lic2")
						if Lic2<>"" then
							TempLic=TempLic & pLL2 & ": " & Lic2 & "<br>" & vbcrlf
						end if
						Lic3=rs19("Lic3")
						if Lic3<>"" then
							TempLic=TempLic & pLL3 & ": " & Lic3 & "<br>" & vbcrlf
						end if
						Lic4=rs19("Lic4")
						if Lic4<>"" then
							TempLic=TempLic & pLL4 & ": " & Lic4 & "<br>" & vbcrlf
						end if
						Lic5=rs19("Lic5")
						if Lic5<>"" then
							TempLic=TempLic & pLL5 & ": " & Lic5 & "<br>" & vbcrlf
						end if
						if TempLic<>"" then
							TempLic=TempLic & "<br>" & vbcrlf
							TempLicStr=TempLicStr & TempLic
						end if
					rs19.movenext
					loop
					if TempLicStr<>"" then
						TempLicStr=dictLanguage.Item(Session("language")&"_sendMail_34") & "<br>" & vbcrlf & "<br>" & vbcrlf & TempLicStr
						CustomerEmail=customerEmail & TempLicStr & "<br>" & vbcrlf
					end if
				end if

				if pAddtoMail<>"" then
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_35") & "<br>" & vbcrlf & "<br>" & vbcrlf & pAddtoMail & "<br>" & vbcrlf & "<br>" & vbcrlf
				end if
			end if
		rs11.MoveNext
		loop
		CustomerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf & "<br>" & vbcrlf
	end if

'********************************************************
'* END - DIGITAL PRODUCTS
'********************************************************


'********************************************************
'* START - GIFT CERTIFICATES
'* Include Gift Certificate information if the order
'* has been processed
'********************************************************

	IF pGCs="1" AND pOrderStatus="3" then
		query="select idproduct from ProductsOrdered WHERE idOrder="& qry_ID
		pidorder=qry_ID
		set rs11=connTemp.execute(query)
		do while not rs11.eof
			query="select products.Description,pcGCOrdered.pcGO_GcCode from Products,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& qry_ID
			set rs=connTemp.execute(query)

			if not rs.eof then
				pIdproduct=rs11("idproduct")
				pName=rs("Description")
				pCode=rs("pcGO_GcCode")
				CustomerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf & "<br>" & vbcrlf
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_42") & "<br>" & vbcrlf
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_43") & pName & "<br>" & vbcrlf

					'// START - Gift Certificate Recipient information
					query="select pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg from Orders WHERE idOrder="& pidorder
					set rs20=Server.CreateObject("ADODB.Recordset")
					set rs20=connTemp.execute(query)
					if not rs20.eof then
						pcvGcRecipientName=rs20("pcOrd_GcReName")
						if pcvGcRecipientName<>"" then
							customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_NotifyRe_3") & pcvGcRecipientName & "<br>" & vbcrlf
						end if
						pcvGcRecipientEmail=rs20("pcOrd_GcReEmail")
						if pcvGcRecipientEmail<>"" then
							customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_NotifyRe_4") & pcvGcRecipientEmail & "<br>" & vbcrlf
						end if
						customerEmail=customerEmail & "<br>" & vbcrlf
					end if
					set rs20=nothing
					'// END - Gift Certificate Recipient information

					'// START - Gift Certificate code(s)
					query="select pcGO_GcCode,pcGO_ExpDate from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & pidorder
					set rs19=connTemp.execute(query)

					do while not rs19.eof
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_44") & rs19("pcGO_GcCode") & "<br>" & vbcrlf
					pExpDate=rs19("pcGO_ExpDate")

					if year(pExpDate)="1900" then
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_45b") & "<br>" & vbcrlf & "<br>" & vbcrlf
					else
					if scDateFrmt="DD/MM/YY" then
					pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
					else
					pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
					end if
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_45") & pExpDate & "<br>" & vbcrlf & "<br>" & vbcrlf
					end if
					rs19.movenext
					loop
					CustomerEmail=customerEmail & "<br>" & vbcrlf
					'// END - Gift Certificate code(s)

			end if
		rs11.MoveNext
		loop
	CustomerEmail=customerEmail & FixedField(80, "R", "====================================================================================================") & "<br>" & vbcrlf & "<br>" & vbcrlf
	END IF
'********************************************************
'* END - GIFT CERTIFICATES
'********************************************************

'Start SDBA
'Back-Ordered products Area%>
<!--#include file="inc_BackOrderEmail.asp"-->
<%
customerEmail=customerEmail & pcv_BackOrderStr
'Create a link to receive customer confirmation about separate shipments
if (scAllowSeparate="1") and (pcv_BackOrderStr<>"") and (pcv_CustomerReceived=0) then
	strPath=Request.ServerVariables("PATH_INFO")
	iCnt=0
	do while iCnt<1
		if mid(strPath,len(strPath),1)="/" then
			iCnt=iCnt+1
		end if
		if iCnt<1 then
			strPath=mid(strPath,1,len(strPath)-1)
		end if
	loop

	strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath

	if Right(strPathInfo,1)="/" then
	else
		strPathInfo=strPathInfo & "/"
	end if

	DO
		Tn1=""
		For dd=1 to 100
			Randomize
			myC=Fix(3*Rnd)
			Select Case myC
				Case 0:
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
				Case 1:
					Randomize
					Tn1=Tn1 & Cstr(Fix(10*Rnd))
				Case 2:
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+97)
			End Select
		Next

		ReqExist=0

		query="SELECT IDOrder FROM Orders WHERE pcOrd_CustRequestStr='" & Tn1 & "'"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rstemp=nothing
			'//close any connections
			call closedb()
			'//redirect to error page
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rstemp.eof then
			ReqExist=1
		end if
		set rstemp=nothing
	LOOP UNTIL ReqExist=0

	query="Update Orders Set pcOrd_CustRequestStr='" & Tn1 & "' WHERE idorder=" & qry_ID
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rstemp=nothing
		'//close any connections
		call closedb()
		'//redirect to error page
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rstemp=nothing

	strPathInfo=strPathInfo & "sds_AllowSeparateShip.asp?req=" & Tn1

	'Add request link to the Customer Confirmation E-mail
	if pcv_CustomerReceived=0 then
	customerEmail=customerEmail & ship_dictLanguage.Item(Session("language")&"_custconfirm_msg_1") & "<br>" & vbcrlf
	customerEmail=customerEmail & strPathInfo & "<br>" & vbcrlf & "<br>" & vbcrlf
	end if

end if
'End SDBA

customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_36") & scCompanyName & "." & "<br>" & vbcrlf & "<br>" & vbcrlf
CustomerEmail=replace(CustomerEmail,"''",chr(39))

'GGG Add-on start
'Del bookmarks
tempMail=split(customerEmail,"AAAAAAAAAA")
customerEmail=replace(customerEmail,"AAAAAAAAAA","")
'GGG Add-on end

'********************************************************
'* START - GIFT REGISTRY - NOTIFY REGISTRY OWNER
'* Send mail to Gift Registry owner if someone purchased
'* gifts for them and the order has been processed
'********************************************************

	IF gIDEvent<>"0" AND pOrderStatus="3" then
		RegEmail=""
		RegEmail=dictLanguage.Item(Session("language")&"_sendMail_58") & gReg & "," & "<br>" & vbcrlf & "<br>" & vbcrlf
		RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_59") & "<br>" & vbcrlf
		RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_55") & geName & "<br>" & vbcrlf
		RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_56") & geDate & "<br>" & vbcrlf
		RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_60") & pCustomerFullName & "<br>" & vbcrlf
		RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_61") & "<br>" & vbcrlf
		RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_62") & scpre+int(qry_ID) & "<br>" & vbcrlf
		RegEmail=RegEmail & tempMail(1) & "<br>" & vbcrlf
		RegEmail=RegEmail & scCompanyName & "<br>" & vbcrlf & "<br>" & vbcrlf
		if geNotify="1" then
			RegEmail = pcf_HtmlEmailWrapper(RegEmail, pcv_HTMLEmailFontFamily)
			call sendmail (scCompanyName, scEmail, gRegemail, "Someone purchased some gifts off your Gift Registry", RegEmail)
			call pcs_hookGROrderEmailSent(gRegemail)
		end if
	END IF

'********************************************************
'* END - GIFT REGISTRY - NOTIFY REGISTRY OWNER
'********************************************************


'********************************************************
'* START - GIFT CERTIFICATES - RECIPIENT NOTIFICATION
'* Send email to recipient of Gift Certificate, if any
'* and if the order has been processed
'********************************************************
	ReciEmail=""
	IF pGCs="1" AND pOrderStatus="3" THEN

		query="select idproduct from ProductsOrdered WHERE idOrder="& qry_ID
		pidorder=qry_ID
		set rs11=connTemp.execute(query)
		do while not rs11.eof
			query="select products.Description,pcGCOrdered.pcGO_GcCode,pcGc.pcGc_EOnly from Products,pcGc,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGC.pcGc_IDProduct=products.idproduct and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& qry_ID
			set rs=connTemp.execute(query)

			if not rs.eof then
				pIdproduct=rs11("idproduct")
				pName=rs("Description")
				pCode=rs("pcGO_GcCode")
				pEOnly=rs("pcGc_EOnly")

					query="select pcGO_Amount,pcGO_GcCode,pcGO_ExpDate from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & pidorder
					set rs19=connTemp.execute(query)

					do while not rs19.eof
					pAmount=rs19("pcGO_Amount")
					if pAmount<>"" then
					else
					pAmount="0"
					end if

					ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_68") & scCurSign & money(pAmount) & "<br>" & vbcrlf

					ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_69") & rs19("pcGO_GcCode") & "<br>" & vbcrlf
					pExpDate=rs19("pcGO_ExpDate")

					if year(pExpDate)="1900" then
					ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_45b") & "<br>" & vbcrlf
					else
					if scDateFrmt="DD/MM/YY" then
					pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
					else
					pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
					end if
					ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_70") & pExpDate & "<br>" & vbcrlf
					end if
					if pEOnly="1" then
					ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_71") & "<br>" & vbcrlf
					end if
					ReciEmail=ReciEmail & "<br>" & vbcrlf
					rs19.movenext
					loop

			end if
		rs11.MoveNext
		loop

		query="select pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg from Orders WHERE idOrder="& qry_ID
		set rs11=connTemp.execute(query)

		GcReName=rs11("pcOrd_GcReName")
		GcReEmail=rs11("pcOrd_GcReEmail")
		GcReMsg=rs11("pcOrd_GcReMsg")
		pCustomerFullNamePlusEmail=pCustomerFullName & " (" & pEmail & ")"

		if GcReEmail<>"" then
			if not GcReName<>"" then
				GcReName=GcReEmail
			end if
			ReciEmail1=replace(dictLanguage.Item(Session("language")&"_sendMail_66"),"<recipient name>",GcReName)
			ReciEmail2=replace(dictLanguage.Item(Session("language")&"_sendMail_67"),"<customer name>",pCustomerFullNamePlusEmail)
			'ReciEmail2=replace(ReciEmail2,"<br>",VbCrLf)
			if GcReMsg<>"" then
				ReciEmail3=replace(dictLanguage.Item(Session("language")&"_sendMail_72"),"<customer name>",pCustomerFullName) & "<br>" & vbcrlf & GcReMsg & "<br>" & vbcrlf
			else
				ReciEmail3=""
			end if
			
			ReciEmail=ReciEmail1 & "<br>" & vbcrlf & "<br>" & vbcrlf & ReciEmail2 & "<br>" & vbcrlf & "<br>" & vbcrlf & ReciEmail & ReciEmail3
			ReciEmail=ReciEmail & "<br>" & vbcrlf & scCompanyName & "<br>" & vbcrlf & scStoreURL & "<br>" & vbcrlf & "<br>" & vbcrlf
			ReciEmail = pcf_HtmlEmailWrapper(ReciEmail, pcv_HTMLEmailFontFamily)
			call sendmail (scCompanyName, scEmail, GcReEmail,pCustomerFullName & dictLanguage.Item(Session("language")&"_sendMail_73"), ReciEmail)
			call pcs_hookGCOrderEmailSent(GcReEmail)
		end if

	END IF
'********************************************************
'* END - GIFT CERTIFICATES - RECIPIENT NOTIFICATION
'********************************************************
%>
