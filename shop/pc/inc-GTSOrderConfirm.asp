<%
gtsPageLang="en_US"
gtsShopID=""
gtsShopAccID=""
gtsCountry="US"
gtsLang="en"
gtsCur="USD"
gtsShipDays=1
gtsHost=Request.ServerVariables("SERVER_NAME")
queryQ="SELECT pcGTS_TurnOn,pcGTS_AccNo,pcGTS_PageLang,pcGTS_ShopAccID,pcGTS_ShopCountry,pcGTS_ShopLang,pcGTS_Currency,pcGTS_ShipDays,pcGTS_DeDays FROM pcGoogleTS;"
set rsQ=connTemp.execute(queryQ)
if not rsQ.eof then
	gtsTurnOn=rsQ("pcGTS_TurnOn")
	gtsAccNo=rsQ("pcGTS_AccNo")
	gtsPageLang=rsQ("pcGTS_PageLang")
	gtsShopAccID=rsQ("pcGTS_ShopAccID")
	gtsCountry=rsQ("pcGTS_ShopCountry")
	gtsLang=rsQ("pcGTS_ShopLang")
	gtsCur=rsQ("pcGTS_Currency")
	gtsShipDays=rsQ("pcGTS_ShipDays")
	gtsDeDays=rsQ("pcGTS_DeDays")
end if
set rsQ=nothing

if (gtsTurnOn="1") then

	'Get Order Details
	gtsDeDate=""
	query="SELECT customers.email,orders.countryCode, orders.total, orders.discountDetails, orders.shipmentDetails, orders.taxamount, orders.ord_DeliveryDate FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	if not rs.eof then
		gtsEmail=rs("email")
		gtsCountry=rs("countryCode")
		gtsTotal=ccur(rs("total"))
		pdiscountDetails=rs("discountDetails")
		pshipmentDetails=rs("shipmentDetails")
		gtsTax=rs("taxamount")
		if IsNull(gtsTax) OR gtsTax="" then
			gtsTax=0
		end if
		gtsTax=ccur(gtsTax)
		gtsDeDate=rs("ord_DeliveryDate")
		set rs=nothing
		if (gtsDeDate<>"") then
			gtsYear=Year(gtsDeDate)
			gtsMon=Month(gtsDeDate)
			if Clng(gtsMon)<10 then
				gtsMon="0" & gtsMon
			end if
			gtsDay=Day(gtsDeDate)
			if Clng(gtsDay)<10 then
				gtsDay="0" & gtsDay
			end if
			gtsDeDate=gtsYear & "-" & gtsMon & "-" & gtsDay
		end if
		
		'// Discounts
		gtsDiscounts=0
		if instr(pdiscountDetails,",") then
			DiscountDetailsArry=split(pdiscountDetails,",")
			intArryCnt = ubound(DiscountDetailsArry)
		else
			intArryCnt = 0
		end if
		For k=0 to intArryCnt
			if intArryCnt=0 then
				pTempDiscountDetails=pdiscountDetails
			else
				pTempDiscountDetails=DiscountDetailsArry(k)
			end if
	
			if instr(pTempDiscountDetails,"- ||") then
				pcv_arryDiscounts = split(pTempDiscountDetails,"- ||")
				discountPrice = pcv_arryDiscounts(1)
				if IsNull(discountPrice) OR discountPrice="" then
					discountPrice=0
				end if
				gtsDiscounts=gtsDiscounts+discountPrice
			end if
		Next
		gtsDiscounts=ccur(-1*gtsDiscounts)
		
		'//Ship Amount
		gtsShipAmount=0
		shipping = split(pshipmentDetails,",")
		if ubound(shipping)>1 then
			if NOT isNumeric(trim(shipping(2))) then
				gtsShipAmount="0"
			else
				Shipper=shipping(0)
				Service=shipping(1)
				gtsShipAmount=trim(shipping(2))
				if ubound(shipping)=>3 then
					serviceHandlingFee=trim(shipping(3))
					if NOT isNumeric(serviceHandlingFee) then
						serviceHandlingFee=0
					end if
				else
					serviceHandlingFee=0
				end if
				gtsShipAmount=Cdbl(gtsShipAmount)+cdbl(serviceHandlingFee)
			end if
		else
			gtsShipAmount="0"
		end if
		gtsShipAmount=ccur(gtsShipAmount)
	end if
	set rs=nothing
	
	'// Check Back-Orders
	gtsBackOrder="N"
	pcvShipNDays=0
	query="SELECT ProductsOrdered.idproduct,Products.pcProd_ShipNDays FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idOrder="& pIdOrder & " AND (((Products.pcProd_BackOrder=1) AND (Products.stock<0)) OR (ProductsOrdered.pcPrdOrd_BackOrder=1));"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		gtsBackOrder="Y"
		pcvShipNDays=rs("pcProd_ShipNDays")
	end if
	set rs=nothing
	
	if gtsBackOrder="Y" then
		query="SELECT ProductsOrdered.idproduct,Products.stock FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idOrder="& pIdOrder & " AND ((Products.stock>0) OR (Products.noStock<>0));"
		set rs=connTemp.execute(query)
	
		if rs.eof then
			if Clng(pcvShipNDays)>0 then
				gtsShipDays=pcvShipNDays
			end if
		end if
		set rs=nothing
	end if
	
	'// Est Ship Date
	gtsShipDate=""
	if IsNull(gtsShipDays) OR gtsShipDays="" then
		gtsShipDays=1
	end if

	gtsShipDate=Date()+Clng(gtsShipDays)
	if (gtsShipDate<>"") then
		gtsYear=Year(gtsShipDate)
		gtsMon=Month(gtsShipDate)
		if Clng(gtsMon)<10 then
			gtsMon="0" & gtsMon
		end if
		gtsDay=Day(gtsShipDate)
		if Clng(gtsDay)<10 then
			gtsDay="0" & gtsDay
		end if
		gtsShipDate=gtsYear & "-" & gtsMon & "-" & gtsDay
	end if
	
	'// Est Delivery Date
	if IsNull(gtsDeDate) OR (gtsDeDate="") then
		if IsNull(gtsDeDays) OR gtsDeDays="" then
			gtsDeDays=1
		end if
	
		gtsDeDate=Date()+(Clng(gtsShipDays)+Clng(gtsDeDays))
	end if
	if (gtsDeDate<>"") then
		gtsYear=Year(gtsDeDate)
		gtsMon=Month(gtsDeDate)
		if Clng(gtsMon)<10 then
			gtsMon="0" & gtsMon
		end if
		gtsDay=Day(gtsDeDate)
		if Clng(gtsDay)<10 then
			gtsDay="0" & gtsDay
		end if
		gtsDeDate=gtsYear & "-" & gtsMon & "-" & gtsDay
	end if
	
	'// Check Digital Goods
	gtsDigital="N"
	query="SELECT ProductsOrdered.idproduct,Products.noshipping,Products.Downloadable,Products.pcprod_GC FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idOrder="& pIdOrder & " AND ((Products.noshipping<>0) OR (Products.Downloadable<>0) OR (Products.pcprod_GC<>0));"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		gtsDigital="Y"
	end if
	set rs=nothing
	
	query="SELECT ProductsOrdered.idProduct, Products.Description, ProductsOrdered.quantity, ProductsOrdered.unitPrice FROM Products INNER JOIN ProductsOrdered ON Products.idProduct=ProductsOrdered.idProduct WHERE ProductsOrdered.idOrder=" & pIdOrder & ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	intCount=-1
	if not rs.eof then
		tmpArr=rs.getRows()
		intCount=ubound(tmpArr,2)
	end if
	set rs=nothing
	%>
	<!-- START Google Trusted Stores Order -->
	<div id="gts-order" style="display:none;" translate="no">
	<!-- start order and merchant information -->
	<span id="gts-o-id"><%=(int(pIdOrder)+scpre)%></span>
	<span id="gts-o-domain"><%=gtsHost%></span>
	<span id="gts-o-email"><%=gtsEmail%></span>
	<span id="gts-o-country"><%=gtsCountry%></span>
	<span id="gts-o-currency"><%=gtsCur%></span>
	<span id="gts-o-total"><%=gtsTotal%></span>
	<span id="gts-o-discounts"><%=gtsDiscounts%></span>
	<span id="gts-o-shipping-total"><%=gtsShipAmount%></span>
	<span id="gts-o-tax-total"><%=gtsTax%></span>
	<span id="gts-o-est-ship-date"><%=gtsShipDate%></span>
	<span id="gts-o-est-delivery-date"><%=gtsDeDate%></span>
	<span id="gts-o-has-preorder"><%=gtsBackOrder%></span>
	<span id="gts-o-has-digital"><%=gtsDigital%></span>
	<!-- end order and merchant information -->
	<!-- start repeated item specific information -->
	<%if intCount>-1 then
	For co=0 to intCount%>
		<span class="gts-item">
		<span class="gts-i-name"><%=tmpArr(1,co)%></span>
		<span class="gts-i-price"><%=ccur(tmpArr(3,co))%></span>
		<span class="gts-i-quantity"><%=tmpArr(2,co)%></span>
		<%if gtsShopAccID<>"" then%>
		<span class="gts-i-prodsearch-store-id"><%=gtsShopAccID%></span>
		<%end if%>
		<span class="gts-i-prodsearch-country"><%=gtsCountry%></span>
		<span class="gts-i-prodsearch-language"><%=gtsLang%></span>
		</span>
	<%Next
	end if%>
	<!-- end repeated item specific information -->
	</div>
	<!-- END Google Trusted Stores Order -->
<%end if%>

