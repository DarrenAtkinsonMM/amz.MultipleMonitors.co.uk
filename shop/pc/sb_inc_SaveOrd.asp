<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by Netsource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of Netsource Commerce.
'Copyright 2001-2006. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'Netsource Commerce. To contact Netsource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<%
'///////////////////////////////////////////////////////////////////////////////////
'//  START: ORDER PROCESSING
'///////////////////////////////////////////////////////////////////////////////////

'***********************************************************************************
' START: Get info we need... segway into PC "SaveOrd.asp" code for managability
'**********************************************************************************
qry_ID=SBIDOrder

query="SELECT pcOrd_Archived,idcustomer, orderdate, Address, city, state, stateCode, zip, CountryCode, paymentDetails, shipmentDetails, shippingAddress, shippingCity, shippingStateCode, shippingState, shippingZip, pcOrd_shippingPhone, pcOrd_ShippingEmail, shippingCountryCode, idAffiliate, affiliatePay, discountDetails, pcOrd_GCDetails,pcOrd_GCAmount, taxAmount,  total, comments, orderStatus, processDate, shipDate, shipvia, trackingNum, returnDate, returnReason, ShippingFullName, ord_DeliveryDate, ord_OrderName, iRewardPoints, iRewardPointsCustAccrued, iRewardValue, address2, shippingCompany, shippingAddress2, taxDetails, adminComments, rmaCredit, DPs, gwAuthCode, gwTransId, gwTransParentId, paymentCode, SRF, ordShipType, ordPackageNum, ord_VAT, pcOrd_CatDiscounts, pcOrd_Payer, pcOrd_PaymentStatus, pcOrd_CustAllowSeparate, pcOrd_CustRequestStr, pcOrd_GCs, pcOrd_GcCode, pcOrd_GcUsed, pcOrd_IDEvent, pcOrd_GWTotal, pcOrd_Time, pcOrd_ShipWeight, pcOrd_GoogleIDOrder, pcOrd_CustomerIP, pcOrd_EligibleForProtection, pcOrd_AVSRespond, pcOrd_CVNResponse, pcOrd_PartialCCNumber, pcOrd_BuyerAccountAge, pcOrd_OrderKey FROM orders WHERE idOrder=" & qry_ID & ";"

Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)

pidorder=qry_ID
pOrdArc=rs("pcOrd_Archived")
if pOrdArc="" OR IsNULL(pOrdArc) then
	pOrdArc=0
end if
pidcustomer=rs("idcustomer")
pcv_strBillingAddress1=rs("Address")
pcv_strBillingCity=rs("city")
pcv_BillingState=rs("state")
pcv_BillingStateCode=rs("stateCode")
pcv_strBillingPostalCode=rs("zip")
pcv_strBillingCountryCode=rs("CountryCode")
ppaymentDetails=trim(rs("paymentDetails"))
pshipmentDetails=rs("shipmentDetails")
pcv_strShippingAddress1=rs("shippingAddress")
pcv_strShippingCity=rs("shippingCity")
pcv_ShippingStateCode=rs("shippingStateCode")
pcv_ShippingState=rs("shippingState")
pcv_strShippingPostalCode=rs("shippingZip")
pcv_strShippingPhone=rs("pcOrd_shippingPhone")
pShippingEmail=rs("pcOrd_ShippingEmail")
pcv_strShippingCountryCode=rs("shippingCountryCode")
pidAffiliate=rs("idaffiliate")
paffiliatePay=rs("affiliatePay")
pdiscountDetails=rs("discountDetails")
GCDetails=rs("pcOrd_GCDetails")
GCAmount=rs("pcOrd_GCAmount")
if GCAmount="" OR IsNull(GCAmount) then
	GCAmount=0
end if
ptaxAmount=rs("taxAmount")
ptotal=rs("total")
pcomments=rs("comments")
porderStatus=rs("orderStatus")
pprocessDate=rs("processDate")
pprocessDate=ShowDateFrmt(pprocessDate)
pshipDate=rs("shipDate")
pshipDate=ShowDateFrmt(pshipDate)
pshipvia=rs("shipvia")
ptrackingNum=rs("trackingNum")
preturnDate=rs("returnDate")
preturnDate=ShowDateFrmt(preturnDate)
preturnReason=rs("returnReason")
pShippingFullName=rs("ShippingFullName")
pord_DeliveryDate=rs("ord_DeliveryDate")
pord_OrderName=rs("ord_OrderName")
if isNULL(pord_OrderName) OR pord_OrderName="" then
	pord_OrderName="No Name"
end if
piRewardPoints=rs("iRewardPoints")
piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
piRewardValue=rs("iRewardValue")
pcv_strBillingAddress2=rs("address2")
pcv_strShippingCompanyName=rs("shippingCompany")
pcv_strShippingAddress2=rs("shippingAddress2")
ptaxDetails=rs("taxDetails")
padminComments=rs("adminComments")
prmaCredit=rs("rmaCredit")
pcDPs=rs("DPs")
pcgwAuthCode=rs("gwAuthCode")
pcgwTransId=rs("gwTransId")
pcgwTransParentId=rs("gwTransParentId")
pcpaymentCode=rs("paymentCode")
pSRF=rs("SRF")

pOrdShipType=rs("ordShipType")

pOrdPackageNum=rs("ordPackageNum")
pVATTotal=rs("ord_VAT")
pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
	pcv_CatDiscounts="0"
end if
pcOrd_Payer=rs("pcOrd_Payer")
pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
if isNULL(pcv_PaymentStatus) OR pcv_PaymentStatus="" then
	pcv_PaymentStatus="0"
end if

pcv_CustAllow=rs("pcOrd_CustAllowSeparate")
	if isNULL(pcv_CustAllow) or pcv_CustAllow="" then
		pcv_CustAllow="0"
	end if

pcv_CustRequestStr=rs("pcOrd_CustRequestStr")
	if isNULL(pcv_CustRequestStr) or pcv_CustRequestStr="" then
		pcv_CustRequestStr="NA"
	end if

'GGG Add-on start
pGCs=rs("pcOrd_GCs")
pGiftCode=rs("pcOrd_GcCode")
pGiftUsed=rs("pcOrd_GcUsed")
gIDEvent=rs("pcOrd_IDEvent")
if gIDEvent<>"" then
else
	gIDEvent="0"
end if
pGWTotal=rs("pcOrd_GWTotal")
if pGWTotal<>"" then
else
	pGWTotal="0"
end if
'GGG Add-on end

'------------------------------
' Order time: retrieve and format
pcv_OrderTime=rs("pcOrd_Time")
if pcv_OrderTime<>"" and not isNull(pcv_OrderTime) then
	if scDateFrmt="DD/MM/YY" then
		pcv_OrderTime = FormatDateTime(pcv_OrderTime, 4)
	else
		pcv_OrderTime = FormatDateTime(pcv_OrderTime, 3)
	end if
else
	pcv_OrderTime=""
end if
'------------------------------
pcOrd_ShipWeight=rs("pcOrd_ShipWeight")
pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder")
pcv_strCustomerIP = rs("pcOrd_CustomerIP")
pcv_strEligibleForProtection = rs("pcOrd_EligibleForProtection")
pcv_strAVSRespond = rs("pcOrd_AVSRespond")
pcv_strCVNResponse = rs("pcOrd_CVNResponse")
pcv_strPartialCCNumber = rs("pcOrd_PartialCCNumber")
pcv_strBuyerAccountAge = rs("pcOrd_BuyerAccountAge")
pcOrderKey=rs("pcOrd_OrderKey")
set rs=nothing

'Is Mixed Order
query="Select SUM((quantity*unitPrice)-QDiscounts) As TotalPrd FROM dbo.ProductsOrdered WHERE idOrder=" & qry_ID &" AND ProductsOrdered.pcSubscription_ID=0;"
set rs=connTemp.execute(query)

if not rs.eof then
	set rs=nothing
	query="Select SUM((quantity*unitPrice)-QDiscounts) As TotalPrd FROM dbo.ProductsOrdered WHERE idOrder=" & qry_ID &" AND ProductsOrdered.pcSubscription_ID>0;"
	set rs=connTemp.execute(query)
	if not (IsNull(rs("TotalPrd"))) then
		pTotal=Cdbl(rs("TotalPrd"))
		pshipmentDetails="No shipping charge (or no shipping required)."
		ptaxDetails=""
		ptaxAmount=0
	end if
	
end if

set rs=nothing

'// Calculate total adjusted for credits
if trim(prmaCredit)="" or IsNull(prmaCredit) then
	prmaCredit=0
end if
pTotalAdj=pTotal-prmaCredit
pTotal=pTotal-prmaCredit

'// Check if the Customer is European Union
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)


set rs=nothing

query="SELECT [name],lastName,pcCust_Guest,customerCompany,phone,email,customerType,idrefer,fax FROM customers WHERE idcustomer="& pidcustomer
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)
if not rs.eof then
	pcv_strFirstName=rs("name")
	pcv_strLastName=rs("lastName")
	pcvGuest=rs("pcCust_Guest")
	if pcvGuest="" OR IsNull(pcvGuest) then
		pcvGuest=0
	end if
	pcv_strBillingCompanyName=rs("customerCompany")
	pcv_strBillingPhone=rs("phone")
	pcv_strBillingEmail=rs("email")
	pcv_strBillingFax=rs("fax")
end if
set rs=nothing



pcv_strFirstName=replace(pcv_strFirstName,"''","'")
pcv_strLastName=replace(pcv_strLastName,"''","'")
pcv_strBillingCompanyName=replace(pcv_strBillingCompanyName,"''","'")
pcv_strBillingAddress1=replace(pcv_strBillingAddress1,"''","'")
pcv_strBillingAddress2=replace(pcv_strBillingAddress2,"''","'")
pcv_strBillingCity=replace(pcv_strBillingCity,"''","'")

'// Billing
pFirstName=getUserInput(pcv_strFirstName,0)
pLastName=getUserInput(pcv_strLastName,0)
pCustomerCompany=getUserInput(pcv_strBillingCompanyName,100)
pPhone=pcv_strBillingPhone
pEmail=pcv_strBillingEmail
pAddress=getUserInput(pcv_strBillingAddress1,0)
pZip=pcv_strBillingPostalCode
pStateCode=pcv_BillingStateCode
pState=getUserInput(pcv_BillingState,0)
pCity=getUserInput(pcv_strBillingCity,0)
pCountryCode=pcv_strBillingCountryCode
pAddress2=getUserInput(pcv_strBillingAddress2,0)
pFax=pcv_strBillingFax

'// Shipping
pShippingCompany=getUserInput(pcv_strShippingCompanyName,0)
pShippingAddress=getUserInput(pcv_strShippingAddress1,0)
pShippingAddress2=getUserInput(pcv_strShippingAddress2,0)
pShippingCity=getUserInput(pcv_strShippingCity,0)
pShippingStateCode=getUserInput(pcv_ShippingStateCode,0)
pShippingState=getUserInput(pcv_ShippingState,0)
pShippingZip=getUserInput(pcv_strShippingPostalCode,0)
pShippingCountryCode=getUserInput(pcv_strShippingCountryCode,0)
pShippingPhone=getUserInput(pcv_strShippingPhone,0)

if pZip="" then
	pZip="NA"
end if
if pShippingZip="" then
	pShippingZip="NA"
end if
'***********************************************************************************
' END: Get info from sessions and customers
'***********************************************************************************



'***********************************************************************************
' START: Order Information
'***********************************************************************************

'// Package Details
if pOrdPackageNum="" then
	pOrdPackageNum=1
end if

'// Misc.
pIDRefer=0
pRewardsBalance=0

pShipping=pshipmentDetails
pComments=""
pShippingReferenceId="0"
pShippingFax=""

'// Discount Flag
pDiscountUsed=""
'***********************************************************************************
' END: Order Information
'***********************************************************************************


'***********************************************************************************
' START: ADDITIONAL ORDER INFORMATION
'***********************************************************************************
pIdPayment= 0

ptaxDetailsString = ptaxDetails

'// Rewards Total
If session("pcSFRewardsDollarValue")<>"" then
	piRewardValue = session("pcSFRewardsDollarValue")
	session("pcSFRewardsDollarValue")=""
End if
'***********************************************************************************
' END: ADDITIONAL ORDER INFORMATION
'***********************************************************************************

pRewardReferral=0
pRewardRefId=0

'***********************************************************************************
' START: REBUILD THE CART
'***********************************************************************************
ppcCartIndex=Session("pcCartIndex")
'***********************************************************************************
' END: REBUILD THE CART
'***********************************************************************************


'***********************************************************************************
' START: AFFILIATES
'***********************************************************************************
'// Retrieve affiliate ID from session
pIdAffiliate=session("idAffiliate")
if pIdAffiliate="" then
	pIdAffiliate=1
end if
pAffiliatePay=0
'***********************************************************************************
' END: AFFILIATES
'***********************************************************************************



'***********************************************************************************
' START: VARIABLES
'***********************************************************************************
'// Details
pDetails=Cstr("")
'// Totals
pSubtotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
pCartTotalWeight=int(calculateCartWeight(pcCartArray, ppcCartIndex))
pCartQuantity=int(calculateCartQuantity(pcCartArray, ppcCartIndex))
pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
pAffiliateSubTotal=pSubtotal
'// Date
pDateOrder=Date()
if SQL_Format="1" then
	pDateOrder=Day(pDateOrder)&"/"&Month(pDateOrder)&"/"&Year(pDateOrder)
else
	pDateOrder=Month(pDateOrder)&"/"&Day(pDateOrder)&"/"&Year(pDateOrder)
end if
'// State and Province
If pStateCode <> "" and (pCountryCode="US" or pCountryCode="CA") then
	pState=""
end if
If pShippingStateCode <> "" and (pShippingCountryCode="US" or pShippingCountryCode="CA") then
	pShippingState=""
end if
'***********************************************************************************
' END: VARIABLES
'***********************************************************************************


'***********************************************************************************
' START: COMPILE MEMO FIELD
'***********************************************************************************
for f=1 to ppcCartIndex 
	' if item is not deleted from cart 
	if pcCartArray(f,10) = 0 then 
		tempAmt=Cdbl( pcCartArray(f,2) * (pcCartArray(f,5)+pcCartArray(f,3)) )
		if scDecSign="," then
			tempAmt=replace(tempAmt,",",".")
		end if
		pDetails	= pDetails & "  Amount: ||"& tempAmt & " Qty:" &pcCartArray(f,2)& "  SKU #:" &pcCartArray(f,7) & " - " &pcCartArray(f,1)& " " & pcCartArray(f,4) & Vbcrlf      
		pDetails = replace(pDetails,"'","''")
		pDetails=replace(pDetails,"''''","''")    
	end if ' item deleted
next
'***********************************************************************************
' END: COMPILE MEMO FIELD
'***********************************************************************************

'***********************************************************************************
' START: PAYMENT DETAILS
'***********************************************************************************
if pIdPayment=0 then
	
	pPaymentDetails = "Recurring Payment || 0.00"
	pPaymentDesc="Recurring Payment"
	
end if
'***********************************************************************************
' END: PAYMENT DETAILS
'***********************************************************************************


'***********************************************************************************
' START: REWARD DETAILS
'***********************************************************************************
If piRewardValue="" then
	piRewardValue="0"
End If
If Session("pcSFUseRewards")="" then
	Session("pcSFUseRewards")="0"
End If

'// Save order temporarily
IDrefer=session("IDrefer")
if isNull(IDrefer) OR IDrefer="" then
	IDrefer="0"
end if

pord_DeliveryDate=""

pord_OrderName="Recurring Order"

'***********************************************************************************
' END: REWARD DETAILS
'***********************************************************************************


'***********************************************************************************
' START: GENERATE ORDER INSERT QUERY
'***********************************************************************************
'Generate Order Key
pcOrderKey=""
TestedOrderKey=0
do while (TestedOrderKey=0)
	pcOrderKey=generateABC(3) & generate123(10)
	query="SELECT idOrder FROM Orders WHERE pcOrd_OrderKey like '" & pcOrderKey & "';"
	set rs=connTemp.execute(query)
	if rs.eof then
		TestedOrderKey=1
	end if
	set rs=nothing
loop
strInsertQuery="INSERT INTO orders (pcOrd_OrderKey,IDrefer,orderDate,idCustomer, details, total, taxAmount, comments, address, zip, state, stateCode, city, CountryCode, shippingAddress, shippingZip, shippingState, shippingStateCode, shippingCity, shippingCountryCode, shipmentDetails, paymentDetails, discountDetails, randomNumber, orderStatus, pcOrd_shippingPhone, idAffiliate, affiliatePay,shippingFullName, iRewardPoints, iRewardValue,iRewardPointsCustAccrued, address2, shippingCompany, shippingAddress2,taxDetails,SRF,ordShipType, ordPackageNum, ord_OrderName"
if DFShow="1"  and pord_DeliveryDate <> "" then
	strInsertQuery=strInsertQuery&",ord_DeliveryDate"
end if
strInsertQuery=strInsertQuery&",ord_VAT,pcord_CatDiscounts,pcOrd_DiscountsUsed,pcOrd_GcCode,pcOrd_GcUsed,pcOrd_GCs,pcOrd_IDEvent,pcOrd_GWTotal,pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg,pcOrd_shippingFax, pcOrd_ShippingEmail, pcOrd_ShipWeight) VALUES ('" & pcOrderKey & "'," & IDrefer & ","
if scDB="SQL" then
	strInsertQuery=strInsertQuery&"'" & pDateOrder  & "'"
else
	strInsertQuery=strInsertQuery&"#" & pDateOrder  & "#"
end if

strInsertQuery=strInsertQuery&"," & int(Session("idCustomer"))& ",'" &pDetails &"'," &replacecomma(pTotal)& "," &replacecomma(pTaxAmount)& ",N'" &pComments& "',N'" &paddress & "','" &pzip& "',N'" &pState& "','" &pStateCode& "',N'" &pCity& "','" &pCountryCode& "',N'" &pShippingAddress & "','" &pShippingZip& "',N'" &pShippingState& "','" &pShippingStateCode& "',N'" &pShippingCity& "','" &pShippingCountryCode& "','" &pShipmentDetails& "','" &replace(pPaymentDetails,"'","''")& "','" &replace(pDiscountDetails,"'","''")& "',0, 1,' " &pShippingPhone& "'," &pIdAffiliate& ", " &replacecomma(pAffiliatePay)&",N'"&pShippingFullName&"', "& 0 &", " &piRewardValue&", "& 0 &", N'" &paddress2 & "', N'" &pShippingCompany & "', N'" &pShippingAddress2 & "','"&ptaxDetailsString&"',"&pSRF&","&pOrdShipType&","&pOrdPackageNum&",N'"&pord_OrderName&"'"
if DFShow="1" and pord_DeliveryDate <> "" then
	if scDB="SQL" then
		strInsertQuery=strInsertQuery&",'" & pord_DeliveryDate  & "'"
	else
		strInsertQuery=strInsertQuery&",#" & pord_DeliveryDate  & "#"
	end if
end if
if pVATTotal="" then
	pVATTotal=0
end if
strInsertQuery=strInsertQuery&","&replacecomma(pVATTotal)&"," & pcv_CatDiscounts & ",'"&pDiscountUsed&"','" & pDiscountCode & "'," & GCAmount & ",0," & gIDEvent & "," & pGWTotal & ",N'" & pcv_GcReName & "','" & pcv_GcReEmail & "',N'" & pcv_GcReMsg & "', '"&pShippingFax&"', '"&pShippingEmail&"', "&pShipWeight&")"
'***********************************************************************************
' END: GENERATE ORDER INSERT QUERY
'***********************************************************************************



'***********************************************************************************
' START: RUN ORDER QUERY
'***********************************************************************************		

set rs=server.CreateObject("ADODB.RecordSet")
'// Insert Order
set rs=conntemp.execute(strInsertQuery)
set rs=nothing

'***********************************************************************************
' END: RUN ORDER QUERY
'***********************************************************************************


'***********************************************************************************
' START: GET ORDER ID
'***********************************************************************************
query="SELECT idOrder FROM orders WHERE idCustomer=" & session("idCustomer") & " ORDER BY idOrder DESC;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if not rs.eof then		
	pIdorder=rs("idOrder")
	session("idOrderSaved")=pIdorder
end if
set rs=nothing
'***********************************************************************************
' END: GET ORDER ID
'***********************************************************************************

'***********************************************************************************
' START: ORDER STATUS
'***********************************************************************************
gwAuthCode=""
gwTransID=""
paymentCode= "RecurringPayment"

Todaydate=Date()
if SQL_Format="1" then
	Todaydate=Day(Todaydate)&"/"&Month(Todaydate)&"/"&Year(Todaydate)
else
	Todaydate=Month(Todaydate)&"/"&Day(Todaydate)&"/"&Year(Todaydate)
end if
pOrderTime=Now()

query="UPDATE orders SET pcOrd_PaymentStatus=2,orderstatus=3, processDate='"&Todaydate&"',gwAuthCode='"&gwAuthCode&"',gwTransID='"&gwTransID&"',paymentCode='"&paymentCode&"',pcOrd_Payer='"& session("idCustomer") &"', pcOrd_Time='"&pOrderTime&"' WHERE idOrder=" & pIdOrder

set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

pOrderStatus="3"
pPaymentStatus="2"

call pcs_hookOrderProcessed(pIdOrder)
'***********************************************************************************
' END: ORDER STATUS
'***********************************************************************************


'***********************************************************************************
' START: SAVE PRODUCTS ORDERED
'***********************************************************************************
for f=1 to ppcCartIndex 
 	if pcCartArray(f,10)=0 then     
    
		if pcCartArray(f,11)="" or isNull(pcCartArray(f,11)) then
			pcCartArray(f,11)="NULL"
		end if
	  
		if pcCartArray(f,12)="" or isNull(pcCartArray(f,12)) then
			pcCartArray(f,12)="NULL"
		end if
	  
		if pcCartArray(f,14)="" or isNull(pcCartArray(f,14)) then
			pcCartArray(f,14)=0
		end if
	 
		' replace , by .
		pcCartArray(f,14)=replace(pcCartArray(f,14),",",".")
		if pcCartArray(f,16)<>"" or pcCartArray(f,15)<>"0" then
			tempVar1=(pcCartArray(f,5) + pcCartArray(f,17))
		else
			tempVar1=(pcCartArray(f,5) + pcCartArray(f,3))
		end if
		
		If pcCartArray(f,16)="" then
			pcCartArray(f,16)=0
		end If
		
		if pcCartArray(f,15)<>"" then
			QDiscounts=pcCartArray(f,15)
		else
			QDiscounts="0"
		end if
		if pcCartArray(f,30)<>"" then
			ItemsDiscounts=pcCartArray(f,30)
		else
			ItemsDiscounts="0"
		end if
		
		'GGG Add-on start
		if pcCartArray(f,33)<>"" then
		geID=pcCartArray(f,33)
		else
		geID="0"
		end if
		
		if pcCartArray(f,34)<>"" then
		pGWOpt=pcCartArray(f,34)
		else
		pGWOpt="0"
		end if
		
		if pcCartArray(f,35)<>"" then
			pGWOptText=Server.HTMLEncode(pcCartArray(f,35))
			pGWOptText=replace(pGWOptText,"'","''")
			if len(pGWOptText)>240 then
				pGWOptText=left(pGWOptText,240)
			end if
		else
			pGWOptText=""
		end if
		
		if pGWOpt<>"0" then
			query="select pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
			set rsG=connTemp.execute(query)
			pGWPrice=rsG("pcGW_OptPrice")
			if pGWPrice<>"" then
			else
				pGWPrice="0"
			end if
		else
			pGWPrice="0"
		end if
		'GGG Add-on end
		
		pcv_xdetails=pcCartArray(f,21)
		if pcv_xdetails<>"" then
			pcv_xdetails=replace(pcv_xdetails,"<br>","|")
			pcv_xdetails=replace(pcv_xdetails,"'","''")
			pcv_xdetails=replace(pcv_xdetails,"''''","''")
		end if
		
		'// Start SDBA
		query="SELECT serviceSpec,stock,nostock,pcProd_BackOrder,pcDropShipper_ID FROM Products WHERE idproduct=" & pcCartArray(f,0)
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)		
		if not rs.eof then
			pcv_serviceSpec=rs("serviceSpec")
			if IsNull(pcv_pserviceSpec) or pcv_pserviceSpec="" then
				pcv_pserviceSpec="0"
			end if
			pcv_Stock=rs("stock")
			if IsNull(pcv_Stock) or pcv_Stock="" then
				pcv_Stock="0"
			end if
			pcv_NoStock=rs("nostock")
			if IsNull(pcv_NoStock) or pcv_NoStock="" then
				pcv_NoStock="0"
			end if
			pcv_intBackOrder=rs("pcProd_BackOrder")
			if IsNull(pcv_intBackOrder) or pcv_intBackOrder="" then
				pcv_intBackOrder="0"
			end if
			pcv_IDDropShipper=rs("pcDropShipper_ID")
			if IsNull(pcv_IDDropShipper) or pcv_IDDropShipper="" then
				pcv_IDDropShipper="0"
			end if
		else
			pcv_pserviceSpec="0"
			pcv_Stock="0"
			pcv_NoStock="0"
			pcv_intBackOrder="0"
			pcv_IDDropShipper="0"
		end if
		set rs=nothing				
		If (scOutofStockPurchase=-1 AND CLng(pcv_Stock)<1 AND pcv_serviceSpec=0 AND pcv_NoStock=0 AND pcv_intBackOrder=1) OR (pcv_serviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pcv_Stock)<1 AND pcv_NoStock=0 AND pcv_intBackOrder=1) Then
			tmp_BackOrder="1"
		Else
			tmp_BackOrder="0"
		End if
		'// End SDBA
		
		'Get SB Infor
		query="SELECT pcSubscription_ID,pcPO_SubActive FROM ProductsOrdered WHERE idProduct=" & pcCartArray(f,0) & " AND idOrder=" & SBIDOrder & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcvSubscriptionID=0
		if not rs.eof then
			pcvSubscriptionID=rs("pcSubscription_ID")
			pcvSubActive=rs("pcPO_SubActive")
		end if
		set rs=nothing
		
		'// Insert the Line Item
		query="INSERT INTO ProductsOrdered (pcSubscription_ID,pcPO_SubActive,idOrder, idProduct, quantity, unitPrice, unitCost, idconfigSession, xfdetails, QDiscounts,ItemsDiscounts, pcPackageInfo_ID, pcDropShipper_ID, pcPrdOrd_Shipped, pcPrdOrd_BackOrder, pcPrdOrd_SelectedOptions, pcPrdOrd_OptionsPriceArray, pcPrdOrd_OptionsArray, pcPO_EPID,pcPO_GWOpt, pcPO_GWNote, pcPO_GWPrice) VALUES (" & pcvSubscriptionID & "," & pcvSubActive & "," & pIdOrder & "," &pcCartArray(f,0)& "," &pcCartArray(f,2)& "," & replacecomma(tempVar1) & "," & replacecomma(pcCartArray(f,14))& "," &pcCartArray(f,16)& ",N'" &pcv_xdetails& "'," & QDiscounts & "," & ItemsDiscounts & ",0," & pcv_IDDropShipper & ",0," & tmp_BackOrder & ",'" & replace(pcCartArray(f,11),"'","''") & "','" & replace(pcCartArray(f,25),"'","''") & "','" & replace(pcCartArray(f,4),"'","''") &"'," & geID & "," & pGWOpt & ",N'" & pGWOptText & "'," & pGWPrice & ")"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)  
		set rs=nothing
	end if 
next 
'***********************************************************************************
' END: SAVE PRODUCTS ORDERED
'***********************************************************************************


'///////////////////////////////////////////////////////////////////////////////////
'//  END: ORDER PROCESSING
'///////////////////////////////////////////////////////////////////////////////////



'// SET THESE
pIdOrder=session("idOrderSaved")


'***********************************************************************************
' START: GET ORDER DETAILS
'***********************************************************************************
query="SELECT orders.idcustomer, orders.address, orders.City, orders.StateCode, orders.State, orders.zip, orders.CountryCode, orders.shippingAddress, orders.shippingCity, orders.shippingStateCode, orders.shippingState, orders.shippingZip,  orders.shippingCountryCode, orders.pcOrd_shippingPhone, orders.ShipmentDetails, orders.PaymentDetails, orders.discountDetails, orders.taxAmount, orders.total, orders.comments, orders.ShippingFullName, orders.address2, orders.ShippingCompany, orders.ShippingAddress2, orders.taxDetails, orders.orderstatus, orders.iRewardPoints, orders.iRewardValue, orders.iRewardPointsCustAccrued, orders.ordPackageNum, customers.phone, orders.ord_DeliveryDate, orders.ord_VAT, orders.pcOrd_DiscountsUsed, orders.pcOrd_Payer FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" & pIdOrder
set rsObjOrder=server.CreateObject("ADODB.RecordSet")
set rsObjOrder=conntemp.execute(query)

pidcustomer=rsObjOrder("idcustomer")
paddress=rsObjOrder("address")
pCity=rsObjOrder("City")
pStateCode=rsObjOrder("StateCode")
pState=rsObjOrder("State")
if isNULL(pStateCode) OR pStateCode="" then
	pStateCode=pState
end if
pzip=rsObjOrder("zip")
pCountryCode=rsObjOrder("CountryCode")
pshippingAddress=rsObjOrder("shippingAddress")
pshippingCity=rsObjOrder("shippingCity")
pshippingStateCode=rsObjOrder("shippingStateCode")
pshippingState=rsObjOrder("shippingState")
if isNULL(pshippingStateCode) OR pshippingStateCode="" then
	pshippingStateCode=pshippingState
end if
pshippingZip=rsObjOrder("shippingZip")
pshippingCountryCode=rsObjOrder("shippingCountryCode")
pshippingPhone=rsObjOrder("pcOrd_shippingPhone")
pShipmentDetails=rsObjOrder("ShipmentDetails")
pPaymentDetails=rsObjOrder("PaymentDetails")
pdiscountDetails=rsObjOrder("discountDetails")
ptaxAmount=rsObjOrder("taxAmount")
ptotal=rsObjOrder("total")
pcomments=rsObjOrder("comments")
pShippingFullName=rsObjOrder("ShippingFullName")
paddress2=rsObjOrder("address2")
pShippingCompany=rsObjOrder("ShippingCompany")
pShippingAddress2=rsObjOrder("ShippingAddress2")
ptaxDetails=rsObjOrder("taxDetails")
pCurOrderStatus=rsObjOrder("orderStatus")
piRewardPoints=rsObjOrder("iRewardPoints")
piRewardValue=rsObjOrder("iRewardValue")
piRewardPointsCustAccrued=rsObjOrder("iRewardPointsCustAccrued")
pOrdPackageNum=rsObjOrder("ordPackageNum")
pPhone=rsObjOrder("phone")
pord_DeliveryDate=rsObjOrder("ord_DeliveryDate")
pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)
pord_VAT=rsObjOrder("ord_VAT")
strPcOrd_DiscountsUsed=rsObjOrder("pcOrd_DiscountsUsed")
pcOrd_Payer=rsObjOrder("pcOrd_Payer")
set rsObjOrder=nothing
'***********************************************************************************
' END: GET ORDER DETAILS
'***********************************************************************************

ppStatus=0 '// This will allow the code below to execute.


'***********************************************************************************
' START: CUSTOMER ID
'***********************************************************************************
pName= pFirstName
pLName= pLastName 
pIdCustomer = pcv_intCustomerId
'***********************************************************************************
' END: CUSTOMER ID
'***********************************************************************************




'***********************************************************************************
' START: ITERATE THROUGH ORDER ITEMS
'***********************************************************************************
query="SELECT idProduct,quantity,idconfigSession FROM ProductsOrdered WHERE idOrder=" & pIdOrder
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)  
	
do while not rs.eof  
	pIdProduct=rs("idProduct")
	pQuantity=rs("quantity")
	idconfig=rs("idconfigSession")
	
	'// Check if stock is ignored, or not
	query="SELECT noStock FROM products WHERE idProduct=" & pIdProduct
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)   
	pNoStock=rstemp("noStock")

	query="SELECT stock, sales, description FROM products WHERE idProduct=" & pIdProduct
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)       
	pDescription=rstemp("description")
	
	if pNoStock=0 then
		'// Decrease stock 
		if ppStatus=0 then
			query="UPDATE products SET stock=stock-" & pQuantity &" WHERE idProduct=" & pIdProduct
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rsTemp=conntemp.execute(query)
			call pcs_hookStockChanged(pIdProduct, "")
			
			'// Update BTO Items & Additional Charges stock and sales 
			IF (idconfig<>"") and (idconfig<>"0") then
				query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
				set rs1=server.CreateObject("ADODB.RecordSet")
				set rs1=conntemp.execute(query)
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")
					
					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &QtyArr(k)*pQuantity&",sales=sales+" &QtyArr(k)*pQuantity&" WHERE idProduct=" &PrdArr(k)
							set rs1=conntemp.execute(query)
							call pcs_hookStockChanged(PrdArr(k), "")
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")
					
					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &pQuantity&",sales=sales+" &pQuantity&" WHERE idProduct=" &CPrdArr(k)
							set rs1=conntemp.execute(query)
							call pcs_hookStockChanged(CPrdArr(k), "")
						end if
					next
				end if
			END IF
			'// End Update BTO Items & Additional Charges stock and sales 
			
		end if
	end if
				 
	'// Update sales 
	if ppStatus=0 then  
		query="UPDATE products SET sales=sales+" &pQuantity&" WHERE idProduct=" &pIdProduct
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)  
		set rstemp=nothing 
	end if 
	
	
	rs.movenext	
loop
set rs=nothing
set rstemp=nothing
set rs1=nothing
'***********************************************************************************
' END: ITERATE THROUGH ORDER ITEMS
'***********************************************************************************




'***********************************************************************************
' START: REWARD POINTS
'***********************************************************************************
qry_ID=pIdOrder
If piRewardPoints > 0 Then
	if ppStatus=0 then
		'// Even if pending, if a customer uses pts, they must be held as substracted until order is canceled.
		query="SELECT iRewardPointsUsed, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		iRewardPointsUsed=rs("iRewardPointsUsed")
		If IsNull(iRewardPointsUsed) OR iRewardPointsUsed="" Then
			iRewardPointsUsed=0
		end if		
		query = "SELECT iRewardValue FROM orders WHERE idOrder=" & qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		piRewardValue=rs("iRewardValue")
		iNewUsed = iRewardPointsUsed + piRewardPoints		
		query = "UPDATE customers SET iRewardPointsUsed=" & iNewUsed & " WHERE idCustomer=" & pIdCustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
End If
'***********************************************************************************
' END: REWARD POINTS
'***********************************************************************************

'***********************************************************************************
' START: INSERT NEW ORDER ID to SB_Orders Table
'***********************************************************************************
query="INSERT INTO SB_Orders (SB_GUID,idOrder,SB_Terms) VALUES ('" & SBGuid & "'," & pIdOrder & ",N'" & SBTerms & "');"
set rs=connTemp.execute(query)
set rs=nothing
'***********************************************************************************
' END: INSERT NEW ORDER ID to SB_Orders Table
'***********************************************************************************


'///////////////////////////////////////////////////////////////////////////////////
'// START: EMAILS
'///////////////////////////////////////////////////////////////////////////////////
'Send as HTML Emails
session("News_MsgType")="1"
%>

<!--#include file="adminNewOrderEmail.asp"-->

<%
if ppStatus=0 then
	dim strNewOrderSubject
	strNewOrderSubject=dictLanguage.Item(Session("language")&"_storeEmail_9")&(Clng(scpre) + Clng(pIdOrder))
	if pcOrderKey<>"" then
		storeAdminEmail=storeAdminEmail & "<br>" & vbCrLf
		storeAdminEmail=storeAdminEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbCrLf
		storeAdminEmail=storeAdminEmail & "ORDER CODE: " & pcOrderKey & "<br>" & vbCrLf
		storeAdminEmail=storeAdminEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbCrLf
	storeAdminEmail=storeAdminEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_SB_37") & "<br>" & vbCrLf
	storeAdminEmail=storeAdminEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbCrLf
	storeAdminEmail = pcf_HtmlEmailWrapper(storeAdminEmail, pcv_HTMLEmailFontFamily)
	call sendmail (scCompanyName, scEmail, scFrmEmail, strNewOrderSubject, storeAdminEmail)
	call pcs_hookNewOrderEmailSent(scFrmEmail)
end if

'// START - Order is processed when placed -> Send order confirmation
IF pOrderStatus="3" THEN 
	'order processed
	if ppStatus=0 then
		'Variable to generate Customer Order Confirmation Email
		pcv_CustomerReceived=0 %>
		<!--#include file="customerOrderConfirmEmail.asp"-->
		<% 
		pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_2") & " - " & dictLanguage.Item(Session("language")&"_sendMail_1") & (Clng(scpre) + Clng(pIdOrder))
		customerEmail=customerEmail & "<br>" & vbCrLf
		customerEmail=customerEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbcrlf & "<br>" & vbcrlf
		customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_SB_37") & "<br>" & vbCrLf
		customerEmail=customerEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbcrlf & "<br>" & vbcrlf
		customerEmail = pcf_HtmlEmailWrapper(customerEmail, pcv_HTMLEmailFontFamily)
		call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, customerEmail)
		call pcs_hookOrderConfirmationEmailSent(pEmail)
		
		'Start SDBA%>
		<!--#include file="inc_DropShipperNotificationEmail.asp"-->
		<%'End SDBA
	end if
ELSE ' Order is pendng -> Send order received e-mail
	if ppStatus=0 then
		'Variable to generate Customer Order Received Email
		pcv_CustomerReceived=1%>
		<!--#include file="customerOrderConfirmEmail.asp"-->
		<%pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_1") & " - " & dictLanguage.Item(Session("language")&"_sendMail_1") & (Clng(scpre) + Clng(pIdOrder))
		customerEmail=customerEmail & "<br>" & vbCrLf
		customerEmail=customerEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbcrlf & "<br>" & vbcrlf
		customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_SB_37") & "<br>" & vbCrLf
		customerEmail=customerEmail & FixedField(100, "R", "----------------------------------------------------------------------------------------------------") & "<br>" & vbcrlf & "<br>" & vbcrlf
		customerEmail = pcf_HtmlEmailWrapper(customerEmail, pcv_HTMLEmailFontFamily)
		call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, customerEmail)
		call pcs_hookOrderReceivedEmailSent(pEmail)
	end if
END IF
'// END - Order is processed when placed -> Send order confirmation
session("News_MsgType")="0"
'***********************************************************************************
' START: SDBA - Send Low Inventory Notification
'***********************************************************************************
%>
<!--#include file="inc_StockNotificationEmail.asp"-->
<%
'***********************************************************************************
' END: SDBA - Send Low Inventory Notification
'***********************************************************************************


'///////////////////////////////////////////////////////////////////////////////////
'// END: EMAILS
'///////////////////////////////////////////////////////////////////////////////////


'***********************************************************************************
' START: CLEAR DATA
'***********************************************************************************
Session("pcCartIndex")=Cint(0)
session("iOrderTotal")=""
session("continueRef")=""
session("pcSFCartRewards")=Cint(0)
session("pcSFUseRewards")=Cint(0)
session("IDRefer")=""
session("specialdiscount")=""
session("EPN_idOrder")=""
session("pc_pidOrder")=""
session("GWAuthCode")=""
session("GWTransId")=""
session("GWPaymentId")=""
session("GWTransType")=""
session("GWOrderId")=""
session("GWSessionID")=""
session("GWOrderDone")=""
'GGG Add-on start
session("Cust_BuyGift")=""
session("Cust_IDEvent")=""
'GGG Add-on end
Session.Abandon()
'***********************************************************************************
' END: CLEAR DATA
'***********************************************************************************

'///////////////////////////////////////////////////////////////////////////////////
'//  END: ORDER STATUS AND PAYMENT STATUS
'///////////////////////////////////////////////////////////////////////////////////
%>