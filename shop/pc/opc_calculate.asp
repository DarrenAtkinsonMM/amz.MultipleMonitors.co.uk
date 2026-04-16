<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="inc_sb.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/GCConstants.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp" -->

<!--#include file="opc_contentType.asp" -->
<%
Call SetContentType()

Dim TurnOffDiscountCodesWhenHasSale
TurnOffDiscountCodesWhenHasSale = scDisableDiscountCodes

Dim HavePrdsOnSale
HavePrdsOnSale = 0

Dim pcvFromCart
pcvFromCart=getUserInput(request("fromcart"),0)
if pcvFromCart="" then
	if (session("idCustomer")="0") OR (session("idCustomer")="") then
		pcvFromCart="1"
	else
		pcvFromCart="0"
	end if
end if


%><!--#include file="ggg_inc_calGW.asp" --><% 
intGCIncludeShipping = GC_INCSHIPPING


Dim intTaxExemptZoneFlag
intTaxExemptZoneFlag = "1"   
'// Change to 0 if you want to tax any tax zone exempt products when they are added to the cart with taxable products. 
'// "1" will ensure that tax zone exempt products are never taxed for that zone.


dim pcCartArray, ppcCartIndex

'SB S
Dim pcIsSubscription , StrandSub 
pcIsSubscription = session("pcIsSubscription")

Dim pcv_sbTax
pcv_sbTax = getUserInput(request("sbTax"), 0)
'SB E

'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=session("pcCartIndex")




'//////////////////////////////////////////////////////////////
'// START - LOAD CUSTOMER SESSION DATA
'//////////////////////////////////////////////////////////////
If Not (session("pcSFIdDbSession") = "" OR session("pcSFRandomKey") = "") Then

    query="SELECT customers.name, customers.lastName, customers.customerCompany, customers.email, customers.phone, customers.fax, customers.address, customers.address2, customers.zip, customers.stateCode, customers.state, customers.city, customers.countryCode, customers.iRewardPointsAccrued, customers.iRewardPointsUsed, customers.pcCust_VATID, pcCustSession_ShippingFirstName, pcCustSession_ShippingLastName, pcCustSession_ShippingCompany, pcCustSession_ShippingAddress, pcCustSession_ShippingAddress2, pcCustSession_ShippingCity, pcCustSession_ShippingStateCode, pcCustSession_ShippingProvince, pcCustSession_ShippingPostalCode, pcCustSession_ShippingCountryCode, pcCustSession_ShippingPhone, pcCustSession_ShippingNickName, pcCustSession_TaxShippingAlone, pcCustSession_TaxShippingAndHandlingTogether, pcCustSession_TaxLocation, pcCustSession_TaxProductAmount, pcCustSession_OrdPackageNumber, pcCustSession_ShippingArray, pcCustSession_ShippingResidential, pcCustSession_IdPayment, pcCustSession_Comment, pcCustSession_discountcode, pcCustSession_UseRewards, pcCustSession_RewardsBalance,pcCustSession_NullShipper,pcCustSession_NullShipRates,pcCustSession_TF1,pcCustSession_DF1,pcCustSession_OrderName,pcCustSession_ShowShipAddr,pcCustSession_ShippingEmail,pcCustSession_ShippingFax,pcCustSession_GCDetails FROM pcCustomerSessions INNER JOIN customers ON pcCustomerSessions.idCustomer = customers.idcustomer WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&")) ORDER BY pcCustomerSessions.idDbSession DESC;"
    Set rs = server.CreateObject("ADODB.RecordSet")
    Set rs = connTemp.execute(query)
    If Not rs.Eof Then
        pcStrBillingFirstName=rs("name")
        pcStrBillingLastName=rs("lastName")
        pcStrBillingCompany=rs("customerCompany")
        pcStrBillingEmail=rs("email")
        pcStrBillingPhone=rs("phone")
        pcStrBillingfax=rs("fax")
        pcStrBillingAddress=rs("address")
        pcStrBillingAddress2=rs("address2")
        pcStrBillingPostalCode=rs("zip")
        pcStrBillingStateCode=rs("stateCode")
        pcStrBillingProvince=rs("state")
        pcStrBillingCity=rs("city")
        pcStrBillingCountryCode=rs("CountryCode")
		pcStrBillingVATID=rs("pcCust_VATID")
        pcIntRewardPointsAccrued = rs("iRewardPointsAccrued")
        pcIntRewardPointsUsed = rs("iRewardPointsUsed")
        pcStrShippingFirstName=rs("pcCustSession_ShippingFirstName")
        pcStrShippingLastName=rs("pcCustSession_ShippingLastName")
        pcStrShippingCompany=rs("pcCustSession_ShippingCompany")
        pcStrShippingAddress=rs("pcCustSession_ShippingAddress")
        pcStrShippingAddress2=rs("pcCustSession_ShippingAddress2")
        pcStrShippingCity=rs("pcCustSession_ShippingCity")
        pcStrShippingStateCode=rs("pcCustSession_ShippingStateCode")
        pcStrShippingProvince=rs("pcCustSession_ShippingProvince")
        pcStrShippingPostalCode=rs("pcCustSession_ShippingPostalCode")
        pcStrShippingCountryCode=rs("pcCustSession_ShippingCountryCode")
        pcStrShippingPhone=rs("pcCustSession_ShippingPhone")
        pcStrShippingNickName=rs("pcCustSession_ShippingNickName")
        TAX_SHIPPING_ALONE=rs("pcCustSession_TaxShippingAlone")
        TAX_SHIPPING_AND_HANDLING_TOGETHER=rs("pcCustSession_TaxShippingAndHandlingTogether")
        If Not ISNULL(rs("pcCustSession_TaxLocation")) Then
            ptaxLoc=Cdbl(rs("pcCustSession_TaxLocation"))
        End If
        If Not ISNULL(rs("pcCustSession_TaxProductAmount")) Then
            ptaxPrdAmount =ccur(rs("pcCustSession_TaxProductAmount"))
        End If
        pcIntOrdPackageNumber=rs("pcCustSession_OrdPackageNumber")
        pcShippingArray=rs("pcCustSession_ShippingArray")
        pOrdShipType=rs("pcCustSession_ShippingResidential")
        pcIdPayment=rs("pcCustSession_IdPayment")
        savOrderComments=rs("pcCustSession_Comment")
        savdiscountcode=rs("pcCustSession_discountcode")
        savUseRewards=rs("pcCustSession_UseRewards")
        savNullShipper=rs("pcCustSession_NullShipper")
        savNullShipRates=rs("pcCustSession_NullShipRates")
        savTF1=rs("pcCustSession_TF1")
        savDF1=rs("pcCustSession_DF1")
        savOrderNickName=rs("pcCustSession_OrderName")
        pcShowShipAddr=rs("pcCustSession_ShowShipAddr")
        pcStrShippingEmail=rs("pcCustSession_ShippingEmail")
        pcStrShippingFax=rs("pcCustSession_ShippingFax")
        savGCs=rs("pcCustSession_GCDetails")
        If savGCs<>"" Then
            GCArr = split(savGCs,"|g|")
            savGCs=""
            For y=0 To ubound(GCArr)
                If GCArr(y)<>"" Then
                    GCInfo=split(GCArr(y),"|s|")
                    If savGCs<>"" Then
                        savGCs=savGCs & ","
                    End If
                    savGCs=savGCs & GCInfo(0)
                End If
            Next
            If savdiscountcode<>"" Then
                If Right(savdiscountcode,1) <> "," Then
                    savdiscountcode = savdiscountcode & ","
                End If
            End If
            savdiscountcode = savdiscountcode & savGCs
        end if	
    Else
        '// Stop when cart is not saved and not "viewcart.asp"
        If pcvFromCart<>"1" Then
            response.End() 			
        Else
            savdiscountcode=session("pcSFCust_discountcode")
        End If
    End If
    Set rs = Nothing
    
Else
	savdiscountcode=session("pcSFCust_discountcode")
End If
'//////////////////////////////////////////////////////////////
'// END - LOAD CUSTOMER SESSION DATA
'//////////////////////////////////////////////////////////////





' NOTE:  This gets tax zone data and will likely be moved into the core cart methods
%> <!--#include file="pcTaxZone.asp"--> <%


'// Check if the Customer is European Union 
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pcStrShippingCountryCode)




'//////////////////////////////////////////////////////////////
'// START - DISCOUNTS
'//////////////////////////////////////////////////////////////

'// Set AutoDiscount flag to 0
pcIntADCnt=0 
    
'// Check for discounts
Dim displayDiscountCode, tmpPrdList
displayDiscountCode = getDiscountCodes(request("rtype"), savdiscountcode, request("discountcode"), TurnOffDiscountCodesWhenHasSale, scHideDiscField, pcCartArray, ppcCartIndex)

'// If this is first visit, check for Auto discounts
displayDiscountCode = getAutoDiscountCodes(HavePrdsOnSale, displayDiscountCode)

pEnterDC=""

if displayDiscountCode<>"" then
    pDiscountCode=URLDecode(getUserInput(displayDiscountCode,0))
	pEnterDC=URLDecode(getUserInput(request("discountcode"),0))
end if

Select Case Session("daActualDiscount")
	Case 30
		pDiscountCode="Bun25,"&pDiscountCode
		pDiscountCode=Replace(lcase(pDiscountCode),"bun50","")
		pDiscountCode=Replace(lcase(pDiscountCode),"bun100","")
	Case 60
		pDiscountCode="Bun50,"&pDiscountCode
		pDiscountCode=Replace(lcase(pDiscountCode),"bun25","")
		pDiscountCode=Replace(lcase(pDiscountCode),"bun100","")
	Case 120
		pDiscountCode="Bun100,"&pDiscountCode
		pDiscountCode=Replace(lcase(pDiscountCode),"bun25","")
		pDiscountCode=Replace(lcase(pDiscountCode),"bun50","")
	Case Else
		pDiscountCode=pDiscountCode
		pDiscountCode=Replace(lcase(pDiscountCode),"bun25","")
		pDiscountCode=Replace(lcase(pDiscountCode),"bun50","")
		pDiscountCode=Replace(lcase(pDiscountCode),"bun100","")
		pDiscountCode=Replace(lcase(pDiscountCode),"bundle50","")
End Select

'Code to stop multiple discount code abuse
if InStr(lcase(pDiscountCode),"bundle50") > 0 then
		pDiscountCode=Replace(lcase(pDiscountCode),"stand15","")
		pDiscountCode=Replace(lcase(pDiscountCode),"pc25","")
end if

'//////////////////////////////////////////////////////////////
'// END - DISCOUNTS
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - CALCULATE TOTAL, WEIGHT, QUANTITIES
'//////////////////////////////////////////////////////////////

Dim pTaxableTotal, pCartTotalWeight, pEryPassword

'SB S
If len(pcv_sbTax)>0 Then
	pTaxableTotal = ccur(calculateTaxableTotal_SB(pcCartArray, ppcCartIndex))
Else
    pTaxableTotal = ccur(calculateTaxableTotal(pcCartArray, ppcCartIndex))
End If
'SB E

pSubTotal = ccur(calculateCartTotal(pcCartArray, ppcCartIndex))
pCartTotalWeight = Int(calculateCartWeight(pcCartArray, ppcCartIndex))
pCartQuantity = Int(calculateCartQuantity(pcCartArray, ppcCartIndex))
pSFSubTotal = pSubTotal '// This is saved to database. Do not move!

'//////////////////////////////////////////////////////////////
'// END - CALCULATE TOTAL, WEIGHT, QUANTITIES
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - SHIPMENT DATA
'//////////////////////////////////////////////////////////////

Dim pcDblServiceHandlingFee, TempStrNewShipping, pcIntIdShipService, tmpErrorOPCReady '// TempStrNewShipping = Passthrough?

If len(pcShippingArray)>0 Then
    pcShipmentPriceToAdd = calculateShippingPrice(savNullShipper, savNullShipRates, pcShippingArray)
    pcDblServiceHandlingFee = calculateServiceHandlingFee(savNullShipper, savNullShipRates, pcShippingArray)
Else
    pcShipmentPriceToAdd = 0
    pcDblServiceHandlingFee = 0
End If

If pcShipmentPriceToAdd > 0 Then 
    pcDblShipmentTotal = pcShipmentPriceToAdd     
Else
    pcDblShipmentTotal = 0
End If

'//////////////////////////////////////////////////////////////
'// END - SHIPMENT DATA
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - PAYMENT DATA
'//////////////////////////////////////////////////////////////
Dim paymentTotal, taxPaymentTotal
If pcvFromCart<>"1" Then
	'// Get and Set the Payment Gateway Id
	pidPayment = getSetPaymentId(getUserInput(request("idpayment"),0), pcIdPayment)
	If len(pidPayment)>0 Then
    	Session("DefaultIdPayment") = pidPayment
	End If 

	'// Calculate Payment Gateway Fees
	paymentTotal = calculatePaymentGatewayFees(pidPayment, pcIsSubscription)
Else
	paymentTotal=0
End if

'// Add Payment Gateway Fees to SubTotal
pSubTotal = pSubTotal + paymentTotal

'//////////////////////////////////////////////////////////////
'// END - PAYMENT DATA
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - REWARD POINTS
'//////////////////////////////////////////////////////////////

'// Note: v5 - Reward points moved below promotions and cat discounts.

'//////////////////////////////////////////////////////////////
'// END - REWARD POINTS
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - CART LOOP STUFF
'//////////////////////////////////////////////////////////////

Dim GiftWrapPaymentTotal, strBundleArray, pSFstrBundleArray

Dim pcProductList(100, 5)

'// The following routine calculates f 40, which is row total
call ReCalculateCartRows(ppcCartIndex, pcCartArray)

pSFstrBundleArray = strBundleArray 
                
'//////////////////////////////////////////////////////////////
'// END - CART LOOP STUFF
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - PRODUCT PROMOTIONS
'//////////////////////////////////////////////////////////////

'// Get Total Promotions
TotalPromotions = calculateTotalPromotions(Session("pcPromoIndex"), Session("pcPromoSession"))

'// Adjust SubTotal with Total Promotions
pSubTotal = pSubTotal - TotalPromotions
       
'//////////////////////////////////////////////////////////////
'// END - PRODUCT PROMOTIONS
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - DISCOUNTS BY CATEGORIES
'//////////////////////////////////////////////////////////////

Dim CatDiscTotal
CatDiscTotal = calculateCategoryDiscountTotal(ppcCartIndex, pcCartArray)

'// Adjust SubTotal with Category Discounts Total
pSubTotal = pSubTotal - CatDiscTotal 
pSFCatDiscTotal = CatDiscTotal
                
'//////////////////////////////////////////////////////////////
'// END - DISCOUNTS BY CATEGORIES
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - REWARD POINTS
'//////////////////////////////////////////////////////////////

pcIntBalance = calculateRewardsBalance(RewardsActive, pcIntRewardPointsAccrued, pcIntRewardPointsUsed, RewardsPercent)

pcSFUseRewards = IsUsingRewards(request("rtype"), request("UseRewards"), savUseRewards, pcIntBalance, RewardsIncludeWholesale)

pcSFCartRewards = calculateSFCartRewards(pcSFUseRewards, pcCartArray, ppcCartIndex, RewardsIncludeWholesale)

iDollarValue = calculateRewardsTotal(RewardsActive, pcSFUseRewards, RewardsPercent, pSubTotal)

'// Adjust SubTotal with Rewards
if pSubTotal<>0 then
    'pSubTotal = pSubTotal - iDollarValue
else
    pSubTotal=0
end if

'// Adjust Taxable Total with Rewards
pTaxableTotal=pTaxableTotal-iDollarValue
If pTaxableTotal<0 Then
    pTaxableTotal=0
End If
If session("customerType")=1 And ptaxwholesale=0 Then
    pTaxableTotal=0
End If

'// Reward Points are being used against the purchase
RewardsDollarValue=0
If RewardsActive=1 And pcSFUseRewards<>"" Then 
    RewardsDollarValue = iDollarValue							
End If 

session("SF_RewardPointTotal") = RewardsDollarValue

'//////////////////////////////////////////////////////////////
'// END - REWARD POINTS
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
' START - SPLIT GIFT CERTS AND DISCOUNTS
'//////////////////////////////////////////////////////////////

pGCError=Cstr("")
pDiscountShowCode=Cstr("")
discountTotal=ccur(0)
passDiscountCnt=-1 '// (this is very important and gets saved)
noCode=""
intCodeCnt=-1 '// (this is very important and gets saved)
'intGCCnt=-1 '// (not being used)

Dim pTempGC, pDiscountCode, DiscountCodeArry
call separateDiscountsAndGiftCodes()

'//////////////////////////////////////////////////////////////
' END - SPLIT GIFT CERTS AND DISCOUNTS
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - DISCOUNTS BY CODE
'//////////////////////////////////////////////////////////////

Dim pDiscountError '// Passthrough
pDiscountError=Cstr("")


Dim intArryCnt
intArryCnt=0
					
'// Set filter variables 
CatCount=1
CatFound=0

'pDiscountCode = "Bun50,"&pDiscountCode
				
If pDiscountCode <> "" Then

    DiscountTableRow=""
					
    '// This array changes below
    DiscountCodeArry=Split(pDiscountCode,",")
    intCodeCnt=ubound(DiscountCodeArry)
    
    '// This is the same as above now, but the one above changes during sorting and this marks the original
    DiscountCodeArryO=Split(pDiscountCode,",")
    intCodeCntO=ubound(DiscountCodeArry)

    '// Sort Discount Codes for Validation
    Dim intCodeCnt                    
    call sortDiscountCodes()

    '// Check Invalid Discount Codes
    pcGlobalDiscError=Cstr("")
    pcGlobalDiscError = checkInvalidCodes(intCodeCntO, DiscountCodeArryO, intCodeCnt)


	pcv_HaveSeparateCode = 0

    For i=0 To intCodeCnt
    
        pcv_Filters=0
        pcv_FResults=0
        pcv_ProTotal=0
        
        If trim(DiscountCodeArry(i)) <> "" Then

            pTempDiscCode = DiscountCodeArry(i)
            Session("DiscountTotal" & pTempDiscCode) = 0
            Session("DiscountRow" & pTempDiscCode) = ""
            
            
            '// Check if discount code has already been used for this store
            Dim UsedDiscountCodes
            UsedDiscountCodes = getUsedDiscountCodes(UsedDiscountCodes, pTempDiscCode)


                            
							query="SELECT iddiscount, onetime,expDate, idProduct, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, DiscountDesc, priceToDiscount, percentageToDiscount, pcSeparate, pcDisc_Auto, pcDisc_StartDate, pcRetailFlag, pcWholesaleFlag, pcDisc_PerToFlatCartTotal, pcDisc_PerToFlatDiscount,pcDisc_IncExcPrd,pcDisc_IncExcCat,pcDisc_IncExcCust,pcDisc_IncExcCPrice FROM discounts WHERE discountcode='" &pTempDiscCode& "' AND active=-1;"
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)
	
							if rs.eof then
								pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4")
								pDiscountDesc=""
							else
								pcv_IDDiscount=rs("iddiscount")
								pcv_IDDiscount1=rs("iddiscount")
								pcv_OneTime=rs("onetime")
								expDate=rs("expDate")
								dcIdProduct=rs("idProduct")
								dcQuantityFrom=rs("quantityFrom")
								dcQuantityUntil=rs("quantityUntil")
								dcWeightFrom=rs("weightFrom")
								dcWeightUntil=rs("weightUntil")
								dcPriceFrom=rs("priceFrom")
								dcPriceUntil=rs("priceUntil")
								pDiscountDesc=rs("DiscountDesc")
								session("DiscountDesc" & pTempDiscCode)= pDiscountDesc
								pPriceToDiscount=ccur(rs("priceToDiscount"))
								ppercentageToDiscount=rs("percentageToDiscount")
								intPcSeparate=rs("pcSeparate")
								intPcAuto=rs("pcDisc_Auto")
								pcv_startDate=rs("pcDisc_StartDate")
								pcv_retail = rs("pcRetailFlag")
								pcv_wholeSale = rs("pcWholeSaleFlag")
								pcv_PerToFlatCartTotal = rs("pcDisc_PerToFlatCartTotal")
								pcv_PerToFlatDiscount = rs("pcDisc_PerToFlatDiscount")
								pcIncExcPrd=rs("pcDisc_IncExcPrd")
								pcIncExcCat=rs("pcDisc_IncExcCat")
								pcIncExcCust=rs("pcDisc_IncExcCust")
								pcIncExcCPrice=rs("pcDisc_IncExcCPrice")
                                If ((clng(intPcSeparate)=0) And (passDiscountCode<>"")) Then
									pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_39")
								End If

                                If ((clng(pcv_HaveSeparateCode)=1) And (passDiscountCode<>"")) Then
									pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_39")
								End If
								
								'// Discount has been used for one use only for this customer specified?
								If pcv_OneTime<>0 Then
									
									'check used discounts in database with iddiscount
									query="SELECT A.idcustomer FROM used_discounts A,customers B WHERE A.iddiscount=" & pcv_IDDiscount1 & " AND A.idcustomer=B.idcustomer AND B.email IN (SELECT C.email FROM customers C WHERE C.idcustomer="&session("IDCustomer")&");"
									set rsCheck=server.CreateObject("ADODB.RecordSet")
									set rsCheck=connTemp.execute(query)									
									if err.number<>0 then
										call LogErrorToDatabase()
									end if
									
									varOneTimePresent=0
									if NOT rsCheck.eof then
										'discount has been used already by the customer
										pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
										varOneTimePresent=1
									end if
									set rsCheck=nothing
									
									If expDate<>"" then
										If datediff("d", Now(), expDate) <= 0 Then
											if varOneTimePresent=0 then
												pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
											end if
										end if
									end if
									
									'check to see if discount has start date
									If pcv_startDate<>"" then
										StartDate=pcv_startDate
										If datediff("d", Now(), StartDate) > 0 Then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_43")
										End If
									end if
								Else
									'check to see if discount code has expired
									If expDate<>"" then
										If datediff("d", Now(), expDate) <= 0 Then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
										end if
									end if
									
									'check to see if discount has start date
									If pcv_startDate<>"" then
										StartDate=pcv_startDate
										If datediff("d", Now(), StartDate) > 0 Then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_43")
										End If
									end if
								end if

								Dim pcv_dblSubTotalAdjusted
								pcv_dblSubTotalAdjusted = pSubTotal - paymentTotal - discountTotal
								If pcv_dblSubTotalAdjusted<0 Then
									pcv_dblSubTotalAdjusted=0
								End If
								
								If pDiscountError="" Then
									if Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil) and Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil) and ccur(pcv_dblSubTotalAdjusted)>=ccur(dcPriceFrom) and ccur(pcv_dblSubTotalAdjusted)<=ccur(dcPriceUntil) then

									else
									
										if NOT (Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5b")
										elseif NOT (Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5a")
										elseif NOT (ccur(pcv_dblSubTotalAdjusted)>=ccur(dcPriceFrom) and ccur(pcv_dblSubTotalAdjusted)<=ccur(dcPriceUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5c")
										else
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
										end if
										
									end if
								End If
							
							end if
							set rs=nothing


							IF pcv_IDDiscount<>"" AND pDiscountError="" THEN
								
								'// START: Filter by Products
								pcv_ProductFilter = 0

								query="select pcFPro_IDProduct from PcDFProds where pcFPro_IDDiscount=" & pcv_IDDiscount1
								set rs=server.CreateObject("ADODB.RecordSet")	
								set rs=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
								end if
								if not rs.eof then									
									pcv_Filters=pcv_Filters+1
									tmpIDArr=rs.getRows()									
									intIDCount=ubound(tmpIDArr,2)
									pcv_ProductFilter = 1
								end if
								set rs=nothing

								If pcv_ProductFilter=1 Then
								
									for f=1 to ppcCartIndex
										if pcProductList(f,1)=0 then
											tmpgotit=0

											for ik=0 to intIDCount
												if clng(pcf_ProductIdFromArray(pcCartArray, f))=clng(tmpIDArr(0,ik)) then
													tmpgotit = 1
													exit for
												end if
											next
                                            
											if (pcIncExcPrd="0") AND (tmpgotit=1) then
												pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
												pcv_FResults=1
											else
												if (pcIncExcPrd="1") AND (tmpgotit=0) then
													pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
													pcv_FResults=1
												end if
											end if
										end if
									next '// for f=1 to ppcCartIndex
									if NOT (ccur(pcv_ProTotal)>=ccur(dcPriceFrom) and ccur(pcv_ProTotal)<=ccur(dcPriceUntil)) then
										pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5c")
									end if
									
								End If '// If intIDCount>0 Then								
								'// END: Filter by Products


								'// START: Filter by Categories
								If pcv_Filters=0 Then
									
									pcv_CatFilter = 0

									query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=connTemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
									end if						
									if not rs.eof then
										pcv_CatFilter = 1
									end if 
									set rs=nothing 
                                    
                                 
									If pcv_CatFilter=1 Then
										
										pcv_Filters=pcv_Filters+1
										
										for f=1 to ppcCartIndex
											
											if pcProductList(f,1)=0 then

												query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcf_ProductIdFromArray(pcCartArray, f)
												set rs2=server.CreateObject("ADODB.RecordSet")
												set rs2=connTemp.execute(query)
												if err.number<>0 then
													call LogErrorToDatabase()
												end if
												intCatCount=-1
												if not rs2.eof then                                                	
													tmpCatArr=rs2.getRows()
                                                    intCatCount=ubound(tmpCatArr,2)
                                                    tmpgotit=0													
												end if
                                                set rs2=nothing
                                                
                                             
												If intCatCount>=0 Then
												
													'Check assigned categories
                                                    For ik=0 to intCatCount
													
														pcv_IDCat=tmpCatArr(o,ik)

														query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1 & " and pcFCat_IDCategory=" & pcv_IDCat
														set rstemp=server.CreateObject("ADODB.RecordSet")
														set rstemp=connTemp.execute(query)
														if err.number<>0 then
															call LogErrorToDatabase()
														end if
														if not rstemp.eof then															
                                                        	set rstemp=nothing
															tmpgotit=1
                                                            exit for
														end if
                                                        set rstemp=nothing
                                                        
                                                     
                                                        'Check parent-categories
                                                        if (tmpgotit=0) AND (pcv_IDCat<>"1") then
                                                        	pcv_ParentIDCat=pcv_IDCat
															do while (tmpgotit=0) and (pcv_ParentIDCat<>"1")
 
																query="select idParentCategory from categories where idcategory=" & pcv_ParentIDCat
																set rstemp=server.CreateObject("ADODB.RecordSet")
																set rstemp=connTemp.execute(query)
																if err.number<>0 then
																	call LogErrorToDatabase()
																end if														
																if not rstemp.eof then																	
																	pcv_ParentIDCat=rstemp("idParentCategory")
																	if pcv_ParentIDCat<>"1" then
																		
																		query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1 & " and pcFCat_IDCategory=" & pcv_ParentIDCat & " and pcFCat_SubCats=1;"
																		set rsFCat=server.CreateObject("ADODB.RecordSet")
																		set rsFCat=connTemp.execute(query)
																		if err.number<>0 then
																			call LogErrorToDatabase()
																		end if
																		if not rsFCat.eof then
																			tmpgotit=1
																		end if
																		set rsFCat=nothing
																		
																	end if
																end if
                                                                set rstemp=nothing
																
                                                             
															loop '// do while (tmpgotit=0) and (pcv_ParentIDCat<>"1")
                                                        end if

                                                        if tmpgotit=1 then
                                                            exit for
														end if
														
													Next '//  For ik=0 to intCatCount
													
													if (pcIncExcCat="0") AND (tmpgotit=1) then
															pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
														pcv_FResults=1
													else
														if (pcIncExcCat="1") AND (tmpgotit=0) then
																pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
															pcv_FResults=1
														end if
													end if

												End If '// If intCatCount>0 Then
											end if '// if pcProductList(f,1)=0 then  (Not deleted product)											
										next '// for f=1 to ppcCartIndex
									End If '// If pcv_CatFilter=1 Then
								End If '// If pcv_Filters=0 Then
								
								'// END: Filter by Categories

								tmpDiscErr=""
								
								'// START: Filter by Customers
								pcv_CustFilter=0

								query="select pcFCust_IDCustomer from PcDFCusts where pcFCust_IDDiscount=" & pcv_IDDiscount1
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
								end if								
								if not rs.eof then
									pcv_Filters=pcv_Filters+1
									pcv_CustFilter=1
								end if
								set rs=nothing

								if pcv_CustFilter=1 then
								
									if session("IDCustomer")="0" OR session("IDCustomer")="" then
										tmpDiscErr=dictLanguage.Item(Session("language")&"_DiscCart_3")
									else
										query="select pcFCust_IDCustomer from PcDFCusts where pcFCust_IDDiscount=" & pcv_IDDiscount1 & " and pcFCust_IDCustomer=" & session("IDCustomer")
										set rs=server.CreateObject("ADODB.RecordSet")
										set rs=connTemp.execute(query)
										if err.number<>0 then
											call LogErrorToDatabase()
										end if							
										if not rs.eof then
											if (pcIncExcCust="0") then
												pcv_FResults=pcv_FResults+1
											end if
										else
											if (pcIncExcCust="1") then
												pcv_FResults=pcv_FResults+1
											end if
										end if
										set rs=nothing
									end if
								end if
								'// END: Filter by Customers

								'// START: Customer Categories
								pcv_CustCatFilter=0

								query="select pcFCPCat_IDCategory from pcDFCustPriceCats where pcFCPCat_IDDiscount=" & pcv_IDDiscount1
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
								end if								
								if not rs.eof then
									pcv_Filters=pcv_Filters+1
									pcv_CustCatFilter=1
								end if
								set rs=nothing
								
								if pcv_CustCatFilter=1 then
									if session("IDCustomer")="0" OR session("IDCustomer")="" then
										tmpDiscErr=dictLanguage.Item(Session("language")&"_DiscCart_3")
									else
										query="select pcDFCustPriceCats.pcFCPCat_IDCategory from pcDFCustPriceCats, Customers where pcDFCustPriceCats.pcFCPCat_IDDiscount=" & pcv_IDDiscount1 & " and pcDFCustPriceCats.pcFCPCat_IDCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
										set rs=server.CreateObject("ADODB.RecordSet")
										set rs=connTemp.execute(query)
										if err.number<>0 then
											call LogErrorToDatabase()
										end if							
										if not rs.eof then
											if (pcIncExcCPrice="0") then
												pcv_FResults=pcv_FResults+1
											end if
										else
											if (pcIncExcCPrice="1") then
												pcv_FResults=pcv_FResults+1
											end if
										end if
										set rs=nothing
									end if
								end if
								'// END: Filter by Customer Categories


								'// START: Filter by reatil or wholesale
		                        if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
							    	pcv_Filters=pcv_Filters+1
									if pcv_wholeSale="1" AND (session("IDCustomer")="0" OR session("IDCustomer")="") then
										tmpDiscErr=dictLanguage.Item(Session("language")&"_DiscCart_3")
									else
										if pcv_wholeSale = "1" and session("customertype") = 1 then
											pcv_FResults=pcv_FResults+1	
										end if 
										if pcv_retail = "1" and session("customertype") <> 1 Then
											pcv_FResults=pcv_FResults+1
										end if
									end if
							    end if 
								'// END: Filter by reatil or wholesale

								if tmpDiscErr<>"" then
									pDiscountError=tmpDiscErr
								else
									if pcv_Filters<>pcv_FResults then
										pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_38")
									end if
								end if
							END IF
		

							'// If Not Errors from Filters
							if pDiscountError="" then
                            
                                '// Set this flag if coupon is not used with others.
                                If Not (intPcSeparate="" Or IsNull(intPcSeparate)) Then
                                    If clng(intPcSeparate)=0 Then
                                        pcv_HaveSeparateCode=1
                                    End If
                                End If
                                 
								pTempPriceToDiscount=pPriceToDiscount
								pTempPercentageToDiscount=pPercentageToDiscount
								pTempIdDiscount=pcv_IDDiscount

								' calculate discount. Note: percentage does not affect shipment and payment prices
								if pTempPriceToDiscount>0 or pTempPercentageToDiscount>0 then
									if pcv_ProTotal=0 then
										pcv_ProTotal=pSubTotal-paymentTotal 
									else
										pcv_ProTotal=pcv_ProTotal-CatDiscTotal
									end if
									if pcv_PerToFlatCartTotal<>0 AND pcv_ProTotal>pcv_PerToFlatCartTotal then
										tempPercentageToDiscount=pcv_PerToFlatDiscount
									else
										tempPercentageToDiscount=(pTempPercentageToDiscount*(pcv_ProTotal)/100)
										tempPercentageToDiscount=RoundTo(tempPercentageToDiscount,.01)
									end if
									pcv_ProTotal=0
									tempDiscountAmount=pTempPriceToDiscount + tempPercentageToDiscount
									discountTotal=discountTotal + tempDiscountAmount
									Session("DiscountTotal"&pTempDiscCode)=tempDiscountAmount
									pCheckSubtotal=pSubtotal-discountTotal
									if pCheckSubTotal<0 then
										tempDiscountAmount=tempDiscountAmount+pChecksubTotal
									end if
									if discountTotal<=0 then
										pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
									else
										if intArryCnt=0 then
											discountAmount=tempDiscountAmount
											passDiscountCode=pTempDiscCode
											passDiscountCnt=passDiscountCnt+1
											intArryCnt=intArryCnt+1
										else
											discountAmount=discountAmount&","&tempDiscountAmount
											passDiscountCode=passDiscountCode&","&pTempDiscCode
											passDiscountCnt=passDiscountCnt+1
											intArryCnt=intArryCnt+1
										end if
										pSFDiscountCodeTotal = discountTotal
									end if
										
								else '// else is "Free Shipping Coupon"
									if pcv_ProTotal=0 then
										pcv_ProTotal=pSubTotal-paymentTotal
									else
										if pcv_FResults=1 then
											pcv_ProTotal=pcv_ProTotal 
										else
											pcv_ProTotal=pcv_ProTotal-CatDiscTotal						
										end if
									end if

									if Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil) and Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil) and ccur(pcv_ProTotal)>=ccur(dcPriceFrom) and ccur(pcv_ProTotal)<=ccur(dcPriceUntil) then

										if pcIntIdShipService<>"" then

											query="select pcFShip_IDShipOpt from pcDFShip where pcFShip_IDDiscount=" & pTempIdDiscount & " and pcFShip_IDShipOpt=" & pcIntIdShipService
											set rs=server.CreateObject("ADODB.RecordSet")
											set rs=connTemp.execute(query)
											if err.number<>0 then
												call LogErrorToDatabase()
											end if								
											if not rs.eof then
												if intArryCnt=0 then
													discountAmount=ccur(pcDblShipmentTotal)
													passDiscountCode=pTempDiscCode
													passDiscountCnt=passDiscountCnt+1
													intArryCnt=intArryCnt+1
												else
													discountAmount=discountAmount&","&ccur(pcDblShipmentTotal)
													passDiscountCode=passDiscountCode&","&pTempDiscCode
													passDiscountCnt=passDiscountCnt+1
													intArryCnt=intArryCnt+1
												end if
												Session("DiscountTotal"&pTempDiscCode)=discountAmount
												pcDblShipmentTotal=0
												pcv_FREESHIP="ok"
											else
												pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_36")
											end if
											set rs=nothing

										else
											if (session("IDCustomer")="0") OR (session("IDCustomer")="") then
												pDiscountError=dictLanguage.Item(Session("language")&"_DiscCart_4")
											else
												pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_36")
											end if
										end if
									
									else
									
										if NOT (Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5b")
										elseif NOT (Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5a")
										elseif NOT (ccur(pcv_ProTotal)>=ccur(dcPriceFrom) and ccur(pcv_ProTotal)<=ccur(dcPriceUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5c")
										else
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
										end if
										
									end if
									
								end if
							end if
						
							if ((pDiscountDesc <> "") OR (pDiscountError<>"")) AND (noCode<>"1") Then

								if (pcIntADCnt>0) AND ((intPcAuto=1) AND ((InStr(pEnterDC,pTempDiscCode)=0) OR (pEnterDC="") OR (request("rtype")<>"1"))) AND (pDiscountError<>"") then
									pDiscountDesc=""
								else
									TableRowStr=""
									if pDiscountError<>"" then
										pcGlobalDiscError=pcGlobalDiscError & "<li>" & pDiscountError & " (<b>"&pTempDiscCode&"</b>)</li>"
									else
										TableRowStr="<tr><td colspan=""3""><p><b>"
										TableRowStr=TableRowStr&dictLanguage.Item(Session("language")&"_orderverify_14")&"</b>"
										TableRowStr=TableRowStr&"&nbsp;"&pDiscountDesc
										TableRowStr=TableRowStr&"</p></td><td nowrap align=""right""><p>"
									end if
									
									if pDiscountDesc <> "" then
										if tempDiscountAmount>0 then
											TableRowStr=TableRowStr&"-"&scCurSign & money(tempDiscountAmount)
										end if
									else
										if TableRowStr<>"" then
											TableRowStr=TableRowStr&"&nbsp;"
										end if
									end If
									if TableRowStr<>"" then
										TableRowStr=TableRowStr&"</p></td></tr>"
									end if
									if NOT pDiscountError<>"" then
										pDiscountShowCode=pDiscountShowCode&pTempDiscCode&","
									end if
									DiscountTableRow=DiscountTableRow&TableRowStr
									Session("DiscountRow"&pTempDiscCode)=TableRowStr
								end if
							end if
							pDiscountError=""
							tempDiscountAmount=0
							
							END IF 'DiscountCodeArry(i)<>""
						
						Next

				END IF
				
				'// Start: Double check the discounts are still valid after all discounts have been applied
				if pDiscountError="" then
					dim AdjustedSubTotal
					AdjustedSubTotal=pSubTotal 
					AdjustedSubTotal=AdjustedSubTotal - discountTotal
					if AdjustedSubTotal<0 then
						AdjustedSubTotal=0
						discountTotal=pSubTotal						
					end if
				end if

				IF (pDiscountCode<>"" AND pcGlobalDiscError="") AND (len(passDiscountCode)>0) THEN
					tmpDiscountCodeArry = split(passDiscountCode,",")
					pcvCodeCnt=ubound(tmpDiscountCodeArry)
					If pcvCodeCnt > 0 Then 
						For i=0 to pcvCodeCnt			
							IF trim(tmpDiscountCodeArry(i))<>"" THEN							
								pTempDiscCode=tmpDiscountCodeArry(i)								

								query="SELECT priceFrom, priceUntil, priceToDiscount, PercentageToDiscount, pcDisc_Auto FROM discounts WHERE discountcode='" & pTempDiscCode & "' AND active=-1;"
								set rs2=server.CreateObject("ADODB.RecordSet")
								set rs2=connTemp.execute(query)	
								if NOT rs2.eof then
									
									dcPriceFrom=rs2("priceFrom")
									dcPriceUntil=rs2("priceUntil")
									pPriceToDiscount=ccur(rs2("priceToDiscount"))
									pPercentageToDiscount=ccur(rs2("PercentageToDiscount"))	
									tmpPcAuto=rs2("pcDisc_Auto")									
									tmpPcAuto=clng(tmpPcAuto)
									
									'// Only double check the discount code if it free shipping								
									if NOT ( (ccur(AdjustedSubTotal)>=ccur(dcPriceFrom)) AND (ccur(AdjustedSubTotal)<=ccur(dcPriceUntil)) ) then	
						
										if NOT (pPriceToDiscount>0 or pPercentageToDiscount>0) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5c")

											if pcIntADCnt>0 AND tmpPcAuto=1 and pDiscountError<>"" then
												pDiscountDesc=""
                                                pPriceToDiscount=0
                                                pPercentageToDiscount=0
											else
												pcGlobalDiscError=pcGlobalDiscError & "<li>" & pDiscountError & " (<b>"&pTempDiscCode&"</b>)</li>"										
												pcv_FREESHIP="" '// disable free shipping
												pcDblShipmentTotal=Session("DiscountTotal"&pTempDiscCode) '// add shipping total back									
												'// remove the invalid discount
												discountAmount=replace(discountAmount, ","&tempDiscountAmount, "")	
												passDiscountCode=replace(passDiscountCode, ","&pTempDiscCode, "")	
												discountAmount=replace(discountAmount, tempDiscountAmount, "")	
												passDiscountCode=replace(passDiscountCode, pTempDiscCode, "")									
												DiscountTableRow=replace(DiscountTableRow, Session("DiscountRow"&pTempDiscCode), "")
												passDiscountCnt=passDiscountCnt-1
												intArryCnt=intArryCnt-1		
											end if
										end if
									end if	
									
								end if	
								set rs2=nothing	

								pDiscountError=""
								tempDiscountAmount=0
								Session("DiscountTotal"&pTempDiscCode)=""
								Session("DiscountRow"&pTempDiscCode)=""										
							END IF					
						Next
					End If
				END IF
				'// END: Double check the discounts are still valid after all discounts have been applied
				


				if pDiscountError="" then

					dim tSubTotal
					tSubTotal=pSubTotal 
					pSubTotal=pSubTotal - discountTotal
					if pSubTotal<0 then
						pSubTotal=0
						discountTotal=tSubTotal						
					end if
				end if

				session("SF_DiscountTotal")= discountTotal
                
'//////////////////////////////////////////////////////////////
'// END - DISCOUNTS BY CODE
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - GIFT WRAPPING
'//////////////////////////////////////////////////////////////
If Session("Cust_GW")="1" Then    
    GWTotal = calGWTotal()
    pSubTotal = pSubTotal + ccur(GWTotal)
End If 
'//////////////////////////////////////////////////////////////
'// END - GIFT WRAPPING
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - TAX CALCULATIONS
'//////////////////////////////////////////////////////////////

Dim ptaxDetailsString, ptaxAvalaraOrder
ptaxAvalaraOrder = 0

'// CALCULATE TAX OR VAT
If ptaxVAT="1" Then

    VATTotal = CalculateTax( _
                                Session("customerType"), _
                                ptaxwholesale, _
                                TAX_SHIPPING_ALONE, _
                                pTaxonCharges, _
                                pcDblServiceHandlingFee, _
                                pcDblShipmentTotal, _
                                pTaxonFees, _
                                TAX_SHIPPING_AND_HANDLING_TOGETHER, _
                                ptaxCanada,  _
                                session("SFTaxZoneRateCnt"),  _
                                pTaxableTotal,  _
                                Session("Cust_GW"),  _
                                GWTotal,  _
                                taxPaymentTotal,  _
                                ptaxVAT,  _
                                discountTotal,  _
                                pcCartArray,  _
                                ppcCartIndex,  _
                                CatDiscTotal,  _
                                GiftWrapPaymentTotal,  _
                                pcv_IsEUMemberState,  _
                                ptaxPrdAmount,  _
                                intTaxExemptZoneFlag _
                                )
    pTaxAmount = 0

Else

    VATTotal = 0
    pTaxAmount = CalculateTax( _
                                Session("customerType"), _
                                ptaxwholesale, _
                                TAX_SHIPPING_ALONE, _
                                pTaxonCharges, _
                                pcDblServiceHandlingFee, _
                                pcDblShipmentTotal, _
                                pTaxonFees, _
                                TAX_SHIPPING_AND_HANDLING_TOGETHER, _
                                ptaxCanada,  _
                                session("SFTaxZoneRateCnt"),  _
                                pTaxableTotal,  _
                                Session("Cust_GW"),  _
                                GWTotal,  _
                                taxPaymentTotal,  _
                                ptaxVAT,  _
                                discountTotal,  _
                                pcCartArray,  _
                                ppcCartIndex,  _
                                CatDiscTotal,  _
                                GiftWrapPaymentTotal,  _
                                pcv_IsEUMemberState,  _
                                ptaxPrdAmount,  _
                                intTaxExemptZoneFlag _
                                )
    
End If

'// Adjust SubTotal for VAT Total
If pcv_IsEUMemberState = 0 Then
    pSubTotal = pSubTotal - VATTotal                    
End If
'//////////////////////////////////////////////////////////////
'// END - TAX CALCULATIONS
'//////////////////////////////////////////////////////////////




      
'//////////////////////////////////////////////////////////////
'// START - GIFT CERTIFICATES
'//////////////////////////////////////////////////////////////

ListGCs = "" ' (This value is very important and gets saved)
ListUsedGCs = "" ' (this is just used briefly below)
TotalGCAmount = 0 ' (this is just used briefly below)
passGCCode = "" ' (this is used further down this section, but not saved)


'// OUT
' 
' ListGCs
' GCAmount
If pcvFromCart="1" AND pTempGC<>"" Then
	pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_DiscCart_2") & " (" & "<b>" &  pTempGC& "</b>)</li>"
	GCAmount=0
Else
    IF pTempGC<>"" THEN
        if intGCIncludeShipping=1 then
            pGCSubTotal=pSubTotal + pcDblShipmentTotal + pcDblServiceHandlingFee + RoundTo(pTaxAmount,.01)
        else
            pGCSubTotal=pSubTotal + pcDblServiceHandlingFee + RoundTo(pTaxAmount,.01)
        end if
        pSubTotal=pGCSubTotal
    
        GCArr=split(pTempGC,",")
        pTempGC=""
        For i=0 to ubound(GCArr)
            if GCArr(i)<>"" AND cdbl(pSubTotal)>0 then
                intDiscMatchFound=0
                if ListUsedGCs<>"" then
                    UsedGCArry=split(ListUsedGCs,",")
                    for t=0 to (ubound(UsedGCArry)-1)
                        if GCArr(i)=UsedGCArry(t) then
                            intDiscMatchFound=1
                        end if
                    next
                end if
            
                if intDiscMatchFound=0 then
            
                    ListUsedGCs=ListUsedGCs&GCArr(i)&","
            
                    query="SELECT pcGCOrdered.pcGO_ExpDate, pcGCOrdered.pcGO_Amount, pcGCOrdered.pcGO_Status, products.Description FROM pcGCOrdered, products WHERE pcGCOrdered.pcGO_GcCode='"&GCArr(i)&"' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"
                    set rsQ=Server.CreateObject("ADODB.Recordset")
                    set rsQ=conntemp.execute(query)
    
                    IF rsQ.eof then
                        pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_orderverify_4") & ": " & "<b>" & GCArr(i) & "</b></li>"
                    ELSE
                        mTest=0
                        pGCExpDate=rsQ("pcGO_ExpDate")
                        pGCAmount=rsQ("pcGO_Amount")
                        if len(pGCAmount)<0 then
                            pGCAmount=0
                        end if
        
                        pGCStatus=rsQ("pcGO_Status")
                        pDiscountDesc=rsQ("Description")
                        if mTest=0 AND ccur(pGCAmount)<=0 then
                            mTest=1
                            pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_msg_1") & "<b>" & pDiscountDesc & "</b>" & " (<b>" & GCArr(i) & "</b>)" & dictLanguage.Item(Session("language")&"_msg_3") & "</li>"
                        end if
                        if mTest=0 AND cint(pGCStatus)<>1 then
                            mTest=1
                            pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_msg_1") & "<b>" & pDiscountDesc & "</b>" & " (<b>" & GCArr(i) & "</b>)" & dictLanguage.Item(Session("language")&"_msg_1a") & "</li>"
                        end if
                        if mTest=0 AND year(pGCExpDate)<>"1900" then
                            if Date()>pGCExpDate then
                                mTest=1
                                pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_msg_1") & "<b>" & pDiscountDesc & "</b>" & " (<b>" & GCArr(i) & "</b>)" & dictLanguage.Item(Session("language")&"_msg_2") & "</li>"
                            end if
                        end if
                        if mTest=0 then
                        'Have Available Amount
                            GCAmount=pGCAmount
                            pTempSubTotal1=pSubTotal - GCAmount
                            tempGCAmount=0
                            if pTempSubTotal1<0 then
                                pGCAmount=pGCAmount-pSubTotal
                                TotalGCAmount=TotalGCAmount+pSubTotal
                                if ListGCs<>"" then
                                    ListGCs=ListGCs & "|g|"
                                end if
                                ListGCs=ListGCs & GCArr(i) & "|s|" & pDiscountDesc & "|s|" & pSubTotal
                                tempGCAmount=pSubTotal
                                if passGCCode<>"" then
                                    passGCCode=passGCCode & ","
                                end if
                                pSubTotal=0
                                passGCCode=passGCCode & GCArr(i)
                            else
                                pSubTotal=pTempSubTotal1
                                TotalGCAmount=TotalGCAmount+pGCAmount
                                pGCAmount=0
                                if ListGCs<>"" then
                                    ListGCs=ListGCs & "|g|"
                                end if
                                ListGCs=ListGCs & GCArr(i) & "|s|" & pDiscountDesc & "|s|" & GCAmount
                                tempGCAmount=GCAmount
                                if passGCCode<>"" then
                                    passGCCode=passGCCode & ","
                                end if
                                passGCCode=passGCCode & GCArr(i)
                            end if
                            if pTempGC<>"" then
                                pTempGC=pTempGC & ","
                            end if
                            pTempGC=pTempGC & GCArr(i)%>
    
                        
            
                        <%end if
                    END IF
                set rs=nothing
                end if 'intDiscMatchFound
            end if
        Next 'GCArr
    
        if pGCError<>"" then
            pGCError="<ul>" & pGCError & "</ul>"
        end if
    
        GCAmount=TotalGCAmount
    END IF
                    
    If pSubTotal<0 Then
        pSubTotal=0
    End If
    
End If

                
If GCAmount=0 Then
    pSubTotal = pSubTotal + pcDblShipmentTotal + pcDblServiceHandlingFee + RoundTo(pTaxAmount, .01)
Else
    If intGCIncludeShipping=0 Then
        pSubTotal = pSubTotal + pcDblShipmentTotal
    End If
End If
                    
'//////////////////////////////////////////////////////////////
'// END - GIFT CERTIFICATES
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - FREE SHIPPING CHECK
'//////////////////////////////////////////////////////////////

Dim pSubTotalCheckFreeShipping
pSubTotalCheckFreeShipping = pSubTotal - pcDblServiceHandlingFee

session("OPCReady") = "" 
session("OPCReady") = IsOPCReady(pSubTotal, pcDblServiceHandlingFee, GCAmount, serviceFreeOverAmt, pcDblShipmentTotal, pcv_FREESHIP, pSubTotalCheckFreeShipping, pcIntIdShipService)

'//////////////////////////////////////////////////////////////
'// END - FREE SHIPPING CHECK
'//////////////////////////////////////////////////////////////




'//////////////////////////////////////////////////////////////
'// START - SAVE CUSTOMER SESSION DATA
'//////////////////////////////////////////////////////////////
response.Clear()
If session("OPCReady")="NO" Then
	Call SetContentType()
	response.Write(tmpErrorOPCReady)
End If

if pcSFUseRewards="" then
	pcSFUseRewards=0
end if

if pTaxAmount="" then
	pTaxAmount=0
end if

if pSubTotal="" then
	pSubTotal=0
end if

if discountAmount="" then
	discountAmount=0
end if

if passDiscountCnt="" then
	passDiscountCnt=0
end if

if VATTotal="" then
	VATTotal =0
end if

if RewardsDollarValue="" then
	RewardsDollarValue=0
end if

if pSFDiscountCodeTotal="" then
	pSFDiscountCodeTotal=0
end if

if pSFSubTotal="" then
	pSFSubTotal=0
end if

if GWTotal="" then
	GWTotal="0"
end if

if pSFCatDiscTotal="" then
	pSFCatDiscTotal=0
end if

if pcSFCartRewards="" then
	pcSFCartRewards=0
end if

if pcIntBalance="" then
	pcIntBalance=0
end if

if GCAmount="" then
	GCAmount=0
end if

if pSubTotal<0 then
    pSubTotal=0
end if



'DA-EDIT REMOVE VAT FROM EU orders with VAt number

''///////////////////////////////////////////////////////////////////////////////////////THIS IS BREAKING SOMETHING!!?!?!?!?

If pcv_IsEUMemberState = 1 Then
'EU Country
		if pcStrBillingVATID <> "" then
		'Got a VAT ID
    		pSubTotal = pSubTotal/1.2
		end if
End If

pSubTotal = RoundTo(pSubTotal, .01)
If pSubTotal=0 AND Not pcIsSubscription Then
    chkPayment="FREE"
End If


If (Not (session("pcSFIdDbSession") = "" OR session("pcSFRandomKey") = "")) AND (session("idCustomer")>"0") Then

    If pidPayment="" then
        pidPayment=0
    End if
    
    If len(pcv_sbTax)>0 Then
        
        query = "UPDATE pcCustomerSessions SET "
        query = query & "pcCustSession_SB_taxAmount=" & pTaxAmount & ", "
        query = query & "pcCustSession_VATTotal=" & VATTotal & ", "
        query = query & "pcCustSession_taxDetailsString='" & ptaxDetailsString & "' "
        query = query & "WHERE pcCustomerSessions.idDbSession=" & session("pcSFIdDbSession") & " "
        query = query & "AND pcCustomerSessions.randomKey=" & session("pcSFRandomKey") & " "
        query = query & "AND pcCustomerSessions.idCustomer=" & session("idCustomer" ) & ";"
        
    Else
    
        query="UPDATE pcCustomerSessions SET "
        query = query & "pcCustSession_GCDetails='" & replace(ListGCs,"'","''") & "', "
        query = query & "pcCustSession_GCTotal=" & GCAmount & ", "
        query = query & "pcCustSession_strBundleArray='" & pSFstrBundleArray & "', "
        query = query & "pcCustSession_CatDiscTotal=" & pSFCatDiscTotal & ", "
        query = query & "pcCustSession_DiscountCodeTotal=" & pSFDiscountCodeTotal & ", "
        query = query & "pcCustSession_pSubTotal=" & pSFSubTotal & ", "
        query = query & "pcCustSession_chkPayment='" & chkPayment & "', "
        query = query & "pcCustSession_RewardsDollarValue=" & RewardsDollarValue & ", "
        query = query & "pcCustSession_GWTotal=" & GWTotal & ", "
        query = query & "pcCustSession_taxAmount=" & pTaxAmount & ", "
        query = query & "pcCustSession_total=" & pSubTotal & ", "
        query = query & "pcCustSession_discountAmount='" & discountAmount & "', "
        query = query & "pcCustSession_intCodeCnt=" & passDiscountCnt & ", "
        query = query & "pcCustSession_VATTotal=" & VATTotal & ", "
        query = query & "pcCustSession_taxDetailsString='" & ptaxDetailsString & "', "
        query = query & "pcCustSession_discountcode='" & passDiscountCode & "', "
        query = query & "pcCustSession_UseRewards=" & pcSFUseRewards & ", "
        query = query & "pcCustSession_RewardsBalance=" & pcIntBalance & ", "
        query = query & "pcCustSession_IdPayment=" & pidPayment & ", "
        query = query & "pcCustSession_CartRewards=" & pcSFCartRewards & ", "
		query = query & "pcCustSession_Avalara=" & ptaxAvalaraOrder & " "
        query = query & "WHERE pcCustomerSessions.idDbSession=" & session("pcSFIdDbSession") & " "
        query = query & "AND pcCustomerSessions.randomKey=" & session("pcSFRandomKey") & " AND "
        query = query & "pcCustomerSessions.idCustomer=" & session("idCustomer") & ";"
        
    End If
    
    Set rs = connTemp.execute(query)
    Set rs = Nothing
    
    session("pcSFCust_FromCart")=""
    session("pcSFCust_discountcode")=""
    session("pcSFCust_DiscountCodeTotal")=""
    session("pcSFCust_discountAmount")=""
    session("pcSFCust_total")=""
    
Else 

    session("pcSFCust_FromCart")=pcvFromCart
    session("pcSFCust_discountcode")=passDiscountCode
    session("pcSFCust_DiscountCodeTotal")=pSFDiscountCodeTotal
    session("pcSFCust_discountAmount")=discountAmount
    session("pcSFCust_total")=pSubTotal
    
End if
'//////////////////////////////////////////////////////////////
'// END - SAVE CUSTOMER SESSION DATA
'//////////////////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////
'// START - OUTPUT
'//////////////////////////////////////////////////////////////
tmpStr=""
pcGlobalDiscError=pGCError & pcGlobalDiscError
if pcGlobalDiscError<>"" then
	tmpStr="|***|ERROR"
	pcGlobalDiscError = "<ul>" & pcGlobalDiscError & "</ul>"
else
	tmpStr="|***|OK"
end if
if passGCCode<>"" then
	if passDiscountCode<>"" then
		if Right(passDiscountCode,1)<>"," then
			passDiscountCode=passDiscountCode & ","
		end if
	end if
	passDiscountCode=passDiscountCode & passGCCode
end if
	
tmpStr=tmpStr & "|***|" & passDiscountCode & "|***|" & pcGlobalDiscError & "|***|" & pcSFUseRewards & "|***|" & chkPayment & "|***|" & pIdPayment & "|***|" & scCurSign & money(pSubTotal) & "|***|" & session("OPCReady") & "|***|" & pSubTotalCheckFreeShipping

response.write tmpStr
'//////////////////////////////////////////////////////////////
'// END - OUTPUT
'//////////////////////////////////////////////////////////////
%>