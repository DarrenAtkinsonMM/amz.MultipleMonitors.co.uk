<%
Function CheckTaxEpt(pcv_IDPro,pcv_StateCode)
	
    Dim rsD, pcArrayD, intCountD, pcv_mc

    query="SELECT pcTEpt_StateCode FROM pcTaxEpt WHERE pcTEpt_StateCode='" & pcv_StateCode & "' AND ((pcTEpt_ProductList like '" & pcv_IDPro & ",%') OR (pcTEpt_ProductList like '%," & pcv_IDPro & ",%'))"
	set rsD = server.CreateObject("ADODB.RecordSet")
    set rsD = connTemp.execute(query)	
	If Not rsD.eof Then		
		CheckTaxEpt=1
        Set rsD = Nothing        
		Exit Function
	end if
	Set rsD = Nothing
	
	query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcv_IDPro & ";"
	set rsD = server.CreateObject("ADODB.RecordSet")
    set rsD = connTemp.execute(query)	
	If Not rsD.eof Then
		
        pcArrayD = rsD.getRows()
		intCountD = ubound(pcArrayD,2)
		Set rsD = Nothing
		
		For pcv_mc=0 to intCountD
			
            query="SELECT pcTEpt_StateCode FROM pcTaxEpt WHERE pcTEpt_StateCode='" & pcv_StateCode & "' AND ((pcTEpt_CategoryList like '" & pcArrayD(0,pcv_mc) & ",%') OR (pcTEpt_CategoryList like '%," & pcArrayD(0,pcv_mc) & ",%'))"
			set rsD = server.CreateObject("ADODB.RecordSet")
            set rsD = connTemp.execute(query)
			If Not rsD.eof Then				
				CheckTaxEpt = 1
                Set rsD = Nothing
				Exit Function
			End If
            
		Next
        
	End If    
    set rsD=nothing

End Function


Public Function CalculateTax(customerType, ptaxwholesale, TaxShippingAlone, pTaxonCharges, pcDblServiceHandlingFee, pcDblShipmentTotal, pTaxonFees, TaxShippingWithHandling, ptaxCanada, SFTaxZoneRateCnt, pTaxableTotal, Cust_GW, GWTotal, taxPaymentTotal, ptaxVAT, discountTotal, pcCartArray, ppcCartIndex, CatDiscTotal, GiftWrapPaymentTotal, pcv_IsEUMemberState, ptaxPrdAmount, intTaxExemptZoneFlag)

    '// Notes:  What do we need to save for tax:  pTaxAmount / VATTotal / ptaxDetailsString

    Dim taxCalcAmt
	taxCalAmt=0
    VATTotal = 0

    If customerType=1 And (ptaxwholesale=0 And pTaxAvalaraEnabled <> 1) And (ptaxCanada<>"1" Or (ptaxCanada="1" And SFTaxZoneRateCnt=0)) Then
        ptaxPrdAmount=ccur(0)
    End If


	If customerType<>1 OR (customerType=1 AND ptaxwholesale=1) OR (pTaxAvalaraEnabled = 1) Then

        If TaxShippingAlone="NA" Then
            
            If pTaxonCharges=1 Then
                taxCalAmt = taxCalAmt + pcDblShipmentTotal
            End If
            If pTaxonFees=1 Then
                taxCalAmt = taxCalAmt + pcDblServiceHandlingFee
            End If
            
        Else '// If TaxShippingAlone="NA" Then

            If TaxShippingWithHandling="Y" Then
                taxCalAmt = taxCalAmt + pcDblShipmentTotal + pcDblServiceHandlingFee
            Else
            
                If TAX_SHIPPING_ALONE="Y" Then
                    taxCalAmt = taxCalAmt + pcDblShipmentTotal
                End If
                
            End If

        End If '// If TaxShippingAlone="NA" Then
			
        '//////////////////////////////////
        '// CANADA
        '////////////////////////////////// 
            		
        If ptaxCanada="1" and SFTaxZoneRateCnt>0 Then
            taxCalAmt = 0
		End If 

        Dim ptaxLocAmount
        ptaxLocAmount = 0

        If (cdbl(taxCalAmt)=0 AND cdbl(pTaxableTotal)=0 AND (ptaxCanada<>"1" OR (ptaxCanada="1" AND SFTaxZoneRateCnt=0))) AND (pTaxAvalaraEnabled <> 1) then

            ptaxLocAmount=0
            
        Else

            '// Gift
            If Cust_GW="1" Then
                pTaxableTotal = pTaxableTotal + ccur(GWTotal)
            End If

            '// If VAT
            Dim VATTotal
            VATTotal=0
            
            If taxPaymentTotal="" Then
                taxPaymentTotal = 0
            End If
						
            '////////////////////////////////////////////////////////////////////////////////////////
            '// START: VAT
            '////////////////////////////////////////////////////////////////////////////////////////
  
            If ptaxVAT="1" Then
                
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '// Start: Discount Distribution %
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				Dim ApplicableDisountTotal
				ApplicableDisountTotal = pTaxableTotal + taxCalAmt + taxPaymentTotal
							
				'// Shipping and Handling represents what % of the Total Discount?  							
				Proportional_taxCalAmt = RoundTo((taxCalAmt/ApplicableDisountTotal),.01)
							
				'// Payment Charges represents what % of the Total Discount? 							
				Proportional_taxPaymentTotal = RoundTo((taxPaymentTotal/ApplicableDisountTotal),.01)
							
				'// Product Pricing represents what % of the Total Discount?
				'	NOTE: Product Level Distributions are calculated at the line item via "calculateVATTotal"
							
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '// End: Discount Distribution %
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '// Start: Distribute Discounts based off % above
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
				'// Shipping and Handling after discount
				ApplicableDisount_taxCalAmt = (discountTotal * Proportional_taxCalAmt)
				taxCalAmt = (taxCalAmt - ApplicableDisount_taxCalAmt)
							
				'// Payment Charges after discount
				ApplicableDisount_taxPaymentTotal = (discountTotal * Proportional_taxPaymentTotal)
				taxPaymentTotal = (taxPaymentTotal - ApplicableDisount_taxPaymentTotal)
							
			    '// Products Price after discount
			    '	NOTE: Discount are distributed to Products at the line item via "calculateVATTotal"
							
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '// End: Distribute Discounts based off % above
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							

							
				'// VAT TAXED AMOUNT - ORDER LEVEL ("Discounts" Removed Proportionately)
				VatTaxedAmount_OrderLevel = taxCalAmt + taxPaymentTotal	
							
				'// VAT TAXED AMOUNT - PRODUCT LEVEL ("Discounts" and "Category Discounts" Removed Proportionately)						
				VatTaxedAmount_ProductLevel = ccur( calculateVATTotal(pcCartArray, ppcCartIndex, discountTotal, CatDiscTotal, ApplicableDisountTotal) )
							
				'// VAT TAXED AMOUNT
				VatTaxedAmount = VatTaxedAmount_OrderLevel + VatTaxedAmount_ProductLevel
							
				'// Shipping and Handling "VATable" Total - Uses Default Rate							
				taxCalAmtNoVAT = RoundTo(pcf_RemoveVAT(taxCalAmt,""),.01)
				CalAmtTotal = RoundTo(taxCalAmt-taxCalAmtNoVAT,.01)
										
				'// Payment Charges "Always VATable" Total - Uses Default Rate
				taxPaymentTotalNoVAT = RoundTo(pcf_RemoveVAT(taxPaymentTotal,""),.01)
				tPaymentTotal = RoundTo(taxPaymentTotal-taxPaymentTotalNoVAT,.01)
							
				'// Gift Wrapping Charges "Always VATable" Total - Uses Default Rate
				taxGiftWrapNoVAT = RoundTo(pcf_RemoveVAT(GiftWrapPaymentTotal,""),.01)
				tGiftWrapTotal = RoundTo(GiftWrapPaymentTotal-taxGiftWrapNoVAT,.01)							


                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '// Start: VAT Totals
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
				'// Order Level Total VAT ("Discounts" Removed)
				VATTotal_OrderLevel = CalAmtTotal + tPaymentTotal + tGiftWrapTotal	'// Payment VAT + Shipping/Handling VAT					
							
				'// Product Level Total VAT ("Discounts" and "Category Discounts" Removed)
				NoVATTotal=ccur( calculateNoVATTotal(pcCartArray, ppcCartIndex, discountTotal, CatDiscTotal, ApplicableDisountTotal) )	
				NoVATTotal=RoundTo(NoVATTotal,.01)
				VATTotal_ProductLevel=RoundTo(VatTaxedAmount_ProductLevel-NoVATTotal,.01)
							
				'// Total VAT
				VATTotal=VATTotal_OrderLevel+VATTotal_ProductLevel							
							
				'// NOTE: CalAmtTotal is included in the Subtotal for display purposes, but technically is applied to Order Level.
							
				'// Display the correct Sub Total when outside the EU
				'If pcv_IsEUMemberState = 0 Then
                
                    '// TO DO: MOVE THIS OUTSIDE
				    'pSubTotal = pSubTotal - VATTotal_ProductLevel - tPaymentTotal - tGiftWrapTotal - CalAmtTotal '// Remove VAT charges from the total.
                    'pSubTotal = pSubTotal - VATTotal
                    
				'End If
                
                            
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '// END: VAT Totals
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
				'// For reference determine the Total VAT Taxed Amount
				'	NOTE: Includes "Shipping and Handling" + "Payment Charges" - "Discounts" - "Cat Discounts" (discounts applied proportionately)
				VATTaxedAmount=RoundTo(VatTaxedAmount,.01)
				If VATTaxedAmount<0 Then
				    VATTaxedAmount=0
				End If
							
				'// If outside EU then specifiy how much VAT was removed
				VATRemovedTotal=0
				If pcv_IsEUMemberState=0 Then
				    VATRemovedTotal = VATTotal
					VATTaxedAmount=0
				End If	
             
            '////////////////////////////////////////////////////////////////////////////////////////
            '// END: VAT
            '////////////////////////////////////////////////////////////////////////////////////////


            ElseIf ptaxAvalara = 1 AND pTaxAvalaraEnabled = 1 Then


				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start AVALARA TAX SYSTEM
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
                '// Avalara Doc
                pDocCode = session("idOrderSaved")
				if pDocCode = "" then
					query = "SELECT TOP 1 idOrder FROM orders ORDER BY idOrder DESC"
					set rs = server.CreateObject("ADODB.RecordSet")
					set rs = conntemp.execute(query)
					
					if not rs.eof then
						pDocCode = rs("idOrder") + 1
					else
						pDocCode = 1
					end if
				end if				
				pDocDate = year(date) & "-" & format_zeros(month(date), 2) & "-" & format_zeros(day(date), 2)

				'// Avalara Request
				postdata = postdata & "<GetTaxRequest>" & vbCrlf
				postdata = postdata & "<Client>a0o33000003xKBM</Client>" & vbCrlf
				postdata = postdata & "<DocDate>" & pDocDate & "</DocDate>" & vbCrlf
				postdata = postdata & "<CustomerCode>" & "CUST-" & session("idCustomer") & "</CustomerCode>" & vbCrlf
				postdata = postdata & "<DocCode>" & scpre+int(pDocCode) & "</DocCode>" & vbCrlf
                
                'postdata = postdata & "<DocType>" & "SalesInvoice" & "</DocType>" & vbCrlf
				postdata = postdata & "<DocType>" & "SalesOrder" & "</DocType>" & vbCrlf
				
				query = "SELECT pcCust_AvalaraExemptionNo FROM customers WHERE idcustomer = " & session("idCustomer")
				set rs = server.CreateObject("ADODB.RecordSet")
				set rs = conntemp.execute(query)
                If Not rs.Eof Then
                    pcv_strAvalaraExemptionNo = rs("pcCust_AvalaraExemptionNo")
                End If
				
                
                '// There is a customer excemption number
                If len(pcv_strAvalaraExemptionNo)>0 Then
                    postdata = postdata & "<ExemptionNo>" & rs("pcCust_AvalaraExemptionNo") & "</ExemptionNo>" & vbCrlf
                Else
                '// Customer has no excemption number
                    If ptaxwholesale = 0 And customerType = 1 Then
                        postdata = postdata & "<CustomerUsageType>" & ptaxAvalaraReason & "</CustomerUsageType>" & vbCrlf
                    End If
                End If
                set rs = nothing
				
                '// Addresses
                postdata = postdata & pcf_avaAddresses(pcStrShippingAddress, pcStrShippingAddress2, pcStrShippingCity, pcStrShippingStateCode, pcStrShippingPostalCode, pcStrBillingAddress, pcStrBillingAddress2, pcStrBillingCity, pcStrBillingStateCode, pcStrBillingPostalCode)

                '// Lines
                postdata = postdata & "<Lines>" & vbCrlf
				
				'// Product Tax
				for f=1 to pcCartIndex
					if pcCartArray(f,10)=0 then
                    
                        pcv_intProductLineTotal = 0
                        
                        '// Bundle Discounts
                        If (pcCartArray(f,27)>0) And (pcCartArray(f,28)>0) Then
                            pcv_intBundleTotal = ((ccur(pcCartArray(f,28) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1)) 
                            pcv_intProductLineTotal = pcCartArray(f,2) * (pcCartArray(f,5) + pcCartArray(f,3) + pcv_intBundleTotal)
                        Else
                            pcv_intProductLineTotal = pcCartArray(f,2) * (pcCartArray(f,5) + pcCartArray(f,3))
                        End If

                        '// Item Discounts
                        If pcCartArray(f,16) <> "" Then
                            If pcCartArray(f,30) <> 0 Then
                                pcv_intProductLineTotal =  pcv_intProductLineTotal - pcCartArray(f,30)
                            End If                
                        End If 
                        
                        '// Additional Charges
                        pcv_intAddCharges = 0
                        If pcCartArray(f,16) <> "" Then
                            query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
                            Set rs = server.CreateObject("ADODB.RecordSet")	
                            Set rs = conntemp.execute(query)
                            If Not rs.Eof Then
                                    stringCProducts=rs("stringCProducts")
                                    stringCValues=rs("stringCValues")
                                    ArrCProduct=Split(stringCProducts, ",")
                                    ArrCValue=Split(stringCValues, ",")
                            End If
                            Set rs = Nothing                    
                            If ArrCProduct(0)<>"na" Then
                                For i=lbound(ArrCProduct) To (UBound(ArrCProduct)-1)
                                    pcv_intAddCharges = pcv_intAddCharges + ArrCValue(i)
                                Next                                
                            End If 
                            If pcv_intAddCharges > 0 Then
                                pcv_intProductLineTotal =  pcv_intProductLineTotal + pcv_intAddCharges
                            End If          
                        End If 
                    
                        '// Product Line Item
                        postdata = postdata & pcf_avaLineItem(f, pcCartArray(f, 7), pcCartArray(f, 1), pcCartArray(f, 2), pcv_intProductLineTotal, pcCartArray(f,44), pcCartArray(f,45))
                        
					end if
				next 
                
				'// Payment Fees
				postdata = postdata & pcf_avaPaymentTotal(f, taxPaymentTotal)               
				
				'// Shipping charge Tax
				postdata = postdata & pcf_avaShipmentTotal(f, pcDblShipmentTotal)
				
				'// Handling fee Tax
				postdata = postdata & pcf_avaServiceHandlingFee(f, pcDblServiceHandlingFee)

                '// Gift Wrapping
                postdata = postdata & pcf_avaGiftWrap(f, GWTotal)

				'// Reward Points
                postdata = postdata & pcf_avaRewardPoints(f, session("SF_RewardPointTotal"))

				postdata = postdata & "</Lines>" & vbCrlf
                
                '// Discounts
				postdata = postdata & pcf_avaDiscounts(f, CatDiscTotal, discountTotal, TotalPromotions)

				postdata = postdata & "</GetTaxRequest>"

                '// Get Tax Total
				ptaxPrdAmount = pcf_GetAvalaraTaxAmount(postdata)

				ptaxDetailsString = "Tax|" & ptaxPrdAmount & ","
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// END AVALARA TAX SYSTEM
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
			Else '// If ptaxVAT="1" Then


                If ptaxCanada="1" And SFTaxZoneRateCnt>0 Then
                
                    '// Calculate Zone Taxes
					ptaxDetailsString=""
					ptaxLocAmount=0
                    
					pTempTaxableCanadaTotal=0
                    
                    For u=1 To SFTaxZoneRateCnt
                    
                        taxCalAmt=0
                        session("taxAmount"&u)=0
                        pcv_IntTaxZoneRateID = session("SFTaxZoneRateID"&u)
                        Dim pTaxZoneExemption
                        
                        If intTaxExemptZoneFlag="1" Then
                        
                            '// Recalculate taxable total
                            If pTempTaxableCanadaTotal>0 Then
                                pTaxableCanadaTotal = pTempTaxableCanadaTotal
							Else
							    pTaxableCanadaTotal = ccur(calculateTaxableZoneTotal(pcCartArray, ppcCartIndex, pcv_IntTaxZoneRateID))
							End If

                            If session("SFTaxZoneRateApplyToSH"&u) Then
                                taxCalAmt = Cdbl(pcDblShipmentTotal) + Cdbl(pcDblServiceHandlingFee)
                            End If

                            '// Change per rewards points
                            If pTaxableCanadaTotal>0 Or taxCalAmt>0 Then
                            
                                ptaxLoc = session("SFTaxZoneRateRate"&u)
                                tempTAmt = ((pTaxableCanadaTotal + taxCalAmt + taxPaymentTotal - discountTotal - CatDiscTotal) * ptaxLoc)
                                tempTAmt = roundTo(tempTAmt,.01)
                                
                                If tempTAmt<0 Then
                                    tempTAmt=0
                                End If

                                ptaxLocAmount = ptaxLocAmount + tempTAmt
								session("taxAmount"&u) = tempTAmt
								session("taxDesc"&u) = session("SFTaxZoneRateName"&u)

                            Else '// If pTaxableCanadaTotal>0 Or taxCalAmt>0 Then
                                
                                tempTAmt=0
                                
							End If
                            
                            If session("SFTaxZoneRateTaxable"&u)="1" Then
                                pTempTaxableCanadaTotal=pTaxableCanadaTotal+tempTAmt
                            End If

                        Else '// If intTaxExemptZoneFlag="1" Then
                        
                            pTaxZoneExemption = checkTaxExempt(pcCartArray, ppcCartIndex, pcv_IntTaxZoneRateID)
										
                            If pTaxZoneExemption=0 Then

                                ptaxLoc = session("SFTaxZoneRateRate"&u)
                                If session("SFTaxZoneRateApplyToSH"&u) Then
                                    taxCalAmt = taxCalAmt + pcDblShipmentTotal + pcDblServiceHandlingFee
                                End If

                                tempTAmt=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal-TotalPromotions) * ptaxLoc)
                                tempTAmt=roundTo(tempTAmt,.01)
                                If tempTAmt<0 Then
                                    tempTAmt=0
                                End If
								
                                ptaxLocAmount = ptaxLocAmount + tempTAmt
								session("taxAmount"&u) = tempTAmt
								session("taxDesc"&u) = session("SFTaxZoneRateName"&u)
								
                                If session("SFTaxZoneRateTaxable"&u)="1" Then
								    pTempTaxableCanadaTotal = pTaxableCanadaTotal + tempTAmt								
                                End If

                            Else '// If pTaxZoneExemption=0 Then
                            
							    session("taxAmount"&u)=0
								session("taxDesc"&u)=session("SFTaxZoneRateName"&u)
                                
							End If '// If pTaxZoneExemption=0 Then
                                        
                        End If '// If intTaxExemptZoneFlag="1" Then
						
                        ptaxDetailsString = ptaxDetailsString & replace(session("taxDesc"&u),",","") & "|" & session("taxAmount"&u) & ","
                    
                    Next '// For u=1 To SFTaxZoneRateCnt

                    session("taxCnt")=session("SFTaxZoneRateCnt")

                Else '// If ptaxCanada="1" And SFTaxZoneRateCnt>0 Then
                
					if session("taxCnt")>"0" then
						ptaxDetailsString=""
						ptaxLocAmount=0
						for i=1 to session("taxCnt")
							ptaxLoc=session("tax"&i)
							tempTAmt=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal-TotalPromotions) * ptaxLoc)
							tempTAmt=roundTo(tempTAmt,.01)
							if tempTAmt<0 then
								tempTAmt=0
							end if
							ptaxLocAmount=ptaxLocAmount+tempTAmt
							session("taxAmount"&i)=tempTAmt
							ptaxDetailsString=ptaxDetailsString&replace(session("taxDesc"&i),",","")&"|"&session("taxAmount"&i)&","
						next
					else
						ptaxLocAmount=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal-TotalPromotions) * ptaxLoc)
						ptaxLocAmount = CCur(pTaxLocAmount)
						ptaxLocAmount=RoundTo(ptaxLocAmount,.01)
						if ptaxLocAmount<0 then
							ptaxLocAmount=0
						end if
					end if

                End If '// If ptaxCanada="1" And SFTaxZoneRateCnt>0 Then 
                            
            End If '// If ptaxVAT="1" Then 
                        
        End If '// If cdbl(taxCalAmt)=0 AND cdbl(pTaxableTotal)=0 AND (ptaxCanada<>"1" OR (ptaxCanada="1" AND SFTaxZoneRateCnt=0)) then 

             
    Else '// If customerType<>1 OR (customerType=1 AND ptaxwholesale=1) Then

        ptaxLocAmount = 0
        
    End If '// If customerType<>1 OR (customerType=1 AND ptaxwholesale=1) Then

    CalculateTax = ptaxPrdAmount + ptaxLocAmount + VATTotal
    
End Function





'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'// START:  v4 Methods - Need to consolidate and merge with v5 above
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Function CheckTaxEptZone(pcv_IDPro,pcv_ZoneRateID)
	Dim query, rsD,pcArrayD,intCountD,pcv_mc
	CheckTaxEptZone=0
	
	query="SELECT pcTaxEpt.pcTaxZoneRate_ID FROM pcTaxEpt WHERE pcTaxEpt.pcTaxZoneRate_ID=" & pcv_ZoneRateID & " AND ((pcTEpt_ProductList like '" & pcv_IDPro & ",%') OR (pcTEpt_ProductList like '%," & pcv_IDPro & ",%'))"
	set rsD=connTemp.execute(query)
	
	if not rsD.eof then
		set rsD=nothing
		CheckTaxEptZone=1
		Exit Function
	end if
	set rsD=nothing
	
	query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcv_IDPro & ";"
	set rsD=connTemp.execute(query)
	if not rsD.eof then
		pcArrayD=rsD.getRows()
		intCountD=ubound(pcArrayD,2)
		set rsD=nothing
		
		For pcv_mc=0 to intCountD
			query="SELECT pcTaxZoneRate_ID FROM pcTaxEpt WHERE pcTaxZoneRate_ID=" & pcv_ZoneRateID & " AND ((pcTEpt_CategoryList like '" & pcArrayD(0,pcv_mc) & ",%') OR (pcTEpt_CategoryList like '%," & pcArrayD(0,pcv_mc) & ",%'))"
			set rsD=connTemp.execute(query)
			if not rsD.eof then
				set rsD=nothing
				CheckTaxEptZone=1
				Exit Function
			end if
		Next
	end if
set rsD=nothing

End Function


function NontaxableItems(tmpIDConfigSession,parentQty)
	Dim rs,query,i,tmpResult
	Dim stringProducts,stringValues,stringCategories,Qstring,Pstring,ArrProduct,ArrValue,ArrCategory,ArrQuantity,ArrPrice
	Dim TempDiscount,TempD1,QFrom,QTo,DUnit,QPercent,DWUnit

	tmpResult=0

	IF tmpIDConfigSession<>"" AND IsNumeric(tmpIDConfigSession) THEN

		query="SELECT stringProducts, stringValues, stringCategories,stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & tmpIDConfigSession
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
							
		if not rs.eof then
			stringProducts=rs("stringProducts")
			stringValues=rs("stringValues")
			stringCategories=rs("stringCategories")
			Qstring=rs("stringQuantity")
			Pstring=rs("stringPrice")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			ArrQuantity=Split(Qstring,",")
			ArrPrice=split(Pstring,",")
			set rs=nothing
						
			if ArrProduct(0)="na" then
			else
				For i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					if ArrProduct(i)<>"" then
						query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i) & " AND notax<>0;"
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpResult=tmpResult+ArrQuantity(i)*ArrPrice(i)*parentQty
							
							query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i)
							set rs=connTemp.execute(query)
		 
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
								if (clng(ArrQuantity(i)*parentQty)>=clng(QFrom)) and (clng(ArrQuantity(i)*parentQty)<=clng(QTo)) then
									if QPercent="-1" then
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DUnit
										end if
									else
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*DUnit
										end if
									end if
								end if
								TempDiscount=TempDiscount+TempD1
								rs.movenext
							loop
							set rs=nothing
							tmpResult=tmpResult-TempDiscount
							
						end if
						set rs=nothing
					end if
				Next
			end if
		end if
		set rs=nothing

	END IF

	NontaxableItems=tmpResult

End function


function NontaxableZoneItems(tmpIDConfigSession,parentQty,zoneRateID)

	Dim rs,query,i,tmpResult
	Dim stringProducts,stringValues,stringCategories,Qstring,Pstring,ArrProduct,ArrValue,ArrCategory,ArrQuantity,ArrPrice
	Dim TempDiscount,TempD1,QFrom,QTo,DUnit,QPercent,DWUnit

	tmpResult=0

	IF tmpIDConfigSession<>"" AND IsNumeric(tmpIDConfigSession) THEN
		query="SELECT stringProducts, stringValues, stringCategories,stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & tmpIDConfigSession
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
							
		if not rs.eof then
			stringProducts=rs("stringProducts")
			stringValues=rs("stringValues")
			stringCategories=rs("stringCategories")
			Qstring=rs("stringQuantity")
			Pstring=rs("stringPrice")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			ArrQuantity=Split(Qstring,",")
			ArrPrice=split(Pstring,",")
			set rs=nothing
						
			if ArrProduct(0)="na" then
			else
				For i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					if ArrProduct(i)<>"" then
						ztest=CheckTaxEptZone(ArrProduct(i),zoneRateID)
						
						if ztest=1 then
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i)
						else
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i) & " AND notax<>0;"
						end if
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpResult=tmpResult+ArrQuantity(i)*ArrPrice(i)*parentQty
							
							query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i)
							set rs=connTemp.execute(query)
		 
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
								if (clng(ArrQuantity(i)*parentQty)>=clng(QFrom)) and (clng(ArrQuantity(i)*parentQty)<=clng(QTo)) then
									if QPercent="-1" then
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DUnit
										end if
									else
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*DUnit
										end if
									end if
								end if
								TempDiscount=TempDiscount+TempD1
								rs.movenext
							loop
							set rs=nothing
							tmpResult=tmpResult-TempDiscount
							
						end if
						set rs=nothing
					end if
				Next
			end if
		end if
		set rs=nothing
	END IF


	IF tmpIDConfigSession<>"" AND IsNumeric(tmpIDConfigSession) THEN
		query="SELECT stringQuantity, stringPrice, stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & tmpIDConfigSession
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
							
		if not rs.eof then
			stringProducts=rs("stringCProducts")
			stringValues=rs("stringCValues")
			stringCategories=rs("stringCCategories")
			Qstring=rs("stringQuantity")
			Pstring=rs("stringPrice")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			ArrQuantity=Split(Qstring,",")
			ArrPrice=split(Pstring,",")
			set rs=nothing
						
			if ArrProduct(0)="na" then
			else
				For i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					if ArrProduct(i)<>"" then
						ztest=CheckTaxEptZone(ArrProduct(i),zoneRateID)
						
						if ztest=1 then
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i)
						else
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i) & " AND notax<>0;"
						end if
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpResult=tmpResult+1*ArrValue(i)
						end if
						set rs=nothing
					end if
				Next
			end if
		end if
		set rs=nothing
	END IF

	NontaxableZoneItems=tmpResult

End function


' Cart Taxable Amount
function calculateTaxableTotal(pcCartArray, indexCart)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	'SB S
	Dim subInstArr
	'SB E
	
	total=0

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			'SB S
			If Not (len(pcCartArray(f,38))>0) Then pcCartArray(f,38)=0
			if pcCartArray(f,38) > 0 then 
				subInstArr = split(getSubInstallVals(pcCartArray(f,38)),",")
			else
				subInstArr = split("0,0,0,0",",") 
			end if
			'SB E
			if pcCartArray(f,16)<>"" then
				'SB S
				if subInstArr(2) = "1" Then 
					total = total + (pcCartArray(f,2) * cdbl(subInstArr(3))) - NontaxableItems(pcCartArray(f,16),pcCartArray(f,2))
				else
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31) - NontaxableItems(pcCartArray(f,16),pcCartArray(f,2))
				end if
				'SB E
			else
				'SB S
				If subInstArr(2) = "1" Then
					total = total + (pcCartArray(f,2) * cdbl(subInstArr(3)))
				Else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if
				End If
				'SB E
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if

		end if
	next
	
	calculateTaxableTotal=total
	set f=nothing
	set total=nothing 
end function


'SB S
function calculateTaxableTotal_SB(pcCartArray, indexCart)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31) - NontaxableItems(pcCartArray(f,16),pcCartArray(f,2))
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if

		end if
	next
	
	calculateTaxableTotal_SB=total
	set f=nothing
	set total=nothing 
end function
'SB E


' Cart Taxable Amount
function calculateTaxableZoneTotal(pcCartArray, indexCart, zoneRateID)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		mtest=0
		
		mtest=CheckTaxEptZone(pcCartArray(f,0),zoneRateID)

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31) - NontaxableZoneItems(pcCartArray(f,16),pcCartArray(f,2),zoneRateID)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if

		end if
	next
	
	calculateTaxableZoneTotal=total

	set f=nothing
	set total=nothing 
end function


function checkTaxExempt(pcCartArray, indexCart, TaxZoneRateID)
	Dim f, checkVar, pcv_StateCode, pcv_Country
	
	strCheckVar=""

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND (pcv_StateCode<>"") and (ucase(pcv_Country)="CA") then
			strCheckVar=strCheckVar&CheckTaxEptZone(pcCartArray(f,0),TaxZoneRateID)&","
		end if
	next
	
	if instr(strCheckVar,"1") then
		if instr(strCheckVar,"0") then
			checkVar=0
		else
			checkVar=1
		end if
	else
		checkVar=0
	end if
	
	checkTaxExempt=checkVar
end function


'////////////////////////////////////////////////////////////////////////////////////////
'// START: VAT CALCULATIONS
'////////////////////////////////////////////////////////////////////////////////////////

'// Removes VAT from UnitPrice when purchased outside the European Union.
Function pcf_VAT(Price, ProductID)
	Dim notax
	If pcv_IsEUMemberState=1 OR pcv_IsEUMemberState=0 Then
		notax="-1"
		If not validNum(ProductID) Then 
			ProductID = 0
		End If		
		If ProductID<>0 Then '// Product			
			query="SELECT products.notax FROM products WHERE idProduct=" & ProductID & " AND configOnly=0 AND removed=0 " 
			set rsVAT=server.CreateObject("ADODB.RecordSet")
			set rsVAT=connTemp.execute(query)
			If NOT rsVAT.eof Then
				notax=rsVAT("notax")
			End If
			set rsVAT=nothing
			If ptaxVAT="1" AND notax <> "-1" Then
				if pcv_IsEUMemberState=0 then
					pcf_VAT=pcf_RemoveVAT(Price, ProductID)
					Exit Function
				else
					pcf_VAT=Price
					Exit Function
				end if
			Else
				pcf_VAT=Price
				Exit Function
			End If
		Else '// VAT Item - No Product ID		
			If ptaxVAT="1" Then
				if pcv_IsEUMemberState=0 then
					pcf_VAT=pcf_RemoveVAT(Price, ProductID)
					Exit Function
				else
					pcf_VAT=Price
					Exit Function
				end if
			Else
				pcf_VAT=Price
				Exit Function
			End If
		End If
	Else '// If pcv_IsEUMemberState=1 OR pcv_IsEUMemberState=0 Then
		pcf_VAT=Price
		Exit Function
	End If
End Function


'// Removes VAT from Price
Function pcf_RemoveVAT(Price,ProductID)
	pcf_RemoveVAT=Price/(1+pcf_VATRate(ptaxVATRate_Code, ProductID)/100)
End Function


'// Determines the correct VAT from a Product ID
Function pcf_VATRate(StateCode,ProductID)
	If StateCode="0" OR StateCode="" Then
		pcf_VATRate=ptaxVATrate '// Default Rate
		Exit Function
	Else
		if not validNum(ProductID) then ProductID = 0
		if ProductID="" OR ProductID=0 OR isNULL(ProductID)=True then
			pcf_VATRate=ptaxVATrate '// Default Rate
			Exit Function
		else
			query="SELECT pcProductsVATRates.pcVATRate_ID, pcVATRates.pcVATRate_Rate, pcVATRates.pcVATRate_ID "
			query=query&"FROM pcProductsVATRates, pcVATRates "
			query=query&"WHERE pcProductsVATRates.pcVATRate_ID=pcVATRates.pcVATRate_ID AND pcProductsVATRates.idProduct="&ProductID	
			Set rsVAT=Server.CreateObject("ADODB.Recordset")  
			set rsVAT=connTemp.execute(query)
			if not rsVAT.eof then
				pcf_VATRate=rsVAT("pcVATRate_Rate") '// Category Rate
			else
				pcf_VATRate=ptaxVATrate '// Default Rate
			end if
			set rsVAT=nothing			
			Exit Function
		end if		
	End If
End Function


'// Determines if a country is apart of the European Union
' 1 = YES
' 0 = NO
' 2 = Not Applicable
' 5 = GB
Function pcf_IsEUMemberState(CountryCode)
	pcf_IsEUMemberState=2
	if CountryCode="GB" then
		pcf_IsEUMemberState=5
		Exit Function
	end if
	If (CountryCode<>"" AND isNULL(CountryCode)=False) Then
		if (UCASE(CountryCode)=UCASE(scCompanyCountry)) AND ptaxVAT="1" then
			pcf_IsEUMemberState=1
			Exit Function
		end if	
	End If
	If ptaxVAT="1" AND (CountryCode<>"" AND isNULL(CountryCode)=False) Then		
		query="SELECT pcVATCountries.pcVATCountry_Code From pcVATCountries WHERE pcVATCountries.pcVATCountry_Code='"&UCASE(CountryCode)&"';"
		set rsVAT=Server.CreateObject("ADODB.Recordset")
		set rsVAT=connTemp.execute(query)
		if not rsVAT.eof then
			pcf_IsEUMemberState=1
		else
			pcf_IsEUMemberState=0
		end if
		set rsVAT=nothing
	End If
End Function


' Cart VAT Total and Remove VAT
function calculateNoVATTotal(pcCartArray, indexCart, TotalStandardDiscount, TotalCategoryDiscount, ApplicableDisountTotal)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0
	grandtotal=0
	
	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		total=0
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if	 
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total = ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if
		end if

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// This Line Item represents what % of the Total Discount?  							
		Proportional_total = RoundTo((total/ApplicableDisountTotal),.01)
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Line Item after discount
		ApplicableDisount_total = (TotalStandardDiscount * Proportional_total)
		total = (total - ApplicableDisount_total)	
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		if pcv_strrApplicableCategories<>"" then
			pcArray_ApplicableCategories = split(pcv_strrApplicableCategories, ",")
			for y=0 to ubound(pcArray_ApplicableCategories)-1 '// For Each Category Discount Available
				pcArray_ApplicableCategory = split(pcArray_ApplicableCategories(y), chr(124))
				tmpApplicableCategoryID = pcArray_ApplicableCategory(1)
				tmpCategorySubTotal = pcArray_ApplicableCategory(0)
			
				ApplicableCategoryItem=False
				if pcv_strApplicableProducts <> "" then
					pcArray_ApplicableProducts = split(pcv_strApplicableProducts, ",")
					for x=0 to ubound(pcArray_ApplicableProducts)-1 '// Loop through all Products						
						pcArray_ApplicableProduct = split(pcArray_ApplicableProducts(x), chr(124))
						tmpProductID = pcArray_ApplicableProduct(0)
						tmpCategoryID = pcArray_ApplicableProduct(1)				
						if (tmpProductID = pcCartArray(f,0)) AND (tmpCategoryID = tmpApplicableCategoryID) then '// This Product is Applicable to this Category
							ApplicableCategoryItem=True
						end if
					next
				end if  '// if pcv_strApplicableProducts <> "" then
				
				If ApplicableCategoryItem=True AND tmpCategorySubTotal>0 Then
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// This Line Item represents what % of the Total Category?  							
				ProportionalCat_total = RoundTo((total/tmpCategorySubTotal),.01)
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// Line Item after Category
				ApplicableCatDisount_total = (TotalCategoryDiscount * ProportionalCat_total)
				total = (total - ApplicableCatDisount_total)	
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				End If
				
			next
		end if '// if pcv_strrApplicableCategories<>"" then	



		total = pcf_RemoveVAT(total, pcCartArray(f,0))
		grandtotal = grandtotal + total
	next
	
	calculateNoVATTotal=grandtotal
	set f=nothing
	set total=nothing 
end function


' Cart VAT Total
function calculateVATTotal(pcCartArray, indexCart, TotalStandardDiscount, TotalCategoryDiscount, ApplicableDisountTotal)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0
	grandtotal=0
	
	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		total=0
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3))
				end if 	 
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total = ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if
		end if
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// This Line Item represents what % of the Total Discount?  			
		Proportional_total = RoundTo((total/ApplicableDisountTotal),.01)
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Line Item after discount
		ApplicableDisount_total = (TotalStandardDiscount * Proportional_total)
		total = (total - ApplicableDisount_total)	
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			



		if pcv_strrApplicableCategories<>"" then
			pcArray_ApplicableCategories = split(pcv_strrApplicableCategories, ",")
			for y=0 to ubound(pcArray_ApplicableCategories)-1 '// For Each Category Discount Available
				pcArray_ApplicableCategory = split(pcArray_ApplicableCategories(y), chr(124))
				tmpApplicableCategoryID = pcArray_ApplicableCategory(1)
				tmpCategorySubTotal = pcArray_ApplicableCategory(0)
			
				ApplicableCategoryItem=False
				if pcv_strApplicableProducts <> "" then
					pcArray_ApplicableProducts = split(pcv_strApplicableProducts, ",")
					for x=0 to ubound(pcArray_ApplicableProducts)-1 '// Loop through all Products						
						pcArray_ApplicableProduct = split(pcArray_ApplicableProducts(x), chr(124))
						tmpProductID = pcArray_ApplicableProduct(0)
						tmpCategoryID = pcArray_ApplicableProduct(1)					
						if (tmpProductID = pcCartArray(f,0)) AND (tmpCategoryID = tmpApplicableCategoryID) then '// This Product is Applicable to this Category
							ApplicableCategoryItem=True
						end if
					next
				end if  '// if pcv_strApplicableProducts <> "" then
				
				If ApplicableCategoryItem=True AND tmpCategorySubTotal>0 Then
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// This Line Item represents what % of the Total Category?  							
				ProportionalCat_total = RoundTo((total/tmpCategorySubTotal),.01)
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// Line Item after Category
				ApplicableCatDisount_total = (TotalCategoryDiscount * ProportionalCat_total)
				total = (total - ApplicableCatDisount_total)	
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				End If
				
			next
		end if '// if pcv_strrApplicableCategories<>"" then			



		grandtotal = grandtotal + total
	next

	calculateVATTotal=grandtotal
	set f=nothing
	set total=nothing 
end function
'////////////////////////////////////////////////////////////////////////////////////////
'// END: VAT CALCULATIONS
'////////////////////////////////////////////////////////////////////////////////////////

Function format_zeros(value, zeros)
  If Len(value) < zeros Then
    Do Until Len(value) = zeros
      value = "0" & value
    Loop
  End If
  format_zeros = value
End Function
%>