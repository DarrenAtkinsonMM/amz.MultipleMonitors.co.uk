<%
'// Query the correct Product Id from the Cart Array
Public Function pcf_ProductIdFromArray(pcCartArray, f)

    If (pcCartArray(f, 43)<>"") Then
        pcv_intProductId = pcCartArray(f, 43)
    Else
        pcv_intProductId = pcCartArray(f, 0)
    End If
    
    pcf_ProductIdFromArray = pcv_intProductId

End Function


'// Query the correct Product Id from the sub-product
Public Function pcf_GetParentId(id)
	Dim rsQ, query

	intParentPrd = id
	
	If statusAPP="1" then
		query="SELECT products.pcProd_ParentPrd FROM products WHERE products.idProduct = " & id
		Set rsQ = server.CreateObject("ADODB.RecordSet")
		Set rsQ=conntemp.execute(query)
		If Not rsQ.Eof Then    
			intParentPrd=rsQ("pcProd_ParentPrd")
			If Not intParentPrd>"0" Then
				intParentPrd = id
			End If
		End If
		Set rsQ=nothing
	end if
    
    pcf_GetParentId = intParentPrd

End Function



'// OPC Re-Calculate Cart Rows
'
' NOTE:  This assumes 30 (item discounts) has already been calculated, which is why we have pre save below.  It also needs the trial amount added.  It also assumes 31 is calculated.
' TO DO:  Remove pcProductList |  Remove cart calculation as no need to redo it here... nothing is new else it would be passed into the method.
'
Public Sub ReCalculateCartRows(ppcCartIndex, pcCartArray)

    strBundleArray = ""
    pSFstrBundleArray=""
    GiftWrapPaymentTotal=0
    
    For f=1 To ppcCartIndex

    
        pcProductList(f,0)=pcCartArray(f,0)
        pcProductList(f,1)=pcCartArray(f,10)
        pcProductList(f,3)=pcCartArray(f,2)
        pcProductList(f,4)=0

			
        If pcCartArray(f,10) = 0 Then

            pBTOValues = calculateBTOValues(pcCartArray(f,16), pcCartArray(f, 0), pcCartArray(f, 2))

            '// START 10th Row - Cross Sell Bundle Discount
            If (pcCartArray(f,27)>0) And (pcCartArray(f,28)>0) Then

                crossSellBundleDiscount = ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2) 
                strBundleArray = strBundleArray & pcCartArray(f,0) & "," & pcCartArray(f,27) & "," & pcCartArray(f,28) & "," & crossSellBundleDiscount & "||"

            End If 
            '// END 10th Row - Cross Sell Bundle Discount

            '// START - GGG Add-on start
            If Session("Cust_GW") = "1" Then

                query="Select pcPE_IDProduct From pcProductsExc Where pcPE_IDProduct=" & pcf_ProductIdFromArray(pcCartArray, f)
                Set rsG = server.CreateObject("ADODB.RecordSet")
                Set rsG = connTemp.execute(query)    
                If rsG.Eof Then
    
                    If Not ((pcCartArray(f,34)="") or (pcCartArray(f,34)="0")) Then
    
                        query = "Select pcGW_OptName,pcGW_OptPrice From pcGWOptions Where pcGW_IDOpt=" & pcCartArray(f,34)
                        Set rsG2 = server.CreateObject("ADODB.RecordSet")
                        Set rsG2 = connTemp.execute(query)    
                        If Not rsG2.eof Then
                            pcv_strOptPrice = rsG2("pcGW_OptPrice")
                            GiftWrapPaymentTotal = GiftWrapPaymentTotal + pcv_strOptPrice
                        end if 
                        Set rsG2 = Nothing
    
                    End If
                    
                End If
                Set rsG = Nothing
  
            End If 
            '// END - GGG Add-on start 
                        
        End If


        If pcv_IsEUMemberState = 0 Then
        
            pcProductList(f, 2) = ( pcCartArray(f,2) * pcCartArray(f,17) ) 
            
            pcCartArray(f, 40) = ( pcCartArray(f,2) * pcCartArray(f,17) ) 
            
        Else
        
            pcProductList(f, 2) = calculateRowTotal( _
                                                    pcCartArray(f,2), _
                                                    pcCartArray(f,17), _
                                                    pcCartArray(f,5), _
                                                    pcCartArray(f,30), _
                                                    pcCartArray(f,31), _
                                                    pcCartArray(f,15), _
                                                    pcCartArray(f,27), _
                                                    pcCartArray(f,28), _
                                                    pcCartArray(f,38), _
                                                    pcv_curTrialAmount, _
                                                    pBTOValues, _
                                                    pcCartArray _
                                                    )
                                                    
            pcCartArray(f, 40) = calculateRowTotal( _
                                                    pcCartArray(f,2), _
                                                    pcCartArray(f,17), _
                                                    pcCartArray(f,5), _
                                                    pcCartArray(f,30), _
                                                    pcCartArray(f,31), _
                                                    pcCartArray(f,15), _
                                                    pcCartArray(f,27), _
                                                    pcCartArray(f,28), _
                                                    pcCartArray(f,38), _
                                                    pcv_curTrialAmount, _
                                                    pBTOValues, _
                                                    pcCartArray _
                                                    )
                                                        
        End If '// If pcv_IsEUMemberState = 0 Then

    Next '// For f=1 To ppcCartIndex
                
    pSFstrBundleArray = strBundleArray
                
End Sub


Public Function calculateRowTotal(quantity, price, optionsprice, itemdiscounts, bto, qtydiscounts, parentindex, bundlediscount, subscription, trialamount, pBTOvalues, pcCartArray)

    '// Step 1: UNIT PRICE  (Quantity * Price  ||  f2 * f17)
    pRowPrice = ccur(quantity * price)
    
    '// Step 2: + OPTIONS PRICE  (Quantity * Options Price  ||  f2 * f5)
    pRowPrice = pRowPrice + ccur(quantity * optionsprice)
    
    '// Step 3: - ITEM DISCOUNTS  (itemdiscounts  ||  f30)
    pRowPrice = pRowPrice - ccur(itemdiscounts)
    
    '// Step 4: + BTO CHARGES  (bto  ||  f31)
    pRowPrice = pRowPrice + ccur(bto)
    
    '// Step 5: + QTY DISCOUNTS  (qtydiscounts  ||  f15)
    pRowPrice = pRowPrice - ccur(qtydiscounts) 
    
    '// SUBSCRIPTION PRICING RESET
    if (subscription) > 0  then

        pSubscriptionID = subscription
        
        %><!--#include file="../pcSBDataInc.asp" --><%
        
        '// If there's a trial set the line total to the trial price
        if pcv_intIsTrial = "1" Then
            pRowPrice = trialamount
        else
        '   // This should be the normal price...
        '    pRowPrice = ccur(quantity * price) - ccur(pBTOvalues) 
        end if 
         
    end if 

    '// Step 6: - CROSS SELL BUNDLE DISCOUNT  (bundlediscount  ||  f28)
    If (pcCartArray(f,27)>0) And (pcCartArray(f,28)>0) Then
        pRowPrice = ( ccur(pRowPrice) + ccur(pcCartArray(cint(parentindex), 40)) ) - ( ( ccur(bundlediscount) + ccur(pcCartArray(cint(parentindex),28) ) ) * quantity )  
    End If
    
    calculateRowTotal = pRowPrice

End Function


Public Function calculateBTOValues(IsBTO, specProduct, btoqty)
	Dim rs, rsQ, query
	Dim pBTOValues
    pBTOValues = 0

    If trim(IsBTO) <> "" Then 

        query="SELECT stringProducts, stringValues, stringCategories,stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(IsBTO)
        Set rs = server.CreateObject("ADODB.RecordSet")
        Set rs = conntemp.execute(query)
        If Not rs.Eof Then
                stringProducts = rs("stringProducts")
                stringValues = rs("stringValues")
                stringCategories = rs("stringCategories")
                ArrProduct = Split(stringProducts, ",")
                ArrValue = Split(stringValues, ",")
                ArrCategory = Split(stringCategories, ",")
                Qstring = rs("stringQuantity")
                ArrQuantity = Split(Qstring,",")
                Pstring = rs("stringPrice")
                ArrPrice = split(Pstring,",")							
        End If
        Set rs = Nothing
    
        If Not ArrProduct(0) = "na" Then
        
            For i=lbound(ArrProduct) To (UBound(ArrProduct)-1)

                tmpMinQty=1
                pcv_intProductId = pcf_GetParentId(ArrProduct(i))

                query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & pcv_intProductId

                Set rsQ = server.CreateObject("ADODB.RecordSet")
                Set rsQ = connTemp.execute(query)
                If Not rsQ.Eof Then
                
                    tmpMinQty = rsQ("pcprod_minimumqty")
                            
                    If IsNull(tmpMinQty) Or tmpMinQty="" Then
                        tmpMinQty=1
                    Else
                    
                        If tmpMinQty="0" Then
                            tmpMinQty=1
                        End If
                        
                    End If
                            
                End If
                Set rsQ = Nothing

                tmpDefault = 0

                query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & specProduct & " AND configProduct=" & pcv_intProductId & " AND cdefault<>0;"
                Set rsQ = server.CreateObject("ADODB.RecordSet")
                Set rsQ = connTemp.execute(query)
                If Not rsQ.Eof Then
                
                    tmpDefault = rsQ("cdefault")
                    
                    If IsNull(tmpDefault) Or tmpDefault="" Then
                        tmpDefault=0
                    Else
                    
                        If tmpDefault<>"0" Then
                            tmpDefault=1
                        End If
                        
                    End If
                    
                End If
                Set rsQ = Nothing
         
                If (ccur(ArrValue(i))<>0) Or ((((clng(ArrQuantity(i))-clng(tmpMinQty)<>0) And (tmpDefault=1)) Or ((clng(ArrQuantity(i))-1<>0) And (tmpDefault=0))) And (ArrPrice(i)<>0)) Then

                    If (clng(ArrQuantity(i))-clng(tmpMinQty))>=0 Then
                    
                        If tmpDefault=1 Then
                            UPrice = (clng(ArrQuantity(i)) - clng(tmpMinQty)) * ArrPrice(i)
                        Else
                            UPrice = (clng(ArrQuantity(i))-1) * ArrPrice(i)
                        End If
                            
                    Else
                        UPrice=0
                    End If
                    
                    pBTOValues = pBTOValues + ccur((ArrValue(i) + UPrice) * btoqty)
                
                End If

            Next '// For i=lbound(ArrProduct) To (UBound(ArrProduct)-1)

        End If '// If Not ArrProduct(0) = "na" Then
                                
    End If '// If trim(IsBTO) <> "" Then 
            
    calculateBTOValues = pBTOValues
        
End Function  


Public Function CalculateCartRows(pcCartIndex, pcCartArray)
 		Dim rs, query
        total = 0
        pRowPrice = 0
        
        for f=1 to pcCartIndex
        
            'pcCartArray(f,40)=""
            
            if pcCartArray(f,10)=0 then

                if trim(pcCartArray(f,27))="" then
                    pcCartArray(f,27)=0
                end if
                
                if trim(pcCartArray(f,28))="" then
                    pcCartArray(f,28)=0
                end if

                '// BTO Values
                pBTOValues = calculateBTOValues(pcCartArray(f, 16), pcCartArray(f, 0), pcCartArray(f, 2))
                pcCartArray(f,41) = pBTOValues

                '// Step 1: UNIT PRICE  (Quantity * Price  ||  f2 * f17)
                pRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))

                if trim(pcCartArray(f,16))<>"" then 

                    query="SELECT stringProducts, stringValues, stringCategories, stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
                    set rs=server.CreateObject("ADODB.RecordSet")
                    set rs=conntemp.execute(query)
 
                    stringProducts=rs("stringProducts")
                    stringValues=rs("stringValues")
                    stringCategories=rs("stringCategories")
                    ArrProduct=Split(stringProducts, ",")
                    ArrValue=Split(stringValues, ",")
                    ArrCategory=Split(stringCategories, ",")
                    Qstring=rs("stringQuantity")
                    ArrQuantity=Split(Qstring,",")
                    Pstring=rs("stringPrice")
                    ArrPrice=split(Pstring,",")
                    set rs=nothing

                End If

                '// Step 2: UNIT PRICE + OPTIONS PRICE  (Quantity * Options Price  ||  f2 * f5)
                pRowPrice = pRowPrice + (ccur(pcCartArray(f,2) * pcCartArray(f,5)))	

                If pcCartArray(f,16) <> "" Then
                       
                    itemsDiscounts = getItemDiscounts(pcCartArray(f,16), pcCartArray(f,2))
        
                    If ItemsDiscounts > 0 Then
        
                        ItemsDiscounts = round(ItemsDiscounts + 0.001, 2)
                        pcCartArray(f,30) = ItemsDiscounts
                        
                        '// Step 3: - ITEM DISCOUNTS  (itemdiscounts  ||  f30)
                        pRowPrice = pRowPrice - ItemsDiscounts 

                    Else
                        pcCartArray(f,30)=0
                    End If
                        
                End If 

                If trim(pcCartArray(f,16)) <> "" Then 
    
                    query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
                    Set rs = server.CreateObject("ADODB.RecordSet")	
                    Set rs = conntemp.execute(query)
                    If Not rs.Eof Then
                            stringCProducts=rs("stringCProducts")
                            stringCValues=rs("stringCValues")
                            stringCCategories=rs("stringCCategories")
                            ArrCProduct=Split(stringCProducts, ",")
                            ArrCValue=Split(stringCValues, ",")
                            ArrCCategory=Split(stringCCategories, ",")
                    End If
                    Set rs = Nothing
      
                    If ArrCProduct(0)<>"na" Then
        
                        '// Step 4: + BTO CHARGES (bto  ||  f31)
                        pRowPrice = pRowPrice + ccur(pcCartArray(f,31)) 
        
        
                    End If '// If ArrCProduct(0)<>"na" Then
                    
                End If  '// If trim(pcCartArray(f,16)) <> "" Then

                If trim(pcCartArray(f,15)) <> "" And trim(pcCartArray(f,15)) > 0 Then
                
                    '// Step 5: + QTY DISCOUNTS  (qtydiscounts  ||  f15)
                    pRowPrice = pRowPrice - ccur(pcCartArray(f,15))
              
                End If 
                pcCartArray(f,42) = pRowPrice

                If (pcCartArray(f,27)>0) And (pcCartArray(f,28)>0) Then
                
                    '// Step 6: - CROSS SELL BUNDLE DISCOUNT  (bundlediscount  ||  f28)
                    pRowPrice = ( ccur(pRowPrice) + ccur(pcCartArray(cint(pcCartArray(f,27)), 40)) ) - ( ( ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28) ) ) * pcCartArray(f,2) ) 
       
                End If


                pcCartArray(f, 40) = pRowPrice
                

                pcv_HaveBundles=0
                if pcCartArray(f,27)=-1 then
                    for mc=1 to pcCartIndex
                        if (pcCartArray(mc,27)<>"") AND (pcCartArray(mc,12)<>"") then
                            if cint(pcCartArray(mc,27))=f AND cint(pcCartArray(mc,12))="0" then
                                pcv_HaveBundles=1
                                exit for
                            end if
                        end if
                    next
                end if
                if (pcCartArray(f,27)>-1) OR (pcv_HaveBundles=0) then
                    total = total + pRowPrice
                end if
        
                if Cint(pcCartArray(f,9))>totalDeliveringTime then
                    totalDeliveringTime=Cint(pcCartArray(f,9))
                end if


            end if '// if pcCartArray(f,10)=0 then
            
        next '// for f=1 to pcCartIndex


        session("pcCartSession")=pcCartArray

    CalculateCartRows = total
       
End Function





'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'// START:  v4 Methods - Need to consolidate and merge with v5 above
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'// Calculates the Cart Subtotal, minus BTO & Discounts
Function calculateCartTotal(pcCartArray, indexCart)
	
    dim f, total
	'SB S
	Dim subInstArr
	'SB E
	total=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then  
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
			   		'// Trial price no discounts 
					total = total + (pcCartArray(f,2) * cdbl(subInstArr(3)))
			 	else
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
				end if  
				'SB E
			else
				'SB S
		    	if subInstArr(2) = "1" Then 
			   		'// Trial price no discounts 
					total = total + (pcCartArray(f,2)* cdbl(subInstArr(3)))
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3))  
				end if  
			end if
				'SB E
			end if
			if (pcCartArray(f,27)>"0") AND (pcCartArray(f,28)>"0") then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if
		end if
	next
	calculateCartTotal=ccur(total)
	set f=nothing
	set total=nothing 
    
End Function


Public Function pcf_CheckNumberRange(n)

    '// This needs reviewed in v5, but was added several years ago to handle out of range integers with euro decimal formatting.
    
    if InStr(Cstr(10/3),",")>0 then
        pcv_SpecialServer=1
        if Instr(n,".")>0 then
            if IsNumeric(n) then
                n=replace(n,".",",")
            end if
        end if
    else
        if scDecSign="," then
            pcv_SpecialServer=0
            if Instr(n,",")>0 then
                if IsNumeric(n) then
                    n=replace(n,",",".")
                end if
            end if
        end if
    end if	
    
    pcf_CheckNumberRange = n 

End Function


Public Function pcf_QtyIDEvent(n)
	Dim rsG, query
    'GGG Add-on start
    gRemain=0
    if (n<>"") and (n<>"0") then
        query="SELECT pcEP_Qty, pcEP_HQty FROM pcEvProducts WHERE pcEP_ID=" & n
        set rsG=connTemp.execute(query)
        if not rsG.eof then
            gRemain = cdbl(rsG("pcEP_Qty")) - cdbl(rsG("pcEP_HQty"))
        end if
        set rsG=nothing
    end if
    'GGG Add-on end
    
    pcf_QtyIDEvent = gRemain
    
End Function 


Public Function pcf_GiftWrapAvailable(n)
	Dim rs1, query
    If session("Cust_GW")="1" Then
        PrdCanGWchecks=0

        query="SELECT pcGC_EOnly FROM pcGC WHERE pcGC_idproduct=" & n
        set rs1=connTemp.execute(query)
        If rs1.eof Then
            query="SELECT pcPE_IDProduct FROM pcProductsExc WHERE pcPE_IDProduct=" & n
            set rs1=connTemp.execute(query)
            If rs1.eof Then
                pcf_GiftWrapAvailable=1
            Else 
                pcf_GiftWrapAvailable=0
            End If
        End If

    End If 
        
End Function 


Function findProduct(pcCartArray, indexCart, pIdProduct)

    Dim f 
    findProduct=Cint(0) 
    If indexCart>0 Then

        For f=1 To indexCart
            If pcCartArray(f,10)=0 And int(pcCartArray(f,0))=int(pIdProduct) Then
                findProduct=-1
            End If
        Next
           
    End If  
    Set f = nothing   
    
End Function


' Cart Weight
function calculateCartWeight(pcCartArray, indexCart)
	dim f, totalWeight
	totalWeight=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then   
			totalWeight=totalWeight + (pcCartArray(f,6)*pcCartArray(f,2))
		end if
	next  
	if cdbl(totalWeight)>0 AND cdbl(totalWeight)<1 then
		totalWeight=1
	end if
	totalWeight=round(totalWeight,0)
	calculateCartWeight=totalWeight  
	set f=nothing
	set totalWeight=nothing
end function

' Cart Weight
function calculateShipWeight(pcCartArray, indexCart)
	dim f, totalWeight
	totalWeight=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND pcCartArray(f,20)=0 then
			totalWeight=totalWeight + (pcCartArray(f,6)*pcCartArray(f,2))
		end if
	next  
	if cdbl(totalWeight)>0 AND cdbl(totalWeight)<1 then
		totalWeight=1
	end if
	totalWeight=round(totalWeight,0) 
	calculateShipWeight=totalWeight  
	set f=nothing
	set totalWeight=nothing
end function


' Cart Surcharge
function calculateTotalProductSurcharge(pcCartArray, indexCart)
	dim f, totalSurcharge
	dim fQty, fSurcharge1, fSurcharge2
	totalSurcharge=0
	'Create a new temporary array to group like ProductIDs regardless of option
	Dim SCArray(100,5)
	G = Cint(0)
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then
			if G>0 then
				var_update = 0
				for h = 0 to G
					if SCArray(h,0)=pcCartArray(f,0) then 
						SCArray(h,1)=Cint(SCArray(h,1))+Cint(pcCartArray(f,2))
						var_update = 1
					end if
				next
				if var_update = 0 then
					SCArray(G,0)=pcCartArray(f,0)
					SCArray(G,1)=pcCartArray(f,2)
					SCArray(G,2)=pcCartArray(f,36)
					SCArray(G,3)=pcCartArray(f,37)
				end if
			else
				SCArray(G,0)=pcCartArray(f,0)
				SCArray(G,1)=pcCartArray(f,2)
				SCArray(G,2)=pcCartArray(f,36)
				SCArray(G,3)=pcCartArray(f,37)
			end if
			G = G + 1
		end if
	Next
	
	for t=0 to G
		fQty = SCArray(t,1) 'quantity ordered of product
		fSurcharge1 = SCArray(t,2) 'Initial Surcharge
		fSurcharge2 = SCArray(t,3) 'Additional Surcharge
        If fSurcharge1="" Then
            fSurcharge1 = 0
        End If
        If fSurcharge2="" Then
            fSurcharge2 = 0
        End If
        If fSurcharge1>0 Then
		    totalSurcharge=ccur(totalSurcharge) + ccur(fSurcharge1)
        End If
		if fQty > 1 AND fSurcharge2>0 then
			totalSurcharge=ccur(totalSurcharge) + (ccur(SCArray(t,3))*(CLng(fQty)-1))
		end if
	next 
	calculateTotalProductSurcharge=totalSurcharge  
	set f=nothing
	set totalSurcharge=nothing
end function


' Cart Product Quantity
function calculateCartQuantity(pcCartArray, indexCart)
	dim f, totalQuantity
	totalQuantity=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then   
			totalQuantity=totalQuantity + pcCartArray(f,2)
		end if
	next  
	calculateCartQuantity=totalQuantity  
	set f=nothing
	set totalQuantity=nothing
end function


' Cart Product Quantity
function calculateCartShipQuantity(pcCartArray, indexCart)
	dim f, totalShipQuantity
	totalShipQuantity=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND pcCartArray(f,20)=0 then   
			totalShipQuantity=totalShipQuantity + pcCartArray(f,2)
		end if
	next  
	calculateCartShipQuantity=totalShipQuantity  
	set f=nothing
	set totalShipQuantity=nothing
end function



'////////////////////////////////////////////////////////////////////////////////////////
'// START: (Re)check stock levels of current cart
'////////////////////////////////////////////////////////////////////////////////////////

function checkCartStockLevels(pcCartArray, indexCart, aryBads)
	Dim intCCSLindex, intCCSLcounter, intCCSLfound
	Dim strCCSLSQL, strCCSLWarn
	Dim aryCCSLitems, aryCCSLids
	Dim objCCSLrs

    ' Initialize the 'bads' array (an array tied to the index of pcCartArray, indicating which lines are bad)
	ReDim aryBads(indexCart)

    ' If cart configuration allows purchasing out-of-stock items, we're all done here
	If scOutofstockpurchase=0 Then Exit Function

	ReDim aryCCSLitems(1,0)

    intCCSLcounter = -1
	strCCSLwarn = ""

	for intCCSLIndex=1 to indexCart

	    intCCSLfound = -1								' Init to -1 indicating 'not found'

		if pcCartArray(intCCSLindex,10)=0 then			' If this product has not been deleted from cart, then
		      intCCSLcounter = intCCSLcounter + 1
			  ReDim preserve aryCCSLitems(1, intCCSLcounter)
		      aryCCSLitems(0,intCCSLcounter) = pcCartArray(intCCSLindex, 0)
		      aryCCSLitems(1,intCCSLcounter) = pcCartArray(intCCSLindex, 2)		   
		End If
		
	Next
	
	' If there were no viable items found in the cart array (e.g. all were deleted?) then exit now
	If intCCSLcounter = -1 Then Exit Function

	' Unspool/Serialize the idProducts in the array
	ReDim aryCCSLids(intCCSLcounter)
	For intCCSLindex=0 to intCCSLcounter
	   aryCCSLids(intCCSLIndex) = aryCCSLitems(0, intCCSLindex)
	Next
	strCCSLSQL = "SELECT idproduct, stock, Description, noStock, pcProd_BackOrder FROM products WHERE idproduct in (" & Join(aryCCSLids, ",") & ") and noStock=0 and pcProd_BackOrder=0"
	Set objCCSLrs = connTemp.execute(strCCSLSQL)
	If objCCSLrs.eof Then
	   objCCSLrs.close
	   Set objCCSLrs = Nothing
	   Exit Function
	Else
	   aryCCSLrecs = objCCSLrs.getrows()
	End If
	objCCSLrs.close
	Set objCCSLrs = Nothing

	ReDim aryBads(UBound(aryCCSLitems,2))
	For intCCSLindex=0 to UBound(aryCCSLitems,2)

       aryBads(intCCSLindex) = 0							' Flag this line as OK (for now)

	   For intCCSLindex2 = 0 To UBound(aryCCSLrecs, 2)

	      If CLng(aryCCSLitems(0, intCCSLindex)) = CLng(aryCCSLrecs(0, intCCSLindex2)) Then
		     If CLng(aryCCSLitems(1, intCCSLindex)) > CLng(aryCCSLrecs(1, intCCSLindex2)) Then ' Is overstock!
			    strCCSLwarn = strCCSLwarn & ("<li>" & aryCCSLrecs(2, intCCSLindex2) & " (we currently have " & aryCCSLrecs(1, intCCSLindex2) & " in stock)</li>")

			    aryBads(intCCSLindex) = -1				' Flag this line as insufficient stock level

			 End If
		     Exit For
		     
	      End If

	   next
	   
	Next

	If Len(strCCSLwarn)>0 Then checkCartStockLevels = dictLanguage.Item(Session("language")&"__alert_14") & "<ul>" & strCCSLwarn & "</ul>"

end function


' Check session lost
function checkSessionLost(pcCartArray, pcCartIndex) 
	if pcCartIndex="" then
		' session is lost, initialize all variables
		Session.Timeout=25
		Session("idCustomer")=Cint(0)
		Session("language")=Cstr("english")
		Session("pcCartIndex")=Cint(0)
		ReDim pcCartArray(50, 18)
		Session("pcCartSession")=pcCartArray      
		checkSessionLost=1
	else
		checkSessionLost=0
	end if
end function


'////////////////////////////////////////////////////////////////////////////////////////
'// START: get customer ID from order ID
'////////////////////////////////////////////////////////////////////////////////////////
function getCustIDfromOrder(idOrder)
	Dim rsCFO, query

	query="SELECT idcustomer FROM orders WHERE idOrder="&idOrder
	set rsCFO=server.CreateObject("ADODB.RecordSet")
	set rsCFO=connTemp.execute(query)
	if rsCFO.eof then
		getCustIDfromOrder=0
	else
		getCustIDfromOrder=rsCFO("idcustomer")
	end if
	set rsCFO=nothing  
end function
'////////////////////////////////////////////////////////////////////////////////////////
'// END: get customer ID from order ID
'////////////////////////////////////////////////////////////////////////////////////////


' count cart Rows
function countCartRows(pcCartArray, indexCart)
 
 dim cont, f
 
 cont=Cint(0)
 if indexCart>0 then
  for f=1 to indexCart
    if pcCartArray(f,10)=0 then
     cont=cont+1
    end if
  next
 else
  cont=0
 end if
 
 countCartRows=cont
 set f=nothing 
 set cont=nothing
 
end function


Public Sub pcs_CartStackTracking
	' ------------------------------------------------------
	' START - CartStack Tracking	// viewcart.asp and onepagecheckout.asp
	' ------------------------------------------------------
	subtotal = CalculateCartRows(session("pcCartIndex"), session("pcCartSession"))
	
	Response.Write vbCrlf & _
		"<script src=""https://api.cartstack.com/js/cs.js"" type=""text/javascript""></script>" & vbCrlf & vbCrlf & _
		"<script language=""javascript"">" & vbCrlf & _
			"var _cartstack = _cartstack || [];" & vbCrlf & _
			"_cartstack.push(['setSiteID', '" & scCartStack_SiteId & "']);" & vbCrlf & _
			"_cartstack.push(['setAPI', 'tracking']);" & vbCrlf & _
			"_cartstack.push(['setCartTotal', '" & subtotal & "']);" & vbCrlf
	
	pcs_CartStackItem
	
	Response.Write "</script>" & vbCrlf & vbCrlf
	' ------------------------------------------------------
	' END - CartStack Tracking
	' ------------------------------------------------------
End Sub


Public Sub pcs_CartStackConfirmation
	' ------------------------------------------------------
	' START - CartStack Confirmation	// ordercomplete.asp
	' ------------------------------------------------------
	Response.Write vbCrlf & _
		"<script language=""javascript"">" & vbCrlf & _
			"var _cartstack = _cartstack || [];" & vbCrlf & _
			"_cartstack.push(['setSiteID', '" & scCartStack_SiteId & "']);" & vbCrlf & _
			"_cartstack.push(['setAPI', 'confirmation']);" & vbCrlf & _
		"</script>" & vbCrlf & vbCrlf & _
		"<script src=""https://api.cartstack.com/js/cartstack.js"" type=""text/javascript""></script>" & vbCrlf & vbCrlf
	' ------------------------------------------------------
	' START - CartStack Confirmation
	' ------------------------------------------------------
End Sub


Public Sub pcs_CartStackItem
Dim rsImg,rsImgParent,query
	' ------------------------------------------------------
	' START - CartStack Cart
	' ------------------------------------------------------
	pcCartIndex = session("pcCartIndex")
	pcCartArray = session("pcCartSession")
	
	if scSSL = "1" then
		pcv_storeURL = scSslURL &"/"& scPcFolder
	else
		pcv_storeURL = scStoreURL &"/"& scPcFolder
	end if
	
	if scPcFolder <> ""	then
		pcv_storeURL = pcv_storeURL & "/"
	end if
	
	For f = 1 To pcCartIndex
	
		query="SELECT sku, smallImageUrl, imageUrl FROM products WHERE idProduct=" & pcCartArray(f,0)
		set rsImg=Server.CreateObject("ADODB.Recordset")
		set rsImg=conntemp.execute(query)
		If Not rsImg.Eof Then
			pcvStrSku = rsImg("sku")
			pcvStrSmallImage = rsImg("smallImageUrl")
		End If
		
		'APP-S
		if Instr(Ucase(scSubVersion),"A")>0 then
			pcv_HaveApparel=1
		else
			pcv_HaveApparel=0
		end if
		'APP-E
		
		'APP-S
		if (pcCartArray(f,32)<>"") then
			pcvStrSmallImage = trim(rsImg("imageUrl"))
		end if
		
		if (trim(pcvStrSmallImage) = "") OR (pcvStrSmallImage = "no_image.gif") OR (IsNull(pcvStrSmallImage)) then
			if (pcCartArray(f,32)<>"") then
				pcv_tmpPPrd=split(pcCartArray(f,32),"$$")
				pcv_tmpPPrdTemp=pcv_tmpPPrd(ubound(pcv_tmpPPrd))
				query="SELECT smallImageUrl FROM products WHERE idProduct=" & pcv_tmpPPrdTemp
				set rsImgParent=Server.CreateObject("ADODB.Recordset")
				set rsImgParent=conntemp.execute(query)
				If Not rsImgParent.Eof Then
					pcvStrSmallImage=rsImgParent("smallImageUrl")
				End If
				set rsImgParent=nothing
			end if
		end if
		'APP-E
			  
		'APP-S
		if (pcCartArray(f,32)<>"") then
			pcv_tmpPPrd = split(pcCartArray(f,32),"$$")
			pIdProduct = pcv_tmpPPrd(ubound(pcv_tmpPPrd))
		else 
			pIdProduct = pcCartArray(f,0)
		end if
		'APP-E	
				
		pSKU = replace(pcCartArray(f,7),"&quot;","""")
		pName = replace(pcCartArray(f,1),"&quot;","""")
		pUnitPrice = pcCartArray(f,42)
		pQuantity = cdbl(pcCartArray(f,2))
		pImageUrl = pcvStrSmallImage
		If pImageUrl = "" Then
			pImageUrl = "no_image.gif"
		End If
		pDesc = ""
		
		pName = replace(pName,"'","\'")
		pDesc = replace(pDesc,"'","\'")
		
		Response.Write vbCrlf & _
		"_cartstack.push(['setCartItem', {" & vbCrlf & _
			"'quantity':'" & pQuantity & "'," & vbCrlf & _
			"'productID':'" & pIdProduct & "'," & vbCrlf & _
			"'productName':'" & pName & "'," & vbCrlf & _
			"'productDescription':'" & pDesc & "'," & vbCrlf & _
			"'productURL':'" & pcv_storeURL & "pc/viewPrd.asp?idproduct=" & pIdProduct & "'," & vbCrlf & _
			"'productImageURL':'" & pcv_storeURL & "pc/catalog/" & pImageUrl & "'," & vbCrlf & _
			"'productPrice':'" & pUnitPrice & "'" & vbCrlf & _
		"}]);" & vbCrlf
		
	Next
	' ------------------------------------------------------
	' END - CartStack
	' ------------------------------------------------------
End Sub
%>