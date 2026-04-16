<%
Function getAutoDiscountCodes(HavePrdsOnSale, displayDiscountCode)

    If HavePrdsOnSale=0 Then

        pcStrAutoDiscCode = ""

        query="SELECT discountcode FROM discounts WHERE pcDisc_Auto=1 AND active=-1 ORDER BY percentagetodiscount DESC,pricetodiscount DESC;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        If Not rs.eof Then
        
            Do Until rs.eof
                
                pcIntADCnt = pcIntADCnt + 1
                
                If pcIntADCnt = 1 Then
                    pcStrAutoDiscCode = pcStrAutoDiscCode & rs("discountcode")
                Else
                    pcStrAutoDiscCode = pcStrAutoDiscCode & "," & rs("discountcode")
                End If

                rs.movenext
            Loop

        End If
        set rs=nothing

        If len(displayDiscountCode)>0 Then
            displayDiscountCode = pcStrAutoDiscCode & "," & trim(displayDiscountCode)
        Else
            displayDiscountCode = pcStrAutoDiscCode	
        End If        
        
    END IF

    getAutoDiscountCodes = displayDiscountCode

End Function




Public Function getDiscountCodes(rtype, savdiscountcode, discountcode, TurnOffDiscountCodesWhenHasSale, scHideDiscField, pcCartArray, ppcCartIndex)

    '// Request Discount Code from Input to recalculate
    If rtype<>"1" Then
        displayDiscountCode = savdiscountcode
    Else
        displayDiscountCode = discountcode
    End If
	
    If displayDiscountCode<>"" Then
        displayDiscountCode = replace(displayDiscountCode, ", ", ",")
        displayDiscountCode = replace(displayDiscountCode, " ,", ",")
    End If
            
    '// Set AutoDiscount flag to 0
    pcIntADCnt=0
            
    If (TurnOffDiscountCodesWhenHasSale="1") And (Not scHideDiscField="1") Then

        Dim tmpPrdList
        tmpPrdList=""

        For f=1 To ppcCartIndex
                    
            If pcCartArray(f,10)=0 Then

                If tmpPrdList<>"" Then
                    tmpPrdList = tmpPrdList & ","
                End If
                tmpPrdList = tmpPrdList & pcCartArray(f,0)
                
            End If
                    
        Next
        
        If tmpPrdList="" Then
            tmpPrdList="0"
        End If
        
        tmpPrdList = "(" & tmpPrdList & ")"

        query = "SELECT idProduct FROM Products WHERE idProduct IN " & tmpPrdList & " AND pcSC_ID>0;"
        set rsQ = connTemp.execute(query)
        If Not rsQ.eof Then
            HavePrdsOnSale = 1
        End If
        set rsQ = Nothing
     
    End If
	    
    getDiscountCodes = displayDiscountCode

End Function




Public Function checkInvalidCodes(intCodeCntO, DiscountCodeArryO, intCodeCnt)
                    
    Dim FoundInArr
    For ik=0 To intCodeCntO

        If trim(DiscountCodeArryO(ik))<>"" Then

            FoundInArr = 0
            For i = 0 To intCodeCnt
                If trim(ucase(DiscountCodeArryO(ik)))=trim(ucase(DiscountCodeArry(i))) Then
                    FoundInArr = 1
                End If
            Next

            If FoundInArr=0 Then
                pcGlobalDiscError = pcGlobalDiscError & "<li>" & dictLanguage.Item(Session("language")&"_orderverify_4") & " (<b>" & DiscountCodeArryO(ik) & "</b>)</li>"
            End If
        
        END IF

    Next
    
    checkInvalidCodes = pcGlobalDiscError
                    
End Function




Public Sub sortDiscountCodes()
                    
    If (pDiscountCode<>"") And (InStr(pDiscountCode,",")>0) Then

        tmpDC=""
        For i=0 To intCodeCnt

            If trim(DiscountCodeArry(i))<>"" Then

                If tmpDC <> "" Then
                    tmpDC = tmpDC & ","
                End If
				tmpDC = tmpDC & "'" & trim(DiscountCodeArry(i)) & "'"
                
            End If

        Next '// For i=0 To intCodeCnt

        If tmpDC<>"" Then

            query = "SELECT discountcode FROM discounts WHERE discountcode IN (" & tmpDC & ") ORDER BY pcDisc_Auto DESC,pcSeparate DESC;"
            set rsQ=connTemp.execute(query)
            If Not rsQ.eof Then

                pDiscountCode=""
                tmpDCArr = rsQ.getRows()
                intCountD = ubound(tmpDCArr, 2)
                
                For i=0 To intCountD

                    If pDiscountCode <> "" Then
                        pDiscountCode = pDiscountCode & ","
                    End If
                    pDiscountCode = pDiscountCode & tmpDCArr(0,i)

                Next

                DiscountCodeArry = Split(pDiscountCode,",")
                intCodeCnt = ubound(DiscountCodeArry)

            End If
			Set rsQ = Nothing

            If pDiscountCode="" Then
                intCodeCnt = -1
            End If

        End If '// If tmpDC<>"" Then

    End If '// If (pDiscountCode<>"") And (InStr(pDiscountCode,",")>0) Then
                    
End Sub 




Public Function getUsedDiscountCodes(UsedDiscountCodes, pTempDiscCode)
    
    '// Check if discount code has already been used for this store
    
    UsedDiscountCodes = ""
    
    intDiscMatchFound = 0	
                            
    If UsedDiscountCodes <> "" Then
    
        UsedDiscountCodeArry = split(UsedDiscountCodes,",")
        
        For t=0 to (ubound(UsedDiscountCodeArry)-1)
            If pTempDiscCode = UsedDiscountCodeArry(t) Then
                intDiscMatchFound = 1
                pDiscountError = dictLanguage.Item(Session("language")&"_orderverify_40") 
            End If
        Next
        
    End If
    
    if intDiscMatchFound=0 then
        UsedDiscountCodes = UsedDiscountCodes & pTempDiscCode & ","
    end if
    
    getUsedDiscountCodes = UsedDiscountCodes 
    
End Function




Public Sub separateDiscountsAndGiftCodes()

    pTempGC=""
    pTempDiscountCode=""
    pCodeTotal=""
    
    If pDiscountCode = "" Then
        noCode = "1"
    End If
				
    If noCode="" Then

        DiscountCodeArry = Split(pDiscountCode,",")
        
        For i=0 To ubound(DiscountCodeArry)

            If DiscountCodeArry(i) <> "" Then
		
                query="SELECT pcGCOrdered.pcGO_ExpDate, pcGCOrdered.pcGO_Amount, pcGCOrdered.pcGO_Status, products.Description FROM pcGCOrdered, products WHERE pcGCOrdered.pcGO_GcCode='" & DiscountCodeArry(i) & "' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"
				Set rsQ = server.CreateObject("ADODB.RecordSet")
				Set rsQ = conntemp.execute(query)
				If Not rsQ.eof Then
        
                    If pTempGC<>"" Then
					    pTempGC = pTempGC & ","
					End If
					pTempGC = pTempGC & DiscountCodeArry(i)
                    
				Else

                    If pTempDiscountCode <> "" Then
                        pTempDiscountCode = pTempDiscountCode & ","
					End If
                    
					pTempDiscountCode = pTempDiscountCode & DiscountCodeArry(i)
				End If
				Set rsQ = nothing
							
            End If '// If DiscountCodeArry(i) <> "" Then

        Next '// For i=0 To ubound(DiscountCodeArry)
					
		pDiscountCode = pTempDiscountCode

		pCodeTotal = pTempDiscountCode
        
		If pTempGC <> "" Then
		
            If pCodeTotal <> "" Then
			    pCodeTotal = pCodeTotal & ","
			End If
			pCodeTotal = pCodeTotal & pTempGC
            
		End If
						
		If displayDiscountCode <> pCodeTotal Then
		    displayDiscountCode = pCodeTotal
		    session("DCODE") = displayDiscountCode
		End If

    End If '// If noCode="" Then

End Sub




Public Function calculateCategoryDiscountTotal(ppcCartIndex, pcCartArray)

    CatDiscTotal=0

	query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	
	Dim tmpPrdIDs, tmpPrdIDs2
	
	Do While not rs.eof

        CatSubQty=0
		CatSubTotal=0
		CatSubDiscount=0
        CanNotRun=0
        
		ApplicableCategoryID = rs("IDCat")
		IDCat=rs("IDCat")
					
        query="SELECT categories_products.idcategory FROM categories_products INNER JOIN pcPrdPromotions ON categories_products.idproduct=pcPrdPromotions.idproduct WHERE categories_products.idcategory=" & IDCat & ";"
		Set rsQ = server.CreateObject("ADODB.RecordSet")
        Set rsQ = connTemp.execute(query)
		If Not rsQ.eof Then
		    CanNotRun=1
		End If
		Set rsQ=nothing
					
        If CanNotRun=0 Then
			
			tmpPrdIDs2 = tmpPrdIDs
			
            For f=1 To ppcCartIndex

                If (pcCartArray(f, 10) = 0) AND Instr(tmpPrdIDs,pcf_ProductIdFromArray(pcCartArray, f) & ",") = 0 Then 
				'If (pcProductList(f,1)=0) And (pcProductList(f,4)=0) Then 
 		
                    query="select idproduct from categories_products where idcategory=" & IDCat & " and idproduct=" & pcf_ProductIdFromArray(pcCartArray, f)
					Set rstemp=server.CreateObject("ADODB.RecordSet")
					Set rstemp=connTemp.execute(query)							
                    If Not rstemp.eof Then
						tmpPrdIDs=tmpPrdIDs & pcf_ProductIdFromArray(pcCartArray, f) & ","
					    CatSubQty = CatSubQty + pcCartArray(f, 2) '// pcProductList(f,3)
						CatSubTotal = CatSubTotal + pcCartArray(f, 42) '// pcProductList(f,2)
						'pcCartArray(f, 4) = 1 '// The 0 is being reset to 1... this will need a new location in our main array.			
					End If
					Set rstemp=nothing
								
                End If
							
            Next '// For f=1 To ppcCartIndex
			
            If CatSubQty>0 Then
	
                query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & IDCat & " AND pcCD_quantityFrom<=" & CatSubQty & " AND pcCD_quantityUntil>=" & CatSubQty
                Set rstemp = server.CreateObject("ADODB.RecordSet")
				Set rstemp = conntemp.execute(query)
				If Not rstemp.eof Then

                    '// There are quantity discounts defined for that quantity 
                    pDiscountPerUnit = rstemp("pcCD_discountPerUnit")
                    pDiscountPerWUnit = rstemp("pcCD_discountPerWUnit")
                    pPercentage = rstemp("pcCD_percentage")
                    pbaseproductonly = rstemp("pcCD_baseproductonly")
					
                    If session("customerType") <> 1 Then  '// Customer is a normal user

                        If pPercentage = "0" Then 
                            CatSubDiscount = pDiscountPerUnit * CatSubQty
						Else
                            CatSubDiscount = (pDiscountPerUnit / 100) * CatSubTotal
						End If

                    Else  '// Customer is a wholesale customer

                        If pPercentage="0" Then 
                            CatSubDiscount = pDiscountPerWUnit * CatSubQty
						Else
                            CatSubDiscount = (pDiscountPerWUnit / 100) * CatSubTotal
                        End If

                    End If
				
				Else
				
					tmpPrdIDs = tmpPrdIDs2

                End If '// If Not rstemp.eof Then					
				Set rstemp = nothing		
											
            End If '// if CatSubQty>0 then
	
            CatDiscTotal = CatDiscTotal + CatSubDiscount
					
        END IF 'CanNotRun
					
        rs.MoveNext
    Loop '// Do While not rs.eof
	Set rs = Nothing
    
    '// Round the Category Discount to two decimals
	If CatDiscTotal<>"" And isNumeric(CatDiscTotal) Then
        CatDiscTotal = RoundTo(CatDiscTotal,.01)
    End If
                
    calculateCategoryDiscountTotal = CatDiscTotal
				
End Function




Public Function getItemDiscounts(id, BTOqty)
Dim itemsDiscounts,rs,query
    itemsDiscounts = 0

    query="SELECT stringProducts, stringValues, stringCategOries, stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(id)
    set rs=server.CreateObject("ADoDB.RecOrdSet")
    set rs=conntemp.execute(query)
    If Not rs.Eof Then
        stringProducts=rs("stringProducts")
        stringValues=rs("stringValues")
        stringCategOries=rs("stringCategOries")
        ArrProduct=Split(stringProducts, ",")
        ArrValue=Split(stringValues, ",")
        ArrCategory=Split(stringCategOries, ",")
        Qstring=rs("stringQuantity")
        ArrQuantity=Split(Qstring,",")
        Pstring=rs("stringPrice")
        ArrPrice=split(Pstring,",")    
    End If
    set rs=Nothing 
               
    For i = lbound(ArrProduct) To (UBound(ArrProduct)-1)

        intParentPrd = pcf_GetParentId(ArrProduct(i))
		
		query="SELECT quantityFrom,quantityUntil,discountperUnit,percentage,discountperWUnit FROM discountsPerQuantity WHERE IDProduct=" & ArrProduct(i) & " OR IdProduct=" & intParentPrd & ";"
		'APP-E
        Set rs = server.CreateObject("ADODB.RecordSet")
        Set rs = connTemp.execute(query)
        
        TempDiscount = 0

        Do While Not rs.Eof

            QFrom = rs("quantityFrom")
            QTo = rs("quantityUntil")
            DUnit = rs("discountperUnit")
            QPercent = rs("percentage")
            DWUnit = rs("discountperWUnit")
            
            If (DWUnit=0) And (DUnit>0) Then
                DWUnit = DUnit
            End If
            
            TempD1=0

            If (clng(ArrQuantity(i)*BTOqty)>=clng(QFrom)) And (clng(ArrQuantity(i)*BTOqty)<=clng(QTo)) Then
                               
                                if QPercent="-1" then
                                    if session("customerType")=1 then
                                        TempD1=ArrQuantity(i)*BTOqty*ArrPrice(i)*0.01*DWUnit
                                    else
                                        TempD1=ArrQuantity(i)*BTOqty*ArrPrice(i)*0.01*DUnit
                                    end if
                                else
                                    if session("customerType")=1 then
                                        TempD1=ArrQuantity(i)*BTOqty*DWUnit
                                    else
                                        TempD1=ArrQuantity(i)*BTOqty*DUnit
                                    end if
                                end if

            End If

            TempDiscount = TempDiscount + TempD1

            rs.movenext
        Loop
        Set rs = Nothing
    
        itemsDiscounts = ItemsDiscounts + TempDiscount

    Next '// For i = lbound(ArrProduct) To (UBound(ArrProduct) - 1)

    getItemDiscounts = itemsDiscounts
    
End Function 





'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'// START:  v4 Methods - Need to consolidate and merge with v5 above
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function calculateCategoryDiscounts(pcCartArray, indexCart)

	Dim TmpProList(100,5)
	for f=1 to indexCart
		TmpProList(f,0)=pcCartArray(f,0)
		TmpProList(f,1)=pcCartArray(f,10)
		TmpProList(f,3)=pcCartArray(f,2)
		TmpProList(f,4)=0
		if pcCartArray(f,10)=0 then
			'Get RowPrice
			pRowPrice=0
			pRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))
			pRowPrice=pRowPrice + ccur(pcCartArray(f,2) * pcCartArray(f,5))	
			if trim(pcCartArray(f,30))<>"" AND trim(pcCartArray(f,30))>"0" then
				pRowPrice=pRowPrice-pcCartArray(f,30)
			end if
			if trim(pcCartArray(f,31))<>"" AND trim(pcCartArray(f,31))>"0" then
				pRowPrice=pRowPrice+ccur(pcCartArray(f,31))
			end if
			if trim(pcCartArray(f,15))<>"" AND trim(pcCartArray(f,15))>"0" then
				pRowPrice=pRowPrice-ccur(pcCartArray(f,15))
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then 
				pRowPrice = ( ccur(pRowPrice) + ccur(TmpProList(cint(pcCartArray(f,27)),2)) ) - ( ( ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28) ) ) * pcCartArray(f,2) )
			end if
			TmpProList(f,2)=pRowPrice
		end if
	next
			
	' ------------------------------------------------------
	' START - Calculate category-based quantity discounts
	' ------------------------------------------------------
	CatDiscTotal=0
	
	query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
	set rsCDObj=server.CreateObject("ADODB.RecordSet")
	set rsCDObj=conntemp.execute(query)

	Do While not rsCDObj.eof
		CatSubQty=0
		CatSubTotal=0
		CatSubDiscount=0
	
		For f=1 to indexCart
			if (TmpProList(f,1)=0) and (TmpProList(f,4)=0) then 

				query="select idproduct from categories_products where idcategory=" & rsCDObj("IDCat") & " and idproduct=" & pcf_ProductIdFromArray(pcCartArray, f)
				set rsCDObjtemp=server.CreateObject("ADODB.RecordSet")
				set rsCDObjtemp=connTemp.execute(query)
				
				if not rsCDObjtemp.eof then
					CatSubQty=CatSubQty+TmpProList(f,3)
					CatSubTotal=CatSubTotal+TmpProList(f,2)
					TmpProList(f,4)=1
				end if
				set rsCDObjtemp=nothing
                
			end if
		Next

		if CatSubQty>0 then
			query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rsCDObj("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
			set rsCDObjtemp=server.CreateObject("ADODB.RecordSet")
			set rsCDObjtemp=conntemp.execute(query)
	
			if not rsCDObjtemp.eof then
				' there are quantity discounts defined for that quantity 
				pDiscountPerUnit=rsCDObjtemp("pcCD_discountPerUnit")
				pDiscountPerWUnit=rsCDObjtemp("pcCD_discountPerWUnit")
				pPercentage=rsCDObjtemp("pcCD_percentage")
				pbaseproductonly=rsCDObjtemp("pcCD_baseproductonly")
				set rsCDObjtemp=nothing
				
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
		rsCDObj.MoveNext
		loop
		set rsCDObj=nothing				
		'// Round the Category Discount to two decimals
		if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
			CatDiscTotal = Round(CatDiscTotal,2)
		end if
		' ------------------------------------------------------
		' END - Calculate category-based quantity discounts
		' ------------------------------------------------------
	calculateCategoryDiscounts=CatDiscTotal
end function 
%>