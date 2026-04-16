<%
Public Function pcf_GetAvalaraResponse(postURL, postdata)

	'response.Clear()
    'response.Write(postdata)  
    'response.End()
	
	if ptaxAvalaraLog = 1 then
		call pcs_logEventUTF8("avalara.log", "REQUEST : " & postdata)
	end if
	
	Set srvAvalaraXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP" & scXML)
	srvAvalaraXmlHttp.open "POST", postURL, False
	srvAvalaraXmlHttp.SetRequestHeader "Content-Type", "text/xml"
	srvAvalaraXmlHttp.SetRequestHeader "Content-Length", Len(postdata)
	srvAvalaraXmlHttp.SetRequestHeader "Authorization", "Basic " & Base64_Encode(ptaxAvalaraAccount & ":" & ptaxAvalaraLicense)
	srvAvalaraXmlHttp.Send postdata
	
	xmlResponse = srvAvalaraXmlHttp.responseText
	Set srvAvalaraXmlHttp = nothing

	if ptaxAvalaraLog = 1 then
		call pcs_logEventUTF8("avalara.log", "RESPONSE : " & xmlResponse)
	end if
	
	'response.Clear()
    'response.Write(xmlResponse)
    'response.End()
	
	if err.number<>0 then
        err.number=0
        err.description=""
    end if
	
	pcf_GetAvalaraResponse = xmlResponse
	
End Function


Public Function pcf_avaDiscounts(i, CatDiscTotal, discountTotal, TotalPromotions)
    Dim postdata
    postdata = ""
    
    '// Category Discounts
    'If CatDiscTotal > 0 Then
    '	postdata = postdata & "<Line>" & vbCrlf
    '	postdata = postdata & "<LineNo>" & i & "-CD</LineNo>" & vbCrlf
    '	postdata = postdata & "<ItemCode>DISCOUNT</ItemCode>" & vbCrlf
    '	postdata = postdata & "<Description>Category Discounts</Description>" & vbCrlf
    '	postdata = postdata & "<Qty>1</Qty>" & vbCrlf
    '	postdata = postdata & "<Amount>-"& CatDiscTotal &"</Amount>" & vbCrlf
    '	postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
    '	postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
    '	postdata = postdata & "</Line>" & vbCrlf
    'End If
    
    '// Total Discounts
    'If discountTotal > 0 Then
    '	postdata = postdata & "<Line>" & vbCrlf
    '	postdata = postdata & "<LineNo>" & i & "-DT</LineNo>" & vbCrlf
    '	postdata = postdata & "<ItemCode>DISCOUNT</ItemCode>" & vbCrlf
    '	postdata = postdata & "<Description>Discounts</Description>" & vbCrlf
    '	postdata = postdata & "<Qty>1</Qty>" & vbCrlf
    '	postdata = postdata & "<Amount>-"& discountTotal &"</Amount>" & vbCrlf
    '	postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
    '	postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
    '	postdata = postdata & "</Line>" & vbCrlf
    'End If
    
    '// Total Promotions
    'If TotalPromotions > 0 Then
    '	postdata = postdata & "<Line>" & vbCrlf
    '	postdata = postdata & "<LineNo>" & i & "-PT</LineNo>" & vbCrlf
    '	postdata = postdata & "<ItemCode>PROMOTIONS</ItemCode>" & vbCrlf
    '	postdata = postdata & "<Description>Promotions</Description>" & vbCrlf
    '	postdata = postdata & "<Qty>1</Qty>" & vbCrlf
    '	postdata = postdata & "<Amount>-"& TotalPromotions &"</Amount>" & vbCrlf
    '	postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
    '	postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
    '	postdata = postdata & "</Line>" & vbCrlf
    'End If
    
    postdata = postdata & "<Discount>" & (CatDiscTotal + discountTotal + TotalPromotions) & "</Discount>" & vbCrlf
                
    pcf_avaDiscounts = postdata
End Function


Public Function pcf_avaServiceHandlingFee(i, pcDblServiceHandlingFee)
    Dim postdata
    postdata = ""
    If pcDblServiceHandlingFee > 0 Then
        postdata = postdata & "<Line>" & vbCrlf
        postdata = postdata & "<LineNo>" & i & "-HF</LineNo>" & vbCrlf
        postdata = postdata & "<ItemCode>HANDLING</ItemCode>" & vbCrlf
        postdata = postdata & "<Description>Handling Fee</Description>" & vbCrlf
        postdata = postdata & "<Qty>1</Qty>" & vbCrlf
        postdata = postdata & "<Amount>"& pcDblServiceHandlingFee &"</Amount>" & vbCrlf
        postdata = postdata & "<Discounted>false</Discounted>" & vbCrlf
        postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
        postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
        If ptaxAvalaraHandlingCode <> "" Then
            postdata = postdata & "<TaxCode>" & ptaxAvalaraHandlingCode & "</TaxCode>" & vbCrlf
        End If
        postdata = postdata & "</Line>" & vbCrlf
    End If
    pcf_avaServiceHandlingFee = postdata
End Function


Public Function pcf_avaShipmentTotal(i, pcDblShipmentTotal)
    Dim postdata
    postdata = ""
    If pcDblShipmentTotal > 0 Then
        postdata = postdata & "<Line>" & vbCrlf
        postdata = postdata & "<LineNo>" & i & "-FR</LineNo>" & vbCrlf
        postdata = postdata & "<ItemCode>FREIGHT</ItemCode>" & vbCrlf
        postdata = postdata & "<Description>Shipping Charge</Description>" & vbCrlf
        postdata = postdata & "<Qty>1</Qty>" & vbCrlf
        postdata = postdata & "<Amount>"& pcDblShipmentTotal &"</Amount>" & vbCrlf
        postdata = postdata & "<Discounted>false</Discounted>" & vbCrlf
        postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
        postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
        If ptaxAvalaraShippingCode <> "" Then
            postdata = postdata & "<TaxCode>" & ptaxAvalaraShippingCode & "</TaxCode>" & vbCrlf
        End If
        postdata = postdata & "</Line>" & vbCrlf
    End If
    pcf_avaShipmentTotal = postdata
End Function


Public Function pcf_avaPaymentTotal(i, taxPaymentTotal)
    Dim postdata
    postdata = ""
    If taxPaymentTotal > 0 Then
        postdata = postdata & "<Line>" & vbCrlf
        postdata = postdata & "<LineNo>" & i & "-PF</LineNo>" & vbCrlf
        postdata = postdata & "<ItemCode>PROCESSING</ItemCode>" & vbCrlf
        postdata = postdata & "<Description>Processing Fees</Description>" & vbCrlf
        postdata = postdata & "<Qty>1</Qty>" & vbCrlf
        postdata = postdata & "<Amount>"& taxPaymentTotal &"</Amount>" & vbCrlf
        postdata = postdata & "<Discounted>false</Discounted>" & vbCrlf
        postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
        postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
        postdata = postdata & "</Line>" & vbCrlf
    End If 
    pcf_avaPaymentTotal = postdata
End Function


Public Function pcf_avaRewardPoints(i, pRewardPoints)
    Dim postdata
    postdata = ""
    If pRewardPoints > 0 Then
        postdata = postdata & "<Line>" & vbCrlf
        postdata = postdata & "<LineNo>" & i & "-RP</LineNo>" & vbCrlf
        postdata = postdata & "<ItemCode>REWARDPOINTS</ItemCode>" & vbCrlf
        postdata = postdata & "<Description>Reward Points</Description>" & vbCrlf
        postdata = postdata & "<Qty>1</Qty>" & vbCrlf
        postdata = postdata & "<Amount>-"& pRewardPoints &"</Amount>" & vbCrlf
        postdata = postdata & "<Discounted>false</Discounted>" & vbCrlf
        postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
        postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
        postdata = postdata & "<TaxCode>NT</TaxCode>" & vbCrlf
        postdata = postdata & "</Line>" & vbCrlf
    End If
    pcf_avaRewardPoints = postdata
End Function


Public Function pcf_avaGiftWrap(i, pcGWTotal)
    Dim postdata
    postdata = ""
    If pcGWTotal > 0 Then
        postdata = postdata & "<Line>" & vbCrlf
        postdata = postdata & "<LineNo>" & i & "-GW</LineNo>" & vbCrlf
        postdata = postdata & "<ItemCode>GIFTWRAPPING</ItemCode>" & vbCrlf
        postdata = postdata & "<Description>Gift Wrapping</Description>" & vbCrlf
        postdata = postdata & "<Qty>1</Qty>" & vbCrlf
        postdata = postdata & "<Amount>"& pcGWTotal &"</Amount>" & vbCrlf
        postdata = postdata & "<Discounted>false</Discounted>" & vbCrlf
        postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
        postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
        postdata = postdata & "</Line>" & vbCrlf
    End If
    pcf_avaGiftWrap = postdata
End Function


Public Function pcf_avaLineItem(i, ItemCode, pDescription, Quantity, LineTotal, pTaxCode, cTaxCode)
    Dim postdata
    postdata = ""
    postdata = postdata & "<Line>" & vbCrlf
    postdata = postdata & "<LineNo>" & i & "</LineNo>" & vbCrlf
    postdata = postdata & "<ItemCode>" & ItemCode & "</ItemCode>" & vbCrlf
    postdata = postdata & "<Description>" & pDescription & "</Description>" & vbCrlf
    postdata = postdata & "<Qty>" & Quantity & "</Qty>" & vbCrlf
    postdata = postdata & "<Amount>" & LineTotal & "</Amount>" & vbCrlf
    postdata = postdata & "<Discounted>true</Discounted>" & vbCrlf
    postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
    postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
    If pTaxCode <> "" Then
        postdata = postdata & "<TaxCode>" & pTaxCode & "</TaxCode>" & vbCrlf
    ElseIf cTaxCode <> "" Then
        postdata = postdata & "<TaxCode>" & cTaxCode & "</TaxCode>" & vbCrlf
    ElseIf ptaxAvalaraProductCode <> "" Then
        postdata = postdata & "<TaxCode>" & ptaxAvalaraProductCode & "</TaxCode>" & vbCrlf
    End If
	postdata = postdata & "</Line>" & vbCrlf
    pcf_avaLineItem = postdata
End Function


Public Function pcf_GetAvalaraTaxAmount(postdata)

    postURL = ptaxAvalaraURL & "/1.0/tax/get"
	xmlResponse = pcf_GetAvalaraResponse(postURL, postdata)

	ptaxAvalaraOrder = 0
    Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
    if xmlDoc.loadXML(xmlResponse) then
        Set result = xmlDoc.selectSingleNode("GetTaxResult/TotalTaxCalculated")
        if not result is nothing then
            ptaxPrdAmount = result.text
            ptaxAvalaraOrder = 1
        end if
    end if

    pcf_GetAvalaraTaxAmount = ptaxPrdAmount

End Function


Public Function pcf_avaAddresses(pcStrShippingAddress, pcStrShippingAddress2, pcStrShippingCity, pcStrShippingStateCode, pcStrShippingPostalCode, pcStrBillingAddress, pcStrBillingAddress2, pcStrBillingCity, pcStrBillingStateCode, pcStrBillingPostalCode)
    Dim postdata
    postdata = ""
    postdata = postdata & "<Addresses>" & vbCrlf
    postdata = postdata & "<Address>" & vbCrlf
    postdata = postdata & "<AddressCode>" & "2" & "</AddressCode>" & vbCrlf        
    if scShipFromAddress1 <> "" then
        postdata = postdata & "<Line1>" & scShipFromAddress1 & "</Line1>" & vbCrlf
        if scShipFromAddress2 <> "" then
            postdata = postdata & "<Line2>" & scShipFromAddress2 & "</Line2>" & vbCrlf
        end if
        postdata = postdata & "<City>" & scShipFromCity & "</City>" & vbCrlf
        postdata = postdata & "<Region>" & scShipFromState & "</Region>" & vbCrlf
        postdata = postdata & "<PostalCode>" & scShipFromPostalCode & "</PostalCode>" & vbCrlf
    else
        postdata = postdata & "<Line1>" & scCompanyAddress & "</Line1>" & vbCrlf
        postdata = postdata & "<City>" & scCompanyCity & "</City>" & vbCrlf
        postdata = postdata & "<Region>" & scCompanyState & "</Region>" & vbCrlf
        postdata = postdata & "<PostalCode>" & scCompanyZip & "</PostalCode>" & vbCrlf
    end if	                    
    postdata = postdata & "</Address>" & vbCrlf        
    postdata = postdata & "<Address>" & vbCrlf
    postdata = postdata & "<AddressCode>" & "1" & "</AddressCode>" & vbCrlf
    if scAlwAltShipAddress = "0" then
        postdata = postdata & "<Line1>" & pcStrShippingAddress & "</Line1>" & vbCrlf
        if pcStrShippingAddress2 <> "" then
            postdata = postdata & "<Line2>" & pcStrShippingAddress2 & "</Line2>" & vbCrlf
        end if
        postdata = postdata & "<City>" & pcStrShippingCity & "</City>" & vbCrlf
        postdata = postdata & "<Region>" & pcStrShippingStateCode & "</Region>" & vbCrlf
        postdata = postdata & "<PostalCode>" & pcStrShippingPostalCode & "</PostalCode>" & vbCrlf
    else
        postdata = postdata & "<Line1>" & pcStrBillingAddress & "</Line1>" & vbCrlf
        if pcStrBillingAddress2 <> "" then
            postdata = postdata & "<Line2>" & pcStrBillingAddress2 & "</Line2>" & vbCrlf
        end if
        postdata = postdata & "<City>" & pcStrBillingCity & "</City>" & vbCrlf
        postdata = postdata & "<Region>" & pcStrBillingStateCode & "</Region>" & vbCrlf
        postdata = postdata & "<PostalCode>" & pcStrBillingPostalCode & "</PostalCode>" & vbCrlf
    end if
    postdata = postdata & "</Address>" & vbCrlf        
    postdata = postdata & "</Addresses>" & vbCrlf
    pcf_avaAddresses = postdata
End Function


Sub pcs_CommitAvalara(pIdOrder, action)

	query = "SELECT Avalara_orders.id, idOrderCounter, status, orders.idOrder, idCustomer, orderDate, address, address2, city, stateCode, zip, shippingAddress, shippingAddress2, shippingCity, shippingStateCode, shippingZip, pcOrd_Avalara, shipmentdetails, pcOrd_CatDiscounts, paymentdetails, discountdetails, pcOrd_GWTotal, iRewardValue FROM orders INNER JOIN Avalara_orders On orders.idOrder = Avalara_orders.idOrder WHERE orders.idOrder = " & pIdOrder & " ORDER BY idOrderCounter DESC"
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs = conntemp.execute(query)	
	If Not rs.Eof Then    
        pcv_intAvalaraOrder = rs("pcOrd_Avalara")
		aId = rs("id")
		aIdOrder = rs("idOrderCounter")
        If IsNull(aIdOrder) OR aIdOrder=0 Then
            aIdOrder = pIdOrder
        End If
		aStatus = rs("status")
		pIdCustomer = rs("idCustomer")
		pOrderDate = rs("orderDate")
		pcStrBillingAddress = rs("address")
		pcStrBillingAddress2 = rs("address2")
		pcStrBillingCity = rs("city")
		pcStrBillingStateCode = rs("stateCode")
		pcStrBillingPostalCode = rs("zip")
		pcStrShippingAddress = rs("shippingaddress")
		pcStrShippingAddress2 = rs("shippingaddress2")
		pcStrShippingCity = rs("shippingcity")
		pcStrShippingStateCode = rs("shippingStateCode")
		pcStrShippingPostalCode = rs("shippingZip")
		shipmentArray = rs("shipmentdetails")
		pcCatDiscTotal = rs("pcOrd_CatDiscounts")
		ppaymentDetails = rs("paymentdetails")
		pdiscountdetails = rs("discountdetails")
		pcGWTotal = rs("pcOrd_GWTotal")
		pRewardPoints = rs("iRewardValue")
    End If
    Set rs = Nothing
    
    
	if pcv_intAvalaraOrder = 1 then

		pDocDate = year(pOrderDate) & "-" & format_zeros(month(pOrderDate), 2) & "-" & format_zeros(day(pOrderDate), 2)

		'// Get Tax Document
		postdata = postdata & "<GetTaxRequest>" & vbCrlf
		postdata = postdata & "<Client>a0o33000003xKBM</Client>" & vbCrlf
		postdata = postdata & "<DocDate>" & pDocDate & "</DocDate>" & vbCrlf
		postdata = postdata & "<CustomerCode>" & "CUST-" & pIdCustomer & "</CustomerCode>" & vbCrlf 

        '// Update
		if (action = "") OR (lcase(aStatus) = "uncommitted") OR (lcase(action) = "update") then
			
            postdata = postdata & "<DocCode>" & scpre+aIdOrder & "</DocCode>" & vbCrlf
		
        '// Edit
        elseif lcase(action) = "edit" then
        
			postdata = postdata & "<DocCode>" & scpre+aIdOrder+0.1 & "</DocCode>" & vbCrlf
			postdata = postdata & "<ReferenceCode>" & scpre+aIdOrder & "</ReferenceCode>" & vbCrlf
		
        '// Commit
        elseif lcase(action) = "commit" then
        
			postdata = postdata & "<DocCode>" & scpre+aIdOrder & "</DocCode>" & vbCrlf
			postdata = postdata & "<ReferenceCode>" & scpre+aIdOrder-0.1 & "</ReferenceCode>" & vbCrlf
		end if
		
		postdata = postdata & "<DocType>SalesInvoice</DocType>" & vbCrlf
		
		if lcase(action) <> "update" AND lcase(action) <> "edit" then
			postdata = postdata & "<Commit>" & "true" & "</Commit>" & vbCrlf
		end if

		query = "SELECT pcCust_AvalaraExemptionNo, customerType FROM customers WHERE idcustomer = " & pIdCustomer
		set rstemp = server.CreateObject("ADODB.RecordSet")
		set rstemp = conntemp.execute(query)
        If Not rstemp.Eof Then
            pcv_strAvalaraExemptionNo = rstemp("pcCust_AvalaraExemptionNo")
            pcv_intCustomerType = rstemp("customerType")
        End If
        set rstemp = nothing
        
		if len(pcv_strAvalaraExemptionNo)>0 then
			postdata = postdata & "<ExemptionNo>" & pcv_strAvalaraExemptionNo & "</ExemptionNo>" & vbCrlf
		else
			if ptaxwholesale = 0 AND pcv_intCustomerType = 1 then
				postdata = postdata & "<CustomerUsageType>" & ptaxAvalaraReason & "</CustomerUsageType>" & vbCrlf
			end if
		end if
		set rstemp = nothing

        '// Addresses
        postdata = postdata & pcf_avaAddresses(pcStrShippingAddress, pcStrShippingAddress2, pcStrShippingCity, pcStrShippingStateCode, pcStrShippingPostalCode, pcStrBillingAddress, pcStrBillingAddress2, pcStrBillingCity, pcStrBillingStateCode, pcStrBillingPostalCode)

        '// Lines
		postdata = postdata & "<Lines>" & vbCrlf
		
        '// Product Tax
		query = "SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, unitPrice, QDiscounts, products.sku, description, pcProd_AvalaraTaxCode, pcprod_ParentPrd, idconfigSession,pcPrdOrd_BundledDisc, ItemsDiscounts FROM ProductsOrdered INNER JOIN products ON ProductsOrdered.idProduct = products.idProduct WHERE ProductsOrdered.idOrder = " & pIdOrder
		set rs = server.CreateObject("ADODB.RecordSet")
		set rs = connTemp.execute(query)

		i = 1
		do until rs.eof
        
            pcv_intProductLineTotal = (rs("quantity") * rs("unitPrice"))
            
            '// Bundle Discounts
            pcv_intProductLineTotal = (pcv_intProductLineTotal - rs("pcPrdOrd_BundledDisc"))

            '// Item Discounts
            pcv_intProductLineTotal = (pcv_intProductLineTotal - rs("ItemsDiscounts"))
            
            '// Additional Charges
            Charges = 0
            If scBTO=1 then
                pidConfigSession = rs("idconfigSession")
                if pidConfigSession<>"0" then

                    query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
                    set rsConfigObj=server.CreateObject("ADODB.RecordSet")
                    set rsConfigObj=connTemp.execute(query)
                    stringCProducts=rsConfigObj("stringCProducts")
                    stringCValues=rsConfigObj("stringCValues")
                    stringCCategories=rsConfigObj("stringCCategories")                    
                    Set rsConfigObj = Nothing
                    
                    ArrCProduct=Split(stringCProducts, ",")
                    ArrCValue=Split(stringCValues, ",")
                    ArrCCategory=Split(stringCCategories, ",")
                    if ArrCProduct(0)<>"na" then

                        for j=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
                            
                            query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(j)&") AND ((products.idProduct)="&ArrCProduct(j)&"))"
                            set rsConfigObj=server.CreateObject("ADODB.RecordSet")
                            set rsConfigObj=connTemp.execute(query)
                            pcategoryDesc=rsConfigObj("categoryDesc")
                            pdescription=rsConfigObj("description")
                            psku=rsConfigObj("sku")
                            pItemWeight=rsConfigObj("weight")
                            intTotalWeight=intTotalWeight+pItemWeight
                            if (CDbl(ArrCValue(j))>0)then
                                Charges = Charges + cdbl(ArrCValue(j))
                            end if 
                            set rsConfigObj = nothing
                        next

                    end if
                end if
            end if 
            pcv_intProductLineTotal = (pcv_intProductLineTotal + Charges)
            
            '// Quantity Discounts (?)
            pcv_intProductLineTotal = (pcv_intProductLineTotal - rs("QDiscounts"))

            pcv_ParentPrd = rs("pcprod_ParentPrd")
            pcv_AvalaraProductTaxCode = rs("pcProd_AvalaraTaxCode")
            
            '// Avalara Category Tax Code
            If pcv_ParentPrd>0 Then
                query = "SELECT pcCats_AvalaraTaxCode FROM categories c JOIN categories_products cp ON cp.idCategory = c.idCategory WHERE cp.idProduct = " & pcv_ParentPrd & " AND pcCats_AvalaraTaxCode <> ''"
            Else
                query = "SELECT pcCats_AvalaraTaxCode FROM categories c JOIN categories_products cp ON cp.idCategory = c.idCategory WHERE cp.idProduct = " & rs("idProduct") & " AND pcCats_AvalaraTaxCode <> ''"
            End If
            set rsCat = server.CreateObject("ADODB.RecordSet")
            set rsCat = conntemp.execute(query)
            if not rsCat.eof then 
                pcv_AvalaraCategoryTaxCode = rsCat("pcCats_AvalaraTaxCode")
            end if
            set rsCat = nothing

            '// Product Line Item
            postdata = postdata & pcf_avaLineItem(i, rs("sku"), rs("description"), rs("quantity"), pcv_intProductLineTotal, pcv_AvalaraProductTaxCode, pcv_AvalaraCategoryTaxCode)

            rs.MoveNext
			i = i + 1
		loop
		set rs = nothing
		
        '// Payment Fees
		If ppaymentDetails <> "" then
			pcArrayPayment = split(trim(ppaymentDetails),"||")
			paymentFee = trim(pcArrayPayment(1))
			If paymentFee<>"" then
				postdata = postdata & pcf_avaPaymentTotal(i, paymentFee)
			End if
		End if
		
        '// Shipping charge Tax
		shipSplit = split(shipmentArray, ",")
		if ubound(shipSplit) > 1 then
			if isNumeric(trim(shipSplit(2))) then
				pcDblShipmentTotal = trim(shipSplit(2))
				pcDblShipmentTotal = replace(pcDblShipmentTotal, CHR(13), "")
				if ubound(shipSplit) => 3 then
					pcDblServiceHandlingFee = trim(shipSplit(3))
					if NOT isNumeric(pcDblServiceHandlingFee) then
						pcDblServiceHandlingFee = 0
					end if
				else
					pcDblServiceHandlingFee = 0
				end if
			end if
		else
			pcDblShipmentTotal = 0
		end if
        postdata = postdata & pcf_avaShipmentTotal(i, pcDblShipmentTotal)

		'// Handling fee Tax
        postdata = postdata & pcf_avaServiceHandlingFee(i, pcDblServiceHandlingFee)

		'// Discounts
		discountTotal = 0
		promotionTotal = 0
		discountPriceTotal = 0
		pdiscountDetails=trim(pdiscountdetails)
		
		intArryCnt=-1
		discountTotalPrice=0
        discountArray = pdiscountDetails
		if discountArray<>"" AND instr(discountArray, "||") then 

			if instr(discountArray,",") then
				DiscountDetailsArry=split(discountArray,",")
				intArryCnt=ubound(DiscountDetailsArry)
			else
				intArryCnt=0
			end if

			strDiscountTableRow=""
			for k=0 to intArryCnt
				if intArryCnt=0 then
					pTempDiscountDetails=discountArray
				else
					pTempDiscountDetails=DiscountDetailsArry(k)
				end if
				discountPrice = 0
				if instr(pTempDiscountDetails,"- ||") then
					discounts = split(pTempDiscountDetails,"- ||")
					discountType = discounts(0)
					discountPrice = discounts(1)
					discountTotalPrice=discountTotalPrice+discountPrice
				end if

				if discountPrice<>0 then
                 
                    discountPriceTotal = discountPriceTotal + discountPrice
          
                end if
			next  

		end if

		'// Gift Wrapping
        postdata = postdata & pcf_avaGiftWrap(i, pcGWTotal)
		
		'// Reward Points
        postdata = postdata & pcf_avaRewardPoints(i, pRewardPoints)
		
		postdata = postdata & "</Lines>" & vbCrlf  
        
        '// Discounts
        postdata = postdata & pcf_avaDiscounts(i, pcCatDiscTotal, discountPriceTotal, promotionTotal)
    
		postdata = postdata & "</GetTaxRequest>"

        '// Get Tax Total
		postURL = ptaxAvalaraURL & "/1.0/tax/get"
		xmlResponse = pcf_GetAvalaraResponse(postURL, postdata)

		Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
		if xmlDoc.loadXML(xmlResponse) then
			pcv_ErrorMessage = ""
			Set result = xmlDoc.selectSingleNode("GetTaxResult/ResultCode")

			if lcase(result.text) = "success" then

                Set result = xmlDoc.selectSingleNode("GetTaxResult/TotalTaxCalculated")
                if not result is nothing then
                    ptaxAmount = result.text
                    ptaxDetailsString = "Tax|" & ptaxAmount & ","
                    
                    '// Save Tax Amounts
                    if lcase(action) = "update" then
                        
                        query = "UPDATE orders SET taxAmount = " & ptaxAmount & ", taxdetails = '" & ptaxDetailsString & "' WHERE idOrder = " & pIdOrder
                        conntemp.execute(query)
                        
                    elseif lcase(action) = "edit" then
                        
                        query = "UPDATE orders SET taxAmount = " & ptaxAmount & ", taxdetails = '" & ptaxDetailsString & "', total = total+" & ptaxAmount & " WHERE idOrder = " & pIdOrder
                        conntemp.execute(query)
                        
                    end if

                    '// Keep Count and Status
                    if lcase(aStatus) <> "uncommitted" then
                        query = "INSERT INTO Avalara_orders (idOrder, idOrderCounter, status, updatedDate) VALUES(" & pIdOrder & ", " & aIdOrder + 0.1 & ", 'Uncommitted', '" & Now() & "')"
                        conntemp.execute(query)
                    end if
                    
                    if lcase(action) = "commit" then
    
                        query = "UPDATE Avalara_orders SET idOrderCounter = " & aIdOrder & ", status = 'Committed', updatedDate = '" & Now() & "' WHERE id = " & aId
                        conntemp.execute(query)
                        
                    end if
                    
                end if
			
			else
				pcv_ErrorMessage = "Sorry, there was an error occurred while processing Avalara transaction."
			end if
		end if
	end if
	Set rs = nothing

End Sub


Sub pcs_VoidAvalara(aIdOrder)

	postdata = postdata & "<CancelTaxRequest>" & vbCrlf
	postdata = postdata & "<Client>a0o33000003xKBM</Client>" & vbCrlf
	postdata = postdata & "<CompanyCode>" & ptaxAvalaraCode & "</CompanyCode>" & vbCrlf
	postdata = postdata & "<DocType>" & "SalesInvoice" & "</DocType>" & vbCrlf
	postdata = postdata & "<DocCode>" & scpre+aIdOrder & "</DocCode>" & vbCrlf
	postdata = postdata & "<CancelCode>DocVoided</CancelCode>" & vbCrlf
	postdata = postdata & "</CancelTaxRequest>"
	
	postURL = ptaxAvalaraURL & "/1.0/tax/cancel"
	xmlResponse = pcf_GetAvalaraResponse(postURL, postdata)
	
	Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
	if xmlDoc.loadXML(xmlResponse) then
		pcv_ErrorMessage = ""
		Set result = xmlDoc.selectSingleNode("CancelTaxResponse/CancelTaxResult/ResultCode")
		if lcase(result.text) = "success" then
			query = "UPDATE Avalara_orders SET status = 'Voided', updatedDate = '" & Now() & "' WHERE idOrderCounter = "& aIdOrder
			conntemp.execute(query)
		else
			pcv_ErrorMessage = "Sorry, there was an error while processing Avalara transaction. Please contact support for more information."
		end if
	end if
	
End Sub


Sub pcs_ReturnAvalara(pIdOrder, products)

	query = "SELECT orders.idOrder, idOrderCounter, status, idCustomer, orderDate, address, address2, city, stateCode, zip, shippingAddress, shippingAddress2, shippingCity, shippingStateCode, shippingZip, pcOrd_Avalara FROM orders INNER JOIN Avalara_orders On orders.idOrder = Avalara_orders.idOrder WHERE orders.idOrder = " & pIdOrder & " ORDER BY idOrderCounter DESC"
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs = conntemp.execute(query)	
	If Not rs.Eof Then

        aIdOrder = rs("idOrderCounter")
        pcv_intAvalaraOrder = rs("pcOrd_Avalara")
		pIdCustomer = rs("idCustomer")
		pOrderDate = rs("orderDate")
		pcStrBillingAddress = rs("address")
		pcStrBillingAddress2 = rs("address2")
		pcStrBillingCity = rs("city")
		pcStrBillingStateCode = rs("stateCode")
		pcStrBillingPostalCode = rs("zip")
		pcStrShippingAddress = rs("shippingaddress")
		pcStrShippingAddress2 = rs("shippingaddress2")
		pcStrShippingCity = rs("shippingcity")
		pcStrShippingStateCode = rs("shippingStateCode")
		pcStrShippingPostalCode = rs("shippingZip")
        
    End If
    Set rs = Nothing
    
    If pcv_intAvalaraOrder = 1 Then
		
		pDocDate = year(date) & "-" & format_zeros(month(date), 2) & "-" & format_zeros(day(date), 2)
		pTaxDate = year(pOrderDate) & "-" & format_zeros(month(pOrderDate), 2) & "-" & format_zeros(day(pOrderDate), 2)

		'// Get Tax Document
		postdata = postdata & "<GetTaxRequest>" & vbCrlf
		postdata = postdata & "<Client>a0o33000003xKBM</Client>" & vbCrlf
		postdata = postdata & "<DocDate>" & pDocDate & "</DocDate>" & vbCrlf
		postdata = postdata & "<CustomerCode>" & "CUST-" & pIdCustomer & "</CustomerCode>" & vbCrlf
		postdata = postdata & "<DocCode>" & scpre+int(aIdOrder) & "</DocCode>" & vbCrlf
		postdata = postdata & "<ReferenceCode>" & scpre+int(pIdOrder) & "</ReferenceCode>" & vbCrlf
		postdata = postdata & "<DocType>ReturnInvoice</DocType>" & vbCrlf
		postdata = postdata & "<Commit>" & "true" & "</Commit>" & vbCrlf

        '// Get Tax Override
		postdata = postdata & "<TaxOverride>"
		postdata = postdata & "<TaxOverrideType>TaxDate</TaxOverrideType>"
		postdata = postdata & "<TaxDate>" & pTaxDate & "</TaxDate>"
		postdata = postdata & "<Reason>Return</Reason>"
		postdata = postdata & "</TaxOverride>"

        '// Addresses
        postdata = postdata & pcf_avaAddresses(pcStrShippingAddress, pcStrShippingAddress2, pcStrShippingCity, pcStrShippingStateCode, pcStrShippingPostalCode, pcStrBillingAddress, pcStrBillingAddress2, pcStrBillingCity, pcStrBillingStateCode, pcStrBillingPostalCode)
        
		'// Lines
		postdata = postdata & "<Lines>" & vbCrlf
		
		arrProducts = split(products, ",")
		returnedQty = 0
		For i=0 To (ubound(arrProducts)-1)
			
            query = "SELECT ProductsOrdered.rmaSubmitted, unitPrice, products.sku, description, pcProd_AvalaraTaxCode, pcprod_ParentPrd FROM ProductsOrdered INNER JOIN products ON ProductsOrdered.idProduct = products.idProduct WHERE ProductsOrdered.idProductOrdered = " & arrProducts(i)
			set rs = server.CreateObject("ADODB.RecordSet")
			set rs = connTemp.execute(query)
			
			if rs("rmaSubmitted") > 0 then
				returnedQty = 1
				postdata = postdata & "<Line>" & vbCrlf
				postdata = postdata & "<LineNo>" & i+1 & "</LineNo>" & vbCrlf
				postdata = postdata & "<ItemCode>" & rs("sku") & "</ItemCode>" & vbCrlf
				postdata = postdata & "<Description>" & rs("description") & "</Description>" & vbCrlf
				postdata = postdata & "<Qty>" & rs("rmaSubmitted") & "</Qty>" & vbCrlf
				postdata = postdata & "<Amount>-" & rs("rmaSubmitted") * rs("unitPrice") & "</Amount>" & vbCrlf
				postdata = postdata & "<Discounted>true</Discounted>" & vbCrlf
				postdata = postdata & "<OriginCode>" & "2" & "</OriginCode>" & vbCrlf
				
				pcv_ParentPrd = rs("pcprod_ParentPrd")
				pcv_AvalaraProductTaxCode = rs("pcProd_AvalaraTaxCode")
				
				'// Avalara Category Tax Code
				If pcv_ParentPrd>0 Then
					query = "SELECT pcCats_AvalaraTaxCode FROM categories c JOIN categories_products cp ON cp.idCategory = c.idCategory WHERE cp.idProduct = " & pcv_ParentPrd & " AND pcCats_AvalaraTaxCode <> ''"
				Else
					query = "SELECT pcCats_AvalaraTaxCode FROM categories c JOIN categories_products cp ON cp.idCategory = c.idCategory WHERE cp.idProduct = " & arrProducts(i) & " AND pcCats_AvalaraTaxCode <> ''"
				End If
				set rsCat = server.CreateObject("ADODB.RecordSet")
				set rsCat = conntemp.execute(query)
				if not rsCat.eof then 
					pcv_AvalaraCategoryTaxCode = rsCat("pcCats_AvalaraTaxCode")
				end if
				set rsCat = nothing
				
				if pcv_AvalaraProductTaxCode <> "" then
					postdata = postdata & "<TaxCode>" & pcv_AvalaraProductTaxCode & "</TaxCode>" & vbCrlf
				elseif pcv_AvalaraCategoryTaxCode <> "" then
					postdata = postdata & "<TaxCode>" & pcv_AvalaraCategoryTaxCode & "</TaxCode>" & vbCrlf
				elseif ptaxAvalaraProductCode <> "" then
					postdata = postdata & "<TaxCode>" & ptaxAvalaraProductCode & "</TaxCode>" & vbCrlf
				end if
				
				postdata = postdata & "<DestinationCode>" & "1" & "</DestinationCode>" & vbCrlf
				postdata = postdata & "</Line>" & vbCrlf
			end if
            
            set rs = nothing
            
		Next

		postdata = postdata & "</Lines>" & vbCrlf
		postdata = postdata & "</GetTaxRequest>"
		
		if returnedQty = 1 then
			postURL = ptaxAvalaraURL & "/1.0/tax/get"
			xmlResponse = pcf_GetAvalaraResponse(postURL, postdata)
			
			Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
			If xmlDoc.loadXML(xmlResponse) Then
				Set result = xmlDoc.selectSingleNode("GetTaxResult/ResultCode")
				If lcase(result.text) = "success" Then
					
					'// Update Order
					query = "UPDATE Avalara_orders SET status = 'Returned', updatedDate = '" & Now() & "' WHERE idOrderCounter = " & aIdOrder
					conntemp.execute(query)
					
				End If
			End If
		End If
        
	End If '// If pcv_intAvalaraOrder = 1 Then
	
End Sub
%>