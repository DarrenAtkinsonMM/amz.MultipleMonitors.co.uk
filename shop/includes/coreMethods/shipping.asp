<%
Dim pcv_boolHasMapping, pcv_boolShowFilteredRates
Dim CntFree, DCnt, serviceFree, serviceFreeOverAmt, serviceCode, OrderTotal, shipArray, shipDetailsArray, tempRate, tempRateDisplay
Dim HasDefaultM,tMArr,tSeArr,MCount
Dim tSArr(100,10)
Dim tSCount,FedExWSCnt,USPSCnt,UPSCnt,CPCnt,CUSTOMCnt,pcv_Default
Dim pshipDetailsArray2
Dim strFEDEX,strUSPS,strUPS,strCP,strCUSTOM
Dim serviceTypeClass,deliveryTimeClass,rateClass,showDeliveryCol
Dim strDefaultProvider,iShipmentTypeCnt,strOptionShipmentType,availableShipStr,strTabShipmentType,iFedExWSFlag,iCPFlag,iCustomFlag
Dim UPS_active,UPS_userid,UPS_password,UPS_license_key
Dim CP_active,CP_server,CP_username,CP_password,CP_custNo
Dim pcv_strMethodNameWS, pcv_strMethodReplyWS, fedex_postdataWS, objFedExWSClass, objOutputXMLDocWS, srvFEDEXWSXmlHttp, FEDEXWS_result, pcv_strErrorMsgWS
Dim usps_userid,usps_server,usps_active
Dim FedEXWS_server,FedEXWS_active,FedEXWS_AccountNumber,FedEXWS_MeterNumber,FEDEXWS_Environment
Dim ErrPageName
Dim pcv_intIncludeDiscounts
Dim pcBillingStateCode,pcBillingAddress,pcBillingCity,pcBillingProvince,pcBillingPostalCode,pcBillingCountryCode,pcCustomerEmail,pResidentialShipping,pcShippingAddress,pcShippingCity,pcShippingStateCode,pcShippingProvince,pcShippingPostalCode,pcShippingCountryCode
Dim CountryCode,StateCode,Province,city,zip
Dim pShipSubTotal,pShipCDSubTotal,TotalPromotions,pSubTotal
Dim pShipWeight,intUniversalWeight,pCartQuantity,pCartShipQuantity,pCartSurcharge
Dim Universal_destination_provOrState,Universal_destination_country,Universal_destination_postal,Universal_destination_city,Universal_destination_address
Dim shipcompany,pcv_NoDynamicShipping,flagShp,intIdFlatShipType,pShpObjType

Public Function pcf_ShowFilteredRates()

    pcv_boolHasMapping = False
    
    pcv_boolShowFilteredRates = scUseShipMap
    If pcv_boolShowFilteredRates = "" Then
        pcv_boolShowFilteredRates = "0"
    End If
    
    If pcv_boolShowFilteredRates = "1" Then
        queryQ = "SELECT TOP 1 pcSM_ID, pcSM_Name, pcSM_Type, pcSM_Order FROM pcShippingMap;"
        Set rsQ = connTemp.execute(queryQ)
        If rsQ.Eof Then
            pcv_boolShowFilteredRates = "0"
        Else
            pcv_boolHasMapping = True
        End If
        Set rsQ = nothing
    end if    
    pcf_ShowFilteredRates = pcv_boolShowFilteredRates
    
End Function

Public Function pcf_HasShipMapping()

    pcv_boolHasMapping = False

    queryQ = "SELECT TOP 1 pcSM_ID, pcSM_Name, pcSM_Type, pcSM_Order FROM pcShippingMap;"
    Set rsQ = connTemp.execute(queryQ)
    If Not rsQ.Eof Then
        pcv_boolHasMapping = True
    End If
    Set rsQ = nothing
  
    pcf_HasShipMapping = pcv_boolHasMapping
    
End Function

Function calculateShippingPrice(savNullShipper, savNullShipRates, pcShippingArray)

    If savNullShipper="Yes" then
        pcShipmentPriceToAdd = "0"
    Else

        If savNullShipRates="Yes" Then
            pcShipmentPriceToAdd="0"
        Else

            TempStrNewShipping = ""
            pcSplitShipping = split(pcShippingArray,",")

            TempStrShipper=pcSplitShipping(0)
            TempStrService=pcSplitShipping(1)
            TempDblPostage=pcSplitShipping(2)
                
            If ubound(pcSplitShipping)>4 Then

                query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceCode='" & pcSplitShipping(5) & "';"
                set rs=server.CreateObject("ADODB.RecordSet")
                set rs=connTemp.execute(query)
                If Not rs.eof Then
                    pcIntIdShipService = rs("idshipservice")
                    serviceFreeOverAmt = rs("serviceFreeOverAmt")
                End If
                set rs=nothing
     
            End If
            TempStrNewShipping = TempStrNewShipping & TempStrShipper & "," & TempStrService & "," & TempDblPostage
                
            If TempStrService="" Then

                If pcIntIdShipService="" Then

                    query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceDescription like '%" & TempStrShipper & "%'"
                    Set rs = server.CreateObject("ADODB.RecordSet")
                    Set rs = connTemp.execute(query)                        
                    If Not rs.eof Then
                        pcIntIdShipService=rs("idshipservice")
                        serviceFreeOverAmt=rs("serviceFreeOverAmt")
                    End If                      
                    Set rs=nothing

                End If
                    
            Else '// If TempStrService="" Then
                
                If pcIntIdShipService="" Then

                    query = "SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceDescription LIKE '%" & TempStrService & "%'"
                    Set rs = server.CreateObject("ADODB.RecordSet")
                    Set rs = connTemp.execute(query)
                    If Not rs.eof Then
                        pcIntIdShipService=rs("idshipservice")
                        serviceFreeOverAmt=rs("serviceFreeOverAmt")
                    End If
                    Set rs = Nothing 

                End If

            End If '// If TempStrService="" Then
                
            pcShipmentPriceToAdd = TempDblPostage

            If ubound(pcSplitShipping) = 3 Or ubound(pcSplitShipping) > 3 Then

                pcDblServiceHandlingFee = pcSplitShipping(3)
                TempStrNewShipping = TempStrNewShipping & "," & pcSplitShipping(3)
                
                If ubound(pcSplitShipping) = 4 Then
                    pcDblIncHandlingFee = pcSplitShipping(4)
                Else
                    pcDblIncHandlingFee = 0
                End If
                    
            Else '// If ubound(pcSplitShipping) = 3 Or ubound(pcSplitShipping) > 3 Then
            
                pcDblServiceHandlingFee = 0
                pcDblIncHandlingFee = 0
                
            End If '// If ubound(pcSplitShipping) = 3 Or ubound(pcSplitShipping) > 3 Then
                
        End If '// If savNullShipRates="Yes" Then

    End If '// If savNullShipper="Yes" then

    calculateShippingPrice = pcShipmentPriceToAdd
    
End Function



Function calculateServiceHandlingFee(savNullShipper, savNullShipRates, pcShippingArray)

    If Not savNullShipper="Yes" then

        If Not savNullShipRates="Yes" Then

            pcSplitShipping = split(pcShippingArray,",")

            If ubound(pcSplitShipping) = 3 Or ubound(pcSplitShipping) > 3 Then
                pcDblServiceHandlingFee = pcSplitShipping(3)
            Else             
                pcDblServiceHandlingFee = 0
            End If 
                
        End If 

    End If 

    calculateServiceHandlingFee = pcDblServiceHandlingFee
    
End Function



Public Function IsOPCReady(pSubTotal, pcDblServiceHandlingFee, GCAmount, serviceFreeOverAmt, pcDblShipmentTotal, pcv_FREESHIP, pSubTotalCheckFreeShipping, pcIntIdShipService)

    Dim pOPCReady

    Dim IsFreeShippingAvailable
    IsFreeShippingAvailable = ((cdbl(pSubTotalCheckFreeShipping) + cdbl(GCAmount)) < cdbl(serviceFreeOverAmt))

    If (pcDblShipmentTotal=0) AND (IsFreeShippingAvailable) AND (pcv_FREESHIP<>"ok") Then
        tmpErrorOPCReady = dictLanguage.Item(Session("language")&"_opc_43")
        pOPCReady="NO"
    Else
       pOPCReady="YES"

        If pcIntIdShipService<>"" Then        
            
            query="SELECT serviceCode FROM shipService WHERE idshipservice=" & pcIntIdShipService & ";"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=connTemp.execute(query)						
            If Not rs.eof Then
                pcServiceCode = rs("serviceCode")
            End If
			set rs=nothing

            If pcServiceCode<>"" Then
            
                If Left(pcServiceCode,1) = "C" And IsNumeric(mid(pcServiceCode, 2, len(pcServiceCode))) Then
                
                    tmpCheckShipValue = pSubTotal - pcDblShipmentTotal - pcDblServiceHandlingFee
			
                    query="SELECT shippingPrice FROM FlatShipTypeRules INNER JOIN FlatShipTypes ON FlatShipTypeRules.idFlatShipType=FlatShipTypes.idFlatShipType WHERE FlatShipTypes.WQP='P' AND FlatShipTypes.idFlatShipType=" & mid(pcServiceCode,2,len(pcServiceCode)) & " AND FlatShipTypeRules.quantityFrom<=" & tmpCheckShipValue & " AND FlatShipTypeRules.quantityTo>=" & tmpCheckShipValue & ";"
                    Set rs=connTemp.execute(query)
                    If Not rs.eof Then
                        
						pCartSurcharge=Cdbl(calculateTotalProductSurcharge(pcCartArray, ppcCartIndex))
                        tmpShipPrice=rs("shippingPrice")+pCartSurcharge

                        If (Cdbl(tmpShipPrice) <> Cdbl(pcDblShipmentTotal)) AND (IsFreeShippingAvailable) AND (pcv_FREESHIP<>"ok") Then
                            tmpErrorOPCReady = dictLanguage.Item(Session("language")&"_opc_43a")
                            pOPCReady = "NO"
                        End If

                    End If
					Set rs = Nothing
                
                End If '// If Left(pcServiceCode,1) = "C" And IsNumeric(mid(pcServiceCode, 2, len(pcServiceCode))) Then

            End If '// If pcServiceCode<>"" Then
            
        End If '// If pcIntIdShipService<>"" Then

    End If

    IsOPCReady = pOPCReady
    
End Function





'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'// START:  v4 Methods - Need to consolidate and merge with v5 above
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////


' Cart Ship Amount
function calculateShipCartTotal(pcCartArray, indexCart)
	dim f, total
	total=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND pcCartArray(f,20)=0 then   
			if pcCartArray(f,16)<>"" then
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3))  
				end if  
			end if  
		end if
	next
	calculateShipCartTotal=total  
	set f=nothing
	set total=nothing 
end function


'If Editing an existing order
function eoversizecheck(pcOSCheckOrderNumber)
	dim f 
	query="SELECT ProductsOrdered.quantity, products.OverSizeSpec, products.weight, ProductsOrdered.idOrder FROM ProductsOrdered INNER JOIN products ON ProductsOrdered.idProduct = products.idProduct WHERE (((ProductsOrdered.idOrder)="&pcOSCheckOrderNumber&"));"
	set rsOSC=server.CreateObject("ADODB.RecordSet")
	set rsOSC=connTemp.execute(query)
	eoversizecheck=""
	do until rsOSC.eof
		pcOS_Quantity=rsOSC("quantity")
		pcOS_OverSizeSpec=rsOSC("OverSizeSpec")
		pcOS_weight=rsOSC("weight")
		if pcOS_OverSizeSpec<>"NO" then
			OSArray=split(pcOS_OverSizeSpec,"||")
			if ubound(OSArray)>3 then
				for i=1 to pcOS_Quantity
					eoversizecheck=eoversizecheck&"1|||"&pcOS_OverSizeSpec&"||"&pcOS_weight&","
				next
			end if
		end if
		rsOSC.MoveNext
	loop
	set rsOSC=nothing
	set f=nothing   
end function


'check for oversized product
function oversizecheck(pcCartArray, indexCart)
	dim f 
	if indexCart>0 then
		oversizecheck=""
		for f=1 to indexCart
			If pcCartArray(f,20)="-1" Then
				'response.write ""
			Else
				if pcCartArray(f,10)=0 and pcCartArray(f,23)<>"NO" then
					OSArray=split(pcCartArray(f,23),"||")
					if ubound(OSArray)>3 then
						for i=1 to pcCartArray(f,2)
							oversizecheck=oversizecheck&"1|||"&pcCartArray(f,23)&"||"&pcCartArray(f,6)&"||"&pcCartArray(f,3)&","
						next
					end if
				end if
				if pcCartArray(f,16)<>"" then
					'//Get config info
					query="SELECT stringProducts, stringPrice, stringQuantity FROM configSessions WHERE idconfigSession="&pcCartArray(f,16)&";"
					set rsBTOChkObj=server.CreateObject("ADODB.RecordSet")
					set rsBTOChkObj=conntemp.execute(query)
					pcv_itemString = rsBTOChkObj(0)
					pcv_itemPrice = rsBTOChkObj(1)
					pcv_itemQty = rsBTOChkObj(2)
					pcv_itemStringArry = split(pcv_itemString,",")
					pcv_itemPriceArry = split(pcv_itemPrice,",")
					pcv_itemQtyArry = split(pcv_itemQty,",")
					
					for iOSChkCnt=lbound(pcv_itemStringArry) to ubound(pcv_itemStringArry)-1
						pcv_tempPrdChk = pcv_itemStringArry(iOSChkCnt)
						query="SELECT weight, oversizeSpec FROM products WHERE idProduct="&pcv_tempPrdChk&";"
						set rsOSChkObj=server.CreateObject("ADODB.RecordSet")
						set rsOSChkObj=conntemp.execute(query)
						pcv_OSChkWeight=rsOSChkObj(0)
						pcv_OSChkSpec=rsOSChkObj(1)
						if pcv_OSChkSpec<>"NO" then
							for i=1 to pcCartArray(f,2)
								for iICnt=1 to pcv_itemQtyArry(iOSChkCnt)
									'//oversized, get array
									oversizecheck=oversizecheck&"2|||"&pcv_OSChkSpec&"||"&pcv_OSChkWeight&"||"&pcv_itemPriceArry(iOSChkCnt)&","
								next
							next
						end if
					next
				end if
			End If
		next
	end if  
	set f=nothing   
end function


'// START:  Tokenize Billing Address
Public Function pcf_TokenizeBillingAddress()

    pcf_TokenizeBillingAddress = pcStrBillingAddress & "|" & pcStrBillingAddress2 & "|" & pcStrBillingCity & "|" & pcStrBillingStateCode & "|" & pcStrBillingPostalCode & "|" & pcStrBillingCountryCode

End Function
'// END:  Tokenize Billing Address


'// START:  Tokenize Shipping Address
Public Function pcf_TokenizeShippingAddress()

    pcf_TokenizeShippingAddress = pcStrShippingAddress & "|" & pcStrShippingAddress2 & "|" & pcStrShippingCity & "|" & pcStrShippingStateCode & "|" & pcStrShippingPostalCode & "|" & pcStrShippingCountryCode

End Function
'// END:  Tokenize Shipping Address


'// START:  Validate Billing Address
Public Sub pcs_ValidateBillingAddress()

    Dim pcv_strAddressToken
    pcv_strAddressToken = getUserInput(Request("billingAddressToken"), 0)

    '// Validate anytime the saved billing address does not match what is submit.
    If (pcv_strAddressToken <> pcf_TokenizeBillingAddress()) Then

        If (ptaxAvalara = 1 AND ptaxAvalaraEnabled = 1 AND ptaxAvalaraAddressValidation = 1) OR USPS_AddressValidation = 1 Then

            Set billingAddress = Server.CreateObject("Scripting.Dictionary")
            billingAddress.Add "FirstName", pcStrBillingFirstName
            billingAddress.Add "LastName", pcStrBillingLastName
            billingAddress.Add "Company", pcStrBillingCompany
            billingAddress.Add "Address", pcStrBillingAddress
            billingAddress.Add "Address2", pcStrBillingAddress2
            billingAddress.Add "Country", pcStrBillingCountryCode
            billingAddress.Add "City", pcStrBillingCity
            billingAddress.Add "Region", pcStrBillingStateCode
            billingAddress.Add "PostalCode", pcStrBillingPostalCode
            billingAddress.Add "Phone", pcStrBillingPhone
            billingAddress.Add "Fax", pcStrBillingFax
            billingAddress.Add "Token", pcf_TokenizeBillingAddress()
            billingAddress.Add "Type", "B"  
            
            '// Keep Address in Session until they confirm
            Set Session("origAddress") = billingAddress
                     
            If USPS_AddressValidation = 1 Then            
                call USPS_validateAddress(billingAddress)
            Else
                call Avalara_validateAddress(billingAddress)
            End if

            '// Check if recommended address is the same...
            pcv_IsAddressMatch = False
            If (pcStrBillingCity  = Session("validAddress").Item("City")) And (pcStrBillingAddress  = Session("validAddress").Item("Address")) And (pcStrBillingAddress2= Session("validAddress").Item("Address2")) And (pcStrBillingCountryCode= Session("validAddress").Item("Country")) And (pcStrBillingStateCode= Session("validAddress").Item("Region")) And (pcStrBillingPostalCode= Session("validAddress").Item("PostalCode")) Then
                pcv_IsAddressMatch = True
            End If            
            Set billingAddress = nothing

            '// Use command "CHECK_ADDRESS" to stop and display confirmation dialog.
            'If (Session("validAddress").Item("Status") = "VALID") And (Not pcv_IsAddressMatch) Then
            If (Not pcv_IsAddressMatch) Then
                response.Clear()
                response.Write("CHECK_ADDRESS")
                response.End()
            End If
            
        End If
    End If
    
End Sub
'// END:  Validate Billing Address


'// START:  Validate Shipping Address
Public Sub pcs_ValidateShippingAddress()

    Dim pcv_strAddressToken
    pcv_strAddressToken = getUserInput(Request("shippingAddressToken"), 0)

    '// Validate anytime the saved shipping address does not match what is submit.
    If (pcv_strAddressToken <> pcf_TokenizeShippingAddress()) Then


        If pcShipOpt <> "-1" AND ((ptaxAvalara = 1 AND ptaxAvalaraEnabled = 1 AND ptaxAvalaraAddressValidation = 1) OR USPS_AddressValidation = 1) Then
    
            Set shippingAddress = Server.CreateObject("Scripting.Dictionary")
            shippingAddress.Add "FirstName", pcStrShippingFirstName
            shippingAddress.Add "LastName", pcStrShippingLastName
            shippingAddress.Add "Company", pcStrShippingCompany
            shippingAddress.Add "Address", pcStrShippingAddress
            shippingAddress.Add "Address2", pcStrShippingAddress2
            shippingAddress.Add "Country", pcStrShippingCountryCode
            shippingAddress.Add "City", pcStrShippingCity
            shippingAddress.Add "Region", pcStrShippingStateCode
            shippingAddress.Add "PostalCode", pcStrShippingPostalCode
            shippingAddress.Add "Phone", pcStrShippingPhone
            shippingAddress.Add "Fax", pcStrShippingFax
            shippingAddress.Add "Token", pcf_TokenizeShippingAddress()
            shippingAddress.Add "Type", "S"
            
            '// Keep Address in Session until they confirm
            Set Session("origAddress") = shippingAddress
             
            If USPS_AddressValidation = 1 Then
                Dim USPS_postdata, USPS_result, srvUSPSXmlHttp, objOutputXMLDoc
                call USPS_validateAddress(shippingAddress)
            Else
                call Avalara_validateAddress(shippingAddress)
            End if
            
            '// Check if recommended address is the same...
            pcv_IsAddressMatch = False
            If (pcStrShippingCity  = Session("validAddress").Item("City")) And (pcStrShippingAddress  = Session("validAddress").Item("Address")) And (pcStrShippingAddress2= Session("validAddress").Item("Address2")) And (pcStrShippingCountryCode= Session("validAddress").Item("Country")) And (pcStrShippingStateCode= Session("validAddress").Item("Region")) And (pcStrShippingPostalCode= Session("validAddress").Item("PostalCode")) Then
                pcv_IsAddressMatch = True
            End If            
            Set billingAddress = nothing

            '// Use command "CHECK_ADDRESS" to stop and display confirmation dialog.
            If (Session("validAddress").Item("Status") = "VALID") And (Not pcv_IsAddressMatch) Then
                response.Clear()
                response.Write("CHECK_ADDRESS")
                response.End()
            End If
            
        End If
    End If
    
End Sub
'// END:  Validate Shipping Address


'// START:  Avalara Validate Address
Function Avalara_validateAddress(pcAddressArray)

    Dim validAddress

	pcv_Address = pcAddressArray.Item("Address")
	pcv_Address2 = pcAddressArray.Item("Address2")
	pcv_Country = pcAddressArray.Item("Country")
	pcv_City = pcAddressArray.Item("City")
	pcv_Region = pcAddressArray.Item("Region")
	pcv_PostalCode = pcAddressArray.Item("PostalCode")

    pcv_Status = "INVALID"
    pcv_ErrorDesc = ""
    pcv_ReturnText = ""
    
	If pcv_Country <> "US" Then	
		pcv_Status = "INVALID"		
	Else
    
        strURL = ptaxAvalaraURL & "/1.0/address/validate?Line1=" & server.URLEncode(pcv_Address) & "&Line2=" & server.URLEncode(pcv_Address2) & "&City=" & server.URLEncode(pcv_City) & "&Region=" & server.URLEncode(pcv_Region) & "&PostalCode=" & server.URLEncode(pcv_PostalCode)
        
        Set srvAvalaraXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP" & scXML)
        srvAvalaraXmlHttp.open "GET", strURL, False
        srvAvalaraXmlHttp.SetRequestHeader "Content-Type", "text/xml"
        srvAvalaraXmlHttp.SetRequestHeader "Authorization", "Basic " & Base64_Encode(ptaxAvalaraAccount & ":" & ptaxAvalaraLicense)
        srvAvalaraXmlHttp.Send
        
        xmlResponse = srvAvalaraXmlHttp.responseText

        Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
        if xmlDoc.loadXML(xmlResponse) then
            Set result = xmlDoc.selectSingleNode("ValidateResult/ResultCode")
        end if
        
        if lcase(result.text) = "success" then
            Set validAddress = Server.CreateObject("Scripting.Dictionary")
            
            if Not xmlDoc.selectSingleNode("ValidateResult/Address/Line2") Is Nothing then
                pcv_Address = getUserInput(xmlDoc.selectSingleNode("ValidateResult/Address/Line2").text, 0)
                pcv_Address2 = getUserInput(xmlDoc.selectSingleNode("ValidateResult/Address/Line1").text, 0)
            else
                pcv_Address = getUserInput(xmlDoc.selectSingleNode("ValidateResult/Address/Line1").text, 0)
            end if            
            pcv_City = getUserInput(xmlDoc.selectSingleNode("ValidateResult/Address/City").text, 0)
            pcv_Region = getUserInput(xmlDoc.selectSingleNode("ValidateResult/Address/Region").text, 0)
            pcv_PostalCode = getUserInput(xmlDoc.selectSingleNode("ValidateResult/Address/PostalCode").text, 0)            
            pcv_Status = "VALID"
            
        else

            pcv_Address = ""
            pcv_Status = "INVALID"
            pcv_ErrorDesc = trim(xmlDoc.selectSingleNode("ValidateResult/Messages/Message/Summary").Text)
            
        end if
    
    End If

    '// Set Parsed Address to Session
    Set validAddress = Server.CreateObject("Scripting.Dictionary")
    
    validAddress.Add "Address", pcv_Address
    validAddress.Add "Address2", pcv_Address2
    validAddress.Add "Country", pcv_Country
    validAddress.Add "City", pcv_City
    validAddress.Add "Region", pcv_Region
    validAddress.Add "PostalCode", pcv_PostalCode
    validAddress.Add "Status", pcv_Status
    validAddress.Add "ErrorDesc", pcv_ErrorDesc
    validAddress.Add "ReturnText", pcv_ReturnText  
    validAddress.Add "Token", pcv_Address & "|" & pcv_Address2 & "|" & pcv_City & "|" & pcv_Region & "|" & pcv_PostalCode & "|" & pcv_Country
    validAddress.Add "Type", pcAddressArray.Item("Type") 
    
    Set Session("validAddress") = validAddress
    Set validAddress = Nothing
     
End Function
'// END:  Avalara Validate Address



'// START: USPS Validate Address
Function USPS_validateAddress(pcAddressArray)

    Dim objUSPSClass, validAddress

	pcv_Address = pcAddressArray.Item("Address")
	pcv_Address2 = pcAddressArray.Item("Address2")
	pcv_Country = pcAddressArray.Item("Country")
	pcv_City = pcAddressArray.Item("City")
	pcv_Region = pcAddressArray.Item("Region")
	pcv_PostalCode = pcAddressArray.Item("PostalCode")
    pcv_Status = "INVALID"
    pcv_ErrorDesc = ""
    pcv_ReturnText = ""

	If pcv_Country <> "US" Then	
		pcv_Status = "INVALID"		
	Else

        '// Validate Address
		Set objUSPSClass = New pcUSPSClass
		objUSPSClass.NewXMLTransaction "Verify", "AddressValidateRequest", USPS_Id
		strXMLClosingTag="AddressValidateRequest"
		
		objUSPSClass.WriteParent "Address", ""
			objUSPSClass.AddNewNode "Address1", pcv_Address, 1
			objUSPSClass.AddNewNode "Address2", pcv_Address2, 1
			objUSPSClass.AddNewNode "City", pcv_City, 1
			objUSPSClass.AddNewNode "State", pcv_Region, 1
			objUSPSClass.AddNewNode "Zip5", pcv_PostalCode, 1
			objUSPSClass.WriteEmptyParent "Zip4", "/"
		objUSPSClass.WriteParent "Address", "/"
		
		ObjUSPSClass.WriteParent strXMLClosingTag, "/"
		
		USPS_postdata=replace(USPS_postdata, "&XML", "andXML")
		USPS_postdata=replace(USPS_postdata, "&", "and")
		USPS_postdata=replace(USPS_postdata, "andamp;", "and")
		USPS_postdata=replace(USPS_postdata, "andXML", "&XML")
		
		call objUSPSClass.SendXMLRequest(USPS_postdata, USPS_AccessLicense)

		Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
        
        '// Parse Address
		If xmlDoc.loadXML(USPS_result) Then
			If xmlDoc.selectSingleNode("Error/Number") Is Nothing Then
				if xmlDoc.selectSingleNode("AddressValidateResponse/Address/Error/Number") Is Nothing then
					if xmlDoc.selectSingleNode("AddressValidateResponse/Address/Address1") Is Nothing then
						pcv_Address = getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/Address2").text, 0)
					else
						pcv_Address = getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/Address2").text, 0)
						pcv_Address2 = getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/Address1").text, 0)
					end if
					
					pcv_City = getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/City").text, 0)
					pcv_Region = getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/State").text, 0)
					pcv_PostalCode = getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/Zip5").text, 0) & "-" & getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/Zip4").text, 0)
					
					if Not xmlDoc.selectSingleNode("AddressValidateResponse/Address/ReturnText") is Nothing then
						pcv_ReturnText = getUserInput(xmlDoc.selectSingleNode("AddressValidateResponse/Address/ReturnText").text, 0)
					end if

				else					
					pcv_ErrorDesc = trim(xmlDoc.selectSingleNode("AddressValidateResponse/Address/Error/Description").Text)
				end if
				
				pcv_Status = "VALID"
			else
            
                pcv_Address = ""
				pcv_ErrorDesc = xmlDoc.selectSingleNode("AddressValidateResponse/Address/Error/Description").Text
				If len(pcv_ErrorDesc)=0 Then
                    pcv_ErrorDesc = xmlDoc.selectSingleNode("Error/Description").Text
                End If                
                pcv_Status = "INVALID"
                pcv_ErrorDesc = trim(pcv_ErrorDesc)
                if err.number<>0 then
                    pcv_Status = "VALID" '// Let pass if unhandled scenario.
                    error.clear
                end if                                
			end if
		end if
		
	End If

    '// Set Parsed Address to Session
    Set validAddress = Server.CreateObject("Scripting.Dictionary")
    
    validAddress.Add "Address", pcv_Address
    validAddress.Add "Address2", pcv_Address2
    validAddress.Add "Country", pcv_Country
    validAddress.Add "City", pcv_City
    validAddress.Add "Region", pcv_Region
    validAddress.Add "PostalCode", pcv_PostalCode
    validAddress.Add "Status", pcv_Status
    validAddress.Add "ErrorDesc", pcv_ErrorDesc
    validAddress.Add "ReturnText", pcv_ReturnText  
    validAddress.Add "Token", pcv_Address & "|" & pcv_Address2 & "|" & pcv_City & "|" & pcv_Region & "|" & pcv_PostalCode & "|" & pcv_Country
    validAddress.Add "Type", pcAddressArray.Item("Type") 
    
    Set Session("validAddress") = validAddress
    Set validAddress = Nothing
	
End Function
'// END: USPS Validate Address

Public Function pcf_GetShipServiceName(tmpService,dType)
Dim tmpSName,queryM,rsM

	pcv_boolShowFilteredRates = pcf_ShowFilteredRates()
	tmpSName=tmpService
											
	if (pcv_boolShowFilteredRates="1") AND (tmpSName<>"") then
		queryM="SELECT pcShippingMap.pcSM_Name FROM pcShippingMap INNER JOIN (pcSMRel INNER JOIN shipService ON pcSMRel.idshipservice=shipService.idshipservice) ON pcShippingMap.pcSM_ID=pcSMRel.pcSM_ID WHERE shipService.serviceDescription LIKE '%" & tmpSName & "%'"
		set rsM=connTemp.execute(queryM)
		if not rsM.eof then
			if dType="1" then
				tmpSName=rsM("pcSM_Name") & " (" & tmpSName & ")"
			else
				tmpSName=rsM("pcSM_Name")
			end if
		end if
		set rsM=nothing
	end if
	
	pcf_GetShipServiceName=tmpSName
End Function

Sub pcs_PreCalShipRates()

shipmentTotal=Cdbl(0)

'//UPS Variables
query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if not rs.eof then
	UPS_active=rs("active")
	UPS_userid=trim(rs("userID"))
	UPS_password=trim(rs("password"))
	UPS_license_key=trim(rs("AccessLicense"))
end if
set rs=nothing

'//CPS Variables
query="SELECT active, shipServer, userID, password, AccessLicense FROM ShipmentTypes WHERE idshipment=7;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if not rs.eof then
	CP_active=rs("active")
	CP_server=trim(rs("shipserver"))
	CP_username=trim(rs("userID"))
	if rs("password")<>"" then
		CP_password=trim(enDeCrypt(rs("password"), scCrypPass))
	end if
	CP_custNo=trim(rs("AccessLicense"))
end if

'// FedEX Variables WS
query="SELECT active, shipServer, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=9;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
If not rs.eof Then
	FedEXWS_server=trim(rs("shipserver"))
	FedEXWS_active=rs("active")
	FedEXWS_AccountNumber=trim(rs("userID"))
	FedEXWS_MeterNumber=trim(rs("password"))
	FEDEXWS_Environment=rs("AccessLicense")
End If
set rs=nothing

'//USPS Variables
query="SELECT active, shipServer, userID FROM ShipmentTypes WHERE idshipment=4;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if not rs.eof then
	usps_userid=trim(rs("userID"))
	usps_server=trim(rs("shipserver"))
	usps_active=rs("active")
end if
set rs=nothing

err.number=0

If PgType="2" then
	ErrPageName = "login.asp"
Else
	'// page name
	pcStrPageName = "estimateShipCost.asp"
	'//set error to zero
	pcv_intErr=0
End if


'// Ship Mapping Settings
pcv_boolShowFilteredRates = pcf_ShowFilteredRates()

If PgType="2" then
	'Retreive the saved shipping information from the customer sessions table
	query="SELECT pcCustomerSessions.idDbSession, pcCustomerSessions.randomKey, pcCustomerSessions.idCustomer, pcCustomerSessions.pcCustSession_BillingStateCode, pcCustomerSessions.pcCustSession_BillingAddress, pcCustomerSessions.pcCustSession_BillingCity, pcCustomerSessions.pcCustSession_BillingProvince, pcCustomerSessions.pcCustSession_BillingPostalCode, pcCustomerSessions.pcCustSession_BillingCountryCode, pcCustomerSessions.pcCustSession_CustomerEmail, pcCustomerSessions.pcCustSession_ShippingResidential, pcCustomerSessions.pcCustSession_ShippingAddress, pcCustomerSessions.pcCustSession_ShippingCity, pcCustomerSessions.pcCustSession_ShippingStateCode, pcCustomerSessions.pcCustSession_ShippingProvince, pcCustomerSessions.pcCustSession_ShippingPostalCode, pcCustomerSessions.pcCustSession_ShippingCountryCode FROM pcCustomerSessions WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&")) ORDER BY idDbSession DESC;"
	
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	pcBillingStateCode=rs("pcCustSession_BillingStateCode")
	pcBillingAddress=rs("pcCustSession_BillingAddress")
	pcBillingCity=rs("pcCustSession_BillingCity")
	pcBillingProvince=rs("pcCustSession_BillingProvince")
	pcBillingPostalCode=rs("pcCustSession_BillingPostalCode")
	if NOT isNull(pcBillingPostalCode) then
		pcBillingPostalCode=pcf_PostCodes(pcBillingPostalCode)
	end if
	pcBillingCountryCode=rs("pcCustSession_BillingCountryCode")
	pcCustomerEmail=rs("pcCustSession_CustomerEmail")
	pResidentialShipping=rs("pcCustSession_ShippingResidential")
	pcShippingAddress=rs("pcCustSession_ShippingAddress")
	pcShippingCity=rs("pcCustSession_ShippingCity")
	pcShippingStateCode=rs("pcCustSession_ShippingStateCode")
	pcShippingProvince=rs("pcCustSession_ShippingProvince")
	pcShippingPostalCode=rs("pcCustSession_ShippingPostalCode")
	if NOT isNull(pcShippingPostalCode) then
		pcShippingPostalCode=pcf_PostCodes(pcShippingPostalCode)
	end if
	pcShippingCountryCode=rs("pcCustSession_ShippingCountryCode")
	
	If Not len(pcShippingCity)>0 Then
		pcShippingCity = pcBillingCity
	End If
	If Not len(pcShippingStateCode)>0 Then
		pcShippingStateCode = pcBillingStateCode
	End If
	If Not len(pcShippingProvince)>0 Then
		pcShippingProvince = pcBillingProvince
	End If
	If Not len(pcShippingPostalCode)>0 Then
		pcShippingPostalCode = pcBillingPostalCode
	End If
	If Not len(pcShippingCountryCode)>0 Then
		pcShippingCountryCode = pcBillingCountryCode
	End If
	
	set rs=nothing
	
	
	'// Do you want the cart shippable total to include discounts?
	pcv_intIncludeDiscounts = 1

Else
	pResidentialShipping=request("residentialShipping")
    
	pcs_ValidateTextField "CountryCode",  pcv_isShipCountryCodeRequired , 4
	pcs_ValidateTextField "zip", pcv_isZipRequired, 10
	if request("ddjumpflag")="YES"then
	pcs_ValidateTextField "city", false, 30
	pcs_ValidateStateProvField  "StateCode", false, 4
	pcs_ValidateStateProvField  "Province", false, 50
	else
	pcs_ValidateTextField "city", pcv_isCityRequired, 30
	pcs_ValidateStateProvField  "StateCode", pcv_isShipStateCodeRequired, 4
	pcs_ValidateStateProvField  "Province", pcv_isShipProvinceCodeRequired, 50
	end if
	CountryCode=Session("pcSFCountryCode")
	StateCode=Session("pcSFStateCode")
	Province=Session("pcSFProvince")
	city=Session("pcSFcity")
	zip=Session("pcSFzip")
	
	If pcv_intErr>0 Then
		response.Clear()
		response.Write("")
		call closedb()
		response.end
	Else
		If DeliveryZip = "1" Then
			query="SELECT * from zipcodevalidation WHERE zipcode='" &zip& "';"
			set rsZipCodeObj=server.CreateObject("ADODB.RecordSet")
			set rsZipCodeObj=conntemp.execute(query)
			if rsZipCodeObj.eof then
				set rsZipCodeObj=nothing
				%>
				<form action="estimateShipCost.asp" id="ShipChargeForm" name="ShipChargeForm" data-target="#QuickViewDialog" method="get">
				<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_Custmoda_23")%></div>
				<button class="btn btn-default" type="button" data-ng-click="showEstShip()" name="SubmitShip" id="SubmitShip">Change Postal Code</button>
				</form>		
				<%
				call closedb()
				response.end
			end if
			set rsZipCodeObj=nothing
		End If
		  
		pcCartArray=Session("pcCartSession")
		ppcCartIndex=Session("pcCartIndex")
	End If
	
End if

If PgType="2" then
	pShipSubTotal=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
	pShipCDSubTotal=Cdbl(calculateCategoryDiscounts(pcCartArray, ppcCartIndex))
	if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
		'// Calculate Promo Price
		TotalPromotions=pcf_GetPromoTotal(Session("pcPromoSession"),Session("pcPromoIndex"))
	end if
	pSubTotal=trim(URLDecode(getUserInput(request.QueryString("pSubTotalCheckFreeShipping"),20)))
	if pSubTotal = "" or isNull(pSubTotal) then
		' Not coming from orderVerify.asp, so calculate normally
		if session("SF_DiscountTotal")="" then session("SF_DiscountTotal")=0
		if session("SF_RewardPointTotal")="" then session("SF_RewardPointTotal")=0
		pSubTotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
		pSubTotal=pSubTotal-pShipCDSubTotal-TotalPromotions-session("SF_DiscountTotal")-session("SF_RewardPointTotal")
		If pcv_intIncludeDiscounts=1 Then
			pShipSubTotal=pSubTotal
		End If
	else
		' Coming from orderVerify.asp, so overwrite pShipSubTotal
		' The sub total is updated on orderVerify.asp so we do not need to subtract anything
		pShipSubTotal=pSubTotal
	end if
Else
	' calculate total price of the order, total weight and product total quantities
	pSubTotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
	pShipCDSubTotal=Cdbl(calculateCategoryDiscounts(pcCartArray, ppcCartIndex))
	pSubTotal=pSubTotal-pShipCDSubTotal
	pShipSubTotal=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
End if

pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
intUniversalWeight=pShipWeight
pCartQuantity=Int(calculateCartQuantity(pcCartArray, ppcCartIndex))
pCartShipQuantity=Int(calculateCartShipQuantity(pcCartArray, ppcCartIndex))
pCartSurcharge=Cdbl(calculateTotalProductSurcharge(pcCartArray, ppcCartIndex))

If PgType="2" then
	' check if state was entered for Shipping Address (only if Canada/US)
	if pcShippingCountryCode<>"" then
		' use shipping codes
		If len(pcShippingStateCode)>0 Then
			pcShippingProvince=pcShippingStateCode
		end if
		Universal_destination_provOrState=pcShippingProvince
		Universal_destination_country=pcShippingCountryCode
		Universal_destination_postal=pcShippingPostalCode
		Universal_destination_city=pcShippingCity
		Universal_destination_address=pcShippingAddress
	else
		' use billing
		if pcBillingProvince="" then
			pcBillingProvince=pcBillingStateCode
		end if
		Universal_destination_provOrState=pcBillingProvince
		Universal_destination_country=pcBillingCountryCode
		Universal_destination_postal=pcBillingPostalCode
		Universal_destination_city=pcBillingCity
		Universal_destination_address=pcBillingAddress
	end if
Else
	if Province="" then
	  Province=StateCode
	end if
	
	Universal_destination_provOrState=Province
	Universal_destination_country=CountryCode
	if CountryCode<>"" then
	  session("DestinationCountry")=CountryCode
	end if
	if Universal_destination_country="" then
	  Universal_destination_country=session("DestinationCountry")
	end if
	Universal_destination_postal=zip
	Universal_destination_city=city
End if

' if customer use anotherState, insert a dummy state code to simplify SQL sentence
if Universal_destination_provOrState="" then
   Universal_destination_provOrState="**"
end if

shipcompany=scShipService

If pShipWeight="0" Then

	If PgType="2" then
		query="SELECT active FROM ShipmentTypes WHERE active<>0"
		set rs=connTemp.execute(query)
		if rs.eof then '// There are NO active dynamic shipping services
			pcv_NoDynamicShipping="1"
		end if
	End if

	query="SELECT idFlatShiptype,WQP FROM FlatShipTypes"
	set rsShpObj=server.CreateObject("ADODB.RecordSet")
	set rsShpObj=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsShpObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if rsShpObj.eof then
		If PgType="2" then
			call UpdateNullShipper("Yes")
		Else
			Session("nullShipper")="Yes"
		End if
			
		set rsShpObj=nothing	
		If PgType="2" then
			If pcv_NoDynamicShipping="1" Then
				response.Clear()
				Call SetContentType()
				call closeDb()
				response.write "OK|*|<div class='pcSuccessMessage'>" & dictLanguage.Item(Session("language")&"_opc_ship_1") & "</div>"
				response.end
			End If 
		End if
	else
		flagShp=0
		do until rsShpObj.eof
			intIdFlatShipType=rsShpObj("idFlatShiptype")
			pShpObjType=rsShpObj("WQP")
			select case pShpObjType
			case "Q"
				flagShp=1
			case "P"
				flagShp=1
			case "O"
				flagShp=1
			case "I"
				flagShp=1
			case "W"
				'do nothing
			end select
			rsShpObj.movenext
		loop
		set rsShpObj=nothing

		if flagShp=0 then
			If PgType="2" then
				call UpdateNullShipper("Yes")
			Else
				Session("nullShipper")="Yes"
			End if
		else
			If PgType="2" then
				call UpdateNullShipper("No")
			Else
				Session("nullShipper")="No"
			End if
		End if
	end if
Else
	If PgType="2" then
		call UpdateNullShipper("No")
	Else
		Session("nullShipper")="No"
	End if
End If

If pCartShipQuantity=0 then
	If PgType="2" then
		call UpdateNullShipper("Yes")
		response.Clear()
		Call SetContentType()
		response.write "OK|*|<div class='pcSuccessMessage'>" & dictLanguage.Item(Session("language")&"_opc_ship_1") & "</div>"
		response.end
	Else
		Session("nullShipper")="Yes"
	End if
end if

iShipmentTypeCnt=0

If PgType="1" then
	if session("provider")="" OR request("provider")<>"" then
	  session("provider")=request("provider")
	end if
	
	session("availableShipStr")=""
End if

End Sub

Sub pcs_ProcessShipMethods()
			
	tSCount=0
	
	CntFree=0
	DCnt=0
	FedExWSCnt=0 '// WS
	USPSCnt=0
	UPSCnt=0
	CPCnt=0
	CUSTOMCnt=0
	pcv_Default=0
	do until rs.eof
		serviceCode=rs("serviceCode")
		serviceFree=rs("serviceFree")
		serviceFreeOverAmt=rs("serviceFreeOverAmt")
		serviceHandlingFee=rs("serviceHandlingFee")
		serviceHandlingIntFee=rs("serviceHandlingIntFee")
		serviceShowHandlingFee=rs("serviceShowHandlingFee")
		serviceLimitation=rs("serviceLimitation")
		customerLimitation=0
		if serviceLimitation<>0 then
			if serviceLimitation=1 then
				if Universal_destination_country=scShipFromPostalCountry then
					customerLimitation=1
				end if
			end if
			if serviceLimitation=2 then
				if Universal_destination_country<>scShipFromPostalCountry then
					customerLimitation=1
				end if
			end if
			if serviceLimitation=3 then
				if ucase(trim(Universal_destination_country))<>"US" then
					customerLimitation=1
				else
					if ucase(trim(Universal_destination_provOrState))="AK" OR ucase(trim(Universal_destination_provOrState))="HI" OR ucase(trim(Universal_destination_provOrState))="AS" OR ucase(trim(Universal_destination_provOrState))="BVI" OR ucase(trim(Universal_destination_provOrState))="GU" OR ucase(trim(Universal_destination_provOrState))="MPI" OR ucase(trim(Universal_destination_provOrState))="MP" OR ucase(trim(Universal_destination_provOrState))="PR" OR ucase(trim(Universal_destination_provOrState))="VI" then
						customerLimitation=1
					end if
				end if
			end if
			if serviceLimitation=4 then
				if ucase(trim(Universal_destination_country))<>"US" then
					customerLimitation=1
				else
					if ucase(trim(Universal_destination_provOrState))<>"AK" AND ucase(trim(Universal_destination_provOrState))<>"HI" then
						customerLimitation=1
					end if
				end if
			end if
		end if

		if customerLimitation=0 then
			shipArray=split(availableShipStr,"|?|")
			for i=lbound(shipArray) to (Ubound(shipArray))
				shipDetailsArray=split(shipArray(i),"|")

				if ubound(shipDetailsArray)>0 then
					if shipDetailsArray(1)=serviceCode then
						tempRate=shipDetailsArray(3)
						if ubound(shipDetailsArray)>4 then
							pcvNegRate=shipDetailsArray(5)
							if ucase(shipDetailsArray(0))="UPS" then
								if pcv_UseNegotiatedRates=1 AND pcvNegRate<>"NONE"  then
									tempRate=pcvNegRate
								end if
							end if
						end if
						tempRate=(cDbl(tempRate)+cDbl(pCartSurcharge))

						tempRateDisplay=scCurSign&money(tempRate)
						If serviceShowHandlingFee="0" then
							tempRate=(cDbl(tempRate)+cDbl(serviceHandlingFee))
							tempRateDisplay=scCurSign&money(tempRate/1.2)
							serviceHandlingFee="0"
						End If
						if ((ucase(shipDetailsArray(0))=ucase(session("provider"))) OR (pcv_boolShowFilteredRates="1")) OR (pgType="2") then
						'if (ucase(shipDetailsArray(0))=ucase(session("provider"))) then
						
						tmpEstShipOpt=""
						if session("pcEstShipping")<>"" then
							tmpEstShipping=split(session("pcEstShipping"),",")
							tmpEstShipOpt=tmpEstShipping(5)
						end if
						
						IF PgType="2" THEN
							If serviceFree="-1" and Cdbl(pSubTotal)>Cdbl(serviceFreeOverAmt) then
								tempRate="0"
								tempRateDisplay= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_f")
								CntFree=CntFree+1
							End If
							pshipDetailsArray2= shipDetailsArray(2)
							pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&reg;</sup>","")
							pshipDetailsArray2= replace(pshipDetailsArray2,"&reg;","")
							pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>SM</sup>","")
							pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&#8482;</sup>","")
							pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&trade;</sup>","")
							
							if pcv_Default=0 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
								pcv_Default=1
							else
								x_checked="XCHECK"
							end if
							
							'get default classes
							serviceTypeClass = "pcShip_ServiceType"
							deliveryTimeClass = "pcShip_DeliveryTime"
							rateClass = "pcShip_Rate"
							
							'decide whether to hide delivery column
							showDeliveryCol = pcHideEstimateDeliveryTimes <> "-1"
							if not showDeliveryCol then
								serviceTypeClass = "pcShip_ServiceTypeL"
							end if
							
							'replace "NA" delivery time column
							if shipDetailsArray(4)="NA" then
								shipDetailsArray(4)="&nbsp;"
							end if

							select case ucase(shipDetailsArray(0))
							case "FEDEXWS"
								DCnt=DCnt+1
								FedExWSCnt=FedExWSCnt+1
								if FedExWSCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
									x_checked=" checked"
								else
									if FedExWSCnt=1 AND pcv_Default=0 then
										x_checked="FCHECK"
									end if
								end if
								strFEDEX=strFEDEX&"<div class='pcTableRow'>"
								tSCount=tSCount+1
								tSArr(tSCount,0)=shipDetailsArray(1)
								tSArr(tSCount,1)=tempRate
								tSArr(tSCount,2)=shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)
								tSArr(tSCount,3)=shipDetailsArray(0)
								tSArr(tSCount,4)=shipDetailsArray(2)
								tSArr(tSCount,5)=pcv_Default
								tSArr(tSCount,6)=shipDetailsArray(4)
								if (tmpEstShipOpt<>"") AND (tmpEstShipOpt=shipDetailsArray(1)) then
									x_checked=" checked"
								end if
								strFEDEX=strFEDEX&"<div class='" & serviceTypeClass & "'><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</div>"										
								if showDeliveryCol then
									strFEDEX=strFEDEX&"<div class='" & deliveryTimeClass & "'>" &shipDetailsArray(4)&"</div>"
								end if
								strFEDEX=strFEDEX&"<div class='" & rateClass & "'>"&tempRateDisplay&"</div>"
								strFEDEX=strFEDEX&"</div>"
							case "USPS"
								DCnt=DCnt+1
								USPSCnt=USPSCnt+1
								if USPSCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
									x_checked=" checked"
								else
									if USPSCnt=1 AND pcv_Default=0 then
										x_checked="FCHECK"
									end if
								end if
								tSCount=tSCount+1
								tSArr(tSCount,0)=shipDetailsArray(1)
								tSArr(tSCount,1)=tempRate
								tSArr(tSCount,2)=shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)
								tSArr(tSCount,3)=shipDetailsArray(0)
								tSArr(tSCount,4)=shipDetailsArray(2)
								tSArr(tSCount,5)=pcv_Default
								tSArr(tSCount,6)=shipDetailsArray(4)
								if (tmpEstShipOpt<>"") AND (tmpEstShipOpt=shipDetailsArray(1)) then
									x_checked=" checked"
								end if
								strUSPS=strUSPS&"<div class='pcTableRow'>"
								strUSPS=strUSPS&"<div class='" & serviceTypeClass & "'><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</div>"
								if showDeliveryCol then
									strUSPS=strUSPS&"<div class='" & deliveryTimeClass & "'>"&shipDetailsArray(4)&"</div>"
								end if
								strUSPS=strUSPS&"<div class='" & rateClass & "'>"&tempRateDisplay&"</div>"
								strUSPS=strUSPS&"</div>"
							case "UPS"
								DCnt=DCnt+1
								UPSCnt=UPSCnt+1
								if UPSCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
									x_checked=" checked"
								else
									if UPSCnt=1 AND pcv_Default=0 then
										x_checked="FCHECK"
									end if
								end if
								tSCount=tSCount+1
								tSArr(tSCount,0)=shipDetailsArray(1)
								tSArr(tSCount,1)=tempRate
								tSArr(tSCount,2)=shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)
								tSArr(tSCount,3)=shipDetailsArray(0)
								tSArr(tSCount,4)=shipDetailsArray(2)
								tSArr(tSCount,5)=pcv_Default
								tSArr(tSCount,6)=shipDetailsArray(4)
								if (tmpEstShipOpt<>"") AND (tmpEstShipOpt=shipDetailsArray(1)) then
									x_checked=" checked"
								end if
								strUPS=strUPS&"<div class='pcTableRow'>"
								strUPS=strUPS&"<div class='" & serviceTypeClass & "'><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</div>"
								if showDeliveryCol then
									strUPS=strUPS&"<div class='" & deliveryTimeClass & "'>"&shipDetailsArray(4)&"</div>"
								end if
								strUPS=strUPS&"<div class='" & rateClass & "'>"&tempRateDisplay&"</div>"
								strUPS=strUPS&"</div>"
							case "CP"
								DCnt=DCnt+1
								CPCnt=CPCnt+1
								if CPCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
									x_checked=" checked"
								else
									if CPCnt=1 AND pcv_Default=0 then
										x_checked="FCHECK"
									end if
								end if
								tSCount=tSCount+1
								tSArr(tSCount,0)=shipDetailsArray(1)
								tSArr(tSCount,1)=tempRate
								tSArr(tSCount,2)=shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)
								tSArr(tSCount,3)=shipDetailsArray(0)
								tSArr(tSCount,4)=shipDetailsArray(2)
								tSArr(tSCount,5)=pcv_Default
								tSArr(tSCount,6)=shipDetailsArray(4)
								if (tmpEstShipOpt<>"") AND (tmpEstShipOpt=shipDetailsArray(1)) then
									x_checked=" checked"
								end if
								strCP=strCP&"<div class='pcTableRow'>"
								strCP=strCP&"<div class='" & serviceTypeClass & "'><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</div>"
								if showDeliveryCol then
									strCP=strCP&"<div class='" & deliveryTimeClass & "'>"&shipDetailsArray(4)&"</div>"
								end if
								strCP=strCP&"<div class='" & rateClass & "'>"&tempRateDisplay&"</div>"
								strCP=strCP&"</div>"
							case "CUSTOM"
								DCnt=DCnt+1
								CUSTOMCnt=CUSTOMCnt+1
								if CUSTOMCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
									x_checked=" checked"
								else
									if CUSTOMCnt=1 AND pcv_Default=0 then
										x_checked="FCHECK"
									end if
								end if
								tSCount=tSCount+1
								tSArr(tSCount,0)=shipDetailsArray(1)
								tSArr(tSCount,1)=tempRate
								tSArr(tSCount,2)=shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)
								tSArr(tSCount,3)=shipDetailsArray(0)
								tSArr(tSCount,4)=shipDetailsArray(2)
								tSArr(tSCount,5)=pcv_Default
								tSArr(tSCount,6)=shipDetailsArray(4)
								if (tmpEstShipOpt<>"") AND (tmpEstShipOpt=shipDetailsArray(1)) then
									x_checked=" checked"
								end if
								'DA - Edit - Set Free shipping if needed on bundle or array
								if Session("daBunArrFreeShip") then
									if shipDetailsArray(1) = "C30" then
										tempRate=0
										tempRateDisplay="Free"
										daDelPreNoon = " - Pre 12 / Morning Delivery"
									else
										daDelPreNoon = ""
									end if
								End if
								strCUSTOM=strCUSTOM&"<div class='pcTableRow'>"
								strCUSTOM=strCUSTOM&"<div class='" & serviceTypeClass & "'><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder opcShipRadio'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</div>"
								if showDeliveryCol then
									'strCUSTOM=strCUSTOM&"<div class='" & deliveryTimeClass & "'>"&shipDetailsArray(4)&"</div>"
									'DA - Edit - Show delivery estimate instead
									strCUSTOM=strCUSTOM&"<div class='" & deliveryTimeClass & "'>"&daFunDelDateReturn(Session("daNumPCOrder"),CInt(shipDetailsArray(4)))&daDelPreNoon&"</div>"
								end if

								strCUSTOM=strCUSTOM&"<div class='" & rateClass & "'>"&tempRateDisplay&"</div>"
								strCUSTOM=strCUSTOM&"</div>"
							end select
						ELSE 'pgType
							If serviceFree="-1" and Cdbl(pSubTotal)>Cdbl(serviceFreeOverAmt) then
								tempRate="0"
								tempRateDisplay= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_f")
								CntFree=CntFree+1
							  End If
							  DCnt=DCnt+1%>
							  <% 
								pshipDetailsArray2= shipDetailsArray(2)
								pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&reg;</sup>","")
								pshipDetailsArray2= replace(pshipDetailsArray2,"&reg;","")
								pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>SM</sup>","")
								pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&#8482;</sup>","")
								pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&trade;</sup>","")
							  tSCount=tSCount+1
							  tSArr(tSCount,0)=shipDetailsArray(1)
							  tSArr(tSCount,1)=tempRate
							  tSArr(tSCount,2)=shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)
							  tSArr(tSCount,3)=shipDetailsArray(0)
							  tSArr(tSCount,4)=shipDetailsArray(2)
							  tSArr(tSCount,5)=0
							  tSArr(tSCount,6)=shipDetailsArray(4)
							  if (tmpEstShipOpt<>"") AND (tmpEstShipOpt=shipDetailsArray(1)) then
									x_checked=" checked"
								end if
							  if (pcv_boolShowFilteredRates<>"1") then%>
							  <div class="row">
								<div class="<%= col_ServiceTypeClass %>">
									<label class="radio-inline" for="pcShipSelection">
										<input type="radio" name="Shipping" value="<%=tSArr(tSCount,2)%>" <%if tSArr(tSCount,2)=session("pcEstShipping") then%>checked<%end if%> class="clearBorder" data-ng-click="updateShippingMethod()"><%=shipDetailsArray(2) %>
									</label>
								</div>
								<div class="col-xs-5"></div>
								<div class="<%= col_RateClass %>"><%=tempRateDisplay%></div>
							  </div>
							  <%end if%>
						<%END IF 'pgType
						end if 'End If
					end if
				end if
			next
								
			tempRate=""
			tempRateDisplay=""
		end if
		rs.movenext
	loop
	set rs=nothing

End Sub

Sub pcs_MapShip()
	HasDefaultM=0
	queryQ="SELECT 0,pcSM_ID,pcSM_Name,'',0,pcSM_Type,0,'' FROM pcShippingMap ORDER BY pcSM_Order ASC, pcSM_Name ASC;"
	set rsQ=connTemp.execute(queryQ)
	if not rsQ.eof then
		tMArr=rsQ.getRows()
		set rsQ=nothing
		MCount=ubound(tMArr,2)
		For iM=0 to MCount
			queryQ="SELECT shipService.serviceCode FROM shipService INNER JOIN pcSMRel ON shipService.idshipservice=pcSMRel.idshipservice WHERE (shipService.serviceActive=-1) AND (pcSMRel.pcSM_ID=" & tMArr(1,iM) & ");"
			set rsQ=connTemp.execute(queryQ)
			if not rsQ.eof then
				tSeArr=rsQ.getRows()
				set rsQ=nothing
				tSeCount=ubound(tSeArr,2)
				For iSe=0 to tSeCount
					For iAv=1 to tSCount
						if (tSArr(iAv,0)=tSeArr(0,iSe)) AND (tSArr(iAv,0)<>"") then
							if ((tMArr(5,iM)="1") AND (Cdbl(tSArr(iAv,1))>Cdbl(tMArr(4,iM)))) OR ((tMArr(5,iM)="0") AND ((Cdbl(tSArr(iAv,1))<Cdbl(tMArr(4,iM))) OR (Cdbl(tMArr(4,iM))=0))) then
								tMArr(0,iM)=1
								tMArr(3,iM)=tSArr(iAv,2)
								tMArr(4,iM)=tSArr(iAv,1)
								tMArr(7,iM)=tSArr(iAv,6)
								if (tSArr(iAv,5)="1") AND (HasDefaultM=0) then
									tMArr(6,iM)=1
									HasDefaultM=1
								else
									tMArr(6,iM)=0
								end if
								exit for
							end if
						end if
					Next
				Next
			end if
			set rsQ=nothing
		Next
	end if
	set rsQ=nothing
End Sub
%>