<%@  language="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "service.asp"
' This page outputs a JSON representation of the shopping cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp" -->
<!--#include file="../includes/dateinc.asp" -->
<!--#include file="inc_sb.asp"-->

<!--#include file="service_init.asp"-->

<% 
response.Clear()
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

Dim subtotal, totalDeliveringTime, totalQuantity, pRowWeight, totalRowWeight	
totalQuantity = Cint(0)
subtotal = Cint(0)
totalDeliveringTime = Cint(0)
HaveGcsTest = Cint(0)
PrdCanGW = Cint(0)

dim pcCartArray, paymentTotal
paymentTotal = Cint(0)

dim jsonService : set jsonService = JSON.parse("{}")
redim shoppingcartrows(pcCartIndex - 1)



'////////////////////////////////////////////////////////////////////
'// START: LOAD CUSTOMER SESSION
'////////////////////////////////////////////////////////////////////

If Not (session("pcSFIdDbSession") = "" OR session("pcSFRandomKey") = "") Then

    query="SELECT customers.name, customers.lastName, customers.customerCompany, customers.email, customers.phone, customers.fax, customers.address, customers.address2, customers.zip, customers.stateCode, customers.state, customers.city, customers.countryCode,  customers.pcCust_VATID, customers.pcCust_SSN,customers.iRewardPointsAccrued, customers.iRewardPointsUsed, pcCustSession_ShippingFirstName, pcCustSession_ShippingLastName, pcCustSession_ShippingCompany, pcCustSession_ShippingAddress, pcCustSession_ShippingAddress2, pcCustSession_ShippingCity, pcCustSession_ShippingStateCode, pcCustSession_ShippingProvince, pcCustSession_ShippingPostalCode, pcCustSession_ShippingCountryCode, pcCustSession_ShippingPhone, pcCustSession_ShippingNickName, pcCustSession_TaxShippingAlone, pcCustSession_TaxShippingAndHandlingTogether, pcCustSession_TaxLocation, pcCustSession_TaxProductAmount, pcCustSession_OrdPackageNumber, pcCustSession_ShippingArray, pcCustSession_ShippingResidential, pcCustSession_IdPayment, pcCustSession_Comment, pcCustSession_discountcode, pcCustSession_discountAmount, pcCustSession_UseRewards, pcCustSession_RewardsBalance, pcCustSession_RewardsDollarValue, pcCustSession_NullShipper,pcCustSession_NullShipRates,pcCustSession_TF1,pcCustSession_DF1,pcCustSession_OrderName,pcCustSession_ShowShipAddr,pcCustSession_ShippingEmail,pcCustSession_ShippingFax, pcCustSession_SB_taxAmount, pcCustSession_taxAmount, pcCustSession_VATTotal, pcCustSession_CartRewards, pcCustSession_GWTotal, pcCustSession_GCDetails, pcCustSession_CatDiscTotal, pcCustSession_total, pcCustSession_DiscountCodeTotal "
    
    query = query & "FROM pcCustomerSessions "
    query = query & "INNER JOIN customers ON pcCustomerSessions.idCustomer = customers.idcustomer "
    query = query & "WHERE (pcCustomerSessions.idDbSession=" & session("pcSFIdDbSession") & ") "
    query = query & "AND (pcCustomerSessions.randomKey=" & session("pcSFRandomKey") & ") "
    If session("idCustomer") > 0 Then
        query = query & "AND (pcCustomerSessions.idCustomer=" & session("idCustomer") & ") "
    End If
    query = query & "ORDER BY pcCustomerSessions.idDbSession DESC;"
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=connTemp.execute(query)
    If Not rs.Eof Then
    
        IsCartSaved = true
        
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
        pcStrBillingSSN=rs("pcCust_SSN")
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
        If Not IsNull(rs("pcCustSession_TaxLocation")) Then
            ptaxLoc=Cdbl(rs("pcCustSession_TaxLocation"))
        End If
        If Not IsNull(rs("pcCustSession_TaxProductAmount")) Then
            ptaxPrdAmount =ccur(rs("pcCustSession_TaxProductAmount"))
        End If
        pcIntOrdPackageNumber=rs("pcCustSession_OrdPackageNumber")
        pcShippingArray=rs("pcCustSession_ShippingArray")
        pOrdShipType=rs("pcCustSession_ShippingResidential")
        pcIdPayment=rs("pcCustSession_IdPayment")
        savOrderComments=rs("pcCustSession_Comment")
        savdiscountcode=rs("pcCustSession_discountcode")
        savdiscountcodeamount=rs("pcCustSession_discountAmount")
        pSFDiscountCodeTotal=rs("pcCustSession_DiscountCodeTotal")
        savUseRewards=rs("pcCustSession_UseRewards")
        savNullShipper=rs("pcCustSession_NullShipper")
        savNullShipRates=rs("pcCustSession_NullShipRates")
        savTF1=rs("pcCustSession_TF1")
        savDF1=rs("pcCustSession_DF1")
        savOrderNickName=rs("pcCustSession_OrderName")
        pcShowShipAddr=rs("pcCustSession_ShowShipAddr")
        pcStrShippingEmail=rs("pcCustSession_ShippingEmail")
        pcStrShippingFax=rs("pcCustSession_ShippingFax")
        pcSFCartRewards = rs("pcCustSession_CartRewards")
        iDollarValue = rs("pcCustSession_RewardsDollarValue") 
        pTaxAmount = rs("pcCustSession_taxAmount")
        VATTotal = rs("pcCustSession_VATTotal") 
        pidPayment = rs("pcCustSession_IdPayment")
        GWTotal=rs("pcCustSession_GWTotal")       
        savGCs=rs("pcCustSession_GCDetails")
        CatDiscTotal = rs("pcCustSession_CatDiscTotal")
        cartTotal = rs("pcCustSession_total")
        if savGCs<>"" then
            GCArr=split(savGCs,"|g|")
            savGCs=""
            for y=0 to ubound(GCArr)
                if GCArr(y)<>"" then
                    GCInfo=split(GCArr(y),"|s|")
                    if savGCs<>"" then
                        savGCs=savGCs & ","
                    end if
                    savGCs=savGCs & GCInfo(0)
                end if
            next
            if savdiscountcode<>"" then
                if Right(savdiscountcode,1)<>"," then
                    savdiscountcode=savdiscountcode & ","
                end if
            end if
            savdiscountcode=savdiscountcode & savGCs
        end if	


        pcv_DebugVars = false
        
        If pcv_DebugVars = true Then
        
            response.Write("Billing First Name:  " & pcStrBillingFirstName & "<br />")
            response.Write("Billing Last Name:  " & pcStrBillingLastName & "<br />")
            response.Write("Billing Company:  " & pcStrBillingCompany & "<br />")
            response.Write("Billing Email:  " & pcStrBillingEmail & "<br />")
            response.Write("Billing Phone:  " & pcStrBillingPhone & "<br />")
            response.Write("Billing Fax:  " & pcStrBillingfax & "<br />")
            response.Write("Billing Address:  " & pcStrBillingAddress & "<br />")
            response.Write("Billing Address 2:  " & pcStrBillingAddress2 & "<br />")
            response.Write("Billing Post Code:  " & pcStrBillingPostalCode & "<br />")
            response.Write("Billing State Code:  " & pcStrBillingStateCode & "<br />")
            response.Write("Billing Province:  " & pcStrBillingProvince & "<br />")
            response.Write("Billing City:  " & pcStrBillingCity & "<br />")
            response.Write("Billing Country Code:  " & pcStrBillingCountryCode & "<br />")
            response.Write("Rewards Points Accrued:  " & pcIntRewardPointsAccrued & "<br />")
            response.Write("Rewards Points Used:  " & pcIntRewardPointsUsed & "<br />")
            response.Write("Shipping First Name:  " & pcStrShippingFirstName & "<br />")
            response.Write("Shipping Last Name:  " & pcStrShippingLastName & "<br />")
            response.Write("Shipping Company:  " & pcStrShippingCompany & "<br />")
            response.Write("Shipping Address:  " & pcStrShippingAddress & "<br />")
            response.Write("Shipping Address 2:  " & pcStrShippingAddress2 & "<br />")
            response.Write("Shipping City:  " & pcStrShippingCity & "<br />")
            response.Write("Shipping State Code:  " & pcStrShippingStateCode & "<br />")
            response.Write("Shipping Province:  " & pcStrShippingProvince & "<br />")
            response.Write("Shipping Postal Code:  " & pcStrShippingPostalCode & "<br />")
            response.Write("Shipping Country Code:  " & pcStrShippingCountryCode & "<br />")
            response.Write("Shipping Phone:  " & pcStrShippingPhone & "<br />")
            response.Write("Shipping Nickname:  " & pcStrShippingNickName & "<br />")
            response.Write("Shipping Tax Alone???:  " & TAX_SHIPPING_ALONE & "<br />")
            response.Write("Shipping Tax with Handling???:  " & TAX_SHIPPING_AND_HANDLING_TOGETHER & "<br />")
            response.Write("Tax Location:  " & ptaxLoc & "<br />")
            response.Write("Tax Prd Amount:  " & ptaxPrdAmount & "<br />")
            response.Write("Package Number:  " & pcIntOrdPackageNumber & "<br />")
            response.Write("Shipping ARray:  " & pcShippingArray & "<br />")
            response.Write("Ship Type:  " & pOrdShipType & "<br />")
            response.Write("Id Payment:  " & pcIdPayment & "<br />")
            response.Write("Order Comments:  " & savOrderComments & "<br />")
            response.Write("Discount Code:  " & savdiscountcode & "<br />")
            response.Write("User Rewards???:  " & savUseRewards & "<br />")
            response.Write("Null Shipper:  " & savNullShipper & "<br />")
            response.Write("Null Ship Rates:  " & savNullShipRates & "<br />")
            response.Write("TF1:  " & savTF1 & "<br />")
            response.Write("DF1:  " & savDF1 & "<br />")
            response.Write("Order Nickname:  " & savOrderNickName & "<br />")
            response.Write("Show Ship Address:  " & pcShowShipAddr & "<br />")
            response.Write("Shipping Email:  " & pcStrShippingEmail & "<br />")
            response.Write("Shipping Fax:  " & pcStrShippingFax & "<br />")
            response.Write("Save Gift Codes???:  " & savGCs & "<br />")
            
            response.End()

        End If

   
    Else
    
        'response.Write("")
        'response.End()		
    
    End If
    set rs=nothing

End If
if (session("pcEstShipping")<>"") then
	pcShippingArray=session("pcEstShipping")
end if

'////////////////////////////////////////////////////////////////////
'// END: LOAD CUSTOMER SESSION
'////////////////////////////////////////////////////////////////////







'////////////////////////////////////////////////////////////////////
'// START: LOAD CUSTOMER ADDRESSES
'////////////////////////////////////////////////////////////////////

If IsCartSaved Then

    If pcIntHideAddresses=0 Then

        showShippingAddress = true '//pcShowShipAddr = "1" AND session("gHideAddress") <> "1"
        
        
        jsonService.set "showAddresses", true
        if showShippingAddress then
            jsonService.set "showShippingAddresses", true
        else
            jsonService.set "showShippingAddresses", false
        end if


        '// Company Address
        dim companyAddress : set companyAddress = JSON.parse("{}")
        companyAddress.set "Name", scCompanyName
        companyAddress.set "Address", scCompanyAddress
        companyAddress.set "City", scCompanyCity
        companyAddress.set "State", scCompanyState
        companyAddress.set "Zip", scCompanyZip
        companyAddress.set "Country", scCompanyCountry
        jsonService.set "companyAddress", companyAddress

        
        if showShippingAddress then
            ' pcStrShippingFirstName & " " & pcStrShippingLastName
        end if


        '// Shipping Address
        dim shippingAddress : set shippingAddress = JSON.parse("{}")    
        
        shippingAddress.set "FirstName", pcStrShippingFirstName
        shippingAddress.set "LastName", pcStrShippingLastName
        shippingAddress.set "companyName", pcStrShippingCompany
        shippingAddress.set "address", pcStrShippingAddress
        shippingAddress.set "address2", pcStrShippingAddress2
        shippingAddress.set "city", pcStrShippingCity
				If pcStrShippingProvince="" Then
            shippingAddress.set "state", pcStrShippingStateCode 
            shippingAddress.set "province", "" 
        Else
            shippingAddress.set "state", "" 
            shippingAddress.set "province", pcStrShippingProvince 
        End If
        shippingAddress.set "postalCode", pcStrShippingPostalCode
        shippingAddress.set "country", pcStrShippingCountryCode
		'shippingAddress.set "token", pcStrShippingAddress & "|" & pcStrShippingAddress2 & "|" & pcStrShippingCity & "|" & pcStrShippingStateCode & "|" & pcStrShippingPostalCode & "|" & pcStrShippingCountryCode
        jsonService.set "shippingAddress", shippingAddress


        '// Billing Address
        dim billingAddress : set billingAddress = JSON.parse("{}")  
        billingAddress.set "FirstName", pcStrBillingFirstName
        billingAddress.set "LastName", pcStrBillingLastName
        billingAddress.set "company", pcStrBillingCompany
        billingAddress.set "email", pcStrBillingEmail
        billingAddress.set "phone", pcStrBillingPhone
        billingAddress.set "fax", pcStrBillingFax
        billingAddress.set "address", pcStrBillingAddress
        billingAddress.set "address2", pcStrBillingAddress2
        billingAddress.set "city", pcStrBillingCity
		If pcStrBillingProvince="" Then
            billingAddress.set "state", pcStrBillingStateCode 
            billingAddress.set "province", "" 
        Else
            billingAddress.set "province", pcStrBillingProvince
            billingAddress.set "state", "" 
        End If
        billingAddress.set "postalCode", pcStrBillingPostalCode
        billingAddress.set "country", pcStrBillingCountryCode
        billingAddress.set "VATID", pcStrBillingVATID
        billingAddress.set "SSN", pcStrBillingSSN
		'billingAddress.set "token", pcStrBillingAddress & "|" & pcStrBillingAddress2 & "|" & pcStrBillingCity & "|" & pcStrBillingStateCode & "|" & pcStrBillingPostalCode & "|" & pcStrBillingCountryCode
        jsonService.set "billingAddress", billingAddress


    End If '// If pcIntHideAddresses=0 Then
    
    
    
		' ------------------------------------------------------
		'Start SDBA - Notify Drop-Shipping
		' ------------------------------------------------------
    If scShipNotifySeparate="1" And pcCartIndex>1 Then
			
        tmp_showmsg = 0

        For f = 1 To pcCartIndex

            tmp_idproduct = pcCartArray(f,0)
            
            query="SELECT pcProd_IsDropShipped FROM products WHERE idproduct=" & tmp_idproduct & " AND pcProd_IsDropShipped=1;"
		    Set rs = server.CreateObject("ADODB.RecordSet")
			Set rs = connTemp.execute(query)
            If Not rs.Eof Then
                tmp_showmsg = 1
				Exit For
            End If
            Set rs = Nothing
         
        Next

        If tmp_showmsg = 1 Then
            jsonService.set "IsDropShipping", true
        End If
			
    End If
		' ------------------------------------------------------
		'End SDBA - Notify Drop-Shipping
		' ------------------------------------------------------
    

    If (savOrderNickName<>"" AND savOrderNickName<>"No Name") or savDF1<>"" or len(savOrderComments)>3 then 
        jsonService.set "displayDivider", true
    End If 


    If savOrderNickName<>"" And savOrderNickName<>"No Name" Then
        jsonService.set "orderNickName", savOrderNickName
    End If 
    

    If savDF1<>"" Then
        jsonService.set "showDateFrmt", showDateFrmt(savDF1)
        If savTF1<>"" Then
            jsonService.set "savTF1", savTF1
        End If 
    End If


    If len(savOrderComments)>3 Then
        jsonService.set "orderComments", savOrderComments
    End If  

End If '// If Not (session("pcSFIdDbSession") = "" OR session("pcSFRandomKey") = "") Then

'////////////////////////////////////////////////////////////////////
'// END: LOAD CUSTOMER ADDRESSES
'////////////////////////////////////////////////////////////////////







 


'// Calculate Rows
subtotal = CalculateCartRows(pcCartIndex, pcCartArray)

'////////////////////////////////////////////////////////////////////
'// START: CART LOOP
'////////////////////////////////////////////////////////////////////
rowCount = 0
tmpStrList = 0
totalRowWeight=0 

'DA - Edit - Setup variables to work out products in the cart
daNumMonitors = 0
daNumStands = 0
daNumPC = 0

For f = 1 To pcCartIndex

    if pcCartArray(f,10)=0 then

                    
        Dim shoppingcartrow : set shoppingcartrow = JSON.parse("{}")
        
        shoppingcartrow.set "row", f
        shoppingcartrow.set "id", pcCartArray(f,0)
		
		'DA - Edit Loop through cart fields to work out what we have
		'Monitors Check
		if not InStr(pcCartArray(f,7), "MM-M") = 0 Then
			daNumMonitors = daNumMonitors + pcCartArray(f,2)
		end if
		'Stand Check
		if not InStr(pcCartArray(f,7), "MM-S") = 0 Then
			daNumStands = daNumStands + pcCartArray(f,2)
		end if
 		'PC Checks
		if not InStr(pcCartArray(f,7), "MM-PRO1") = 0 Then
			daNumPC = daNumPC + pcCartArray(f,2)
		end if
		if not InStr(pcCartArray(f,7), "MM-ULT1") = 0 Then
			daNumPC = daNumPC + pcCartArray(f,2)
		end if
		if not InStr(pcCartArray(f,7), "MM-EXT1") = 0 Then
			daNumPC = daNumPC + pcCartArray(f,2)
		end if
		if not InStr(pcCartArray(f,7), "MM-TRA1") = 0 Then
			daNumPC = daNumPC + pcCartArray(f,2)
		end if
		if not InStr(pcCartArray(f,7), "MM-TRP1") = 0 Then
			daNumPC = daNumPC + pcCartArray(f,2)
		end if
		if not InStr(pcCartArray(f,7), "MM-CHA1") = 0 Then
			daNumPC = daNumPC + pcCartArray(f,2)
		end if
		        
        'APP-S
        if (pcCartArray(f,32)<>"") then
            pIdProduct = pcCartArray(f,43)
        else
			pIdProduct = pcCartArray(f,0)
        end if
        'APP-E
		
		shoppingcartrow.set "productID", pIdProduct
        shoppingcartrow.set "sku", replace(pcCartArray(f,7),"&quot;","""")
		
		pDescription = replace(pcCartArray(f,1),"&quot;","""")
        shoppingcartrow.set "description", pDescription
		
		'// Call SEO Routine
		pcGenerateSeoLinks
		
        shoppingcartrow.set "productURL", Server.HtmlEncode(pcStrPrdLink)
        
        If Not (pcCartArray(f,27)>0) Then
            rowCount = rowCount + 1
        End If 
        If rowCount mod 2 = 1 Then
            shoppingcartrow.set "rowClass", "odd"
        Else
            shoppingcartrow.set "rowClass", "even"
        End If
        
        'APP-S
        if (pcCartArray(f,32)<>"") then            
            shoppingcartrow.set "IsApparel", True
        else 
            shoppingcartrow.set "IsApparel", False
        end if
        'APP-E

        '/////////////////////////////////////////////////////////////////////////////////////
        '// START:  Set Default Values
        '/////////////////////////////////////////////////////////////////////////////////////
        if trim(pcCartArray(f,27))="" then
            pcCartArray(f,27)=0
        end if
        if trim(pcCartArray(f,28))="" then
            pcCartArray(f,28)=0
        end if
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END:  Set Default Values
        '/////////////////////////////////////////////////////////////////////////////////////


        '/////////////////////////////////////////////////////////////////////////////////////
        '// START:  Load Product Details
        '/////////////////////////////////////////////////////////////////////////////////////        
        'APP-S
        if (pcCartArray(f,32)<>"") then
            pcv_tmpPPrd1 = pcCartArray(f,43)
        else
            pcv_tmpPPrd1=pcCartArray(f,0)
        end if 
        'APP-E

        'APP-S
        query="SELECT serviceSpec,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcProd_multiQty,pcUrl FROM products WHERE idproduct=" & pcv_tmpPPrd1
        'APP-E
		
		
        set rs=server.CreateObject("ADODB.RecordSet")									
        set rs=connTemp.execute(query)
        If Not rs.Eof Then
            IsBTO=rs("serviceSpec")
            pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
            pcv_intQtyValidate=rs("pcprod_QtyValidate")
            pcv_lngMinimumQty=rs("pcprod_MinimumQty")
            pcv_lngMultiQty=rs("pcProd_multiQty")
		
			'DA - Edit
			shoppingcartrow.set "daproductURL", "/products/" & rs("pcUrl") & "/"

        End If        
        
        if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
            pcv_intHideBTOPrice="0"
        end if        
        
        if isNULL(pcv_intQtyValidate) OR pcv_intQtyValidate="" then
            pcv_intQtyValidate="0"
        end if				
        
        if isNULL(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
            pcv_lngMinimumQty="0"
        end if
        
        if isNULL(pcv_lngMultiQty) OR pcv_lngMultiQty="" then
            pcv_lngMultiQty="0"
        end if 
        set rs=nothing 


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
                pcv_tmpPPrdTemp=pcCartArray(f,43)
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
                                    
        if pcvStrSmallImage = "" or pcvStrSmallImage = "no_image.gif" then
            pcvStrSmallImage = "hide"
        end if        
        set rsImg = nothing


        pcv_FinalizedQuote=0
        if Instr(session("sf_FQuotes"),"****" & pcCartArray(f,0) & "****")>0 then
            pcv_FinalizedQuote=1
        end if

        '/////////////////////////////////////////////////////////////////////////////////////
        '// END:  Load Product Details
        '/////////////////////////////////////////////////////////////////////////////////////


        '/////////////////////////////////////////////////////////////////////////////////////
        '// START:  SubscriptionBridge
        '/////////////////////////////////////////////////////////////////////////////////////       
        If (pcCartArray(f,38)>0) then
            pcIsSubscription = True
			pSubscriptionID = (pcCartArray(f,38))
            %>
            <!--#include file="../includes/pcSBDataInc.asp" -->
            <%								
        End if 
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END:  SubscriptionBridge
        '/////////////////////////////////////////////////////////////////////////////////////




        '/////////////////////////////////////////////////////////////////////////////////////
        '// START: QUANTITY FIELDS
        '
        '   Notes:  Set Quantity field to transparent if it's a child (Cross Sell) or it's a Finalized Quote
        '
        '/////////////////////////////////////////////////////////////////////////////////////
        
        shoppingcartrow.set "quantity", cdbl(pcCartArray(f,2))
		shoppingcartrow.set "savequantity", cdbl(pcCartArray(f,2))
        
        totalQuantity = totalQuantity + pcCartArray(f,2)
        
        if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) OR (pcv_FinalizedQuote=1) OR (pcCartArray(f,38)>0) then
            'if CLng(aryBadItems(f-1))<>0 then
            '    response.Write("""quantityClass"":""transparentField redBackground"",") '// #fcc
            'else
                shoppingcartrow.set "quantityClass", "transparentField"
            'end if
            shoppingcartrow.set "quantityReadOnly", "true"
            shoppingcartrow.set "quantityValidate", "false"
            shoppingcartrow.set "MinimumQty", pcv_lngMinimumQty            
            if pcv_intQtyValidate<>"1" then
                shoppingcartrow.set "QtyValidate", 0
            else
                shoppingcartrow.set "QtyValidate", 1
            end if
            if session("Cust_IDEvent")<>"" then
                shoppingcartrow.set "QtyIDEvent", pcf_QtyIDEvent(pcCartArray(f,33))
            else
                shoppingcartrow.set "QtyIDEvent", 0
            end if
            shoppingcartrow.set "MultiQty", pcv_lngMultiQty
            
        else
            'if CLng(aryBadItems(f-1))<>0 then
            '    response.Write("""quantityClass"":""redBackground"",") '// #fcc
            'else
                shoppingcartrow.set "quantityClass", ""
            'end if
            shoppingcartrow.set "quantityReadOnly", false
            shoppingcartrow.set "quantityValidate", true
            shoppingcartrow.set "MinimumQty", pcv_lngMinimumQty            
            if pcv_intQtyValidate<>"1" then
                shoppingcartrow.set "QtyValidate", 0
            else
                shoppingcartrow.set "QtyValidate", 1
            end if
            if session("Cust_IDEvent")<>"" then
                shoppingcartrow.set "QtyIDEvent", pcf_QtyIDEvent(pcCartArray(f,33))
            else
                shoppingcartrow.set "QtyIDEvent", 0
            end if
            shoppingcartrow.set "MultiQty", pcv_lngMultiQty
            
        end if
    
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END: QUANTITY FIELDS
        '/////////////////////////////////////////////////////////////////////////////////////


        '/////////////////////////////////////////////////////////////////////////////////////
        '// START: GIFT WRAP
        '/////////////////////////////////////////////////////////////////////////////////////
        grCode = ""
        if Session("Cust_BuyGift")<>"" then

            query="select pcEv_Code from pcEvents where pcEv_IDEvent=" & session("Cust_IDEvent")
            set rsG=conntemp.execute(query)
            grCode=rsG("pcEv_Code")

        end if
        if len(grCode)>"" then            
            shoppingcartrow.set "geID", pcCartArray(f,33)
        end if
				
        
        If session("Cust_GW") = "1" Then

            PrdCanGWchecks=0
            
            query="SELECT pcGC_EOnly FROM pcGC WHERE pcGC_idproduct=" & pcCartArray(f,0)
			Set rs = connTemp.execute(query)
			If rs.Eof Then
			
                query="SELECT pcPE_IDProduct FROM pcProductsExc WHERE pcPE_IDProduct=" & pcCartArray(f,0)
				Set rs1 = connTemp.execute(query)
				If rs1.Eof Then
				    PrdCanGWchecks = 1
				Else 
				    PrdCanGWchecks = 0
				End If
                Set rs1 = Nothing

            End If
            Set rs = Nothing

            If PrdCanGWchecks=1 Then
                tmpStrList = 1
                shoppingcartrow.set "giftWrapStatus", true 
            Else
                shoppingcartrow.set "giftWrapStatus", false 
            End If

        End If

        PrdCanGWchecks=0

        query="SELECT pcGC_EOnly FROM pcGC WHERE pcGC_idproduct=" & pcCartArray(f,0)
        set rs1=connTemp.execute(query)
        if rs1.eof then
            query="SELECT pcPE_IDProduct FROM pcProductsExc WHERE pcPE_IDProduct=" & pcCartArray(f,0)
            set rs1=connTemp.execute(query)
            if rs1.eof then
                PrdCanGWchecks=1
            else 
                PrdCanGWchecks=0                
            end if
        end if
 
        if PrdCanGWchecks=1 then
        
            PrdCanGW = PrdCanGW + 1
									
            if (pcCartArray(f,34)<>"") and (pcCartArray(f,34)<>"0") then
                pcv_IsGiftWrapped = true
            else
                pcv_IsGiftWrapped = false
            end if		

        end if
        
        shoppingcartrow.set "IsGiftWrapped", pcv_IsGiftWrapped
        shoppingcartrow.set "PrdCanGWchecks", PrdCanGWchecks
        shoppingcartrow.set "totalGiftWrapProducts", PrdCanGW       
      
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END: GIFT WRAP
        '/////////////////////////////////////////////////////////////////////////////////////


        '/////////////////////////////////////////////////////////////////////////////////////
        '// START: IMAGE FIELDS
        '/////////////////////////////////////////////////////////////////////////////////////
        If Session("Cust_BuyGift")="" Then
            if pcvStrSmallImage = "hide" then
                pcv_imageURL = "no_image.gif"
            else
                pcv_imageURL = pcvStrSmallImage
            end if
        Else
            if pcvStrSmallImage = "hide" then
                pcv_imageURL = "no_image.gif"
            else
               pcv_imageURL = pcvStrSmallImage              
            end if
        End If	
        shoppingcartrow.set "ImageURL", pcv_imageURL
        shoppingcartrow.set "SmallImageWidth", pcIntSmImgWidth
        If pcvStrSmallImage = "hide" Then
            shoppingcartrow.set "ShowImage", false
        Else
            shoppingcartrow.set "ShowImage", true
        End if
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END: IMAGE FIELDS
        '/////////////////////////////////////////////////////////////////////////////////////




        '/////////////////////////////////////////////////////////////////////////////////////
        '// START: ITEM UNIT PRICE & ROW PRICE
        '/////////////////////////////////////////////////////////////////////////////////////

        'if pcv_IsEUMemberState=0 then
        '	tmpRowPrice=ccur( pcCartArray(f,2) * pcCartArray(f,17) )
        'end if

        if pcCartArray(f,20)=0 then
            pRowWeight=pcCartArray(f,2)*pcCartArray(f,6)
        else
            pRowWeight=0
        end if
        totalRowWeight = totalRowWeight + pRowWeight 
 

        RowUnit = pcCartArray(f,17)-ccur(ccur(pcCartArray(f,41))/pcCartArray(f,2))
        RowSubTotal = ccur(pcCartArray(f,2) * pcCartArray(f,17)) - ccur(pcCartArray(f,41))
        
        if pcv_intHideBTOPrice<>"1" then
            if pcCartArray(f,17) > 0 then
                shoppingcartrow.set "UnitPrice", scCurSign & money(RowUnit)
				shoppingcartrow.set "DAUnitPrice", scCurSign & money(RowUnit/1.2)
            end if
        end if

        if RowSubTotal > 0 then 
            shoppingcartrow.set "RowPrice", scCurSign & money(RowSubTotal) '//  pcf_CheckNumberRange(money(pExtRowPrice))
			shoppingcartrow.set "DARowPrice", scCurSign & money(RowSubTotal/1.2) '//  pcf_CheckNumberRange(money(pExtRowPrice))
        end if
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END: ITEM UNIT PRICE
        '/////////////////////////////////////////////////////////////////////////////////////


        '/////////////////////////////////////////////////////////////////////////////////////
        '// START: REMOVE BUTTON
        '
        ' NOTE:  Show Remove Button if it's NOT a Required child (Cross Sell)
        '
        '/////////////////////////////////////////////////////////////////////////////////////
        If (pcCartArray(f,12)<>"-2") Then 
            shoppingcartrow.set "IsRemoveable", true
        Else
            shoppingcartrow.set "IsRemoveable", false
        End If
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END: REMOVE BUTTON
        '/////////////////////////////////////////////////////////////////////////////////////


        '/////////////////////////////////////////////////////////////////////////////////////
        '// START: RECONFIGURE BUTTON
        '/////////////////////////////////////////////////////////////////////////////////////
        If (IsBTo=-1) And pcCartArray(f,16)="" And pcv_FinalizedQuote=0 Then

            queryQ="SELECT TOP 1 configProduct FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & ";"
            set rsQ=connTemp.execute(queryQ)
            If Not rsQ.eof Then 
                shoppingcartrow.set "IsReconfigurable", true
            Else
                shoppingcartrow.set "IsReconfigurable", false
            End If
            set rsQ=Nothing

        Else
            shoppingcartrow.set "IsReconfigurable", false
        End If
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END: RECONFIGURE BUTTON
        '/////////////////////////////////////////////////////////////////////////////////////


        '/////////////////////////////////////////////////////////////////////////////////////
        '// START: BTO PRODUCT DETAILS
        '/////////////////////////////////////////////////////////////////////////////////////      

        If trim(pcCartArray(f,16)) <> "" Then 

            query="SELECT stringProducts, stringValues, stringCategOries, stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
            set rs=server.CreateObject("ADoDB.RecOrdSet")
            set rs=conntemp.execute(query)

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
            set rs=Nothing 
            
            '// BTO Product Details
            If ArrProduct(0)="na" Then
                shoppingcartrow.set "BToConfigTitle", true
            Else 
                shoppingcartrow.set "BToConfigTitle", false

                Redim btoConfiguration(UBound(ArrProduct)-1) 
                 
                For i=lbound(ArrProduct) To (UBound(ArrProduct)-1)
                
                    myBToConfigPriceString = ""
                    
                    Dim btoLineItem : Set btoLineItem = JSON.parse("{}")
                    
                    If pcv_SpecialServer=1 Then
                      ArrValue(i)=replace(ArrValue(i),".",",")
                      ArrPrice(i)=replace(ArrPrice(i),".",",")
                    End If
                
                    'APP-S
					if pcv_HaveApparel=1 then
						tmpQ=", products.pcProd_ParentPrd"
					else
						tmpQ=""
					end if
					intParentPrd=ArrProduct(i)
                    query="SELECT products.description" & tmpQ & " FROM products WHERE products.idProduct="&ArrProduct(i)&";"
                    'APP-E
                    set rs=server.CreateObject("ADoDB.RecOrdSet")
                    set rs=conntemp.execute(query)
					
                    if not rs.eof then
						strDescription=rs("description")
						'APP-S
						if pcv_HaveApparel=1 then
							intParentPrd=rs("pcProd_ParentPrd")
							if intParentPrd>"0" then
							else
								intParentPrd=ArrProduct(i)
							end if
						end if
						'APP-E
					end if
					
                    set rs=Nothing
					
					'APP-S
					query="SELECT categories.categoryDesc FROM categories INNER JOIN Categories_Products ON categories.idCategory=Categories_Products.idCategory WHERE Categories_Products.idCategory=" & ArrCategory(i) & " AND Categories_Products.idProduct=" & intParentPrd & ";"
                    'APP-E
                    set rs=server.CreateObject("ADoDB.RecOrdSet")
                    set rs=conntemp.execute(query)
					
                    if not rs.eof then
						strCategoryDesc=rs("categoryDesc")
					end if
					set rs=nothing
					  
                    'APP-S
                    query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&intParentPrd&" and specProduct=" & pcCartArray(f,0) 
                    'APP-E
                    set rs=server.CreateObject("ADoDB.RecOrdSet")
                    set rs=conntemp.execute(query)
                                          
                    btDisplayQF=rs("displayQF")
                    set rs=Nothing
                      
                    'APP-S											
                    query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & intParentPrd & ";"
                    'APP-E	
                    set rsQ=connTemp.execute(query)
                    tmpMinQty=1
                    If Not rsQ.eof Then
                      tmpMinQty=rsQ("pcprod_minimumqty")
                      If IsNull(tmpMinQty) Or tmpMinQty="" Then
                        tmpMinQty=1
                      Else
                        If tmpMinQty="0" Then
                          tmpMinQty=1
                        End If
                      End If
                    End If
                    set rsQ=Nothing
                    tmpDefault=0
                    'APP-S
                    query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & " AND configProduct=" & intParentPrd & " AND cdefault<>0;"
                    'APP-E
                    set rsQ=connTemp.execute(query)
                    If Not rsQ.eof Then
                      tmpDefault=rsQ("cdefault")
                      If IsNull(tmpDefault) Or tmpDefault="" Then
                        tmpDefault=0
                      Else
                        If tmpDefault<>"0" Then
                          tmpDefault=1
                        End If
                      End If
                    End If
                    set rsQ=Nothing 

                  btoLineItem.set "BToConfigCatDescription", strCategoryDesc
                  
                  If btDisplayQF=True And clng(ArrQuantity(i))>1 Then
                    btoLineItem.set "BToConfigQuantity", ArrQuantity(i)
                  End If
                  
                    btoLineItem.set "BToConfigDescription", strDescription
                
                    If (ccur(ArrValue(i))<>0) Or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) And (tmpDefault=1)) Or ((ArrQuantity(i)-1<>0) And (tmpDefault=0))) And (ArrPrice(i)<>0)) Then
                      If (ArrQuantity(i)-clng(tmpMinQty))>=0 Then
                        If tmpDefault=1 Then
                          UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
                        Else
                          UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
                        End If
                      Else
                        UPrice=0
                      End If 
                  
                      myBToConfigPriceString = "" & ""&scCurSign & money(ccur((ArrValue(i)+UPrice)*pcCartArray(f,2)))&""
                    Else
                      If tmpDefault=1 Then 
                        myBToConfigPriceString = "" & ""&dictLanguage.Item(Session("language")&"_defaultNotice_1")&""
                      End If
                    End If
                    
					btoLineItem.set "BToConfigUnitPrice",scCurSign & money(ccur(ArrPrice(i)))
                    btoLineItem.set "BToConfigPrice", myBToConfigPriceString
                    
                    Set btoConfiguration(i) = btoLineItem
                    Set btoLineItem = Nothing
                  
                  Next
                  shoppingcartrow.Set "btoConfiguration", btoConfiguration
                  
                  
                
                  
                End If 

          End If 
        '/////////////////////////////////////////////////////////////////////////////////////
        '// END: BTO PRODUCT DETAILS
        '/////////////////////////////////////////////////////////////////////////////////////




        '// START 4th Row - Product Options
        Dim pcv_strOptionsArray, pcv_inToptionLoopSize, pcv_inToptionLoopCounter, tempPrice, tAprice
        Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelecteDoptions
        
        pcv_strOptionsArray = trim(pcCartArray(f,4))
                
        If len(pcv_strOptionsArray)>0 Then 

          '// Generate Our Local Arrays from our STored Arrays  
          
          ' Column 11) pcv_strSelecteDoptions '// Array of Individual Selected Options Id Numbers	
          pcArray_strSelecteDoptions = ""					
          pcArray_strSelecteDoptions = Split(trim(pcCartArray(f,11)),chr(124))
          
          ' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
          pcArray_strOptionsPrice = ""
          pcArray_strOptionsPrice = Split(trim(pcCartArray(f,25)),chr(124))
          
          ' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
          pcArray_strOptions = ""
          pcArray_strOptions = Split(trim(pcv_strOptionsArray),chr(124))
          
          ' Get Our Loop Size
          pcv_inToptionLoopSize = 0
          pcv_inToptionLoopSize = Ubound(pcArray_strSelecteDoptions)
                            
          ' Start in Position One
          pcv_inToptionLoopCounter = 0
            
          ' Display Our Options
          Redim productoptions(pcv_inToptionLoopSize)          
          For pcv_inToptionLoopCounter = 0 To pcv_inToptionLoopSize 

                Dim productoption : Set productoption = JSON.parse("{}")
                
                '// Display Our Options Prices - If any
                tempPrice = pcArray_strOptionsPrice(pcv_inToptionLoopCounter)
                If Not (tempPrice="" Or tempPrice=0) Then 
                    tAprice = (tempPrice*ccur(pcCartArray(f,2))) 
                    productoption.set "name", pcArray_strOptions(pcv_inToptionLoopCounter)
                    productoption.set "unitprice", scCurSign & money(tempPrice)
                    productoption.set "daunitprice", scCurSign & money(tempPrice/1.2)
                    productoption.set "price", scCurSign & money(tAprice)
                    productoption.set "daprice", scCurSign & money(tAprice/1.2)
                Else
                    tAprice = 0
                    productoption.set "name", pcArray_strOptions(pcv_inToptionLoopCounter)
                    'productoption.set "unitprice", scCurSign & money(tempPrice)
                    'productoption.set "price", scCurSign & money(tAprice)
                End If

                Set productoptions(pcv_inToptionLoopCounter) = productoption
                Set productoption = Nothing
          Next
          shoppingcartrow.Set "productoptions", productoptions

          '// If there are product options And Not GGG, show link To edit them
          If trim(pcCartArray(f,16))="" And Session("Cust_BuyGIft")="" Then  
                shoppingcartrow.set "HasProductOptions", true
          Else
                shoppingcartrow.set "HasProductOptions", false
          End If
          
        End If       
        '// END 4th Row - Product Options


        '// START 5th Row - Custom Input Fields
        'If trim(pcCartArray(f,21)) <> "" Then
        '    shoppingcartrow.set "xField", pcCartArray(f,21)
        'Else
        '    shoppingcartrow.set "xField", ""
        'End If
        customFields = ""
        If trim(pcCartArray(f,21)) <> "" Then
            xfdetails = pcCartArray(f,21)
            If len(xfdetails)>3 then
                xfarray=split(xfdetails,"<br>")
                Redim customFields(ubound(xfarray))   
                for q=lbound(xfarray) to ubound(xfarray)            
                    Dim customField : Set customField = JSON.parse("{}")            
                    customField.set "xField", replace(xfarray(q),"''","'")
                    Set customFields(q) = customField
                    Set customField = Nothing  
                    shoppingcartrow.set "editInputField", true          
                next
            End If 
        End If
        shoppingcartrow.Set "customFields", customFields


        '// If there are custom input fields And NO product options, And Not GGG, And Not BTO, show EDIT here
        If Not (len(pcv_strOptionsArray) = 0 And trim(pcCartArray(f,16)) = "" And Session("Cust_BuyGift") = "") Then
            shoppingcartrow.set "editInputField", false
        End If
        '// END 5th Row - Custom Input Fields


        '// START 6th Row - BTO Item Discounts
        If pcCartArray(f,16) <> "" Then
            If pcCartArray(f,30) <> 0 Then
                shoppingcartrow.set "itemDiscountRowTotal", scCurSign & "-" & money(pcCartArray(f,30)) '// money(ItemsDiscounts)
            End If                
        End If     
        '// END 6th Row - BTO Item Discounts


        '// START 7th Row - BTO Additional Charges
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

                Redim additionalCharges(UBound(ArrCProduct)-1)
                
                For i=lbound(ArrCProduct) To (UBound(ArrCProduct)-1)
                                            
                    dim additionalCharge : set additionalCharge = JSON.parse("{}")   
                 
                    If pcv_SpecialServer=1 Then
                        ArrCValue(i) = replace(ArrCValue(i), ".", ",")
                    End If

                    query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
                    Set rs = server.CreateObject("ADODB.RecordSet")
                    Set rs = conntemp.execute(query)
                    If Not rs.Eof Then                     
                        strCategoryDesc=rs("categoryDesc")
                        strDescription=rs("description")
                        
                        additionalCharge.set "categoryDesc", strCategoryDesc
                        additionalCharge.set "description", strDescription
                    End If
                    Set rs = Nothing

                    additionalCharge.set "total", scCurSign & money(ArrCValue(i))
                    
                    Set additionalCharges(i) = additionalCharge
                    Set additionalCharge = Nothing
                    
                Next
                
                shoppingcartrow.set "additionalCharges", additionalCharges
                
            End If '// If ArrCProduct(0)<>"na" Then
            
        End If  '// If trim(pcCartArray(f,16)) <> "" Then
        '// END 7th Row - BTO Additional Charges
        
        
        '// START 8th Row - Quantity Discounts
        If trim(pcCartArray(f,15)) <> "" And trim(pcCartArray(f,15)) > 0 Then
            shoppingcartrow.set "itemQuantityDiscountRowTotal", scCurSign & "-" & money(pcCartArray(f,15))
        End If      
        
        pcv_tmpID = pcCartArray(f,0)
        'APP-S
        If pcCartArray(f,32)<>"" Then
            pcv_tmpID=pcCartArray(f,43)
        
            query="SELECT discountPerUnit FROM discountsPerQuantity WHERE idProduct=" & pcv_tmpID & ";"
            set rsQ=connTemp.execute(query)
            if rsQ.eof then
                pcv_tmpID=pcCartArray(f,0)
            end if
            set rsQ=nothing
        end if  
        'APP-E        
        shoppingcartrow.set "itemQuantityDiscountRowID", pcv_tmpID
        '// END 8th Row - Quantity Discounts

 
        '// START 9th Row - Product Subtotal
        If RowSubTotal <> pcCartArray(f,42) Then
            shoppingcartrow.set "productSubTotal", scCurSign &  money(pcCartArray(f, 42))
        Else
            shoppingcartrow.set "productSubTotal", ""
        End If         
        '// END 9th Row - Product Subtotal  


        '// START 10th Row - Cross Sell Bundle Discount
        shoppingcartrow.set "ChildBundleID", pcCartArray(f, 8)        
         
        
        If (pcCartArray(f, 27) = -1) And (Not pcCartArray(f, 8)="") Then
            
            If (pcCartArray(f, 8)>0) Then
                shoppingcartrow.set "BundleTitle", pcCartArray(f, 1) & " + " & pcCartArray(f + 1, 1)
                shoppingcartrow.set "IsParentofBundle", true
            Else
                shoppingcartrow.set "IsParentofBundle", false
            End If
                 
        Else
            shoppingcartrow.set "IsParentofBundle", false
        End If

        If (pcCartArray(f,27)>0) And (pcCartArray(f,28)>0) Then
            shoppingcartrow.set "xSellBundleDiscount", scCurSign &  money( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2))
        End If
        '// END 10th Row - Cross Sell Bundle Discount


        '// START 11th Row - Cross Sell Bundle Subtotal
        If (pcCartArray(f,27)>0) And (pcCartArray(f,28)>0) Then
            shoppingcartrow.set "xSellBundleSubTotal", scCurSign &  money(pcCartArray(f, 40))
        End If
        '// START 11th Row - Cross Sell Bundle Subtotal
        
        
        
        '// START 12th Row - Gift Wrapping
        If Session("Cust_GW") = "1" Then

            GWmsg = ""
            'GWmsg = "" & dictLanguage.Item(Session("language")&"_orderverify_36a") & " "

            query="select pcPE_IDProduct from pcProductsExc where pcPE_IDProduct=" & pcCartArray(f,0)
            Set rsG = server.CreateObject("ADODB.RecordSet")
            Set rsG = connTemp.execute(query)
            If Not rsG.Eof Then    
                'GWmsg = GWmsg & dictLanguage.Item(Session("language")&"_orderverify_38a")
            Else
                
                If (pcCartArray(f,34)="") or (pcCartArray(f,34)="0") Then                
                    'GWmsg = GWmsg & dictLanguage.Item(Session("language")&"_orderverify_37a")                    
                Else

                    query="select pcGW_OptName,pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & pcCartArray(f,34)
                    set rsG2=server.CreateObject("ADODB.RecordSet")
                    set rsG2=connTemp.execute(query)
    
                    If Not rsG2.Eof Then
                    
                        pcv_strOptName = rsG2("pcGW_OptName")
                        pcv_strOptPrice = rsG2("pcGW_OptPrice")
                        GWmsg = GWmsg & pcv_strOptName & " - " & scCurSign & money(pcv_strOptPrice)
                        shoppingcartrow.set "giftWrapMessage", GWmsg
                        
                    End If
                    Set rsG2 = Nothing
                  
                End If
                    
            End If
            Set rsG = Nothing

        End If 
        
        
        '// START Delivery Time
        if Cint(pcCartArray(f,9))>totalDeliveringTime then
            totalDeliveringTime=Cint(pcCartArray(f,9))
        end if
        '// END Delivery Time


        '// START Gift Cert Check        
        query="select pcprod_Gc from Products where idproduct=" & pcCartArray(f,0) & " AND pcprod_Gc=1"
        Set rsGc = server.CreateObject("ADODB.RecordSet")
        Set rsGc = conntemp.execute(query)
        If Not rsGc.eof Then
            HaveGcsTest = 1
        End If
        Set rsGc = Nothing
        '// END Gift Cert Check   

        set shoppingcartrows(f-1) = shoppingcartrow
        
    end if '// if pcCartArray(f,10)=0 then

             
Next '// For f=1 To pcCartIndex

'////////////////////////////////////////////////////////////////////
'// END: CART LOOP
'////////////////////////////////////////////////////////////////////

'DA - Edit - Work out bundles, cables, shipping and cart messages    
daMsgBool = true
daMsgHeader = ""
daMsgBody = ""
daMsgUrl = ""
daMsgButText = ""
daCblBool = false
daDelCharge = 0
daBunDisc = 0
Session("daActualDiscount")=0
Session("daBunArrFreeShip")=False

if daNumStands = 0 then
'No stands, check for screens
	if daNumMonitors = 0 then
	'No stand or screens
		if daNumPC = 0 then
		'No stand, screens or PC, something is wrong!
		else
		'PC only cart
		'!!!! OPTION LINE 6 !!!!
		
			'Set cart message
			daMsgHeader = "Save up to £100 on your order with a bundle!"
			daMsgBody = "Add a Monitor Stand and some Screens to qualify for free delivery, free cables and get a bundle discount of up to £100 applied to your order."
			daMsgButText = "View monitor arrays here"
			daMsgUrl = "/display-systems/"
			
			'Set delivery charge
			daDelCharge=12
		
		end if
	else
	'No stand, Got Monitors
		if daNumPC = 0 then
		'Monitors only cart
		'!!!! OPTION LINE 4 !!!!
		
			'Set cart message
			daMsgHeader = "NEED A STAND TO HOLD YOUR SCREENS?"
			daMsgBody = "Add a Synergy Multi-Screen Monitor Stand to your order and get FREE 3m long digital video & power cables and qualify for FREE delivery."
			daMsgButText = "View Synergy Stands here"
			daMsgUrl = "/stands/"
			
			'Set delivery charge
			daDelCharge=12
		
		else
		'Monitors & PC Only cart
		'!!!! OPTION LINE 5 !!!!
		
			'Set cart message
			daMsgHeader = "NEED A STAND TO HOLD YOUR SCREENS?"
			daMsgBody = "Add a Synergy Multi-Screen Monitor Stand to your order and get FREE 3m long digital video & power cables, FREE delivery and a Bundle Discount!."
			daMsgButText = "View Synergy Stands here"
			daMsgUrl = "/stands/"
			
			'Set delivery charge
			daDelCharge=24
		
		end if
	end if
else
'Got stand
	if daNumMonitors = 0 then
	'Got stand but no screens
		if daNumPC = 0 then
		'Stand only cart
		'!!!! OPTION LINE 1 !!!!
		
			'Set cart message
			daMsgHeader = "NEED SOME MONITORS TO GO WITH YOUR STAND?"
			daMsgBody = "Add some screens to your order and qualify for FREE 3m long digital cables and FREE UK delivery!"
			daMsgButText = "View Monitors here"
			daMsgUrl = "/monitors/"
			
			'Set delivery charge
			daDelCharge=12
		
		else
		'Stand & PC only cart
		'!!!! OPTION LINE 2 !!!!
		
			'Set cart message
			daMsgHeader = "NEED SOME MONITORS TO GO WITH YOUR STAND & COMPUTER?"
			daMsgBody = "Add some screens to your order and qualify for FREE 3m long digital cables, FREE UK delivery and a Bundle Discount!"
			daMsgButText = "View Monitors here"
			daMsgUrl = "/monitors/"
			
			'Set delivery charge
			daDelCharge=24
		
		end if
	else
	'Got stand and screens
		if daNumPC = 0 then
		'Stand and monitor cart only
		'!!!! OPTION LINE 3 !!!!
		
			'Set cart message
			daMsgHeader = "NEED A COMPUTER TO RUN YOUR MULTI-SCREEN ARRAY?"
			daMsgBody = "Add a multi-screen capable PC to your order and make connecting your screen easy, qualify for a bundle discount as well! "
			daMsgButText = "View Computers here"
			daMsgUrl = "/computers/"
			
			'Set cables row
			daCblBool = true
			
			'Set delivery charge
			daDelCharge=0
			
			'Set session variable to get free shipping based on array purchase
			Session("daBunArrFreeShip")=True
		
		else
		'Got a full bundle
		'!!!! OPTION LINE 7 !!!!
		
			'Set cart message display off
			daMsgBool = false
			
			'Set cables row
			daCblBool = true
			
			'Set delivery charge
			daDelCharge=0
			
			'Set bundle discount
			select case daNumMonitors
				case 0
					daBunDisc = 0
 				case 1
					daBunDisc = 0
				case 2
					daBunDisc = 30
				case 3
					daBunDisc = 30
				case 4
					daBunDisc = 60
				case 5
					daBunDisc = 60
				case 6
					daBunDisc = 120
				case 8
					daBunDisc = 120
				case else
					daBunDisc = 120
			end select

			Session("daActualDiscount")=daBunDisc
			Session("daBunArrFreeShip")=True
			
		end if
	end if
end if

'Set session for PC in basket to use in shipping estimate calculation on OPC
Session("daNumPCOrder")=daNumPC

'Set cart messages
jsonService.set "daMsgBool", daMsgBool
if daMsgBool then
	jsonService.set "daMsgHeader", daMsgHeader
	jsonService.set "daMsgBody", daMsgBody
	jsonService.set "daMsgButText", daMsgButText
	jsonService.set "daMsgUrl", daMsgUrl
end if

'Set cables row
jsonService.set "daCblBool", daCblBool
if daCblBool then
	jsonService.set "daCblValue", scCurSign & money(daNumMonitors*15)
end if

'Set delivery charge
if daDelCharge = 0 then
	jsonService.set "daDelCharge", "FREE"
else
	jsonService.set "daDelCharge", scCurSign & money(daDelCharge/1.2)
end if

'Set bundle discount display settings
if daBunDisc > 0 then
	jsonService.set "daBundleDiscount", scCurSign & money(daBunDisc/1.2)
	jsonService.set "daBundleDiscountApplied", true
else
	jsonService.set "daBundleDiscountApplied", false
end if

'Test to set delivery estimate
jsonService.set "daFunDelDateBlockTest", daFunDelDateBlockTest(daNumPC,0)
jsonService.set "daDelCutOff", daFunDelCutOff()
jsonService.set "daDelDate", daFunDelDateReturn(daNumPC,0)

'// Timestamp
jsonService.set "date", showDateFrmt(Date())



'// Display Applied Product Promotions (if any)
TotalPromotions=0

If Session("pcPromoIndex")<>"" And Session("pcPromoIndex")>"0" Then

    PromoArr1=Session("pcPromoSession")
    PromoIndex=Session("pcPromoIndex")
    
    Redim promotions(PromoIndex-1)
    For m=1 to PromoIndex
       
        dim promotion : set promotion = JSON.parse("{}") 
        
        '// Show message and add to total if promotion discount is > 0
        If PromoArr1(m, 2) > 0 Then
        
            TotalPromotions = TotalPromotions + cdbl(PromoArr1(m, 2))
            promotion.set "name", PromoArr1(m, 1)
            promotion.set "price", scCurSign &  "-" & money(cdbl(PromoArr1(m, 2)))
            
            Set promotions(m-1) = promotion
            Set promotion = Nothing
            
        End If

    Next
    jsonService.set "promotions", promotions
    
End If



'// Category Discounts
Dim CatDiscTotal
CatDiscTotal = calculateCategoryDiscountTotal(pcCartIndex, pcCartArray)
If CatDiscTotal > 0 Then 
    jsonService.set "categoryDiscountTotal", scCurSign & "-" & money(CatDiscTotal)
End If



'// Promotions



'// Payment Total
If Session("customerType")=1 Then
    If Len(pPaymentDesc)>0 Then
        jsonService.set "paymentDescription", pPaymentDesc
    End If
End If



paymentTotal = calculatePaymentGatewayFees(pidPayment, pcIsSubscription)
If paymentTotal > "0" Then
    jsonService.set "paymentTotal", scCurSign &  money(paymentTotal)     
End If



'// Discounts when not logged in...
If (savdiscountcode="") And (Not IsCartSaved) Then
	If session("pcSFCust_FromCart")="1" Then
		jsonService.set "haveDiscounts",false
		savdiscountcode=session("pcSFCust_discountcode")
		If savdiscountcode<>"" Then
			jsonService.set "haveDiscounts",true
			If Right(savdiscountcode,1)<>"," Then
				savdiscountcode=savdiscountcode & ","
			End If
		End If
		savdiscountcodeamount=session("pcSFCust_discountAmount")
	Else
		jsonService.set "haveDiscounts",false
	End If
Else
	If (savdiscountcode<>"") Then
		jsonService.set "haveDiscounts",true
	End if
End If

daOPCDiscountTotal = 0

'// Discount Row
If savdiscountcode <> "" Then

    DiscountCodeArry=Split(savdiscountcode,",")
    intCodeCnt=ubound(DiscountCodeArry)
   
    DiscountAmountArry=Split(savdiscountcodeamount,",")
    intCodeCnt=ubound(DiscountAmountArry)
    
    
    Redim discounts(intCodeCnt)
    For i=0 to intCodeCnt

        If DiscountAmountArry(i) > 0 Then
        
            dim discount : set discount = JSON.parse("{}") 
            
            '// Name
            discount.set "name", session("DiscountDesc" & DiscountCodeArry(i)) & " (" & DiscountCodeArry(i) & ")"
            
            '// Price
            discount.set "price", scCurSign &  "-" & money(DiscountAmountArry(i))
			
			daOPCDiscountTotal = daOPCDiscountTotal + CDbl(DiscountAmountArry(i))
        
            Set discounts(i) = discount
            Set discount = Nothing
        
        End If

    Next
    jsonService.set "discounts", discounts
    
End If



'// Reward Points being used
If RewardsActive=1 And savUseRewards<>"" Then 
    If iDollarValue>0 Then
        jsonService.set "rewardPointsUsedLabel", savUseRewards & " " & RewardsLabel
        jsonService.set "rewardPointsUsedTotal", scCurSign & "-" & money(iDollarValue)
    End If
End If 



'// Reward Points being accrued
If RewardsActive=1 And pcIntUseRewards=0 And pcSFCartRewards > 0 Then 
    jsonService.set "rewardPointsAccrued", pcSFCartRewards
End If 



'// Gift Wrap Total
If Session("Cust_GW")="1" Then
    If GWTotal > 0 Then
        jsonService.set "giftWrapTotal", scCurSign & money(GWTotal)
    End If
End If



'// Shipment Data
If savNullShipper="Yes" Then

    pcStrShipmentDesc = ship_dictLanguage.Item(Session("language")&"_noShip_a")
    pcShipmentPriceToAdd = "0"
    
Else '// If savNullShipper="Yes" Then

    If savNullShipRates="Yes" Then
        pcStrShipmentDesc = ship_dictLanguage.Item(Session("language")&"_noShip_b")
        pcShipmentPriceToAdd = "0"
    Else '// If savNullShipRates="Yes" Then
    
        If len(pcShippingArray)>0 Then
           
            pcSplitShipping = split(pcShippingArray,",")
            TempStrShipper=pcSplitShipping(0)
            TempStrService=pcSplitShipping(1)
            TempDblPostage=pcSplitShipping(2)
            
            If ubound(pcSplitShipping)>4 Then
                
                query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceCode='"&pcSplitShipping(5)&"';"
                set rs=server.CreateObject("ADODB.RecordSet")
                set rs=connTemp.execute(query)
                if not rs.eof then
                    pcIntIdShipService=rs("idshipservice")
                    serviceFreeOverAmt=rs("serviceFreeOverAmt")
                end if
                
                set rs=nothing

            End If

            If TempStrService = "" Then
                pcStrShipmentDesc = TempStrShipper	
                
                If pcIntIdShipService="" Then

                    query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceDescription like '%" & TempStrShipper & "%'"
                    set rs=server.CreateObject("ADODB.RecordSet")
                    set rs=connTemp.execute(query)
    
                    if not rs.eof then
                        pcIntIdShipService=rs("idshipservice")
                        serviceFreeOverAmt=rs("serviceFreeOverAmt")
                    end if
                    
                    set rs=nothing

                End IF
                		
            Else
                
                If pcIntIdShipService = "" Then               

                    query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceDescription LIKE '%" & TempStrService & "%'"
                    set rs = server.CreateObject("ADODB.RecordSet")
                    set rs = connTemp.execute(query)				
                    If Not rs.Eof Then
                        pcIntIdShipService = rs("idshipservice")
                        serviceFreeOverAmt = rs("serviceFreeOverAmt")
                    End If				
                    Set rs = nothing

                End If
        
                If TempStrShipper = "UPS" then
                    serviceVar = TempStrService
                    select case serviceVar
                    case "UPS Next Day Air "
                        TempStrService="UPS Next Day Air&reg;"
                    case "UPS 2nd Day Air "
                        TempStrService="UPS 2nd Day Air&reg;"
                    case "UPS Ground"
                        TempStrService="UPS Ground"
                    case "UPS Worldwide Express "
                        TempStrService="UPS Worldwide Express<sup>SM</sup>"
                    case "UPS Worldwide Expedited "
                        TempStrService="UPS Worldwide Expedited<sup>SM</sup>"
                    case "UPS Standard To Canada"
                        TempStrService="UPS Standard To Canada"
                    case "UPS 3 Day Select "
                        TempStrService="UPS 3 Day Select<sup>SM</sup>"
                    case "UPS Next Day Air Saver "
                        TempStrService="UPS Next Day Air Saver&reg;"
                    case "UPS Next Day Air Early A.M.&reg;"
                        TempStrService="UPS Next Day Air Early A.M.&reg;"
                    case "UPS Next Day Air Early A.M. "
                        TempStrService="UPS Next Day Air&reg; Early A.M.&reg;"
                    case "UPS 2nd Day Air A.M. "
                        TempStrService="UPS 2nd Day Air A.M.&reg;"
                    case "UPS Express Saver "
                        TempStrService="UPS Express Saver <sup>SM</sup>"
                    end select	
                End If
				
				pcv_boolShowFilteredRates = pcf_ShowFilteredRates()
				
				if (pcv_boolShowFilteredRates="1") AND (pcIntIdShipService<>"") then
					queryM="SELECT pcSM_Name FROM pcShippingMap INNER JOIN pcSMRel ON pcShippingMap.pcSM_ID=pcSMRel.pcSM_ID WHERE pcSMRel.idshipservice=" & pcIntIdShipService & ";"
					set rsM=connTemp.execute(queryM)
					if not rsM.eof then
						TempStrService=rsM("pcSM_Name")
					end if
					set rsM=nothing
				end if
        
                pcStrShipmentDesc = TempStrService
                
                pcShipmentPriceToAdd=TempDblPostage
        
                if ubound(pcSplitShipping)=3 OR ubound(pcSplitShipping)>3 then
                    pcDblServiceHandlingFee=pcSplitShipping(3)
                    TempStrNewShipping=TempStrNewShipping&","&pcSplitShipping(3)
                    if ubound(pcSplitShipping)=4 then
                        pcDblIncHandlingFee=pcSplitShipping(4)
                    else
                        pcDblIncHandlingFee=0
                    end if
                else
                    pcDblServiceHandlingFee=0
                    pcDblIncHandlingFee=0
                end if
                
            
            End If '// If TempStrService = "" Then
            
        End If '// If savNullShipRates="Yes" Then
          
    End If
    
    If pcShipmentPriceToAdd > 0 Then 
	    pcDblShipmentTotal=pcShipmentPriceToAdd     
    Else
	    pcDblShipmentTotal=0
    End If

    If pcDblShipmentTotal > 0 Then
        '// response.write scCurSign & money(pcDblShipmentTotal)
        jsonService.set "shipmentTotal", scCurSign & money(pcDblShipmentTotal)
		jsonService.set "dashipmentTotal", scCurSign & money(pcDblShipmentTotal/1.2)
    Else

        If pcv_FREESHIP="ok" Then
            '// response.write dictLanguage.Item(Session("language")&"_orderverify_37")
            shiparr=split(TempStrNewShipping,",")
            TempStrNewShipping=""

            For i=lbound(shiparr) To ubound(shiparr)

                If i=2 Then
                    TempStrNewShipping=TempStrNewShipping & "0,"
                Else
                    If i=ubound(shiparr) then
                        TempStrNewShipping = TempStrNewShipping & shiparr(i)
                    Else
                        TempStrNewShipping = TempStrNewShipping & shiparr(i) & ","
                    End If
                End If
                
            Next
            
        Else '// If pcv_FREESHIP="ok" Then
        
            'response.write dictLanguage.Item(Session("language")&"_orderverify_37")
			if (session("idCustomer")>"0") then
	            jsonService.set "shipmentTotal", dictLanguage.Item(Session("language")&"_orderverify_37")
			end if

        End If '// If pcv_FREESHIP="ok" Then

    End If '// If pcDblShipmentTotal > 0 Then

End If '// If savNullShipper="Yes" Then
jsonService.set "shippingMethod", pcStrShipmentDesc



If pcDblServiceHandlingFee<>0 Then
    jsonService.set "serviceHandlingFee", scCurSign & money(pcDblServiceHandlingFee)
End If



If ptaxVAT<>"1" And pTaxAmount>0 Then

    Redim taxes(session("taxCnt")) 
    Dim tax : Set tax = JSON.parse("{}")

    If (ptaxseparate="1" Or (ptaxCanada="1" And session("SFTaxZoneRateCnt")>0)) And session("taxCnt")<>0 Then

            
        For i=1 To session("taxCnt") 
			Set tax = JSON.parse("{}")
            If (session("taxAmount" & i) > 0) then
                tax.set "name", Session("taxDesc" & i) & ":"
                tax.set "amount", scCurSign & money(Session("taxAmount" & i))
            End If
            
            If ccur(ptaxPrdAmount)>0 Then
                tax.set "name", dictLanguage.Item(Session("language")&"_orderverify_44")
                tax.set "amount", scCurSign & money(ptaxPrdAmount)
            End If

            Set taxes(i-1) = tax

        Next '// For i=1 To session("taxCnt") 
        
        
    Else

        tax.set "name", dictLanguage.Item(Session("language")&"_orderverify_16")
        tax.set "amount", scCurSign & money(pTaxAmount)
        Set taxes(0) = tax
        
    End If
    
    jsonService.Set "taxes", taxes
    Set tax = Nothing

End If '// If ptaxVAT<>"1" And pTaxAmount>0 Then



If (savGCs<>"") Then

    GCArr=split(savGCs,",")
    
    pTempGC=""

    Redim giftCerts(ubound(GCArr))    
    For i=0 To ubound(GCArr)

        If (GCArr(i) <> "") Then

            Dim giftcert : Set giftcert = JSON.parse("{}")

            query = "SELECT pcGCOrdered.pcGO_ExpDate, pcGCOrdered.pcGO_Amount, pcGCOrdered.pcGO_Status, products.Description FROM pcGCOrdered, products WHERE pcGCOrdered.pcGO_GcCode='"&GCArr(i)&"' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"

            Set rsQ = Server.CreateObject("ADODB.Recordset")
            Set rsQ = conntemp.execute(query)
        
            If Not rsQ.eof Then

                pGCExpDate = rsQ("pcGO_ExpDate")
                pGCAmount = rsQ("pcGO_Amount")
                
                If Len(pGCAmount)<0 Then
                    pGCAmount=0
                End If
          
                pGCStatus=rsQ("pcGO_Status")
                pDiscountDesc=rsQ("Description")

                If pDiscountDesc <> "" Then

                    If pGCAmount>0 Then   
                        giftcert.set "name", dictLanguage.Item(Session("language")&"_orderverify_46") & " " & pDiscountDesc & " (" & GCArr(i) & ")"
                        giftcert.set "amount", scCurSign & "-" & money(pGCAmount)
                    End If
                    
                End If 

            End If '// If Not rsQ.eof Then
            Set rsQ = nothing

            Set giftCerts(i) = giftcert

        End If '// If (GCArr(i) <> "") And (Cdbl(pSubTotal) > 0) Then

    Next '// For i=0 To ubound(GCArr)
   jsonService.Set "giftCerts", giftCerts
    
End If '// If (savGCs<>"") Then



If ptaxVAT = "1" And VATTotal  >0 Then

    If VATRemovedTotal = 0 Then
        jsonService.Set "vatName", dictLanguage.Item(Session("language")&"_orderverify_35")  
        jsonService.Set "vatTotal", scCurSign & money(VATTotal) 
    Else
        jsonService.Set "vatName", dictLanguage.Item(Session("language")&"_orderverify_42") 
        jsonService.Set "vatTotal", scCurSign & money(VATTotal) 
    End If
    
End If

jsonService.set "daOCVAT", scCurSign &  money(0)

'// Sub Total
'DA - Edit
jsonService.set "subTotalBeforeDiscounts", scCurSign &  money(subtotal)
jsonService.set "daQuickCart", scCurSign &  money((subtotal- daBunDisc + daDelCharge))
jsonService.set "dasubTotalBeforeDiscounts", scCurSign &  money((subtotal- daBunDisc)/1.2)
jsonService.set "daVAT", scCurSign &  money((subtotal - daBunDisc + daDelCharge)-((subtotal - daBunDisc + daDelCharge)/1.2))
jsonService.set "daFinalTotal", scCurSign &  money(subtotal - daBunDisc + daDelCharge)
jsonService.set "daTotalPromotions", "- " & scCurSign &  money(daOPCDiscountTotal/1.2)
if daOPCDiscountTotal > 0 then
	jsonService.set "daOPCDiscountApplied", true
else
	jsonService.set "daOPCDiscountApplied", false
end if
jsonService.set "dasubTotalOPCSummary", scCurSign &  money((subtotal)/1.2)


subtotal = subtotal - CatDiscTotal - TotalPromotions
'// Add Discounts for View Cart
If scDispDiscCart = "1" Then
	If session("pcSFCust_DiscountCodeTotal")>"0" Then
		subtotal = subtotal - Cdbl(session("pcSFCust_DiscountCodeTotal"))
    Else
		subtotal = subtotal - Cdbl(pSFDiscountCodeTotal)
	End If
End If

'// Check for FREE shipping discount code
If Not (session("pcSFIdDbSession") = "" OR session("pcSFRandomKey") = "") Then
	query="SELECT pcFShip_IDDiscount FROM pcDFShip WHERE pcFShip_IDDiscount = (SELECT iddiscount FROM discounts WHERE discountcode = (SELECT pcCustSession_discountcode FROM pcCustomerSessions WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&"));"
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	if rs.eof then
		'// Add Shipping for View Cart
		If pcShipmentPriceToAdd > 0 Then
			subtotal = subtotal + Cdbl(pcShipmentPriceToAdd)
		End If
	end if
End if
If subtotal < 0 Then
	subtotal = 0
End If
jsonService.set "subtotal", scCurSign &  money(subtotal)


'// Show Gift 
gHaveGR=0

       
query="select pcEv_IDEvent from pcEvents where pcEv_IDCustomer=" & Session("idCustomer") & " and pcEv_Active=1"
Set rs1 = conntemp.execute(query)
If Not rs1.Eof Then
    gHaveGR=1
End If
Set rs1 = Nothing
            
If (Not (Session("idCustomer")=0)) And (gHaveGR=1) And (session("Cust_buyGift")="") Then 
    jsonService.set "haveGR", true
Else
    jsonService.set "haveGR", false
End If



'// Weight
If cdbl(totalRowWeight) > 0 And cdbl(totalRowWeight) < 1 Then
    totalRowWeight = 1
End If
totalRowWeight = round(totalRowWeight, 0)

If scShowCartWeight="-1" Then
    jsonService.set "showCartWeight", true
Else
    jsonService.set "showCartWeight", false
End If

If scShipFromWeightUnit="KGS" Then

    pKilos = Int(totalRowWeight/1000)
    pWeight_g = totalRowWeight-(pKilos*1000)

    jsonService.set "kilos", pKilos & " kg "
    
    If pWeight_g > 0 Then 
        jsonService.set "weightG", pWeight_g & " g"
    End If 

Else

    pPounds = Int(totalRowWeight/16)
    pWeight_oz = totalRowWeight-(pPounds*16)
          
    jsonService.set "pounds", pPounds & " lbs. "

    If pWeight_oz > 0 Then
        jsonService.set "weightOZ", pWeight_oz & " oz."
    End If
    
End If
jsonService.set "weight", totalRowWeight


'// Gift Wrap
jsonService.set "grCode", grCode



'// Checkout Stage (Please Holder)
jsonService.set "checkoutStage", true

'//VAT EU Calculations
boolOPCVATRemoved = 0

Dim pcv_IsEUMemberState
Dim pcv_ShipIsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pcStrBillingCountryCode)
pcv_ShipIsEUMemberState = pcf_IsEUMemberState(pcStrShippingCountryCode)
if pcv_IsEUMemberState = 0 then
'Non EU Country
	'Now check del is not UK, if it is charge VAT, if not zero VAT
	if pcv_ShipIsEUMemberState = 0 then
		VATTotal = 0
		boolOPCVATRemoved = 1
	end if
end if

if pcv_IsEUMemberState = 1 then
'Billing is EU
	if pcStrBillingVATID <> "" then
	'EU country with a VAT number
		VATTotal = 0
		boolOPCVATRemoved = 1
	end if
end if


'// Total
If IsCartSaved Then
    jsonService.set "total", scCurSign &  money(cartTotal)
	if boolOPCVATRemoved = 1 then
		jsonService.set "datotalFinalCheckout", scCurSign &  money(cartTotal)
	else
		jsonService.set "datotalFinalCheckout", scCurSign &  money(cartTotal/1.2)
	end if
	jsonService.set "daOPCVAT", scCurSign &  money(VATTotal)
Else
    jsonService.set "total", scCurSign &  money(subtotal)
	if boolOPCVATRemoved = 1 then
		jsonService.set "datotalFinalCheckout", scCurSign &  money(subtotal)
	else
		jsonService.set "datotalFinalCheckout", scCurSign &  money(subtotal/1.2)
	end if
	jsonService.set "daOPCVAT", scCurSign &  money(VATTotal)
End If




'// Total Quantity
jsonService.set "totalQuantity", totalQuantity

If totalQuantity>=100 Then
    jsonService.set "totalQuantityDisplay", "99+"
Else
    jsonService.set "totalQuantityDisplay", totalQuantity
End If

jsonService.set "shoppingcartrow", shoppingcartrows












'/////////////////////////////////////////////////////
'// OPC Display Settings
'/////////////////////////////////////////////////////

'// Check Gift Registry
If session("Cust_IDEvent")<>"" Then

    query="select pcEv_IDCustomer, pcEv_Delivery, pcEv_MyAddr, pcEv_HideAddress from PcEvents where pcEv_IDEvent=" & session("Cust_IDEvent")
    set rstemp=connTemp.execute(query)					
    If Not rstemp.eof Then
    
        gIDCustomer = rstemp("pcEv_IDCustomer")
        gDelivery = rstemp("pcEv_Delivery")
        If Not (gDelivery<>"") then
            gDelivery  =0
        End If
        
        gMyAddr = rstemp("pcEv_MyAddr")
        If Not (gMyAddr<>"") Then
            gMyAddr = 0
        End If
        If gDelivery="1" Then
            GRTest = 1
        End If
        gHideAddress = rstemp("pcEv_HideAddress")
        session("gHideAddress") = gHideAddress						

    End If
    Set rstemp = Nothing

End If


'// Set global shipping mode to local
pcv_AlwAltShipAddress = scAlwAltShipAddress

pcv_NOShippingAtAll = "1" 


query="SELECT active FROM ShipmentTypes WHERE active<>0"
set rs = connTemp.execute(query)
If rs.eof Then '// There are NO active dynamic shipping services
    pcv_NoDynamicShipping="1"
End If

query="SELECT idFlatShipType FROM FlatShipTypes"
set rs = connTemp.execute(query)
If rs.eof Then '// There are NO active custom shipping services
    pcv_NoCustomShipping="1"
end if
If pcv_NoDynamicShipping="1" And pcv_NoCustomShipping="1" Then '// There are NO active shipping options
    pcv_AlwAltShipAddress = "1"
    pcv_NOShippingAtAll = "2"
End If


'// If No shipping at all is still set to "1" - check if products in cart qualify for no shipping and if so - is the store owner hiding the address?
pShipTotal = Cdbl(calculateShipCartTotal(pcCartArray, pcCartIndex))
pShipWeight = Cdbl(calculateShipWeight(pcCartArray, pcCartIndex))
pShipQuantity = Int(calculateCartShipQuantity(pcCartArray, pcCartIndex))

If session("Cust_IDEvent")="" And pShipTotal=0 And pShipWeight=0 And pShipQuantity=0 And scHideShipAddress="1" And pcv_NOShippingAtAll="1" Then
    pcv_NOShippingAtAll = "2"
End If

If session("Cust_IDEvent")<>"" Then
    pcv_AlwAltShipAddress = "2"
End If

If (Session("idCustomer")=0 Or Session("idCustomer")="") Or (session("idCustomer")>"0" And session("CustomerGuest")>"0") Then    
    jsonService.set "displayLogin", true          
Else    
    jsonService.set "displayLogin", false         
End If

If session("CustomerGuest")="0" Then           
    jsonService.set "guestCustomer", false 
Else
    jsonService.set "guestCustomer", true
End If

'// Guest Checkout Status
If scGuestCheckoutOpt=0 Or scGuestCheckoutOpt="" Then
    jsonService.set "guestCheckoutStatus", 0
ElseIf scGuestCheckoutOpt="1" Then
    jsonService.set "guestCheckoutStatus", 1
ElseIf scGuestCheckoutOpt="2" Then
    jsonService.set "guestCheckoutStatus", 2
End If 

'// Guest Checkout Fields
If Not (Session("idCustomer")>0 And session("CustomerGuest")="0") Then
    jsonService.set "displayGuestFields", true  
Else
    jsonService.set "displayGuestFields", false  
End If

'// Allow Passwords
If scGuestCheckoutOpt=0 Or scGuestCheckoutOpt="" Or scGuestCheckoutOpt=1 Then 
    jsonService.set "allowPassword", true
Else 
   jsonService.set "allowPassword", true 
End If 

'// Guest Checkout
jsonService.set "GuestCheckout", session("CustomerGuest") 

'// New Customer (previously tmpNewCust)
If Session("CustomerGuest")="0" And Session("IdCustomer")>"0" Then 
    jsonService.set "newCustomer", 0
Else 
    jsonService.set "newCustomer", 1
End If 

'// Show Optional Password Box
If Not (session("idCustomer")>"0" AND (session("CustomerGuest")="0" OR session("CustomerGuest")="2")) Then
    jsonService.set "DisplayOptionalPassword", true
Else
    jsonService.set "DisplayOptionalPassword", false
End IF

'// Allow Nicknames
If scOrderName = "1" Then
    jsonService.set "AllowNicknames", true
Else
    jsonService.set "AllowNicknames", false
End If


'// Gift Wrapping                            
query = "SELECT pcGWSet_Show, pcGWSet_Overview, pcGWSet_HTML FROM pcGWSettings;"
Set rs = server.CreateObject("ADODB.RecordSet")
Set rs = connTemp.execute(query)
pcvGW=0
pcvOverview=0
If Not rs.Eof Then
    pcvGW = rs("pcGWSet_Show")
    if IsNull(pcvGW) OR pcvGW="" then
        pcvGW="0"
    end if			
    session("Cust_GW")=pcvGW					
    pcvOverview=rs("pcGWSet_Overview")
    if pcvOverview="0" then
        pcvOverview=""
    end if
    session("Cust_GWText")=pcvOverview
    pcvGWDetails=rs("pcGWSet_HTML")
    jsonService.set "giftWrappingDetails", pcf_FixHTMLContentPaths(pcvGWDetails)
    If pcvGW="0" Then
        jsonService.set "ShowGiftWrapOptions", false
    Else
        jsonService.set "ShowGiftWrapOptions", true
    End If
Else
    jsonService.set "ShowGiftWrapOptions", false
End If
Set rs = Nothing
   
If tmpStrList = 1 Then
    jsonService.set "showGiftWrapProductList", true
Else
    jsonService.set "showGiftWrapProductList", false
End If

'// Customer Logged In
If session("idcustomer")>0 Then
    jsonService.set "IsLoggedIn", true
Else
    jsonService.set "IsLoggedIn", false
End If

'// Time Field
If TFShow="1" Then
    jsonService.set "displayTimeField", true
Else
    jsonService.set "displayTimeField", false   
End If

'// Time Field
If TFShow="1" Then
    jsonService.set "timeFieldLabel", TFLabel
Else
    jsonService.set "timeFieldLabel", ""   
End If

'// Time Field Required
If TFReq="1" Then
    jsonService.set "timeFieldRequired", true
Else
    jsonService.set "timeFieldRequired", false
End IF

'// Time Field Value
If savTF1<>"" Then
    jsonService.set "timeFieldValue", savTF1
Else
    jsonService.set "timeFieldValue", ""
End If

'// Date Field
If DFShow="1" Then
    jsonService.set "displayDateField", true
Else
    jsonService.set "displayDateField", false   
End If

'// Date Label
If DFShow="1" Then
    jsonService.set "dateFieldLabel", DFLabel
Else
    jsonService.set "dateFieldLabel", ""   
End If

'// Date Field Value
If savDF1<>"" Then
    jsonService.set "dateFieldValue", savDF1
Else
    jsonService.set "dateFieldValue", ""
End If

'// Date Field Required
If DFReq="1" Then
    jsonService.set "dateFieldRequired", true
Else                                                             
    jsonService.set "dateFieldRequired", false
End If

'// Delivery Area Required
If DFShow="1" Or TFShow="1" Then
    jsonService.set "HaveDeliveryArea", 1
Else
    jsonService.set "HaveDeliveryArea", 0   
End If

'// Delivery Blackout Dates                      

'// If the store is using blackout dates, show a message here and a link a list of dates
Dim blackoutdates
query="SELECT * FROM Blackout ORDER BY Blackout_Date ASC;"
Set rs = connTemp.execute(query)
If rs.eof Then
    blackoutdates="0"
Else
    blackoutdates="1"
End If
Set rs = nothing
        
If blackoutdates="1" Then 
    jsonService.set "displayBlackOutDates", true
Else
    jsonService.set "displayBlackOutDates", false
End If 
     
'// Delivery Date Message
If (DTCheck="1") Then 
    jsonService.set "displayDeliveryDateMessage", true
Else
    jsonService.set "displayDeliveryDateMessage", true
End If 

'// Express Checkout
Dim hasExpressCheckout
If pcIsSubscription then		
	strAndSub = "AND (pcPayTypes_Subscription = 1)"
Else		
	strAndSub = ""		
End if
query="SELECT idPayment FROM paytypes WHERE active=-1 AND (gwCode=999999 OR gwCode=46 OR gwCode=53 OR gwCode=80 OR gwCode=99 OR gwCode=88)" & strAndSub
Set rs = connTemp.execute(query)
hasExpressCheckout = false
If Not rs.eof Then
  hasExpressCheckout = true
End If
Set rs = nothing

jsonService.set "hasExpressCheckout", hasExpressCheckout

If session("ExpressCheckoutPayment") <> "YES" Then
    jsonService.set "ExpressCheckoutInUse", false
Else
    jsonService.set "ExpressCheckoutInUse", true
End If

If session("PayWithAmazon") <> "YES" Then
    jsonService.set "PayWithAmazonInUse", false
Else
    jsonService.set "PayWithAmazonInUse", true
End If

If session("AmazonFirstTime") <> "1" Then
    jsonService.set "AmazonFirstTime", false
Else
    jsonService.set "AmazonFirstTime", true
End If

'// Gift Registry in Cart
If Session("Cust_BuyGift")<>"" Then
     jsonService.set "IsBuyGift", true       
Else
     jsonService.set "IsBuyGift", false 
End If

If (pcv_NOShippingAtAll = "2" And pcv_AlwAltShipAddress="0") Or (pcv_AlwAltShipAddress="1") Then
    jsonService.set "displayShippingAddress", false
Else
    jsonService.set "displayShippingAddress", true
    
    If pcv_AlwAltShipAddress="0" Or pcv_AlwAltShipAddress="2" Then
        
        If (session("Cust_IDEvent")="") Or (session("Cust_IDEvent")<>"" And gDelivery=0) Then
            
                jsonService.set "NeedLoadShipContent", 1 '// var NeedLoadShipContent=1;
                jsonService.set "CanCreateNewShip", 1 '// var CanCreateNewShip=1;
                
        Else
            
                jsonService.set "NeedLoadShipContent", 0 '// var NeedLoadShipContent=0;
                jsonService.set "CanCreateNewShip", 0 '// var CanCreateNewShip=0;
                
        End If
            
    Else
        
            jsonService.set "NeedLoadShipContent", 0 '// var NeedLoadShipContent=0;
            jsonService.set "CanCreateNewShip", 0 '// var CanCreateNewShip=0;
            
    End If

End If

If Not pcv_NOShippingAtAll = "2" Then
    jsonService.set "displayRatesArea", true
Else
    jsonService.set "displayRatesArea", false
End IF


'// Edit Order Mode (currently only applies to SubscriptionBridge)
If (Session("SBEditOrder")<>"") And (Session("SBEditOrderID")<>"") Then
    jsonService.set "IsEditOrder", true
Else
    jsonService.set "IsEditOrder", false
End If

'// Gift In Cart
If HaveGcsTest = 1 Then
    jsonService.set "HaveGcs", true
Else
    jsonService.set "HaveGcs", false
End If

'// Currency Format
If scDecSign = "," Then
    jsonService.set "decimal", ","
Else
    jsonService.set "decimal", "."
End If

'// Express Checkout Buttons
pcv_strShowCheckoutBtn = pcf_PaymentTypes("")
If pcv_strShowCheckoutBtn=1 Then
    jsonService.set "ShowCheckoutBtn", true       
Else
    jsonService.set "ShowCheckoutBtn", false    
End If

'// Currency
jsonService.set "currencySymbol", scCurSign  


'// Hide Shipping Panel
If pcv_NOShippingAtAll = "2" OR scAlwAltShipAddress="1" Then
    jsonService.set "IsHideShippingPanel", true
Else
    jsonService.set "IsHideShippingPanel", false
End If

'// Hide Delivery Panel
If pcv_NOShippingAtAll = "2" Then
    jsonService.set "IsHideDeliveryPanel", true
Else
    jsonService.set "IsHideDeliveryPanel", false
End If

'// ADDRESS TYPE SELECTION AREA
jsonService.set "billingAddressTypeArea", showBillingAddressTypeArea(pcv_NOShippingAtAll, pcv_AlwAltShipAddress, scComResShipAddress) 
jsonService.set "shippingAddressTypeArea", showShippingAddressTypeArea(pcv_NOShippingAtAll, pcv_AlwAltShipAddress, scComResShipAddress)

response.Clear()
If rowCount > 0 Then
    Response.write( JSON.stringify(jsonService, null, 2) & vbNewline )
Else
    Response.Write("")
End If
set Info = nothing

call closeDb()
response.End()
%>