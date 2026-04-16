<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/CashbackConstants.asp"-->
<% pcStrPageName="service_orders.asp" %>
<% 
response.Clear()
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

dim jsonService : set jsonService = JSON.parse("{}")

dim pIdOrder, pOID, pnValid, pOrderStatus, pcv_noDoubleTracking
%>
<!--#include file="prv_getsettings.asp"-->
<% 
If len(session("idOrder"))>0 Then
	pOID=session("idOrder")
	session("idOrderConfirm")=pOID
Else
	pOID=session("idOrderConfirm")
	pcv_noDoubleTracking=1
End If
if pOID = "" then
	pOID = 0
	pnValid=1
end if
session("idOrder")=""
session("GWOrderId")="" '// PayPal Standard
if NOT validNum(pOID) then
	pnValid=1
end if

' Start Order Details section
pIdOrder=pOID

query="SELECT customers.idCustomer, customers.email,customers.fax,customers.pcCust_VATID,orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr,orders.orderDate, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.customerType, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.comments, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_shippingPhone, orders.shippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.idOrder, orders.rmaCredit, orders.ordPackageNum, orders.ord_DeliveryDate, orders.ord_OrderName, orders.ord_VAT,orders.pcOrd_CatDiscounts, orders.paymentDetails, orders.gwAuthCode, orders.gwTransId, orders.paymentCode, orders.pcOrd_GWTotal FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number<>0 then
    call LogErrorToDatabase()
    set rs=nothing
    call closedb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
    set rs=nothing
    call closeDb()
    response.redirect "msg.asp?message=35"     
end if 

dim pidCustomer, porderDate, pfirstname, plastname,pcustomerCompany, pphone, paddress, pzip, pstate, pcity, pcountryCode, pcomments, pshippingAddress, pshippingState, pshippingCity, pshippingCountryCode, pshippingZip, paddress2, pshippingFullName, pshippingCompany, pshippingAddress2, pshippingPhone, pcustomerType


pEmail=rs("email")
pFax=rs("fax")
pshippingEmail=rs("pcOrd_ShippingEmail")
pshippingFax=rs("pcOrd_ShippingFax")
pcShowShipAddr=rs("pcOrd_ShowShipAddr")
if IsNull(pcShowShipAddr) OR (pcShowShipAddr="") then
    pcShowShipAddr=0
end if
pidCustomer=rs("idCustomer")
pVATID=rs("pcCust_VATID")
porderDate=rs("orderDate")
porderDate=showdateFrmt(porderDate)
pfirstname=rs("name")
plastName=rs("lastName")
pCustomerName=pfirstname& " " & plastName
pcustomerCompany=rs("customerCompany")
pphone=rs("phone")
pcustomerType=rs("customerType")
paddress=rs("address")
pzip=rs("zip")
pstate=rs("stateCode")
if pstate="" then
    pstate=rs("state")
end if
pcity=rs("city")
pcountryCode=rs("countryCode")
pcomments=rs("comments")
pshippingAddress=rs("shippingAddress")
pshippingState=rs("shippingStateCode")
if pshippingState="" then
    pshippingState=rs("shippingState")
end if
pshippingCity=rs("shippingCity")
pshippingCountryCode=rs("shippingCountryCode")
pshippingZip=rs("shippingZip")
pshippingPhone=rs("pcOrd_shippingPhone")
pshippingFullName=rs("shippingFullName")
paddress2=rs("address2")
pshippingCompany=rs("shippingCompany")
pshippingAddress2=rs("shippingAddress2")
pidOrder=rs("idOrder")
pRmaCredit=rs("rmaCredit")
pOrdPackageNum=rs("ordPackageNum")
pord_DeliveryDate=rs("ord_DeliveryDate")
pord_OrderName=rs("ord_OrderName")
pord_VAT=rs("ord_VAT")
pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
    pcv_CatDiscounts="0"
end if
pcpaymentDetails=trim(rs("paymentDetails"))
pcgwAuthCode=rs("gwAuthCode")
pcgwTransId=rs("gwTransId")
pcpaymentCode=rs("paymentCode")
'GGG Add-on start
pGWTotal=rs("pcOrd_GWTotal")
if pGWTotal<>"" then
else
pGWTotal="0"
end if
'GGG Add-on end

'// Check if the Customer is European Union 
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)
boolVAT = CheckVAT(pcv_IsEUMemberState, pshippingCountryCode, pVATID)

query="SELECT ProductsOrdered.idProduct, ProductsOrdered.pcSubscription_ID, ProductsOrdered.quantity, ProductsOrdered.unitPrice, ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts  "
'CONFIGURATOR ADDON-S
If scBTO=1 then
    query=query&", ProductsOrdered.idconfigSession"
End If
'CONFIGURATOR ADDON-E
query=query&", pcPO_GWOpt, pcPO_GWPrice, products.description, products.sku, orders.total, orders.paymentDetails, orders.taxamount, orders.shipmentDetails, orders.discountDetails, orders.pcOrd_GCDetails, orders.orderstatus,orders.processDate, orders.shipdate, orders.shipvia, orders.trackingNum, orders.returnDate, orders.returnReason, orders.iRewardPoints, orders.iRewardValue, orders.iRewardPointsCustAccrued, orders.taxdetails, orders.dps, ProductsOrdered.xfdetails, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, pcPrdOrd_BundledDisc, pcPO_GWNote FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idCustomer=" & pidCustomer & " AND orders.idOrder=" & pIdOrder

set rsOrdObj = server.CreateObject("ADODB.RecordSet")
rsOrdObj.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
                    
dim pidProduct, pquantity, punitPrice, pxfdetails, pidconfigSession, pdescription, pSku, pcDPs, ptotal, ppaymentDetails,ptaxamount,pshipmentDetails, pdiscountDetails
dim pprocessDate, pshipdate, pshipvia, ptrackingNum, preturnDate, preturnReason, piRewardPoints, piRewardValue, piRewardPointsCustAccrued,ptaxdetails, pOpPrices, rsObjOptions, pRowPrice, count, rsConfigObj,stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory,i, s,OptPrice,xfdetails, xfarray, q
Dim GCDetails
Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
Dim subTotal
subTotal=0

'////////////////////////////////////////////////////////////////////
'// START: CART LOOP
'////////////////////////////////////////////////////////////////////
rowCount = rsOrdObj.RecordCount 
f = 1
redim shoppingcartrows(rowCount - 1)
do while not rsOrdObj.eof
    pidProduct=rsOrdObj("idProduct")
    pcSubscription_ID=rsOrdObj("pcSubscription_ID")
    pquantity=rsOrdObj("quantity")
    punitPrice=rsOrdObj("unitPrice")
    QDiscounts=rsOrdObj("QDiscounts")
    ItemsDiscounts=rsOrdObj("ItemsDiscounts")
    'CONFIGURATOR ADDON-S
    if scBTO=1 then
        pidconfigSession=rsOrdObj("idconfigSession")
        if pidconfigSession="" then
            pidconfigSession="0"
        end if
    End If
    'CONFIGURATOR ADDON-E
    'GGG Add-on start  
    pGWOpt=rsOrdObj("pcPO_GWOpt")
    if pGWOpt<>"" then
    else
        pGWOpt="0"
    end if
    pGWPrice=rsOrdObj("pcPO_GWPrice")
    if pGWPrice<>"" then
    else
        pGWPrice="0"
    end if
    'GGG Add-on end
    
    pdescription=rsOrdObj("description")
    pSku=rsOrdObj("sku")
    ptotal=rsOrdObj("total")
    ppaymentDetails=trim(rsOrdObj("paymentDetails"))
    ptaxamount=rsOrdObj("taxamount")
    pshipmentDetails=rsOrdObj("shipmentDetails")
    pdiscountDetails=rsOrdObj("discountDetails")
    GCDetails=rsOrdObj("pcOrd_GCDetails")
    porderstatus=rsOrdObj("orderstatus")
    pprocessDate=rsOrdObj("processDate")
    pshipdate=rsOrdObj("shipdate")
    pshipdate=ShowDateFrmt(pshipdate)
    pshipvia=rsOrdObj("shipvia")
    ptrackingNum=rsOrdObj("trackingNum")
    preturnDate=rsOrdObj("returnDate")
    preturnDate=ShowDateFrmt(preturnDate)
    preturnReason=rsOrdObj("returnReason")
    piRewardPoints=rsOrdObj("iRewardPoints")
    piRewardValue=rsOrdObj("iRewardValue")
    piRewardPointsCustAccrued=rsOrdObj("iRewardPointsCustAccrued")
    ptaxdetails=rsOrdObj("taxdetails")
    pcDPs=rsOrdObj("DPs")
    pxfdetails=rsOrdObj("xfdetails")
    '// Product Options Arrays
    pcv_strSelectedOptions = rsOrdObj("pcPrdOrd_SelectedOptions") ' Column 11
    pcv_strOptionsPriceArray = rsOrdObj("pcPrdOrd_OptionsPriceArray") ' Column 25
    pcv_strOptionsArray = rsOrdObj("pcPrdOrd_OptionsArray") ' Column 4
    pcPrdOrd_BundledDisc=rsOrdObj("pcPrdOrd_BundledDisc")
    pGWText=rsOrdObj("pcPO_GWNote")
    pprocessDate=ShowDateFrmt(pprocessDate)
    
    pIdConfigSession=trim(pidconfigSession)
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' START: Get the total Price of all options
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    pOpPrices=0
    dim pcv_tmpOptionLoopCounter, pcArray_TmpCounter
    
    If len(pcv_strOptionsPriceArray)>0 then
    
        pcArray_TmpCounter = split(pcv_strOptionsPriceArray,chr(124))
        For pcv_tmpOptionLoopCounter = 0 to ubound(pcArray_TmpCounter)
            pOpPrices = pOpPrices + pcArray_TmpCounter(pcv_tmpOptionLoopCounter)
        Next
        
    end if				

    if NOT isNumeric(pOpPrices) then
        pOpPrices=0
    end if	
    
    '// Apply Discounts to Options Total
    '   >>> call function "pcf_DiscountedOptions(OriginalOptionsTotal, ProductID, Quantity, CustomerType)" from stringfunctions.asp
    Dim pcv_intDiscountPerUnit
    pOpPrices = pcf_DiscountedOptions(pOpPrices, pidProduct, pquantity, pcustomerType)
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' END: Get the total Price of all options
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
    %>
    
    <% 'CONFIGURATOR ADDON-S
    err.number=0
    TotalUnit=0
    If scBTO=1 then
        pIdConfigSession=trim(pidconfigSession)
        if pIdConfigSession<>"0" then 
            query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
            set rsConfigObj=conntemp.execute(query)
            if err.number<>0 then
                call LogErrorToDatabase()
                set rsConfigObj=nothing
                call closedb()
                response.redirect "techErr.asp?err="&pcStrCustRefID
            end if
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
            set rsConfigObj=nothing
            for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
            query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
            set rsConfigObj=conntemp.execute(query)
            if err.number<>0 then
                call LogErrorToDatabase()
                set rsConfigObj=nothing
                call closedb()
                response.redirect "techErr.asp?err="&pcStrCustRefID
            end if
            if NOT validNum(ArrQuantity(i)) then
                pIntQty=1
            else
                pIntQty=ArrQuantity(i)
            end if
            
            query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
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
            query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
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
                TotalUnit=TotalUnit+((ArrValue(i)+UPrice)*pQuantity)
            end if
            set rsConfigObj=nothing
            next
        end if 
    End If 
    'CONFIGURATOR ADDON-E

    if TotalUnit>0 then
        punitPrice1=punitPrice
        if pIdConfigSession<>"0" then
            pRowPrice1=Cdbl(pquantity * ( punitPrice1 )) - TotalUnit
            punitPrice1=Round(pRowPrice1/pquantity,2)
        else
            pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
        end if
    else
        punitPrice1=punitPrice
        if pIdConfigSession<>"0" then
            pRowPrice1=Cdbl(pquantity * ( punitPrice1 ))
        else
            pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
            punitPrice1=Round(pRowPrice1/pquantity,2)
        end if
    end if

    Dim shoppingcartrow : set shoppingcartrow = JSON.parse("{}")

    shoppingcartrow.set "row", f
    shoppingcartrow.set "id", pidProduct
    shoppingcartrow.set "quantity", pQuantity
    shoppingcartrow.set "sku", pSku
    shoppingcartrow.set "description", replace(pdescription, "&quot;", """")
    
    'If Not (pcCartArray(f,27)>0) Then
    '    rowCount = rowCount + 1
    'End If 
    'If rowCount mod 2 = 1 Then
        shoppingcartrow.set "rowClass", "odd"
    'Else
    '    shoppingcartrow.set "rowClass", "even"
    'End If

    If pcv_RWActive="1" Then
            
        query="SELECT pcRE_IDProduct FROM pcRevExc WHERE pcRE_IDProduct=" & pidProduct
        set rsQ=server.CreateObject("ADODB.RecordSet")
        set rsQ=connTemp.execute(query)
        if rsQ.eof then
            Prv_Accept=1
        else
            Prv_Accept=0
        end if
        set rsQ=nothing
        
        IF Prv_Accept=1 THEN
            shoppingcartrow.set "writeReview", true
        End If
        
    End If

    'SB S
    If pcSubscription_ID>0 Then
        query="SELECT SB_Terms FROM SB_Orders WHERE idOrder=" & pIdOrder & ";"
        Set rsSB=Server.CreateObject("ADODB.Recordset")
        Set rsSB=connTemp.execute(query)
        If NOT rsSB.eof Then
            pcv_strTerms = rsSB("SB_Terms")
            if len(pcv_strTerms)>0 then
                'response.Write(pcv_strTerms)
            end if
        End If
    End If
    'SB E

    shoppingcartrow.set "UnitPrice", scCurSign & money(punitPrice1)
	shoppingcartrow.set "DAUnitPrice", scCurSign & money(punitPrice1/1.2)
    shoppingcartrow.set "RowPrice", scCurSign & money(pRowPrice1)
	shoppingcartrow.set "DARowPrice", scCurSign & money(pRowPrice1/1.2)


    '/////////////////////////////////////////////////////////////////////////////////////
    '// START: BTO PRODUCT DETAILS
    '/////////////////////////////////////////////////////////////////////////////////////      
    If scBTO = 1 then

        pIdConfigSession=trim(pidconfigSession)
        
        If pIdConfigSession<>"0" Then 
            
            query="SELECT stringProducts, stringValues, stringCategOries, stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
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
                    query="SELECT categories.categoryDesc, products.description, products.pcProd_ParentPrd FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
                    'APP-E
                    set rs=server.CreateObject("ADoDB.RecOrdSet")
                    set rs=conntemp.execute(query)
                      
                    strCategoryDesc=rs("categoryDesc")
                    strDescription=rs("description")
                    'APP-S
                    intParentPrd=rs("pcProd_ParentPrd")
                    if intParentPrd>"0" then
                    else
                        intParentPrd=ArrProduct(i)
                    end if
                    'APP-E
                    set rs=Nothing
                      
                    'APP-S
                    query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&intParentPrd&" and specProduct=" & pidProduct 
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
                    query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & intParentPrd & " AND cdefault<>0;"
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
                    shoppingcartrow.set "BToConfigQuantity", ArrQuantity(i)
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
                  
                      myBToConfigPriceString = "" & ""&scCurSign & money(ccur((ArrValue(i)+UPrice)*pQuantity))&""
                    Else
                      If tmpDefault=1 Then 
                        myBToConfigPriceString = "" & ""&dictLanguage.Item(Session("language")&"_defaultNotice_1")&""
                      End If
                    End If
                    
                    btoLineItem.set "BToConfigPrice", myBToConfigPriceString
                    
                    Set btoConfiguration(i) = btoLineItem
                    Set btoLineItem = Nothing
                  
                Next
                shoppingcartrow.Set "btoConfiguration", btoConfiguration
 
            End If 

        End If 
        
    End If
    '/////////////////////////////////////////////////////////////////////////////////////
    '// END: BTO PRODUCT DETAILS
    '/////////////////////////////////////////////////////////////////////////////////////

    

    '// START 4th Row - Product Options
    if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
        pcv_strSelectedOptions = ""
    end if

    if len(pcv_strSelectedOptions)>0 then 

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
    
        ' Start in Position One
        pcv_intOptionLoopCounter = 0
    
        ' Display Our Options
        Redim productoptions(pcv_inToptionLoopSize)          
        For pcv_inToptionLoopCounter = 0 To pcv_inToptionLoopSize 
    
            Dim productoption : Set productoption = JSON.parse("{}")
        
            '// Display Our Options Prices - If any
            tempPrice = pcArray_strOptionsPrice(pcv_inToptionLoopCounter)
            If Not (tempPrice="" Or tempPrice=0) Then 
                tAprice=(tempPrice*Cdbl(pquantity))
				tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice)
                productoption.set "name", pcArray_strOptions(pcv_inToptionLoopCounter)
                productoption.set "unitprice", scCurSign & money(tempPrice)
                productoption.set "price", scCurSign & money(tAprice)
                productoption.set "DAunitprice", scCurSign & money(tempPrice/1.2)
                productoption.set "DAprice", scCurSign & money(tAprice/1.2)
            Else
                tAprice = 0
                productoption.set "name", pcArray_strOptions(pcv_inToptionLoopCounter)

            End If
    
            Set productoptions(pcv_inToptionLoopCounter) = productoption
            Set productoption = Nothing
        Next
        shoppingcartrow.Set "productoptions", productoptions
    

        shoppingcartrow.set "HasProductOptions", false

    
    End If       
    '// END 4th Row - Product Options

    customFields=""
    xfdetails=pxfdetails
    If len(xfdetails)>3 then
        xfarray=split(xfdetails,"|")
        Redim customFields(ubound(xfarray))   
        for q=lbound(xfarray) to ubound(xfarray)            
            Dim customField : Set customField = JSON.parse("{}")            
            customField.set "xField", xfarray(q)
            Set customFields(q) = customField
            Set customField = Nothing  
            shoppingcartrow.set "editInputField", true          
        next
    End If 
    shoppingcartrow.Set "customFields", customFields

    err.number=0
    pRowPrice=(punitPrice)*(pquantity)
    pExtRowPrice=pRowPrice
    Charges=0

    '// START 6th Row - BTO Item Discounts
    If scBTO=1 then
        pIdConfigSession = trim(pidconfigSession)        
        If pIdConfigSession<>"0" Then
            ItemsDiscounts = trim(ItemsDiscounts)
            If ItemsDiscounts="" Then
                ItemsDiscounts=0
            End If
            If (ItemsDiscounts<>"") And (Cdbl(ItemsDiscounts)<>"0") Then
                shoppingcartrow.set "itemDiscountRowTotal", scCurSign &  "-" & money(ItemsDiscounts)
                pRowPrice = pRowPrice - Cdbl(ItemsDiscounts)
            End If                
        End If  
    End IF   
    '// END 6th Row - BTO Item Discounts


    '// START 7th Row - BTO Additional Charges
    If scBTO=1 then
        pIdConfigSession = trim(pidconfigSession)
        If pIdConfigSession<>"0" Then
                       
            query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
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
                        If (CDbl(ArrCValue(i))>0) Then
                            Charges=Charges+cdbl(ArrCValue(i))
                        End If 
                        additionalCharge.set "categoryDesc", strCategoryDesc
                        additionalCharge.set "description", strDescription
                    End If
                    Set rs = Nothing
    
                    additionalCharge.set "total", scCurSign & money(ArrCValue(i))
                    
                    Set additionalCharges(i) = additionalCharge
                    Set additionalCharge = Nothing
                    
                Next
                pRowPrice=pRowPrice+Cdbl(Charges)
                shoppingcartrow.set "additionalCharges", additionalCharges
                
            End If '// If ArrCProduct(0)<>"na" Then
            
        End If  '// If trim(pcCartArray(f,16)) <> "" Then
    
    End If
    '// END 7th Row - BTO Additional Charges



    '// START 8th Row - Quantity Discounts
    QDiscounts = trim(QDiscounts)
    If QDiscounts="" Then
        QDiscounts=0
    End If

    If (QDiscounts<>"") And (CDbl(QDiscounts)<>"0") Then
        shoppingcartrow.set "itemQuantityDiscountRowTotal", scCurSign & "-" & money(QDiscounts)
        pRowPrice = pRowPrice - Cdbl(QDiscounts)
    End If    
    
    pcv_tmpID = pidProduct
    'APP-S
    'If pcCartArray(f,32)<>"" Then
    '    pcv_tmpPPrd=split(pcCartArray(f,32),"$$")
    '    pcv_tmpID=pcv_tmpPPrd(ubound(pcv_tmpPPrd))
    
    '    query="SELECT discountPerUnit FROM discountsPerQuantity WHERE idProduct=" & pcv_tmpID & ";"
     '   set rsQ=connTemp.execute(query)
    '    if rsQ.eof then
    '        pcv_tmpID = pidProduct
    '    end if
    '    set rsQ=nothing
    'end if  
    'APP-E        
    shoppingcartrow.set "itemQuantityDiscountRowID", pcv_tmpID
    '// END 8th Row - Quantity Discounts


   
    '// START 9th Row - Product Subtotal
    If pRowPrice1 <> pRowPrice Then
        shoppingcartrow.set "productSubTotal", scCurSign &  money(pRowPrice)
		shoppingcartrow.set "DAproductSubTotal", scCurSign &  money(pRowPrice/1.2)
    Else
        shoppingcartrow.set "productSubTotal", ""
    End If    
    if pRowPrice<>pRowPrice1 then
        subTotal=subTotal + cdbl(pRowPrice)
    else
        subTotal=subTotal + cdbl(pRowPrice1)
    end if    
    '// END 9th Row - Product Subtotal  



    '// START 10th Row - Cross Sell Bundle Discount
    'shoppingcartrow.set "ChildBundleID", pcCartArray(f, 8)        
     
    
    'If (pcCartArray(f, 27) = -1) And (Not pcCartArray(f, 8)="") Then
        
    '    If (pcCartArray(f, 8)>0) Then
    '        shoppingcartrow.set "BundleTitle", pcCartArray(f, 1) & " + " & pcCartArray(f + 1, 1)
    '        shoppingcartrow.set "IsParentofBundle", true
    '    Else
    '        shoppingcartrow.set "IsParentofBundle", false
    '    End If
             
    'Else
    '    shoppingcartrow.set "IsParentofBundle", false
    'End If
    
    If (pcPrdOrd_BundledDisc>0) Then
        shoppingcartrow.set "xSellBundleDiscount", scCurSign &  "-" & money(pcPrdOrd_BundledDisc)
        subTotal=subTotal - cdbl(pcPrdOrd_BundledDisc)
    End If
    '// END 10th Row - Cross Sell Bundle Discount
    
    
    '// START 11th Row - Cross Sell Bundle Subtotal
    'If (pcCartArray(f,27)>0) And (pcCartArray(f,28)>0) Then
    If (pcPrdOrd_BundledDisc>0) Then
        'shoppingcartrow.set "xSellBundleSubTotal", scCurSign &  money(pRowPrice - pcPrdOrd_BundledDisc)
        shoppingcartrow.set "productSubTotal", scCurSign &  money(pRowPrice - pcPrdOrd_BundledDisc)
        shoppingcartrow.set "DAproductSubTotal", scCurSign &  money((pRowPrice - pcPrdOrd_BundledDisc)/1.2)
    End If
    'End If
    '// START 11th Row - Cross Sell Bundle Subtotal



    '// START 12th Row - Gift Wrapping
    If pGWOpt <> "0" Then

        GWmsg = ""
        'GWmsg = "" & dictLanguage.Item(Session("language")&"_orderverify_36a") & " "

        query="select pcPE_IDProduct from pcProductsExc where pcPE_IDProduct=" & pidProduct
        Set rsG = server.CreateObject("ADODB.RecordSet")
        Set rsG = connTemp.execute(query)
        If Not rsG.Eof Then    
            'GWmsg = GWmsg & dictLanguage.Item(Session("language")&"_orderverify_38a")
        Else
            
            If (pGWOpt="") or (pGWOpt="0") Then                
                'GWmsg = GWmsg & dictLanguage.Item(Session("language")&"_orderverify_37a")                    
            Else

                query="select pcGW_OptName,pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
                set rsG2=server.CreateObject("ADODB.RecordSet")
                set rsG2=connTemp.execute(query)

                If Not rsG2.Eof Then
                
                    'pcv_strOptName = rsG2("pcGW_OptName")
                    'pcv_strOptPrice = rsG2("pcGW_OptPrice")
                    GWmsg = GWmsg & pGWText & " - " & scCurSign & money(pGWPrice)
                    shoppingcartrow.set "giftWrapMessage", GWmsg
                    
                End If
                Set rsG2 = Nothing
              
            End If
                
        End If
        Set rsG = Nothing

    End If 
    '// END 12th Row - Gift Wrapping


    set shoppingcartrows(f-1) = shoppingcartrow

    f = f + 1   
    count=count+1
    If pshippingAddress="" then
        'grab shipping address from shipping...
        pshippingAddress=pAddress
        pshippingAddress2=pAddress2
        pshippingCity=pCity
        pshippingState=pState
        pshippingZip=pZip
        pshippingCountryCode=pCountryCode
    End if

 
    rsOrdObj.movenext  
loop
'////////////////////////////////////////////////////////////////////
'// END: CART LOOP
'////////////////////////////////////////////////////////////////////


jsonService.set "shoppingcartrow", shoppingcartrows


'// Promotions
'// NA in v4.7


'// Payment Total
dim payment, PaymentType, PayCharge
payment = split(ppaymentDetails,"||")
err.clear
on error resume next
PaymentType=payment(0)
If payment(1)="" then
    if err.number<>0 then
        PayCharge=0
    end if
    PayCharge=0
else
    PayCharge=payment(1)
end If
err.number=0

If Session("customerType")=1 Then
    If Len(PaymentType)>0 Then
        jsonService.set "paymentDescription", PaymentType
    End If
End If

If PayCharge>0 Then
    jsonService.set "paymentTotal", scCurSign &  money(PayCharge)   
    'subTotal=subTotal+PayCharge  
End If


'// Category Discounts
If pcv_CatDiscounts>"0" Then 
    jsonService.set "categoryDiscountTotal", scCurSign & "-" & money(pcv_CatDiscounts)
End If


'// Discount Table Row
if instr(pdiscountDetails,",") then
    DiscountDetailsArry=split(pdiscountDetails,",")
    intArryCnt = ubound(DiscountDetailsArry)
else
    intArryCnt = 0
end if
If len(pdiscountDetails)>0 Then
    If pdiscountDetails<>"No discounts applied." Then
        Redim discounts(intArryCnt)
        for k=0 to intArryCnt
            if intArryCnt=0 then
                pTempDiscountDetails=pdiscountDetails
            else
                pTempDiscountDetails=DiscountDetailsArry(k)
            end if
    
            if instr(pTempDiscountDetails,"- ||") then
                pcv_arryDiscounts = split(pTempDiscountDetails,"- ||")
                discountType = pcv_arryDiscounts(0)
                discountPrice = pcv_arryDiscounts(1)
            
                dim discount : set discount = JSON.parse("{}") 
                
                '// Name
                discount.set "name", discountType
    
                '// Price
                discount.set "price", scCurSign &  "-" & money(discountPrice)
                discount.set "DAprice", scCurSign &  "-" & money(discountPrice/1.2)
                
                Set discounts(k) = discount
                Set discount = Nothing
                
            end if
        Next
        jsonService.set "discounts", discounts
    End If
End If
    

'// Reward Points being used
If piRewardPoints>0 Then 
    jsonService.set "rewardPointsUsedLabel", piRewardPoints & " " & RewardsLabel
    jsonService.set "rewardPointsUsedTotal", scCurSign & "-" & money(piRewardValue)
End If 



'// Reward Points being accrued
If piRewardPointsCustAccrued > 0 Then 
    jsonService.set "rewardPointsAccrued", piRewardPointsCustAccrued
End If 


'// Sub Total
if subTotal>0 then
    jsonService.set "subTotalBeforeDiscounts", scCurSign &  money(subtotal)
    jsonService.set "DAsubTotalBeforeDiscounts", scCurSign &  money(subtotal/1.2)
end if

'// Gift Wrap Total
If pGWTotal>0 Then
    jsonService.set "giftWrapTotal", scCurSign & money(pGWTotal)
End If


'// Shipment Data
dim shipping, varShip, Shipper, Service, Postage, serviceHandlingFee
shipping = split(pshipmentDetails,",")
if ubound(shipping)>1 then
    if NOT isNumeric(trim(shipping(2))) then
        varShip="0"
        response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
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
		
		pcv_boolShowFilteredRates = pcf_ShowFilteredRates()
		
		if (pcv_boolShowFilteredRates="1") AND (Service<>"") then
			queryM="SELECT pcShippingMap.pcSM_Name FROM pcShippingMap INNER JOIN (pcSMRel INNER JOIN shipService ON pcSMRel.idshipservice=shipService.idshipservice) ON pcShippingMap.pcSM_ID=pcSMRel.pcSM_ID WHERE shipService.serviceDescription LIKE '%" & Service & "%'"
			set rsM=connTemp.execute(queryM)
			if not rsM.eof then
				Service=rsM("pcSM_Name")
			end if
			set rsM=nothing
		end if
    end if
else
    varShip="0"
end if 

if varShip<>"0" then
    jsonService.set "shippingMethod", Service
    jsonService.set "shipmentTotal", scCurSign & money(Postage)
    jsonService.set "DAshipmentTotal", scCurSign & money(Postage/1.2)
End If

If serviceHandlingFee>0 Then
    jsonService.set "serviceHandlingFee", scCurSign & money(serviceHandlingFee)
End If



'// Tax Data
if pord_VAT>0 then
else

    Redim taxes(session("taxCnt")) 
    Dim tax : Set tax = JSON.parse("{}") 

    if isNull(ptaxDetails) or trim(ptaxDetails)="" then

        If ptaxAmount > 0 Then
        
            tax.set "name", dictLanguage.Item(Session("language")&"_CustviewOrd_12")
            tax.set "amount", scCurSign & money(ptaxAmount)
            Set taxes(0) = tax
        
        End If

    Else

        dim taxArray, taxDesc
        taxArray=split(ptaxDetails,",")
        
        If ubound(taxArray)>0 Then

            for i=0 to ubound(taxArray)
				if Instr(taxArray(i),"|")>0 then
					taxDesc=split(taxArray(i),"|")
					if taxDesc(0)<>"" then
						Set tax = JSON.parse("{}")
						If ccur(taxDesc(1))>0 Then
							tax.set "name", taxDesc(0) & ":"
							tax.set "amount", scCurSign& money(taxDesc(1))
						End If
			
						Set taxes(i) = tax
					
					end if
				end if
            next 

        End If
        
    End If
    
    jsonService.Set "taxes", taxes
    Set tax = Nothing

End If '// If ptaxVAT<>"1" And pTaxAmount>0 Then


'// Gift Certificates
If GCDetails<>"" Then

    GCArry=split(GCDetails,"|g|")
    intArryCnt=ubound(GCArry)

    Redim giftCerts(intArryCnt)    
    for k=0 to intArryCnt

        if GCArry(k)<>"" then

            GCInfo = split(GCArry(k),"|s|")
            if GCInfo(2)="" OR IsNull(GCInfo(2)) then
                GCInfo(2)=0
            end if
            
            If Cdbl(GCInfo(2)) <> 0 Then         
                Dim giftcert : Set giftcert = JSON.parse("{}")
                giftcert.set "name", dictLanguage.Item(Session("language")&"_CustviewOrd_15A") & " " & GCInfo(1) & " (" & GCInfo(0) & ")"
                giftcert.set "amount", scCurSign & "-" & money(GCInfo(2))
                Set giftCerts(k) = giftcert
                Set giftcert = Nothing
            End If                    

        End If '// If (GCArr(i) <> "") And (Cdbl(pSubTotal) > 0) Then

    Next '// For i=0 To ubound(GCArr)
   jsonService.Set "giftCerts", giftCerts
    
End If '// If (savGCs<>"") Then


'// VAT
If pord_VAT>0 Then

    If boolVAT > 0 Then
        jsonService.Set "vatName", dictLanguage.Item(Session("language")&"_orderverify_35")  
        jsonService.Set "vatTotal", scCurSign & money(pord_VAT) 
         jsonService.Set "DAvatTotal", scCurSign & money(pord_VAT) 
   Else
        jsonService.Set "vatName", dictLanguage.Item(Session("language")&"_orderverify_42") 
       jsonService.Set "vatTotal", scCurSign & money(0) 
        jsonService.Set "DAvatTotal", scCurSign & money(0) 
    End If
    
End If


'// Currency Format
If scDecSign = "," Then
    jsonService.set "decimal", ","
Else
    jsonService.set "decimal", "."
End If

'// Currency
jsonService.set "currencySymbol", scCurSign  

'// Total
jsonService.set "total", scCurSign &  money(ptotal)

      
'// Timestamp
jsonService.set "date", showDateFrmt(Date())

   
response.Clear()
Response.write( JSON.stringify(jsonService, null, 2) & vbNewline )
set Info = nothing
call closeDb()
response.End()
%>
