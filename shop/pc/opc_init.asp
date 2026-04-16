<% Dim iAddDefaultPrice, iAddDefaultWPrice %>

<%
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "OnePageCheckout.asp"


'// Discount Configuration
Dim TurnOffDiscountCodesWhenHasSale, HavePrdsOnSale
TurnOffDiscountCodesWhenHasSale = scDisableDiscountCodes
'=1: True - Default
'=0: False
HavePrdsOnSale=0



'*******************************
' Pay Panel Open or Closed
'*******************************
Dim pcv_strPayPanel
pcv_strPayPanel = getUserInput(request("PayPanel"), 2)



'*******************************
' START: Check store on/off, start PC session, check affiliate ID
'*******************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*******************************
' END: Check store on/off, start PC session, check affiliate ID
'*******************************


'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
pcCartArray=Session("pcCartSession")
ppcCartIndex=Session("pcCartIndex")


if countCartRows(pcCartArray, ppcCartIndex)=0 then
	response.redirect "msg.asp?message=9" 
end if


%><!--#include file="inc_checkPrdQtyCart.asp"--><%
Call CheckALLCartStock()
%>
<!--#include file="DBsv.asp"-->
<%
'SB S
session("pcIsSubscription") = ""
pcIsSubscription = findSubscription(pcCartArray, ppcCartIndex)			
session("pcIsSubscription") = pcIsSubscription
'SB E

If session("customerType")=1 Then
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
		Session("message") = dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scWholesaleMinPurchase)& dictLanguage.Item(Session("language")&"_techErr_3") & "<BR><BR><a href=viewcart.asp>"& dictLanguage.Item(Session("language")&"_msg_back") &"</a>"
		response.redirect "msgb.asp"
	end if
Else
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then  
		Session("message") = dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scMinPurchase)&"<BR><BR><a href=viewcart.asp>"& dictLanguage.Item(Session("language")&"_msg_back") &"</a>"
		response.redirect "msgb.asp"
	end if
End If

if (session("ExpressCheckoutPayment") <> "YES") OR (session("PayWithAmazon")="YES") then
%>
<!--#include file="opc_checkpayment.asp"-->
<%
else
	if session("pcSFIdPayment")="" then
		session("pcSFIdPayment")=999999
	end if
end if

'//* Get Bill Information from database if Registered Customer
If Session("idCustomer")>0 Then
    query="SELECT idcustomer, customers.pcCust_Guest, customers.pcCust_VATID, customers.pcCust_SSN, [name], lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, address2, suspend, idCustomerCategory, customerType, RecvNews, fax, pcCust_Locked,pcCust_AgreeTerms FROM customers WHERE ((customers.idcustomer)="& session("idCustomer") &");"
    Set rs = connTemp.execute(query)
    If Not rs.Eof Then
        pcStrBillingFirstName=rs("name")
        pcStrBillingLastName=rs("lastName")
        pcStrBillingCompany=rs("customerCompany")
        pcStrBillingPhone=rs("phone")
        pcStrCustomerEmail=rs("email")
        pcStrBillingAddress=rs("address")
        pcStrBillingPostalCode=rs("zip")
        pcStrBillingStateCode=rs("stateCode")
        pcStrBillingProvince=rs("state")
        pcStrBillingCity=rs("city")
        pcStrBillingCountryCode=rs("countryCode")
        pcStrBillingAddress2=rs("address2")
        pcIntSuspend=rs("suspend")
        pcIntRecvNews=rs("RecvNews")
        pcStrBillingFax=rs("fax")
        pcStrBillingVATID=Trim(rs("pcCust_VATID"))
        pcStrBillingSSN=Trim(rs("pcCust_SSN"))
        pcAgreeTerms=rs("pcCust_AgreeTerms")
        If IsNull(pcAgreeTerms) Or pcAgreeTerms="" Then
            pcAgreeTerms=0
        End If
    End If
    Set rs = Nothing
End If


'//* Get Shipping Information from database if available
If Session("idCustomer")>0 Then
    If Not (session("pcSFIdDbSession") = "" OR session("pcSFRandomKey") = "") Then
        query="SELECT pcCustSession_ShippingCountryCode, pcCustSession_ShippingStateCode, pcCustSession_ShippingProvince FROM pcCustomerSessions "
        query = query & "WHERE (pcCustomerSessions.idDbSession=" & session("pcSFIdDbSession") & ") "
        query = query & "AND (pcCustomerSessions.randomKey=" & session("pcSFRandomKey") & ") "
        Set rs = connTemp.execute(query)
        If Not rs.Eof Then
            pcStrShippingCountryCode=rs("pcCustSession_ShippingCountryCode")
            pcStrShippingStateCode=rs("pcCustSession_ShippingStateCode")
            pcStrShippingProvince=rs("pcCustSession_ShippingProvince")
        End If
        Set rs = Nothing
    End If
End If


'// We'll check the cart stock levels on entry (plus each time the ajax panels slide)
Dim strCCSLCheck

strCCSLcheck = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)

If Len(Trim(strCCSLCheck))>0 Then
    response.redirect "viewcart.asp"
End If

'// START - Check for SSL and redirect to SSL login if not already on HTTPS
call storeSSLRedirect("1")
'// END - check for SSL

'// Set Required Fields / Defaults
If len(pcStrBillingCountryCode)=0 Then
    pcStrBillingCountryCode = scShipFromPostalCountry
End If
If len(pcStrShippingCountryCode)=0 Then
    pcStrShippingCountryCode = scShipFromPostalCountry
End If
pcv_isStateCodeRequired = true
pcv_isCountryCodeRequired = true

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

If (pcv_NOShippingAtAll = "2" And pcv_AlwAltShipAddress="0") Or (pcv_AlwAltShipAddress="1") Then
    displayShippingAddress = false
Else
    displayShippingAddress = true
End If

%>