<% 
'// START - Check for SSL and redirect to SSL login if not already on HTTPS
call storeSSLRedirect("2")
'// END - check for SSL

'// Create Redirect URL
Dim strRedirectSSL
strRedirectSSL="onepagecheckout.asp"
if scSSL="1" AND scIntSSLPage="1" then
	strRedirectSSL=replace((scSslURL&"/"&scPcFolder&"/pc/onepagecheckout.asp"),"//","/")
	strRedirectSSL=replace(strRedirectSSL,"https:/","https://")
	strRedirectSSL=replace(strRedirectSSL,"http:/","http://")
end if
%>

<%
'// Clear One Page Checkout progress
session("pcPay_PayPalExp_OrderTotal") = ""

'// Clear Shipping Sessions
session("availableShipStr")=""
session("provider")=""


'// Clear Express Checkout
if Request("cmd")="_express-checkout" then
	session("ExpressCheckoutPayment")=""	
end if

'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************



'*****************************************************************************************************
' START: Save Cart
'*****************************************************************************************************

%><!--#include file="inc_SaveShoppingCart.asp"--><%

'*****************************************************************************************************
' END: Save Cart
'*****************************************************************************************************



'*****************************************************************************************************
' START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
' END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************



'*****************************************************************************************************
' START: PAGE ON LOAD
'*****************************************************************************************************


'// Cart Count Validation
if countCartRows(pcCartArray, pcCartIndex)=0 then
 	response.redirect "msg.asp?message=1"     
end if


'// Sotck Level Validation
Dim strCCSLCheck

strCCSLcheck = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)

'If Len(Trim(strCCSLCheck))>0 Then
'   response.write "<div class=""pcErrorMessage"">"
'   response.write dictLanguage.Item(Session("language")&"_alert_19") & strCCSLcheck
'   response.write "</div>"
'End If


'// Duplicate Order Validation
If len(session("GWOrderID"))>0 Then

	Dim pcv_intOrderID
	pcv_intOrderID = session("GWOrderID")
	pcv_intOrderID = pcv_intOrderID - int(scPre)

    query="SELECT orderStatus FROM orders WHERE orderStatus>1 AND idOrder=" & pcv_intOrderID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	If NOT rs.eof Then
		set rs=nothing
		call closedb()
		session("SFClearCartURL") = "msg.asp?message=308"
		response.Redirect("CustLOb.asp")
        response.End()
	End If
	set rs=nothing
 
End If

'// Discount Configuration
Dim TurnOffDiscountCodesWhenHasSale, HavePrdsOnSale
TurnOffDiscountCodesWhenHasSale = scDisableDiscountCodes
'=1: True - Default
'=0: False
HavePrdsOnSale=0

'*****************************************************************************************************
' END: PAGE ON LOAD
'*****************************************************************************************************



'*****************************************************************************************************
' START - Load Gift Wrapping settings
'*****************************************************************************************************

query="select pcGWSet_Show,pcGWSet_OverviewCart,pcGWSet_HTMLCart from pcGWSettings"
Set rstemp = server.CreateObject("ADODB.RecordSet")
Set rstemp = connTemp.execute(query)
If Not rstemp.eof Then
	pcv_GW = rstemp("pcGWSet_Show")
    pcv_Overview = rstemp("pcGWSet_OverviewCart")
    pcv_GWDetails = rstemp("pcGWSet_HTMLCart")
    
	If pcv_GW = "0" Then
		pcv_GW = ""
	End If
	session("Cust_GW") = pcv_GW
	
	if pcv_Overview = "0" then
		pcv_Overview = ""
	end if
	session("Cust_GWText") = pcv_Overview

	If trim(pcv_GWDetails) <> "" Then 
        pcv_GWDetails = replace(pcv_GWDetails, "&quot;", chr(34))
    End If
    
Else
	session("Cust_GW")=""
	session("Cust_GWText")=""
End If
Set rstemp = Nothing 

'*****************************************************************************************************
' END - Load Gift Wrapping settings
'*****************************************************************************************************
%>