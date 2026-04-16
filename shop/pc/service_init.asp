<%
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
' START: Display Settings
'*****************************************************************************************************
' 1) The following variable controls the size of the small image shown in the cart
' Move to the Control Panel in a future version

Dim pcIntSmImgWidth
pcIntSmImgWidth = 35

' 2) The following varaible controls whether the SKU is shown in the cart or not.
' Move to the Control Panel in a future version

Dim pcIntShowSku
pcIntShowSku = 1 ' Change to 0 to hide the SKU

'*****************************************************************************************************
' END: Display Settings
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

'// SubscriptionBridge
Dim EditSB
EditSB=0
if (Session("SBEditOrder")<>"") AND (Session("SBEditOrderID")<>"") then
    EditSB=1
end if

'// Calculate Product Promotions
%>
<!--#include file="inc_CalPromotions.asp"-->
<%

'// ONE PAGE CHECKOUT SETTING
'// --> Hide company address, customer billing and shipping addresses in the Order Preview section.

Dim pcIntHideAddresses
pcIntHideAddresses = 0 ' Addresses are hidden

'// ONE PAGE CHECKOUT SETTING - END


Dim TurnOffDiscountCodesWhenHasSale, HavePrdsOnSale

TurnOffDiscountCodesWhenHasSale=scDisableDiscountCodes
'=1: True - Default
'=0: False


HavePrdsOnSale=0



'GGG Add-on start
%>
<!--#include file="ggg_inc_calGW.asp" -->
<% 
intGCIncludeShipping=GC_INCSHIPPING

intTaxExemptZoneFlag="1" 'Change to 0 if you want to tax any tax zone exempt products when they are added to the cart with taxable products. "1" will ensure that tax zone exempt products are never taxed for that zone.  
Dim GiftWrapPaymentTotal
GiftWrapPaymentTotal=0
%>
<%
'GGG Add-on end

'SB S
Dim pcIsSubscription, StrandSub 
pcIsSubscription = False
'pcIsSubscription = session("pcIsSubscription")

Dim pcv_sbTax
pcv_sbTax=getUserInput(request("sbTax"),0)
'SB E

If Not len(pcCartIndex)>0 Then
    pcCartIndex=session("pcCartIndex")
End If

'// GET CUSTOMER SESSION DATA
IsCartSaved = false

'*****************************************************************************************************
' END: PAGE ON LOAD
'*****************************************************************************************************
%>