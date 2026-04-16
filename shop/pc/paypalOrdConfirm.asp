<%@LANGUAGE="VBScript"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="inc_sb.asp"-->
<% 
Dim str
Dim objHttp

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: SEND SUBSCRIPTION REQUESTS TO SB
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error Resume Next

'// Forward to SB
str = Request.Form
set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
xml.open "POST", gv_RootURL & "/PostBack/PayPalSilentProcessor.asp", false
xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xml.send(str)			
if err.number<>0 then
	pcResultErrorMsg = err.description
end if
strStatus = xml.Status		
err.number=0
err.clear
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: SEND SUBSCRIPTION REQUESTS TO SB
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim arrNames(46), arrValues(46)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: REQUEST THE IPN VALUES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
arrNames(0)="Receiver_Email"
arrNames(1)="Business"
arrNames(2)="Item_Name"
arrNames(3)="Item_Number"
arrNames(4)="Quantity"
arrNames(5)="Invoice"
arrNames(6)="Custom"
arrNames(7)="Option_Name1"
arrNames(8)="Option_Selection1"
arrNames(9)="Option_Name2" 
arrNames(10)="Option_Selection2" 
arrNames(11)="Num_Cart_Items" 
arrNames(12)="Payment_Status" 
arrNames(13)="Pending_Reason" 
arrNames(14)="Payment_Date" 
arrNames(15)="Settle_Amount" 
arrNames(16)="Settle_Currency" 
arrNames(17)="Exchange_Rate" 
arrNames(18)="Payment_Gross" 
arrNames(19)="Payment_Fee" 
arrNames(20)="Mc_Gross" 
arrNames(21)="Mc_Fee" 
arrNames(22)="Mc_Currency" 
arrNames(23)="Tax" 
arrNames(24)="Txn_Id" 
arrNames(25)="Txn_Type" 
arrNames(26)="Reason_Code" 
arrNames(27)="For_Auction" 
arrNames(28)="Auction_Buyer_Id" 
arrNames(29)="Auction_Close_Date" 
arrNames(30)="Auction_Multi_Item" 
arrNames(31)="Memo" 
arrNames(32)="First_Name" 
arrNames(33)="Last_Name"
arrNames(34)="Address_Street" 
arrNames(35)="Address_City" 
arrNames(36)="Address_State" 
arrNames(37)="Address_Zip" 
arrNames(38)="Address_Country" 
arrNames(39)="Address_Status" 
arrNames(40)="Payer_Email" 
arrNames(41)="Payer_Id" 
arrNames(42)="Payer_Status" 
arrNames(43)="Payment_Type"
arrNames(44)="Notify_Version" 
arrNames(45)="Verify_Sign" 
arrNames(46)="Parent Transaction ID" 

arrValues(0)=getUserInput(Request.Form("receiver_email"),150)
arrValues(1)=getUserInput(Request.Form("business"),150)
arrValues(2)=getUserInput(Request.Form("item_name"),35)
arrValues(3)=getUserInput(Request.Form("item_number"),10)
arrValues(4)=getUserInput(Request.Form("quantity"),5)
arrValues(5)=getUserInput(Request.Form("invoice"),10)
arrValues(6)=getUserInput(Request.Form("custom"),150)
arrValues(7)=getUserInput(Request.Form("option_name1"),150)
arrValues(8)=getUserInput(Request.Form("option_selection1"),150)
arrValues(9)=getUserInput(Request.Form("option_name2"),150)
arrValues(10)=getUserInput(Request.Form("option_selection2"),150)
arrValues(11)=getUserInput(Request.Form("num_cart_items"),10)
arrValues(12)=getUserInput(Request.Form("payment_status"),50)
arrValues(13)=getUserInput(Request.Form("pending_reason"),50)
arrValues(14)=getUserInput(Request.Form("payment_date"),50)
arrValues(15)=getUserInput(Request.Form("settle_amount"),25)
arrValues(16)=getUserInput(Request.Form("settle_currency"),25)
arrValues(17)=getUserInput(Request.Form("exchange_rate"),25)
arrValues(18)=getUserInput(Request.Form("payment_gross"),25)
arrValues(19)=getUserInput(Request.Form("payment_fee"),25)
arrValues(20)=getUserInput(Request.Form("mc_gross"),25)
arrValues(21)=getUserInput(Request.Form("mc_fee"),25)
arrValues(22)=getUserInput(Request.Form("mc_currency"),25)
arrValues(23)=getUserInput(Request.Form("tax"),25)
arrValues(24)=getUserInput(Request.Form("txn_id"),150)
arrValues(25)=getUserInput(Request.Form("txn_type"),25)
arrValues(26)=getUserInput(Request.Form("reason_code"),50)
arrValues(27)=getUserInput(Request.Form("for_auction"),50)
arrValues(28)=getUserInput(Request.Form("auction_buyer_id"),50)
arrValues(29)=getUserInput(Request.Form("auction_close_date"),50)
arrValues(30)=getUserInput(Request.Form("auction_multi_item"),50)
arrValues(31)=getUserInput(Request.Form("memo"),150)
arrValues(32)=getUserInput(Request.Form("first_name"),50)
arrValues(33)=getUserInput(Request.Form("last_name"),50)
arrValues(34)=getUserInput(Request.Form("address_street"),50)
arrValues(35)=getUserInput(Request.Form("address_city"),50)
arrValues(36)=getUserInput(Request.Form("address_state"),50)
arrValues(37)=getUserInput(Request.Form("address_zip"),10)
arrValues(38)=getUserInput(Request.Form("address_country"),50)
arrValues(39)=getUserInput(Request.Form("address_status"),25)
arrValues(40)=getUserInput(Request.Form("payer_email"),150)
arrValues(41)=getUserInput(Request.Form("payer_id"),150)
arrValues(42)=getUserInput(Request.Form("payer_status"),25)
arrValues(43)=getUserInput(Request.Form("payment_type"),25)
arrValues(44)=getUserInput(Request.Form("notify_version"),25)
arrValues(45)=getUserInput(Request.Form("verify_sign"),150)
arrValues(46)=getUserInput(Request.Form("parent_txn_id"),150)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: REQUEST THE IPN VALUES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: SET LOCAL VALUES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Cheat Sheet:
' - pcv_OrderID = The order ID + the offset
' - pcv_OrderNum = Either the visible order ID or order code,
'                  depending how it's done.
' - pcv_trueOrderID = The actual database order ID
pcv_OrderNum=arrValues(3)
if pcv_OrderNum="" then
	pcv_OrderNum=arrValues(5)
end if
if pcv_OrderNum="" then
	pcv_OrderNum=getUserInput(Request("pcOID"),150)
end if
pcv_OrderID=pcv_OrderNum
pcv_trueOrderID=int(pcv_OrderNum)-scpre
pcv_PaymentStatus=arrValues(12)
pcv_PendingReason=arrValues(13)
pcv_gwTransID = arrValues(24)
pcv_gwTransParentID = arrValues(46)
if pcv_gwTransParentID="" then
	pcv_gwTransParentID = pcv_gwTransID
end if

session("GWSessionID")=randomNumber(99999999)

' read post from PayPal system and add 'cmd'
str = Request.Form & "&cmd=_notify-validate"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: SET LOCAL VALUES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: OPTIONAL LOGGING and DEBUGGING
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
dim pcv_Debug, pcv_Logging

pcv_Debug=0 'Change to 1 to view values sent back from paypal - for testing only
pcv_Logging=0

if pcv_Debug=1 then
	Response.Write "<div style='text-align: center'>"
	Response.Write "<strong>Response For Your Credit Card Transaction</strong>"
	For i=0 To UBound(arrValues)
		If arrValues(i) > "" Then
			Response.Write "<div style='width: 100%; clear: both'>"
			Response.Write "<div style='width: 45%; float: left; padding-right: 10px;'>Array #: "&i&" :" & arrNames(i) & ":&nbsp;</div>"
			Response.Write "<div style='width: 45%; float: left; padding-right: 10px;'>" & arrValues(i) & "</div>"
			Response.Write "</div>"
		End If
	Next
	Response.Write "</div>"
	response.End()
end if

if pcv_Logging=1 then
	TrackBug("Date: " & now())
	TrackBug("Order ID: " & pcv_trueOrderID)
	TrackBug("Order Number: " & pcv_OrderNum)
	TrackBug("Transaction ID: " & pcv_gwTransID)
	TrackBug("Transaction Parent ID: " & pcv_gwTransParentID)
	FOR i=0 TO UBOUND(arrValues)
		IF arrValues(i) > "" THEN
			TrackBug("Array #: " & i & " :" & arrNames(i) & " - " & arrValues(i))
		END IF		
	NEXT
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: OPTIONAL LOGGING and DEBUGGING
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: CHECK FOR INVALID ORDER TYPE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

query="SELECT paymentCode FROM orders WHERE idOrder="& pcv_trueOrderID & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
  call LogErrorToDatabase()
  set rs=nothing
  call closedb()
  response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if NOT rs.eof then
  paymentCode = trim(rs("paymentCode"))
  
	if lcase(paymentCode) <> "paypal" then
		if pcv_Logging=1 then
			TrackBug("Encountered Invalid IPN request with payment code: " & paymentCode)
		end if
		response.End()
	end if
end if
set rs=nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: CHECK FOR INVALID ORDER TYPE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: FRAUD CHECK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if pcv_trueOrderID<>"" AND UCase(pcv_PaymentStatus)="COMPLETED" then
	query="SELECT total FROM orders WHERE idOrder="& pcv_trueOrderID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rs.eof then
		tempOrderTotal = rs("total")
		pcv_gross = arrValues(20)
		If (pcv_gross<>"") AND (tempOrderTotal<>"") Then
			if CCur(pcv_gross) < CCur(tempOrderTotal) then
				'// send email to admin
				%>  <!--#include file="../includes/sendAlarmEmail.asp" -->  <%
				'// display message to customer			
				set rs = nothing
				call closedb()
				Session("message") = dictLanguage.Item(Session("language")&"_paypalOrdConfirm_D")
				response.redirect "msgb.asp"
			end if
		End If
	end if
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: FRAUD CHECK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: CHECK FOR DUPLICATE IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT orders.gwTransID FROM orders WHERE gwTransID='"& pcv_gwTransID &"' OR gwTransID='"& pcv_gwTransParentID &"';"
set rs=server.CreateObject("ADODB.RecordSet")
if pcv_Logging=1 then
	TrackBug("The IPN Query: " & query)
end if
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if NOT rs.eof then
	set rs=nothing
	call closedb()
	if pcv_trueOrderID<>"" then
		'// orderComplete.asp expects database order ID
		session("idOrder")=pcv_trueOrderID
		response.redirect "orderComplete.asp"
	else
		session("GWSessionID")=""
		contURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
		contURL=replace(contURL,"https:/","https://")
		contURL=replace(contURL,"http:/","http://")	
		response.redirect contURL
	end if	
	response.End() '// IPN is resending
else
	if pcv_gwTransID="" AND pcv_trueOrderID<>"" then
		'// orderComplete.asp expects database order ID
	  session("idOrder")=pcv_trueOrderID
	  response.redirect "orderComplete.asp"
	end if
end if
set rs=nothing
if pcv_Logging=1 then
	TrackBug("The Error Description: " & err.description)
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: CHECK FOR DUPLICATE IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: GATEWAY SPECIFIC VARIABLES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT Pay_To, PP_Currency, PP_Sandbox FROM paypal WHERE ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcv_PayTo=rs("Pay_To")
pcv_PPCurrency=rs("PP_Currency")
pcv_PP_Sandbox=rs("PP_Sandbox")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: GATEWAY SPECIFIC VARIABLES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: VALIDATE IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
if pcv_PP_Sandbox=1 then
	objHttp.open "POST", "https://www.sandbox.paypal.com/cgi-bin/webscr", false  '// SandBox Testing	
else
	objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false  '// LIVE
end if
objHttp.setRequestHeader "Host", "www.paypal.com"
objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHttp.setRequestHeader "Content-Length", Len(str)
objHttp.Send str
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: VALIDATE IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: PROCESS IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim Status, Result
Status=objHttp.status
Result=objHttp.ResponseText

'// Check notification validation
if (Status <> 200 ) then
'// Now we see if the payment is pending, verified, or denied 
elseif (objHttp.responseText="VERIFIED") then
	
	session("GWCustomerID")=pcf_RestoreCustomer(pcv_trueOrderID)
	
	if ucase(pcv_PaymentStatus)="COMPLETED" then
		session("GWAuthCode")=pcv_PaymentStatus
		session("GWTransId")=pcv_gwTransID
		session("GWOrderId")=pcv_trueOrderID	'// gwReturn.asp expects database order ID
		
		pcf_AuthorizeCustomer session("GWSessionID"), session("GWOrderId") 
        call closeDb()
		response.Redirect("gwReturn.asp?s=true&gw=PayPal&GWAuthCode="&session("GWAuthCode")&"&GWOrderId="&session("GWOrderId")&"&GWSessionID="&session("GWSessionID")&"&GWTransId="&session("GWTransId")&"&GWCustomerID="&session("GWCustomerID")&"")
		response.end
	elseif ucase(pcv_PaymentStatus)="PENDING" then
		If ucase(pcv_PendingReason)="AUTHORIZATION" Then
			pcv_PaymentStatus = "Authorized"
		End If

		'// SAVE PENDING REASON TO DB
		session("GWAuthCode")=pcv_PaymentStatus
		session("GWTransId")=pcv_gwTransID
		session("GWOrderId")=pcv_trueOrderID	'// gwReturn.asp expects database order ID
			
		'// FLAG AS PENDING
		session("GWTransType")="P"
		pcf_AuthorizeCustomer session("GWSessionID"), session("GWOrderId")  
        call closeDb()
		response.Redirect("gwReturn.asp?s=true&gw=PayPal&GWAuthCode="&session("GWAuthCode")&"&GWOrderId="&session("GWOrderId")&"&GWSessionID="&session("GWSessionID")&"&GWTransId="&session("GWTransId")&"&GWCustomerID="&session("GWCustomerID")&"")
		response.end
	end if
elseif (objHttp.responseText="INVALID" OR objHttp.responseText="FAILED") then
	'// If we get an Invalid response from PayPal, then the payment is messed up and we notify the customer
	if ucase(pcv_PaymentStatus)="COMPLETED" then
		session("GWAuthCode")=pcv_PaymentStatus
		session("GWTransId")=""
		session("GWOrderId")=pcv_trueOrderID	'// gwReturn.asp expects database order ID
		
		pcf_AuthorizeCustomer session("GWSessionID"), session("GWOrderId")  
        call closeDb()
		response.Redirect("gwReturn.asp?s=true&gw=PayPal&GWAuthCode="&session("GWAuthCode")&"&GWOrderId="&session("GWOrderId")&"&GWSessionID="&session("GWSessionID")&"&GWTransId="&session("GWTransId")&"&GWCustomerID="&session("GWCustomerID")&"")
		response.end
	else
		response.redirect "msg.asp?message=73" 
	end if  
else 
'// error
end if
set objHttp=nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: PROCESS IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Public Function TrackBug(message)	
	pcv_tmpLogFile=server.MapPath("header_wrapper.asp")
	pcv_tmpLogFile=left(pcv_tmpLogFile,instr(pcv_tmpLogFile,"\pc\"))
	pcv_tmpLogFile=pcv_tmpLogFile&"includes/paypal_Log.pclog"
    logFilename = pcv_tmpLogFile	
	Dim oFs
	Dim oTextFile
	Set oFs = Server.createobject("Scripting.FileSystemObject")
	Const ioMode = 8
	Set oTextFile = oFs.openTextFile(logFilename, ioMode, True)
	oTextFile.writeLine "Tracking Report: " & message
	oTextFile.close
	Set oTextFile = Nothing
	Set oFS = Nothing	
End Function


Private Function pcf_AuthorizeCustomer(gwTransID, orderID)
	'// Temporarily use the "gwTransID" field to verify the authenticity of the query on "gwReturn.asp"	
	query="UPDATE orders SET gwTransID='"& gwTransID &"' WHERE idOrder=" & orderID &";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing
End Function


Private Function pcf_RestoreCustomer(orderID)
	'// Restore the Customer From the OrderID	
	query="SELECT orders.idCustomer FROM orders WHERE idOrder=" & orderID &";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pidCustomer=rs("idCustomer")		
	pcf_RestoreCustomer=pidCustomer
	set rs=nothing
End Function


Public Function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
End Function
%>
