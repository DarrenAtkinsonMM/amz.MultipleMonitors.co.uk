<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%Dim pcStrPageName
pcStrPageName = "gwAmazonMWS.asp"%>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
tmpscXML=".3.0"
amzCurrencyCode="USD"
session("amzError")=""


'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if

if (session("GWOrderId")="") Or ((session("AmzOrderID")="") And (session("AmzBillAgreementID")="")) then
	response.redirect "onepagecheckout.asp"
end if

'//Declare and Retrieve Customer's IP Address	
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

If Len(session("AmzBillAgreementID")) > 0 Then
    QueryStr=""
    tmpTimeStamp=GenAmazonTimeStamp(UtcNow())
    QueryStr=QueryStr & "AWSAccessKeyId=" & pcAMZAccessKeyID
    QueryStr=QueryStr & "&Action=CreateOrderReferenceForId"
    QueryStr=QueryStr & "&Id=" & session("AmzBillAgreementID")
    QueryStr=QueryStr & "&IdType=BillingAgreement"
    QueryStr=QueryStr & "&SellerId=" & pcAMZSellerID
    QueryStr=QueryStr & "&SignatureMethod=HmacSHA256"
    QueryStr=QueryStr & "&SignatureVersion=2"
    QueryStr=QueryStr & "&Timestamp=" & tmpTimeStamp
    QueryStr=QueryStr & "&Version=2013-01-01"

    StringtoSign="POST" & vbLf
    StringtoSign=StringtoSign & pcAMZHost & vbLf
    StringtoSign=StringtoSign & pcAMZUI & vbLf
    StringtoSign=StringtoSign & QueryStr

    Set sha256 = GetObject( "script:" & Server.MapPath("sha256md5.txt") )
    StringtoSign=Server.URLEncode(sha256.b64_hmac_sha256(pcAMZSecretKey, StringtoSign))

    QueryStr=QueryStr & "&Signature=" & StringtoSign

    Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
    xml.open "POST", pcAMZEndPoint, False
    xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xml.send QueryStr
    strStatus = xml.Status

    'store the response
    strRetVal = xml.responseText

    Set ReXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
    Set ReXML = xml.responseXML
    Set iRoot = ReXML.documentElement

    Set tmpNode=iRoot.selectSingleNode("CreateOrderReferenceForIdResult/OrderReferenceDetails/AmazonOrderReferenceId")

    If (tmpNode is Nothing) OR (tmpNode.Text="") Then
        'Log failed transaction
		call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
	    response.redirect "msg.asp?message=311"
    else
        session("AmzOrderID") = tmpNode.Text
    end if
End If

'Add Order Total
QueryStr=""
tmpTimeStamp=GenAmazonTimeStamp(UtcNow())
QueryStr=QueryStr & "AWSAccessKeyId=" & pcAMZAccessKeyID
QueryStr=QueryStr & "&Action=SetOrderReferenceDetails"
QueryStr=QueryStr & "&AmazonOrderReferenceId=" & session("AmzOrderID")
QueryStr=QueryStr & "&OrderReferenceAttributes.OrderTotal.Amount=" & CCur(pcBillingTotal)
QueryStr=QueryStr & "&OrderReferenceAttributes.OrderTotal.CurrencyCode=" & amzCurrencyCode
QueryStr=QueryStr & "&OrderReferenceAttributes.PlatformId=A2UEDTDSCICEOS"
QueryStr=QueryStr & "&OrderReferenceAttributes.SellerNote="
QueryStr=QueryStr & "&OrderReferenceAttributes.SellerOrderAttributes.SellerOrderId=" & "Order-" & session("GWOrderId")
QueryStr=QueryStr & "&SellerId=" & pcAMZSellerID
QueryStr=QueryStr & "&SignatureMethod=HmacSHA256"
QueryStr=QueryStr & "&SignatureVersion=2"
QueryStr=QueryStr & "&Timestamp=" & tmpTimeStamp
QueryStr=QueryStr & "&Version=2013-01-01"

StringtoSign="POST" & vbLf
StringtoSign=StringtoSign & pcAMZHost & vbLf
StringtoSign=StringtoSign & pcAMZUI & vbLf
StringtoSign=StringtoSign & QueryStr

Set sha256 = GetObject( "script:" & Server.MapPath("sha256md5.txt") )
StringtoSign=Server.URLEncode(sha256.b64_hmac_sha256(pcAMZSecretKey, StringtoSign))

QueryStr=QueryStr & "&Signature=" & StringtoSign

Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
xml.open "POST", pcAMZEndPoint, False
xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xml.send QueryStr
strStatus = xml.Status

'store the response
strRetVal = xml.responseText

Set ReXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
Set ReXML = xml.responseXML
Set iRoot = ReXML.documentElement

Set tmpNode=iRoot.selectSingleNode("SetOrderReferenceDetailsResult/OrderReferenceDetails/OrderTotal/Amount")

If (tmpNode is Nothing) OR (tmpNode.Text="") Then
    'Log failed transaction
	call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
	response.redirect "msg.asp?message=311"
else
	if CCur(pcBillingTotal)<>CCur(tmpNode.Text) then
        'Log failed transaction
		call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
		response.redirect "msg.asp?message=311"
	end if
end if

'Confirm Order
QueryStr=""
tmpTimeStamp=GenAmazonTimeStamp(UtcNow())
QueryStr=QueryStr & "AWSAccessKeyId=" & pcAMZAccessKeyID
QueryStr=QueryStr & "&Action=ConfirmOrderReference"
QueryStr=QueryStr & "&AmazonOrderReferenceId=" & session("AmzOrderID")
QueryStr=QueryStr & "&SellerId=" & pcAMZSellerID
QueryStr=QueryStr & "&SignatureMethod=HmacSHA256"
QueryStr=QueryStr & "&SignatureVersion=2"
QueryStr=QueryStr & "&Timestamp=" & tmpTimeStamp
QueryStr=QueryStr & "&Version=2013-01-01"

StringtoSign="POST" & vbLf
StringtoSign=StringtoSign & pcAMZHost & vbLf
StringtoSign=StringtoSign & pcAMZUI & vbLf
StringtoSign=StringtoSign & QueryStr

Set sha256 = GetObject( "script:" & Server.MapPath("sha256md5.txt") )
StringtoSign=Server.URLEncode(sha256.b64_hmac_sha256(pcAMZSecretKey, StringtoSign))

QueryStr=QueryStr & "&Signature=" & StringtoSign

Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
xml.open "POST", pcAMZEndPoint, False
xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xml.send QueryStr
strStatus = xml.Status

'store the response
strRetVal = xml.responseText

'Authorize Payment
QueryStr=""
tmpTimeStamp=GenAmazonTimeStamp(UtcNow())
QueryStr=QueryStr & "AWSAccessKeyId=" & pcAMZAccessKeyID
QueryStr=QueryStr & "&Action=Authorize"
QueryStr=QueryStr & "&AmazonOrderReferenceId=" & session("AmzOrderID")
QueryStr=QueryStr & "&AuthorizationAmount.Amount=" & CCur(pcBillingTotal)
QueryStr=QueryStr & "&AuthorizationAmount.CurrencyCode=" & amzCurrencyCode
QueryStr=QueryStr & "&AuthorizationReferenceId=Authorize-Order-" & session("GWOrderId")
QueryStr=QueryStr & "&SellerAuthorizationNote=Authorize-Order-" & session("GWOrderId")
QueryStr=QueryStr & "&SellerId=" & pcAMZSellerID
QueryStr=QueryStr & "&SignatureMethod=HmacSHA256"
QueryStr=QueryStr & "&SignatureVersion=2"
QueryStr=QueryStr & "&Timestamp=" & tmpTimeStamp
if x_Mode="1" then
	QueryStr=QueryStr & "&TransactionTimeout=0"
else
	QueryStr=QueryStr & "&TransactionTimeout=60"
end if
QueryStr=QueryStr & "&Version=2013-01-01"

StringtoSign="POST" & vbLf
StringtoSign=StringtoSign & pcAMZHost & vbLf
StringtoSign=StringtoSign & pcAMZUI & vbLf
StringtoSign=StringtoSign & QueryStr

Set sha256 = GetObject( "script:" & Server.MapPath("sha256md5.txt") )
StringtoSign=Server.URLEncode(sha256.b64_hmac_sha256(pcAMZSecretKey, StringtoSign))

QueryStr=QueryStr & "&Signature=" & StringtoSign

Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
xml.open "POST", pcAMZEndPoint, False
xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xml.send QueryStr
strStatus = xml.Status

'store the response
strRetVal = xml.responseText

Set ReXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
Set ReXML = xml.responseXML
Set iRoot = ReXML.documentElement

Set tmpNode=iRoot.selectSingleNode("AuthorizeResult/AuthorizationDetails/AuthorizationStatus/State")

amzAuID=""

If (tmpNode is Nothing) OR (tmpNode.Text="") Then
    'Log failed transaction
	call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
	response.redirect "msg.asp?message=311"
else
	if (UCase(tmpNode.Text)<>"PENDING") AND (UCase(tmpNode.Text)<>"OPEN") then
		session("amzError")=tmpNode.Text
		response.redirect "msg.asp?message=312"
	else
		AutStatus=UCase(tmpNode.Text)
		Set tmpNode=iRoot.selectSingleNode("AuthorizeResult/AuthorizationDetails/AmazonAuthorizationId")
		If (tmpNode is Nothing) OR (tmpNode.Text="") Then
		Else
			amzAuID=tmpNode.Text
		End if
	end if
end if

amz_Capture="0"

'Capture Immediately
if (x_Mode="1") AND (AutStatus="OPEN") AND (amzAuID<>"") then
	QueryStr=""
	tmpTimeStamp=GenAmazonTimeStamp(UtcNow())
	QueryStr=QueryStr & "AWSAccessKeyId=" & pcAMZAccessKeyID
	QueryStr=QueryStr & "&Action=Capture"
	QueryStr=QueryStr & "&AmazonAuthorizationId=" & amzAuID
	QueryStr=QueryStr & "&CaptureAmount.Amount=" & CCur(pcBillingTotal)
	QueryStr=QueryStr & "&CaptureAmount.CurrencyCode=" & amzCurrencyCode
	QueryStr=QueryStr & "&CaptureReferenceId=Capture-Order-" & session("GWOrderId")
	QueryStr=QueryStr & "&SellerCaptureNote=Capture-Order-" & session("GWOrderId")
	QueryStr=QueryStr & "&SellerId=" & pcAMZSellerID
	QueryStr=QueryStr & "&SignatureMethod=HmacSHA256"
	QueryStr=QueryStr & "&SignatureVersion=2"
	QueryStr=QueryStr & "&Timestamp=" & tmpTimeStamp
	QueryStr=QueryStr & "&Version=2013-01-01"
	
	StringtoSign="POST" & vbLf
	StringtoSign=StringtoSign & pcAMZHost & vbLf
	StringtoSign=StringtoSign & pcAMZUI & vbLf
	StringtoSign=StringtoSign & QueryStr
	
	Set sha256 = GetObject( "script:" & Server.MapPath("sha256md5.txt") )
	StringtoSign=Server.URLEncode(sha256.b64_hmac_sha256(pcAMZSecretKey, StringtoSign))
	
	QueryStr=QueryStr & "&Signature=" & StringtoSign
	
	Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
	xml.open "POST", pcAMZEndPoint, False
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send QueryStr
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText
	response.write strRetVal

	Set ReXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
	Set ReXML = xml.responseXML
	Set iRoot = ReXML.documentElement
	
	Set tmpNode=iRoot.selectSingleNode("CaptureResult/CaptureDetails/CaptureStatus/State")
	
	amzAuID=""
	
	If (tmpNode is Nothing) OR (tmpNode.Text="") Then
	else
		if (UCase(tmpNode.Text)="COMPLETED") then
			amz_Capture="1"
		end if
	end if
end if

if amz_Capture="1" then
	session("GWTransType")=2
else
	session("GWTransType")=1
end if

'Log successful transaction
call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 1)

session("GWTransId")=session("AmzOrderID")
session("GWAuthCode")=amzAuID
		
Response.redirect "gwReturn.asp?s=true&gw=AMZ"

%>

<%
Function AmazonURLEnCode(tmpStr)
Dim tmp1

tmp1=tmpStr
tmp1=replace(Server.URLEncode(tmp1),"%2E",".")
tmp1=replace(tmp1,"%5F","_")
tmp1=replace(tmp1,"%2D","-")
tmp1=replace(tmp1,"%7E","~")
tmp1=replace(tmp1,"+","%20")
AmazonURLEnCode=tmp1

End Function

Function AddZero(tmpStr)
	if Clng(tmpStr)<10 then
		AddZero="0" & tmpStr
	else
		AddZero=tmpStr
	end if
End Function

Function GenAmazonTimeStamp(tmpDate)
Dim tmp1
	tmp1=Year(tmpDate) & "-" & AddZero(Month(tmpDate)) & "-" & AddZero(Day(tmpDate)) & "T" & AddZero(Hour(tmpDate)) & ":" & AddZero(Minute(tmpDate)) & ":" & AddZero(Second(tmpDate)) & "Z"
	GenAmazonTimeStamp=AmazonURLEnCode(tmp1)
End Function


Function UtcNow()
UtcNow = serverdate.toUTCString()
UtcNow = CDate(Replace(Right(UtcNow, Len(UtcNow) - Instr(UtcNow, ",")), "UTC", ""))
End Function
%>
<script language="JScript" runat="server">
var serverdate=new Date();
</script>
<!--#include file="footer_wrapper.asp"-->