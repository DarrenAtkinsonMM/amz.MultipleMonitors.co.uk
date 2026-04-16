<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/CashbackConstants.asp"--> 
<!--#include file="chkPrices.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="../includes/pcSBHelperInc.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
Dim xmlResponse
Dim xmlAcknowledgment
Dim biData
biData = Request.BinaryRead(Request.TotalBytes)
Dim nIndex
For nIndex = 1 to LenB(biData)
    xmlResponse = xmlResponse & Chr(AscB(MidB(biData,nIndex,1)))
Next

Dim messageRecognizer,XMLErr,SBGuid,ReErr,SaveOrdErr

XMLErr=0
ReErr=0
SaveOrdErr=0
SBGuid=""

IF xmlResponse="" THEN
	XMLErr=100
END IF

IF XMLErr=0 THEN
	on error resume next
	
	Dim domResponseObj,domMcCallbackObjRoot
	Set domResponseObj = Server.CreateObject("Msxml2.DOMDocument.3.0")
	domResponseObj.loadXml xmlResponse	

	messageRecognizer = domResponseObj.documentElement.tagName
	
	if err.number<>0 then
		XMLErr=101
		err.number=0
		err.description=""
	end if

	if XMLErr=0 then
		Select Case messageRecognizer
			Case "SB_Callback":
			Case Else
				'Incorrect Response XML		
				XMLErr=1
		End Select
	end if

END IF

if XMLErr=0 then

	Dim pcv_strGUID,pcv_EventCode
	pcv_strGUID = ""
	pcv_EventCode=""

	pcv_strGUID = pcf_GetNode(xmlResponse, "Guid", "//SB_Callback")
	pcv_EventCode = pcf_GetNode(xmlResponse, "Event_Code", "//SB_Callback")
	
	'// Incorrect SB Event Code
	if (pcv_EventCode="") OR (pcv_EventCode<>"sb_n_4c" AND pcv_EventCode<>"sb_n_6b") then
		XMLErr=2
	end if
	
	'// Incorrect SB Guid
	if (XMLErr=0) AND (pcv_strGUID="") then
		XMLErr=3
	end if

end if

Set domResponseObj=nothing

Dim SBIDOrder,SBTerms

IF XMLErr=0 THEN
	SBGuid=pcv_strGUID
	SBIDOrder=0
		
	query="SELECT TOP 1 idOrder,SB_Terms FROM SB_Orders WHERE SB_Guid like '" & SBGuid & "' ORDER BY idOrder DESC;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		SBIDOrder=rs("idOrder")
		SBTerms=rs("SB_Terms")
	end if
	set rs=nothing
	
	query="SELECT orderDate FROM orders WHERE idOrder = " & SBIDOrder & " ORDER BY idOrder ASC;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		orderDate=rs("orderDate")
	end if
	set rs=nothing

	'*******************************'
	' START - Creating a New Order
	'*******************************'
	IF DateDiff("d", cdate(orderDate), Now()) > 1 THEN
		IF SBIDOrder>"0" THEN
			
			'Step 1 - Repeat Order/Generate shopping cart array
			ReErr=0
			%>
			<!--#include file="sb_inc_repeatorder.asp"-->
			<%
			
			'Step 2 - Save the cart from memory into a new Order
			If (ReErr=0) OR (ReErr=5) then
				SaveOrdErr=0%>
				<!--#include file="sb_inc_SaveOrd.asp"-->
				<%
			End if
		
		END IF
	ELSE
		response.write "Error: Ordered Date < 1 day."
	END IF
	
	call closedb()
ELSE
	response.write "XML Error: " & XMLErr
END IF

' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function

Function generateABC(keyLength)
Dim sDefaultChars
Dim iCounter
Dim sMyKeys
Dim iPickedChar
Dim iDefaultCharactersLength
Dim ikeyLength

	sDefaultChars="ABCDEFGHIJKLMNOPQRSTUVXYZ"
	ikeyLength=keyLength
	iDefaultCharactersLength = Len(sDefaultChars)
	Randomize
	For iCounter = 1 To ikeyLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)
		sMyKeys = sMyKeys & Mid(sDefaultChars,iPickedChar,1)
	Next
	generateABC = sMyKeys
End Function

Function generate123(keyLength)
Dim sDefaultChars
Dim iCounter
Dim sMyKeys
Dim iPickedChar
Dim iDefaultCharactersLength
Dim ikeyLength

	sDefaultChars="0123456789"
	ikeyLength=keyLength
	iDefaultCharactersLength = Len(sDefaultChars)
	Randomize
	For iCounter = 1 To ikeyLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)
		sMyKeys = sMyKeys & Mid(sDefaultChars,iPickedChar,1)
	Next
	generate123 = sMyKeys
End Function


Function pcf_GetNode(responseXML, nodeName, nodeParent)
	Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)				 
	myXmlDoc.loadXml(responseXML)
	'response.Write(nodeParent)
	Set Nodes = myXmlDoc.selectnodes(nodeParent)	
	For Each Node In Nodes	
		pcf_GetNode = pcf_CheckNode(Node,nodeName,"")				
	Next
	Set Node = Nothing
	Set Nodes = Nothing
	Set myXmlDoc = Nothing
End Function


Function pcf_CheckNode(Node,tagName,default)		
	Dim tmpNode
	Set tmpNode=Node.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		pcf_CheckNode=default
	Else
		pcf_CheckNode=Node.selectSingleNode(tagName).text
	End if
End Function
%>
