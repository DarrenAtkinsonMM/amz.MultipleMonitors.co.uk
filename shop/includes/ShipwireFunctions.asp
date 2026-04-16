<%'Shipwire Integration Functions
Dim srvXmlHttpSHW
Dim SHW_ErrMsg, SHW_SuccessMsg, SHW_XmlDoc
Dim shipwireXmlUrl1,shipwireXmlUrl2,shipwireXmlUrl3

SHW_ErrMsg=""
SHW_SuccessMsg=""

tmpscXML=".3.0"
'// Shipwire Constants

'//Production
private const shipwirePrdXmlUrl1 = "https://api.shipwire.com/exec/FulfillmentServices.php"
private const shipwirePrdXmlUrl2 = "https://api.shipwire.com/exec/TrackingServices.php"
private const shipwirePrdXmlUrl3 = "https://api.shipwire.com/exec/InventoryServices.php"

'//Test
private const shipwireTestXmlUrl1 = "https://api.beta.shipwire.com/exec/FulfillmentServices.php"
private const shipwireTestXmlUrl2 = "https://api.beta.shipwire.com/exec/TrackingServices.php"
private const shipwireTestXmlUrl3 = "https://api.beta.shipwire.com/exec/InventoryServices.php"

private const shwRequestName1 = "OrderListXML"
private const shwRequestName2 = "TrackingUpdateXML"
private const shwRequestName3 = "InventoryUpdateXML"

private const shwRequestType1 = "OrderList"
private const shwRequestType2 = "TrackingUpdate"
private const shwRequestType3 = "InventoryUpdate"

private const shwResponseType1 = "SubmitOrderResponse"
private const shwResponseType2 = "TrackingUpdateResponse"
private const shwResponseType3 = "InventoryUpdateResponse"

shwStatus = "Status"
shwError = "ErrorMessage"
shwWarnList= "WarningList"
shwWarn = "Warning"
shwTransID = "TransactionId"

Dim SHWMaxRequestTime,SHWStopHTTPRequests

Dim shwUser,shwPass,shwMode,shwOnOff,shwReXML,iRoot,shwDupID,shwOrdID, shwOrdStatus,shwStatus
Dim HadshwSettings
Dim SHWshipped,SHWshipper,SHWshipperFullName
Dim SHWshipDate,SHWexpectedDeliveryDate
Dim SHWhandling,SHWshipping,SHWpackaging,SHWtotal
Dim SHWreturned,SHWreturnDate,SHWreturnCondition
Dim SHWOrdURL,SHWmanuallyEdited
Dim SHWTrackingNumber,SHWTrackcarrier,SHWTrackURL

'maximum seconds for each HTTP request time
SHWMaxRequestTime=30

SHWStopHTTPRequests=0

HadshwSettings=0
shwUser=""
shwPass=""
shwMode="Test"
shwOnOff=0

shipwireXmlUrl1=shipwireTestXmlUrl1
shipwireXmlUrl2=shipwireTestXmlUrl2
shipwireXmlUrl3=shipwireTestXmlUrl3

Sub GetSHWSettings()
Dim queryQ,rsQ

if HadshwSettings=0 then

queryQ="SELECT pcSWS_Username,pcSWS_Password,pcSWS_Mode,pcSWS_OnOff FROM pcShipwireSettings;"
set rsQ=connTemp.execute(queryQ)
if not rsQ.eof then
	shwUser=rsQ("pcSWS_Username")
	shwPass=enDeCrypt(rsQ("pcSWS_Password"), scCrypPass)
	shwMode="Test"
	shipwireXmlUrl1=shipwireTestXmlUrl1
	shipwireXmlUrl2=shipwireTestXmlUrl2
	shipwireXmlUrl3=shipwireTestXmlUrl3
	if rsQ("pcSWS_Mode")="1" then
		shwMode="Production"
		shipwireXmlUrl1=shipwirePrdXmlUrl1
		shipwireXmlUrl2=shipwirePrdXmlUrl2
		shipwireXmlUrl3=shipwirePrdXmlUrl3
	end if
	if rsQ("pcSWS_OnOff")="1" then
		shwOnOff=1
	end if
end if
set rsQ=nothing
HadshwSettings=1
end if

End Sub

Public Function pcf_SHWIsResponseGood()
	On Error Resume Next
	
	If srvXmlHttpSHW.readyState <> 4  Then
		if not IsNull(SHWMaxRequestTime) then
			TransactionReady = srvXmlHttpSHW.waitForResponse(SHWMaxRequestTime)
		else
			TransactionReady = srvXmlHttpSHW.waitForResponse(5)
		end if
		If TransactionReady = False Then
			pcf_SHWIsResponseGood=False
			srvXmlHttpSHW.Abort
			Exit Function
		End If
	End If  	

	If Err.Number <> 0 then
		pcf_SHWIsResponseGood=False
		srvXmlHttpSHW.Abort
		Exit Function
	Else
		If (srvXmlHttpSHW.readyState <> 4) Then
			pcf_SHWIsResponseGood=False
			srvXmlHttpSHW.Abort
			Exit Function
		Else
			pcf_SHWIsResponseGood=True
			Exit Function
		End If	
	End If		

	If Err.Number <> 0 then
		pcf_SHWIsResponseGood=False
		srvXmlHttpSHW.Abort
		Exit Function
	End If
	
	On Error Goto 0
End Function

Function XMLReplace(tmpData)
Dim tmp1
	tmp1=tmpData
	tmp1=replace(tmp1,"&","&amp;")
	tmp1=replace(tmp1,"<","&lt;")
	tmp1=replace(tmp1,">","&gt;")
	tmp1=replace(tmp1,"""","&quot;")
	tmp1=replace(tmp1,"'","&apos;")
	XMLReplace=tmp1
End Function

Function XMLNumber(tmpData)
Dim tmp1
	tmp1=tmpData
	if scDecSign="," then
		tmp1=replace(tmp1,".","")
		tmp1=replace(tmp1,",",".")
	else
		tmp1=replace(tmp1,",","")
	end if
	XMLNumber=tmp1
End Function

Function SHWConnectServer(tmpURL,tmpMethod,tmpContentType,tmpSOAPHead,tmpData)
Dim rersult1,tmpStatus

on error resume next
Set srvXmlHttpSHW=nothing

IF SHWStopHTTPRequests<>"1" THEN
	Set srvXmlHttpSHW = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
	if not IsNull(SHWMaxRequestTime) then
		srvXmlHttpSHW.open tmpMethod, tmpURL, True
	else
		srvXmlHttpSHW.open tmpMethod, tmpURL, False
	end if
	if tmpContentType="" then
		srvXmlHttpSHW.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	else
		srvXmlHttpSHW.setRequestHeader "Content-Type", tmpContentType
	end if
	srvXmlHttpSHW.setRequestHeader "CharSet", "UTF-8"
	if tmpSOAPHead<>"" then
	srvXmlHttpSHW.setRequestHeader "SOAPAction", tmpSOAPHead
	end if
	srvXmlHttpSHW.send tmpData
	
	if not IsNull(SHWMaxRequestTime) then
		if pcf_SHWIsResponseGood()=False then
			SHWConnectServer="TIMEOUT"
			exit function
			SHWStopHTTPRequests=1
		end if
	end if
	
	result1 = srvXmlHttpSHW.responseText

	if err.number<>0 then
		err.number=0
		err.description=""
		SHWConnectServer="ERROR"
		if not IsNull(SHWMaxRequestTime) then
			SHWStopHTTPRequests=1
		end if
		set srvXmlHttpSHW=nothing
	else
		tmpStatus=srvXmlHttpSHW.Status
		if (tmpStatus<>200) then
			if result1<>"" then
				if Instr(result1,"Response>")=0 then
					SHWConnectServer="ERROR"
					if not IsNull(SHWMaxRequestTime) then
						SHWStopHTTPRequests=1
					end if
				else
					SHWConnectServer="OK"
					Set shwReXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
					set shwReXML = srvXmlHttpSHW.responseXML
					set iRoot=shwReXML.documentElement
				end if
			else
				SHWConnectServer="ERROR"
				if not IsNull(SHWMaxRequestTime) then
					SHWStopHTTPRequests=1
				end if
			end if
		else
			SHWConnectServer="OK"
			Set shwReXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
			set shwReXML = srvXmlHttpSHW.responseXML
			set iRoot=shwReXML.documentElement
		end if
		set srvXmlHttpSHW=nothing
	end if
ELSE
	SHWConnectServer="TIMEOUT"
END IF
End Function

Function SHWFindStatusCode(tmpXML)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"<Status>")>0 then
			tmp1=split(tmpXML,"<Status>")
			tmp2=split(tmp1(1),"</Status>")
			SHWFindStatusCode=Clng(tmp2(0))
		else
			if Instr(tmpXML,"&lt;Status&gt;")>0 then
				tmp1=split(tmpXML,"&lt;Status&gt;")
				tmp2=split(tmp1(1),"&lt;/Status&gt;")
				SHWFindStatusCode=Clng(tmp2(0))
			else
				SHWFindStatusCode=0
			end if
		end if
	else
		SHWFindStatusCode=0
	end if
End Function

Function SHWFindXMLValue(tmpXML,tmpName)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"<" & tmpName & ">")>0 then
			tmp1=split(tmpXML,"<" & tmpName & ">")
			tmp2=split(tmp1(1),"</" & tmpName & ">")
			SHWFindXMLValue=tmp2(0)
		else
			if Instr(tmpXML,"&lt;" & tmpName & "&gt;")>0 then
				tmp1=split(tmpXML,"&lt;" & tmpName & "&gt;")
				tmp2=split(tmp1(1),"&lt;/" & tmpName & "&gt;")
				SHWFindXMLValue=tmp2(0)
			else
				SHWFindXMLValue=""
			end if
		end if
	else
		SHWFindXMLValue=""
	end if
End Function

Function SHWFindErrMsg(tmpXML)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"<" & shwError & ">")>0 then
			tmp1=split(tmpXML,"<" & shwError & ">")
			tmp2=split(tmp1(1),"</" & shwError & ">")
			SHWFindErrMsg=GetSHWmsg(tmp2(0))
		else
			if Instr(tmpXML,"&lt;" & shwError & "&gt;")>0 then
				tmp1=split(tmpXML,"&lt;" & shwError & "&gt;")
				tmp2=split(tmp1(1),"&lt;/" & shwError & "&gt;")
				SHWFindErrMsg=GetSHWmsg(tmp2(0))
			else
				SHWFindErrMsg=""
			end if
		end if
	else
		SHWFindErrMsg=""
	end if
End Function

Function GetSHWmsg(tmpMsg)
Dim tmpStr1,tmpStr2
	tmpStr1=tmpMsg
	if InStr(tmpStr1,"<![CDATA[")>0 then
		tmpStr2=split(tmpStr1,"<![CDATA[")
		tmpStr1=replace(tmpStr2(1),"]]>","")
	end if
	GetSHWmsg=tmpStr1
End Function

Function SHWGetErrorList()
Dim tmpStr,i,tmpStr1
	tmpStr=""
		set myElementList1 = shwReXML.getElementsByTagName(shwError)
		For i = 0 To (myElementList1.length - 1)
			tmpStr1=GetSHWmsg(myElementList1.Item(i).Text)
			tmpStr=tmpStr & "<li>" & tmpStr1 & "</li>"
		Next
		if tmpStr<>"" then
			tmpStr="<ul>" & tmpStr & "</ul>"
		end if
	SHWGetErrorList=tmpStr
End Function

Function SHWGetWarningList()
Dim tmpStr,i,tmpStr1
	shwDupID=""
	tmpStr=""
		set myElementList1 = shwReXML.getElementsByTagName(shwWarn)
		For i = 0 To (myElementList1.length - 1)
			tmpStr1=GetSHWmsg(myElementList1.Item(i).Text)
			tmpStr=tmpStr & "<li>" & tmpStr1 & "</li>"
			if Instr(ucase(tmpStr1),ucase("Order appears to be a duplicate of "))>0 then
				shwDupID=replace(ucase(tmpStr1),ucase("Order appears to be a duplicate of "),"")
				shwDupID=trim(replace(ucase(shwDupID),ucase("; ignoring this order"),""))
			end if
		Next
		if tmpStr<>"" then
			tmpStr="<ul>" & tmpStr & "</ul>"
		end if
	SHWGetWarningList=tmpStr
End Function

Function SHWGetOrderShipInfo()
Dim tmpStr,i,parentNode
	tmpStr=""
		Set parentNode = iRoot.selectSingleNode("OrderInformation/Order/Shipping")
		If parentNode is Nothing then
		Else
			Set ChildNodes = parentNode.childNodes
			For Each strNode In ChildNodes
				if strNode.text<>"" then
					tmpStr=tmpStr & "<li>" & strNode.nodeName & ": " & GetSHWmsg(strNode.text) & "</li>"
				end if
			Next
			if tmpStr<>"" then
				tmpStr="<ul>" & tmpStr & "</ul>"
			end if
		End if
	SHWGetOrderShipInfo=tmpStr
End Function

Function SHWGetOrderExcInfo()
Dim tmpStr,i,parentNode,tmpStr1
	tmpStr=""
		Set parentNode = iRoot.selectSingleNode("OrderInformation/Order/Exception")
		If parentNode is Nothing then
		Else
			if parentNode.text<>"" then
				tmpStr1=split(GetSHWmsg(parentNode.text),";")
				For i=lbound(tmpStr1) to ubound(tmpStr1)
					if trim(tmpStr1(i))<>"" then
						tmpStr=tmpStr & "<li>" & trim(tmpStr1(i)) & "</li>"
					end if
				Next
				if tmpStr<>"" then
					tmpStr="<ul>" & tmpStr & "</ul>"
				end if
			end if
		End if
	SHWGetOrderExcInfo=tmpStr
End Function

Function SHWGetSentOrderInfo()
Dim tmpStr,i,parentNode
	tmpStr=""
		Set parentNode = iRoot.selectSingleNode("OrderInformation/Order")
		If parentNode is Nothing then
		Else
			Set sNodeAttributes = parentNode.attributes
			For Each strAtt In sNodeAttributes
				if strAtt.value<>"" then
					Select Case ucase(strAtt.name)
						Case "NUMBER":	tmpStr=tmpStr & "<li>Your Order ID#: " & strAtt.value & "</li>"
						Case "ID":	tmpStr=tmpStr & "<li>SHIPWIRE Order ID#: " & strAtt.value & "</li>"
						shwOrdID=strAtt.value
						Case "STATUS":
							shwOrdStatus=ucase(strAtt.value)
							if ucase(strAtt.value)<>"ACCEPTED" then
								tmpStr=tmpStr & "<li>SHIPWIRE Status: <font color=red><b>" & ucase(strAtt.value) & "</b></font></li>"
							else
								tmpStr=tmpStr & "<li>SHIPWIRE Status: <font color=green><b>" & ucase(strAtt.value) & "</b></font></li>"
							end if
						Case Else: tmpStr=tmpStr & "<li>" & strAtt.name & ": " & strAtt.value & "</li>"
					End Select
				end if
			Next
			if tmpStr<>"" then
				tmpStr="<ul>" & tmpStr & "</ul>"
			end if
		End if
	SHWGetSentOrderInfo=tmpStr
End Function


Function CheckExistTag(tagName)
Dim tmpNode
	Set tmpNode=iRoot.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		CheckExistTag=False
	Else
		CheckExistTag=True
	End if
End Function

Sub SHWGetRequestStatus()
Dim tmpNode
	shwStatus=""
	Set tmpNode=iRoot.selectSingleNode("Status")
	shwStatus=tmpNode.Text
End Sub

Function CheckExistTagEx(parentNode,tagName)
Dim tmpNode
	Set tmpNode=parentNode.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		CheckExistTagEx=False
	Else
		CheckExistTagEx=True
	End if
End Function

Function SHWGetInventoryStatus(tmpSKU)
Dim tmpQty,tmpData,xmlRequest
Dim tmpNode

tmpQty=-1

call GetSHWSettings()

if shwOnOff=1 then

xmlRequest="<" & shwRequestType3 & ">" &_
 "<Username>" & shwUser & "</Username>" &_
 "<Password>" & shwPass & "</Password>" &_
 "<Server>" & shwMode & "</Server>" &_
 "<ProductCode>" & tmpSKU & "</ProductCode>" &_
 "</" & shwRequestType3 & ">"
 xmlRequest = Server.URLEncode(xmlRequest)
 
xmlResult=SHWConnectServer(shipwireXmlUrl3,"POST","","",shwRequestName3 & "=" & xmlRequest)

if xmlResult="OK" then
	call SHWGetRequestStatus()
	if shwStatus<>"0" then
		tmpQty=-1
	else
		Set tmpNode=iRoot.selectSingleNode("TotalProducts")
		if tmpNode.Text="0" then
			tmpQty=-1
		else
			Set tmpNode=iRoot.selectSingleNode("Product")
			Set sNodeAttributes = tmpNode.attributes
			CorrectSKU=0
			For Each strAtt In sNodeAttributes
				if strAtt.name="code" then
					if ucase(strAtt.value)=ucase(tmpSKU) then
						CorrectSKU=1
						Exit For
					end if
				end if
			Next
			if CorrectSKU=1 then
				For Each strAtt In sNodeAttributes
					if strAtt.name="quantity" then
						tmpQty=strAtt.value
						Exit For
					end if
				Next
			else
				tmpQty=-1
			end if
		end if
	end if
end if

end if

SHWGetInventoryStatus=tmpQty

End Function

Sub SHWSyncAllInventoryStatus()
Dim tmpQty,tmpData,xmlRequest
Dim tmpNode,i,tmpSKU,queryQ,rsQ

tmpQty=0

call GetSHWSettings()

if shwOnOff=1 then

xmlRequest="<" & shwRequestType3 & ">" &_
 "<Username>" & shwUser & "</Username>" &_
 "<Password>" & shwPass & "</Password>" &_
 "<Server>" & shwMode & "</Server>" &_
 "</" & shwRequestType3 & ">"
 xmlRequest = Server.URLEncode(xmlRequest)
 
xmlResult=SHWConnectServer(shipwireXmlUrl3,"POST","","",shwRequestName3 & "=" & xmlRequest)

if xmlResult="OK" then
	call SHWGetRequestStatus()
	if shwStatus="0" then
		set myElementList1 = shwReXML.getElementsByTagName("Product")
		For i = 0 To (myElementList1.length - 1)
			Set tmpNode=myElementList1.Item(i)
			Set sNodeAttributes = tmpNode.attributes
			tmpSKU=""
			tmpQty=-1
			HaveBoth=0
			For Each strAtt In sNodeAttributes
				if strAtt.name="code" then
					tmpSKU=strAtt.value
					HaveBoth=HaveBoth+1
				end if
				if strAtt.name="quantity" then
					tmpQty=strAtt.value
					HaveBoth=HaveBoth+1
				end if
				if HaveBoth=2 then
					exit for
				end if
			Next
			if (tmpSKU<>"") AND (Clng(tmpQty)>=0) then
				queryQ="UPDATE Products SET stock=" & tmpQty & " WHERE SKU Like '" & tmpSKU & "';"
				set rsQ=connTemp.execute(queryQ)
				set rsQ=nothing
				call pcs_hookStockChanged("", tmpSKU)
			end if
		Next
		Set tmpNode=nothing
		Set sNodeAttributes=nothing
		Set strAtt=nothing
		Set myElementList1=nothing
	end if
else
	shwStatus="ERROR"
end if

end if

End Sub


Function SHWGetPackStatus(tmpShipwireID)
Dim tmpData,xmlRequest
Dim tmpNode

tmpData=""

SHWshipped=""
SHWshipper=""
SHWshipperFullName=""
SHWshipDate=""
SHWexpectedDeliveryDate=""
SHWhandling=""
SHWshipping=""
SHWpackaging=""
SHWtotal=""
SHWreturned=""
SHWreturnDate=""
SHWreturnCondition=""
SHWOrdURL=""
SHWmanuallyEdited=""
SHWTrackingNumber=""
SHWTrackcarrier=""
SHWTrackURL=""

call GetSHWSettings()

if shwOnOff=1 then

xmlRequest="<" & shwRequestType2 & ">" &_
 "<Username>" & shwUser & "</Username>" &_
 "<Password>" & shwPass & "</Password>" &_
 "<Server>" & shwMode & "</Server>" &_
 "<ShipwireId>" & tmpShipwireID & "</ShipwireId>" &_
 "</" & shwRequestType2 & ">"
 xmlRequest = Server.URLEncode(xmlRequest)
 
xmlResult=SHWConnectServer(shipwireXmlUrl2,"POST","","",shwRequestName2 & "=" & xmlRequest)

if xmlResult="OK" then
	call SHWGetRequestStatus()
	if shwStatus<>"0" then
		tmpData=""
	else
		Set tmpNode=iRoot.selectSingleNode("Order")
		Set sNodeAttributes = tmpNode.attributes
		CorrectOrder=0
		For Each strAtt In sNodeAttributes
			if strAtt.name="shipwireId" then
				if ucase(strAtt.value)=ucase(tmpShipwireID) then
					CorrectOrder=1
					Exit For
				end if
			end if
		Next
		if CorrectOrder=1 then
			For Each strAtt In sNodeAttributes
				Select Case strAtt.name
					Case "shipped": SHWshipped=strAtt.value
					Case "shipper": SHWshipper=strAtt.value
					Case "shipperFullName": SHWshipperFullName=strAtt.value
					Case "shipDate": SHWshipDate=strAtt.value
					Case "expectedDeliveryDate": SHWexpectedDeliveryDate=strAtt.value
					Case "handling": SHWhandling=strAtt.value
					Case "shipping": SHWshipping=strAtt.value
					Case "packaging": SHWpackaging=strAtt.value
					Case "total": SHWtotal=strAtt.value
					Case "returned": SHWreturned=strAtt.value
					Case "returnDate": SHWreturnDate=strAtt.value
					Case "returnCondition": SHWreturnCondition=strAtt.value
					Case "href": SHWOrdURL=strAtt.value
					Case "manuallyEdited": SHWmanuallyEdited=strAtt.value
				End Select
			Next
			
			Set tmpNode=iRoot.selectSingleNode("Order/TrackingNumber")
			if not (tmpNode is Nothing) then
				SHWTrackingNumber=tmpNode.value
				Set tmpNode=iRoot.selectSingleNode("Order/TrackingNumber")
				Set sNodeAttributes = tmpNode.attributes
				For Each strAtt In sNodeAttributes
					Select Case strAtt.name
						Case "carrier": SHWTrackcarrier=strAtt.value
						Case "href": SHWTrackURL=strAtt.value
					End Select
				Next
			end if
			
			tmpData="<tr><td colspan=""2""><b>Shipping Details</b></td></tr>"
			tmpData=tmpData & "<tr><td>Shipped: </td><td><b>" & SHWshipped & "</b></td></tr>"
			if SHWshipper<>"" then
				tmpData=tmpData & "<tr><td>Shipper: </td><td>" & SHWshipper & "</td></tr>"
			end if
			if SHWshipperFullName<>"" then
				tmpData=tmpData & "<tr><td nowrap>Shipper Fullname: </td><td>" & SHWshipperFullName & "</td></tr>"
			end if
			if SHWshipDate<>"" then
				tmpData=tmpData & "<tr><td nowrap>Ship Date: </td><td>" & SHWshipDate & "</td></tr>"
			end if
			if SHWexpectedDeliveryDate<>"" then
				tmpData=tmpData & "<tr><td nowrap>Expected Delivery Date: </td><td>" & SHWexpectedDeliveryDate & "</td></tr>"
			end if
			if SHWtotal<>"" then
				tmpData=tmpData & "<tr><td colspan=""2""><b>Fees</b></td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Shipping Fees: </td><td>" & SHWshipping & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Handling Fees: </td><td>" & SHWhandling & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Packaging Fees: </td><td>" & SHWpackaging & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Total: </td><td><b>" & SHWtotal & "</b></td></tr>"
			end if
			if SHWreturned="YES" then
				tmpData=tmpData & "<tr><td colspan=""2""><b>Return Information</b></td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Return Date: </td><td>" & SHWreturnDate & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Return Condition: </td><td>" & SHWreturnCondition & "</td></tr>"
			end if
			if (SHWOrdURL<>"") OR (SHWmanuallyEdited<>"") then
				tmpData=tmpData & "<tr><td colspan=""2""><b>Other Information</b></td></tr>"
				if SHWOrdURL<>"" then
					tmpData=tmpData & "<tr><td nowrap>SHIPWIRE Order URL: </td><td><a href=""" & SHWOrdURL & """ target=""_blank"">" & SHWOrdURL & "</href></td></tr>"
				end if
				if SHWmanuallyEdited<>"" then
					tmpData=tmpData & "<tr><td nowrap>Manually Edited: </td><td>" & SHWmanuallyEdited & "</td></tr>"
				end if
			end if
			if (SHWTrackingNumber<>"") then
				tmpData=tmpData & "<tr><td colspan=""2""><b>Tracking Information</b></td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Tracking Number: </td><td>" & SHWTrackingNumber & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Carrier: </td><td>" & SHWTrackcarrier & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Tracking URL: </td><td><a href=""" & SHWTrackURL & """ target=""_blank"">" & SHWTrackURL & "</href></td></tr>"
			end if
		else
			tmpData=""
		end if
	end if
end if

end if

SHWGetPackStatus=tmpData

End Function

Function SHWGetPackStatusCust(tmpShipwireID)
Dim tmpData,xmlRequest
Dim tmpNode

tmpData=""

SHWshipped=""
SHWshipper=""
SHWshipperFullName=""
SHWshipDate=""
SHWexpectedDeliveryDate=""
SHWhandling=""
SHWshipping=""
SHWpackaging=""
SHWtotal=""
SHWreturned=""
SHWreturnDate=""
SHWreturnCondition=""
SHWOrdURL=""
SHWmanuallyEdited=""
SHWTrackingNumber=""
SHWTrackcarrier=""
SHWTrackURL=""

call GetSHWSettings()

if shwOnOff=1 then

xmlRequest="<" & shwRequestType2 & ">" &_
 "<Username>" & shwUser & "</Username>" &_
 "<Password>" & shwPass & "</Password>" &_
 "<Server>" & shwMode & "</Server>" &_
 "<ShipwireId>" & tmpShipwireID & "</ShipwireId>" &_
 "</" & shwRequestType2 & ">"
 xmlRequest = Server.URLEncode(xmlRequest)
 
xmlResult=SHWConnectServer(shipwireXmlUrl2,"POST","","",shwRequestName2 & "=" & xmlRequest)

if xmlResult="OK" then
	call SHWGetRequestStatus()
	if shwStatus<>"0" then
		tmpData=""
	else
		Set tmpNode=iRoot.selectSingleNode("Order")
		Set sNodeAttributes = tmpNode.attributes
		CorrectOrder=0
		For Each strAtt In sNodeAttributes
			if strAtt.name="shipwireId" then
				if ucase(strAtt.value)=ucase(tmpShipwireID) then
					CorrectOrder=1
					Exit For
				end if
			end if
		Next
		if CorrectOrder=1 then
			For Each strAtt In sNodeAttributes
				Select Case strAtt.name
					Case "shipped": SHWshipped=strAtt.value
					Case "shipper": SHWshipper=strAtt.value
					Case "shipperFullName": SHWshipperFullName=strAtt.value
					Case "shipDate": SHWshipDate=strAtt.value
					Case "expectedDeliveryDate": SHWexpectedDeliveryDate=strAtt.value
					Case "handling": SHWhandling=strAtt.value
					Case "shipping": SHWshipping=strAtt.value
					Case "packaging": SHWpackaging=strAtt.value
					Case "total": SHWtotal=strAtt.value
					Case "returned": SHWreturned=strAtt.value
					Case "returnDate": SHWreturnDate=strAtt.value
					Case "returnCondition": SHWreturnCondition=strAtt.value
					Case "href": SHWOrdURL=strAtt.value
					Case "manuallyEdited": SHWmanuallyEdited=strAtt.value
				End Select
			Next
			
			Set tmpNode=iRoot.selectSingleNode("Order/TrackingNumber")
			if not (tmpNode is Nothing) then
				SHWTrackingNumber=tmpNode.value
				Set tmpNode=iRoot.selectSingleNode("Order/TrackingNumber")
				Set sNodeAttributes = tmpNode.attributes
				For Each strAtt In sNodeAttributes
					Select Case strAtt.name
						Case "carrier": SHWTrackcarrier=strAtt.value
						Case "href": SHWTrackURL=strAtt.value
					End Select
				Next
			end if
			
			tmpData="<tr><td colspan=""2""><b>Shipping Details</b></td></tr>"
			tmpData=tmpData & "<tr><td>Shipped: </td><td><b>" & SHWshipped & "</b></td></tr>"
			if SHWshipper<>"" then
				tmpData=tmpData & "<tr><td>Shipper: </td><td>" & SHWshipper & "</td></tr>"
			end if
			if SHWshipperFullName<>"" then
				tmpData=tmpData & "<tr><td nowrap>Shipper Fullname: </td><td>" & SHWshipperFullName & "</td></tr>"
			end if
			if SHWshipDate<>"" then
				tmpData=tmpData & "<tr><td nowrap>Ship Date: </td><td>" & SHWshipDate & "</td></tr>"
			end if
			if SHWexpectedDeliveryDate<>"" then
				tmpData=tmpData & "<tr><td nowrap>Expected Delivery Date: </td><td>" & SHWexpectedDeliveryDate & "</td></tr>"
			end if
			if SHWreturned="YES" then
				tmpData=tmpData & "<tr><td colspan=""2""><b>Return Information</b></td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Return Date: </td><td>" & SHWreturnDate & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Return Condition: </td><td>" & SHWreturnCondition & "</td></tr>"
			end if
			if (SHWOrdURL<>"") then
				tmpData=tmpData & "<tr><td colspan=""2""><b>Other Information</b></td></tr>"
				if SHWOrdURL<>"" then
					tmpData=tmpData & "<tr><td nowrap>SHIPWIRE Order URL: </td><td><a href=""" & SHWOrdURL & """ target=""_blank"">" & SHWOrdURL & "</href></td></tr>"
				end if
			end if
			if (SHWTrackingNumber<>"") then
				tmpData=tmpData & "<tr><td colspan=""2""><b>Tracking Information</b></td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Tracking Number: </td><td>" & SHWTrackingNumber & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Carrier: </td><td>" & SHWTrackcarrier & "</td></tr>"
				tmpData=tmpData & "<tr><td nowrap>Tracking URL: </td><td><a href=""" & SHWTrackURL & """ target=""_blank"">" & SHWTrackURL & "</href></td></tr>"
			end if
		else
			tmpData=""
		end if
	end if
end if

end if

SHWGetPackStatusCust=tmpData

End Function


%>