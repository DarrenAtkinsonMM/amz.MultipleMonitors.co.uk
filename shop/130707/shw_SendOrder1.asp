<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by Netsource Commerce. ProductCart, its source code, the ProductCart name and logo are property of Netsource Commerce. Copyright 2001-2010. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Netsource Commerce. To contact Netsource Commerce, please visit www.productcart.com.
%>
<%pageTitle="SHIPWIRE Shipping Wizard" %>
<% response.Buffer=true %>
<% section="orders" %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
' Validate order ID
pcv_IdOrder=getUserInput(request("idorder"),10)
if not validNum(pcv_IdOrder) then pcv_IdOrder=0
if pcv_IdOrder=0 then 
	call closeDb()
	response.redirect "menu.asp"
end if

if request("action")<>"send" then 
	call closeDb()
	response.redirect "menu.asp"
end if
tmpPrdList=""
tmpShipDetails=""
PrdCount=request("count")
cPrd=-1
For i=1 to PrdCount
	If request("C" & i)="1" then
		tmpID=request("IDPrd" & i)
		query="SELECT Products.sku,Products.idProduct, ProductsOrdered.quantity,Products.Description FROM Products INNER JOIN ProductsOrdered ON Products.idProduct=ProductsOrdered.idProduct WHERE ProductsOrdered.idorder=" & pcv_IdOrder & " AND ProductsOrdered.idProductOrdered=" & tmpID & " AND pcPrdOrd_Shipped=0;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			cPrd=cPrd+1
			tmpSKU=rs("SKU")
			tmpQty=rs("Quantity")
			tmpName=rs("Description")
			tmpPrdList=tmpPrdList & "<Item num=""" & cPrd & """>" & "<Code>" & tmpSKU & "</Code>" & "<Quantity>" & tmpQty & "</Quantity></Item>"
			tmpShipDetails= tmpShipDetails & tmpName & " (" & tmpSKU & ") - Qty: " & tmpQty & "<br>"
		end if
		set rs=nothing
	End if
Next

if tmpPrdList="" then 
	call closeDb()
	response.redirect "shw_SendOrder.asp?idOrder=" & pcv_IdOrder
end if
tmpAddressInfo=""
query="SELECT IdCustomer,Address, address2, city, state, stateCode, zip, CountryCode,ShippingFullName,shippingAddress,shippingAddress2,shippingStateCode,shippingState,shippingCity,shippingCountryCode,shippingZip,pcOrd_shippingPhone,pcOrd_ShippingEmail FROM Orders WHERE idOrder=" & pcv_IdOrder & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	if rs("shippingAddress")<>"" then
		tmpAddressInfo=tmpAddressInfo & "<AddressInfo type=""ship"">" &_
		"<Name>" &_
			"<Full>" & rs("ShippingFullName") & "</Full>" &_
		"</Name>" &_
		"<Address1>" & rs("shippingAddress") & "</Address1>" &_
		"<Address2>" & rs("shippingAddress2") & "</Address2>" &_
		"<City>" & rs ("shippingCity") & "</City>" &_
		"<State>" & rs("shippingStateCode") & rs("shippingState") & "</State>" &_
		"<Country>" & rs("shippingCountryCode") & "</Country>" &_
		"<Zip>" & rs("shippingZip") & "</Zip>" &_
		"<Phone>" & rs("pcOrd_shippingPhone") & "</Phone>" &_
		"<Email>" & rs("pcOrd_ShippingEmail") & "</Email>" &_
		"</AddressInfo>"
	else
		pidcustomer=rs("IdCustomer")
		query="SELECT [name],lastName,phone,email FROM customers WHERE idcustomer="& pidcustomer
		Set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			pname=rsQ("name")
			plastName=rsQ("lastName")
			pphone=rsQ("phone")
			pemail=rsQ("email")
		end if
		set rsQ=nothing
		tmpAddressInfo=tmpAddressInfo & "<AddressInfo type=""ship"">" &_
		"<Name>" &_
			"<Full>" & pname & " " & plastName & "</Full>" &_
		"</Name>" &_
		"<Address1>" & rs("Address") & "</Address1>" &_
		"<Address2>" & rs("Address2") & "</Address2>" &_
		"<City>" & rs ("City") & "</City>" &_
		"<State>" & rs("StateCode") & rs("State") & "</State>" &_
		"<Country>" & rs("CountryCode") & "</Country>" &_
		"<Zip>" & rs("Zip") & "</Zip>" &_
		"<Phone>" & pphone & "</Phone>" &_
		"<Email>" & pemail & "</Email>" &_
		"</AddressInfo>"
	end if
end if
set rs=nothing

call GetSHWSettings()

if (shwOnOff=0) then 
	call closeDb()
	response.redirect "shw_SendOrder.asp?idOrder=" & pcv_IdOrder
end if

tmpidorder=scpre+int(pcv_IdOrder)

tmpShwMethod=request("shwMethod")
if tmpShwMethod<>"" then
	tmpShwMethod="<Shipping>" & tmpShwMethod & "</Shipping>"
end if

xmlRequest="<" & shwRequestType1 & ">" &_
	"<EmailAddress>" & shwUser & "</EmailAddress>" &_
	"<Password>" & shwPass & "</Password>" &_
	"<Server>" & shwMode & "</Server>" &_
	"<Order id=""" & tmpidorder & """>" &_
	"<Warehouse>00</Warehouse>" & tmpAddressInfo & tmpShwMethod & tmpPrdList &_
	"</Order>" &_
	"</" & shwRequestType1 & ">"
xmlRequest = Server.URLEncode(xmlRequest)
	
xmlResult=SHWConnectServer(shipwireXmlUrl1,"POST","","",shwRequestName1 & "=" & xmlRequest)

call SHWGetRequestStatus()

tmpWarning=SHWGetWarningList()
tmpShipInfo=SHWGetOrderShipInfo()
tmpOrdInfo=SHWGetSentOrderInfo()
tmpExcInfo=SHWGetOrderExcInfo()
tmpErrorList=SHWGetErrorList()


IF xmlResult<>"OK" THEN%>
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessage">
			Cannot send selected products to SHIPWIRE Server.<br>
			<a href="shw_SendOrder.asp?idOrder=<%=pcv_IdOrder%>">Click here</a> to try again.
		</div>
	</td>
</tr>
</table>
<%ELSE
IF (shwStatus<>"0") OR (tmpErrorList<>"") THEN%>
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessage">
			Sent selected products to SHIPWIRE Server but received error message(s).<br>
			<a href="shw_SendOrder.asp?idOrder=<%=pcv_IdOrder%>">Click here</a> to try again.
		</div>
	</td>
</tr>
</table>
<%ELSE%>
<table class="pcCPcontent">
<tr>
	<td>
		<%if (shwDupID<>"") then
			HaveErrors=1
			shwOrdID=shwDupID
		else
			if (tmpExcInfo<>"") then
				HaveErrors=2
			else
				HaveErrors=0
			end if
		end if%>
		<div <%if HaveErrors=1 then%>class="pcCPmessageInfo"<%else%><%if HaveErrors=2 then%>class="pcCPmessage"<%else%>class="pcCPmessageSuccess"<%end if%><%end if%>>
			Sent selected products to SHIPWIRE Server successfully.
			<%if (shwDupID<>"") then%>
			<br>But this SHIPWIRE Order appears to be a duplicate of <%=shwDupID%>. So, it has been ignored.
			<%end if%>
			<%if (shwOrdStatus<>"ACCEPTED") then%>
				<br>But this SHIPWIRE Order isn't accepted yet. Please log-in to your <a href="http://www.shipwire.com" target="_blank">Shipwire account</a> to fix it.
			<%end if%>
			<%if (tmpExcInfo<>"") then%>
				<br>You can see the "Exceptions" area below for some reasons.
			<%end if%>
		</div>
		<br>
		<input type="button" class="btn btn-default"  class="btn btn-primary" value="Back to Order Details" onclick="location='ordDetails.asp?id=<%=pcv_IdOrder%>';">
		<br>
		<br>
	</td>
</tr>
</table>
<%
if (shwOrdID<>"") then
	query="SELECT pcSWO_ID FROM pcShipwireOrders WHERE idOrder=" & pcv_IdOrder & " AND pcSWO_ShipwireID like '" & shwOrdID & "';"
	set rs=connTemp.execute(query)
	if rs.eof then
		query="INSERT INTO pcShipwireOrders (idOrder,pcSWO_ShipwireID,pcSWO_ShipwireDetails) VALUES (" & pcv_IdOrder & ",'" & shwOrdID & "','" & tmpShipDetails & "');"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	set rs=nothing
	For i=1 to PrdCount
		If request("C" & i)="1" then
			tmpID=request("IDPrd" & i)
			query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1 WHERE idOrder=" & pcv_IdOrder & " AND idProductOrdered=" & tmpID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		End if
	Next
	query="SELECT idProductOrdered FROM ProductsOrdered WHERE idOrder=" & pcv_IdOrder & " AND pcPrdOrd_Shipped=0;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		query="UPDATE Orders SET OrderStatus=7 WHERE idOrder=" & pcv_IdOrder & ";"
		set rs=connTemp.execute(query)
	else
		query="UPDATE Orders SET OrderStatus=4 WHERE idOrder=" & pcv_IdOrder & ";"
		set rs=connTemp.execute(query)
	end if
	set rs=nothing
end if

END IF%>
<table class="pcCPcontent">
<tr>
	<th colspan="2">SHIPWIRE Response</th>
</tr>
<%if tmpErrorList<>"" then%>
<tr>
	<td colspan="2"><b>Error(s)</b></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td width="90%"><%=tmpErrorList%></td>
</tr>
<%end if%>
<%if tmpOrdInfo<>"" then%>
<tr>
	<td colspan="2"><b>SHIPWIRE Order Information</b></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td width="90%"><%=tmpOrdInfo%></td>
</tr>
<%end if%>
<%if tmpExcInfo<>"" then%>
<tr>
	<td colspan="2"><b>Exceptions</b></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td width="90%"><%=tmpExcInfo%></td>
</tr>
<%end if%>
<%if tmpWarning<>"" then%>
<tr>
	<td colspan="2"><b>Warnings</b></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td width="90%"><%=tmpWarning%></td>
</tr>
<%end if%>
<%if tmpShipInfo<>"" then%>
<tr>
	<td colspan="2"><b>Shipping</b></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td width="90%"><%=tmpShipInfo%></td>
</tr>
<%end if%>
</table>
<%END IF%>
<!--#include file="AdminFooter.asp"-->