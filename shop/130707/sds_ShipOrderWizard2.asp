<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pageTitle="Shipping Wizard" %>
<% response.Buffer=true %>
<% section="orders" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%
'// PACKAGE COUNT
pcv_strPackageCount=request("PackageCount")
pcv_strSessionPackageCount=Session("pcAdminPackageCount")
if pcv_strSessionPackageCount="" OR len(pcv_strPackageCount)>0 then
	pcPackageCount=pcv_strPackageCount
	Session("pcAdminPackageCount")=pcPackageCount
else
	pcPackageCount=pcv_strSessionPackageCount
end if
if pcPackageCount="" then
	pcPackageCount=1
end if

pcArraySize = (pcPackageCount -1)
						
'// GET ORDER ID
dim intResetSessions
intResetSessions=0
pcv_IdOrder=request("idorder")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_IdOrder)>0 then
	pcv_intOrderID=pcv_IdOrder
	'// Reset all sessions
	if pcv_strSessionOrderID<>pcv_IdOrder then
		intResetSessions=1
	end if
else
	pcv_intOrderID=pcv_strSessionOrderID
end if
Session("pcAdminOrderID")=pcv_intOrderID
'// REDIRECT
if pcv_intOrderID="" then
	call closeDb()
response.redirect "menu.asp"
end if

'// ITEM COUNT
pcv_count=Request("count")
if pcv_count="" then
	pcv_count=0
end if

'// CREATE THE ARRAY
Dim pcLocalArray()

'// SIZE THE ARRAY
ReDim pcLocalArray(pcArraySize)

'// POPULATE THE ARRAY
if request.form("submit")<>"" OR request.form("submit1")<>"" then
	For xPackageCount=0 to pcArraySize
		pcLocalArray(xPackageCount) = Request("pcAdminPrdList" & (xPackageCount+1))	
	Next 
else
	if Session("pcGlobalArray")<>"" then
		pcArray_TmpGlobalReturn = split(Session("pcGlobalArray"), chr(124))
		For xPackageCount = LBound(pcArray_TmpGlobalReturn) TO UBound(pcArray_TmpGlobalReturn)
			pcLocalArray(xPackageCount) = pcArray_TmpGlobalReturn(xPackageCount)	 
		Next
	end if
end if

'// UPDATE ARRAY
pcv_PrdList=""
If pcv_count <> 0 Then	
	For i=1 to pcv_count
		if request("C" & i)="1" then
			pcv_PrdList=pcv_PrdList & request("IDPrd" & i) & ","
		end if		
	Next
	pcLocalArray((pcPackageCount-1)) = pcv_PrdList
End If
'// CONVERT ARRAY TO SESSIONS
For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
	Session("pcAdminPrdList"&(xArrayCount+1)) = pcLocalArray(xArrayCount)  
Next

'// ARRAY TO PASS TO OTHER PAGES
pcv_strItemsList = join(pcLocalArray, chr(124))

'// SESSION FOR REDIRECTS
Session("pcGlobalArray") = pcv_strItemsList
'----------------------------------------------------------------------------------------------------
	%>
	<table class="pcCPcontent">
	<tr>
		<td valign="top">
		<table  border="0" cellpadding="0" cellspacing="0" width="60%">
		<tr>
			<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_IdOrder))%></b></td>
		</tr>
		<tr>
			<td><b>Steps</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
			<td width="95%"><font color="#A8A8A8">Select products</font></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step2a.gif"></td>
			<td><b>Specify Shipment Details</b></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step3.gif"></td>
			<td><font color="#A8A8A8">Finalize Shipment</font></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
	
<%
	' Look up shipping method
	
	Dim pshipmentDetails, pSRF, pShippingMethod
	query="SELECT shipmentDetails, SRF FROM orders WHERE idOrder=" & pcv_IdOrder & ";"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	pshipmentDetails=rs("shipmentDetails")
	pSRF=rs("SRF")
	set rs=nothing
		
	If pSRF="1" then
		pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_b")
	else
		shipping=split(pshipmentDetails,",")
		if ubound(shipping)>1 then
			if NOT isNumeric(trim(shipping(2))) then
				varShip="0"
				pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
			else
				Shipper=shipping(0)
				Service=shipping(1)
				Postage=trim(shipping(2))
				ShipID=shipping(5)
				if ubound(shipping)=>3 then
					serviceHandlingFee=trim(shipping(3))
					if NOT isNumeric(serviceHandlingFee) then
						serviceHandlingFee=0
					end if
				else
					serviceHandlingFee=0
				end if
			end if
		else
			varShip="0"
			pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
		end if 
	end if
			
	if pSRF="1" then
		pShippingMethod=pshipmentDetails
	else
		if varShip<>"0"  then
			pShippingMethod=Service
		else
			pShippingMethod=pshipmentDetails
		end if 
	end if
		
	' Look up today's date
	Dim varMonth, varDay, varYear
	varMonth=Month(Date)
	varDay=Day(Date)
	varYear=Year(Date) 
	dim dtInputStr
	dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
	if scDateFrmt="DD/MM/YY" then
		dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
	end if
			
			
	' Setup default Order Shipped message
		
	' Get customer information 
	query="SELECT idcustomer,orderDate FROM orders WHERE idOrder="& pcv_IdOrder
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	pIdCustomer=rs("idcustomer")
	pcv_orderDate=rs("orderDate")
	set rs=nothing
	
	query="SELECT name,lastname FROM customers WHERE idcustomer="& pIdCustomer
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	pcv_CustomerName = rs("name")&" "&rs("lastname")
	
	' Prepare message
	customerShippedEmail=""
	personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
	personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
	personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
	personalmessage=replace(personalmessage,"<TODAY_DATE>",dtInputStr)
	personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",pcv_CustomerName)
	personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(pcv_IdOrder)))
	personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pcv_orderDate))
	If scShippedEmail<>"" Then
		customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf & vbCrLf
	end if
	CustomerShippedEmail=replace(CustomerShippedEmail,"//","/")
	CustomerShippedEmail=replace(CustomerShippedEmail,"http:/","http://")
	CustomerShippedEmail=replace(CustomerShippedEmail,"https:/","https://")
	CustomerShippedEmail=replace(CustomerShippedEmail,"''",chr(39))
	
	'Work out transit days and email text type
	Select Case ShipID
		Case "C30"
			DAEditEmailTxtType = "UK"
			ShipDays = "0"
		Case "C22"
		 	DAEditEmailTxtType = "UK-SAT"
			ShipDays = "0"
		Case "C4" 'Northern Ireland
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C7" 'Republic of Ireland
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C8" 'Austria
		 	DAEditEmailTxtType = "INT"
			ShipDays = "2"
		Case "C9" 'Belgium
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C10" 'Denmark
		 	DAEditEmailTxtType = "INT"
			ShipDays = "2"
		Case "C11" 'France
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C12" 'Germany
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C13" 'Italy
		 	DAEditEmailTxtType = "INT"
			ShipDays = "4"
		Case "C14" 'Luxembourg
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C15" 'Monaco
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C16" 'Netherlands
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case "C17" 'Norway
		 	DAEditEmailTxtType = "INT"
			ShipDays = "4"
		Case "C18" 'Portugal
		 	DAEditEmailTxtType = "INT"
			ShipDays = "4"
		Case "C19" 'Spain
		 	DAEditEmailTxtType = "INT"
			ShipDays = "3"
		Case "C20" 'Sweden
		 	DAEditEmailTxtType = "INT"
			ShipDays = "4"
		Case "C21" 'Switzerland
		 	DAEditEmailTxtType = "INT"
			ShipDays = "2"
		Case "C23" 'Romania
		 	DAEditEmailTxtType = "INT"
			ShipDays = "4"
		Case "C24" 'Malta
		 	DAEditEmailTxtType = "INT"
			ShipDays = "4"
		Case "C27" 'Poland
		 	DAEditEmailTxtType = "INT"
			ShipDays = "3"
		Case "C28" 'Greece
		 	DAEditEmailTxtType = "INT"
			ShipDays = "6"
		Case "C29" 'Finland
		 	DAEditEmailTxtType = "INT"
			ShipDays = "4"
		Case "C31" 'Channel Islands
		 	DAEditEmailTxtType = "INT"
			ShipDays = "1"
		Case Else
			DAEditEmailTxtType = "UK"
			ShipDays = "0"
	End Select

	DAEditEmailIntTxt = "Dear <CUSTOMER_NAME>,<br/><br/>We are pleased to let you know that your order number <ORDER_ID> has been completed and is now prepared for dispatch.<br/><br/>Your delivery is scheduled to take <TRANSIT_DAYS> working days however we recommend using the courier tracking link below to monitor the shipment to get a more accurate delivery estimate. Please note that tracking details only become live once the couriers have collected your items from our workshop."
	DAEditEmailIntTxt=replace(DAEditEmailIntTxt,"<CUSTOMER_NAME>",pcv_CustomerName)
	DAEditEmailIntTxt=replace(DAEditEmailIntTxt,"<ORDER_ID>",(scpre + int(pcv_IdOrder)))
	DAEditEmailIntTxt=replace(DAEditEmailIntTxt,"<TRANSIT_DAYS>",ShipDays + 1)
	
	DAEditEmailUKTxt = "Dear <CUSTOMER_NAME>,<br/><br/>We are pleased to let you know that your order number <ORDER_ID> has been completed and is now prepared for dispatch.<br/><br/>Delivery will be made on the next working day (Monday - Friday, excluding bank holidays) after this message.<br/><br/>On the morning of delivery the couriers will email / txt you a 1 hour delivery window during which your delivery will take place.<br/><br/>You can also use the tracking link below to view and track your delivery online. Please note that tracking details only become live once the couriers have collected your items from our workshop."
	DAEditEmailUKTxt=replace(DAEditEmailUKTxt,"<CUSTOMER_NAME>",pcv_CustomerName)
	DAEditEmailUKTxt=replace(DAEditEmailUKTxt,"<ORDER_ID>",(scpre + int(pcv_IdOrder)))
	
	DAEditEmailUKSATTxt = "Dear <CUSTOMER_NAME>,<br/><br/>We are pleased to let you know that your order number <ORDER_ID> has been completed and is now prepared for dispatch.<br/><br/>Delivery will be made on the next Saturday after this message.<br/><br/>On the morning of delivery the couriers will email / txt you a 1 hour delivery window during which your delivery will take place.<br/><br/>You can also use the tracking link below to view and track your delivery online. Please note that tracking details only become live once the couriers have collected your items from our workshop."
	DAEditEmailUKSATTxt=replace(DAEditEmailUKSATTxt,"<CUSTOMER_NAME>",pcv_CustomerName)
	DAEditEmailUKSATTxt=replace(DAEditEmailUKSATTxt,"<ORDER_ID>",(scpre + int(pcv_IdOrder)))
	
	Select Case DAEditEmailTxtType
		Case "UK"
			DAEditEmailTxt = DAEditEmailUKTxt
		Case "UK-SAT"
			DAEditEmailTxt = DAEditEmailUKSATTxt
		Case "INT"
			DAEditEmailTxt = DAEditEmailIntTxt
	End Select
%>
	
	<Form name="form1" method="post" action="sds_ShipOrderWizard3.asp?action=add" class="pcForms">
		<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Specify Shipment Details</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
        <tr>
            <td colspan="2">
            <script type=text/javascript>
            function FaxSelected<%=k%>(){
            
            var selectValDom = document.forms['form1'];
            if (selectValDom.FaxLetter<%=k%>.checked == true) {
            document.getElementById('FaxTable<%=k%>').style.display='';
            }else{
            document.getElementById('FaxTable<%=k%>').style.display='none';
            }
            }
			</script>
			<%
			if Session("pcAdminFaxLetter"&k)="true" then
				pcv_strDisplayStyle="style=""display:visible"""
			else
				pcv_strDisplayStyle="style=""display:none"""
			end if
			%>
			<input onClick="FaxSelected<%=k%>();" name="FaxLetter<%=k%>" id="FaxLetter<%=k%>" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("FaxLetter"&k, "true")%>>
			Click Here to view <b>package contents</b>.
			
            <table class="pcCPcontent" ID="FaxTable<%=k%>" <%=pcv_strDisplayStyle%>>
                <tr>
                    <td width="14%">&nbsp;</td>
                    <td width="86%">
	  				<% 	for k=1 to pcPackageCount 

						xProductDisplayArray = split(Session("pcAdminPrdList"&k),",")
						For pcv_xCounter=0 to (ubound(xProductDisplayArray)-1)
							pcv_intPackageInfo_ID = xProductDisplayArray(pcv_xCounter)
							' GET THE PACKAGE CONTENTS
							' >>> Tables: products, ProductsOrdered
							query = "SELECT ProductsOrdered.pcPackageInfo_ID, ProductsOrdered.quantity , products.description, products.idProduct  "
							query = query & "FROM ProductsOrdered "
							query = query & "INNER JOIN products "
							query = query & "ON ProductsOrdered.idProduct = products.idProduct "
							query = query & "WHERE ProductsOrdered.idProductOrdered=" & pcv_intPackageInfo_ID &" "  
												
							set rs2=server.CreateObject("ADODB.RecordSet")
							set rs2=conntemp.execute(query)		
							
							if err.number<>0 then
								'// handle admin error
							end if
							
							if NOT rs2.eof then
								Do until rs2.eof	
									pcv_strProductQty = rs2("quantity")
									pcv_strProductDescription = rs2("description")
									
									%>
									<li><%=pcv_strProductQty&"&nbsp;"&pcv_strProductDescription%></li>
									<%
								rs2.movenext
								Loop								
							end if	
						Next
					next				
					%>
                    </td>
                </tr>
            </table>
            </td>
        </tr>
          
          
		<tr>
			<td width="15%">Shipment Method:</td>
			<td width="85%"><input type="text" name="pcv_method" value="<%=pShippingMethod%>" size="40"></td>
		</tr>
		<tr>
			<td>Tracking Number:</td>
			<td><input type="text" name="pcv_tracking" value="" size="40"></td>
		</tr>
		<tr>
			<td>Shipped Date:</td>
			<td><input type="text" name="pcv_shippedDate" value="<%=dtInputStr%>" size="40"> <span class="pcCPnotes">Date Format: <%=scDateFrmt%></span></td>
		</tr>
		<tr>
			<td valign="top">Comments:</td>
			<td valign="top">
			<textarea name="pcv_AdmComments" size="40" rows="10" cols="65"><%=DAEditEmailTxt%></textarea>
			<div style="margin: 10px 15px 15px 0;" class="pcCPnotes">Please note that additional text will appear in the message that is emailed to the customer depending on whether this is a partial or final shipment, and depending on which shipping provider was used for the shipment, if any. The additional text can be edited by editing the file &quot;includes/languages_ship.asp". We recommend that you ship a few test orders in different scenarios to become familiar with the way the final message appears.</div>			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>
			<input type="submit" name="submit1" value="Finalize Shipment" class="btn btn-primary">
			&nbsp;<input type="button" class="btn btn-default" name="Back" value="Back" onclick="javascript:history.back();">
			<input type=hidden name="PrdList" value="<%=pcv_PrdList%>">
			<input type=hidden name="idorder" value="<%=pcv_IdOrder%>">	
            
            <input type="hidden" name="PackageCount" value="<%=pcPackageCount%>">
            <input type="hidden" name="ItemsList" value="<%=pcv_strItemsList%>">
           	</td>
		</tr>
		</table>
	</Form>
<!--#include file="AdminFooter.asp"-->
