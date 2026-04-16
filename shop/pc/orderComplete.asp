<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/CashbackConstants.asp"-->
<!--#include file="header_wrapper.asp"-->
<%pcStrPageName="OrderComplete.asp"%>
<!--#include file="inc_AmazonHeader.asp"-->
<% 
err.number=0
dim pIdOrder, pOID, pnValid, pOrderStatus, pcv_noDoubleTracking
%>
<!--#include file="prv_getsettings.asp"-->
<script type=text/javascript>
	function openbrowser(url) {
			self.name = "productPageWin";
			popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
			if (navigator.appName == 'Netscape') {
			popUpWin.focus();
		}
	}
</script>
<%pcv_RWActive=pcv_Active
pnValid=0
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

' Create "View Previous Order" link
if scSSL="1" AND scIntSSLPage="1" then
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/CustviewPastD.asp"),"//","/")
else
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/CustviewPastD.asp"),"//","/")
end if
tempURL=replace(tempURL,"https:/","https://")
tempURL=replace(tempURL,"http:/","http://")
tempURL=tempURL & "?idOrder=" & (int(pOID)+scpre)
	
' clear cart data
if len(session("pcSFIdDbSession"))>0 then
	on error resume next
	query="DELETE FROM pcCustomerSessions WHERE idDbSession="&session("pcSFIdDbSession")&" AND randomKey="&session("pcSFRandomKey")&" AND idCustomer="&session("idCustomer")&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	err.number=0
	err.clear
end if

dim pcCartArray2(100,45)
Session("pcCartSession")=pcCartArray2
Session("pcCartIndex")=Cint(0)
session("pcSFIdDbSession")=""
session("pcSFRandomKey")=""
session("iOrderTotal")=""
session("pcSFCartRewards")=Cint(0)
session("pcSFUseRewards")=Cint(0)
session("IDRefer")=""
session("specialdiscount")=""
Session("ContinueRef")=""
session("TF1")=""
session("DF1")=""
session("shippingFullName")=""
session("shippingCompany")=""
session("shippingAddress")=""
session("shippingAddress2")=""
session("shippingStateCode")=""
session("shippingState")=""
session("shippingZip")=""
session("shippingPhone")=""
session("shippingCity")=""
session("shippingCountryCode")=""
session("DCODE")=""
session("idOrderSaved")=""
session("ExpressCheckoutPayment")=""
session("GWOrderDone")=""
session("redirectPage")=""
Session("SFStrRedirectUrl")=""
session("idGWSubmit")=""
session("idGWSubmit2")=""
session("idGWSubmit3")=""
session("Gateway")=""
session("SaveOrder")=""
Session("pcPromoSession")=""
Session("pcPromoIndex")=""
session("Entered-" & session("GWPaymentId"))=""
session("SF_DiscountTotal")=""
session("SF_RewardPointTotal")=""
session("pcSFIdPayment")=""
session("PPSAID") = ""
IDSC=0
tmpGUID=getUserInput(Request.Cookies("SavedCartGUID"),0)
IF tmpGUID<>"" THEN
	query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		IDSC=rsQ("SavedCartID")
		HasSavedCart=1
	end if
	set rsQ=nothing
	if HasSavedCart=1 then
		query="DELETE FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		query="DELETE FROM pcSavedCarts WHERE SavedCartID=" & IDSC & ";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	end if
	Response.Cookies("SavedCartGUID")=""
END IF

call pcs_hookOrderCompleted(pOID)
%>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Thank You For Your Order">Your Order Number Is: <%=(int(pOID)+scpre)%></h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->
	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
                    <div id="pcMain" class="pcOrderComplete">

	<div class="pcMainContent">
		
		<% if pnValid=1 then 'Order number not valid %>
		<div> 
			<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_viewPostings_a")%></div>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
		</div>
		<% else 'Order number is valid
		
			' Get order status and customer ID
			query = "SELECT orders.idCustomer, orders.orderStatus, orders.pcOrd_OrderKey, orders.total FROM orders WHERE orders.idOrder =" & pOID
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
					%>
					<div> 
						<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_viewPostings_a")%></div>
						<p>&nbsp;</p>
						<p>&nbsp;</p>
						<p>&nbsp;</p>
						<p>&nbsp;</p>
						<p>&nbsp;</p>
					</div>
			<% end if 
			
			'Get the customer ID if the session is empty
			if int(Session("idcustomer")) = 0 then
				Session("idcustomer") = rs("idCustomer")
			end if
				
			'Get the order status
			pOrderStatus=rs("orderStatus")
			pcOrderKey=rs("pcOrd_OrderKey")
            ptotal=rs("total")
			set rs=nothing
			
			'If order has already been processed, show corresponding message
			if pOrderStatus="3" then %>
			<h2>Thank you for your order.</h2>
			<% else %>
			<h2>Thank you for your order.</h2>
		<% 
			end if 'End if order has already been processed
		%>
		<h4><%=dictLanguage.Item(Session("language")&"_orderComplete_2")%></h4>

				<div class="pcTable hidden-xs">
					<div class="pcTableRow">
						<div class="pcTableRowFull">
						    <% 
							pcv_intNewAcct=getUserInput(Request("newAcct"),0)
							if pcv_intNewAcct="1" then 'New Account Created %>
							<div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_common_7")%></div>
						    <% end if %>
							<%if pcOrderKey<>"" then%>
								<div id="OrderCodeArea" class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_common_1")%>&nbsp;<%=pcOrderKey%></div>
								<p><%=dictLanguage.Item(Session("language")&"_opc_common_9")%></p>
							<%end if%>
						</div>
					</div>

					<% 
					' Start Order Details section
					pIdOrder=pOID
					
					query="SELECT customers.email,customers.fax,orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr,orders.orderDate, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.customerType, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.comments, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_shippingPhone, orders.shippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.idOrder, orders.rmaCredit, orders.ordPackageNum, orders.ord_DeliveryDate, orders.ord_OrderName, orders.ord_VAT,orders.pcOrd_CatDiscounts, orders.paymentDetails, orders.gwAuthCode, orders.gwTransId, orders.paymentCode, orders.pcOrd_GWTotal FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
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
					pidCustomer=Session("idcustomer")
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
					%>
					<div class="pcTableRow">
						<div class="pcTableRowFull">
							<div style="width: 50%;float: left;" class="daOCTblHead">
								<%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_14")%>
								<%response.write porderDate%> - 
								<%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_9")&": "&(int(pIdOrder)+scpre)%>
							</div>
							<div class="pcSmallText daOCTblHead" style="float: left;width: 49%;text-align: right;">
								<a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><img src="<%=pcf_getImagePath("images","document.gif")%>" alt="<%= dictLanguage.Item(Session("language")&"_CustviewOrd_33") %>"></a> 
								<a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><%= dictLanguage.Item(Session("language")&"_CustviewOrd_33") %></a>
							</div>
						</div>
					</div>
					
					<% if (pord_DeliveryDate<>"") then
						if scDateFrmt="DD/MM/YY" then
							pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
						else
							pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
						end if
						pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
						%>
						<div class="pcTableRow">
							<div>
							
							<%=dictLanguage.Item(Session("language")&"_CustviewOrd_39")%><%=pord_DeliveryDate%> <% If pord_DeliveryTime <> "00:00" Then %><%=", " & pord_DeliveryTime%><% End If %>
								
							</div>
						</div>
						<div class="pcTableRow">
							<div>&nbsp;</div>
						</div>
					<%end if%>
					
					<div class="pcTableHeader">
						<div class="pcOrderComplete_AddressTitle">&nbsp;</div>
						<div class="pcOrderComplete_BillingAddress"><%response.write dictLanguage.Item(Session("language")&"_orderverify_23")%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								response.write dictLanguage.Item(Session("language")&"_orderverify_24")
							end if%>
						</div>
					</div>
					
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle">
							<b><% response.write replace(dictLanguage.Item(Session("language")&"_orderverify_7"),"''","'")%></b>
						</div>
						<div class="pcOrderComplete_BillingAddress"> 
							<% response.write pFirstName&" "&plastname %>
						</div>
						<div class="pcOrderComplete_ShippingAddress">
							<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then%>
								<% response.write pshippingFullName %>
							<% end if%>
						</div>
					</div>
					
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle">
							<b><% response.write dictLanguage.Item(Session("language")&"_orderverify_8")%></b>
						</div>
						<div class="pcOrderComplete_BillingAddress"><%=pcustomerCompany%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								if pshippingCompany<>"" then
									response.write pshippingCompany
								else
									if (pshippingFullName = "" or pshippingFullName = pCustomerName) and pCustomerCompany <> "" then
										response.write pCustomerCompany
									end if
								end if
							end if %>
						</div>
					</div>
					
					<%if pEmail<>pshippingEmail AND pshippingEmail<>"" then%>
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle"><b> 
							<%=dictLanguage.Item(Session("language")&"_opc_5")%>
						</b></div>
						<div class="pcOrderComplete_BillingAddress"><%=pEmail%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								response.write pshippingEmail
							end if %>
						</div>
					</div>
					<%end if%>
					
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle"><b> 
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_9")%>
						</b></div>
						<div class="pcOrderComplete_BillingAddress"><%=pPhone%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								response.write pshippingPhone
							end if %>
						</div>
					</div>
					
					<%if pFax<>"" OR pshippingFax<>"" then%>
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle"><b> 
							<%=dictLanguage.Item(Session("language")&"_opc_18")%>
						</b></div>
						<div class="pcOrderComplete_BillingAddress"><%=pFax%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								response.write pshippingFax
							end if %>
						</div>
					</div>
					<%end if%>
					
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle"><b> 
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_10")%>
						</b></div>
						<div class="pcOrderComplete_BillingAddress"><%=paddress%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								if pshippingAddress="" then
									response.write "Same as Billing Address"
								else
									response.write pshippingAddress
								end if
							else
								if pcShowShipAddr="0" AND session("gHideAddress")<>"1" then
									response.write "Same as Billing Address"
								end if
							end if %>
						 </div>
					</div>
					
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle">&nbsp;</div>
						<div class="pcOrderComplete_BillingAddress"><%=paddress2%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								if pshippingAddress2<>"" then
									response.write pshippingAddress2
								end if
							end if %>
						</div>
					</div>
					
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle">&nbsp;</div>
						<div class="pcOrderComplete_BillingAddress"><%=pCity&", "&pState&" "&pzip%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								if pshippingAddress<>"" then
									response.write pShippingCity&", "&pshippingState
									If pshippingState="" then
										response.write pshippingStateCode
									End If
									response.write " "&pshippingZip
								end if
							end if %>
						</div>
					</div>
					
					<div class="pcTableRow">
						<div class="pcOrderComplete_AddressTitle">&nbsp;</div>
						<div class="pcOrderComplete_BillingAddress"><%=pCountryCode%></div>
						<div class="pcOrderComplete_ShippingAddress">
							<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
								response.write pshippingCountryCode
								strFedExCountryCode=pshippingCountryCode
							else
								strFedExCountryCode=pCountryCode
							end if %>
						</div>
					</div>
				
					<% ' Start of payment details
					payment = split(pcpaymentDetails,"||")
					PaymentType=trim(payment(0))
					
					'Get payment nickname
					query="SELECT paymentDesc, paymentNickName FROM paytypes WHERE paymentDesc = '" & replace(PaymentType,"'","''") & "';"
					Set rsTemp=Server.CreateObject("ADODB.Recordset")
					Set rsTemp=connTemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsTemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					if not rsTemp.EOF then
						PaymentName=trim(rsTemp("paymentNickName"))
						else
						PaymentName=""
					end if
					Set rsTemp = nothing
					'End get payment nickname
					
					'Get authorization and transaction IDs, if any
					varTransID=""
					varTransName= dictLanguage.Item(Session("language")&"_CustviewPastD_102")
					varAuthCode=""
					varAuthName= dictLanguage.Item(Session("language")&"_CustviewPastD_103")
				
					if NOT isNull(pcpaymentCode) AND pcpaymentCode<>"" then 
						varShowCCInfo=0
						select case pcpaymentCode
						case "LinkPoint"
							varAry=split(pcgwAuthCode,":")
							varTransName="Approval Number"
							varAuthName="Reference Number"
							varTransID=left(varAry(1),6)
							varAuthCode=right(varAry(1),10)
						case "PFLink", "PFPro", "PFPRO", "PFLINK"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
							varShowCCInfo=1
							varGWInfo="P"
						case "Authorize"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
							varShowCCInfo=1
							if instr(ucase(PaymentType),"CHECK") then
								varShowCCInfo=0
							end if
							varGWInfo="A"
						case "twoCheckout"
							varTransName="2Checkout Order No"
							varTransID=pcgwTransId
						case "BOFA"
							varTransName="Order No"
							varAuthName="Authorization Code"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
						case "WorldPay"
							varTransID=""
							varAuthCode=""
						case "iTransact"
							varTransName="Transaction ID"
							varAuthName="Authorization Code"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
						case "PSI", "PSIGate"
							varTransName="Transaction ID"
							varAuthName="Authorization Code"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
						case "fasttransact", "FastTransact", "FAST","CyberSource"
							varTransName="Transaction ID"
							varAuthName="Authorization Code"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
						case "USAePay","FastCharge"
							varTransName="Transaction reference code"
							varAuthName="Authorization code"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
						case "PxPay"
							varTransName="DPS Transaction Reference Number"
							varAuthName="Authorization code"
							varTransID=pcgwTransId
							varAuthCode=pcgwAuthCode
						 case "Moneris2"								     					  
							  Dim varIDEBIT_ISSCONF, varIDEBIT_ISSNAME,varRespName,varResponseCode
							   varTransName="Sequence Number"
							   varAuthName="Approval Code"
							   varRespName="Response / ISO Code"
							   varTransID=pcgwTransId
							   varAuthCode=pcgwAuthCode
							
							   query = "Select pcPay_MOrder_responseCode, pcPay_MOrder_ISOcode, pcPay_MOrder_IDEBIT_ISSCONF, pcPay_MOrder_IDEBIT_ISSNAME from pcPay_OrdersMoneris Where pcPay_MOrder_TransId='"& pcgwTransId &"';" 
							   set rstemp=server.CreateObject("ADODB.RecordSet")
							   set rstemp=conntemp.execute(query)											  
								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
				
							if not rs.eof then
							   varResponseCode = RStemp("pcPay_MOrder_responseCode")
							   varISO_Code = RStemp("pcPay_MOrder_ISOcode")
							   varIDEBIT_ISSCONF = rstemp("pcPay_MOrder_IDEBIT_ISSCONF")
							   varIDEBIT_ISSNAME = rstemp("pcPay_MOrder_IDEBIT_ISSNAME")								 							
							end if
							set rstemp=nothing
							  
						end select
				
					end if
					
					'End get authorization and transaction IDs
					
					on error resume next
					If payment(1)="" then
					 if err.number<>0 then
						PayCharge=0
					 end if
						PayCharge=0
					else
						PayCharge=payment(1)
					end If
					err.number=0
					if instr(PaymentType,"FREE") AND len(PaymentType)<6 then
					else %>
						<div class="pcTableRow">
							<div class="pcTableRowFull"><hr></div>
						</div>
						<div class="pcTableRow">
							<div class="pcTableRowFull">
							<%=dictLanguage.Item(Session("language")&"_CustviewPastD_101")%>
							<%
								if PaymentName <> "" and PaymentName <> PaymentType then
									Dim pcv_strPaymentType
									Select Case PaymentType
										Case "PayPal Website Payments Pro": pcv_strPaymentType=PaymentName
										Case Else: pcv_strPaymentType=PaymentName & " (" & PaymentType & ")"
									End Select
									Response.Write pcv_strPaymentType
									else
									Response.Write PaymentType
								end if
							%>
							<% if PayCharge>0 then %>
								<br><%=dictLanguage.Item(Session("language")&"_CustviewOrd_14b")%><%= " " & scCurSign&money(PayCharge)%>
							<% end if %>
							<% if varTransID<>"" then %>
							<br><%=varTransName%>: <%=varTransID%>
							<% end if %>
							<% if varAuthCode<>"" then %>
							<br><%=varAuthName%>: <%=varAuthCode%>
							<% end if %>
							<%if varResponseCode <> ""  or varISO_Code <> "" Then%>
							<BR><%=varRespName%>&nbsp;<%=varResponseCode%>/<%=varISO_Code%>
							<%end if %>
							<% if varIDEBIT_ISSCONF <> ""  and varIDEBIT_ISSNAME <> "" then %>
							<br><%=dictLanguage.Item(Session("language")&"_CustviewOrd_48")%>
							<BR><%=dictLanguage.Item(Session("language")&"_CustviewOrd_49")%>&nbsp;<%=varIDEBIT_ISSNAME%>
							<BR><%=dictLanguage.Item(Session("language")&"_CustviewOrd_50")%>&nbsp;<%=varIDEBIT_ISSCONF%>						
							<% end if%>
							
							<br><br>
							</div>
						</div>
					<% end if
						' End of payment details
					%>
				
					<% ' Start of order comments
						if len(pcomments)>3 then %>
						<div class="pcTableRow">
							<div class="pcTableRowFull"> 
								<b><% response.write dictLanguage.Item(Session("language")&"_orderverify_11")%></b> 
								<%=pcomments%><br>
								<br>
							</div>
						</div>
					<% end if 
						' End of order comments
					%>
		</div>
        <!--#include file="inc_ordercomplete.asp"-->
		<div class="pcTable">
		<%' ------------------------------------------------------
		'Start SDBA - Notify Drop-Shipping
		' ------------------------------------------------------
		if scShipNotifySeparate="1" then
			tmp_showmsg=0
			query="SELECT products.pcProd_IsDropShipped FROM products INNER JOIN productsOrdered ON (products.idproduct=productsOrdered.idproduct AND products.pcProd_IsDropShipped=1) WHERE ProductsOrdered.idOrder=" & pIdOrder & ";"
			set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if not rs.eof then
				tmp_showmsg=1
			end if
			set rs=nothing
			if tmp_showmsg=1 then%>
			<div class="pcTableRow"> 
				<div class="pcTableRowFull">
					<hr>
				</div>
			</div>
			<div class="pcTableRow">
				<div class="pcTableRowFull">
					<div class="pcTextMessage"><%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%></div>
				</div>
			</div>
			<%end if
		end if
		' ------------------------------------------------------
		'End SDBA - Notify Drop-Shipping
		' ------------------------------------------------------%>
		
	<%
		end if 'End if order number is valid
	%>
		<%if Session("CustomerGuest")="1" then%>
		<div class="pcTableRow">
			<div class="pcTableRowFull">
				<div id="PwdArea">
					<form id="PwdForm" name="PwdForm">
						<div class="pcShowContent">
							<div class="pcFormItem">
								<div class="pcSectionTitle"><%=dictLanguage.Item(Session("language")&"_opc_common_2")%></div>
							</div>
							<div class="pcFormItem">
								<div><%=dictLanguage.Item(Session("language")&"_opc_common_3")%></div>
							</div>
							<div class="pcFormItem">
								<div style="width:20%;"><%=dictLanguage.Item(Session("language")&"_opc_6")%></div>
								<div style="width:30%;"><input type="password" name="newPass1" id="newPass1" size="20"></div>
								<div style="width:20%;"><%=dictLanguage.Item(Session("language")&"_opc_38")%></div>
								<div style="width:30%;"><input type="password" name="newPass2" id="newPass2" size="20"></div>
							</div>
							<div class="pcFormItem">
								<div style="padding-top: 10px;"></div>
							</div>
							<div class="pcFormItem">
								<div style="padding-top: 10px;"><input type="button" name="PwdSubmit" id="PwdSubmit" value="<%=dictLanguage.Item(Session("language")&"_opc_common_4")%>" class="submit2"></div>
							</div>
						</div>
					</form>
					<div id="PwdLoader" style="display:none"></div>
				</div>
			</div>
		</div>
		<%end if%>
		<%if Session("CustomerGuest")="2" then
			Session("JustPurchased")="1"
		end if%>
		<div class="pcTableRow">
			<div class="pcTableRowFull">
				<script type=text/javascript>
				$pc(document).ready(function()
				{
					//jQuery.validator.setDefaults({
						//success: function(element) {
							//$pc(element).parent("td").children("input, textarea").addClass("success")
						//}
					//});

					<%if Session("CustomerGuest")="1" then%>
					//*Validate Password Form
					$pc("#PwdForm").validate({
						rules: {
							newPass1: 
							{
								required: true,
							},
							newPass2:
							{
								required: true,
								equalTo: "#newPass1"
							}
						},
						messages: {
							newPass1: {
								required: "<%=dictLanguage.Item(Session("language")&"_opc_js_4")%>",
								minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>"
							},
							newPass2: {
								required: "<%=dictLanguage.Item(Session("language")&"_opc_js_47")%>",
								minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>",
								equalTo: "<%=dictLanguage.Item(Session("language")&"_opc_js_48")%>"
							}
						}
					})
					
					$pc('#PwdSubmit').click(function(){
						if ($pc('#PwdForm').validate().form())
						{
							$pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","ajax-loader1.gif")%>" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_5")%>');
							$pc("#PwdLoader").show();	
							$pc.ajax({
								type: "POST",
								url: "opc_createacc.asp",
								data: $pc('#PwdForm').formSerialize() + "&action=create",
								timeout: 5000,
								success: function(data, textStatus){
									if (data=="SECURITY")
									{
										$pc("#PwdArea").html("");
										$pc("#PwdArea").hide();
										$pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_error_small.png")%>" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_6")%>');
										var callbackPwd=function (){setTimeout(function(){$pc("#PwdLoader").hide();},1000);}
										$pc("#PwdLoader").effect('pulsate',{},500,callbackPwd);
									}
									else
									{
									if ((data=="OK") || (data=="REG") || (data=="OKA") || (data=="REGA"))
									{
										location='orderComplete.asp?newAcct=1';
									}
									else
									{
										$pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_error_small.png")%>" align="absmiddle"> '+data);
										var callbackPwd=function (){setTimeout(function(){$pc("#PwdLoader").hide();},1000);}
										$pc("#PwdLoader").effect('pulsate',{},500,callbackPwd);
									}
									}
								}
					 		});
							return(false);
						}
						return(false);
					});
					<%end if%>

					<%if pcOrderKey<>"" then%>
					//var callbackOCA=function(){};
					//$pc("#OrderCodeArea").effect('pulsate',{},800,callbackOCA);
					<%end if%>
				});
				</script>
			</div>
		</div>
		<% ' Continue shopping button %>
		<% ' End Continue shopping button %>
	</div> 
</div> <!-- pcMainContent -->
</div> <!-- pcMain -->
<% 
'// Tell the system that this is the order completed page
Dim pcv_intOrderComplete
pcv_intOrderComplete=1

'// Tell the system that there has been a page refresh
if pcv_noDoubleTracking=1 then
	pcv_intOrderComplete=0
end if

%>
<!--#include file="inc-GTSOrderConfirm.asp"-->
<!--#include file="orderCompleteTracking.asp"-->
<!--#include file="inc-Cashback.asp"-->
<%
session("ExpressCheckoutPayment")=""
Session("PayPalExpressToken")=""
session("gHideAddress")=""
session("AmazonFirstTime")=""
session("Amz_scope")=""
session("Amz_expires_in")=""
session("Amz_token_type")=""
session("Amz_access_token")=""
session("AmzOrderID")=""
session("AmzBillAgreementID")=""
session("PPSAID") = ""

on error resume next
If Session("Payer")&""<>"" Then

	Session("Payer") = ""
    session("ExpressCheckoutPayment") = ""
	
	' clear cart data
	redim pcCartArray2(100,45)
	Session("pcCartSession")=pcCartArray2
	Session("pcCartIndex")=Cint(0)
    err.clear
    
End If

'Log successful transaction
call pcs_LogTransaction(Session("idcustomer"), pIdOrder, 1)

' Reset Failed Payment Count
call pcs_clearFailedPaymentAttempt(Session("idcustomer"))

'// Google Analytics (GA)
'// Inform GA script that this is the Order Completed page
'// If GA is not used, this code does not need to be removed as it is harmless
Dim pcGAorderComplete 
pcGAorderComplete=1
'APP-S
'Update parent products inventory levels if necessary%>
<!--#include file="app-updstock.asp"-->
<%' End update parent products inventory levels
'APP-E%>
<%IF session("PayWithAmazon")="YES" THEN%>
<script type=text/javascript>
    amazon.Login.logout();
</script>
<%END IF
session("PayWithAmazon")=""%>
					</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->
<!--#include file="footer_wrapper.asp"-->