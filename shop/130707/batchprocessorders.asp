<%
Server.ScriptTimeout = 5400
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
Dim pageTitle, Section
pageTitle="Batch Process Orders"
Section="orders"
%>
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>
	function openwin(file)
	{
		msgWindow=open(file,'win1','scrollbars=yes,resizable=yes,width=500,height=400');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
</script>

<%
if request.QueryString("capture")<>"" then
	
	cIdOrder=request.QueryString("capture")
	cGateway=request.QueryString("GW")

	select case cGateway
		case "authorders"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idauthorder="&cIdOrder			
		case "pcPay_PayPal_Authorize"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idPayPal_Authorize="&cIdOrder	
		case "pcPay_PFL_Authorize"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idPFL_Authorize="&cIdOrder
		case "pcPay_LinkPointAPI"
			query="UPDATE "&cGateway&" SET pcPay_LPAPI_Captured=1 WHERE pcPay_LPAPI_ID="&cIdOrder	
		case "pfporders"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idpfporder="&cIdOrder
		case "netbillorders"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idnetbillorder="&cIdOrder
		case "pcPay_USAePay_Orders"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idePayOrder="&cIdOrder
		case "pcPay_EIG_Authorize"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idauthorder="&cIdOrder
		case "pcPay_EPN_Authorize"
			query="UPDATE "&cGateway&" SET captured=1 WHERE idauthorder="&cIdOrder
		case else				
			query="UPDATE "&cGateway&" SET captured=1 WHERE idOrder="&cIdOrder
	end select

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing

	call closeDb()
response.redirect "batchprocessorders.asp"
	response.End()
end if

PFPOrder=request("PFPOrder")
PFPSort=request("PFPSort")
if PFPOrder="" then
	PFPOrder="orders.idOrder"
	PFPSort="DESC"
end if

AuthOrder=request.QueryString("AuthOrder")
AuthSort=request.QueryString("AuthSort")
if AuthOrder="" then
	AuthOrder="orders.idOrder"
end if
if AuthSort="" then
	AuthSort="DESC"
end if

PFApiOrder=request.QueryString("PFApiOrder")
PFApiSort=request.QueryString("PFApiSort")
if PFApiOrder="" then
	PFApiOrder="orders.idOrder"
	PFApiSort="DESC"
end if


PayPalOrder=request.QueryString("PayPalOrder")
PayPalSort=request.QueryString("PayPalSort")
if PayPalOrder="" then
	PayPalOrder="orders.idOrder"
	PayPalSort="DESC"
end if

LinkOrder=request("LinkOrder")
LinkSort=request("LinkSort")
if LinkOrder="" then
	LinkOrder="idOrder"
	LinkSort="DESC"
end if

NetbillOrder=request.QueryString("NetbillOrder")
NetbillSort=request.QueryString("NetbillSort")
if NetbillOrder="" then
	NetbillOrder="orders.idOrder"
	NetbillSort="DESC"
end if

USAePayOrder=request.QueryString("NetbillOrder")
USAePaySort=request.QueryString("NetbillSort")
if USAePayOrder="" then
	USAePayOrder="orders.idOrder"
	USAePayOrderSort="DESC"
end if

GenOrder=request("GenOrder")
GenSort=request("GenSort")
if GenOrder="" then
	GenOrder="idOrder"
	GenSort="DESC"
end if

PEYGenOrder=request("PEYGenOrder")
PEYGenSort=request("PEYGenSort")
if PEYGenOrder="" then
	PEYGenOrder="idOrder"
	PEYGenSort="DESC"
end if

Dim iCnt, gwa, gwvpfp, gwpp, gwpsi, gwit, gwlp, gwvpfl, gwwp, gwmoneris, gwbofa, gw2Checkout, gwAIM, gwnetbill, gwUSAePay, gwEIG, varTemp, varActive, actGW


'// Get Payment Code
query="SELECT paymentDesc, gwCode, active FROM paytypes"
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	
	call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if NOT rs.eof then 
	
	iCnt=1
	gwa=0
	gwaDPM=0
	gwvpfp=0
	gwpp=0
	gwpsi=0
	gwit=0
	gwlp=0
	gwvpfl=0
	gwwp=0
	gwmoneris=0
	gwbofa=0
	gw2Checkout=0
	gwAIM=0
	gwnetbill=0
	gwUSAePay=0
	gwEIG=0
	gwEPN=0
	PayPalDP=0
	PayPalExp=0
	PayPalPPA=0
	PayPalPPP=0
	PayPalPFL=0
	actGW=0
	gwPEY=0
	
	do until rs.eof
		varTemp=rs("gwCode")
		varActive=rs("active")
		varPaymentDesc=rs("paymentDesc")
		
		if varTemp<>"0" then			
			select case varTemp
				case 1
					gwa=1
					actGW=1
				case 1103
					gwaDPM=1
					actGW=1
				case 2
					gwvpfp=1
					actGW=1
				case 3
					gwpp=1
				case 4
					gwpsi=1
				case 5
					gwit=1 
				case 8
					gwlp=1
				case 10
					gwwp=1
				case 11
					gwmoneris=1
				case 12
					gwbofa=1
				case 13
					gw2Checkout=1
				case 14
					gwAIM=1
				case 27
					gwnetbill=1
					actGW=1
				case 35
					gwUSAePay=1
					actGW=1
				case 42
					gwEPN=1
				case 67
					gwEIG=1
					actGW=1
				case 46
					PayPalDP=1
					PayPalDPName=varPaymentDesc
					actGW=1
				case 53
					PayPalPPP=1
					PayPalPPPName=varPaymentDesc
				case 80
					PayPalPPA=1
					PayPalPPAName=varPaymentDesc
				case 9,99
					PayPalPFL=1
					PayPalPFLName=varPaymentDesc
				case 999999
					PayPalExp=1
					actGW=1
				case 1101
					gwPEY=1
			end select
		end if '// if varTemp<>"0" then
		rs.moveNext
	loop
end if
set rs=nothing

'////////////////////////////////////////////////////
'// START: Payeezy
'////////////////////////////////////////////////////
IF gwPEY=1 THEN
query="SELECT orders.idOrder, orders.orderDate, orders.orderstatus, orders.total, orders.idCustomer, orders.paymentCode, orders.paymentDetails, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany FROM orders INNER JOIN customers ON orders.idCustomer = customers.idCustomer WHERE orderstatus=2 AND paymentCode='Payeezy' ORDER BY "&PEYGenOrder&" "&PEYGenSort&";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
%>
<form name="PEYform" method="post" action="batchprocess_Payeezy.asp" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="7"><h2>Payeezy Orders</h2></td>
	</tr>
	<tr>
		<th>Process</th>
		<th nowrap="nowrap">Send Email</th>
		<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PEYGenOrder=orders.orderdate&PEYGenSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PEYGenOrder=orders.orderdate&PEYGenSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
		<th nowrap="nowrap">Payment Type</th>
		<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PEYGenOrder=idOrder&PEYGenSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=pfporders.idOrder&PFPSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
		<th>Customer</th>
		<th>Total</th>
	</tr>
    <tr>
        <td colspan="8" class="pcCPspacer"></td>
    </tr>
	<% dim noPEYORDRec
	noPEYORDRec=0
	if rs.eof then
		noPEYORDRec=1 
		%>
		<tr> 
			<td colspan="7"><div class="pcCPmessage">No records found</div></td>
		</tr>
	<% end if %>
	<% PEYcheckboxCnt=0
	do until rs.eof
		idOrder=rs("idOrder")
		orderDate=rs("orderDate")
		orderStatus=rs("orderstatus")
		total=rs("total")
		idCustomer=rs("idCustomer")
		paymentCode=rs("paymentCode")
		ppaymentDetails=trim(rs("paymentDetails"))
		pcv_custcomments=trim(rs("comments"))
		pcv_admcomments=trim(rs("admincomments"))
		customerName=rs("name") & " " & rs("lastName")
		customerCompany=rs("customerCompany")
			if trim(customerCompany)<>"" then
				customerInfo=customerName & " (" & customerCompany & ")"
				else
				customerInfo=customerName
			end if
		PEYcheckboxCnt=PEYcheckboxCnt+1
		%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td>
				<div align="center"><input name="checkOrd<%=PEYcheckboxCnt%>" type="checkbox" id="checkOrd<%=PEYcheckboxCnt%>" value="YES" class="clearBorder"></div>
			</td>
			<td>
				<div align="center"><input name="checkEmail<%=PEYcheckboxCnt%>" type="checkbox" id="checkEmail<%=PEYcheckboxCnt%>" value="YES" checked class="clearBorder"></div>
			</td>
			<td><%=ShowDateFrmt(orderDate)%></td>
			<td>
				<input type="hidden" name="orderstatus<%=PEYcheckboxCnt%>" value="<%=orderStatus%>">
				<input type="hidden" name="idOrder<%=PEYcheckboxCnt%>" value="<%=idOrder%>">
				<input type="hidden" name="amt<%=PEYcheckboxCnt%>" value="<%=total%>">
				<%=PaymentType%>
			</td>
			<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
			<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
			<td><div align="center"><%=scCurSign&money(total)%></div></td>
		</tr>
		<%
		rs.moveNext
	loop
	set rs=nothing
	
	if noPEYORDRec=0 AND PEYcheckboxCnt=0 then
		noPEYORDRec=1 
		%>
		<tr> 
			<td colspan="7"><div class="pcCPmessage">No pending records found</div></td>
		</tr>
	<% end if %>
	<input type="hidden" name="PEYcheckboxCnt" value="<%=PEYcheckboxCnt%>">
	<tr>
				<td nowrap="nowrap">
				<%if PEYcheckboxCnt>"0" then%>
				<input type=hidden name="PEYCheck1" value="0">
				<input type="checkbox" name="PEYCheck1a" value="1" onClick="javascript:testcheckPEY_1()" class="clearBorder"> Check All 
				<script type=text/javascript>
				function checkAllPEY_1() {
				for (var j = 1; j <= <%=PEYcheckboxCnt%>; j++) {
				box = eval("document.PEYform.checkOrd" + j); 
				if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAllPEY_1() {
				for (var j = 1; j <= <%=PEYcheckboxCnt%>; j++) {
				box = eval("document.PEYform.checkOrd" + j); 
				if (box.checked == true) box.checked = false;
				   }
				}
				
				function testcheckPEY_1() {
				if (document.PEYform.PEYCheck1.value=="0") {
				document.PEYform.PEYCheck1.value="1";
				checkAllPEY_1();
					}
				else
					{
				document.PEYform.PEYCheck1.value="0";
				uncheckAllPEY_1();
					}
				}
				</script>
				<%end if%>
				</td>
				<td nowrap="nowrap">
				<%if PEYcheckboxCnt>"0" then%>
				<input type=hidden name="PEYCheck2" value="1">
				<input type="checkbox" name="PEYCheck2a" checked value="1" onClick="javascript:testcheckPEY_2()" class="clearBorder"> Check All
				<script type=text/javascript>
				function checkAllPEY_2() {
				for (var j = 1; j <= <%=PEYcheckboxCnt%>; j++) {
				box = eval("document.PEYform.checkEmail" + j); 
				if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAllPEY_2() {
				for (var j = 1; j <= <%=PEYcheckboxCnt%>; j++) {
				box = eval("document.PEYform.checkEmail" + j); 
				if (box.checked == true) box.checked = false;
				   }
				}
				
				function testcheckPEY_2() {
				if (document.PEYform.PEYCheck2.value=="0") {
				document.PEYform.PEYCheck2.value="1";
				checkAllPEY_2();
					}
				else
					{
				document.PEYform.PEYCheck2.value="0";
				uncheckAllPEY_2();
					}
				}
				</script>
				<%end if%>
				</td>
				<td colspan="5">&nbsp;</td>
		</tr>
	<% if noPEYORDRec=0  then %>
		<tr>
			<td colspan="7" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="7"><input type="submit" name="PendingSubmit" value="Process Selected Orders" class="btn btn-primary"></td>
		</tr>
		<tr>
			<td colspan="7">&nbsp;</td>
		 </tr>
	<% end if %>
</table>
</form>
<%END IF
'////////////////////////////////////////////////////
'// END: Payeezy
'////////////////////////////////////////////////////


'////////////////////////////////////////////////////
'// START: Authorize Net
'////////////////////////////////////////////////////
IF gwa=1 THEN
	
	query="SELECT orders.paymentCode FROM orders WHERE paymentCode='Authorize';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If
	set rs=nothing '// If NOT rs.EOF Then

	If pcv_strDisplaySection = -1 Then
		'// Check for authorize.net orders
		query="SELECT authorders.idOrder, authorders.idauthorder, authorders.amount, authorders.paymentmethod, authorders.transtype, authorders.authcode, authorders.ccnum, authorders.ccexp, authorders.idCustomer, authorders.fname, authorders.lname, authorders.address, authorders.zip, authorders.pcSecurityKeyID, orders.orderDate, orders.orderStatus, orders.gwTransId, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email FROM customers INNER JOIN (authorders INNER JOIN orders ON authorders.idOrder = orders.idOrder) ON (authorders.idCustomer = customers.idcustomer) AND (customers.idcustomer = orders.idCustomer) WHERE (((authorders.transtype)='AUTH_ONLY') AND ((authorders.captured)=0)) AND authorders.paymentmethod!='DPM' ORDER BY "&AuthOrder&" "&AuthSort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
		<form name="form1" method="post" action="batchprocess_auth.asp" class="pcForms">
			<table class="pcCPcontent">
				<tr>
					<td colspan="8"><h2>Authorize.Net Orders</h2></td>
				</tr>
				<tr>
					<th>Process</th>
					<th nowrap="nowrap">Send Email</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authorders.idOrder&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authorders.idOrder&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
					<th>Customer</th>
					<th colspan="2" align="left">Total</th>
				</tr>
                <tr>
                	<td colspan="8" class="pcCPspacer"></td>
                </tr>
				<% dim noAuthRec
				noAuthRec=0
				if rs.eof then
					noAuthRec=1 
					%>
					<tr> 
						<td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
					</tr>
				<% end if %>
				<% dim checkboxCnt
				checkboxCnt=0
				do until rs.eof
					checkboxCnt=checkboxCnt+1
					idOrder=rs("idOrder")
					idauthorder=rs("idauthorder")
					amount=rs("amount")
					paymentmethod=rs("paymentmethod")
					transtype=rs("transtype")
					authcode=rs("authcode")
					ccnum=rs("ccnum")
					ccexp=rs("ccexp")
					idCustomer=rs("idCustomer")
					fname=rs("fname")
					lname=rs("lname")
					address=rs("address")
					zip=rs("zip")
					pcv_SecurityKeyID=rs("pcSecurityKeyID")
					orderDate=rs("orderDate")
					orderStatus=rs("orderstatus")
					gwTransId=rs("gwTransId")
					stateCode=rs("stateCode")
					if stateCode="" then
						stateCode=rs("State")
					end if
					City=rs("city")
					countryCode=rs("countryCode")
					shippingAddress=rs("shippingAddress")
					shippingStateCode=rs("shippingStateCode")
					shippingState=rs("shippingState")
					shippingCity=rs("shippingCity")
					shippingCountryCode=rs("shippingCountryCode")
					shippingZip=rs("shippingZip")
					shippingFullName=rs("shippingFullName")
					address2=rs("address2")
					shippingCompany=rs("shippingCompany")
					shippingAddress2=rs("shippingAddress2")
					pcv_custcomments=trim(rs("comments"))
					pcv_admcomments=trim(rs("admincomments"))
					customerName=rs("name") & " " & rs("lastName")
					customerCompany=rs("customerCompany")
						if trim(customerCompany)<>"" then
							customerInfo=customerName & " (" & customerCompany & ")"
							else
							customerInfo=customerName
						end if
					phone=rs("phone")
					email =rs("email")
					
					pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
					ccnum2=enDeCrypt(ccnum, pcv_SecurityPass)
										
					'// Get amount from orders table
					query="SELECT total from orders WHERE idOrder="&idOrder&";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					curTotal=rstemp("total")
					set rstemp=nothing %>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td>
						<div align="center">
						<input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
						</div></td>
						<td>
						<div align="center">
						<input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
						</div></td>
						<td><%=ShowDateFrmt(orderDate)%></td>
						<td><%=authcode%>
						
						<input type="hidden" name="idauthorder<%=checkboxCnt%>" value="<%=idauthorder%>">
						<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=int(idOrder)+scpre%>">
						<input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
						</td>
						<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
						<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
						<td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
						<td><div align="center"><a href="batchprocessorders.asp?capture=<%=idauthorder%>&GW=authorders">Remove</a></div></td>
					</tr>
					<% rs.moveNext
				loop
				set rs=nothing
				%>
			<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
			<tr>
				<td nowrap="nowrap">
					<%if checkboxCnt>"0" then%>
					<input type=hidden name="Check1" value="0">
					<input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck1_1()" class="clearBorder"> Check All
					<script type=text/javascript>
					function checkAll1_1() {
					for (var j = 1; j <= <%=checkboxCnt%>; j++) {
					box = eval("document.form1.checkOrd" + j); 
					if (box.checked == false) box.checked = true;
							}
					}

					function uncheckAll1_1() {
					for (var j = 1; j <= <%=checkboxCnt%>; j++) {
					box = eval("document.form1.checkOrd" + j); 
					if (box.checked == true) box.checked = false;
							 }
					}
					
					function testcheck1_1() {
					if (document.form1.Check1.value=="0") {
					document.form1.Check1.value="1";
					checkAll1_1();
							}
					else
							{
					document.form1.Check1.value="0";
					uncheckAll1_1();
							}
					}
					</script>
					<%end if%>				
			</td>
			<td nowrap="nowrap">
					<%if checkboxCnt>"0" then%>
						<input type=hidden name="Check2" value="1">
						<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck1_2()"  class="clearBorder"> Check All
						<script type=text/javascript>
						function checkAll1_2() {
						for (var j = 1; j <= <%=checkboxCnt%>; j++) {
						box = eval("document.form1.checkEmail" + j); 
						if (box.checked == false) box.checked = true;
								}
						}

						function uncheckAll1_2() {
						for (var j = 1; j <= <%=checkboxCnt%>; j++) {
						box = eval("document.form1.checkEmail" + j); 
						if (box.checked == true) box.checked = false;
								 }
						}
						
						function testcheck1_2() {
						if (document.form1.Check2.value=="0") {
						document.form1.Check2.value="1";
						checkAll1_2();
								}
						else
								{
						document.form1.Check2.value="0";
						uncheckAll1_2();
								}
						}
						</script>
					<%end if%>				
				</td>
				<td colspan="6" class="pcCPspacer"></td>
			</tr>
			<% if noAuthRec=0  then %>
			<tr>
				<td colspan="8">
					<input type="submit" name="AuthSubmit" value="Process Selected Authorize.Net Orders" class="btn btn-primary">
				</td>
			</tr>
			<tr>
			  <td colspan="8">&nbsp;</td>
			</tr>
			<% end if %>
		</table>
		</form>
		<% 
	End If '// pcv_strDisplaySection = -1
		
END IF 
'////////////////////////////////////////////////////
'// END: Authorize Net
'////////////////////////////////////////////////////

'////////////////////////////////////////////////////
'// START: Authorize Net DPM
'////////////////////////////////////////////////////
IF gwaDPM=1 THEN
	
	query="SELECT orders.paymentCode FROM orders WHERE paymentCode='AuthorizeDPM';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If
	set rs=nothing '// If NOT rs.EOF Then

	If pcv_strDisplaySection = -1 Then
		'// Check for authorize.net orders
		query="SELECT authorders.idOrder, authorders.idauthorder, authorders.amount, authorders.paymentmethod, authorders.transtype, authorders.authcode, authorders.ccnum, authorders.ccexp, authorders.idCustomer, authorders.fname, authorders.lname, authorders.address, authorders.zip, authorders.pcSecurityKeyID, orders.orderDate, orders.orderStatus, orders.gwTransId, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email FROM customers INNER JOIN (authorders INNER JOIN orders ON authorders.idOrder = orders.idOrder) ON (authorders.idCustomer = customers.idcustomer) AND (customers.idcustomer = orders.idCustomer) WHERE (((authorders.transtype)='AUTH_ONLY') AND ((authorders.captured)=0)) AND authorders.paymentmethod='DPM' ORDER BY "&AuthOrder&" "&AuthSort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
		<form name="DPMform" method="post" action="batchprocess_authDPM.asp" class="pcForms">
			<table class="pcCPcontent">
				<tr>
					<td colspan="8"><h2>Authorize.Net DPM Orders</h2></td>
				</tr>
				<tr>
					<th>Process</th>
					<th nowrap="nowrap">Send Email</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authorders.idOrder&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authorders.idOrder&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
					<th>Customer</th>
					<th colspan="2" align="left">Total</th>
				</tr>
                <tr>
                	<td colspan="8" class="pcCPspacer"></td>
                </tr>
				<%
				noAuthRec=0
				if rs.eof then
					noAuthRec=1 
					%>
					<tr> 
						<td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
					</tr>
				<% end if %>
				<%
				checkboxCnt=0
				do until rs.eof
					checkboxCnt=checkboxCnt+1
					idOrder=rs("idOrder")
					idauthorder=rs("idauthorder")
					amount=rs("amount")
					paymentmethod=rs("paymentmethod")
					transtype=rs("transtype")
					authcode=rs("authcode")
					ccnum=rs("ccnum")
					ccexp=rs("ccexp")
					idCustomer=rs("idCustomer")
					fname=rs("fname")
					lname=rs("lname")
					address=rs("address")
					zip=rs("zip")
					pcv_SecurityKeyID=rs("pcSecurityKeyID")
					orderDate=rs("orderDate")
					orderStatus=rs("orderstatus")
					gwTransId=rs("gwTransId")
					stateCode=rs("stateCode")
					if stateCode="" then
						stateCode=rs("State")
					end if
					City=rs("city")
					countryCode=rs("countryCode")
					shippingAddress=rs("shippingAddress")
					shippingStateCode=rs("shippingStateCode")
					shippingState=rs("shippingState")
					shippingCity=rs("shippingCity")
					shippingCountryCode=rs("shippingCountryCode")
					shippingZip=rs("shippingZip")
					shippingFullName=rs("shippingFullName")
					address2=rs("address2")
					shippingCompany=rs("shippingCompany")
					shippingAddress2=rs("shippingAddress2")
					pcv_custcomments=trim(rs("comments"))
					pcv_admcomments=trim(rs("admincomments"))
					customerName=rs("name") & " " & rs("lastName")
					customerCompany=rs("customerCompany")
						if trim(customerCompany)<>"" then
							customerInfo=customerName & " (" & customerCompany & ")"
							else
							customerInfo=customerName
						end if
					phone=rs("phone")
					email =rs("email")
					
					pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
					ccnum2=enDeCrypt(ccnum, pcv_SecurityPass)
										
					'// Get amount from orders table
					query="SELECT total from orders WHERE idOrder="&idOrder&";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					curTotal=rstemp("total")
					set rstemp=nothing %>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td>
						<div align="center">
						<input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
						</div></td>
						<td>
						<div align="center">
						<input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
						</div></td>
						<td><%=ShowDateFrmt(orderDate)%></td>
						<td><%=authcode%>
						
						<input type="hidden" name="idauthorder<%=checkboxCnt%>" value="<%=idauthorder%>">
						<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=int(idOrder)+scpre%>">
						<input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
						</td>
						<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
						<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
						<td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
						<td><div align="center"><a href="batchprocessorders.asp?capture=<%=idauthorder%>&GW=authorders">Remove</a></div></td>
					</tr>
					<% rs.moveNext
				loop
				set rs=nothing
				%>
			<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
			<tr>
				<td nowrap="nowrap">
					<%if checkboxCnt>"0" then%>
					<input type=hidden name="DPMCheck1" value="0">
					<input type="checkbox" name="DPMCheck1a" value="1" onClick="javascript:testcheckDPM_1()" class="clearBorder"> Check All
					<script type=text/javascript>
						function checkAllDPM_1() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
								box = eval("document.DPMform.checkOrd" + j); 
								if (box.checked == false) box.checked = true;
							}
						}
	
						function uncheckAllDPM_1() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
								box = eval("document.DPMform.checkOrd" + j); 
								if (box.checked == true) box.checked = false;
							}
						}
						
						function testcheckDPM_1() {
						if (document.DPMform.DPMCheck1.value=="0") {
							document.DPMform.DPMCheck1.value="1";
							checkAllDPM_1();
						}
						else
						{
							document.DPMform.DPMCheck1.value="0";
							uncheckAllDPM_1();
						}
					}
					</script>
					<%end if%>				
			</td>
			<td nowrap="nowrap">
					<%if checkboxCnt>"0" then%>
						<input type=hidden name="DPMCheck2" value="1">
						<input type="checkbox" name="DPMCheck2a" checked value="1" onClick="javascript:testcheckDPM_2()"  class="clearBorder"> Check All
						<script type=text/javascript>
						function checkAllDPM_2() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
								box = eval("document.DPMform.checkEmail" + j); 
								if (box.checked == false) box.checked = true;
							}
						}

						function uncheckAllDPM_2() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
								box = eval("document.DPMform.checkEmail" + j); 
								if (box.checked == true) box.checked = false;
							}
						}
						
						function testcheckDPM_2() {
							if (document.DPMform.DPMCheck2.value=="0") {
								document.DPMform.DPMCheck2.value="1";
								checkAllDPM_2();
							}
							else
							{
								document.DPMform.DPMCheck2.value="0";
								uncheckAllDPM_2();
							}
						}
						</script>
					<%end if%>				
				</td>
				<td colspan="6" class="pcCPspacer"></td>
			</tr>
			<% if noAuthRec=0  then %>
			<tr>
				<td colspan="8">
					<input type="submit" name="AuthDPMSubmit" value="Process Selected Authorize.Net DPM Orders" class="btn btn-primary">
				</td>
			</tr>
			<tr>
			  <td colspan="8">&nbsp;</td>
			</tr>
			<% end if %>
		</table>
		</form>
		<% 
	End If '// pcv_strDisplaySection = -1
		
END IF 
'////////////////////////////////////////////////////
'// END: Authorize Net DPM
'////////////////////////////////////////////////////

'////////////////////////////////////////////////////
'// START: PayPal Classic APIs
'////////////////////////////////////////////////////
IF PayPalExp=1 OR PayPalDP=1 THEN
	AddPayPalTable("PayPalExp")
	AddPayPalTable("PayPalDP")

	Function AddPayPalTable(paymentCode)

		If paymentCode = "PayPalExp" Then
			paymentMethods = "'PayPalExp'"
			paymentName = "PayPal Express Checkout"
		Else
			paymentMethods = " 'PayPalDP', 'PayPalWP'"
			paymentName = PayPalDPName
		End If
		query="SELECT orders.paymentCode FROM orders WHERE paymentCode IN (" & paymentMethods & ");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcv_strDisplaySection = 0
		If NOT rs.EOF Then	
			pcv_strDisplaySection = -1
		End If
		set rs=nothing '// If NOT rs.EOF Then
		If pcv_strDisplaySection = -1 Then
	
			'// Check for PayPal orders
			query="SELECT pcPay_PayPal_Authorize.idPayPal_Authorize, pcPay_PayPal_Authorize.idOrder, orders.orderDate, orders.orderStatus, orders.gwTransId, pcPay_PayPal_Authorize.amount, pcPay_PayPal_Authorize.paymentmethod, pcPay_PayPal_Authorize.transtype, pcPay_PayPal_Authorize.authcode, pcPay_PayPal_Authorize.idCustomer, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany FROM customers INNER JOIN (pcPay_PayPal_Authorize INNER JOIN orders ON pcPay_PayPal_Authorize.idOrder = orders.idOrder) ON (pcPay_PayPal_Authorize.idCustomer = customers.idcustomer) AND (customers.idcustomer = orders.idCustomer) WHERE ((pcPay_PayPal_Authorize.transtype='Authorization') AND (pcPay_PayPal_Authorize.captured=0) AND (pcPay_PayPal_Authorize.paymentmethod IN (" & paymentMethods & "))) ORDER BY "&PayPalOrder&" "&PayPalSort&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if NOT rs.eof then
			%>
			<form name="form<%=paymentCode%>" method="post" action="batchprocess_paypal.asp" class="pcForms">
				<table class="pcCPcontent">
					<tr>
						<td colspan="8"><h2><%= paymentName %> Orders</h2></td>
					</tr>
					<tr>
						<th>Process</th>
						<th nowrap="nowrap">Send Email</th>
						<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&PayPalOrder=orders.orderdate&PayPalSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&PayPalOrder=orders.orderdate&PayPalSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
						<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&PayPalOrder=gwTransID&PayPalSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&PayPalOrder=gwTransID&PayPalSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
						<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&PayPalOrder=idOrder&PayPalSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&PayPalOrder=idOrder&PayPalSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
						<th>Customer</th>
						<th colspan="2" align="left">Total</th>
					</tr>
									<tr>
                		<td colspan="8" class="pcCPspacer"></td>
									</tr>
					<% 
					noAuthRec=0 
					checkboxCnt=0
					ActivecheckboxCnt=0
					do until rs.eof
						checkboxCnt=checkboxCnt+1
						idauthorder=rs("idPayPal_Authorize")
						idOrder=rs("idOrder")						
						orderDate=rs("orderDate")
						orderStatus=rs("orderStatus")
						gwTransId=rs("gwTransId")
						amount=rs("amount")
						paymentmethod=rs("paymentmethod")
						transtype=rs("transtype")
						authcode=rs("authcode")
						pcv_custcomments=trim(rs("comments"))
						pcv_admcomments=trim(rs("admincomments"))
						customerName=rs("name") & " " & rs("lastName")
						customerCompany=rs("customerCompany")
							if trim(customerCompany)<>"" then
								customerInfo=customerName & " (" & customerCompany & ")"
								else
								customerInfo=customerName
							end if
						'// Get amount from orders table
						query="SELECT total from orders WHERE idOrder="&idOrder&";"
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
						curTotal=rstemp("total")
						set rstemp=nothing 
					
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Start: Check Reauthorization needed
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						Dim pcv_strReAuthorizeFlag					
						pcv_strReAuthorizeFlag = 0
					
						'// Testing Mode
						'orderDate=dateadd("d",-5,Date())
					
						'// Check Date is within the Honor Period
						if Date() > dateadd("d",3,orderDate) then
							pcv_strReAuthorizeFlag = 1
						end if
						if Date() > dateadd("d",29,orderDate) then
							pcv_strReAuthorizeFlag = 2
						end if
					
						'// Testing Mode
						'amount=1076.1
						'curTotal=3002.25
					
						'// Check Current Amount
						if curTotal > amount then
							pcv_CurPriceDifference = abs(amount - curTotal)
							pcv_MaxAllowedPrice = (amount*1.15)
							'// no greater than $75.00 increase
							if pcv_CurPriceDifference > 75 then
								pcv_strReAuthorizeFlag = 3
							end if
							'// no greater 115% increase
							if curTotal => pcv_MaxAllowedPrice then
								pcv_strReAuthorizeFlag = 3
							end if
						end if					
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// End: Check Reauthorization needed
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
						'// Count the number of Updatable Orders
						if pcv_strReAuthorizeFlag = 0 OR pcv_strReAuthorizeFlag = 1 then
							ActivecheckboxCnt=ActivecheckboxCnt+1
						end if
					
						If gwTransId<>"" and isNULL(gwTransId)=False Then
						%>
						<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
							<td>
							<div align="center">
							<% if pcv_strReAuthorizeFlag = 2 then %>
								<input name="checkOrd<%=checkboxCnt%>_disabled" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" disabled class="clearBorder">
							<% else %>
								<input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">							
							<% end if %>						
							</div></td>
							<td>
							<div align="center">
							<input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
							</div></td>
							<td><%=ShowDateFrmt(orderDate)%></td>
							<td><a href="javascript:openwin('popup_PayPalTransSearch.asp?TransID=<%=gwTransId%>');"><%=gwTransId%></a>						
							<input type="hidden" name="Amount<%=checkboxCnt%>" value="<%=curTotal%>">
							<input type="hidden" name="idAuthOrder<%=checkboxCnt%>" value="<%=gwTransId%>">
							<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=int(idOrder)%>">					
							<input type="hidden" name="Note<%=checkboxCnt%>" value="<%=pcv_custcomments%>">
							<input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">					
							<input type="hidden" name="idCustomer<%=checkboxCnt%>" value="<%=idCustomer%>">
							<input type="hidden" name="authamount<%=checkboxCnt%>" value="<%=amount%>">
							<input type="hidden" name="authcode<%=checkboxCnt%>" value="<%=authcode%>">
							<input type="hidden" name="transid<%=checkboxCnt%>" value="<%=gwTransId%>">
							<input type="hidden" name="ccnum<%=checkboxCnt%>" value="<%=ccnum2%>">
							<input type="hidden" name="ccexp<%=checkboxCnt%>" value="<%=ccexp%>">
							<input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
						
							<input type="hidden" name="fName<%=checkboxCnt%>" value="<%=fname%>">
							<input type="hidden" name="lName<%=checkboxCnt%>" value="<%=lname%>">
							<input type="hidden" name="address<%=checkboxCnt%>" value="<%=address%>">
							<input type="hidden" name="zip<%=checkboxCnt%>" value="<%=zip%>">
							<input type="hidden" name="stateCode<%=checkboxCnt%>" value="<%=stateCode%>">					
							<input type="hidden" name="City<%=checkboxCnt%>" value="<%=city%>">
							<input type="hidden" name="countryCode<%=checkboxCnt%>" value="<%=countryCode%>">
							<input type="hidden" name="shippingAddress<%=checkboxCnt%>" value="<%=shippingAddress%>">
							<input type="hidden" name="shippingStateCode<%=checkboxCnt%>" value="<%=shippingStateCode%>">
							<input type="hidden" name="shippingState<%=checkboxCnt%>" value="<%=shippingState%>">
							<input type="hidden" name="shippingCity<%=checkboxCnt%>" value="<%=shippingCity%>">
							<input type="hidden" name="shippingCountryCode<%=checkboxCnt%>" value="<%=shippingCountryCode%>">
							<input type="hidden" name="shippingZip<%=checkboxCnt%>" value="<%=shippingZip%>">
							<input type="hidden" name="shippingFullName<%=checkboxCnt%>" value="<%=shippingFullName%>">
							<input type="hidden" name="address2<%=checkboxCnt%>" value="<%=address2%>">
							<input type="hidden" name="shippingCompany<%=checkboxCnt%>" value="<%=shippingCompany%>">
							<input type="hidden" name="shippingAddress2<%=checkboxCnt%>" value="<%=shippingAddress2%>"> 
							<input type="hidden" name="customerCompany<%=checkboxCnt%>" value="<%=customerCompany%>"> 
							<input type="hidden" name="phone<%=checkboxCnt%>" value="<%=phone%>"> 
							<input type="hidden" name="email<%=checkboxCnt%>" value="<%=email%>">					

							</td>
							<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
							<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
							<td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
							<td><div align="center"><a href="batchprocessorders.asp?capture=<%=idauthorder%>&GW=pcPay_PayPal_Authorize">Remove</a></div></td>
						</tr>
						<% if pcv_strReAuthorizeFlag = 1 then %>
						<tr>
							<td colspan="8" style="padding-bottom:8px"><span class="pcCPnotes">
								This order listed above is over 3 days. PayPal honors 100% of authorized funds for three days. You can settle without a reauthorization from day 4 to day 29 of the authorization period, but PayPal cannot ensure that 100% of the funds will be available after the three-day honor period.
							</span></td>
						</tr>
						<% elseif pcv_strReAuthorizeFlag = 2 then %>
						<tr>
							<td colspan="8" style="padding-bottom:8px"><span class="pcCPnotes">
							This order listed above is over 29 days and must be reauthorized. When your buyer approves an authorization, the buyer's balance can be placed on hold for a 29-day period to ensure the availability of the authorization amount for capture. You can reauthorize a transaction only once, up to 115% of the originally authorized amount (not to exceed an increase of $75 USD). <strong>You can Reauthorize from within the <a href="Orddetails.asp?id=<%=idOrder%>">Details</a>.</strong>
							</span></td>
						</tr>
						<% elseif pcv_strReAuthorizeFlag = 3 then %>
						<tr>
							<td colspan="8" style="padding-bottom:8px"><span class="pcCPnotes">
							This order's Total has increased over the original authorization max and must be reauthorized. You can reauthorize a transaction only once, up to 115% of the originally authorized amount (not to exceed an increase of $75 USD). <strong>You can Reauthorize from within the <a href="Orddetails.asp?id=<%=idOrder%>">Details</a>.</strong>
							</span></td>
						</tr>
						<% end if %>
						<% 
						End If
						rs.moveNext
					loop
					set rs=nothing
					%>
				
						<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
						<tr>
							<td nowrap="nowrap">
							<%if checkboxCnt>"0" AND ActivecheckboxCnt>0 then%>
							<input type=hidden name="Check1" value="0">
							<input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck<%=paymentCode%>_1()" class="clearBorder"> Check All
							<script type=text/javascript>
							function checkAll<%=paymentCode%>_1() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
							box = eval("document.form<%=paymentCode%>.checkOrd" + j); 
							if (box.checked == false) box.checked = true;
								}
							}
		
							function uncheckAll<%=paymentCode%>_1() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
							box = eval("document.form<%=paymentCode%>.checkOrd" + j); 
							if (box.checked == true) box.checked = false;
								 }
							}
						
							function testcheck<%=paymentCode%>_1() {
							if (document.form<%=paymentCode%>.Check1.value=="0") {
							document.form<%=paymentCode%>.Check1.value="1";
							checkAll<%=paymentCode%>_1();
								}
							else
								{
							document.form<%=paymentCode%>.Check1.value="0";
							uncheckAll<%=paymentCode%>_1();
								}
							}
							</script>
						<%end if%>				
						</td>
						<td nowrap="nowrap">
						<%if checkboxCnt>"0" AND ActivecheckboxCnt>0 then%>
							<input type=hidden name="Check2" value="1">
							<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck<%=paymentCode%>_2()"  class="clearBorder"> Check All
							<script type=text/javascript>
							function checkAll<%=paymentCode%>_2() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
							box = eval("document.form<%=paymentCode%>.checkEmail" + j); 
							if (box.checked == false) box.checked = true;
								}
							}
		
							function uncheckAll<%=paymentCode%>_2() {
							for (var j = 1; j <= <%=checkboxCnt%>; j++) {
							box = eval("document.form<%=paymentCode%>.checkEmail" + j); 
							if (box.checked == true) box.checked = false;
								 }
							}
						
							function testcheck<%=paymentCode%>_2() {
							if (document.form<%=paymentCode%>.Check2.value=="0") {
							document.form<%=paymentCode%>.Check2.value="1";
							checkAll<%=paymentCode%>_2();
								}
							else
								{
							document.form<%=paymentCode%>.Check2.value="0";
							uncheckAll<%=paymentCode%>_2();
								}
							}
							</script>
						<%end if%>
						</td>
						<td colspan="6" class="pcCPspacer"></td>
					</tr>
					<% if noAuthRec=0 AND ActivecheckboxCnt>0 then %>
					<tr>
						<td colspan="8">
							<input type="submit" name="AuthSubmit" value="Process Selected PayPal Orders" class="submit7 btn btn-primary">
						</td>
					</tr>
					<tr>
						<td colspan="8">&nbsp;</td>
					</tr>
					<% end if %>
				</table>
			</form>
			<% 
			end if
		
		End If '// pcv_strDisplaySection = -1
	End Function
END IF 
'////////////////////////////////////////////////////
'// END: PayPal Classic APIs
'////////////////////////////////////////////////////


'////////////////////////////////////////////////////
'// START: Link Point
'////////////////////////////////////////////////////
dim LinkPointOn
LinkPointAPIOn=0
query= "SELECT lp_yourpay FROM linkpoint where id=1"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number <> 0 then
	strErrorDescription=err.description
	set rs=nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
end If
lp_yourpay=rs("lp_yourpay")
if lp_yourpay="API" then
	LinkPointAPIOn=1
end if

If LinkPointAPIOn=1 then 
	query="select orders.*, pcPay_LinkPointAPI.pcPay_LPAPI_ID, pcPay_LinkPointAPI.pcPay_LPAPI_amount, pcPay_LinkPointAPI.pcPay_LPAPI_ccnum, pcPay_LinkPointAPI.pcPay_LPAPI_ccexpmonth, pcPay_LinkPointAPI.pcPay_LPAPI_ccexpyear, pcPay_LinkPointAPI.pcPay_LPAPI_RTDate, pcSecurityKeyID from orders, pcPay_LinkPointAPI WHERE orders.paymentCode='LinkPointApi' and orders.pcOrd_PaymentStatus<>2 and orders.idorder=pcPay_LinkPointAPI.idorder and pcPay_LinkPointAPI.pcPay_LPAPI_captured<>1;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if NOT rs.eof then
	
	%>
	<form name="form2" method="post" action="batchprocess_lp.asp" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="11"><h2>LinkPoint API Orders</h2></td>
		</tr>
		<tr>
		  <th>Process</th>
		  <th>Send Email</th>
			<th><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></th>
			<th>Date</th>
			<th><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=transid&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=transid&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></th>
			<th nowrap="nowrap">Transaction ID</th>
			<th><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authorders.idOrder&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authorders.idOrder&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></th>
			<th nowrap="nowrap">Order ID</th>
			<th colspan="3">Total</th>
		</tr>
        <tr>
            <td colspan="8" class="pcCPspacer"></td>
        </tr>
		<%
		noLPRec=0
		if rs.eof then
			noLPRec=1 
			%>
			<tr> 
				<td colspan="11"><div class="pcCPmessage">No pending records found</div></td>
			</tr>
		<% end if %>
		<% checkboxCnt=0
		do until rs.eof
			checkboxCnt=checkboxCnt+1
			idOrder=rs("idOrder")
			idCustomer=rs("idCustomer")
			orderDate=rs("orderDate")
			orderstatus=rs("orderstatus")
			paymentmethod=rs("paymentmethod")
			transtype=rs("transtype")
			authcode=rs("authcode")
			fname=rs("fname")
			lname=rs("lname")
			address=rs("address")
			zip=rs("zip")
			pcPay_LPAPI_ID=rs("pcPay_LPAPI_ID")
			amount=rs("pcPay_LPAPI_amount")
			ccnum=rs("pcPay_LPAPI_ccnum")
			ccexpmonth=rs("pcPay_LPAPI_ccexpmonth")
			ccexpyear=rs("pcPay_LPAPI_ccexpyear")
			lpdate = rs("pcPay_LPAPI_RTDate")
			pcv_SecurityKeyID=rs("pcSecurityKeyID")
			
			pCardNumber=ccnum
			
			pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
			ccnum=enDeCrypt(pCardNumber, pcv_SecurityPass)
				
			'get amount from orders table
			query="SELECT total from orders WHERE idOrder="&idOrder&";"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=conntemp.execute(query)
			curTotal=rstemp("total")
			set rstemp=nothing %>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td>
					<div align="center">
						<input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
					</div></td>
				<td>
					<div align="center">
					  <input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">				
				  </div></td>
				<td><%=ShowDateFrmt(orderDate)%></td>
				<td><%=authcode%>
				<input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
				<input type="hidden" name="fName<%=checkboxCnt%>" value="<%=fname%>">
				<input type="hidden" name="lName<%=checkboxCnt%>" value="<%=lname%>">
				<input type="hidden" name="address<%=checkboxCnt%>" value="<%=address%>">
				<input type="hidden" name="zip<%=checkboxCnt%>" value="<%=zip%>">
				<input type="hidden" name="idauthorder<%=checkboxCnt%>" value="<%=idauthorder%>">
				<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=idOrder%>">
				<input type="hidden" name="lpidOrder<%=checkboxCnt%>" value="<%=idOrder%>">
				<input type="hidden" name="lpamount<%=checkboxCnt%>" value="<%=amount%>">
				<input type="hidden" name="authcode<%=checkboxCnt%>" value="<%=authcode%>">
				<input type="hidden" name="ccnum<%=checkboxCnt%>" value="<%=ccnum%>">
				<input type="hidden" name="ccexpmonth<%=checkboxCnt%>" value="<%=ccexpmonth%>">
				<input type="hidden" name="ccexpyear<%=checkboxCnt%>" value="<%=ccexpyear%>">
				<input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
				<input type="hidden" name="lpdate<%=checkboxCnt%>" value="<%=lpdate%>">
				</td>
				<td colspan="2"><%=int(idOrder)+scpre%></td>
				<td><div align="center"><%=scCurSign&money(curTotal)%> </div></td>
				<td><a href="Orddetails.asp?id=<%=idOrder%>">View Details</a><%if pcv_custcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view customer comments"></a><%end if%></td>
				<td><div align="center"><a href="batchprocessorders.asp?capture=<%=pcPay_LPAPI_ID%>&GW=pcPay_LinkPointAPI">Remove</a></div></td>
			</tr>
			<% rs.moveNext
			loop
			set rs=nothing
			%>
			<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
			<tr>
				<td nowrap="nowrap">
				<%if checkboxCnt>"0" then%>
					<input type=hidden name="Check1" value="0">
					<input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck2_1()" class="clearBorder"> Check All
					<script type=text/javascript>
					function checkAll2_1() {
					for (var j = 1; j <= <%=checkboxCnt%>; j++) {
					box = eval("document.form2.checkOrd" + j); 
					if (box.checked == false) box.checked = true;
						}
					}
	
					function uncheckAll2_1() {
					for (var j = 1; j <= <%=checkboxCnt%>; j++) {
					box = eval("document.form2.checkOrd" + j); 
					if (box.checked == true) box.checked = false;
					   }
					}
					
					function testcheck2_1() {
					if (document.form2.Check1.value=="0") {
					document.form2.Check1.value="1";
					checkAll2_1();
						}
					else
						{
					document.form2.Check1.value="0";
					uncheckAll2_1();
						}
					}
					</script>
				<%end if%></td>
				<td nowrap="nowrap">
				<%if checkboxCnt>"0" then%>
					<input type=hidden name="Check2" value="1">
					<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck2_2()" class="clearBorder"> Check All
					<script type=text/javascript>
					function checkAll2_2() {
					for (var j = 1; j <= <%=checkboxCnt%>; j++) {
					box = eval("document.form2.checkEmail" + j); 
					if (box.checked == false) box.checked = true;
						}
					}
	
					function uncheckAll2_2() {
					for (var j = 1; j <= <%=checkboxCnt%>; j++) {
					box = eval("document.form2.checkEmail" + j); 
					if (box.checked == true) box.checked = false;
					   }
					}
					
					function testcheck2_2() {
					if (document.form2.Check2.value=="0") {
					document.form2.Check2.value="1";
					checkAll2_2();
						}
					else
						{
					document.form2.Check2.value="0";
					uncheckAll2_2();
						}
					}
					</script>
				<%end if%>
				</td>
				<td colspan="6">&nbsp;</td>
			</tr>
			<% if noLPRec=0 then %>
				<tr>
					<td colspan="8" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="8">
						<input type="submit" name="LPSubmit" value="Process Selected LinkPoint Orders" class="btn btn-primary"></td>
				</tr>
				<tr>
				  <td colspan="8">&nbsp;</td>
				</tr>
			<% end if %>
		</table>
		</form>
		<% 
	End If '// pcv_strDisplaySection = -1
		
END IF 
'////////////////////////////////////////////////////
'// END: Link Point
'////////////////////////////////////////////////////

'////////////////////////////////////////////////////
'// START: Payflow Editions
'////////////////////////////////////////////////////
IF PayPalPPA=1 OR PayPalPPP=1 OR PayPalPFL=1 THEN
	If PayPalPPA=1 Then
		TmpPFApiTitle=PayPalPPAName
		TmppaySource="PayPalAdvanced"
	ElseIf PayPalPPP=1 Then
		TmpPFApiTitle=PayPalPPPName
		TmppaySource="PayPalPro"
	ElseIf PayPalPFL=1 Then
		TmpPFApiTitle=PayPalPFLName
		TmppaySource="PFL"
	End If

	query="SELECT orders.paymentCode FROM orders WHERE paymentCode='" & TmppaySource & "';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If

	TmpPFApiAction="batchprocess_pfapi.asp"
	TmpProcessBtn="Process Selected Orders"

	If pcv_strDisplaySection = -1 Then
		'// Check for Payflow API orders
		query="SELECT pcPay_PFL_Authorize.idPFL_Authorize, pcPay_PFL_Authorize.idOrder, pcPay_PFL_Authorize.amount, pcPay_PFL_Authorize.paySource, pcPay_PFL_Authorize.paymentmethod, pcPay_PFL_Authorize.transtype, pcPay_PFL_Authorize.authcode, pcPay_PFL_Authorize.fraudcode, pcPay_PFL_Authorize.gwCode, orders.orderDate, orders.orderstatus, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany, customers.idcustomer, customers.email FROM customers INNER JOIN (pcPay_PFL_Authorize INNER JOIN orders ON pcPay_PFL_Authorize.idOrder = orders.idOrder) ON customers.idcustomer = orders.idCustomer WHERE (pcPay_PFL_Authorize.transtype='A' AND orders.orderStatus>=2 AND orders.orderstatus<>5 AND pcPay_PFL_Authorize.captured=0) ORDER BY "&PFApiOrder&" "&PFApiSort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		%>
		<form name="form3" method="post" action="<%=TmpPFApiAction%>" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="8"><h2><%=TmpPFApiTitle%> Orders</h2></td>
			</tr>
			<tr>
				<th>Process</th>
				<th nowrap="nowrap">Send Email</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFApiOrder=orders.orderdate&PFApiSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFApiOrder=orders.orderdate&PFApiSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFApiOrder=pcPay_PFL_Authorize.authcode&PFApiSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFApiOrder=pcPay_PFL_Authorize.authcode&PFApiSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFApiOrder=orders.idOrder&PFApiSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFApiOrder=orders.idOrder&PFApiSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
				<th>Customer</th>
				<th colspan="2" align="left">Total</th>
            </tr>
            <tr>
                <td colspan="8" class="pcCPspacer"></td>
            </tr>
            <% dim noPPARec
            noPPARec=0
            if rs.eof then
                noPPARec=1 
                %>
                <tr> 
                    <td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
                </tr>
            <% end if %>
            <% checkboxCnt=0
						checkboxRvwCnt=0
            do until rs.eof
                checkboxCnt=checkboxCnt+1
                
                idPFL_Authorize=rs("idPFL_Authorize")
                idOrder=rs("idOrder")
                amt=rs("amount")
                paySource=rs("paySource")
                paymentmethod=rs("paymentmethod")
                trxtype=rs("transtype")
                origid=rs("authcode")
                fraudcode=rs("fraudcode")
								gwCode=rs("gwCode")
                'acct=rs("acct")
                'expdate=rs("expdate")
                orderDate=rs("orderDate")
                orderStatus=rs("orderstatus")
                pcv_custcomments=trim(rs("comments"))
                pcv_admcomments=trim(rs("admincomments"))
                customerName=rs("name") & " " & rs("lastName")
                customerCompany=rs("customerCompany")
                if trim(customerCompany)<>"" then
                    customerInfo=customerName & " (" & customerCompany & ")"
                else
                    customerInfo=customerName
                end if
                idcustomer=rs("idcustomer")
                customeremail=rs("email")
                'pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
                'acct2=enDeCrypt(acct, pcv_SecurityPass)
							
								If paySource = "PayPalExp" Then
									paymentMethodTitle = getPFApiExpressTitle(gwCode)
								Else
									paymentMethodTitle = TmpPFApiTitle
								End If

								If IsNull(fraudcode) Then
									fraudcode = ""
								End If
                
                'get amount from orders table
                query="SELECT total from orders WHERE idOrder="&idOrder&";"
                set rstemp=server.CreateObject("ADODB.RecordSet")
                set rstemp=conntemp.execute(query)
								If Not rstemp.eof Then
									curTotal=rstemp("total")
								End If
                set rstemp=nothing

								if fraudcode<>"126" then %>
                    <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                        <td>
                        <div align="center">
                        <input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
                        </div></td>
                        <td>
                            <div align="center">
                                <input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
                        </div></td>
                        <td><%=ShowDateFrmt(orderDate)%></td>
                        <td>
													<a href="#" onClick="return false" title="<%= paymentMethodTitle %>"><%=origid%></a>
													<input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
													<input type="hidden" name="fullName<%=checkboxCnt%>" value="<%=customerName%>">
													<input type="hidden" name="pfpidorder<%=checkboxCnt%>" value="<%=idPFL_Authorize%>">
													<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=idOrder%>">
													<input type="hidden" name="pfpamount<%=checkboxCnt%>" value="<%=amt%>">
													<input type="hidden" name="origid<%=checkboxCnt%>" value="<%=origid%>">
													<input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
													<input type="hidden" name="email<%=checkboxCnt%>" value="<%=customeremail%>">
													<input type="hidden" name="idCustomer<%=checkboxCnt%>" value="<%=idCustomer%>">
                        </td>
                        <td><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
												<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
                        <td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
                        <td><div align="center"><a href="batchprocessorders.asp?capture=<%=idPFL_Authorize%>&GW=pcPay_PFL_Authorize">Remove</a></div></td>
                    </tr>
                <% else 
										checkboxRvwCnt=checkboxRvwCnt+1 
										If checkboxRvwCnt=1 Then %>
                        <tr>
                            <td colspan="8" style="background-color:#FF9"><h2><%=TmpPFApiTitle%> Pending Review</h2>
                        	</td>
                        </tr>
                        <tr>
                            <th colspan="8" style="background-color:#FF9">For all "Pending Review" orders it is recommended that you manually review each order before processing the order. If you choose to batch process one or more of these orders through this tool you will be automatically approving the order for capture and it will be processed. If you decide to not approve an order, you can click on the "remove" link to remove it from this list and update the order status from the order details page.</th>
                        </tr>
            		<% End If %>

					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist" style="background-color:#FF9"> 
                        <td>
                        <div align="center">
                        <input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
                        </div></td>
                        <td>
                            <div align="center">
                                <input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
                        </div></td>
                        <td><%=ShowDateFrmt(orderDate)%></td>
                        <td><%=origid%>
                        <input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
                        <input type="hidden" name="fullName<%=checkboxCnt%>" value="<%=customerName%>">
                        <input type="hidden" name="pfpidorder<%=checkboxCnt%>" value="<%=idPFL_Authorize%>">
                        <input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=idOrder%>">
                        <input type="hidden" name="fraudmode<%=checkboxCnt%>" value="1">
                        <input type="hidden" name="pfpamount<%=checkboxCnt%>" value="<%=amt%>">
                        <input type="hidden" name="origid<%=checkboxCnt%>" value="<%=origid%>">
                        <input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
                        <input type="hidden" name="email<%=checkboxCnt%>" value="<%=customeremail%>">
                        <input type="hidden" name="idCustomer<%=checkboxCnt%>" value="<%=idCustomer%>">
                        </td>
                        <td><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
												<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
                        <td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
                        <td><div align="center"><a href="batchprocessorders.asp?capture=<%=idPFL_Authorize%>&GW=pcPay_PFL_Authorize">Remove</a></div></td>
                    </tr>
				<% end if %>
                <% rs.moveNext
            loop
            set rs=nothing
            %>
            <input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
            <tr>
                <td nowrap="nowrap">
                <%if checkboxCnt>"0" then%>
                    <input type=hidden name="Check1" value="0">
                    <input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck3_1()" class="clearBorder"> Check All
                    <script type=text/javascript>
                    function checkAll3_1() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form3.checkOrd" + j); 
                    if (box.checked == false) box.checked = true;
                        }
                    }
    
                    function uncheckAll3_1() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form3.checkOrd" + j); 
                    if (box.checked == true) box.checked = false;
                         }
                    }
            
                    function testcheck3_1() {
                    if (document.form3.Check1.value=="0") {
                    document.form3.Check1.value="1";
                    checkAll3_1();
                        }
                    else
                        {
                    document.form3.Check1.value="0";
                    uncheckAll3_1();
                        }
                    }
                    </script>
                <%end if%>
                </td>
                <td nowrap="nowrap">
                <%if checkboxCnt>"0" then%>
                    <input type=hidden name="Check2" value="1">
                    <input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck3_2()" class="clearBorder"> Check All
                    <script type=text/javascript>
                    function checkAll3_2() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form3.checkEmail" + j); 
                    if (box.checked == false) box.checked = true;
                        }
                    }
    
                    function uncheckAll3_2() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form3.checkEmail" + j); 
                    if (box.checked == true) box.checked = false;
                         }
                    }
                    
                    function testcheck3_2() {
                    if (document.form3.Check2.value=="0") {
                    document.form3.Check2.value="1";
                    checkAll3_2();
                        }
                    else
                    {
                    document.form3.Check2.value="0";
                    uncheckAll3_2();
                        }
                    }
                    </script>
                <%end if%>
                </td>
                <td colspan="6">&nbsp;</td>
            </tr>
            <% if noPPARec=0  then %>
                <tr>
                    <td colspan="8" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <td colspan="8">
                        <input type="submit" name="PFPSubmit" value="<%=TmpProcessBtn %>" class="btn btn-primary">						</td>
                </tr>
                <tr>
                  <td colspan="8">&nbsp;</td>
                  </tr>
            <% end if %>
		</table>
		</form>
	<% End If
END IF 
'////////////////////////////////////////////////////
'// END: Payflow Editions
'////////////////////////////////////////////////////




'////////////////////////////////////////////////////
'// START: Payflo Pro
'////////////////////////////////////////////////////
IF gwvpfp=1 THEN 
	
	query="SELECT orders.paymentCode FROM orders WHERE paymentCode='PFPRO';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If
	set rs=nothing '// If NOT rs.EOF Then

	If pcv_strDisplaySection = -1 Then
		'// Check for payflow pro orders
		query="SELECT pfporders.idpfporder, pfporders.idOrder, pfporders.amt, pfporders.tender, pfporders.trxtype, pfporders.origid, pfporders.acct, pfporders.expdate, pfporders.idCustomer, pfporders.fullname, pfporders.street, pfporders.state, pfporders.email, pfporders.zip, orders.orderDate, orders.orderstatus, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany FROM customers INNER JOIN (pfporders INNER JOIN orders ON pfporders.idOrder = orders.idOrder) ON customers.idcustomer = orders.idCustomer WHERE (((pfporders.trxtype)='A') AND ((orders.orderStatus)>=2) AND (orders.orderstatus<>5) AND ((pfporders.captured)=0)) ORDER BY "&PFPOrder&" "&PFPSort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
		<form name="form3" method="post" action="batchprocess_pfp.asp" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="8"><h2>PayPal Payflow Pro Orders</h2></td>
			</tr>
			<tr>
				<th>Process</th>
				<th nowrap="nowrap">Send Email</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=orders.orderdate&PFPSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=orders.orderdate&PFPSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=origid&PFPSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=origid&PFPSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=pfporders.idOrder&PFPSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=pfporders.idOrder&PFPSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
				<th>Customer</th>
				<th colspan="2" align="left">Total</th>
				</tr>
                <tr>
                	<td colspan="8" class="pcCPspacer"></td>
                </tr>
				<% dim noPFPRec
				noPFPRec=0
				if rs.eof then
					noPFPRec=1 
					%>
					<tr> 
						<td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
					</tr>
				<% end if %>
				<% checkboxCnt=0
				do until rs.eof
					checkboxCnt=checkboxCnt+1
					idpfporder=rs("idpfporder")
					idOrder=rs("idOrder")
					amt=rs("amt")
					tender=rs("tender")
					trxtype=rs("trxtype")
					origid=rs("origid")
					acct=rs("acct")
					expdate=rs("expdate")
					idCustomer=rs("idCustomer")
					fullname=rs("fullname")
					street=rs("street")
					state=rs("state")
					email=rs("email")
					zip=rs("zip")
					orderDate=rs("orderDate")
					orderStatus=rs("orderstatus")
					pcv_custcomments=trim(rs("comments"))
					pcv_admcomments=trim(rs("admincomments"))
					customerName=rs("name") & " " & rs("lastName")
					customerCompany=rs("customerCompany")
					if trim(customerCompany)<>"" then
						customerInfo=customerName & " (" & customerCompany & ")"
						else
						customerInfo=customerName
					end if
					
					pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
					acct2=enDeCrypt(acct, pcv_SecurityPass)
					
					'get amount from orders table
					query="SELECT total from orders WHERE idOrder="&idOrder&";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					curTotal=rstemp("total")
					set rstemp=nothing  %>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td>
						<div align="center">
						<input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
						</div></td>
						<td>
							<div align="center">
								<input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
						</div></td>
						<td><%=ShowDateFrmt(orderDate)%></td>
						<td><%=origid%>
						<input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
						<input type="hidden" name="fullName<%=checkboxCnt%>" value="<%=fullname%>">
						<input type="hidden" name="street<%=checkboxCnt%>" value="<%=street%>">
						<input type="hidden" name="zip<%=checkboxCnt%>" value="<%=zip%>">
						<input type="hidden" name="state<%=checkboxCnt%>" value="<%=state%>">
						<input type="hidden" name="pfpidorder<%=checkboxCnt%>" value="<%=idpfporder%>">
						<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=idOrder%>">
						<input type="hidden" name="pfpamount<%=checkboxCnt%>" value="<%=amt%>">
						<input type="hidden" name="origid<%=checkboxCnt%>" value="<%=origid%>">
						<input type="hidden" name="acct<%=checkboxCnt%>" value="<%=acct2%>">
						<input type="hidden" name="expdate<%=checkboxCnt%>" value="<%=expdate%>">
						<input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
						<input type="hidden" name="email<%=checkboxCnt%>" value="<%=email%>">
						<input type="hidden" name="idCustomer<%=checkboxCnt%>" value="<%=idCustomer%>">
						</td>
						<td colspan="2"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
						<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
						<td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
						<td><div align="center"><a href="batchprocessorders.asp?capture=<%=idpfporder%>&GW=pfporders">Remove</a></div></td>
					</tr>
					<% rs.moveNext
				loop
				set rs=nothing
				%>
				<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
				<tr>
					<td nowrap="nowrap">
					<%if checkboxCnt>"0" then%>
						<input type=hidden name="Check1" value="0">
						<input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck3_1()" class="clearBorder"> Check All
						<script type=text/javascript>
						function checkAll3_1() {
						for (var j = 1; j <= <%=checkboxCnt%>; j++) {
						box = eval("document.form3.checkOrd" + j); 
						if (box.checked == false) box.checked = true;
							}
						}
		
						function uncheckAll3_1() {
						for (var j = 1; j <= <%=checkboxCnt%>; j++) {
						box = eval("document.form3.checkOrd" + j); 
						if (box.checked == true) box.checked = false;
							 }
						}
				
						function testcheck3_1() {
						if (document.form3.Check1.value=="0") {
						document.form3.Check1.value="1";
						checkAll3_1();
							}
						else
							{
						document.form3.Check1.value="0";
						uncheckAll3_1();
							}
						}
						</script>
					<%end if%>
					</td>
					<td nowrap="nowrap">
					<%if checkboxCnt>"0" then%>
						<input type=hidden name="Check2" value="1">
						<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck3_2()" class="clearBorder"> Check All
						<script type=text/javascript>
						function checkAll3_2() {
						for (var j = 1; j <= <%=checkboxCnt%>; j++) {
						box = eval("document.form3.checkEmail" + j); 
						if (box.checked == false) box.checked = true;
							}
						}
		
						function uncheckAll3_2() {
						for (var j = 1; j <= <%=checkboxCnt%>; j++) {
						box = eval("document.form3.checkEmail" + j); 
						if (box.checked == true) box.checked = false;
							 }
						}
						
						function testcheck3_2() {
						if (document.form3.Check2.value=="0") {
						document.form3.Check2.value="1";
						checkAll3_2();
							}
						else
						{
						document.form3.Check2.value="0";
						uncheckAll3_2();
							}
						}
						</script>
					<%end if%>
					</td>
					<td colspan="6">&nbsp;</td>
				</tr>
				<% if noPFPRec=0  then %>
					<tr>
						<td colspan="8" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="8">
							<input type="submit" name="PFPSubmit" value="Process Selected PayFlow Pro Orders" class="btn btn-primary">						</td>
					</tr>
					<tr>
					  <td colspan="8">&nbsp;</td>
					  </tr>
				<% end if %>
		</table>
		</form>
		<% 
	End If '// pcv_strDisplaySection = -1
		
END IF 
'////////////////////////////////////////////////////
'// START: Payflo Pro
'////////////////////////////////////////////////////




'////////////////////////////////////////////////////
'// START: Net Bill
'////////////////////////////////////////////////////
IF gwnetbill=1 THEN

	query="SELECT orders.idOrder, orders.orderDate, orders.orderstatus, orders.total, orders.idCustomer, orders.paymentCode FROM orders WHERE paymentCode='Netbill' and orderStatus=2;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If
	set rs=nothing '// If NOT rs.EOF Then

	If pcv_strDisplaySection = -1 Then
		'// Check for Netbilling orders
		query="SELECT netbillorders.idOrder, netbillorders.idnetbillorder, netbillorders.amount, netbillorders.paymentmethod, netbillorders.transtype, netbillorders.authcode, netbillorders.trans_id, netbillorders.ccnum, netbillorders.ccexp, netbillorders.idCustomer, netbillorders.fname, netbillorders.lname, netbillorders.address, netbillorders.zip, netbillorders.pcSecurityKeyID, orders.orderDate, orders.orderStatus, orders.gwTransId, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.comments, orders.adminComments, customers.name, customers.lastName, customers.customerCompany FROM customers INNER JOIN (netbillorders INNER JOIN orders ON netbillorders.idOrder = orders.idOrder) ON customers.idcustomer = orders.idCustomer WHERE (((netbillorders.transtype)='A') AND ((orders.orderStatus)=2) AND ((netbillorders.captured)=0)) ORDER BY "&NetbillOrder&" "&NetbillSort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
		<form name="form4" method="post" action="batchprocess_netbill.asp">
		<table class="pcCPcontent">
			<tr>
				<td colspan="8"><h2>Netbilling Orders</h2></td>
			</tr>
			<tr>
				<th>Process</th>
				<th nowrap="nowrap">Send Email</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=orders.orderdate&NetbillSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=orders.orderdate&NetbillSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=origid&NetbillSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=origid&NetbillSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=netbillorders.idOrder&NetbillSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=netbillorders.idOrder&NetbillSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
				<th>Customer</th>
				<th colspan="2" align="left">Total</th>
				</tr>
                <tr>
                	<td colspan="8" class="pcCPspacer"></td>
                </tr>
			<% dim noNetbillRec
			noNetbillRec=0
			if rs.eof then
				noNetbillRec=1 
				%>
				<tr> 
					<td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
				</tr>
			<% else %>
				<% checkboxCnt=0
                do until rs.eof
                    checkboxCnt=checkboxCnt+1
					idOrder=rs("idOrder")
                    idnetbillorder=rs("idnetbillorder")
					amount=rs("amount")
					paymentmethod=rs("paymentmethod")
					transtype=rs("transtype")
					authcode=rs("authcode")
					trans_id=rs("trans_id")
					ccnum=rs("ccnum")
					ccexp=rs("ccexp")
					idCustomer=rs("idCustomer")
					fname=rs("fname")
					lname=rs("lname")
					address=rs("address")
					zip=rs("zip")
					pcv_SecurityKeyID=rs("pcSecurityKeyID")
                    orderDate=rs("orderDate")
                    orderStatus=rs("orderstatus")
					gwTransId=rs("gwTransId")
                    stateCode=rs("stateCode")
                    if stateCode="" then
                        stateCode=rs("State")
                    end if
					City=rs("city")
					countryCode=rs("countryCode")
					shippingAddress=rs("shippingAddress")
					shippingStateCode=rs("shippingStateCode")
					shippingState=rs("shippingState")
					shippingCity=rs("shippingCity")
					shippingCountryCode=rs("shippingCountryCode")
					shippingZip=rs("shippingZip")
					shippingFullName=rs("shippingFullName")
					address2=rs("address2")
					shippingCompany=rs("shippingCompany")
					shippingAddress2=rs("shippingAddress2")
					pcv_custcomments=trim(rs("comments"))
					pcv_admcomments=trim(rs("admincomments"))
					customerName=rs("name") & " " & rs("lastName")
					customerCompany=rs("customerCompany")
					if trim(customerCompany)<>"" then
						customerInfo=customerName & " (" & customerCompany & ")"
						else
						customerInfo=customerName
					end if

					'get amount from orders table
					query="SELECT total from orders WHERE idOrder="&idOrder&";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					curTotal=rstemp("total")
					set rstemp=nothing  %>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                        <td>
                            <div align="center">
                                <input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
                            </div></td>
                        <td>
                            <div align="center">
                                <input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
                        </div></td>
                        <td><%=ShowDateFrmt(orderDate)%></td>
                        <td><%=trans_id%>
                        <% 
						pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
						ccnum2=enDeCrypt(ccnum, pcv_SecurityPass) %>
                        <input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
                        <input type="hidden" name="fName<%=checkboxCnt%>" value="<%=fname%>">
                        <input type="hidden" name="lName<%=checkboxCnt%>" value="<%=lname%>">
                        <input type="hidden" name="address<%=checkboxCnt%>" value="<%=address%>">
                        <input type="hidden" name="zip<%=checkboxCnt%>" value="<%=zip%>">
                        <input type="hidden" name="idnetbillorder<%=checkboxCnt%>" value="<%=idnetbillorder%>">
                        <input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=int(idOrder)+scpre%>">
                        <input type="hidden" name="netbillamount<%=checkboxCnt%>" value="<%=amount%>">
                        <input type="hidden" name="authcode<%=checkboxCnt%>" value="<%=authcode%>">
                        <input type="hidden" name="trans_id<%=checkboxCnt%>" value="<%=trans_id%>">
                        <input type="hidden" name="ccnum<%=checkboxCnt%>" value="<%=ccnum2%>">
                        <input type="hidden" name="ccexp<%=checkboxCnt%>" value="<%=ccexp%>">
                        <input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
                        <input type="hidden" name="stateCode<%=checkboxCnt%>" value="<%=stateCode%>">
                        <input type="hidden" name="idCustomer<%=checkboxCnt%>" value="<%=idCustomer%>">
                        <input type="hidden" name="City<%=checkboxCnt%>" value="<%=city%>">
                        <input type="hidden" name="countryCode<%=checkboxCnt%>" value="<%=countryCode%>">
                        <input type="hidden" name="shippingAddress<%=checkboxCnt%>" value="<%=shippingAddress%>">
                        <input type="hidden" name="shippingStateCode<%=checkboxCnt%>" value="<%=shippingStateCode%>">
                        <input type="hidden" name="shippingState<%=checkboxCnt%>" value="<%=shippingState%>">
                        <input type="hidden" name="shippingCity<%=checkboxCnt%>" value="<%=shippingCity%>">
                        <input type="hidden" name="shippingCountryCode<%=checkboxCnt%>" value="<%=shippingCountryCode%>">
                        <input type="hidden" name="shippingZip<%=checkboxCnt%>" value="<%=shippingZip%>">
                        <input type="hidden" name="shippingFullName<%=checkboxCnt%>" value="<%=shippingFullName%>">
                        <input type="hidden" name="address2<%=checkboxCnt%>" value="<%=address2%>">
                        <input type="hidden" name="shippingCompany<%=checkboxCnt%>" value="<%=shippingCompany%>">
                        <input type="hidden" name="shippingAddress2<%=checkboxCnt%>" value="<%=shippingAddress2%>">
                        </td>
                        <td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
                        <td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
                        <td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
                        <td><div align="center"><a href="batchprocessorders.asp?capture=<%=idnetbillorder%>&GW=netbillorders">Remove</a></div></td>
                    </tr>
                    <% rs.moveNext
                loop
                set rs=nothing
                %>
                <input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
                <tr>
                    <td nowrap="nowrap">
                    <%if checkboxCnt>"0" then%>
                    <input type=hidden name="Check1" value="0">
                    <input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck4_1()" class="clearBorder"> Check All
                    <script type=text/javascript>
                    function checkAll4_1() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form4.checkOrd" + j); 
                    if (box.checked == false) box.checked = true;
                        }
                    }
    
                    function uncheckAll4_1() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form4.checkOrd" + j); 
                    if (box.checked == true) box.checked = false;
                       }
                    }
                    
                    function testcheck4_1() {
                    if (document.form4.Check1.value=="0") {
                    document.form4.Check1.value="1";
                    checkAll4_1();
                        }
                    else
                        {
                    document.form4.Check1.value="0";
                    uncheckAll4_1();
                        }
                    }
                    </script>
                    <%end if%>
                    </td>
                    <td nowrap="nowrap">
                    <%if checkboxCnt>"0" then%>
                    <input type=hidden name="Check2" value="1">
                    <input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck4_2()" class="clearBorder"> Check All
                    <script type=text/javascript>
                    function checkAll4_2() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form4.checkEmail" + j); 
                    if (box.checked == false) box.checked = true;
                        }
                    }
    
                    function uncheckAll4_2() {
                    for (var j = 1; j <= <%=checkboxCnt%>; j++) {
                    box = eval("document.form4.checkEmail" + j); 
                    if (box.checked == true) box.checked = false;
                       }
                    }
                    
                    function testcheck4_2() {
                    if (document.form4.Check2.value=="0") {
                    document.form4.Check2.value="1";
                    checkAll4_2();
                        }
                    else
                        {
                    document.form4.Check2.value="0";
                    uncheckAll4_2();
                        }
                    }
                    </script>
                    <%end if%>
                    </td>
                    <td colspan="6">&nbsp;</td>
                </tr>
				<% if noNetbillRec=0  then %>
                <tr>
                    <td colspan="8" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <td colspan="8"><input type="submit" name="NetbillSubmit" value="Process Selected Netbilling Orders" class="btn btn-primary">				</td>
                </tr>
                <tr>
                  <td colspan="8">&nbsp;</td>
                  </tr>
                <% end if
			end if
			%>
		</table>
		</form>
		<% 
	End If '// pcv_strDisplaySection = -1
		
END IF 
'////////////////////////////////////////////////////
'// END: Net Bill
'////////////////////////////////////////////////////





'////////////////////////////////////////////////////
'// START: USA ePay
'////////////////////////////////////////////////////
IF gwUSAePay=1 THEN
	
    query="SELECT orders.idOrder, orders.orderDate, orders.orderstatus, orders.total, orders.idCustomer, orders.paymentCode FROM orders WHERE paymentCode='USAePay' and orderStatus=2;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If
	set rs=nothing '// If NOT rs.EOF Then

	If pcv_strDisplaySection = -1 Then
	
		'// Check for USAePay orders
		query="SELECT pcPay_USAePay_Orders.idOrder, pcPay_USAePay_Orders.idePayOrder, pcPay_USAePay_Orders.Amount, pcPay_USAePay_Orders.paymentmethod, pcPay_USAePay_Orders.transtype, pcPay_USAePay_Orders.RefNum, pcPay_USAePay_Orders.ccCard, pcPay_USAePay_Orders.ccExp, pcPay_USAePay_Orders.idCustomer, pcPay_USAePay_Orders.fname, pcPay_USAePay_Orders.lname, pcPay_USAePay_Orders.address, pcPay_USAePay_Orders.zip, pcPay_USAePay_Orders.pcSecurityKeyID, orders.orderDate, orders.orderStatus, orders.gwTransId, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.comments, orders.adminComments, customers.name, customers.lastName, customers.customerCompany FROM customers INNER JOIN (pcPay_USAePay_Orders INNER JOIN orders ON pcPay_USAePay_Orders.idOrder = orders.idOrder) ON customers.idcustomer = orders.idCustomer WHERE (((pcPay_USAePay_Orders.transtype)='0') AND ((orders.orderStatus)=2) AND ((pcPay_USAePay_Orders.captured)=0)) ORDER BY "&USAePayOrder&" "&USAePaySort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
		<form name="form5" method="post" action="batchprocess_USAePay.asp" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="8"><h2>USAePay Orders</h2></td>
			</tr>
			<tr>
				<th>Process</th>
				<th nowrap="nowrap">Send Email</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=orders.orderdate&NetbillSort=ASC&NetbillOrder=orders.orderdate&NetbillSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=orders.orderdate&NetbillSort=Desc&NetbillOrder=orders.orderdate&NetbillSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=origid&NetbillSort=ASC&NetbillOrder=origid&NetbillSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=origid&NetbillSort=Desc&NetbillOrder=origid&NetbillSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
				<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=pcPay_USAePay_Orders.idOrder&NetbillSort=ASC&NetbillOrder=pcPay_USAePay_Orders.idOrder&NetbillSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&NetbillOrder=pcPay_USAePay_Orders.idOrder&NetbillSort=Desc&NetbillOrder=pcPay_USAePay_Orders.idOrder&NetbillSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
				<th>Customer</th>
				<th colspan="2">Total</th>
				</tr>
                <tr>
                	<td colspan="8" class="pcCPspacer"></td>
                </tr>
			<% dim noUSAePayRec
			noUSAePayRec=0
			if rs.eof then
				noUSAePayRec=1 
				%>
				<tr> 
					<td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
				</tr>
			<% end if %>
			<% checkboxCnt=0
			do until rs.eof
				checkboxCnt=checkboxCnt+1
				idOrder=rs("idOrder")
				idePayOrder=rs("idePayOrder")
				amount=rs("amount")
				paymentmethod=rs("paymentmethod")
				transtype=rs("transtype")
				RefNum=rs("RefNum")
				ccCard=rs("ccCard")
				ccExp=rs("ccExp")
				idCustomer=rs("idCustomer")
				fname=rs("fname")
				lname=rs("lname")
				address=rs("address")
				zip=rs("zip")
				pcv_SecurityKeyID=rs("pcSecurityKeyID")
				orderDate=rs("orderDate")
				orderStatus=rs("orderstatus")
				gwTransId=rs("gwTransId")
				stateCode=rs("stateCode")
				if stateCode="" then
					stateCode=rs("State")
				end if
				City=rs("city")
				countryCode=rs("countryCode")
				shippingAddress=rs("shippingAddress")
				shippingStateCode=rs("shippingStateCode")
				shippingState=rs("shippingState")
				shippingCity=rs("shippingCity")
				shippingCountryCode=rs("shippingCountryCode")
				shippingZip=rs("shippingZip")
				shippingFullName=rs("shippingFullName")
				address2=rs("address2")
				shippingCompany=rs("shippingCompany")
				shippingAddress2=rs("shippingAddress2")
				pcv_custcomments=trim(rs("comments"))
				pcv_admcomments=trim(rs("admincomments"))
				customerName=rs("name") & " " & rs("lastName")
				customerCompany=rs("customerCompany")
				if trim(customerCompany)<>"" then
					customerInfo=customerName & " (" & customerCompany & ")"
					else
					customerInfo=customerName
				end if
					
				'get amount from orders table
				query="SELECT total from orders WHERE idOrder="&idOrder&";"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				curTotal=rstemp("total")
				set rstemp=nothing  %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td>
						<div align="center">
							<input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder">
						</div></td>
					<td>
						<div align="center">
							<input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
					</div></td>
					<td><%=ShowDateFrmt(orderDate)%></td>
					<td><%=RefNum%>
					<% 
					pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

					ccCard2=enDeCrypt(ccCard, pcv_SecurityPass) %>
					<input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
					<input type="hidden" name="fName<%=checkboxCnt%>" value="<%=fname%>">
					<input type="hidden" name="lName<%=checkboxCnt%>" value="<%=lname%>">
					<input type="hidden" name="address<%=checkboxCnt%>" value="<%=address%>">
					<input type="hidden" name="zip<%=checkboxCnt%>" value="<%=zip%>">
					<input type="hidden" name="idePayOrder<%=checkboxCnt%>" value="<%=idePayOrder%>">
					<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=int(idOrder)+scpre%>">
					<input type="hidden" name="USAePayamount<%=checkboxCnt%>" value="<%=amount%>">
					<input type="hidden" name="RefNum<%=checkboxCnt%>" value="<%=RefNum%>">
					<input type="hidden" name="ccCard<%=checkboxCnt%>" value="<%=ccCard2%>">
					<input type="hidden" name="ccExp<%=checkboxCnt%>" value="<%=ccExp%>">
					<input type="hidden" name="curamount<%=checkboxCnt%>" value="<%=curTotal%>">
					<input type="hidden" name="stateCode<%=checkboxCnt%>" value="<%=stateCode%>">
					<input type="hidden" name="idCustomer<%=checkboxCnt%>" value="<%=idCustomer%>">
					<input type="hidden" name="City<%=checkboxCnt%>" value="<%=city%>">
					<input type="hidden" name="countryCode<%=checkboxCnt%>" value="<%=countryCode%>">
					<input type="hidden" name="shippingAddress<%=checkboxCnt%>" value="<%=shippingAddress%>">
					<input type="hidden" name="shippingStateCode<%=checkboxCnt%>" value="<%=shippingStateCode%>">
					<input type="hidden" name="shippingState<%=checkboxCnt%>" value="<%=shippingState%>">
					<input type="hidden" name="shippingCity<%=checkboxCnt%>" value="<%=shippingCity%>">
					<input type="hidden" name="shippingCountryCode<%=checkboxCnt%>" value="<%=shippingCountryCode%>">
					<input type="hidden" name="shippingZip<%=checkboxCnt%>" value="<%=shippingZip%>">
					<input type="hidden" name="shippingFullName<%=checkboxCnt%>" value="<%=shippingFullName%>">
					<input type="hidden" name="address2<%=checkboxCnt%>" value="<%=address2%>">
					<input type="hidden" name="shippingCompany<%=checkboxCnt%>" value="<%=shippingCompany%>">
					<input type="hidden" name="shippingAddress2<%=checkboxCnt%>" value="<%=shippingAddress2%>">
					</td>
					<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
					<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
					<td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
					<td><div align="center"><a href="batchprocessorders.asp?capture=<%=idePayOrder%>&GW=pcPay_USAePay_Orders">Remove</a></div></td>
				</tr>
				<% rs.moveNext
			loop
			set rs=nothing
			%>
			<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
			<tr>
				<td nowrap="nowrap">
				<%if checkboxCnt>"0" then%>
				<input type=hidden name="Check1" value="0">
				<input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck5_1()" class="clearBorder"> Check All
				<script type=text/javascript>
				function checkAll5_1() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form5.checkOrd" + j); 
				if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAll5_1() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form5.checkOrd" + j); 
				if (box.checked == true) box.checked = false;
				   }
				}
				
				function testcheck5_1() {
				if (document.form5.Check1.value=="0") {
				document.form5.Check1.value="1";
				checkAll5_1();
					}
				else
					{
				document.form5.Check1.value="0";
				uncheckAll5_1();
					}
				}
				</script>
				<%end if%>
				</td>
				<td nowrap="nowrap">
				<%if checkboxCnt>"0" then%>
				<input type=hidden name="Check2" value="1">
				<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck5_2()" class="clearBorder"> Check All
				<script type=text/javascript>
				function checkAll5_2() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form5.checkEmail" + j); 
				if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAll5_2() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form5.checkEmail" + j); 
				if (box.checked == true) box.checked = false;
				   }
				}
				
				function testcheck5_2() {
				if (document.form5.Check2.value=="0") {
				document.form5.Check2.value="1";
				checkAll5_2();
					}
				else
					{
				document.form5.Check2.value="0";
				uncheckAll5_2();
					}
				}
				</script>
				<%end if%>
				</td>
				<td colspan="6">&nbsp;</td>
			</tr>
			<% if noUSAePayRec=0  then %>
			<tr>
				<td colspan="8" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="8">
					<input type="submit" name="USAePaySubmit" value="Process Selected USAePay Orders" class="btn btn-primary">
				</td>
			</tr>
			<tr>
			  <td colspan="8">&nbsp;</td>
			  </tr>
			<% end if %>
		</table>
		</form>
		<% 
	End If '// pcv_strDisplaySection = -1
		
END IF 
'////////////////////////////////////////////////////
'// START: USA ePay
'////////////////////////////////////////////////////


'////////////////////////////////////////////////////
'// START: EIG
'////////////////////////////////////////////////////
IF gwEIG=1 THEN
	
	query="SELECT orders.paymentCode FROM orders WHERE paymentCode='EIG';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If
	set rs=nothing '// If NOT rs.EOF Then

	If pcv_strDisplaySection = -1 Then
		'// Check for EIG orders
		query="SELECT pcPay_EIG_Authorize.idOrder, pcPay_EIG_Authorize.vaultToken, pcPay_EIG_Authorize.idauthorder, pcPay_EIG_Authorize.amount, pcPay_EIG_Authorize.paymentmethod, pcPay_EIG_Authorize.transtype, pcPay_EIG_Authorize.authcode, pcPay_EIG_Authorize.ccnum, pcPay_EIG_Authorize.ccexp, pcPay_EIG_Authorize.idCustomer, pcPay_EIG_Authorize.fname, pcPay_EIG_Authorize.lname, pcPay_EIG_Authorize.address, pcPay_EIG_Authorize.zip, pcPay_EIG_Authorize.pcSecurityKeyID, orders.orderDate, orders.orderStatus, orders.gwTransId, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email FROM customers INNER JOIN (pcPay_EIG_Authorize INNER JOIN orders ON pcPay_EIG_Authorize.idOrder = orders.idOrder) ON (pcPay_EIG_Authorize.idCustomer = customers.idcustomer) AND (customers.idcustomer = orders.idCustomer) WHERE (((pcPay_EIG_Authorize.transtype)='AUTH_ONLY') AND ((pcPay_EIG_Authorize.captured)=0)) ORDER BY "&AuthOrder&" "&AuthSort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
        <form name="formEIG" method="post" action="batchprocess_EIG.asp" class="pcForms">		
        	<table class="pcCPcontent">
				<tr>
					<td colspan="8"><h2><h2>NetSource Commerce Gateway Orders</h2></td>
				</tr>
				<tr>
					<th>Process</th>
					<th nowrap="nowrap">Send Email</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=pcPay_EIG_Authorize.idOrder&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=pcPay_EIG_Authorize.idOrder&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
					<th>Customer</th>
					<th colspan="2" align="left">Total</th>
				</tr>
                <tr>
                	<td colspan="8" class="pcCPspacer"></td>
                </tr>
				<% dim noAuthRecEIG
				noAuthRecEIG=0
				if rs.eof then
					noAuthRecEIG=1 
					%>
					<tr> 
						<td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
					</tr>
				<% end if %>
				<% dim EIGcheckboxCnt
				EIGcheckboxCnt=0
				do until rs.eof
					EIGcheckboxCnt=EIGcheckboxCnt+1
					idOrder=rs("idOrder")
					vaultToken=rs("vaultToken")
					idauthorder=rs("idauthorder")
					amount=rs("amount")
					paymentmethod=rs("paymentmethod")
					transtype=rs("transtype")
					authcode=rs("authcode")
					ccnum=rs("ccnum")
					ccexp=rs("ccexp")
					idCustomer=rs("idCustomer")
					fname=rs("fname")
					lname=rs("lname")
					address=rs("address")
					zip=rs("zip")
					pcv_SecurityKeyID=rs("pcSecurityKeyID")
					orderDate=rs("orderDate")
					orderStatus=rs("orderstatus")
					gwTransId=rs("gwTransId")
					stateCode=rs("stateCode")
					if stateCode="" then
						stateCode=rs("State")
					end if
					City=rs("city")
					countryCode=rs("countryCode")
					shippingAddress=rs("shippingAddress")
					shippingStateCode=rs("shippingStateCode")
					shippingState=rs("shippingState")
					shippingCity=rs("shippingCity")
					shippingCountryCode=rs("shippingCountryCode")
					shippingZip=rs("shippingZip")
					shippingFullName=rs("shippingFullName")
					address2=rs("address2")
					shippingCompany=rs("shippingCompany")
					shippingAddress2=rs("shippingAddress2")
					pcv_custcomments=trim(rs("comments"))
					pcv_admcomments=trim(rs("admincomments"))
					customerName=rs("name") & " " & rs("lastName")
					customerCompany=rs("customerCompany")
						if trim(customerCompany)<>"" then
							customerInfo=customerName & " (" & customerCompany & ")"
							else
							customerInfo=customerName
						end if
					phone=rs("phone")
					email =rs("email")
					
					pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
					ccnum2=enDeCrypt(ccnum, pcv_SecurityPass)
					vaultToken=enDeCrypt(vaultToken, pcv_SecurityPass)
										
					'// Get amount from orders table
					query="SELECT total from orders WHERE idOrder="&idOrder&";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					curTotal=rstemp("total")
					set rstemp=nothing %>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td>
						<div align="center">
						<input name="checkOrd<%=EIGcheckboxCnt%>" type="checkbox" id="checkOrd<%=EIGcheckboxCnt%>" value="YES" class="clearBorder">
						</div></td>
						<td>
						<div align="center">
						<input name="checkEmail<%=EIGcheckboxCnt%>" type="checkbox" id="checkEmail<%=EIGcheckboxCnt%>" value="YES" checked class="clearBorder">
						</div></td>
						<td><%=ShowDateFrmt(orderDate)%></td>
						<td><%=authcode%>
						
						<input type="hidden" name="orderstatus<%=EIGcheckboxCnt%>" value="<%=orderStatus%>">
                        <input type="hidden" name="vaultToken<%=EIGcheckboxCnt%>" value="<%=vaultToken%>">
                        <input type="hidden" name="SecurityKeyID<%=EIGcheckboxCnt%>" value="<%=pcv_SecurityKeyID%>">
						<input type="hidden" name="fName<%=EIGcheckboxCnt%>" value="<%=fname%>">
						<input type="hidden" name="lName<%=EIGcheckboxCnt%>" value="<%=lname%>">
						<input type="hidden" name="address<%=EIGcheckboxCnt%>" value="<%=address%>">
						<input type="hidden" name="zip<%=EIGcheckboxCnt%>" value="<%=zip%>">
						<input type="hidden" name="idauthorder<%=EIGcheckboxCnt%>" value="<%=idauthorder%>">
						<input type="hidden" name="idOrder<%=EIGcheckboxCnt%>" value="<%=int(idOrder)+scpre%>">
						<input type="hidden" name="authamount<%=EIGcheckboxCnt%>" value="<%=amount%>">
						<input type="hidden" name="authcode<%=EIGcheckboxCnt%>" value="<%=authcode%>">
						<input type="hidden" name="transid<%=EIGcheckboxCnt%>" value="<%=gwTransId%>">
						<input type="hidden" name="ccnum<%=EIGcheckboxCnt%>" value="<%=ccnum2%>">
						<input type="hidden" name="ccexp<%=EIGcheckboxCnt%>" value="<%=ccexp%>">
						<input type="hidden" name="curamount<%=EIGcheckboxCnt%>" value="<%=curTotal%>">
						<input type="hidden" name="stateCode<%=EIGcheckboxCnt%>" value="<%=stateCode%>">
						<input type="hidden" name="idCustomer<%=EIGcheckboxCnt%>" value="<%=idCustomer%>">
						<input type="hidden" name="City<%=EIGcheckboxCnt%>" value="<%=city%>">
						<input type="hidden" name="countryCode<%=EIGcheckboxCnt%>" value="<%=countryCode%>">
						<input type="hidden" name="shippingAddress<%=EIGcheckboxCnt%>" value="<%=shippingAddress%>">
						<input type="hidden" name="shippingStateCode<%=EIGcheckboxCnt%>" value="<%=shippingStateCode%>">
						<input type="hidden" name="shippingState<%=EIGcheckboxCnt%>" value="<%=shippingState%>">
						<input type="hidden" name="shippingCity<%=EIGcheckboxCnt%>" value="<%=shippingCity%>">
						<input type="hidden" name="shippingCountryCode<%=EIGcheckboxCnt%>" value="<%=shippingCountryCode%>">
						<input type="hidden" name="shippingZip<%=EIGcheckboxCnt%>" value="<%=shippingZip%>">
						<input type="hidden" name="shippingFullName<%=EIGcheckboxCnt%>" value="<%=shippingFullName%>">
						<input type="hidden" name="address2<%=EIGcheckboxCnt%>" value="<%=address2%>">
						<input type="hidden" name="shippingCompany<%=EIGcheckboxCnt%>" value="<%=shippingCompany%>">
						<input type="hidden" name="shippingAddress2<%=EIGcheckboxCnt%>" value="<%=shippingAddress2%>"> 
						<input type="hidden" name="customerCompany<%=EIGcheckboxCnt%>" value="<%=customerCompany%>"> 
						<input type="hidden" name="phone<%=EIGcheckboxCnt%>" value="<%=phone%>"> 
						<input type="hidden" name="email<%=EIGcheckboxCnt%>" value="<%=email%>">
						</td>
						<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
						<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
						<td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
						<td><div align="center"><a href="batchprocessorders.asp?capture=<%=idauthorder%>&GW=pcPay_EIG_Authorize">Remove</a></div></td>
					</tr>
					<% rs.moveNext
				loop
				set rs=nothing
				%>
			<input type="hidden" name="EIGcheckboxCnt" value="<%=EIGcheckboxCnt%>">
			<tr>
				<td nowrap="nowrap">
					<%if EIGcheckboxCnt>"0" then%>
					<input type=hidden name="Check1" value="0">
					<input type="checkbox" name="Check1a" value="1" onClick="javascript:EIGtestcheck1_1()" class="clearBorder"> Check All
					<script type=text/javascript>
					function EIGcheckAll1_1() {
					for (var j = 1; j <= <%=EIGcheckboxCnt%>; j++) {
					box = eval("document.formEIG.checkOrd" + j); 
					if (box.checked == false) box.checked = true;
							}
					}

					function EIGuncheckAll1_1() {
					for (var j = 1; j <= <%=EIGcheckboxCnt%>; j++) {
					box = eval("document.formEIG.checkOrd" + j); 
					if (box.checked == true) box.checked = false;
							 }
					}
					
					function EIGtestcheck1_1() {
					if (document.formEIG.Check1.value=="0") {
					document.formEIG.Check1.value="1";
					EIGcheckAll1_1();
							}
					else
							{
					document.formEIG.Check1.value="0";
					EIGuncheckAll1_1();
							}
					}
					</script>
					<%end if%>				
			</td>
			<td nowrap="nowrap">
					<%if EIGcheckboxCnt>"0" then%>
						<input type=hidden name="Check2" value="1">
						<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:EIGtestcheck1_2()"  class="clearBorder"> Check All
						<script type=text/javascript>
						function EIGcheckAll1_2() {
						for (var j = 1; j <= <%=EIGcheckboxCnt%>; j++) {
						box = eval("document.formEIG.checkEmail" + j); 
						if (box.checked == false) box.checked = true;
								}
						}

						function EIGuncheckAll1_2() {
						for (var j = 1; j <= <%=EIGcheckboxCnt%>; j++) {
						box = eval("document.formEIG.checkEmail" + j); 
						if (box.checked == true) box.checked = false;
								 }
						}
						
						function EIGtestcheck1_2() {
						if (document.formEIG.Check2.value=="0") {
						document.formEIG.Check2.value="1";
						EIGcheckAll1_2();
								}
						else
								{
						document.formEIG.Check2.value="0";
						EIGuncheckAll1_2();
								}
						}
						</script>
					<%end if%>				
				</td>
				<td colspan="6" class="pcCPspacer"></td>
			</tr>
			<% if noAuthRecEIG=0  then %>
			<tr>
				<td colspan="8">
					<input type="submit" name="AuthSubmit" value="Process Selected NetSource Commerce Gateway Orders" class="btn btn-primary">
				</td>
			</tr>
			<tr>
			  <td colspan="8">&nbsp;</td>
			</tr>
			<% end if %>
		</table>
		</form>
		<% 
	End If '// pcv_strDisplaySection = -1
		
END IF 
'////////////////////////////////////////////////////
'// END: EIG
'////////////////////////////////////////////////////


'////////////////////////////////////////////////////
'// START: EPN
'////////////////////////////////////////////////////
IF gwEPN=1 THEN
	query="SELECT orders.paymentCode FROM orders WHERE paymentCode='eProcessingNetwork';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_strDisplaySection = 0
	If NOT rs.EOF Then	
		pcv_strDisplaySection = -1
	End If
	set rs=nothing '// If NOT rs.EOF Then
	
	If pcv_strDisplaySection = -1 Then
		'// Check for EIG orders
		query="SELECT pcPay_EPN_Authorize.idauthorder, pcPay_EPN_Authorize.idOrder, pcPay_EPN_Authorize.amount, pcPay_EPN_Authorize.transtype, pcPay_EPN_Authorize.authcode, pcPay_EPN_Authorize.idCustomer, orders.orderDate, orders.orderStatus, orders.gwTransId, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email FROM customers INNER JOIN (pcPay_EPN_Authorize INNER JOIN orders ON pcPay_EPN_Authorize.idOrder = orders.idOrder) ON (pcPay_EPN_Authorize.idCustomer = customers.idcustomer) AND (customers.idcustomer = orders.idCustomer) WHERE (((pcPay_EPN_Authorize.transtype)='AuthOnly') AND ((pcPay_EPN_Authorize.captured)=0)) ORDER BY "&AuthOrder&" "&AuthSort&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
        <form name="formEPN" method="post" action="batchprocess_EPN.asp" class="pcForms">		
        	<table class="pcCPcontent">
				<tr>
					<td colspan="8"><h2>eProcessing Network Gateway Orders</h2></td>
				</tr>
				<tr>
					<th>Process</th>
					<th nowrap="nowrap">Send Email</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=orders.orderdate&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=authcode&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Trans. ID</th>
					<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=pcPay_EPN_Authorize.idOrder&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&PFLOrder=<%=PFLOrder%>&LinkOrder=<%=LinkOrder%>&AuthOrder=pcPay_EIG_Authorize.idOrder&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
					<th>Customer</th>
					<th colspan="2" align="left">Total</th>
				</tr>
        <tr>
          <td colspan="8" class="pcCPspacer"></td>
        </tr>
				<% dim noAuthRecEPN
				noAuthRecEPN=0
				if rs.eof then
					noAuthRecEPN=1 
					%>
					<tr> 
						<td colspan="8"><div class="pcCPmessage">No pending records found</div></td>
					</tr>
				<% end if %>
				<% dim EPNcheckboxCnt
				EPNcheckboxCnt=0
				do until rs.eof
					EPNcheckboxCnt=EPNcheckboxCnt+1
					idOrder=rs("idOrder")
					idauthorder=rs("idauthorder")
					amount=rs("amount")
					transtype=rs("transtype")
					authcode=rs("authcode")
					idCustomer=rs("idCustomer")
					address=rs("address")
					zip=rs("zip")
					orderDate=rs("orderDate")
					orderStatus=rs("orderstatus")
					gwTransId=rs("gwTransId")
					stateCode=rs("stateCode")
					if stateCode="" then
						stateCode=rs("State")
					end if
					City=rs("city")
					countryCode=rs("countryCode")
					shippingAddress=rs("shippingAddress")
					shippingStateCode=rs("shippingStateCode")
					shippingState=rs("shippingState")
					shippingCity=rs("shippingCity")
					shippingCountryCode=rs("shippingCountryCode")
					shippingZip=rs("shippingZip")
					shippingFullName=rs("shippingFullName")
					address2=rs("address2")
					shippingCompany=rs("shippingCompany")
					shippingAddress2=rs("shippingAddress2")
					pcv_custcomments=trim(rs("comments"))
					pcv_admcomments=trim(rs("admincomments"))
					customerName=rs("name") & " " & rs("lastName")
					customerCompany=rs("customerCompany")
						if trim(customerCompany)<>"" then
							customerInfo=customerName & " (" & customerCompany & ")"
							else
							customerInfo=customerName
						end if
					phone=rs("phone")
					email =rs("email")
					
					'// Get amount from orders table
					query="SELECT total from orders WHERE idOrder="&idOrder&";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					curTotal=rstemp("total")
					set rstemp=nothing
					%>
					
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td>
						<div align="center">
						<input name="checkOrd<%=EPNcheckboxCnt%>" type="checkbox" id="checkOrd<%=EPNcheckboxCnt%>" value="YES" class="clearBorder">
						</div></td>
						<td>
						<div align="center">
						<input name="checkEmail<%=EPNcheckboxCnt%>" type="checkbox" id="checkEmail<%=EPNcheckboxCnt%>" value="YES" checked class="clearBorder">
						</div></td>
						<td><%=ShowDateFrmt(orderDate)%></td>
						<td nowrap><%=authcode%>
						
						<input type="hidden" name="orderstatus<%=EPNcheckboxCnt%>" value="<%=orderStatus%>">
						<input type="hidden" name="address<%=EPNcheckboxCnt%>" value="<%=address%>">
						<input type="hidden" name="zip<%=EPNcheckboxCnt%>" value="<%=zip%>">
						<input type="hidden" name="idauthorder<%=EPNcheckboxCnt%>" value="<%=idauthorder%>">
						<input type="hidden" name="idOrder<%=EPNcheckboxCnt%>" value="<%=int(idOrder)+scpre%>">
						<input type="hidden" name="authamount<%=EPNcheckboxCnt%>" value="<%=amount%>">
						<input type="hidden" name="authcode<%=EPNcheckboxCnt%>" value="<%=authcode%>">
						<input type="hidden" name="transid<%=EPNcheckboxCnt%>" value="<%=gwTransId%>">
						<input type="hidden" name="curamount<%=EPNcheckboxCnt%>" value="<%=curTotal%>">
						<input type="hidden" name="stateCode<%=EPNcheckboxCnt%>" value="<%=stateCode%>">
						<input type="hidden" name="idCustomer<%=EPNcheckboxCnt%>" value="<%=idCustomer%>">
						<input type="hidden" name="City<%=EPNcheckboxCnt%>" value="<%=city%>">
						<input type="hidden" name="countryCode<%=EPNcheckboxCnt%>" value="<%=countryCode%>">
						<input type="hidden" name="shippingAddress<%=EPNcheckboxCnt%>" value="<%=shippingAddress%>">
						<input type="hidden" name="shippingStateCode<%=EPNcheckboxCnt%>" value="<%=shippingStateCode%>">
						<input type="hidden" name="shippingState<%=EPNcheckboxCnt%>" value="<%=shippingState%>">
						<input type="hidden" name="shippingCity<%=EPNcheckboxCnt%>" value="<%=shippingCity%>">
						<input type="hidden" name="shippingCountryCode<%=EPNcheckboxCnt%>" value="<%=shippingCountryCode%>">
						<input type="hidden" name="shippingZip<%=EPNcheckboxCnt%>" value="<%=shippingZip%>">
						<input type="hidden" name="shippingFullName<%=EPNcheckboxCnt%>" value="<%=shippingFullName%>">
						<input type="hidden" name="address2<%=EPNcheckboxCnt%>" value="<%=address2%>">
						<input type="hidden" name="shippingCompany<%=EPNcheckboxCnt%>" value="<%=shippingCompany%>">
						<input type="hidden" name="shippingAddress2<%=EPNcheckboxCnt%>" value="<%=shippingAddress2%>"> 
						<input type="hidden" name="customerCompany<%=EPNcheckboxCnt%>" value="<%=customerCompany%>"> 
						<input type="hidden" name="phone<%=EPNcheckboxCnt%>" value="<%=phone%>"> 
						<input type="hidden" name="email<%=EPNcheckboxCnt%>" value="<%=email%>">
						</td>
						<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
						<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
						<td><div align="center"><%=scCurSign&money(curTotal)%></div></td>
						<td><div align="center"><a href="batchprocessorders.asp?capture=<%=idauthorder%>&GW=pcPay_EPN_Authorize">Remove</a></div></td>
					</tr>
					<% rs.moveNext
				loop
				set rs=nothing
				%>
			<input type="hidden" name="EPNcheckboxCnt" value="<%=EPNcheckboxCnt%>">
			<tr>
				<td nowrap="nowrap">
					<%if EPNcheckboxCnt>"0" then%>
					<input type=hidden name="Check1" value="0">
					<input type="checkbox" name="Check1a" value="1" onClick="javascript:EPNtestcheck1_1()" class="clearBorder"> Check All
					<script type=text/javascript>
					function EPNcheckAll1_1() {
					for (var j = 1; j <= <%=EPNcheckboxCnt%>; j++) {
					box = eval("document.formEPN.checkOrd" + j); 
					if (box.checked == false) box.checked = true;
							}
					}

					function EPNuncheckAll1_1() {
					for (var j = 1; j <= <%=EPNcheckboxCnt%>; j++) {
					box = eval("document.formEPN.checkOrd" + j); 
					if (box.checked == true) box.checked = false;
							 }
					}
					
					function EPNtestcheck1_1() {
					if (document.formEPN.Check1.value=="0") {
					document.formEPN.Check1.value="1";
					EPNcheckAll1_1();
							}
					else
							{
					document.formEPN.Check1.value="0";
					EPNuncheckAll1_1();
							}
					}
					</script>
					<%end if%>				
			</td>
			<td nowrap="nowrap">
					<%if EPNcheckboxCnt>"0" then%>
						<input type=hidden name="Check2" value="1">
						<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:EPNtestcheck1_2()"  class="clearBorder"> Check All
						<script type=text/javascript>
						function EPNcheckAll1_2() {
						for (var j = 1; j <= <%=EPNcheckboxCnt%>; j++) {
						box = eval("document.formEPN.checkEmail" + j); 
						if (box.checked == false) box.checked = true;
								}
						}

						function EPNuncheckAll1_2() {
						for (var j = 1; j <= <%=EPNcheckboxCnt%>; j++) {
						box = eval("document.formEPN.checkEmail" + j); 
						if (box.checked == true) box.checked = false;
								 }
						}
						
						function EPNtestcheck1_2() {
						if (document.formEPN.Check2.value=="0") {
						document.formEPN.Check2.value="1";
						EPNcheckAll1_2();
								}
						else
								{
						document.formEPN.Check2.value="0";
						EPNuncheckAll1_2();
								}
						}
						</script>
					<%end if%>				
				</td>
				<td colspan="6" class="pcCPspacer"></td>
			</tr>
			<% if noAuthRecEPN=0  then %>
			<tr>
				<td colspan="8">
					<input type="submit" name="AuthSubmit" value="Process Selected EPN Orders" class="btn btn-primary">
				</td>
			</tr>
			<tr>
			  <td colspan="8">&nbsp;</td>
			</tr>
			<% end if %>
    </table>
    </form>
   <%
	End If
End If
'////////////////////////////////////////////////////
'// END: EPN
'////////////////////////////////////////////////////


'////////////////////////////////////////////////////
'// START: OTHERS
'////////////////////////////////////////////////////

query="SELECT orders.idOrder, orders.orderDate, orders.orderstatus, orders.total, orders.idCustomer, orders.paymentCode, orders.paymentDetails, orders.comments, orders.admincomments, customers.name, customers.lastName, customers.customerCompany FROM orders INNER JOIN customers ON orders.idCustomer = customers.idCustomer WHERE orderstatus=2 ORDER BY "&GenOrder&" "&GenSort&";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
%>
<form name="form6" method="post" action="batchprocess_pending.asp" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="7"><h2>Pending Orders (No Payment Gateway)</h2></td>
	</tr>
	<tr>
		<th>Process</th>
		<th nowrap="nowrap">Send Email</th>
		<th nowrap="nowrap"><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&GenOrder=orders.orderdate&GenSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?iPageCurrent=<%=iPageCurrent%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&GenOrder=orders.orderdate&GenSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
		<th nowrap="nowrap">Payment Type</th>
		<th nowrap="nowrap"><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&GenOrder=idOrder&GenSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="batchprocessorders.asp?AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&PFPOrder=pfporders.idOrder&PFPSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order ID</th>
		<th>Customer</th>
		<th>Total</th>
	</tr>
    <tr>
        <td colspan="8" class="pcCPspacer"></td>
    </tr>
	<% dim noORDRec
	noORDRec=0
	if rs.eof then
		noORDRec=1 
		%>
		<tr> 
			<td colspan="7"><div class="pcCPmessage">No pending records found</div></td>
		</tr>
	<% end if %>
	<% checkboxCnt=0
	do until rs.eof
		idOrder=rs("idOrder")
		orderDate=rs("orderDate")
		orderStatus=rs("orderstatus")
		total=rs("total")
		idCustomer=rs("idCustomer")
		paymentCode=rs("paymentCode")
		ppaymentDetails=trim(rs("paymentDetails"))
		pcv_custcomments=trim(rs("comments"))
		pcv_admcomments=trim(rs("admincomments"))
		customerName=rs("name") & " " & rs("lastName")
		customerCompany=rs("customerCompany")
			if trim(customerCompany)<>"" then
				customerInfo=customerName & " (" & customerCompany & ")"
				else
				customerInfo=customerName
			end if
		pcArrayPayment = split(ppaymentDetails,"||")
		PaymentType=pcArrayPayment(0)
		if paymentCode="Payeezy" OR paymentCode="Authorize" OR paymentCode="PFPRO" OR paymentCode="Netbill" OR paymentCode="USAePay" OR paymentCode="Google" OR paymentCode="PayPalWP" OR paymentCode="PayPalExp" OR paymentCode="EIG" OR paymentCode="eProcessingNetwork" OR paymentCode="PayPalAdvanced" OR paymentCode="PFLink" then
		else
		checkboxCnt=checkboxCnt+1
		%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td>
				<div align="center"><input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" class="clearBorder"></div>
			</td>
			<td>
				<div align="center"><input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder"></div>
			</td>
			<td><%=ShowDateFrmt(orderDate)%></td>
			<td>
				<input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
				<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=idOrder%>">
				<input type="hidden" name="amt<%=checkboxCnt%>" value="<%=total%>">
				<%=PaymentType%>
			</td>
			<td align="center"><a href="Orddetails.asp?id=<%=idOrder%>"><%=int(idOrder)+scpre%></a><%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=idOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%></td>
			<td><a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank"><%=customerInfo%></a></td>
			<td><div align="center"><%=scCurSign&money(total)%></div></td>
		</tr>
		<% end if
		rs.moveNext
	loop
	set rs=nothing
	
	if noORDRec=0 AND checkboxCnt=0 then
		noORDRec=1 
		%>
		<tr> 
			<td colspan="7"><div class="pcCPmessage">No pending records found</div></td>
		</tr>
	<% end if %>
	<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
	<tr>
				<td nowrap="nowrap">
				<%if checkboxCnt>"0" then%>
				<input type=hidden name="Check1" value="0">
				<input type="checkbox" name="Check1a" value="1" onClick="javascript:testcheck6_1()" class="clearBorder"> Check All 
				<script type=text/javascript>
				function checkAll6_1() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form6.checkOrd" + j); 
				if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAll6_1() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form6.checkOrd" + j); 
				if (box.checked == true) box.checked = false;
				   }
				}
				
				function testcheck6_1() {
				if (document.form6.Check1.value=="0") {
				document.form6.Check1.value="1";
				checkAll6_1();
					}
				else
					{
				document.form6.Check1.value="0";
				uncheckAll6_1();
					}
				}
				</script>
				<%end if%>
				</td>
				<td nowrap="nowrap">
				<%if checkboxCnt>"0" then%>
				<input type=hidden name="Check2" value="1">
				<input type="checkbox" name="Check2a" checked value="1" onClick="javascript:testcheck6_2()" class="clearBorder"> Check All
				<script type=text/javascript>
				function checkAll6_2() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form6.checkEmail" + j); 
				if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAll6_2() {
				for (var j = 1; j <= <%=checkboxCnt%>; j++) {
				box = eval("document.form6.checkEmail" + j); 
				if (box.checked == true) box.checked = false;
				   }
				}
				
				function testcheck6_2() {
				if (document.form6.Check2.value=="0") {
				document.form6.Check2.value="1";
				checkAll6_2();
					}
				else
					{
				document.form6.Check2.value="0";
				uncheckAll6_2();
					}
				}
				</script>
				<%end if%>
				</td>
				<td colspan="5">&nbsp;</td>
		</tr>
	<% if noORDRec=0  then %>
		<tr>
			<td colspan="7" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="7"><input type="submit" name="PendingSubmit" value="Process Selected Orders" class="btn btn-primary"></td>
		</tr>
		<tr>
			<td colspan="7">&nbsp;</td>
		 </tr>
	<% end if %>
</table>
</form>
<%
'////////////////////////////////////////////////////
'// END: OTHERS
'////////////////////////////////////////////////////
%>
<!--#include file="adminfooter.asp"-->
