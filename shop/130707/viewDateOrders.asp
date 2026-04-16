<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
'read start and end dates
Dim strTDateVar, strTDateVar2, DateVar, DateVar2, dateRange
strTDateVar=Request.QueryString("FromDate")
DateVar=strTDateVar
if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
end if
strTDateVar2=Request.QueryString("ToDate")
DateVar2=strTDateVar2
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	if err.number<>0 then
		DateVar=Request.QueryString("FromDate")
		DateVar2=Request.QueryString("ToDate")
	end if
end if

if SQL_Format="1" then
	DateVar=Day(DateVar)&"/"&Month(DateVar)&"/"&Year(DateVar)
	DateVar2=Day(DateVar2)&"/"&Month(DateVar2)&"/"&Year(DateVar2)
else
	DateVar=Month(DateVar)&"/"&Day(DateVar)&"/"&Year(DateVar)
	DateVar2=Month(DateVar2)&"/"&Day(DateVar2)&"/"&Year(DateVar2)
end if

if (DateVar<>"") and IsDate(DateVar) then
	dateRange=" orders.orderDate >='" & DateVar & "' "
else
	dateRange=""
end if

if (DateVar2<>"") and IsDate(DateVar2) then
	dateRange= dateRange & " AND orders.orderDate <='" & DateVar2 & "' "
END IF 
pageTitle="Total sales recorded from " & strTDateVar & " to " & strTDateVar2
pageIcon="pcv4_icon_people.png"
section= "mngAcc"
pcInt_ShowOrderLegend = 1
%>
<%PmAdmin="7*9*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 

Const iPageSize=15
Dim iPageCurrent
if request.querystring("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request.QueryString("iPageCurrent")
end if

'sorting order
Dim strORD
strORD=request("order")
if strORD="" then
	strORD="idorder"
End If

strSort=request("sort")
if strSort="" Then
	strSort="DESC"
End If 

dim shiptemp

		' Get order totals
			TotalOrdered=0
			query="SELECT sum(total-rmaCredit) As TotalAmount, Sum(Total) AS TotalLessRMA FROM orders WHERE ((orderStatus>2 AND orderStatus<5) OR (orderStatus>6 AND orderStatus<=9) OR (orderStatus=10 OR orderStatus=12)) AND " & dateRange &";"
			
			set rstemp=Server.CreateObject("ADODB.Recordset") 
			rstemp.Open query, conntemp
			if err.number <> 0 then
				call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error calculating order total: "&Err.Description) 
			end If
			if not rstemp.eof then
				TotalOrdered=rstemp("TotalAmount")
				TotalLessRMA=rstemp("TotalLessRMA")
			end if
			if isNull(TotalOrdered) then
				TotalOrdered = TotalLessRMA
				if isNull(TotalOrdered) then
					TotalOrdered=0
				end if
			end if
			set rstemp=nothing
			
		' Count total orders except for incomplete
			Dim pcvIntTotalOrders
			query = "SELECT Count(*) AS intTotal FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND " & dateRange
			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			rsTemp.Open query, conntemp
			pcvIntTotalOrders = rsTemp("intTotal")
			set rsTemp = nothing
		
		' Count incomplete orders
			Dim pcvIntIncOrders
			query = "SELECT Count(*) AS intTotal FROM orders WHERE orderStatus=1 AND " & dateRange
			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			rsTemp.Open query, conntemp
			pcvIntIncOrders = rsTemp("intTotal")
			set rsTemp = nothing

    ' Get shipment Info
			query="SELECT orders.shipmentDetails FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND" & dateRange & " ORDER BY "& strORD &" "& strSort
			set shiptemp=Server.CreateObject("ADODB.Recordset")     
			shiptemp.Open query, conntemp

		' Get orders placed by this customer
			query="SELECT orders.idorder, orderDate, total, orderstatus,orders.pcOrd_PaymentStatus,orders.comments,orders.admincomments,orders.details,orders.rmaCredit, customers.name, customers.lastName, orders.idcustomer, orders.shipmentDetails FROM orders, customers WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.idcustomer=customers.idcustomer AND " & dateRange & " ORDER BY "& strORD &" "& strSort
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			rstemp.CursorLocation=adUseClient
			rstemp.CacheSize=iPageSize
			rstemp.PageSize=iPageSize
			rstemp.Open query, conntemp
			
			if err.number <> 0 then
				call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
			end If
			
	if rstemp.eof then
		dim showLinks
		showLinks = 1
	else
	
		rstemp.MoveFirst
		' get the max number of pages
		Dim iPageCount
			iPageCount=rstemp.PageCount
			If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
		If iPageCurrent < 1 Then iPageCurrent=1
		
		' set the absolute page
		rstemp.AbsolutePage=iPageCurrent
		Dim count
		Count=0
	end if

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
IF showlinks = 1 THEN '// NO ORDERS
%>
	<table class="pcCPcontent">
		<tr>
			<td>
			<div class="pcCPmessage">There are no orders associated with this customer account.</div>
			<ul>
			<li><a href="modCusta.asp?idcustomer=<%=pIdCustomer%>">View customer details</a></li>
			<li><a href="adminPlaceOrder.asp?idcustomer=<%=pidcustomer%>" target="_blank">Place an order on behalf of this customer</a><br><br></li>
			<li><a href="viewCusta.asp">Look for another customer</a></li>
			<li><a href="viewCustb.asp?mode=ALL">View all customers</a></li>
			<li><a href="javascript: history.go(-1)">Back</a></li>
			</ul>
			</td>
		</tr>
	</table>
<% 
ELSE
%>

	
	<table class="pcCPcontent" style="margin-top: 10px;">
	<tr> 
        <th nowrap align="center">Status</th>
        <th width="2%" nowrap><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=idorder&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=idorder&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order#</th>
        <th width="2%" nowrap><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
        <th nowrap>Customer Name</th>
        <th nowrap>Total</th>
        <th colspan="2" nowrap>Products Ordered</th>
	</tr>
    <tr>
        <td colspan="6" class="pcCPspacer"></td>
    </tr>

	<%
  gTotalshipfees=0
  gTotalhandfees=0
  do while not shiptemp.eof
    pShipmentDetail = shiptemp("shipmentDetails")
    shipping=split(pShipmentDetail,",")
	  if ubound(shipping)>1 then
		  if NOT isNumeric(trim(shipping(2))) then
			  shipfees=0
		  else
			  shipfees=cdbl(trim(shipping(2)))
		  end if	
		  if NOT isNumeric(trim(shipping(2))) then
			  HandFees=0
		  else
			  HandFees=cdbl(trim(shipping(3)))
		  end if
	  else
		  shipfees=0
		  Handfees=0
	  end if
	  gTotalshipfees=gTotalshipfees + shipFees
	  gTotalhandfees=gTotalhandfees + HandFees
    shiptemp.movenext
  loop
  set shiptemp = nothing

	do while not rstemp.eof And Count < rstemp.PageSize
	
		pidorder=rstemp("idorder")
		porderDate=rstemp("orderDate")
    pCustomerName = rstemp("name")&" "&rstemp("lastName")
    pCustomerID = rstemp("idcustomer")

		porderDate=ShowDateFrmt(porderDate)
		ptotal=rstemp("total")
		prmaCredit=rstemp("rmaCredit")
			'// Calculate total adjusted for credits
			if trim(prmaCredit)="" or IsNull(prmaCredit) then
				prmaCredit=0
			end if
			pTotalAdj=pTotal-prmaCredit
		porderstatus=rstemp("orderStatus")
		'Start SDBA
		pcv_PaymentStatus=rstemp("pcOrd_PaymentStatus")
		if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
			pcv_PaymentStatus=0
		end if
		'End SDBA
		pcv_custcomments=trim(rstemp("comments"))
		pcv_admcomments=trim(rstemp("admincomments"))
		pcv_details=trim(rstemp("details"))
			if len(pcv_details)>180 then
				pcv_details=left(pcv_details,180) & "..."
			end if
		pcv_details=replace(pcv_details," ||",""&scCurSign&"")
	%>
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
		<td align="center" valign="top" width="5%"><!--#include file="inc_orderStatusIcons.asp"--></td>
    	<td align="center" valign="top" width="5%"><% if porderstatus="1" then %><a href="OrdDetailsIncomplete.asp?id=<%=pidOrder%>"><% else %><a href="Orddetails.asp?id=<%=pidOrder%>"><% end if %><strong><%response.write (scpre+int(pIdOrder))%></strong></a></td>
		<td valign="top" width="5%"><%response.write pOrderDate%></td>
    <td valign="top" align="center" width="5%"><a href="viewCustOrders.asp?idcustomer=<%=pCustomerID%>" ><%response.write pCustomerName%></a></td>
		<td valign="top" width="5%"><%response.write(scCurSign & money(ptotal))%></td>
    	<td valign="top" width="70%">
    
			<%
                query="SELECT ProductsOrdered.idProduct, ProductsOrdered.idOrder, products.description, products.sku, products.idProduct, orders.idOrder FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idOrder=" &pIdOrder
                set rs=Server.CreateObject("ADODB.Recordset")
                set rs=conntemp.execute(query)
            
                While Not rs.EOF
                pIdProduct=rs("idProduct") 
                pSku=rs("sku")
                pDescription=rs("description")
                %>
                <div style="margin-bottom: 3px;"><%=psku%> - <%=pDescription %></div>
                <%
                rs.MoveNext
                Wend
                set rs = nothing
            %>

        </td>
        <td align="right" nowrap valign="top" width="10%">
            <% if porderstatus="1" then %>
             <a href="OrdDetailsIncomplete.asp?id=<%=pidOrder%>">Review</a>
            <% else %>
             <a href="Orddetails.asp?id=<%=pidOrder%>"><img src="images/pcIconNext.jpg" width="12" height="12" alt="View and Process"></a>&nbsp;<a href="OrdInvoice.asp?id=<%response.write pIdOrder%>" target="_blank"><img src="images/print_xsmall.gif" alt="Printer Friendly Version" border="0"></a>
            <% end if %>
            <%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=pidOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%>
        </td>
	</tr>
														
	<%
	 rstemp.movenext
	 Count=Count + 1
	Loop
	set rstemp=nothing
	%>
    <tr>
        <td colspan="6" class="pcCPspacer"></td>
    </tr>
	<tr>
	<td colspan="6">
	<form method="post" action="viewDateOrders.asp" name="" class="pcForms">
	<%=("Page "& iPageCurrent & " of "& iPageCount)%>
    <br />
	<%
	'Display Next / Prev buttons
	if iPageCurrent > 1 then
	'We are not at the beginning, show the prev button %>
		<a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&status=<%=pOrderStatus%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
	<% end If
	If iPageCount <> 1 then
		For I=1 To iPageCount
			If I=iPageCurrent Then %>
				<%=I%> 
			<% Else %>
				<a href="viewDateOrders.asp?FromDate=<%=strTDateVar%>&ToDate=<%=strTDateVar2%>&status=<%=pOrderStatus%>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
			<% End If %>
		<% Next %>
	<% end if %>
	<% if CInt(iPageCurrent) <> CInt(iPageCount) then
	'We are not at the end, show a next link %>
		<a href="viewDateOrders.asp?FromDate=<%=strTDateVar%>&ToDate=<%=strTDateVar2%>&status=<%=pOrderStatus%>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
	<% end If %>
	</form>
	</td>
	</tr>
</table>
	<table class="pcCPcontent">
		<tr>
			<td><strong>Number of orders:</strong>&nbsp;&nbsp;<%=pcvIntTotalOrders%></td>
			<td>Totals:</td>
      <td><strong><%=scCurSign & money(TotalOrdered)%></strong>&nbsp;&nbsp;|&nbsp;Shipping: <strong><%=scCurSign & money(gTotalshipfees)%></strong></td>
		</tr>
</table>
<% 
END IF
%><!--#include file="Adminfooter.asp"-->
