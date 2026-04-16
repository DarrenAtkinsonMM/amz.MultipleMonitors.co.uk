<%'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Sales Reports"
Response.Buffer = False
Server.ScriptTimeout = 8000 %>

<% Section="genRpts" %>
<%PmAdmin=10%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="pcCharts.asp"-->
<%
recordsToShow=Request.QueryString("resultCnt")
srcVar=Request.QueryString("src")
FromDate=Request.QueryString("FromDate")
ToDate=Request.QueryString("ToDate")

query1=""

if (FromDate<>"") and (not (IsDate(FromDate))) then
	FromDate=Date()
end if
if (ToDate<>"") and (not (IsDate(ToDate))) then
	ToDate=Date()
end if

err.clear

Dim strTDateVar, strTDateVar2, DateVar, DateVar2
strTDateVar=FromDate

if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	FromDate=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
end if

strTDateVar2=ToDate
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	ToDate=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	if err.number<>0 then
		FromDate=(day(FromDate)&"/"&month(FromDate)&"/"&year(FromDate))
		ToDate=(day(ToDate)&"/"&month(ToDate)&"/"&year(ToDate))
	end if
end if

err.clear

tmpDate=request("basedon")
tmpD=""
tmpD1=""
tmpD2=""
Select case tmpDate
	Case "2": tmpD="orders.processDate"
		tmpD1="processDate"
		tmpD2="Processed On"
	Case "3": tmpD="pcPackageInfo.pcPackageInfo_ShippedDate"
		tmpD1="pcPackageInfo_ShippedDate"
		tmpD2="Shipped On"
	Case Else: tmpD="orders.orderDate"
		tmpD1="processDate"
		tmpD2="Processed On"
End Select

TempSQL1=""
TempSQL2=""

if SQL_Format="1" then
	FromDate=Day(FromDate)&"/"&Month(FromDate)&"/"&Year(FromDate)
	ToDate=Day(ToDate)&"/"&Month(ToDate)&"/"&Year(ToDate)
else
	FromDate=Month(FromDate)&"/"&Day(FromDate)&"/"&Year(FromDate)
	ToDate=Month(ToDate)&"/"&Day(ToDate)&"/"&Year(ToDate)
end if

if (FromDate<>"") and (IsDate(FromDate)) then
	TempSQL1 = " AND " & tmpD & " >='" & FromDate & "'"
end if

if (ToDate<>"") and (IsDate(ToDate)) then
	TempSQL2 = " AND " & tmpD & " <='" & ToDate & "'"
end if

TempSpecial=""
if tmpDate="3" then
	tmpStr1=""
	if TempSQL1<>"" then
		tmpStr1=replace(TempSQL1,tmpD,"orders.shipDate")
		tmpStr1=replace(tmpStr1," AND ","")
	end if
	tmpStr2=""
	if TempSQL2<>"" then
		tmpStr2=replace(TempSQL2,tmpD,"orders.shipDate")
		tmpStr2=replace(tmpStr2," AND ","")
	end if
	tmpD="orders.processDate"
	
	TempSpecial=" AND "
	if tmpStr1 & tmpStr2 <> "" then
		TempSpecial=TempSpecial & " ((" & tmpStr1
		if tmpStr2<>"" then
			if tmpStr1<>"" then
				TempSpecial=TempSpecial & " AND "
			end if
			TempSpecial=TempSpecial & tmpStr2 & ") OR "
		end if
	end if
	
	TempSpecial=TempSpecial & " (orders.idorder IN (SELECT DISTINCT idorder FROM pcPackageInfo"
	if TempSQL1<>"" or TempSQL2<>"" then
		TempSpecial=TempSpecial & " WHERE pcPackageInfo_ID>0 " & TempSQL1 & TempSQL2
	end if
	TempSQL1=""
	TempSQL2=""
	TempSpecial=TempSpecial & "))"
	if tmpStr1 & tmpStr2 <> "" then
		TempSpecial=TempSpecial & ")"
	end if
end if

query1=query1 & TempSQL1 & TempSQL2 '& TempSpecial

'// Top Viewed Products
if srcVar="2" then
	query="SELECT TOP " & recordsToShow & " IDProduct, Description, Visits FROM products WHERE products.visits >0 ORDER BY products.visits DESC;"
end if

'// Top 'Wish List' Products
if srcVar="4" then
	query="SELECT TOP " & recordsToShow & " Products.IDProduct, Products.Description, COUNT(*) AS TotalCount FROM Products INNER JOIN WishList ON Products.IDProduct=WishList.IDProduct GROUP BY Products.IDProduct, Products.Description ORDER BY TotalCount DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	pcv_havelist=0
	intCount=-1
	if not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
		pcv_havelist=1
	end if	
	set rs=nothing
	
End if

'// Top Selling Products
if srcVar="1" then
	query="UPDATE products SET Sales=0" 
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	
	query = "UPDATE Products SET Products.sales=(SELECT SUM(ProductsOrdered.Quantity) FROM ProductsOrdered " 
	query = query & "INNER JOIN orders on Orders.IDOrder = ProductsOrdered.IDOrder "
	query = query & "WHERE ProductsOrdered.idProduct=Products.idProduct AND Orders.IDOrder=ProductsOrdered.IDOrder AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12))" & query1 & ")"

	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
	
	query="SELECT TOP " & recordsToShow & " IDProduct, description, sales FROM Products WHERE products.sales >0 ORDER BY products.sales DESC;"
end if 

'//Top Customers
if srcVar="3" then
	query="UPDATE Customers SET TotalOrders=0, TotalSales=0"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	
	query="UPDATE Customers SET Customers.TotalOrders=(SELECT count(*) FROM Orders WHERE Orders.idCustomer=Customers.idCustomer AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12))" & TempSpecial & query1 &"), Customers.TotalSales=(SELECT sum(total) FROM Orders WHERE Orders.idCustomer=Customers.idCustomer AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12))" & TempSpecial & query1 &")"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rstemp=nothing
	
	query="SELECT TOP " & recordsToShow & " idcustomer, name, lastname, customerCompany, Totalorders, Totalsales FROM customers WHERE TotalOrders>0 ORDER BY TotalSales DESC"
	
end if

' Our Recordset Object
if srcVar<>"4" then
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
end if

Dim rcount, i, x
Dim tmpChartName,tmpLine1
tmpChartName=""
tmpLine1=""
%>
<table class="pcCPcontent" width="100%">
<tr valign="top">
<td width="100%">
<table class="pcCPcontent" style="width:auto;">
	<tr>
		<td colspan="5" nowrap>
		<% if srcVar<>"2" then %>
			<%if FromDate<>"" then%>
				From:&nbsp;<%=FromDate%>
			<%end if%>
			<%if ToDate<>"" then%>
				&nbsp;To:&nbsp;<%=ToDate%>
			<%end if%>
		<%end if%>
		</td>
	</tr>
	<% if srcVar="1" then
	tmpChartName="Top Selling Products" %>
		<tr> 
			<th colspan="2" nowrap>Top Selling Products</th>
			<th nowrap colspan="2">Amount Sold</th>
		</tr>
	<% end if %>

	<% if srcVar="2" then
	tmpChartName="Most Viewed Products" %>
		<tr> 
			<th colspan="2" nowrap>Most Viewed Products</th>
			<th nowrap colspan="2">Total Views</th>
        </tr>
	<% end if %>

	<% if srcVar="4" then
	tmpChartName="Top ""Wish List"" Products" %>
		<tr> 
			<th colspan="2" nowrap>Top "Wish List" Products</th>
			<th nowrap colspan="2">Number of Wish Lists</th>
		</tr>
	<% end if %>

	<% if srcVar="3" then
	tmpChartName="Best Customers" %>
		<tr> 
			<th colspan="2" nowrap>Best Customers</th>
			<th nowrap>Number of Orders</th>
			<th colspan="2" nowrap>Orders Total</th>
		</tr>
	<% end if %>
	
	<% if srcVar="3" then %>
		<tr> 
			<td colspan="5" class="pcCPspacer"></td>
		</tr>
	<%else%>
		<tr> 
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
	<%end if%>

	<% IF srcVar="4" THEN
		if pcv_havelist<>"1" then %>
			<tr> 	
				<td colspan="5"> 
					<div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20">No Results Found</div>
				</td>
			</tr>
		<% else
			rCount=0
			For i=0 to intCount
				pIDProduct=pcArr(0,i)
				PDesc=pcArr(1,i)
				TotalCount=pcArr(2,i)
				rCount=rCount+1
				if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & PDesc & "'," & Clng(TotalCount) & "]"
				%>
				<tr bgcolor="<%= strCol %>"> 		
					<td width="6%"><%Response.Write rcount %></td>
					<td width="44%"><a href="FindProductType.asp?id=<%=pIDProduct%>" target="_blank"><%=PDesc%></a></td>
					<td colspan="2"><%=TotalCount%></td>
				</tr>
		
			<%Next
		end if
	ELSE
		If rs.EOF Then %>
			<tr> 	
                <td colspan="5" height="22"> 
                        <div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20">No Results Found</div>
                </td>
			</tr>
		<%
		Else 
			' Showing relevant records
			rcount=0
			pcArr=rs.getRows()
			intCount=ubound(pcArr,2)
			
			if intCount>=clng(recordsToShow) then
				intCount=clng(recordsToShow)-1
			end if
			
			For i=0 to intCount
				rcount=rcount+1
				IF srcVar="1" then %>
						
                    <tr bgcolor="<%= strCol %>"> 
                        <td width="6%"><%Response.Write rcount %></td>
                        <td width="44%"><a href="FindProductType.asp?id=<%=pcArr(0,i)%>" target="_blank"><%=pcArr(1,i)%></a></td>
                        <td colspan="2"><%=pcArr(2,i)%></td>
                    </tr>
					

				<%
				if tmpline1<>"" then
					tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & pcArr(1,i) & "'," & Clng(pcArr(2,i)) & "]"
				END IF

				IF srcVar="2" then %>
		
					<tr bgcolor="<%= strCol %>"> 		
						<td width="6%"><%Response.Write rcount %></td>
						<td width="44%"><a href="FindProductType.asp?id=<%=pcArr(0,i)%>" target="_blank"><%=pcArr(1,i)%></a></td>
						<td colspan="2"><%=pcArr(2,i)%></td>
					</tr>
	
				<%
				if tmpline1<>"" then
					tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & pcArr(1,i) & "'," & Clng(pcArr(2,i)) & "]"
				END IF

				IF srcVar="3" THEN %>
							
					<tr bgcolor="<%= strCol %>"> 			
						<td width="6%"><%Response.Write rcount %></td>
						<td width="44%" nowrap><a href="modCusta.asp?idcustomer=<%=pcArr(0,i)%>" target="_blank"><%=pcArr(1,i)%>&nbsp;<%=pcArr(2,i)%></a>
                        <%
						pcvCustomerCompany = pcArr(3,i)
						if pcvCustomerCompany<>"" and not isNull(pcvCustomerCompany) then
						%>
						&nbsp;(<%=pcvCustomerCompany%>)
						<%
						end if
						%>
						</td>
						<td align="center"><%=pcArr(4,i)%></td>
						<td colspan="2"><%=scCurSign%> <%=money(pcArr(5,i))%></td>
					</tr>
						
				<%
				if tmpline1<>"" then
					tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & pcArr(1,i) & " " & pcArr(2,i) & "'," & Clng(pcArr(5,i)) & "]"
				END IF
			Next
			
		End If 'Have Records
		
		If srcVar="1" then
			
				'reset sales
				query="UPDATE Products SET Sales=0" 
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				set rstemp=nothing

				query="UPDATE Products SET Products.sales=(SELECT SUM(ProductsOrdered.Quantity) FROM ProductsOrdered " 
				query = query & "INNER JOIN orders on Orders.IDOrder = ProductsOrdered.IDOrder "
				query = query & "WHERE ProductsOrdered.idProduct=Products.idProduct AND Orders.IDOrder=ProductsOrdered.IDOrder AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)))"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				set rstemp=nothing
				
		End If
	END IF 'Top Wish List Products
	%>
    <tr>
        <td colspan="4">&nbsp;</td>
    </tr>
</table>
</td>
</tr>
<tr>
<td width="100%">
<%if tmpline1<>"" then%>
	<div id="chartTop" style="height:330px; "></div>
	<script type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
		
	<script type=text/javascript>$pc(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $pc.jqplot('chartTop', [line1], {
    	title: '<%=tmpChartName%>',
    	seriesDefaults:{renderer:$pc.jqplot.PieRenderer, rendererOptions:{showDataLabels: true,sliceMargin:0}},
    	legend:{show:true}
		});});
	</script>
<%end if%>
</td>
</tr>
</table>
<%  ' Done. Now release Objects
set rs=nothing

%>
<!--#include file="AdminFooter.asp"-->