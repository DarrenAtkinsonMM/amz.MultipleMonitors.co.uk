<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Updating the prices of additional charges assigned to configurable products" %>
<% section="services" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<% dim iPageCurrent
if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request("iPageCurrent")
end If

iPageSize=50

function FixPrice(tmpprice)
	Dim tmp1,tmp2
	tmp1=tmpprice
	tmp2=tmp1
	if cdbl(fix(tmp1))<>cdbl(tmp2) then
		FixPrice=money(tmp1)
	else
		FixPrice=tmp1
	end if
end function

Function RemoveCurSign(tmpPrice)
	Dim tmp1,tmp2
	tmp1=tmpPrice
	if tmp1<>"" then
		if Instr(tmp1,",")>0 then
			if Instr(tmp1,".")<Instr(tmp1,",") then
				tmp1=replace(tmp1,".","")
				tmp1=replace(tmp1,",",".")
			end if
		end if
	
		if Instr(tmp1,".")>0 then
			if Instr(tmp1,",")<Instr(tmp1,".") then
				tmp1=replace(tmp1,",","")
			end if
		end if
	end if
	
	RemoveCurSign=tmp1

End Function

'sorting order
Dim strORD
strORD=request("order")
if strORD="" then
	strORD="products.description"
End If

dim strSort
strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If 

Dim i, query1, query2

if request("action")="update" then
	count=request("count")

	for i=1 to count
		if request("C" & i)="1" then
			query1=""
			query2=""

			OPrice=RemoveCurSign(request("Price" & i))
			if OPrice<>"" then
				query1=query1 & "price=" & OPrice
				query2=query2 & "price=" & OPrice
			end if

			WPrice=RemoveCurSign(request("WPrice" & i))
			if WPrice<>"" then
				if query1<>"" then
					query1=query1 & ",Wprice=" & WPrice
					query2=query2 & ",btoBprice=" & WPrice
				else
					query1=query1 & "Wprice=0"
					query2=query2 & "btoBprice=0"
				end if
			end if

			if query1<>"" then
				query="UPDATE configSpec_Charges SET " & query1 & " WHERE configproduct="& request("ID" & i)
				set rstemp=server.CreateObject("ADODB.RecordSet")
				Set rstemp=conntemp.execute(query)
				msg="Additional Charges have been updated successfully!<br /><br />Note that you will not see those changes in the prices listed below (in case you changed any of them). This page shows the original prices associated with these <em>Additional Charges</em>. The prices you entered have been saved to the affected product configurations."
				msgtype=1
				set rstemp=nothing
			end if
			
			pcv_PriceCats=request("PricingCatCount")
			if pcv_PriceCats>"0" then
				pcv_tmpPrd=request("ID" & i)
				For j=1 to pcv_PriceCats
					pcv_Cat=request("Cat_" & j & "_" & i)
					pcv_Price=RemoveCurSign(request("Price_" & j & "_" & i))
					query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory=" & pcv_Cat & " AND idproduct=" & pcv_tmpPrd & ";"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					if not rs.eof then
						query="UPDATE pcCC_Pricing SET pcCC_Price=" & pcv_Price & " WHERE idcustomerCategory=" & pcv_Cat & " AND idproduct=" & pcv_tmpPrd & ";"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
					else
						query="INSERT INTO pcCC_Pricing (idcustomerCategory,idproduct,pcCC_Price) values (" & pcv_Cat & "," & pcv_tmpPrd & "," & pcv_Price & ");"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
					end if
					set rs=nothing
					
					call updPrdEditedDate(pcv_tmpPrd)

					query="UPDATE pcCC_BTO_Pricing SET pcCC_BTO_Price=" & pcv_Price & " WHERE idcustomerCategory=" & pcv_Cat & " AND idBTOItem=" & pcv_tmpPrd & ";"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					set rs=nothing
				Next
			end if
		end if	
	next 
end if
%>
<!--#include file="AdminHeader.asp"-->
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="POST" name="checkboxform" action="updateBTOACprices.asp?action=update&iPageCurrent=<%=request("iPageCurrent")%>&order=<%=request("order")%>&sort=<%=request("sort")%>" onSubmit="return Form1_Validator(this)" class="pcForms">

		<% 
		'// Count Pricing Categories and load array of names
		pcv_PriceCats=0
		query="SELECT idCustomerCategory,pcCC_Name FROM pcCustomerCategories;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		intCount=-1
		if not rs.eof then
			pcv_PCArr=rs.getRows()
			intCount=ubound(pcv_PCArr,2)
			pcv_PriceCats=intCount+1
			set rs=nothing
		end if
		%>

			
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" align="right">Prices:&nbsp;</td>
			<td colspan="2" align="center" style="border-left: 1px solid #CCC;">Original regular &amp; wholesale</td>
			<td colspan="2" align="center" style="border-left: 1px solid #CCC;">Assigned regular &amp; wholesale</td>
			<% if pcv_PriceCats>0 then %>
			<td colspan="<%=pcv_PriceCats%>" style="border-left: 1px solid #CCC;">Pricing Categories</td>
			<% end if %>
		</tr>
		
		<tr> 
			<th nowrap><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;SKU</th>
			<th nowrap><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Item Name</th>
			<th nowrap><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=price&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=price&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Reg.</th>
			<th nowrap><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=btoBprice&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent%>&order=btoBprice&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;W.</th>
			<th nowrap>Reg.</th>
			<th nowrap>W.</th>
			<%
				For i=0 to intCount%>
					<th style="font-size: 10px;"><%=pcv_PCArr(1,i)%></th>
			<%
				Next
			%>
			<th nowrap>&nbsp;</th>
		</tr>
		<tr> 
			<td colspan="<%=Clng(pcv_PriceCats)+7%>" class="pcCPspacer"></td>
		</tr>
		
		<%
		if request("view")<>"all" then
			query="SELECT products.idproduct, products.sku, products.description, products.price, products.btoBprice FROM products WHERE idproduct IN (SELECT DISTINCT idproduct FROM (configSpec_Charges INNER JOIN products ON configSpec_Charges.configProduct=products.idProduct) WHERE products.configOnly=1 AND products.serviceSpec=0 AND products.removed=0) ORDER BY " & strORD & " " & strSort
			Set rsInv=Server.CreateObject("ADODB.Recordset")
			rsInv.CacheSize=iPageSize
			rsInv.PageSize=iPageSize
			rsInv.Open query, connTemp, adOpenStatic, adLockReadOnly
		else
			query="SELECT products.idproduct, products.sku, products.description, products.price, products.btoBprice FROM products WHERE idproduct IN (SELECT DISTINCT idproduct FROM (configSpec_Charges INNER JOIN products ON configSpec_Charges.configProduct=products.idProduct) WHERE products.configOnly=1 AND products.serviceSpec=0 AND products.removed=0) ORDER BY " & strORD & " " & strSort
			Set rsInv=Server.CreateObject("ADODB.Recordset")
			set rsInv=connTemp.execute(query)
		end if													

		If rsInv.eof Then %>
			<tr> 
				<td colspan="<%=Clng(pcv_PriceCats)+7%>">This store currently does not have additional charges assigned to Configurable Products.</td>
			</tr>
		<% Else
		
			if request("view")<>"all" then
				Dim iPageCount
				iPageCount=rsInv.PageCount
				If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
				If iPageCurrent < 1 Then iPageCurrent=1
			
				' set the absolute page
				rsInv.AbsolutePage=iPageCurrent
				pcPrdArr=rsInv.getRows(iPageSize)
				intPrdCount=ubound(pcPrdArr,2)
			else
				iPageCount=1
				iPageCurrent=1
				pcPrdArr=rsInv.getRows()
				intPrdCount=ubound(pcPrdArr,2)
			end if
			set rsInv=nothing
			
			Count=0
			
			For k=0 to intPrdCount
				count=count + 1
				pcIdProduct=pcPrdArr(0,k)
				strSKU=pcPrdArr(1,k)
				strDescription=pcPrdArr(2,k)
				dbPrice=pcPrdArr(3,k)
				dbBtoBPrice=pcPrdArr(4,k)
				temp_IDProduct=pcIdProduct
				%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td><%=strSKU%><input type="hidden" name="ID<%=count%>" value="<%=pcIdProduct%>"></td>
					<td><a href="FindProductType.asp?id=<%=pcIdProduct%>"><%=strDescription%></a></td>
					<td><%=scCurSign & money(dbPrice)%></td>
					<td><%=scCurSign & money(dbBtoBPrice)%></td>                          
					<td><input type="text" name="Price<%=count%>" size="5" value="<%=FixPrice(dbPrice)%>"></td>
					<td><input type="text" name="WPrice<%=count%>" size="5" value="<%=FixPrice(dbBtoBPrice)%>"></td>
						<% 
						For i=0 to pcv_PriceCats-1
							pcv_CatPrice=0
	
							query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory=" & pcv_PCArr(0,i) & " AND idproduct=" & temp_IDProduct & ";"
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)
							if not rs.eof then
								pcv_CatPrice=rs("pcCC_Price")
								pcv_CatPrice=pcf_Round(pcv_CatPrice, 2)
							else
								query="SELECT  idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory=" & pcv_PCArr(0,i) & ";"
								set rs=server.CreateObject("ADODB.RecordSet")
								SET rs=conntemp.execute(query)
								intIdcustomerCategory=rs("idcustomerCategory")
								strpcCC_Name=rs("pcCC_Name")
								strpcCC_CategoryType=rs("pcCC_CategoryType")
								intpcCC_ATBPercentage=rs("pcCC_ATB_Percentage")
								intpcCC_ATB_Off=rs("pcCC_ATB_Off")
								if intpcCC_ATB_Off="Retail" then
									intpcCC_ATBPercentOff=0
								else
									intpcCC_ATBPercentOff=1
								end if
								
								if (dbBtoBPrice>"0") then
									tempPrice=dbBtoBPrice
								else
									tempPrice=dbPrice
								end if

								' Calculate the "across the board" price
								if strpcCC_CategoryType="ATB" then
									if intpcCC_ATBPercentOff=0 then
										ATBPrice=dbPrice-(pcf_Round(dbPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
									else
										ATBPrice=tempPrice-(pcf_Round(tempPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
									end if
									pcv_CatPrice=ATBPrice
								end if
							end if
							set rs=nothing%>
							<td>
								<input type="text" name="Price_<%=i+1%>_<%=count%>" size="5" value="<%=FixPrice(pcv_CatPrice)%>">
								<input type="hidden" name="Cat_<%=i+1%>_<%=count%>" value="<%=pcv_PCArr(0,i)%>">
							</td>
						<%Next%>
					<td><input type="checkbox" name="C<%=count%>" value="1" class="clearBorder"></td>
				</tr>
			<%Next
			set rsInv=nothing 
			
			%>

			<tr>
				<td colspan="<%=Clng(pcv_PriceCats)+7%>" class="pcCPspacer"><input type=hidden name="PricingCatCount" value="<%=pcv_PriceCats%>"></td>
			</tr>
			<tr>
	<td colspan="<%=Clng(pcv_PriceCats)+7%>" align="right" class="cpLinksList">
			<input type=hidden name=count value=<%=count%>>
				<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
				<%if request("view")<>"all" then%>
				&nbsp;|&nbsp;<a href="updateBTOACprices.asp?view=all">Show All Items</a>
				<%end if%>
				</td>
			</tr>
			<tr>
				<td colspan="<%=Clng(pcv_PriceCats)+7%>" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="<%=Clng(pcv_PriceCats)+7%>">
				<input type="submit" name="submit" value="Update Assigned Additional Charges Prices" class="btn btn-primary">&nbsp;
				<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
				</td>
			</tr>
		<%End If%>
	</table>
	                  

	<% If iPageCount>1 Then %>
		<br>
		<table class="pcCPcontent">                   
			<tr> 
				<td><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%></td>
			</tr>
			<tr>                   
				<td> 
				<%' display Next / Prev buttons
				if iPageCurrent > 1 then %>
					<a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
				<% end If
				For I=1 To iPageCount
					If Cint(I)=Cint(iPageCurrent) Then %>
						<b><%=I%></b> 
					<% Else %>
						<a href="updateBTOACprices.asp?iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"> 
					<%=I%></a> 
					<% End If %>
				<% Next %>
				<% if CInt(iPageCurrent) < CInt(iPageCount) then %>
					<a href="updateBTOACprices.asp?iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
				<% end If %>
				</td>
			</tr>
		</table>
	<% End If %>
</form>
<% if count>0 then%>
	<script type=text/javascript>
	function checkAll() {
	for (var j = 1; j <= <%=count%>; j++) {
	box = eval("document.checkboxform.C" + j); 
	if (box.checked == false) box.checked = true;
		 }
	}
	
	function uncheckAll() {
	for (var j = 1; j <= <%=count%>; j++) {
	box = eval("document.checkboxform.C" + j); 
	if (box.checked == true) box.checked = false;
		 }
	}
		
	function isDigit(s)
	{
	var test=""+s;
	if(test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
			{
			return(true) ;
			}
			return(false);
		}
		
	function allDigit(s)
		{
			var test=""+s ;
			for (var k=0; k <test.length; k++)
			{
				var c=test.substring(k,k+1);
				if (isDigit(c)==false)
				{
					return (false);
				}
			}
			return (true);
		}
	
	function Form1_Validator(theForm)
	{
		if (confirm('You are about to update prices within all configurable products that include the additional charges that you have checked on this page. The current price and wholesale price will be updated with the new prices listed above. Would you like to continue?'))
		{
		}
		else
		{
		return(false);
		}
		
		for (var j = 1; j <= <%=count%>; j++) 
		{
		box = eval("document.checkboxform.C" + j); 
		if (box.checked == true)
		{
		qtt= eval("document.checkboxform.Price" + j);
			if (qtt.value != "")
			{
				if (allDigit(qtt.value) == false)
				{
					alert("Please enter a numeric value for this Field.");
					qtt.focus();
					return (false);
					}
				}
			qtt1= eval("document.checkboxform.WPrice" + j);
			if (qtt1.value != "")
			{
				if (allDigit(qtt1.value) == false)
				{
					alert("Please enter a numeric value for this Field.");
					qtt1.focus();
					return (false);
					}
				}
	
		}
		}
	
	return (true);
	}
	</script>
<% end if 'Count > 0%>
<!--#include file="AdminFooter.asp"-->
