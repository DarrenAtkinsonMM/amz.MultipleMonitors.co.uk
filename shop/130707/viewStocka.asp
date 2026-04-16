<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="View and Update Inventory Levels for Multiple Products"
pageIcon="pcv4_icon_inventoryAdded.gif"
section="products" 
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<%
Dim rsOrd, pid, pcv_ShowStock

pcv_ShowStock=1

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

session("cp_lct_form_iPageCurrent")=iPageCurrent


if request("order")<>"" then
	if IsNumeric(request("order")) then
		session("cp_lct_form_order")=request("order")
	end if
end if

if request("action")="update" then

	session("intShowBOonly")=request.Form("showBOonly")
		if session("intShowBOonly")="" then
			session("intShowBOonly")=0
		end if
	
	session("intShowOOSonly")=request.Form("showOOSonly")
		if session("intShowOOSonly")="" then
			session("intShowOOSonly")=0
		end if

 count=request("count")
 for i=1 to count
  query="SELECT stock FROM products WHERE idproduct=" & request("ID" & i)
  set rs=Server.CreateObject("ADODB.Recordset")
  Set rs=conntemp.execute(query) 
  stock=clng(rs("stock"))
  set rs=nothing
  newstock=stock
 	if request("C" & i)="1" then
		if request("total")<>"" then
			newstock=newstock+clng(request("total"))
  		else
	  		newstock=newstock+clng(request("Q" & i))
		end if
  end if
  if newstock<>stock then
		query="UPDATE products SET stock="& newstock &"  WHERE idproduct="& request("ID" & i)
		Set rs=conntemp.execute(query)
		Set rs=nothing
		
		call pcs_hookStockChanged(request("ID" & i), "")

		'BackInStock-S		
		Call pcs_hookInStockEvent(request("ID" & i), "")
		'BackInStock-E
		
		call updPrdEditedDate(request("ID" & i))
  end if
 next 
	
	If statusAPP="1" OR scAPP=1 Then
		'// Update parent products inventory levels if necessary 
		%>
		<!--#include file="../pc/app-updstock.asp"-->
		<% 
	End If

 msg="Select products updated successfully"
 msgType=1
end if
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" name="checkboxform" action="viewStocka.asp?action=update&iPageCurrent=<%=request("iPageCurrent")%>&order=<%=request("order")%>&sort=<%=request("sort")%>" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="8" class="pcCPspacer" align="center">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
		</td>
	</tr>
	<tr> 
		<th nowrap colspan="2" width="10%"><a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent%>&order=2"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent%>&order=3"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;SKU</th>
		<th nowrap colspan="2" width="70%"><a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent%>&order=4"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent%>&order=5"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Product</th>
		<th nowrap colspan="2" width="8%"><a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent%>&order=6"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent%>&order=7"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;In Stock</th>
		<th nowrap width="10%">+/- Units</th>
		<th nowrap width="2%">Select</th>
	</tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
                      
		<!--#include file="inc_srcPrdQuery.asp"-->
		<%Set rsInv=Server.CreateObject("ADODB.Recordset")
		rsInv.CacheSize=session("cp_lct_form_iPageSize")
		rsInv.PageSize=session("cp_lct_form_iPageSize")
		rsInv.Open query, connTemp, adOpenStatic, adLockReadOnly
		If rsInv.eof Then
		pcv_ShowStock=0
		%>
			<tr> 
				<td colspan="8">
					<div class="pcCPmessage">No products found. <a href="viewStock.asp">New search &gt;&gt;</a></div>
                    <div>The reason could be that...</div>
					<ul>
						<li>This store currently does not have products for which it is tracking inventory</li>
						<li>The filters that you have selected returned no results. Run a <a href="viewStock.asp">new search &gt;&gt;</a></li>
					</ul>
				</td>
			</tr>
		<%
		Else 			
			rsInv.MoveFirst

			' get the max number of pages
			Dim iPageCount
			iPageCount=rsInv.PageCount
			If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
			If iPageCurrent < 1 Then iPageCurrent=1
				
			' set the absolute page
			rsInv.AbsolutePage=iPageCurrent  
			
			Count=0
			Count1=0
			Do While NOT rsInv.EOF And Count1 < rsInv.PageSize
			count=count + 1
			count1=count1 + 1
		
			pcv_ReorderLevel=rsInv("pcProd_ReorderLevel")
			if IsNull(pcv_ReorderLevel) or pcv_ReorderLevel="" then
				pcv_ReorderLevel="0"
			end if

			if statusAPP="1" then
				pcv_Apparel=rsInv("pcprod_Apparel")
				if IsNull(pcv_Apparel) or pcv_Apparel="" then
					pcv_Apparel="0"
				end if
			else
				pcv_Apparel="0"
			end if
			%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td colspan="2" nowrap><%=rsInv("sku")%></td>
					<td colspan="2">
						<a href="FindProductType.asp?id=<%=rsInv("idproduct")%>" target="_blank"><%=rsInv("description")%></a>
					</td>
					<%
					PR_Name=rsInv("description")%>
					<td colspan="2"><%if pcv_Apparel="0" then%><%=rsInv("stock")%><%end if%></td>
					<td>
						<%if pcv_Apparel="0" then%>
							<input type="text" name="Q<%=count%>" size="4">
						<%else%>
							<input type="hidden" name="Q<%=count%>" size="10" value="0">
						<%end if%>
					</td>
					<td>
						<%if pcv_Apparel="0" then%>
							<input type="checkbox" name="C<%=count%>" value="1" class="clearBorder">
						<%else%>
							<input type="hidden" name="C<%=count%>" value="">
						<%end if%>
					<input type="hidden" name="ID<%=count%>" value="<%=rsInv("idproduct")%>">
					</td>				
				</tr>
                      
				<%
				if (pcv_Apparel<>"0") then
					query="SELECT description,stock,idproduct FROM products WHERE removed=0 AND pcprod_ParentPrd=" & rsInv("idproduct") & "  ORDER BY idproduct asc;"
					set rs1=connTemp.execute(query)
	
					do while not rs1.eof
						count=count+1%>
						<tr>
							<td colspan="2" align="center" bgcolor="<%= strCol %>">&nbsp;</td>
							<td colspan="2" bgcolor="<%= strCol %>" >&nbsp;&nbsp;
								<%
								SPname=trim(replace(rs1("description"),PR_Name,""))
								SPname=mid(SPname,2,len(SPname))
								SPname=Left(SPname,len(SPname)-1)
								%><%=SPname%>
							</td>
							<td colspan="2" bgcolor="<%= strCol %>" > 
								<%if (clng(rs1("stock"))<=pcv_ReorderLevel) or (rs1("stock")="0") then%><font color=#FF0000><b><%end if%><%=rs1("stock")%>
								<%if (clng(rs1("stock"))<=pcv_ReorderLevel) or (rs1("stock")="0") then%></b></font><%end if%>
							</td>
							<td width="64" bgcolor="<%= strCol %>"> 
								<input type="text" name="Q<%=count%>" size="10">
							</td>
							<td width="68" bgcolor="<%= strCol %>"> 
								<input type="checkbox" name="C<%=count%>" value="1">
								<input type=hidden name="ID<%=count%>" value="<%=rs1("idproduct")%>">
							</td>
						</tr>
						<%
						rs1.MoveNext
					loop
					set rs1=nothing
				end if

			rsInv.MoveNext
			Loop
			%>
			<tr>
				<td colspan="8" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="8" align="right" class="cpLinksList">
				<input type="hidden" name="count" value=<%=count%>>
				<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></td>
			</tr>
			<tr>
				<td colspan="8"><hr></td>
			</tr>	  
			<tr>
				<td colspan="8">Change inventory for <u>all checked products</u> on this page by the following number of units:</td>
			</tr>
			<tr>
				<td colspan="8">+/- Units: <input type="text" name="total" size="10">
				</td>
			</tr>
			<tr>
				<td colspan="8"><hr></td>
			</tr> 
			<tr>
				<td colspan="8">
					<input type="submit" name="submit" value="Update" class="btn btn-primary">&nbsp;
					<input type="button" class="btn btn-default"  name="back" value="New Search" onClick="location='viewstock.asp';">
				</td>
			</tr>
	<%
	End If
	
	If iPageCount>1 Then
	%>
  <tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>                            
	<tr> 
		<td colspan="8"><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%></td>
	</tr>
	<tr>                   
	<td colspan="8"> 
		<%' display Next / Prev buttons
		if iPageCurrent > 1 then %>
		<a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
		<%
		end If
		For I=1 To iPageCount
		If Cint(I)=Cint(iPageCurrent) Then %>
			<b><%=I%></b> 
		<%
		Else
		%>
			<a href="viewStocka.asp?iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
		<%
		End If
		Next
			if CInt(iPageCurrent) < CInt(iPageCount) then %>
				<a href="viewStocka.asp?iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
		<%
			end If
		%>
	</td>
	</tr>
<% End If %>
</table>
</form>
<% if pcv_ShowStock<>0 then %>
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
if(test=="+"||test=="-"||test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
  if (theForm.total.value == "")
  {
	for (var j = 1; j <= <%=count%>; j++) {
	box = eval("document.checkboxform.C" + j); 
	if (box.checked == true)
	{
	qtt= eval("document.checkboxform.Q" + j);
		if (qtt.value == "")
	  	{
	    alert("Please enter a value for this field!");
	    qtt.focus();
	    return (false);
		}
		else
		{
			if (allDigit(qtt.value) == false)
			{
		    alert("Please enter a numeric value for this Field.");
		    qtt.focus();
		    return (false);
		    }
	    }
	}
	}
  }
  else
  {
	  if (allDigit(theForm.total.value) == false)
	  {
	    alert("Please enter a right value for this field.");
	    theForm.total.focus();
	    return (false);
	  }
  }

return (true);
}
</script>
<% end if %>
<!--#include file="AdminFooter.asp"-->