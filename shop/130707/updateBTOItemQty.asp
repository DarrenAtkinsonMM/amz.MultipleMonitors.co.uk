<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Configurable Item Inventory Management" %>
<% section="services" %>
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

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="description"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If 
	

if request("action")="clearFilters" then
	session("intShowOOSonly")=0
	call closeDb()
response.redirect("updateBTOItemQty.asp")
	response.End()
end if

if request("action")="update" then

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
		
		call updPrdEditedDate(request("ID" & i))
  end if
 next 
end if
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" name="checkboxform" action="updateBTOItemQty.asp?action=update&iPageCurrent=<%=request("iPageCurrent")%>&order=<%=request("order")%>&sort=<%=request("sort")%>" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="8" class="pcCPspacer" align="center">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
		</td>
	</tr>
	<tr> 
		<th nowrap colspan="2"><a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;SKU</th>
		<th nowrap colspan="2"><a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Product</th>
		<th nowrap colspan="2"><a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent%>&order=stock&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent%>&order=stock&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;In Stock</th>
		<th nowrap>+/- Units</th>
		<th nowrap>Select</th>
	</tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
                      
		<%
		if session("intShowOOSonly")<>0 then
			query2=" AND stock<1"
		end if
		query="SELECT idproduct,sku,description,stock FROM products WHERE active=-1 AND removed=0 AND configOnly=1 AND noStock=0"&query1&query2&" ORDER BY "& strORD &" "& strSort
		Set rsInv=Server.CreateObject("ADODB.Recordset")
		rsInv.CacheSize=15
		rsInv.PageSize=15
		rsInv.Open query, connTemp, adOpenStatic, adLockReadOnly
		If rsInv.eof Then
		pcv_ShowStock=0
		%>
			<tr> 
				<td colspan="8">
					<div class="pcCPmessage">No Configurable Items found. <a href="updateBTOItemQty.asp?action=clearFilters">Clear Filters &gt;&gt;</a>.
					<ul class="pcListIcon">
						<li>This store currently does not have Configurable Items for which it is tracking inventory</li>
						<li>Or the filters that you have selected returned no results (<a href="updateBTOItemQty.asp?action=clearFilters">Clear Filters &gt;&gt;</a>)</li>
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
			Do While NOT rsInv.EOF And Count < rsInv.PageSize
			count=count + 1
			%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td colspan="2"><%=rsInv("sku")%></td>
					<td colspan="2">
						<a href="FindProductType.asp?id=<%=rsInv("idproduct")%>" target="_blank"><%=rsInv("description")%></a>
					</td>
					<td colspan="2"><%=rsInv("stock")%></td>
					<td><input type="text" name="Q<%=count%>" size="10"></td>
					<td><input type="checkbox" name="C<%=count%>" value="1" class="clearBorder"><input type="hidden" name="ID<%=count%>" value="<%=rsInv("idproduct")%>"></td>
				</tr>
                      
			<% 
			rsInv.MoveNext
			Loop
			%>
			<tr>
				<td colspan="8" align="right" class="cpLinksList">
				<input type="hidden" name="count" value=<%=count%>>
				<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></td>
			</tr>
			<tr>
				<td colspan="8"><hr></td>
			</tr>	  
			<tr>
				<td colspan="8">Change inventory for <u>all checked Configurable Items</u> in this page by the following number of units:</td>
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
					<input type="checkbox" value="1" name="showOOSonly" class="clearBorder" <%if session("intShowOOSonly")<>0 then%>checked<%end if%>> Show Out-of-Stock Items only
					&nbsp;&nbsp;			
				</td>
			</tr>
			<tr>
				<td colspan="8"><hr></td>
			</tr>	 
			<tr>
				<td colspan="8">
					<input type="submit" name="submit" value="Update" class="btn btn-primary">&nbsp;
					<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
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
		<a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
		<%
		end If
		For I=1 To iPageCount
		If Cint(I)=Cint(iPageCurrent) Then %>
			<b><%=I%></b> 
		<%
		Else
		%>
			<a href="updateBTOItemQty.asp?iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"> 
			<%=I%></a> 
		<%
		End If
		Next
			if CInt(iPageCurrent) < CInt(iPageCount) then %>
				<a href="updateBTOItemQty.asp?iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
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
<% end if %><!--#include file="AdminFooter.asp"-->
