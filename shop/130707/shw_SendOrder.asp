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
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%
' Validate order ID
pcv_IdOrder=getUserInput(request("idorder"),10)
if not validNum(pcv_IdOrder) then pcv_IdOrder=0
if pcv_IdOrder=0 then 
	call closeDb()
	response.redirect "menu.asp"
end if

query="SELECT shippingCountryCode FROM Orders WHERE idorder=" & pcv_IdOrder & ";"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

UseLocalMethod=0
if not rs.eof then
	pcv_CountryCode=rs("shippingCountryCode")
	if (Ucase(scShipFromPostalCountry)=UCase(pcv_CountryCode)) AND ((Ucase(pcv_CountryCode)="US") OR (Ucase(pcv_CountryCode)="UK") OR (Ucase(pcv_CountryCode)="CA") OR (Ucase(pcv_CountryCode)="GB")) then
		UseLocalMethod=1
	end if
end if
set rs=nothing
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_IdOrder))%></b></td>
	</tr>
</table>
	<%
	query="SELECT Products.idproduct, Products.Description, Products.sku, ProductsOrdered.idProductOrdered, ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_Shipped FROM Products INNER JOIN ProductsOrdered ON Products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & pcv_IdOrder & " AND pcPrdOrd_Shipped=0;"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	IF rs.eof THEN
		set rs=nothing%>
		<table class="pcCPcontent">
		<tr>
		<td colspan="2">
		<div class="pcCPmessage">
			No products found.<br /><br />
			<a href="#" onClick="javascript:history.back();">Back</a>
		</div>
		</td>
		</tr>
		</table>
	<% ELSE %>
		<form name="form1" method="post" action="shw_SendOrder1.asp?action=send" onSubmit="return Form1_Validator(this)" class="pcForms">
			<table class="pcCPcontent">
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th width="5%">&nbsp;</th>
					<th>SKU</th>
					<th width="70%">Product Name</th>
					<th><div align="right">Quantity</div></th>
				</tr>
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<%
				pcv_count=0
				pcv_available=0
				Do while not rs.eof
					pcv_cancheck=0
					pcv_count=pcv_count+1
					pcv_IDProduct=rs("idproduct")
					pcv_IDProductOrdered=rs("idProductOrdered")
					pcv_Description=rs("description")
					pcv_Sku=rs("sku")
					pcv_Qty=rs("quantity")
					if IsNull(pcv_Qty) or pcv_Qty="" then
						pcv_Qty=0
					end if
					pcv_Shipped=rs("pcPrdOrd_Shipped")
					if IsNull(pcv_Shipped) or pcv_Shipped="" then
						pcv_Shipped=0
					end if
					
					shwQty=SHWGetInventoryStatus(pcv_Sku)
					if clng(shwQty)<>-1 then
						queryQ="UPDATE Products SET stock=" & shwQty & " WHERE idProduct=" & pcv_IDProduct & ";"
						set rsQ=connTemp.execute(queryQ)
						set rsQ=nothing
						call pcs_hookStockChanged(pcv_IDProduct, "")
					end if
					if clng(shwQty)>=clng(pcv_Qty) then
						pcv_cancheck=1
						pcv_available=pcv_available+1
					end if
					%>
					<tr valign="top">
						<td align="center" width="5%">
							<% dim pcv_showLink
							pcv_showLink = 1
							%>
							<input type="checkbox" name="C<%=pcv_count%>" value="1" <%if pcv_cancheck=1 then%>checked<%else%>disabled<%end if%> class="clearBorder">
							<input type="hidden" name="IDPrd<%=pcv_count%>" value="<%=pcv_IDProductOrdered%>">
						</td>
						<td align="left" nowrap><a href="FindProductType.asp?id=<%=pcv_IDProduct%>" target="_blank"><%=pcv_Sku%></a></td>
						<td align="left" width="60%"><a href="FindProductType.asp?id=<%=pcv_IDProduct%>" target="_blank"><%=pcv_Description%></a></td>
						<td align="right" width="35%" nowrap><b><%=pcv_Qty%></b><br>
						<%
						If Clng(shwQty)<Clng(pcv_Qty) AND Clng(shwQty)>0 then%>
						<font color=red><i>(Low Stock. Current Stock at Shipwire: <%=shwQty%>)</i></font>
						<%else
						Select Case Clng(shwQty)
						Case -1:%><font color=red><i>(Cannot find this product at Shipwire)</i></font>
						<%Case 0:%><font color=red><i>(Out of stock at Shipwire)</i></font>
						<%Case Else: %><i>(Current Stock at Shipwire: <%=shwQty%>)</i>
						<%End Select
						end if%>
						</td>
						
					</tr>

				<%
				rs.MoveNext
				loop
				set rs=nothing%>
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="4" align="left">
					<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
					<br />
					<br />
					<script type=text/javascript>
					function Form1_Validator(theForm)
					{
						var hase=0;
						for (var j = 1; j <= <%=pcv_count%>; j++) {
							box = eval("document.form1.C" + j);
							if ((box.checked == true) && (box.disabled == false))
							{ hase=1; break; }
						 }
						 if (hase==0)
						 {alert("Please select at least 1 product to send");
							return (false);}
						pcf_Open_ShipwirePop();
						return (true);
					}

					function checkAll() {
					for (var j = 1; j <= <%=pcv_count%>; j++) {
					box = eval("document.form1.C" + j);
					if ((box.checked == false) && (box.disabled == false)) box.checked = true;
						 }
					}

					function uncheckAll() {
					for (var j = 1; j <= <%=pcv_count%>; j++) {
					box = eval("document.form1.C" + j);
					if ((box.checked == true) && (box.disabled == false)) box.checked = false;
						 }
					}
					</script>
					<%if pcv_available>0 then%>
					<b>Shipping Method:</b> <select name="shwMethod" id="shwMethod">
					<%if UseLocalMethod=1 then%>
					<option value="">Auto</option>
					<option value="GD">Ground</option>
					<option value="2D">2 Day</option>
					<option value="1D">1 Day</option>
					<%else%>
					<option value="">Auto</option>
					<option value="E-INTL">International Economy</option>
					<option value="INTL">International Standard</option>
					<option value="PL-INTL">International Plus</option>
					<option value="PM-INTL">International Premium</option>
					<%end if%>
					</select><br><br>
					<input type="submit" name="submit1" value="Send Selected Products to SHIPWIRE" class="btn btn-primary">
					<%end if%>
					&nbsp;<input type="button" class="btn btn-default"  name="Back" value=" Back " onClick="javascript:history.back();">
					<input type="hidden" name="count" value="<%=pcv_count%>">
					<input type="hidden" name="idorder" value="<%=pcv_IdOrder%>">
					</td>
				</tr>
			</table>
		</Form>
	<%END IF%>
<%%>
<%Response.write(pcf_ModalWindow("Connecting to SHIPWIRE Server... ","ShipwirePop", 300))%>
<!--#include file="AdminFooter.asp"-->