<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
if request.form("submit")<>"" then
	CP_Service=request.form("CP_Service")
	Session("ship_CP_Service")=CP_Service
	if CP_Service="" then
		call closeDb()
response.redirect "4_Step2.asp?msg="&Server.URLEncode("Select at least one service.")
		response.end
	end if
	freeshipStr=""
	handlingStr=""
	
	shipServiceArray=split(CP_Service,", ")
	for i=0 to ubound(shipServiceArray)
		If request.form("free"&shipServiceArray(i))="YES" then
			freeamt=request.form("amt"&shipServiceArray(i))
			freeshipStr=freeshipStr&shipServiceArray(i)&"|"&replacecomma(freeamt)&","
		End if
		If request.form("handling"&shipServiceArray(i))<>"0" AND request.form("handling"&shipServiceArray(i))<>"" then
			If isNumeric(request.form("handling"&shipServiceArray(i)))=true then
				handlingStr=handlingStr&shipServiceArray(i)&"|"&replacecomma(request.form("handling"&shipServiceArray(i)))&"|"&request.form("shfee"&shipServiceArray(i))&","
			End If
		End if
	next

	Session("ship_CP_freeshipStr")=freeshipStr
	Session("ship_CP_handlingStr")=handlingStr
	call closeDb()
response.redirect "4_Step3.asp"
	response.end
else %>
	<form name="form1" method="post" action="4_Step2.asp" class="pcForms">
		<table class="pcCPcontent">
            <tr>
                <td colspan="2" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
			<tr> 
				<td colspan="2">Choose one or more shipping services to offer to your customers.</td>
			</tr>

			<%
				query="SELECT serviceCode, serviceDescription FROM shipService WHERE idShipment=7"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
				
				while not rs.eof
					serviceCode = rs("serviceCode")
					serviceName = rs("serviceDescription")
					rs.movenext
				%>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
				<tr bgcolor="#DDEEFF">
					<td align="right"> 
						<input type="checkbox" name="CP_Service" value="<%=serviceCode%>">
					</td>
					<td><font color="#3366CC"><b><%=serviceName%></b></font></td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>
						<input name="free<%=serviceCode%>" type="checkbox" id="free<%=serviceCode%>" value="YES">
						Offer free shipping for orders over <%=scCurSign%> 
						<input name="amt<%=serviceCode%>" type="text" id="amt<%=serviceCode%>" size="10" maxlength="10">
					</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>Add Handling Fee <%=scCurSign%> 
					<input name="handling<%=serviceCode%>" type="text" id="handling<%=serviceCode%>" size="10" maxlength="10">
					</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>
					<input type="radio" name="shfee<%=serviceCode%>" value="-1" checked>
					Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
					<input type="radio" name="shfee<%=serviceCode%>" value="0">
					Integrate into shipping rate.</td>
				</tr>
			<% wend %>        

            <tr>
                <td colspan="2" class="pcCPspacer"><hr></td>
            </tr>
                            
			<tr> 
				<td>&nbsp;</td>
				<td>
				<input type="submit" name="Submit" value="Submit" class="btn btn-primary"></td>
			</tr>
		</table>
	</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->
