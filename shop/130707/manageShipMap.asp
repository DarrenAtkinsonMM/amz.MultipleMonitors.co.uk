<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Manage Shipping Filters"
pageIcon="pcv4_icon_ship.png"
Section="shipOpt"
%>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
if request("action")="upd" then
	FCount=getUserInput(request("FCount"),0)
	if FCount="" then
		FCount=0
	end if
	For i=1 to FCount
		tmpSMID=getUserInput(request("ID" & i),0)
		tmpSMOrder=getUserInput(request("O" & i),0)
		
		if (tmpSMID<>"") AND (tmpSMOrder<>"") then
			query="UPDATE pcShippingMap SET pcSM_Order=" & tmpSMOrder & " WHERE pcSM_ID=" & tmpSMID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	Next
	
	msg="Shipping filter order has beeen updated successfully!"
	msgType=1
end if

if scUseShipMap<>"1" then
	query="SELECT TOP 1 pcSM_ID,pcSM_Name,pcSM_Type,pcSM_Order FROM pcShippingMap;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		msg="Shipping rates are not being filtered. You can enable filters on the <a href=""modFromShipper.asp"" target=""_blank"">Shipping Settings</a> page."
		msgType=0
	end if
	set rs=nothing
end if
	
 %>
<!--#include file="AdminHeader.asp"-->
<form method="post" action="manageShipMap.asp?action=upd" name="mngSM" class="pcForms">
<table class="pcCPcontent">
    <tr>
    	<td colspan="6">
        	<!--#include file="pcv4_showMessage.asp"-->
        </td>
	</tr>
<%
	noShipServices=0
	intCount=-1
	query="SELECT [idshipservice],[serviceActive] FROM shipService WHERE serviceActive<>0;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if rs.EOF then
		set rs=nothing
		noShipServices=1
%>      
      <tr> 
        <td colspan="6"><div class="pcCPmessage">No active shipping methods found. Please add shipping methods before using this page.</div></td>
      </tr>
      <tr>
        <td colspan="6" class="pcCPspacer"></td>
      </tr>                
<% 
	Else                    
	set rs=nothing
	query="SELECT pcSM_ID,pcSM_Name,pcSM_Type,pcSM_Order FROM pcShippingMap ORDER BY pcSM_Order ASC, pcSM_Name ASC;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if rs.EOF then
		set rs=nothing
		
%>      
      <tr> 
        <td colspan="6"><div class="pcCPmessage">No shipping filters are defined.</div></td>
      </tr>
      <tr>
        <td colspan="6" class="pcCPspacer"></td>
      </tr>                
<% 

	Else 
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)%>
		<tr>
			<th width="10&" align="right">ID</td>
			<th width="25%">Filter Name</td>
			<th width="30%">Shipping Methods</td>
			<th width="15%">Display Type</td>
			<th width="10%">Order</td>
			<th width="10%">&nbsp;<input type="hidden" name="FCount" value="<%=Clng(intCount)+1%>"></td>
		</tr>
		<%For i=0 to intCount
		tmpSMID=tmpArr(0,i)
		tmpSMName=tmpArr(1,i)
		tmpSMType=tmpArr(2,i)
		tmpSMOrder=tmpArr(3,i)
		
		tmpRel=""
		queryQ="SELECT shipService.serviceDescription FROM shipService INNER JOIN pcSMRel ON shipService.idshipservice=pcSMRel.idshipservice WHERE pcSMRel.pcSM_ID=" & tmpSMID & " ORDER BY shipService.serviceDescription ASC;"
		set rsQ=connTemp.execute(queryQ)
		if not rsQ.eof then
			RelArr=rsQ.getRows()
			intR=ubound(RelArr,2)
			For iR=0 to intR
				if tmpRel<>"" then
					tmpRel=tmpRel & "<br>"
				end if
				tmpRel=tmpRel & RelArr(0,iR)
			Next
		end if
		set rsQ=nothing
		%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist" valign="top">
			<td width="10&"><%=tmpSMID%></td>
			<td width="20%"><a href="AddEditShipMap.asp?id=<%=tmpSMID%>"><%=tmpSMName%></a></td>
			<td width="20%"><%=tmpRel%></a></td>
			<td width="30%"><%if tmpSMType="1" then%>Highest Rate Method<%else%>Lowest Rate Method<%end if%></td>
			<td width="10%">
				<input type="text" name="O<%=i+1%>" id="O<%=i+1%>" size="4" value="<%=tmpSMOrder%>">
				<input type="hidden" name="ID<%=i+1%>" id="ID<%=i+1%>" value="<%=tmpSMID%>">
			</td>
			<td width="10%" nowrap class="cpLinksList">
				<a href="AddEditShipMap.asp?id=<%=tmpSMID%>">Edit</a> | <a href="javascript:if (confirm('You are about to remove this filter from your database. Are you sure you want to complete this action?')) location='delShipMap.asp?id=<%=tmpSMID%>';">Delete</a>
			</td>
		</tr>
		<%Next%>
		<tr>
		
	<%End If
	set rs=nothing
	End if
%>

<%if noShipServices=0 then%>
<tr>
	<td colspan="6">
        <%if intCount>=0 then%>
        <strong>Note:</strong> <i>When shipping filters are enabled, only shipping services mapped above can be used for the check-out process.</i><br><br>
        <input type="submit" name="Submit" value="Save Filter Order" class="btn btn-primary">&nbsp;
        <%end if%>
        <%AvailableMethods=0
        query="SELECT shipService.idshipservice FROM shipService WHERE (shipService.serviceActive<>0) AND (shipService.idshipservice NOT IN (SELECT idshipservice FROM pcSMRel));"
        set rs=connTemp.execute(query)
        if not rs.eof then
            AvailableMethods=1
        end if
        set rs=nothing
        if AvailableMethods=1 then%>
            <input type="button" <%if intCount<0 then%>class="btn btn-primary"<%else%>class="btn btn-default"<%end if%>  name="Go" value="Add New Filter" onClick="javascript:location='AddEditShipMap.asp';">&nbsp;
        <%end if%>
        <input type="button" class="btn btn-default"  name="Go" value="Manage Shipping Methods" onClick="javascript:location='viewShippingOptions.asp';">
        <input type="button" class="btn btn-default"  name="Go2" value="Manage Shipping Settings" onClick="javascript:location='modFromShipper.asp';">
	</td>
</tr>
<%else%>
<tr>
	<td colspan="6">
		<input type="button" class="btn btn-primary"  name="Go" value="Manage Shipping Methods" onClick="javascript:location='viewShippingOptions.asp';">
		<input type="button" class="btn btn-default"  name="Go2" value="Manage Shipping Settings" onClick="javascript:location='modFromShipper.asp';">
	</td>
</tr>
<%end if%>    
</table>
</form>
<br /><br />
<!--#include file="AdminFooter.asp"-->