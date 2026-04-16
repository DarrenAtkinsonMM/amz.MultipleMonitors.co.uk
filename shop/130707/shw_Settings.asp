<%
'This file is part of ProductCart, an ecommerce application developed and sold by Netsource Commerce. ProductCart, its source code, the ProductCart name and logo are property of Netsource Commerce. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Netsource Commerce. To contact Netsource Commerce, please visit www.productcart.com.
%>
<% pageTitle="SHIPWIRE Settings" %>
<% response.Buffer=false %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%
pcPageName="shw_Settings.asp"

IF request("action")="sync" THEN
	call SHWSyncAllInventoryStatus()
	IF shwStatus="0" then
		if scDB="SQL" then
			queryQ="UPDATE pcShipwireSettings SET pcSWS_SyncDate='" & Now() & "';"
		else
			queryQ="UPDATE pcShipwireSettings SET pcSWS_SyncDate=#" & Now() & "#;"
		end if
		set rsQ=connTemp.execute(queryQ)
		set rsQ=nothing
		msg="Synchronized Inventory Status successfully!"
		msgType=1
	ELSE
		tmpError=SHWGetErrorList()
		msg="Cannot synchronize Inventory Status. Please try again."
		if tmpError<>"" then
			msg=msg & "<br>Error message(s):<br>" & tmpError
		end if
		msgType=0
	END IF
END IF
	
IF request("action")="update" THEN
	shwUser=getUserInput(request("Username"),0)
	shwPass=enDeCrypt(getUserInput(request("Password"),0), scCrypPass)
	shwMode=request("shwMode")
	if shwMode="" then
		shwMode=0
	end if
	shwOnOff=request("turnon")
	if shwOnOff="" then
		shwOnOff=0
	end if
	queryQ="SELECT pcSWS_Username,pcSWS_Password,pcSWS_Mode,pcSWS_OnOff,pcSWS_SyncDate FROM pcShipwireSettings;"
	set rsQ=connTemp.execute(queryQ)
	if rsQ.eof then
		queryQ="INSERT INTO pcShipwireSettings (pcSWS_Username,pcSWS_Password,pcSWS_Mode,pcSWS_OnOff) VALUES ('" & shwUser & "','" & shwPass & "'," & shwMode & "," & shwOnOff & ");"
		set rsQ=connTemp.execute(queryQ)
		set rsQ=nothing
	else
		queryQ="UPDATE pcShipwireSettings SET pcSWS_Username='" & shwUser & "',pcSWS_Password='" & shwPass & "',pcSWS_Mode=" & shwMode & ",pcSWS_OnOff=" & shwOnOff & ";"
		set rsQ=connTemp.execute(queryQ)
		set rsQ=nothing
	end if
	set rsQ=nothing
	msg="SHIPWIRE settings were updated successfully!"
	msgType=1
END IF	

shwUser=""
shwPass=""
shwMode=0
shwOnOff=0
shwSyncDate=""
queryQ="SELECT pcSWS_Username,pcSWS_Password,pcSWS_Mode,pcSWS_OnOff,pcSWS_SyncDate FROM pcShipwireSettings;"
set rsQ=connTemp.execute(queryQ)
if not rsQ.eof then
	shwUser=rsQ("pcSWS_Username")
	shwPass=enDeCrypt(rsQ("pcSWS_Password"), scCrypPass)
	shwMode=0
	if rsQ("pcSWS_Mode")="1" then
		shwMode=1
	end if
	if rsQ("pcSWS_OnOff")="1" then
		shwOnOff=1
	end if
	shwSyncDate=rsQ("pcSWS_SyncDate")
end if
set rsQ=nothing

%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
	<script type=text/javascript>
	function checkFormA(tmpForm)
	{
		if (tmpForm.username.value=="")
		{
			alert("Please enter a value for 'Username' field");
			tmpForm.username.focus();
			return(false);
		}
		if (tmpForm.password.value == false)
		{
			alert("Please enter a value for 'Password' field");
			tmpForm.password.focus();
			return(false);
		}
		return(true);
	}
	</script>
	<form name="formA" method="post" action="<%=pcPageName%>?action=update" onsubmit="javascript: return(checkFormA(this));">
	<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2"><b>Turn SHIPWIRE Services On & Off</b></th>
	</tr>
	<tr>
		<td colspan="2"><input type="radio" name="turnon" value="1"  <%if shwOnOff=1 then%>checked<%end if%> class="clearBorder"> Turn ON&nbsp;&nbsp;&nbsp;<input type="radio" name="turnon" value="0" <%if shwOnOff=0 then%>checked<%end if%> class="clearBorder"> Turn OFF</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2"><b>SHIPWIRE Account Information</b></th>
	</tr>
	<tr valign="top">
		<td>User name:</td>
		<td>
			<input type="text" name="username" size="30" value="<%=shwUser%>"> <img src="images/sample/pc_icon_required.gif" border="0">
		</td>
	</tr>
	<tr>
		<td>Password:</td>
		<td>
			<input type="password" name="password" size="30" value="<%=shwPass%>"> <img src="images/sample/pc_icon_required.gif" border="0">
		</td>
	</tr>
	<tr valign="top">
		<td nowrap>Mode:</td>
		<td>
			<select name="shwmode">
				<option value="0" <%if shwMode=0 then%>selected<%end if%>>Test</option>
				<option value="1" <%if shwMode=1 then%>selected<%end if%>>Production</option>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<input type="submit" name="submitFormA" value=" Save Settings " class="btn btn-primary">
		</td>
	</tr>
	</table>
	</form>
	<%if shwOnOff=1 then%>
	<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2"><b>Synchronizing Store with Inventory Status at Shipwire</b></th>
	</tr>
	<tr>
		<td colspan="2">
			<%if not (IsNull(shwSyncDate)) then%>
				Last Synchronized Date: <font color=blue><%=shwSyncDate%></font><br>
			<%end if%>
			Click on the button below to synchronize all products of your store with the inventory status at Shipwire.<br><br>
			<input type="submit" name="syncprd" value=" Synchronize Inventory Status " onclick="pcf_Open_ShipwirePop();location='<%=pcPageName%>?action=sync';" class="btn btn-primary">
		</td>
	</tr>
	</table>
	<%end if%>
<%%>
<%Response.write(pcf_ModalWindow("Connecting to Shipwire Server... ","ShipwirePop", 300))%>
<!--#include file="AdminFooter.asp"-->