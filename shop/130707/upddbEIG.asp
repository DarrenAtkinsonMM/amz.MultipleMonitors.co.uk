<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next %>
<% '--Updater File-- %>
<% pageTitle = "NetSource Commerce Gateway - Database Update" %>
<% Section="paymntOpt" %>
<%PmAdmin=5%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f, iCnt
dim pcArr, intCount,i,j,query,pcArr1, intCount1,pcArr2, intCount2

Response.Cookies("AgreeLicense") = ""
IF request("action")="sql" then
	if request.querystring("hmode")="2" then
		SSIP=request("SSIP")
		UID=request("UID")
		PWD=request("PWD")
		SSDB=request("SSDB")
		if SSIP="" or UID="" or PWD="" then
			call closeDb()
response.redirect "upddbEIG.asp?mode=3"
			response.End
		end if
		set connTemp=server.createobject("adodb.connection")
		connTemp.Open "Driver={SQL Server};Server="&SSIP&";Address="&SSIP&",1433;Network=DBMSSOCN;Database="&SSDB&";Uid="&UID&";Pwd="&PWD
		if err.number <> 0 then
			call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error while opening database")
		end if
	else
		if instr(ucase(scDSN),"DSN=") then
			call closeDb()
response.redirect "upddbEIG.asp?mode=1"
			response.End
		end if
		
	end if
	
	iCnt=0
	ErrStr=""		

	'// Create table pcPay_EIG
	query="CREATE TABLE [dbo].[pcPay_EIG] ("
	query=query&"[pcPay_EIG_ID] [int] NULL  DEFAULT (1),"
	query=query&"[pcPay_EIG_Username] [nvarchar] (100) NULL ,"
	query=query&"[pcPay_EIG_Password] [nvarchar] (100) NULL ,"
	query=query&"[pcPay_EIG_Key] [nvarchar] (100) NULL ,"
	query=query&"[pcPay_EIG_Type] [nvarchar] (50) NULL ,"
	query=query&"[pcPay_EIG_Version] [nvarchar] (4) NULL ,"
	query=query&"[pcPay_EIG_Curcode] [nvarchar] (4) NULL ,"
	query=query&"[pcPay_EIG_CVV] [int] NULL DEFAULT(0) ,"
	query=query&"[pcPay_EIG_SaveCards] [int] NULL DEFAULT(0) ,"
	query=query&"[pcPay_EIG_UseVault] [int] NULL DEFAULT(0) ,"
	query=query&"[pcPay_EIG_TestMode] [int] NULL DEFAULT(0)"
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcPay_EIG")
	else
		query="INSERT INTO pcPay_EIG (pcPay_EIG_ID, pcPay_EIG_CVV, pcPay_EIG_TestMode) VALUES (1,0,1);"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if
	
	'// Create table pcPay_EIG_Vault
	query="CREATE TABLE [dbo].[pcPay_EIG_Vault] ("
	query=query&"[pcPay_EIG_Vault_ID] [INT] IDENTITY (1, 1) NOT FOR REPLICATION NOT NULL ,"
	query=query&"[idOrder] [int] NULL  DEFAULT (1),"
	query=query&"[idCustomer] [int] NULL  DEFAULT (1),"
	query=query&"[IsSaved] [int] NULL  DEFAULT (1),"
	query=query&"[pcPay_EIG_Vault_CardNum] [nvarchar] (25) NULL ,"
	query=query&"[pcPay_EIG_Vault_CardType] [nvarchar] (10) NULL ,"
	query=query&"[pcPay_EIG_Vault_CardExp] [nvarchar] (10) NULL ,"
	query=query&"[pcPay_EIG_Vault_Token] [nvarchar] (50) NULL "
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcPay_EIG_Vault")
	end if		
	
	'// Create table pcPay_EIG_Authorize
	query="CREATE TABLE [dbo].[pcPay_EIG_Authorize] ("
	query=query&"[idauthorder] [INT] IDENTITY (1, 1) NOT FOR REPLICATION NOT NULL ,"
	query=query&"[idOrder] [int] NULL  DEFAULT (1),"
	query=query&"[idCustomer] [int] NULL  DEFAULT (1),"
	query=query&"[captured] [int] NULL  DEFAULT (1),"
	query=query&"[pcSecurityKeyID] [int] NULL  DEFAULT (1),"
	query=query&"[vaultToken] [nvarchar] (50) NULL ,"
	query=query&"[amount] [money] NULL DEFAULT(0) ,"	
	query=query&"[paymentmethod] [nvarchar] (250) NULL ,"	
	query=query&"[transtype] [nvarchar] (250) NULL ,"
	query=query&"[authcode] [nvarchar] (250) NULL ,"
	query=query&"[ccnum] [nvarchar] (250) NULL ,"
	query=query&"[ccexp] [nvarchar] (10) NULL ,"
	query=query&"[cctype] [nvarchar] (25) NULL ,"	
	query=query&"[fname] [nvarchar] (250) NULL ,"	
	query=query&"[lname] [nvarchar] (250) NULL ,"	
	query=query&"[address] [nvarchar] (250) NULL ,"
	query=query&"[zip] [nvarchar] (25) NULL ,"	
	query=query&"[trans_id] [nvarchar] (250) NULL "
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcPay_EIG_Vault")
	end if	

	set rs=nothing
	set connTemp=nothing
	
	Function TrapSQLError(varTableName)		
		'// -2147217900 = Table 'x' already exists.
		'// -2147217887 = Field 'x' already exists in table 'x'.
		if ((Err.Number=-2147217900) OR (Err.Number=-2147217887)) then
			Err.Description=""
			err.number=0
		else
			ErrStr = ErrStr & "Error Creating Table "&varTableName&": "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	End Function

	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if

END IF

%>
<!--#include file="AdminHeader.asp"-->
<form action="upddbEIG.asp" method="get" name="form1" class="pcForms">
<%
if mode="complete" then 
	call closeDb()
response.redirect "pcConfigurePayment.asp?gwchoice=66"
	response.end
else %>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

		<% if mode="errors" then %>
			<tr>
				<td align="center">
					<div class="pcCPmessage">The following errors occurred while trying to update your database:<br><br>
				<%=ErrStr%></div>
				</td>
			</tr>
		<% end if %>

		<tr>
			<td>
                <div class="pcCPnotes" style="padding:10px;">
				<span style="font-weight: bold">You must update your database before you can add NetSource Commerce Gateway.<u></u></span><br>
				In order to activate NetSource Commerce Gateway you will need to update your ProductCart database to add the required table.</div>				<p><strong><br>
				    <br>
				  You are about to update the store database. Please read the following carefully before proceeding.</strong></p>
				<p style="padding-top:10px;">Although we have tested this update script in a variety of environments, there is always the possibility of something going wrong. Make sure to <span style="font-weight: bold">backup your database</span> prior to executing this update. To do so:
				<ul>
				<li>Depending on how the database has been setup, you may be able to either perform the backup yourself or have your Web hosting company do it for you. Note: your SQL database is likely being automatically backed up every day: confirm that this is the case by asking your Web host when the last back up occurred.</li>
				</ul>
				</p>
				<br />
			<table class="pcCPcontent" width="80%">
			<% if request.querystring("mode")="1" OR request.querystring("mode")="3" then %>
				<tr>
					<td>
					It appears that you are using a DSN connection to connect to your SQL server. In order to complete this update, please enter your SQL Server Information below:
					<% if request.querystring("mode")="1" then %>
						<br>
						<strong>*All fields are required.</strong>
					<% end if %>					</td>
				</tr>
				<tr>
					<td>Server Domain/IP:	<input name="SSIP" type="text" id="SSIP" size="30"></td>
				</tr>
				<tr>
					<td>Databas  Name:	<input name="SSDB" type="text" id="SSDB" size="30"></td>
				</tr>
				<tr>
					<td>User ID: <input name="UID" type="text" id="UID" size="30"></td>
				</tr>
				<tr>
					<td>Password: <input name="PWD" type="password" id="PWD" size="30"></td>
				</tr>
				<input name="hmode" type="hidden" id="hmode" value="2">
				<input name="action" type="hidden" id="action" value="sql">
			<% end if %>
				<tr>
					<td align="center">
					<% if request.querystring("mode")="1" OR request.querystring("mode")="3" then %>
							<input name="access2" type="submit" id="access2" value="Update Your ProductCart MS SQL Database" class="btn btn-primary">
						<% else %>
					  <input type="button" class="btn btn-default"  name="access2" value="Update Your ProductCart MS SQL Database" onClick="location='upddbEIG.asp?action=sql';" class="btn btn-primary">
						<% end if %>
					</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<% end if %>
</form>
<!--#include file="AdminFooter.asp"-->