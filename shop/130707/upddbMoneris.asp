<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next %>
<% '--Updater File-- %>
<% pageTitle = "Moneris Gateway -   Database Update" %>
<% Section="paymntOpt" %>
<%PmAdmin=5%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f, iCnt
dim pcArr, intCount,i,j,pcArr1, intCount1,pcArr2, intCount2

tmp_mode=request("mode")
tmp_id=request("id")

Response.Cookies("AgreeLicense") = ""
IF request("action")="sql" then
	if request.querystring("hmode")="2" then
		SSIP=request("SSIP")
		UID=request("UID")
		PWD=request("PWD")
		SSDB=request("SSDB")
		if SSIP="" or UID="" or PWD="" then
			call closeDb()
response.redirect "upddbMoneris.asp?mode=3"
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
response.redirect "upddbMoneris.asp?mode=1"
			response.End
		end if
		
	end if
	
	iCnt=0
	ErrStr=""		
  
	'//  Create column pcPay_Moneris_Meth for table pcPay_Moneris
	query="ALTER TABLE [pcPay_Moneris] ADD [pcPay_Moneris_Meth] [int] NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Moneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	else
		query="UPDATE [pcPay_Moneris] SET pcPay_Moneris_Meth=0;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		query="ALTER TABLE [pcPay_Moneris] ADD CONSTRAINT [DF_pcPay_Moneris_pcPay_Moneris_Meth] DEFAULT (0) FOR [pcPay_Moneris_Meth];"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if
  
	'//  Create column pcPay_Moneris_Interac for table pcPay_Moneris 
	query="ALTER TABLE [pcPay_Moneris] ADD [pcPay_Moneris_Interac] [int] NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Moneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	else
		query="UPDATE pcPay_Moneris SET pcPay_Moneris_Interac=0;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		query="ALTER TABLE [pcPay_Moneris] ADD CONSTRAINT [DF_pcPay_Moneris_pcPay_Moneris_Interac] DEFAULT (0) FOR [pcPay_Moneris_Interac];"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if


	'//  Create column pcPay_Moneris_Interac_MerchantID for table pcPay_Moneris 
	query="ALTER TABLE [pcPay_Moneris] ADD [pcPay_Moneris_Interac_MerchantID] [nvarchar] (20) NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Moneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if
	
	
	query="ALTER TABLE [pcPay_OrdersMoneris] ADD [pcPay_MOrder_IDEBIT_ISSCONF] [nvarchar] (30) NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_OrdersMoneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if
	
	
	query="ALTER TABLE [pcPay_OrdersMoneris] ADD [pcPay_MOrder_IDEBIT_ISSNAME] [nvarchar] (30) NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_OrdersMoneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
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
		strMode="errors"
	else
		strMode="complete"
	end if

END IF

%>
<!--#include file="AdminHeader.asp"-->
<form action="upddbMoneris.asp" method="get" name="form1" class="pcForms">
<% 
'mode=Edit&id=219&gwchoice=11
if strMode="complete" then
	call closeDb()
response.redirect "pcConfigurePayment.asp?mode="&tmp_mode&"&id="&tmp_id&"&gwchoice=11"
	response.end
else %>
	<input type="hidden" name="mode" value="<%=tmp_mode%>">
	<input type="hidden" name="id" value="<%=tmp_id%>">
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

		<% if strMode="errors" then %>
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
				<span style="font-weight: bold">You must update your database before you can add/modify Moneris Gateway to support Verify By Visa And Interac&reg; online.<u></u></span><br>
				In order to activate Moneris real-time payment gateway you will need to update your ProductCart database to add the required table.</div>				<p><strong><br>
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
					<td>Database	Name:	<input name="SSDB" type="text" id="SSDB" size="30"></td>
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
					  <input type="button" class="btn btn-default"  name="access2" value="Update Your ProductCart MS SQL Database" onClick="location='upddbMoneris.asp?action=sql&mode=<%=tmp_mode%>&id=<%=tmp_id%>';" class="btn btn-primary">
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
