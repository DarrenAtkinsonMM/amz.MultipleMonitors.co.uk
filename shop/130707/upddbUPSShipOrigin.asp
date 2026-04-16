<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next %>
<% '--Updater File-- %>
<% pageTitle = "UPS, Canada Origin - Database Update" %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
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
response.redirect "upddbUPSShipOrigin.asp?mode=3"
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
response.redirect "upddbUPSShipOrigin.asp?mode=1"
			response.End
		end if
		
	end if
	
	iCnt=0
	ErrStr=""		

	'**** Alter serviceDescription for table shipService ***************************************
	If request("AlterOrigin")="CA" then
		query="DELETE FROM shipService WHERE serviceCode='03';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="DELETE FROM shipService WHERE serviceCode='59';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="DELETE FROM shipService WHERE serviceCode='64';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="DELETE FROM shipService WHERE serviceCode='65';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="UPDATE shipService SET serviceDescription='UPS Express<sup>SM</sup>' WHERE serviceCode='01';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="UPDATE shipService SET serviceDescription='UPS Expedited<sup>SM</sup>' WHERE serviceCode='02';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="UPDATE shipService SET serviceDescription='UPS Express Saver<sup>&reg;</sup>' WHERE serviceCode='13';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="UPDATE shipService SET serviceDescription='UPS Express Saver<sup>&reg;</sup> Early A.M.<sup>&reg;</sup>' WHERE serviceCode='14';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		query="UPDATE shipService SET serviceDescription='UPS Worldwide Express Plus<sup>SM</sup>' WHERE serviceCode='54';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	else
		If request("AlterOrigin")="US" then
			query="SELECT * FROM shipService WHERE serviceCode='03';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			if NOT rs.eof then
				query="UPDATE shipService SET serviceDescription='UPS Ground<sup>&reg;</sup>' WHERE serviceCode='03';"
			else
				query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation, idShipment) VALUES (0, '03', 0, 'UPS Ground<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0, 3);"
			end if
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
	
			query="SELECT * FROM shipService WHERE serviceCode='59';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			if NOT rs.eof then
				query="UPDATE shipService SET serviceDescription='UPS 2<sup>nd</sup> Day Air A.M.<sup>&reg;</sup>' WHERE serviceCode='59';"
			ELSE
				query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation, idShipment) VALUES (0, '59', 0, 'UPS 2<sup>nd</sup> Day Air A.M.<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0, 3);"
			END IF
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
		
			query="SELECT * FROM shipService WHERE serviceCode='65';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			if NOT rs.eof then
				query="UPDATE shipService SET serviceDescription='UPS Express Saver<sup>&reg;</sup>' WHERE serviceCode='65';"
			ELSE
				query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation, idShipment) VALUES (0, '65', 0, 'UPS Express Saver<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0, 3);"
			END IF
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
	
			query="UPDATE shipService SET serviceDescription='UPS Next Day Air<sup>&reg;</sup>' WHERE serviceCode='01';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		
			query="UPDATE shipService SET serviceDescription='UPS 2<sup>nd</sup> Day Air<sup>&reg;</sup>' WHERE serviceCode='02';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		
			query="UPDATE shipService SET serviceDescription='UPS Next Day Air Saver<sup>&reg;</sup>' WHERE serviceCode='13';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		
			query="UPDATE shipService SET serviceDescription='UPS Next Day Air<sup>&reg;</sup> Early A.M.<sup>&reg;</sup>' WHERE serviceCode='14';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		
			query="UPDATE shipService SET serviceDescription='UPS Worldwide Express Plus<sup>SM</sup>' WHERE serviceCode='54';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		End If
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
<form action="upddbUPSShipOrigin.asp" method="get" name="form1" class="pcForms">
<%
if mode="complete" then 
	call closeDb()
response.redirect "viewShippingOptions.asp"
	response.end
else
	dim pcv_AllwUSOption
	pcv_AllwUSOption=0
	
	
	query="SELECT * FROM shipService WHERE serviceCode='03';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	If rs.eof then
		'// Currently set to Canada origin, give US option
		pcv_AllwUSOption=1
	end if
	 %>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

		<% if mode="errors" then %>
			<tr>
				<td align="center">
					<div class="pcCPmessage">The following errors occurred while trying to update your database:<br><br>
				<%=ErrStr%></div>				</td>
			</tr>
		<% end if %>

		<tr>
			<td>
                <div class="pcCPnotes" style="padding:10px;">
				<span style="font-weight: bold">You must update your database before you can use Canada as your UPS Shipping Origin.<u></u></span><br>
				In order to use Canada as your UPS Shipping Origin you will need to update your ProductCart database to add the required shipping types.</div>				<p><strong><br>
				    <br>
				  You are about to update the store database. Please read the following carefully before proceeding.</strong></p>
				<p style="padding-top:10px;">Although we have tested this update script in a variety of environments, there is always the possibility of something going wrong. Make sure to <span style="font-weight: bold">backup your database</span> prior to executing this update. To do so:
				<ul>
				<li>Depending on how the database has been setup, you may be able to either perform the backup yourself or have your Web hosting company do it for you. Note: your SQL database is likely being automatically backed up every day: confirm that this is the case by asking your Web host when the last back up occurred.</li>
				</ul>
				<% if pcv_AllwUSOption=1 then %><p><strong>It appears that your Datebase is set to use Canada as the Shipping Origin.</strong></p>
				<p><br>
				  <input type="radio" name="AlterOrigin" id="AlterOrigin" value="US" class="clearBorder"> Use United States as Shipping Origin<br>
				  <input type="radio" name="AlterOrigin" id="AlterOrigin" value="CA" class="clearBorder" checked> Use Canada as Shipping Origin<br>
				<% else %>
                	<input type="hidden" name="AlterOrigin" value="CA">
               	<% end if %>
</p>
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
                            <input name="submit" type="submit" value="Update Your ProductCart MS SQL Database" class="btn btn-primary">
					<% else %>
							<input name="submit" type="submit" value="Update Your ProductCart MS SQL Database" class="btn btn-primary">
					<% end if %>
					</td>
			</tr>
			</table>
		</td>
	</tr>
		<tr>
		  <td>&nbsp;</td>
		  </tr>
</table>
<% end if %>
</form>
<!--#include file="AdminFooter.asp"-->