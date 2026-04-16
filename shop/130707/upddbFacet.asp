<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next %>
<% '--Updater File-- %>
<% pageTitle = "ProductCart Facet Groups - Database Update" %>
<% Section = "" %>
<%PmAdmin=19%> 
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f, iCnt
dim pcArr, intCount,i,j,pcArr1, intCount1,pcArr2, intCount2

Response.Cookies("AgreeLicense") = ""
IF request("action")="sql" then
	if request.querystring("hmode")="2" then
		SSIP=request("SSIP")
		UID=request("UID")
		PWD=request("PWD")
		SSDB=request("SSDB")
		if SSIP="" or UID="" or PWD="" then
			call closeDb()
response.redirect "upddbFacet.asp?mode=3"
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
response.redirect "upddbFacet.asp?mode=1"
			response.End
		end if
		
	end if
	
	iCnt=0
	ErrStr=""
	
	if not TableExists("pcFacetGroups") then
		query="CREATE TABLE pcFacetGroups ("
		query=query&"pcFG_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"pcFG_Name [nvarchar](250) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcFacetGroups")
        end if
        set rs=nothing
    end if
	
	if not TableExists("pcFacets") then
		query="CREATE TABLE pcFacets ("
		query=query&"pcFC_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"pcFG_ID [INT] DEFAULT(0) NULL,"
		query=query&"pcFC_Code [nvarchar](250) NULL,"
		query=query&"pcFC_Name [nvarchar](250) NULL,"
		query=query&"pcFC_Img [nvarchar](250) NULL,"
		query=query&"pcFC_Order [INT] DEFAULT(0) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcFacets")
        end if
        set rs=nothing
    end if
	
	if not TableExists("pcFGOG") then
		query="CREATE TABLE pcFGOG ("
		query=query&"pcFO_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"pcFG_ID [INT] DEFAULT(0) NULL,"
		query=query&"idOptionGroup [INT] DEFAULT(0) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcFGOG")
        end if
        set rs=nothing
    end if

	if not TableExists("pcFCAttr") then
		query="CREATE TABLE pcFCAttr ("
		query=query&"pcFA_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"idOption [INT] DEFAULT(0) NULL,"
		query=query&"pcFC_ID [INT] DEFAULT(0) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcFCAttr")
        end if
        set rs=nothing
    end if	

	set rs=nothing
	set connTemp=nothing

	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if

END IF

%>
<!--#include file="AdminHeader.asp"-->
<form action="upddbFacet.asp" method="get" name="form1" class="pcForms">
<%
if mode="complete" then 
	call closeDb()
	response.redirect "ManageFacetGroups.asp"
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
					  <input type="button" class="btn btn-default"  name="access2" value="Update Your ProductCart MS SQL Database" onClick="location='upddbFacet.asp?action=sql';" class="btn btn-primary">
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