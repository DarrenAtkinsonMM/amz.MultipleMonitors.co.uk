<% 
PmAdmin=19
pageTitle = "ProductCart Defender - Database Update" 
Section = "" 
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="fixedNTextConst.asp"-->
<% 
On Error Resume Next
dim conntemp1

IF request("action")="sql" then
	if request("hmode")="2" then
		SSIP=request("SSIP")
		UID=request("UID")
		PWD=request("PWD")
		SSDB=request("SSDB")
		if SSIP="" or UID="" or PWD="" then
			call closeDb()
            response.redirect "upddb_v5.1.00.asp?mode=3"
			response.End
		end if
		set connTemp=server.createobject("adodb.connection")
		connTemp.Open "Provider=sqloledb;Data Source="&SSIP&";Initial Catalog="&SSDB&";User Id="&UID&";Password="&PWD
		if err.number <> 0 then
			call closeDb()
            response.redirect "techErr.asp?error="&Server.Urlencode("Error while opening database")
		end if
	else
		if instr(ucase(scDSN),"DSN=") then
			call closeDb()
            response.redirect "upddb_v5.1.00.asp?mode=1"
			response.End
		end if
		
	end if
	
	iCnt=0
	ErrStr=""

	'========================================================================
	'// START:  DB UPDATES
	'========================================================================	

	query="CREATE TABLE [dbo].[pcDefinitions] ("
	query=query&"[pcDef_Id] [int] NULL  DEFAULT (1),"
	query=query&"[pcDef_Key] [nvarchar] (25) NULL ,"
	query=query&"[pcDef_Desc] [nvarchar] (100) NULL ,"
	query=query&"[pcDef_Pattern] [nvarchar] (500) NULL ,"
	query=query&"[pcDef_Replace] [nvarchar] (100) NULL ,"
	query=query&"[pcDef_IsGlobal] [int] NULL DEFAULT(0) ,"
    query=query&"[pcDef_IgnoreCase] [int] NULL DEFAULT(0) ,"
	query=query&"[pcDef_Type] [nvarchar] (10) NULL ,"
	query=query&"[pcDef_ContinueOnError] [int] NULL DEFAULT(0) ,"
	query=query&"[pcDef_Priority] [int] NULL DEFAULT(0) ,"
	query=query&"[pcDef_Active] [int] NULL DEFAULT(0) "
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcDefinitions")
	end if

    call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_AdminLastLogin","[datetime]",2,"1/1/2013","0")
    
    Dim pcv_strIsOverRide
    pcv_strIsOverRide = 1
    call pcs_updateDefinitions()

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
<form action="upddbPCD.asp" method="post" name="form1" id="form1" class="pcForms">
<%
if mode="complete" then
%>
<table class="pcCPcontent">
    <tr>
		<td>
            <div class="pcCPmessageSuccess">
                Your store database was successfully updated. 
            </div>            
        </td>
    </tr>
    <tr>
  	    <td class="pcCPspacer"></td>
    </tr>
</table>
<%
else 
%>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

		<% if (session("PmAdmin")<>"19") then %>
			<tr>
				<td align="center">
					<div class="pcCPmessage">
                        Full administrative permissions are required to complete the database update.  Please logout and log back in with the admin user account.
                    </div>
				</td>
			</tr>
		<% end if %>
		<% if mode="errors" then %>
			<tr>
				<td align="center">
					<div class="pcCPmessage">The following errors occurred while updating the store database. Try running the database update script again. If the errors persist, please open a support ticket:
                        <br><br>
					    <%=ErrStr%>
                    </div>
				</td>
			</tr>
		<% end if %>


		<tr>
			<td>
            
                <h1 class="page-header">Welcome to ProductCart Defender</h1>
                <p class="lead">
                    ProductCart Defender is a new feature that helps to protect your store against threats. Be sure to read the <a href="https://productcart.desk.com/customer/portal/articles/2142176-critical-security-patch-september-30-2015-" target="_blank">Update Guide</a> prior to updating.
                </p>

				
					<% 
                    dim findit
                    if PPD="1" then
                        PageName="/"&scPcFolder&"/includes/diagtxt.txt"
                    else
                        PageName="../includes/diagtxt.txt"
                    end if
                    findit=Server.MapPath(PageName)
                    
                    Dim fso, f, errpermissions, errdelete_includes, errwrite_includes, errwrite_others
                    errpermissions=0
                    errdelete_includes=0
                    errwrite_includes=0
                    errwrite_others=0
                    Set fso=server.CreateObject("Scripting.FileSystemObject")
                    Set f=fso.GetFile(findit)
                    Err.number=0
                    f.Delete
                    if Err.number>0 then
                        errdelete_includes=1
                        errpermissions=1
                        Err.number=0
                    end if
                    'Set f=nothing
                                
                    Set f=fso.OpenTextFile(findit, 2, True)
                    f.Write "test done"
                    if Err.number>0 then
                        errwrite_includes=1
                        errpermissions=1
                        Err.number=0
                    end if
                    
                    if PPD="1" then
                        PageName="/"&scPcFolder&"/pc/diagtxt.txt"
                    else			
                        PageName="../pc/diagtxt.txt"
                    end if
                    findit=Server.MapPath(PageName)
                    Set f=fso.OpenTextFile(findit, 2, True)
                    f.Write "test done"
                    if Err.number>0 then
                        errwrite_others=1
                        errpermissions=1
                        Err.number=0
                    end if
                                
                    f.Close
                    Set fso=nothing
                    Set f=nothing
                    if errpermissions=0 then %>
 
					<% else %>
                    
                        <div class="pcCPmessageWarning">
                        
                        <h2>Please correct these issues before you begin:</h2>

                        <% if scDB<>"SQL" then %> 
                            <table>
                                <tr> 
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">ProductCart v5 only works with MS SQL databases.  The Access database is been deprecated for security and performance reasons.  <a href="https://www.productcart.com/support/v5/article.asp?id=3" target="_blank">Click here</a> to ask for a quote to convert your Access database to SQL.</font></td>
                                </tr>
                            </table>
                        <% end if %>
                        
					    <% if errwrite_others=1 or errwrite_includes=1 then %> 
                            <table>
                                <tr> 
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">You need to assign 'read/write' permissions to the 'productcart' folder and all of its subfolders.</font></td>
                                </tr>
                            </table>
						<% end if

                            if errdelete_includes=1 then 
                                %>
                                <table>
                                <tr> 
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">You need to assign 'read/write/delete' permissions to the 'productcart/includes' folder and all of its subfolders.</font></td>
                                </tr> 
                            </table>
                            <% 
                            end if
                            %>
                            </div>
                            <%
				    end if 
                    %>
                
                    <div class="bs-callout bs-callout-info">
                        <h4>Read Me</h4>
                        <p>
                            Click "Update Now" to update your MS SQL Database.  
                        </p>
                    </div>              
    
                    <div class="bs-callout bs-callout-warning">
                        <h4>Backup Your Database</h4>
                        <p>
                            Although we have tested this update script in a variety of environments, there is always the possibility of something going wrong. 
                            Make sure to <span style="font-weight: bold">backup your database</span> prior to executing this update.
                            Depending on how the database has been setup, you may be able to either perform the backup yourself or have your Web hosting company do it for you. 
                            Note: Your SQL database is likely being automatically backed up every day: confirm that this is the case by asking your Web host when the last back up occurred.
                        </p>
                    </div>

			<table class="pcCPcontent" width="80%">
			<% if request.querystring("mode")="1" OR request.querystring("mode")="3" then %>
				<tr>
					<td>
						It appears that you are using a DSN connection to connect to your SQL server. In order to complete this update, please enter your SQL Server Information below:
						<% if request.querystring("mode")="1" then %>
							<br>
							<strong>*All fields are required.</strong>
						<% end if %>

						<input name="hmode" type="hidden" id="hmode" value="2">	
					</td>
				</tr>
				<tr>
					<td>Server Domain/IP: <input name="SSIP" type="text" id="SSIP" size="30"></td>
				</tr>
				<tr>
					<td>Database Name: <input name="SSDB" type="text" id="SSDB" size="30"></td>
				</tr>
				<tr>
					<td>User ID: <input name="UID" type="text" id="UID" size="30"></td>
				</tr>
				<tr>
					<td>Password: <input name="PWD" type="password" id="PWD" size="30"></td>
				</tr>

			<% end if %>
				<tr>
					<td align="center">
						<input name="action" type="hidden" id="action" value="sql">

                        <% if errpermissions=0 then %>
                                <input type="button" name="access2" value=" Update Now " onClick="$pc('#form1').submit();" class="btn btn-primary">
                        <% else %>
                                <input type="button" name="access2" value=" Update Now " class="btn btn-primary disabled" disabled>
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