<% 
PmAdmin=19
pageTitle = "ProductCart v5.1.01 - Database Update" 
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

'IF request("action")="sql" then
	call opendb()
	
	iCnt=0
	ErrStr=""


	'========================================================================
	'// START:  CHECK FOR APPAREL FIELDS
	'========================================================================
    
    If request("action")="sql" Then
    
        '========================================================================
        '// START:  APPAREL
        '========================================================================
       
        '// Create table pcApparelSettings
        if not TableExists("pcApparelSettings") then
            query="CREATE TABLE [dbo].[pcApparelSettings] ("
            query=query&"[pcAS_ID] [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL,"
            query=query&"[pcAS_HideUItems] [int] NULL DEFAULT(1) ,"
            query=query&"[pcAS_PriceDiff] [int] NULL DEFAULT(0) ,"
            query=query&"[pcAS_TurnWB] [int] NULL DEFAULT(0) ,"
            query=query&"[pcAS_WMsg] [nvarchar] (250) NULL "
            query=query&");"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if

        if err.number <> 0 then
            TrapSQLError("pcApparelSettings")
        else
            query="INSERT INTO pcApparelSettings (pcAS_HideUItems,pcAS_PriceDiff,pcAS_TurnWB,pcAS_WMsg) VALUES (1,0,0,'Please wait...');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
        '// Add Products column pcprod_AppDefault if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_AppDefault] [INT] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_AppDefault]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Products column pcprod_Apparel if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_Apparel] [INT] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description& " (" & err.number & ")<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_Apparel]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Products column pcprod_ParentPrd if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_ParentPrd] [INT] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_ParentPrd]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Products column pcprod_Relationship if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_Relationship] [nvarchar](max) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Products column pcprod_ShowStockMsg if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_ShowStockMsg] [INT] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_ShowStockMsg]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Products column pcprod_StockMsg if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_StockMsg] [nvarchar](max) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Products column pcprod_SizeLink if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_SizeLink] [nvarchar](max) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Products column pcprod_SizeInfo if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_SizeInfo] [nvarchar](max) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Products column pcprod_SizeImg if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_SizeImg] [nvarchar] (150) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Products column pcprod_SizeURL if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_SizeURL] [nvarchar](max) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Products column pcprod_AddPrice if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_AddPrice] [float] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
    
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_AddPrice]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Products column pcprod_SentNotice if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_SentNotice] [INT] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_SentNotice]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Options column pcOpt_Img if doesn't exist
        query="ALTER TABLE [Options] ADD [pcOpt_Img] [nvarchar] (150) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Options - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Options column pcOpt_Code if doesn't exist
        query="ALTER TABLE [Options] ADD [pcOpt_Code] [nvarchar] (150) NULL;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Options - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    
    
        '// Add Products column pcProd_SPInActive if doesn't exist
        query="ALTER TABLE [products] ADD [pcProd_SPInActive] [INT] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcProd_SPInActive]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Products column pcprod_AddWPrice if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_AddWPrice] [float] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_AddWPrice]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    
    
        '// Add Products column pcprod_ApparelRadio if doesn't exist
        query="ALTER TABLE [products] ADD [pcprod_ApparelRadio] [INT] NULL DEFAULT '0';"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            if err.number = -2147217900 then    ' "COLUMN NAMES IN EACH TABLE MUST BE UNIQUE"
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE Products - Error: "&Err.Description&"<BR>"
                err.number=0
                iCnt=iCnt+1
            end if
        else
            query="UPDATE [products] SET [pcprod_ApparelRadio]=0;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
        
        
        '// Start converting ntext to nvarchar(MAX)
        call AlterTableSQL("products","ALTER COLUMN","pcprod_Relationship ","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("products","ALTER COLUMN","pcprod_StockMsg","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("products","ALTER COLUMN","pcprod_SizeLink","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("products","ALTER COLUMN","pcprod_SizeInfo","[nvarchar](max)", 0, "","0")
        call AlterTableSQL("products","ALTER COLUMN","pcprod_SizeURL","[nvarchar](max)", 0, "","0")        
        '// Sales Manager
        SavedFile = "SalesManager_APP.sql"
        findit = Server.MapPath(Savedfile)
        Set fso = server.CreateObject("Scripting.FileSystemObject")
        Err.number=0
        Set f = fso.OpenTextFile(findit, 1)
        Flines = f.ReadAll
        f.close
        Set f=nothing
        Set fso=nothing
    
        tmp1=split(Flines,"GO" & vbcrlf)
        For i=0 to ubound(tmp1)
            if trim(tmp1(i))<>"" then
                set rs=connTemp.execute(tmp1(i))
                Err.number=0
            end if
            set rs=nothing
        Next
        
        '========================================================================
        '// END: APPAREL
        '======================================================================== 
        
    End If

    If request("action")<>"skip" Then
        err.clear
        pcv_IsApparelError = False
        query="SELECT TOP 1 pcprod_AppDefault FROM [dbo].[products]"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            iCnt = iCnt + 1
            pcv_IsApparelError = True
        end if
        err.clear
    End If

	'========================================================================
	'// END:  CHECK FOR APPAREL FIELDS
	'========================================================================


	'========================================================================
	'// START:  PRODUCTCART DEFENDER
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
    
	query="DELETE FROM [dbo].[pcDefinitions]"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
    
    call pcs_updateDefinitions()
    
	'========================================================================
	'// END:  PRODUCTCART DEFENDER
	'========================================================================


	'========================================================================
	'// START:  DB UPDATES FOR v5.1.01
	'========================================================================

	'// ALTER EXISTING TABLES    
    call AlterTableSQL("orders","ALTER COLUMN","pcOrd_CVNResponse","[nvarchar](50)", 0, "","0")
	call AlterTableSQL("pcSales","ALTER COLUMN","pcSales_Desc","[nvarchar](max)", 0, "","0")
	call AlterTableSQL("pcSales_Completed","ALTER COLUMN","pcSC_SaveDesc","[nvarchar](max)", 0, "","0")
    
    '// PAYEEZY UPDATE
	query="ALTER TABLE [pcPay_Payeezy] ADD [pcPEY_JSKey] [nvarchar](250)"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Payeezy - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if
	
	query="ALTER TABLE [pcPay_Payeezy] ADD [pcPEY_TAToken] [nvarchar](250)"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Payeezy - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if


	call closedb()

	'========================================================================
	'// END:  DB UPDATES FOR v5.1.01
	'========================================================================

	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if

'END IF
%>
<!--#include file="AdminHeader.asp"-->
<%if mode="complete" then
	call closeDb()
	response.redirect "upddb_v5.1.01_complete.asp"
	response.end()	
else%>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>
        
		<% if mode="errors" And ErrStr<>"" then %>
			<tr>
				<td align="center">
					<p class="bs-callout bs-callout-danger">
                        The following errors occurred while updating the store database. Try running the database update script again. If the errors persist, please open a support ticket:
                        <br><br>
					    <%=ErrStr%>
                    </p>
				</td>
			</tr>
		<% end if %>
        
		<% if mode="errors" And pcv_IsApparelError = True then %>
			<tr>
				<td align="center">
					<p class="bs-callout bs-callout-danger">
                        We detected a database error. Click the "Fix Database Errors" button below to repair the problem. If you receive the error message again, then use the "Continue without Fix" button and open a support ticket.
                        
                        <br /><br />

                        <input name="action" type="button" class="btn btn-success"  id="action" value=" Fix Database Errors " class="btn btn-primary" onclick="javascript:location='upddb_v5.1.01.asp?action=sql';">
                        
                        <input name="action" type="button" class="btn btn-default"  id="action" value=" Continue without Fix " class="btn btn-primary" onclick="javascript:location='upddb_v5.1.01.asp?action=skip';">
                        
                    </p>
				</td>
			</tr>
		<% end if %>
        
	</table>
<%end if%>

<!--#include file="AdminFooter.asp"-->
