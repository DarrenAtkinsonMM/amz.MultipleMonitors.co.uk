<% 
PmAdmin=19
pageTitle = "ProductCart v5.x to v5.2.00 - Database Update" 
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

If request("action")="sql" Then

	iCnt=0
	ErrStr=""


	'========================================================================
	'// START:  CHECK FOR APPAREL FIELDS
	'========================================================================
    
    If request("action")="sql2" Then
    
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
	'// START:  DB UPDATES FOR v5.0
	'========================================================================

    '// <removed> MAX convertions

    '// Create table gwAmazon
    if not TableExists("gwAmazon") then
        query="CREATE TABLE gwAmazon ("
        query=query&"gwAMZ_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
        query=query&"gwAMZ_SellerID [nvarchar] (250) NULL, "
        query=query&"gwAMZ_AccessKey [nvarchar] (250) NULL, "
        query=query&"gwAMZ_SecretKey [nvarchar] (250) NULL, "
        query=query&"gwAMZ_ClientID [nvarchar] (250) NULL, "
        query=query&"gwAMZ_ClientSecret [nvarchar] (250) NULL, "
        query=query&"gwAMZ_Mode [int] NULL DEFAULT(0) ,"
        query=query&"gwAMZ_TestMode [int] NULL DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("gwAmazon")
        end if
        set rs=nothing
    end if

    '// Create table pcShipwireSettings
    if not TableExists("pcShipwireSettings") then
        query="CREATE TABLE pcShipwireSettings ("
        query=query&"pcSWS_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
        query=query&"pcSWS_UserName [nvarchar] (250) NULL ,"
        query=query&"pcSWS_Password [nvarchar] (250) NULL ,"
        query=query&"pcSWS_OnOff [int] NULL DEFAULT(0) ,"
        query=query&"pcSWS_Mode [int] NULL DEFAULT(0) ,"
        query=query&"pcSWS_SyncDate [datetime] NULL "
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcShipwireSettings")
        end if
        set rs=nothing
    end if

    if not TableExists("pcShipwireOrders") then
        query="CREATE TABLE pcShipwireOrders ("
        query=query&"pcSWO_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
        query=query&"idOrder [int] NULL DEFAULT(0) ,"
        query=query&"pcSWO_ShipwireID [nvarchar] (250) NULL ,"
        query=query&"pcSWO_ShipwireDetails [varchar] (8000) NULL "
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcShipwireOrders")
        end if
        set rs=nothing
    end if
    
    if not TableExists("pcContactPageSettings") then
        query="CREATE TABLE pcContactPageSettings ("
        query=query&"pcCPage_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
        query=query&"pcCPage_PageDesc [nvarchar](max) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcContactPageSettings")
        end if
    end if
    
    if not TableExists("pcPrdXFields") then
        query="CREATE TABLE pcPrdXFields ("
        query=query&"pcPXF_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
        query=query&"IdProduct [int] NULL DEFAULT(0) ,"
        query=query&"IdXfield [int] NULL DEFAULT(0) ,"
        query=query&"pcPXF_XReq [int] NULL DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcPrdXFields")
        end if
        set rs=nothing
    end if

    call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_DisplayQuickView","[int]",1,"0","0")
    call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_AdminLastLogin","[datetime]",2,"1/1/2013","0")

    call AlterTableSQL("categories","ADD","pcCats_ProductOrder","[nvarchar] (4)", 0, "","0")

    call AlterTableSQL("Products","ADD","pcProd_Top","[nvarchar] (800)", 0, "","0")
    call AlterTableSQL("Products","ADD","pcProd_TopLeft","[nvarchar] (800)", 0, "","0")
    call AlterTableSQL("Products","ADD","pcProd_TopRight","[nvarchar] (800)", 0, "","0")
    call AlterTableSQL("Products","ADD","pcProd_Middle","[nvarchar] (800)", 0, "","0")
    call AlterTableSQL("Products","ADD","pcProd_Bottom","[nvarchar] (800)", 0, "","0")
    call AlterTableSQL("Products","ADD","pcProd_Tabs","[nvarchar] (max)", 0, "","0")

    call AlterTableSQL("paypal","ADD","PP_PaymentAction","[int]", 1, "1","0")

    call AlterTableSQL("pcPay_PFL_Authorize","ADD","gwCode","[int]", 1, "1","0")
    call AlterTableSQL("pcPay_PFL_Authorize","ADD","fraudcode","[int]", 1, "1","0")
    
    '// Default Product Layout
    if not TableExists("pcDefaultPrdLayout") then
        query="CREATE TABLE pcDefaultPrdLayout ("
        query=query&"pcDPL_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
        query=query&"pcDPL_idProduct [INT] NULL ,"
        query=query&"pcDPL_Name [nvarchar] (255) NULL,"
        query=query&"pcDPL_Top [nvarchar] (800) NULL,"
        query=query&"pcDPL_TopLeft [nvarchar] (800) NULL,"
        query=query&"pcDPL_TopRight [nvarchar] (800) NULL,"
        query=query&"pcDPL_Middle [nvarchar] (800) NULL,"
        query=query&"pcDPL_Bottom [nvarchar] (800) NULL,"
        query=query&"pcDPL_Tabs [nvarchar] (max) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcDefaultPrdLayout")

            '// Add new column if the table already exists
            call AlterTableSQL("pcDefaultPrdLayout","ADD","pcDPL_Middle","[nvarchar] (800)", 0, "","0")
            call AlterTableSQL("pcDefaultPrdLayout","ADD","pcDPL_Name","[nvarchar] (255)", 0, "","0")
            call AlterTableSQL("pcDefaultPrdLayout","ADD","pcDPL_idProduct","[int]", 0, "","0")
        end if
    end if
    
    '// Slideshow Feature
    if not TableExists("pcSlideShow") then
        query="CREATE TABLE pcSlideShow ("
        query=query&"idSlide [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
        query=query&"slideImage [nvarchar](255) NOT NULL,"
        query=query&"slideCaption [nvarchar](MAX) NULL,"
        query=query&"slideUrl [nvarchar](500) NULL,"
        query=query&"slideAlt [nvarchar](255) NULL,"
        query=query&"slideOrder [int] NULL,"
        query=query&"slideDateUploaded datetime NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcSlideShow")
        else
        end if
        set rs=nothing
    end if

    call AlterTableSQL("pcSlideShow","ADD","idSetting","[int]",0,"","0")

    '// Slideshow Settings
    if not TableExists("pcSlideShowSettings") then
        query="CREATE TABLE pcSlideShowSettings ("
        query=query&"id [int] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
        query=query&"slideWidth [int] NOT NULL,"
        query=query&"slideHeight [int] NOT NULL,"
        query=query&"effect [nvarchar](50) NULL,"
        query=query&"pauseTime [int] NULL,"
        query=query&"animSpeed [int] NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcSlideShowSettings")
        else
        end if
        set rs=nothing
    end if

    call AlterTableSQL("pcSlideShowSettings","ADD","idSetting","[int]",0,"","0")
    call AlterTableSQL("pcSlideShowSettings","ADD","useDefault","[int]",1,"0","0")

    '// Add default slideshow configuration
    '<removed>
    
    '// Add slideshow config for the mobile
    '<removed>
    
    '// Accepted Payments
    if not TableExists("pcAcceptedPayments") then
        query="CREATE TABLE pcAcceptedPayments ("
        query=query&"pcAcceptedPayment_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
        query=query&"pcAcceptedPayment_Name [nvarchar](50) NOT NULL,"
        query=query&"pcAcceptedPayment_Image [nvarchar](200) NOT NULL,"
        query=query&"pcAcceptedPayment_CustomImage [nvarchar](200) NULL,"
        query=query&"pcAcceptedPayment_Alt [nvarchar](255) NULL,"
        query=query&"pcAcceptedPayment_Active [bit] NULL,"
        query=query&"pcAcceptedPayment_Order [int] NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcAcceptedPayments")
        end if
        set rs=nothing
    end if

    '// Add default accepted payments
    '<removed>
    
    '// Google Trusted Store
    if not TableExists("pcGoogleTS") then
        query="CREATE TABLE pcGoogleTS ("
        query=query&"pcGTS_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
        query=query&"pcGTS_TurnOn [INT] NULL DEFAULT(0),"
        query=query&"pcGTS_AccNo [nvarchar](50) NULL,"
        query=query&"pcGTS_PageLang [nvarchar](50) NULL,"
        query=query&"pcGTS_ShopAccID [nvarchar](50) NULL,"
        query=query&"pcGTS_ShopCountry [nvarchar](50) NULL,"
        query=query&"pcGTS_ShopLang [nvarchar](50) NULL,"
        query=query&"pcGTS_Currency [nvarchar](5) NULL,"
        query=query&"pcGTS_ShipDays [INT] NULL DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcGoogleTS")
        end if
        set rs=nothing
    end if
    
    call AlterTableSQL("pcGoogleTS","ADD","pcGTS_DeDays","[int]", 1, "0","0")
    
    '// Social Links
    if not TableExists("pcSocialLinks") then
        query="CREATE TABLE pcSocialLinks ("
        query=query&"pcSocialLink_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
        query=query&"pcSocialLink_Name [nvarchar](50) NOT NULL,"
        query=query&"pcSocialLink_Image [nvarchar](200) NOT NULL,"
        query=query&"pcSocialLink_CustomImage [nvarchar](200) NULL,"
        query=query&"pcSocialLink_Url [nvarchar](500) NULL,"
        query=query&"pcSocialLink_Alt [nvarchar](255) NULL,"
        query=query&"pcSocialLink_Order [int] NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcSocialLinks")
        end if
        set rs=nothing
    end if

    '// Add default social links
    '<removed>
    
    '// Create table pcFacebookSettings
    if not TableExists("pcFacebookSettings") then
        query="CREATE TABLE [pcFacebookSettings] ("
        query=query&"[pcFBS_id] [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL ,"
        query=query&"[pcFBS_TurnOnOff] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_OffMsg] [nvarchar] (400) NULL ,"
        query=query&"[pcFBS_AppID] [nvarchar] (100) NULL ,"
        query=query&"[pcFBS_RedirectURL] [nvarchar] (250) NULL ,"
        query=query&"[pcFBS_Header] [nvarchar] (max) NULL ,"
        query=query&"[pcFBS_Footer] [nvarchar] (max) NULL ,"
        query=query&"[pcFBS_PageWidth] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_CustomDisplay] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_CatImages] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_CatRow] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_CatRowsPerPage] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_BType] [nvarchar] (5) NULL ,"
        query=query&"[pcFBS_PrdRow] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_PrdRowsPerPage] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_ShowSKU] [int] NULL DEFAULT(0) ,"
        query=query&"[pcFBS_ShowSmallImg] [int] NULL DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcFacebookSettings")
        end if
        set rs=nothing
    end if

    '========================================================================
    '// START: MOBILE
    '======================================================================== 
    
    '// Create table pcMobileSettings
    if not TableExists("pcMobileSettings") then
        query="CREATE TABLE pcMobileSettings ("
        query=query&"pcMS_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL,"
        query=query&"pcMS_Logo [nvarchar] (250) NULL ,"
        query=query&"pcMS_ShowHomeNav [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_ShowHomeSP [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_ShowHomeNA [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_ShowHomeBS [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_ShowHomeFP [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_ShowNavTop [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_ShowNavBot [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_IsApparelAddOn [INT] NULL DEFAULT(0) ,"
        query=query&"pcMS_PayPalCardTypes [nvarchar] (50) NULL"
        query=query&");"
        set rs=conntemp.execute(query)
        set rs=nothing
    end if
    if err.number <> 0 then
        TrapSQLError("pcMobileSettings")
    end if


    call AlterTableSQL("pcMobileSettings","ADD","pcMS_Pay","[int]",1,"0","0")
    call AlterTableSQL("pcMobileSettings","ADD","pcMS_TurnOn","[int]",1,"0","0")
    call AlterTableSQL("Orders","ADD","pcOrd_MobileSF","[int]",1,"0","0")
    
    '========================================================================
    '// END: MOBILE
    '======================================================================== 


    call AlterTableSQL("Customers","ADD","pcCust_FBId","[nvarchar] (100)", 0, "","0")
    call AlterTableSQL("Customers","ADD","pcCust_AmazonId","[nvarchar] (200)", 0, "","0")
    
    call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_PNButtons","[int]",1,"1","0")
    call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_ConURL","[int]",1,"0","0")
    call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_GAType","[int]",1,"0","0")
    call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_ThemeFolder","[nvarchar] (100)",0,"","0")

    '// Add shipping service shipment ID
    call AlterTableSQL("shipService","ADD","idShipment","[int]",0,"","1")
    
    '// Create table pcPackageLabel
    if not TableExists("pcPackageLabel") then
        query="CREATE TABLE pcPackageLabel ("
        query=query&"pcPackageLabel_ID [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL,"
        query=query&"pcPackageInfo_ID [INT] NOT NULL ,"
        query=query&"pcPackageLabel_Name [nvarchar] (100) NOT NULL ,"
        query=query&"pcPackageLabel_File [nvarchar] (255) NOT NULL ,"
        query=query&"pcPackageLabel_FileType [nvarchar] (50) NULL ,"
        query=query&"pcPackageLabel_Resolution [INT] NULL ,"
        query=query&"pcPackageLabel_Type [nvarchar] (50) NULL ,"
        query=query&"pcPackageLabel_Date [datetime] NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcPackageLabel")
        end if
        set rs=nothing
    end if

    '// UPS
    query = "UPDATE shipService SET idShipment = 3 WHERE serviceCode IN ('01','02','03','07','08','11','12','13','14','54','59','65')"
    set rs=conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
    set rs=nothing

    '// USPS
    query = "UPDATE shipService SET idShipment = 4 WHERE serviceCode IN ('9901','9902','9903','9904','9905','9906','9907','9908','9909','9910','9911','9912','9914','9915','9916','9917')"
    set rs=conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
    set rs=nothing
    
    '// Change "USPS Parcel" to "USPS Standard Post"
    query="UPDATE shipService SET serviceDescription = 'USPS Standard Post<sup>&reg;</sup>' WHERE serviceCode = '9903';"
    set rs=conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
    set rs=nothing

    '// Canada Post
    query = "UPDATE shipService SET idShipment = 7 WHERE serviceCode IN ('1010','1020','1130','1030','1040','1120','1220','1230','2010','2020','2030','2040','2050','3010','3020','3040','2005','2015','2025','3005','3015','3025','3050')"
    set rs=conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
    set rs=nothing
'// FedEx Web Services
    query = "UPDATE shipService SET idShipment = 9 WHERE serviceCode IN ('FIRST_OVERNIGHT','FEDEX_FIRST_FREIGHT','PRIORITY_OVERNIGHT','STANDARD_OVERNIGHT','FEDEX_2_DAY','FEDEX_2_DAY_AM','FEDEX_EXPRESS_SAVER','FEDEX_FREIGHT_PRIORITY','FEDEX_FREIGHT_ECONOMY','FEDEX_GROUND','GROUND_HOME_DELIVERY','INTERNATIONAL_GROUND','INTERNATIONAL_FIRST','INTERNATIONAL_PRIORITY','INTERNATIONAL_ECONOMY','FEDEX_1_DAY_FREIGHT','FEDEX_2_DAY_FREIGHT','FEDEX_3_DAY_FREIGHT','INTERNATIONAL_PRIORITY_FREIGHT','INTERNATIONAL_ECONOMY_FREIGHT','FEDEX_FREIGHT','FEDEX_NATIONAL_FREIGHT','SMART_POST','FEDEX_ECONOMY_CANADA','EUROPE_FIRST_INTERNATIONAL_PRIORITY')"
    set rs=conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
    set rs=nothing
    
    '// Add missing FedEx shipping services
    '<removed>

    '// Update incorrect country names
    query="UPDATE countries SET countryName = 'Aland Islands' WHERE countryCode = 'AX'"
    conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if

    query="UPDATE countries SET countryName = 'Cote D''Ivoire' WHERE countryCode = 'CI'"
    conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if

    query="UPDATE countries SET countryName = 'Curacao' WHERE countryCode = 'CW'"
    conntemp.execute(query)
    if err.number<>0 then
        err.number=0
    end if

    query="UPDATE countries SET countryName = 'Reunion' WHERE countryCode = 'RE'"
    conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
    
    query="UPDATE countries SET countryName = 'Saint Barthelemy' WHERE countryCode = 'BL'"
    conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if

    query="UPDATE countries SET countryName = 'Viet Nam' WHERE countryCode = 'VN'"
    conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if

    '// Delete non-existent countries
    query="DELETE FROM countries WHERE countryCode = 'AN'"
    conntemp.execute(query)
    if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
    Err.Description=""



    '// Update icon resources
    call UpdateTableIfValue("pcRevSettings", "pcRS_Img1", "", "smileygreen.gif", "smileygreen.png")
    call UpdateTableIfValue("pcRevSettings", "pcRS_Img2", "", "smileyred.gif", "smileyred.png")
    call UpdateTableIfValue("pcRevSettings", "pcRS_Img3", "", "fullstar.gif", "fullstar.png")
    call UpdateTableIfValue("pcRevSettings", "pcRS_Img4", "", "halfstar.gif", "halfstar.png")
    call UpdateTableIfValue("pcRevSettings", "pcRS_Img5", "", "emptystar.gif", "emptystar.png")

    call UpdateTableIfValue("icons", "discount", "where id=1", "images/sample/pc_icon_discount.gif", "images/sample/pc_icon_discount.png")
    call UpdateTableIfValue("icons", "erroricon", "where id=1", "images/sample/pc_icon_error.gif", "images/sample/pc_icon_error.png")
    call UpdateTableIfValue("icons", "zoom", "where id=1", "images/sample/pc_icon_zoom.gif", "images/sample/pc_icon_zoom.png")

    '// Create table pcUpdateLog
    if not TableExists("pcUpdateLog") then
        query="CREATE TABLE pcUpdateLog ("
        query=query&"id [INT] IDENTITY (1, 1) PRIMARY KEY CLUSTERED NOT NULL,"
        query=query&"name [nvarchar] (1000) NULL ,"
        query=query&"filename [nvarchar] (1000) NULL ,"
        query=query&"date_installed [datetime] NULL DEFAULT (GETDATE()) ,"
        query=query&"notes [nvarchar] (max) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcUpdateLog")
        end if
        set rs=nothing
    end if
		
	
	
	'========================================================================
	'// END:  DB UPDATES FOR v5.0
	'========================================================================



	'========================================================================
	'// START:  DB UPDATES FOR v5.02
	'========================================================================

    call AlterTableSQL("emailSettings","ADD","FontSize","[nvarchar](10)", 2, "13px","0")

	'========================================================================
	'// END:  DB UPDATES FOR v5.02
	'========================================================================



	'========================================================================
	'// START:  DB UPDATES FOR v5.03
	'========================================================================

    call AlterTableSQL("pcPay_PFL_Authorize","ADD","fraudcode","[nvarchar](250)", 1, "1","0")
    call AlterTableSQL("products","ADD","pcprod_AppDefault","[int]", 1, "1","0")
	call AlterTableSQL("pcSavedCarts","ADD","SavedCartQuotes","[nvarchar](500)", 0, "","0")
	call AlterTableSQL("categories","ADD","pcCats_NotImg","[int]", 1, "0","0")
       
    '// Mobile Settings
    if not TableExists("pcMobileSettings") then
        query="CREATE TABLE pcMobileSettings ("
        query=query&"pcMS_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
        query=query&"pcMS_Logo [nvarchar](250) NULL,"
        query=query&"pcMS_ShowHomeNav [INT] NULL DEFAULT(0),"
        query=query&"pcMS_ShowHomeSP [INT] DEFAULT(0),"
        query=query&"pcMS_ShowHomeNA [INT] DEFAULT(0),"
        query=query&"pcMS_ShowHomeBS [INT] DEFAULT(0),"
        query=query&"pcMS_ShowHomeFP [INT] DEFAULT(0),"
        query=query&"pcMS_ShowNavTop [INT] DEFAULT(0),"
        query=query&"pcMS_ShowNavBot [INT] DEFAULT(0),"
        query=query&"pcMS_IsApparelAddOn [INT] NULL DEFAULT(0),"        
        query=query&"pcMS_PayPalCardTypes [nvarchar](50) NULL,"
        query=query&"pcMS_Pay [INT] NULL DEFAULT(0),"
        query=query&"pcMS_TurnOn [INT] NULL DEFAULT(0) "
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcMobileSettings")
        end if
        set rs=nothing
    end if

	'========================================================================
	'// END:  DB UPDATES FOR v5.03
	'========================================================================




	'========================================================================
	'// START:  DB UPDATES FOR UNIFIED BUILD
	'========================================================================
	query="CREATE TABLE [dbo].[pcBTORules] ("
	query=query&"[pcBR_ID] [int] IDENTITY (1, 1) NOT NULL,"
	query=query&"[pcBR_IDBTOPrd] [int] NULL DEFAULT '0',"
	query=query&"[pcBR_IDSourcePrd] [int] NULL DEFAULT '0',"
	query=query&"[pcBR_isCAT] [int] NULL DEFAULT '0',"
	query=query&"[pcBR_Must_Exists] [int] NULL DEFAULT '0',"
	query=query&"[pcBR_CanNot_Exists] [int] NULL DEFAULT '0',"
	query=query&"[pcBR_CatMust_Exists] [int] NULL DEFAULT '0',"
	query=query&"[pcBR_CatCanNot_Exists] [int] NULL DEFAULT '0'"
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"ALREADY AN OBJECT NAMED") then
			Err.Description=""
			err.number=0
		else
			response.write "Error Creating table pcBTORules: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if

	set rs=nothing
	
	'**** Create table pcBRMust ***************************************
	query="CREATE TABLE [dbo].[pcBRMust] ("
	query=query&"[pcBRMust_ID] [int] IDENTITY (1, 1) NOT NULL,"
	query=query&"[pcBR_ID] [int] NULL DEFAULT '0',"
	query=query&"[pcBRMust_Item] [int] NULL DEFAULT '0'"
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"ALREADY AN OBJECT NAMED") then
			Err.Description=""
			err.number=0
		else
			response.write "Error Creating table pcBRMust: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if

	set rs=nothing
	
	'**** Create table pcBRCanNot ***************************************
	query="CREATE TABLE [dbo].[pcBRCanNot] ("
	query=query&"[pcBRCanNot_ID] [int] IDENTITY (1, 1) NOT NULL,"
	query=query&"[pcBR_ID] [int] NULL DEFAULT '0',"
	query=query&"[pcBRCanNot_Item] [int] NULL DEFAULT '0'"
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"ALREADY AN OBJECT NAMED") then
			Err.Description=""
			err.number=0
		else
			response.write "Error Creating table pcBRCanNot: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if

	set rs=nothing
	
	'**** Create table pcBRCatMust ***************************************
	query="CREATE TABLE [dbo].[pcBRCatMust] ("
	query=query&"[pcBRCatMust_ID] [int] IDENTITY (1, 1) NOT NULL,"
	query=query&"[pcBR_ID] [int] NULL DEFAULT '0',"
	query=query&"[pcBRCatMust_Item] [int] NULL DEFAULT '0'"
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"ALREADY AN OBJECT NAMED") then
			Err.Description=""
			err.number=0
		else
			response.write "Error Creating table pcBRCatMust: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if

	set rs=nothing
	
	'**** Create table pcBRCatCanNot ***************************************
	query="CREATE TABLE [dbo].[pcBRCatCanNot] ("
	query=query&"[pcBRCatCanNot_ID] [int] IDENTITY (1, 1) NOT NULL,"
	query=query&"[pcBR_ID] [int] NULL DEFAULT '0',"
	query=query&"[pcBRCatCanNot_Item] [int] NULL DEFAULT '0'"
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"ALREADY AN OBJECT NAMED") then
			Err.Description=""
			err.number=0
		else
			response.write "Error Creating table pcBRCatCanNot: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if
	
	'**** Add new field into the table Products ***************************************
	query="ALTER TABLE [Products] ADD [pcProd_ShowBTOCMMsg] [int] NULL;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("configSpec_products")
	else
		query="UPDATE products SET pcProd_ShowBTOCMMsg=0;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		query="UPDATE products SET pcProd_ShowBTOCMMsg=1 WHERE serviceSpec<>0;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		query="ALTER TABLE [products] ADD CONSTRAINT [DF_products_pcProd_ShowBTOCMMsg] DEFAULT (0) FOR [pcProd_ShowBTOCMMsg];"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if
	'========================================================================
	'// END:  DB UPDATES FOR UNIFIED BUILD
	'========================================================================



	'========================================================================
	'// START:  DB UPDATES FOR v5.1.00
	'========================================================================

    '// Payeezy Gateway
	if not TableExists("pcPay_Payeezy") then
		query="CREATE TABLE pcPay_Payeezy ("
		query=query&"pcPEY_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"pcPEY_MerchantID [nvarchar](250) NULL,"
		query=query&"pcPEY_MToken [nvarchar](250) NULL,"
		query=query&"pcPEY_APIKey [nvarchar](250) NULL,"
		query=query&"pcPEY_APISKey [nvarchar](250) NULL,"
        query=query&"pcPEY_Mode [INT] DEFAULT(0),"
        query=query&"pcPEY_TestMode [INT] DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcPay_Payeezy")
        end if
        set rs=nothing
    end if
	
	'========================================================================
	'// END:  DB UPDATES FOR v5.1.00
	'========================================================================



	'========================================================================
	'// START:  DB UPDATES FOR v5.1.01
	'========================================================================

    call AlterTableSQL("orders","ALTER COLUMN","pcOrd_CVNResponse","[nvarchar](50)", 0, "","0")
	call AlterTableSQL("pcSales","ALTER COLUMN","pcSales_Desc","[nvarchar](max)", 0, "","0")
	call AlterTableSQL("pcSales_Completed","ALTER COLUMN","pcSC_SaveDesc","[nvarchar](max)", 0, "","0")

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

	'========================================================================
	'// END:  DB UPDATES FOR v5.1.01
	'========================================================================



	'========================================================================
	'// START:  Images Alt Tag Text
	'========================================================================

	call AlterTableSQL("products", "ADD", "pcProd_AdditionalImages", "[INT]", 1, "0", "1")
	call AlterTableSQL("products", "ADD", "pcProd_AltTagText", "[NVarChar](255)", 0, "", "0")
	call AlterTableSQL("pcProductsImages", "ADD", "pcProdImage_AltTagText", "[NVarChar](255)", 0, "", "0")

	'========================================================================
	'// END:  Images Alt Tag Text
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
    
    Dim pcv_strIsOverRide
    pcv_strIsOverRide = True
    
    call pcs_updateDefinitions()
    
	'========================================================================
	'// END:  PRODUCTCART DEFENDER
	'========================================================================



	'========================================================================
	'// START:  DB UPDATES FOR v5.2
	'========================================================================

    '// Change "USPS Standard Post" to "USPS Retail Ground"
    query="UPDATE shipService SET serviceDescription = 'USPS Retail Ground<sup>&trade;</sup>' WHERE serviceCode = '9903';"
    set rs=conntemp.execute(query)
	
	'// Replace "®" sign that causing a question mark character issue, with HTML format "&reg;"
	query="UPDATE shipService SET serviceDescription = 'UPS Next Day Air<sup>&reg;</sup>' WHERE serviceCode='01';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS 2<sup>nd</sup> Day Air<sup>&reg;</sup>' WHERE serviceCode='02';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS Ground<sup>&reg;</sup>' WHERE serviceCode='03';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS Standard To Canada<sup>&reg;</sup>' WHERE serviceCode='11';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS 3 Day Select<sup>&reg;</sup>' WHERE serviceCode='12';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS Next Day Air Saver<sup>&reg;</sup>' WHERE serviceCode='13';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS Next Day Air<sup>&reg;</sup> Early A.M.<sup>&reg;</sup>' WHERE serviceCode='14';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS 2<sup>nd</sup> Day Air A.M.<sup>&reg;</sup>' WHERE serviceCode='59';"
	set rs=conntemp.execute(query)
	query="UPDATE shipService SET serviceDescription = 'UPS Express Saver<sup>&reg;</sup>' WHERE serviceCode='65';"
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
        err.number=0
        Err.Description=""
    end if
	set rs=nothing

     '// pcReCaSettings
	if not TableExists("pcReCaSettings") then
		query="CREATE TABLE pcReCaSettings ("
		query=query&"pcRCS_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"pcRCS_SiteKey [nvarchar](100) NULL,"
		query=query&"pcRCS_Secret [nvarchar](100) NULL,"
		query=query&"pcRCS_Theme [nvarchar](50) NULL,"
		query=query&"pcRCS_Type [nvarchar](50) NULL,"
		query=query&"pcRCS_Size [nvarchar](50) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcReCaSettings")
        end if
        set rs=nothing
    end if
	
	'// pcUsedPassHistory
	if not TableExists("pcUsedPassHistory") then
		query="CREATE TABLE pcUsedPassHistory ("
		query=query&"pcUP_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"IdCustomer [INT] DEFAULT(0),"
		query=query&"pcUP_UsedPass [nvarchar](400) NULL,"
		query=query&"pcUP_IPAddress [nvarchar](50) NULL,"
		query=query&"pcUP_CreatedDate [datetime] NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcUsedPassHistory")
        end if
        set rs=nothing
    end if
	
	'// pcPassResetHistory
	if not TableExists("pcPassResetHistory") then
		query="CREATE TABLE pcPassResetHistory ("
		query=query&"pcPassID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"IdCustomer [INT] DEFAULT(0),"
		query=query&"pcPassResetGuid [nvarchar](400) NULL,"
		query=query&"pcPassResetTimeout [datetime] NULL,"
		query=query&"pcPassResetIPAddress [nvarchar](50) NULL,"
		query=query&"pcPassResetTime [datetime] NULL,"
		query=query&"pcPassResetSuccess [INT] DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcUsedPassHistory")
        end if
        set rs=nothing
    end if
	
	'// pcPassResetHistory
	if not TableExists("pcPassResetHistory") then
		query="CREATE TABLE pcPassResetHistory ("
		query=query&"pcPassID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"IdCustomer [INT] DEFAULT(0),"
		query=query&"pcPassResetGuid [nvarchar](400) NULL,"
		query=query&"pcPassResetTimeout [datetime] NULL,"
		query=query&"pcPassResetIPAddress [nvarchar](50) NULL,"
		query=query&"pcPassResetTime [datetime] NULL,"
		query=query&"pcPassResetSuccess [INT] DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcUsedPassHistory")
        end if
        set rs=nothing
    end if
	
	'// pcLoginHistory
	if not TableExists("pcLoginHistory") then
		query="CREATE TABLE pcLoginHistory ("
		query=query&"pcLH_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"IdCustomer [INT] DEFAULT(0),"
		query=query&"pcLH_IPAddress [nvarchar](50) NULL,"
		query=query&"pcLH_DateTime [datetime] NULL,"
		query=query&"pcLH_Failed [INT] DEFAULT(0)"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcLoginHistory")
        end if
        set rs=nothing
    end if
	
	'// pcShippingMap
	if not TableExists("pcShippingMap") then
		query="CREATE TABLE pcShippingMap ("
		query=query&"pcSM_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"pcSM_Name [nvarchar](400) NULL,"
		query=query&"pcSM_Type [INT] DEFAULT(0) NULL,"
		query=query&"pcSM_Order [INT] DEFAULT(0) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcShippingMap")
        end if
        set rs=nothing
    end if
	
	'// pcSMRel
	if not TableExists("pcSMRel") then
		query="CREATE TABLE pcSMRel ("
		query=query&"pcSMR_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"pcSM_ID [INT] DEFAULT(0) NULL,"
		query=query&"idshipservice [INT] DEFAULT(0) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcSMRel")
        end if
        set rs=nothing
    end if

	'// ALTER EXISTING TABLES
	call AlterTableSQL("Customers", "ADD", "pcCust_LockUntil", "[DateTime]", 0, "", "0")	
	call AlterTableSQL("Customers", "ADD", "pcCust_LockMinutes","[INT]","1","0","0")	

	call AlterTableSQL("pcStoreSettings", "ADD", "pcStoreSettings_DispDiscCart","[INT]","1","0","0")
	call AlterTableSQL("pcStoreSettings", "ADD", "pcStoreSettings_KeepSession","[INT]","1","0","0")
	call AlterTableSQL("pcStoreSettings", "ADD", "pcStoreSettings_BTOShowMustIm","[INT]","1","0","0")
	
	call AlterTableSQL("pcSales","ALTER COLUMN","pcSales_Desc","[nvarchar](max)", 0, "","0")
	call AlterTableSQL("pcSales_Completed","ALTER COLUMN","pcSC_SaveDesc","[nvarchar](max)", 0, "","0")

	'========================================================================
	'// END:  DB UPDATES FOR v5.2
	'========================================================================



	'========================================================================
	'// START: Advanced Security
	'========================================================================
    if not TableExists("pcTransactionLogs") then
        query="CREATE TABLE pcTransactionLogs ("
        query=query&"id [int] IDENTITY (1,1) PRIMARY KEY NOT NULL,"
        query=query&"datetime [datetime] NULL,"
		query=query&"IP [nvarchar] (50) NOT NULL,"
		query=query&"customerId [int] NOT NULL,"
		query=query&"orderId [int] NOT NULL,"
        query=query&"isSuccess [bit] NOT NULL,"
		query=query&"gatewayId [int] NOT NULL"    
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcTransactionLogs")
        end if
        set rs=nothing
    else
        call AlterTableSQL("pcTransactionLogs","ALTER COLUMN","IP ","[nvarchar](50)", 0, "","0")
    end if 
    
    call AlterTableSQL("Customers", "ADD", "pcCust_FailedPaymentCount","[INT]","1","0","0")
	'========================================================================
	'// END: Advanced Security
	'========================================================================




	'========================================================================
	'// START: Theme Chooser
	'========================================================================

	'// Create table pcThemes
    if not TableExists("pcThemes") then
        query="CREATE TABLE pcThemes ("
        query=query&"pcThemes_Id [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL ,"
        query=query&"pcThemes_Name [nvarchar] (100) NOT NULL ,"
        query=query&"pcThemes_Active [bit] NULL DEFAULT (0) ,"
        query=query&"pcThemes_DateUploaded [datetime] NULL"
        query=query&");"
        err.clear
        on error resume next
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            Err.Description=""
            err.number=0
        end if
    end if
  
	'// Add default themes
    call pcs_IndexThemeFolder()
    
    '// Save new theme settings file
    If len(scThemeFolder)>0 Then
        call pcs_SaveThemeToSettings(scThemeFolder)
    End If

	'========================================================================
	'// END: Theme Chooser
	'========================================================================



	'========================================================================
	'// START OF DB UPDATES FOR Facet Search
	'========================================================================
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

	Function TableExists(tableName)
        query="IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'" & tableName & "') SELECT 1 ELSE SELECT 0;"
        set rs=conntemp.execute(query)
		If rs(0) = "1" Then
            TableExists = true
        Else
            TableExists = false
        End If
		set rs=nothing
        err.clear
        err.number = 0
    End Function
	'========================================================================
	'// END OF DB UPDATES FOR Facet Search
	'========================================================================



	'========================================================================
	'// START OF DB UPDATES FOR ProductCart Apps
	'========================================================================

	'// Create table pcWebServiceSettings
    if not TableExists("pcWebServiceSettings") then

        query="CREATE TABLE pcWebServiceSettings ("
        query=query&"pcPCWS_Id [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL ,"
        query=query&"pcPCWS_TurnOnOff [INT] DEFAULT(0) NULL,"
        query=query&"pcPCWS_IsActive [INT] DEFAULT(0) NULL,"
        query=query&"pcPCWS_Url [nvarchar] (250) NULL ,"
        query=query&"pcPCWS_Uid [nvarchar] (250) NULL ,"
        query=query&"pcPCWS_Username [nvarchar] (250) NULL ,"
        query=query&"pcPCWS_Fullname [nvarchar] (250) NULL ,"
        query=query&"pcPCWS_Email [nvarchar] (250) NULL ,"
        query=query&"pcPCWS_EmailConfirmed [INT] DEFAULT(0) NULL,"
        query=query&"pcPCWS_Level [INT] DEFAULT(0) NULL,"
        query=query&"pcPCWS_JoinDate [nvarchar] (250) NULL ,"
        query=query&"pcPCWS_LicenseKey [nvarchar] (250) NULL ,"
        query=query&"pcPCWS_AuthToken [nvarchar] (MAX) NULL ,"
        query=query&"pcPCWS_Password [nvarchar] (250) NULL "
        query=query&");"
        err.clear
        on error resume next
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            Err.Description=""
            err.number=0
        end if
    end if
    
	'// Create table pcWebServiceFeatures
    if not TableExists("pcWebServiceFeatures") then

        query="CREATE TABLE pcWebServiceFeatures ("
        query=query&"pcPCWS_FeatureId [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL ,"
        query=query&"pcPCWS_FeatureCode [nvarchar] (50) NULL ,"
        query=query&"pcPCWS_IsActive [INT] DEFAULT(0) NULL,"
        query=query&"pcPCWS_IsEnabled [INT] DEFAULT(0) NULL,"
        query=query&"pcPCWS_IsProvisioned [INT] DEFAULT(0) NULL "
        query=query&");"
        err.clear
        on error resume next
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            Err.Description=""
            err.number=0
        end if
    end if
    
	'========================================================================
	'// END OF DB UPDATES FOR ProductCart Apps
	'========================================================================



	'========================================================================
	'// START OF DB UPDATES FOR Facet Search
	'========================================================================
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
	'========================================================================
	'// END OF DB UPDATES FOR Facet Search
	'========================================================================



	'========================================================================
	'// START OF DB UPDATES FOR Enable Bundling and Javascript Optimization
	'========================================================================

	'//  Add column pcStoreSettings_EnableBundling for table "pcStoreSettings"
	call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_EnableBundling","[int]",1,"0","0")
	
	'//  Add column pcStoreSettings_OptimizeJavascript for table "pcStoreSettings"
	call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_OptimizeJavascript","[int]",1,"0","0")

	'========================================================================
	'// END OF DB UPDATES FOR Enable Bundling and Javascript Optimization
	'========================================================================



	'========================================================================
	'// START OF DB UPDATES FOR CartStack
	'========================================================================
	

	'//  Add column pcStoreSettings_CartStack for table "pcStoreSettings"
	call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_CartStack","[int]",1,"0","0")

	
	'//  Add column pcStoreSettings_CSSiteId for table "pcStoreSettings"
	call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_CSSiteId","[nvarchar](20)", 0, "","0")

	'========================================================================
	'// END OF DB UPDATES FOR CartStack
	'========================================================================



	'========================================================================
	'// START:  Google Tag Manager
	'========================================================================

	'//  Add column pcStoreSettings_GoogleTagManager for table "pcStoreSettings"
	call AlterTableSQL("pcStoreSettings", "ADD", "pcStoreSettings_GoogleTagManager", "[NVarChar](50)", 0, "", "0")	

	'========================================================================
	'// END:  Google Tag Manager
	'========================================================================



	'========================================================================
	'// START:  Sales Manager for Apparel Add-On
	'========================================================================
	
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
	'// END:  Sales Manager for Apparel Add-On
	'========================================================================



	'========================================================================
	'// START:  Avalara
	'========================================================================
    
	'// Create table Avalara_orders
	if not TableExists("Avalara_orders") then
		query="CREATE TABLE Avalara_orders ("
		query=query&"id [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL ,"
		query=query&"idOrder [int] DEFAULT(0) NOT NULL,"
        query=query&"idOrderCounter [float] DEFAULT(0) NOT NULL,"
		query=query&"status [nvarchar] (20),"
		query=query&"updatedDate [datetime] NULL"
		query=query&");"
		err.clear
		on error resume next
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			Err.Description=""
			err.number=0
		end if
    else
        call AlterTableSQL("Avalara_orders", "ADD", "idOrderCounter", "float", 1, 0, "1")
	end if
    
    call AlterTableSQL("orders", "ADD", "pcOrd_Avalara", "int", 1, 0, "1")

    call AlterTableSQL("customers", "ADD", "pcCust_AvalaraExemptionNo","[nvarchar](20)", 0, "","0")
	
    call AlterTableSQL("products", "ADD", "pcProd_AvalaraTaxCode","[nvarchar](20)", 0, "","0")

    'call AlterTableSQL("categories", "ADD", "pcProd_AvalaraTaxCode","[nvarchar](20)", 0, "","0")

    call AlterTableSQL("pcCustomerSessions", "ADD", "pcCustSession_Avalara", "int", 1, 0, "1")
    
    call AlterTableSQL("categories", "ADD", "pcCats_AvalaraTaxCode","[nvarchar](20)", 0, "","0")


	'========================================================================
	'// END:  Avalara
	'========================================================================



	'========================================================================
	'// START:  Control Panel Misc.
	'========================================================================
    call AlterTableSQL("pcStoreSettings", "ADD", "pcStoreSettings_SPhoneReq", "[tinyint]", 1, 0, "1")
    call AlterTableSQL("pcSlideShow", "ADD", "slideNewWindow", "int", 1, 0, "1")
	'========================================================================
	'// END:  Control Panel Misc.
	'========================================================================



	'========================================================================
	'// START:  PAYEEZY
	'========================================================================
	if not TableExists("pcPayeezyLogs") then
		query="CREATE TABLE pcPayeezyLogs ("
		query=query&"pcPEYLg_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
		query=query&"idOrder [INT] DEFAULT(0) NULL,"
		query=query&"idCustomer [INT] DEFAULT(0) NULL,"
		query=query&"pcPEYLg_Status [INT] DEFAULT(0) NULL,"
		query=query&"pcPEYLg_TransID [nvarchar](250) NULL,"
		query=query&"pcPEYLg_TransTag [nvarchar](250) NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcPayeezyLogs")
        end if
        set rs=nothing
    end if
	'========================================================================
	'// END:  PAYEEZY
	'======================================================================== 



	'========================================================================
	'// START:  AUTHORIZE.NET DPM
	'========================================================================
	If Not TableExists("pcPay_AuthorizeDPM") Then
        query="CREATE TABLE [dbo].[pcPay_AuthorizeDPM] ("
        query=query&"[pcAuNet_id] [int] IDENTITY(1,1) NOT NULL,"
        query=query&"[id] [int] NULL,"
        query=query&"[x_Type] [nvarchar](50) NULL,"
        query=query&"[x_Login] [nvarchar](100) NULL,"
        query=query&"[x_Password] [nvarchar](100) NULL,"
        query=query&"[x_version] [nvarchar](4) NULL,"
        query=query&"[x_Curcode] [nvarchar](4) NULL,"
        query=query&"[x_Method] [nvarchar](4) NULL,"
        query=query&"[x_DPMType] [nvarchar](50) NULL,"
        query=query&"[x_CVV] [int] NULL,"
        query=query&"[x_testmode] [int] NULL,"
        query=query&"[x_eCheck] [int] NULL,"
        query=query&"[x_secureSource] [int] NULL,"
        query=query&"[x_eCheckPending] [int] NULL"
        query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcPay_AuthorizeDPM")
        else
            query="INSERT INTO pcPay_AuthorizeDPM (id, x_Type, x_Login, x_Password, x_version, x_Curcode, x_Method, x_DPMType, x_CVV, x_testmode) VALUES (1, 'AUTH_ONLY', 'testdriver', 'testdriver', '3.1', 'USD', 'DPM', 'PASS', 0, 1);"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
        end if
    End If
	'========================================================================
	'// END:  AUTHORIZE.NET DPM
	'========================================================================

	%>
		<!-- #include file="pcAdminRetrieveSettings.asp" -->
	<%
	pcIntScUpgrade = 0
	%>
		<!-- #include file="pcAdminSaveSettings.asp" -->
	<%

	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if

End If
%>
<!--#include file="AdminHeader.asp"-->
<form action="upddb_v5.2.asp" method="post" name="form1" id="form1" class="pcForms">
<%
if mode="complete" then
	call closeDb()
	response.redirect "upddb_v5.2_complete.asp?CanUpd=" & CanUpd
	response.end()	
else
%>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>
        
        <%
        pcv_boolIsStoreClosed = True
        If (scStoreOff = "0") Then
            'pcv_boolIsStoreClosed = False
            'pcv_boolIsError = 1
        End If
        %> 
        
        <%
        pcv_boolIsNET = False
		tmpResult = pcf_PasswordHash("!2500LmB!..")
        If (instr(tmpResult, "NSPC") = 0) Then
            pcv_boolIsNET = True
            pcv_boolIsError = 1
        End If
        %>       

        <%
        pcv_boolIs2005 = False
        query="SELECT @@VERSION AS 'Version';"        
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        If Not rs.Eof Then
            pcv_strVersion = rs("Version")
            If instr(pcv_strVersion, "Microsoft SQL Server 2005")>0 Then
                pcv_boolIs2005 = True            
            End If
        End If
        set rs=Nothing
        %>
        
 		<% if (pcv_boolIs2005 = True) then %>
			<tr>
				<td align="center">
					<p class="bs-callout bs-callout-danger">
                        It looks like you are using SQL 2000.  ProductCart v5 requires 2005 or greater.
                    </p>
				</td>
			</tr>
		<% end if %>

		<% if (session("PmAdmin")<>"19") then %>
			<tr>
				<td align="center">
					<p class="bs-callout bs-callout-danger">
                        Full administrative permissions are required to complete the database update.  Please logout and log back in with the admin user account.
                    </p>
				</td>
			</tr>
		<% end if %>


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


		<% if request("s")="88" then %>
			<tr>
				<td align="center">
					<div class="pcCPmessageSuccess">Updated SQL database successfully to use the data type: 'Nvarchar(Max)' instead of 'NText'</div>
				</td>
			</tr>
		<% end if %>


        <% If mode="errors" And pcv_IsApparelError = True Then %>
			<tr>
				<td align="center">
					<p class="bs-callout bs-callout-danger">
                        We detected a database error. Click the "Fix Database Errors" button below to repair the problem. If you receive the error message again, then use the "Continue without Fix" button and open a support ticket.
                        
                        <br /><br />

                        <input name="action" type="button" class="btn btn-success"  id="action" value=" Fix Database Errors " class="btn btn-primary" onclick="javascript:location='upddb_v5.2.asp?action=apparel';">
                        
                        <input name="action" type="button" class="btn btn-default"  id="action" value=" Continue without Fix " class="btn btn-primary" onclick="javascript:location='upddb_v5.2.asp?action=skip';">
                        
                    </p>
				</td>
			</tr>
		<% End If %>


		<% If scFixedNText = 0 Then %>
			<tr>
				<td align="center">
					<p>
                        From ProductCart v5.0, we don't use the field data type: 'NText' anymore for store database because the next versisons of MS SQL Server won't support it.
                        <br />
					    You need to update store database to use the data type: 'Nvarchar(Max)' instead of 'NText'.
					</p>
					<br><br>
					<input name="fixntext" type="button" class="btn btn-default"  id="fixntext" value="Update Your ProductCart MS SQL Database" class="btn btn-primary" onclick="javascript:location='upddb_fixNtext.asp';">
					<br><br>					
				</td>
			</tr>
		<% Else %>

            <tr>
			    <td>
            
                    <h1 class="page-header">Welcome to ProductCart 5.2.00</h1>
                    <p class="lead">
                        ProductCart 5.2.00 is a full feature release, but also contains miscellaneous bug fixes and improvements for ProductCart v5.1.00. 
                        Be sure to read the <a href="https://productcart.desk.com/customer/portal/articles/2492501-updating-productcart-v5-1-0x-to-v5-2-00" target="_blank">v5.2.00 Update Guide</a>.
                    </p>
                    <br />
                    <br />
					<% 
                    dim findit
                    if PPD="1" then
                        PageName="/"&scPcFolder&"/includes/diagtxt.txt"
                    else
                        PageName="../includes/diagtxt.txt"
                    end if
                    findit=Server.MapPath(PageName)
                    
                    Dim fso, f, pcv_boolIsError, errdelete_includes, errwrite_includes, errwrite_others
                    errdelete_includes=0
                    errwrite_includes=0
                    errwrite_others=0
                    Set fso=server.CreateObject("Scripting.FileSystemObject")
                    Set f=fso.GetFile(findit)
                    Err.number=0
                    f.Delete
                    if Err.number>0 then
                        errdelete_includes=1
                        pcv_boolIsError=1
                        Err.number=0
                    end if
                    'Set f=nothing
                                
                    Set f=fso.OpenTextFile(findit, 2, True)
                    f.Write "test done"
                    if Err.number>0 then
                        errwrite_includes=1
                        pcv_boolIsError=1
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
                        pcv_boolIsError=1
                        Err.number=0
                    end if
                                
                    f.Close
                    Set fso=nothing
                    Set f=nothing
                    if pcv_boolIsError=0 then %>
 
					<% else %>
                    
                        <div class="pcCPmessageWarning">
                        
                        <h3>Please correct these issues before you begin:</h3>
                        
                        <% if (pcv_boolIsStoreClosed = False) then %>
                            <table>
                                <tr>
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">
                                            Your store must be closed during the update.  You can close the store using the <a href="AdminSettings.asp" target="_blank">Store Settings</a>.
                                    </font></td>
                                </tr>
                            </table>
                        <% end if %> 
                        
                        <% if (pcv_boolIsNET = True) then %>
                            <table>
                                <tr>
                                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">
                                            It looks like you don't have ASP.NET installed, or there is a problem.
                                            Please contact technical support and open a ticket.
                                    </font></td>
                                </tr>
                            </table>
                        <% end if %> 

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


                    <div class="bs-callout bs-callout-info upgrade">
                        <h4>Pre-Update Review</h4>
                        Please take a moment to go over the following pre-update action items.
                        <ul class="list-group">
                            <li class="list-group-item"><span class="glyphicon glyphicon-check checklist"></span>&nbsp;&nbsp;Create a backup of the entire "store" folder.</li>
                            <li class="list-group-item"><span class="glyphicon glyphicon-check checklist"></span>&nbsp;&nbsp;Create a backup of the database.</li>
                            <li class="list-group-item"><span class="glyphicon glyphicon-check checklist"></span>&nbsp;&nbsp;Review the <a href="https://productcart.desk.com/customer/portal/articles/2492501-updating-productcart-v5-1-0x-to-v5-2-00" target="_blank">v5.2.00 Update Guide</a> and <a href="https://productcart.desk.com/customer/portal/articles/2492904-productcart-v5-2-00-change-log" target="_blank">Change Log</a>.</li>
                        </ul>
                    </div>              


                    <div class="bs-callout bs-callout-warning upgrade">
                        <h4>Important Notes</h4>
                        <ul class="list-group">
                            <li class="list-group-item"><span class="glyphicon glyphicon-thumbs-up defaultlist"></span>&nbsp;&nbsp;v5.2.00 leverages ASP.NET.  We've already checked and you have it!</li>
                            <li class="list-group-item"><span class="glyphicon glyphicon-bookmark warninglist"></span>&nbsp;&nbsp;Just in case you missed it, here's a bookmark to an <a href="https://productcart.desk.com/customer/portal/articles/2492502-v5-2-update-before-you-begin" target="_blank">important message</a>.</li>    
                            <li class="list-group-item"><span class="glyphicon glyphicon-cloud infolist"></span>&nbsp;&nbsp;After you update you'll find ProductCart Apps under the "Settings" menu.</li>
                        </ul>
                    </div>


                    <table class="pcCPcontent" width="80%">
                            <% If request.querystring("mode")="1" Or request.querystring("mode")="3" Then %>
                                <tr>
                                    <td>
                                        It appears that you are using a DSN connection to connect to your SQL server. In order to complete this update, please enter your SQL Server Information below:
                                        <% If request.querystring("mode")="1" Then %>
                                            <br>
                                            <strong>*All fields are required.</strong>
                                        <% End If %>            
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
                            <% End If %>
                            <tr>
                                <td align="center">
                                    <input name="action" type="hidden" id="action" value="sql">
            
                                    <% if pcv_boolIsError=0 then %>
                                            <input type="button" name="access2" value=" Update Now " onClick="$pc('#form1').submit();" class="btn btn-primary">
                                    <% else %>
                                            <input type="button" name="access2" value=" Update Now " class="btn btn-default disabled" disabled>
                                    <% end if %>
                                </td>
                            </tr>
                    </table>


			    </td>
		    </tr>
		<%END IF%>
	</table>
<% end if %>
</form>
<!--#include file="AdminFooter.asp"-->
