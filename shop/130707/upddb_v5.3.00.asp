<% 
PmAdmin=19
pageTitle = "ProductCart v5.2.1 to v5.3.00 - Database Update" 
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
    '// START:  v5.2.10
    '========================================================================
    
	'query="DROP TABLE pcWidgets"
	'set rs=server.CreateObject("ADODB.RecordSet")
	'set rs=conntemp.execute(query)
    
    If Not TableExists("pcWidgets") Then
		query="CREATE TABLE pcWidgets ("
		query=query&"widget_ID [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL,"
		query=query&"widget_Shortcode [nvarchar] (250),"
		query=query&"widget_Desc [nvarchar] (MAX),"
		query=query&"widget_Type [nvarchar] (250),"
		query=query&"widget_Uri [nvarchar] (MAX),"
		query=query&"widget_Method [nvarchar] (250),"
		query=query&"widget_Lang [nvarchar] (250) "
		query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcWidgets")
        else

            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdAT', 'Add This Zone', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdATC', 'Add To Cart Zone', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdBOM', 'Back-Order Message', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdBrand', 'Brand Name', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdConfig', 'Product Configuration', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('CatTree', 'Category Breadcrumbs', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdCS', 'Cross Selling Zone', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdInput', 'Custom Input Fields', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdNoShip', 'Non-Shipping Item Message', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdBtns', 'Next &amp; Back buttons', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdOSM', 'Out of Stock Message', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdDesc', 'Product Description', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdStock', 'Product Inventory', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdImg', 'Product Images', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdLDesc', 'Product Long Description', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdName', 'Product Name', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdOpt', 'Product Options', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdPrice', 'Product Prices', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdPromo', 'Product Promo Message', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdRate', 'Product Rating', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdRev', 'Product Reviews', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdRP', 'Product Reward Points', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdSearch', 'Product Search Fields', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdSKU', 'Product SKU', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdW', 'Product Weight', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdQDisc', 'Quantity Discounts', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdSB', 'Subscription Bridge', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('PrdWL', 'Wish List Zone', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('CUSTOMHTML', 'Custom HTML Element', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
            query="INSERT INTO pcWidgets (widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('CUSTOMHTML', 'Custom HTML Element', 'Core', '', '', 'ASP');"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            
     
        end if
    End If
    
    If Not TableExists("pcHooks") Then
		query="CREATE TABLE pcHooks ("
		query=query&"hook_ID [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL,"
		query=query&"hook_Shortcode [nvarchar] (250),"
		query=query&"hook_Type [nvarchar] (250),"
		query=query&"hook_Uri [nvarchar] (MAX),"
		query=query&"hook_Method [nvarchar] (250),"
		query=query&"hook_Lang [nvarchar] (250),"
		query=query&"hook_Desc [nvarchar] (MAX), "
		query=query&"hook_Event [nvarchar] (250) "
		query=query&");"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("pcHooks")
        end if
    End If
	
	'// ALTER EXISTING TABLES
	call AlterTableSQL("pcStoreSettings","ADD","pcStoreSettings_EnableGCT","[int]", 1, "0","1")
	call AlterTableSQL("crossSelldata","ADD","cs_showNFS","[int]", 1, "1","1")
	

    '========================================================================
    '// START:  v5.2.10 Patch
    '========================================================================
    
    '// ALTER EXISTING TABLES
	call AlterTableSQL("pcPay_PayPal","ADD","pcPay_PayPal_Layout","[nvarchar](20)", 2, "vertical","0")
	call AlterTableSQL("pcPay_PayPal","ADD","pcPay_PayPal_Shape","[nvarchar](20)", 2, "pill","0")
	call AlterTableSQL("pcPay_PayPal","ADD","pcPay_PayPal_Size","[nvarchar](20)", 2, "medium","0")
	call AlterTableSQL("pcPay_PayPal","ADD","pcPay_PayPal_Color","[nvarchar](20)", 2, "gold","0")
    
    '========================================================================
    '// END:  v5.2.10 Patch
    '========================================================================
    '========================================================================
    '// END:  v5.2.10
    '========================================================================



	'========================================================================
	'// START:  DB UPDATES FOR v5.3.0
	'========================================================================

    '// <removed> MAX convertions

    '// Create table mod_bannermanagement
    if not TableExists("mod_bannermanagement") then
        query="CREATE TABLE [dbo].[mod_bannermanagement]( "
        query=query & " [bannerid] [int] IDENTITY(1,1) NOT NULL, "
        query=query & " [startdate] [datetime] NULL, "
        query=query & " [enddate] [datetime] NULL, "
        query=query & " [active] [int] NULL, "
        query=query & " [background] [nvarchar](50) NULL, "
        query=query & " [html] [nvarchar](max) NULL, "
        query=query & " CONSTRAINT [PK_mod_bannermanagement] PRIMARY KEY CLUSTERED "
        query=query & " ( "
        query=query & " [bannerid] ASC "
        query=query & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] "
        query=query & " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] "

        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            TrapSQLError("mod_bannermanagement")
        end if
        set rs=nothing
    end if


	'// ALTER EXISTING TABLES
	call AlterTableSQL("products","ALTER COLUMN","description","NVARCHAR(500)", 0, "","0")
	call AlterTableSQL("products","ADD","pcProdImage_AltTagText","NVARCHAR(255)", 0, "","0")
	call AlterTableSQL("pcStoreSettings","ADD","pcEnableBulkAdd","INT", 0, "","0")
	call AlterTableSQL("pcSlideShow","ADD","slideStart","DATETIME", 0, "","0")
	call AlterTableSQL("pcSlideShow","ADD","slideEnd","DATETIME", 0, "","0")
	call AlterTableSQL("crossSelldata","ADD","csw_status","[int]", 1, "0","1")
	call AlterTableSQL("authorizeNet", "ADD", "x_accountType", "[INT]", 1, "0", "1")
	

	query="UPDATE pcStoreSettings SET pcEnableBulkAdd=0"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("Settings Column Update - pcEnableBulkAdd FAIL")
	end if
	set rs=nothing
	
	'set default dates for images
	query="UPDATE pcSlideShow SET slideStart='5/31/2018', slideEnd='10/31/2020'"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("Settings Column Add - pcSlideShow FAIL")
	end if
	set rs=nothing					

	'turn off offline credit card
	query= "UPDATE payTypes SET active=0 WHERE gwCode=6"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=conntemp.execute(query)		
	set rs=nothing
	
	'update UPS API URL
	query= "UPDATE ShipmentTypes SET shipserver='https://onlinetools.ups.com/ups.app/xml/Rate' WHERE idShipment=3"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	set rs=nothing
	
	
	'========================================================================
	'// START OF DB UPDATES FOR Canada Post Services
	'========================================================================
	
	query="DELETE FROM shipService WHERE idShipment=7"
	conntemp.execute(query)
	
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'DOM.RP', 'Canada Post Regular Parcel')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'DOM.EP', 'Canada Post Expedited Parcel')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'DOM.XP', 'Canada Post Xpresspost')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'DOM.XP.CERT', 'Canada Post Xpresspost Certified')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'DOM.PC', 'Canada Post Priority')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'DOM.DT', 'Canada Post Delivered Tonight')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'DOM.LIB', 'Canada Post Library Materials')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.EP', 'Canada Post Expedited Parcel USA')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.PW.ENV', 'Canada Post Priority Worldwide Envelope USA')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.PW.PAK', 'Canada Post Priority Worldwide Pak USA')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.PW.PARCEL', 'Canada Post Priority Worldwide Parcel USA')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.SP.AIR', 'Canada Post Small Packet USA Air')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.TP', 'Canada Post Tracked Packet – USA')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.TP.LVM', 'Canada Post Tracked Packet – USA (Large Volume Mailers)')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'USA.XP', 'Canada Post Xpresspost USA')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.XP', 'Canada Post Xpresspost International')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.IP.AIR', 'Canada Post International Parcel Air')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.IP.SURF', 'Canada Post International Parcel Surface')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.PW.ENV', 'Canada Post Priority Worldwide Envelope International')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.PW.PAK', 'Canada Post Priority Worldwide Pak International')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.PW.PARCEL', 'Canada Post Priority Worldwide Parcel International')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.SP.AIR', 'Canada Post Small Packet International Air')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.SP.SURF', 'Canada Post Small Packet International Surface')"
	conntemp.execute(query)
	query="INSERT INTO shipService(idShipment, serviceCode, serviceDescription) VALUES(7, 'INT.TP', 'Canada Post Tracked Packet – International')"
	conntemp.execute(query)
		
	'========================================================================
	'// END OF DB UPDATES FOR Canada Post Services
	'========================================================================
    
	'========================================================================
	'// END:  DB UPDATES FOR v5.3
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
<form action="upddb_v5.3.00.asp" method="post" name="form1" id="form1" class="pcForms">
<%
if mode="complete" then
	call closeDb()
	response.redirect "upddb_v5.3.00_complete.asp?CanUpd=" & CanUpd
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
        
 		<% if len(ErrStr&"") > 0 then %>
			<tr>
				<td align="center">
					<p class="bs-callout bs-callout-danger">
                      Error: <%=ErrStr%>
                    </p>
				</td>
			</tr>
		<% end if %>

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
				</td>v5.3.
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

            <tr>
			    <td>
            
                    <h1 class="page-header">Welcome to ProductCart 5.3.00</h1>
                    <p class="lead">
                        ProductCart 5.3.00 is a full feature release, but also contains miscellaneous bug fixes and improvements for ProductCart v5.2.1. 
                        Be sure to read the <a href="https://productcart.desk.com/customer/portal/articles/2959589-updating-productcart-v5-0-to-v5-3-00" target="_blank">v5.3.00 Update Guide</a>.
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
                            <li class="list-group-item"><span class="glyphicon glyphicon-check checklist"></span>&nbsp;&nbsp;Review the <a href="https://productcart.desk.com/customer/portal/articles/2959589-updating-productcart-v5-0-to-v5-3-00" target="_blank">v5.3.00 Update Guide</a> and <a href="https://productcart.desk.com/customer/portal/articles/2957357-productcart-5-3-change-log" target="_blank">Change Log</a>.</li>
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

	</table>
<% end if %>
</form>
<!--#include file="AdminFooter.asp"-->

<%
    
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
    %>
