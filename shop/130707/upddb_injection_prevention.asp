<%
PmAdmin=19
pageTitle = "ProductCart SQL Injection Prevention Script"
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
			response.redirect "upddb_injection_prevention.asp?mode=3"
			response.End
		end if
		set connTemp=server.createobject("adodb.connection")
		connTemp.Open scDSN
		if err.number <> 0 then
			call closeDb()
			response.redirect "techErr.asp?error="&Server.Urlencode("Error while opening database")
		end if
	else
		if instr(ucase(scDSN),"DSN=") then
			call closeDb()
			response.redirect "upddb_injection_prevention.asp?mode=1"
			response.End
		end if

	end if

	iCnt=0
	ErrStr=""

	'========================================================================
	'// BEGIN SQL INJECTION PREVENTION SCRIPT
	'========================================================================

	'***************************************************************
	' Begin dropping constraints if they exist.
	'***************************************************************
	'Drop constraints for the layout Table

	query="IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_head]')) begin ALTER TABLE Layout DROP CONSTRAINT  CK_Layout_head; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_recal]')) begin ALTER TABLE Layout DROP CONSTRAINT  CK_Layout_recal; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_cont]')) begin ALTER TABLE Layout DROP CONSTRAINT  CK_Layout_cont; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_check]')) begin ALTER TABLE Layout DROP CONSTRAINT  CK_Layout_check; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_subm]')) begin ALTER TABLE Layout DROP CONSTRAINT  CK_Layout_subm; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_more]')) begin ALTER TABLE Layout DROP CONSTRAINT  CK_Layout_more; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_view]')) begin ALTER TABLE Layout DROP CONSTRAINT  CK_Layout_view; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_checko]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_checko; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_addt]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_addt; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_addto]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_addto; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_register]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_register; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_canc]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_canc; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_remo]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_remo; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_add2]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_add2; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_logi]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_logi; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_login]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_login; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_back]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_back; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_regicheck]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_regicheck; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_cust]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_cust; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_recon]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_recon; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_reset]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_reset; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_save]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_save; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_revo]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_revo; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_submit]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_submit; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_reqq]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_reqq; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_place]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_place; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_checkout]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_checkout; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_proce]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_proce; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_final]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_final; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_backto]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_backto; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_previ]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_previ; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_next]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_next; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_crer]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_crer; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_delregistry]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_delregistry; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_addtor]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_addtor; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_updr]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_updr; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_send]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_send; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_retr]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_retr; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_update]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_update; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_Layout_savecart]')) begin ALTER TABLE Layout DROP CONSTRAINT CK_Layout_savecart; end; "

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while Dropping Layout Constraints")
	end if

	'Drop constraints for the customers Table

	query=" IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_name]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_name; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_lastname]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_lastname; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_custcomp]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_custcomp; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_phone]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_phone; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_email]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_email; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_address]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_address; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_zip]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_zip; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_state]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_state; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_city]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_city; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_sadd]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_sadd; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_scity]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_scity; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_sstate]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_sstate; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_szip]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_szip; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_ci1]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_ci1; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_ci2]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_ci2; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_add2]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_add2; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_scomp]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_scomp; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_sadd2]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_sadd2; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_fax]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_fax; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_semail]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_semail; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_vatid]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_vatid; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_sphone]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_sphone; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_sfax]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_sfax; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_consol]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_consol; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_notes]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_notes; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_fbid]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_fbid; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_customers_avalara]')) begin ALTER TABLE customers DROP CONSTRAINT CK_customers_avalara; end; "

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while Dropping customers Constraints")
	end if

	'Drop constraints for the pcSavedCarts Table
	query="IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_pcSavedCarts_SavedCart]')) begin ALTER TABLE pcSavedCarts DROP CONSTRAINT CK_pcSavedCarts_SavedCart; end; "
	query=query & " IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CK_pcSavedCarts_SavedQ]')) begin ALTER TABLE pcSavedCarts DROP CONSTRAINT CK_pcSavedCarts_SavedQ; end; "

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while pcSavedCarts Constraints")
	end if

	'***************************************************************
	'// DROP/CREATE FUNCTION
	'***************************************************************
	query="IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[screen_string_inputs]')) begin Drop Function screen_string_inputs; end; "
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while dropping the function")
	end if

	query="CREATE FUNCTION [dbo].screen_string_inputs(@string VARCHAR(max)) " & _
	"RETURNS int " & _
	"AS " & _
	"BEGIN " & _
	"	DECLARE @rstring nvarchar(max) " & _
	"	DECLARE @strResult int " & _
	"	select @rstring=REPLACE(REPLACE(REPLACE(@string, ' ', '*^'), '^*', ''), '*^', ' '); " & _
	"	IF UPPER(@rstring) like '%SCRIPT SRC%' " & _
	"	BEGIN " & _
	"		SET @strResult=1 " & _
	"	END " & _
	"	ELSE " & _
	"	BEGIN " & _
	"		SET @strResult=0 " & _
	"	END " & _
	"	RETURN (@strResult) " & _
		"END; "
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while creating the function")
	end if


	'***************************************************************
	' Begin adding constraints.
	'***************************************************************
	'layout Table
	query="IF COL_LENGTH('dbo.layout', 'headerid') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_head CHECK ([dbo].screen_string_inputs(headerid)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'recalculate') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_recal CHECK ([dbo].screen_string_inputs(recalculate)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'continueshop') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_cont CHECK ([dbo].screen_string_inputs(continueshop)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'checkout') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_check CHECK ([dbo].screen_string_inputs(checkout)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'submit') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_subm CHECK ([dbo].screen_string_inputs(submit)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'morebtn') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_more CHECK ([dbo].screen_string_inputs(morebtn)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'viewcartbtn') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_view CHECK ([dbo].screen_string_inputs(viewcartbtn)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'checkoutbtn') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_checko CHECK ([dbo].screen_string_inputs(checkoutbtn)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'addtocart') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_addt CHECK ([dbo].screen_string_inputs(addtocart)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'addtowl') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_addto CHECK ([dbo].screen_string_inputs(addtowl)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'register') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_register CHECK ([dbo].screen_string_inputs(register)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'cancel') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_canc CHECK ([dbo].screen_string_inputs(cancel)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'remove') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_remo CHECK ([dbo].screen_string_inputs(remove)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'add2') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_add2 CHECK ([dbo].screen_string_inputs(add2)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'login') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_logi CHECK ([dbo].screen_string_inputs(login)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'login_checkout') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_login CHECK ([dbo].screen_string_inputs(login_checkout)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'back') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_back CHECK ([dbo].screen_string_inputs(back)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'register_checkout') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_regicheck CHECK ([dbo].screen_string_inputs(register_checkout)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'customize') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_cust CHECK ([dbo].screen_string_inputs(customize)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', '[reconfigure]') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_recon CHECK ([dbo].screen_string_inputs([reconfigure])=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'resetdefault') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_reset CHECK ([dbo].screen_string_inputs(resetdefault)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'savequote') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_save CHECK ([dbo].screen_string_inputs(savequote)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'RevOrder') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_revo CHECK ([dbo].screen_string_inputs(RevOrder)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'SubmitQuote') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_submit CHECK ([dbo].screen_string_inputs(SubmitQuote)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_requestQuote') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_reqq CHECK ([dbo].screen_string_inputs(pcLO_requestQuote)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_placeOrder') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_place CHECK ([dbo].screen_string_inputs(pcLO_placeOrder)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_checkoutWR') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_checkout CHECK ([dbo].screen_string_inputs(pcLO_checkoutWR)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_processShip') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_proce CHECK ([dbo].screen_string_inputs(pcLO_processShip)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_finalShip') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_final CHECK ([dbo].screen_string_inputs(pcLO_finalShip)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_backtoOrder') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_backto CHECK ([dbo].screen_string_inputs(pcLO_backtoOrder)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_previous') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_previ CHECK ([dbo].screen_string_inputs(pcLO_previous)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_next') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_next CHECK ([dbo].screen_string_inputs(pcLO_next)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'CreRegistry') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_crer CHECK ([dbo].screen_string_inputs(CreRegistry)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'DelRegistry') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_delregistry CHECK ([dbo].screen_string_inputs(DelRegistry)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'AddToRegistry') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_addtor CHECK ([dbo].screen_string_inputs(AddToRegistry)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'UpdRegistry') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_updr CHECK ([dbo].screen_string_inputs(UpdRegistry)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'SendMsgs') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_send CHECK ([dbo].screen_string_inputs(SendMsgs)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'RetRegistry') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_retr CHECK ([dbo].screen_string_inputs(RetRegistry)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_Update') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_update CHECK ([dbo].screen_string_inputs(pcLO_Update)=0) END; "
	query=query & "IF COL_LENGTH('dbo.layout', 'pcLO_Savecart') IS NOT NULL BEGIN ALTER TABLE Layout WITH NOCHECK ADD CONSTRAINT CK_Layout_savecart CHECK ([dbo].screen_string_inputs(pcLO_Savecart)=0) END; "

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while Adding Layout Constraints")
	end if

	query=" IF COL_LENGTH('dbo.customers', 'name') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_name CHECK ([dbo].screen_string_inputs(name)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'lastName') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_lastname CHECK ([dbo].screen_string_inputs(lastName)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'customerCompany') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_custcomp CHECK ([dbo].screen_string_inputs(customerCompany)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'phone') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_phone CHECK ([dbo].screen_string_inputs(phone)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'email') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_email CHECK ([dbo].screen_string_inputs(email)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'address') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_address CHECK ([dbo].screen_string_inputs(address)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'zip') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_zip CHECK ([dbo].screen_string_inputs(zip)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'state') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_state CHECK ([dbo].screen_string_inputs(state)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'city') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_city CHECK ([dbo].screen_string_inputs(city)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingaddress') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_sadd CHECK ([dbo].screen_string_inputs(shippingaddress)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingcity') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_scity CHECK ([dbo].screen_string_inputs(shippingcity)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingState') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_sstate CHECK ([dbo].screen_string_inputs(shippingState)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingZip') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_szip CHECK ([dbo].screen_string_inputs(shippingZip)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'CI1') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_ci1 CHECK ([dbo].screen_string_inputs(CI1)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'CI2') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_ci2 CHECK ([dbo].screen_string_inputs(CI2)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'address2') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_add2 CHECK ([dbo].screen_string_inputs(address2)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingCompany') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_scomp CHECK ([dbo].screen_string_inputs(shippingCompany)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingAddress2') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_sadd2 CHECK ([dbo].screen_string_inputs(shippingAddress2)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'fax') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_fax CHECK ([dbo].screen_string_inputs(fax)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingEmail') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_semail CHECK ([dbo].screen_string_inputs(shippingEmail)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'pcCust_VATID') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_vatid CHECK ([dbo].screen_string_inputs(pcCust_VATID)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingPhone') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_sphone CHECK ([dbo].screen_string_inputs(shippingPhone)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'shippingFax') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_sfax CHECK ([dbo].screen_string_inputs(shippingFax)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'pcCust_ConsolidateStr') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_consol CHECK ([dbo].screen_string_inputs(pcCust_ConsolidateStr)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'pcCust_Notes') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_notes CHECK ([dbo].screen_string_inputs(pcCust_Notes)=0) END; "
	query=query & " IF COL_LENGTH('dbo.customers', 'pcCust_FBId') IS NOT NULL BEGIN ALTER TABLE customers WITH NOCHECK ADD CONSTRAINT CK_customers_fbid CHECK ([dbo].screen_string_inputs(pcCust_FBId)=0) END; "

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while Adding customers Constraints")
	end if

	'Add constraints to the pcSavedCarts table
	query=" IF COL_LENGTH('dbo.pcSavedCarts', 'SavedCartName') IS NOT NULL BEGIN ALTER TABLE pcSavedCarts WITH NOCHECK ADD CONSTRAINT CK_pcSavedCarts_SavedCart CHECK ([dbo].screen_string_inputs(SavedCartName)=0) END; "
	query=query & " IF COL_LENGTH('dbo.pcSavedCarts', 'SavedCartQuotes') IS NOT NULL BEGIN ALTER TABLE pcSavedCarts WITH NOCHECK ADD CONSTRAINT CK_pcSavedCarts_SavedQ CHECK ([dbo].screen_string_inputs(SavedCartQuotes)=0) END; "

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		response.redirect "techErr.asp?error="&Server.Urlencode("Error while Adding pcSavedCarts Constraints")
	end if
	'========================================================================
	'// END OF SQL INJECTION PREVENTION UPDATES
	'========================================================================
	set rs=nothing
	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if

END IF
%>
<!--#include file="AdminHeader.asp"-->
<form action="upddb_injection_prevention.asp" method="post" name="form1" id="form1" class="pcForms">
<%
if mode="complete" then
	call closeDb()
	response.redirect "upddb_injection_prevention.asp?status=complete"
	response.end()
else
%>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

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

		<%IF request("status")<>"complete" then%>


		<tr>
			<td>

                <h1 class="page-header">Prevent SQL Injection v1.0</h1>
                <p class="lead">
                    This update script is used to help prevent any unauthorized injection attempts on your database.
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
                            Click "Upgrade Now" to update your SQL Injection script to v1.0.
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
                            <input type="button" name="access2" value=" Upgrade Now " onClick="$pc('#form1').submit();" class="btn btn-primary">
                        <% else %>
                            <input type="button" name="access2" value=" Upgrade Now " class="btn btn-primary disabled" disabled>
                        <% end if %>
					</td>
				</tr>
			</table>
			</td>
		</tr>
        <%else %>
        <p>The database has been successfully updated.</p>
		<%END IF%>
	</table>
<% end if %>
</form>
<!--#include file="AdminFooter.asp"-->
