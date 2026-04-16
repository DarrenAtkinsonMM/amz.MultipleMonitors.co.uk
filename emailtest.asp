<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%private const scEmail="sales@multiplemonitors.co.uk"
private const scFrmEmail="website@multiplemonitors.uk"
private const scCustServEmail="sales@multiplemonitors.co.uk"
private const scEmailComObj="CDOSYS"
private const scSMTP="guk1022.siteground.eu"
private const scLocalOrRemote="2"
private const scPort="465"
private const scSMTPAuthentication="Y"
private const scSMTPUID="website@multiplemonitors.uk"
private const scSMTPPWD="1Ci1~g1y#@@{"
private const scEmailFontSize="13px"
private const scConfirmEmail="Dear <CUSTOMER_NAME><br><br>We wanted to let you know that order number <ORDER_ID> that you placed on <TODAY_DATE> has been processed and will be shipped soon.<br><br>This is your order confirmation. Order details are listed below.<br><br>If you have any questions, please do not hesitate to contact us."
private const scReceivedEmail="Dear <CUSTOMER_NAME><br><br>Thank you for shopping at <COMPANY>.<br><br>We received your order on <TODAY_DATE>. Your order number is <ORDER_ID>.<br><br>Note that this is not an order confirmation. You will receive a detailed confirmation message once your order has been processed. You can check the status of your order by logging into your account at <COMPANY_URL>/shop/pc/custpref.asp<br><br>If you have any questions, please do not hesitate to contact us.<br><br>Thank you for being a <COMPANY> customer.<br><br>Best Regards,<br><COMPANY>"
private const scShippedEmail="Dear <CUSTOMER_NAME><br><br>We thought you may like to know that your order number <ORDER_ID> has been shipped. Shipping details are listed below.<br><br>If you have any questions, please do not hesitate to contact us."
private const scNoticeNewCust="0"
private const scCancelledEmail="This message is to inform you that order number <ORDER_ID> that you submitted in this store on <ORDER_DATE> has been cancelled."
Dim mail 
	Dim iConf 
	Dim Flds
	Dim localOrRemote 

	on error resume next 
	
	localOrRemote = scLocalOrRemote
	if(localOrRemote = "") then
	    localOrRemote = "1"
	end if
	
	Set mail = CreateObject("CDO.Message") 'calls CDO message COM object
	Set iConf = CreateObject("CDO.Configuration") 'calls CDO configuration COM object
	Set Flds = iConf.Fields
	Flds( "http://schemas.microsoft.com/cdo/configuration/sendusing") = localOrRemote   ' "1" tells cdo we're using the local smtp service, use "2" if not local
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = scSMTP
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = scPort
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 20
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup" 'verify that this path is correct
	Flds.Update 'updates CDO's configuration database
	'if smtp authentication is required
	'==================================
	if scSMTPAuthentication="Y" then
		Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 ' cdoBasic
		Flds("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
		Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = scSMTPUID
		Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = scSMTPPWD
		Flds.Update 'updates CDO's configuration database
	end if
'==================================
	Set mail.Configuration = iConf 'sets the configuration for the message
	mail.BodyPart.Charset = "UTF-8" 
	mail.TextBodyPart.Charset = "UTF-8" 
	mail.HTMLBodyPart.Charset = "UTF-8"

	mail.To = "sales@multiplemonitors.co.uk"
	mail.From = "website@multiplemonitors.uk"
	mail.Sender = "website@multiplemonitors.uk"
	mail.ReplyTo = "sales@multiplemonitors.co.uk"
	mail.Subject = "Test" 
		  mail.HTMLBody="<P>Test Email</p>"
	'If session("News_MsgType")="1" Then
			HTMLBody=HTML
			mail.HTMLBody = strEmailHeader & strEmailH1 & strEmailMiddle & body & strEmailFooter 
	'Else
			'TextBody=Plain
			'mail.TextBody = body
	'End If
	mail.Send 'commands CDO to send the message
	if err then
		pcv_errMsg = err.Description
	end if
	set mail=nothing 
		  %>
	<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Untitled Document</title>
</head>

<body>
	<p>Email Send Test - <%=pcv_errMsg%></p>
</body>
</html>
