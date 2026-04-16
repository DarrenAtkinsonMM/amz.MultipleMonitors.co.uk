<!--#include file="adminv.asp"-->
<!--#include file="common.asp"-->
<!--#include file="utilities.asp"-->
<%
Dim PageName, Body
Dim FS, f, findit
dim tLocalOrRemote, tPort, tCustServEmail

' request values
q=Chr(34)

tLocalOrRemote = getUserInput(Session("pcAdminoptLocalRemote"), 0)
tPort = getUserInput(Session("pcAdminoptPort"), 0)
townerEmail = getUserInput(Session("pcAdminownerEmail"), 0)
tfrmEmail = getUserInput(Session("pcAdminfrmEmail"), 0)
tCustServEmail = getUserInput(Session("pcAdminCustServEmail"), 0)
if trim(tCustServEmail) = "" then tCustServEmail=tfrmEmail
tNoticeNewCust = getUserInput(Session("pcAdminNoticeNewCust"), 0)
if tNoticeNewCust="" then tNoticeNewCust="0"	
tEmailFontSize = getUserInput(Session("pcAdminFontSize"), 0)

tConfirmEmail = replace(Session("pcAdminConfirmEmail"),"""","&quot;")
tConfirmEmail		= replace(tConfirmEmail,"&lt;","<")
tConfirmEmail		= replace(tConfirmEmail,"&gt;",">")
tConfirmEmail = replace(tConfirmEmail, vbCrLf, "<br>")

tReceivedEmail = replace(Session("pcAdminReceivedEmail"),"""","&quot;")
tReceivedEmail	= replace(tReceivedEmail,"&lt;","<")
tReceivedEmail	= replace(tReceivedEmail,"&gt;",">")
tReceivedEmail = replace(tReceivedEmail, vbCrLf, "<br>")

tShippedEmail = replace(Session("pcAdminShippedEmail"),"""","&quot;")
tShippedEmail		= replace(tShippedEmail,"&lt;","<")
tShippedEmail		= replace(tShippedEmail,"&gt;",">")
tShippedEmail = replace(tShippedEmail, vbCrLf, "<br>")

tCancelledEmail	= Session("pcAdminCancelledEmail")
If tCancelledEmail = "" Then
	tCancelledEmail = "This message is to inform you that order number <ORDER_ID> that you submitted in this store on <ORDER_DATE> has been cancelled."
End If
tCancelledEmail	= replace(tCancelledEmail,"""","&quot;")
tCancelledEmail	= replace(tCancelledEmail,"&lt;","<")
tCancelledEmail	= replace(tCancelledEmail,"&gt;",">")
tCancelledEmail	= replace(tCancelledEmail, vbCrLf, "<br>")

tEmailComObj = getUserInput(Session("pcAdminEmailComObj"), 0)

tSMTPAuthenticationTemp = getUserInput(Session("pcAdminSmtpAuth"), 0)
If tSMTPAuthenticationTemp = "1" Then
	tSMTPAuthentication = "Y"
Else
	tSMTPAuthentication = "N"
End If
	
tSMTP =	getUserInput(Session("pcAdminSMTP"), 0)
tSMTPUID = getUserInput(Session("pcAdminSmtpAuthUID"), 0)
tSMTPPWD = getUserInput(Session("pcAdminSmtpAuthPWD")	, 0)

tPayPalEmail = getUserInput(Session("pcAdminPayPalEmail"), 0)
If tPayPalEmail = "" Then
	tPayPalEmail = "We have received your order and we are awaiting payment confirmation from PayPal, the payment option that you selected. As soon as payment confirmation is received, your order will be processed and you will receive an order receipt at this email address."
End If
tPayPalEmail = replace(tPayPalEmail, vbCrLf, "<br>")

If (len(tEmailFontSize)=0) Then
	response.End() '// Not a valid form.
End If	
		
query="UPDATE emailsettings SET ownerEmail='"&townerEmail&"',frmEmail='"&tfrmEmail&"',FontSize='" & tEmailFontSize & "',ConfirmEmail=N'"&tConfirmEmail&"',ReceivedEmail=N'"&tReceivedEmail&"',ShippedEmail=N'"&tShippedEmail&"',CancelledEmail=N'"&tCancelledEmail&"',PayPalEmail='"&tPayPalEmail&"' WHERE id=1"

set rs=Server.CreateObject("ADODB.Recordset")     
set rs=conntemp.execute(query)

if err.number <> 0 then
    response.write "Error in PageCreateEmailSettings.asp: "&Err.Description
end if

PageName="emailSettings.asp"

set StringBuilderObj = new StringBuilder

StringBuilderObj.append CHR(60)&CHR(37)&"private const scEmail="&q&townerEmail&q&CHR(10)
StringBuilderObj.append "private const scFrmEmail="&q&tfrmEmail&q&CHR(10)
StringBuilderObj.append "private const scCustServEmail="&q&tCustServEmail&q&CHR(10)
StringBuilderObj.append "private const scEmailComObj="&q&tEmailComObj&q&CHR(10)
StringBuilderObj.append "private const scSMTP="&q&tSMTP&q&CHR(10)
StringBuilderObj.append "private const scLocalOrRemote="&q&tLocalOrRemote&q&CHR(10)
StringBuilderObj.append "private const scPort="&q&tPort&q&CHR(10)
StringBuilderObj.append "private const scSMTPAuthentication="&q&tSMTPAuthentication&q&CHR(10)
StringBuilderObj.append "private const scSMTPUID="&q&tSMTPUID&q&CHR(10)
StringBuilderObj.append "private const scSMTPPWD="&q&tSMTPPWD&q&CHR(10)
StringBuilderObj.append "private const scEmailFontSize="&q&tEmailFontSize&q&CHR(10)
StringBuilderObj.append "private const scConfirmEmail="&q&tConfirmEmail&q&CHR(10)
StringBuilderObj.append "private const scReceivedEmail="&q&tReceivedEmail&q&CHR(10)
StringBuilderObj.append "private const scShippedEmail="&q&tShippedEmail&q&CHR(10)
StringBuilderObj.append "private const scNoticeNewCust="&q&tNoticeNewCust&q&CHR(10)
StringBuilderObj.append "private const scCancelledEmail="&q&tCancelledEmail&q&CHR(37)&CHR(62)

' create the file using the FileSystemObject
call pcs_SaveUTF8(PageName, PageName, StringBuilderObj.toString)

set StringBuilderObj=nothing

response.redirect "../"&scAdminFolderName&"/emailsettings.asp?s=1&message="&Server.URLEncode("The e-mail settings were updated successfully.")
%>
