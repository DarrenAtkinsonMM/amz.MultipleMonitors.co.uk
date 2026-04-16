<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../pdf/fpdf.asp"-->
<!--#include file="../includes/html-email.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Order Invoice or Packing Slip</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin: 10px; background-image: none; font-size:14px;">

    
    <table class="pcCPcontent">
	<tr>
		<td>
			<div class="pcCPmessageSuccess" style="font-size:14px; font-family: 'Helvetica Neue',Helvetica,Arial,sans-serif;">
				Customer Name: <%=Request("tbCustName")%><br/>
				Customer Email: <%=Request("tbCustEmail")%><br/>
				Tracking Number: <%=Request("tbTracking")%><br/>
				Order Type: <%=Request("ddlOrdType")%><br/>
			</div>
		</td>
	</tr>
	<tr> 
		<td valign="top">&nbsp;</td>
	</tr>
	<tr> 
		<td valign="top">&nbsp;</td>
	</tr>
	</table>
	
	<%
	'SEND EMAIL SECTION
	Dim mail 
	Dim iConf 
	Dim Flds
	Dim localOrRemote
	Dim attachement
	Dim attachement2 
	Dim attachement3
	Dim attachement4 
	
	Dim CustomerName = Request("tbCustName")
	Dim CustomerEmail = Request("tbCustEmail")
	Dim daTrackingNumber = Request("tbTracking")
	
	'attachement = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\shop\130707\pdf-invoices\" & tmpPDFFile

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
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
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
	mail.To = CustomerEmail
	mail.From = "website@multiplemonitors.uk"
	mail.ReplyTo = "sales@multiplemonitors.co.uk"
	'mail.Subject = "Quick Start Guides and Invoice" 
	'strEmailH1 = "Your Quick Start Guides &amp Invoice."
	'strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, you should have received an email dispatch note which contains a tracking code for the delivery.<br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />We have also attached some quick start guides tailored specifically for your order, please <strong>read before you start assembling</strong> your new equipment as they contain helpful pointers on getting up and running as quickly as possible.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
	mail.Bcc = "sales@multiplemonitors.co.uk"
	'mail.AddAttachment attachement
	Select Case Request("ddlOrdType")
		Case "pc"
			attachement2 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\PC-Quickstart.pdf"
			mail.AddAttachment attachement2
			mail.Subject = "Delivery Info & Guides"
			strEmailH1 = "Your Delivery Info & Quick Start Guide"
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, your tracking link can be found below, and the couriers will contact you directly with delivery updates and a 1 hour delivery window.<br /><br /><strong>Tracking Link:</strong> <a href=""https://www.dpdlocal.co.uk/apps/tracking/?reference=" & daTrackingNumber & """>DPD Consignment: " & daTrackingNumber & "</a><br /><br />Attached is a quick start guide for your new PC, please <strong>read before you start assembling</strong> your new equipment as the guide contains helpful pointers on getting up and running as quickly as possible. In addition to this you can view a more detailed <strong>'Getting Started Guide'</strong> on our website here: <a href=""http://www.multiplemonitors.co.uk/pages/getting-started/"">www.multiplemonitors.co.uk/pages/getting-started/</a>.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
		Case "pcstand"
			attachement2 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\PC-Quickstart.pdf"
			attachement3 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\Synergy-Assembly.pdf"
			mail.AddAttachment attachement2
			mail.AddAttachment attachement3
			mail.Subject = "Delivery Info, Invoice, & Guides"
			strEmailH1 = "Your Delivery Info, Invoice, & Quick Start Guides"
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, your tracking link can be found below, and the couriers will contact you directly with delivery updates and a 1 hour delivery window.<br /><br /><strong>Tracking Link:</strong> <a href=""https://www.dpdlocal.co.uk/apps/tracking/?reference=" & daTrackingNumber & """>DPD Consignment: " & daTrackingNumber & "</a><br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />Attached is a quick start guide for your new PC, please <strong>read before you start assembling</strong> your new equipment as the guide contains helpful pointers on getting up and running as quickly as possible. In addition to this you can view a more detailed <strong>'Getting Started Guide'</strong> on our website here: <a href=""http://www.multiplemonitors.co.uk/pages/getting-started/"">www.multiplemonitors.co.uk/pages/getting-started/</a>.<br /><br />An assembly guide for your Synergy Stand is also attached, there will be a hard copy supplied in with the stand as well, again we strongly advise you review it in full before beginning assembly of your stand.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
		Case "bundle"
			attachement2 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\PC-Quickstart.pdf"
			attachement3 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\Synergy-Assembly.pdf"
			mail.AddAttachment attachement2
			mail.AddAttachment attachement3
			mail.Subject = "Delivery Info, Invoice, & Guides"
			strEmailH1 = "Your Delivery Info, Invoice, & Quick Start Guides"
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, your tracking link can be found below, and the couriers will contact you directly with delivery updates and a 1 hour delivery window.<br /><br /><strong>Tracking Link:</strong> <a href=""https://www.dpdlocal.co.uk/apps/tracking/?reference=" & daTrackingNumber & """>DPD Consignment: " & daTrackingNumber & "</a><br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />Attached is a quick start guide for your new PC, please <strong>read before you start assembling</strong> your new equipment as the guide contains helpful pointers on getting up and running as quickly as possible. In addition to this you can view a more detailed <strong>'Getting Started Guide'</strong> on our website here: <a href=""http://www.multiplemonitors.co.uk/pages/getting-started/"">www.multiplemonitors.co.uk/pages/getting-started/</a>.<br /><br />An assembly guide for your Synergy Stand is also attached, there will be a hard copy supplied in with the stand as well, again we strongly advise you review it in full before beginning assembly of your stand.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
		Case "bundle2"
			attachement2 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\PC-Quickstart.pdf"
			attachement3 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\Synergy-Assembly.pdf"
			mail.AddAttachment attachement2
			mail.AddAttachment attachement3
			mail.Subject = "Delivery Info, Invoice, & Quick Start Guides"
			strEmailH1 = "Your Delivery Info, Invoice, & Quick Start Guides"
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, you should have received an email dispatch note which contains a tracking code for the delivery.<br /><br /><strong>Important Delivery Message: </strong>You will receive <strong>two deliveries on the delivery date</strong>. Your new PC and Synergy Stand along with all cabling will be delivered via DPD Local. Your monitors will be arriving in a separate delivery direct from our distributor. Both couriers should contact you directly with tracking information.<br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />Also attached is a quick start guide for your new PC, please <strong>read before you start assembling</strong> your new equipment as the guide contains helpful pointers on getting up and running as quickly as possible. In addition to this you can view a more detailed <strong>'Getting Started Guide'</strong> on our website here: <a href=""http://www.multiplemonitors.co.uk/pages/getting-started/"">www.multiplemonitors.co.uk/pages/getting-started/</a>.<br /><br />Finally an assembly guide for your Synergy Stand is attached, there will be a hard copy supplied in with the stand as well, again we strongly advise you review it in full before beginning assembly of your stand.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
		Case "stand"
			attachement2 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\Synergy-Assembly.pdf"
			mail.AddAttachment attachement2
			mail.Subject = "Delivery Info, Invoice, & Assembly Guide"
			strEmailH1 = "Your Delivery Info, Invoice, & Assembly Guide"
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, your tracking link can be found below, and the couriers will contact you directly with delivery updates and a 1 hour delivery window.<br /><br /><strong>Tracking Link:</strong> <a href=""https://www.dpdlocal.co.uk/apps/tracking/?reference=" & daTrackingNumber & """>DPD Consignment: " & daTrackingNumber & "</a><br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />We have also attached the Assembly Guide for your Synergy Stand, please <strong>read before you start assembling</strong> your new equipment as the guide contains helpful pointers on getting up and running as quickly as possible.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
		Case "array"
			attachement2 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\Synergy-Assembly.pdf"
			mail.AddAttachment attachement2
			mail.Subject = "Delivery Info, Invoice, & Assembly Guide"
			strEmailH1 = "Your Delivery Info, Invoice, & Assembly Guide"
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, your tracking link can be found below, and the couriers will contact you directly with delivery updates and a 1 hour delivery window.<br /><br /><strong>Tracking Link:</strong> <a href=""https://www.dpdlocal.co.uk/apps/tracking/?reference=" & daTrackingNumber & """>DPD Consignment: " & daTrackingNumber & "</a><br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />We have also attached the Assembly Guide for your Synergy Stand, please <strong>read before you start assembling</strong> your new equipment as the guide contains helpful pointers on getting up and running as quickly as possible.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
		Case "array2"
			attachement2 = "c:\inetpub\wwwroot\MultipleMonitors.co.uk\guides\Synergy-Assembly.pdf"
			mail.AddAttachment attachement2
			mail.Subject = "Delivery Info, Invoice, & Assembly Guide"
			strEmailH1 = "Your Delivery Info, Invoice, & Assembly Guide"
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, you will receive an email directly from the courier DPD Local which contains a tracking code for the delivery.<br /><br /><strong>Important Delivery Message: </strong>You will receive <strong>two deliveries on the delivery date</strong>. Your new Synergy Stand along with all cabling will be delivered via DPD Local. Your monitors will be arriving in a separate delivery direct from our distributor. Both couriers should contact you directly with tracking information.<br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />We have also attached the Assembly Guide for your Synergy Stand, please <strong>read before you start assembling</strong> your new equipment as the guide contains helpful pointers on getting up and running as quickly as possible.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
		Case "none"
			mail.Subject = "Your Invoice"
			strEmailH1 = "Your Invoice."
			strEmailP = "<p>Dear " & CustomerName & ",<br /><br />The invoice for your recent order is attached to this email, please save it for your records.<br /><br />Thank you again for your order and if you do have any  questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
	End Select
	HTMLBody=HTML
	mail.HTMLBody = strEmailHeader & strEmailH1 & strEmailMiddle & strEmailP & strEmailFooter
	mail.Send 'commands CDO to send the message
	if err then
		pcv_errMsg = err.Description
	end if
	set mail = nothing


call closedb()
%>
</body>
</html>