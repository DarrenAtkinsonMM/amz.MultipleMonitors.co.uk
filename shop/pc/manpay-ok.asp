<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
    <%
	'SEND EMAIL SECTION
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
	mail.To = "sales@multiplemonitors.co.uk"
	mail.From = "Multiple Monitors <website@multiplemonitors.uk>"
	mail.ReplyTo = "sales@multiplemonitors.co.uk"
	'mail.Subject = "Quick Start Guides and Invoice" 
	'strEmailH1 = "Your Quick Start Guides &amp Invoice."
	'strEmailP = "<p>Dear " & CustomerName & ",<br /><br /> We are pleased to let you know that your order has been released for delivery, you should have received an email dispatch note which contains a tracking code for the delivery.<br /><br />The invoice for your order is attached to this email, please save it for your records.<br /><br />We have also attached some quick start guides tailored specifically for your order, please <strong>read before you start assembling</strong> your new equipment as they contain helpful pointers on getting up and running as quickly as possible.<br /><br />Thank you again for your order and if you do have any setup questions or problems just let us know, we are here to help!<br /><br />Best Regards,<br /><br />Multiple Monitors</p>"
	mail.Bcc = "sales@multiplemonitors.co.uk"
			mail.Subject = "Manual Payment Received - " & Request.QueryString("payid")
			strEmailH1 = "Manual Payment Received"
			strEmailP = "<p>A manual payment has been sent to Sage Pay</p>"
	HTMLBody=HTML
	mail.HTMLBody = strEmailHeader & strEmailH1 & strEmailMiddle & strEmailP & strEmailFooter
	mail.Send 'commands CDO to send the message

	%>

<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Payment Confirmation">Quick Payment Page - Step: 3 / 3</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->

	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s"><div id="pcMain">
<div id="pcMain">
<h1>Thank You For Your Payment</h1><br />
<p>Your payment details have been received successfully and are now been processed by our payment processor.</p><br />
<p>If you had been asked to make this payment via a member of our team over email you can email them to let them know you have paid if you wish however they will be notified of this payment automatically.</p><br />
<p>If a receipt is needed it will be issued via email when the payment has been processed.</p><br />
<p>You may now close this page.</p>
</div>
					</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->
<!--#include file="footer_wrapper.asp"-->