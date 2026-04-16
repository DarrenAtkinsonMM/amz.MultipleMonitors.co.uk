<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Set fso=server.CreateObject("Scripting.FileSystemObject")

if request("action")="update" then

	pcv_Security=request("pcv_Security")
	if pcv_Security="" then
		pcv_Security=0
	end if
	pcv_EnforceAdmin=request("pcv_EnforceAdmin")
	if pcv_EnforceAdmin="" then
		pcv_EnforceAdmin=0
	end if
	pcv_UserLogin=request("pcv_UserLogin")
	if pcv_UserLogin="" then
		pcv_UserLogin=0
	end if
	pcv_UserReg=request("pcv_UserReg")
	if pcv_UserReg="" then
		pcv_UserReg=0
	end if
	pcv_AffLogin=request("pcv_AffLogin")
	if pcv_AffLogin="" then
		pcv_AffLogin=0
	end if
	pcv_AffReg=request("pcv_AffReg")
	if pcv_AffReg="" then
		pcv_AffReg=0
	end if
	pcv_Review=request("pcv_Review")
	if pcv_Review="" then
		pcv_Review=0
	end if
	pcv_Contact=request("pcv_Contact")
	if pcv_Contact="" then
		pcv_Contact=0
	end if
	pcv_AdminLogin=request("pcv_AdminLogin")
	if pcv_AdminLogin="" then
		pcv_AdminLogin=0
	end if
	pcv_UseImgs=request("pcv_UseImgs")
	if pcv_UseImgs="" then
		pcv_UseImgs=0
	end if
	pcv_UseImgs2=request("pcv_UseImgs2")
	if pcv_UseImgs2="" then
		pcv_UseImgs2=0
	end if
	pcv_UseImgsGC=request("pcv_UseImgsGC")
	if pcv_UseImgsGC="" then
		pcv_UseImgsGC=0
	end if
	pcv_AlarmMsg=request("pcv_AlarmMsg")
	if pcv_AlarmMsg="" then
		pcv_AlarmMsg=0
	end if
	pcv_AttackCount=request("pcv_AttackCount")
	if pcv_AttackCount="" then
		pcv_AttackCount=0
	end if
    pcv_GWLockAttempts=request("pcv_GWLockAttempts")
	if pcv_GWLockAttempts="" then
		pcv_GWLockAttempts=0
	end if
    pcv_GWSecurity=request("pcv_GWSecurity")
	if pcv_GWSecurity="" then
		pcv_GWSecurity=0
	end if
	pcv_StrongPass=request("pcv_StrongPass")
	if pcv_StrongPass="" then
		pcv_StrongPass=0
	end if
	pcv_CheckSamePass=request("pcv_CheckSamePass")
	if pcv_CheckSamePass="" then
		pcv_CheckSamePass=0
	end if
	pcv_ResetPassMail=request("pcv_ResetPassMail")
	if pcv_ResetPassMail="" then
		pcv_ResetPassMail=0
	end if
	pcv_SaveLogins=request("pcv_SaveLogins")
	if pcv_SaveLogins="" then
		pcv_SaveLogins=0
	end if
	pcv_LockFailedUser=request("pcv_LockFailedUser")
	if pcv_LockFailedUser="" then
		pcv_LockFailedUser=0
	end if
	pcv_LockFailedTime=request("pcv_LockFailedTime")
	if pcv_LockFailedTime="" then
		pcv_LockFailedTime=0
	end if
	pcv_LockFailedCount=request("pcv_LockFailedCount")
	if pcv_LockFailedCount="" then
		if pcv_LockFailedUser="1" then
			pcv_LockFailedCount=5
		else
			pcv_LockFailedCount=0
		end if
	end if
	pcv_LockFailedMin=request("pcv_LockFailedMin")
	if pcv_LockFailedMin="" then
		if pcv_LockFailedUser="1" then
			pcv_LockFailedCount=1
		else
			pcv_LockFailedMin=0
		end if
	end if
	pcv_LockFailedIP=request("pcv_LockFailedIP")
	if pcv_LockFailedIP="" then
		pcv_LockFailedIP=0
	end if
	pcv_LockFailedIPTime=request("pcv_LockFailedIPTime")
	if pcv_LockFailedIPTime="" then
		pcv_LockFailedIPTime=0
	end if
	pcv_LockFailedIPCount=request("pcv_LockFailedIPCount")
	if pcv_LockFailedIPCount="" then
		if pcv_LockFailedIP="1" then
			pcv_LockFailedIPCount=5
		else
			pcv_LockFailedIPCount=0
		end if
	end if
	pcv_LockFailedIPMin=request("pcv_LockFailedIPMin")
	if pcv_LockFailedIPMin="" then
		if pcv_LockFailedIP="1" then
			pcv_LockFailedIPMin=1
		else
			pcv_LockFailedIPMin=0
		end if
	end if
	pcv_ResetPassCapt=request("pcv_ResetPassCapt")
	if pcv_ResetPassCapt="" then
		pcv_ResetPassCapt=0
	end if

	pcv_CaptchaType=request("pcv_CaptchaType")
	if pcv_CaptchaType="" then
		pcv_CaptchaType="0"
	end if

	if PPD="1" then
		findit=Server.MapPath("/"&scPcFolder&"/includes/securitysettings.asp")
	else
		findit=Server.MapPath("../includes/securitysettings.asp")
	end if

	Set f = fso.CreateTextFile(FindIt,True)
	fBody=CHR(60)&CHR(37)
	fBody=fBody&"private const scSecurity=" & pcv_Security & VBCrlf
	fBody=fBody&"private const scUserLogin=" & pcv_UserLogin & VBCrlf
	fBody=fBody&"private const scEnforceAdmin=" & pcv_EnforceAdmin & VBCrlf
	fBody=fBody&"private const scUserReg=" & pcv_UserReg & VBCrlf
	fBody=fBody&"private const scAffLogin=" & pcv_AffLogin & VBCrlf
	fBody=fBody&"private const scAffReg=" & pcv_AffReg & VBCrlf
	fBody=fBody&"private const scReview=" & pcv_Review & VBCrlf
	fBody=fBody&"private const scContact=" & pcv_Contact & VBCrlf
	fBody=fBody&"private const scAdminLogin=" & pcv_AdminLogin & VBCrlf
	fBody=fBody&"private const scUseImgs=" & pcv_UseImgs & VBCrlf
	fBody=fBody&"private const scUseImgs2=" & pcv_UseImgs2 & VBCrlf
	fBody=fBody&"private const scUseImgsGC=" & pcv_UseImgsGC & VBCrlf
	fBody=fBody&"private const scAlarmMsg=" & pcv_AlarmMsg & VBCrlf
	fBody=fBody&"private const scAttackCount=" & pcv_AttackCount & VBCrlf
    fBody=fBody&"private const scGWLockAttempts=" & pcv_GWLockAttempts & VBCrlf
    fBody=fBody&"private const scGWSecurity=" & pcv_GWSecurity & VBCrlf
	fBody=fBody&"private const scStrongPass=" & pcv_StrongPass & VBCrlf
	fBody=fBody&"private const scCheckSamePass=" & pcv_CheckSamePass & VBCrlf
	fBody=fBody&"private const scResetPassMail=" & pcv_ResetPassMail & VBCrlf
	fBody=fBody&"private const scSaveLogins=" & pcv_SaveLogins & VBCrlf
	fBody=fBody&"private const scLockFailedUser=" & pcv_LockFailedUser & VBCrlf
	fBody=fBody&"private const scLockFailedTime=" & pcv_LockFailedTime & VBCrlf
	fBody=fBody&"private const scLockFailedCount=" & pcv_LockFailedCount & VBCrlf
	fBody=fBody&"private const scLockFailedMin=" & pcv_LockFailedMin & VBCrlf
	fBody=fBody&"private const scLockFailedIP=" & pcv_LockFailedIP & VBCrlf
	fBody=fBody&"private const scLockFailedIPTime=" & pcv_LockFailedIPTime & VBCrlf
	fBody=fBody&"private const scLockFailedIPCount=" & pcv_LockFailedIPCount & VBCrlf
	fBody=fBody&"private const scLockFailedIPMin=" & pcv_LockFailedIPMin & VBCrlf
	fBody=fBody&"private const scResetPassCapt=" & pcv_ResetPassCapt & VBCrlf
	fBody=fBody&"private const scCaptchaType=" & pcv_CaptchaType & VBCrlf
	fBody=fBody&CHR(37)&CHR(62)
	f.write fBody
	f.Close
	Set f=nothing

	call closeDb()
	response.redirect "AdminSecuritySettings.asp?s=1&msg=Security Settings were updated successfully!"

end if

set fso=nothing


pageTitle="Advanced Security Settings"
pageIcon="pcv4_security.png"
section="layout"
%>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="AdminSecuritySettings.asp?action=update" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
			<th colspan="3">Overview</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td colspan="3">
		<p>Use this feature to tun on and off additional security filters when customers and Control Panel users login. The CAPTCHA feature requires a working XML Parser: <a href="pcTSUtility.asp" target="_blank">review your XML Parser settings</a>. The settings listed below apply only when &quot;Advanced Security&quot; is turned &quot;On&quot;, <u>except for</u> the one for the &quot;Contact Us&quot; form, which works independently of the others.&nbsp;<a href="http://wiki.productcart.com/productcart/settings-security-settings" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Learn more about this topic"></a></p></td>
	</tr>
		<tr>
			<td colspan="3">
			<p> Turn Advanced Security Settings <input type="radio" name="pcv_Security" value="1" <%if scSecurity=1 then%>checked<%end if%> class="clearBorder">On <input type="radio" name="pcv_Security" value="0" <%if scSecurity<>1 then%>checked<%end if%> class="clearBorder">Off</p>
            </td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="3"><p></p></td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
			<tr>
			<th colspan="3">PC Defender</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_EnforceAdmin" value="1" <%if scEnforceAdmin=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Lock Suspicious User Accounts</td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Storefront</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_UserLogin" value="1" <%if scUserLogin=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to User <strong>Login</strong> pages</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_UserReg" value="1" <%if scUserReg=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to User <strong>Registration</strong> pages</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_UseImgs" value="1" <%if scUseImgs=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%" valign="top">Add <strong>CAPTCHA</strong> (random security code) to <strong>Login/Registration</strong> pages in the storefront.<br /><a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
 		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_UseImgsGC" value="1" <%if scUseImgsGC=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%" valign="top">Add <strong>CAPTCHA</strong> (random security code) to <strong>Guest Checkout</strong> page in the storefront.<br /><a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
 		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_Review" value="1" <%if scReview=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add <strong>CAPTCHA</strong> (random security code) to the <strong>Product Review</strong> submission page. <br /> <a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
 		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_Contact" value="1" <%if scContact=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add <strong>CAPTCHA</strong> (random security code) to the <strong><a href="../pc/contact.asp" target="_blank">Contact Us</a></strong> form. <br /><a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
    	<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_AffLogin" value="1" <%if scAffLogin=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to Affiliate Login pages</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_AffReg" value="1" <%if scAffReg=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to Affiliate Registration pages</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_GWSecurity" value="1" <%if scGWSecurity=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">
		  Add advanced security to the gateway pages. Lock out customer on &nbsp; <input type="text" name="pcv_GWLockAttempts" value="<%=scGWLockAttempts%>" size="1" /> &nbsp; failed payment attempts.
		</td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Control Panel</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_AdminLogin" value="1" <%if scAdminLogin=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to Control Panel Login page</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_UseImgs2" value="1" <%if scUseImgs2=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%" valign="top">Add <strong>CAPTCHA</strong> (random security code) to the Control Panel <strong>Login</strong> page.<br /><a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Alerts</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_AlarmMsg" value="1" <%if scAlarmMsg=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Send a notification e-mail to the store administrator when someone attempts to log into the store more than the number of attempts listed below. This feature can alert you of a script-based attacked performed against the store. This applies to any login form in the storefront and in the Control Panel.</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td>&nbsp;</td>
		<td width="95%">Number of Consecutive Attempts: <input type="text" name="pcv_AttackCount" size="4" value="<%=scAttackCount%>"></td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Password & Login</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_StrongPass" value="1" <%if scStrongPass=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Enforce strong password security</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_CheckSamePass" value="1" <%if scCheckSamePass=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Prevent customers from using the same password more than once</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_ResetPassMail" value="1" <%if scResetPassMail=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Send an email to customer when the password is successfully reset</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_SaveLogins" value="1" <%if scSaveLogins=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Record all login attempts</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_LockFailedUser" value="1" <%if scLockFailedUser=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Keeping track of failed attempts (per user account) and locking the account <input type="text" name="pcv_LockFailedTime" size="4" value="<%=scLockFailedTime%>"> minutes if</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td>&nbsp;</td>
		<td width="95%">Number of failed Login Attempts: <input type="text" name="pcv_LockFailedCount" size="4" value="<%=scLockFailedCount%>"> within <input type="text" name="pcv_LockFailedMin" size="4" value="<%=scLockFailedMin%>"> minutes</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_LockFailedIP" value="1" <%if scLockFailedIP=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Keeping track of failed attempts (per IP Address) and locking the account <input type="text" name="pcv_LockFailedIPTime" size="4" value="<%=scLockFailedIPTime%>"> minutes if</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td>&nbsp;</td>
		<td width="95%">Number of failed Login Attempts: <input type="text" name="pcv_LockFailedIPCount" size="4" value="<%=scLockFailedIPCount%>"> within <input type="text" name="pcv_LockFailedIPMin" size="4" value="<%=scLockFailedIPMin%>"> minutes</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_ResetPassCapt" value="1" <%if scResetPassCapt=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Use Captcha on the password reset request screen</td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">CAPTCHA Settings</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="3">Which CAPTCHA do you want to use?</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="radio" name="pcv_CaptchaType" value="0" <%if scCaptchaType<>"1" then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Default CAPTCHA</td>
		</tr>
		<tr>
        <%
        call pcs_getReCaSettingsNoAuth(0)
        If prcs_Had=1 Then
        %>
            <td width="5%">&nbsp;</td>
            <td valign="top"><input type="radio" name="pcv_CaptchaType" value="1" <%if scCaptchaType="1" then%>checked<%end if%> class="clearBorder"></td>
            <td width="95%">Use Google reCAPTCHA&nbsp;&nbsp;(<a href="reCaptchaSettings.asp" target="_blank">settings</a>)</td>
        <% Else %>
            <td>&nbsp;</td>
            <td valign="top"><input type="radio" name="pcv_CaptchaTypeDisabled" value="1" disabled="disabled" class="disabled"></td>
            <td width="95%"><span class="disabled">Google reCAPTCHA</span>&nbsp;&nbsp;(<a href="reCaptchaSettings.asp" target="_blank">Click here to enable</a>)</td>
            <% End If %>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="3"><hr></td>
		</tr>
		<tr>
		<td>&nbsp;</td>
		<td></td>
		<td>
			<input type="submit" name="submit" value="Update Settings" class="btn btn-primary">
            &nbsp;<input type="button" class="btn btn-default"  name="back" value="Back" onClick="JavaScript:history.go(-1);">
        </td>
        </tr>
		<tr>
			<td colspan="3">&nbsp;</td>
		</tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->
