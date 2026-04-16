<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%

if request("action")="update" then
	prcs_SiteKey=request("prcs_SiteKey")
	if prcs_SiteKey<>"" then
		prcs_SiteKey=enDeCrypt(prcs_SiteKey, scCrypPass)
	end if
	prcs_Secret=request("prcs_Secret")
	if prcs_Secret<>"" then
		prcs_Secret=enDeCrypt(prcs_Secret, scCrypPass)
	end if
	prcs_Theme=request("prcs_Theme")
	if prcs_Theme="" then
		prcs_Theme="light"
	end if
	prcs_Type=request("prcs_Type")
	if prcs_Type="" then
		prcs_Type="image"
	end if
	prcs_Size=request("prcs_Size")
	if prcs_Size="" then
		prcs_Size="normal"
	end if
		
	query="SELECT pcRCS_ID FROM pcReCaSettings;"
	set rs=connTemp.execute(query)

	if not rs.eof then
		tmp1=""
		if prcs_SiteKey<>"" then
			tmp1="pcRCS_SiteKey='" & prcs_SiteKey & "' "
		end if
		if prcs_Secret<>"" then
			if tmp1<>"" then
				tmp1=tmp1 & ", "
			end if
			tmp1=tmp1 & "pcRCS_Secret='" & prcs_Secret & "' "
		end if
		if tmp1<>"" then
			tmp1=tmp1 & ", "
		end if
		query="UPDATE pcReCaSettings SET " & tmp1 & " pcRCS_Theme='" & prcs_Theme & "', pcRCS_Type='" & prcs_Type & "', pcRCS_Size='" & prcs_Size & "';"
		set rs=connTemp.execute(query)
		set rs=nothing
	else
		if prcs_SiteKey="" AND prcs_Secret="" then
			call closeDb()
			response.redirect "reCaptchaSettings.asp?s=0"
		end if
		query="INSERT INTO pcReCaSettings (pcRCS_SiteKey,pcRCS_Secret,pcRCS_Theme,pcRCS_Type,pcRCS_Size) VALUES ('" & prcs_SiteKey & "','" & prcs_Secret & "','" & prcs_Theme & "','" & prcs_Type & "','" & prcs_Size & "');"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	set rs=nothing

	call closeDb()
	response.redirect "reCaptchaSettings.asp?s=1"

end if

prcs_SiteKey=""
prcs_Secret=""
prcs_Theme="light"
prcs_Type="image"
prcs_Size="normal"

query="SELECT pcRCS_SiteKey,pcRCS_Secret,pcRCS_Theme,pcRCS_Type,pcRCS_Size FROM pcReCaSettings;"
set rs=connTemp.execute(query)
if not rs.eof then
	prcs_SiteKey=rs("pcRCS_SiteKey")
	prcs_Secret=rs("pcRCS_Secret")
	prcs_Theme=rs("pcRCS_Theme")
	prcs_Type=rs("pcRCS_Type")
	prcs_Size=rs("pcRCS_Size")
end if
set rs=nothing

if (prcs_SiteKey<>"") AND (prcs_Secret<>"") then
	prcs_Had=1
else
	prcs_Had=0
end if

pageTitle="Google reCAPTCHA Settings" 
pageIcon="pcv4_icon_settings.png"
section="layout" 
%>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="reCaptchaSettings.asp?action=update" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any
				if request("s")="1" then
					msgType=1
					msg="Google reCAPTCHA Settings were updated successfully!"
				else
					if prcs_Had=0 then
						msgType=2
						msg="You need to have Google reCAPTCHA &quot;Site Key&quot; and &quot;Secret Key&quot; to use this feature. If you don't have them yet, you can <a href=""https://www.google.com/recaptcha/admin"" target=""_blank"">click here</a> to get them.</p>"
					end if
				end if%>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% ' END show message %>
                
                <% if prcs_Had=1 and scCaptchaType=0 then %>                
                    <div class="bs-callout bs-callout-info">
                        Google reCAPTCHA is enabled, but is not your default CAPTCHA. <a href="AdminSecuritySettings.asp">Click here </a> to modify your advanced security settings.
                    </div>
                <% end if %>
                
                <% if prcs_Had=1 and scCaptchaType=1 and scSecurity=0 then %>                
                    <div class="bs-callout bs-callout-warning">
                        Google reCAPTCHA is enabled as your default CAPTCHA, but will not be displayed until you <a href="AdminSecuritySettings.asp">turn on advanced security</a>.
                    </div>
                <% end if %>
                
                <% if prcs_Had=1 and scCaptchaType=1 and scSecurity=1 then %>                
                    <div class="bs-callout bs-callout-success">
                        Google reCAPTCHA is enabled as your default CAPTCHA for the following pages:
                        <% If scUseImgs=1 Or scReview=1 Or scContact=1 Or scAdminLogin=1 Then %>
                            <ul>
                                <% If scUseImgs=1 Then %>
                                    <li>Login/Registration pages in the storefront</li>
                                <% End If %>
                                <% If scReview=1 Then %>
                                    <li>Product Review submission page</li>
                                <% End If %>
                                <% If scContact=1 Then %>
                                    <li>Contact Us form</li>
                                <% End If %>
                                <% If scAdminLogin=1 Then %>
                                    <li> Control Panel Login page</li>
                                <% End If %>
                            </ul>
                            If you follow the instruction outlined <a href="https://productcart.desk.com/customer/portal/articles/search?q=captcha" target="_blank">here</a> then a sample reCaptcha should be display below this text.
                            <div id="gcaptcha"></div>
                        <% Else %>
                            <ul><li>No pages selected. <a href="AdminSecuritySettings.asp">Click here </a> to modify your advanced security settings.</li></ul>
                        <% End If %>                        
                    </div>
                <% end if %>
                
            </td>
        </tr>
		<tr>
			<td colspan="2" class="pcCPspacer">
				<p></p>
			</td>
		</tr>
		<tr>
			<th colspan="2">reCAPTCHA Keys</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer">
			<%if prcs_Had=1 then%>
				<p>Your current "Site Key" and "Secret Key" are hidden for security reasons. If you want to update them, please enter new keys into the fields below.</p>
			<%end if%>
			</td>
		</tr>
		<tr>
		<td width="15%">Site Key:</td>
		<td width="85%"><input type="text" name="prcs_SiteKey" size="30" value=""></td>
		</tr>
		<td width="15%">Secret Key:</td>
		<td width="85%"><input type="text" name="prcs_Secret" size="30" value=""></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">reCAPTCHA Settings</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr valign="top">
		<td width="15%">Widget Theme:</td>
		<td width="85%">
				<input type="radio" name="prcs_Theme" value="light" checked>&nbsp;light<br>
				<input type="radio" name="prcs_Theme" value="dark" <%if prcs_Theme="dark" then%>checked<%end if%>>&nbsp;dark
		</td>
		</tr>
		<tr valign="top">
		<td width="15%">Widget Type:</td>
		<td width="85%">
			<input type="radio" name="prcs_Type" value="image" checked>&nbsp;image<br>
			<input type="radio" name="prcs_Type" value="audio" <%if prcs_Type="audio" then%>checked<%end if%>>&nbsp;audio
            <br>
            <div class="help-block">When extra validation is needed Google will display an image or play audio.</div>
		</td>
		</tr>
		<tr valign="top">
		<td width="15%">Widget Size:</td>
		<td width="85%">
			<input type="radio" name="prcs_Size" value="normal" checked>&nbsp;normal<br>
			<input type="radio" name="prcs_Size" value="compact" <%if prcs_Size="compact" then%>checked<%end if%>>&nbsp;compact
		</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr>
		<td>&nbsp;</td>
		<td>
			<input type="submit" name="submit" value="Update Settings" class="btn btn-primary">
            &nbsp;
            <input type="button" class="btn btn-default"  name="security" value="Advanced Security Settings" onclick="location.href='AdminSecuritySettings.asp'">
            &nbsp;
            <input type="button" class="btn btn-default"  name="back" value="Back" onClick="JavaScript:history.go(-1);">
        </td>
        </tr>
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->
