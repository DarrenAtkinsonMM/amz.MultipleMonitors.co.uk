<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcCPLog.asp" -->
<%
if scDSN="" then 
    response.redirect "../setup/default.asp"
end if
%>
<% 
'Checks for cookie
Dim CookieVar, ShowAgreement
ShowAgreement=0
CookieVar=Request.Cookies("AgreeLicense45")

if request("RedirectURL")<>"" then
	Session("RedirectURL")=getUserInput(request("RedirectURL"),0)
end if

If CookieVar="Agreed" then
Else
	ShowAgreement=1
End If
If request.form("Submit2")<>"" then
	AgreeVar=request.form("agree")
	If AgreeVar=1 then
		'place cookie
		Response.Cookies("AgreeLicense45")="Agreed"
		Response.Cookies("AgreeLicense45").Expires=Date() + 365
		MyCookiePath=Request.ServerVariables("PATH_INFO")
		do while not (right(MyCookiePath,1)="/")
		MyCookiePath=mid(MyCookiePath,1,len(MyCookiePath)-1)
		loop
		Response.Cookies("AgreeLicense45").Path=MyCookiePath
		call closeDb()
response.redirect "login_1.asp"
	else
		'send message to agree
		AgreeMsg="Agree to the terms and conditions of the ProductCart End User License Agreement to continue."
		call closeDb()
response.redirect "login_1.asp?AM="&server.URLEncode(AgreeMsg)
	end if
End If
' verifies if admin is logged, so as not send to login page
if session("admin")<>0 then
if Session("RedirectURL")<>"" then
	RedirectURL=Session("RedirectURL")
	Session("RedirectURL")=""
	call closeDb()
response.redirect RedirectURL
else
 call closeDb()
response.redirect "menu.asp"
end if 
end if
pageTitle="Login"
pageIcon="pcv4_icon_login.png"
%>
<!--#include file="../includes/validation.asp" -->
<%
if (request.form("submitf")="1") and (Session("cp_Adminlogin")="1") then
	if (scSecurity=1) and (scAdminLogin=1) and (scUseImgs2=1) then
		if scCaptchaType="1" then
			blnCAPTCHAcodeCorrect=pcf_checkReCaptcha()
		else%>
			<!-- Include file for CAPTCHA configuration -->
			<!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
			 
			<!-- Include file for CAPTCHA form processing -->
			<!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
		<%end if
		If not blnCAPTCHAcodeCorrect then
			If scAlarmMsg=1 then
				if session("AttackCount")="" then
					session("AttackCount")=0
				end if
				session("AttackCount")=session("AttackCount")+1
				if session("AttackCount")>=scAttackCount then
					session("AttackCount")=0%>
					<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if	
			End if
			Session("cp_postnum")=""
			call closeDb()
response.redirect "login_1.asp?msg="& Server.Urlencode(dictLanguage.Item(Session("language")&"_security_3"))
		End if
	End if
End if
%>
<% validateForm "login.asp" %>
<!--#include file="AdminHeader.asp"-->
<% if ShowAgreement=0 then %>

	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
	<div style="float: right; width: 325px">
    	<div style="float:right; margin: 0 5px;"><img src="images/twitter_newbird_boxed_blueonwhite.png" alt="Twitter"></div>
    	<div style="margin-bottom: 10px;">The latest news from the ProductCart world, via Twitter</div>
    	<a class="twitter-timeline"  href="https://twitter.com/productcart"  data-widget-id="388668860498325504" width="325">Tweets by @productcart</a>
			<script type=text/javascript>
            try {
            !function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0],p=/^http:/.test(d.location)?'http':'https';if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src=p+"://platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");
            } catch(err) { }
            </script>

			<div id="fb-root"></div>
			<script type=text/javascript>
            try {
            (function(d, s, id) {
			  var js, fjs = d.getElementsByTagName(s)[0];
			  if (d.getElementById(id)) return;
			  js = d.createElement(s); js.id = id;
			  js.src = "//connect.facebook.net/en_US/all.js#xfbml=1&status=0";
			  fjs.parentNode.insertBefore(js, fjs);
			}(document, 'script', 'facebook-jssdk'));
            } catch(err) { }
            </script>
			
			<br />
			<div class="fb-follow" data-href="http://www.facebook.com/productcartsoftware" data-width="325" data-height="The pixel height of the plugin" data-colorscheme="light" data-layout="standard" data-show-faces="true"></div>		
            
    </div>

    <form method="post" action="login_1.asp" name="login" class="pcForms">
        <table class="pcCPcontent" style="width: 400px;">
            <tr> 
                <td colspan="4">&nbsp;<% validateError %></td>
            </tr>
            <tr> 
                <td align="right">User:</td>
                <td>
				<% textbox "idadmin", "", 12, "textbox"
				validate "idadmin", "positiveNumber" %>
				</td>
                <td align="right">Password:</td>
                <td> 
				<% textbox "password","", 12, "password"
                validate "password", "required" %>
                </td>
            </tr>
            <tr>
            	<td></td>
            	<td><a href="forgot_username.asp" style="text-decoration: none; color:#777;">Forgot User Name?</a></td>
            	<td></td>
                <td><a href="forgot_password.asp" style="text-decoration: none; color:#777;">Forgot Password?</a></td>
            </tr>
			<%
            Session("cp_Adminlogin")="1"
            Session("cp_postnum")=""
            session("cp_num")="      "%>
            <%if (scSecurity=1) and (scAdminLogin=1) and (scUseImgs2=1) then%>
            	<tr>
                	<td colspan="4" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <td></td>
                    <td colspan="3">
						<%if scCaptchaType="1" then
							call pcs_genReCaptcha()
						else%>
							<!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
						<%end if%>
					</td>
                </tr>
            <%end if%>
            <tr> 
                <td colspan="4">&nbsp;</td>
            </tr>
            <tr> 
                <td>&nbsp;</td>
                <td colspan="3">
                <input type="hidden" name="submitf" value="1">
                <input type="submit" name="Submit" value="Submit" class="btn btn-primary">
                </td>
            </tr>
        </table>
    </form>
    <script type=text/javascript>
     document.login.idadmin.focus();
    </script>
    
<% else %>
    <form action="login_1.asp" method="post" name="IAgree" id="IAgree" class="pcForms">
        <table class="pcCPcontent">
            <tr> 
                <td colspan="2">
                <% if request.querystring("AM")<>"" then %>
                <div class="pcCPmessage">
                <%=request.querystring("AM")%>
                </div>
                <% end if %>
                </td>
            </tr>
            <tr>
                <td colspan="2">
				<!--#include file="inc_EULA.asp"-->
                </td>
            </tr>
            <tr> 
                <td colspan="2">
                <input type="checkbox" name="agree" value="1" class="clearBorder"> I agree to the terms and conditions of the <strong>ProductCart End User License Agreement</strong>, which are listed above.
                </td>
            </tr>
            <tr> 
                <td colspan="2" style="padding-top: 10px;">
                <input type="submit" name="Submit2" value="Continue" class="btn btn-primary">
                </td>
            </tr>
        </table>
    </form> 
<% end if %>
<!--#include file="AdminFooter.asp"-->
