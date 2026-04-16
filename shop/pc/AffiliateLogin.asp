<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/validation.asp" -->
<% if request.form("SubmitCO")<>"" then
	ErrCnt=0
	EP=0
	if (scSecurity=1) and (scAffLogin=1) and (scUseImgs=1) then
	Session("store_affpostnum")=replace(request("postnum"),"'","''")
	else
	Session("store_affpostnum")=""
	end if
	'form is submitted
	Email=replace(request.form("Email"),"'","''")
	session("Email")=Email
	if Email="" then
		ErrCnt=ErrCnt+1
	End if
	password=request.form("password")
	if password="" then
		ErrCnt=ErrCnt+1
		EP=1
	End if
	
	if (scSecurity=1) and (scAffLogin=1) and (scUseImgs=1) then

		If scCaptchaType="1" Then
			blnCAPTCHAcodeCorrect=pcf_checkReCaptcha()
		Else%>
			<!-- Include file for CAPTCHA configuration -->
			<!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
			 
			<!-- Include file for CAPTCHA form processing -->
			<!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
		<%End If
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
			Session("store_affpostnum")=""
			response.redirect "AffiliateLogin.asp?EP="&EP&"&msg=1"
		End if
	End if

	If ErrCnt>0 then
	If (scSecurity=1) and (scAffLogin=1) and (scAlarmMsg=1) then
				if session("AttackCount")="" then
					session("AttackCount")=0
				end if
				session("AttackCount")=session("AttackCount")+1
				if session("AttackCount")>=scAttackCount then
				session("AttackCount")=0%>
				<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if	
	End if
		response.redirect "AffiliateLogin.asp?EP="&EP&"&err=2"
	Else
		erypassword=encrypt(password, 9286803311968)
		session("erypassword")=erypassword
		response.redirect "AffiliateLoginB.asp" 
	End if

end if
%>
<% ' if customer already login
if (Session("pc_idAffiliate")<>0) then
 response.redirect "AffiliateMain.asp"
end if
%>
<!--#include file="header_wrapper.asp"-->
<%
pcRequestRedirect=getUserInput(request("redirectUrl"),250)
if len(pcRequestRedirect)>0 then
	session("redirectUrlLI")=pcRequestRedirect
end if
%>
<div id="pcMain">
	<div class="pcMainContent">
		<form method="post" name="auth" action="AffiliateLogin.asp" class="pcForms">
			<% 
				msg = ""
				code = getUserInput(Request.QueryString("msg"), 0)
				Select Case code
				Case "1" : msg = dictLanguage.Item(Session("language")&"_security_3")
				Case "2" : msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
				End Select

				If msg<>"" Then
					%><div class="pcErrorMessage"><%= msg %></div><% 
				End If 
			%>

			<!-- start of login section -->
			<div id="pcAffLogin" class="pcShowContent">
				<h2><%= dictLanguage.Item(Session("language")&"_AffLogin_1")%></h2>
				<p><%= dictLanguage.Item(Session("language")&"_AffLogin_2")%></p>

				<% '// Login %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_AffLogin_3")%></div>
					<div class="pcFormField">
						<input type="text" name="Email" size="25" maxlength="150" value="<%=session("Email")%>">
						<% if msg="" then %>
							<img src="<%=pcf_getImagePath("",rsIconObj("requiredicon"))%>">
						<% else
								if session("email")="" then %>
								<img src="<%=pcf_getImagePath("",rsIconObj("errorfieldicon"))%>">
								<% end if %>
						<% end if %>
					</div>
				</div>
						
				<% '// Login Password %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_AffLogin_4")%></div>
					<div class="pcFormField">
						<input type="password" name="password" size="25" maxlength="150">
						<% if msg="" then %>
						<img src="<%=pcf_getImagePath("",rsIconObj("requiredicon"))%>">
						<% else
								if request("EP")="1" then %>
								<img src="<%=pcf_getImagePath("",rsIconObj("errorfieldicon"))%>">
								<% end if %>
						<% end if %>
					</div>
				</div>

				<%
				Session("store_afflogin")="1"
				Session("store_affpostnum")=""
				session("store_affnum")="      "
				%>
				<%if (scSecurity=1) and (scAffLogin=1) and (scUseImgs=1) then%>
					<%if scCaptchaType="1" then
						call pcs_genReCaptcha()
					else%>
						<!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
					<%end if%>
				<%end if%>
				<div class="pcFormButtons">
					<button class="pcButton pcButtonLogin" name="SubmitCO" value="Submit" id="submit">
						<img src="<%=pcf_getImagePath("",rslayout("login"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_login") %>" />
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_login") %></span>
					</button>
				</div>
			</div>
			<!-- end of login section -->
				
			<!-- start of register section -->
			<div id="pcAffRegister" class="pcShowContent">
				<h2><%= dictLanguage.Item(Session("language")&"_AffLogin_5")%></h2>
				<p><%= dictLanguage.Item(Session("language")&"_AffLogin_6")%></p>
				<div class="pcFormButtons">
					<a class="pcButton pcButtonRegister" href="NewAffa.asp">
						<img src="<%=pcf_getImagePath("",rslayout("register"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_register") %>" />
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_register") %></span>
					</a>
				</div>
			</div>
			<!-- end of register section -->
		</form>
		
		<div class="pcClear"></div>
		<div class="pcFormItem">
			<hr>
		</div>

		<!-- start of password request -->
		<div class="pcShowContent">
			<% if request.querystring("s")="1" then %>
				<div class="pcErrorMessage">
					<%= dictLanguage.Item(Session("language")&"_AffLogin_7")%>
				</div>
			<% else 
			pcRequestRedirect=getUserInput(request("redirectUrl"),250)
			%>
				<div class="pcFormItem">
					<%= dictLanguage.Item(Session("language")&"_AffLogin_8")%>
					<a href="Affiliatefpassword.asp?redirectUrl=<%if trim(pcRequestRedirect)<>"" then%><%=Server.UrlEncode(pcRequestRedirect)%><%else%><%=Server.UrlEncode(session("redirectUrlLI"))%><%end if%>&frURL=AffiliateLogin.asp"><%= dictLanguage.Item(Session("language")&"_AffLogin_9")%></a>
				</div>
			<% end if %>
		</div>
		<!-- end of password request -->

	</div>
</div>
<%call clearLanguage()%>
<!--#include file="footer_wrapper.asp"-->
