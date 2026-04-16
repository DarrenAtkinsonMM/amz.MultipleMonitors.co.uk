<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<%
Dim pcEmail, pcPassword, securityCode, CAPTCHA_Postback, guestCheckout
Dim tmpResult,tmpResult1

'// clear alternative payment method
session("PayWithAmazon")=""
session("ExpressCheckoutPayment")=""

pcEmail=getUserInput(request("email"),0)
pcPassword=getUserInput(request("password"),0)
securityCode=getUserInput(request("securityCode"),0)
CAPTCHA_Postback=getUserInput(request("CAPTCHA_Postback"),0)
guestCheckout=getUserInput(request("guestCheckout"),0)

pcErrMsg=""

if scSecurity=1 AND (scUserLogin=1 OR scUserReg=1) then
	pcv_Test=0
	'// Remote access attempt
	if (session("store_userlogin")<>"1") AND (session("store_adminre")<>"1") then
		session("store_userlogin")=""
		session("store_adminre")=""
		pcv_Test=1
	end if
	if pcv_Test=0 AND scUseImgs=1 then
		if scCaptchaType="1" then
			blnCAPTCHAcodeCorrect=pcf_checkReCaptcha()
		else%>
			<!-- Include file for CAPTCHA configuration -->
			<!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
			 
			<!-- Include file for CAPTCHA form processing -->
			<!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
		<%end if	
		If not blnCAPTCHAcodeCorrect then
			pcv_Test=2
			pcErrMsg=dictLanguage.Item(Session("language")&"_security_3")
		end if
	end if

	if pcv_Test=1 then
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
		pcErrMsg=dictLanguage.Item(Session("language")&"_security_2")
	end if					
end if

if guestCheckout = "" then
	if pcErrMsg="" then
		if pcEmail="" OR pcPassword="" then
			pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checklogin_1")
		else
            pcStrLoginPassword=pcPassword
			tmpResult="false"
			query="SELECT idcustomer,suspend,pcCust_Locked,pcCust_Guest,[password] FROM customers WHERE email like '" & pcEmail & "'"
			set rs=connTemp.execute(query)
			if not rs.eof then
				if rs("pcCust_Guest")="1" then
					pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checklogin_5")
				else
					if pcErrMsg="" then
						tmpResult = pcf_CheckPassH(pcStrLoginPassword, rs("password"))
						if Ucase(""&tmpResult)<>"TRUE" then
							pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checklogin_4")
							pcIntIdcustomer=rs("idcustomer")
							tmpResult1=pcf_SaveLoginLockFailed(pcIntIdcustomer,"1")
							if tmpResult1>"0" then
								pcErrMsg=dictLanguage.Item(Session("language")&"_newpass_12") & session("pcSFLockMinutes") & dictLanguage.Item(Session("language")&"_newpass_12a")
								session("pcSFLockMinutes")=""
							end if
						end if
					end if
					
					if pcErrMsg="" then
						tmpResult3 = pcf_CheckNewPassH("", pcEmail)
						if tmpResult3="0" then
							pcErrMsg=dictLanguage.Item(Session("language")&"_resetpass_2")
						end if
						tmpResult3=pcf_CheckUnlockUser("",pcEmail)
						if tmpResult3="1" then
							pcErrMsg=dictLanguage.Item(Session("language")&"_newpass_11") & session("pcSFLockMinutes") & dictLanguage.Item(Session("language")&"_newpass_11a")
							session("pcSFLockMinutes")=""
						end if
					end if
				end if
			end if

			if pcErrMsg="" then
				if (rs.eof) OR (Ucase(""&tmpResult)<>"TRUE") then
					pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checklogin_2")
				else
					if rs("suspend")="1" then
						pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_3")
					end if
					if rs("pcCust_Locked")="1" then
						pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_4")
					end if
				end if
			end if
		end if
		set rs=nothing
	end if
	if pcErrMsg="" then
		session("pcSFLoginEmail")=pcEmail
		session("pcSFLoginPassword")=pcPassword
		erypassword=encrypt(pcPassword, 9286803311968)
		session("pcSFPassNotEnter")="0"
		session("pcSFPassWordExists")="YES"
		session("pcSFEryPassword")=erypassword
		response.redirect "login.asp?lmode=0&opc=1"
	end if
else
	if pcErrMsg = "" then
		pcErrMsg = "OK"
	end if
end if
%>
<%response.write pcErrMsg%>
<%
call closeDb()
%>

