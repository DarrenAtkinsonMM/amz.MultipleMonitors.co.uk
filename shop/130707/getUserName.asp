<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="pcCPLog.asp" -->
<% 
Dim SPath
SPath=Request.ServerVariables("PATH_INFO")
SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
	strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
Else
	strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
End if

' attack was not submitted from the forgot_username page   close them out  
if Session("cp_Forgotusername")<>"1" then		  
	Session("cp_Forgotusername")=""
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1	
	
	call closeDb()
response.redirect "forgot_username.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end
end if
		
' attack was not submitted from this site  close them out 
if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "forgot_username.asp")) <>1 then
	Session("cp_Forgotusername")=""
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end  if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1		
	call closeDb()
response.redirect "forgot_username.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end			
end if
			
IF session("ForgotAttackCount") => 5 THEN 
	call closeDb()
response.redirect "forgot_username.asp"
	response.end()
END IF  
  
dim pemail, pusername, pAdminemail

pAdminEmail = getUserInput(request.querystring("email"), 150)

If lcase(Trim(pAdminEmail)) = lcase(Trim(scFrmEmail)) Then 

	err.clear
	' authenticated and charge session
	query="SELECT TOP 1 IDAdmin FROM admins WHERE AdminLevel='19';" 
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
		
	if err.number>0 then
					
		if session("ForgotAttackCount")="" then
			session("ForgotAttackCount")=0
		end if
		session("ForgotAttackCount")=session("ForgotAttackCount")+1
							
		call closeDb()
response.redirect "forgot_username.asp?msg=" & dictLanguageCP.Item(Session("language")&"_forgotusernameadminDBerror") & err.number
		response.end()
	end if

	if rs.eof then
		
		if session("ForgotAttackCount")="" then
			session("ForgotAttackCount")=0
		end if
		session("ForgotAttackCount")=session("ForgotAttackCount")+1						
	
		call closeDb()
response.redirect "forgot_username.asp?msg=" & dictLanguageCP.Item(Session("language")&"_forgotusernameadminerror") 
		response.end()
	else
		
		Dim IDAdmin,fromName,from,rcpt,subject,body
		session("ForgotAttackCount")=0
		IDAdmin = rs("IDAdmin")	
			
		
		fromName = scCompanyName  		
		from = scFrmEmail
		rcpt = scEmail
		subject = dictLanguageCP.Item(Session("language")&"_forgotusernameadminmailsubject") 
		body = Replace(dictLanguageCP.Item(Session("language")&"_forgotusernameadminmailbody1"),"#username",IDAdmin ) 	
	
		call sendMail (fromName, from, rcpt, subject, body)
		
		' SEnd an email to the store Admin 
		if Session("RedirectURL")<>"" then
			RedirectURL=Session("RedirectURL")
			Session("RedirectURL")=""
			call closeDb()
response.redirect RedirectURL
		else
			call closeDb()
response.redirect "login_1.asp?s=1&msg=" & dictLanguageCP.Item(Session("language")&"_forgotusernameadminsuccess") 
		end if
	end if 
Else
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1
		
	if Session("RedirectURL")<>"" then
			RedirectURL=Session("RedirectURL")
			Session("RedirectURL")=""
			call closeDb()
response.redirect RedirectURL
	else
			call closeDb()
response.redirect "forgot_username.asp?msg=" & dictLanguageCP.Item(Session("language")&"_forgotusernameadminerror") 
	end if
	
	response.end()
	
End if 
%>
<!--#include file="AdminFooter.asp"-->
