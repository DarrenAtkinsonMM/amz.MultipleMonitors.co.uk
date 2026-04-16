<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
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
           
' attack was not submitted from the forgot_password page   close them out  
if Session("cp_Forgotpassword")<>"1" then		  
	Session("cp_Forgotpassword")=""
	if session("ForgotAttackCount")="" then
	session("ForgotAttackCount")=0
	end if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1	
	
	call closeDb()
    response.redirect "forgot_password.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end
end if
		
' attack was not submitted from this site  close them out 
if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "forgot_password.asp")) <>1 then
	Session("cp_Forgotpassword")=""
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end  if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1		
	call closeDb()
    response.redirect "forgot_password.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end			
end if
	
IF session("ForgotAttackCount") => 5 THEN 
	call closeDb()
    response.redirect "forgot_password.asp"
	response.end()
END IF    

dim pemail, ppassword, pAdminPassword

pAdminUser = getUserInput(request.querystring("user"), 16)

err.clear

'// Start: Validate User
pcv_boolIsValid = False
If IsNumeric(pAdminUser) Then
    query = "SELECT idadmin, adm_ContactEmail FROM admins WHERE idadmin = " & pAdminUser & ";" 
    Set rs = server.CreateObject("ADODB.RecordSet")		
    Set rs = conntemp.execute(query)	
    If Not rs.Eof Then
        pcv_boolIsValid = True
        pcv_intIdAdmin = rs("IDAdmin")	
        pcv_strEmail = rs("adm_ContactEmail")
        If len(pcv_strEmail)=0 Or IsNull(pcv_strEmail) Then
            pcv_strEmail = scEmail
        End If
    End If
    Set rs = Nothing
End If
'// End: Validate User


If pcv_boolIsValid = True Then

	'// Success, check for errors...
    tmpResult = pcf_CreatePRGuidAdmin(pcv_intIdAdmin, pcv_strEmail)
			
    if err.number > 0 OR tmpResult = "0" then
        set rs=nothing
                
        if session("ForgotAttackCount")="" then
            session("ForgotAttackCount")=0
        end if
        session("ForgotAttackCount")=session("ForgotAttackCount")+1
                        
        call closeDb()
        response.redirect "forgot_password.asp?msg=" & dictLanguageCP.Item(Session("language")&"_forgotpasswordadminDBerror")& err.number
        response.end()
    end if

	'// Success, send email and redirect...
    if Session("RedirectURL")<>"" then
        RedirectURL=Session("RedirectURL")
        Session("RedirectURL")=""
        call closeDb()
        response.redirect RedirectURL
    else
        call closeDb()
        response.redirect "login_1.asp?s=1&msg=" & server.URLEncode(dictLanguageCP.Item(Session("language")&"_forgotpasswordadminsuccess"))
    end if
    
Else

    '// Failure, log and redirect...
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1								
	call closeDb()
    response.redirect "forgot_password.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end() 
     
End If
%>
