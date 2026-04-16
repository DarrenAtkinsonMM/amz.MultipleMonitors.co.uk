<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<!--#include file="../shop/includes/settings.asp"-->
<!--#include file="../shop/includes/storeconstants.asp"-->  
<!--#include file="../shop/includes/openDb.asp"-->
<!--#include file="../shop/includes/stringfunctions.asp"-->
<!--#include file="../shop/includes/currencyformatinc.asp"--> 
<!--#include file="../shop/includes/languages.asp" -->
<!--#include file="../shop/includes/emailsettings.asp"-->
<!--#include file="../shop/includes/securitysettings.asp" -->
<!--#include file="../shop/includes/sendmail.asp"-->
<!--#include file="../shop/includes/productcartFolder.asp"-->
<!--#include file="../shop/includes/secureadminfolder.asp"-->
<!--#include file="../shop/includes/pcServerSideValidation.asp" -->  
<!--#include file="../shop/includes/javascripts/pcClientSideValidation.asp" -->
<%
pcStrPageName = "contact.asp"

'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

dim conntemp, query, rs, rs2, ErrCheckEmail, pcv_intSuccess

Dim TurnOnSecurity 
TurnOnSecurity=scContact '1 = Turn On (Default) | 0 = Turn Off
	if not validNum(TurnOnSecurity) then
		TurnOnSecurity=0
	end if

Dim pcSecurityPath, strSiteSecurityURL

IF TurnOnSecurity=1 THEN
	pcSecurityPath=Request.ServerVariables("PATH_INFO")
	pcSecurityPath=mid(pcSecurityPath,1,InStrRev(pcSecurityPath,"/")-1)
	If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
		strSiteSecurityURL="http://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
	Else
		strSiteSecurityURL="https://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
	End if
END IF

pIdCustomer=session("idCustomer")
	if not validNum(pIdCustomer) then
		pIdCustomer=0
	end If

msg=getUserInput(request.querystring("msg"),0)

pcv_isNameRequired=True
pcv_isEmailRequired=True
pcv_isTitleRequired=True
pcv_isBodyRequired=True

if request.form("updatemode")="1" then

	'//set error to zero
	pcv_intErr=0

	pcs_ValidateEmailField	"FromEmail", pcv_isEmailRequired, 0
	pcs_ValidateTextField	"FromName", pcv_isNameRequired, 0
	pcs_ValidateTextField	"MsgTitle", pcv_isTitleRequired, 0
	pcs_ValidateTextField	"MsgBody", pcv_isBodyRequired, 0

	IF TurnOnSecurity=1 THEN
	
%>
    <!-- Include file for CAPTCHA configuration -->
    <!-- #include file="../shop/CAPTCHA/CAPTCHA_configuration.asp" --> 
     
    <!-- Include file for CAPTCHA form processing -->
    <!-- #include file="../shop/CAPTCHA/CAPTCHA_process_form.asp" -->   
<%	
	If not blnCAPTCHAcodeCorrect then
		response.redirect "contact.asp?msg=security2"
	else
		Session("store_postnum")=replace(request("postnum"),"'","''")
		pcv_Test=0
		if InStr(ucase(Request.ServerVariables("HTTP_REFERER")),ucase(strSiteSecurityURL & pcStrPageName))<>1 then
			session("store_postnum")=""
			session("store_num")=""
			pcv_test=1
		end if
		
		if pcv_Test=1 then
			if session("AttackCount")="" then
				session("AttackCount")=0
			end if
			session("AttackCount")=session("AttackCount")+1
			if session("AttackCount")>=scAttackCount then
					session("AttackCount")=0%>
					<!--#include file="../shop/includes/sendAlarmEmail.asp" -->
			<%end if	
			response.redirect pcStrPageName & "?msg=security1"
			response.end
		end if
		
		if pcv_Test=0 then
			if (session("store_num")="") OR (session("store_num")&"" <> Session("store_postnum")&"") then
				session("store_postnum")=""
				session("store_num")=""
				pcv_test=1
			end if
		end if

		if pcv_Test=1 then
			if session("AttackCount")="" then
				session("AttackCount")=0
			end if
			session("AttackCount")=session("AttackCount")+1
			if session("AttackCount")>=scAttackCount then
					session("AttackCount")=0%>
					<!--#include file="../shop/includes/sendAlarmEmail.asp" -->
			<%end if	
		end if
	end if
	END IF

	'//Email error for page
	If Session("ErrFromEmail")="" OR isNULL(Session("ErrFromEmail")) Then Session("ErrFromEmail")=0
	if Session("ErrFromEmail")="1" then
		pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_sendpassword_1"))
	else	
		'//generic error for page
		pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	end if
		
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError
	else
		CustName=Session("pcSFFromName")	
		CustEmail=Session("pcSFFromEmail")
		MsgTitle=dictLanguage.Item(Session("language")&"_Contact_9") & Session("pcSFMsgTitle")
		MsgTitle=replace(MsgTitle,"''","'")
					
		'// Add variables to body
		MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_Contact_6") & CustName & vbcrlf
		MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_Contact_7") & CustEmail & vbcrlf
		
			'// IF customer is logged in, add more information
			Dim pcHideCPlink
			pcHideCPlink=1 ' Change to 0 to include the link in the message to the store administrator
			if pIdCustomer>0 and pcHideCPlink=0 then
				'//	Generate link to customer edit page
				SPath1=Request.ServerVariables("PATH_INFO")
				mycount1=0
				do while mycount1<2
					if mid(SPath1,len(SPath1),1)="/" then
						mycount1=mycount1+1
					end if
					if mycount1<2 then
						SPath1=mid(SPath1,1,len(SPath1)-1)
					end if
				loop
				SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
				if Right(SPathInfo,1)="/" then
				else
					SPathInfo=SPathInfo & "/"
				end if
				dURL=SPathInfo & scAdminFolderName & "/login_1.asp?redirectUrl=" & Server.URLEnCode(SPathInfo & scAdminFolderName &  "/modcusta.asp?idcustomer=" & pIdCustomer)
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_Contact_8") & dURL & vbcrlf & vbcrlf
			end if
			'// END IF customer is logged in
			
		MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_Contact_5") & vbcrlf & vbcrlf
		MsgBody=MsgBody & Session("pcSFMsgBody")
		MsgBody=replace(MsgBody,"''","'")
		
		'// Prevent issues with Customer Service E-mail not being set (v4.5)
		Dim strCustServEmail
		strCustServEmail=scCustServEmail
		if trim(strCustServEmail)="" then strCustServEmail=scFrmEmail
		
		call sendmail (CustName,CustEmail,strCustServEmail,MsgTitle,MsgBody)
		pcv_intSuccess=1
	End If

End If

if pIdCustomer>0 AND msg="" then
	call openDb()
	query="SELECT name,lastName,email FROM customers WHERE idCustomer=" &pIdCustomer
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error on contact.asp: "&Err.Description) 
	end if
	Session("pcSFFromName")=rs("name") & " " & rs("lastName")
	Session("pcSFFromEmail")=rs("email")
	Session("pcSFMsgTitle")=""
	Session("pcSFMsgBody")=""
	set rs=nothing
	call closeDB()
end if

if not request("sub") = "" Then
	strSubject = request("sub")
end if

%>
 <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>International Delivery | Multiple Monitors</title>
<META NAME="Robots" CONTENT="noindex,follow" />
<link href="/styles/popcss.css" rel="stylesheet" media="screen, projection" type="text/css" />
<!--[if lt IE 8]>
<link href="/popie.css" rel="stylesheet"  media="screen,projection" type="text/css" />
<![endif]-->
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-5648327-3']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</head>
<body>

<script language="JavaScript">
<!--
	
function Form1_Validator(theForm)
{
	if (theForm.FromName.value == "")
  	{
			alert("<%response.write dictLanguage.Item(Session("language")&"_security_20")%>");
		    theForm.FromName.focus();
		    return (false);
	}

	if (theForm.FromEmail.value == "")
  	{
			alert("<%response.write dictLanguage.Item(Session("language")&"_security_21")%>");
		    theForm.FromEmail.focus();
		    return (false);
	}
	
	if (theForm.MsgTitle.value == "")
  	{
			alert("<%response.write dictLanguage.Item(Session("language")&"_security_22")%>");
		    theForm.MsgTitle.focus();
		    return (false);
	}
	
	if (theForm.MsgBody.value == "")
  	{
			alert("<%response.write dictLanguage.Item(Session("language")&"_security_23")%>");
		    theForm.MsgBody.focus();
		    return (false);
	}
	
	if (theForm.postnum.value == "")
  	{
			alert("<%response.write dictLanguage.Item(Session("language")&"_security_6")%>");
		    theForm.postnum.focus();
		    return (false);
	}
	
return (true);
}
//-->
</script>

<% if pcv_intSuccess<>1 then %>

					<h1>Email an Enquiry</h1>
		
				<% If msg<>"" or request("msg")<>"" then %>
					<%if request("msg")="security1" then%>
						<div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_security_2")%></div>
					<%else
						if request("msg")="security2" then%>
						<div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_security_6")%></div>
						<%else
							if msg<>"" then%>
								<div class="pcErrorMessage"><%=msg%></div>
							<%end if
						end if
					end if%>
				<% end if %>
				
                <%
				' Build secure URL
				dim strActionSSL
				strActionSSL="contact.asp"
				if scSSL="1" AND scIntSSLPage="1" then
					strActionSSL=replace((scSslURL&"/pop-pages/contact-pop.asp"),"//","/")
					strActionSSL=replace(strActionSSL,"https:/","https://")
					strActionSSL=replace(strActionSSL,"http:/","http://")
				end if
				strActionSSL = "/pop-pages/contact-pop.asp?email=ok"
				%>

				
				<form method="post" name="contact" action="<%=strActionSSL%>" onSubmit="return Form1_Validator(this)" class="pcForms">
				<input type="hidden" name="updatemode" value="1">
				<table class="pop-contact-tbl">
					<tr>
						<td width="25%">
							<p><%response.write dictLanguage.Item(Session("language")&"_Contact_2")%></p>
						</td>
						<td width="75%"> 
							<p>
							<input type="text" name="FromName" size="35" maxlength="70"
							value="<%=pcf_FillFormField ("FromName", pcv_isNameRequired) %>">
						<img src="/shop/pc/images/sample/pc_icon_required.gif" style="float:none;" /></p>
						</td>
					</tr>
					<tr>
						<td>
						<p><%response.write dictLanguage.Item(Session("language")&"_Contact_3")%></p>
						</td>
						<td>
						<p>
						<input type="text" name="FromEmail" size="35" maxlength="70"
						value="<%=pcf_FillFormField ("FromEmail", pcv_isEmailRequired) %>">					
						<img src="/shop/pc/images/sample/pc_icon_required.gif" style="float:none;" /></p>
						</td>
					</tr>
					<tr>
						<td>
						<p>
						<%response.write dictLanguage.Item(Session("language")&"_Contact_4")%>
						</p>
						</td>
						<td>
						<p>
							<input type="text" name="MsgTitle" size="35" maxlength="70"	value="<%=strSubject%>">	
						<img src="/shop/pc/images/sample/pc_icon_required.gif" style="float:none;" /></p>
						</td>
					</tr>
					<tr>
						<td valign="top">
							<p><%response.write dictLanguage.Item(Session("language")&"_Contact_5")%></p>
						</td>
						<td valign="top">
							<p>
							<textarea rows="10" name="MsgBody" cols="35"><%=pcf_FillFormField ("MsgBody", pcv_isBodyRequired) %></textarea>
						<img src="/shop/pc/images/sample/pc_icon_required.gif" style="float:none;" /></p>
						</td>
					</tr>

					<tr> 
						 <td>&nbsp;</td>
                        <td>
							<p><input type="image" name="submit" value="Send message" src="/shop/pc/images/pc/mm-send-enquiry.gif" id="submit"></p>
						</td>
					</tr>
				</table>
				</form>

<% else %>
<div id="pageContent">
<div class="detail-holder">
    <div class="product-detail">
	<div id="pcMain">		
		<table class="pcMainTable">
			<tr>
				<td>
					<h1>Thank You For your Enquiry</h1>
				</td>
			</tr>
			<tr>
				<td class="pcSectionTitle">
					<p>Thank you for the enquiry, we will get back to you as soon as possible.</p>
                    <p>In the mean time why not see if your question has been answered on our blog:<br /> <a href="/blog/" target="_parent">Multiple Monitors Blog</a></p>
				</td>
			</tr>
			<tr> 
				<td> 
					<p><a href="<% if pIdCustomer>0 then %>custpref.asp<% else %>default.asp<% end if %>"><img src="<%=rslayout("submit")%>"></a></p>
				</td>
			</tr>
		</table>
	</div>
    <div class="clear"></div>
</div><!-- END: product-detail -->
</div><!-- END: detail-holder --> 
<div class="clear"></div>
</div><!-- END: pageContent -->
<div class="clear"></div>
<% end if %>
</body>
</html>
