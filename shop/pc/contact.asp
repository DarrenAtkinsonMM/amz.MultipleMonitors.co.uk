<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="../includes/pcFormHelpers.asp" -->
<%
pcStrPageName = "contact.asp"

dim rs2, ErrCheckEmail, pcv_intSuccess

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

code=getUserInput(request.querystring("err"),0)
If code & "" = "" Then
	code = 0
End If

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
		if scCaptchaType="1" then
			blnCAPTCHAcodeCorrect=pcf_checkReCaptcha()
		else%>
			<!-- Include file for CAPTCHA configuration -->
			<!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
			 
			<!-- Include file for CAPTCHA form processing -->
			<!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
		<%end if
		If not blnCAPTCHAcodeCorrect then
			session("store_num")=""
			response.redirect "contact.asp?err=2"
		else
			pcv_Test=0
			if InStr(ucase(Request.ServerVariables("HTTP_REFERER")),ucase(strSiteSecurityURL & pcStrPageName))<>1 then
				session("store_num")=""
				pcv_test=1
			end if
			
			if (IsNull(session("store_num"))) OR (session("store_num")="") then
				session("store_num")=""
				pcv_test=1
			else
				pcv_test=0
			end if
			
			if pcv_Test=1 then
				if session("AttackCount")="" then
					session("AttackCount")=0
				end if
				session("AttackCount")=session("AttackCount")+1
				if session("AttackCount")>=scAttackCount then
						session("AttackCount")=0%>
						<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if	
				response.redirect pcStrPageName & "?err=1"
				response.end
			end if
			
		end if
	END IF
	
	session("store_num")=""

	'//Email error for page
	If Session("ErrFromEmail")="" OR isNULL(Session("ErrFromEmail")) Then Session("ErrFromEmail")=0
	if Session("ErrFromEmail")="1" then
		pcv_strGenericPageError = dictLanguage.Item(Session("language")&"_sendpassword_1")
	else	
		'//generic error for page
		pcv_strGenericPageError = dictLanguage.Item(Session("language")&"_Custmoda_18")
	end if
		
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		Session("message") = pcv_strGenericPageError
		response.redirect pcStrPageName
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
		call pcs_hookContactUsEmailSent(strCustServEmail)
		pcv_intSuccess=1
	End If

End If

if pIdCustomer>0 AND msg="" then
	
	query="SELECT name,lastName,email FROM customers WHERE idCustomer=" &pIdCustomer
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		call LogErrorToDatabase()
        set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	Session("pcSFFromName")=rs("name") & " " & rs("lastName")
	Session("pcSFFromEmail")=rs("email")
	Session("pcSFMsgTitle")=""
	Session("pcSFMsgBody")=""
	set rs=nothing
	
end if

query = "SELECT pcCPage_PageDesc FROM pcContactPageSettings;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
If Not rs.Eof Then
    pageContent = rs("pcCPage_PageDesc")
End If
set rs=nothing

%>
 
<script type=text/javascript>
function Form1_Validator(theForm)
{
	if (theForm.FromName.value == "")
  	{
			alert("<%= dictLanguage.Item(Session("language")&"_security_20")%>");
		    theForm.FromName.focus();
		    return (false);
	}

	if (theForm.FromEmail.value == "")
  	{
			alert("<%= dictLanguage.Item(Session("language")&"_security_21")%>");
		    theForm.FromEmail.focus();
		    return (false);
	}
	
	if (theForm.MsgTitle.value == "")
  	{
			alert("<%= dictLanguage.Item(Session("language")&"_security_22")%>");
		    theForm.MsgTitle.focus();
		    return (false);
	}
	
	if (theForm.MsgBody.value == "")
  	{
			alert("<%= dictLanguage.Item(Session("language")&"_security_23")%>");
		    theForm.MsgBody.focus();
		    return (false);
	}
	
return (true);
}
</script>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Contact Us">Contact Us</h3>
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
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
<div id="pcMain" class="container-fluid pcContact">
    <div class="row">
  
        <% if pcv_intSuccess<>1 then %>
      
			<% If pcf_HasHTMLContent(pageContent) Then %>
                <div class="pcPageDesc">
					<%= pcf_FixHTMLContentPaths(pageContent) %>
				</div>
			<% End If %>
      
			<%
				msg = ""
				Select Case code
				Case 1 : msg = dictLanguage.Item(Session("language")&"_security_2")
				Case 2 : msg = dictLanguage.Item(Session("language")&"_security_6")
				Case Else
					If Session("message") <> "" Then
						msg = Session("message")
						Session("message") = ""
					End If
				End Select

				If msg <> "" Then
					%><div class="pcErrorMessage"><%= msg %></div><%
				End If
				%>

            <%
            ' Build secure URL
            dim strActionSSL
            strActionSSL="contact.asp"
            if scSSL="1" AND scIntSSLPage="1" then
                strActionSSL=replace((scSslURL&"/"&scPcFolder&"/pc/contact.asp"),"//","/")
                strActionSSL=replace(strActionSSL,"https:/","https://")
                strActionSSL=replace(strActionSSL,"http:/","http://")
            end if
            
            Tn1=""
            Tn1=Tn1 & Year(Date()) & Month(Date()) & Day(Date()) & Hour(Time()) & Minute(Time())
            LenTn1=Len(Tn1)
                For dd=LenTn1+1 to 50
                    Randomize
                    myC=Fix(3*Rnd)
                    Select Case myC
                        Case 0: 
                            Randomize
                            Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
                        Case 1: 
                            Randomize
                            Tn1=Tn1 & Cstr(Fix(10*Rnd))
                        Case 2: 
                            Randomize
                            Tn1=Tn1 & Chr(Fix(26*Rnd)+97)		
                    End Select		
                    Randomize		
                Next
            
            session("store_num")=Tn1
			%>
 
			<form method="post" name="contact" action="<%=strActionSSL%>" onSubmit="return Form1_Validator(this)" class="form">
				<input type="hidden" name="updatemode" value="1">
      	

        	        <%
					formItemStart dictLanguage.Item(Session("language")&"_Contact_2"), "FromName", True
					%>
                    <input class="form-control" type="text" id="FromName" name="FromName" size="35" maxlength="70" value="<%=pcf_FillFormField ("FromName", pcv_isNameRequired) %>">
                    <% pcs_RequiredImageTagHorizontal "FromName", pcv_isNameRequired %>
           	        <%
					formItemEnd
					%>
          
        	        <%
                    formItemStart dictLanguage.Item(Session("language")&"_Contact_3"), "FromEmail", True
                    %>
                    <input class="form-control" type="text" id="FromEmail" name="FromEmail" size="35" maxlength="70"
                    value="<%=pcf_FillFormField ("FromEmail", pcv_isEmailRequired) %>">					
                    <% pcs_RequiredImageTagHorizontal "FromEmail", pcv_isEmailRequired %>
           	        <%
                    formItemEnd
					%>
          
        	        <%
					formItemStart dictLanguage.Item(Session("language")&"_Contact_4"), "MsgTitle", True
					%>
					<input class="form-control" type="text" id="MsgTitle" name="MsgTitle" size="35" maxlength="70"	value="<%=pcf_FillFormField ("MsgTitle", pcv_isTitleRequired) %>">	
					<% pcs_RequiredImageTagHorizontal "MsgTitle", pcv_isTitleRequired %>
           	        <%
					formItemEnd
					%>
          
        	        <%
					formItemStart dictLanguage.Item(Session("language")&"_Contact_5"), "MsgBody", True
					%>
					<textarea class="form-control" rows="10" id="MsgBody" name="MsgBody" cols="35"><%=pcf_FillFormField ("MsgBody", pcv_isBodyRequired) %></textarea>
					<% pcs_RequiredImageTagHorizontal "MsgBody", pcv_isBodyRequired %>
           	        <%
					formItemEnd
					%>
          
					<% IF TurnOnSecurity=1 THEN %>
                        <div class="form-group">

                            <%if scCaptchaType="1" then
                                call pcs_genReCaptcha()
                            else%>
                                <!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
                            <%end if%>

						</div>
					<% END IF%>
          
                    <div class="form-group">
                        <button class="pcButton pcButtonContinue btn btn-skin btn-wc btn-contact" id="FormSubmit" name="FormSubmit">
                            <span class="pcButtonText">Send Message</span>
                        </button>                       
                    </div>
      </form>

    <% else %>

        <p>
            <%= dictLanguage.Item(Session("language")&"_Contact_10")%>
        </p>
      
        <div class="pcFormItem"><hr></div>
              
        <div class="pcFormButtons">
      	    <a class="pcButton pcButtonContinue btn btn-skin btn-wc" href="<% if pIdCustomer>0 then %>custpref.asp<% else %>default.asp<% end if %>">
        	    <img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>" />
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
            </a>
        </div>
        
    <% end if %>
  </div>
</div>
</div>
				</div>
		    </div>
			 
    </section>	
    <!-- /Section: Welcome -->
<!--#include file="footer_wrapper.asp"-->
