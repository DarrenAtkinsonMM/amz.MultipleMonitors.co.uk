<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pcStrPageName="checkout.asp" %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<%
'// START - Check for SSL and redirect to SSL login if not already on HTTPS
call storeSSLRedirect("1")
'// END - check for SSL

'Capture any redirects
dim pcRequestRedirect
pcRequestRedirect=getUserInput(request("redirectUrl"),250)

If Session("SFStrRedirectUrl")<>"" AND pcRequestRedirect="" Then
Else
  Session("SFStrRedirectUrl")=pcRequestRedirect
End If

session("REGidCustomer")=""

dim pcPageMode

pcPageName="checkout.asp"

pcPageMode=request("cmode")
If pcPageMode="" Then
  pcPageMode=0
Else
  If NOT validNum(pcPageMode) Then
    pcPageMode=0
  End If
End If

'// Check If only PayPal Express or Pay with Amazon is enabled - begin       
If session("customerType")=1 Then
	query="SELECT gwcode FROM paytypes WHERE active=-1;"
Else
  query="SELECT gwcode FROM paytypes WHERE active=-1 and Cbtob=0;"
End If
set rsGWObj=server.CreateObject("ADODB.RecordSet")
set rsGWObj=conntemp.execute(query)

dim intPPECheck, intPPECnt, intPPEOnly
intPPECheck=0
intAmzCheck=0
intGWCnt=0
intAltCheckoutOnly=0
do until rsGWObj.eof
	gwCode=rsGWObj("gwcode")
	If gwCode="999999" Then
		intPPECheck=1
    ElseIf gwCode="88" Then
        intAmzCheck=1
	End If
	intGWCnt=intGWCnt+1
	rsGWObj.movenext
loop
set rsGWObj=nothing
If intGWCnt=1 AND (intPPECheck=1 OR intAmzCheck=1) Then
	intAltCheckoutOnly=1
End If

'FB-S
session("pcFBS_TurnOnOff")=0
query="SELECT pcFBS_TurnOnOff,pcFBS_AppID FROM pcFacebookSettings;"
set rs=connTemp.execute(query)
if not rs.eof then
	session("pcFBS_TurnOnOff")=rs("pcFBS_TurnOnOff")
	if IsNull(session("pcFBS_TurnOnOff")) OR session("pcFBS_TurnOnOff")="" then
		session("pcFBS_TurnOnOff")="0"
	end if
	session("pcFBS_AppID")=rs("pcFBS_AppID")
end if
set rs=nothing

If (pcPageMode=0 And request("EmailNotFound")="") And (session("pcFBS_TurnOnOff")="0") Then
  If intAltCheckoutOnly = 1 Then
  	response.redirect "viewcart.asp"
  End If
  response.redirect "onepagecheckout.asp"
End If
'// Check If only PayPal Express is enabled - End

'pcPageMode
'0=checkout
'1=login
'2=retreive password
'3=autologin
'4=retreive order code(s)

Dim strCCSLCheck
strCCSLcheck = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)
If Len(Trim(strCCSLCheck))>0 Then
    call closedb()
    response.redirect "viewcart.asp"
End If

If pcPageMode=2 Then
  pcFromPageMode=getUserInput(request("fmode"),1)
  If pcFromPageMode="" Then
    pcFromPageMode=2
  Else
    If not validNum(pcFromPageMode) Then
      pcFromPageMode=0
    End If
  End If
End If

If pcPageMode=4 Then
  pcFromPageMode=getUserInput(request("fmode"),1)
  If pcFromPageMode="" Then
    pcFromPageMode=4
  Else
    If not validNum(pcFromPageMode) Then
      pcFromPageMode=0
    End If
  End If
End If

If (Session("SFStrRedirectUrl")<>"" AND pcPageMode<>0) AND (session("idCustomer")<>0 and session("idCustomer")<>"") Then
  response.redirect "Login.asp?lmode=2"
End If

session("pcSFCMode")=pcPageMode

'Get path for Advanced Security
If scSecurity=1 Then
  Dim pcSecurityPath, strSiteSecurityURL

  pcSecurityPath=Request.ServerVariables("PATH_INFO")
  pcSecurityPath=mid(pcSecurityPath,1,InStrRev(pcSecurityPath,"/")-1)
  If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" Then
    strSiteSecurityURL="http://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
  Else
    strSiteSecurityURL="https://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
  End If
End If

If scResetPassCapt="1" Or scUseImgs=1 Then
	If (request.form("SubmitCO")<>"") OR (request("SubmitPM")<>"") Then
		if scCaptchaType="1" then
			blnCAPTCHAcodeCorrect=pcf_checkReCaptcha()
		else%>
			<!-- Include file for CAPTCHA configuration -->
			<!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
		 
			<!-- Include file for CAPTCHA form processing -->
			<!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" --> 
		<%end if
	End If
End If

'// Check If email is passed to create reset password GUID
If pcPageMode=2 And request("SubmitPM")<>"" Then

	If scResetPassCapt="1" Then
		
		If Not blnCAPTCHAcodeCorrect Then
			session("store_userlogin")=""
			session("store_adminre")=""
			session("store_num")=""
			call closedb()
			response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=6")
		End If

	End If
	
  pcv_intErr=0 'set to zero
  pcs_ValidateEmailField  "LoginEmail", true, 250
  pcStrEmail = Session("pcSFLoginEmail")

  query="SELECT idcustomer, name, lastname, email, [password] from customers WHERE email='" &pcStrEmail& "' AND (pcCust_Guest=0 OR pcCust_Guest=2)"
  
  set rs=server.CreateObject("ADODB.RecordSet")
  set rs=conntemp.execute(query)  
  If not rs.eof Then
    pcIntCustomerID=rs("idcustomer")
    pcStrName=rs("name")
    pcStrLastName=rs("lastname")
    pcStrEmail=rs("email")
    pcStrPassword=rs("Password")

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '// START No password, add now
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
    If trim(pcStrPassword)="" or IsNull(pcStrPassword) Then
      ' Generate random passwords:
      function randomNumber(limit)
      	randomize
        randomNumber=int(rnd*limit)+2
      End function
      pcStrCustomerPassword=randomNumber(99999999)
	  pcStrPassword=pcStrCustomerPassword
      pcStrCustomerPassword=pcf_PasswordHash(pcStrCustomerPassword)
      query="UPDATE customers SET [password]='"&pcStrCustomerPassword&"' WHERE idCustomer="& pcIntCustomerID
      set rsTemp=server.CreateObject("ADODB.RecordSet")
      set rsTemp=conntemp.execute(query)
      set rsTemp=nothing  
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '// START No password, add now
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  	set rs=nothing
    
    tmpResult=pcf_CreatePRGuid(pcIntCustomerID,pcStrEmail)
	
    call closedb()
	
	If tmpResult="0" Then
		If pcFromPageMode=2 Then
        	response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=&EmailNotFound=1"
      	Else
        	response.redirect "checkout.asp?cmode="&pcFromPageMode&"&EmailNotFound=1"
      	End If
	Else
		If pcFromPageMode=2 Then
		  response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=&EmailNotFound=0"
		Else
		  response.redirect "checkout.asp?cmode="&pcFromPageMode&"&EmailNotFound=0"
		End If
	End if      
  Else		
    'Guest
    set rs=nothing
    query="SELECT idcustomer from customers WHERE email='" &pcStrEmail& "' AND pcCust_Guest=1;"
    set rs=connTemp.execute(query)
    If not rs.eof Then
      set rs=nothing      
      call closedb()
      If pcFromPageMode=2 Then
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=&msgmode=7"
      Else
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&msgmode=7"
      End If      
    Else    
      set rs=nothing      
      call closedb()
      If pcFromPageMode=2 Then
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=&EmailNotFound=1"
      Else
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&EmailNotFound=1"
      End If      
    End If
    
  End If
End If

if pcPageMode=4 AND request("SubmitPM")<>"" then
  pcv_intErr=0 'set to zero
  pcs_ValidateEmailField  "LoginEmail", true, 250
  pcStrEmail = Session("pcSFLoginEmail")

  query="SELECT idcustomer, name, lastname, email, [password], pcCust_Guest from customers WHERE email LIKE '" &pcStrEmail& "';"  
 
  set rs=server.CreateObject("ADODB.RecordSet")
  set rs=conntemp.execute(query)  
  If not rs.eof Then
    pcIntCustomerID=rs("idcustomer")
    pcStrName=rs("name")
    pcStrLastName=rs("lastname")
    pcStrEmail=rs("email")
    pcStrPassword=rs("Password")
    pcv_Guest=rs("pcCust_Guest")
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '// START No password, add now
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
    If trim(pcStrPassword)="" or IsNull(pcStrPassword) Then
      ' Generate random passwords:
      function randomNumber(limit)
        randomize
        randomNumber=int(rnd*limit)+2
      End function
      pcStrCustomerPassword=randomNumber(99999999)
	  pcStrPassword=pcStrCustomerPassword
      pcStrCustomerPassword=pcf_PasswordHash(pcStrCustomerPassword)
      query="UPDATE customers SET [password]='"&pcStrCustomerPassword&"' WHERE idCustomer="& pcIntCustomerID
      set rsTemp=server.CreateObject("ADODB.RecordSet")
      set rsTemp=conntemp.execute(query)
      set rsTemp=nothing  
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '// END No password, add now
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    query="SELECT pcOrd_OrderKey,orderDate FROM Orders INNER JOIN Customers ON Orders.IdCustomer=Customers.IdCustomer WHERE Customers.email LIKE '" &pcStrEmail& "' AND Orders.OrderStatus>1 ORDER BY Orders.OrderDate ASC;"
    set rs=connTemp.execute(query)
    If not rs.eof Then
      pcArr=rs.getRows()
      intCount=ubound(pcArr,2)
      tmpOrderCodes=""
      For i=0 to intCount
        tmpOrderCodes=tmpOrderCodes & dictLanguage.Item(Session("language")&"_forgotordercodesmailbody1") & pcArr(0,i) & " - " & dictLanguage.Item(Session("language")&"_forgotordercodesmailbody2") & pcArr(1,i) & vbcrlf
      Next
      pcStrSubject=dictLanguage.Item(Session("language")&"_forgotordercodesmailsubject")
      pcStrBody=dictLanguage.Item(Session("language")&"_forgotordercodesmailbody")
      pcStrBody=replace(pcStrBody,"#ordercodes",tmpOrderCodes)  
      pcStrBody=replace(pcStrBody,"#firstname",pcStrName)      
      pcStrBody=replace(pcStrBody,"#lastname",pcStrLastName)
      call sendmail (scEmail, scEmail, pcStrEmail, pcStrSubject, pcStrBody) 
	  
	  call pcs_hookForgotOrderCodeEmailSent(pcStrEmail)
	  
      set rs=nothing
      call closedb()
      If pcFromPageMode=4 Then
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=4&ENotFound=0"
      Else
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&ENotFound=0"
      End If
    Else
      set rs=nothing
      call closedb()
      If pcFromPageMode=4 Then
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=4&ENotFound=2"
      Else
        response.redirect "checkout.asp?cmode="&pcFromPageMode&"&ENotFound=2"
      End If
    End If
      
  Else
    'customer not found..
    set rs=nothing
    call closedb()
    If pcFromPageMode=4 Then
      response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=4&ENotFound=1"
    Else
      response.redirect "checkout.asp?cmode="&pcFromPageMode&"&ENotFound=1"
    End If
  End If
End If

session("availableShipStr")=""
session("provider")=""
pcAutoLoginAllowed=0

'FB-S/E
if (request.form("SubmitCO")<>"") OR (pcPageMode=3) OR (((pcPageMode=1) OR (pcPageMode=0)) AND (request.form("fb")="1")) then
  pcv_intErr=0 'set to zero
  FBAutoLogin=0
  'Autologin
  If pcPageMode=3 Then
    'check If admin is logged in
    If session("admin")=-1 Then
      pcAutoLoginAllowed=1
    End If
    
    If pcAutoLoginAllowed=1 Then

      '// Request "LoginPassword", trim, and set to Session
      pcStrLoginPassword = session("ppassword")
      If len(pcStrLoginPassword)>0 Then
        session("pcSFPassWordExists")="YES"
		session("pcSFPassNotEnter")="1"
        session("pcSFLoginPassword") = pcStrLoginPassword
        session("pcSFLoginPassword")=Decrypt(session("pcSFLoginPassword"),9286803311968)
      End If
      session("ppassword") = ""
      If len(session("pcSFLoginEmail"))<1 AND session("idCustomer")=0 Then
        response.redirect("checkout.asp?cmode=1&msgcode=1")
      End If
    Else
      response.redirect("checkout.asp?cmode=1")
    End If
    'End Autologin
  Else
		'FB-S
		IF (getUserInput(request("fl"),250)<>"") AND (session("pcFBS_TurnOnOff")="1") THEN
			pcs_ValidateEmailField "fe", true, 0
			session("pcSFLoginEmail")=session("pcSFFe")
			if len(session("pcSFLoginEmail"))<1 AND session("idCustomer")=0 then
				response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=1")
			end if
			pcs_ValidateTextField "fl", true, 0
			
			query="SELECT [password] FROM Customers WHERE [email] LIKE '" & session("pcSFLoginEmail") & "' AND pcCust_FBId LIKE '" & session("pcSFFl") & "';"
			set rs=connTemp.execute(query)
			if not rs.eof then
				session("pcSFPassWordExists")="YES"
				session("pcSFLoginPassword")=rs("password")
				session("pcSFPassNotEnter")="1"
				set rs=nothing
			else
				pcs_ValidateTextField	"ffn", true, 0
				pcs_ValidateTextField	"fln", true, 0
				Tn1=""
				Tn1=Tn1 & Year(Date()) & Month(Date()) & Day(Date()) & Hour(Time()) & Minute(Time())
				LenTn1=Len(Tn1)
				For dd=LenTn1+1 to 25
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
				pcStrCustomerPassword=Tn1
				session("pcSFPassWordExists")="YES"
				Session("pcSFLoginPassword")=pcStrCustomerPassword
				session("pcSFPassNotEnter")="0"
				query="INSERT INTO Customers ([Email],[Password],[name],lastName,pcCust_FBId) VALUES ('" & session("pcSFLoginEmail") & "','" & pcf_PasswordHash(Session("pcSFLoginPassword")) & "',N'" & Session("pcSFFFn") & "',N'" & Session("pcSFFln") & "','" & Session("pcSFFl") & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			FBAutoLogin=1
		ELSE
		'FB-E
    pcs_ValidateEmailField  "LoginEmail", true, 0
		'pcStrLoginEmail=replace(request.form("LoginEmail"),"'","''")
		'session("pcSFLoginEmail")=pcStrLoginEmail
		'if pcStrLoginEmail="" then
		'	pcv_intErr=pcv_intErr+1
		'End if
    
    '// Request "LoginPassword", trim, and set to Session
    pcs_ValidateTextField "LoginPassword", false, 0
		'pcStrLoginPassword=request.form("LoginPassword")		
		'session("pcSFLoginPassword")=pcStrLoginPassword

    pcs_ValidateTextField "PassWordExists", false, 0
		'session("pcSFPassWordExists")=request.Form("PassWordExists")
    
    'If pcStrLoginPassword="" AND session("pcSFPassWordExists")="YES" Then
    If session("pcSFLoginPassword")="" AND session("pcSFPassWordExists")="YES" Then   
      pcv_intErr=pcv_intErr+1
    End If

    'If len(pcStrLoginEmail)<1 AND session("idCustomer")=0 Then
    If len(session("pcSFLoginEmail"))<1 AND session("idCustomer")=0 Then
      response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=1")
    End If    
		'FB-S
		END IF
		'FB-E	
	end if
  
  If (session("ErrLoginEmail")="") AND (pcAutoLoginAllowed=0) AND (FBAutoLogin=0) Then

    If scSecurity=1 AND ((scUserLogin=1 AND session("pcSFPassWordExists")="YES") OR (scUserReg=1 AND session("pcSFPassWordExists")<>"YES")) Then
      pcv_Test=0
      If (session("store_userlogin")<>"1") AND (session("store_adminre")<>"1") Then
        session("store_userlogin")=""
        session("store_adminre")=""
        pcv_test=1
      End If
      If pcv_Test=0 AND session("store_adminre")<>"1" Then
        If InStr(ucase(Request.ServerVariables("HTTP_REFERER")),ucase(strSiteSecurityURL & "checkout.asp"))<>1 Then
          session("store_userlogin")=""
          session("store_adminre")=""
          pcv_test=1
        End If
        session("store_adminre")=""
      End If
      If pcv_Test=0 AND scUseImgs=1 Then
        If not blnCAPTCHAcodeCorrect Then
          session("store_userlogin")=""
          session("store_adminre")=""
          pcv_test=1
          response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=6")
        End If
      End If

      If pcv_Test=1 Then
        If scAlarmMsg=1 Then
          If session("AttackCount")="" Then
            session("AttackCount")=0
          End If
          session("AttackCount")=session("AttackCount")+1
          If session("AttackCount")>=scAttackCount Then
            session("AttackCount")=0%>
            <!--#include file="../includes/sEndAlarmEmail.asp" -->
          <%End If  
        End If

        response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=4")
      End If          
    End If

  End If
  
  If pcv_intErr=0 Then
    erypassword=encrypt(session("pcSFLoginPassword"), 9286803311968)
    session("pcSFEryPassword")=erypassword		
    If pcPageMode=0 Then
			'FB-S
			if (session("pcFBS_TurnOnOff")="1") THEN
				response.redirect "login.asp?lmode=0"
			else
				response.redirect "onepagecheckout.asp"
			end if
			'FB-E
    Else
      'just logging in
      response.redirect "login.asp?lmode=2"
    End If
  Else
    '// handle error
  End If
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Section C - Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script type=text/javascript>"&vbcrlf
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

pcs_JavaTextField "LoginEmail", pcv_isLoginEmailRequired, dictLanguage.Item(Session("language")&"_validate_1"), ""
  
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: FORM VALIDATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Customer Support Area">Customer Support Area</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->

	<section id="pc-multisection" class="pc-multisection paddingbot-40">
		<div class="container">
			<div class="row">

<div id="pcMain" class="container-fluid pcCheckout">
  <div class="row">
  <%	
    '// Generate payment type query
    If session("customerType")=1 Then
      query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 AND (gwCode=999999 OR gwCode=46 OR gwCode=53) ORDER by paymentPriority;"
    Else
      query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 and Cbtob=0 AND (gwCode=999999 OR gwCode=46 OR gwCode=53) ORDER by paymentPriority;"
    End If  
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)    
		        
    If NOT rs.eof Then
      intPayPalExp=1
      '// Determine which API to use (US or UK)
      query="SELECT pcPay_PayPal.pcPay_PayPal_Partner, pcPay_PayPal.pcPay_PayPal_VEndor FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
      set rsPayPalType=Server.CreateObject("ADODB.Recordset")
      set rsPayPalType=conntemp.execute(query)
      pcPay_PayPal_Partner=rsPayPalType("pcPay_PayPal_Partner")
      pcPay_PayPal_VEndor=rsPayPalType("pcPay_PayPal_VEndor")
      If pcPay_PayPal_Partner<>"" AND pcPay_PayPal_VEndor<>"" Then  
        pcPay_PayPal_Version = "UK"     
      Else
        pcPay_PayPal_Version = "US"           
      End If
      set rsPayPalType=nothing          
    Else
      intPayPalExp=0
    End If
    set rs=nothing
    
    '====================================
    ' START: PayPal Express Button
    '==================================== 
    If intPayPalExp=1 AND pcPageMode=0 Then %>
			<form name="PayPalCheckout" class="pcForms">
        <div id="pcPayPalCheckout">
          <div class="nsFormItem">
            <div class="nsFormGroupLabelLong">
              <label>
              <% '// Display the API Button Code
              If pcPay_PayPal_Version = "US" Then %>
                <a href="pcPay_ExpressPay_Start.asp"><img src="<%=pcf_getImagePath("https://www.paypal.com/en_US/i/btn","btn_xpressCheckout.gif")%>" alt="Acceptance Mark"></a>
              <% Else %>
                <a href="pcPay_ExpressPayUK_Start.asp"><img src="<%=pcf_getImagePath("https://www.paypal.com/en_US/i/btn","btn_xpressCheckout.gif")%>" alt="Acceptance Mark"></a>
              <% End If %>
              </label>
            </div>
          </div>
        </div>
      </form>   
    <% End If
    '====================================
    ' END: PayPal Express Button
    '====================================
    %>

    
    <% 
			'====================================
			' START: Error Message
			'==================================== 
			msgMode=getUserInput(request.querystring("msgmode"),1)
			select case msgMode
				case "2"
					msg=dictLanguage.Item(Session("language")&"_validate_2")
					msgClass="pcErrorMessage"
				case "3"
					msg=dictLanguage.Item(Session("language")&"_validate_3")
					msgClass="pcErrorMessage"
				case "4"
					msg=dictLanguage.Item(Session("language")&"_validate_2")
					msgClass="pcErrorMessage"
				case "5"
					msg=dictLanguage.Item(Session("language")&"_validate_4")
					msgClass="pcInfoMessage"
				case "6"
					msg=dictLanguage.Item(Session("language")&"_security_3")
					msgClass="pcErrorMessage"
				case "7"
					msg=dictLanguage.Item(Session("language")&"_validate_5")
					msgClass="pcInfoMessage"
					tmpemail=pcf_FillFormField("LoginEmail", true)
				case "8"
					msg=dictLanguage.Item(Session("language")&"_validate_6")
					msgClass="pcErrorMessage"
					tmpemail=pcf_FillFormField("LoginEmail", true)
				case Else
					msg=""
			End select
			
      If msg<>"" Then 
			%>
      	<div class="<%=msgClass%>">
        	<%=msg%>
        </div>
      <% 
      End If
      
    '====================================
    ' END: Error Message
    '====================================
    colClass = "col-md-6"
    If pcPageMode=2 Or pcPageMode=4 Then
        colClass = "col-md-12"
    End If
    %>
    <div id="pcLoginForm" class="<%= colClass %>">   
            
        <form name="LoginForm" id="LoginForm" method="post" action="<%=pcPageName%>" onSubmit="return Form1_Validator(this)" class="form" role="form">

            <input type="hidden" name="cmode" value="<%=pcPageMode%>">
            <%'FB-S
            if pcPageMode="0" OR pcPageMode="1" then %>
            <input type="hidden" id="fe" name="fe" value="">
            <input type="hidden" id="fl" name="fl" value="">
            <input type="hidden" id="ffn" name="ffn" value="">
            <input type="hidden" id="fln" name="fln" value="">
            <input type="hidden" id="fb" name="fb" value="1">
            <% end if
            'FB-E%>
          
            <%
            '====================================
            ' START: Page Title
            '==================================== 
            %>
            <%
            If pcPageMode=2 Then
              pcPageTitle=dictLanguage.Item(Session("language")&"_newpass_16")
            Else
              If pcPageMode=4 Then
                pcPageTitle=dictLanguage.Item(Session("language")&"_checkout_29")
              Else
                pcPageTitle=dictLanguage.Item(Session("language")&"_checkout_23")
              End If
            End If
            %> 
            <div class="pcFormItem"><h1><%=pcPageTitle%></h1></div>
            <div class="pcSpacer"></div>
			<%If (pcPageMode=2) AND (request("EmailNotFound")<>"0") Then%>
				<%If request("new")="1" then%>
				<div class="pcFormItem">
					<p><%=dictLanguage.Item(Session("language")&"_resetpass_2")%></p>
				</div>
				<%End if%>
				<div class="pcFormItem">
					<p><%=dictLanguage.Item(Session("language")&"_resetpass_1")%></p>
				</div>
			<%End if%>
            <%
            '====================================
            ' END: Page Title
            '==================================== 
            %>


            <%'FB-S
            IF (session("Facebook")="1" OR session("pcFBS_TurnOnOff")="1") AND (pcPageMode<>2) THEN%>
            <div class="pcFormItem">
                <div class="pcFormLabel">
                    <div id="fb-root"></div>
                    <script type=text/javascript>
					function statusChangeCallback(response) {
						if (response.status === 'connected') {
					  		// Logged into your app and Facebook.
					  		loginPCAPI();
						} else if (response.status === 'not_authorized') {
						} else {
						}
					}
					
					function checkLoginState() {
						FB.getLoginStatus(function(response) {
					  		statusChangeCallback(response);
						});
					}
                      window.fbAsyncInit = function() {
                      FB.init({
                        appId      : '<%=session("pcFBS_AppID")%>', // App ID
                        cookie     : true, // enable cookies to allow the server to access the session
                        xfbml      : true,  // parse XFBML
						version    : 'v2.5' // use version 2.5
                      });
                      
						FB.getLoginStatus(function(response) {
						statusChangeCallback(response);
						});

                      };
                      (function(d, s, id) {
						var js, fjs = d.getElementsByTagName(s)[0];
						if (d.getElementById(id)) return;
						js = d.createElement(s); js.id = id;
						js.src = "//connect.facebook.net/en_US/sdk.js";
						fjs.parentNode.insertBefore(js, fjs);
						}(document, 'script', 'facebook-jssdk'));
					  
                  
						function loginPCAPI() {
							FB.api('/me', {fields: 'email,name,first_name,last_name,id'},function(response) {
							document.getElementById("fe").value=response.email;
							document.getElementById("fl").value=response.id;
							document.getElementById("ffn").value=response.first_name;
							document.getElementById("fln").value=response.last_name;
							document.LoginForm.submit();
                        });
                      }
                    </script>
					<fb:login-button show-faces="false" size="medium" width="200" max-rows="1" scope="email" onlogin="checkLoginState();">Login with Facebook</fb:login-button>
                </div>
            </div>
            <div class="pcFormItem">
                <div class="pcFormLabel">
                <h3>Or <%=dictLanguage.Item(Session("language")&"_opc_3")%></h3>
                </div>
            </div>
            <%END IF
            'FB-E
            %>



			<%
            '====================================
            ' START: Account Login/Forgot Fields
            '==================================== 
            If pcPageMode=2 Then
              pcIntEmailNotFound=getUserInput(request("EmailNotFound"),1)
              If Not ValidNum(pcIntEmailNotFound) Then
                pcIntEmailNotFound=""
              End If
              If pcIntEmailNotFound<>"" Then %>
                <% If pcIntEmailNotFound=1 Then %>
                  <div class="pcErrorMessage">
                    <%=dictLanguage.Item(Session("language")&"_forgotpasswordexec_2") %>
                  </div>
                <% Else %>
                  <div class="pcSuccessMessage">
                    <%= dictLanguage.Item(Session("language")&"_checkout_11")%>
                  </div>
                <% End If %>
              <%End If
            End If
            
            pcIntOPCEmailNotFound=getUserInput(request("ENotFound"),1)
            If session("ErrLoginEmail")<>"" Then
              session("PCErrLoginEmail")="1"
              %>
                <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_Custmoda_16")%></div>
              <% 
            Else
              If pcPageMode=4 Then
                If Not ValidNum(pcIntOPCEmailNotFound) Then
                  pcIntOPCEmailNotFound=""
                End If
                
                If pcIntOPCEmailNotFound<>"" Then
                  notFoundClass = "pcErrorMessage"
                  
                  Select Case pcIntOPCEmailNotFound
                  Case 1:
                    notFoundText = dictLanguage.Item(Session("language")&"_checkout_30")
                  Case 2:
                    notFoundText = dictLanguage.Item(Session("language")&"_checkout_31")
                  Case 3:
                    notFoundText = dictLanguage.Item(Session("language")&"_checkout_35")
                  Case Else:
                    notFoundClass = "pcSuccessMessage"
                    notFoundText = dictLanguage.Item(Session("language")&"_checkout_32")
                  End Select
                  %>
                    <div class="<%= notFoundClass %>">
                      <%= notFoundText %>
                    </div>
                  <%
                End If
              End If
              
              If pcIntEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"3" Then
                %>
                <div class="form-group">
                    <label for="LoginEmail"><%=dictLanguage.Item(Session("language")&"_Custmoda_4")%><div class="pcRequiredIcon"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div></label>
                    <input name="LoginEmail" type="email" class="form-control" maxlength="254" placeholder="you@domain.com" autocomplete="off" value="<%=pcf_FillFormField("LoginEmail", true)%>" />
                    <% pcs_RequiredImageTagHorizontal "LoginEmail", true %>
                </div>
				<%If pcPageMode=2 AND scResetPassCapt="1" Then %>
                        <div class="form-group">
							<%if scCaptchaType="1" then
								call pcs_genReCaptcha()
							else%>
								<!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
							<%end if%>
						</div>
				<%End If%>
                <%
              End If
            End If
            
            If pcPageMode<>2 And pcPageMode<>4 Then
            %>
 
              <div class="pcFormItem">
                <div class="pcFormField">
                  <%=dictLanguage.Item(Session("language")&"_checkout_25")%>
                </div>
              </div>

            <%
              '// Show or hide user login option
              If scSecurity=1 AND scUserLogin=1 Then
                pcShowLoginStyle="" 
              Else
                pcShowLoginStyle="none"
              End If
              
              '// Show or hide user registration option
              If scSecurity=1 AND scUserReg=1 Then
                pcShowStyle=""
              Else
                pcShowStyle="none"
              End If 
            %>            
              <div class="pcFormItem">
                <div class="pcFormLabel">
                  <input name="PassWordExists" type="radio" value="YES" checked="checked" <% If scUseImgs=1 Then%>onClick="document.getElementById('show_security').style.display='<%=pcShowLoginStyle%>'"<% End If%> class="clearBorder">
                  <label>&nbsp;<%=dictLanguage.Item(Session("language")&"_checkout_26")%></label>
                </div>
                <div class="pcFormField">
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input class="form-control form-control-inline" name="LoginPassword" type="password" maxlength="254" placeholder="" autocomplete="off" data-hint="" onFocus="document.LoginForm.PassWordExists[0].checked=true" />
                </div>
              </div>
        
                <div class="pcFormItem">
                  <div class="pcFormLabel">
                    <input name="PassWordExists" type="radio" value="NO" onFocus="document.LoginForm.LoginPassword.value=''" <% If scUseImgs=1 Then%>onClick="document.getElementById('show_security').style.display='<%=pcShowStyle%>'"<% End If %> class="clearBorder">
                    <label>&nbsp;<%=dictLanguage.Item(Session("language")&"_checkout_27")%></label>
                  </div>
                </div>		
                <% 
                '// Advanced Security (CAPTCHA)
                If scSecurity=1 Then
                    Session("store_userlogin")="1"
                    session("store_adminre")="1"  
                    If (scUserLogin=1 OR scUserReg=1) and (scUseImgs=1) Then %>
						<%if scCaptchaType="1" then
							call pcs_genReCaptcha()
						else%>
	                        <!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
						<%end if%>
                    <% Else 
                        response.write "<div id=""show_security""></div>"
                    End If
                Else
                    response.write "<div id=""show_security""></div>"
                End If
                
        End If
        %>

        <div class="pcSpacer"></div>

			<%
            If pcPageMode=2 OR pcPageMode=4 Then
            %>
              <div class="pcFormButtons">
                <input type="hidden" name="fmode" value="<%=pcFromPageMode%>">
                <%
                  If pcIntEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"3" Then
                  %>
                    <button class="pcButton pcButtonSubmit btn btn-skin btn-wc btn-contact" id="SubmitPM" name="SubmitPM" Value="Submit">
                      <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
                    </button>
                  <% 
                  End If 
                %>
                <a class="pcButton pcButtonBack" href="javascript:<%If pcFromPageMode=2 Then%>location='onepagecheckout.asp';<%Else%><%If pcFromPageMode=4 Then%>location='checkout.asp?cmode=1';<%Else%>history.go(-1);<%End If%><%End If%>">
                  <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
                </a>   
              </div>
            <%
            Else
            %>
              <div class="pcFormButtons">
                <% if pcPageMode=1 OR pcPageMode=3 then %>
                  <button class="pcButton pcButtonSubmit" id="SubmitCO" name="SubmitCO" value="Submit">
                    <img src="<%=pcf_getImagePath("", rslayout("submit"))%>" alt="Submit" />
                    <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
                  </button>
                <% else %>
                  <button class="pcButton pcButtonSubmit" id="SubmitCO" name="SubmitCO" value="Submit">
                    <img src="<%=pcf_getImagePath("",rslayout("login_checkout"))%>" alt="Submit" />
                    <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_checkout") %></span>
                  </button>
               	<% end if %>
              </div>
            <%
            End If
						'====================================
						' END: Account Login/Forgot Fields
						'==================================== 
					%>
							 
					<%
						'====================================
						' START: Forgot Password Link
						'====================================
						If pcPageMode<>2 AND pcPageMode<>4 Then
							pcIntEmailNotFound=getUserInput(request("EmailNotFound"),1)
							If Not ValidNum(pcIntEmailNotFound) Then
								pcIntEmailNotFound=""
							End If
				
							If pcIntEmailNotFound<>"" AND session("PCErrLoginEmail")<>"1" Then
								If pcIntEmailNotFound=1 Then
								%>
									<div class="pcErrorMessage">
										<% = dictLanguage.Item(Session("language")&"_forgotpassworderror") %>
									</div>
								<%
								Else 
								%>
									<div class="pcSuccessMessage">
										<% = dictLanguage.Item(Session("language")&"_checkout_11" )%>
									</div>
								<%
								End If
							Else
							%>
								<div class="pcSpacer"></div>
											
                <div class="pcFormItem">
                  <p>
                    <img src="<%=pcf_getImagePath("images","pc_icon_info.png")%>" alt="<% = dictLanguage.Item(Session("language")&"_Custva_3")%>" style="margin-right: 5px;padding-top: 2px;float: left;" />
                    <span>
                      <% = dictLanguage.Item(Session("language")&"_Custva_3")%><br/>
                      <a href="<%=Server.HtmlEncode("checkout.asp?cmode=2&fmode="&pcPageMode)%>"><%=dictLanguage.Item(Session("language")&"_Custva_8")%></a>
                    </span>
                  </p>
                </div>
							<%
							End If
							
							session("PCErrLoginEmail")=""
						End If
						'====================================
						' END: Forgot Password Link
						'====================================
					%>
        </div>
      </form>
                
    <%
    '====================================
    ' START: Order Review Section
    '====================================
    If (not (pcPageMode=4 and pcFromPageMode=1)) AND _
        (not (pcPageMode=4 and pcFromPageMode=4)) AND _
        (not (pcPageMode=2 and pcFromPageMode=1)) AND _
        (not (pcPageMode=2 and pcFromPageMode=2)) AND _
        (request("orderReview")<>"no") Then
        %>
        <div id="pcOrderReviewForm" class="<%=colClass %>">    
      	
            <form id="ORVForm" name="ORVForm" class="form" role="form">

                <% '// Order Review Header %>
                <div class="pcFormItem">
                    <h1><%=dictLanguage.Item(Session("language")&"_opc_checkout_1")%></h1>
                </div>
				
                <div class="pcSpacer"></div>
            
                <% '// Order Review Description %>
                <div class="pcFormItem">
                    <p><%=dictLanguage.Item(Session("language")&"_opc_checkout_2")%></p>
                </div>
    
                <div class="pcSpacer"></div>
                
                <% '// Order Review Email %>
                <div class="form-group">
                    <label for="email"><%=dictLanguage.Item(Session("language")&"_opc_5")%></label>
                    <input class="form-control" name="custemail" type="email" maxlength="254" placeholder="you@domain.com" autocomplete="off" data-hint="" value="<%=tmpemail%>" />
                </div>

                          
                <% '// Order Review Order Code %>
                <div class="form-group">
                    <label for="email"><%=dictLanguage.Item(Session("language")&"_opc_checkout_3")%></label>
                    <input class="form-control" name="ordercode" type="text" maxlength="254" placeholder="" autocomplete="off" data-hint="" />
                </div>                

    
                  <% '// Order Review Loader %>
                  <div id="ORVLoader" style="display:none"></div>
        
                  <div class="pcSpacer"></div>

                  <% '// Order Review Submit Button %>
                  <div class="pcFormButtons">
                    <button class="pcButton pcButtonSubmit" id="ORVSubmit" name="ORVSubmit">
                      <img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="Submit" />
                      <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
                    </button>            
                  </div>
           
                
        	<% 
            '====================================
            ' START: Forgot Order Number(s) Link
            '====================================
            If pcPageMode<>2 AND pcPageMode<>4 Then
                pcIntEmailNotFound=getUserInput(request("ENotFound"),1)
                If Not ValidNum(pcIntEmailNotFound) Then
                    pcIntEmailNotFound=""
                End If
                %>
            <div class="pcSpacer"></div>
               
                <div class="pcFormItem">
                  <p>
                    <img src="<%=pcf_getImagePath("images","pc_icon_info.png")%>" alt="<% = dictLanguage.Item(Session("language")&"_checkout_33")%>" style="margin-right: 5px;padding-top: 2px;float: left;">
                    <span>
                      <%= dictLanguage.Item(Session("language")&"_checkout_33")%><br/>
                      <a href="<%=Server.HtmlEncode("checkout.asp?cmode=4&fmode=4")%>"><%= dictLanguage.Item(Session("language")&"_checkout_34")%></a>
                    </span>
                  </p>
                </div>
          		<%  
						End If
						'====================================
						' START: Forgot Order Number(s) Link
						'====================================
        	%>
        	</div>
      	</form>
    	<% End If %>

  <script type=text/javascript>
    $pc(document).ready(function()
    {

        //*Validate Order Review Form
        $pc("#ORVForm").validate({
            rules: {
                custemail: 
                {
                    required: true,
                    email: true
                },
                ordercode: "required"
            },
            messages: {
                custemail: {
                    required: "<%=dictLanguage.Item(Session("language")&"_opc_js_2")%>",
                    email: "<%=dictLanguage.Item(Session("language")&"_opc_js_3")%>"
                },
                ordercode: {
                    required: "<%=dictLanguage.Item(Session("language")&"_opc_checkout_4")%>"
                }
            }
        });
        $pc('#ORVSubmit').click(function(){
            if ($pc('#ORVForm').validate().form())
            {
                $pc("#ORVLoader").html('<img src="<%=pcf_getImagePath("images","ajax-loader1.gif")%>" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_checkout_5")%>');
                $pc("#ORVLoader").show(); 
                $pc.ajax({
                    type: "POST",
                    url: "opc_checkORV.asp",
                    data: $pc('#ORVForm').formSerialize(),
                    timeout: 5000,
                    success: function(data, textStatus){
                        if (data.indexOf("OK")>=0)
                        {
                            var tmpArr=data.split("|*|")
                            $pc("#ORVLoader").html('<div class=pcSuccessMessage><%=dictLanguage.Item(Session("language")&"_opc_checkout_6")%></div>');
                            var callbackBill=function (){setTimeout(function(){$pc("#ORVLoader").hide();},1000);}
                            location=tmpArr[1];
                        }
                        else
                        {
                            $pc("#ORVLoader").html('<div class=pcErrorMessage> '+data+' </div>');
                            var callbackBill=function (){setTimeout(function(){$pc("#ORVLoader").hide();},1000);}
                        }
                    }
                });
                return(false);
            }
            return(false);
        });
    });
    </script>
    </div>  
</div>

<% 
'// Managed Form Sessions Auto-Cleared
session("ErrLoginEmail")=""
session("pcSFLoginEmail")=""

'// Clear Un-Managed Sessions
session("pcSFLoginPassword")=""
session("pcSFPassWordExists")=""
session("pcSFEryPassword")=""
%></div>		
		</div>
	</section>
	<!-- /Section: custom-order -->

<!--#include file="footer_wrapper.asp"-->
