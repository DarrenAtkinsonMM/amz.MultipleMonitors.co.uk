<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->

<!--#include file="../includes/sendmail.asp"-->

<!--#include file="header_wrapper.asp"-->

<!--#include file="../includes/pcAffConstants.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<%
lmode=0 '// Display Form

pcStrPageName = "NewAffa.asp"

'// Set Required Fields
pcv_isnameRequired=true
pcv_iscompanyRequired=false
pcv_isemailRequired=true
pcv_ispasswordRequired=true
pcv_iscountryRequired=true

'// Use the Request object to toggle State (based of Country selection)
pcv_isstateRequired=true
pcv_strStateCodeRequired=getUserInput(request("pcv_isStateCodeRequired"),250)
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isstateRequired=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
pcv_isprovinceRequired=false
pcv_strProvinceCodeRequired=getUserInput(request("pcv_isProvinceCodeRequired"),250)
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isprovinceRequired=pcv_strProvinceCodeRequired
end if

pcv_isaddressRequired=true
pcv_isaddress2Required=false  
pcv_iscityRequired=true
pcv_iszipRequired=true
pcv_isphoneRequired=true
pcv_isfaxRequired=false
pcv_iswebsiteRequired=true
%>
<% 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script type=text/javascript>"&vbcrlf
	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf
pcs_JavaTextField	"name", pcv_isnameRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"company", pcv_iscompanyRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"email", pcv_isemailRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"password", pcv_ispasswordRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"country", pcv_iscountryRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
'// Do not show, invisible controls
'pcs_JavaTextField	"state", pcv_isstateRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
'pcs_JavaTextField	"province", pcv_isprovinceRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"address", pcv_isaddressRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"city", pcv_iscityRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"zip", pcv_iszipRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"phone", pcv_isphoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"fax", pcv_isfaxRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"website", pcv_iswebsiteRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf

response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IF request.Form("Submit")<>"" THEN
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0

	pcs_ValidateTextField	"name", pcv_isnameRequired, 150
	pcs_ValidateTextField	"company", pcv_iscompanyRequired, 150
	pcs_ValidateEmailField	"email", pcv_isemailRequired, 50
	pcs_ValidateTextField	"password", pcv_ispasswordRequired, 100
	pcs_ValidateTextField	"address", pcv_isaddressRequired, 70
	pcs_ValidateTextField	"address2", pcv_isaddress2Required, 150 
	pcs_ValidateTextField	"country", pcv_iscountryRequired, 150
	pcs_ValidateTextField	"state", pcv_isstateRequired, 150
	pcs_ValidateTextField	"province", pcv_isprovinceRequired, 150
	pcs_ValidateTextField	"city", pcv_iscityRequired, 150
	pcs_ValidateTextField	"zip", pcv_iszipRequired, 12
	pcs_ValidatePhoneNumber	"phone", pcv_isphoneRequired, 30
	pcs_ValidatePhoneNumber	"fax", pcv_isfaxRequired, 30
	pcs_ValidateTextField	"website", pcv_iswebsiteRequired, 150

	'// run additional checks and functions on the our sessions
	'if NOT validNum(Session("pcSFzip")) then
	'	Session("pcSFzip")=0
	'end if

	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg=1"
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		Dim SPath
		SPath=Request.ServerVariables("PATH_INFO")
		SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
		If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
			strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
		Else
			strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
		End if
		
		IF scSecurity=1 THEN
			if scAffReg=1 then
				pcv_Test=0
				if Session("store_AffReg")<>"1" then
					Session("store_AffReg")=""
					Session("store_AffRegpostnum")=""
					Session("store_AffRegnum")=""
					pcv_Test=1
				end if
				if pcv_Test=0 then
					if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "NewAffa.asp"))<>1 then
						Session("store_AffReg")=""
						Session("store_AffRegpostnum")=""
						Session("store_AffRegnum")=""
						pcv_Test=1
					end if
				end if
				
				if pcv_Test=0 then
					if scUseImgs=1 then
						if scCaptchaType="1" then
							blnCAPTCHAcodeCorrect=pcf_checkReCaptcha()
						else%>            
							<!-- Include file for CAPTCHA configuration -->
							<!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
                            
							<!-- Include file for CAPTCHA form processing -->
							<!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
						<%end if
						If not blnCAPTCHAcodeCorrect then
							Session("store_AffReg")=""
							pcv_Test=1
							response.redirect "NewAffa.asp?msg=2"
						end if
					end if
				end if
				
				if pcv_Test=1 then
					If scAlarmMsg=1 then
						if session("AttackCount")="" then
							session("AttackCount")=0
						end if
						session("AttackCount")=session("AttackCount")+1
						if session("AttackCount")>=scAttackCount then
						%>
							<!--#include file="../includes/sendAlarmEmail.asp" -->
						<%
						end if	
					End if
					response.write dictLanguage.Item(Session("language")&"_security_2")
					response.end
				end if
			end if
		END IF
		
		Session("store_AffReg")=""
		Session("store_AffRegpostnum")=""
		Session("store_AffRegnum")=""
		
		' form parameters
		Dim pname, pphone, pcommission

		If Session("pcSFcompany")="" then
			Session("pcSFcompany")=Null
		end if
		
		If Session("pcSFaddress2")="" then
			Session("pcSFaddress2")=Null
		end if
		
		If Session("pcSFfax")="" then
			Session("pcSFfax")=Null
		end if
		
		If Session("pcSFwebsite")="" then
			Session("pcSFwebsite")=Null
		end if
		
		If Session("pcSFprovince")<>"" then
			pcv_strStateOrProvince = Session("pcSFprovince")
		Else
			pcv_strStateOrProvince = Session("pcSFstate")
		End If
		
		'// ProductCart 3.5 - Use default commission from new affiliate settings
		if isNumeric(scAffDefaultCom) and trim(scAffDefaultCom)<>"0" then
			pcommission=scAffDefaultCom
			else
			pcommission="0"
		end if

		'// ProductCart 3.5 - Check auto-approve preference from new affiliate settings		
		if scAffAutoApprove="1" then
			pactive="1"
			else
			pactive="0"
		end if
			
		Session("pcSFPassword")=enDeCrypt(Session("pcSFPassword"), scCrypPass)
		
		' insert product in to db
		query="INSERT INTO affiliates (affiliatename, affiliateEmail, affiliateaddress, affiliateaddress2, affiliatecity, affiliatestate, affiliateCountryCode, affiliatezip, affiliatephone, affiliatefax, affiliatecompany, commission,pcAff_Password,pcAff_Active,pcAff_website) VALUES (N'" &Session("pcSFname")& "','" &Session("pcSFemail")& "',N'" &Session("pcSFaddress")& "',N'" &Session("pcSFaddress2")& "',N'" &Session("pcSFcity")& "',N'" &pcv_strStateOrProvince& "',N'" &Session("pcSFcountry")& "','" &Session("pcSFzip")& "','" &Session("pcSFphone")& "','" &Session("pcSFfax")& "',N'" &Session("pcSFcompany")& "','" & pcommission & "','" & Session("pcSFPassword") & "'," & pactive & ",'" &Session("pcSFwebsite")& "')"
		
		set rstemp=server.createObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		query="Select idAffiliate from affiliates where affiliateEmail='" & Session("pcSFemail") & "' order by idAffiliate desc"
		
		set rstemp=conntemp.execute(query)
		
		pidAffiliate=rstemp("idAffiliate")
		
		pcvStrSubject = dictLanguage.Item(Session("language")&"_storeEmail_4")
		MsgBody=""
		MsgBody=MsgBody & "A new affiliate registered on your store. Below are the affiliate's details." & VBCRLF
		if pactive="0" then
			MsgBody=MsgBody & "This new affiliate account is inactive and will remain so until you review it and activate it using the Affiliates section of the Control Panel. You can turn ON automatical approval using the Affiliate Settings section of the Control Panel." & VBCRLF
			else
			MsgBody=MsgBody & "This affiliate account was automatically approved and is already active. You can turn OFF automatical approval using the Affiliate Settings section of the Control Panel." & VBCRLF
		end if			
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=MsgBody & "==========================================" & VBCRLF
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=MsgBody & "Affiliate ID: #" & pidAffiliate & VBCRLF
		MsgBody=MsgBody & "Affiliate Name: " & Session("pcSFName") & VBCRLF
		MsgBody=MsgBody & "Affiliate Company: " & Session("pcSFCompany") & VBCRLF
		MsgBody=MsgBody & "Phone: " & Session("pcSFPhone") & VBCRLF
		MsgBody=MsgBody & "E-mail: " & Session("pcSFEmail") & VBCRLF
		MsgBody=MsgBody & "Address: " & Session("pcSFAddress") & VBCRLF
		if Session("pcSFAddress2")<>"" then
			MsgBody=MsgBody & "         " & Session("pcSFAddress") & VBCRLF
		end if
		MsgBody=MsgBody & "City: " & Session("pcSFCity") & VBCRLF
		if Session("pcSFState")="" then
			Session("pcSFState")="Not Available"
		end if
		MsgBody=MsgBody & "State/Province: " & Session("pcSFState") & VBCRLF
		MsgBody=MsgBody & "Postal Code: " & Session("pcSFZip") & VBCRLF
		MsgBody=MsgBody & "Country Code: " & Session("pcSFCountry") & VBCRLF
		MsgBody=MsgBody & "Web site: " & Session("pcSFwebsite") & VBCRLF		
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=MsgBody & "==========================================" & VBCRLF
		MsgBody=MsgBody & "" & VBCRLF

		call sendmail (scCompanyName, scEmail, scFrmEmail, pcvStrSubject, MsgBody)		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		lmode=1 '// success
		
		'// Clear the sessions
		pcs_ClearAllSessions
		
	End If	
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim pcStrAffDesc
if pactive="0" then
	pcStrAffDesc=dictLanguage.Item(Session("language")&"_NewAffa_1c")
elseif pactive="1" then
	pcStrAffDesc=dictLanguage.Item(Session("language")&"_NewAffa_2c")
end if	
%>
<div id="pcMain">		
	<div class="pcMainContent">
		<h1><%=dictLanguage.Item(Session("language")&"_NewAffa_1")%></h1>
      
			<%
			' Show information for new affiliates
			if lmode=1 then
			%>
			<div class="pcShowContent">
				<div class="pcTextMessage">
					<p><%=dictLanguage.Item(Session("language")&"_NewAffb_1")%></p>
				</div>
				<div class="pcSpacer"></div>
				<div class="pcFormItem">
					<p><%=pcStrAffDesc%></p>
				</div>
				<div class="pcSpacer"></div>
				<div class="pcFormButtons">
					<a class="pcButton pcButtonSubmit" href="AffiliateLogin.asp">
						<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
					</a>
				</div>
			</div>
			<% 
			end if
			%>
		<%
		msg = ""
		code = getUserInput(request.QueryString("msg"),0)
		Select Case code
			Case "1" : msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
			Case "2" : msg = dictLanguage.Item(Session("language")&"_security_3")
		End Select
		
		If scAffProgramActive = "0" Then
			msg = dictLanguage.Item(Session("language")&"_NewAffb_3")
			lmode = 1
		End If
		
		If msg<>"" then
		%>
			<div class="pcErrorMessage"><%=msg%></div>
		<%
		end if
		%> 
			
		<form method="post" name="addaffiliate" action="<%=pcStrPageName%>" onSubmit="return Form1_Validator(this)" class="pcForms">
			
			<% IF lmode=0 THEN %>	
				<div class="pcShowContent">
        	<div class="pcPageDesc"><%=dictLanguage.Item(Session("language")&"_NewAffa_terms")%></div>

					<% '// Affiliate Name %>
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_2")%></div>
						<div class="pcFormField">
							<input type="text" name="name" value="<% =pcf_FillFormField ("name", pcv_isnameRequired) %>" size="30" maxlength="50">
							<% pcs_RequiredImageTag "name", pcv_isnameRequired %>
						</div>
					</div>

					<% '// Company %>
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_3")%></div>
						<div class="pcFormField">
							<input type="text" name="company" value="<% =pcf_FillFormField ("company", pcv_iscompanyRequired) %>" size="30" maxlength="50"> 
							<% pcs_RequiredImageTag "company", pcv_iscompanyRequired %>
						</div>
					</div>

					<%	'// Email Custom Error
					if session("Erremail")<>"" then %>
						<div class="pcFormItem"> 
							<div class="pcFormLabel">&nbsp;</div>
							<div class="pcFormField">
								<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
							</div>
						</div>
						<% 
						session("Erremail") = ""
					end if 
					%>
                   
					<% '// Email %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_4")%></div>
						<div class="pcFormField">
							<input type="text" name="email" value="<% =pcf_FillFormField ("email", pcv_isemailRequired) %>" size="30" maxlength="150">
							<% pcs_RequiredImageTag "email", pcv_isemailRequired %>
						</div>
					</div>

					<% '// Password %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_5")%></div>
						<div class="pcFormField">
							<input type="password" name="password" value="<% =pcf_FillFormField ("password", pcv_ispasswordRequired) %>" size="30" maxlength="150" autocomplete="off">
							<% pcs_RequiredImageTag "password", pcv_ispasswordRequired %>
						</div>
					</div>

					<%
					'///////////////////////////////////////////////////////////
					'// START: COUNTRY AND STATE/ PROVINCE CONFIG
					'///////////////////////////////////////////////////////////
					' 
					' 1) Place this section ABOVE the Country field
					' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
					' 3) Additional Required Info
						
					'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
					pcv_isStateCodeRequired = pcv_isstateRequired '// determines if validation is performed (true or false)
					pcv_isProvinceCodeRequired = pcv_isprovinceRequired '// determines if validation is performed (true or false)
					pcv_isCountryCodeRequired = pcv_iscountryRequired '// determines if validation is performed (true or false)
						
					'// #3 Additional Required Info
					pcv_strTargetForm = "addaffiliate" '// Name of Form
					pcv_strCountryBox = "country" '// Name of Country Dropdown
					pcv_strTargetBox = "state" '// Name of State Dropdown
					pcv_strProvinceBox =  "province" '// Name of Province Field
						
					'// Set local Country to Session
					'if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
					'	Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
					'end if
						
					'// Set local State to Session
					'if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
					'	Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&pcv_strTargetBox)
					'end if
						
					'// Set local Province to Session
					'if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
					'	Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&pcv_strProvinceBox)
					'end if
					%>					
					<!--#include file="../includes/javascripts/pcStateAndProvince5.asp"-->
					<%
					'///////////////////////////////////////////////////////////
					'// END: COUNTRY AND STATE/ PROVINCE CONFIG
					'///////////////////////////////////////////////////////////
					%>
						
					<%
					'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince5.asp)
					pcs_CountryDropdown
					%>	

					<% '// Address %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_6")%></div>
						<div class="pcFormField">
							<input type="text" name="address" value="<% =pcf_FillFormField ("address", pcv_isaddressRequired) %>" size="30" maxlength="150"> 
							<% pcs_RequiredImageTag "address", pcv_isaddressRequired %>
						</div>
					</div> 

					<% '// Address 2 %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_7")%></div>
						<div class="pcFormField">
							<input type="text" name="address2" value="<% =pcf_FillFormField ("address2", pcv_isaddress2Required) %>" size="30" maxlength="150">
							<% pcs_RequiredImageTag "address2", pcv_isaddress2Required %>
						</div>
					</div>

					<% '// City %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_8")%></div>
						<div class="pcFormField">
							<input type="text" name="city" value="<% =pcf_FillFormField ("city", pcv_iscityRequired) %>" size="20" maxlength="50">
							<% pcs_RequiredImageTag "city", pcv_iscityRequired %>
						</div>
					</div>        
						          
					<%
					'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince5.asp)
					pcs_StateProvince
					%>
						
					<% '// Postal Code %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_11")%></div>
						<div class="pcFormField">
							<input type="text" name="zip" value="<% =pcf_FillFormField ("zip", pcv_iszipRequired) %>" size="10" maxlength="50">
							<% pcs_RequiredImageTag "zip", pcv_iszipRequired %>
						</div>
					</div>

					<%	'// Phone Custom Error
					if session("Errphone")<>"" then %>
						<div class="pcFormItem">
							<div class="pcFormLabel">&nbsp;</div>
							<div class="pcFormField">
								<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
							</div>
						</div>
						<% 
						session("Errphone") = ""
					end if 
					%>
					       
					<% '// Phone %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_12")%></div>
						<div class="pcFormField">
							<input type="text" name="phone" value="<% =pcf_FillFormField ("phone", pcv_isphoneRequired) %>" size="20" maxlength="20"> 
							<% pcs_RequiredImageTag "phone", pcv_isphoneRequired %>
						</div>
					</div>

					<%	'// Fax Custom Error
					if session("Errfax")<>"" then %>
						<div class="pcFormItem">
							<div class="pcFormLabel">&nbsp;</div>
							<div class="pcFormField">
								<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
							</div>
						</div>
						<% 
						session("Errfax") = ""
					end if 
					%>
                   
					<% '// Fax %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_13")%></div>
						<div class="pcFormField">
							<input type="text" name="fax" value="<% =pcf_FillFormField ("fax", pcv_isfaxRequired) %>" size="20" maxlength="20">
							<% pcs_RequiredImageTag "fax", pcv_isfaxRequired %>
						</div>
					</div> 
					 
					<% '// Website URL %> 
					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_NewAffa_15")%></div>
						<div class="pcFormField">
							<input type="text" name="website" value="<%=pcf_FillFormField ("website", pcv_iswebsiteRequired) %>" size="30" maxlength="150">
							<% pcs_RequiredImageTag "website", pcv_iswebsiteRequired %>
						</div>
					</div>

					<%
					Session("store_AffReg")="1"
					Session("store_AffRegpostnum")=""
					session("store_AffRegnum")="      "
					%>

					<%if (scSecurity=1) and (scAffReg=1) and (scUseImgs=1) then%>
						<div class="pcFormItem">
							<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_security_1")%></div>
							<div class="pcFormField">
								<%if scCaptchaType="1" then
									call pcs_genReCaptcha()
								else%>
									<!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
								<%end if%>
							</div>
						</div>
					<%end if%>

					<div class="pcFormButtons">
						<button class="pcButton pcButtonSubmit" name="Submit" id="submit" value="Submit">
							<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
						</button>
					</div>
				</div>

				<div class="pcSpacer"></div>

			<% END IF %>
	
		</form>
	</div>	
</div>
<!--#include file="footer_wrapper.asp"-->
