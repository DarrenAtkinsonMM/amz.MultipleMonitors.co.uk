<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<%
if session("pc_sdsIsDropShipper")="1" then
	pcv_pageType="0"
	pcv_Table="pcSupplier"
else
	pcv_pageType="1"
	pcv_Table="pcDropShipper"
end if

pcStrPageName="pcnidsdsA2.asp"

'*****************************************************************	
' START: Declare Page Requirements
'*****************************************************************
pcv_sdsCompanyRequired = true
pcv_sdsFirstNameRequired = true
pcv_sdsLastNameRequired = true
pcv_sdsPhoneRequired = true
pcv_sdsEmailRequired = true
pcv_sdsURLRequired = false
pcv_sdsIsDropShipperRequired = false
if pcv_PageType="1" then
	pcv_sdsFromAddressRequired = true
	pcv_sdsFromAddress2Required = false
	pcv_sdsFromCityRequired = true
	pcv_sdsFromZipRequired = true
	pcv_sdsFromCountrycodeRequired = true		
	pcv_sdsFromState1Required = true
	pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
	if  len(pcv_strStateCodeRequired)>0 then
		pcv_sdsFromState1Required=pcv_strStateCodeRequired
	end if		
	pcv_sdsFromState2Required = false
	pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
	if  len(pcv_strProvinceCodeRequired)>0 then
		pcv_sdsFromState2Required=pcv_strProvinceCodeRequired
	end if			
	pcv_sdsUsernameRequired = true
	pcv_sdsPasswordRequired = true	
else
	pcv_sdsFromAddressRequired =  false
	pcv_sdsFromAddress2Required =  false
	pcv_sdsFromCityRequired = false
	pcv_sdsFromZipRequired = false
	pcv_sdsFromCountrycodeRequired = false	
	pcv_sdsFromState1Required = false
	pcv_sdsFromState2Required = false	
	pcv_sdsUsernameRequired = false
	pcv_sdsPasswordRequired = false
end if
'*****************************************************************	
' END: Declare Page Requirements
'*****************************************************************
%>
<%
'*****************************************************************	
' START: Load the Form
'*****************************************************************	
query="SELECT " & pcv_Table & "_ID," & pcv_Table & "_Username," & pcv_Table & "_Password," & pcv_Table & "_FirstName," & pcv_Table & "_LastName," & pcv_Table & "_Company," & pcv_Table & "_Phone," & pcv_Table & "_Email," & pcv_Table & "_URL," & pcv_Table & "_FromAddress," & pcv_Table & "_FromAddress2," & pcv_Table & "_FromCity," & pcv_Table & "_FromStateProvinceCode," & pcv_Table & "_FromZip," & pcv_Table & "_FromCountryCode," & pcv_Table & "_BillingAddress," & pcv_Table & "_BillingAddress2," & pcv_Table & "_BillingCity," & pcv_Table & "_BillingStateProvinceCode," & pcv_Table & "_BillingZip," & pcv_Table & "_BillingCountryCode," & pcv_Table & "_NoticeEmail," & pcv_Table & "_NoticeType," & pcv_Table & "_NoticeMsg," & pcv_Table & "_NotifyManually," & pcv_Table & "_CustNotifyUpdates"
	if pcv_pageType="0" then
		query=query & ","  & pcv_Table & "_IsDropShipper"
	end if	
	query=query & " FROM " & pcv_Table & "s WHERE " & pcv_Table & "_ID=" & session("pc_idsds")
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	

	Session("pcSFpcv_IDsds")=pcf_ResetFormField(Session("pcSFpcv_IDsds"),rs(pcv_Table & "_ID"))
	Session("pcSFpcv_sdsUsername")=pcf_ResetFormField(Session("pcSFpcv_sdsUsername"),rs(pcv_Table & "_Username"))
	Session("pcSFpcv_sdsPassword")=pcf_ResetFormField(Session("pcSFpcv_sdsPassword"),rs(pcv_Table & "_Password"))
	if Session("pcSFpcv_sdsPassword")<>"" then
		Session("pcSFpcv_sdsPassword")=enDeCrypt(Session("pcSFpcv_sdsPassword"), scCrypPass)
	end if
	
	
	Session("pcSFpcv_sdsFirstName")=pcf_ResetFormField(Session("pcSFpcv_sdsFirstName"),rs(pcv_Table & "_FirstName"))
	Session("pcSFpcv_sdsLastName")=pcf_ResetFormField(Session("pcSFpcv_sdsLastName"),rs(pcv_Table & "_LastName"))
	Session("pcSFpcv_sdsCompany")=pcf_ResetFormField(Session("pcSFpcv_sdsCompany"),rs(pcv_Table & "_Company"))
	Session("pcSFpcv_sdsPhone")=pcf_ResetFormField(Session("pcSFpcv_sdsPhone"),rs(pcv_Table & "_Phone"))
	Session("pcSFpcv_sdsEmail")=pcf_ResetFormField(Session("pcSFpcv_sdsEmail"),rs(pcv_Table & "_Email"))
	
	Session("pcSFpcv_sdsURL")=pcf_ResetFormField(Session("pcSFpcv_sdsURL"),rs(pcv_Table & "_URL"))
	Session("pcSFpcv_sdsFromAddress")=pcf_ResetFormField(Session("pcSFpcv_sdsFromAddress"),rs(pcv_Table & "_FromAddress"))
	Session("pcSFpcv_sdsFromAddress2")=pcf_ResetFormField(Session("pcSFpcv_sdsFromAddress2"),rs(pcv_Table & "_FromAddress2"))
	Session("pcSFpcv_sdsFromCity")=pcf_ResetFormField(Session("pcSFpcv_sdsFromCity"),rs(pcv_Table & "_FromCity"))
	Session("pcSFpcv_sdsFromStateProvinceCode")=pcf_ResetFormField(Session("pcSFpcv_sdsFromStateProvinceCode"),rs(pcv_Table & "_FromStateProvinceCode"))
	Session("pcSFpcv_sdsFromZip")=pcf_ResetFormField(Session("pcSFpcv_sdsFromZip"),rs(pcv_Table & "_FromZip"))
	Session("pcSFpcv_sdsFromCountrycode")=pcf_ResetFormField(Session("pcSFpcv_sdsFromCountrycode"),rs(pcv_Table & "_FromCountrycode"))
	Session("pcSFpcv_sdsBillingAddress")=pcf_ResetFormField(Session("pcSFpcv_sdsBillingAddress"),rs(pcv_Table & "_BillingAddress"))
	Session("pcSFpcv_sdsBillingAddress2")=pcf_ResetFormField(Session("pcSFpcv_sdsBillingAddress2"),rs(pcv_Table & "_BillingAddress2"))
	Session("pcSFpcv_sdsBillingCity")=pcf_ResetFormField(Session("pcSFpcv_sdsBillingCity"),rs(pcv_Table & "_BillingCity"))
	Session("pcSFpcv_sdsBillingStateProvinceCode")=pcf_ResetFormField(Session("pcSFpcv_sdsBillingStateProvinceCode"),rs(pcv_Table & "_BillingStateProvinceCode"))
	Session("pcSFpcv_sdsBillingZip")=pcf_ResetFormField(Session("pcSFpcv_sdsBillingZip"),rs(pcv_Table & "_BillingZip"))
	Session("pcSFpcv_sdsBillingCountrycode")=pcf_ResetFormField(Session("pcSFpcv_sdsBillingCountrycode"),rs(pcv_Table & "_BillingCountrycode"))
	Session("pcSFpcv_sdsNoticeEmail")=pcf_ResetFormField(Session("pcSFpcv_sdsNoticeEmail"),rs(pcv_Table & "_NoticeEmail"))
	Session("pcSFpcv_sdsNoticeType")=pcf_ResetFormField(Session("pcSFpcv_sdsNoticeType"),rs(pcv_Table & "_NoticeType"))
	if (Session("pcSFpcv_sdsNoticeType")="") or (not Isnumeric(Session("pcSFpcv_sdsNoticeType"))) then
		Session("pcSFpcv_sdsNoticeType")=0
	end if
	
	Session("pcSFpcv_sdsNoticeMsg")=pcf_ResetFormField(Session("pcSFpcv_sdsNoticeMsg"),rs(pcv_Table & "_NoticeMsg"))
	Session("pcSFpcv_sdsNotifyManually")=pcf_ResetFormField(Session("pcSFpcv_sdsNotifyManually"),rs(pcv_Table & "_NotifyManually"))
	if (Session("pcSFpcv_sdsNotifyManually")="") or (not Isnumeric(Session("pcSFpcv_sdsNotifyManually"))) then
		Session("pcSFpcv_sdsNotifyManually")=0
	end if
	
	Session("pcSFpcv_sdsCustNotifyUpdates")=pcf_ResetFormField(Session("pcSFpcv_sdsCustNotifyUpdates"),rs(pcv_Table & "_CustNotifyUpdates"))
	if (Session("pcSFpcv_sdsCustNotifyUpdates")="") or (not Isnumeric(Session("pcSFpcv_sdsCustNotifyUpdates"))) then
		Session("pcSFpcv_sdsCustNotifyUpdates")=0
	end if
	
	Session("pcSFpcv_sdsIsDropShipper")=0
	if pcv_pageType="0" then
		Session("pcSFpcv_sdsIsDropShipper")=pcf_ResetFormField(Session("pcSFpcv_sdsIsDropShipper"),rs(pcv_Table & "_IsDropShipper"))
		if (Session("pcSFpcv_sdsIsDropShipper")="") or (not Isnumeric(Session("pcSFpcv_sdsIsDropShipper"))) then
			Session("pcSFpcv_sdsIsDropShipper")=0
		end if
	end if

	set rs=nothing	
'*****************************************************************	
' END: Load the Form
'*****************************************************************

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script type=text/javascript>"&vbcrlf
	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf
pcs_JavaTextField	"pcv_sdsCompany", pcv_iscompanyRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"pcv_sdsFirstName",pcv_sdsFirstNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"pcv_sdsLastName",pcv_sdsLastNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"pcv_sdsEmail", pcv_sdsEmailRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
pcs_JavaTextField	"pcv_sdsPhone", pcv_sdsPhoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
if pcv_PageType="0" then
	response.write "if (theForm.pcv_sdsIsDropShipper.checked==true)"&vbcrlf
	response.write "{"&vbcrlf
end if
	pcs_JavaTextField	"pcv_sdsFromAddress", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcv_sdsFromCity", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcv_sdsFromZip", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcv_sdsFromCountrycode", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcv_sdsUsername", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
	pcs_JavaTextField	"pcv_sdsPassword", true, dictLanguage.Item(Session("language")&"_NewCust_3"), ""
if pcv_PageType="0" then
	response.write "}"&vbcrlf
end if
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf

response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>	

<div id="pcMain">
	<div class="pcMainContent">
		<h1><%response.write dictLanguage.Item(Session("language")&"_ModsdsA_1")%></h1>
		<div class="pcFormItem">
			<div class="pcFormItemFull">
				<%response.write dictLanguage.Item(Session("language")&"_ModsdsA_1b")%>
			</div>
		</div>
		<%
		msg = ""
		code = getUserInput(request.QueryString("msg"),0)
		Select Case code
			Case "1" : msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
		End Select
		If msg<>"" then
		%>
			<div class="pcErrorMessage"><%=msg%></div>
		<%
		end if
		%> 
		<form method="post" name="form1" action="pcmodsdsB.asp?action=upd" onSubmit="return Form1_Validator(this)" class="pcForms">
			<div class="pcShowContent">
				<h4><%= dictLanguage.Item(Session("language")&"_ModsdsA_3")%></h4>

				<% '// Company %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_4")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsCompany" size="50" value="<% =pcf_FillFormField ("pcv_sdsCompany", pcv_sdsCompanyRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>
				
				<% '// First Name %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_5")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsFirstName" size="50" value="<% =pcf_FillFormField ("pcv_sdsFirstName", pcv_sdsFirstNameRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>

				<% '// Last Name %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_6")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsLastName" size="50" value="<% =pcf_FillFormField ("pcv_sdsLastName", pcv_sdsLastNameRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>
		
				<%	'// Phone Custom Error
				if session("Errpcv_sdsPhone")<>"" then %>
					<div class="pcFormItem">
						<div class="pcFormLabel">&nbsp;</div>
						<div class="pcFormField">
							<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
						</div>
					</div>
					<% 
					session("Errpcv_sdsPhone") = ""
				end if 
				%>

				<% '// Phone %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_7")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsPhone" size="50" value="<% =pcf_FillFormField ("pcv_sdsPhone", pcv_sdsPhoneRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>
		
				<%	'// Email Custom Error
				if session("Errpcv_sdsEmail")<>"" then %>
					<div class="pcFormItem">
						<div class="pcFormLabel">&nbsp;</div>
						<div class="pcFormField">
							<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
						</div>
					</div>
					<% 
					session("Errpcv_sdsEmail") = ""
				end if 
				%>
		
				<% '// Email %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_8")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsEmail" size="50" value="<% =pcf_FillFormField ("pcv_sdsEmail", pcv_sdsEmailRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>
				
				<% '// Website URL %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_9")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsURL" size="50" value="<% =pcf_FillFormField ("pcv_sdsURL", pcv_sdsURLRequired) %>">
					</div>
				</div>

				<h4><%= dictLanguage.Item(Session("language")&"_ModsdsA_10")%></h4>

				<%
				'///////////////////////////////////////////////////////////
				'// START: COUNTRY AND STATE/ PROVINCE CONFIG
				'///////////////////////////////////////////////////////////
				' 
				' 1) Place this section ABOVE the Country field
				' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
				' 3) Additional Required Info
		
				'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
				pcv_isStateCodeRequired = pcv_sdsFromState1Required '// determines if validation is performed (true or false)
				pcv_isProvinceCodeRequired = pcv_sdsFromState2Required '// determines if validation is performed (true or false)
				pcv_isCountryCodeRequired = pcv_sdsFromCountrycodeRequired '// determines if validation is performed (true or false)
		
				'// #3 Additional Required Info
				pcv_strTargetForm = "form1" '// Name of Form
				pcv_strCountryBox = "pcv_sdsFromCountrycode" '// Name of Country Dropdown
				pcv_strTargetBox = "pcv_sdsFromState1" '// Name of State Dropdown
				pcv_strProvinceBox =  "pcv_sdsFromState2" '// Name of Province Field
		
				'// Set local Country to Session
				if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
				end if
		
				'// Set local State to Session
				if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&"pcv_sdsFromStateProvinceCode")
				end if
		
				'// Set local Province to Session
				if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&"pcv_sdsFromStateProvinceCode")
				end if
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
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_11")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsFromAddress" size="50" value="<% =pcf_FillFormField ("pcv_sdsFromAddress", pcv_sdsFromAddressRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>

				<% '// Address 2 %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_12")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsFromAddress2" size="50" value="<% =pcf_FillFormField ("pcv_sdsFromAddress2", pcv_sdsFromAddress2Required) %>">
					</div>
				</div>

				<% '// City %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_13")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsFromCity" size="50" value="<% =pcf_FillFormField ("pcv_sdsFromCity", pcv_sdsFromCityRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>
				
				<%
				'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince5.asp)
				pcs_StateProvince
				%>

				<% '// Postal Code %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_16")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsFromZip" size="10" value="<% =pcf_FillFormField ("pcv_sdsFromZip", pcv_sdsFromZipRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>

				<h4><%response.write dictLanguage.Item(Session("language")&"_ModsdsA_18")%></h4>
				
				<% '// Username %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_19")%></div>
					<div class="pcFormField">
						<%if (pcv_clone<>"1") and (trim(Session("pcSFpcv_sdsUsername"))<>"") then%>
							<b><%=Session("pcSFpcv_sdsUsername")%></b>
							<input type="hidden" name="pcv_sdsUsername" value="<% =pcf_FillFormField ("pcv_sdsUsername", pcv_sdsUsernameRequired) %>">
						<%else%>
							<input type="text" name="pcv_sdsUsername" size="50" value="<% =pcf_FillFormField ("pcv_sdsUsername", pcv_sdsUsernameRequired) %>">
							&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
						<%end if%>
					</div>
				</div>

				<% '// Password %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_20")%></div>
					<div class="pcFormField">
						<input type="password" name="pcv_sdsPassword" size="50" value="<% =pcf_FillFormField ("pcv_sdsPassword", pcv_sdsPasswordRequired) %>">
						&nbsp;<img src="<%=pcf_getImagePath("images","pc_required.gif")%>" alt="required field">
					</div>
				</div>

				<% '// Order Notification Email %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_21")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsNoticeEmail" size="50" value="<% =pcf_FillFormField ("pcv_sdsNoticeEmail", pcv_sdsNoticeEmailRequired) %>">
					</div>
				</div>

				<h4><%response.write dictLanguage.Item(Session("language")&"_ModsdsA_22")%></h4>

				<%
				'///////////////////////////////////////////////////////////
				'// START: COUNTRY AND STATE/ PROVINCE CONFIG
				'///////////////////////////////////////////////////////////
				' 
				' 1) Place this section ABOVE the Country field
				' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
				' 3) Additional Required Info
		
				'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
				pcv_isStateCodeRequired = false '// determines if validation is performed (true or false)
				pcv_isProvinceCodeRequired = false '// determines if validation is performed (true or false)
				pcv_isCountryCodeRequired = false '// determines if validation is performed (true or false)
		
				'// #3 Additional Required Info
				pcv_strTargetForm = "form1" '// Name of Form
				pcv_strCountryBox = "pcv_sdsBillingCountrycode" '// Name of Country Dropdown
				pcv_strTargetBox = "pcv_sdsBillingState1" '// Name of State Dropdown
				pcv_strProvinceBox =  "pcv_sdsBillingState2" '// Name of Province Field
		
				'// Set local Country to Session
				if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
				end if
		
				'// Set local State to Session
				if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&"pcv_sdsBillingStateProvinceCode")
				end if
		
				'// Set local Province to Session
				if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&"pcv_sdsBillingStateProvinceCode")
				end if

				'// Declare the instance number if greater than 1
				pcv_strFormInstance = "2"  

				'///////////////////////////////////////////////////////////
				'// END: COUNTRY AND STATE/ PROVINCE CONFIG
				'///////////////////////////////////////////////////////////
				%>
		
				<%
				'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince5.asp)
				pcs_CountryDropdown
				%>	

				<% '// Billing Address %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_11")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsBillingAddress" size="50" value="<% =pcf_FillFormField ("pcv_sdsBillingAddress", false) %>">
					</div>
				</div>

				<% '// Billing Address 2 %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_12")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsBillingAddress2" size="50" value="<% =pcf_FillFormField ("pcv_sdsBillingAddress2", false) %>">
					</div>
				</div>

				<% '// Billing City %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_13")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsBillingCity" size="50" value="<% =pcf_FillFormField ("pcv_sdsBillingCity", false) %>">
					</div>
				</div>

				<%
				'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince5.asp)
				pcs_StateProvince
				%>

				<% '// Billing Postal Code %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_ModsdsA_16")%></div>
					<div class="pcFormField">
						<input type=text name="pcv_sdsBillingZip" size="10" value="<% =pcf_FillFormField ("pcv_sdsBillingZip", false) %>">
					</div>
				</div>

				<div class="pcFormItem">
					<div class="pcFormItemFull"><hr /></div>
				</div>

				<div class="pcFormButtons">
					<button class="pcButton pcButtonSubmit" name="Submit" id="submit">
						<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
					</button>

					<a class="pcButton pcButtonBack" href="javascript:history.go(-1)">
						<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="Cancel" />
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
					</a>
				</div>
			</div>
		</form>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
