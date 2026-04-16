<!DOCTYPE html>
<html lang="en">
<head>
<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="include-metatags.asp"-->
<%
if pcv_PageName<>"" then%>
<title><%=pcv_PageName%></title>
<%end if%>
<%GenerateMetaTags()%>
<%Response.Buffer=True%> 
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!--#include file="inc_headerv5.asp" -->
</head>
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<% '// Check if store is turned off and return message to customer
'// Page Name
pcStrPageName = "CustAddShipPop.asp"

'// Check if are coming from the address book
'	>>> If we are coming from the address book we will modify the back button to go to the checkout page
pcv_intMode = request("mode")
if pcv_intMode="" then
	pcv_intMode=0
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Look Up Default Address
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT address, city, state, stateCode, shippingaddress, shippingcity, shippingState, shippingStateCode FROM customers WHERE (((idcustomer)="&session("idCustomer")&"));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
pcStrDefaultShipAddress=rs("shippingAddress")
If len(pcStrDefaultShipAddress)<1 then
	pcStrDefaultShipAddress=pcDefaultAddress
	pcStrDefaultShipCity=pcDefaultCity
	pcStrDefaultShipState=pcDefaultState
	pcStrDefaultShipStateCode=pcDefaultStateCode
Else
	pcStrDefaultShipCity=rs("shippingCity")
	pcStrDefaultShipState=rs("shippingState")
	pcStrDefaultShipStateCode=rs("shippingStateCode") 
End if
set rs=nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Look Up Default Address
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


pcv_isShipFirstNameRequired=True
pcv_isShipLastNameRequired=True
pcv_isShipNickNameRequired=False
pcv_isShipCompanyRequired=False
pcv_isShipAddressRequired=True
pcv_isShipCityRequired=True

'// Use the Request object to toggle State (based of Country selection)
pcv_isShipStateCodeRequired=True
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isShipStateCodeRequired=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
pcv_isShipProvinceCodeRequired=False
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isShipProvinceCodeRequired=pcv_strProvinceCodeRequired
end if

pcv_isShipZipRequired=True
pcv_isShipCountryCodeRequired=True
pcv_isShipPhoneRequired=True
pcv_isShipFaxRequired=False
pcv_isShipEmailRequired=False

if request.form("updatemode")="1" then
	'//set error to zero
	pcv_intErr=0
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ValidateTextField "shipFirstName", pcv_isShipFirstNameRequired, 0
	pcs_ValidateTextField "shipLastName", pcv_isShipLastNameRequired, 0
	pcs_ValidateTextField "shipNickName", pcv_isShipNickNameRequired, 0
	pcs_ValidateTextField "ShipCompany", pcv_isShipCompanyRequired, 0
	pcs_ValidateTextField "ShipAddress", pcv_isShipAddressRequired, 0
	pcs_ValidateTextField "ShipAddress2", false, 0
	pcs_ValidateTextField "ShipCity", pcv_isShipCityRequired, 0
	pcs_ValidateTextField "ShipState", pcv_isShipProvinceCodeRequired, 0
	pcs_ValidateTextField "ShipStateCode", pcv_isShipStateCodeRequired, 0
	pcs_ValidateTextField "ShipZip", pcv_isShipZipRequired, 0
	pcs_ValidateTextField "ShipCountryCode", pcv_isShipCountryCodeRequired, 0
	pcs_ValidatePhoneNumber "ShipPhone", pcv_isShipPhoneRequired, 14
	pcs_ValidatePhoneNumber "ShipFax", pcv_isShipFaxRequired, 14
	pcs_ValidateEmailField "ShipEmail", pcv_isShipEmailRequired, 0
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg=1"
	Else
		
		'// Save collected data in database		
		'// Set Local Variables for recipient
		pcStrShipFirstName = Session("pcSFshipFirstName")
		pcStrShipLastName = Session("pcSFshipLastName")	
		pcStrShipNickName = Session("pcSFshipNickName")
		pcStrShipCompany = Session("pcSFShipCompany")
		pcStrShipAddress = Session("pcSFShipAddress")
		pcStrShipAddress2 = Session("pcSFShipAddress2")
		pcStrShipCity = Session("pcSFShipCity")
		pcStrShipState = Session("pcSFShipState")
		pcStrShipStateCode = Session("pcSFShipStateCode")
		pcStrShipZip = Session("pcSFShipZip")
		pcStrShipCountryCode = Session("pcSFShipCountryCode")
		pcStrShipPhone = Session("pcSFShipPhone")
		pcStrShipFax = Session("pcSFShipFax")
		pcStrShipEmail = Session("pcSFShipEmail")
		pcStrShipFullName=pcStrShipFirstName&" "&pcStrShipLastName
		
		if len(pcStrShipNickName)<1 then
			pcStrShipNickName=pcStrShipFullName
		end if
		
		If pcStrShipState<>"" then
			pcStrShipStateCode = ""
		End If
		
		pcStrShipNickNameTaken=0
		query="SELECT recipients.idRecipient FROM recipients WHERE recipient_NickName='"&pcStrShipNickName&"' AND idCustomer="&session("idCustomer")&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if NOT rs.eof then
			'// Nickname in use already
			pcStrShipNickNameTaken=1
		end if
		set rs=nothing
	
		'// start: check if address matches the default, or any existing nickname
		if ((ucase(pcStrShipAddress)=ucase(pcStrDefaultShipAddress) AND ucase(pcStrShipCity)=ucase(pcStrDefaultShipCity) AND ucase(pcStrShipStateCode)=ucase(pcStrDefaultShipStateCode)) OR (pcStrShipNickNameTaken=1)) then
			if pcStrShipNickNameTaken=1 then
				'// Alert that this address is already existing.	
				response.redirect pcStrPageName&"?msg=2"
			else
				'// Alert that this address is already existing as the default.	
				response.redirect pcStrPageName&"?msg=3"
			end if			
		else
			query="INSERT INTO recipients (idCustomer,recipient_FullName,recipient_Address,recipient_City,recipient_StateCode,recipient_State,recipient_Zip,recipient_CountryCode,recipient_Company,recipient_Address2, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Phone, recipient_Fax, recipient_Email) VALUES ("&session("idCustomer")&",N'" & pcStrShipFullName & "',N'" & pcStrShipAddress & "',N'" & pcStrShipCity & "','" & pcStrShipStateCode & "',N'" & pcStrShipState & "','" & pcStrShipZip & "','" & pcStrShipCountryCode & "',N'" & pcStrShipCompany & "',N'" & pcStrShipAddress2 & "',N'" & pcStrShipNickName & "',N'" & pcStrShipFirstName & "',N'" & pcStrShipLastName & "','" & pcStrShipPhone & "','" & pcStrShipFax & "','" & pcStrShipEmail & "');"
		
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				'// clear the sessions
				pcs_ClearAllSessions
				response.redirect "techErr.asp?err="&pcStrCustRefID
			else
				'// Clear the sessions
				pcs_ClearAllSessions
				%>
                <script type=text/javascript>
					window.close();
				</script>                
                <%
			end if
		end if
		'// end: check if address matches the default
				
	End If
end if

%>
<body id="pcPopup">
<div id="pcMain">		
	<div class="pcMainContent">
    <h1><%= dictLanguage.Item(Session("language")&"_CustSAmanage_1")%></h1>
    <div class="pcSectionTitle"><%= dictLanguage.Item(Session("language")&"_CustAddModShip_17")%></div>
    
		<% 
			msg = ""
			code = getUserInput(request.QueryString("msg"), 0)
			Select Case code
				Case "1" : msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
				Case "2" : msg = dictLanguage.Item(Session("language")&"_CustSAmanage_14")
				Case "3" : msg = dictLanguage.Item(Session("language")&"_CustSAmanage_13")
			End Select

			If msg<>"" then 
				%><div class="pcErrorMessage"><%= msg %></div><%
			end if 
		%>
    
			<form action="<%=pcStrPageName%>" method="post" name="shippingform" class="pcForms">
				<input type="hidden" name="updatemode" value="1">
				<%
				'// The mode param below means this customer was just on the address book and is checking out.
				'   If this param is set to "1" we will re-direct to 'checkout.asp'.
				%>
				<input type="hidden" name="mode" value="<%=pcv_intMode%>">
				<div class="pcShowContent">
        
        	<% 'Recipient Nickname %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_16")%>
            </div>
            <div class="pcFormField">
              <input type="text" name="shipNickName" id="shipNickName" size="20" value="<%=pcf_FillFormField ("shipNickName", pcv_isShipNickNameRequired) %>">
              <% pcs_RequiredImageTag "shipNickName", pcv_isShipNickNameRequired %>
            </div>
          </div>
          
        	<% 'First Name %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_12")%>
            </div>
            <div class="pcFormField">
							<input type="text" name="shipFirstName" id="shipFirstName" size="20" value="<%=pcf_FillFormField ("shipFirstName", pcv_isShipFirstNameRequired) %>">
							<% pcs_RequiredImageTag "shipFirstName", pcv_isShipFirstNameRequired %>
            </div>
          </div>
          
        	<% 'Last Name %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_13")%>
            </div>
            <div class="pcFormField">
							<input type="text" name="shipLastName" id="shipLastName" size="20" value="<%=pcf_FillFormField ("shipLastName", pcv_isShipLastNameRequired) %>">
							<% pcs_RequiredImageTag "shipLastName", pcv_isShipLastNameRequired %>
            </div>
          </div>
          
        	<% 'Company %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_9")%>
            </div>
            <div class="pcFormField">
							<input type="text" name="ShipCompany" id="ShipCompany" size="20" value="<% =pcf_FillFormField ("ShipCompany", pcv_isShipCompanyRequired) %>">
							<% pcs_RequiredImageTag "ShipCompany", pcv_isShipCompanyRequired %>
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
					pcv_isStateCodeRequired = pcv_isShipStateCodeRequired '// determines if validation is performed (true or false)
					pcv_isProvinceCodeRequired = pcv_isShipProvinceCodeRequired '// determines if validation is performed (true or false)
					pcv_isCountryCodeRequired = pcv_isShipCountryCodeRequired '// determines if validation is performed (true or false)
					
					'// #3 Additional Required Info
					pcv_strTargetForm = "shippingform" '// Name of Form
					pcv_strCountryBox = "ShipCountryCode" '// Name of Country Dropdown
					pcv_strTargetBox = "ShipStateCode" '// Name of State Dropdown
					pcv_strProvinceBox =  "ShipState" '// Name of Province Field
					
					'// Set local Country to Session
					if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrShipCountryCode
					end if
					
					'// Set local State to Session
					if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrShipStateCode
					end if
					
					'// Set local Province to Session
					if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pcStrShipState
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

					
        	<% 'Address %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_3")%>
            </div>
            <div class="pcFormField">
							<input type="text" name="ShipAddress" id="ShipAddress" size="20" value="<% =pcf_FillFormField ("ShipAddress", pcv_isShipAddressRequired) %>">
						<% pcs_RequiredImageTag "ShipAddress", pcv_isShipAddressRequired %>
            </div>
          </div>
          
        	<% 'Address 2 %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
            	&nbsp;
            </div>
            <div class="pcFormField">
							<input type="text" name="ShipAddress2" id="ShipAddress2" size="20" value="<% =pcf_FillFormField ("ShipAddress2", false) %>">
            </div>
          </div>
          
        	<% 'City %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_4")%>
            </div>
            <div class="pcFormField">
							<input type="text" name="ShipCity" id="ShipCity" size="20" value="<% =pcf_FillFormField ("ShipCity", pcv_isShipCityRequired) %>">
							<% pcs_RequiredImageTag "ShipCity", pcv_isShipCityRequired %>
            </div>
          </div>
          
					<%
					'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince5.asp)
					pcs_StateProvince
					%>

        	<% 'Zip %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_7")%>
            </div>
            <div class="pcFormField">
							<input type="text" name="ShipZip" id="ShipZip" size="20" value="<% =pcf_FillFormField ("ShipZip", pcv_isShipZipRequired) %>">
							<% pcs_RequiredImageTag "ShipZip", pcv_isShipZipRequired %>
							<span class="pcSmallText"><%= dictLanguage.Item(Session("language")&"_checkout_12")%></span>
            </div>
          </div>

					<%	'// Phone Custom Error
					if session("ErrShipPhone")<>"" then %>            
            <div class="pcFormItem">
              <div class="pcFormLabel">
                &nbsp;
              </div>
              <div class="pcFormField">
								<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
              </div>
            </div>
          
						<% session("ErrShipPhone") = ""
					end if %>

        	<% 'Phone Number %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel">
							<%= dictLanguage.Item(Session("language")&"_CustAddModShip_10")%>
            </div>
            <div class="pcFormField">
							<input type="text" name="ShipPhone" id="ShipPhone" size="20" value="<% =pcf_FillFormField ("ShipPhone", pcv_isShipPhoneRequired) %>">
						  <% pcs_RequiredImageTag "ShipPhone", pcv_isShipPhoneRequired %>
            </div>
          </div>
				<%	'// Phone Custom Error
				if session("ErrShipFax")<>"" then %>
          <div class="pcFormItem">
            <div class="pcFormLabel">
              &nbsp;
            </div>
            <div class="pcFormField">
            	<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
            </div>
          </div>
				<% end if %>
        
				<% 'Fax %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%= dictLanguage.Item(Session("language")&"_CustAddModShip_14")%>
          </div>
          <div class="pcFormField">
            <input type="text" name="ShipFax" id="ShipFax" size="20" value="<% =pcf_FillFormField ("ShipFax", pcv_isShipFaxRequired) %>">
            <% pcs_RequiredImageTag "ShipFax", pcv_isShipFaxRequired %>
          </div>
        </div>
        
				<%	'// Email Custom Error
				if session("ErrShipEmail")<>"" then %>
          <div class="pcFormItem">
            <div class="pcFormLabel">
              &nbsp;
            </div>
            <div class="pcFormField">
            	<img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
            </div>
          </div>
				<% end if %>
				<% 'Fax %>
        <div class="pcFormItem">
          <div class="pcFormLabel">
            <%= dictLanguage.Item(Session("language")&"_CustAddModShip_15")%>
          </div>
          <div class="pcFormField">
						<input type="text" name="ShipEmail" id="ShipEmail" size="20" value="<% =pcf_FillFormField ("ShipEmail", pcv_isShipEmailRequired) %>">
						<% pcs_RequiredImageTag "ShipEmail", pcv_isShipEmailRequired %>
          </div>
        </div>
        
        <div class="pcSpacer"></div>
        
				<div class="pcFormButtons">
					<button class="pcButton pcButtonContinue" name="submitship" id="submit" value="<%= dictLanguage.Item(Session("language")&"_CustAddModShip_11")%>" >
           	<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_CustAddModShip_11")%>" />
            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit")%></span>
          </button>
          
					<a class="pcButton pcButtonBack" href="javascript:window.close();">
          	<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back")%></span>
          </a>
				</div>
			</div>
		</form>
	</div>
</div>
<% call closeDb() %>
</body>
</html>
<% 
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing %>