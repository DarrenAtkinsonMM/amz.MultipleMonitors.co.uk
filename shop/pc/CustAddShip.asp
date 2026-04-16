<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<!--#include file="../includes/pcFormHelpers.asp"-->
<% '// Check if store is turned off and return message to customer
'// Page Name
pcStrPageName = "CustAddShip.asp"

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
		response.redirect pcStrPageName & "?msg=1"
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
				if pcv_intMode = 1 then
					'// checkout can detect the mode param to pass a msg to login.asp					
					'// Express Checkout			
					if session("ExpressCheckoutPayment")<>"YES" then
						response.redirect "checkout.asp"
					else
						response.redirect "login.asp"
					end if
				else
					response.redirect "CustSAmanage.asp?msg=1"					
				end if
			end if
		end if
		'// end: check if address matches the default
				
	End If
end if

%>
<div id="pcMain" class="container">		
	<div class="row">
		<h1><%= dictLanguage.Item(Session("language")&"_CustSAmanage_1")%></h1>

        <!-- <div class="pcSectionTitle"><%= dictLanguage.Item(Session("language")&"_CustAddModShip_17")%></div> -->
        
		<% 
			msg = ""
			code = getUserInput(request.QueryString("msg"),0)
			Select Case code
				Case "1" : msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
				Case "2" : msg = dictLanguage.Item(Session("language")&"_CustSAmanage_14")
				Case "3" : msg = dictLanguage.Item(Session("language")&"_CustSAmanage_13")
			End Select

			If msg<>"" then 
				%><div class="pcErrorMessage"><%= msg %></div><%
			end if 
		%> 

    <form action="CustAddShip.asp" method="post" name="shippingform" class="form">
      <input type="hidden" name="updatemode" value="1">
      <%
      '// The mode param below means this customer was just on the address book and is checking out.
      '   If this param is set to "1" we will re-direct to 'checkout.asp'.
      %>
      <input type="hidden" name="mode" value="<%=pcv_intMode%>">

				<%
          '// Shipping Nick Name
          addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_16"), "text", "shipNickName", "shipNickName", pcStrShipNickName, 20, pcv_isShipNickNameRequired, NULL
					
					'// Shipping First Name
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_12"), "text", "shipFirstName", "shipFirstName", pcStrShipFirstName, 20, pcv_isShipFirstNameRequired, NULL
					
					'// Shipping Last Name
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_13"), "text", "shipLastName", "shipLastName", pcStrShipLastName, 20, pcv_isShipLastNameRequired, NULL
					
					'// Shipping Company
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_9"), "text", "ShipCompany", "ShipCompany", pcStrShipCompany, 20, pcv_isShipCompanyRequired, NULL
        %>
  
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
        <!--#include file="../includes/javascripts/opc_pcStateAndProvince.asp"-->
        <%
        '///////////////////////////////////////////////////////////
        '// END: COUNTRY AND STATE/ PROVINCE CONFIG
        '///////////////////////////////////////////////////////////
        %>
  
        <%
        '// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince5.asp)
        pcs_CountryDropdown
        %>	
        <%
					'// Shipping Address
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_3"), "text", "ShipAddress", "ShipAddress", pcStrShipAddress, 20, pcv_isShipAddressRequired, NULL
					
					'// Shipping Address 2
					addFormInput dictLanguage.Item(Session("language")&"_opc_14"), "text", "ShipAddress2", "ShipAddress2", pcStrShipAddress2, 20, NULL, NULL
					
					'// Shipping City
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_4"), "text", "ShipCity", "ShipCity", pcStrShipCity, 20, pcv_isShipCityRequired, NULL
					
					'// Shipping State/Province
					pcs_StateProvince
					
					'// Shipping Zip
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_7"), "text", "ShipZip", "ShipZip", pcStrShipZip, 20, pcv_isShipZipRequired, dictLanguage.Item(Session("language")&"_checkout_12")
					
					'// Phone Custom Error
					if session("ErrShipPhone")<>"" then %>
						<div class="pcFormItem"><img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%></div>
                        <%
						session("ErrShipPhone") = ""
					end if 
						
					'// Recipient Phone
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_10"), "text", "ShipPhone", "ShipPhone", pcStrShipPhone, 20, pcv_isShipPhoneRequired, NULL
					
					if reID <> "0" then
						'// Fax Custom Error
						if session("ErrShipFax")<>"" then %>
							<div class="pcFormItem">
							    <img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
                            </div>
                            <%
						end if 
						
						'// Recipient Fax
						addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_14"), "text", "ShipFax", "ShipFax", pcStrShipFax, 20, pcv_isShipFaxRequired, NULL
					end if
					
					'// Email Custom Error
					if session("ErrShipEmail")<>"" then %>
						<div class="pcFormItem">
                            <img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
                        </div>
                        <%
					end if 
						
					'// Recipient Email
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_15"), "text", "ShipEmail", "ShipEmail", pcStrShipEmail, 20, pcv_isShipEmailRequired, NULL
        %>

        <div class="pcFormButtons">
			<button class="pcButton pcButtonSubmit" id="submit" name="submitship">
				<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_CustAddModShip_11")%>" />
				<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
			</button>
					&nbsp;
					<%
						if pcv_intMode = 1 then
							'// Express Checkout			
							if session("ExpressCheckoutPayment")<>"YES" then
								pcv_strMode = "checkout.asp"
							else
								pcv_strMode = "login.asp"
							end if
						else
							pcv_strMode = "CustSAmanage.asp"
						end if
					%>
			<a class="pcButton pcButtonBack" href="javascript:location='<%=pcv_strMode%>'">
				<img src="<%=pcf_getImagePath("",rslayout("back")) %>" alt="Cancel" />
				<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
			</a>
        </div>

    </form>
    <div class="pcSpacer"></div>
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->
