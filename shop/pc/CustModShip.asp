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
'// Get recipient ID
reID=getUserInput(request("reID"),0)
if not reID<>"" then
	response.redirect "CustSAmanage.asp"
end if

'// Page Name
pcStrPageName="CustModShip.asp"

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
	IF reID<>"0" then
		pcs_ValidateTextField "shipFirstName", pcv_isShipFirstNameRequired, 0
		pcs_ValidateTextField "shipLastName", pcv_isShipLastNameRequired, 0
		pcs_ValidateTextField "shipNickName", pcv_isShipNickNameRequired, 0
		pcs_ValidatePhoneNumber "ShipFax", pcv_isShipFaxRequired, 14
	End If
	pcs_ValidatePhoneNumber "ShipPhone", pcv_isShipPhoneRequired, 14
	pcs_ValidateEmailField "ShipEmail", pcv_isShipEmailRequired, 0
	pcs_ValidateTextField "ShipCompany", false, 0
	pcs_ValidateTextField "ShipAddress", pcv_isShipAddressRequired, 0
	pcs_ValidateTextField "ShipAddress2", false, 0
	pcs_ValidateTextField "ShipCity", pcv_isShipCityRequired, 0
	pcs_ValidateTextField "ShipState", pcv_isShipProvinceCodeRequired, 0
	pcs_ValidateTextField "ShipStateCode", pcv_isShipStateCodeRequired, 0
	pcs_ValidateTextField "ShipZip", pcv_isShipZipRequired, 0
	pcs_ValidateTextField "ShipCountryCode", pcv_isShipCountryCodeRequired, 0
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Set Local Variables for recipient
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	IF reID<>"0" then
		pcStrShipFirstName = Session("pcSFshipFirstName")
		pcStrShipLastName = Session("pcSFshipLastName")
		pcStrShipNickName = Session("pcSFshipNickName")
	end if
	pcStrShipCompany = Session("pcSFShipCompany")
	pcStrShipAddress = Session("pcSFShipAddress")
	pcStrShipAddress2 = Session("pcSFShipAddress2")
	pcStrShipCity = Session("pcSFShipCity")
	pcStrShipState = Session("pcSFShipState")
	pcStrShipStateCode = Session("pcSFShipStateCode")
	pcStrShipZip = Session("pcSFShipZip")
	pcStrShipCountryCode = Session("pcSFShipCountryCode")
	pcStrShipEmail = Session("pcSFShipEmail")
	pcStrShipPhone = Session("pcSFShipPhone")
	IF reID<>"0" then
		pcStrShipFax = Session("pcSFShipFax")
		pcStrShipFullName=pcStrShipFirstName&" "&pcStrShipLastName
	end if
	
	if len(pcStrShipNickName)<1 then
		pcStrShipNickName=pcStrShipFullName
	end if
	
	If pcStrShipState<>"" then
		pcStrShipStateCode = ""
	End If
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Set Local Variables for recipient
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	If pcv_intErr>0 Then	
		response.redirect pcStrPageName&"?reID="&reID&"&msg=1"
	Else
		
		'//Open database to update data	
		IF reID="0" then
			query="UPDATE customers SET shippingAddress=N'" & pcStrShipAddress & "', shippingCity=N'" & pcStrShipCity & "', shippingState=N'" & pcStrShipState & "', shippingStateCode='" & pcStrShipStateCode & "', shippingZip='" & pcStrShipZip & "', shippingCountryCode='" & pcStrShipCountryCode & "', shippingCompany=N'" & pcStrShipCompany & "', shippingAddress2=N'" & pcStrShipAddress2 & "', shippingPhone='" & pcStrShipPhone & "', shippingEmail='" & pcStrShipEmail & "' WHERE IDCustomer=" & session("idCustomer") &";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		ELSE
		
			'// Check the Nickname
			pcStrShipNickNameTaken=0
			query="SELECT recipients.idRecipient FROM recipients WHERE recipient_NickName='"&pcStrShipNickName&"' AND idCustomer="&session("idCustomer")&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if NOT rs.eof then
				pcv_stridRecipient = rs("idRecipient")
				if (pcv_stridRecipient=cint(reID))=False then
					'// Nickname in use already
					pcStrShipNickNameTaken=1
				end if
			end if
			set rs=nothing
			
			'// If Nickname is in use, redirect with a message.
			if pcStrShipNickNameTaken=1 then
				'// Alert that this address is already existing in the database.	
				response.redirect pcStrPageName&"?reID="&reID&"&msg=2"
			else
				query="update recipients set recipient_FullName=N'" & pcStrShipFullName & "',recipient_Address=N'" & pcStrShipAddress & "',recipient_City=N'" & pcStrShipCity & "',recipient_StateCode='" & pcStrShipStateCode & "',recipient_State=N'" & pcStrShipState & "',recipient_Zip='" & pcStrShipZip & "',recipient_CountryCode='" & pcStrShipCountryCode & "',recipient_Company=N'" & pcStrShipCompany & "',recipient_Address2=N'" & pcStrShipAddress2 & "', recipient_NickName=N'" & pcStrShipNickName & "', recipient_FirstName=N'" & pcStrShipFirstName & "', recipient_LastName=N'" & pcStrShipLastName & "', recipient_Phone='" & pcStrShipPhone & "', recipient_Fax='" & pcStrShipFax & "', recipient_Email='" & pcStrShipEmail & "' where idRecipient=" & reID & " and IDCustomer=" & session("idCustomer")
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
			end if
			
		END if
	
		set rs=nothing
		'//Close database before redirecting 		
		call closedb()
		
		'// Clear all sessions
		pcs_ClearAllSessions()
		
		response.redirect "CustSAmanage.asp?msg=2"
	End If
end if

IF reID="0" then
	
	query="SELECT Address, City, State, Statecode, Zip, CountryCode, customerCompany, phone, email, shippingAddress, shippingCity, shippingState, shippingStateCode, shippingZip, shippingCountryCode, shippingCompany, shippingAddress2, shippingPhone, shippingEmail FROM customers WHERE idCustomer=" &session("idCustomer")& ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	
	if not rs.eof then
		pcStrAddress=rs("Address")
		pcStrCity=rs("City")
		pcStrState=rs("State")
		pcStrStatecode=rs("Statecode")
		pcStrZip=rs("Zip")
		pcStrCountryCode=rs("CountryCode")
		pcStrcustomerCompany=rs("customerCompany")
		pcStrphone=rs("phone")
		pcStremail=rs("email")
		pcStrShipAddress=rs("shippingAddress")
		pcStrShipCity=rs("shippingCity")
		pcStrShipState=rs("shippingState")
		pcStrShipStateCode=rs("shippingStateCode") 
		pcStrShipZip=rs("shippingZip")
		pcStrShipCountryCode=rs("shippingCountryCode")
		pcStrShipCompany=rs("shippingCompany")
		pcStrShipAddress2=rs("shippingAddress2")
		pcStrShipPhone=rs("shippingPhone")
		pcStrShipEmail=rs("shippingEmail")
		if rs("shippingAddress")<>"" then
		else
			pcStrShipAddress=pcStrAddress
			pcStrShipZip=pcStrZip
			pcStrShipState=pcStrState
			pcStrShipStateCode=pcStrStatecode
			pcStrShipCity=pcStrCity
			pcStrShipCountryCode=pcStrCountryCode
			pcStrShipCompany=pcStrcustomerCompany
			pcStrShipPhone=pcStrphone
			pcStrShipEmail=pcStremail
		end if
		set rs=nothing
	else
		set rs=nothing
		call closeDb()
		response.redirect "CustSAmanage.asp"
	end if	
ELSE
	query="SELECT recipient_FullName, recipient_Address, recipient_City, recipient_StateCode, recipient_State, recipient_Zip, recipient_CountryCode, recipient_Company, recipient_Address2, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Phone, recipient_Fax, recipient_Email FROM recipients WHERE (((idRecipient)=" & reID & ") AND ((idCustomer)=" & session("idCustomer") & "));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rs.eof then
		set rs=nothing
		call closeDb()
		response.redirect "CustSAmanage.asp"
	else
		pcStrShipFullName=rs("recipient_FullName")
		pcStrShipAddress=rs("recipient_Address")
		pcStrShipCity=rs("recipient_City")
		pcStrShipStateCode=rs("recipient_StateCode")
		pcStrShipState=rs("recipient_State")
		pcTempShipZip=rs("recipient_Zip")
		pcTempSplitZip=split(pcTempShipZip,"||")
		if ubound(pcTempSplitZip)>-1 then
			pcStrShipZip=pcTempSplitZip(0)
			if ubound(pcTempSplitZip)>0 then
				pcStrShipPhone=pcTempSplitZip(1)
			end if
		end if
		pcStrShipCountryCode=rs("recipient_CountryCode")
		pcStrShipCompany=rs("recipient_Company")
		pcStrShipAddress2=rs("recipient_Address2")
		pcStrShipNickName=rs("recipient_NickName")
		pcStrShipFirstName=rs("recipient_FirstName")
		pcStrShipLastName=rs("recipient_LastName")
		pcStrShipPhone=rs("recipient_Phone")
		pcStrShipFax=rs("recipient_Fax")
		pcStrShipEmail=rs("recipient_Email")
		set rs=nothing
		
		'//If First and Last Names are not present, parse FullName
		If len(pcStrShipFirstName)<1 AND len(pcStrShipLastName)<1 AND len(pcStrShipFullName)>0 then
			pcStrShipFullNameArray=split(pcStrShipFullName, " ")
			pcStrShipFirstName=pcStrShipFullNameArray(0)
			if ubound(pcStrShipFullNameArray)>0 then
				pcStrShipLastName=pcStrShipFullNameArray(1)
			end if
		end if
		If len(pcStrShipNickname)<1 then
			pcStrShipNickName=pcStrShipFirstName&" "&pcStrShipLastName
		End if
		
	end if
END IF



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Re-Set the Variables
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


IF reID<>"0" then	
	pcStrShipFirstName = pcf_ResetFormField(Session("pcSFshipFirstName"), pcStrShipFirstName)	
	pcStrShipLastName = pcf_ResetFormField(Session("pcSFshipLastName"), pcStrShipLastName)	
	pcStrShipNickName = pcf_ResetFormField(Session("pcSFshipNickName"), pcStrShipNickName)	
end if
pcStrShipCompany = pcf_ResetFormField(Session("pcSFShipCompany"), pcStrShipCompany)
pcStrShipAddress = pcf_ResetFormField(Session("pcSFShipAddress"), pcStrShipAddress)
pcStrShipAddress2 = pcf_ResetFormField(Session("pcSFShipAddress2"), pcStrShipAddress2)
pcStrShipCity = pcf_ResetFormField(Session("pcSFShipCity"), pcStrShipCity)
pcStrShipState = pcf_ResetFormField(Session("pcSFShipState"), pcStrShipState)
pcStrShipStateCode = pcf_ResetFormField(Session("pcSFShipStateCode"), pcStrShipStateCode)
pcStrShipZip = pcf_ResetFormField(Session("pcSFShipZip"), pcStrShipZip)
pcStrShipCountryCode = pcf_ResetFormField(Session("pcSFShipCountryCode"), pcStrShipCountryCode)
pcStrShipPhone = pcf_ResetFormField(Session("pcSFShipPhone"), pcStrShipPhone)
pcStrShipEmail = pcf_ResetFormField(Session("pcSFShipEmail"), pcStrShipEmail)
IF reID<>"0" then
	pcStrShipFax = pcf_ResetFormField(Session("pcSFShipFax"), pcStrShipFax)
	pcStrShipFullName=pcStrShipFirstName&" "&pcStrShipLastName
end if
if len(pcStrShipNickName)<1 then
	pcStrShipNickName=pcStrShipFullName
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Re-Set the Variables
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<div id="pcMain" class="container-fluid">		

      <div class="row">  
      
		<h1><%= dictLanguage.Item(Session("language")&"_CustSAmanage_1")%></h1>
    
		<!--<div class="pcSectionTitle"><%= dictLanguage.Item(Session("language")&"_CustAddModShip_1")%></div>-->
    
		<% 
			msg = ""
			code = getUserInput(request.QueryString("msg"),0)
			Select Case code
				Case "1" : msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
				Case "2" : msg = dictLanguage.Item(Session("language")&"_CustSAmanage_14")
			End Select

			If msg<>"" then 
				%><div class="pcErrorMessage"><%= msg %></div><%
			end if 
		%>
        
		<form action="<%=pcStrPageName%>" method="post" name="shippingform" class="form">
			<input type="hidden" name="updatemode" value="1">
			<input type=hidden name="reID" value="<%=ReID%>">  
      

				<% 
					if reID<>"0" then
						'// Shipping Nick Name
						addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_16"), "text", "shipNickName", "shipNickName", pcStrShipNickName, 20, pcv_isShipNickNameRequired, NULL
												 
						'// Shipping First Name
						addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_12"), "text", "shipFirstName", "shipFirstName", pcStrShipFirstName, 20, pcv_isShipFirstNameRequired, NULL
						
						'// Shipping Last Name
						addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_13"), "text", "shipLastName", "shipLastName", pcStrShipLastName, 20, pcv_isShipLastNameRequired, NULL
        	end if
				%>
        
        <%
					'// Shipping Company
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_9"), "text", "ShipCompany", "ShipCompany", pcStrShipCompany, 20, NULL, NULL	
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
					
					'// Shipping Zip
					addFormInput dictLanguage.Item(Session("language")&"_CustAddModShip_7"), "text", "ShipZip", "ShipZip", pcStrShipZip, 20, pcv_isShipZipRequired, dictLanguage.Item(Session("language")&"_checkout_12")
				%>
					
				<%
					'// Shipping State/Province
					pcs_StateProvince
				%>
        
        <%
					'// Phone Custom Error
					if session("ErrShipPhone")<>"" then %>
						<div class="pcFormItem">
                            <img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
                        </div>
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
			<a class="pcButton pcButtonBack" href="javascript:location='CustSAmanage.asp'">
				<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="Cancel" />
				<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
			</a>
      </div>
    </form>
    <div class="pcSpacer"></div>
  
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->
