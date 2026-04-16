<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 'Gateway specific files %>
<%
'SB-S
If session("SB_SkipPayment")="1" then
	Response.redirect "gwReturn.asp?s=true&gw=EIG"
End if
'SB-E

'SB S
Dev_Testmode = 1
msg=getUserInput(request.querystring("message"),0)
msg=replace(msg, "&lt;BR&gt;", "<BR>")
msg=replace(msg, "&lt;br&gt;", "<br>")
msg=replace(msg, "&lt;b&gt;", "<b>")
msg=replace(msg, "&lt;/b&gt;", "</b>")
msg=replace(msg, "&lt;/font&gt;", "</font>")
msg=replace(msg, "&lt;a href", "<a href")
msg=replace(msg, "&gt;Back&lt;/a&gt;", ">Back</a>")
msg=replace(msg, "&lt;font", "<font")
msg=replace(msg, "&gt;<b>Error&nbsp;</b>:", "><b>Error&nbsp;</b>:")
msg=replace(msg, "&gt;&lt;img src=", "><img src=")
msg=replace(msg, "&gt;&lt;/a&gt;", "></a>")
msg=replace(msg, "&gt;<b>", "><b>")
msg=replace(msg, "&lt;/a&gt;", "</a>")
msg=replace(msg, "&gt;View Cart", ">View Cart")
msg=replace(msg, "&gt;Continue", ">Continue")
msg=replace(msg, "&lt;u>", "<u>")
msg=replace(msg, "&lt;/u>", "</u>")
msg=replace(msg, "&lt;ul&gt;", "<ul>")
msg=replace(msg, "&lt;/ul&gt;", "</ul>")
msg=replace(msg, "&lt;li&gt;", "<li>")
msg=replace(msg, "&lt;/li&gt;", "</li>")
msg=replace(msg, "&gt;", ">")
msg=replace(msg, "&lt;", "<")
'SB E
' Determine BACK button url
If scSSL="1" And scIntSSLPage="1" Then
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
Else
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
%>

<div id="PleaseWaitDialog" title="" style="display:none">
	<div id="PleaseWaitMsg" class="ui-main"></div>
</div>
<script type="text/javascript">
	$pc(document).ready(function() {
		//*Please Wait Dialog
		//$pc("#PleaseWaitDialog").dialog({
		//	bgiframe: true,
		//	autoOpen: false,
		//	resizable: false,
		//	width: 250,
		//	minHeight: 50,
		//	modal: true
		//});
	});
</script>
<div id="pcMain">
	<div class="container-fluid">
		<%
		Dim pcCustIpAddress
		pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

		' Get Order ID
		if session("GWOrderId")="" then
			session("GWOrderId")=getUserInput(request("idOrder"),0)
		end if

		pcGatewayDataIdOrder=session("GWOrderID")
		%>
		<!--#include file="pcGateWayData.asp"-->
		<% session("idCustomer")=pcIdCustomer
		pcv_IncreaseCustID=(scCustPre + int(pcIdCustomer))
		pcTrueOrdnum=(int(session("GWOrderId"))-scpre) %>
		<%
		'// LOAD SSETTINGS		
		query="SELECT pcPay_EIG_Type, pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key, pcPay_EIG_Curcode, pcPay_EIG_CVV, pcPay_EIG_TestMode, pcPay_EIG_SaveCards, pcPay_EIG_UseVault FROM pcPay_EIG WHERE pcPay_EIG_ID=1"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		x_Username=rs("pcPay_EIG_Username")
        If len(x_Username)>0 Then
		    x_Username=enDeCrypt(x_Username, scCrypPass)
		End If
        x_Password=rs("pcPay_EIG_Password")
        If len(x_Password)>0 Then
		    x_Password=enDeCrypt(x_Password, scCrypPass)
		End If
        x_Key=rs("pcPay_EIG_Key")
        If len(x_Key)>0 Then
		    x_Key=enDeCrypt(x_Key, scCrypPass)
        End If
		x_CVV=rs("pcPay_EIG_CVV")
		x_Type=rs("pcPay_EIG_Type")
        If len(x_Type)>0 Then
		    x_TypeArray=Split(x_Type,"||")
            x_TransType=x_TypeArray(0)
        End If		
		x_Curcode=rs("pcPay_EIG_Curcode")
		x_TestMode=rs("pcPay_EIG_TestMode")
		x_SaveCards=rs("pcPay_EIG_SaveCards")
		x_UseVault=rs("pcPay_EIG_UseVault")
		set rs=nothing


		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// START:  PROCESS RESULTS WHEN SUBSCRIPTION
		'/////////////////////////////////////////////////////////////////////////////////////////////
		If Request.Form("PaymentGWEIG")="Go" Then %>

			<% session("redirectPage")="gwEIGateway.asp" %>

			<%
			dim tempReturnURL
			If scSSL="" OR scSSL="0" Then
				tempReturnURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
				tempReturnURL=replace(tempReturnURL,"https:/","https://")
				tempReturnURL=replace(tempReturnURL,"http:/","http://")
			Else
				tempReturnURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
				tempReturnURL=replace(tempReturnURL,"https:/","https://")
				tempReturnURL=replace(tempReturnURL,"http:/","http://")
			End If

			'SB S
			'// By pass EIG if the immediate order value is 0
			If pcBillingTotal<0 Then
				pcBillingTotal=0
			End If
			If (pcIsSubscription) AND (pcBillingTotal=0) Then

				session("reqCardNumber")=getUserInput(request.Form("billing-cc-number"),16)
				session("reqExpMonth")=getUserInput(request.Form("billing-cc-exp1"),0)
				session("reqExpYear")=getUserInput(request.Form("billing-cc-exp2"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("billing-cvv"),4)
				pExpiration=getUserInput(request("billing-cc-exp1"),0) & "/01/" & getUserInput(request("billing-cc-exp2"),0)

				'// Validates expiration
				if DateDiff("d", Month(Now)&"/01/"&Year(now),pExpiration)<=-1 then
                    call closeDb()
                    Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_6")
                    Session("backbuttonURL") = tempURL & "?psslurl=gwEIGateway.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"
				end if

				'// Validate card
				if not IsCreditCard(session("billing-cc-number"), request.form("x_Card_Type")) then
                    call closeDb()
                    Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_5")
                    Session("backbuttonURL") = tempURL & "?psslurl=gwEIGateway.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"
				end if

                call closeDb()
				Response.Redirect("gwReturn.asp?s=true&gw=EIG&GWError=1")
				Response.End

			Else

				'// Normal Payment, Let Pass
				session("reqCardNumber")=getUserInput(request.Form("billing-cc-number"),16)
				session("reqExpMonth")=getUserInput(request.Form("billing-cc-exp1"),0)
				session("reqExpYear")=getUserInput(request.Form("billing-cc-exp2"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("billing-cvv"),4)

			End if
			'SB E
			%>


			<%
			if pcBillingTotal > 0 then

				If x_CVV="1" Then
					if not isnumeric(session("reqCVV")) or len(session("reqCVV")) < 3 or len(session("reqCVV")) > 4 Then

                        call closeDb()
                        Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_7")&dictLanguage.Item(Session("language")&"_paymntb_c_4")
                        Session("backbuttonURL") = tempURL & "?psslurl=gwEIGateway.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                        response.redirect "msgb.asp?back=1"
                        
					End If
				End if

				Dim objXMLHTTP, xml

				'// Send the request to the Authorize.NET processor.
				stext=""
				stext=stext & "username=" & x_Username
				stext=stext & "&password=" & x_Password
                
                '// Customer Information for Direct Authorizations Only
                stext=stext & "&firstname=" & pcf_FixXML(pcBillingFirstName)
                stext=stext & "&lastname=" & pcf_FixXML(pcBillingLastName)
                stext=stext & "&company=" & pcf_FixXML(pcBillingCompany)
                stext=stext & "&address1=" & pcf_FixXML(pcBillingAddress)
                stext=stext & "&address2=" & pcf_FixXML(pcBillingAddress2)
                stext=stext & "&city=" & pcf_FixXML(pcBillingCity)
                stext=stext & "&state=" & pcf_FixXML(pcBillingState)
                stext=stext & "&zip=" & pcf_FixXML(pcBillingPostalCode)
                stext=stext & "&country=" & pcf_FixXML(pcBillingCountryCode)
                stext=stext & "&phone=" & pcf_FixXML(pcBillingPhone)
                stext=stext & "&email=" & pcf_FixXML(pcCustomerEmail) 
                stext=stext & "&orderid=" & pcf_FixXML(session("GWOrderId"))

				If x_TransType="AUTH_ONLY" Then
					stext=stext & "&type=auth"
				Else
					stext=stext & "&type=sale"
				End If
				stext=stext & "&amount=" & pcBillingTotal
				stext=stext & "&ccnumber=" & session("reqCardNumber")
				stext=stext & "&ccexp=" & session("reqExpMonth")&session("reqExpYear")
				If x_CVV="1" Then
					stext=stext & "&cvv=" & session("reqCVV")
				End If

				'// Send the transaction info as part of the querystring
				set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
				xml.open "POST", "https://secure.networkmerchants.com/api/transact.php", false
				xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				xml.send stext
				strStatus = xml.Status

				'// Store the response
				strRetVal = xml.responseText
				Set xml = Nothing

				'// Check for success
				Set resArray = deformatNVP(strRetVal)
				ack = resArray("response")
				ackDesc = resArray("responsetext")

				If ack="1" Then

					pcv_SecurityPass = scCrypPass
					pcv_SecurityKeyID = pcs_GetKeyID

					dim pCardNumber, pCardNumber2
					pCardNumber=session("reqCardNumber")
					pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)

					'// Save Batch Processing Record
					If x_TransType="AUTH_ONLY" Then

						session("GWAuthCode") = resArray("authcode")
						session("GWTransId") = resArray("transactionid")

						query="INSERT INTO pcPay_EIG_Authorize (idOrder, amount, vaultToken, paymentmethod, transtype, authcode, ccnum, ccexp, cctype, idCustomer, fname, lname, address, zip, captured, trans_id, pcSecurityKeyID) VALUES ("& pcTrueOrdnum &", "& pcBillingTotal &", '', 'CC', '"& x_TransType &"', '"& session("GWAuthCode") &"', '"& pCardNumber2 &"', '"& session("reqExpMonth")&session("reqExpYear") &"', '"& session("reqCardType") &"', "& session("idCustomer") &", N'"&replace(pcBillingFirstName,"'","''")&"', N'"&replace(pcBillingLastName,"'","''")&"', N'"&replace(pcBillingAddress,"'","''")&"', '"& pcBillingPostalCode &"', 0, '"& session("GWTransId") &"', "& pcv_SecurityKeyID &");"

						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						set rs=nothing

						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if

					End If
					set rs=nothing
					
					'Log successful transaction
					call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 1)


					'// Clear all card sessions before redirect
					'SB S
					if not pcIsSubscription then
						session("reqCardNumber")=""
						session("reqExpMonth")=""
						session("reqExpYear")=""
						session("reqCVV")=""
					End if
					'SB E

					session("x_response_code")=""
					session("x_response_subcode")=""
					session("x_response_reason_code")=""
					session("x_response_reason_text")=""
					session("x_avs_code")=""
					session("x_description")=""
					session("x_amount")=""
					'session("x_method")=""
					'session("x_type")=""
					session("x_cust_id")=""
					session("x_first_name")=""
					session("x_last_name")=""
					session("x_company")=""
					session("x_address")=""
					session("x_city")=""
					session("x_state")=""
					session("x_zip")=""
					session("x_country")=""
					session("x_phone")=""
					session("x_fax")=""
					session("x_email")=""
					session("x_ship_to_first_name")=""
					session("x_ship_to_last_name")=""
					session("x_ship_to_company")=""
					session("x_ship_to_address")=""
					session("x_ship_to_city")=""
					session("x_ship_to_state")=""
					session("x_ship_to_zip")=""
					session("x_ship_to_country")=""

					If pcIsSubscription Then
						Response.redirect "gwReturn.asp?s=true&gw=EIG"
					Else
						Response.redirect "gwReturn.asp?s=false&gw=EIG"
					End If


				Else '// If ack="1" Then
				  	
					'Log failed transaction
					call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)

                    call closeDb()
                    Session("message") = ackDesc
                    Session("backbuttonURL") = tempURL & "?psslurl=gwEIGateway.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"

				End If '// If ack="1" Then

			End if '// if pcBillingTotal > 0 then


		End If
		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// END:  PROCESS RESULTS WHEN SUBSCRIPTION
		'/////////////////////////////////////////////////////////////////////////////////////////////




		'// CHECK FOR TOKEN
		Dim TokenID
		TokenID=getUserInput(request("token-id"),0)
		IF len(TokenID)>0 THEN

		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// START:  PROCESS RESULTS WHEN TOKEN EXISTS
		'/////////////////////////////////////////////////////////////////////////////////////////////

			'// COMPLETE ACTION
			strTest = ""
			strTest = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
			strTest = strTest & "<complete-action>"
			strTest = strTest & "<api-key>" & x_Key & "</api-key>"
			strTest = strTest & "<token-id>" & TokenID & "</token-id>"
			strTest = strTest & "</complete-action>"

			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			xml.open "POST", "https://secure.nmi.com/api/v2/three-step", false
			xml.setRequestHeader "Content-Type", "text/xml"
			xml.send strTest
			strStatus = xml.Status
			strRetVal = xml.responseText
			Set xml = Nothing

			strResult = pcf_GetNode(strRetVal, "result", "*")
			strResultText = pcf_GetNode(strRetVal, "result-text", "*")
			strTransactionID = pcf_GetNode(strRetVal, "transaction-id", "*")
			strResultCode = pcf_GetNode(strRetVal, "result-code", "*")
			strAuthorizationCode = pcf_GetNode(strRetVal, "authorization-code", "*")
			pcv_strCustomerVaultID = pcf_GetNode(strRetVal, "customer-vault-id", "*")

			'response.Write(strResult & ".<br />")
			'response.Write(strResultText & ".<br />")
			'response.Write(strTransactionID & ".<br />")
			'response.Write(strResultCode & ".<br />")
			'response.Write(authorization-code & ".<br />")
			'response.Write(pcv_strCustomerVaultID & ".<br />")
			'response.End

			'// PROCESS RESULTS
			'// 1 = Transaction Approved
			'// 2 = Transaction Declined
			'// 3 = Error in transaction data or system error
			If strResult="1" Then


				If (x_TransType="AUTH_ONLY") OR (x_SaveCards="1" AND Session("SF_IsSaved")="true") Then

					Dim pcv_CardNum, pcv_CardType, pcv_CardExp

					If (len(Session("CustomerVaultID"))=0) Then '// A. New Card was used.

						If (x_UseVault=1) OR (x_SaveCards="1" AND Session("SF_IsSaved")="true") Then '// Vault Storage Enabled - OR - Customer Opt-In (Grab from Secure Vault)

							strTest = ""
							strTest = strTest & "username=" & x_Username
							strTest = strTest & "&password=" & x_Password
							strTest = strTest & "&transaction_id=" & strTransactionID

							set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
							xml.open "POST", "https://secure.networkmerchants.com/api/query.php", false
							xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
							xml.send strTest
							strStatus = xml.Status
							strRetVal = xml.responseText
							Set xml = Nothing

							pcv_CardNum = pcf_GetNode(strRetVal, "cc_number", "//nm_response/transaction")
							pcv_CardType = Session("CardType")
							pcv_CardExp = pcf_GetNode(strRetVal, "cc_exp", "//nm_response/transaction")
							If len(pcv_strCustomerVaultID)=0 Then
								pcv_strCustomerVaultID = pcf_GetNode(strRetVal, "customerid", "//nm_response/transaction")
							End If
							If len(pcv_strCustomerVaultID)>0 Then
								Session("CustomerVaultID")=pcv_strCustomerVaultID
							End If
							pcv_strCustomerVaultID2=enDeCrypt(Session("CustomerVaultID"), scCrypPass)

							'// Save Vault Record
							query="SELECT idOrder FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_CardNum='"& pcv_CardNum &"' AND pcPay_EIG_Vault_CardExp='"& pcv_CardExp &"'"
							set rs=Server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)
							if rs.eof then

								If Session("SF_IsSaved")="true" Then
									pcv_tmpIsSaved = 1
								Else
									pcv_tmpIsSaved = 0
								End If
								query="INSERT INTO pcPay_EIG_Vault (idOrder, idCustomer, IsSaved, pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardType, pcPay_EIG_Vault_CardExp, pcPay_EIG_Vault_Token) VALUES ("&pcTrueOrdnum&", "&session("idCustomer")&", "&pcv_tmpIsSaved&", '"&pcv_CardNum&"', '"& pcv_CardType &"', '"& pcv_CardExp &"', '"& pcv_strCustomerVaultID2 &"');"
								set rs2=server.CreateObject("ADODB.RecordSet")
								set rs2=connTemp.execute(query)
								set rs2=nothing

							end if
							set rs=nothing

						Else '// Vault Storage Disabled (Grab from Session)

							pcv_CardNum = Session("CardNum")
							pcv_CardType = Session("CardType")
							pcv_CardExp = Session("CardExp")
							pcv_strCustomerVaultID2="" '// No vault record

						End If

					Else '// B. Saved Card was used.

						query="SELECT pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardType, pcPay_EIG_Vault_CardExp FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_ID="& Session("VaultID") &""
						set rs=Server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if NOT rs.eof then
							pcv_CardNum = rs("pcPay_EIG_Vault_CardNum")
							pcv_CardType = rs("pcPay_EIG_Vault_CardType")
							pcv_CardExp = rs("pcPay_EIG_Vault_CardExp")
						end if
						set rs=nothing
						pcv_strCustomerVaultID2=enDeCrypt(Session("CustomerVaultID"), scCrypPass)

					End If


					'// Save Batch Processing Record
					If x_TransType="AUTH_ONLY" Then

						pcv_CardNum=enDeCrypt(pcv_CardNum, scCrypPass)

						query="INSERT INTO pcPay_EIG_Authorize (idOrder, amount, vaultToken, paymentmethod, transtype, authcode, ccnum, ccexp, cctype, idCustomer, fname, lname, address, zip, captured, trans_id, pcSecurityKeyID) VALUES ("& pcTrueOrdnum &", "& pcBillingTotal &", '"& pcv_strCustomerVaultID2 &"', 'CC', '"& x_TransType &"', '"& strAuthorizationCode &"', '"& pcv_CardNum &"', '"& pcv_CardExp &"', '"& pcv_CardType &"', "& session("idCustomer") &", N'"&replace(pcBillingFirstName,"'","''")&"', N'"&replace(pcBillingLastName,"'","''")&"', N'"&replace(pcBillingAddress,"'","''")&"', '"& pcBillingPostalCode &"', 0, '"& strTransactionID &"', "& pcs_GetKeyID &");"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						set rs=nothing

					End If

				End If
				
				'Log successful transaction
				call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 1)

				Session("CardType") = ""
				Session("CardNum") = ""
				Session("CardExp") = ""
				Session("CustomerVaultID") = ""
				Session("VaultID") = ""
				Session("SF_IsSaved") = ""
				session("GWAuthCode") = strAuthorizationCode
				session("GWTransId") = strTransactionID
				session("GWTransType") = x_TransType

				Response.redirect "gwReturn.asp?s=true&gw=EIG"

			Else '// If strResult="1" Then

				'Log failed transaction
				call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)
				
				response.Redirect("gwEIGateway.asp?Error=" & strResultText)

			End If '// If strResult="1" Then
			response.End()

		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// END:  PROCESS RESULTS WHEN TOKEN EXISTS
		'/////////////////////////////////////////////////////////////////////////////////////////////

		END IF

		If len(strError)=0 Then
			strError=getUserInput(Request("Error"),0)
		End If
		%>

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
            
            <% If strError<>"" Then %>
                <div class="pcErrorMessage"><%=strError%></div>
            <% End If %>

                    
            <% call pcs_showBillingAddress %>

            <%					
            Dim pcSavedCardsCount
            pcSavedCardsCount=0
            query="SELECT pcPay_EIG_Vault_ID, pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardExp FROM pcPay_EIG_Vault WHERE idCustomer="& Session("idCustomer") &" AND IsSaved=1"
            set rs=Server.CreateObject("ADODB.RecordSet")
            set rs=connTemp.execute(query)
            if NOT rs.eof then
                pcArray_SavedCards = rs.GetRows()
                pcSavedCardsCount = Ubound(pcArray_SavedCards)
            end if
            set rs=nothing
            %>
            <% If pcSavedCardsCount>0 AND (NOT pcIsSubscription) Then %>

                    <div class="row">
                        <div class="col-sm-8">   
                            <h3><%=dictLanguage.Item(Session("language")&"_EIG_7")%></h3>
                        </div>
                    </div>                    <%
                    '// Saved Credit Card
                    %>
                    <form method="POST" name="form-saved-card" id="form-saved-card" class="form-horizontal">

                    <div class="row">
                        <div class="col-sm-4">   

                            <select class="form-control" name="VaultID" id="VaultID">
                            <%
                            For SavedCardsCounter=0 to ubound(pcArray_SavedCards,2)
                                %>
                                <option value="<%=pcArray_SavedCards(0,SavedCardsCounter)%>"><%=pcArray_SavedCards(1,SavedCardsCounter)%> (<%=dictLanguage.Item(Session("language")&"_GateWay_8")%> <%=pcArray_SavedCards(2,SavedCardsCounter)%>)</option>
                                <%
                            Next
                            %>
                            </select>&nbsp;&nbsp;<a href="CustviewPayment.asp" target="_blank"><%=dictLanguage.Item(Session("language")&"_EIG_9")%></a> 

                        </div>
                    </div>
                    <div class="row">
                        <div class="col-sm-4">   
                     
                            <a class="pcButton pcButtonBack" href="<%=tempURL%>">
                                <img src="<%=pcf_getImagePath("",rslayout("back"))%>">
                                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
                            </a>
                            &nbsp;
                            <button class="pcButton pcButtonSubmit" id="submit-saved-card">
                                <img src="<%=pcf_getImagePath("",rslayout("pcLO_placeOrder"))%>" name="submit-saved-card" style="border:none; cursor:pointer"> 
                                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>                                    
                            </button>
                            <script type="text/javascript">
                                $pc(document).ready(function() {
                                    $pc('#submit-saved-card').click(function() {
                                        //$pc("#PleaseWaitMsg").html('<img src="<%=pcf_getImagePath("images","ajax-loader1.gif")%>" width="20" height="20" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_EIG_19"))%>');
                                        //$pc("#PleaseWaitDialog").dialog('open');
                                        //$pc(".ui-dialog-titlebar").css({'display' : 'none'});
                                        //$pc("#PleaseWaitDialog").css({'min-height' : '50px'});
                                        var tmpdata="";
                                        tmpdata=$pc('#VaultID').val();
                                        $pc.ajax(
                                               {
                                                type: "GET",
                                                url: "gwEIGatewayURL.asp",
                                                data: "VaultID=" + tmpdata + '&token=' + (new Date()).getTime(),
                                                timeout: 45000,
                                                success: function(data, textStatus){
                                                    if (data.indexOf("OK||")>=0) {
                                                        var tmpArr=data.split("||")
                                                        $pc("#form-saved-card").attr("action", tmpArr[1]);
                                                        $pc("#form-saved-card").submit();
                                                        return true;
                                                    } else {
                                                        window.location.href = 'gwEIGateway.asp?Error=' + data;
                                                        return false;
                                                    }
                                                }
                                        });
                                    });
                                });
                            </script> 
                    
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-sm-4">   
                            <h3><%=dictLanguage.Item(Session("language")&"_EIG_18")%></h3>
                        </div>
                    </div>
                </form>

            <% End If %>

            <%
            '// New Credit Card
            %>
            <% If NOT pcIsSubscription Then %>
                <form method="POST" name="form-new-card" id="form-new-card" class="form">
            <% Else %>
                <% 'SB S %>
                <form method="POST" name="form-new-card" id="form-new-card" class="form" action="gwEIGateway.asp">
                <input type="hidden" name="PaymentGWEIG" value="Go">
                <% 'SB E %>
            <% End If %>

            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label>Card Type:</label>               
                        <select class="form-control" name="x_Card_Type" id="x_Card_Type">
                            <% 	
                            x_TypeArray=Split(x_Type,"||")
                            If ubound(x_TypeArray)=1 Then
                                x_Type2=x_TypeArray(1)
                                cardTypeArray=split(x_Type2,", ")
                                i=ubound(cardTypeArray)
                                cardCnt=0
                                do until cardCnt=i+1
                                    cardVar=cardTypeArray(cardCnt)
                                    select case cardVar
                                        case "V"
                                            response.write "<option value=""V"" selected>Visa</option>"
                                            cardCnt=cardCnt+1
                                        case "M"
                                            response.write "<option value=""M"">MasterCard</option>"
                                            cardCnt=cardCnt+1
                                        case "A"
                                            response.write "<option value=""A"">American Express</option>"
                                            cardCnt=cardCnt+1
                                        case "D"
                                            response.write "<option value=""D"">Discover</option>"
                                            cardCnt=cardCnt+1
                                    end select
                                loop
                            End If 
                            %>
                        </select>            
                    </div> 
                </div>
            </div>   
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></label>
                        <input class="form-control" type="text" name="billing-cc-number" id="billing-cc-number" value="" autocomplete="off">
                    </div> 
                </div>
            </div>  
                      
            <div class="row">

                <!--<label><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></label>-->
                <div class="col-xs-2">
                    <div class="form-group">
                        <input type="hidden" name="billing-cc-exp" id="billing-cc-exp" value="">
                        <% dtCurYear=Year(date()) %>
                        <label><%=dictLanguage.Item(Session("language")&"_GateWay_9")%></label>
                        <select class="form-control" name="billing-cc-exp1" id="billing-cc-exp1">
                            <option value="01">1</option>
                            <option value="02">2</option>
                            <option value="03">3</option>
                            <option value="04">4</option>
                            <option value="05">5</option>
                            <option value="06">6</option>
                            <option value="07">7</option>
                            <option value="08">8</option>
                            <option value="09">9</option>
                            <option value="10">10</option>
                            <option value="11">11</option>
                            <option value="12">12</option>
                        </select>
                    </div>
                </div>
                <div class="col-xs-2">
                    <div class="form-group">
                        <label><%=dictLanguage.Item(Session("language")&"_GateWay_10")%></label>
                        <select class="form-control" name="billing-cc-exp2" id="billing-cc-exp2">
                            <option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
                            <option value="<%=right(dtCurYear+1,2)%>"><%=dtCurYear+1%></option>
                            <option value="<%=right(dtCurYear+2,2)%>"><%=dtCurYear+2%></option>
                            <option value="<%=right(dtCurYear+3,2)%>"><%=dtCurYear+3%></option>
                            <option value="<%=right(dtCurYear+4,2)%>"><%=dtCurYear+4%></option>
                            <option value="<%=right(dtCurYear+5,2)%>"><%=dtCurYear+5%></option>
                            <option value="<%=right(dtCurYear+6,2)%>"><%=dtCurYear+6%></option>
                            <option value="<%=right(dtCurYear+7,2)%>"><%=dtCurYear+7%></option>
                            <option value="<%=right(dtCurYear+8,2)%>"><%=dtCurYear+8%></option>
                            <option value="<%=right(dtCurYear+9,2)%>"><%=dtCurYear+9%></option>
                            <option value="<%=right(dtCurYear+10,2)%>"><%=dtCurYear+10%></option>
                        </select>
                    </div>
                </div>

            </div>
                            
            <% If x_CVV="1" Then %>
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></label>
                        <input name="billing-cvv" type="text" id="billing-cvv" value="" size="4" maxlength="4" class="form-control">
                    </div>
                </div>
            </div> 
                
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155">
                    </div>
                </div>
            </div>
                
            <% End If %>

            <%
            'SB S
            if pcIsSubscription Then
            %>
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label><%=scSBLang7%></label>
                        <%= money((pcBillingTotal + pcBillingSubScriptionTotal))%>
                    </div>
                </div>
            </div>
                
            <% Else %>
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></label>
                        <%= scCurSign & money(pcBillingTotal)%> 
                    </div>
                </div>
            </div>
                
            <%
            End if
            'SB E
            %>

            <%'SB S
            If pcIsSubscription Then %>
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label><%=scSBLang8%></label>
                        <!--#include file="inc_sb_widget.asp"-->
                    </div>
                </div>
            </div>
            
            <% End If
            'SB E %>

            <%
            'SB S
            If pcIsSubscription AND scSBaymentPageText <>"" Then %>
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label><%=scSBLang9%></label>
                        <%=scSBaymentPageText%>
                    </div>
                </div>
            </div>
                
            <% End If %>
            
            <% If pcIsSubscription AND pcv_intIsTrial AND scSBPaymentPageTrialText <> "" Then %>
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                        <label><%=scSBLang10%></label>
                        <%=scSBPaymentPageTrialText%>
                    </div>
                </div>
            </div>
                
            <%
            End if
            'SB E
            %>

            <% If x_SaveCards="1" AND (NOT pcIsSubscription) Then %>
            
            <div class="row">
                <div class="col-sm-4">                          
                    <div class="form-group">
                
                        <label>
                            <%=dictLanguage.Item(Session("language")&"_EIG_1")%>
                        </label>
    
                        <input name="x_SaveCards" id="x_SaveCards" type="checkbox" value="x_SaveCards" <%if Session("SF_IsSaved")="true" then response.Write("checked")%> class="clearBorder"/>
                        <img src="<%=pcf_getImagePath("images","pc_icon_info.png")%>" width="20" height="20" alt="<%=dictLanguage.Item(Session("language")&"_EIG_1")%>" title="<%=dictLanguage.Item(Session("language")&"_EIG_2")%>">
                    </div>
                </div>
            </div>

            <% End If %>

            <% 'SB S
             If (pcIsSubscription) Then %>

                <div class="pcFormButtons">
                    <!--#include file="inc_gatewayButtons.asp"-->
                </div>

            <% 'SB E
            Else %>

                <div class="pcFormButtons">
                    <a class="pcButton pcButtonBack" href="<%=tempURL%>">
                      <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
                      <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
                    </a>
                    &nbsp;
                    <button class="<%= "pcButton " & buttonClass %>" name="submit-new-card" id="submit-new-card">
                      <img src="<%=pcf_getImagePath("",rslayout("pcLO_placeOrder"))%>">
                      <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_placeorder") %></span>
                    </button>
                </div>
                
                <script type="text/javascript">
                    $pc(document).ready(function() {
                        $pc('#submit-new-card').click(function() {
                            //$pc("#PleaseWaitMsg").html('<img src="<%=pcf_getImagePath("images","ajax-loader1.gif")%>" width="20" height="20" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_EIG_19"))%>');
                            //$pc("#PleaseWaitDialog").dialog('open');
                            //$pc(".ui-dialog-titlebar").css({'display' : 'none'});
                            //$pc("#PleaseWaitDialog").css({'min-height' : '50px'});
                            var cardType = $pc('#x_Card_Type').val();
                            <% If (x_UseVault<>1) AND (x_TransType="AUTH_ONLY") Then %>
                                var CardNum = $pc('#billing-cc-number').val();
                            <% Else %>
                                var CardNum = '';
                            <% End If %>
                            var exp1 = $pc('#billing-cc-exp1').val();
                            var exp2 = $pc('#billing-cc-exp2').val();
                            var exp_date = exp1 + exp2;
                            $pc('#billing-cc-exp').val(exp_date);
                            var tmpdata="";
                            if ($pc('#x_SaveCards').prop('checked')) {
                                tmpdata=true
                            } else {
                                tmpdata=false
                            }

                            $pc.ajax(
                                   {
                                    type: "GET",
                                    url: "gwEIGatewayURL.asp",
                                    data: "IsSaved=" + tmpdata + '&CardType=' + cardType + '&CardNum=' + CardNum + '&ExpDate=' + exp_date + '&token=' + (new Date()).getTime(),
                                    timeout: 45000,
                                    success: function(data, textStatus){
                                        if (data.indexOf("OK||")>=0) {
                                            var tmpArr=data.split("||")
                                            $pc("#form-new-card").attr("action", tmpArr[1]);
                                            $pc("#form-new-card").submit();
                                            return true;
                                        } else {
                                            window.location.href = 'gwEIGateway.asp?Error=' + data;
                                            return false;
                                        }
                                    }
                            });
                        });
                    });
                </script>

            <% End If %>
            <%
            '// Secure Form Fields
            %>
            <input type="hidden" name="billing-first-name" value="<%=pcf_FixXML(pcBillingFirstName)%>">
            <input type="hidden" name="billing-last-name" value="<%=pcf_FixXML(pcBillingLastName)%>">
            <input type="hidden" name="billing-address1" value="<%=pcf_FixXML(pcBillingAddress)%>">
            <input type="hidden" name="billing-address2" value="<%=pcf_FixXML(pcBillingAddress2)%>">
            <input type="hidden" name="billing-city" value="<%=pcf_FixXML(pcBillingCity)%>">
            <input type="hidden" name="billing-state" value="<%=pcf_FixXML(pcBillingState)%>">
            <input type="hidden" name="billing-postal" value="<%=pcf_FixXML(pcBillingPostalCode)%>">
            <input type="hidden" name="billing-country" value="<%=pcf_FixXML(pcBillingCountryCode)%>">
            <input type="hidden" name="billing-phone" value="<%=pcf_FixXML(pcBillingPhone)%>">
            <input type="hidden" name="billing-fax" value="<%=pcf_FixXML(pcShippingFax)%>">
            <input type="hidden" name="billing-email" value="<%=pcf_FixXML(pcCustomerEmail)%>">
            <input type="hidden" name="billing-company" value="<%=pcf_FixXML(pcBillingCompany)%>">
            
            <% If len(pcShippingAddress)>0 Then %>
                <input type="hidden" name="shipping-address1" value="<%=pcf_FixXML(pcShippingAddress)%>">
                <input type="hidden" name="shipping-address2" value="<%=pcf_FixXML(pcShippingAddress2)%>">
                <input type="hidden" name="shipping-city" value="<%=pcf_FixXML(pcShippingCity)%>">
                <input type="hidden" name="shipping-state" value="<%=pcf_FixXML(pcShippingState)%>">
                <input type="hidden" name="shipping-postal" value="<%=pcf_FixXML(pcShippingPostalCode)%>">
                <input type="hidden" name="shipping-country" value="<%=pcf_FixXML(pcShippingCountryCode)%>">
                <input type="hidden" name="shipping-phone" value="<%=pcf_FixXML(pcShippingPhone)%>">
                <input type="hidden" name="shipping-fax" value="<%=pcf_FixXML(pcShippingFax)%>">
                <input type="hidden" name="shipping-email" value="<%=pcf_FixXML(pcShippingEmail)%>">
                <input type="hidden" name="shipping-company" value="<%=pcf_FixXML(pcShippingCompany)%>">
            <% End If %>

        </form>
    </div>
</div>

<% '// Functions
 function IsCreditCard(ByRef anCardNumber, ByRef asCardType)
	Dim lsNumber		' Credit card number stripped of all spaces, dashes, etc.
	Dim lsChar			' an individual character
	Dim lnTotal			' Sum of all calculations
	Dim lnDigit			' A digit found within a credit card number
	Dim lnPosition		' identifies a character position In a String
	Dim lnSum			' Sum of calculations For a specific Set

	' Default result is False
	IsCreditCard = False

	' ====
	' Strip all characters that are Not numbers.
	' ====

	' Loop through Each character inthe card number submited
	For lnPosition = 1 To Len(anCardNumber)
		' Grab the current character
		lsChar = Mid(anCardNumber, lnPosition, 1)
		' if the character is a number, append it To our new number
		if validNum(lsChar) Then lsNumber = lsNumber & lsChar

	Next ' lnPosition

	' ====
	' The credit card number must be between 13 and 16 digits.
	' ====
	' if the length of the number is less Then 13 digits, then Exit the routine
	if Len(lsNumber) < 13 Then Exit function

	' if the length of the number is more Then 16 digits, then Exit the routine
	if Len(lsNumber) > 16 Then Exit function

	' Choose action based on Type of card
	Select Case LCase(asCardType)
		' VISA
		Case "visa", "v", "V"
			' if first digit Not 4, Exit function
			if Not Left(lsNumber, 1) = "4" Then Exit function
		' American Express
		Case "american express", "americanexpress", "american", "ax", "A"
			' if first 2 digits Not 37, Exit function
			if Not Left(lsNumber, 2) = "37" AND Not Left(lsNumber, 2) = "34" Then Exit function
		' Mastercard
		Case "mastercard", "master card", "master", "M"
			' if first digit Not 5, Exit function
			if Not Left(lsNumber, 1) = "5" Then Exit function
		' Discover
		Case "discover", "discovercard", "discover card", "D"
			' if first digit Not 6, Exit function
			if Not Left(lsNumber, 1) = "6" Then Exit function

		Case Else
	End Select ' LCase(asCardType)

	' ====
	' if the credit card number is less Then 16 digits add zeros
	' To the beginning to make it 16 digits.
	' ====
	' Continue Loop While the length of the number is less Then 16 digits
	While Not Len(lsNumber) = 16

		' Insert 0 To the beginning of the number
		lsNumber = "0" & lsNumber

	Wend ' Not Len(lsNumber) = 16

	' ====
	' Multiply Each digit of the credit card number by the corresponding digit of
	' the mask, and sum the results together.
	' ====

	' Loop through Each digit
	For lnPosition = 1 To 16

		' Parse a digit from a specified position In the number
		lnDigit = Mid(lsNumber, lnPosition, 1)

		' Determine if we multiply by:
		'	1 (Even)
		'	2 (Odd)
		' based On the position that we are reading the digit from
		lnMultiplier = 1 + (lnPosition Mod 2)

		' Calculate the sum by multiplying the digit and the Multiplier
		lnSum = lnDigit * lnMultiplier

		' (Single digits roll over To remain single. We manually have to Do this.)
		' if the Sum is 10 or more, subtract 9
		if lnSum > 9 Then lnSum = lnSum - 9

		' Add the sum To the total of all sums
		lnTotal = lnTotal + lnSum

	Next ' lnPosition

	' ====
	' Once all the results are summed divide
	' by 10, if there is no remainder Then the credit card number is valid.
	' ====
	IsCreditCard = ((lnTotal Mod 10) = 0)

End function ' IsCreditCard
%>
<!--#include file="footer_wrapper.asp"-->
<%
Function pcf_GetNode(responseXML, nodeName, nodeParent)
	Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)
	myXmlDoc.loadXml(responseXML)
	Set Nodes = myXmlDoc.selectnodes(nodeParent)
	For Each Node In Nodes
		pcf_GetNode = pcf_CheckNode(Node,nodeName,"")
	Next
	Set Node = Nothing
	Set Nodes = Nothing
	Set myXmlDoc = Nothing
End Function

Function pcf_CheckNode(Node,tagName,default)
	Dim tmpNode
	Set tmpNode=Node.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		pcf_CheckNode=default
	Else
		pcf_CheckNode=Node.selectSingleNode(tagName).text
	End if
End Function

Function pcf_FixXML(str)
	str=replace(str, "&","and")
	pcf_FixXML=str
End Function

Public Function pcf_EIGChars(pgwTransId)
	pgwTransId=replace(pgwTransId,chr(0),"")
	pgwTransId=replace(pgwTransId,chr(13),"")
	pgwTransId=replace(pgwTransId,chr(10),"")
	pgwTransId=replace(pgwTransId,chr(34),"")
	pcf_EIGChars=trim(pgwTransId)
End Function

Public Function deformatNVP(nvpstr)
	On Error Resume Next

	Dim AndSplitedArray, EqualtoSplitedArray, Index1, Index2, NextIndex
	Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
	AndSplitedArray = Split(nvpstr, "&", -1, 1)
	NextIndex=0
	For Index1 = 0 To UBound(AndSplitedArray)
		EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
		For Index2 = 0 To UBound(EqualtoSplitedArray)
			NextIndex=Index2+1
			NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
			'response.Write(URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex)) & "<br />")
			Index2=Index2+1
		Next
	Next
	Set deformatNVP = NvpCollection

End Function

Function URLDecode(str)
	On Error Resume Next

	str = Replace(str, "+", " ")
	For i = 1 To Len(str)
	sT = Mid(str, i, 1)
		If sT = "%" Then
			sR = sR & Chr(CLng("&H" & Mid(str, i+1, 2)))
			i = i+2
		Else
			sR = sR & sT
		End If
	Next
	URLDecode = sR
End Function
%>
