<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>

<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 'Gateway specific files %>
<%
'SB S
Dev_Testmode = 2
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
<div id="pcMain">
	<div class="pcMainContent">
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
		query="SELECT gwBT_MerchantID, gwBT_PublicKey, gwBT_PrivateKey, gwBT_CurCode, gwBT_CVV, gwBT_Mode, gwBT_TestMode FROM gwBrainTree Where gwBT_ID=1"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		x_Type= "M,V,A,D"  
		
		x_Login=rs("gwBT_MerchantID")
		x_Login=enDeCrypt(x_Login, scCrypPass)
		
		x_Password=rs("gwBT_PublicKey")
		x_Password=enDeCrypt(x_Password, scCrypPass)
		
		x_Key=rs("gwBT_PrivateKey")
		x_Key=enDeCrypt(x_Key, scCrypPass)
		
		x_Curcode=rs("gwBT_CurCode")
		
		'x_AIMType=rs("x_AIMType")  '//  ????????????????????????????????
		
		x_CVV=rs("gwBT_CVV")
		
		x_testmode=rs("gwBT_TestMode")
		
		'x_secureSource=rs("x_secureSource")  '//  ????????????????????????????????
		
		x_TransType =rs("gwBT_Mode")
		set rs=nothing
	
		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// END:  PROCESS RESULTS 
		'/////////////////////////////////////////////////////////////////////////////////////////////
		If Request.Form("PaymentGWBraintree")="Go" Then %>
        
        	<% session("redirectPage")="gwBrainTree.asp" %>
            
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
			'// By pass BrainTree if the immediate order value is 0 
			If pcBillingTotal<0 Then
				pcBillingTotal=0
			End If
			If (pcIsSubscription) AND (pcBillingTotal=0) Then 

				session("reqCardNumber")=getUserInput(request.Form("cardNumber"),16)
				session("reqExpMonth")=getUserInput(request.Form("expMonth"),0)
				session("reqExpYear")=getUserInput(request.Form("expYear"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("CVV"),4)					
				pExpiration=getUserInput(request("expMonth"),0) & "/01/" & getUserInput(request("expYear"),0)				
				
				'// Validates expiration
			    if DateDiff("d", Month(Now)&"/01/"&Year(now),pExpiration)<=-1 then
                    call closeDb()
                    Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_6")
                    Session("backbuttonURL") = tempURL & "?psslurl=gwBrainTree.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"
			    end if
		       	
				'// Validate card
			    if not IsCreditCard(session("reqCardNumber"), request.form("x_Card_Type")) then
                    call closeDb()
                    Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_5")
                    Session("backbuttonURL") = tempURL & "?psslurl=gwBrainTree.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1" 
			    end if 
				
				session("GWAuthCode")	= "AUTH-BT" 
				session("GWTransId")	= "0" 

                call closeDb()
				Response.Redirect("gwReturn.asp?s=true&gw=BT&GWError=1")
				Response.End 
				
			Else

				'// Normal Payment Required, Let Pass
				session("reqCardNumber")=getUserInput(request.Form("cardNumber"),16)
				session("reqExpMonth")=getUserInput(request.Form("expMonth"),0)
				session("reqExpYear")=getUserInput(request.Form("expYear"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("CVV"),4)	
				
			End if 
			'SB E
			%>



			<% 
			if pcBillingTotal > 0 then
			
				If x_CVV="1" Then			 
					if not isnumeric(session("reqCVV")) or len(session("reqCVV")) < 3 or len(session("reqCVV")) > 4 Then				 
                        call closeDb()
                        Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_7")&dictLanguage.Item(Session("language")&"_paymntb_c_4")
                        Session("backbuttonURL") = tempURL & "?psslurl=gwBrainTree.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                        response.redirect "msgb.asp?back=1"
					End If
				 End if 
	
				'// Send the request to the SubscriptionBridge BT.
	
				If x_testmode="1" Then
					
				Else
				
				End If
	


				
				If x_CVV="1" Then
				'	stext=stext & "&x_Card_Code=" & session("reqCVV")
				End If
				
				'stext=stext & "&x_customer_ip=" & pcCustIpAddress
	
				If x_TransType = 0 Then
				'	stext=stext & "&x_Type=" & x_TransType '// Sale
				Else
				
				End If
				
				
				


	
				'// Send the transaction info as part of the querystring
				Set objSB = NEW pcARBClass
				
				
				
				if x_testmode="1" then

				else

				end if
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// TRANSACTION
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				objSB.CartRegularAmt = pcBillingTotal
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// CREDIT CARD
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				objSB.PayInfoExpMonth= session("reqExpMonth")
				If len(session("reqExpYear"))=2 Then
					objSB.PayInfoExpYear = "20" & session("reqExpYear")		
				Else
					objSB.PayInfoExpYear = session("reqExpYear")	
				End If
				objSB.PayInfoCardNumber = left(session("reqCardNumber"),16)
				objSB.PayInfoAccountNumber = right(PayInfoCardNumber,4)
				objSB.PayInfoCardType = session("reqCardType")
				objSB.PayInfoCVVNumber = session("reqCVV")
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// BILLING ADDRESS
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'stext=stext & "&x_Currency_Code=" & x_Curcode
				'stext=stext & "&x_Description=" & replace(scCompanyName,",","-") & " Order: " & session("GWOrderID")
				'stext=stext & "&x_Invoice_Num=" & session("GWOrderID")
				'stext=stext & "&x_Cust_ID=" & pcv_IncreaseCustID

				objSB.BillingFirstName = pcBillingFirstName
				objSB.BillingLastName = pcBillingLastName
				objSB.BillingCompany = replace(pcBillingCompany,",","||")
				objSB.BillingAddress = replace(pcBillingAddress,",","||")
				objSB.BillingAddress2 = pcBillingAddress2
				objSB.BillingCity = pcBillingCity
				objSB.BillingPostalCode = pcBillingPostalCode
				objSB.BillingStateCode = pcBillingState
				objSB.BillingProvince = pcBillingState
				objSB.BillingCountryCode = pcBillingCountryCode
				objSB.BillingPhone = pcBillingPhone
				objSB.CustomerEmail = pcCustomerEmail
			
			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// SHIPPING ADDRESS
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				if pcShippingFullName<>"" then
					pcShippingNameArry=split(pcShippingFullName, " ")
					if ubound(pcShippingNameArry)>0 then
						pcShippingFirstName=pcShippingNameArry(0)
						if ubound(pcShippingNameArry)>1 then
							 tmpShipFirstName = pcShippingFirstName&" "
							 pcShippingLastName = replace(pcShippingFullName,tmpShipFirstName,"")
						else
							pcShippingLastName=pcShippingNameArry(1)
						end if
					else
						pcShippingFirstName=pcShippingFullName
						pcShippingLastName=pcShippingFullName
					end if
				else
					pcShippingFirstName=pcBillingFirstName
					pcShippingLastName=pcBillingLastName
				end if

				objSB.ShippingFirstName = pcShippingFirstName
				objSB.ShippingLastName = pcShippingLastName
				objSB.ShippingCompany = pcShippingCompany
				objSB.ShippingAddress = replace(pcShippingAddress,",","||")
				objSB.ShippingAddress2 = pcShippingAddress2
				objSB.ShippingCity = pcShippingCity
				objSB.ShippingPostalCode = pcShippingPostalCode
				objSB.ShippingStateCode = pcShippingState
				objSB.ShippingProvince = pcShippingState
				objSB.ShippingCountryCode = pcShippingCountryCode
				objSB.ShippingPhone = pcShippingPhone
				objSB.ShippingEmail = pcShippingEmail
				
				query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
				set rsAPI=connTemp.execute(query)
				if not rsAPI.eof then
					Setting_APIUser=rsAPI("Setting_APIUser")
					Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
					Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
				end if
				set rsAPI=nothing
	
				result = objSB.TransactionRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)
				If len(result)>0 AND instr(result,"ErrorDetail")=0 Then
					ack = objSB.pcf_GetNode(result,"Ack","//TransactionResponse")
					if len(ack)>0 then ack=UCase(ack)
				Else
					ackDesc = objSB.pcf_GetNode(result,"ErrorDetail","//Error")													
				End if		

				'// Check the ErrorCode to make sure that the component was able to talk to the authorization network
				If ack="SUCCESS" then					
								
					pcv_SecurityPass = scCrypPass
					pcv_SecurityKeyID = pcs_GetKeyID
				
					dim pCardNumber, pCardNumber2
					pCardNumber=session("reqCardNumber")
					pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)

					session("GWAuthCode") = objSB.pcf_GetNode(result,"ProcessorResponseCode","//TransactionResponse")
					session("GWTransId") = objSB.pcf_GetNode(result,"TransactionID","//TransactionResponse")
					session("GWTransType") = x_TransType
					
					'// Save Batch Processing Record	
					If x_TransType = 0 Then

						'query="INSERT INTO pcPay_EIG_Authorize (idOrder, amount, vaultToken, paymentmethod, transtype, authcode, ccnum, ccexp, cctype, idCustomer, fname, lname, address, zip, captured, trans_id, pcSecurityKeyID) VALUES ("& pcTrueOrdnum &", "& pcBillingTotal &", '', 'CC', '"& x_TransType &"', '"& session("GWAuthCode") &"', '"& pCardNumber2 &"', '"& session("reqExpMonth")&session("reqExpYear") &"', '"& session("reqCardType") &"', "& session("idCustomer") &", '"&replace(pcBillingFirstName,"'","''")&"', '"&replace(pcBillingLastName,"'","''")&"', '"&replace(pcBillingAddress,"'","''")&"', '"& pcBillingPostalCode &"', 0, '"& session("GWTransId") &"', "& pcv_SecurityKeyID &");"

						'set rs=server.CreateObject("ADODB.RecordSet")
						'set rs=connTemp.execute(query)
						'set rs=nothing
	
						'if err.number<>0 then
						'	call LogErrorToDatabase()
						'	set rs=nothing
						'	call closedb()
						'	response.redirect "techErr.asp?err="&pcStrCustRefID
						'end if
						
					End If
	
					set rs=nothing


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

					If pcIsSubscription Then
						Response.redirect "gwReturn.asp?s=true&gw=BT"
					Else
						Response.redirect "gwReturn.asp?s=false&gw=BT"
					End If
					
					
				Else '// If ack="SUCCESS" then
				
					If NOT len(ackDesc)>0 Then
						ackDesc = objSB.pcf_GetNode(result,"Reason","//TransactionResponse")
					End If

                    call closeDb()
                    Session("message") = ackDesc
                    Session("backbuttonURL") = tempURL & "?psslurl=gwBrainTree.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"
						
				End If '// If ack="SUCCESS" then
				
			End if '// if pcBillingTotal > 0 then
			
		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// END:  PROCESS RESULTS 
		'/////////////////////////////////////////////////////////////////////////////////////////////

		Else

		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// START:  SHOW FORM
		'/////////////////////////////////////////////////////////////////////////////////////////////

				query="SELECT gwBT_Mode, gwBT_CVV FROM gwBrainTree Where gwBT_ID=1"
				set rs=Server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				x_TransType = rs("gwBT_Mode")
				x_Type = "V,M,A,D"
				x_CVV = rs("gwBT_CVV")
				
				M="0"
				V="0"
				A="0"
				D="0"
				%>	
	
				<% If x_CVV="1" Then 
					response.write "<form method=""POST"" action=""gwBrainTree.asp"" name=""form1"" class=""pcForms"">"
				Else %>
					<form action="gwBrainTree.asp" method="POST" name="form1" class="pcForms">
				<% End If %>

				<input type="hidden" name="PaymentGWBraintree" value="Go">

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                    
                    <% call pcs_showBillingAddress %>


            <div class="pcFormItem">
                <%=dictLanguage.Item(Session("language")&"_GateWay_5")%>
            </div>

            <div class="pcFormItem">
                <div class="pcFormLabel">Card Type:</div>
                <div class="pcFormField">
							<select name="x_Card_Type">
							<% 	
							If len(x_Type)>0 Then
								cardTypeArray=split(x_Type,",")
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
							End If %>
							</select>
                </div>
            </div>

            <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></div>
                <div class="pcFormField"><input type="text" name="CardNumber" value="" autocomplete="off"></div>
            </div>

					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></div>
						<div class="pcFormField"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="expMonth">
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
							<% dtCurYear=Year(date()) %>
							&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%> 
							<select name="expYear">
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
                    
					<% If x_CVV="1" Then %>
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155"></div>
                </div>
					<% End If %>
                    
                    <% 
					'SB S 
					if pcIsSubscription Then
					%>
                        <div class="pcFormItem"> 
                            <div class="pcFormLabel"><%=scSBLang7%></div>
                            <div class="pcFormField"><%= money((pcBillingTotal + pcBillingSubScriptionTotal))%></div> 
                        </div>
                    <% Else %>
                        <div class="pcFormItem"> 
                            <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
                            <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
                        </div>
                    <% 
					End if
					'SB E 
					%>
                    
                    <%'SB S 
					If pcIsSubscription Then %>
                        <div class="pcFormItem"> 
                            <div class="pcFormLabel"><%=scSBLang8%></div>
                            <div class="pcFormField">
                                <!--#include file="inc_sb_widget.asp"-->
                            </div> 
                        </div>
                    <% End If
					'SB E %>
					
					<%
                    'SB S
                    If pcIsSubscription AND scSBaymentPageText <>"" Then %>
                        <div class="pcFormItem"> 
                            <div class="pcFormLabel"><%=scSBLang9%></div>
                            <div class="pcFormField">
                                <%=scSBaymentPageText%>
                            </div> 
                        </div>
					<% End If %>
                    <% If pcIsSubscription AND pcv_intIsTrial AND scSBPaymentPageTrialText <> "" Then %>
                        <div class="pcFormItem"> 
                            <div class="pcFormLabel"><%=scSBLang10%></div>
                            <div class="pcFormField">
                                <%=scSBPaymentPageTrialText%>
                            </div> 
                        </div>
                    <% 
                    End if 
                    'SB E
                    %>					

            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
		<% 
		end if 
		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// END:  SHOW FORM
		'/////////////////////////////////////////////////////////////////////////////////////////////
		%>
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
