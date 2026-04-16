<%
Public Sub pcs_DisplayFraudDetails()
    On Error Resume Next
    Dim rs

    If scFraud_IsEnabled = 1 Then
    
        query = "SELECT faAccountId, faRiskScore "
        query = query & "FROM orders "
        query = query & "WHERE idOrder = " & Session("ehidOrder") & ";"
        Set rs=connTemp.execute(query)
        If Not rs.eof Then
            pcv_strRiskScore = rs("faRiskScore")
            pcv_intAccountId = rs("faAccountId") 
            
            If pcv_strRiskScore <> "0" Then
                %>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <th colspan="2">Advanced Fraud Screening</th>
                </tr>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <td colspan="2">Risk Score: <%=pcv_strRiskScore %></td>
                </tr>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <%  
            End If
                   
        End If
        Set rs=nothing  
             
    End If
       
End Sub


Public Sub pcs_CheckForFraud()
    'On Error Resume Next
    Dim rs

    '// Check for Fraud
    '// 0 = error message stored in pcv_strErrorMsg;	
    '// 1 = redirect to msgb.asp;
    pcv_strRedirect = 1
    
    If scFraud_IsEnabled = 1 Then

        If session("reqCardNumber") = "" Then
            session("reqCardNumber") = getUserInput(request.Form("cardNumber"),16)
        End If

        If session("reqCardNumber")<>"" And Session("pcBillingCountryCode")<>"" Then
    
            pcv_riskScore = 0
        
            query = "SELECT * FROM pcWebServiceFraud ORDER BY pcPay_FA_Id DESC"
            set rs=connTemp.execute(query)
            
            pcPay_FA_RiskScore = rs("pcPay_FA_RiskScore")
            pcPay_FA_LicenseKey = rs("pcPay_FA_LicenseKey")
            pcPay_FA_SendShipping = rs("pcPay_FA_SendShipping")
            pcPay_FA_SendEmail = rs("pcPay_FA_SendEmail")
            pcPay_FA_SendPhone = rs("pcPay_FA_SendPhone")
            pcPay_FA_RiskScoreLock = rs("pcPay_FA_RiskScoreLock")
            pcPay_FA_RiskScoreEmail = rs("pcPay_FA_RiskScoreEmail")

            dim emailDomain, phonePrefix
            pcCustomerEmail = Session("pcCustomerEmail")
            If InStr(1, pcCustomerEmail, "@", 1) > 0 Then
                emailDomain = Right(pcCustomerEmail, Len(pcCustomerEmail) - InStr(1, pcCustomerEmail, "@", 1))
            End If
            
            If Session("pcBillingCountryCode") = "US" AND Session("pcBillingPhone") <> "" Then
                Dim myRegExp, regExpResult
                Set myRegExp = New RegExp
                myRegExp.Pattern = "[^\d]*"
                myRegExp.Global = true
                phonePrefix = myRegExp.Replace(Session("pcBillingPhone"), "")
                If Len(phonePrefix) >= 6 Then 
                    phonePrefix = Left(phonePrefix, 6)
                    phonePrefix = Left(phonePrefix, 3) &"-" &Right(phonePrefix, 3)
                Else
                    phonePrefix = ""
                End If
            End If
        
            '// IP Address
            pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
            If pcCustIpAddress="::1" Then pcCustIpAddress = "216.255.247.166"


            '// 1) Gen Request
            Dim jsonService : Set jsonService = JSON.parse("{}")
            
            '// Required Fields
            jsonService.Set "ip", pcCustIpAddress
            jsonService.Set "useragent", cstr(Request.ServerVariables("HTTP_USER_AGENT"))
            jsonService.Set "transactionid", session("pcGatewayDataIdOrder")
            
            '// Account
            jsonService.Set "userid", scPCWS_Uid
            jsonService.Set "username", Session("pcIdCustomer")
            
            '// Billing
            jsonService.Set "firstname", Session("pcBillingFirstName")
            jsonService.Set "lastname", Session("pcBillingLastName")
            jsonService.Set "company", Session("pcBillingCompany")
            jsonService.Set "address", Session("pcBillingAddress")
            jsonService.Set "address2", Session("pcBillingAddress2")   
            jsonService.Set "city", Session("pcBillingCity")
            If len(Session("pcBillingStateCode"))>=0 Then
                jsonService.Set "region", Session("pcBillingStateCode")
            Else
                jsonService.Set "region", Session("pcBillingProvince")
            End If
            jsonService.Set "postal", Session("pcBillingPostalCode")
            jsonService.Set "country", Session("pcBillingCountryCode")   
            
            '// Shipping
            If pcPay_FA_SendShipping = 1 Then
                jsonService.Set "shipfirstname", Session("pcShippingFirstName")
                jsonService.Set "shiplastname", Session("pcShippingLastName")
                jsonService.Set "shipcompany", Session("pcShippingCompany")
                jsonService.Set "shipaddress", Session("pcShippingAddress")  
                jsonService.Set "shipaddress2", Session("pcShippingAddress2")  
                jsonService.Set "shipcity", Session("pcShippingCity")
                If len(Session("pcShippingStateCode"))>=0 Then
                    jsonService.Set "shipregion", Session("pcShippingStateCode")
                Else
                    jsonService.Set "shipregion", Session("pcShippingProvince")
                End If
                jsonService.Set "shippostal", Session("pcShippingPostalCode")
                jsonService.Set "shipcountry", Session("pcShippingCountryCode")        
            End If
            
            '// User Data
            If pcPay_FA_SendPhone = 1 Then
                jsonService.Set "phoneNumber", phonePrefix
                jsonService.Set "shipphoneNumber", "" 
            End If
            
            If pcPay_FA_SendEmail = 1 Then
                jsonService.Set "email", emailDomain
            End If
            
            '// Bank Data
            jsonService.Set "issuerIdNumber", Left(session("reqCardNumber"), 6)
        
            '// Transaction Information    
            jsonService.Set "amount", Session("pcBillingTotal")
            jsonService.Set "currency", "USD"
            jsonService.Set "shopid", scPCWS_Uid
            jsonService.Set "txn_type", "creditcard"
            
            '// Credit Card Check
            jsonService.Set "avs_result", "Y"
            jsonService.Set "cvv_result", "N"'
            
            '// Miscellaneous
            jsonService.Set "requested_type", "standard"
            jsonService.Set "forwardedIP", Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            
            Dim jsonObj
            jsonObj = JSON.stringify(jsonService, null, 2)    
            jsonObj = "" & jsonObj & ""

            '// 2) Send off the Service Activation
            query="SELECT pcPCWS_Uid, pcPCWS_AuthToken, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
            Set rs=connTemp.execute(query)
            If Not rs.eof Then
                pcv_strUid = rs("pcPCWS_Uid")
                pcv_AuthToken = rs("pcPCWS_AuthToken")  
                pcv_strUsername = rs("pcPCWS_Username")  
                pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
            End If
            Set rs=nothing            
            
            session("reqCardNumber") = ""             
            cfuResult = pcf_PostRequest(jsonObj, pcv_marketURL & "api/fraudalert/score", pcv_AuthToken)

            '// 3) Parse Response
            pcv_strMessage = ""    
            If len(cfuResult)>0 Then 
            
                If instr(cfuResult, "Message") > 0 Then
                    Dim Info2 : Set Info2 = JSON.parse(cfuResult)   
                    pcv_strMessage = Info2.Message
                End If   
                
                If instr(cfuResult, "score")>0 Then
                    Dim Info3 : Set Info3 = JSON.parse(cfuResult)   
                    pcv_riskScore = Info3.score
                    pcv_distance = Info3.distance
                    pcv_countryMatch = Info3.countryMatch
                    pcv_highRiskCountry = Info3.highRiskCountry
                End If 

                If pcv_riskScore <> "" And pcv_riskScore > 0 Then
                
                    query="UPDATE orders SET faAccountId='" & scPCWS_Uid & "', faRiskScore=" & pcv_riskScore & " WHERE idOrder=" & session("pcGatewayDataIdOrder")
                    set rs=server.CreateObject("ADODB.RecordSet")
                    set rs=connTemp.execute(query)
                    set rs=nothing

                    pcv_strErrorMsg = ""
                    If (Round(cdbl(pcv_riskScore), 2) > Round(cdbl(pcPay_FA_RiskScore), 2)) OR (pcv_distance = 0) OR (lcase(pcv_countryMatch) = "no") OR (lcase(pcv_highRiskCountry) = "yes") Then
                        
                        pcv_strErrorMsg = dictLanguage.Item(Session("language")&"_pcFraud")
                        
                        If (Round(cdbl(pcv_riskScore), 2) > Round(cdbl(pcPay_FA_RiskScoreLock), 2)) Then 
                            '// Lock Account
                            call pcs_lockCustomerAccount(Session("pcIdCustomer"))
                            session("SFClearCartURL") = "msg.asp?message=56" 
                            response.Redirect("CustLO.asp") 
                        End If
                        
                        If (Round(cdbl(pcv_riskScore), 2) > Round(cdbl(pcPay_FA_RiskScoreEmail), 2)) Then
                            '// Send Email
                            fSubject=dictLanguage.Item(Session("language")&"_security_4")
                            fBody="Store administrator, someone is attempting to place an order, which received a risk score of " & pcv_riskScore & ". Their email is " & Session("pcCustomerEmail") & ". We recommend that you lock this account.  "
                            'call sendmail (scEmail, scEmail, scFrmEmail, fSubject, fBody)
                        End If

                        If pcv_strRedirect = 1 Then
                            Session("message") = pcv_strErrorMsg
                            Session("backbuttonURL") = Session("tempURL") & "?psslurl=" & session("redirectPage") & "&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId") & "&ordertotal=" & Session("pcBillingTotal")
                            session("reqCardNumber") = ""
                            response.redirect "msgb.asp?back=1"
                        End If
                        
                    End If

                End If
                 
            End If
            
        End If
    
    End If

End Sub
%>