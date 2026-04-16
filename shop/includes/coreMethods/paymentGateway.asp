<%
Public Sub pcs_showBillingAddress

    if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
        <div class="pcSpacer"></div>
        <p><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></p>
    <% end if %>
    <div class="pcSpacer"></div>
    <div class="daOPCCCBillAdd">
    <div class="daOPCCCPaySecTitle">
        Card Billing Address:
    </div>
    <div class="pcSpacer"></div>
    <p><%=pcBillingFirstName&" "&pcBillingLastName%></p>
    <p><%=pcBillingAddress%></p>
    <% if pcBillingAddress2<>"" then %>
      <p><%=pcBillingAddress2%></p>
    <% end if %>
    <p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p>
    <p><a href="onepagecheckout.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p>
</div>
    <% 
    If x_testmode="1" _
                Or WP_testmode="YES" _
                Or pcPay_VM_Testmode="1" _
                Or pcPay_Uep_TestMode="1" _
                Or x_testmode="1" _
                Or pcPay_TD_TestMode=1 _
                Or pcPay_Cys_TestMode="0" _
                Or pcPay_TW_Testmode=1 _
                Or pcPay_TGS_Testmode="1" _
                Or TCTestMode=1 _
                Or pcPay_SkipJack_TestMode="1" _
                Or pcPay_SecPay_TestMode = 1 _
                Or v2co_TestMode=1 _
                Or pcPay_ACH_TestMode=1 _
                Or pcv_testmode="1" _
                Or pcPay_BS_Testmode=1 _
                Or pcBPTestmode="TEST" _
                Or x_testmode="1" _
                Or pcvPay_CBN_test = 1 _
                Or CP_testmode="YES" _
                Or pcPay_Cys_TestMode="0" _
                Or pcPay_Dow_TestMode="1" _
                Or pcPay_EM_Testmode=1 _
                Or pcPay_EPN_TestMode=1 _
                Or pcEwayTestmode=1 _
                Or pcPay_GP_Testmode=1 _
                Or pcPay_HSBC_TestMode="1" _
                Or IsTestmode="1" _
                Or lp_testmode="YES" _
                Or pcPay_Cys_TestMode="0" _
                Or pcPay_Moneris_TestMode="1" _
                Or pcPay_OG_TestMode=1 _
                Or pcPay_OMG_Testmode=1 _
                Or pcPay_ParaData_TestMode = 1 _
                Or pcPay_PJ_Testmode=1 _
                Or pcPay_PT_Testing="1" _
                Or pcPay_PaymentExpress_TestMode = 1 _
                Or pcPay_ACH_TestMode=1 _
                Or ProtxTestmode="1" OR ProtxTestmode="2" _
                Or pcv_StrProtxTestmode<>0 _
                Or psi_testmode="YES" _
                Or pcv_PSI_TestMode="YES" _
                Or psi_XMLTestmode="YES" _
    Then %>

        <div class="pcSpacer"></div>
        <div class="pcInfoMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div>

    <% End If %>
    
    <div class="pcSpacer"></div>
    <div class="daOPCCCPaySecTitle">
        Card Payment Details:
    </div>
    <div class="pcSpacer"></div>
    <%           
End Sub


Function getSetPaymentId(idpayment, pcIdPayment)

    ' pidPayment = (this is saved at the end and very important)
    ' pcIdPayment = (this is the value that comes from the current database session

    '// If in session, then set local from session.
    '// If NOT in session, then set local from database
    If session("pcSFIdPayment")<>"" And session("pcSFIdPayment")<>"0" Then
        pidPayment = session("pcSFIdPayment")
    Else
        pidPayment = pcIdPayment
    End If

    '// If has user input, then over-ride existing 
    If idpayment<>"" Then
	    pidPayment = URLDecode(idpayment)
    End If

    If pidPayment = "" Then
	    pidPayment = pcIdPayment
    Else
	    If Not IsNumeric(pidPayment) Then
		    pidPayment = pcIdPayment
	    End If
    End If

    '// Check for Empty
    If len(pidPayment)=0 OR pidPayment=0 Then
        pidPayment = Session("DefaultIdPayment") 
    End If 
    
    '// Set current value to session
    session("pcSFIdPayment") = pidPayment
    
    'pcIdPayment = pidPayment
    getSetPaymentId = pidPayment

End Function

Public Function calculatePaymentGatewayFees(pidPayment, pcIsSubscription)
    
    paymentTotal=0
    
    'SB S
    strAndSub = ""
    If pcIsSubscription = True Then
        strAndSub = " pcPayTypes_Subscription <> 0 "
    Else
        strAndSub = " idPayment=" & pidPayment
    End if 
    'SB E
    
    If pidPayment<>0 And pidPayment<>"" And pidPayment<>999999 Then
     
        query="SELECT paymentDesc, priceToAdd, percentageToAdd FROM paytypes WHERE " & StrandSub
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=connTemp.execute(query)        
        If Not rs.Eof Then    
            pPaymentPriceToAdd=rs("priceToAdd")
            pPaymentpercentageToAdd=rs("percentageToAdd")
        End If        
        set rs=nothing
             
    ElseIf pidPayment=0 Then '// If pidPayment<>0 And pidPayment<>"" And pidPayment<>999999 Then
                
        'SB S
        strAndSub = ""
        If pcIsSubscription = True Then
            strAndSub = " AND pcPayTypes_Subscription = 1 ORDER by pcPayTypes_Subscription, paymentPriority"
        Else
            strAndSub = " ORDER by paymentPriority"
        End If 
        'SB E
   
        If session("customerType")=1 Then
            query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE active=-1 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
        Else
            query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE active=-1 AND Cbtob=0 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
        End If         
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=connTemp.execute(query)

        If NOT rs.eof Then
            pPaymentPriceToAdd=rs("priceToAdd")
            pPaymentpercentageToAdd=rs("percentageToAdd")
        End If
        set rs=nothing
          
        pcIdPayment = pidPayment
                    
    Else '// If pidPayment<>0 And pidPayment<>"" And pidPayment<>999999 Then
                
        If pidPayment=999999 Then

            query="SELECT paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE gwcode=46 OR gwcode=53 OR gwcode=999999;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=connTemp.execute(query)
                
            If rs.eof Then
                pPaymentPriceToAdd=0
                pPaymentpercentageToAdd=0
            Else
                pPaymentPriceToAdd=rs("priceToAdd")
                pPaymentpercentageToAdd=rs("percentageToAdd")
            End If
            set rs=nothing

        End If
                    
    End If '// If pidPayment<>0 And pidPayment<>"" And pidPayment<>999999 Then
                
    If pSubTotal>0 Then
        intCalPaymnt = pSubTotal
    Else
        intCalPaymnt = subtotal
    End If
                
    '// Add payment amount
    If ccur(pPaymentPriceToAdd)<>0 Or ccur(pPaymentpercentageToAdd)<>0 Then 
                
        tempPercentageToAdd = (pPaymentpercentageToAdd*intCalPaymnt/100)
        tempPercentageToAdd=roundTo(tempPercentageToAdd,.01)
        tempTaxPercentageToAdd=(pPaymentpercentageToAdd*pTaxableTotal/100)
        tempTaxPercentageToAdd=roundTo(tempTaxPercentageToAdd,.01)
        
        paymentTotal = pPaymentPriceToAdd + tempPercentageToAdd 
        taxPaymentTotal = pPaymentPriceToAdd + tempTaxPercentageToAdd '// Processing fees on taxable total (only if percentage)           

    End If
    
    calculatePaymentGatewayFees = paymentTotal

End Function
	
Function getIsPayPalClassicEnabled()
	query = "SELECT gwCode FROM payTypes WHERE gwCode IN (46, 999999)"
	Set rsPayPal = conntemp.execute(query)
	If Not rsPayPal.eof Then
		getIsPayPalClassicEnabled = true
	Else
		getIsPayPalClassicEnabled = false
	End If
	Set rsPayPal = Nothing
End Function

Function getPFApiGatewayCode(orderID)
	getPFApiGatewayCode = 0
	
	'// Get gateway code
	'// PayPal Payments Pro, PayPal Payments Advanced, or Payflow Link
	query="SELECT idOrder, gwCode FROM pcPay_PFL_Authorize WHERE idOrder=" & orderID & " AND gwCode IN (53, 80, 9, 99);"
	set rsQ=connTemp.execute(query)
	If Not rsQ.eof Then
		getPFApiGatewayCode = cint(rsQ("gwCode"))
	End If
	set rsQ=nothing
End Function

Function getIsPFApiEnabled()
	query = "SELECT gwCode FROM payTypes WHERE gwCode IN (53, 80, 9, 99)"
	Set rsPayPal = conntemp.execute(query)
	If Not rsPayPal.eof Then
		getIsPFApiEnabled = true
	Else
		getIsPFApiEnabled = false
	End If
	Set rsPayPal = Nothing
End Function

Function getPFApiExpressTitle(gwCode)
	paymentTitle = "PayPal Express Checkout"

	'// Re-generate title based on gateway code
	Select Case gwCode
	Case 53
		paymentTitle = "PayPal Payments Pro (Express Checkout)"
	Case 80
		paymentTitle = "PayPal Payments Advanced (Express Checkout)"
	Case 9, 99
		paymentTitle = "PayPal Payflow Link (Express Checkout)"
	End Select

	getPFApiExpressTitle = paymentTitle
End Function

%>