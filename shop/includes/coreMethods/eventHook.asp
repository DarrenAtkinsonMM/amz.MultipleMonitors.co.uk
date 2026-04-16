<%
'////////////////////////////////////////////////////////////////////////////////////////
'// START:  EVENT HOOKS
'////////////////////////////////////////////////////////////////////////////////////////

'// Control Panel Menu
Public Sub pcs_hookPreCPanelMenu()
    call pcs_doEventHook("PreCPanelMenu")
    If len(Session("msg"))>0 Then
        msg=msg & Session("msg")
    End If
    Session("msg")=""
End Sub


'// Control Payment Info Tab
Public Sub pcs_hookCPanelTabPaymentInfo()    
    call pcs_doEventHook("CPanelTabPaymentInfo")
End Sub


'// Control Panel Footer JS
Public Sub pcs_hookCPanelFooterJS()
    call pcs_doEventHook("CPanelFooterJS")
End Sub


'// Pre-Payment Hook
Public Sub pcs_hookPrePayment(id)
    call pcs_doEventHook("PrePayment")
End Sub


'// Product Modified Hook
Public Sub pcs_hookProductModified(id, sku)
    call pcs_doEventHook("ProductModified")
End Sub


'// Product Removed Hook
Public Sub pcs_hookProductRemoved(id, sku)
    call pcs_doEventHook("ProductRemoved")
End Sub


'// Product Purged Hook
Public Sub pcs_hookProductPurged(id, sku)
    call pcs_doEventHook("ProductPurged")
End Sub


'// Stock Changed Hook
Public Sub pcs_hookStockChanged(id, sku)
    call pcs_doEventHook("StockChanged")
End Sub


'// "In Stock" Event Hook
Public Sub pcs_hookInStockEvent(id, sku) 
    call pcs_doEventHook("InStockEvent")
End Sub


'// Order is Processed Hook
Public Sub pcs_hookOrderProcessed(id)
    call pcs_doEventHook("OrderProcessed")
End Sub


'// Order is Completed Hook
Public Sub pcs_hookOrderCompleted(id)
    call pcs_doEventHook("OrderCompleted")
End Sub


'// Reset Password Email Sent
Public Sub pcs_hookCustResetPassEmailSent(email)
    call pcs_doEventHook("CustResetPassEmailSent")
End Sub


'// Send Alarm Email Sent
Public Sub pcs_hookSendAlarmEmailSent(email)
    call pcs_doEventHook("SendAlarmEmailSent")
End Sub


'// New Customer Email Sent
Public Sub pcs_hookNewCustEmailSent(email)
    call pcs_doEventHook("NewCustEmailSent")	
End Sub


'// Affiliate Retrieve Password Email Sent
Public Sub pcs_hookAffRetrievePassEmailSent(email)
    call pcs_doEventHook("AffRetrievePassEmailSent")
End Sub


'// Forgot Order Code Email Sent
Public Sub pcs_hookForgotOrderCodeEmailSent(email)
    call pcs_doEventHook("ForgotOrderCodeEmailSent")
End Sub


'// Contact Us Email Sent
Public Sub pcs_hookContactUsEmailSent(email)
    call pcs_doEventHook("ContactUsEmailSent")
End Sub


'// New Order Email Sent
Public Sub pcs_hookNewOrderEmailSent(email)
    call pcs_doEventHook("NewOrderEmailSent")	
End Sub


'// Order Confirmation Email Sent
Public Sub pcs_hookOrderConfirmationEmailSent(email)
    call pcs_doEventHook("OrderConfirmationEmailSent")	
End Sub


'// Order Received Email Sent
Public Sub pcs_hookOrderReceivedEmailSent(email)
    call pcs_doEventHook("OrderReceivedEmailSent")	
End Sub


'// Order Shipped Email Sent
Public Sub pcs_hookOrderShippedEmailSent(email)
    call pcs_doEventHook("OrderShippedEmailSent")	
End Sub


'// Order Partially Shipped Email Sent
Public Sub pcs_hookOrderPartShippedEmailSent(email)
    call pcs_doEventHook("OrderPartShippedEmailSent")	
End Sub


'// Gift Registry Order Email Sent
Public Sub pcs_hookGROrderEmailSent(email)
    call pcs_doEventHook("GROrderEmailSent")	
End Sub


'// Gift Certificate Order Email Sent
Public Sub pcs_hookGCOrderEmailSent(email)
    call pcs_doEventHook("GCOrderEmailSent")	
End Sub


'// Affiliate Order Email Sent
Public Sub pcs_hookAffOrderEmailSent(email)
    call pcs_doEventHook("AffOrderEmailSent")
End Sub


'// Help Desk Email Sent
Public Sub pcs_hookHelpDeskEmailSent(email)
    call pcs_doEventHook("HelpDeskEmailSent")
End Sub


'// Product Review Email Sent
Public Sub pcs_hookProductReviewEmailSent(email)
    call pcs_doEventHook("ProductReviewEmailSent")
End Sub


'////////////////////////////////////////////////////////////////////////////////////////
'// END:  EVENT HOOKS
'////////////////////////////////////////////////////////////////////////////////////////
%>