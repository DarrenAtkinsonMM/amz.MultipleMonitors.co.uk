<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->

<%
If Request("x_response_code") <> "" Then

	pcIdCustomer = Request("x_cust_id")
	x_TransType = UCase(Request("x_type"))
	pcGatewayDataIdOrder = Request("x_invoice_num")
	
	Session("DefaultIdPayment") = 1103
	
	If Request("x_response_code") = 1 Then
		
		If x_TransType="AUTH_ONLY" Then
			
			pcTrueOrdnum = (CLng(pcGatewayDataIdOrder)-scpre)
			pcv_SecurityKeyID = pcs_GetKeyID
			
			If (pcTrueOrdnum="") OR (IsNull(pcTrueOrdnum)) OR (pcTrueOrdnum<"0") then
				Response.Write "An error occurred during processing. Cannot get OrderID#. Please try again later."
				Response.end
			End if
			If Not(IsNumeric(pcTrueOrdnum)) then
				Response.Write "An error occurred during processing. Cannot get OrderID#. Please try again later."
				Response.end
			End if
			
			query="INSERT INTO authorders (idOrder, amount, paymentmethod, transtype, authcode, ccnum, ccexp, idCustomer, fname, lname, address, zip, captured, pcSecurityKeyID) VALUES ("&pcTrueOrdnum&", "&Request("x_amount")&", 'DPM', '"&x_TransType&"', '"&Request("x_auth_code")&"', '"&Request("x_account_number")&"', '"&Request("expMonth")&Request("expYear")&"', "&pcIdCustomer&", '"&replace(Request("x_first_name"),"'","''")&"', '"&replace(Request("x_last_name"),"'","''")&"', '"&replace(Request("x_address"),"'","''")&"', '"&Request("x_zip")&"', 0, "&pcv_SecurityKeyID&");"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
		
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
		
		If scSSL="" OR scSSL="0" Then
			x_redirectURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwReturn.asp?s=true&gw=DPM&type=" & Request("x_type") & "&auth="&Request("x_auth_code") & "&trans="&Request("x_trans_id") & "&avs=" & Request("x_avs_code") & "&cvv=" & Request("x_cvv2_resp_code") & "&cav=" & Request("x_cavv_resp_code")),"//","/")
			x_redirectURL=replace(x_redirectURL,"https:/","https://")
			x_redirectURL=replace(x_redirectURL,"http:/","http://") 
		Else
			x_redirectURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwReturn.asp?s=true&gw=DPM&type=" & Request("x_type") & "&auth="&Request("x_auth_code") & "&trans="&Request("x_trans_id") & "&avs=" & Request("x_avs_code") & "&cvv=" & Request("x_cvv2_resp_code") & "&cav=" & Request("x_cavv_resp_code")),"//","/")
			x_redirectURL=replace(x_redirectURL,"https:/","https://")
			x_redirectURL=replace(x_redirectURL,"http:/","http://")
		End If

	Else
		
		'Log failed transaction
		call pcs_LogTransaction(pcIdCustomer, pcGatewayDataIdOrder, 0)

		'call closeDb()
		message = "<b>Error code " & Request("x_response_code") & ": " & Request("x_response_reason_text")
		
		If scSSL="" OR scSSL="0" Then
			x_redirectURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwAuthorizeDPM_error.asp?msg=" & message & "&idCust=" & pcIdCustomer & "&idOrder=" & pcGatewayDataIdOrder & "&amount=" & Request("x_amount")),"//","/")
			x_redirectURL=replace(x_redirectURL,"https:/","https://")
			x_redirectURL=replace(x_redirectURL,"http:/","http://") 
		Else
			x_redirectURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwAuthorizeDPM_error.asp?msg=" & message & "&idCust=" & pcIdCustomer & "&idOrder=" & pcGatewayDataIdOrder & "&amount=" & Request("x_amount")),"//","/")
			x_redirectURL=replace(x_redirectURL,"https:/","https://")
			x_redirectURL=replace(x_redirectURL,"http:/","http://")
		End If
		
	End If
End If
%>

<html>
<head>
	<script type="text/javascript" charset="utf-8">
		window.location='<%=x_redirectURL%>';
 	</script>
    <noscript>
    	<meta http-equiv="refresh" content="1;url=<%=x_redirectURL%>">
    </noscript>
</head>
<body></body>
</html>