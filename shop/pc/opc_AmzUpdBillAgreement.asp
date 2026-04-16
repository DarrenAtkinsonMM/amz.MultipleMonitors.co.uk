<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%Dim pcStrPageName
pcStrPageName = "opc_AmzUpdBillAgreement.asp"%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_AmazonHeader.asp" -->
<%
	tmpscXML=".3.0"

	AmzBillingAgreementId=getUserInput(request("id"),0)
	if AmzBillingAgreementId="" then
		response.write "Error: Cannot get Amazon Billing Agreement ID#"
		response.End()
	end if
	session("AmzBillAgreementID")=AmzBillingAgreementId
%>
