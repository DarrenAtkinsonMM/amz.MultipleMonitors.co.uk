<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%
ClearCartURL = session("SFClearCartURL")
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"--> 
<!--#include file="../includes/CashbackConstants.asp"--> 
<%
' If coming from custLO.asp, clear customer session
' (addresses issue with shared SSL scenario)
if request.QueryString("action") = "clear" then
	if session("admin")<>0 then
		session("idcustomer")=""
		session("pcStrCustName")=""
		session("customerCategory")= Cint(0)
		session("customerType")= Cint(0)
		session("ATBCustomer")= Cint(0)
		session("ATBPercentOff")= Cint(0)
		session("customerCategoryType")=""
		session("CustomerGuest")=""
	else
		Session.Abandon
	end if
end if
%>
<!--#include file="pcStartSession.asp"-->
<%
call closeDb()

'if there is an alternate homepage set, then redirect to it, else redirect to the default page (index.asp)
If ClearCartURL<>"" then
    session("SFClearCartURL") = ""
	Response.Status = "301 Moved Permanently" 
	Response.AddHeader "Location", ClearCartURL
	Response.End
Else
	If scURLredirect <>"" then
		Response.Status = "301 Moved Permanently" 
		Response.AddHeader "Location", scURLredirect
		Response.End
	else
		Response.Status = "301 Moved Permanently" 
		Response.AddHeader "Location", "home.asp"
		Response.End
	end If
End if
%>
