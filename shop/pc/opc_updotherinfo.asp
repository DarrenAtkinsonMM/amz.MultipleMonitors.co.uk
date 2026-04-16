<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/sendmail.asp"-->

<!--#include file="pcStartSession.asp" -->
<!--#include file="opc_contentType.asp" -->
<%
Call SetContentType()

Call pcs_CheckLoggedIn()

pcErrMsg=""

pcStrOrderNickName=URLDecode(getUserInput(request("OrderNickName"),250))
pcStrOrderComments=URLDecode(getUserInput(request("OrderComments"),0))

pcStrGcReName=URLDecode(getUserInput(request("GcReName"),250))
pcStrGcReEmail=URLDecode(getUserInput(request("GcReEmail"),250))
pcStrGcReMsg=URLDecode(getUserInput(request("GcReMsg"),0))

if pcStrGcReEmail<>"" then
	pcStrGcReEmail=replace(pcStrGcReEmail," ","")
	if instr(pcStrGcReEmail,"@")=0 or instr(pcStrGcReEmail,".")=0 then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_70")&"</li>"
	end if
end if

if pcErrMsg="" then
	query="UPDATE pcCustomerSessions SET pcCustSession_OrderName=N'" & pcStrOrderNickName & "',pcCustSession_Comment=N'" & pcStrOrderComments & "',pcCustSession_GcReName=N'" & pcStrGcReName & "',pcCustSession_GcReEmail='" & pcStrGcReEmail & "',pcCustSession_GcReMsg=N'" & pcStrGcReMsg & "' WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
	set rs=connTemp.execute(query)
	set rs=nothing
	OKmsg="OK"
end if

if pcErrMsg<>"" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_71")&"<br><ul>" & pcErrMsg & "</ul>"
	response.write pcErrMsg
else
	response.write OKmsg
end if

call closeDb()
%>


