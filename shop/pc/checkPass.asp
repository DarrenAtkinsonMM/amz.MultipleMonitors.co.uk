<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="opc_contentType.asp" -->
<%Dim tmpResult
if tmpPass="" then
	tmpPass=getUserInput(request("pass"),0)
end if
tmpType=getUserInput(request("passtype"),0)
tmpEmail=getUserInput(request("email"),0)
pIdCustomer=""
pEmail=""
if ((session("idCustomer")>"0") OR (Session("adminidcustomer")>"0") OR (tmpEmail<>"")) AND (tmpType<>"Cp") then
	pIdCustomer=session("idCustomer")
	if (pIdCustomer="") OR (pIdCustomer<="0") then
		pIdCustomer=Session("adminidcustomer")
	end if
	if (pIdCustomer="") OR (pIdCustomer<="0") then
		pIdCustomer="0"
	end if
	if pIDCustomer="0" and tmpEmail<>"" then
		query="SELECT idCustomer,[email] FROM Customers WHERE email LIKE '" & tmpEmail & "';"
	else
		query="SELECT idCustomer,[email] FROM Customers WHERE idCustomer=" & pIdCustomer & ";"
	end if
	set rs=connTemp.execute(query)
	if not rs.eof then
		pIdCustomer=rs("idCustomer")
		pEmail=rs("email")
	end if
	set rs=nothing
end if

if (tmpPass="") OR ((tmpType<>"R") AND (tmpType<>"Rs") AND (tmpType<>"Cp")) then
	response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_0") & """}"
	response.End()
end if

if (tmpType="R") OR (tmpType="Rs") OR (tmpType="Cp") then
	tmpResult=pcf_CheckCommonPass(tmpPass)
	if tmpResult="1" then
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_01") & """}"
		response.End()
	end if

	tmpResult=pcf_CheckStrongPass(tmpPass,pEmail)
	Select Case tmpResult
	Case "1":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_1") & """}"
		response.End()
	Case "2":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_2") & """}"
		response.End()
	Case "3":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_3") & """}"
		response.End()
	Case "4":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_4") & """}"
		response.End()
	Case "5":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_5") & """}"
		response.End()
	Case "6":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_6") & """}"
		response.End()
	Case "7":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_7") & """}"
		response.End()
	Case "8":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_8") & """}"
		response.End()
	Case "9":
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_9") & """}"
		response.End()
	End Select
end if

if (tmpType="Rs") then
	tmpResult=pcf_CheckUsedPassH(pIdCustomer,pEmail,tmpPass)
	
	if tmpResult="1" then
		response.write  "{""isError"": ""true"",""errorMessage"": """ & dictLanguage.Item(Session("language")&"_newpass_10") & """}"
		response.End()
	end if
end if

response.write  "{""isError"": ""false"",""errorMessage"": """" }"%>