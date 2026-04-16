<% response.Buffer=true %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
		<% CP_Server=Session("ship_CP_Server")
		CP_Username=Session("ship_CP_Username")
		CP_Password=Session("ship_CP_Password")
		CP_CustNo=Session("ship_CP_CustNo")
		CP_ServiceStr=Session("ship_CP_Service")
		CP_freeshipStr=Session("ship_CP_freeshipStr")
		CP_handlingStr=Session("ship_CP_handlingStr")
		CP_EMPackage=Session("ship_CP_EMPackage")
		CP_PMPackage=Session("ship_CP_PMPackage")
		CP_Height=Session("ship_CP_Height")
		CP_Width=Session("ship_CP_Width")
		CP_Length=Session("ship_CP_Length")

		set rs=Server.CreateObject("ADODB.Recordset")
		mySQL="UPDATE ShipmentTypes SET shipServer='"&CP_Server&"', userID='"&CP_Username&"', [password]='"&CP_Password&"', AccessLicense='"&CP_CustNo&"', active=-1 WHERE idShipment=7;"
		set rs=connTemp.execute(mySQL)
		
		'clear all informatin out of shipService for Canada Post
		mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE idShipment=7;"
		set rs=connTemp.execute(mySQL)
		
		Dim i
		shipServiceArray=split(CP_ServiceStr,", ")
		for i=0 to ubound(shipServiceArray)
			mySQL="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
			'response.write mySQL
			set rs=connTemp.execute(mySQL)
		next
		
		freeshipStrArray=split(CP_freeshipStr,",")
		for i=0 to (ubound(freeshipStrArray)-1)
			freeoveramt=split(freeshipStrArray(i),"|")
			mySQL="UPDATE shipService SET serviceFreeOverAmt="&freeoveramt(1)&" WHERE serviceCode='"&freeoveramt(0)&"';"
			'response.write mySQL
			set rs=connTemp.execute(mySQL)
		next
		
		handlingStrArray=split(CP_handlingStr,",")
		for i=0 to (ubound(handlingStrArray)-1)
			shiphandamt=split(handlingStrArray(i),"|")
			mySQL="UPDATE shipService SET serviceHandlingFee="&shiphandamt(1)&", serviceShowHandlingFee="&shiphandamt(2)&" WHERE serviceCode='"&shiphandamt(0)&"';"
			'response.write mySQL
			set rs=connTemp.execute(mySQL)
		next

		set rs=nothing
		
		call closeDb()
response.redirect "../includes/PageCreateCPConstants.asp"
		%>
<!--#include file="AdminFooter.asp"-->
