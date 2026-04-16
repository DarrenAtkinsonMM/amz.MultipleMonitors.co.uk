<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
idnews=request("idnews")
query="select CustFile from News where idnews=" & idnews
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)

CustFile=rstemp("CustFile")
findit = Server.MapPath("newslists/" & CustFile)
Set fso = server.CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(findit)
f.Delete
Set fso = nothing
Set f = nothing
	
query="delete from News where idnews=" & idnews
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)

set rstemp=nothing


call closeDb()
response.redirect "manageNews.asp"
%>
