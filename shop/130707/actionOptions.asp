<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Dim intidoption, intidoptiongrp, strOptionName

If Request.QueryString("delete") <> "" Then
	intidoption=Request.QueryString("delete")
	intidoptiongrp=Request.QueryString("idOptionGroup")

	If statusAPP="1" OR scAPP=1 Then

		query="SELECT idoptoptgrp FROM options_optionsGroups WHERE idoption="&intidoption&" AND idOptionGroup="&intidoptiongrp&";"
		set rstemp=connTemp.execute(query)
		do while not rstemp.eof
			idoptoptgrp=rstemp("idoptoptgrp")
	
			query="UPDATE Products SET removed=-1,active=0 WHERE removed=0 AND ((pcprod_Relationship like '%[_]" & idoptoptgrp & "[_]%') OR (pcprod_Relationship like '%[_]" & idoptoptgrp & "'))"
			set rstemp1=conntemp.execute(query)
			set rstemp1=nothing
	
			rstemp.MoveNext
		loop
		set rstemp=nothing

	End If
	query="Delete From options_optionsGroups WHERE idoption="&intidoption&" AND idOptionGroup="&intidoptiongrp&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	query="Delete From optGrps WHERE idoption="&intidoption&" AND idOptionGroup="&intidoptiongrp&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing
	
	call closeDb()
	response.redirect "modOptGrpa.asp?s=1&idOptionGroup="&intidoptiongrp&"&msg="&server.URLencode("Option attribute successfully deleted.")
End If
%>