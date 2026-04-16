<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
pidOption=request.form("idOption")
poptionDescrip=replace(request.form("optionDescrip"),"'","''")
pidOptionGroup=request.form("idOptionGroup")
predirectURL=request.form("redirectURL")
pmode=request.form("mode")
pboton=request.form("modify")

pcv_OptImg=request.form("OptImg")
pcv_OptCode=request.form("OptCode")


	query="UPDATE options SET optionDescrip=N'" &poptionDescrip& "',pcOpt_Img='" & pcv_OptImg & "',pcOpt_Code='" & pcv_OptCode & "' WHERE idOption=" &pidOption

	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	set rstemp=nothing
	
	FCount=getUserInput(request("FCount"),0)
	if FCount>"0" then
		queryQ="DELETE FROM pcFCAttr WHERE idOption=" & pidOption & ";"
		set rsQ=connTemp.execute(queryQ)
		set rsQ=nothing

		For ik=1 to FCount
			tFC=getUserInput(request("FC" & ik),0)
			if tFC="" then
				tFC="0"
			end if
			if tFC>"0" then
				queryQ="INSERT INTO pcFCAttr (IdOption,pcFC_ID) VALUES (" & pidOption & "," & tFC & ");"
				set rsQ=connTemp.execute(queryQ)
				set rsQ=nothing
			end if
		Next
	end if

	if predirectURL<>"" then 
		call closeDb()
		response.redirect predirectURL&"&mode="&pmode
	else
		call closeDb()
		response.redirect "modOptGrpa.asp?idOptionGroup="&pidOptionGroup&"&s=1&msg="&Server.URLEncode("Attribute successfully updated.")
	end if
%>