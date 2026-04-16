<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f,g,z, randomnum

pidOptionGroup=request.form("idOptionGroup")
tmpFGID=getUserInput(request("pcFGid"),0)
if tmpFGID="" then
	tmpFGID=0
end if
saveFGID=getUserInput(request("saveFG"),0)
if saveFGID="" then
	saveFGID=0
end if
poptionGroupDesc=trim(replace(request.form("optionGroupDesc"),"'","''"))
poptionGroupDesc=replace(poptionGroupDesc,"""","&quot;")
if poptionGroupDesc = "" then
	call closeDb()
    response.redirect "modOptGrpa.asp?idOptionGroup="&pidOptionGroup&"&s=1&msg="&Server.URLEncode("Please enter a description for this option group")
end if

	
	'ensure that the new name does not already exist
	query="SELECT * FROM optionsGroups WHERE optionGroupDesc='"&poptionGroupDesc&"'"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	nomatch="0"
	do until rstemp.eof
		comparestring=StrComp(poptionGroupDesc, rstemp("optionGroupDesc"), 1) 
		if comparestring=-1 then
			nomatch="1"
		end if
	rstemp.movenext
	loop
	
	if nomatch="1" then
		set rstemp=nothing
		
		call closeDb()
response.redirect "modOptGrpa.asp?msg="&Server.Urlencode("Unable to rename group. There is already a group that uses this name.")&"&idOptionGroup="&pidOptionGroup
		response.end
	end if
	
	query="UPDATE optionsGroups SET optionGroupDesc=N'" &poptionGroupDesc& "' WHERE idOptionGroup=" &pidOptionGroup
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		set rstemp=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error modifying the option group name on modOptGrpb.asp") 
	end If
	
	set rstemp=nothing
	
	if tmpFGID<>"" then
		if Clng(saveFGID)<>Clng(tmpFGID) then
			query="DELETE FROM pcFCAttr WHERE idOption IN (SELECT idOption FROM OptGrps WHERE idOptionGroup=" & pidOptionGroup & ");"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
		
			query="DELETE FROM pcFGOG WHERE idOptionGroup=" &pidOptionGroup
			set rstemp=connTemp.execute(query)
			set rstemp=nothing 
			if tmpFGID>"0" then
				query="INSERT INTO pcFGOG (pcFG_ID,idOptionGroup) VALUES (" & tmpFGID & "," & pidOptionGroup & ");"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing
			end if
		end if
	end if
	
	call closeDb()
response.redirect "modOptGrpa.asp?idOptionGroup="&pidOptionGroup&"&s=1&msg="&Server.URLEncode("Option Group successfully updated!")
%>
