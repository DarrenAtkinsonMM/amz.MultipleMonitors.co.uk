<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% On Error Resume Next %>
<!--#include file="pcCPLog.asp" -->
<%
pcv_strAdminPrefix="1"

If (session("admin") = 0 OR session("admin") = 1 OR session("admin") = "") _
	OR _
	((instr(session("PmAdmin"),"*")=0 And instr(session("PmAdmin"),"19")=0)) _
	OR _
	(len(session("CUID"))=0) _
	OR _
	(session("admin." & pcf_getAdminToken()) <> Session.SessionID) Then

	call closeDb()
    response.redirect "login_1.asp?RedirectURL=" & Server.URLEncode(pcv_filePath)
    
end if

Dim cpAccessArr, cpAccessArrCount, pcUserArr, pcUserArrCount, pcFoundPermission, pcFoundPermissionTotal

'// Get array of user level permissions allowed on the current page
'// Permission level 0 means page open to the public
if instr(PmAdmin,"*")=0 then 
	PmAdmin=PmAdmin&"*"
end if
cpAccessArr=split(PmAdmin,"*")
cpAccessArrCount=ubound(cpAccessArr)-1

'// Find out if this is the Master User or if the page is open to the public
if (session("PmAdmin") = "19") or (not isNull(findUser(cpAccessArr,0,cpAccessArrCount))) then

'// Display page

else

	'// Get the array of permissions assigned to the user
	pcUserArr = split(session("PmAdmin"),"*")
	pcUserArrCount = ubound(pcUserArr)-1

	'// Loop through to see if any of them match the page permissions
	pcFoundPermissionTotal=0
	For k=0 To pcUserArrCount
		if isNull(findUser(cpAccessArr,pcUserArr(k),cpAccessArrCount)) then
			pcFoundPermission=0
		else
			pcFoundPermission=1
		end if
		pcFoundPermissionTotal=pcFoundPermissionTotal + pcFoundPermission
	Next

	'// None of the permissions match: no access to the page and redirect
	if pcFoundPermissionTotal = 0 then
		call closeDb()
		response.redirect "menu.asp?msg=" & server.URLEncode("You do not have enough permissions to access the selected page.")
	end if
		
end if

Private Function pcf_getAdminToken()
	pcv_strLocalAddress = Request.ServerVariables("LOCAL_ADDR") 
	pcv_strLocalSessionID = Session.SessionID
	pcv_strAdminToken = pcv_strLocalAddress & "." & pcv_strLocalSessionID
	pcf_getAdminToken = pcv_strAdminToken
End Function
%>
