<% 
'// Check for same-page message
IF msg<>"" THEN
	pcStrMsg=trim(msg)
	pcvMessageType=msgType
'// Check for querystrings and forms
ELSE
	pcStrMsg=trim(request.querystring("msg"))
	if pcStrMsg="" then
		pcStrMsg=trim(request.querystring("message"))
	end if
	pcvMessageType=request.querystring("s")
END IF

'check msg from login
If pcStrMsg = "" AND Session("pcCPCheckText") <> "" Then 
	pcvMessageType = Session("pcCPCheckCode")
	pcStrMsg = Session("pcCPCheckText")
	
	Session("pcCPCheckCode") = ""
	Session("pcCPCheckText") = ""
End If
	
if pcStrMsg<>"" then
	if not validNum(pcvMessageType) then pcvMessageType=0
	if pcvMessageType=1 then %>
	<div class="pcCPmessageSuccess"><%=pcStrMsg%></div>
<% 
	elseif pcvMessageType=2 then %>
	<div class="pcCPmessageWarning"><%=pcStrMsg%></div>
<% 
	else 
%>
	<div class="pcCPmessage"><%=pcStrMsg%></div>
<% 
	end if 
end if
%>