<%if Ucase(scDB)<>"SQL" then
	response.Clear()
	call closeDb()
response.redirect "sm_Access.asp"
end if%>
