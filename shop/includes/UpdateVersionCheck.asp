<%
if InStr(scVersion,"5.3.00")=0 then
	updtrigger=1
	updDBScript="upddb_v5.3.00.asp"
	updSubVersion=""
else
	updtrigger=0
	updDBScript=""
end if
%>