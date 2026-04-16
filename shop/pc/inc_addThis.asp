<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: AddThis
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddThis
	Dim rs,query,pcStrAddThisCode
	query="SELECT pcStoreSettings_AddThisCode FROM pcStoreSettings WHERE (((pcStoreSettings_ID)=1));"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcStrAddThisCode=rs("pcStoreSettings_AddThisCode")
	end if
	if trim(pcStrAddThisCode)<>"" and not IsNull(pcStrAddThisCode) then
		if scAddThisDisplay=1 then
			pcStrAddThisClass = "pcAddThisRight"
		else
			pcStrAddThisClass = "pcAddThis"
		end if
		
		Response.Write "<div class=""" & pcStrAddThisClass & """>"
		Response.Write pcStrAddThisCode

		'remove extra characters appended to URL
		Response.Write "<script type=""text/javascript"">" &vbNewLine
		Response.Write "var addthis_config = addthis_config||{};" &vbNewLine
		Response.Write "addthis_config.data_track_addressbar = false;" &vbNewLine
		Response.Write "</script>" &vbNewLine
		Response.Write "</div>"
	end if
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: AddThis
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>