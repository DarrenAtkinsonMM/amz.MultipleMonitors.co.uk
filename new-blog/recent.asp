<%
	call openDb()
	query2="SELECT TOP 5 pcCont_IDPage,pcCont_PageName,pcUrl FROM pcContents WHERE pcCont_Blog=1 AND pcCont_InActive=0 ORDER BY pcCont_PubDate DESC"
	set rsPop=server.CreateObject("ADODB.Recordset")
	set rsPop=connTemp.execute(query2)
	
	'loop through return records to build up list
	Do While Not rsPop.EOF
		strList = strList & "<li><a href=""/blog/" & rsPop("pcUrl") & "/"">" & rsPop("pcCont_PageName") & "</a></li>"
		rsPop.MoveNext()
	Loop
	
	set rsPop=nothing	
	call closeDB()
	
	response.write(strList)
%>
								