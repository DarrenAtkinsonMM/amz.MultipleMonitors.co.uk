<%
'Save UTF-8 data to file
Public Sub pcs_SaveUTF8(filepath1,filepath2,filedata)
    On Error Resume Next
    Dim objFS
    Dim objFile
    Dim pcStrFileName

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	
	if PPD="1" then
		pcStrFileName=Server.Mappath (filepath1)
	else
		pcStrFileName=Server.Mappath (filepath2)
	end if

	Set objFile = CreateObject("ADODB.Stream")
	objFile.CharSet = "utf-8"
	objFile.Open

	objFile.WriteText filedata
	
	Call TraceStack("&FileName=" & pcStrFileName & "&FileData=" & filedata)

	objFile.SaveToFile TraceFileName(pcStrFileName), 2
	
	objFile.Close
	set objFS=nothing
	set objFile=nothing

End Sub

'Append UTF-8 data to file
Public Sub pcs_AppendUTF8(filepath1,filepath2,filedata)
    On Error Resume Next
	Dim objFS
	Dim objFile
	Dim pcStrFileName

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	
	if PPD="1" then
		pcStrFileName=Server.Mappath (filepath1)
	else
		pcStrFileName=Server.Mappath (filepath2)
	end if
	
	Set objFile = CreateObject("ADODB.Stream")
	objFile.Type=2
    objFile.mode=3
	objFile.CharSet = "utf-8"
	objFile.Open
	objFile.LoadFromFile(TraceFileName(pcStrFileName))
	
	strData = objFile.ReadText()
	
	Call TraceStack("&FileName=" & pcStrFileName & "&FileData=" & filedata)
	
	objFile.Position= 0
	objFile.WriteText=strData & filedata
		
	objFile.SaveToFile TraceFileName(pcStrFileName), 2
	
	objFile.Close
	set objFS=nothing
	set objFile=nothing

End Sub

'Open UTF-8 data to file
Public Function pcf_OpenUTF8(filepath1,filepath2)
    On Error Resume Next
    Dim objFS
    Dim objFile
    Dim pcStrFileName

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	
	if PPD="1" then
		pcStrFileName=Server.Mappath (filepath1)
	else
		pcStrFileName=Server.Mappath (filepath2)
	end if

	Set objFile = CreateObject("ADODB.Stream")
	objFile.Type=2
    objFile.mode=3
	objFile.CharSet = "utf-8"
	objFile.Open
	objFile.LoadFromFile(TraceFileName(pcStrFileName))
	
	strData = objFile.ReadText()
	
	pcf_OpenUTF8 = strData
	
	objFile.Close
	set objFS=nothing
	set objFile=nothing
    err.clear
End Function


Public Sub pcs_logEventUTF8(filename, data)
	Dim objFile, objFSO, findit

	if PPD = "1" then
		findit = Server.MapPath("/" & scPcFolder & "/includes/logs/" & filename)
	else
		findit = Server.MapPath("../includes/logs/" & filename)
	end if
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	If not objFSO.FileExists(findit) Then
		objFSO.CreateTextFile(findit)
	End If
	Set objFSO = nothing
	
	Set objFile = CreateObject("ADODB.Stream")
	objFile.Type=2
    objFile.mode=3
	objFile.CharSet = "utf-8"
	objFile.Open
	objFile.LoadFromFile(findit)
	
	strData = objFile.ReadText()
	
	objFile.Position= 0
	objFile.WriteText = strData & vbCrlf & vbCrlf & replace(data, vbCrlf, "")
		
	objFile.SaveToFile findit, 2
	
	objFile.Close
	set objFile=nothing
End Sub
%>