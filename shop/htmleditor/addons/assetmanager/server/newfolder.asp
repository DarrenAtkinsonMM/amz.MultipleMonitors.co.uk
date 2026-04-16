<!-- #include file="../inc_productcart.asp" -->
<%
Dim savefile
Dim tempFile

tempPath = getUserInput(Request.Form("folder"), 0)
				
'// ProductCart Mod: Add path offset - START
If imageFullPath <> "true" Then
	tempPath = "..\" & tempPath
End If
'// ProductCart Mod: Add path offset - END

savePath = Server.MapPath(tempPath)
Dim fs
Set fs=Server.CreateObject("Scripting.FileSystemObject")
if fs.FolderExists(savePath) = false then
    fs.CreateFolder(savePath)
End If
Set fs=nothing

Response.Status = 200
%>
