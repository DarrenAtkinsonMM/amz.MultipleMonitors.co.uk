<!-- #include file="../inc_productcart.asp" -->
<%
Dim savefile
Dim tempFile

tempFile = getUserInput(Request.Form("file"), 0)

'// ProductCart Mod: Add path offset - START
If imageFullPath <> "true" Then
	tempFile = "..\" & tempFile
End If
'// ProductCart Mod: Add path offset - END

savefile = Server.MapPath(tempFile)

Dim fs
Set fs=Server.CreateObject("Scripting.FileSystemObject")
If fs.FileExists(savefile) = true Then
    fs.DeleteFile(savefile)
End If
Set fs=nothing

Response.Status = 200
%>
