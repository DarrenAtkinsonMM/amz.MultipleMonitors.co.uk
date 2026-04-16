<!-- #include file="../inc_productcart.asp" -->
<!-- #include file="uploader.asp" -->
<%
on error resume next

Dim savefile
Dim tempFile
Dim message
dim UploadObject
Dim UploadErrorCode, UploadErrorFile, UploadErrorStr
	
Function GetErrorMessage(FileName, FileType, Code)
	Select Case Trim(Code)
	Case "FILENAME"
		GetErrorMessage = FileName & " - The name of this file is invalid. Please check the filename and try again."
	Case "FILETYPE"
		GetErrorMessage = FileName & " - You are not allowed to upload files of type <strong>*." + FileType + "</strong> to this store. Please use a different file format and try again."
	Case "NOTALLOWED"
		GetErrorMessage = "You do you not have the neccessary permissions to upload files. Please login or try again later."
	End Select
End Function

tempPath = getUserInput(Request("folder"), 0) 
showMessage = getUserInput(Request("message"), 0) 


'// ProductCart Mod: Add path offset - START
If imageFullPath <> "true" Then
	tempPath = "..\" & tempPath
End If
'// ProductCart Mod: Add path offset - END

savePath = Server.MapPath(tempPath)

Set UploadObject = New Uploader

UploadObject.allowedTypes = "gif|jpg|png|wma|wmv|swf|doc|zip|pdf|txt|mp4|ogv|webm|csv"

UploadObject.Save(savePath)

If UploadErrorCode&"" <> "" Then
	If showMessage = "1" Then
		UploadErrorStr = GetErrorMessage(UploadErrorFile, UploadObject.GetExtension(UploadErrorFile), UploadErrorCode)
	Else
		Response.Write UploadErrorCode
		Response.End
	End If
End If

%>