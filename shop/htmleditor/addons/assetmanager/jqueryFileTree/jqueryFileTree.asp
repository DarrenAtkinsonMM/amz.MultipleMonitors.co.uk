<!--#include file="../inc_productcart.asp"-->
<%
'
' jQuery File Tree ASP (VBS) Connector
' Copyright 2008 Chazzuka
' programmer@chazzuka.com
' http://www.chazzuka.com/
'
Function URLDecode(str)
	str = Replace(str, "+", " ")
	For i = 1 To Len(str)
		sT = Mid(str, i, 1)
		If sT = "%" Then
			If i+2 < Len(str) Then
				sR = sR & _
					Chr(CLng("&H" & Mid(str, i+1, 2)))
				i = i+2
			End If
		Else
			sR = sR & sT
		End If
	Next
	URLDecode = sR
End Function

pcv_strDirectoryPath = getUserInput(Request("dir"), 0)
pcv_strSearchTerms = getUserInput(Request("search"), 0)
cacheCount = getUserInput(Request("count"), 0) 

' retrive base directory
dim BaseFileDir:BaseFileDir=URLDecode(pcv_strDirectoryPath)
' if blank give default value
if len(BaseFileDir)=0 then BaseFileDir="/userfiles/"

dim IsImg:IsImg=false
if(Request.QueryString("img")="yes") then
	IsImg = true
end if

searchTerms = URLDecode(pcv_strSearchTerms)
If IsNull(searchTerms) Then
	searchTerms = ""
End If
searchTerms = LCase(searchTerms)

page = Request.QueryString("page")
If Not IsNumeric(page) Or Len(page) < 1 Then
	page = 0
Else
	page = CInt(page)
End If

If Not IsNumeric(cacheCount) Or Len(cacheCount) < 1 Then
	cacheCount = 0
Else
	cacheCount = CLng(cacheCount)
End If

BaseFilePrefix = ""
If imageFullPath <> "true" Then
	BaseFilePrefix = "..\"
End If

folderCount = 0
startFile = 0
fileCount = 0
maxFiles = 3 * 100		'// Any multiple of 3 since images are usually generated in groups of 3
starttime = timer()

If page > 0 Then
	startFile = page * maxFiles
End If

Function GetFolderFiles(Folder, ShowFiles)
	loadedFiles = 0

	'LOOP THROUGH FILES
	For Each ObjFile In Folder.Files
	
		addFile = false
	
		If Len(searchTerms) > 0 Then
			i__Name=ObjFile.name
			i__Ext = LCase(Mid(i__Name, InStrRev(i__Name, ".", -1, 1) + 1))
			If InStr(lcase(i__Name), searchTerms) > 0 Then
				If loadedFiles < maxFiles Then
					addFile = true
				Else
					'// Don't bother searching for more files
					If cacheCount > 0 Then
						fileCount = cacheCount
						Exit For
					End If
				End If
	
				If ShowFiles Then
					fileCount = fileCount + 1
				End If
			End If
		Else
			If IsImg=True Then
				i__Name=ObjFile.name
				i__Ext = LCase(Mid(i__Name, InStrRev(i__Name, ".", -1, 1) + 1))

				If i__Ext="jpeg" or i__Ext="jpg" or i__Ext="png" or i__Ext="gif" Then
					If loadedFiles < maxFiles Then
						addFile = true
					End If
	
					If ShowFiles Then
						fileCount = fileCount + 1
					End If
				End If
			Else
				If loadedFiles < maxFiles Then
					i__Name=ObjFile.name
					i__Ext = LCase(Mid(i__Name, InStrRev(i__Name, ".", -1, 1) + 1))

					addFile = true
				End If
	
				If ShowFiles Then
					fileCount = fileCount + 1
				End If
			End If	
		End If

		If addFile Then
			If fileCount >= startFile Then
				If ShowFiles Then
					i__Date=ObjFile.DateLastModified
			
					Html = Html + "<li class=""file ext_"&i__Ext&""">"
					Html = Html + "  <a href=""#"" rel="""+(BaseFileDir+Replace(i__Name, "#", "%23"))+""">"&i__Name
					Html = Html + "    <span class=""dateModified"">(Last Modified: " & i__Date & ")</span>"
					Html = Html + "  </a>"
					Html = Html + "</li>"
					Html = Html + VBCRLF
				End If

				loadedFiles = loadedFiles + 1
			End If
		End If
	Next

	GetFolderFiles = loadedFiles
End Function

dim ObjFSO,BaseFile,Html
' resolve the absolute path
BaseFile = Server.MapPath(BaseFilePrefix & BaseFileDir)&"\"
' create FSO
Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
' if given folder is exists
if ObjFSO.FolderExists(BaseFile) then
	dim ObjFolder,ObjSubFolder,ObjFile,i__Name,i__Ext

	If page = 0 Then
		Html = Html +  "<ul class=""jqueryFileTree"" style=""display: none;"">"&VBCRLF
	End If
			
  Set ObjFolder = ObjFSO.GetFolder(BaseFile)

	If page = 0 Then
		' LOOP THROUGH SUBFOLDER
		For Each ObjSubFolder In ObjFolder.SubFolders
			FolderName=ObjSubFolder.name

			addFolder = false
			If Len(searchTerms) > 0 Then
				folderFiles = GetFolderFiles(ObjSubFolder, False)
				If folderFiles > 0 Then
					addFolder = true
				End If
			Else
				addFolder = true
			End If

			If addFolder Then
				Html = Html + "<li class=""directory collapsed"">"&_
																	"<a href=""#"" rel="""+(BaseFileDir+FolderName+"/")+""">"&_
																	(FolderName)+"</a></li>"&VBCRLF
				folderCount = folderCount + 1
			End If
		Next
	End If

	loadedFileCount = GetFolderFiles(ObjFolder, True)

	If IsImg = True Then
		delim = "&"
	Else
		delim = "?"
	End If

	If startFile + loadedFileCount < fileCount Then
		Html = Html + "<li class=""load_more""><a href=""" & delim &"page=" & page + 1 & "&count=" & fileCount & """ rel="""+BaseFileDir+""">Load More...</a></li>"
	End If
	
	If folderCount = 0 And loadedFileCount = 0 Then
			Html = Html + "<li class=""message"">No Files or Folders Found.</li>"
	End If
	
	If page = 0 Then
		Html = Html + "</ul>"&VBCRLF
	End If

	fileWord = "Files"
	If IsImg Then
		fileWord = "Images"
	End If

	If page = 0 Then
		endtime = timer()
		Html = Html +  "<div style=""position: fixed; bottom: 0px; left: 0px; right: 0px; padding: 5px; background-color: #eee;"">"
		Html = Html + fileWord & ": <strong>" & fileCount & "</strong> | Folders: <strong>" & folderCount & "</strong>"
		Html = Html + " | Load Time: <strong>" & Round((endtime - starttime), 2) & " s</strong>" 
		Html = Html + "</div>"&VBCRLF
	End If

end if

Response.Write Html
%>
