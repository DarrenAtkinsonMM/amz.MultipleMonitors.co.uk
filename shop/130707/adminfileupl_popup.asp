<%@ LANGUAGE="VBSCRIPT" %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<%
Session.CodePage = 1252

response.write "<b>Please wait while your files are being uploaded ...</b><br><br>"
Response.Expires = 0
'response.Buffer=true
'Response.Clear
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
'//DETAILED MSG
dim pcTmpErr, pcTmpErrSize
pcTmpErr = Cstr("")
pcTmpErrSize = Cint(0)

pcTmpErr = err.description

If pcTmpErr&""<>"" Then
	If instr(pcTmpErr, "007") Then
		pcTmpErrSize = 1
	End If
End If
'//END DETAILED MSG

Dim UploadRequest
Set UploadRequest = Server.CreateObject("Scripting.Dictionary")

BuildUploadRequest RequestBin

Dim InValidFile, FileCnt
InValidFile = 0
FileCnt = 0
TempName = month(now()) & day(now()) & year(now()) & hour(now()) & minute(now()) & second(now())

pcv_intMaxUploads = 6
pcUploadAllowed = false

For idx = 1 To pcv_intMaxUploads
	contentType = UploadRequest.Item("file_" & idx).Item("ContentType")
	filepathname = UploadRequest.Item("file_" & idx).Item("FileName")
	checkfile = 0
	
	if filepathname <> "" Then
		pcUploadAllowed = IsUploadAllowed(filepathname)
	end if
	
	if pcUploadAllowed then
		extfile = Right(ucase(filepathname),4)
		if (extfile = ".TXT") or (extfile = ".HTM") or (extfile = ".GIF") or (extfile = ".JPG") or (extfile = ".PNG") or (extfile = ".PDF") or (extfile = ".DOC") or (extfile = ".ZIP") then
			checkfile = 1
		else
			extfile = Right(ucase(filepathname),5)
			if (extfile = ".HTML") then
				checkfile = 1
			end if
		end if
	end if

	if checkfile = 1 then
		FileCnt = FileCnt+1
		
		validFile = ""
		if PPD = "1" then
			filename = "/" & scPcFolder & "/pc/Library" & "/" & TempName & "_" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		else
			filename = "../pc/Library" & "/" & TempName & "_" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		end if
		
		if not filename = "" then 
			value = UploadRequest.Item("file_" & idx).Item("Value")
			Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
			Set MyFile = ScriptObject.CreateTextFile(Server.mappath(filename))
			For i = 1 to LenB(value)
				MyFile.Write chr(AscB(MidB(value, i, 1)))
			Next
			MyFile.Close
			set myfile=nothing

			validFile = TempName & "_" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		end if
		
		if validFile <> "" then
			MySQL = "INSERT INTO pcUploadFiles (pcUpld_IDFeedback, pcUpld_FileName) VALUES (" & session("UIDFeedback") & ", '" & validFile & "')"
			set rstemp = connTemp.execute(mySQL)
		end if
	else
		if UploadRequest.Item("file_" & idx).Item("FileName") <> "" then
			FileCnt = FileCnt+1
			InValidFile = InValidFile+1
		end if
	end if
Next

set rstemp=nothing


If InValidFile > 0 then
	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	else
		response.write "<br><div align=center><font face=arial size=2>"&InValidFile&" of your "&FileCnt&" files were not a valid file type. <br>Invalid file types are not allowed to be uploaded to the server.<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	end if
	'//END DETAILED MSG
Else
	if FileCnt>0 then
		session("uploaded") = "1"%>
		<script type=text/javascript>
			location = "adminfileupl_popup_confirm.asp";
		</script>
		<%
	else
		response.write "<br><div align=center><font face=arial size=2>You need to supply at least one file to upload.<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	end if
end if
%>