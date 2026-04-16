<%@ LANGUAGE="VBSCRIPT" %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<%
Session.CodePage = 1252
Server.ScriptTimeout = 5400
Response.Expires = 0
response.Buffer = true
Response.Clear

byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)

'// DETAILED MSG
dim pcTmpErr, pcTmpErrSize
pcTmpErr = Cstr("")
pcTmpErrSize = Cint(0)

pcTmpErr = err.description

If pcTmpErr & "" <> "" Then
	If instr(pcTmpErr, "007") Then
		pcTmpErrSize = 1
	End If
End If
'//END DETAILED MSG

Dim UploadRequest
Set UploadRequest = Server.CreateObject("Scripting.Dictionary")

BuildUploadRequest RequestBin

Dim InValidCSV
InValidCSV = 0

filepathname = UploadRequest.Item("file1").Item("FileName")
contentType = Right(filepathname,4)

if ((right(ucase(filepathname),4) = ".CSV") or (right(ucase(filepathname),4) = ".XLS")) AND IsUploadAllowed(filepathname) then
	if PPD = "1" then
		filename = "/" & scPcFolder & "/pc/catalog" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
	else
		filename = "../pc/catalog" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
	end if
	
	if not filename = "" then 
		value = UploadRequest.Item("file1").Item("Value")
	
    	Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile = ScriptObject.CreateTextFile(Server.mappath(filename))
		For i = 1 to LenB(value)
	    	MyFile.Write chr(AscB(MidB(value, i, 1)))
		Next
		MyFile.Close
		set myfile = nothing

		lgCSV = Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		session("importfile") = lgCSV
				
		call closeDb()
		response.redirect "ship-index_import.asp?s=1&nextstep=1&msg=" & Server.URLEncode("The data file " & ucase(lgCSV) & " was uploaded successfully.")
	end if
else
	if UploadRequest.Item("file1").Item("FileName") <> "" then
		InValidCSV = InValidCSV + 1
	end if
end if

If InValidCSV > 0 then
	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: " & err.description & "<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value."
	else
		call closeDb()
		response.redirect "ship-index_import1.asp?s=0&msg=" & Server.URLEncode("Invalid file type. Only CSV & XLS files can be uploaded to the server.")
	end if
	'//END DETAILED MSG
Else
	call closeDb()
	response.redirect "ship-index_import1.asp?s=0&msg=" & Server.URLEncode("You did not select a file to upload.")
end if
%>