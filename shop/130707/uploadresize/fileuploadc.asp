<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../../includes/common.asp"-->
<!--#include file="../../includes/languagesCP.asp"-->
<!--#include file="../../includes/common_checkout.asp"-->
<!--#include file="../../includes/pcSanitizeUpload.asp"-->
<%
Session.CodePage = 1252
Response.Expires = 0
response.Buffer = true
Response.Clear

byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)

'//DETAILED MSG
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

Dim invalidTextFile, fileCnt
invalidTextFile = 0
fileCnt = 0

contentType = UploadRequest.Item("one").Item("ContentType")
filepathname = UploadRequest.Item("one").Item("FileName")

if instr(ucase(contentType), "TEXT") AND IsUploadAllowed(filepathname) then
	fileCnt = 1
    
	If LCase(right(filepathname,3)) <> "txt" Then
       invalidTextFile = 1
    else
        if PPD = "1" then
            filename = "/" & scPcFolder & "/pc/library" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
        else
            filename = "../../pc/library" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
        end If
		
        if not filename = "" then
            value = UploadRequest.Item("one").Item("Value")
            Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
            Set MyFile = ScriptObject.CreateTextFile(Server.mappath(filename))
            For i = 1 to LenB(value)
            	MyFile.Write chr(AscB(MidB(value, i, 1)))
            Next
            MyFile.Close
            set myfile = nothing
			
            File1 = Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
        end if
        response.write "library/" & File1 & "<br>"
    end If
else
	if UploadRequest.Item("one").Item("FileName") <> "" then
		fileCnt = 1
		invalidTextFile = 1
	end if
end if


If invalidTextFile > 0 then
	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: " & err.description & "<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value."
	else
		call closeDb()
		response.redirect "FileUploada.asp?msg=" & Server.URLEncode(invalidTextFile & " is not a valid TEXT file (with extension .TXT) and could not be uploaded to the server.")
	end if
	'//END DETAILED MSG
Else
	if fileCnt > 0 Then
		call closeDb()
		response.redirect "FileUploada.asp?s=1&f=" & File1 & "&msg=" & Server.URLEncode("Your file was successfully uploaded.")
	else
		call closeDb()
		response.redirect "FileUploada.asp?msg=" & Server.URLEncode("You need to supply a file to upload.")
	end if
end if
%>