<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
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

Dim InValidImage, ImageCnt
InValidImage = 0
ImageCnt = 0

Dim icon(9)
icon(1) = "erroricon"
icon(2) = "requiredicon"
icon(3) = "errorfieldicon"
icon(4) = "previousicon"
icon(5) = "nexticon"
icon(6) = "zoom"
icon(7) = "discount"
icon(8) = "arrowUp"
icon(9) = "arrowDown"

Dim validFile(9)

For idx = 1 to Ubound(icon)
	contentType = UploadRequest.Item(icon(idx)).Item("ContentType")
	filepathname = UploadRequest.Item(icon(idx)).Item("FileName")
	
	if instr(ucase(contentType), "IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt = ImageCnt+1
		
		validFile(idx) = ""
		if PPD = "1" then
			filename = "/" & scPcFolder & "/pc/images/pc" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		else
			filename = "../pc/images/pc" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		end if
		
		if not filename = "" then 
			value = UploadRequest.Item(icon(idx)).Item("Value")
			Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
			Set MyFile = ScriptObject.CreateTextFile(Server.mappath(filename))
			For i = 1 to LenB(value)
				MyFile.Write chr(AscB(MidB(value, i, 1)))
			Next
			MyFile.Close
			set myfile = nothing
			
			validFile(idx) = Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
			response.write "images/" & validFile(idx) & "<br />"
		end if
	else
		if UploadRequest.Item("erroricon").Item("FileName") <> "" then
			ImageCnt = ImageCnt + 1
			InValidImage = InValidImage + 1
		end if
	end if
Next


'response.end
If InValidImage > 0 then

	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
		
	else
		If Cint(InValidImage) = Cint(ImageCnt) then
			call closeDb()
			response.redirect "AdminIcons.asp?msg=" & Server.URLEncode(InValidImage & " of your " & ImageCnt & " images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		else
			call closeDb()
			
			redirectURL = "dbicons.asp?"
			For idx = 1 to Ubound(validFile)
				redirectURL = redirectURL & "file" & idx & "=" & validFile(idx) & "&"
			Next
			
			response.redirect redirectURL & "msg=" & Server.URLEncode(InValidImage & " of your " & ImageCnt & " images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		end if
	end if
	'//END DETAILED MSG

Else
	if ImageCnt > 0 then
		call closeDb()
		
		redirectURL = "dbicons.asp?"
		For idx = 1 to Ubound(validFile)
			redirectURL = redirectURL & "file" & idx & "=" & validFile(idx) & "&"
		Next
		
		response.redirect redirectURL & "s=1&msg=" & Server.URLEncode("Your images were successfully uploaded.")
	else
		call closeDb()
		response.redirect "AdminIcons.asp?msg="&Server.URLEncode("You need to supply at least one file to upload.")
	end if
end if
%>