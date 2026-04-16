<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
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

pcv_intMaxUploads = 6
pcUploadAllowed = false

For idx = 1 To pcv_intMaxUploads
	contentType = UploadRequest.Item("image_" & idx).Item("ContentType")
	filepathname = UploadRequest.Item("image_" & idx).Item("FileName")
	
	if filepathname <> "" Then
		pcUploadAllowed = IsUploadAllowed(filepathname)
	end if
	
	if instr(ucase(contentType), "IMAGE") AND pcUploadAllowed then
		ImageCnt = ImageCnt+1
		
		validFile = ""
		if PPD = "1" then
			filename = "/" & scPcFolder & "/pc/images" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		else
			filename = "../pc/images" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		end if
		
		if not filename = "" then 
			value = UploadRequest.Item("image_" & idx).Item("Value")
			Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
			Set MyFile = ScriptObject.CreateTextFile(Server.mappath(filename))
			For i = 1 to LenB(value)
				MyFile.Write chr(AscB(MidB(value, i, 1)))
			Next
			MyFile.Close
			set myfile = nothing

			validFile = Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		end if
		response.write "images/" & validFile & "<br>"
	else
		if UploadRequest.Item("image_" & idx).Item("FileName") <> "" then
			ImageCnt = ImageCnt + 1
			InValidImage = InValidImage + 1
		end if
	end if
Next

If InValidImage>0 then
	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value.<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	else
		response.write "<br><div align=center><font face=arial size=2>"&InValidImage&" of your "&ImageCnt&" images were not in a valid image format. <br>Invalid image formats are not allowed to be uploaded to the server.<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	end if
	'//END DETAILED MSG
Else
	if ImageCnt>0 then
		call closeDb()
response.redirect "adminimageupl_popup_confirm.html"
	else
		response.write "<br><div align=center><font face=arial size=2>You need to supply at least one file to upload.<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	end if
end if
%>