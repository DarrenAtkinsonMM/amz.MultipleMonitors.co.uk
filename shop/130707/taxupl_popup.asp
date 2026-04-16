<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*6*"%>
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

Dim UploadRequest
Set UploadRequest = Server.CreateObject("Scripting.Dictionary")

BuildUploadRequest RequestBin

contentType = UploadRequest.Item("one").Item("ContentType")
filepathname = UploadRequest.Item("one").Item("FileName")

if (instr(ucase(contentType), "APPLICATION") OR instr(ucase(contentType), "TEXT")) AND IsUploadAllowed(filepathname) then
	if instr(ucase(filepathname), ".CSV") then
		ImageCnt = ImageCnt + 1
		
		if PPD = "1" then
			filename = "/" & scPcFolder & "/pc/tax" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		else
			filename = "../pc/tax" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
		end if
		
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
	else
		if UploadRequest.Item("one").Item("FileName") <> "" then
			ImageCnt = ImageCnt + 1
			InValidImage = InValidImage + 1
		end if
	end if
else
	if UploadRequest.Item("one").Item("FileName") <> "" then
		ImageCnt = ImageCnt + 1
		InValidImage = InValidImage + 1
	end if
end if

If InValidImage > 0 then
	response.write "<br><div align=center><font face=arial size=2>Your file does not appear to be in the correct format. <br>Invalid  formats are not allowed to be uploaded to the server.<br><br><a href=""javascript:history.go(-1)""><font face=" & Link & ">Back</font></a><br><br>If you are certain that your file is of the proper format and you are receiving this error, you will need to manually upload your ""Tax Rate File"" file to your server using your ftp client.<br><br>The file needs to be uploaded to the folder:<br><br> <font color=""FF0000"">""/store/pc/tax/""</font></div><p align=""center""><input type=""button"" value=""Close Window"" onClick=""javascript:window.close();""></p></font>"
Else
	if ImageCnt > 0 then
		call closeDb()
		response.redirect "taxupl_popup_confirm.html"
	else
		response.write "<br><div align=center><font face=arial size=2>You need to supply a file to upload.<br><br><a href=""javascript:history.go(-1)""><font face=" & Link & ">Back</font></a></font></div>"
	end if
end if
%>