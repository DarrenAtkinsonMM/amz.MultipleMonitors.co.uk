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

Dim button(39)
button(1) = "add2"
button(2) = "addtocart"
button(3) = "addtowl"
button(4) = "checkout"
button(5) = "cancel"
button(6) = "continueshop"
button(7) = "morebtn"
button(8) = "login"
button(10) = "recalculate"
button(11) = "register"
button(12) = "remove"
button(14) = "back"
button(16) = "viewcartbtn"
button(18) = "customize"
button(19) = "reconfigure"
button(20) = "resetdefault"
button(21) = "savequote"
button(22) = "revorder"
button(23) = "submitquote"
button(24) = "pcv_requestQuote"
button(25) = "pcv_placeOrder"
button(26) = "pcv_checkoutWR"
button(27) = "pcv_processShip"
button(28) = "pcv_finalShip"
button(29) = "pcv_backtoOrder"
button(30) = "pcv_previous"
button(31) = "pcv_next"
button(32) = "crereg"
button(33) = "delreg"
button(34) = "addreg"
button(35) = "updreg"
button(36) = "sendmsgs"
button(37) = "retreg"
button(38) = "yellowupd"
button(39) = "savecart"

Dim validFile(39)

For idx = 1 to Ubound(button)
	If button(idx) <> "" Then
		contentType = UploadRequest.Item(button(idx)).Item("ContentType")
		filepathname = UploadRequest.Item(button(idx)).Item("FileName")
		
		if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
			ImageCnt = ImageCnt + 1
			
			validFile(idx) = ""
			if PPD = "1" then
				filename = "/" & scPcFolder & "/pc/images/pc" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
			else
				filename = "../pc/images/pc" & "/" & Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
			end if
			
			if not filename = "" then
				value = UploadRequest.Item(button(idx)).Item("Value")
				Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
				Set MyFile = ScriptObject.CreateTextFile(Server.mappath(filename))
				For i = 1 to LenB(value)
					MyFile.Write chr(AscB(MidB(value, i, 1)))
				Next
				MyFile.Close
				set myfile = nothing
	
				validFile(idx) = Right(filepathname, Len(filepathname)-InstrRev(filepathname, "\"))
			end if
			response.write "images/" & validFile(idx) & "<br />"
		else
			if UploadRequest.Item(button(idx)).Item("FileName") <> "" then
				ImageCnt = ImageCnt + 1
				InValidImage = InValidImage + 1
			end if
		end if
	End If
Next



If InValidImage > 0 then
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0

		response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value."
		
	else
		If Cint(InValidImage) = Cint(ImageCnt) then
			call closeDb()
			response.redirect "AdminButtons.asp?msg=" & Server.URLEncode(InValidImage & " of your " & ImageCnt & " images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")

		else
			call closeDb()
		
			redirectURL = "dbbuttons.asp?"
			For idx = 1 to Ubound(validFile)
				redirectURL = redirectURL & "file" & idx & "=" & validFile(idx) & "&"
			Next
			
			response.redirect redirectURL & "msg=" & Server.URLEncode(InValidImage & " of your " & ImageCnt & " images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		end if
	end if
Else
	if ImageCnt > 0 then
		call closeDb()
		
		redirectURL = "dbbuttons.asp?"
		For idx = 1 to Ubound(validFile)
			redirectURL = redirectURL & "file" & idx & "=" & validFile(idx) & "&"
		Next
		
		response.redirect redirectURL & "s=1&msg=" & Server.URLEncode("Your images were successfully uploaded.")
	else
		call closeDb()
		response.redirect "AdminButtons.asp?msg=" & Server.URLEncode("You need to supply at least one file to upload.")
	end if
end if
%>